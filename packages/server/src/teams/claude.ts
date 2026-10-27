import Anthropic from "@anthropic-ai/sdk";
import { brand } from "../../../connector-kit/src/branding.js";
import { ConnectError } from "../../../connector-kit/src/types.js";

// What Teams triggers need from Claude Managed Agents, with the organization's own key.

export type SessionState = {
  status: "rescheduling" | "running" | "idle" | "terminated";
  costCents: number | null;
  // Idle because it reached its spending limit.
  atBudget: boolean;
};

export type Claude = {
  // Proves the key works (ConnectError when Claude refuses it): the Managed Agents it can use.
  listAgents(): Promise<{ id: string; name: string }[]>;
  // The agent's MCP server addresses, or null when there is no such (live) agent.
  agent(id: string): Promise<{ mcpUrls: string[] } | null>;
  environmentExists(id: string): Promise<boolean>;
  // The MCP server addresses of every credential in the vault (allUrls) and of its bearer keys
  // (bearerUrls), or null when there is no vault.
  vault(id: string): Promise<{ allUrls: string[]; bearerUrls: string[] } | null>;
  createSession(input: {
    agentId: string;
    environmentId: string;
    vaultIds: string[];
    budgetCents: number;
    metadata: Record<string, string>;
    text: string;
  }): Promise<string>;
  // The next message in a reused session, with its spending limit raised so this answer gets
  // its own budgetCents on top of what the session already spent.
  sendMessage(sessionId: string, text: string, budgetCents: number): Promise<void>;
  // Null when the session is gone (archived, deleted, or under another key).
  session(sessionId: string): Promise<SessionState | null>;
  interrupt(sessionId: string): Promise<void>;
};

export type ClaudeFactory = (apiKey: string) => Claude;

/** The MCP address that takes only this agent's key (the shared /mcp takes any agent's). */
export const agentMcpUrl = (base: string, agentId: string) => `${base}/mcp/${agentId}`;

/** Whether a URL is this agent's own MCP address, another address of this service that takes
 *  an agent key (the shared /mcp, another agent's, or any with a query or fragment), or not
 *  this service's. Parsed, so a default port, dot segments and case do not hide one. */
export function serviceAddress(base: string, agentId: string) {
  const b = new URL(base);
  const root = `${b.pathname.replace(/\/+$/, "")}/mcp`.toLowerCase();
  const own = `${root}/${agentId.toLowerCase()}`;
  return (url: string): "own" | "other" | null => {
    const u = URL.parse(url);
    if (!u || u.host !== b.host) return null;
    const path = u.pathname.replace(/\/+$/, "").toLowerCase();
    if (path === own && !u.search && !u.hash) return "own";
    return u.search || u.hash || path === root || path.startsWith(`${root}/`) ? "other" : null;
  };
}

/** A trigger's Claude agent answers Teams only through this agent's own MCP address, and its
 *  vaults hold this agent's key and no other address of this service: otherwise whoever the
 *  trigger admits would act under another key's grants. Checked when the trigger is saved and
 *  again before each answer, since both can be edited in Claude's console afterwards. */
export async function verifyTrigger(
  claude: Claude,
  base: string,
  agentId: string,
  trigger: { managedAgentId: string; vaultIds: string[] },
) {
  const mcp = agentMcpUrl(base, agentId);
  const kind = serviceAddress(base, agentId);
  const agent = await claude.agent(trigger.managedAgentId);
  if (!agent)
    throw new ConnectError("Claude has no agent with that id under your organization's key");
  if (!agent.mcpUrls.some((u) => kind(u) === "own"))
    throw new ConnectError(
      `That Claude agent does not have ${brand().productName}'s MCP server (${mcp})`,
    );
  const extra = agent.mcpUrls.find((u) => kind(u) === "other");
  if (extra)
    throw new ConnectError(
      `That Claude agent also has another ${brand().productName} address (${extra}). Keep only ${mcp}.`,
    );
  let holdsKey = false;
  for (const id of trigger.vaultIds) {
    const vault = await claude.vault(id);
    if (!vault) throw new ConnectError(`Claude has no vault ${id} under your organization's key`);
    const foreign = vault.allUrls.find((u) => kind(u) === "other");
    if (foreign)
      throw new ConnectError(
        `Vault ${id} also holds a key for another ${brand().productName} address (${foreign}). Use a vault with only this agent's key.`,
      );
    holdsKey ||= vault.bearerUrls.some((u) => kind(u) === "own");
  }
  if (!holdsKey)
    throw new ConnectError(
      `None of these vaults holds a key for ${mcp}. Add this agent's ${brand().productName} key to one.`,
    );
}

const notFound = (error: unknown) => error instanceof Anthropic.NotFoundError;

export const claudeClient: ClaudeFactory = (apiKey) => {
  const api = new Anthropic({ apiKey });
  return {
    async listAgents() {
      const out: { id: string; name: string }[] = [];
      try {
        for await (const a of api.beta.agents.list({ limit: 100 }))
          if (!a.archived_at) out.push({ id: a.id, name: a.name });
      } catch (error) {
        if (
          error instanceof Anthropic.AuthenticationError ||
          error instanceof Anthropic.PermissionDeniedError
        )
          throw new ConnectError("Claude did not accept this key");
        throw error;
      }
      return out;
    },
    async agent(id) {
      try {
        const agent = await api.beta.agents.retrieve(id);
        return agent.archived_at ? null : { mcpUrls: agent.mcp_servers.map((m) => m.url) };
      } catch (error) {
        if (notFound(error)) return null;
        throw error;
      }
    },
    async vault(id) {
      try {
        const vault = await api.beta.vaults.retrieve(id);
        if (vault.archived_at) return null;
        const allUrls: string[] = [];
        const bearerUrls: string[] = [];
        for await (const c of api.beta.vaults.credentials.list(id)) {
          if (c.archived_at || !("mcp_server_url" in c.auth)) continue;
          allUrls.push(c.auth.mcp_server_url);
          if (c.auth.type === "static_bearer") bearerUrls.push(c.auth.mcp_server_url);
        }
        return { allUrls, bearerUrls };
      } catch (error) {
        if (notFound(error)) return null;
        throw error;
      }
    },
    async environmentExists(id) {
      try {
        await api.beta.environments.retrieve(id);
        return true;
      } catch (error) {
        if (notFound(error)) return false;
        throw error;
      }
    },
    async createSession(input) {
      const session = await api.beta.sessions.create({
        agent: input.agentId,
        environment_id: input.environmentId,
        vault_ids: input.vaultIds,
        budget: {
          type: "limit",
          max_list_cost: { amount: String(input.budgetCents), currency: "USD" },
        },
        metadata: input.metadata,
        initial_events: [{ type: "user.message", content: [{ type: "text", text: input.text }] }],
      });
      return session.id;
    },
    async sendMessage(sessionId, text, budgetCents) {
      const s = await api.beta.sessions.retrieve(sessionId);
      const spent = Number(s.usage?.list_cost?.amount ?? 0);
      await api.beta.sessions.update(sessionId, {
        budget: {
          type: "limit",
          max_list_cost: { amount: String(spent + budgetCents), currency: "USD" },
        },
      });
      await api.beta.sessions.events.send(sessionId, {
        events: [{ type: "user.message", content: [{ type: "text", text }] }],
      });
    },
    async session(sessionId) {
      let s;
      try {
        s = await api.beta.sessions.retrieve(sessionId);
      } catch (error) {
        if (notFound(error)) return null;
        throw error;
      }
      const amount = s.usage?.list_cost?.amount;
      const costCents = amount == null ? null : Number(amount);
      const limit = s.budget ? Number(s.budget.max_list_cost.amount) : null;
      return {
        status: s.status,
        costCents,
        atBudget: s.status === "idle" && limit !== null && (costCents ?? 0) >= limit,
      };
    },
    async interrupt(sessionId) {
      await api.beta.sessions.events.send(sessionId, { events: [{ type: "user.interrupt" }] });
    },
  };
};
