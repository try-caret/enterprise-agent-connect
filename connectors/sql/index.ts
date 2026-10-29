import { createHash } from "node:crypto";
import { z } from "zod";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";
import type {
  Connector,
  ConnectorModule,
  SqlApp,
} from "../../packages/connector-kit/src/connector.js";
import {
  ConnectError,
  type Vendor,
  type VendorConnection,
} from "../../packages/connector-kit/src/types.js";
import {
  assertDriver,
  DEFAULT_PORT,
  openEngine,
  parseStatement,
  type Engine,
  type Rows,
} from "./engines.js";

// The generic module behind every connector.json with an `sql` block: an app reached as a
// database. Each tool runs the one statement its definition holds, with the agent's values
// bound as params.

export const MAX_ANSWER_BYTES = 1024 * 1024;

// What a tool sends its vendor: the values to bind and the agent's answer limits, checked
// inside the transaction so a write outside them is rolled back.
type SqlCall = {
  values: Record<string, unknown>;
  answerLimits: { column: string; values: string[] }[];
};

const message = (error: unknown) => (error instanceof Error ? error.message : String(error));

export function sqlVendor(connector: Connector, open = openEngine): Vendor {
  const app = connector.sql!;
  const port = app.port ?? DEFAULT_PORT[app.engine];
  const tools = new Map(app.tools.map((t) => [t.name, { ...t, statement: parseStatement(t.sql) }]));
  // ponytail: a pool per credential, never closed; an old one only idles at 0 connections.
  const pools = new Map<string, Promise<Engine>>();

  function engine(c: VendorConnection, write: boolean) {
    const target = {
      host: app.host,
      port,
      database: app.database,
      user: c.config[app.auth.user],
      password: c.secret[app.auth.password],
      tls: app.tls,
      // Only SQL Server opens a read-only connection; elsewhere one pool serves both.
      readOnly: app.engine === "sqlserver" && !write,
    };
    const key = createHash("sha256")
      .update(JSON.stringify([target.user, target.password, target.readOnly]))
      .digest("hex");
    let pool = pools.get(key);
    if (!pool) {
      pool = open(app.engine, target);
      pools.set(key, pool);
      pool.catch(() => pools.delete(key));
    }
    return pool;
  }

  return {
    async connect(fields) {
      const config: Record<string, string> = {
        host: `${app.host}:${port}/${app.database}`,
      };
      const secret: Record<string, string> = {};
      for (const f of connector.fields) {
        const value = fields[f.id]?.trim() ?? "";
        if (!value && f.optional) continue;
        if (!value) throw new ConnectError(`Enter ${f.label}`);
        (f.secret ? secret : config)[f.id] = value;
      }
      return { config, secret, accountLabel: config[connector.accountKey] ?? config.host };
    },
    async call(c, request) {
      const tool = tools.get(request.path.slice(1));
      if (!tool) throw new Error(`${connector.id} has no query ${request.path}`);
      const { values, answerLimits } = request.body as SqlCall;
      try {
        const out = await (
          await engine(c, tool.write)
        ).run({
          statement: tool.statement,
          values,
          write: tool.write,
          maxRows: app.maxRows,
          timeoutMs: app.timeoutSeconds * 1000,
          check: ({ rows, rowsAffected }) => {
            if (!answerLimits.length) return;
            const undone = tool.write ? "; nothing was changed" : "";
            // A change returning fewer rows than it made cannot be checked: refused.
            if (rows.length < (rowsAffected ?? 0))
              throw new Denied(
                `This change cannot be checked against this agent's limits${undone}`,
              );
            for (const a of answerLimits)
              if (!rows.every((row) => within(row, a.column, a.values)))
                throw new Denied(
                  `The answer's ${a.column} is outside this agent's limits${undone}`,
                );
          },
        });
        return { status: 200, data: out };
      } catch (error) {
        if (error instanceof Denied) throw error;
        throw new ConnectError(`The database refused the call: ${message(error)}`);
      }
    },
    checks: {
      // A trivial query proving the address, the sign-in and the database name.
      async database(c) {
        try {
          return { values: [await (await engine(c, false)).version()] };
        } catch (error) {
          throw new ConnectError(`Could not reach the database: ${message(error)}`);
        }
      },
    },
  };
}

function within(row: Record<string, unknown>, column: string, values: string[]) {
  if (!Object.hasOwn(row, column)) return false;
  const v = row[column];
  return (
    (typeof v === "string" || typeof v === "number" || typeof v === "bigint") &&
    values.includes(String(v))
  );
}

// Values JSON can carry: dates as ISO text, bytes as base64, big integers as text.
function plain(v: unknown): unknown {
  if (v instanceof Date) return v.toISOString();
  if (Buffer.isBuffer(v) || v instanceof Uint8Array) return Buffer.from(v).toString("base64");
  if (typeof v === "bigint") return v.toString();
  return v;
}

// The rows that fit in maxRows and MAX_ANSWER_BYTES, and whether any were left out.
function trim(rows: Rows, maxRows: number) {
  const out: Rows = [];
  let size = 0;
  for (const row of rows.slice(0, maxRows)) {
    const clean = Object.fromEntries(Object.entries(row).map(([k, v]) => [k, plain(v)]));
    size += Buffer.byteLength(JSON.stringify(clean));
    if (size > MAX_ANSWER_BYTES) break;
    out.push(clean);
  }
  return { rows: out, truncated: out.length < rows.length };
}

const zodType = {
  string: z.string(),
  number: z.number(),
  integer: z.number().int(),
  boolean: z.boolean(),
};

export function sqlTools(app: SqlApp): Toolset {
  return {
    names: app.tools.map((t) => t.name),
    register(server, ctx, allowed) {
      for (const tool of app.tools) {
        if (!allowed.has(tool.name)) continue;
        const shape = Object.fromEntries(
          Object.entries(tool.params).map(([name, p]) => {
            const type = zodType[p.type].describe(p.description);
            return [name, p.required ? type : type.optional()];
          }),
        );
        server.registerTool(
          tool.name,
          {
            description: tool.description,
            inputSchema: z.object(shape).strict(),
            annotations: {
              readOnlyHint: !tool.write,
              destructiveHint: tool.destructive,
              openWorldHint: false,
            },
          },
          (args: Record<string, unknown>) =>
            audited(ctx, tool.name, JSON.stringify(args).slice(0, 300), async (mark) => {
              for (const [name, p] of Object.entries(tool.params)) {
                const allowedValues = p.limit ? listLimit(ctx.limits, p.limit) : null;
                if (!allowedValues) continue;
                // Left out, an optional param would be null, which often means "all of them".
                if (args[name] === undefined)
                  throw new Denied(`${name} is required: this agent is limited to some values`);
                if (!allowedValues.includes(String(args[name])))
                  throw new Denied(`${name} ${String(args[name])} is outside this agent's limits`);
              }
              const call: SqlCall = {
                values: Object.fromEntries(
                  Object.keys(tool.params).map((name) => [name, args[name] ?? null]),
                ),
                answerLimits: tool.answerLimits.flatMap((a) => {
                  const values = listLimit(ctx.limits, a.limit);
                  return values ? [{ column: a.column, values }] : [];
                }),
              };
              let res;
              try {
                res = await ctx.call({ method: "POST", path: `/${tool.name}`, body: call });
              } catch (error) {
                if (error instanceof ConnectError) throw new VendorFailed(502, error.message);
                throw error;
              }
              mark(res.status);
              const { rows, rowsAffected } = res.data as {
                rows: Rows;
                rowsAffected: number | null;
              };
              const answer = trim(rows, app.maxRows);
              return json({
                ...(tool.write ? { rows_affected: rowsAffected } : {}),
                ...answer,
                ...(answer.truncated
                  ? { note: `More rows matched than the ${answer.rows.length} returned` }
                  : {}),
              });
            }),
        );
      }
    },
  };
}

/** Binds to any connector with an `sql` block; createAgentConnect uses it for those. */
export default {
  id: "sql",
  create: ({ connector }) => {
    const app = connector.sql!;
    assertDriver(app.engine, connector.id);
    if (app.engine === "oracle" && app.tls === "insecure")
      throw new Error(
        `${connector.id}: Oracle checks the server's certificate whenever TLS is on: use verify or off`,
      );
    for (const tool of app.tools) {
      const used = new Set(parseStatement(tool.sql).names);
      const declared = Object.keys(tool.params);
      if (used.size !== declared.length || declared.some((n) => !used.has(n)))
        throw new Error(
          `${connector.id}: ${tool.name} declares params ${declared.join(", ") || "none"} but its sql uses ${[...used].map((n) => `:${n}`).join(", ") || "none"}`,
        );
    }
    return { vendor: sqlVendor(connector), toolset: sqlTools(app) };
  },
} satisfies ConnectorModule;
