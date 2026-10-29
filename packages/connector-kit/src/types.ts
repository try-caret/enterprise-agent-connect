// What every app's client gives the rest of the service.

export type Fetch = typeof fetch;

// A grant's value limits by limit id: a list, a text or a map (e.g. a folder per site).
export type Limits = Record<string, string[] | string | Record<string, string>>;

export type VendorConnection = {
  config: Record<string, string>;
  secret: Record<string, string>;
};

export type VendorRequest = {
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  path: string;
  body?: unknown;
  headers?: Record<string, string>;
  // "text": the answer is returned as text, not parsed as JSON (file content).
  responseType?: "text";
};

// `headers` (lower-case names): set by vendors whose tools read them, e.g. a Link header.
export type VendorResponse = { status: number; data: unknown; headers?: Record<string, string> };

export type CheckInput = {
  // The vendor scopes (or token roles) the connection's capabilities need.
  expected: string[];
};

export type Vendor = {
  // Turns the admin's form fields into a working connection, or throws ConnectError. An app
  // signing in through a provider gets that provider's connection and returns only its own
  // config (its secret is the provider's). `previous`: the app's config before a reconnect,
  // for settings kept across it (e.g. sites added after connecting).
  connect(
    fields: Record<string, string>,
    provider?: VendorConnection & { accountLabel: string },
    previous?: Record<string, string>,
  ): Promise<VendorConnection & { accountLabel: string }>;
  call(connection: VendorConnection, request: VendorRequest): Promise<VendorResponse>;
  // Apps whose connector has `siteAccess`: sites added one at a time, each checked live.
  sites?: {
    // Normalized addresses, or ConnectError naming the bad one.
    normalize(raw: string[]): string[];
    // `write`: also prove the app may change files there.
    access(connection: VendorConnection, site: string, write: boolean): Promise<SiteAccess>;
  };
  // Checks named by the connector's `{ check }` probes. `labels`: a label for the admin per
  // value, e.g. a site's name for its id.
  checks?: Record<
    string,
    (
      connection: VendorConnection,
      input: CheckInput,
    ) => Promise<{ values: string[]; warning?: string; labels?: Record<string, string> }>
  >;
};

// Whether the app can open a site. "granted": it read the site's files; "not_granted":
// Microsoft refused (no access given yet); "not_found": no site at that address.
export type SiteAccess = {
  site: string;
  status: "granted" | "not_granted" | "not_found" | "failed";
  name: string | null;
  message: string;
};

// Whether the app can reach one agent's own address, e.g. its mailbox (connector `agentAccess`).
export type AgentAccess = {
  value: string;
  status: "granted" | "not_granted" | "not_found" | "failed";
  message: string;
};

// The sign-in several apps share (connectors/<id>/provider.json).
export type ProviderClient = {
  // Turns the admin's fields into credentials and proves them by signing in. `existing` is what
  // is already kept for this provider, e.g. the key of a certificate made before connecting.
  connect(
    fields: Record<string, string>,
    existing: VendorConnection | null,
  ): Promise<VendorConnection & { accountLabel: string }>;
  // Signs in afresh, e.g. to check a new certificate before switching to it. Throws ConnectError.
  verify(connection: VendorConnection): Promise<void>;
  // Approval of the deployment's own shared app, where the deployment offers it.
  approval?: {
    // The shared app's Application (client) ID: not secret, shown in the admin's setup steps.
    clientId: string;
    consentUrl(input: { state: string; redirectUri: string }): string;
    // After the approval, where the admin signs in to the organization the approval named.
    // The approval's own answer is an unsigned redirect, so the organization is only trusted
    // once this sign-in proves it.
    signInUrl(input: { tenant: string; state: string; redirectUri: string; nonce: string }): string;
    // Redeems the sign-in's code; throws ConnectError unless it proves a member of `tenant`.
    proveTenant(input: {
      tenant: string;
      code: string;
      redirectUri: string;
      nonce: string;
    }): Promise<void>;
    connect(tenant: string): Promise<VendorConnection & { accountLabel: string }>;
  };
};

// A failure the admin can act on; its message is shown to them.
export class ConnectError extends Error {}

// Vendor access tokens, kept until a minute before they expire.
export class TokenCache {
  #tokens = new Map<string, { token: string; expiresAt: number }>();

  async get(key: string, mint: () => Promise<{ token: string; expiresIn: number }>) {
    const hit = this.#tokens.get(key);
    if (hit && hit.expiresAt > Date.now()) return hit.token;
    const { token, expiresIn } = await mint();
    this.#tokens.set(key, { token, expiresAt: Date.now() + (expiresIn - 60) * 1000 });
    return token;
  }

  forget(key: string) {
    this.#tokens.delete(key);
  }
}
