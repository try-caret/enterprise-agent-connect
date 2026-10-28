import type { components, operations } from "./schema.js";

// The admin API's shapes, by name, from the generated schema (openapi.json).

type S = components["schemas"];

/** The JSON a route answers with on success. */
export type Answer<K extends keyof operations> = operations[K]["responses"] extends infer R
  ? {
      [C in keyof R]: C extends 200 | 201
        ? R[C] extends { content: { "application/json": infer J } }
          ? J
          : never
        : never;
    }[keyof R]
  : never;

/** The JSON body a route takes. */
export type Body<K extends keyof operations> = operations[K]["requestBody"] extends
  { content: { "application/json": infer J } } | undefined
  ? J
  : never;

export type Organization = S["Organization"];
export type Connector = S["Connector"];
export type Provider = S["Provider"];
export type ProviderConnection = S["ProviderConnection"];
export type Connection = S["Connection"];
export type CertificateInfo = S["CertificateInfo"];
export type Check = S["Check"];
export type CheckOutcome = S["CheckOutcome"];
export type SignInCheck = S["SignInCheck"];
export type SiteAccess = S["SiteAccess"];
export type AgentAccess = S["AgentAccess"];
export type Limits = S["Limits"];
export type Grant = S["Grant"];
export type Agent = S["Agent"];
export type NewAgent = S["NewAgent"];
export type AuditRow = S["AuditRow"];
export type LiveCheck = S["LiveCheck"];
export type TeamsBot = S["TeamsBot"];
export type TeamsBotRecord = S["TeamsBotRecord"];
export type TeamsPlace = S["TeamsPlace"];
export type AddStep = S["AddStep"];
export type InstallTarget = S["InstallTarget"];
export type TeamsTrigger = S["TeamsTrigger"];
export type TeamsTriggerInput = Body<"setTeamsTrigger">;
export type TeamsRun = S["TeamsRun"];
export type EventSubscription = S["EventSubscription"];
export type EventSubscriptionInput = Body<"subscribeEvents">;
export type EventSubscribed = S["EventSubscribed"];
export type EventTrigger = S["EventTrigger"];
export type EventTriggerInput = Body<"setEventTrigger">;
export type EventRun = S["EventRun"];
export type ClaudeKey = S["ClaudeKey"];
export type ClaudeAgent = S["ClaudeAgent"];

// Parts of a connector and a provider.
export type Step = Connector["setup"][number];
export type AgentStep = Provider["addAgent"][number];
export type Setting = NonNullable<Step["settings"]>[number];
export type ChosenOption = NonNullable<Step["option"]>;
/** A connect form field; a provider's may be left empty (`optional`). */
export type Field = Connector["fields"][number] & { optional?: boolean };
export type FieldOption = NonNullable<Field["options"]>[number];
export type Limit = Connector["limits"][number];
export type Capability = Connector["capabilities"][number];
export type Probe = Connector["verify"];
