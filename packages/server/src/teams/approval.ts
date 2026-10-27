// Approval cards: what an agent posts to ask one person to approve or reject, how the card looks
// before and after the answer, and the answer Teams expects to a click. The click itself is
// handled by the runtime (runtime.ts), after the Bot Framework token checks (jwt.ts: the token
// is checked, the body is not signed). docs/teams.md "Approval cards".
//
// Universal Actions (Action.Execute): a click is an `adaptiveCard/action` invoke to the bot,
// answered in the same HTTP response with a card that replaces the clicker's view; everyone
// else's view changes when the bot edits its message (PUT …/activities/{id}). Sources in
// docs/teams.md [28]-[31].

import { createHash, randomBytes } from "node:crypto";

export const APPROVAL_DAYS = 7;
export const MAX_COMMENT = 1000;

export type ApprovalFact = { name: string; value: string };

export type ApprovalRow = {
  id: string;
  agentId: string;
  botId: string;
  conversationId: string;
  serviceUrl: string;
  threadId: string | null;
  activityId: string | null;
  approverAadId: string;
  approverName: string;
  title: string;
  body: string;
  facts: ApprovalFact[];
  status: "pending" | "approved" | "rejected";
  decidedByAadId: string | null;
  decidedByName: string | null;
  decidedAt: string | null;
  comment: string | null;
  createdAt: string;
  expiresAt: string;
  // Still pending past expiresAt.
  expired: boolean;
};

// What the agent holds. A random value, kept only as its hash; the card carries the row id
// instead, which answers nothing by itself (only the approver's own click counts).
export const newApprovalToken = () => `appr_${randomBytes(32).toString("base64url")}`;
export const hashApprovalToken = (token: string) =>
  createHash("sha256").update(token).digest("hex");
export const looksLikeApprovalToken = (token: string) => /^appr_[A-Za-z0-9_-]{43}$/.test(token);

/** pending, approved, rejected, or expired (pending past its time). */
export const approvalStatus = (a: Pick<ApprovalRow, "status" | "expired">) =>
  a.status === "pending" && a.expired ? ("expired" as const) : a.status;

// Adaptive Cards' DATE()/TIME() take RFC 3339 without fractional seconds, shown in the reader's
// own time zone.
const when = (iso: string) => {
  const t = iso.replace(/\.\d+Z$/, "Z");
  return `{{DATE(${t}, SHORT)}} {{TIME(${t})}}`;
};

const CARD = {
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  type: "AdaptiveCard",
  // Universal Actions need 1.4 or later; Teams' own page says 1.5.
  version: "1.5",
} as const;

/** The card as it should look now: buttons while pending, the decision once made. */
export function approvalCard(a: ApprovalRow): Record<string, unknown> {
  const status = approvalStatus(a);
  const body: Record<string, unknown>[] = [
    { type: "TextBlock", text: a.title, weight: "Bolder", size: "Medium", wrap: true },
    { type: "TextBlock", text: a.body, wrap: true },
  ];
  if (a.facts.length)
    body.push({
      type: "FactSet",
      facts: a.facts.map((f) => ({ title: f.name, value: f.value })),
    });
  if (status === "pending") {
    body.push(
      {
        type: "TextBlock",
        text: `Waiting for ${a.approverName}. Only they can answer.`,
        isSubtle: true,
        wrap: true,
      },
      {
        type: "Input.Text",
        id: "comment",
        placeholder: "Comment (optional)",
        isMultiline: true,
        maxLength: MAX_COMMENT,
      },
      {
        // Wrapped in an ActionSet, as Microsoft recommends for Action.Execute.
        type: "ActionSet",
        actions: [
          {
            type: "Action.Execute",
            title: "Approve",
            verb: "approve",
            style: "positive",
            data: { approval: a.id },
          },
          {
            type: "Action.Execute",
            title: "Reject",
            verb: "reject",
            style: "destructive",
            data: { approval: a.id },
          },
        ],
      },
    );
  } else if (status === "expired") {
    body.push({
      type: "TextBlock",
      text: `Expired ${when(a.expiresAt)} without an answer from ${a.approverName}.`,
      weight: "Bolder",
      wrap: true,
    });
  } else {
    body.push({
      type: "TextBlock",
      text: `${status === "approved" ? "Approved" : "Rejected"} by ${a.decidedByName ?? a.approverName} ${when(a.decidedAt!)}`,
      weight: "Bolder",
      color: status === "approved" ? "Good" : "Attention",
      wrap: true,
    });
    if (a.comment) body.push({ type: "TextBlock", text: a.comment, wrap: true, isSubtle: true });
  }
  return { ...CARD, body };
}

export const cardAttachment = (card: Record<string, unknown>) => ({
  contentType: "application/vnd.microsoft.card.adaptive",
  content: card,
});

// The body of the HTTP 200 answering an adaptiveCard/action invoke.
export const invokeCard = (card: Record<string, unknown>) => ({
  statusCode: 200,
  type: "application/vnd.microsoft.card.adaptive",
  value: card,
});
export const invokeMessage = (text: string) => ({
  statusCode: 200,
  type: "application/vnd.microsoft.activity.message",
  value: text,
});
export const invokeError = (message: string) => ({
  statusCode: 400,
  type: "application/vnd.microsoft.error",
  value: { code: "BadRequest", message },
});
