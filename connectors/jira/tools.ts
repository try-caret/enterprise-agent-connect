import { z } from "zod";
import { isProjectKey, JqlRejected, scopeJql } from "./jql.js";
import {
  audited,
  Denied,
  json,
  listLimit,
  VendorFailed,
  type ToolContext,
  type Toolset,
} from "../../packages/connector-kit/src/common.js";

const ISSUE_KEY = /^[A-Z][A-Z0-9_]*-\d+$/;
const SEARCH_FIELDS = [
  "summary",
  "status",
  "project",
  "issuetype",
  "assignee",
  "priority",
  "updated",
];
const ISSUE_FIELDS = [...SEARCH_FIELDS, "reporter", "created", "labels", "description"];
// jira_get_issue returns the most recent comments only.
const RECENT_COMMENTS = 20;

// Who someone is, beyond the display name a person can set to anything: Atlassian's account id
// and type (atlassian, customer for a service project's customers, app). UserDetails in
// Atlassian's v3 reference. No email: an agent with read would get every customer's address.
function person(user: any) {
  return {
    account_id: typeof user?.accountId === "string" ? user.accountId : null,
    account_type: typeof user?.accountType === "string" ? user.accountType : null,
  };
}

// Only a service project (Jira Service Management) has customers, and internal notes they never
// see. The issue's own project field carries the type (ProjectDetails.projectTypeKey in
// Atlassian's v3 reference), so no project call is needed.
const serviceProject = (issue: Issue) => issue.fields.project?.projectTypeKey === "service_desk";

type Issue = { key: string; fields: Record<string, any> };

async function ok(
  ctx: ToolContext,
  mark: (s: number) => void,
  req: Parameters<ToolContext["call"]>[0],
) {
  const res = await ctx.call(req);
  mark(res.status);
  if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, jiraError(res));
  return res.data as any;
}

function jiraError(res: { status: number; data: unknown }): string {
  const d = res.data as { errorMessages?: string[]; errors?: Record<string, string> } | null;
  const detail = [...(d?.errorMessages ?? []), ...Object.values(d?.errors ?? {})].join("; ");
  return `Jira answered ${res.status}${detail ? `: ${detail}` : ""}`;
}

// Atlassian Document Format to plain text, enough for an agent to read.
export function adfText(node: any): string {
  if (!node || typeof node !== "object") return "";
  if (node.type === "text") return node.text ?? "";
  if (node.type === "hardBreak") return "\n";
  if (node.type === "mention") return node.attrs?.text ?? "";
  const inner = (node.content ?? []).map(adfText).join("");
  return ["paragraph", "heading", "listItem", "codeBlock", "blockquote"].includes(node.type)
    ? `${inner}\n`
    : inner;
}

function adfDoc(text: string) {
  return {
    type: "doc",
    version: 1,
    content: text
      .split(/\n{2,}/)
      .filter((p) => p.trim())
      .map((p) => ({ type: "paragraph", content: [{ type: "text", text: p }] })),
  };
}

function summary(issue: Issue) {
  const f = issue.fields;
  return {
    key: issue.key,
    summary: f.summary,
    project: f.project?.key,
    type: f.issuetype?.name,
    status: f.status?.name,
    priority: f.priority?.name ?? null,
    assignee: f.assignee?.displayName ?? null,
    updated: f.updated,
  };
}

function allowedProject(projects: string[] | null, key: string | undefined): boolean {
  return projects === null || (key !== undefined && projects.includes(key));
}

type Transition = { id: string; name: string; to?: { name?: string } };

async function transitions(ctx: ToolContext, mark: (s: number) => void, key: string) {
  await fetchIssue(ctx, mark, key, ["project"]);
  const data = await ok(ctx, mark, { method: "GET", path: `/rest/api/3/issue/${key}/transitions` });
  return data.transitions as Transition[];
}

// Fields a new issue must have that this tool cannot fill. Jira sets the reporter itself.
const FILLED = new Set(["project", "issuetype", "summary", "description", "reporter"]);

async function fetchIssue(
  ctx: ToolContext,
  mark: (s: number) => void,
  key: string,
  fields: string[],
) {
  if (!ISSUE_KEY.test(key)) throw new Denied("Give an issue key such as ABC-123");
  const issue: Issue = await ok(ctx, mark, {
    method: "GET",
    path: `/rest/api/3/issue/${key}?fields=${fields.join(",")}`,
  });
  if (!allowedProject(listLimit(ctx.limits, "projects"), issue.fields.project?.key))
    throw new Denied(`Issue ${key} is in a project this agent is not allowed to use`);
  return issue;
}

// The issue's changelog comes oldest first, at most 100 changes a page.
const CHANGELOG_PAGE = 100;
// Changelog calls one jira_get_issue makes at most: the first page, then back from the newest.
const CHANGELOG_CALLS = 5;

type LabelAdd = {
  label: string;
  author: string | null;
  author_account_id: string | null;
  author_account_type: string | null;
  added: string | null;
};

// Who last added each current label: the newest changelog entry whose labels item has it in
// toString and not in fromString (Jira's labels are space-free, space-separated there). Read from
// the newest change back, CHANGELOG_CALLS calls at most; a label not found (set when the issue
// was created, or older than what was read) gets nulls. Null when the changelog cannot be read:
// the rest of the issue is still worth returning. GET /rest/api/3/issue/{key}/changelog
// (PageBeanChangelog) in Atlassian's v3 reference.
async function labelsAddedBy(
  ctx: ToolContext,
  key: string,
  labels: string[],
): Promise<LabelAdd[] | null> {
  if (!labels.length) return [];
  const page = async (startAt: number, count: number) => {
    const res = await ctx.call({
      method: "GET",
      path: `/rest/api/3/issue/${key}/changelog?startAt=${startAt}&maxResults=${count}`,
    });
    if (res.status < 200 || res.status >= 300) throw new VendorFailed(res.status, jiraError(res));
    return res.data as any;
  };
  const found = new Map<string, LabelAdd>();
  const words = (s: unknown) => (typeof s === "string" ? s.split(/\s+/).filter(Boolean) : []);
  const scan = (values: any[]) => {
    for (const change of [...values].reverse())
      for (const item of change?.items ?? []) {
        if (item?.field !== "labels") continue;
        const before = new Set(words(item.fromString));
        for (const label of words(item.toString))
          if (!before.has(label) && labels.includes(label) && !found.has(label)) {
            const author = person(change.author);
            found.set(label, {
              label,
              author: change.author?.displayName ?? null,
              author_account_id: author.account_id,
              author_account_type: author.account_type,
              added: change.created ?? null,
            });
          }
      }
  };
  try {
    const first = await page(0, CHANGELOG_PAGE);
    const head = (first?.values ?? []) as any[];
    const total = typeof first?.total === "number" ? first.total : head.length;
    let start = Math.max(total, head.length);
    for (let calls = 1; start > head.length && calls < CHANGELOG_CALLS; calls++) {
      if (found.size === labels.length) break;
      const from = Math.max(head.length, start - CHANGELOG_PAGE);
      scan((await page(from, start - from))?.values ?? []);
      start = from;
    }
    if (start <= head.length) scan(head);
  } catch (error) {
    if (!(error instanceof VendorFailed))
      console.error("[agent-connect] jira changelog failed", error);
    return null;
  }
  return labels.map(
    (label) =>
      found.get(label) ?? {
        label,
        author: null,
        author_account_id: null,
        author_account_type: null,
        added: null,
      },
  );
}

export const jiraTools: Toolset = {
  names: [
    "jira_search_issues",
    "jira_get_issue",
    "jira_list_projects",
    "jira_add_comment",
    "jira_list_transitions",
    "jira_transition_issue",
    "jira_create_issue",
    "jira_delete_issue",
  ],

  register(server, ctx, allowed) {
    const readOnly = { readOnlyHint: true, openWorldHint: true } as const;

    if (allowed.has("jira_search_issues"))
      server.registerTool(
        "jira_search_issues",
        {
          description:
            'Search Jira issues with JQL, e.g. `status = "In Progress" ORDER BY updated DESC`. ' +
            "Results are limited to the projects this agent may use.",
          inputSchema: z.object({
            jql: z.string().max(4000).describe("JQL query. ORDER BY is allowed at the end."),
            max_results: z.number().int().min(1).max(50).default(20),
          }),
          annotations: readOnly,
        },
        ({ jql, max_results }) =>
          audited(ctx, "jira_search_issues", jql, async (mark, retarget) => {
            const projects = listLimit(ctx.limits, "projects");
            let scoped: string;
            try {
              scoped = scopeJql(jql, projects);
            } catch (error) {
              if (error instanceof JqlRejected) throw new Denied(error.message);
              throw error;
            }
            retarget(scoped);
            const data = await ok(ctx, mark, {
              method: "POST",
              path: "/rest/api/3/search/jql",
              body: { jql: scoped, maxResults: max_results, fields: SEARCH_FIELDS },
            });
            const issues = (data.issues as Issue[])
              .filter((i) => allowedProject(projects, i.fields.project?.key))
              .map(summary);
            return json({ issues, more: Boolean(data.nextPageToken) });
          }),
      );

    if (allowed.has("jira_get_issue"))
      server.registerTool(
        "jira_get_issue",
        {
          description:
            "Read one Jira issue: fields, description and the 20 most recent comments (fails if " +
            "the comments cannot be read). People come with account_id and account_type " +
            "(atlassian, customer, app). In a service project each comment's internal is true " +
            "for an internal note the customer never sees; elsewhere it is null. labels_added_by " +
            "names, for each current label, who last added it and when (nulls when the issue " +
            "history does not show it, e.g. set at creation); it is null when the history " +
            "could not be read.",
          inputSchema: z.object({ issue_key: z.string().describe("Issue key, e.g. ABC-123.") }),
          annotations: readOnly,
        },
        ({ issue_key }) =>
          audited(ctx, "jira_get_issue", issue_key, async (mark) => {
            const issue = await fetchIssue(ctx, mark, issue_key, ISSUE_FIELDS);
            const f = issue.fields;
            // The newest comments, put back in date order. Each Comment carries jsdPublic,
            // false for a service project's internal note; Atlassian documents it as true for
            // any comment outside a service project, so there it says nothing and is left null.
            const page = await ok(ctx, mark, {
              method: "GET",
              path: `/rest/api/3/issue/${issue_key}/comment?orderBy=-created&maxResults=${RECENT_COMMENTS}`,
            });
            const recent = ((page?.comments ?? []) as any[]).slice(0, RECENT_COMMENTS).reverse();
            const service = serviceProject(issue);
            const reporter = person(f.reporter);
            const labels = ((f.labels ?? []) as unknown[]).filter(
              (l): l is string => typeof l === "string",
            );
            return json({
              ...summary(issue),
              reporter: f.reporter?.displayName ?? null,
              reporter_account_id: reporter.account_id,
              reporter_account_type: reporter.account_type,
              created: f.created,
              labels,
              labels_added_by: await labelsAddedBy(ctx, issue_key, labels),
              description: adfText(f.description).trim(),
              comments: recent.map((c) => {
                const author = person(c.author);
                return {
                  author: c.author?.displayName ?? null,
                  author_account_id: author.account_id,
                  author_account_type: author.account_type,
                  internal: service && typeof c.jsdPublic === "boolean" ? !c.jsdPublic : null,
                  created: c.created,
                  body: adfText(c.body).trim(),
                };
              }),
            });
          }),
      );

    if (allowed.has("jira_list_projects"))
      server.registerTool(
        "jira_list_projects",
        {
          description: "List the Jira projects this agent may use.",
          inputSchema: z.object({}),
          annotations: readOnly,
        },
        () =>
          audited(ctx, "jira_list_projects", null, async (mark) => {
            const projects = listLimit(ctx.limits, "projects");
            // Only projects whose issues the account can open; others are listed by Jira but
            // their searches come back empty.
            const permitted = await ok(ctx, mark, {
              method: "POST",
              path: "/rest/api/3/permissions/project",
              body: { permissions: ["BROWSE_PROJECTS"] },
            });
            const browsable = new Set((permitted.projects as { key: string }[]).map((p) => p.key));
            // Every page: a granted project can sit past the first hundred.
            const all: { key: string; name: string }[] = [];
            for (let startAt = 0; ;) {
              const page = await ok(ctx, mark, {
                method: "GET",
                path: `/rest/api/3/project/search?maxResults=100&startAt=${startAt}`,
              });
              const values = page.values as { key: string; name: string }[];
              all.push(...values);
              if (page.isLast !== false || values.length === 0) break;
              startAt += values.length;
            }
            return json(
              all
                .filter(
                  (p) =>
                    isProjectKey(p.key) && browsable.has(p.key) && allowedProject(projects, p.key),
                )
                .map((p) => ({ key: p.key, name: p.name })),
            );
          }),
      );

    if (allowed.has("jira_add_comment"))
      server.registerTool(
        "jira_add_comment",
        {
          description:
            "Add a comment to a Jira issue. Blank lines separate paragraphs. In a service project " +
            "(Jira Service Management) a comment can be public: the customer sees it and is emailed " +
            "(whether it is depends on the account's Service Management seat). Set internal to add an " +
            "internal note only the service team sees; internal notes exist only in service " +
            "projects. The answer's internal says what Jira saved (null outside a service " +
            "project); false when internal was asked means it was still posted, publicly: never " +
            "retry it.",
          inputSchema: z.object({
            issue_key: z.string().describe("Issue key, e.g. ABC-123."),
            body: z.string().min(1).max(30000),
            internal: z
              .boolean()
              .default(false)
              .describe(
                "An internal note, not shown or sent to the customer. Service projects only.",
              ),
          }),
          annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
        },
        ({ issue_key, body, internal }) =>
          audited(
            ctx,
            "jira_add_comment",
            `${issue_key}${internal ? " (internal)" : ""}`,
            async (mark, retarget) => {
              const issue = await fetchIssue(ctx, mark, issue_key, ["project"]);
              const service = serviceProject(issue);
              if (internal) {
                if (!service)
                  throw new Denied(
                    `${issue_key} is not in a service project, so it has no internal notes. Leave internal off to comment.`,
                  );
                // The platform comment with the sd.public.comment entity property set to
                // {internal: true} is what Jira Service Management's own agent view sends
                // (Atlassian does not document the property). Tried live 2026-10-25: Create request
                // comment (/rest/servicedeskapi/...) answers 403 to a service account without a
                // Service Management agent seat, while this works with plain Jira access and gives
                // jsdPublic false. jsdPublic in the answer is the guard.
                const created = await ok(ctx, mark, {
                  method: "POST",
                  path: `/rest/api/3/issue/${issue_key}/comment`,
                  body: {
                    body: adfDoc(body),
                    properties: [{ key: "sd.public.comment", value: { internal: true } }],
                  },
                });
                if (created.jsdPublic !== false) {
                  // Already posted, shown to the customer and emailed: no error, which an agent
                  // would retry into a second public copy. The audit row records the write.
                  retarget(`${issue_key} (internal asked, posted PUBLIC as comment ${created.id})`);
                  return json({
                    id: created.id,
                    created: created.created,
                    internal: false,
                    warning:
                      `Jira did not keep this as an internal note: comment ${created.id} WAS ` +
                      `posted on ${issue_key} and is PUBLIC, so the customer can see it and may ` +
                      "have been emailed. Do not retry: that would post it again. Tell a person " +
                      "on the service team.",
                  });
                }
                return json({ id: created.id, created: created.created, internal: true });
              }
              const created = await ok(ctx, mark, {
                method: "POST",
                path: `/rest/api/3/issue/${issue_key}/comment`,
                body: { body: adfDoc(body) },
              });
              // As jira_get_issue reads it: a service project's comment can come out internal
              // anyway (seen live from an account without a Service Management seat).
              return json({
                id: created.id,
                created: created.created,
                internal: service ? created.jsdPublic === false : null,
              });
            },
          ),
      );

    const writes = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
    const issueKey = z.string().describe("Issue key, e.g. ABC-123.");

    if (allowed.has("jira_list_transitions"))
      server.registerTool(
        "jira_list_transitions",
        {
          description: "List the statuses a Jira issue can move to now, for jira_transition_issue.",
          inputSchema: z.object({ issue_key: issueKey }),
          annotations: readOnly,
        },
        ({ issue_key }) =>
          audited(ctx, "jira_list_transitions", issue_key, async (mark) =>
            json(
              (await transitions(ctx, mark, issue_key)).map((t) => ({
                id: t.id,
                name: t.name,
                to: t.to?.name ?? null,
              })),
            ),
          ),
      );

    if (allowed.has("jira_transition_issue"))
      server.registerTool(
        "jira_transition_issue",
        {
          description:
            "Move a Jira issue to another status. Give a transition id or name from jira_list_transitions.",
          inputSchema: z.object({
            issue_key: issueKey,
            transition: z.string().min(1).max(200).describe("Transition id or name."),
          }),
          annotations: writes,
        },
        ({ issue_key, transition }) =>
          audited(ctx, "jira_transition_issue", `${issue_key} → ${transition}`, async (mark) => {
            const want = transition.trim().toLowerCase();
            const found = (await transitions(ctx, mark, issue_key)).find(
              (t) => t.id === want || t.name.toLowerCase() === want,
            );
            if (!found)
              throw new Denied(
                `${issue_key} cannot move with "${transition}" now. List the options with jira_list_transitions.`,
              );
            await ok(ctx, mark, {
              method: "POST",
              path: `/rest/api/3/issue/${issue_key}/transitions`,
              body: { transition: { id: found.id } },
            });
            return json({ moved: true, to: found.to?.name ?? found.name });
          }),
      );

    if (allowed.has("jira_create_issue"))
      server.registerTool(
        "jira_create_issue",
        {
          description:
            "Create a Jira issue in a project this agent may use. Blank lines in the description separate paragraphs.",
          inputSchema: z.object({
            project_key: z.string().describe("Project key, e.g. ABC."),
            issue_type: z.string().min(1).max(100).describe("Issue type name, e.g. Task or Bug."),
            summary: z.string().min(1).max(255),
            description: z.string().max(30000).optional(),
          }),
          annotations: writes,
        },
        ({ project_key, issue_type, summary, description }) =>
          audited(ctx, "jira_create_issue", project_key, async (mark) => {
            if (!isProjectKey(project_key)) throw new Denied("Give a project key such as ABC");
            if (!allowedProject(listLimit(ctx.limits, "projects"), project_key))
              throw new Denied(`This agent is not allowed to use project ${project_key}`);
            const meta = await ok(ctx, mark, {
              method: "GET",
              path: `/rest/api/3/issue/createmeta/${project_key}/issuetypes?maxResults=200`,
            });
            const types = meta.issueTypes as { id: string; name: string; subtask: boolean }[];
            const type = types.find(
              (t) => !t.subtask && t.name.toLowerCase() === issue_type.trim().toLowerCase(),
            );
            if (!type)
              throw new Denied(
                `${project_key} has no issue type ${issue_type}. Use one of: ${types
                  .filter((t) => !t.subtask)
                  .map((t) => t.name)
                  .join(", ")}`,
              );
            const fields = (
              await ok(ctx, mark, {
                method: "GET",
                path: `/rest/api/3/issue/createmeta/${project_key}/issuetypes/${type.id}?maxResults=200`,
              })
            ).fields as {
              fieldId: string;
              name: string;
              required: boolean;
              hasDefaultValue: boolean;
            }[];
            const needed = fields.filter((f) => f.required && !f.hasDefaultValue);
            if (!description && needed.some((f) => f.fieldId === "description"))
              throw new Denied(`${project_key} needs a description for a ${type.name}`);
            const unfillable = needed.filter((f) => !FILLED.has(f.fieldId));
            if (unfillable.length)
              throw new Denied(
                `A ${type.name} in ${project_key} needs fields this tool cannot fill: ${unfillable
                  .map((f) => f.name)
                  .join(", ")}`,
              );
            const created = await ok(ctx, mark, {
              method: "POST",
              path: "/rest/api/3/issue",
              body: {
                fields: {
                  project: { key: project_key },
                  issuetype: { id: type.id },
                  summary,
                  ...(description ? { description: adfDoc(description) } : {}),
                },
              },
            });
            return json({ key: created.key, id: created.id });
          }),
      );

    if (allowed.has("jira_delete_issue"))
      server.registerTool(
        "jira_delete_issue",
        {
          description:
            "Delete a Jira issue for good. An issue with sub-tasks is only deleted with delete_subtasks set, which deletes them too.",
          inputSchema: z.object({
            issue_key: issueKey,
            delete_subtasks: z.boolean().default(false),
          }),
          annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
        },
        ({ issue_key, delete_subtasks }) =>
          audited(ctx, "jira_delete_issue", issue_key, async (mark) => {
            await fetchIssue(ctx, mark, issue_key, ["project"]);
            await ok(ctx, mark, {
              method: "DELETE",
              path: `/rest/api/3/issue/${issue_key}?deleteSubtasks=${delete_subtasks}`,
            });
            return json({ deleted: issue_key });
          }),
      );
  },
};
