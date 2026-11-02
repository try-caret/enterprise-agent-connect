import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { after, before, test } from "node:test";
import type { McpServer } from "@modelcontextprotocol/server";
import postgres from "postgres";
import type { ToolContext } from "../../connector-kit/src/common.ts";
import { connectorSchema, type Connector } from "../../connector-kit/src/connector.ts";
import { ConnectError } from "../../connector-kit/src/types.ts";
import { assertToolsImplemented } from "../src/service.ts";
import sqlConnector, { sqlTools, sqlVendor as realVendor } from "../../../connectors/sql/index.ts";
import {
  assertDriver,
  mySqlEngine,
  openEngine,
  oracleEngine,
  parseStatement,
  sqlServerEngine,
  type Engine,
  type RunInput,
  type Target,
} from "../../../connectors/sql/engines.ts";

// The generic sql module: the Postgres engine against a real database made from the example's
// schema (TEST_DATABASE_URL, see store.test.ts; skipped without one), the other engines
// against fake drivers.

const example = JSON.parse(
  readFileSync(
    new URL("../../../examples/legacy-orders/connectors/erp/connector.json", import.meta.url),
    "utf8",
  ),
);
const SCHEMA_SQL = readFileSync(
  new URL("../../../examples/legacy-orders/schema.sql", import.meta.url),
  "utf8",
);
const url = process.env.TEST_DATABASE_URL;
const DB = "ac_sql_test";
const server = url ? new URL(url) : null;
const admin = url ? postgres(url, { max: 1, onnotice: () => {} }) : null;
const db = server ? postgres({ ...pgOptions(), database: DB, max: 1, onnotice: () => {} }) : null;

function pgOptions() {
  return { host: server!.hostname, port: Number(server!.port || 5432), username: "postgres" };
}

before(async () => {
  if (!admin) return;
  await admin.unsafe(`drop database if exists ${DB}`);
  await admin.unsafe("drop role if exists agents");
  await admin.unsafe(`create database ${DB}`);
  await db!.unsafe(SCHEMA_SQL);
});
// Pools stay open for the process's life; the tests close theirs so the run can end.
const opened: Promise<Engine>[] = [];
const sqlVendor = (connector: Connector) =>
  realVendor(connector, (engine, target) => {
    const pool = openEngine(engine, target);
    opened.push(pool);
    return pool;
  });
after(async () => {
  await Promise.all(opened.map(async (p) => (await p).close()));
  await db?.end();
  await admin?.end();
});

function erp(sql: Record<string, unknown> = {}, extra: Record<string, unknown> = {}) {
  return connectorSchema.parse({
    ...example,
    ...extra,
    sql: {
      ...example.sql,
      ...(server ? { host: server.hostname, port: Number(server.port || 5432), database: DB } : {}),
      ...sql,
    },
  });
}

const AGENT = { config: { db_user: "agents" }, secret: { db_password: "agents-example" } };

// Registers the tools on a server that only records, so a test can call them directly.
async function run(
  connector: Connector,
  tool: string,
  args: Record<string, unknown>,
  { limits = {} as ToolContext["limits"], vendor = sqlVendor(connector) } = {},
) {
  const handlers = new Map<string, (a: Record<string, unknown>) => Promise<unknown>>();
  const recorder = {
    registerTool: (name: string, _c: unknown, h: never) => handlers.set(name, h),
  } as unknown as McpServer;
  const audit: { decision: string; reason: string | null; vendorStatus: number | null }[] = [];
  const ctx = {
    limits,
    config: AGENT.config,
    call: (r) => vendor.call(AGENT, r),
    audit: async (e) => void audit.push(e),
  } as ToolContext;
  const toolset = sqlTools(connector.sql!);
  toolset.register(recorder, ctx, new Set(toolset.names));
  const result = (await handlers.get(tool)!(args)) as {
    isError?: boolean;
    content: { text: string }[];
  };
  const text = result.content[0].text;
  return { result, text, data: result.isError ? null : JSON.parse(text), audit };
}

test("the example connector loads and runs on the generic module", () => {
  const connector = erp();
  assertToolsImplemented([{ connector, ...sqlConnector.create({ connector } as never) }]);
});

test("a destructive tool needs a destructive capability", () => {
  const connector = erp();
  const soft = {
    ...connector,
    capabilities: connector.capabilities.map((c) => ({ ...c, destructive: false })),
  };
  assert.throws(
    () =>
      assertToolsImplemented([
        { connector: soft, toolset: sqlTools(soft.sql!), vendor: sqlVendor(soft) },
      ]),
    /capability cancel has erp_cancel_order/,
  );
});

test("sql definitions are validated", () => {
  const caps = example.capabilities as { id: string; required?: boolean }[];
  // A write in a required capability.
  assert.throws(
    () =>
      erp({}, { capabilities: caps.map((c) => (c.id === "notes" ? { ...c, required: true } : c)) }),
    /write tool is in no required capability/,
  );
  // The password must be a secret field.
  assert.throws(
    () =>
      erp(
        {},
        {
          fields: [
            { id: "db_user", label: "User", secret: false },
            { id: "db_password", label: "Password", secret: false },
          ],
        },
      ),
    /auth names its fields/,
  );
  assert.throws(
    () => erp({}, { verify: { method: "GET", path: "/", label: "x", show: "x" } }),
    /probes are checks/,
  );
  const http = {
    baseUrl: "https://erp.example.com",
    auth: { type: "bearer", token: "db_password" },
    tools: [{ name: "erp_ping", description: "Ping.", method: "GET", path: "/ping" }],
  };
  assert.throws(() => erp({}, { http }), /at most one of http, soap, odata and sql/);
  const tools = example.sql.tools as Record<string, unknown>[];
  assert.throws(
    () => erp({ tools: [{ ...tools[0], destructive: true }, ...tools.slice(1)] }),
    /A destructive tool is a write/,
  );
  // Params declared must be exactly the ones the statement uses.
  const orderNo = (tools[1].params as Record<string, object>).order_no;
  for (const params of [{}, { order_no: orderNo, extra: orderNo }]) {
    const connector = erp({ tools: [tools[0], { ...tools[1], params }, ...tools.slice(2)] });
    assert.throws(
      () => sqlConnector.create({ connector } as never),
      /erp_get_order declares params/,
    );
  }
  assert.throws(
    () => sqlConnector.create({ connector: erp({ engine: "oracle", tls: "insecure" }) } as never),
    /Oracle checks the server's certificate/,
  );
});

test("a missing driver fails startup naming what to install", () => {
  const missing = () => {
    throw new Error("not found");
  };
  assert.throws(() => assertDriver("sqlserver", "erp", missing), /npm install mssql/);
  assert.throws(() => assertDriver("mysql", "erp", missing), /npm install mysql2/);
  assert.throws(() => assertDriver("oracle", "erp", missing), /npm install oracledb/);
  assertDriver("postgres", "erp", missing);
});

test("params are found outside quotes, comments and casts", () => {
  const s = parseStatement(
    "select ':no', \":no\", `:no` -- :no\n/* :no */ from t where a = :a and b = :b::int and c = :a",
  );
  assert.deepEqual(s.names, ["a", "b", "a"]);
  assert.equal(s.parts.length, 4);
  assert.match(s.parts[0], /^select ':no', ":no", `:no` -- :no\n\/\* :no \*\/ from t where a = $/);
  assert.equal(s.parts[2], "::int and c = ");
});

// Postgres, for real.

const pg = { skip: !url && "TEST_DATABASE_URL not set" };

test("the check proves the connection, and a wrong database is refused plainly", pg, async () => {
  const vendor = sqlVendor(erp());
  const conn = await vendor.connect({ db_user: "agents", db_password: "agents-example" });
  assert.equal(conn.accountLabel, `${server!.hostname}:${server!.port}/${DB}`);
  assert.equal(conn.secret.db_password, "agents-example");
  assert.match((await vendor.checks!.database(conn, { expected: [] })).values[0], /PostgreSQL/);
  const wrong = sqlVendor(erp({ database: "no_such_db" }));
  await assert.rejects(
    wrong.checks!.database(conn, { expected: [] }),
    (e) =>
      e instanceof ConnectError && /Could not reach the database: .*no_such_db/.test(e.message),
  );
  await assert.rejects(vendor.connect({ db_user: "agents" }), /Enter Password/);
});

test("a limited param must be one of the agent's values, and may not be left out", pg, async () => {
  const limits = { customers: ["ACME"] };
  const acme = await run(erp(), "erp_list_orders", { customer: "ACME" }, { limits });
  assert.deepEqual(
    acme.data.rows.map((r: { order_no: number }) => r.order_no),
    [1002, 1001],
  );
  assert.equal(acme.data.truncated, false);
  assert.equal(acme.audit[0].decision, "allowed");
  const open = await run(
    erp(),
    "erp_list_orders",
    { customer: "ACME", status: "open" },
    { limits },
  );
  assert.deepEqual(open.data.rows, [
    {
      order_no: 1002,
      customer: "ACME",
      status: "open",
      total: "89.90",
      created_on: "2026-10-20T00:00:00.000Z",
    },
  ]);
  for (const args of [{ customer: "GLOBEX" }, {}]) {
    const { result, audit } = await run(erp(), "erp_list_orders", args, { limits });
    assert.equal(result.isError, true);
    assert.equal(audit[0].decision, "denied");
  }
});

test("values are bound, never pasted into the statement", pg, async () => {
  const { data } = await run(erp(), "erp_list_orders", { customer: "ACME' or '1'='1" });
  assert.deepEqual(data.rows, []);
  const { result } = await run(erp(), "erp_get_order", { order_no: "1001 or 1=1" });
  assert.equal(result.isError, true);
});

test("an answer limit refuses a row outside the agent's values", pg, async () => {
  const limits = { customers: ["ACME"] };
  const mine = await run(erp(), "erp_get_order", { order_no: 1001 }, { limits });
  assert.equal(mine.data.rows[0].customer_name, "Acme Corp");
  const theirs = await run(erp(), "erp_get_order", { order_no: 1003 }, { limits });
  assert.equal(theirs.result.isError, true);
  assert.doesNotMatch(theirs.text, /Globex/);
  assert.equal(theirs.audit[0].decision, "denied");
  // Without values for the limit, the answer is passed on.
  assert.equal((await run(erp(), "erp_get_order", { order_no: 1003 })).data.rows.length, 1);
});

test("a write outside the agent's limits is rolled back", pg, async () => {
  const limits = { customers: ["ACME"] };
  const notes = async () => (await db!`select count(*)::int as n from order_notes`)[0].n;
  const added = await run(erp(), "erp_add_note", { order_no: 1002, note: "call back" }, { limits });
  assert.equal(added.data.rows_affected, 1);
  assert.equal(added.data.rows[0].customer, "ACME");
  assert.equal(await notes(), 1);

  const refused = await run(erp(), "erp_add_note", { order_no: 1003, note: "x" }, { limits });
  assert.equal(refused.result.isError, true);
  assert.equal(refused.audit[0].decision, "denied");
  assert.equal(await notes(), 1);

  const cancel = await run(erp(), "erp_cancel_order", { order_no: 1003 }, { limits });
  assert.equal(cancel.result.isError, true);
  assert.match(cancel.text, /outside this agent's limits; nothing was changed/);
  const status = async (n: number) =>
    (await db!`select status from orders where order_no = ${n}`)[0].status;
  assert.equal(await status(1003), "open");
  const done = await run(erp(), "erp_cancel_order", { order_no: 1002 }, { limits });
  assert.equal(done.data.rows_affected, 1);
  assert.equal(await status(1002), "cancelled");
  // Shipped orders are left alone by the statement itself.
  assert.equal((await run(erp(), "erp_cancel_order", { order_no: 1001 })).data.rows_affected, 0);
});

test("a write whose rows cannot be checked is refused under a limit", pg, async () => {
  const tools = example.sql.tools as Record<string, unknown>[];
  const blind = erp({
    tools: [
      ...tools.slice(0, 3),
      { ...tools[3], sql: "update orders set status = status where order_no = :order_no" },
    ],
  });
  const limits = { customers: ["GLOBEX"] };
  const { result, text } = await run(blind, "erp_cancel_order", { order_no: 1003 }, { limits });
  assert.equal(result.isError, true);
  assert.match(text, /cannot be checked/);
  assert.equal((await run(blind, "erp_cancel_order", { order_no: 1003 })).data.rows_affected, 1);
});

test("a read cannot change anything", pg, async () => {
  const tools = example.sql.tools as Record<string, unknown>[];
  const sneaky = erp({
    tools: [
      {
        ...tools[1],
        sql: "update orders set status = 'lost' where order_no = :order_no returning customer_code as customer",
      },
      ...tools.slice(2),
      { ...tools[0], name: "erp_list_orders" },
    ],
  });
  const { result, text } = await run(sneaky, "erp_get_order", { order_no: 1001 });
  assert.equal(result.isError, true);
  assert.match(text, /read-only transaction/);
  assert.equal((await db!`select status from orders where order_no = 1001`)[0].status, "shipped");
});

test("answers are cut at maxRows, and a slow statement times out", pg, async () => {
  const { data } = await run(erp({ maxRows: 1 }), "erp_list_orders", { customer: "ACME" });
  assert.equal(data.rows.length, 1);
  assert.equal(data.truncated, true);
  assert.match(data.note, /More rows matched/);
  const tools = example.sql.tools as Record<string, unknown>[];
  const slow = erp({
    timeoutSeconds: 1,
    tools: [
      { ...tools[1], sql: "select pg_sleep(3), 'ACME' as customer where 1 = :order_no" },
      ...tools.filter((_, i) => i !== 1),
    ],
  });
  const { result, text, audit } = await run(slow, "erp_get_order", { order_no: 1 });
  assert.equal(result.isError, true);
  assert.match(text, /statement timeout/);
  assert.equal(audit[0].decision, "failed");
});

// The other engines, against fake drivers that record what they were asked.

const target: Target = {
  host: "db.internal",
  port: 1,
  database: "erp",
  user: "agents",
  password: "pw",
  tls: "verify",
  readOnly: true,
};

function input(write: boolean, over: Partial<RunInput> = {}) {
  return {
    statement: parseStatement("select a from t where x = :x and y = :y or z = :x"),
    values: { x: "1", y: null },
    write,
    maxRows: 10,
    timeoutMs: 5000,
    check: () => {},
    ...over,
  } satisfies RunInput;
}

test("SQL Server: named params, read-only intent, ROWCOUNT and rollback on reads", async () => {
  const log: string[] = [];
  let options: Record<string, Record<string, unknown>> = {};
  const answer = { recordset: [{ a: 1 }], rowsAffected: [1] };
  const mssql = {
    ConnectionPool: class {
      constructor(o: typeof options) {
        options = o;
      }
      connect = async () => log.push("connect");
    },
    Transaction: class {
      begin = async () => log.push("begin");
      commit = async () => log.push("commit");
      rollback = async () => log.push("rollback");
    },
    Request: class {
      constructor(_tx: unknown, o: { requestTimeout: number }) {
        log.push(`timeout ${o.requestTimeout}`);
      }
      input = (n: string, v: unknown) => log.push(`input ${n}=${v}`);
      query = async (q: string) => (log.push(q), answer);
    },
  };
  const engine = sqlServerEngine(target, mssql);
  assert.deepEqual(await engine.run(input(false)), { rows: [{ a: 1 }], rowsAffected: null });
  assert.deepEqual(log, [
    "connect",
    "begin",
    "timeout 5000",
    "input x=1",
    "input y=null",
    "set rowcount 11;\nselect a from t where x = @x and y = @y or z = @x",
    "rollback",
  ]);
  assert.equal(options.options.readOnlyIntent, true);
  assert.equal(options.options.encrypt, true);
  assert.equal(options.options.trustServerCertificate, false);
  log.length = 0;
  assert.deepEqual(await engine.run(input(true)), { rows: [{ a: 1 }], rowsAffected: 1 });
  assert.deepEqual(log.slice(-2), ["select a from t where x = @x and y = @y or z = @x", "commit"]);
  log.length = 0;
  const denied = input(true, {
    check: () => {
      throw new Error("outside");
    },
  });
  await assert.rejects(engine.run(denied), /outside/);
  assert.equal(log.at(-1), "rollback");
  assert.ok(!log.includes("commit"));
});

test("MySQL: prepared statements, a read-only transaction and a row cap", async () => {
  const log: unknown[] = [];
  let options: Record<string, unknown> = {};
  const conn = {
    query: async (q: string) => log.push(q),
    execute: async (o: { sql: string; timeout: number }, values: unknown[]) => {
      log.push(o.sql, o.timeout, values);
      return [o.sql.startsWith("select") ? [{ a: 1 }] : { affectedRows: 2 }];
    },
    release: () => log.push("release"),
  };
  const mysql = {
    createPool: (o: typeof options) => ((options = o), { getConnection: async () => conn }),
  };
  const engine = mySqlEngine(target, mysql);
  assert.deepEqual(await engine.run(input(false)), { rows: [{ a: 1 }], rowsAffected: null });
  assert.deepEqual(log, [
    "set session sql_select_limit = 11",
    "start transaction read only",
    "select a from t where x = ? and y = ? or z = ?",
    5000,
    ["1", null, "1"],
    "rollback",
    "release",
  ]);
  assert.deepEqual(options.ssl, { rejectUnauthorized: true });
  log.length = 0;
  const write = input(true, { statement: parseStatement("update t set a = :x") });
  assert.deepEqual(await engine.run(write), { rows: [], rowsAffected: 2 });
  assert.deepEqual(log, [
    "set session sql_select_limit = default",
    "start transaction read write",
    "update t set a = ?",
    5000,
    ["1"],
    "commit",
    "release",
  ]);
});

test("Oracle: named binds, a read-only transaction, maxRows and TLS", async () => {
  const log: unknown[] = [];
  let options: Record<string, unknown> = {};
  const conn = {
    callTimeout: 0,
    execute: async (q: string, binds?: unknown, o?: { maxRows: number; autoCommit: boolean }) => {
      log.push(q, ...(binds ? [binds, o!.maxRows, o!.autoCommit] : []));
      return { rows: [{ A: 1 }], rowsAffected: 3 };
    },
    commit: async () => log.push("commit"),
    rollback: async () => log.push("rollback"),
    close: async () => log.push("close"),
    oracleServerVersionString: "23.4.0.0",
  };
  const oracledb = {
    createPool: async (o: typeof options) => ((options = o), { getConnection: async () => conn }),
    OUT_FORMAT_OBJECT: 4002,
  };
  const engine = oracleEngine(target, oracledb);
  assert.deepEqual(await engine.run(input(false)), { rows: [{ A: 1 }], rowsAffected: null });
  assert.deepEqual(log, [
    "set transaction read only",
    "select a from t where x = :x and y = :y or z = :x",
    { x: "1", y: null },
    11,
    false,
    "rollback",
    "close",
  ]);
  assert.equal(conn.callTimeout, 5000);
  assert.equal(options.connectString, "tcps://db.internal:1/erp");
  assert.equal(options.sslServerDNMatch, true);
  log.length = 0;
  assert.deepEqual(await engine.run(input(true)), { rows: [{ A: 1 }], rowsAffected: 3 });
  assert.deepEqual(log.slice(-2), ["commit", "close"]);
  assert.equal(await engine.version(), "Oracle 23.4.0.0");
});

test("a failed first connection is tried again on the next call", async () => {
  let attempts = 0;
  const oracledb = {
    createPool: async () => {
      if (++attempts === 1) throw new Error("listener down");
      return { getConnection: async () => ({ execute: async () => ({}), close: async () => {} }) };
    },
  };
  const engine = oracleEngine({ ...target, tls: "off" }, oracledb);
  await assert.rejects(engine.version(), /listener down/);
  assert.match(await engine.version(), /Oracle/);
  assert.equal(attempts, 2);
});

// Optional smoke runs against real servers, e.g. TEST_MYSQL_URL=mysql://root:pw@localhost:3306/smoke,
// TEST_SQLSERVER_URL=sqlserver://sa:pw@localhost:1433/master (TLS without checking the
// certificate, as their containers' are self-signed) or TEST_ORACLE_URL=oracle://user:pw@localhost:1521/FREEPDB1
// (plain, and 23ai or later for "drop table if exists").
for (const [engine, env] of [
  ["mysql", "TEST_MYSQL_URL"],
  ["sqlserver", "TEST_SQLSERVER_URL"],
  ["oracle", "TEST_ORACLE_URL"],
] as const) {
  const raw = process.env[env];
  test(
    `${engine}, for real: bound values, row cap, reads changing nothing`,
    {
      skip: !raw && `${env} not set`,
    },
    async () => {
      const u = new URL(raw!);
      const t: Target = {
        host: u.hostname,
        port: Number(u.port),
        database: u.pathname.slice(1),
        user: decodeURIComponent(u.username),
        password: decodeURIComponent(u.password),
        tls: engine === "oracle" ? "off" : "insecure",
        readOnly: false,
      };
      const writer = await openEngine(engine, t);
      const reader = await openEngine(engine, { ...t, readOnly: engine === "sqlserver" });
      // Oracle names columns in upper case.
      const go = async (e: Engine, sql: string, values = {}, over: Partial<RunInput> = {}) => {
        const out = await e.run({
          statement: parseStatement(sql),
          values,
          write: e === writer,
          maxRows: 2,
          timeoutMs: 10_000,
          check: () => {},
          ...over,
        });
        const rows = out.rows.map((r) =>
          Object.fromEntries(Object.entries(r).map(([k, v]) => [k.toLowerCase(), v])),
        );
        return { ...out, rows };
      };
      try {
        assert.ok((await writer.version()).length > 0);
        await go(writer, "drop table if exists smoke_orders");
        await go(writer, "create table smoke_orders (id int primary key, customer varchar(20))");
        for (const [id, customer] of [
          [1, "ACME"],
          [2, "ACME"],
          [3, "ACME"],
          [4, "GLOBEX"],
        ])
          await go(writer, "insert into smoke_orders values (:id, :customer)", { id, customer });
        // Oracle refuses a read-only transaction a table changed shape within seconds of
        // (ORA-01466); a real table is older than that.
        if (engine === "oracle") await new Promise((done) => setTimeout(done, 5000));
        const all = await go(
          reader,
          "select id from smoke_orders where customer = :c order by id",
          {
            c: "ACME",
          },
        );
        assert.deepEqual(
          all.rows.map((r) => r.id),
          [1, 2, 3],
        );
        const capped = await go(reader, "select id from smoke_orders order by id");
        assert.equal(capped.rows.length, 3);
        const sneaky = await go(reader, "select id from smoke_orders where customer = :c", {
          c: "x' or '1'='1",
        });
        assert.deepEqual(sneaky.rows, []);
        const read = go(reader, "update smoke_orders set customer = 'LOST' where id = :id", {
          id: 1,
        });
        // SQL Server has no read-only transaction: the change is rolled back instead.
        if (engine === "sqlserver") await read;
        else await assert.rejects(read, /READ ONLY|ORA-01456/i);
        const refused = go(
          writer,
          "update smoke_orders set customer = 'LOST' where id = :id",
          { id: 2 },
          {
            check: () => {
              throw new Error("outside");
            },
          },
        );
        await assert.rejects(refused, /outside/);
        const after = await go(
          reader,
          "select customer from smoke_orders where id in (1, 2) order by id",
        );
        assert.deepEqual(
          after.rows.map((r) => r.customer),
          ["ACME", "ACME"],
        );
        assert.equal(
          (
            await go(writer, "update smoke_orders set customer = :c where id = :id", {
              c: "X",
              id: 3,
            })
          ).rowsAffected,
          1,
        );
        await go(writer, "drop table smoke_orders");
      } finally {
        await writer.close();
        await reader.close();
      }
    },
  );
}
