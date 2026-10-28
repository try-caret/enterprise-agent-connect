import { createRequire } from "node:module";
import postgres from "postgres";
import type { SqlApp } from "../../packages/connector-kit/src/connector.js";

// One adapter per database engine. Each runs one statement in its own transaction, read-only
// and rolled back unless it is a write, with the connection's own small pool.

/** A statement's text around its :name params: parts.length is names.length + 1. */
export type Statement = { parts: string[]; names: string[] };

export type Rows = Record<string, unknown>[];

export type RunInput = {
  statement: Statement;
  values: Record<string, unknown>;
  write: boolean;
  // Rows wanted at most; adapters fetch one more so a cut-off answer shows.
  maxRows: number;
  timeoutMs: number;
  // Called before commit; a throw rolls the write back.
  check(out: { rows: Rows; rowsAffected: number | null }): void;
};

export type Engine = {
  run(input: RunInput): Promise<{ rows: Rows; rowsAffected: number | null }>;
  /** The server's version, from a trivial query proving the connection. */
  version(): Promise<string>;
  close(): Promise<void>;
};

export type Target = {
  host: string;
  port: number;
  database: string;
  user: string;
  password: string;
  tls: SqlApp["tls"];
  // SQL Server: route to a read-only replica where there is one (ApplicationIntent=ReadOnly).
  readOnly: boolean;
};

export const DEFAULT_PORT = { postgres: 5432, sqlserver: 1433, mysql: 3306, oracle: 1521 };
// The module imported for each engine; its package is the part before any "/".
export const DRIVER = { sqlserver: "mssql", mysql: "mysql2/promise", oracle: "oracledb" } as const;
const POOL_MAX = 3;
const IDLE_SECONDS = 60;
const CONNECT_SECONDS = 10;

/** Splits on :name, skipping quoted text, comments and casts (::type). */
export function parseStatement(sql: string): Statement {
  const parts: string[] = [];
  const names: string[] = [];
  let start = 0;
  let i = 0;
  while (i < sql.length) {
    const c = sql[i];
    const close =
      c === "'" || c === '"' || c === "`"
        ? c
        : sql.startsWith("--", i)
          ? "\n"
          : sql.startsWith("/*", i)
            ? "*/"
            : null;
    if (close) {
      const end = sql.indexOf(close, i + (close.length === 2 ? 2 : 1));
      i = end < 0 ? sql.length : end + close.length;
      continue;
    }
    const name = c === ":" && sql[i - 1] !== ":" && /^:([a-z][a-z0-9_]*)/.exec(sql.slice(i));
    if (name) {
      parts.push(sql.slice(start, i));
      names.push(name[1]);
      i += name[0].length;
      start = i;
      continue;
    }
    i++;
  }
  parts.push(sql.slice(start));
  return { parts, names };
}

const render = (s: Statement, placeholder: (name: string, index: number) => string) =>
  s.parts.reduce((out, part, i) => out + placeholder(s.names[i - 1], i - 1) + part);

/** Throws when the engine's driver is not installed, naming what to install. */
export function assertDriver(
  engine: SqlApp["engine"],
  app: string,
  resolve: (name: string) => string = createRequire(import.meta.url).resolve,
) {
  if (engine === "postgres") return;
  try {
    resolve(DRIVER[engine]);
  } catch {
    throw new Error(
      `${app} reaches ${engine}, whose driver is not installed: npm install ${DRIVER[engine].split("/")[0]}`,
    );
  }
}

// The drivers but postgres are optional peers, loaded on first use.
type Driver = any;
const load = (name: string): Promise<Driver> => import(name).then((m) => m.default ?? m);

export async function openEngine(
  engine: SqlApp["engine"],
  target: Target,
  driver?: Driver,
): Promise<Engine> {
  if (engine === "postgres") return pgEngine(target, driver ?? postgres);
  const loaded = driver ?? (await load(DRIVER[engine]));
  if (engine === "sqlserver") return sqlServerEngine(target, loaded);
  if (engine === "mysql") return mySqlEngine(target, loaded);
  return oracleEngine(target, loaded);
}

// Rolls a read back, whatever it did.
class Rollback extends Error {}

export function pgEngine(t: Target, pg: Driver): Engine {
  const sql = pg({
    host: t.host,
    port: t.port,
    database: t.database,
    username: t.user,
    password: t.password,
    ssl: t.tls === "off" ? false : { rejectUnauthorized: t.tls === "verify", servername: t.host },
    max: POOL_MAX,
    idle_timeout: IDLE_SECONDS,
    connect_timeout: CONNECT_SECONDS,
    onnotice: () => {},
  });
  return {
    async run({ statement, values, write, maxRows, timeoutMs, check }) {
      const order = [...new Set(statement.names)];
      const text = render(statement, (name) => `$${order.indexOf(name) + 1}`);
      const params = order.map((n) => values[n]);
      let out!: { rows: Rows; rowsAffected: number | null };
      await sql
        .begin(write ? "read write" : "read only", async (tx: Driver) => {
          await tx.unsafe(`set local statement_timeout = ${Math.floor(timeoutMs)}`);
          if (write) {
            const result = await tx.unsafe(text, params);
            out = { rows: [...result], rowsAffected: result.count ?? null };
          } else {
            let rows: Rows = [];
            // A cursor stops the server after maxRows + 1 rows.
            for await (const batch of tx.unsafe(text, params).cursor(maxRows + 1)) {
              rows = batch;
              break;
            }
            out = { rows, rowsAffected: null };
          }
          check(out);
          if (!write) throw new Rollback();
        })
        .catch((error: unknown) => {
          if (!(error instanceof Rollback)) throw error;
        });
      return out;
    },
    async version() {
      const [row] = await sql`select version() as v`;
      return String(row.v);
    },
    close: () => sql.end(),
  };
}

export function sqlServerEngine(t: Target, mssql: Driver): Engine {
  const pool = new mssql.ConnectionPool({
    server: t.host,
    port: t.port,
    database: t.database,
    user: t.user,
    password: t.password,
    pool: { max: POOL_MAX, min: 0, idleTimeoutMillis: IDLE_SECONDS * 1000 },
    connectionTimeout: CONNECT_SECONDS * 1000,
    options: {
      encrypt: t.tls !== "off",
      trustServerCertificate: t.tls === "insecure",
      readOnlyIntent: t.readOnly,
    },
  });
  // Connected on first use, and again on the next call after a failed attempt.
  let connecting: Promise<unknown> | null = null;
  const connected = () =>
    (connecting ??= pool.connect().catch((error: unknown) => {
      connecting = null;
      throw error;
    }));
  return {
    async run({ statement, values, write, maxRows, timeoutMs, check }) {
      await connected();
      const tx = new mssql.Transaction(pool);
      await tx.begin();
      try {
        const request = new mssql.Request(tx, { requestTimeout: timeoutMs });
        for (const name of new Set(statement.names)) request.input(name, values[name]);
        // ROWCOUNT stops a read after maxRows + 1 rows; a write is never cut short.
        const text = render(statement, (name) => `@${name}`);
        const result = await request.query(write ? text : `set rowcount ${maxRows + 1};\n${text}`);
        const out = {
          rows: (result.recordset ?? []) as Rows,
          rowsAffected: write ? (result.rowsAffected as number[]).reduce((a, b) => a + b, 0) : null,
        };
        check(out);
        await (write ? tx.commit() : tx.rollback());
        return out;
      } catch (error) {
        // The server may already have rolled it back (e.g. after a timeout).
        await tx.rollback().catch(() => {});
        throw error;
      }
    },
    async version() {
      await connected();
      const result = await pool.request().query("select @@version as v");
      return String(result.recordset[0].v).split("\n")[0];
    },
    close: () => pool.close(),
  };
}

export function mySqlEngine(t: Target, mysql: Driver): Engine {
  const pool = mysql.createPool({
    host: t.host,
    port: t.port,
    database: t.database,
    user: t.user,
    password: t.password,
    ssl: t.tls === "off" ? undefined : { rejectUnauthorized: t.tls === "verify" },
    connectionLimit: POOL_MAX,
    maxIdle: 0,
    idleTimeout: IDLE_SECONDS * 1000,
    connectTimeout: CONNECT_SECONDS * 1000,
    supportBigNumbers: true,
    bigNumberStrings: true,
  });
  return {
    async run({ statement, values, write, maxRows, timeoutMs, check }) {
      const conn = await pool.getConnection();
      try {
        // Session settings outlive the call on a pooled connection, so each call sets its own.
        await conn.query(`set session sql_select_limit = ${write ? "default" : maxRows + 1}`);
        await conn.query(`start transaction ${write ? "read write" : "read only"}`);
        try {
          // execute: a server-side prepared statement, the values never in the text.
          const [result] = await conn.execute(
            { sql: render(statement, () => "?"), timeout: timeoutMs },
            statement.names.map((n) => values[n] ?? null),
          );
          const out = Array.isArray(result)
            ? { rows: result as Rows, rowsAffected: null }
            : { rows: [], rowsAffected: write ? (result.affectedRows as number) : null };
          check(out);
          await conn.query(write ? "commit" : "rollback");
          return out;
        } catch (error) {
          await conn.query("rollback").catch(() => {});
          throw error;
        }
      } finally {
        conn.release();
      }
    },
    async version() {
      const [rows] = await pool.query("select version() as v");
      return String(rows[0].v);
    },
    close: () => pool.end(),
  };
}

export function oracleEngine(t: Target, oracledb: Driver): Engine {
  const options = {
    user: t.user,
    password: t.password,
    connectString: `${t.tls === "verify" ? "tcps://" : ""}${t.host}:${t.port}/${t.database}`,
    sslServerDNMatch: t.tls === "verify",
    poolMin: 0,
    poolMax: POOL_MAX,
    poolIncrement: 1,
    poolTimeout: IDLE_SECONDS,
    connectTimeout: CONNECT_SECONDS,
  };
  let pool: Promise<Driver> | null = null;
  // Large text and binary columns as values, not LOB handles.
  const fetchTypeHandler = (meta: { dbType: unknown }) =>
    meta.dbType === oracledb.DB_TYPE_CLOB || meta.dbType === oracledb.DB_TYPE_NCLOB
      ? { type: oracledb.STRING }
      : meta.dbType === oracledb.DB_TYPE_BLOB
        ? { type: oracledb.BUFFER }
        : undefined;
  const connection = async () => {
    pool ??= oracledb.createPool(options).catch((error: unknown) => {
      pool = null;
      throw error;
    });
    return (await pool!).getConnection();
  };
  return {
    async run({ statement, values, write, maxRows, timeoutMs, check }) {
      const conn = await connection();
      try {
        conn.callTimeout = timeoutMs;
        if (!write) await conn.execute("set transaction read only");
        const binds = Object.fromEntries([...new Set(statement.names)].map((n) => [n, values[n]]));
        const result = await conn.execute(
          render(statement, (name) => `:${name}`),
          binds,
          {
            outFormat: oracledb.OUT_FORMAT_OBJECT,
            maxRows: write ? 0 : maxRows + 1,
            autoCommit: false,
            fetchTypeHandler,
          },
        );
        const out = {
          rows: (result.rows ?? []) as Rows,
          rowsAffected: write ? (result.rowsAffected ?? 0) : null,
        };
        check(out);
        await (write ? conn.commit() : conn.rollback());
        return out;
      } catch (error) {
        await conn.rollback().catch(() => {});
        throw error;
      } finally {
        await conn.close();
      }
    },
    async version() {
      const conn = await connection();
      try {
        await conn.execute("select 1 from dual");
        return `Oracle ${conn.oracleServerVersionString}`;
      } finally {
        await conn.close();
      }
    },
    close: async () => (await pool)?.close(0),
  };
}
