# Add a database without code

Many core systems are best reached through their database: an old ERP, an order system with no
API, a reporting replica. You can let agents use one by describing it in a single
`connector.json`. Each action an agent gets is a query your developer wrote and named, such as
"list a customer's orders". Agents fill in its values; they never write SQL. Your database gets
what every other app gets: an API key per agent, the actions you tick for each agent, value
limits, a confirmation for anything that deletes, and an audit row for every call.

Supported: PostgreSQL, Microsoft SQL Server, MySQL and MariaDB, and Oracle.

## Try the example

`examples/legacy-orders` is a small orders database (`schema.sql`) and its connector. From
`deploy/docker`:

```sh
docker compose -f docker-compose.yml -f docker-compose.sql-example.yml up --build --wait
```

"Orders database" now shows up with the other apps. Connect it with the user `agents` and the
password `agents-example`, give an agent a key with "Read orders" and the customer `ACME`, and
ask it to list ACME's orders. Ask it for order 1003 (another customer's) and it is refused.

## Write your own

Make a folder named after your app's id, put `connector.json` in it, and point the service at
the folder that holds it with `CONNECTOR_DIRS` (several folders: separate them with `:`).

A connector with an `sql` block is like any other connector (name, setup steps, connect form
`fields`, `capabilities`, `limits`), plus:

```json
"verify": { "label": "Database", "check": "database" },
"sql": {
  "engine": "sqlserver",
  "host": "erp-db.corp.example.com",
  "database": "ERP",
  "auth": { "user": "db_user", "password": "db_password" },
  "tools": [
    {
      "name": "erp_list_orders",
      "description": "List a customer's orders created since a date (YYYY-MM-DD).",
      "sql": "select id, status, total from orders where customer = :customer and created >= :since",
      "params": {
        "customer": { "type": "string", "description": "The customer code", "limit": "customers" },
        "since": { "type": "string", "description": "The earliest creation date, YYYY-MM-DD" }
      }
    }
  ]
}
```

- **engine**: `postgres`, `sqlserver`, `mysql` (MySQL and MariaDB) or `oracle`.
- **host**, **port** (the engine's usual one when left out) and **database** (for Oracle, the
  service name): every call goes to this database and nowhere else.
- **auth**: the ids of two connect form fields, a plain one for the user and one with
  `"secret": true` for the password. The password is stored encrypted and never shown again.
- **tls**: `verify` (the default: encrypted, and the server's certificate must be trusted),
  `insecure` (encrypted, any certificate; not offered for Oracle) or `off` (plain, for a
  database on the same private network). For a certificate from your company's own authority,
  give the service its certificate with `NODE_EXTRA_CA_CERTS`.
- **timeoutSeconds** (default 30) and **maxRows** (default 200, at most 1000): a statement
  running longer is stopped, and an answer with more rows is cut off; the agent is told so
  (`truncated: true`) and can narrow its request.
- **verify**: always `{"label": "...", "check": "database"}`. Connecting runs a trivial query
  and shows the server's version.
- **tools**: one per action, each a single statement.
  - `sql` takes `:name` placeholders. Each is bound by the database driver, never pasted into
    the text, so a value can never change the statement. A name may appear more than once.
  - `params` declares exactly the names the statement uses, each with a `type` (`string`,
    `number`, `integer`, `boolean`) and a `description` the agent reads. Params are required
    unless `"required": false`; one left out is bound as null. On PostgreSQL, a param that only
    appears as `:name is null` needs a type: `cast(:status as text) is null`.
  - `write`: a tool that changes data must say `"write": true`. Any other tool runs in a
    read-only transaction that is always rolled back, so even a statement that tries to change
    something cannot.
  - `destructive`: a write that deletes or overwrites must also say `"destructive": true`.

## Capabilities and limits

- **capabilities** group the tools into what you tick per agent. A write tool cannot be in a
  `required` capability, and a destructive one must sit in a capability marked
  `"destructive": true`.
- **limits**: give a param `"limit": "<list limit id>"` and an agent with values in that limit
  can only use those values, and must always send that param.
- **answerLimits**: for a statement that cannot be limited by its params, such as an order by
  number, check the rows instead: `"answerLimits": [{"column": "customer", "limit": "customers"}]`.
  While the agent has values in `customers`, every row's `customer` column must be one of them,
  or the agent gets a refusal and none of the rows. A row without the column is refused too.
  Writes can have them as well: the rows a write returns (`returning` on PostgreSQL and
  MariaDB, `output inserted.*` on SQL Server) are checked before it is committed, and a change
  outside the agent's values is rolled back. A write that changes rows without returning them
  cannot be checked, so under a limit it is refused. Column names are as the database returns
  them (Oracle: upper case unless the alias is quoted).

## The database user

Connect with a user made for agents and give it only what the tools need: `select` on the
tables they read, and `insert` or `update` on exactly the tables and columns the write tools
change (see `examples/legacy-orders/schema.sql`). The read-only transaction stops a read tool
from changing anything on PostgreSQL, MySQL and Oracle. SQL Server has no read-only
transaction: reads are rolled back and open their connection with `ApplicationIntent=ReadOnly`
(which sends them to a read-only replica where there is one), so the user's grants are the real
guard there.

## Drivers

PostgreSQL works out of the box. The other engines' drivers are not installed with the npm
package, so installing it stays small; the container image carries all of them. Embedding the
package yourself, install the one you need beside it: `npm install mssql` (SQL Server),
`npm install mysql2` (MySQL and MariaDB) or `npm install oracledb` (Oracle, in its thin mode: no
Oracle client libraries). A connector naming an engine whose driver is missing stops the service
at startup with the command to run.

Each connection keeps a pool of at most 3 database connections, closed after a minute idle.
