# Add a SOAP service without code

Many core business systems answer only SOAP: SAP (its SOAP web services and enterprise
services), Oracle E-Business Suite, Workday's web services, older Microsoft Dynamics, and most
internal services built on .NET WCF or Java JAX-WS. You can let agents use one by describing its
operations in a single `connector.json` with a `soap` block. The service runs it with its built-in
`soap` connector, and the app gets what every other app gets: an API key per agent, the actions
you tick for each agent, value limits, a confirmation for anything that deletes, and an audit row
for every call.

`examples/legacy-apps/connectors/hr-soap` is a complete example. Put your folder where
`CONNECTOR_DIRS` points (see [http-connector.md](http-connector.md)), or pass the definition in
`connectors`.

## Start from the WSDL

```sh
npx tsx scripts/wsdl-to-connector.ts https://hr.internal.example.com/EmployeeService?wsdl --id hr --prefix hr > connector.json
```

It drafts one tool per operation of the first SOAP port (document/literal, the usual style), with
the endpoint, the version, each operation's SOAPAction and a body template of its fields. It is a
draft, and notes on stderr say what it could not do (nested fields, a verify call). Before use:
write descriptions an agent can act on, keep only the operations agents need, put them in
capabilities (lookups `readOnly`, deletes `destructive`), and add limits. A WSDL behind a sign-in:
download it first (`curl -u user -o service.wsdl <url>`) and pass the file.

## The `soap` block

```json
"soap": {
  "endpoint": "https://hr.internal.example.com/services/EmployeeService",
  "version": "1.1",
  "namespaces": { "hr": "urn:example:hr:v1" },
  "auth": { "type": "basic", "username": "username", "password": "password" },
  "verify": { "tool": "hr_search_employees", "args": { "department": "IT" }, "show": "name" },
  "tools": [
    {
      "name": "hr_get_employee",
      "description": "Read one employee by employee id.",
      "operation": "GetEmployee",
      "soapAction": "urn:example:hr:v1/GetEmployee",
      "body": "<hr:GetEmployee><hr:employeeId>{employee_id}</hr:employeeId></hr:GetEmployee>",
      "params": { "employee_id": { "type": "string", "description": "The employee id" } },
      "result": "GetEmployeeResponse.employee",
      "fields": ["id", "name", "department", "title"],
      "readOnly": true,
      "answerLimits": [{ "field": "department", "limit": "departments" }]
    }
  ]
}
```

- **endpoint**: every call is a POST to this address, and nowhere else; redirects are not
  followed. It must be `https://`; for a service reachable only inside your network over plain
  http, add `"allowInsecureHttp": true`.
- **version**: `1.1` (the default: `text/xml` and a `SOAPAction` header) or `1.2`
  (`application/soap+xml` with the action in the content type).
- **namespaces**: prefixes declared on the envelope, for use in the bodies. `soap`, `wsse` and
  `wsu` are taken.
- **auth** (optional): `header`, `bearer` or `basic`, as in [http-connector.md](http-connector.md).
- **wsSecurity** (optional, with or without auth): a WS-Security UsernameToken in the SOAP header,
  `{"username": "<plain field>", "password": "<secret field>"}`. The password goes as text
  unless `"digest": true` (PasswordDigest: a hash of a fresh nonce, the time and the password).
  Each message carries a new nonce and time. Use text only over https.
- **verify**: the call made when the app is connected, to prove the credentials: a tool, its
  arguments and the answer field to show. The connector's own `verify` is then
  `{"label": "The HR directory answers", "check": "operation"}`.
- **tools**: one per operation.
  - `body` is the content of the SOAP Body, with a `{param}` placeholder where each value goes.
    Every value is XML-escaped (`<`, `&`, quotes) before it is put in, so an agent can never add
    elements or attributes; values with characters XML cannot carry are refused. A body must be
    well-formed XML, or the service refuses to start.
  - `params`: `type` (`string`, `number`, `integer`, `boolean`) and a `description` the agent
    reads. Params are required unless `"required": false`; an optional param must be the whole
    content of one element (`<hr:namePrefix>{name_prefix}</hr:namePrefix>`), and when the agent
    leaves it out, that element is left out of the body.
  - `result`: where the answer is in the Body, as a dot path (default `<operation>Response`).
    End it with `[]` for a list, so one item still comes back as a list and none as `[]`.
  - `fields`: only these fields of the answer (of each item, for a list) reach the agent. Use it:
    SOAP answers are often large, and a smaller answer is cheaper and leaks less.
  - `readOnly`: the operation changes nothing. `destructive`: it deletes or overwrites, and must
    sit in a destructive capability.

Answers come back as plain JSON: element names without their prefixes, repeated elements as a
list, text trimmed, `xsi:nil` as null, unprefixed attributes as fields.

## Limits

As for http apps: a param with `"limit": "<list limit>"` must be one of the agent's values, and
must be sent while the agent has any. `answerLimits` (on `readOnly` tools only) check a field of
the answer instead, for a lookup by id: an answer outside the agent's values, or without the
field, is refused and nothing of it reaches the agent.

A write cannot be checked by its answer, since it has already happened. To keep writes to some
records, make the operation take the value the limit is about (the example's `department` on
`hr_update_work_phone`) and have the service refuse a mismatch, or offer the write through OData,
where an update reads the record first.

## What the service enforces

Bodies under 256 KB, answers under 1 MB and 30 seconds, only XML, JSON or text answers. Answers
are read by a small built-in XML reader that refuses any document type declaration, so external
or expanding entities (XXE, "billion laughs") never run, and refuses unknown entities and
documents nested over 100 levels. A SOAP Fault, of either version and whatever its HTTP status,
reaches the agent as `The service refused the call: <reason> (<code>)`.

Tested against a fake SOAP service only. Live services differ in what they want in the WS-Security
header (some need a Timestamp, which is not sent) and in the SOAPAction they expect: check one call
from the admin page before giving agents the key.
