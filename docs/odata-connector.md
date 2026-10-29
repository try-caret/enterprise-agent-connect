# Add an OData service without code

OData is how SAP (SAP Gateway and S/4HANA, on premise and cloud), Microsoft Dynamics 365
(Dataverse), SAP SuccessFactors and many other enterprise systems publish their data. You can let
agents use an OData v2 or v4 service by describing what they may do in a single `connector.json`
with an `odata` block. The service runs it with its built-in `odata` connector, and the app gets
what every other app gets: an API key per agent, the actions you tick for each agent, value
limits, a confirmation for anything that deletes, and an audit row for every call.

Agents never write OData. Each tool takes typed values, and the service writes every `$filter`,
key and literal itself, so an agent cannot widen a filter or reach another entity set.

`examples/legacy-apps/connectors/sap-sales-orders` is a complete example (SAP's Sales Order API).
Put your folder where `CONNECTOR_DIRS` points (see [http-connector.md](http-connector.md)), or pass
the definition in `connectors`.

## Start from `$metadata`

```sh
curl -u USER -o metadata.xml 'https://s4.example.com/sap/opu/odata/sap/API_SALES_ORDER_SRV/$metadata'
npx tsx scripts/odata-to-connector.ts metadata.xml \
  --root https://s4.example.com/sap/opu/odata/sap/API_SALES_ORDER_SRV --id sap-orders --prefix sap \
  --sets A_SalesOrder > connector.json
```

It drafts a list and a get tool for each entity set (`--sets` keeps only those named), and a tool
for each function or action import. It is a draft, and notes on stderr say what it skipped. Before
use: write descriptions an agent can act on, trim `select` to what agents need, add filters and
limits, and add create and update tools by hand for the few writes agents should make.

## The `odata` block

```json
"odata": {
  "serviceRoot": "https://s4.example.com/sap/opu/odata/sap/API_SALES_ORDER_SRV",
  "version": "2",
  "auth": { "type": "basic", "username": "username", "password": "password" },
  "headers": { "sap-client": "100" },
  "csrf": true,
  "tools": [
    {
      "name": "sap_list_sales_orders",
      "description": "List sales orders, newest first.",
      "kind": "list",
      "entitySet": "A_SalesOrder",
      "filters": {
        "sales_org": { "type": "string", "property": "SalesOrganization", "description": "Sales organization", "limit": "sales_orgs" },
        "created_from": { "type": "datetime", "property": "CreationDate", "op": "ge", "description": "Created on or after" }
      },
      "select": ["SalesOrder", "SalesOrganization", "SoldToParty", "CreationDate"],
      "orderBy": "CreationDate desc",
      "top": 50
    }
  ]
}
```

- **serviceRoot**: every call goes under this address and nowhere else; redirects are not
  followed. https unless `"allowInsecureHttp": true`.
- **version**: `"2"` (SAP Gateway, SuccessFactors) or `"4"` (Dynamics, newer SAP services).
- **auth**: `header`, `bearer`, `basic` or `oauth2` (client credentials, with an optional
  `scope`), as in [http-connector.md](http-connector.md).
- **headers**: sent on every call, e.g. SAP's `sap-client`. Headers the service sets itself
  (authorization, cookies, OData's own) are refused.
- **csrf**: before each write, fetch an `x-csrf-token` (and its session cookies) and send it with
  the write. SAP Gateway refuses writes without one.
- **paging**: `skip` (default: `$top` and `$skip`, and the service's own `$skiptoken` when it gives
  one) or, for v4 services without `$skip` such as Dynamics, `server` (asks for pages of `top`
  with `Prefer: odata.maxpagesize` and follows the service's next links).

### Tools

Every tool has a `name`, a `description` and a `kind`:

| kind | needs | sends |
|------|-------|-------|
| `list` | `entitySet`, optional `filters`, `orderBy` | `GET /Set?$filter&$select&$orderby&$top` |
| `get` | `entitySet`, `key` | `GET /Set(Key='…')?$select` |
| `create` | `entitySet`, `fields` | `POST /Set` with a JSON body |
| `update` | `entitySet`, `key`, `fields` | v4 `PATCH`, v2 `MERGE`, with `If-Match: *` |
| `function` | `function`, optional `params` | `GET` of an unbound function import |
| `action` | `function`, optional `params` | `POST` of an unbound action (v2: a POST function import) |

- **key**, **filters**, **fields** and **params** map the agent's param names to values:
  `type`, a `description` the agent reads, the OData `property` (default the param's own name),
  and an optional `limit`. Types: `string`, `integer`, `number` (Edm.Double), `decimal`,
  `boolean`, `guid`, `datetime` (a date or a date and time; without an offset it is UTC) and
  `datetimeoffset`. Each value is checked against its type and written as the right literal for
  the version: strings quoted with `'` doubled, `datetime'…'`, `guid'…'` and `…M` in v2.
- **filters** (list): each has an `op`: `eq` (default), `ne`, `gt`, `ge`, `lt`, `le`, or for
  strings `contains` and `startswith`. Filters are optional unless `"required": true`, and the
  ones the agent sends are joined with `and`.
- **fields** (create, update): required on create unless `"required": false`; on update the
  agent sends only those it changes. A Dynamics lookup can be bound with a property such as
  `parentaccountid@odata.bind`.
- **select**: the `$select`, and the only properties passed back. OData's own bookkeeping
  (`__metadata`, `@odata.*`, deferred links) is always left out.
- **top**: the most items in a page (1 to 1000, default 50). Lists take `max_results` (at most
  `top`) and answer `{items, next}`; the agent passes `next` back as `cursor` for the following
  page, with the same filters. A cursor is only a position: the service rebuilds the filter itself.
- **destructive**: the tool deletes or overwrites, and must sit in a destructive capability.

## Limits

- A key, field or param with `"limit": "<list limit>"` must be one of the agent's values, and
  must be sent while the agent has any (on update, a limited field the agent sends must be one).
- A limited filter (only `eq`) that the agent leaves out becomes all of the agent's values
  (`(SalesOrganization eq '1010' or SalesOrganization eq '1710')`); any other value is refused.
  Every item of the answer is checked too, so a service that ignores the filter still cannot
  hand over another value's items. When the tool has a `select`, it must include that property.
- `answerLimits` check a field of the answer for reads (get, list, function): an answer outside
  the agent's values, or without the field, is refused. On an **update** the service first reads
  the entity (`$select` of those fields) and changes it only when it is inside the agent's
  values, so a key cannot reach another sales organization's order. Creates and actions cannot
  have them: limit their fields and params instead.

## What the service enforces

Addresses under 8 KB, bodies under 256 KB, answers under 1 MB and 30 seconds, only JSON or text
answers. Errors reach the agent in the service's own words (`The service answered 404: SY/530:
Sales order 1 does not exist`). The audit row's target is the request actually sent, with the
filter the service wrote.

## Systems

SAP S/4HANA and Dynamics 365 also have built-in apps (`sap-s4hana`, `dynamics-365`) with tools
of their own. Use an `odata` block for the services and entity sets those do not cover, or for
other OData systems.

- **SAP S/4HANA and SAP Gateway** (v2, most of SAP's published APIs): `basic` auth with a
  communication or technical user, `csrf: true`, `sap-client` in `headers` on systems with
  several clients.
- **Microsoft Dynamics 365 / Dataverse** (v4): `oauth2` with an Entra app registered as an
  application user in the environment, `tokenUrl`
  `https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/token`, `scope`
  `https://<org>.crm.dynamics.com/.default`, serviceRoot `https://<org>.crm.dynamics.com/api/data/v9.2`
  and `"paging": "server"`.
- **SAP SuccessFactors** (v2): `basic` auth as `user@companyId` against the API server of your
  data center; its change dates are `datetimeoffset`.

Tested against fake SAP Gateway (v2) and Dynamics-like (v4) services only, never against a live
SAP, Dynamics or SuccessFactors system.
