# Running it inside your network

Most installs run next to the systems they connect: an on-premises SAP Gateway, a ServiceNow or
Jira behind your own domain, a database on a private network. This page covers the two things
that usually need care there.

## Systems with certificates from your own authority

Many internal systems use a certificate signed by your company's own certificate authority
rather than a public one. The service checks every certificate, so a call to such a system fails
with a certificate error until the service trusts your authority.

Give Node your authority's certificate (PEM; several in one file are fine) with
`NODE_EXTRA_CA_CERTS`. Public authorities stay trusted. Node reads it once, at start, so it
must be set in the environment before the process starts, not in code. It covers every app the
service calls over HTTPS (built-in, `http`, `soap` and `odata` connectors) and the `sql`
connector's database drivers.

**Docker:**

```sh
docker run -v /etc/pki/corp-ca.pem:/etc/agent-connect/ca.pem:ro \
  -e NODE_EXTRA_CA_CERTS=/etc/agent-connect/ca.pem ... ghcr.io/try-caret/enterprise-agent-connect:<version>
```

**Helm:** put the bundle in a ConfigMap (or a Secret) and name it in `caBundle`. The chart
mounts it and sets `NODE_EXTRA_CA_CERTS`:

```sh
kubectl create configmap corp-ca --from-file=ca.crt=corp-ca.pem
helm upgrade --install agent-connect deploy/helm/agent-connect \
  --set caBundle.configMap=corp-ca   # or caBundle.secret=<name>; caBundle.key defaults to ca.crt
```

**Embedded in your own Node app:** set `NODE_EXTRA_CA_CERTS` for the process that runs it, the
same way.

Don't turn certificate checks off (`NODE_TLS_REJECT_UNAUTHORIZED=0`). That would let anyone on
the network path read and change the calls, including the credentials they carry.

## Behind a proxy or load balancer

The service has two kinds of callers:

- **Agents**, on `/agent-connect/mcp`, and Microsoft Teams and app webhooks, on
  `/agent-connect/teams` and `/agent-connect/events`. These are the only paths to expose.
- **Your admin pages**, on `/agent-connect/admin/*`, with `AGENT_CONNECT_ADMIN_TOKEN`.

The admin API refuses any request carrying `X-Forwarded-For`. Proxies and load balancers add
that header, so the admin API cannot be reached through them, even if a route to it is added by
mistake. This is deliberate and cannot be turned off. Your admin pages reach it by one of these
routes, which never pass through the proxy:

1. **In the same process (embedding).** When your own Node app embeds the service with
   `createAgentConnect`, hand the admin client the service's `fetch`. No request leaves the
   process, so no proxy header can be on it:

   ```ts
   import { createClient } from "@arbium/enterprise-agent-connect-client";

   const admin = createClient({
     baseUrl: "http://agent-connect",
     token: process.env.AGENT_CONNECT_ADMIN_TOKEN!,
     fetch: (input, init) => eac.fetch(new Request(input, init)),
   });
   ```

   Your admin pages call your server, and your server calls `admin`. The browser never talks to
   the admin API.

2. **Over the private network (a separate service).** Run the image, and let your admin app
   call it directly on its internal address, for example the Kubernetes Service
   `http://agent-connect:3000`. Route only the public paths through your ingress or load
   balancer; the chart's Ingress does exactly that.

3. **A private listener in your own app.** If your app is itself behind a proxy and you would
   rather keep the admin API on a network hop, serve `eac.fetch` a second time on an address
   only your admin backend reaches, and forward nothing from the proxy to it:

   ```ts
   import { serve } from "@hono/node-server";

   // Public: behind the proxy, which forwards only /agent-connect/mcp, /teams and /events.
   serve({ fetch: eac.fetch, port: 3000 });
   // Private: bound to the loopback (or an internal interface), never behind the proxy.
   serve({ fetch: eac.fetch, port: 3001, hostname: "127.0.0.1" });
   ```

Whichever you choose, keep the proxy forwarding only the three public paths. A proxy that
removes `X-Forwarded-For` before calling the service would defeat the check, so don't configure
one that way.
