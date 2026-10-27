// A stand-in for a company's internal API: orders, behind an API key. No dependencies.
// node server.mjs (PORT, default 8080; API_KEY, default "example-key")
import { createServer } from "node:http";

const KEY = process.env.API_KEY || "example-key";
const orders = new Map([
  ["1001", { id: "1001", customer: "acme", status: "shipped", notes: [] }],
  ["1002", { id: "1002", customer: "globex", status: "open", notes: [] }],
]);

const send = (res, status, data) => {
  res.writeHead(status, { "content-type": "application/json" });
  res.end(JSON.stringify(data));
};

createServer(async (req, res) => {
  if (req.headers["x-api-key"] !== KEY) return send(res, 401, { error: "bad API key" });
  const url = new URL(req.url, "http://x");
  if (url.pathname === "/api/health") return send(res, 200, { name: "Orders", ok: true });
  if (req.method === "GET" && url.pathname === "/api/orders") {
    const customer = url.searchParams.get("customer");
    return send(
      res,
      200,
      [...orders.values()].filter((o) => !customer || o.customer === customer),
    );
  }
  const match = url.pathname.match(/^\/api\/orders\/([^/]+)(\/notes)?$/);
  const order = match && orders.get(decodeURIComponent(match[1]));
  if (!order) return send(res, 404, { error: "no such order" });
  if (req.method === "GET" && !match[2]) return send(res, 200, order);
  if (req.method === "POST" && match[2]) {
    let raw = "";
    for await (const chunk of req) raw += chunk;
    order.notes.push(JSON.parse(raw).text);
    return send(res, 201, order);
  }
  if (req.method === "DELETE" && !match[2]) {
    orders.delete(order.id);
    return send(res, 200, { deleted: order.id });
  }
  send(res, 405, { error: "not allowed" });
}).listen(Number(process.env.PORT) || 8080);
