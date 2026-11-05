import assert from "node:assert/strict";
import { test } from "node:test";
import { sanitizeHtml } from "./html.ts";

test("sanitizeHtml keeps formatting, tables, https links and layout styles", () => {
  const html =
    '<p style="color: #333; font-weight: bold">Hi <b>there</b></p>' +
    '<table border="1" cellpadding="4"><tr><td colspan="2" style="background-color: rgb(240, 240, 240); padding: 4px">A &amp; B</td></tr></table>' +
    '<a href="https://example.com/x?a=1&amp;b=2">link</a><br/><img src="https://example.com/i.png" alt="logo">';
  assert.equal(
    sanitizeHtml(html),
    '<p style="color: #333; font-weight: bold">Hi <b>there</b></p>' +
      '<table border="1" cellpadding="4"><tr><td colspan="2" style="background-color: rgb(240, 240, 240); padding: 4px">A &amp; B</td></tr></table>' +
      '<a href="https://example.com/x?a=1&amp;b=2" rel="noopener noreferrer">link</a><br><img src="https://example.com/i.png" alt="logo">',
  );
});

test("sanitizeHtml removes scripts, forms, frames, handlers and non-https addresses", () => {
  const cases: [string, string][] = [
    ["<script>alert(1)</script>ok", "ok"],
    ["<SCRIPT src=x></SCRIPT >ok", "ok"],
    ["<script>never closed", ""],
    ["<style>p{}</style><iframe src=https://x></iframe><form><input></form>ok", "ok"],
    ["<svg><script>alert(1)</script></svg>ok", "ok"],
    ['<p onclick="alert(1)" onmouseover=x>t</p>', "<p>t</p>"],
    ['<a href="javascript:alert(1)">t</a>', "<a>t</a>"],
    ['<a href="&#106;avascript:alert(1)">t</a>', "<a>t</a>"],
    ['<a href=" java\tscript:alert(1)">t</a>', "<a>t</a>"],
    ['<a href="http://example.com">t</a>', "<a>t</a>"],
    ['<a href="https:/\\evil.com">t</a>', "<a>t</a>"],
    ['<img src="http://example.com/track.png">', ""],
    ['<img src="data:image/png;base64,AAAA">x', "x"],
    ['<img src=x onerror="alert(1)">', ""],
    ['<p style="background: url(https://x/t.png); color: red">t</p>', '<p style="color: red">t</p>'],
    ['<p style="width: expression(alert(1))">t</p>', "<p>t</p>"],
    ['<p style="color: re&#100;; position: fixed">t</p>', '<p style="color: red">t</p>'],
    ["<!-- <script>x</script> -->ok", "ok"],
    ["<!DOCTYPE html><html><body>ok</body></html>", "ok"],
    ["<scr<script>ipt>alert(1)</script>", "ipt&gt;alert(1)"],
    ['<a href="https://e.com" title="x" style="">t', '<a href="https://e.com" title="x" rel="noopener noreferrer">t'],
    ["1 < 2 & 3 > 2", "1 &lt; 2 &amp; 3 &gt; 2"],
    ['<p title="a&quot;onmouseover=&quot;x">t</p>', "<p>t</p>"],
    ['<meta http-equiv="refresh" content="0;url=https://x">ok', "ok"],
    ['<base href="https://evil.com/">ok', "ok"],
    ['<link rel=stylesheet href="https://x">ok', "ok"],
  ];
  for (const [input, expected] of cases) assert.equal(sanitizeHtml(input), expected, input);
});
