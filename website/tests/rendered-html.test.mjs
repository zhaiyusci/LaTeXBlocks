import assert from "node:assert/strict";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("https://latex-blocks.example/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the LaTeX Blocks product homepage", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<title>LaTeX Blocks — Editable LaTeX for Microsoft Office<\/title>/i);
  assert.match(html, /LaTeX that belongs/);
  assert.match(html, /Right at home in Word/);
  assert.match(html, /Make LaTeX part of your Office workflow/);
  assert.doesNotMatch(html, /productStage|modelCard|sourceDiagram|stepVisual/);
  assert.doesNotMatch(html, /UI screenshot|Office frame[^<]*diagram/i);
  assert.doesNotMatch(html, /og\.png|twitter:image|og:image/);
  assert.doesNotMatch(html, /codex-preview|SkeletonPreview|Your site is taking shape/);
});
