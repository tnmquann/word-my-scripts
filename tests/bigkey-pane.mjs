// Task pane "Công cụ BigKey" (src/ui/bigkey): loads in jsdom with a fake Office/Word; every button runs one job.
import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";
import { root, loadModule, check, done } from "./helpers.mjs";

const html = fs.readFileSync(path.join(root, "src/ui/bigkey/bigkey.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
const dom = new JSDOM(html);
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
let ready;
globalThis.Office = { onReady: (cb) => (ready = cb) };
const runs = [];
globalThis.Word = { run: async (fn) => (runs.push(fn), "Đã xong") };
globalThis.__desktop = false; // Word on the web
await loadModule("ui/bigkey/bigkey.js");
ready();
const $$ = (s) => [...document.querySelectorAll(s)];
const tick = () => new Promise((r) => setTimeout(r, 0));

check($$("button[data-tag]").length === 19, "ID: NB/TH/VD/VDC + 15 group buttons", String($$("button[data-tag]").length));
check($$(".desk").every((b) => b.disabled), "web: Picture Inline disabled");
let clicked = 0;
for (const b of $$("button[data-act], button[data-tag]")) {
  if (b.disabled) continue;
  b.click();
  clicked++;
  await tick();
}
check(runs.length === clicked && clicked >= 38, `every enabled button runs one Word.run job (${clicked})`, String(runs.length));
check(document.getElementById("status").textContent === "Đã xong" && document.getElementById("status").className === "ok", "status shows the result");
done();
