// Render the "Công cụ BigKey" pane as static HTML (icons/descriptions applied, all sections open) for a screenshot:
//   node tests/tools/pane-preview.mjs <out.html>   then   msedge --headless=new --screenshot=out.png --window-size=360,1400 file:///…/out.html
import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";
import { root, loadModule } from "../helpers.mjs";

const out = process.argv[2];
const html = fs.readFileSync(path.join(root, "src/ui/bigkey/bigkey.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
const dom = new JSDOM(html);
Object.assign(globalThis, { window: dom.window, document: dom.window.document });
let ready;
globalThis.Office = { onReady: (cb) => (ready = cb) };
globalThis.Word = { run: async () => "" };
await loadModule("ui/bigkey/bigkey.js");
ready();
for (const d of dom.window.document.querySelectorAll("details")) d.open = true;
fs.writeFileSync(out, "<!doctype html>" + dom.window.document.documentElement.outerHTML);
console.log("preview ->", out);
