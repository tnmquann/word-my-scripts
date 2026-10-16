// Task pane "Xoá cần thiết" (src/ui/xoa): loads in jsdom with a fake Office/Word and wires every button.
import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";
import { root, loadModule, check, done } from "./helpers.mjs";

const html = fs.readFileSync(path.join(root, "src/ui/xoa/xoa.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
const dom = new JSDOM(html);
const { document } = dom.window;
Object.assign(globalThis, { window: dom.window, document });
let ready;
globalThis.Office = { onReady: (cb) => (ready = cb) };
const runs = [];
globalThis.Word = { run: async (fn) => (runs.push(fn), "Đã xong") };
globalThis.__desktop = false; // Word on the web: frame buttons disabled
console.error = () => {}; // the pane logs the expected validation errors

await loadModule("ui/xoa/xoa.js");
ready();
const $ = (s) => document.querySelector(s);
const tick = () => new Promise((r) => setTimeout(r, 0));

check($("#style").options.length === 8 && $("#style").value === "Double", "8 border styles, default Đôi", $("#style").value);
check(!$("#webWarn").hidden && [...document.querySelectorAll(".desk")].every((b) => b.disabled), "web: warning shown, frame buttons disabled");

$('button[data-act="mucDo"]').click();
await tick();
check(runs.length === 0 && /Chưa chọn mức/.test($("#status").textContent), "levels: nothing checked -> message, no Word.run", $("#status").textContent);

document.querySelector('input[name="lv"][value="3"]').checked = true;
$('button[data-act="mucDo"]').click();
await tick();
check(runs.length === 1 && /Đã xong/.test($("#status").textContent), "levels: runs once and shows the result", $("#status").textContent);

$("#nDots").value = "0";
$('button[data-act="dots"]').click();
await tick();
check(/1 đến 30/.test($("#status").textContent), "dotted lines: out-of-range n rejected", $("#status").textContent);

for (const act of ["loiGiai", "suyLuan", "id"]) $(`button[data-act="${act}"]`).click(), await tick();
check(runs.length === 4, "every delete button calls Word.run", String(runs.length));
check([...document.querySelectorAll(".desk")].every((b) => b.disabled), "web: frame buttons stay disabled after a run");
done();
