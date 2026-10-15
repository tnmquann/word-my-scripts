// Mock-Word test of the batched LaTeX -> equation driver (search / insertOoxml / sync count).
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sp = fs.mkdtempSync(path.join(os.tmpdir(), "myscripts-"));
const { build } = createRequire(root + "/package.json")("esbuild");
const stub = path.join(sp, "runtime-stub.js");
fs.writeFileSync(stub, "export const askUser=async o=>globalThis.__choice||o.buttons[0];");
const out = await build({
  entryPoints: [path.join(root, "src/features/equations/latexEquations.js")], bundle: true, format: "esm", write: false, platform: "node",
  plugins: [{ name: "stub", setup(b) { b.onResolve({ filter: /runtime\.js$/ }, () => ({ path: stub })); } }],
});
const mod = await import("data:text/javascript;base64," + Buffer.from(out.outputFiles[0].text).toString("base64"));

const long = "$" + "x+".repeat(140) + "y$"; // > 250 chars
const texts = [
  "Cho $x^2$ và $x^2$ cùng $\\frac{1}{2}$.",   // duplicate formula in one paragraph
  "Giá $5 and $6 không phải công thức.",        // currency must be ignored
  "Công thức dài: " + long,
  "Display: $$\\sum_{i=1}^n i$$ hết.",
];
let syncs = 0;
const inserted = [];
const paras = texts.map((text, pi) => ({
  text,
  search: (q, o) => {
    const needle = q.replace(/\^\^/g, "^");
    const items = []; let k = 0, j;
    while ((j = text.indexOf(needle, k)) !== -1) { items.push({ pi, start: j, end: j + needle.length, getRange(w) { return { pi, start: w === "Start" ? this.start : this.end, end: w === "Start" ? this.start : this.end, expandTo: (b) => mkRange(pi, this.start, b.end) }; } }); k = j + needle.length; }
    return { items: items.map((it) => mkRange(pi, it.start, it.end)), load() {} };
  },
}));
function mkRange(pi, start, end) {
  return { pi, start, end,
    getRange(w) { return mkRange(pi, w === "End" ? end : start, w === "End" ? end : start); },
    expandTo(b) { return mkRange(pi, start, b.end); },
    insertOoxml(x) { inserted.push({ pi, text: texts[pi].slice(start, end), display: x.includes("oMathPara") }); },
    compareLocationWith() { return { value: "Inside" }; } };
}
const ctx = { sync: async () => { syncs++; }, document: {
  getSelection: () => ({ isEmpty: true, load() {} }),
  body: { paragraphs: { items: paras, load() {} } } } };

globalThis.__choice = "Toàn bộ tài liệu";
const res = await mod.latexToEquation(ctx);
console.log("summary:", JSON.stringify(res.summary));
console.log("syncs:", syncs);
for (const i of inserted) console.log(`  p${i.pi} ${i.display ? "[display]" : "[inline] "} ${i.text.length > 60 ? i.text.slice(0, 30) + "…(" + i.text.length + " chars)" : i.text}`);
const ok = inserted.length === 5 && !inserted.some((i) => i.text.includes("$5")) && inserted.filter((i) => i.text === "$x^2$").length === 2
  && new Set(inserted.filter((i) => i.text === "$x^2$").map((i) => i.text + i.pi)).size === 1;
console.log(ok && syncs <= 4 ? "PASS" : "FAIL");
if (!ok || syncs > 4) process.exitCode = 1;
globalThis.__choice = "Huỷ";
console.log("cancel ->", JSON.stringify(await mod.latexToEquation(ctx)));
globalThis.__choice = "Vùng chọn";
console.log("selection while empty ->", JSON.stringify((await mod.latexToEquation(ctx)).summary));
