// Round-trip LaTeX -> OMML -> LaTeX in a real browser DOM (Edge headless --dump-dom).
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import os from "node:os";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sp = fs.mkdtempSync(path.join(os.tmpdir(), "myscripts-"));
const { build } = createRequire(root + "/package.json")("esbuild");
const stub = path.join(sp, "runtime-stub.js");
fs.writeFileSync(stub, "export const askUser=async o=>o.buttons[0];");
const out = await build({
  entryPoints: [path.join(root, "src/features/equations/latexEquations.js")], bundle: true, format: "iife", globalName: "EQ", write: false,
  plugins: [{ name: "stub", setup(b) { b.onResolve({ filter: /runtime\.js$/ }, () => ({ path: stub })); } }],
});
const samples = [
  "\\frac{a+b}{c}", "x^{2}+y_{1}", "\\sum_{i=1}^{n} i^2", "\\int_0^1 f(x)\\,dx", "\\sqrt[3]{x+1}", "\\binom{n}{k}",
  "\\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}", "f(x)=\\begin{cases} x & x>0 \\\\ -x & x\\le 0 \\end{cases}",
  "\\overset{\\frown}{AB}", "\\underbrace{a+b}_{n}", "\\mathbb{R}", "\\mathcal{L}", "\\mathbf{v}", "\\operatorname{rank}(A)",
  "\\lVert x\\rVert", "\\text{khi } x>0", "\\lim_{x\\to 0}\\frac{\\sin x}{x}", "\\alpha+\\beta\\geq\\gamma",
];
const html = `<!doctype html><meta charset="utf-8"><pre id="o"></pre><script>${out.outputFiles[0].text}
const W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const S=${JSON.stringify(samples)};let r=[];
for(const s of S){try{const om=EQ.conv(s);const x='<w:document '+W+'><w:body><w:p><w:r><w:t>a </w:t></w:r><m:oMath>'+om+'</m:oMath></w:p></w:body></w:document>';
const res=EQ.ommlToLatexXml(x);const t=[...new DOMParser().parseFromString(res.xml,'application/xml').getElementsByTagName('w:t')].map(e=>e.textContent).join('');
r.push((om.includes('<m:')?'OK  ':'??  ')+s+'  =>  '+t.replace(/^a /,''));}catch(e){r.push('ERR '+s+' : '+e.message);}}
document.getElementById('o').textContent=r.join('\\n');</script>`;
const page = path.join(sp, "eq-test.html");
fs.writeFileSync(page, html);
const dom = execFileSync("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  ["--headless=new", "--disable-gpu", "--dump-dom", "file:///" + page.replace(/\\/g, "/")], { encoding: "utf8" });
const m = dom.match(/<pre id="o">([\s\S]*?)<\/pre>/);
const res = m ? m[1].replace(/&gt;/g, ">").replace(/&lt;/g, "<").replace(/&amp;/g, "&") : dom.slice(0, 500);
console.log(res);
// every sample must convert (OK) and come back without falling back to \mathrm{<unknown command>}
const lines = res.split("\n").filter(Boolean);
const bad = lines.filter((l) => !l.startsWith("OK") || /\\mathrm\{[a-z]{3,}\}/.test(l));
console.log(lines.length === samples.length && !bad.length ? "PASS" : "FAIL " + bad.join(" | "));
if (lines.length !== samples.length || bad.length) process.exitCode = 1;
