// Fill the ZipGrade 50 template with sample answers and write a .docx to inspect visually (render with render-pdf.ps1).
//   node tests/tools/zipgrade-render.mjs <ans50zg.docx> <out.docx>
import fs from "node:fs";
import { createRequire } from "node:module";
import { root, loadModule } from "../helpers.mjs";

const [, , tpl, out] = process.argv;
const PL = await loadModule("features/bigkey/plugins.js");
const JSZip = createRequire(root + "/package.json")("jszip");
const zip = await JSZip.loadAsync(fs.readFileSync(tpl));
const answers = { 1: "A", 2: "C", 3: "B", 4: "D", 25: "B", 26: "C", 50: "D" };
const res = PL.fillZipGradeXml(await zip.file("word/document.xml").async("string"), answers, "B");
zip.file("word/document.xml", res.xml);
fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer" }));
console.log(`filled ${res.filled}: ${JSON.stringify(answers)}, mã đề B -> ${out}`);
