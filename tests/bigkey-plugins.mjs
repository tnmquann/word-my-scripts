// BigKey plugins (src/features/bigkey/plugins.js): Chemistry Formatter rules + OOXML, Tách LG extraction, ZipGrade fill.
//   node tests/bigkey-plugins.mjs [path/to/ans50zg.docx]   (the ZipGrade template is not in the repo: BigKey plugin file)
import fs from "node:fs";
import { createRequire } from "node:module";
import { root, loadModule, check, done } from "./helpers.mjs";

const PL = await loadModule("features/bigkey/plugins.js");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const pkg = (inner) =>
  `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
  `<w:document xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>${inner}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const p = (...runs) => `<w:p>${runs.join("")}</w:p>`;
const r = (t, rpr = "") => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;

// ---- Chemistry Formatter rules: _x subscript, ^x superscript
const show = (s) => PL.chemTransform(s).map((o) => (o.va === "subscript" ? "_" : o.va === "superscript" ? "^" : "") + o.ch).join("");
const cases = [
  ["H2SO4", "H_2SO_4"], ["Fe3+", "Fe^3^+"], ["Fe2+", "Fe^2^+"], ["SO42-", "SO_4^2^-"], ["NO3-", "NO_3^-"],
  ["Cu(OH)2", "Cu(OH)_2"], ["Al2(SO4)3", "Al_2(SO_4)_3"], ["CuSO4*5H2O", "CuSO_4·5H_2O"],
  ["2H2 + O2 --> 2H2O", "2H_2 + O_2 → 2H_2O"], ["H-Cl", "H-Cl"], ["Na+ và Cl- là ion", "Na^+ và Cl^- là ion"],
  ["6.02E23 hạt", "6.02 × 10^2^3 hạt"], ["x = 2", "x = 2"], ["Câu 12.", "Câu 12."],
];
for (const [a, b] of cases) check(show(a) === b, `chem: ${a} -> ${b}`, show(a));

// ---- Chemistry Formatter OOXML: split runs with vertAlign, keep italic, skip equations, keep tabs
const chemPkg = pkg(p(r("Cho H2SO4", "<w:i/>"), `<w:r><w:tab/></w:r>`, `<m:oMath><m:r><m:t>x2</m:t></m:r></m:oMath>`, r(" và Fe3+")));
const out = PL.chemFormatXml(chemPkg);
const flat = out && [...new DOMParser().parseFromString(out.xml, "application/xml").getElementsByTagNameNS(W, "r")]
  .map((x) => { const v = x.getElementsByTagNameNS(W, "vertAlign")[0]; const t = x.getElementsByTagNameNS(W, "t")[0]; return (v ? (v.getAttributeNS(W, "val") === "subscript" ? "_" : "^") : "") + (t ? t.textContent : x.getElementsByTagNameNS(W, "tab").length ? "⇥" : ""); }).join("|");
check(flat === "Cho H|_2|SO|_4|⇥| và Fe|^3+", "chemFormatXml: runs split, tab kept", flat);
check(out && /<m:t>x2<\/m:t>/.test(out.xml) && (out.xml.match(/<w:i\/>/g) || []).length === 4, "chemFormatXml: equation untouched, italic kept on split runs");
check(PL.chemFormatXml(pkg(p(r("Không có gì")))) === null, "chemFormatXml: nothing to do -> null");

// ---- Tách LG
const solPkg = pkg(
  p(r("Câu 1. Tính?")) + p(r("A. 1\tB. 2")) + p(r("Lời giải")) + p(r("Ta có 1+1")) +
  `<w:tbl><w:tr><w:tc><w:p>${r("bảng")}</w:p></w:tc></w:tr></w:tbl>` + p(r("Chọn A")) +
  p(r("Câu 2. Hỏi?")) + p(r("Lời giải:")) + p(r("Giải câu 2")) +
  p(r("Questions 3-4: Read the following passage")) + p(r("Passage text")) +
  p(r("Câu 3. Không có lời giải")) + p(r("A. x"))
);
const sol = PL.extractSolutionsXml(solPkg);
const sd = sol && new DOMParser().parseFromString(sol.xml, "application/xml");
const kids = sd ? [...sd.getElementsByTagNameNS(W, "body")[0].childNodes].filter((n) => n.nodeType === 1) : [];
const txt = kids.map((n) => (n.localName === "tbl" ? "[tbl]" : [...n.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("")));
check(sol && sol.count === 2 && JSON.stringify(txt) === JSON.stringify(["LỜI GIẢI CHI TIẾT", "Câu 1. Ta có 1+1", "[tbl]", "Chọn A", "Câu 2. Giải câu 2", ""]),
  "Tách LG: label + solution only, table kept, passage and unsolved question skipped, sentinel last", JSON.stringify(txt));
check(sol && !/sectPr/.test(sol.xml), "Tách LG: no section break inserted");
check(PL.extractSolutionsXml(pkg(p(r("Câu 1. Hỏi?")) + p(r("A. 1")))) === null, "Tách LG: no solution -> null");

// ---- Tô phiếu ZG50 (template from the BigKey plugin folder, if given)
const tpl = process.argv[2] || process.env.MYSCRIPTS_ZG50;
if (!tpl || !fs.existsSync(tpl)) console.log("SKIP zipgrade: no ans50zg.docx given");
else {
  const JSZip = createRequire(root + "/package.json")("jszip");
  const zip = await JSZip.loadAsync(fs.readFileSync(tpl));
  const xml = await zip.file("word/document.xml").async("string");
  const res = PL.fillZipGradeXml(xml, { 1: "A", 2: "C", 50: "D" }, "B");
  const off = (x, name, axis) => {
    const a = [...new DOMParser().parseFromString(x, "application/xml").getElementsByTagName("wp:docPr")].find((d) => d.getAttribute("name") === name).parentNode;
    return +[...a.childNodes].find((n) => n.localName === "position" + axis).getElementsByTagName("wp:posOffset")[0].textContent;
  };
  check(res.filled === 3, "zipgrade: 3 bubbles filled", String(res.filled));
  check(off(res.xml, "1", "H") === off(xml, "1", "H") && off(res.xml, "2", "H") - off(xml, "2", "H") === Math.round(2 * 14.6 * 12700) && off(res.xml, "50", "H") - off(xml, "50", "H") === Math.round(3 * 14.6 * 12700),
    "zipgrade: A stays, C/D shifted by 2/3 bubble steps");
  check(off(res.xml, "Ma", "V") - off(xml, "Ma", "V") === Math.round(21.15 * 12700) && /<w:t[^>]*>B<\/w:t>/.test(res.xml), "zipgrade: mã đề B written and its bubble shifted");
  check((res.xml.match(/val="292929"/g) || []).length === 4, "zipgrade: 4 shapes filled dark (3 answers + mã đề)");
}
done();
