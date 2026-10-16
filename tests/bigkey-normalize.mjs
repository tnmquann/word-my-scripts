// Chuẩn hoá BigKey paragraph layout (src/features/bigkey/normalizeXml.js): matches the old VBA bkBonDong
// (Normal, no direct paragraph formatting, indent 1 cm, single spacing, no tab stops / keep with next,
// Times New Roman 12) while keeping bold/italic/underline/colour, equations and section breaks.
//   node tests/bigkey-normalize.mjs [export.json …]
import fs from "node:fs";
import { loadModule, check, done } from "./helpers.mjs";

const NX = await loadModule("features/bigkey/normalizeXml.js");
const OL = await loadModule("features/bigkey/optionLayout.js");
const RB = await loadModule("features/dgnl/removeBreakEng.js");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const pkg = (bodyInner) =>
  `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
  `<w:document xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>${bodyInner}</w:body></w:document>` +
  `</pkg:xmlData></pkg:part></pkg:package>`;

const src = pkg(
  `<w:p><w:pPr><w:pStyle w:val="Title"/><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="3000"/></w:tabs><w:spacing w:after="120" w:line="264" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="Palatino Linotype"/><w:b/><w:sz w:val="26"/></w:rPr><w:t>Đề ôn tập</w:t></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t>Câu 3.</w:t></w:r><w:r><w:tab/><w:t>Cho</w:t></w:r>` +
    `<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>x</m:t></m:r></m:oMath>` +
    `<w:r><w:br w:type="page"/><w:t xml:space="preserve"> hết</w:t><w:tab/><w:t>ok</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:pPr><w:r><w:t>sect</w:t></w:r></w:p>`
);
const doc = RB.parseXml(NX.normalizeParagraphsXml(src));
const ps = RB.bodyParagraphs(RB.documentBody(doc));
const xml = (n) => new XMLSerializer().serializeToString(n);
const pPr0 = xml(ps[0].getElementsByTagNameNS(W, "pPr")[0]);
check(!/pStyle|keepNext|w:tabs|w:jc|w:line="264"/.test(pPr0) && /w:line="240"/.test(pPr0) && /w:left="567"/.test(pPr0) && /w:firstLine="0"/.test(pPr0),
  "pPr reset: no style/keepNext/tabs/jc, single spacing, left 1 cm", pPr0);
const r0 = xml(ps[0].getElementsByTagNameNS(W, "r")[0]);
check(/Times New Roman/.test(r0) && /w:sz w:val="24"/.test(r0) && /<w:b\/>/.test(r0) && !/Palatino|w:val="26"/.test(r0), "run: TNR 12, bold kept", r0);
const p1 = xml(ps[1]);
check(/Câu 3\.<\/w:t>.*<w:t xml:space="preserve"> <\/w:t><w:t>Cho/.test(p1), "tab after 'Câu N.' becomes a space", p1);
check(/<w:i\/>/.test(p1) && /FF0000/.test(p1), "italic/colour kept");
check(/Cambria Math/.test(p1), "equation run font untouched");
check(!/w:type="page"/.test(p1) && /<w:tab\/><w:t>ok/.test(p1), "page break removed, later tabs kept");
check(/sectPr/.test(xml(ps[2])), "section break in pPr kept");
// rPr children stay in schema order (Word rejects out-of-order rPr)
const order = [...ps[1].getElementsByTagNameNS(W, "r")[0].firstChild.childNodes].map((c) => c.localName).join(",");
check(order === "rFonts,i,color,sz,szCs", "run rPr children in schema order", order);
const pOrder = (p) => [...p.getElementsByTagNameNS(W, "pPr")[0].childNodes].map((c) => c.localName).join(",");
check(pOrder(ps[1]) === "spacing,ind,rPr" && pOrder(ps[2]) === "spacing,ind,rPr,sectPr", "pPr children in schema order", pOrder(ps[1]) + " / " + pOrder(ps[2]));

// options: labels and paragraph carry TNR 12 + single spacing (not the document's Normal style)
const opt = OL.rebuildOptionsXml(pkg(`<w:p><w:r><w:t>A. 1</w:t><w:tab/><w:t>B. 2</w:t><w:tab/><w:t>C. 3</w:t><w:tab/><w:t>D. 4</w:t></w:r></w:p>`));
const optPPr = RB.parseXml(opt.xml).getElementsByTagNameNS(W, "pPr")[0];
check([...optPPr.childNodes].map((c) => c.localName).join(",") === "tabs,spacing,ind,rPr", "option pPr in schema order");
check(opt && /w:line="240"/.test(opt.xml) && (opt.xml.match(/Times New Roman/g) || []).length >= 5, "option paragraph + labels: TNR 12, single spacing");

// real exports: every non-table paragraph normalized, objects/equations preserved
for (const file of process.argv.slice(2)) {
  if (!fs.existsSync(file)) { console.log("SKIP", file); continue; }
  const data = JSON.parse(fs.readFileSync(file, "utf8").replace(/^﻿/, ""));
  const before = data.bodyXml;
  const d = RB.parseXml(before);
  const all = RB.bodyParagraphs(RB.documentBody(d));
  all.forEach((p, i) => !data.paragraphs[i].inTable && NX.normalizeParagraph(d, p));
  const after = new XMLSerializer().serializeToString(d);
  const count = (s, re) => (s.match(re) || []).length;
  const name = file.split(/[\\/]/).pop();
  check(count(after, /<w:object/g) === count(before, /<w:object/g) && count(after, /<m:oMath[ >]/g) === count(before, /<m:oMath[ >]/g) && count(after, /<w:drawing/g) === count(before, /<w:drawing/g),
    `${name}: objects, equations, drawings preserved`);
  const bad = all.filter((p, i) => !data.paragraphs[i].inTable && /keepNext|<w:tabs|<w:jc /.test(xml(p.getElementsByTagNameNS(W, "pPr")[0])));
  check(bad.length === 0, `${name}: no keepNext/tab stops/alignment left on ${all.length} paragraphs`, String(bad.length));
}
done();
