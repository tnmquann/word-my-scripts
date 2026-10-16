// Simulate "Chuẩn hoá BigKey (cả đề)" on an exported document (tests/tools/export-docx.ps1) and write the
// result as a Flat OPC .xml that Word opens directly, to compare with the original VBA output.
//   node tests/tools/bigkey-simulate.mjs <export.json> <out.xml> [startNumber=1]
// Uses the real modules for steps 1–3 (normalizeParagraphsXml per paragraph, planBlocks, rebuildOptionsXml);
// steps done with plain Office JS calls in the command (labels, Lời giải, Chọn X, empty lines, page setup)
// are mimicked in XML. Render: open <out.xml> in Word (or pwsh tests/tools/render-pdf.ps1).
import fs from "node:fs";
import { loadModule } from "../helpers.mjs";

const [, , input, out, startArg] = process.argv;
if (!input || !out) throw new Error("usage: node tests/tools/bigkey-simulate.mjs <export.json> <out.xml> [start]");
const RB = await loadModule("features/dgnl/removeBreakEng.js");
const NX = await loadModule("features/bigkey/normalizeXml.js");
const OL = await loadModule("features/bigkey/optionLayout.js");
const CH = await loadModule("features/bigkey/chuanHoaBigKey.js");
const W = NX.W;

const data = JSON.parse(fs.readFileSync(input, "utf8").replace(/^﻿/, ""));
const P = data.paragraphs.map((p) => ({ text: p.text, inTable: p.inTable, list: p.list }));
const doc = RB.parseXml(data.bodyXml);
const body = RB.documentBody(doc);
let paras = RB.bodyParagraphs(body);
if (paras.length !== P.length) throw new Error(`paragraph mismatch ${paras.length} vs ${P.length}`);

// 1) list numbers -> text are already text in most exams; 2) normalize every non-table paragraph
paras.forEach((p, i) => !P[i].inTable && NX.normalizeParagraph(doc, p));

// 3) options, from the end
const plan = CH.planBlocks(P);
const ser = new XMLSerializer();
for (const b of [...plan].reverse()) {
  if (b.optFrom < 0) continue;
  // package with only these paragraphs, like Range.getOoxml()
  const pkg = RB.parseXml(ser.serializeToString(doc));
  const pb = RB.documentBody(pkg);
  const keep = new Set(RB.bodyParagraphs(pb).slice(b.optFrom, b.optTo + 1));
  for (const n of [...pb.childNodes]) if (n.nodeType === 1 && n.localName !== "sectPr" && !keep.has(n)) pb.removeChild(n);
  const res = OL.rebuildOptionsXml(ser.serializeToString(pkg), b.chonLetter);
  if (!res) continue;
  const newPs = RB.bodyParagraphs(RB.documentBody(RB.parseXml(res.xml))).map((n) => doc.importNode(n, true));
  const first = paras[b.optFrom];
  for (const n of newPs) first.parentNode.insertBefore(n, first);
  for (let i = b.optFrom; i <= b.optTo; i++) paras[i].parentNode.removeChild(paras[i]);
  paras.splice(b.optFrom, b.optTo - b.optFrom + 1, ...newPs);
  P.splice(b.optFrom, b.optTo - b.optFrom + 1, ...newPs.map((n) => ({ text: n.textContent, inTable: false, list: "" })));
}

// 4) labels "Câu i:" bold blue + hanging 1 cm; "Lời giải" bold blue; "Chọn X" highlighted
const mk = (name, attrs = {}) => {
  const e = doc.createElementNS(W, "w:" + name);
  for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, "w:" + k, v);
  return e;
};
const setRun = (r, f) => {
  const rPr = NX.setFont(doc, r);
  for (const c of [...rPr.childNodes]) if (["b", "color", "u", "highlight"].includes(c.localName)) rPr.removeChild(c);
  if (f.b) rPr.insertBefore(mk("b"), rPr.childNodes[1] || null);
  if (f.color) rPr.appendChild(mk("color", { val: f.color }));
  if (f.hl) rPr.appendChild(mk("highlight", { val: f.hl }));
};
const plan2 = CH.planBlocks(P);
const start = +(startArg || 1);
plan2.forEach((b, k) => {
  const p = paras[b.q];
  const ind = p.getElementsByTagNameNS(W, "ind")[0];
  ind.removeAttributeNS(W, "firstLine");
  ind.setAttributeNS(W, "w:hanging", "567");
  // label: the first text node holding "Câu N." -> split off a bold blue run "Câu i:"
  for (const t of p.getElementsByTagNameNS(W, "t")) {
    const m = t.textContent.match(/^(\s*)Câu\s*\d+\s*[.:]/);
    if (!m) continue;
    const r = t.parentNode;
    const rest = r.cloneNode(true);
    rest.getElementsByTagNameNS(W, "t")[0].textContent = t.textContent.slice(m[0].length);
    rest.getElementsByTagNameNS(W, "t")[0].setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = `Câu ${start + k}:`;
    setRun(r, { b: true, color: "0000FF" });
    r.parentNode.insertBefore(rest, r.nextSibling);
    break;
  }
  if (b.lg >= 0) for (const r of paras[b.lg].getElementsByTagNameNS(W, "r")) setRun(r, { b: true, color: "0000FF" });
  if (b.chon >= 0) for (const r of paras[b.chon].getElementsByTagNameNS(W, "r")) setRun(r, { hl: "yellow" });
});

// 5) empty paragraphs (no objects) removed; page A4, margins 1.5 cm
paras.forEach((p, i) => {
  if (P[i].inTable || p.textContent.trim() !== "" || /drawing|pict|object|oMath/.test(ser.serializeToString(p))) return;
  if (p.parentNode === body) body.removeChild(p);
});
for (const s of doc.getElementsByTagNameNS(W, "sectPr")) {
  for (const c of [...s.childNodes]) if (c.localName === "pgSz" || c.localName === "pgMar") s.removeChild(c);
  s.insertBefore(mk("pgMar", { top: "850", right: "850", bottom: "850", left: "850", header: "720", footer: "720", gutter: "0" }), s.firstChild);
  s.insertBefore(mk("pgSz", { w: "11906", h: "16838" }), s.firstChild);
}
fs.writeFileSync(out, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n' + ser.serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, ""));
console.log(`questions: ${plan2.length}, written ${out}`);
