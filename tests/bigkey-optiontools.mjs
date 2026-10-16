// optionTools (BigKey option-block tools that keep the layout): mark answer, read answers,
// highlight -> red, split options onto their own paragraphs, add/remove trailing periods.
//   node tests/bigkey-optiontools.mjs [export.json …]
// With exported documents (tests/tools/export-docx.ps1), every function also runs on every real option block.
import fs from "node:fs";
import { loadModule, check, done } from "./helpers.mjs";

const OT = await loadModule("features/bigkey/optionTools.js");
const { RPR_ORDER } = await loadModule("features/bigkey/normalizeXml.js");

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"';
const r = (t, pr = "") => `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const U = '<w:u w:val="single"/>';
const HL = '<w:highlight w:val="yellow"/>';
const tab = "<w:r><w:tab/></w:r>";
const TABS = '<w:pPr><w:tabs><w:tab w:val="left" w:pos="2977"/><w:tab w:val="left" w:pos="5387"/><w:tab w:val="left" w:pos="7797"/></w:tabs><w:ind w:left="567"/></w:pPr>';
const p = (...runs) => `<w:p>${runs.join("")}</w:p>`;
const pt = (...runs) => `<w:p>${TABS}${runs.join("")}</w:p>`;
const pkg = (body) => `<pkg:package ${NS}><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const math = "<m:oMath><m:r><m:t>x^2</m:t></m:r></m:oMath>";
const mathHL = `<m:oMath><m:r><w:rPr>${HL}</w:rPr><m:t>y</m:t></m:r></m:oMath>`;
const obj = '<w:r><w:object><v:shape style="width:20pt"/></w:object></w:r>';

// Render a paragraph: plain text, ⇥ tab, ↵ br, [EQ], [OBJ]; styled runs as <flags:text>
// (b bold, B blue, R red, u underline, H highlight); adjacent runs with the same flags merged.
const kid = (e, name) => e && [...e.childNodes].find((c) => c.localName === name);
function flags(rPr) {
  if (!rPr) return "";
  const c = kid(rPr, "color"), u = kid(rPr, "u"), h = kid(rPr, "highlight");
  const cv = c && c.getAttribute("w:val");
  return (kid(rPr, "b") ? "b" : "") + (cv ? (cv === "FF0000" ? "R" : cv === "0000FF" ? "B" : "?") : "") +
    (u && u.getAttribute("w:val") !== "none" ? "u" : "") + (h && h.getAttribute("w:val") !== "none" ? "H" : "");
}
function renderP(pEl) {
  const segs = [];
  const push = (f, s) => { const l = segs[segs.length - 1]; if (l && l[0] === f) l[1] += s; else segs.push([f, s]); };
  for (const n of pEl.childNodes) {
    if (n.localName === "r") {
      const f = flags(kid(n, "rPr"));
      for (const c of n.childNodes) {
        if (c.localName === "t") push(f, c.textContent);
        else if (c.localName === "tab") push("", "⇥");
        else if (c.localName === "br") push("", "↵");
        else if (c.localName === "object") push("", "[OBJ]");
      }
    } else if (n.localName === "oMath") push("", "[EQ]");
  }
  return segs.map(([f, s]) => (f ? `<${f}:${s}>` : s)).join("");
}
/** Parse an output: paragraphs before the sentinel, plus structural checks shared by every case. */
function parse(xml, name) {
  const d = new DOMParser().parseFromString(xml, "application/xml");
  const body = d.getElementsByTagName("w:body")[0];
  const els = [...body.childNodes].filter((n) => n.nodeType === 1);
  const ps = els.filter((n) => n.localName === "p");
  const sentinel = ps[ps.length - 1];
  const sectLast = !els.some((n) => n.localName === "sectPr") || els[els.length - 1].localName === "sectPr";
  const orderOk = [...d.getElementsByTagName("w:rPr")].every((rp) => {
    const ranks = [...rp.childNodes].filter((c) => c.nodeType === 1).map((c) => RPR_ORDER.indexOf(c.localName));
    return ranks.every((x, i) => i === 0 || x >= ranks[i - 1]);
  });
  if (!(sentinel && !sentinel.hasChildNodes() && sectLast && orderOk)) check(false, `${name}: structure`, `sentinel=${!!sentinel && !sentinel.hasChildNodes()} sectLast=${sectLast} rPrOrder=${orderOk}`);
  return { d, paras: ps.slice(0, -1) };
}
const texts = (xml, name) => parse(xml, name).paras.map(renderP);
const eq = (got, exp) => JSON.stringify(got) === JSON.stringify(exp);
const tabStops = (pEl) => kid(kid(pEl, "pPr"), "tabs") ? kid(kid(pEl, "pPr"), "tabs").childNodes.length : 0;
const noSentinel = (xml) => xml.replace(/<w:p\/>/g, "");

// ---- fixtures ----
const oneLine = pkg(pt(r("A. 1"), tab, r("B", U), r(". 2"), tab, r("C. 3"), tab, r("D. 4")));
const twoLine = pkg(pt(r("A. "), math, tab, r("B. "), obj) + pt(r("C. 3"), tab, r("D. 4")));
const fourPara = pkg(p(r("A) x")) + p(r("B) y")) + p(r("C: z")) + p(r("D. w")));
const splitRuns = pkg(pt(r("A. 1"), tab, r("B", "<w:i/>"), r(".", "<w:i/><w:sz w:val=\"24\"/>"), r(" 2"), tab, r("C. 3"), tab, r("D. 4")));
const notLabel = pkg(pt(r("A. điểm B. không"), tab, r("B. 2"), tab, r("C. 3"), tab, r("D. 4")));
const hlC = pkg(pt(r("A. 1"), tab, r("B", U), r(". 2"), tab, r("C. "), r("3", HL), tab, r("D. 4")));
const hlMath = pkg(pt(r("A. 1"), tab, r("B. 2"), tab, r("C. 3"), tab, r("D. "), mathHL));
const lenient = pkg(p(r("A. 1 B. 2 C. 3 D. 4")));
const inRunBr = pkg(pt('<w:r><w:t xml:space="preserve">A. 1</w:t><w:br/><w:t xml:space="preserve">B. 2</w:t></w:r>', tab, r(" C. 3"), tab, r("D. 4")));

// ---- markAnswerXml ----
const markCases = [
  ["mark C on one tabbed line (B was underlined)", oneLine, "C", ["<bB:A.> 1⇥<bB:B.> 2⇥<bRu:C><bB:.> 3⇥<bB:D.> 4"], "C"],
  ["keep the underlined answer (B)", oneLine, "keep", ["<bB:A.> 1⇥<bRu:B><bB:.> 2⇥<bB:C.> 3⇥<bB:D.> 4"], "B"],
  ["null keeps the underlined answer", oneLine, null, ["<bB:A.> 1⇥<bRu:B><bB:.> 2⇥<bB:C.> 3⇥<bB:D.> 4"], "B"],
  ["no answer mode", oneLine, "none", ["<bB:A.> 1⇥<bB:B.> 2⇥<bB:C.> 3⇥<bB:D.> 4"], ""],
  ["4 paragraphs, 'A)'/'C:' labels -> 'A.'", fourPara, "D", ["<bB:A.> x", "<bB:B.> y", "<bB:C.> z", "<bRu:D><bB:.> w"], "D"],
  ["2 lines with equation and object", twoLine, "A", ["<bRu:A><bB:.> [EQ]⇥<bB:B.> [OBJ]", "<bB:C.> 3⇥<bB:D.> 4"], "A"],
  ["label split across runs", splitRuns, "B", ["<bB:A.> 1⇥<bRu:B><bB:.> 2⇥<bB:C.> 3⇥<bB:D.> 4"], "B"],
  ["'B.' inside option A is not a label", notLabel, "B", ["<bB:A.> điểm B. không⇥<bRu:B><bB:.> 2⇥<bB:C.> 3⇥<bB:D.> 4"], "B"],
];
for (const [name, xml, ans, exp, letter] of markCases) {
  const res = OT.markAnswerXml(xml, ans);
  if (!res) { check(false, name, "null"); continue; }
  const got = texts(res.xml, name);
  check(eq(got, exp) && res.letter === letter, `markAnswerXml: ${name}`, JSON.stringify(got) + " " + res.letter);
  const again = OT.markAnswerXml(res.xml, "keep");
  check(noSentinel(again.xml) === noSentinel(res.xml), `markAnswerXml: ${name}: idempotent`);
}
{
  const { paras } = parse(OT.markAnswerXml(oneLine, "C").xml, "tabs");
  check(paras.length === 1 && tabStops(paras[0]) === 3, "markAnswerXml keeps pPr tab stops");
  const { d } = parse(OT.markAnswerXml(splitRuns, "B").xml, "italic");
  const bRun = [...d.getElementsByTagName("w:r")].find((x) => (kid(x, "t") || {}).textContent === "B");
  check(eq([...kid(bRun, "rPr").childNodes].map((c) => c.localName), ["b", "i", "color", "u"]), "markAnswerXml keeps other rPr (italic) in schema order");
}
check(OT.markAnswerXml(pkg(p(r("Cho hàm số y = x"))), "A") === null, "markAnswerXml: not an option block -> null");

// ---- readAnswersXml ----
check(eq(OT.readAnswersXml(hlC), { underlined: "B", highlighted: "C" }), "readAnswersXml: underlined B, highlighted C");
check(eq(OT.readAnswersXml(hlMath), { underlined: "", highlighted: "D" }), "readAnswersXml: highlight inside an equation");
check(eq(OT.readAnswersXml(oneLine), { underlined: "B", highlighted: "" }), "readAnswersXml: no highlight");

// ---- highlightToRedXml ----
{
  const res = OT.highlightToRedXml(hlC);
  const got = texts(res.xml, "hl");
  check(res.letter === "C" && eq(got, ["<bB:A.> 1⇥<bB:B.> 2⇥<bRu:C><bB:.> <R:3>⇥<bB:D.> 4"]), "highlightToRedXml: C content red, C marked", JSON.stringify(got));
  const res2 = OT.highlightToRedXml(hlMath);
  const { d } = parse(res2.xml, "hlmath");
  const mr = d.getElementsByTagName("m:r")[0];
  check(res2.letter === "D" && !d.getElementsByTagName("w:highlight").length && flags(kid(mr, "rPr")) === "R", "highlightToRedXml: equation highlight -> red, D marked");
  const res3 = OT.highlightToRedXml(oneLine);
  check(res3.letter === "" && eq(texts(res3.xml, "nohl"), texts(oneLine.replace("<w:sectPr/>", "<w:p/><w:sectPr/>"), "orig")), "highlightToRedXml: nothing highlighted -> unchanged");
}

// ---- splitOptionsXml ----
const splitCases = [
  ["one tabbed line", oneLine, ["A. 1", "<u:B>. 2", "C. 3", "D. 4"]],
  ["2 lines with equation and object", twoLine, ["A. [EQ]", "B. [OBJ]", "C. 3", "D. 4"]],
  ["already 4 paragraphs", fourPara, ["A) x", "B) y", "C: z", "D. w"]],
  ["in-run line break and '⇥ C.'", inRunBr, ["A. 1", "B. 2 ", "C. 3", "D. 4"]],
  ["space-separated options are not split", lenient, ["A. 1 B. 2 C. 3 D. 4"]],
];
for (const [name, xml, exp] of splitCases) {
  const res = OT.splitOptionsXml(xml);
  const got = texts(res.xml, name);
  check(eq(got, exp), `splitOptionsXml: ${name}`, JSON.stringify(got));
  check(noSentinel(OT.splitOptionsXml(res.xml).xml) === noSentinel(res.xml), `splitOptionsXml: ${name}: idempotent`);
}
{
  const { paras } = parse(OT.splitOptionsXml(oneLine).xml, "pPr");
  const ind = (x) => kid(kid(x, "pPr"), "ind").getAttribute("w:left");
  check(tabStops(paras[0]) === 3 && paras.slice(1).every((x) => tabStops(x) === 0 && ind(x) === "567"), "splitOptionsXml: new paragraphs copy pPr without tab stops");
}

// ---- optionPeriodsXml ----
const mixed = pkg(pt(r("A. 1."), tab, r("B. 2  "), tab, r("C. "), math, tab, r("D. "), obj));
{
  const add = OT.optionPeriodsXml(mixed, true);
  const got = texts(add.xml, "add");
  check(eq(got, ["A. 1.⇥B. 2.  ⇥C. [EQ].⇥D. [OBJ]."]), "optionPeriodsXml add: text, equation, object; no '..'", JSON.stringify(got));
  check(noSentinel(OT.optionPeriodsXml(add.xml, true).xml) === noSentinel(add.xml), "optionPeriodsXml add: idempotent");
  const rem = OT.optionPeriodsXml(noSentinel(add.xml), false); // input from Word never carries our sentinel
  const got2 = texts(rem.xml, "rem");
  check(eq(got2, ["A. 1⇥B. 2  ⇥C. [EQ]⇥D. [OBJ]"]), "optionPeriodsXml remove: one trailing '.' each, trailing spaces ignored", JSON.stringify(got2));
  const { paras } = parse(rem.xml, "rem");
  check(tabStops(paras[0]) === 3, "optionPeriodsXml keeps pPr tab stops");
  const got3 = texts(OT.optionPeriodsXml(fourPara, true).xml, "4p");
  check(eq(got3, ["A) x.", "B) y.", "C: z.", "D. w."]), "optionPeriodsXml add on 4 paragraphs", JSON.stringify(got3));
}

// ---- real exported documents ----
const files = process.argv.slice(2).filter((a) => fs.existsSync(a));
if (!files.length) console.log("SKIP real option blocks: no exported document given");
if (files.length) {
  const CH = await loadModule("features/bigkey/chuanHoaBigKey.js");
  const RB = await loadModule("features/dgnl/removeBreakEng.js");
  const count = (x, re) => (x.match(re) || []).length;
  const objCounts = (x) => [/<w:object[ >]/g, /<m:oMath[ >]/g, /<w:drawing[ >]/g].map((re) => count(x, re)).join("/");
  const textOf = (x) => [...x.matchAll(/<(w|m):t(?: [^>]*)?>([^<]*)<\/\1:t>/g)].map((m) => m[2]).join("").replace(/[.):\s]/g, "");
  for (const file of files) {
    const data = JSON.parse(fs.readFileSync(file, "utf8").replace(/^\uFEFF/, ""));
    const plan = CH.planBlocks(data.paragraphs.map((x) => ({ text: x.text, inTable: x.inTable, list: x.list })));
    const slice = (from, to) => {
      const doc = RB.parseXml(data.bodyXml);
      const body = RB.documentBody(doc);
      const keep = new Set(RB.bodyParagraphs(body).slice(from, to + 1));
      for (const n of [...body.childNodes]) if (n.nodeType === 1 && n.localName !== "sectPr" && !keep.has(n)) body.removeChild(n);
      return new XMLSerializer().serializeToString(doc);
    };
    const fails = [];
    let blocks = 0;
    for (const b of plan.filter((x) => x.optFrom >= 0)) {
      blocks++;
      const src = slice(b.optFrom, b.optTo);
      const runs = {
        markKeep: OT.markAnswerXml(src, "keep"), markNone: OT.markAnswerXml(src, "none"), hlRed: OT.highlightToRedXml(src),
        split: OT.splitOptionsXml(src), addDot: OT.optionPeriodsXml(src, true), remDot: OT.optionPeriodsXml(src, false),
      };
      const ra = OT.readAnswersXml(src);
      if (!ra) fails.push(`Câu ${b.num} readAnswers null`);
      for (const [k, res] of Object.entries(runs)) {
        if (!res) { fails.push(`Câu ${b.num} ${k} null`); continue; }
        parse(res.xml, `${file} Câu ${b.num} ${k}`);
        if (objCounts(res.xml) !== objCounts(src)) fails.push(`Câu ${b.num} ${k} objects ${objCounts(src)} -> ${objCounts(res.xml)}`);
        if (textOf(res.xml) !== textOf(src)) fails.push(`Câu ${b.num} ${k} text changed`);
      }
      if (runs.markKeep && ra && runs.markKeep.letter !== ra.underlined) fails.push(`Câu ${b.num} markKeep letter ${runs.markKeep.letter} vs ${ra.underlined}`);
      for (const [k, f] of [["markKeep", (x) => OT.markAnswerXml(x, "keep")], ["split", OT.splitOptionsXml], ["addDot", (x) => OT.optionPeriodsXml(x, true)]]) {
        if (runs[k] && noSentinel(f(runs[k].xml).xml) !== noSentinel(runs[k].xml)) fails.push(`Câu ${b.num} ${k} not idempotent`);
      }
    }
    check(blocks > 0 && !fails.length, `real ${file.split(/[\\/]/).pop()}: ${blocks} option blocks`, fails.slice(0, 10).join("; "));
  }
}
done();
