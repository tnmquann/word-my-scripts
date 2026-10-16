// rebuildOptionsXml (BigKey option layout): 1/2/4 options per line, correct answer, equations kept.
import { loadModule, check, done } from "./helpers.mjs";

const OL = await loadModule("features/bigkey/optionLayout.js");

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"';
const r = (t, u) => `<w:r>${u ? '<w:rPr><w:u w:val="single"/></w:rPr>' : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const tab = "<w:r><w:tab/></w:r>";
const p = (...runs) => `<w:p>${runs.join("")}</w:p>`;
const pkg = (body) => `<pkg:package ${NS}><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const math = "<m:oMath><m:r><m:t>x^2+1</m:t></m:r></m:oMath>";

const cases = [
  ["short options, one line with tabs, B underlined", pkg(p(r("A. 1"), tab, r("B", true), r(". 2"), tab, r("C. 3"), tab, r("D. 4"))), { layout: 4, correct: "B", paras: 1, text: ["A. 1⇥B. 2⇥C. 3⇥D. 4"] }],
  ["medium options on 4 paragraphs", pkg(p(r("A. Hàm số đồng biến trên khoảng (0;1)")) + p(r("B. Hàm số nghịch biến trên (0;1)")) + p(r("C. Hàm số đồng biến trên R")) + p(r("D. Hàm số không đổi"))), { layout: 2, correct: "", paras: 2 }],
  ["long options", pkg(p(r("A. " + "Đây là một phương án rất dài để kiểm tra việc xuống dòng ".repeat(2))) + p(r("B. ngắn")) + p(r("C. ngắn")) + p(r("D. ngắn"))), { layout: 1, correct: "", paras: 4 }],
  ["options with OMML equation, D underlined", pkg(p(r("A. "), math, tab, r("B. "), math, tab, r("C. 2"), tab, r("D", true), r(". 5"))), { layout: 4, correct: "D", paras: 1, math: 2 }],
  ["not an option block (text before A.)", pkg(p(r("Cho hàm số A. B. C. D."))), null],
  ["'B.' inside option A content is not a label", pkg(p(r("A. điểm B. không"), tab, r("B. 2"), tab, r("C. 3"), tab, r("D. 4"))), { layout: 4, correct: "", paras: 1, text: ["A. điểm B. không⇥B. 2⇥C. 3⇥D. 4"] }],
  ["space-separated options on one line (lenient)", pkg(p(r("A. 1 B. 2 C. 3 D. 4"))), { layout: 4, correct: "", paras: 1, text: ["A. 1⇥B. 2⇥C. 3⇥D. 4"] }],
];

const textOf = (pEl) =>
  [...pEl.childNodes].map((n) => (n.localName === "r" ? [...n.childNodes].map((x) => (x.localName === "tab" ? "⇥" : x.localName === "t" ? x.textContent : "")).join("") : n.localName === "oMath" ? "[EQ]" : "")).join("");

for (const [name, xml, e] of cases) {
  const res = OL.rebuildOptionsXml(xml);
  if (e === null) { check(res === null, name); continue; }
  if (!res) { check(false, name, "got null"); continue; }
  const d = new DOMParser().parseFromString(res.xml, "application/xml");
  const ps = [...d.getElementsByTagName("w:p")];
  const got = { layout: res.layout, correct: res.correct.join(""), paras: ps.length, math: d.getElementsByTagName("m:oMath").length, text: ps.map(textOf) };
  const ok = got.layout === e.layout && got.correct === e.correct && got.paras === e.paras && (!e.math || got.math === e.math) && (!e.text || JSON.stringify(got.text) === JSON.stringify(e.text));
  check(ok, name, JSON.stringify(got));
}
done();
