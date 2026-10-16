// Round-trip LaTeX -> OMML -> LaTeX for representative formulas (XML DOM from jsdom).
// Every sample must convert, and none may fall back to \mathrm{<unknown command>}.
import { loadModule, check, done } from "./helpers.mjs";

const EQ = await loadModule("features/equations/latexEquations.js");
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const samples = [
  "\\frac{a+b}{c}", "x^{2}+y_{1}", "\\sum_{i=1}^{n} i^2", "\\int_0^1 f(x)\\,dx", "\\sqrt[3]{x+1}", "\\binom{n}{k}",
  "\\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}", "f(x)=\\begin{cases} x & x>0 \\\\ -x & x\\le 0 \\end{cases}",
  "\\overset{\\frown}{AB}", "\\underbrace{a+b}_{n}", "\\mathbb{R}", "\\mathcal{L}", "\\mathbf{v}", "\\operatorname{rank}(A)",
  "\\lVert x\\rVert", "\\text{khi } x>0", "\\lim_{x\\to 0}\\frac{\\sin x}{x}", "\\alpha+\\beta\\geq\\gamma",
];
for (const s of samples) {
  const om = EQ.conv(s);
  const x = `<w:document ${NS}><w:body><w:p><w:r><w:t>a </w:t></w:r><m:oMath>${om}</m:oMath></w:p></w:body></w:document>`;
  const res = EQ.ommlToLatexXml(x);
  const back = [...new DOMParser().parseFromString(res.xml, "application/xml").getElementsByTagName("w:t")].map((e) => e.textContent).join("").replace(/^a /, "");
  check(om.includes("<m:") && !/\\mathrm\{[a-z]{3,}\}/.test(back), s, "=> " + back);
}
done();
