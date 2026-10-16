// Real-document check for Chuẩn hoá BigKey: planBlocks on the exported paragraphs and
// rebuildOptionsXml on each real option block (OOXML sliced from the exported body).
//   pwsh -File tests/tools/export-docx.ps1 -Path <exam.docx> -Out <file.json>
//   node tests/bigkey-real.mjs <file.json>
// Skipped (not failed) when no export is given — real exams are not committed to the repo.
import fs from "node:fs";
import { loadModule, check, done } from "./helpers.mjs";

// --has-errors: the document contains real faults (reported, not asserted clean).
const file = process.argv.slice(2).find((a) => !a.startsWith("--")) || process.env.MYSCRIPTS_REAL_DOC;
const hasErrors = process.argv.includes("--has-errors");
if (!file || !fs.existsSync(file)) {
  console.log("SKIP bigkey-real: no exported document (pass a JSON from tests/tools/export-docx.ps1)");
  process.exit(0);
}
const CH = await loadModule("features/bigkey/chuanHoaBigKey.js");
const OL = await loadModule("features/bigkey/optionLayout.js");
const RB = await loadModule("features/dgnl/removeBreakEng.js");
const data = JSON.parse(fs.readFileSync(file, "utf8").replace(/^\uFEFF/, ""));
const P = data.paragraphs.map((p) => ({ text: p.text, inTable: p.inTable, list: p.list }));

const plan = CH.planBlocks(P);
console.log(`questions: ${plan.length}`);
for (const b of plan) console.log(`  Câu ${b.num}: q=${b.q} options=${b.optFrom}..${b.optTo} lời giải=${b.lg} chọn=${b.chonLetter || "-"}`);
check(plan.length > 0, "questions detected");
check(plan.every((b) => b.optFrom >= 0), "every question has an A–D block");

// Build a package that contains only paragraphs [from..to] (indices in Word's paragraph order).
const doc0 = RB.parseXml(data.bodyXml);
const all = RB.bodyParagraphs(RB.documentBody(doc0));
check(all.length === P.length, "OOXML paragraphs map 1:1 to Word paragraphs", `${all.length} vs ${P.length}`);
function slice(from, to) {
  const doc = RB.parseXml(data.bodyXml);
  const body = RB.documentBody(doc);
  const ps = RB.bodyParagraphs(body);
  const keep = new Set(ps.slice(from, to + 1));
  for (const n of [...body.childNodes]) if (n.nodeType === 1 && n.localName !== "sectPr" && !keep.has(n)) body.removeChild(n);
  return new XMLSerializer().serializeToString(doc);
}
for (const b of plan.filter((x) => x.optFrom >= 0)) {
  const res = OL.rebuildOptionsXml(slice(b.optFrom, b.optTo), b.chonLetter);
  const objs = res ? (res.xml.match(/<w:object/g) || []).length : 0;
  const before = (slice(b.optFrom, b.optTo).match(/<w:object/g) || []).length;
  check(!!res && objs === before && (b.chonLetter ? res.correct.join("") === b.chonLetter : res.correct.length === 1), `Câu ${b.num} options rebuilt`,
    res ? `layout=${res.layout} widths=${res.widths} correct=${res.correct} objects ${before}->${objs}` : "null");
}

// Check lỗi đề on the real document: a clean exam must report no errors…
const CK = await loadModule("features/bigkey/checkLoiDe.js");
const optXml = plan.map((b) => (b.optFrom >= 0 ? slice(b.optFrom, b.optTo) : ""));
const gapXml = plan.map((b) => (b.optTo >= 0 && b.lg > b.optTo + 1 ? [...Array(b.lg - b.optTo - 1)].map((_, i) => slice(b.optTo + 1 + i, b.optTo + 1 + i)) : []));
const errs = CK.findErrors(P, plan, optXml, gapXml);
const hard = errs.filter((e) => !e.warn).map((e) => e.msg);
console.log("  warnings:", JSON.stringify(errs.filter((e) => e.warn).map((e) => e.msg)));
if (hasErrors) console.log("  REPORTED (document faults):", JSON.stringify(hard));
else check(hard.length === 0, "check lỗi đề: no errors", JSON.stringify(hard));
// …and injected faults must be caught.
const P2 = P.map((p) => ({ ...p }));
const iC = [...Array(plan[1].optTo - plan[1].optFrom + 1)].map((_, i) => plan[1].optFrom + i).find((i) => /(^|[\t ])C\s*[.)]/.test(P2[i].text));
P2[iC] = { ...P2[iC], text: P2[iC].text.replace(/(^|[\t ])C(\s*[.)])/, "$1B$2") }; // duplicate B, missing C
P2[plan[2].q] = { ...P2[plan[2].q], text: P2[plan[2].q].text.replace(/Câu\s*\d+/,"Câu " + (plan[2].num + 2)) }; // numbering jump
const plan2 = CK.findErrors ? (await loadModule("features/bigkey/chuanHoaBigKey.js")).planBlocks(P2) : [];
const errs2 = CK.findErrors(P2, plan2, optXml, gapXml).map((e) => e.msg);
check(errs2.some((m) => /thiếu phương án C/.test(m)) && errs2.some((m) => /trùng nhãn phương án B/.test(m)) && errs2.some((m) => /không liên tục/.test(m)),
  "check lỗi đề: injected faults detected", JSON.stringify(errs2));
done();
