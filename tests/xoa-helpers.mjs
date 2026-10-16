// Pure helpers of Xoá cần thiết (src/features/xoa/xoaCanThiet.js) on synthetic paragraphs,
// plus the real exports when given (same JSON as bigkey-real.mjs):
//   node tests/xoa-helpers.mjs [lhp.json eng.json …]
import fs from "node:fs";
import { loadModule, check, done } from "./helpers.mjs";

const X = await loadModule("features/xoa/xoaCanThiet.js");
const reLG = /^\s*(L[ờơo]i\s*gi[ảa]i|Hướng\s*dẫn\s*giải)/i;
const P = (lines) => lines.map((t) => (typeof t === "string" ? { text: t, inTable: false, list: "" } : t));

// --- blocks: a solution ends at the next question / group lead-in, never swallows a passage
const doc = P([
  "Câu 1. [2D1-3] (Sở Hà Nội) Hỏi gì?", "A. 1\tB. 2\tC. 3\tD. 4", "Lời giải", "Chọn A", "Giải thích",
  "Questions 2-3: Read the following passage and choose the best answer.", "Passage line 1", "Passage line 2",
  "Câu 2. [1H2-1] Hỏi tiếp?", "A. x\tB. y\tC. z\tD. t", "Lời giải", { text: "ô bảng", inTable: true, list: "" }, "Chọn B",
  "Đọc đoạn văn sau và trả lời các câu hỏi từ 3 đến 4", "Đoạn văn",
  { text: "Hỏi cuối? [2D3-4]", inTable: false, list: "Câu 3." }, "A. a\tB. b\tC. c\tD. d", "Lời giải", "Chọn C",
  "BẢNG ĐÁP ÁN", { text: "1A", inTable: true, list: "" },
]);
check(JSON.stringify(X.findBlocks(doc, reLG)) === "[[2,4],[10,12],[17,18]]", "solution blocks stop at lead-ins and BẢNG ĐÁP ÁN", JSON.stringify(X.findBlocks(doc, reLG)));
const qb = X.questionBlocks(doc).map((b) => [b.q, b.end, b.num]);
check(JSON.stringify(qb) === "[[0,4,1],[8,12,2],[15,18,3]]", "question blocks (typed + auto-numbered) exclude passages", JSON.stringify(qb));

// --- relative indent (as Xoá ĐACT ĐGNL): a passage at the margin with NO lead-in text still ends the solution
const ind = (text, left) => ({ text, inTable: false, list: "", left });
const passage = [ind("Câu 1. Hỏi?", 28), ind("A. 1\tB. 2\tC. 3\tD. 4", 28), ind("Lời giải", 28), ind("Giải thích", 28),
  ind("Some reading passage without any lead-in line.", 0), ind("Second passage paragraph.", 0), ind("Câu 2. Hỏi?", 28)];
check(JSON.stringify(X.findBlocks(passage, reLG)) === "[[2,3]]", "solution block stops at a less-indented passage", JSON.stringify(X.findBlocks(passage, reLG)));
const flat = passage.map((p) => ({ ...p, left: 28 })); // BigKey layout: everything at 1 cm -> text rules only
check(JSON.stringify(X.findBlocks(flat, reLG)) === "[[2,5]]", "equal indents: block runs to the next question (text rules only)");

// --- Xoá dòng viết LG: after options / Lời giải / dotted lines, stops at HẾT and less-indented passages
const wr = [ind("Câu 1. Hỏi?", 28), ind("A. 1\tB. 2\tC. 3\tD. 4", 28), ind("....................", 28), ind("....................", 28),
  ind("Câu 2. Tự luận: chứng minh?", 28), ind("Gợi ý: dùng định lí", 28), ind("..........................", 28),
  ind("Câu 3. Hỏi?", 28), ind("A. 1", 28), ind("B. 2", 28), ind("C. 3", 28), ind("D. 4", 28), ind("Lời giải", 28), ind("Chọn A", 28),
  ind("----- HẾT -----", 28), ind("Passage at the margin", 0)];
check(JSON.stringify(X.writingBlocks(wr)) === "[[2,3],[6,6],[12,13]]", "writingBlocks: MC dots, essay dots (stem kept), Lời giải; HẾT kept", JSON.stringify(X.writingBlocks(wr)));
check(X.findBlocks([ind("Lời giải", 28), ind("x", 28), ind("HẾT", 28)], reLG).join() === "0,1", "HẾT ends a solution block");
check(X.findBlocks([ind("Lời giải", 28), ind("HẾT HẠN nộp bài là…", 28)], reLG).join() === "0,1", "'HẾT HẠN…' text is not a boundary");

// --- tags
check(X.leadingSourceTags(" [2D1-3] (Sở Hà Nội) Hỏi gì?") === "[2D1-3] (Sở Hà Nội)", "leading source tags");
check(!X.authorParts("Tác giảng viên").wholeParagraph, "author: no false match on longer word");
check(X.leadingSourceTags("\tHỏi (x) gì?") === "", "no tag when text comes first");
check(X.levelOf("[2D1-3] abc") === 3 && X.levelOf("[1H2-1]") === 1 && X.levelOf("abc") === null, "levelOf");
check(JSON.stringify(X.idTags("a [2D1-3] b [H3-4] c [x] [1D2]")) === '["[2D1-3]","[H3-4]","[1D2]"]', "idTags");
check(X.authorParts("Tác giả: Nguyễn A").wholeParagraph, "author line");
check(JSON.stringify(X.authorParts("Câu 1. [(Tác giả: B)] Hỏi").inline) === '["[(Tác giả: B)]"]', "inline author", JSON.stringify(X.authorParts("Câu 1. [(Tác giả: B)] Hỏi").inline));

// --- real documents: no block may contain another question or a passage lead-in
for (const file of process.argv.slice(2)) {
  if (!fs.existsSync(file)) { console.log("SKIP", file); continue; }
  const R = JSON.parse(fs.readFileSync(file, "utf8").replace(/^﻿/, "")).paragraphs.map((p) => ({ text: p.text, inTable: p.inTable, list: p.list }));
  const lg = X.findBlocks(R, reLG);
  const qs = X.questionBlocks(R);
  const bad = lg.filter(([a, b]) => R.slice(a + 1, b + 1).some((p) => X.isQuestion(p) || /^\s*Questions?\s*\d/.test(p.text)));
  check(lg.length === qs.length && !bad.length, `${file.split(/[\\/]/).pop()}: ${lg.length} solution blocks for ${qs.length} questions, none crosses a question/lead-in`, JSON.stringify(bad));
  const passages = R.map((p, i) => i).filter((i) => /^\s*Questions?\s*\d/.test(R[i].text));
  check(passages.every((i) => !lg.some(([a, b]) => a <= i && i <= b) && !qs.some((q) => q.q <= i && i <= q.end)), `${file.split(/[\\/]/).pop()}: lead-ins kept (${passages.length})`);
}
done();
