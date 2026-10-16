// Check lỗi đề (BigKey)  (spec: docs/specs/bigkey/Chuan_Hoa_BigKey.md)
// Port + viết lại bkCheckLoiDe của BigKeyUI.dotm. Bản VBA dừng ở lỗi ĐẦU TIÊN và dựa vào toạ độ dòng/trang;
// bản này kiểm tra theo cấu trúc đoạn, liệt kê TẤT CẢ lỗi và chọn sẵn câu lỗi đầu tiên.
// Lỗi: dữ liệu trước Câu 1, số câu trùng/không liên tục, thiếu/trùng/sai thứ tự phương án A–D,
//      thiếu đáp án (không gạch chân và không có "Chọn X"), hình nằm lạc giữa phương án và "Lời giải".
// Bảng: chỉ cảnh báo (Chuẩn hoá BigKey giữ nguyên bảng).

import { planBlocks } from "./chuanHoaBigKey.js";
import { rebuildOptionsXml } from "./optionLayout.js";

const reLabel = (L) => new RegExp("(^|[\\t \\n])" + L + "\\s*[.)]", "g"); // các đoạn được ghép bằng \n
const hasObject = (xml) => /<w:drawing|<w:pict|<w:object/.test(xml);

/**
 * Thuần: P = [{text, inTable, list}], plan = planBlocks(P),
 * optXml[k] = OOXML khối phương án của câu k (hoặc ""), gapXml[k] = mảng OOXML các đoạn giữa phương án và Lời giải.
 * Trả về [{ q (chỉ số đoạn), msg, warn? }] — warn: chỉ cảnh báo (vd. tiêu đề/hướng dẫn trước câu đầu).
 */
export function findErrors(P, plan, optXml, gapXml) {
  const errs = [];
  if (!plan.length) return [{ q: 0, msg: 'Không tìm thấy câu hỏi nào dạng "Câu N."' }];
  const before = P.slice(0, plan[0].q).filter((p) => !p.inTable && p.text.trim() !== "");
  if (before.length) errs.push({ q: 0, warn: true, msg: `Có ${before.length} đoạn ngoài câu hỏi ở đầu văn bản (trước Câu ${plan[0].num}) — kiểm tra lại nếu không phải tiêu đề/hướng dẫn.` });

  const seen = new Map();
  plan.forEach((b, k) => {
    const dupNum = seen.has(b.num);
    seen.set(b.num, (seen.get(b.num) || 0) + 1);
    if (dupNum) errs.push({ q: b.q, msg: `Trùng số câu: "Câu ${b.num}" xuất hiện ${seen.get(b.num)} lần.` });
    else if (k > 0 && b.num !== plan[k - 1].num + 1)
      errs.push({ q: b.q, msg: `Số câu không liên tục: Câu ${plan[k - 1].num} → Câu ${b.num}.` });

    const end = b.lg >= 0 ? b.lg : k + 1 < plan.length ? plan[k + 1].q : P.length;
    const text = P.slice(b.q + 1, end).filter((p) => !p.inTable).map((p) => p.text).join("\n");
    const counts = "ABCD".split("").map((L) => (text.match(reLabel(L)) || []).length);
    const missing = "ABCD".split("").filter((L, i) => counts[i] === 0);
    const dup = "ABCD".split("").filter((L, i) => counts[i] > 1);
    if (missing.length) errs.push({ q: b.q, msg: `Câu ${b.num}: thiếu phương án ${missing.join(", ")}.` });
    if (dup.length) errs.push({ q: b.q, msg: `Câu ${b.num}: trùng nhãn phương án ${dup.join(", ")}.` });
    if (!missing.length && !dup.length) {
      const pos = "ABCD".split("").map((L) => text.search(reLabel(L)));
      if (!(pos[0] < pos[1] && pos[1] < pos[2] && pos[2] < pos[3])) errs.push({ q: b.q, msg: `Câu ${b.num}: thứ tự phương án bị sai.` });
    }

    const marked = optXml[k] ? (rebuildOptionsXml(optXml[k]) || { correct: [] }).correct : [];
    if (!b.chonLetter && !marked.length) errs.push({ q: b.q, msg: `Câu ${b.num}: thiếu đáp án (không gạch chân, không có "Chọn X").` });
    if (b.chonLetter && marked.length && marked[0] !== b.chonLetter)
      errs.push({ q: b.q, msg: `Câu ${b.num}: gạch chân ${marked[0]} nhưng lời giải ghi "Chọn ${b.chonLetter}".` });

    if ((gapXml[k] || []).some(hasObject))
      errs.push({ q: b.q, msg: `Câu ${b.num}: có hình nằm giữa phương án và "Lời giải" (không thuộc lời dẫn, phương án hay lời giải).` });
  });
  return errs;
}

export async function checkLoiDe(context) {
  const ps = context.document.body.paragraphs;
  ps.load("items/text,items/isListItem,items/tableNestingLevel");
  await context.sync();
  const lis = ps.items.map((p) => (p.isListItem ? p.listItemOrNullObject : null));
  lis.forEach((l) => l && l.load("listString,isNullObject"));
  await context.sync();
  const P = ps.items.map((p, i) => ({ text: p.text, inTable: p.tableNestingLevel > 0, list: lis[i] && !lis[i].isNullObject ? lis[i].listString : "" }));
  const plan = planBlocks(P);

  const optX = plan.map((b) => (b.optFrom >= 0 ? ps.items[b.optFrom].getRange("Whole").expandTo(ps.items[b.optTo].getRange("Whole")).getOoxml() : null));
  const gapX = plan.map((b) => {
    if (b.optTo < 0 || b.lg < 0) return [];
    const out = [];
    for (let i = b.optTo + 1; i < b.lg; i++) if (!P[i].inTable) out.push(ps.items[i].getOoxml());
    return out;
  });
  await context.sync();
  const all = findErrors(P, plan, optX.map((x) => (x ? x.value : "")), gapX.map((a) => a.map((x) => x.value)));

  const nTableParas = P.filter((p) => p.inTable).length;
  const tableNote = nTableParas ? `ℹ Đề có bảng (${nTableParas} ô) — Chuẩn hoá BigKey giữ nguyên bảng.` : "";
  if (!errs.length) return [`Không phát hiện lỗi trong ${plan.length} câu.`, tableNote].filter(Boolean).join("\n");

  ps.items[errs[0].q].select(); // đưa con trỏ tới câu lỗi đầu tiên
  await context.sync();
  return {
    ok: false,
    summary: [`Phát hiện ${errs.length} lỗi trong ${plan.length} câu (đã chọn câu lỗi đầu tiên):`, ...errs.slice(0, 6).map((e) => "• " + e.msg), errs.length > 6 ? `… và ${errs.length - 6} lỗi khác (xem Chi tiết).` : "", ...warns, tableNote].filter(Boolean).join("\n"),
    details: errs.map((e) => e.msg).join("\n"),
  };
}
