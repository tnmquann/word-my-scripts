// Xoá cần thiết  (spec: docs/specs/xoa/Xoa_Can_Thiet.md)
// Port + viết lại nhóm "Xóa cần thiết" của NHOM-WORD-BIEN-SOAN.dotm (xoa_can_thiet.bas, form HO_tro "Xóa mức độ").
// Bản VBA dùng chuỗi Find/Replace với ký tự đánh dấu tạm (#, ~, #$) — hỏng nếu văn bản có sẵn các ký tự đó.
// Bản này làm theo cấu trúc đoạn: tìm khối -> xoá bảng trong khối -> xoá khối (xem dgnl/xoaDactDgnl.js).

import { blockRange, deleteFast, deleteSlow } from "../dgnl/xoaDactDgnl.js";
import { isDesktop } from "../runtime.js";

const reQ = /^\s*Câu\s*(\d+)\s*[.:]/i;
const reLG = /^\s*(L[ờơo]i\s*gi[ảa]i|Hướng\s*dẫn\s*giải)/i;
const reSuyLuan = /^\s*Phân\s*tích\s*(hướng\s*dẫn|lời)\s*giải/i;
// ranh giới khối: câu mới, chuyên đề "CĐ1…", "PHẦN …", "BẢNG ĐÁP ÁN", và lời dẫn nhóm câu
// ("Questions 46-52: Read the following passage…", "Đọc đoạn văn…", "…trả lời các câu hỏi từ X đến Y")
// — nếu không, xoá lời giải/câu cuối nhóm sẽ xoá luôn đoạn văn đọc hiểu của nhóm sau.
export const reBoundary = /^\s*(CĐ\s*\d+|PHẦN\s+[IVXLC\d]+|BẢNG\s*ĐÁP\s*ÁN|[-–—_.*\s]*HẾT[-–—_.!*\s]*$|Questions?\s*\d+\s*[-–—]|Read\s+the\s+following|Mark\s+the\s+letter|Đọc\s+(đoạn|bài|văn\s*bản|thông\s*tin)|Dựa\s+vào|Sử\s+dụng\s+(các\s+)?(thông\s*tin|dữ\s*kiện)|Dùng\s+(các\s+)?(thông\s*tin|dữ\s*kiện))|trả\s+lời\s+(các\s+)?câu(\s*hỏi)?\s*(từ\s*)?\d+\s*(đến|[-–—])/i;

const full = (p) => ((p.list ? p.list + " " : "") + (p.text || "")).normalize("NFC");
export const isQuestion = (p) => !p.inTable && reQ.test(full(p));

/**
 * Các khối [a,b] bắt đầu ở đoạn khớp startRe, kết thúc trước câu mới / ranh giới / endRe (nếu có), hoặc trước đoạn
 * thụt lề ÍT hơn đoạn bắt đầu > 5 pt (đoạn văn / lời dẫn sát lề — như Xoá ĐACT ĐGNL; cần P[i].left). Thuần.
 */
export function findBlocks(P, startRe, endRe) {
  const out = [];
  for (let i = 0; i < P.length; i++) {
    if (P[i].inTable || !startRe.test(full(P[i]))) continue;
    let j = i + 1;
    while (j < P.length) {
      if (!P[j].inTable) {
        const t = full(P[j]);
        if (isQuestion(P[j]) || reBoundary.test(t) || (endRe && endRe.test(t)) || startRe.test(t)) break;
        if (P[j].left != null && P[i].left != null && P[j].left < P[i].left - 5) break;
      }
      j++;
    }
    out.push([i, j - 1]);
    i = j - 1;
  }
  return out;
}

/** Các khối câu hỏi [q, cuối câu] (đến trước câu mới / ranh giới). Thuần. */
export function questionBlocks(P) {
  const qs = [];
  for (let i = 0; i < P.length; i++) if (isQuestion(P[i])) qs.push(i);
  return qs.map((q, k) => {
    let e = k + 1 < qs.length ? qs[k + 1] : P.length;
    for (let j = q + 1; j < e; j++) if (!P[j].inTable && reBoundary.test(full(P[j]))) { e = j; break; }
    return { q, end: e - 1, num: +full(P[q]).match(reQ)[1] };
  });
}

/** Cụm trích nguồn / mã ID đứng đầu câu: "[2D1-3] (Sở GD Hà Nội) …" -> chuỗi cần xoá (hoặc ""). Thuần. */
export function leadingSourceTags(textAfterLabel) {
  const m = textAfterLabel.match(/^[\s\t]*((?:\[[^\]\n]{0,200}\]|\([^)\n]{0,200}\))(?:[\s\t]*(?:\[[^\]\n]{0,200}\]|\([^)\n]{0,200}\)))*)/);
  return m ? m[1] : "";
}

/** Mức độ của câu từ mã ID "[2D1-3]" -> 3 (null nếu không có). Thuần. */
export const levelOf = (text) => {
  const m = text.match(/\[[^\]\n]*?-(\d)\]/);
  return m ? +m[1] : null;
};

/** Các mã ID dạng [2D1-3], [1H2-1], [D1-2], [H3-4]… trong văn bản. Thuần. */
export const idTags = (text) => text.match(/\[[0-2]?[DH][^\]\n]{0,20}\]/g) || [];

/** Dòng / cụm "Tác giả": { wholeParagraph } hoặc { inline: [chuỗi] }. Thuần. */
export function authorParts(text) {
  // không dùng \b cạnh chữ Việt (JS coi "ả" là ký tự ngoài từ)
  if (/^[\s\[(\-–—]*Tác\s*giả(?=$|[\s:.,;)\]\-–—])/i.test(text)) return { wholeParagraph: true, inline: [] };
  return { wholeParagraph: false, inline: text.match(/\[?[\[(]\s*Tác\s*giả[^\])\n]{0,150}[\])]\]?/gi) || [] };
}

// ---------------------------------------------------------------- Word helpers

async function loadParas(context) {
  const ps = context.document.body.paragraphs;
  ps.load("items/text,items/isListItem,items/tableNestingLevel,items/leftIndent");
  await context.sync();
  const lis = ps.items.map((p) => (p.isListItem ? p.listItemOrNullObject : null));
  lis.forEach((l) => l && l.load("listString,isNullObject"));
  await context.sync();
  return { it: ps.items, P: ps.items.map((p, i) => ({ text: p.text, inTable: p.tableNestingLevel > 0, left: p.leftIndent, list: lis[i] && !lis[i].isNullObject ? lis[i].listString : "" })) };
}

const wq = (s) => s.replace(/\^/g, "^^"); // ^ là ký tự đặc biệt của Word Search

/** Xoá các khối find(P); fast path, lỗi -> quét lại (fast path có thể đã xoá một phần) rồi xoá từng khối. */
async function deleteBlocks(context, find) {
  const { it, P } = await loadParas(context);
  const blocks = find(P);
  if (!blocks.length) return 0;
  try {
    await deleteFast(context, it, blocks);
    return blocks.length;
  } catch (e) {
    console.warn("deleteFast failed, fallback", e);
    const again = await loadParas(context);
    const { ok } = await deleteSlow(context, again.it, find(again.P));
    return ok;
  }
}

/** Đánh lại số "Câu N." liên tục từ số câu đầu tiên (chỉ câu gõ tay; câu đánh số tự động Word tự đánh). */
async function renumber(context) {
  const { it, P } = await loadParas(context);
  const qs = questionBlocks(P);
  if (!qs.length) return 0;
  const first = qs[0].num;
  const hits = qs.map((b, k) => {
    if (P[b.q].list) return null;
    const m = P[b.q].text.match(/^\s*Câu\s*\d+/i);
    if (!m || +m[0].match(/\d+/)[0] === first + k) return null;
    const s = it[b.q].search(m[0].trim(), { matchCase: false });
    s.load("items");
    return { s, k };
  }).filter(Boolean);
  if (!hits.length) return 0;
  await context.sync();
  for (const { s, k } of hits) if (s.items[0]) s.items[0].insertText(`Câu ${first + k}`, "Replace");
  await context.sync();
  return hits.length;
}

// ---------------------------------------------------------------- ribbon commands

export async function xoaTrichNguon(context) {
  const { it, P } = await loadParas(context);
  const jobs = [];
  P.forEach((p, i) => {
    if (!isQuestion(p)) return;
    const after = p.list ? p.text : p.text.replace(reQ, "");
    const tags = leadingSourceTags(after);
    if (!tags || tags.length > 250) return;
    const s = it[i].search(wq(tags.trim()), { matchCase: true });
    s.load("items");
    jobs.push(s);
  });
  if (!jobs.length) return "Không có trích nguồn / mã nào ở đầu câu hỏi.";
  await context.sync();
  let n = 0;
  for (const s of jobs) if (s.items[0]) (s.items[0].delete(), n++);
  await context.sync();
  return `Đã xoá trích nguồn ở ${n} câu.`;
}

export async function xoaLoiGiai(context) {
  const n = await deleteBlocks(context, (P) => findBlocks(P, reLG));
  if (!n) return 'Không tìm thấy "Lời giải" / "Hướng dẫn giải" nào.';
  return `Đã xoá ${n} khối lời giải.`;
}

export async function xoaTacGia(context) {
  const { it, P } = await loadParas(context);
  const whole = [];
  const inline = [];
  P.forEach((p, i) => {
    if (p.inTable) return;
    const a = authorParts(p.text);
    if (a.wholeParagraph) whole.push(i);
    else for (const s of a.inline) if (s.length <= 250) {
      const r = it[i].search(wq(s), { matchCase: false });
      r.load("items");
      inline.push(r);
    }
  });
  if (!whole.length && !inline.length) return 'Không tìm thấy "Tác giả" nào.';
  if (inline.length) await context.sync();
  let n = 0;
  for (const r of inline) for (const x of r.items) (x.delete(), n++);
  for (let k = whole.length - 1; k >= 0; k--) it[whole[k]].delete();
  await context.sync();
  return `Đã xoá ${whole.length} dòng và ${n} cụm "Tác giả".`;
}

// ---------------------------------------------------------------- task pane "Xoá mức độ…"

export async function xoaTheoMucDo(context, levels) {
  const n = await deleteBlocks(context, (P) => questionBlocks(P).filter((b) => levels.includes(levelOf(full(P[b.q])))).map((b) => [b.q, b.end]));
  if (!n) return `Không có câu nào thuộc mức ${levels.join(", ")} (mã dạng [..-k]).`;
  const r = await renumber(context);
  return `Đã xoá ${n} câu mức ${levels.join(", ")}.` + (r ? ` Đánh lại số ${r} câu.` : "");
}

export async function xoaID(context) {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  const tags = [...new Set(idTags(body.text))];
  if (!tags.length) return "Không có mã ID nào.";
  const hits = tags.map((t) => {
    const s = body.search(wq(t), { matchCase: true });
    s.load("items");
    return s;
  });
  await context.sync();
  let n = 0;
  hits.forEach((s) => s.items.forEach((x) => (x.delete(), n++)));
  await context.sync();
  return `Đã xoá ${n} mã ID.`;
}

export async function xoaHuongSuyLuan(context) {
  const n = await deleteBlocks(context, (P) => findBlocks(P, reSuyLuan, reLG));
  if (!n) return 'Không có "Phân tích hướng dẫn giải" nào.';
  return `Đã xoá ${n} phần hướng suy luận.`;
}

// 8 kiểu viền của form HO_tro (wdLineStyle 7, 2–6, 19, 20); mặc định "Đôi" như bản VBA
export const BORDER_STYLES = { "Đôi": "Double", "Chấm": "Dot", "Gạch ngắn": "DashSmallGap", "Gạch dài": "DashLargeGap", "Gạch chấm": "DashDot", "Gạch chấm chấm": "DashDotDot", "Sóng đôi": "DoubleWavy", "Gạch chấm nổi": "DashDotStroked" };

/** Khung quanh từng câu (phần đề, trước "Lời giải"). opts: { style, borderColor, fill } — fill "" = không tô. */
export async function taoKhung(context, opts) {
  if (!isDesktop()) return { ok: false, summary: "Tạo khung chỉ chạy trên Word desktop (Windows/Mac).", details: "" };
  const { it, P } = await loadParas(context);
  const qs = questionBlocks(P);
  const targets = [];
  for (const b of qs) {
    for (let i = b.q; i <= b.end; i++) {
      if (!P[i].inTable && i > b.q && reLG.test(full(P[i]))) break;
      if (!P[i].inTable) targets.push(it[i]);
    }
  }
  const bs = targets.map((p) => {
    const c = p.borders;
    c.load("items/inside");
    return c;
  });
  await context.sync();
  targets.forEach((p, k) => {
    // cùng thụt lề cho cả khối để Word gộp viền các đoạn liền nhau thành 1 khung (như bản VBA)
    p.leftIndent = 49.6;
    p.firstLineIndent = -49.6;
    for (const b of bs[k].items) {
      if (b.inside) b.lineStyle = "None";
      else if (opts.style === "None") b.lineStyle = "None";
      else {
        b.lineStyle = opts.style;
        b.color = opts.borderColor;
        b.lineWidth = "Pt075";
      }
    }
    if (opts.fill) p.shading.backgroundPatternColor = opts.fill;
    else p.shading.backgroundPatternColorIndex = "Auto"; // không tô
  });
  await context.sync();
  return `${opts.style === "None" ? "Đã huỷ khung" : "Đã tạo khung"} cho ${qs.length} câu.`;
}

export const huyKhung = (context) => taoKhung(context, { style: "None", fill: "" });

/** Thay lời giải bằng N dòng chấm ở mỗi câu (học sinh làm bài). */
export async function themDongCham(context, n) {
  const del = await deleteBlocks(context, (P) => findBlocks(P, reLG));
  const { it: it2, P: P2 } = await loadParas(context);
  const qs = questionBlocks(P2);
  const dots = ".".repeat(150);
  for (let k = qs.length - 1; k >= 0; k--) {
    let last = it2[qs[k].end];
    for (let j = 0; j < n; j++) {
      // câu kết thúc trong bảng -> chèn sau bảng, không chèn vào ô
      last = j === 0 && P2[qs[k].end].inTable ? last.parentTable.insertParagraph(dots, "After") : last.insertParagraph(dots, "After");
      last.font.color = "#808080";
      last.font.size = 12;
      last.font.bold = false;
      last.font.underline = "None";
      last.font.highlightColor = null;
      last.leftIndent = 49.6;
      last.firstLineIndent = 0;
      last.lineSpacing = 18; // ~1,5 dòng
      last.spaceBefore = 0;
      last.spaceAfter = 0;
    }
  }
  await context.sync();
  return `Đã thêm ${n} dòng chấm cho ${qs.length} câu` + (del ? ` (đã xoá ${del} lời giải).` : ".");
}

// ---------------------------------------------------------------- plugin BigKey "Xóa dòng viết LG"

const reOptD = /(^|[\t ])D\s*[.)]/;
const reDots = /^[\s.…_\-–]{8,}$/; // dòng chấm / gạch để học sinh viết

/**
 * Phần "viết lời giải" của mỗi câu: từ "Lời giải" (nếu có), nếu không thì từ đoạn ngay sau phương án D (trắc nghiệm)
 * hoặc từ dòng chấm đầu tiên (tự luận), đến hết câu. Dừng ở câu mới / lời dẫn nhóm / đoạn thụt lề ít hơn đề > 5 pt
 * (đoạn văn đọc hiểu sát lề). Thuần. Trả về các khối [a, b].
 */
export function writingBlocks(P) {
  const out = [];
  for (const b of questionBlocks(P)) {
    let start = -1;
    for (let i = b.q + 1; i <= b.end && start < 0; i++) if (!P[i].inTable && reLG.test(full(P[i]))) start = i;
    if (start < 0) {
      let d = -1;
      for (let i = b.q; i <= b.end; i++) if (!P[i].inTable && reOptD.test(full(P[i]))) d = i;
      if (d >= 0) start = d + 1;
      else for (let i = b.q + 1; i <= b.end && start < 0; i++) if (!P[i].inTable && reDots.test(P[i].text)) start = i;
    }
    if (start < 0 || start > b.end) continue;
    let end = b.end;
    const base = P[b.q].left;
    for (let i = start; i <= b.end; i++) if (!P[i].inTable && base != null && P[i].left != null && P[i].left < base - 5) { end = i - 1; break; }
    if (end >= start) out.push([start, end]);
  }
  return out;
}

/**
 * Xoá dòng viết LG (plugin BigKey dnqXoadongvietLG = Chèn lời giải + Xóa lời giải): xoá phần sau phương án / lời giải
 * / dòng chấm của mọi câu. Bản VBA chèn "Lời giải" giả rồi xoá bằng ký tự đánh dấu ~ — dễ xoá nhầm đoạn văn.
 */
export async function xoaDongVietLG(context) {
  const n = await deleteBlocks(context, writingBlocks);
  return n ? `Đã xoá phần viết lời giải ở ${n} câu.` : "Không có phần viết lời giải nào (sau phương án / Lời giải / dòng chấm).";
}
