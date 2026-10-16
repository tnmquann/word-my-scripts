// Công cụ hỗ trợ của nhóm Chuẩn hoá BigKey  (spec: docs/specs/bigkey/BigKey_Tools.md)
// Port các menu con của "Chuẩn hóa" trong BigKeyUI.dotm (bản 2020 + 2024): ID, Lời giải, Số thứ tự, Hình ảnh, Câu 0.
// Bản VBA dùng Find/Replace với ký tự tạm (#, ~) và lưu trạng thái Ẩn/Hiện trong %APPDATA%\BigKey\showhide.dat;
// bản này tìm theo cấu trúc đoạn, trạng thái Ẩn/Hiện suy ra từ tài liệu.

import { askUser, isDesktop } from "../runtime.js";
import { loadParas, planBlocks, listToText } from "./chuanHoaBigKey.js";
import { findBlocks, questionBlocks } from "../xoa/xoaCanThiet.js";
import { bodyHasObject } from "./normalizeXml.js";
import { markAnswerXml, readAnswersXml, splitOptionsXml, optionPeriodsXml, highlightToRedXml } from "./optionTools.js";

const BLUE = "#0000FF";
const PINK = "#FF00FF"; // wdColorPink
const GREEN = "#008000"; // wdGreen (ColorIndex)
const reLG = /^\s*(L[ờơo]i\s*gi[ảa]i|Hướng\s*dẫn\s*giải)/i;
// Mã ID kiểu BigKey: [2D1-3], [0H2-1.3-2], [1]…[15] — không có khoảng trắng/dấu ; (tránh đoạn [0;1] trong Toán)
const reID = /\[[0-9A-Za-z][0-9A-Za-z.\-]{0,30}\]/g;
const wq = (s) => s.replace(/\^/g, "^^");

/** Các mã ID khác nhau trong văn bản. Thuần. */
export const idTagsBK = (text) => [...new Set(text.match(reID) || [])];

async function searchAll(context, tags) {
  const body = context.document.body;
  const hits = tags.map((t) => {
    const s = body.search(wq(t), { matchCase: true });
    s.load("items");
    return s;
  });
  await context.sync();
  return hits.reduce((a, s) => a.concat(s.items), []);
}

async function bodyTags(context) {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  return idTagsBK(body.text);
}

// ---------------------------------------------------------------- ID

export async function toMauID(context) {
  const items = await searchAll(context, await bodyTags(context));
  if (!items.length) return "Không có mã ID nào (dạng [2D1-3], [1]…).";
  items.forEach((r) => {
    r.font.bold = true;
    r.font.color = PINK;
  });
  await context.sync();
  return `Đã tô hồng ${items.length} mã ID.`;
}

export async function xoaIDBK(context) {
  const items = await searchAll(context, await bodyTags(context));
  if (!items.length) return "Không có mã ID nào (dạng [2D1-3], [1]…).";
  for (let k = items.length - 1; k >= 0; k--) items[k].delete();
  await context.sync();
  await collapseSpaces(context);
  return `Đã xoá ${items.length} mã ID.`;
}

/** Xoá nguồn tác giả: cụm "(…)" màu xanh lá (BigKey "Gắn Nguồn" tô xanh lá) — cụm ngoặc khác giữ nguyên. */
export async function xoaNguonBK(context) {
  const s = context.document.body.search("\\(*\\)", { matchWildcards: true });
  s.load("items/text,items/font/color");
  await context.sync();
  const green = s.items.filter((r) => (r.font.color || "").toUpperCase() === GREEN && r.text.length <= 200);
  if (!green.length) return 'Không có nguồn tác giả (cụm "(…)" màu xanh lá) nào.';
  for (let k = green.length - 1; k >= 0; k--) green[k].delete();
  await context.sync();
  await collapseSpaces(context);
  return `Đã xoá ${green.length} nguồn tác giả.`;
}

/** Ẩn mã ID còn hiện; nếu không còn mã nào hiện -> hiện lại toàn bộ (như bản VBA: Word không tìm được chữ đang ẩn). */
export async function anHienID(context) {
  const items = await searchAll(context, await bodyTags(context));
  if (items.length) {
    items.forEach((r) => (r.font.hidden = true));
    await context.sync();
    return `Đã ẩn ${items.length} mã ID. Bấm lại để hiện.`;
  }
  context.document.body.font.hidden = false;
  await context.sync();
  return "Đã hiện lại toàn bộ chữ ẩn (mã ID).";
}

/**
 * Gắn ID mức độ / nhóm (task pane, thay form Bkf_GanIDNhom): ghi "[k]" ngay sau "Câu N:" của câu tại con trỏ
 * (thay mã [k] cũ nếu có), rồi chọn nhãn câu kế tiếp để bấm tiếp. tag: "[1]".."[15]".
 */
export async function ganIDNhom(context, tag) {
  if (!/^\[\d{1,2}\]$/.test(tag)) throw new Error("Mã không hợp lệ: " + tag);
  const P = await loadParas(context);
  const plan = planBlocks(P).filter((b) => !P[b.q].list);
  if (!plan.length) return { ok: false, summary: 'Không có câu hỏi "Câu N:" gõ tay (hãy Chuẩn hoá BigKey trước).', details: "" };
  const sel = context.document.getSelection().getRange("Start");
  const cmp = plan.map((b) => P[b.q].p.getRange("Start").compareLocationWith(sel));
  await context.sync();
  let k = 0;
  cmp.forEach((c, i) => ["Before", "Equal", "AdjacentBefore"].includes(c.value) && (k = i));
  const label = (b) => P[b.q].p.search(P[b.q].text.match(/^\s*Câu\s*\d+\s*[.:]/i)[0].trim(), { matchCase: false });
  const cur = label(plan[k]);
  const old = /^\s*Câu\s*\d+\s*[.:]\s*(\[\d{1,2}\])/i.exec(P[plan[k].q].text);
  const oldHit = old ? P[plan[k].q].p.search(old[1], { matchCase: true }) : null;
  const next = k + 1 < plan.length ? label(plan[k + 1]) : null;
  [cur, oldHit, next].forEach((s) => s && s.load("items"));
  await context.sync();
  if (oldHit && oldHit.items[0]) oldHit.items[0].insertText(tag, "Replace");
  else {
    const r = cur.items[0].insertText(" " + tag, "After");
    r.font.bold = false;
    r.font.color = "#000000";
  }
  if (next && next.items[0]) next.items[0].select();
  await context.sync();
  return `Câu ${plan[k].num}: ${old ? old[1] + " → " : ""}${tag}` + (next ? ` · tiếp: Câu ${plan[k + 1].num}` : " · hết đề");
}

// ---------------------------------------------------------------- Đáp án

/** Đáp án từng câu: chữ cái phương án đang gạch chân, nếu không có thì theo "Chọn X". "" nếu không rõ. */
async function readAnswers(context, P, plan) {
  const ox = plan.map((b) => (b.optFrom >= 0 ? P[b.optFrom].p.getRange("Whole").expandTo(P[b.optTo].p.getRange("Whole")).getOoxml() : null));
  await context.sync();
  return plan.map((b, k) => {
    const r = ox[k] && readAnswersXml(ox[k].value);
    return (r && r.underlined) || b.chonLetter || "";
  });
}

/** Tạo bảng đáp án BigKey ở cuối đề (10 cột, hàng đầu "MÃ ĐỀ: 001"); đã có bảng đáp án -> ghi đè. */
export async function taoBangKeyBK(context) {
  const P = await loadParas(context);
  const plan = planBlocks(P).filter((b) => b.optFrom >= 0);
  if (!plan.length) return 'Không có câu trắc nghiệm A–D nào ("Câu N." + phương án).';
  const ans = await readAnswers(context, P, plan);
  const missing = plan.filter((b, k) => !ans[k]).map((b) => b.num);
  const cells = plan.map((b, k) => `${b.num}.${ans[k] || "?"}`);
  const rows = [];
  for (let i = 0; i < cells.length; i += 10) rows.push(cells.slice(i, i + 10).concat(Array(10).fill("")).slice(0, 10));

  const tables = context.document.body.tables;
  tables.load("items/values,items/rowCount");
  await context.sync();
  const key = keyTable(tables.items);
  if (key) {
    // ghi đè vùng đáp án (từ hàng 2), thêm hàng nếu thiếu
    if (key.rowCount - 1 < rows.length) key.addRows("End", rows.length - (key.rowCount - 1), rows.slice(key.rowCount - 1));
    rows.slice(0, key.rowCount - 1).forEach((r, i) => r.forEach((v, c) => (key.getCell(i + 1, c).value = v)));
    for (let i = rows.length + 1; i < key.rowCount; i++) for (let c = 0; c < 10; c++) key.getCell(i, c).value = "";
  } else {
    const t = context.document.body.insertTable(rows.length + 1, 10, "End", [["MÃ ĐỀ: 001", "", "", "", "", "", "", "", "", ""], ...rows]);
    t.font.name = "Times New Roman";
    t.font.size = 12;
    const head = t.rows.getFirst();
    head.shadingColor = "#008000";
    head.font.color = "#FFFFFF";
    head.font.bold = true;
    t.mergeCells(0, 0, 0, 9); // WordApi 1.4 (yêu cầu tối thiểu từ 2.1.0)
  }
  await context.sync();
  const msg = `${key ? "Đã cập nhật" : "Đã tạo"} bảng đáp án ${plan.length} câu.`;
  return missing.length
    ? { ok: false, summary: `${msg}\n⚠ Thiếu đáp án (ghi "?"): Câu ${missing.join(", ")}.`, details: "Đáp án lấy từ chữ cái phương án gạch chân, nếu không có thì từ dòng \"Chọn X\"." }
    : msg;
}

// ---------------------------------------------------------------- Phương án / gạch chân đáp án (optionTools.js)

/**
 * Áp fn(xml, b) lên OOXML từng khối phương án A–D (giữ bố cục), ghi lại từ cuối lên đầu, rồi xoá đoạn rỗng
 * sentinel còn lại ngay sau khối (chỉ khi không chứa hình/công thức). fn trả về { xml, letter } | null (bỏ qua).
 * Trả về { n, done: [{b, res}] }.
 */
async function transformOptions(context, fn) {
  const P = await loadParas(context);
  const plan = planBlocks(P).filter((b) => b.optFrom >= 0);
  const ranges = plan.map((b) => P[b.optFrom].p.getRange("Whole").expandTo(P[b.optTo].p.getRange("Whole")));
  const ox = ranges.map((r) => r.getOoxml());
  await context.sync();
  const done = [];
  for (let k = plan.length - 1; k >= 0; k--) {
    const res = fn(ox[k].value, plan[k]);
    if (!res) continue;
    ranges[k].insertOoxml(res.xml, "Replace");
    done.unshift({ b: plan[k], res });
  }
  if (done.length) {
    await context.sync();
    const P2 = await loadParas(context);
    const cand = planBlocks(P2)
      .filter((b) => b.optFrom >= 0 && P2[b.optTo + 1] && !P2[b.optTo + 1].inTable && P2[b.optTo + 1].text.trim() === "")
      .map((b) => P2[b.optTo + 1].p);
    const cx = cand.map((p) => p.getOoxml());
    await context.sync();
    cand.forEach((p, i) => !bodyHasObject(cx[i].value) && p.delete());
    await context.sync();
  }
  return { n: plan.length, done };
}

const none = (n) => (n ? "" : 'Không có câu trắc nghiệm A–D nào ("Câu N." + phương án).');

/** Chuẩn hoá màu PA-ĐA-LG: nhãn đậm xanh, đáp án (đang gạch chân) đỏ gạch chân, "A)"/"A:" -> "A.", "Lời giải" đậm xanh. */
export async function chuanHoaMauPA(context) {
  const { n, done } = await transformOptions(context, (x) => markAnswerXml(x, "keep"));
  if (!n) return none(n);
  const P = await loadParas(context);
  const lgs = planBlocks(P)
    .filter((b) => b.lg >= 0)
    .map((b) => P[b.lg].p.search(P[b.lg].text.normalize("NFC").match(reLG)[1], { matchCase: false }));
  lgs.forEach((s) => s.load("items"));
  await context.sync();
  lgs.forEach((s) => s.items[0] && ((s.items[0].font.bold = true), (s.items[0].font.color = BLUE)));
  await context.sync();
  return `Đã chuẩn hoá màu ${done.length}/${n} câu (${done.filter((d) => d.res.letter).length} câu có đáp án gạch chân).`;
}

/** Bỏ đánh dấu đáp án: mọi nhãn chỉ còn đậm xanh. */
export async function chuyenDAThanhPA(context) {
  const { n, done } = await transformOptions(context, (x) => markAnswerXml(x, "none"));
  return none(n) || `Đã bỏ đánh dấu đáp án ở ${done.length}/${n} câu.`;
}

export async function xuongDongPA(context) {
  const { n, done } = await transformOptions(context, splitOptionsXml);
  return none(n) || `Đã tách phương án xuống dòng ở ${done.length}/${n} câu.`;
}

export async function chamCuoiPA(context) {
  const { n, done } = await transformOptions(context, (x) => optionPeriodsXml(x, true));
  return none(n) || `Đã thêm dấu chấm cuối phương án ở ${done.length}/${n} câu.`;
}

export async function xoaChamCuoiPA(context) {
  const { n, done } = await transformOptions(context, (x) => optionPeriodsXml(x, false));
  return none(n) || `Đã xoá dấu chấm cuối phương án ở ${done.length}/${n} câu.`;
}

/** Gạch chân đáp án theo map num -> chữ cái; báo câu không có đáp án. */
async function gachChanTheo(context, answerOf, source) {
  const miss = [];
  const { n, done } = await transformOptions(context, (x, b) => {
    const a = answerOf(b);
    if (!a) return miss.push(b.num), null;
    return markAnswerXml(x, a);
  });
  if (!n) return none(n);
  const msg = `Đã gạch chân đáp án ${done.length}/${n} câu (theo ${source}).`;
  return miss.length ? { ok: false, summary: `${msg}\n⚠ Không có đáp án: Câu ${miss.join(", ")}.`, details: "" } : msg;
}

export async function gachChanTuBangKey(context) {
  const tables = context.document.body.tables;
  tables.load("items/values");
  await context.sync();
  const key = keyTable(tables.items);
  if (!key) return { ok: false, summary: 'Không tìm thấy bảng đáp án (bảng cuối 10 cột, ô dạng "1.A").', details: "" };
  const map = {};
  key.values.forEach((row) => row.forEach((v) => {
    const m = /^\s*(\d+)\s*\.\s*([A-D])\s*$/.exec(v);
    if (m) map[+m[1]] = m[2];
  }));
  return gachChanTheo(context, (b) => map[b.num], "bảng đáp án");
}

export const gachChanTuLoiGiai = (context) => gachChanTheo(context, (b) => b.chonLetter, '"Chọn X" trong lời giải');

export async function gachChanHighlightDo(context) {
  const { n, done } = await transformOptions(context, (x) => {
    const r = highlightToRedXml(x);
    return r && r.letter ? r : null;
  });
  return none(n) || `Đã gạch chân đáp án theo highlight ở ${done.length}/${n} câu (chữ tô sáng -> chữ đỏ).`;
}

// ---------------------------------------------------------------- Lời giải

/**
 * Chèn "Lời giải" cho câu chưa có: câu trắc nghiệm -> "Lời giải" + "Chọn X" (X = đáp án gạch chân, chưa có -> A);
 * câu tự luận -> "Lời giải" + dòng "Nội dung lời giải". writeAnswers: ghi cả đáp án vào câu đã có lời giải
 * (sửa "Chọn ?" hoặc thêm dòng "Chọn X" sau "Lời giải") — lệnh "Ghi đáp án vào lời giải".
 */
async function chenLG(context, writeAnswers) {
  const P = await loadParas(context);
  const plan = planBlocks(P);
  if (!plan.length) return 'Không tìm thấy câu hỏi "Câu N." nào.';
  const ends = questionBlocks(P); // cuối mỗi câu (dừng ở lời dẫn nhóm, PHẦN…)
  const ans = await readAnswers(context, P, plan);
  const addLines = (after, lines) => {
    let p = after;
    lines.forEach(([text, head]) => {
      p = p.insertParagraph(text, "After");
      p.firstLineIndent = 0;
      p.font.bold = !!head;
      p.font.italic = false;
      p.font.underline = "None";
      p.font.color = head ? BLUE : "#000000";
      p.font.highlightColor = null;
    });
  };
  let nNew = 0;
  let nAns = 0;
  plan.forEach((b, k) => {
    const mc = b.optFrom >= 0;
    if (b.lg < 0) {
      const end = (ends.find((e) => e.q === b.q) || { end: b.optTo }).end;
      addLines(P[mc ? b.optTo : end].p, [["Lời giải", true], mc ? [`Chọn ${ans[k] || "A"}`] : ["Nội dung lời giải"]]);
      nNew++;
    } else if (writeAnswers && mc && ans[k]) {
      if (b.chon >= 0) {
        if (b.chonLetter !== ans[k]) (P[b.chon].p.getRange("Content").insertText(`Chọn ${ans[k]}`, "Replace"), nAns++);
      } else (addLines(P[b.lg].p, [[`Chọn ${ans[k]}`]]), nAns++);
    }
  });
  if (!nNew && !nAns) return writeAnswers ? "Mọi câu đã có lời giải và đáp án khớp." : "Mọi câu đã có lời giải.";
  await context.sync();
  const miss = plan.filter((b, k) => b.optFrom >= 0 && b.lg < 0 && !ans[k]).map((b) => b.num);
  return `Chèn lời giải: ${nNew} câu.` + (writeAnswers ? ` Ghi đáp án: ${nAns} câu.` : "") + (miss.length ? `\n⚠ Chưa có đáp án, tạm ghi "Chọn A": Câu ${miss.join(", ")}.` : "");
}

export const chenLoiGiai = (context) => chenLG(context, false);
export const ghiDapAnVaoLoiGiai = (context) => chenLG(context, true);

export async function anHienLoiGiai(context) {
  const P = await loadParas(context);
  const blocks = findBlocks(P, reLG);
  if (blocks.length) {
    blocks.forEach(([a, b]) => (P[a].p.getRange("Whole").expandTo(P[b].p.getRange("Whole")).font.hidden = true));
    await context.sync();
    return `Đã ẩn ${blocks.length} lời giải. Bấm lại để hiện.`;
  }
  context.document.body.font.hidden = false;
  await context.sync();
  return "Đã hiện lại toàn bộ lời giải (chữ ẩn).";
}

// ---------------------------------------------------------------- Số thứ tự

/** "Câu N." gõ tay -> "Câu i:" đậm xanh, i từ start. Đánh lại cả bảng đáp án (10 cột "n.X") nếu có. */
/** fixedStart: số bắt đầu cho sẵn (task pane có ô nhập); không có -> hỏi bằng popup (lệnh ribbon). */
export async function danhSoThuTu(context, fixedStart) {
  const P = await loadParas(context);
  const plan = planBlocks(P);
  if (!plan.length) return 'Không tìm thấy câu hỏi "Câu N." nào.';
  let start = fixedStart;
  if (start === undefined) {
    const ans = await askUser({
      title: "Đánh số thứ tự",
      lines: [`Tìm thấy ${plan.length} câu (câu đầu hiện là số ${plan[0].num}).`, "Đánh số lại liên tục, bắt đầu từ số:"],
      buttons: ["Đánh số", "Huỷ"],
      input: { label: "Bắt đầu từ", value: 1 },
    });
    const button = typeof ans === "string" ? ans : ans.button;
    if (button !== "Đánh số") return "Đã huỷ, tài liệu không thay đổi.";
    start = typeof ans === "string" ? 1 : Number(ans.value);
  }
  if (!Number.isInteger(start) || start < 1 || start > 9999) return { ok: false, summary: `Số bắt đầu không hợp lệ: "${start}". Nhập số nguyên từ 1 đến 9999.`, details: "" };

  const hits = plan.map((b) => {
    if (P[b.q].list) return null; // đánh số tự động: Word tự đánh
    const m = P[b.q].text.match(/^\s*Câu\s*\d+\s*[.:]/i);
    if (!m) return null;
    const s = P[b.q].p.search(m[0].trim(), { matchCase: false });
    s.load("items");
    return s;
  });
  const tables = context.document.body.tables;
  tables.load("items/values,items/rowCount");
  await context.sync();
  hits.forEach((s, k) => {
    const r = s && s.items[0] && s.items[0].insertText(`Câu ${start + k}:`, "Replace");
    if (r) (r.font.bold = true), (r.font.color = BLUE), (r.font.underline = "None");
  });
  // bảng đáp án (BigKey): bảng cuối, 10 cột, ô "n.X"
  const key = keyTable(tables.items);
  let nKey = 0;
  if (key) {
    let n = start;
    key.values.forEach((row, ri) =>
      row.forEach((v, ci) => {
        const m = /^\s*\d+\s*\.\s*([A-D])\s*$/.exec(v);
        if (!m) return;
        key.getCell(ri, ci).value = `${n++}.${m[1]}`;
        nKey++;
      })
    );
  }
  await context.sync();
  return `Đã đánh số ${hits.filter(Boolean).length} câu từ ${start}.` + (nKey ? ` Bảng đáp án: ${nKey} ô.` : "");
}

/** Bảng đáp án kiểu BigKey: bảng cuối có 10 cột, hàng 2 ô 1 dạng "n.X". */
export function keyTable(tables) {
  const t = tables[tables.length - 1];
  if (!t || !t.values || t.values.length < 2 || t.values[1].length !== 10) return null;
  return /^\s*\d+\s*\.\s*[A-D]\s*$/.test(t.values[1][0]) ? t : null;
}

export async function sttAutoToText(context) {
  const n = await listToText(context, await loadParas(context), true);
  return n ? `Đã chuyển ${n} số tự động thành chữ.` : "Không có đánh số tự động nào.";
}

// ---------------------------------------------------------------- Hình ảnh

/** Hình nổi (Picture) -> nằm trên dòng chữ (Inline). Chỉ Word desktop. */
export async function pictureInline(context) {
  if (!isDesktop()) return { ok: false, summary: "Picture Inline chỉ chạy trên Word desktop (Windows/Mac).", details: "" };
  const shapes = context.document.body.shapes;
  shapes.load("items/type");
  await context.sync();
  const pics = shapes.items.filter((s) => s.type === "Picture");
  if (!pics.length) return "Không có hình nổi (floating) nào.";
  pics.forEach((s) => (s.textWrap.type = "Inline"));
  await context.sync();
  return `Đã chuyển ${pics.length} hình thành Inline with text.`;
}

/** Căn giữa đoạn chỉ chứa hình (không căn đoạn có chữ — bản VBA căn cả dòng chữ chứa hình công thức). */
export async function canGiuaHinh(context) {
  const pics = context.document.body.inlinePictures;
  pics.load("items");
  await context.sync();
  const paras = pics.items.map((p) => {
    const para = p.paragraph;
    para.load("text,tableNestingLevel");
    return para;
  });
  await context.sync();
  const targets = paras.filter((p) => p.text.trim() === "" && p.tableNestingLevel === 0);
  targets.forEach((p) => (p.alignment = "Centered"));
  await context.sync();
  return targets.length ? `Đã căn giữa ${targets.length} hình.` : "Không có đoạn nào chỉ chứa hình.";
}

// ---------------------------------------------------------------- Câu 0

/** Chèn "Câu 0: " ở đầu đoạn tại con trỏ (đánh dấu kết thúc đề của BigKey). */
export async function chenCau0(context) {
  const para = context.document.getSelection().paragraphs.getFirst();
  const r = para.getRange("Start").insertText("Câu 0: ", "Before");
  r.font.name = "Times New Roman";
  r.font.bold = false;
  r.font.color = "#000000";
  para.alignment = "Left";
  await context.sync();
  return 'Đã chèn "Câu 0: " tại con trỏ.';
}

export async function xoaCau0(context) {
  const s = context.document.body.search("Câu 0:", { matchCase: true });
  s.load("items");
  await context.sync();
  if (!s.items.length) return 'Không có "Câu 0:" nào.';
  for (let k = s.items.length - 1; k >= 0; k--) s.items[k].delete();
  await context.sync();
  return `Đã xoá ${s.items.length} "Câu 0:".`;
}

// ---------------------------------------------------------------- chung

async function collapseSpaces(context) {
  const body = context.document.body;
  for (let pass = 0; pass < 5; pass++) {
    const dbl = body.search("  ", { matchCase: true });
    dbl.load("items");
    await context.sync();
    if (!dbl.items.length) return;
    dbl.items.forEach((r) => r.insertText(" ", "Replace"));
    await context.sync();
  }
}
