// Chuẩn hoá BigKey  (spec: docs/specs/bigkey/Chuan_Hoa_BigKey.md)
// Port + viết lại phần "Chuẩn hóa" của BigKeyUI.dotm (bkAutoTab / bkChuanHoaBoiDen):
//   1. Đánh số tự động -> chữ; bố cục đoạn như bản VBA (normalizeXml.js): style Normal, xoá định dạng đoạn,
//      lề trái 1 cm, giãn dòng đơn, không tab stop / keep with next, Times New Roman 12 (giữ đậm/nghiêng…)
//      (cả tài liệu: trang A4 dọc, lề 1,5 cm — chỉ Word desktop)
//   2. Dàn phương án A–D 1/2/4 dòng theo độ rộng (optionLayout.js), nhãn đậm xanh, đáp án đúng đỏ gạch chân
//   3. "Câu N." -> "Câu i:" đậm xanh, thụt lề treo 1 cm; "Lời giải" đậm xanh; dòng "Chọn X" highlight vàng
//   4. Xoá dòng trống, gộp khoảng trắng kép
// Khác bản VBA: KHÔNG đổi bảng thành chữ (bản VBA crash khi đề có bảng) — đoạn trong bảng được giữ nguyên.

import { rebuildOptionsXml } from "./optionLayout.js";
import { normalizeParagraphsXml, bodyHasObject } from "./normalizeXml.js";
import { isDesktop, askUser } from "../runtime.js";

const CM = 28.35; // pt
const CHUNK = 40; // số đoạn mỗi lần ghi OOXML
const BLUE = "#0000FF";
const reQ = /^\s*Câu\s*(\d+)\s*[.:]/i;
const reLG = /^\s*(L[ờơo]i\s*gi[ảa]i|Hướng\s*dẫn\s*giải)/i;
const reChon = /^\s*Chọn\s*[\t ]*([A-D])\b/;
const reA = /^\s*A\s*[.)]/;
const reD = /(^|[\t ])D\s*[.)]/;

/**
 * Lập kế hoạch từ danh sách đoạn {text, inTable, list}. Thuần — test được trên dữ liệu thật.
 * Trả về các câu: { q, num, optFrom, optTo, lg, chon, chonLetter }.
 */
export function planBlocks(P) {
  const full = (p) => ((p.list ? p.list + " " : "") + (p.text || "")).normalize("NFC");
  const isQ = (i) => !P[i].inTable && reQ.test(full(P[i]));
  const qs = [];
  for (let i = 0; i < P.length; i++) if (isQ(i)) qs.push(i);
  return qs.map((q, k) => {
    const end = k + 1 < qs.length ? qs[k + 1] : P.length;
    const item = { q, num: +full(P[q]).match(reQ)[1], optFrom: -1, optTo: -1, lg: -1, chon: -1, chonLetter: "" };
    for (let i = q + 1; i < end; i++) {
      if (P[i].inTable) continue;
      const t = full(P[i]);
      if (item.lg < 0 && reLG.test(t)) item.lg = i;
      if (item.lg < 0 && item.optFrom < 0 && reA.test(t)) {
        // tối đa 4 đoạn có chữ (đoạn trống/chỉ có hình không tính — bỏ sót câu khi có dòng trống chen giữa)
        for (let j = i, seen = 0; j < end && seen < 4; j++) {
          if (P[j].inTable || reLG.test(full(P[j]))) break;
          if (!P[j].text.trim()) continue;
          seen++;
          if (reD.test(full(P[j]))) {
            item.optFrom = i;
            item.optTo = j;
            break;
          }
        }
      }
      const c = t.match(reChon);
      if (c && item.chon < 0) {
        item.chon = i;
        item.chonLetter = c[1];
      }
    }
    return item;
  });
}

/** Đánh số tự động -> chữ (listString + " "), bỏ khỏi danh sách. withTables: cả đoạn trong bảng. Trả về số đoạn. */
export async function listToText(context, P, withTables = false) {
  let n = 0;
  for (const x of P) {
    if ((x.inTable && !withTables) || !x.list) continue;
    x.p.getRange("Start").insertText(x.list + " ", "Before");
    x.p.detachFromList();
    n++;
  }
  if (n) await context.sync();
  return n;
}

/**
 * Xoá đoạn trống ngoài bảng (không chứa hình/công thức/đối tượng/bảng), trừ đoạn cuối tài liệu.
 * Gọi ngay sau mỗi bước insertOoxml: đoạn sentinel còn lại nằm giữa các phương án làm bước sau bỏ sót câu.
 */
async function dropEmpty(context, scope, useSelection) {
  const P = await loadParas(context, scope);
  const empties = P.filter((x) => !x.inTable && x.text.trim() === "");
  if (!empties.length) return 0;
  const eox = empties.map((x) => x.p.getOoxml());
  await context.sync();
  let n = 0;
  empties.forEach((x, k) => {
    if (bodyHasObject(eox[k].value)) return;
    if (P.indexOf(x) === P.length - 1 && !useSelection) return; // đoạn cuối tài liệu không xoá được
    x.p.delete();
    n++;
  });
  if (n) await context.sync();
  return n;
}

export async function loadParas(context, scopeRange) {
  const ps = scopeRange ? scopeRange.paragraphs : context.document.body.paragraphs;
  ps.load("items/text,items/isListItem,items/tableNestingLevel,items/leftIndent");
  await context.sync();
  const lis = ps.items.map((p) => {
    if (!p.isListItem) return null;
    const l = p.listItemOrNullObject;
    l.load("listString,isNullObject");
    return l;
  });
  await context.sync();
  return ps.items.map((p, i) => ({ p, text: p.text, inTable: p.tableNestingLevel > 0, left: p.leftIndent, list: lis[i] && !lis[i].isNullObject ? lis[i].listString : "" }));
}

async function chuanHoa(context, useSelection) {
  const lines = [];
  const desktop = isDesktop();
  let scope = null;
  if (useSelection) {
    scope = context.document.getSelection();
    scope.load("isEmpty");
    await context.sync();
    if (scope.isEmpty) return { ok: false, summary: "Chưa bôi đen câu hỏi muốn chuẩn hoá.", details: "" };
  }

  // 1) đánh số tự động -> chữ (ngoài bảng)
  let P = await loadParas(context, scope);
  const nList = await listToText(context, P);

  // 2) bố cục đoạn như bản VBA (normalizeXml.js): OOXML theo cụm đoạn liền nhau ngoài bảng, từ cuối lên đầu.
  //    Cụm không ghi được OOXML -> đặt bằng Office JS (thiếu: xoá tab stop, keep with next).
  P = await loadParas(context, scope);
  // Cụm chỉ cắt ở đầu một câu hỏi (không cắt đôi khối phương án), tối đa ~CHUNK đoạn.
  // Đoạn cuối cùng KHÔNG đi qua OOXML: Word không thay được dấu ¶ cuối tài liệu (bản beta: câu cuối phải chạy 2 lần).
  const isQ = (i) => reQ.test(((P[i].list ? P[i].list + " " : "") + P[i].text).normalize("NFC"));
  const chunks = [];
  for (let i = 0; i < P.length - 1; i++) {
    if (P[i].inTable) continue;
    const last = chunks[chunks.length - 1];
    const full = last && i - last[0] >= CHUNK && (isQ(i) || i - last[0] >= 3 * CHUNK);
    if (last && last[1] === i - 1 && !full) last[1] = i;
    else chunks.push([i, i]);
  }
  const direct = (p) => {
    p.styleBuiltIn = "Normal";
    p.font.name = "Times New Roman";
    p.font.size = 12;
    p.alignment = "Left";
    p.lineSpacing = 12; // Office JS: điểm, UI chia 12 -> giãn dòng đơn
    p.leftIndent = CM;
    p.firstLineIndent = 0;
    p.rightIndent = 0;
    p.spaceBefore = 0;
    p.spaceAfter = 0;
  };
  const cRanges = chunks.map(([a, b]) => P[a].p.getRange("Whole").expandTo(P[b].p.getRange("Whole")));
  const cXml = cRanges.map((r) => r.getOoxml());
  for (const x of P) if (x.inTable) (x.p.font.name = "Times New Roman"), (x.p.font.size = 12); // bảng: chỉ font
  const tail = P[P.length - 1];
  if (tail && !tail.inTable) direct(tail.p);
  await context.sync();
  let nFallback = 0;
  for (let k = chunks.length - 1; k >= 0; k--) {
    try {
      cRanges[k].insertOoxml(normalizeParagraphsXml(cXml[k].value), "Replace");
      await context.sync();
    } catch (e) {
      console.warn("normalize chunk failed, fallback", chunks[k], e);
      nFallback++;
      for (let i = chunks[k][0]; i <= chunks[k][1]; i++) direct(P[i].p);
      await context.sync();
    }
  }
  if (nFallback) lines.push(`⚠ ${nFallback} cụm đoạn định dạng bằng cách dự phòng (có thể còn tab stop / keep with next cũ).`);
  const nEmpty2 = await dropEmpty(context, scope, useSelection); // sentinel của các cụm
  if (!useSelection) {
    if (desktop) {
      const secs = context.document.sections;
      secs.load("items");
      await context.sync();
      for (const s of secs.items) {
        const ps = s.pageSetup;
        ps.orientation = "Portrait";
        ps.pageWidth = 21 * CM;
        ps.pageHeight = 29.7 * CM;
        ps.topMargin = ps.bottomMargin = ps.leftMargin = ps.rightMargin = 1.5 * CM;
        ps.headerDistance = ps.footerDistance = 1.27 * CM;
        ps.gutter = 0;
      }
      await context.sync();
      // Không bắt buộc (WordApiDesktop 1.4, một số bản Word báo GeneralException): mỗi mục 1 lần sync riêng,
      // lỗi thì bỏ qua — không được làm hỏng cả lệnh.
      const optional = [
        () => secs.items.forEach((s) => (s.pageSetup.lineNumbering.isActive = false)),
        () => (context.document.defaultTabStop = 1.5 * CM), // như bản VBA
      ];
      for (const set of optional) {
        try {
          set();
          await context.sync();
        } catch (e) {
          console.warn("optional page setting skipped", e);
        }
      }
    } else lines.push("⚠ Word Web: bỏ qua khổ giấy/lề trang (chỉ Word desktop).");
  }
  await context.sync();

  // 3) dàn phương án (OOXML), từ cuối lên đầu — đọc lại: bước 2 đã thay các đoạn
  P = await loadParas(context, scope);
  let plan = planBlocks(P);
  const blocks = plan.filter((b) => b.optFrom >= 0);
  const ranges = blocks.map((b) => P[b.optFrom].p.getRange("Whole").expandTo(P[b.optTo].p.getRange("Whole")));
  const oxs = ranges.map((r) => r.getOoxml());
  if (ranges.length) await context.sync();
  let laid = 0;
  const skipped = [];
  for (let k = blocks.length - 1; k >= 0; k--) {
    const res = rebuildOptionsXml(oxs[k].value, blocks[k].chonLetter);
    if (!res) {
      skipped.push(blocks[k].num);
      continue;
    }
    ranges[k].insertOoxml(res.xml, "Replace");
    laid++;
  }
  if (laid) await context.sync();
  const nEmpty3 = laid ? await dropEmpty(context, scope, useSelection) : 0; // sentinel của các khối phương án

  // 4) đánh số câu, Lời giải, Chọn X (đọc lại sau khi dàn phương án)
  P = await loadParas(context, scope);
  plan = planBlocks(P);
  // Như bản VBA: cả đề đánh lại từ 1. Đề bắt đầu từ số khác (vd. Câu 31–60 có lời dẫn "Questions 31-35")
  // -> hỏi số bắt đầu (mặc định 1). Vùng chọn: tiếp tục từ số của câu đầu vùng chọn.
  let start = useSelection && plan.length ? plan[0].num : 1;
  if (!useSelection && plan.length && plan[0].num !== 1) {
    const ans = await askUser({
      title: "Chuẩn hoá BigKey",
      lines: [`Đề bắt đầu từ Câu ${plan[0].num} (${plan.length} câu).`, "Đánh số lại liên tục, bắt đầu từ số:"],
      buttons: ["Đánh số", `Giữ từ ${plan[0].num}`],
      input: { label: "Bắt đầu từ", value: 1 },
    });
    // runtime cũ (chưa khởi động lại Word) trả về chuỗi: không có ô nhập
    const button = typeof ans === "string" ? ans : ans.button;
    const v = typeof ans === "string" ? 1 : Number(ans.value);
    start = button === "Đánh số" && Number.isInteger(v) && v >= 1 && v <= 9999 ? v : plan[0].num;
  }
  const lbl = plan.map((b) => {
    const m = P[b.q].text.match(/^\s*Câu\s*\d+\s*[.:]/i);
    return m ? P[b.q].p.search(m[0].trim(), { matchCase: false }) : null;
  });
  const lgs = plan.map((b) => (b.lg >= 0 ? P[b.lg].p.search(P[b.lg].text.normalize("NFC").match(reLG)[1], { matchCase: false }) : null));
  [...lbl, ...lgs].forEach((s) => s && s.load("items"));
  await context.sync();
  plan.forEach((b, k) => {
    const p = P[b.q].p;
    p.leftIndent = CM;
    p.firstLineIndent = -CM;
    const hit = lbl[k] && lbl[k].items[0];
    if (hit) {
      const r = hit.insertText(`Câu ${start + k}:`, "Replace");
      r.font.bold = true;
      r.font.color = BLUE;
      r.font.underline = "None";
    }
    const lg = lgs[k] && lgs[k].items[0];
    if (lg) {
      lg.font.bold = true;
      lg.font.color = BLUE;
    }
    if (b.chon >= 0) {
      const c = P[b.chon];
      if (/^\s*Chọn[\t ]+[A-D]\.?\s*$/.test(c.text)) c.p.getRange("Content").insertText(`Chọn ${b.chonLetter}`, "Replace");
    }
  });
  await context.sync();
  plan.forEach((b) => {
    if (b.chon < 0) return;
    const c = P[b.chon].p;
    c.getRange("Content").font.highlightColor = "Yellow";
  });
  await context.sync();

  // 5) dọn: dòng trống (không chứa hình/công thức/đối tượng), khoảng trắng kép
  const nEmpty = nEmpty2 + nEmpty3 + (await dropEmpty(context, scope, useSelection));
  const target = scope || context.document.body;
  for (let pass = 0; pass < 5; pass++) {
    const dbl = target.search("  ", { matchCase: true });
    dbl.load("items");
    await context.sync();
    if (!dbl.items.length) break;
    dbl.items.forEach((r) => r.insertText(" ", "Replace"));
    await context.sync();
  }

  P = await loadParas(context, scope);
  const nTables = P.filter((x) => x.inTable).length;
  lines.unshift(
    `Chuẩn hoá ${plan.length} câu${useSelection ? " (vùng chọn)" : ""}.`,
    `Dàn phương án: ${laid} câu.` + (skipped.length ? ` Bỏ qua ${skipped.length} câu (không đủ A–D liền nhau).` : "")
  );
  if (nTables) lines.push(`Giữ nguyên ${nTables} đoạn trong bảng.`);
  return {
    ok: true,
    summary: lines.join("\n"),
    details: [
      `Đánh số tự động -> chữ: ${nList} đoạn; xoá ${nEmpty} dòng trống.`,
      skipped.length ? `Câu không dàn được phương án: ${skipped.join(", ")}` : "",
    ].filter(Boolean).join("\n"),
  };
}

export const chuanHoaBigKeyDe = (context) => chuanHoa(context, false);
export const chuanHoaBigKeyBoiDen = (context) => chuanHoa(context, true);
