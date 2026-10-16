// Chuẩn hoá BigKey  (spec: docs/specs/bigkey/Chuan_Hoa_BigKey.md)
// Port + viết lại phần "Chuẩn hóa" của BigKeyUI.dotm (bkAutoTab / bkChuanHoaBoiDen):
//   1. Đánh số tự động -> chữ; style Normal, Times New Roman 12, lề trái 1 cm, không giãn đoạn
//      (cả tài liệu: trang A4 dọc, lề 1,5 cm — chỉ Word desktop)
//   2. Dàn phương án A–D 1/2/4 dòng theo độ rộng (optionLayout.js), nhãn đậm xanh, đáp án đúng đỏ gạch chân
//   3. "Câu N." -> "Câu i:" đậm xanh, thụt lề treo 1 cm; "Lời giải" đậm xanh; dòng "Chọn X" highlight vàng
//   4. Xoá dòng trống, gộp khoảng trắng kép
// Khác bản VBA: KHÔNG đổi bảng thành chữ (bản VBA crash khi đề có bảng) — đoạn trong bảng được giữ nguyên.

import { rebuildOptionsXml } from "./optionLayout.js";
import { isDesktop } from "../runtime.js";

const CM = 28.35; // pt
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
        for (let j = i; j < Math.min(end, i + 4); j++) {
          if (P[j].inTable || reLG.test(full(P[j]))) break;
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

async function loadParas(context, scopeRange) {
  const ps = scopeRange ? scopeRange.paragraphs : context.document.body.paragraphs;
  ps.load("items/text,items/isListItem,items/tableNestingLevel");
  await context.sync();
  const lis = ps.items.map((p) => {
    if (!p.isListItem) return null;
    const l = p.listItemOrNullObject;
    l.load("listString,isNullObject");
    return l;
  });
  await context.sync();
  return ps.items.map((p, i) => ({ p, text: p.text, inTable: p.tableNestingLevel > 0, list: lis[i] && !lis[i].isNullObject ? lis[i].listString : "" }));
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
  let nList = 0;
  for (const x of P) {
    if (x.inTable || !x.list) continue;
    x.p.getRange("Start").insertText(x.list + " ", "Before");
    x.p.detachFromList();
    nList++;
  }
  if (nList) await context.sync();

  // 2) định dạng chung
  P = await loadParas(context, scope);
  for (const x of P) {
    if (x.inTable) continue;
    x.p.styleBuiltIn = "Normal";
    x.p.font.name = "Times New Roman";
    x.p.font.size = 12;
    x.p.leftIndent = CM;
    x.p.firstLineIndent = 0;
    x.p.rightIndent = 0;
    x.p.spaceBefore = 0;
    x.p.spaceAfter = 0;
  }
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
      }
    } else lines.push("⚠ Word Web: bỏ qua khổ giấy/lề trang (chỉ Word desktop).");
  }
  await context.sync();

  // 3) dàn phương án (OOXML), từ cuối lên đầu
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

  // 4) đánh số câu, Lời giải, Chọn X (đọc lại sau khi dàn phương án)
  P = await loadParas(context, scope);
  plan = planBlocks(P);
  // đánh số liên tục từ số của câu đầu (bản VBA luôn từ 1 — sai với đề tách phần, vd. Câu 31–60 có lời dẫn "Questions 31-35")
  const start = plan.length ? plan[0].num : 1;
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
    c.alignment = "Justified";
  });
  await context.sync();

  // 5) dọn: dòng trống (không chứa hình/công thức/đối tượng), khoảng trắng kép
  P = await loadParas(context, scope);
  const empties = P.filter((x) => !x.inTable && x.text.trim() === "");
  const eox = empties.map((x) => x.p.getOoxml());
  if (empties.length) await context.sync();
  let nEmpty = 0;
  empties.forEach((x, k) => {
    if (/<w:drawing|<w:pict|<w:object|<m:oMath|<w:tbl/.test(eox[k].value)) return;
    if (P.indexOf(x) === P.length - 1 && !useSelection) return; // đoạn cuối tài liệu không xoá được
    x.p.delete();
    nEmpty++;
  });
  if (nEmpty) await context.sync();
  const target = scope || context.document.body;
  for (let pass = 0; pass < 5; pass++) {
    const dbl = target.search("  ", { matchCase: true });
    dbl.load("items");
    await context.sync();
    if (!dbl.items.length) break;
    dbl.items.forEach((r) => r.insertText(" ", "Replace"));
    await context.sync();
  }

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
