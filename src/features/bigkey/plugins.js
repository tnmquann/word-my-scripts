// Plugin của BigKey (D:\BIGKEY\Apps\Bigkey\Plugins): Tách LG, Chemistry Formatter, Tô phiếu ZG50 từ ĐA.
// ("Xóa dòng viết LG" ở xoa/xoaCanThiet.js.) Spec: docs/specs/bigkey/BigKey_Tools.md
// Phần thuần (OOXML / chuỗi) test được ngoài Word; phần Office JS chỉ đọc/ghi.

import JSZip from "jszip";
import { parseXml, documentBody } from "../dgnl/removeBreakEng.js";
import { W, putInRPr, sentinelParagraph } from "./normalizeXml.js";
import { reBoundary, xoaLoiGiai } from "../xoa/xoaCanThiet.js";
import { loadParas, planBlocks } from "./chuanHoaBigKey.js";
import { readAnswersXml } from "./optionTools.js";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const reQ = /^\s*(Câu\s*\d+\s*[.:])/i;
const reLG = /^\s*(L[ờơo]i\s*gi[ảa]i|Hướng\s*dẫn\s*giải)\s*[:.]?\s*$/i;

const el = (doc, name, attrs = {}) => {
  const e = doc.createElementNS(W, "w:" + name);
  for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, "w:" + k, v);
  return e;
};
const kids = (e) => [...e.childNodes].filter((n) => n.nodeType === 1);
const pText = (p) => [...p.getElementsByTagNameNS(W, "t")].map((t) => t.textContent).join("");
function textRun(doc, text, props = {}) {
  const r = el(doc, "r");
  const rPr = el(doc, "rPr");
  if (props.b) putInRPr(rPr, el(doc, "b"));
  if (props.color) putInRPr(rPr, el(doc, "color", { val: props.color }));
  r.appendChild(rPr);
  const t = el(doc, "t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = text;
  r.appendChild(t);
  return r;
}

// ================================================================ Tách LG

/**
 * Tách LG (dnqTachloigiai): từ gói OOXML các câu đã chọn, giữ "Câu N." + nội dung sau "Lời giải" của từng câu
 * (bỏ đề, phương án, dòng "Lời giải"), đặt dưới tiêu đề "LỜI GIẢI CHI TIẾT". Bảng/hình trong lời giải giữ nguyên.
 * opts.markMissing (plugin TACHLOIGIAI_BIGKEY): câu không có lời giải -> dòng "Câu N. THIẾU LỜI GIẢI Ở ĐÂY" (đỏ).
 * Trả về { xml, count, missing: [nhãn] } hoặc null nếu không có câu nào có lời giải. Thuần.
 */
export function extractSolutionsXml(pkgXml, opts = {}) {
  const doc = parseXml(pkgXml);
  const body = documentBody(doc);
  const nodes = kids(body).filter((n) => n.namespaceURI === W && (n.localName === "p" || n.localName === "tbl"));
  const keep = [];
  let label = "";
  let inSol = false;
  let first = false;
  let count = 0;
  let solved = true; // câu hiện tại đã có lời giải chưa
  const missing = [];
  const closeQuestion = () => {
    if (label && !solved) {
      missing.push(label);
      if (opts.markMissing) {
        const mp = el(doc, "p");
        mp.appendChild(textRun(doc, label + " ", { b: true, color: "0000FF" }));
        mp.appendChild(textRun(doc, "THIẾU LỜI GIẢI Ở ĐÂY", { b: true, color: "FF0000" }));
        keep.push(mp);
      }
    }
  };
  for (const n of nodes) {
    const isP = n.localName === "p";
    const t = isP ? pText(n) : "";
    const q = isP && reQ.exec(t);
    if (q) {
      closeQuestion();
      label = q[1].replace(/\s+/g, " ").trim();
      solved = false;
      inSol = false;
      continue;
    }
    if (isP && reBoundary.test(t)) {
      closeQuestion();
      label = "";
      inSol = false;
      continue;
    }
    if (isP && label && !inSol && reLG.test(t)) {
      inSol = true;
      solved = true;
      first = true;
      count++;
      continue;
    }
    if (!inSol) continue;
    if (first) {
      first = false;
      if (isP) {
        const pPr = kids(n).find((c) => c.localName === "pPr");
        n.insertBefore(textRun(doc, label + " ", { b: true, color: "0000FF" }), pPr ? pPr.nextSibling : n.firstChild);
      } else {
        const lp = el(doc, "p");
        lp.appendChild(textRun(doc, label, { b: true, color: "0000FF" }));
        keep.push(lp);
      }
    }
    keep.push(n);
  }
  closeQuestion();
  if (!count) return null;
  for (const n of [...body.childNodes]) body.removeChild(n); // bỏ cả sectPr: chèn vào cuối tài liệu không được tạo ngắt section
  const head = el(doc, "p");
  const hPPr = el(doc, "pPr");
  hPPr.appendChild(el(doc, "jc", { val: "center" }));
  head.appendChild(hPPr);
  head.appendChild(textRun(doc, "LỜI GIẢI CHI TIẾT", { b: true }));
  body.appendChild(head);
  for (const n of keep) {
    for (const s of [...n.getElementsByTagNameNS(W, "sectPr")]) s.parentNode.removeChild(s);
    body.appendChild(n);
  }
  body.appendChild(sentinelParagraph(doc));
  return { xml: new XMLSerializer().serializeToString(doc), count, missing };
}

/** Tách LG: lời giải của các câu đang chọn -> cuối tài liệu dưới "LỜI GIẢI CHI TIẾT" (đề giữ nguyên). */
export async function tachLG(context) {
  const sel = context.document.getSelection();
  sel.load("isEmpty");
  const ps = sel.paragraphs;
  ps.load("items");
  await context.sync();
  if (sel.isEmpty || !ps.items.length) return { ok: false, summary: "Chưa bôi đen các câu hỏi cần tách lời giải.", details: "" };
  const ox = ps.items[0].getRange("Whole").expandTo(ps.items[ps.items.length - 1].getRange("Whole")).getOoxml();
  await context.sync();
  const res = extractSolutionsXml(ox.value);
  if (!res) return { ok: false, summary: 'Vùng chọn không có câu "Câu N." nào có dòng "Lời giải".', details: "" };
  const body = context.document.body;
  body.insertParagraph("", "End"); // đoạn trống mới làm điểm chèn: không dính vào đoạn cuối hiện có
  body.insertOoxml(res.xml, "End");
  await context.sync();
  return `Đã tách lời giải của ${res.count} câu xuống cuối tài liệu.`;
}

/**
 * Chuyển lời giải xuống cuối (plugin TACHLOIGIAI_BIGKEY): mọi lời giải của đề được CHUYỂN (không chép) xuống cuối tài
 * liệu dưới "LỜI GIẢI CHI TIẾT"; câu thiếu lời giải có dòng "THIẾU LỜI GIẢI Ở ĐÂY" và được liệt kê.
 * Bản VBA: thêm "Câu cuối" giả, dùng Find/Replace wildcard với "Câu n" theo số thứ tự (sai khi số câu không liên tục).
 */
export async function chuyenLoiGiaiXuongCuoi(context) {
  const body = context.document.body;
  const ox = body.getOoxml();
  await context.sync();
  const res = extractSolutionsXml(ox.value, { markMissing: true });
  if (!res) return 'Không có câu "Câu N." nào có dòng "Lời giải".';
  await xoaLoiGiai(context); // cùng ranh giới khối (lời dẫn nhóm, PHẦN, HẾT, đoạn sát lề)
  body.insertParagraph("", "End");
  body.insertOoxml(res.xml, "End");
  await context.sync();
  const msg = `Đã chuyển lời giải của ${res.count} câu xuống cuối tài liệu.`;
  return res.missing.length
    ? { ok: false, summary: `${msg}\n⚠ Thiếu lời giải (đã đánh dấu ở cuối): ${res.missing.slice(0, 10).join(", ")}${res.missing.length > 10 ? "…" : ""}`, details: res.missing.join(", ") }
    : msg;
}

// ================================================================ Chemistry Formatter

const isLetterish = (c) => !!c && /[A-Za-z)\]]/.test(c);
const isDigit = (c) => !!c && c >= "0" && c <= "9";
const isDigit19 = (c) => !!c && c >= "1" && c <= "9"; // VBA: Asc > 48 And < 58 (không gồm 0)
const isSign = (c) => c === "+" || c === "-" || c === "–";
const isBreak = (c) => c === " " || c === "\t" || c === "\n" || c === "\v" || c === "\r";

/**
 * Luật của hdpnChemistryFormatter: s -> [{ ch, src, va }] (va: "" | "subscript" | "superscript"; src: vị trí ký tự gốc).
 *  - "-->" -> "→", "*" -> "·"
 *  - chữ số sau chữ cái / ")" / "]" -> chỉ số dưới (H2SO4, Cu(OH)2)
 *  - điện tích +/- sau chữ/số (không có dấu cách trước) -> chỉ số trên, kèm 1–2 chữ số trước (Fe3+, SO42-; NO3- giữ 3 dưới)
 *  - 6.02E23 -> "6.02 × 10" + "23" trên
 * Thuần.
 */
export function chemTransform(s, dec = ".") {
  // 1) thay ký tự
  const t = [];
  for (let i = 0; i < s.length; i++) {
    if (s.startsWith("-->", i)) {
      t.push({ ch: "→", src: i, va: "" });
      i += 2;
    } else t.push({ ch: s[i] === "*" ? "·" : s[i], src: i, va: "" });
  }
  const c = (i) => (t[i] ? t[i].ch : undefined);
  const n = t.length;
  // 2) chỉ số dưới
  let sub = false;
  for (let i = 0; i < n; i++) {
    const ch = c(i);
    if (isDigit(ch) || ch === dec) {
      if (sub) {
        if (ch !== dec) t[i].va = "subscript";
        else if (!isBreak(c(i + 1)) && c(i + 1) !== ".") {
          if (isLetterish(c(i + 1))) sub = false;
          else t[i].va = "subscript";
        }
      }
    } else sub = isLetterish(ch);
  }
  // 3) điện tích (chỉ số trên)
  const sup = (i) => t[i] && (t[i].va = "superscript");
  let supRegion = isLetterish(c(0));
  for (let i = 1; i < n; i++) {
    const ch = c(i), prev = c(i - 1), prev2 = c(i - 2), next = c(i + 1);
    if (supRegion && isSign(ch) && prev !== " ") {
      let need = true;
      const prevDigitSup = () => {
        if (prev2 === ")" || !isDigit19(prev)) return;
        if (isDigit19(prev2) || prev === "1" || prev === "2" || (prev === "3" && prev2 !== "O")) sup(i - 1);
      };
      if (i === n - 1) prevDigitSup();
      else if (isDigit19(next)) sup(i + 1);
      else if (prev2 !== ")" && isDigit19(prev)) prevDigitSup();
      else if (prev2 !== ")" && !isDigit19(prev) && next !== " " && !"(,.:;?!\t\r\n\v".includes(next)) {
        need = false; // gạch nối trong chữ (H-Cl), không phải điện tích
        supRegion = false;
      }
      if (need) sup(i);
    }
    const nx = c(i + 1);
    if (isBreak(nx)) supRegion = false;
    else if (isLetterish(nx)) supRegion = true;
  }
  // 4) ký hiệu E: 6.02E23 -> 6.02 × 10^23
  for (let i = n - 2; i >= 1; i--) {
    if (c(i) !== "E" && c(i) !== "e") continue;
    const nx = c(i + 1);
    if (!(isDigit(nx) || isSign(nx))) continue;
    let len = 1;
    while (isDigit(c(i + 1 + len))) len++;
    if (len === 1 && isSign(nx)) continue;
    const numBefore = isDigit(c(i - 1)) || (c(i - 1) === dec && isDigit(c(i - 2)));
    if (!numBefore) continue;
    for (let k = i + 1; k <= i + len; k++) t[k].va = "superscript";
    const src = t[i].src;
    t.splice(i, 1, ...[..." × 10"].map((ch) => ({ ch, src, va: "" })));
  }
  return t;
}

/** Áp chemTransform lên mọi đoạn của gói OOXML (bỏ qua công thức OMML, giữ hình/đối tượng). Thuần. */
export function chemFormatXml(pkgXml) {
  const doc = parseXml(pkgXml);
  let changed = 0;
  for (const p of [...doc.getElementsByTagNameNS(W, "p")]) {
    const runs = [...p.getElementsByTagNameNS(W, "r")].filter((r) => {
      for (let x = r.parentNode; x; x = x.parentNode) if (x.namespaceURI === M) return false;
      return true;
    });
    // 1) mỗi run chỉ còn một phần tử nội dung (t / tab / br / drawing…), rPr nhân bản
    const parts = [];
    for (const r of runs) {
      const rPr = kids(r).find((x) => x.localName === "rPr");
      const content = kids(r).filter((x) => x !== rPr);
      content.forEach((c, k) => {
        let nr = r;
        if (k > 0) {
          nr = el(doc, "r");
          if (rPr) nr.appendChild(rPr.cloneNode(true));
          nr.appendChild(c);
          parts[parts.length - 1].r.parentNode.insertBefore(nr, parts[parts.length - 1].r.nextSibling);
        }
        parts.push({ r: nr, c });
      });
    }
    // 2) chuỗi của đoạn (tab/xuống dòng/đối tượng là dấu ngăn) + vị trí của từng w:t
    let s = "";
    const texts = [];
    for (const { r, c } of parts) {
      if (c.localName === "t") {
        texts.push({ r, c, from: s.length, len: c.textContent.length });
        s += c.textContent;
      } else s += c.localName === "tab" ? "\t" : c.localName === "br" ? "\v" : "\u0001";
    }
    if (!texts.length) continue;
    const out = chemTransform(s);
    if (out.length === s.length && out.every((o, i) => o.ch === s[i] && !o.va)) continue;
    changed++;
    // 3) thay mỗi run chữ bằng các run theo đoạn định dạng (vertAlign)
    for (const it of texts) {
      const segs = [];
      for (const o of out) {
        if (o.src < it.from || o.src >= it.from + it.len) continue;
        const last = segs[segs.length - 1];
        if (last && last.va === o.va) last.text += o.ch;
        else segs.push({ va: o.va, text: o.ch });
      }
      const rPr = kids(it.r).find((x) => x.localName === "rPr");
      for (const sg of segs) {
        const nr = el(doc, "r");
        const pr = rPr ? rPr.cloneNode(true) : el(doc, "rPr");
        for (const v of kids(pr)) if (v.localName === "vertAlign") pr.removeChild(v);
        if (sg.va) putInRPr(pr, el(doc, "vertAlign", { val: sg.va }));
        if (kids(pr).length) nr.appendChild(pr);
        const t = el(doc, "t");
        t.setAttributeNS(XML_NS, "xml:space", "preserve");
        t.textContent = sg.text;
        nr.appendChild(t);
        it.r.parentNode.insertBefore(nr, it.r);
      }
      it.r.parentNode.removeChild(it.r);
    }
  }
  return changed ? { xml: new XMLSerializer().serializeToString(doc), changed } : null;
}

/** Chemistry Formatter: định dạng công thức hoá học trong vùng chọn. */
export async function chemFormatter(context) {
  const sel = context.document.getSelection();
  sel.load("isEmpty");
  await context.sync();
  if (sel.isEmpty) return { ok: false, summary: "Hãy bôi đen đoạn có công thức hoá học (vd. H2SO4, Fe3+, CuSO4*5H2O).", details: "" };
  const ox = sel.getOoxml();
  await context.sync();
  const res = chemFormatXml(ox.value);
  if (!res) return "Vùng chọn không có gì cần định dạng hoá học.";
  sel.insertOoxml(res.xml, "Replace");
  await context.sync();
  return `Đã định dạng hoá học ${res.changed} đoạn.`;
}

// ================================================================ Tô phiếu ZipGrade 50 câu từ đáp án

const EMU_PT = 12700;
const BUBBLE_STEP = 14.6; // pt giữa hai ô A–B–C–D (bản VBA)
const MADE_STEP = 21.15; // pt giữa hai ô mã đề A–E

/** Hình neo (wp:anchor) có docPr name = name trong document.xml. */
function anchorByName(doc, name) {
  for (const pr of doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "docPr"))
    if (pr.getAttribute("name") === name) return pr.parentNode;
  return null;
}
function fillDark(anchor) {
  const spPr = anchor.getElementsByTagNameNS("http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "spPr")[0];
  if (!spPr) return false;
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const noFill = [...spPr.childNodes].find((n) => n.localName === "noFill");
  const solid = spPr.ownerDocument.createElementNS(A, "a:solidFill");
  const clr = spPr.ownerDocument.createElementNS(A, "a:srgbClr");
  clr.setAttribute("val", "292929");
  solid.appendChild(clr);
  if (noFill) spPr.replaceChild(solid, noFill);
  else spPr.insertBefore(solid, [...spPr.childNodes].find((n) => n.localName === "ln") || null);
  return true;
}
function shift(anchor, axis, pt) {
  const pos = [...anchor.childNodes].find((n) => n.localName === (axis === "H" ? "positionH" : "positionV"));
  const off = pos && [...pos.childNodes].find((n) => n.localName === "posOffset");
  if (!off) return false;
  off.textContent = String(Math.round(+off.textContent + pt * EMU_PT));
  return true;
}

/**
 * Tô phiếu ZipGrade 50 câu trên document.xml của mẫu ans50zg.docx: ô "n" của câu n được tô đen và dời sang cột
 * đáp án (A +0, B +14,6 pt…); mã đề A–E (tuỳ chọn): ghi chữ vào hộp "Made", tô và dời ô "Ma". Thuần.
 * answers: { [num]: "A".."D" }. Trả về { xml, filled }.
 */
export function fillZipGradeXml(docXml, answers, made = "") {
  const doc = parseXml(docXml);
  let filled = 0;
  for (const [num, letter] of Object.entries(answers)) {
    const a = anchorByName(doc, String(num));
    const i = "ABCD".indexOf(letter);
    if (!a || i < 0) continue;
    if (fillDark(a) && shift(a, "H", BUBBLE_STEP * i)) filled++;
  }
  if (made) {
    // hộp "Made" của mẫu là hộp chữ rỗng (một đoạn không có run): ghi chữ vào đoạn đầu của w:txbxContent
    const box = anchorByName(doc, "Made");
    const bp = box && box.getElementsByTagNameNS(W, "txbxContent")[0] && box.getElementsByTagNameNS(W, "txbxContent")[0].getElementsByTagNameNS(W, "p")[0];
    if (bp) {
      for (const old of [...bp.getElementsByTagNameNS(W, "r")]) old.parentNode.removeChild(old);
      bp.appendChild(textRun(doc, made, { b: true }));
    }
    const ma = anchorByName(doc, "Ma");
    if (ma) fillDark(ma) && shift(ma, "V", MADE_STEP * "ABCDE".indexOf(made));
  }
  return { xml: new XMLSerializer().serializeToString(doc), filled };
}

/**
 * Tô phiếu ZG50 từ ĐA: đáp án gạch chân (hoặc "Chọn X") của câu 1–50 -> mở phiếu ZipGrade 50 câu đã tô trong cửa sổ
 * mới (lưu / in PDF từ đó). template: ArrayBuffer của ans50zg.docx. Bản VBA: lỗi khi có mã đề (biến arMade chưa
 * khai báo), xuất PDF cạnh file đề.
 */
export async function toPhieuZG50(context, made = "", template = null) {
  if (!template) return { ok: false, summary: "Chưa chọn file mẫu phiếu ans50zg.docx (trong thư mục plugin TophieuZG50tuDA của BigKey).", details: "" };
  if (made && !/^[A-E]$/.test(made)) throw new Error("Mã đề phải là một chữ A–E: " + made);
  const P = await loadParas(context);
  const plan = planBlocks(P).filter((b) => b.optFrom >= 0);
  if (!plan.length) return 'Không có câu trắc nghiệm A–D nào ("Câu N." + phương án).';
  const over = plan.filter((b) => b.num < 1 || b.num > 50).map((b) => b.num);
  if (over.length) return { ok: false, summary: `Phiếu ZipGrade chỉ có câu 1–50; đề có câu ${over.slice(0, 5).join(", ")}. Hãy đánh số lại.`, details: "" };
  const ox = plan.map((b) => P[b.optFrom].p.getRange("Whole").expandTo(P[b.optTo].p.getRange("Whole")).getOoxml());
  await context.sync();
  const answers = {};
  const missing = [];
  plan.forEach((b, k) => {
    const r = readAnswersXml(ox[k].value);
    const a = (r && r.underlined) || b.chonLetter;
    if (a) answers[b.num] = a;
    else missing.push(b.num);
  });
  // mẫu do người dùng chọn trong task pane (không phát hành file của plugin BigKey lên Pages)
  const zip = await JSZip.loadAsync(template);
  if (!zip.file("word/document.xml")) throw new Error("File mẫu không phải .docx hợp lệ.");
  const { xml, filled } = fillZipGradeXml(await zip.file("word/document.xml").async("string"), answers, made);
  zip.file("word/document.xml", xml);
  const b64 = await zip.generateAsync({ type: "base64" });
  context.application.createDocument(b64).open();
  await context.sync();
  const msg = `Đã tô ${filled} câu${made ? `, mã đề ${made}` : ""} trên phiếu ZipGrade 50 (mở trong cửa sổ mới — lưu hoặc in PDF từ đó).`;
  return missing.length ? { ok: false, summary: `${msg}\n⚠ Thiếu đáp án: Câu ${missing.join(", ")}.`, details: "" } : msg;
}
