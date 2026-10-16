// Chuẩn hoá BigKey — bố cục đoạn như bản VBA (bkBonDong): Office JS không đặt được tab stop hay
// "Keep with next" cho từng đoạn, nên định dạng đoạn được viết lại trong OOXML.
// Mỗi đoạn: style Normal, xoá định dạng đoạn trực tiếp (canh lề, giãn dòng, tab stop, keep with next…),
// lề trái 1 cm, phải 0, dòng đầu 0, cách đoạn 0, giãn dòng đơn; mọi run chữ Times New Roman 12
// (giữ đậm/nghiêng/gạch chân/màu; không đụng công thức OMML). Giữ section break trong pPr.
// Tab ngay sau "Câu N." -> dấu cách (như bản VBA). Bỏ ngắt trang thủ công.

import { parseXml, documentBody, bodyParagraphs } from "../dgnl/removeBreakEng.js";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
export const CM = 567; // twip
export const FONT = "Times New Roman";
export const HALF_PT = "24"; // 12 pt

// Thứ tự phần tử con của w:rPr theo schema (Word từ chối rPr sai thứ tự khi insertOoxml)
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
  "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
  "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
  "specVanish", "oMath"];

const el = (doc, name, attrs = {}) => {
  const e = doc.createElementNS(W, "w:" + name);
  for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, "w:" + k, v);
  return e;
};
const kids = (e, name) => [...e.childNodes].filter((n) => n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name));

/** Chèn child vào rPr đúng vị trí theo RPR_ORDER. */
function putInRPr(rPr, node) {
  const rank = RPR_ORDER.indexOf(node.localName);
  const after = kids(rPr).find((c) => RPR_ORDER.indexOf(c.localName) > rank);
  rPr.insertBefore(node, after || null);
}

/** Đặt Times New Roman 12 cho rPr (tạo rPr nếu cần), giữ các định dạng khác. */
export function setFont(doc, owner) {
  let rPr = kids(owner, "rPr")[0];
  if (!rPr) {
    rPr = el(doc, "rPr");
    // w:r: rPr đứng đầu; w:pPr: rPr đứng sau spacing/ind…, trước sectPr/pPrChange (thứ tự schema)
    const before = owner.localName === "pPr" ? kids(owner).find((c) => c.localName === "sectPr" || c.localName === "pPrChange") : owner.firstChild;
    owner.insertBefore(rPr, before || null);
  }
  for (const c of kids(rPr)) if (c.localName === "rFonts" || c.localName === "sz" || c.localName === "szCs") rPr.removeChild(c);
  putInRPr(rPr, el(doc, "rFonts", { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT }));
  putInRPr(rPr, el(doc, "sz", { val: HALF_PT }));
  putInRPr(rPr, el(doc, "szCs", { val: HALF_PT }));
  return rPr;
}

/** pPr chuẩn BigKey: [tabs] + cách đoạn 0, giãn dòng đơn + lề trái 1 cm. tabsTwip: vị trí tab trái. */
export function standardPPr(doc, { tabsTwip = [] } = {}) {
  const pPr = el(doc, "pPr");
  if (tabsTwip.length) {
    const tabs = el(doc, "tabs");
    for (const pos of tabsTwip) tabs.appendChild(el(doc, "tab", { val: "left", pos: String(pos) }));
    pPr.appendChild(tabs);
  }
  pPr.appendChild(el(doc, "spacing", { before: "0", after: "0", line: "240", lineRule: "auto" }));
  pPr.appendChild(el(doc, "ind", { left: String(CM), right: "0", firstLine: "0" }));
  return pPr;
}

const inMath = (n) => {
  for (let p = n.parentNode; p; p = p.parentNode) if (p.namespaceURI === M) return true;
  return false;
};
const reCauLabel = /^\s*Câu\s*\d+\s*[.:]\s*$/;

/** Chuẩn hoá một w:p tại chỗ. */
export function normalizeParagraph(doc, p) {
  const old = kids(p, "pPr")[0];
  const pPr = standardPPr(doc);
  if (old) {
    const markRPr = kids(old, "rPr")[0];
    const sect = kids(old, "sectPr")[0];
    if (markRPr) pPr.appendChild(markRPr);
    if (sect) pPr.appendChild(sect);
    p.replaceChild(pPr, old);
  } else p.insertBefore(pPr, p.firstChild);
  setFont(doc, pPr);

  let text = "";
  let tabFixed = false;
  for (const r of [...p.getElementsByTagNameNS(W, "r")]) {
    if (inMath(r)) continue;
    setFont(doc, r);
    for (const c of kids(r)) {
      if (c.localName === "t") text += c.textContent;
      else if (c.localName === "br" && c.getAttributeNS(W, "type") === "page") r.removeChild(c);
      else if (c.localName === "tab" && !tabFixed) {
        tabFixed = true;
        if (reCauLabel.test(text)) {
          // "Câu 1.<tab>Nội dung" -> "Câu 1. Nội dung"
          const t = el(doc, "t");
          t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
          t.textContent = /\s$/.test(text) ? "" : " ";
          r.replaceChild(t, c);
          text += t.textContent;
        }
      }
    }
  }
}

/**
 * Đoạn rỗng "hy sinh" đặt cuối gói OOXML trước khi insertOoxml("Replace"): Word gộp đoạn CUỐI của nội dung
 * chèn vào dấu ¶ của đoạn đích, và (Office JS) giữ định dạng đoạn CŨ của dấu ¶ đó -> mất tab stop / lề mới
 * (bug thật 2.0.0-beta.4: phương án A–D không có tab). Đoạn rỗng này nhận dấu ¶ cũ; bước dọn dòng trống xoá nó.
 */
export function sentinelParagraph(doc) {
  return doc.createElementNS(W, "w:p");
}

/** Chuẩn hoá mọi đoạn (cấp body) của một gói OOXML (flat OPC từ getOoxml), thêm đoạn sentinel cuối. Thuần. */
export function normalizeParagraphsXml(pkgXml) {
  const doc = parseXml(pkgXml);
  const body = documentBody(doc);
  const ps = bodyParagraphs(body);
  for (const p of ps) normalizeParagraph(doc, p);
  const last = ps[ps.length - 1];
  if (last) last.parentNode.insertBefore(sentinelParagraph(doc), last.nextSibling);
  return new XMLSerializer().serializeToString(doc);
}
