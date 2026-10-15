// Fix lỗi break dòng T.Anh  (spec: docs/specs/dgnl/Remove_Break_Eng.md)
// Gỡ "Keep with next" thừa khiến Word đẩy cả khối sang trang sau, để lại khoảng trắng lớn.
//
// Giữ nguyên (không đụng tới):
//   - Dòng đề "Câu N." và mọi list item khác
//   - Dòng "Questions …" (luôn đi cùng câu đầu tiên)
//   - Dòng đáp án A–D mà ngay sau vẫn là dòng đáp án (chuỗi đáp án không bị tách)
//   - Đoạn trong bảng
// Tắt "Keep with next" (ghi <w:keepNext w:val="0"/>) ở mọi đoạn còn lại đang bật — kể cả bật do style:
// "Lời giải", giải thích, bài đọc, dòng đáp án cuối.
//
// Office JS không có Paragraph.keepWithNext -> đọc/ghi qua OOXML.

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

const reAnswer = /^\s*[A-D]\s*\./;
const reKeepHeading = /^\s*Questions?\b/i;
const CHUNK = 25; // số đoạn insertOoxml mỗi lần sync
const MAX_PER_PARA_SCAN = 600; // giới hạn khi phải đọc OOXML từng đoạn

// Đoạn chứa các cấu trúc này không round-trip OOXML an toàn -> bỏ qua
const RISKY = [
  [W_NS, "sectPr"], // section break nằm trong pPr
  [W_NS, "fldChar"],
  [W_NS, "fldSimple"],
  [W_NS, "object"], // OLE / MathType
  [W_NS, "ins"],
  [W_NS, "del"], // track changes
  [W_NS, "commentRangeStart"],
  [W_NS, "commentRangeEnd"],
];

// ---------- OOXML helpers ----------
export function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("Không đọc được OOXML của tài liệu.");
  return doc;
}

/** <w:body> của part /word/document.xml trong flat-OPC package */
export function documentBody(doc) {
  for (const part of doc.getElementsByTagNameNS(PKG_NS, "part")) {
    if (part.getAttributeNS(PKG_NS, "name") === "/word/document.xml") {
      return part.getElementsByTagNameNS(W_NS, "body")[0] || null;
    }
  }
  return doc.getElementsByTagNameNS(W_NS, "body")[0] || null;
}

/** Các <w:p> thuộc body (bỏ đoạn trong text box và nhánh mc:Fallback) theo thứ tự tài liệu */
export function bodyParagraphs(body) {
  const out = [];
  if (!body) return out;
  for (const p of body.getElementsByTagNameNS(W_NS, "p")) {
    let nested = false;
    for (let n = p.parentNode; n && n !== body; n = n.parentNode) {
      if ((n.namespaceURI === W_NS && n.localName === "txbxContent") || (n.namespaceURI === MC_NS && n.localName === "Fallback")) {
        nested = true;
        break;
      }
    }
    if (!nested) out.push(p);
  }
  return out;
}

const childW = (el, name) => {
  for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.namespaceURI === W_NS && c.localName === name) return c;
  }
  return null;
};

/** Trạng thái keepNext đặt trực tiếp: true / false / null (không đặt -> theo style) */
function directKeepNext(p) {
  const k = childW(childW(p, "pPr"), "keepNext");
  if (!k) return null;
  const v = k.getAttributeNS(W_NS, "val");
  return !(v && /^(0|false|off)$/i.test(v));
}

const isRisky = (p) => RISKY.some(([ns, n]) => p.getElementsByTagNameNS(ns, n).length > 0);

/** Chữ/số trong đoạn, để đối chiếu OOXML <-> Word */
const sig = (s) =>
  s
    .toLowerCase()
    .replace(/[^0-9a-zÀ-ỹ]+/g, "")
    .slice(0, 40);
const xmlText = (p) =>
  Array.from(p.getElementsByTagNameNS(W_NS, "t"))
    .concat(Array.from(p.getElementsByTagNameNS(M_NS, "t")))
    .map((t) => t.textContent)
    .join("");

/** Ghi <w:keepNext w:val="0"/> vào đoạn đầu tiên của package. Trả về XML mới, hoặc null nếu bỏ qua. */
function disableKeepNext(pkgXml) {
  const doc = parseXml(pkgXml);
  const p = bodyParagraphs(documentBody(doc))[0];
  if (!p || isRisky(p)) return null;
  let pPr = childW(p, "pPr");
  if (!pPr) {
    pPr = doc.createElementNS(W_NS, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  let k = childW(pPr, "keepNext");
  if (!k) {
    k = doc.createElementNS(W_NS, "w:keepNext");
    // Thứ tự schema: pStyle, keepNext, ...
    const pStyle = childW(pPr, "pStyle");
    pPr.insertBefore(k, pStyle ? pStyle.nextSibling : pPr.firstChild);
  }
  k.setAttributeNS(W_NS, "w:val", "0");
  return new XMLSerializer().serializeToString(doc);
}

/** Map tên style (local) -> keepWithNext. Cần WordApi 1.5; không hỗ trợ thì trả về map rỗng. */
async function styleKeepMap(context) {
  const map = new Map();
  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) return map;
  const styles = context.document.getStyles();
  styles.load("items/nameLocal,items/type,items/paragraphFormat/keepWithNext");
  await context.sync();
  for (const s of styles.items) {
    if (s.type === "Paragraph" || s.type === "Linked") map.set(s.nameLocal, !!s.paragraphFormat.keepWithNext);
  }
  return map;
}

// ---------- main ----------
export async function removeBreakEng(context) {
  const body = context.document.body;
  const ps = body.paragraphs;
  ps.load("items/text,items/style,items/isListItem,items/tableNestingLevel");
  const bodyXml = body.getOoxml();
  await context.sync();
  const items = ps.items;
  const N = items.length;
  const styleKeep = await styleKeepMap(context);

  // 1) Chọn đoạn ứng viên theo cấu trúc
  const isPlain = (i) => i < N && !items[i].isListItem && items[i].tableNestingLevel === 0;
  const isAnswer = (i) => isPlain(i) && reAnswer.test(items[i].text);
  const candidates = [];
  for (let i = 0; i < N; i++) {
    if (!isPlain(i)) continue;
    if (reKeepHeading.test(items[i].text)) continue;
    if (isAnswer(i) && isAnswer(i + 1)) continue;
    candidates.push(i);
  }
  const fromStyle = (i) => styleKeep.get(items[i].style) === true;

  // 2) Xác định đoạn đang bật keepNext (trực tiếp hoặc do style)
  let xmlParas = bodyParagraphs(documentBody(parseXml(bodyXml.value)));
  if (xmlParas.length === N + 1 && xmlText(xmlParas[N]) === "") xmlParas = xmlParas.slice(0, N);
  const mapped =
    xmlParas.length === N &&
    candidates.every((i) => sig(xmlText(xmlParas[i])) === sig(items[i].text) || sig(items[i].text) === "");

  let targets;
  let skipped = 0;
  if (mapped) {
    targets = candidates.filter((i) => {
      const d = directKeepNext(xmlParas[i]);
      if (!(d === true || (d === null && fromStyle(i)))) return false;
      if (isRisky(xmlParas[i])) {
        skipped++;
        return false;
      }
      return true;
    });
  } else {
    // Không đối chiếu được OOXML với danh sách đoạn -> đọc OOXML từng ứng viên (chậm hơn)
    console.warn(`OOXML/paragraph mapping failed (${xmlParas.length} vs ${N}); reading per paragraph`);
    if (candidates.length > MAX_PER_PARA_SCAN) {
      throw new Error(
        `Không đối chiếu được cấu trúc tài liệu và có ${candidates.length} đoạn cần kiểm tra (> ${MAX_PER_PARA_SCAN}). Hãy chạy trên từng phần nhỏ hơn của file.`
      );
    }
    const oo = candidates.map((i) => items[i].getOoxml());
    await context.sync();
    targets = [];
    candidates.forEach((i, k) => {
      const p = bodyParagraphs(documentBody(parseXml(oo[k].value)))[0];
      if (!p) return;
      const d = directKeepNext(p);
      if (!(d === true || (d === null && fromStyle(i)))) return;
      if (isRisky(p)) skipped++;
      else targets.push(i);
    });
  }
  if (!targets.length) {
    return skipped
      ? `Không sửa đoạn nào. Bỏ qua ${skipped} đoạn có field/section break/track changes — sửa tay các đoạn này.`
      : "Không có đoạn nào bị \"Keep with next\" thừa.";
  }

  // 3) Ghi lại theo từng chunk, từ cuối tài liệu lên đầu
  const origEmpty = (i) => i < N && items[i].text.trim() === "";
  let done = 0;
  let failed = 0;
  let cleaned = 0;
  const order = targets.slice().reverse();
  for (let c = 0; c < order.length; c += CHUNK) {
    const chunk = order.slice(c, c + CHUNK);
    try {
      const oo = chunk.map((i) => items[i].getOoxml());
      await context.sync();

      const checks = [];
      chunk.forEach((i, k) => {
        const xml = disableKeepNext(oo[k].value);
        if (!xml) {
          skipped++;
          return;
        }
        const r = items[i].insertOoxml(xml, "Replace");
        const inner = r.paragraphs;
        inner.load("items/text");
        const next = inner.getLast().getNextOrNullObject();
        next.load("text");
        checks.push({ i, inner, next });
      });
      await context.sync();
      done += checks.length;

      // insertOoxml đôi khi sinh thêm 1 đoạn rỗng -> xoá
      for (const { i, inner, next } of checks) {
        const ps2 = inner.items;
        if (ps2.length > 1 && ps2[ps2.length - 1].text === "" && !origEmpty(i)) {
          ps2[ps2.length - 1].delete();
          cleaned++;
        } else if (!next.isNullObject && next.text === "" && !origEmpty(i + 1)) {
          next.delete();
          cleaned++;
        }
      }
      await context.sync();
    } catch (e) {
      console.error("removeBreakEng chunk failed", e);
      failed += chunk.length;
    }
  }

  // 4) Kiểm tra số đoạn
  const after = body.paragraphs;
  after.load("items/tableNestingLevel");
  await context.sync();
  const lines = [`Đã tắt "Keep with next" ở ${done} đoạn.`];
  if (skipped) lines.push(`Bỏ qua ${skipped} đoạn có field/section break/track changes (sửa tay nếu cần).`);
  if (failed) lines.push(`⚠ Không sửa được ${failed} đoạn.`);
  if (after.items.length !== N) {
    lines.push(`⚠ Số đoạn thay đổi ${N} → ${after.items.length}, hãy kiểm tra lại (Ctrl+Z để hoàn tác).`);
  }
  return lines.join("\n");
}
