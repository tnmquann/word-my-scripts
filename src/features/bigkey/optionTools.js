// Công cụ sửa khối phương án A–D tại chỗ (thuần OOXML, test được ngoài Word) — port các tool BigKey VBA.
// Input: flat-OPC package của 1–4 đoạn chứa nhãn A. B. C. D. (như Range.getOoxml()).
// Khác rebuildOptionsXml: GIỮ bố cục (đoạn, pPr/tab stop, run nội dung), chỉ sửa đúng phần mô tả.
// Sửa từ cuối khối lên đầu để tham chiếu token phía trước vẫn đúng.
// Mọi hàm trả về null nếu không tìm đủ nhãn A–D theo thứ tự; xml trả về luôn kết thúc bằng đoạn sentinel.

import { parseXml, documentBody, bodyParagraphs } from "../dgnl/removeBreakEng.js";
import { tokenize, findLabels, isBlank } from "./optionLayout.js";
import { W, putInRPr, sentinelParagraph } from "./normalizeXml.js";

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const BLUE = "0000FF";
const RED = "FF0000";

const child = (e, name) => {
  for (let c = e && e.firstElementChild; c; c = c.nextElementSibling) if (c.namespaceURI === W && c.localName === name) return c;
  return null;
};

function load(pkgXml) {
  const doc = parseXml(pkgXml);
  const body = documentBody(doc);
  const paras = bodyParagraphs(body).filter((p) => p.parentNode === body);
  if (!paras.length) return null;
  const toks = tokenize(paras);
  const labels = findLabels(toks, true, ".):") || findLabels(toks, false, ".):");
  if (!labels) return null;
  // phạm vi token nội dung của từng phương án: [a, b) đã bỏ khoảng trắng hai đầu
  const ranges = labels.map((lb, k) => {
    let a = lb.end;
    let b = k < 3 ? labels[k + 1].start : toks.length;
    while (a < b && isBlank(toks[a])) a++;
    while (b > a && isBlank(toks[b - 1])) b--;
    return [a, b];
  });
  return { doc, body, toks, labels, ranges };
}

/** Thêm sentinel làm đoạn cuối (trước sectPr cấp body) rồi serialize. */
function finish({ doc, body }) {
  body.insertBefore(sentinelParagraph(doc), child(body, "sectPr"));
  return new XMLSerializer().serializeToString(doc);
}

const mkT = (doc, s) => {
  const t = doc.createElementNS(W, "w:t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = s;
  return t;
};

/** Cắt w:t tại offset; trả về nút con bắt đầu phần sau (null nếu không còn gì). */
function cutText(tEl, off) {
  const s = tEl.textContent;
  if (off <= 0) return tEl;
  if (off >= s.length) return tEl.nextSibling;
  const t2 = mkT(tEl.ownerDocument, s.slice(off));
  tEl.textContent = s.slice(0, off);
  tEl.setAttributeNS(XML_NS, "xml:space", "preserve");
  tEl.parentNode.insertBefore(t2, tEl.nextSibling);
  return t2;
}

const isEmptyRun = (run) => ![...run.childNodes].some((c) => c.nodeType === 1 && !(c.namespaceURI === W && c.localName === "rPr"));

/** Tách run trước nút con `at`: run gốc giữ phần trước, run mới (rPr nhân bản) chứa at..cuối, chèn ngay sau. */
function splitRun(run, at) {
  if (!at) return null;
  const r2 = run.cloneNode(false);
  const rPr = child(run, "rPr");
  if (rPr) r2.appendChild(rPr.cloneNode(true));
  while (at) {
    const nx = at.nextSibling;
    r2.appendChild(at);
    at = nx;
  }
  run.parentNode.insertBefore(r2, run.nextSibling);
  if (isEmptyRun(run)) run.parentNode.removeChild(run);
  return r2;
}

/** Tách ký tự của token thành run riêng; trả về run đó. */
function isolate(tk) {
  splitRun(tk.run, cutText(tk.node, tk.idx + tk.ch.length));
  return splitRun(tk.run, cutText(tk.node, tk.idx));
}

/** Đặt (val chuỗi, "" = không thuộc tính val) hoặc xoá (val null) thuộc tính `name` trong rPr, đúng thứ tự schema. */
function setProp(rPr, name, val) {
  for (const c of [...rPr.childNodes]) if (c.nodeType === 1 && c.namespaceURI === W && c.localName === name) rPr.removeChild(c);
  if (val === null) return;
  const e = rPr.ownerDocument.createElementNS(W, "w:" + name);
  if (val) e.setAttributeNS(W, "w:val", val);
  putInRPr(rPr, e);
}

function rPrOf(run) {
  let rPr = child(run, "rPr");
  if (!rPr) {
    rPr = run.ownerDocument.createElementNS(W, "w:rPr");
    run.insertBefore(rPr, run.firstChild);
  }
  return rPr;
}

function style(run, color, underline) {
  const rPr = rPrOf(run);
  setProp(rPr, "b", "");
  setProp(rPr, "color", color);
  setProp(rPr, "u", underline ? "single" : null);
}

/** Tô 4 nhãn: chữ cái + "." đậm xanh; chữ cái đáp án đậm đỏ gạch chân. "A)"/"A:" -> "A.". */
function mark(st, letter) {
  for (let k = 3; k >= 0; k--) {
    const lb = st.labels[k];
    const pt = st.toks[lb.end - 1];
    const pr = isolate(pt);
    if (pt.ch !== ".") child(pr, "t").textContent = ".";
    style(pr, BLUE, false);
    style(isolate(st.toks[lb.start]), lb.letter === letter ? RED : BLUE, lb.letter === letter);
  }
}

const hlOn = (h) => !!h && (h.getAttributeNS(W, "val") || "none") !== "none";

/** Các w:highlight đang bật trong token (ký tự: rPr của run; node: rPr run + mọi highlight bên trong). */
function highlights(tk) {
  const out = [];
  if (tk.t === "ch") out.push(child(tk.rPr, "highlight"));
  else if (tk.t === "node") {
    if (tk.run) out.push(child(child(tk.run, "rPr"), "highlight"));
    out.push(...tk.node.getElementsByTagNameNS(W, "highlight"));
  }
  return out.filter(hlOn);
}

const currentAnswer = (st) => (st.labels.find((l) => l.correct) || { letter: "" }).letter;

function highlightedOption(st) {
  const k = st.ranges.findIndex(([a, b]) => st.toks.slice(a, b).some((tk) => !isBlank(tk) && highlights(tk).length));
  return k < 0 ? "" : "ABCD"[k];
}

/**
 * Đánh dấu đáp án trên nhãn (giữ nguyên bố cục). answer:
 *   "A".."D"         -> phương án đó là đáp án;
 *   "none"           -> không có đáp án (mọi nhãn đậm xanh, bỏ gạch chân) — "Chuyển đáp án thành phương án";
 *   khác ("keep", "", null) -> giữ đáp án đang gạch chân (nếu có) — "Chuẩn hoá màu PA-ĐA-LG".
 * Trả về { xml, letter } (letter = đáp án sau khi đánh dấu, "" nếu không có) hoặc null.
 */
export function markAnswerXml(pkgXml, answer = "keep") {
  const st = load(pkgXml);
  if (!st) return null;
  const letter = answer === "none" ? "" : /^[ABCD]$/.test(answer || "") ? answer : currentAnswer(st);
  mark(st, letter);
  return { xml: finish(st), letter };
}

/** Đọc đáp án: underlined = nhãn đầu tiên có chữ cái gạch chân; highlighted = phương án đầu tiên có nội dung tô sáng. */
export function readAnswersXml(pkgXml) {
  const st = load(pkgXml);
  if (!st) return null;
  return { underlined: currentAnswer(st), highlighted: highlightedOption(st) };
}

/** Nội dung phương án tô sáng -> bỏ tô sáng, chữ đỏ; nhãn của phương án đó thành đáp án. Trả về { xml, letter } hoặc null. */
export function highlightToRedXml(pkgXml) {
  const st = load(pkgXml);
  if (!st) return null;
  const letter = highlightedOption(st);
  const hs = new Set();
  for (const [a, b] of st.ranges) for (const tk of st.toks.slice(a, b)) for (const h of highlights(tk)) hs.add(h);
  for (const h of hs) {
    const rPr = h.parentNode;
    rPr.removeChild(h);
    setProp(rPr, "color", RED);
  }
  if (letter) mark(st, letter);
  return { xml: finish(st), letter };
}

/** Mỗi phương án một đoạn: tách đoạn trước nhãn B/C/D đứng sau tab/xuống dòng (bỏ tab/br đó); đoạn mới nhận pPr gốc trừ w:tabs. */
export function splitOptionsXml(pkgXml) {
  const st = load(pkgXml);
  if (!st) return null;
  const { doc, toks, labels } = st;
  for (let k = 3; k > 0; k--) {
    let i = labels[k].start - 1;
    while (i >= 0 && toks[i].t === "ch" && toks[i].ch === " ") i--;
    const sep = toks[i];
    if (!sep || !sep.node || (sep.t !== "tab" && sep.t !== "br")) continue; // đã ở đầu đoạn
    const lt = toks[labels[k].start];
    const p = lt.run.parentNode;
    const start = splitRun(lt.run, cutText(lt.node, lt.idx));
    const np = doc.createElementNS(W, "w:p");
    const pPr = child(p, "pPr");
    if (pPr) {
      const c = pPr.cloneNode(true);
      for (const n of [...c.childNodes]) if (n.nodeType === 1 && n.namespaceURI === W && (n.localName === "tabs" || n.localName === "sectPr")) c.removeChild(n);
      // ngắt section (nếu có) đi theo phần cuối của đoạn
      const sect = child(pPr, "sectPr");
      if (sect) c.insertBefore(sect, child(c, "pPrChange"));
      np.appendChild(c);
    }
    for (let n = start; n; ) {
      const nx = n.nextSibling;
      np.appendChild(n);
      n = nx;
    }
    p.parentNode.insertBefore(np, p.nextSibling);
    sep.run.removeChild(sep.node);
    if (isEmptyRun(sep.run)) sep.run.parentNode.removeChild(sep.run);
  }
  return { xml: finish(st) };
}

/** add=true: nội dung mỗi phương án kết thúc bằng "." (không thành ".."); add=false: bỏ một "." cuối (bỏ qua khoảng trắng cuối). */
export function optionPeriodsXml(pkgXml, add) {
  const st = load(pkgXml);
  if (!st) return null;
  const { doc, toks } = st;
  for (let k = 3; k >= 0; k--) {
    const [a, b] = st.ranges[k];
    if (a >= b) continue;
    const last = toks[b - 1];
    if (last.t === "ch") {
      const s = last.node.textContent;
      const end = last.idx + last.ch.length;
      if (add && last.ch !== ".") {
        last.node.textContent = s.slice(0, end) + "." + s.slice(end);
      } else if (!add && last.ch === ".") {
        if (s.length > 1) last.node.textContent = s.slice(0, last.idx) + s.slice(end);
        else {
          last.run.removeChild(last.node);
          if (isEmptyRun(last.run)) last.run.parentNode.removeChild(last.run);
        }
      }
    } else if (add && last.t === "node") {
      // sau công thức / đối tượng: run mới, định dạng theo ký tự gần nhất trước đó trong phương án
      const r = doc.createElementNS(W, "w:r");
      let i = b - 2;
      while (i >= a && toks[i].t !== "ch") i--;
      if (i >= a && toks[i].rPr) r.appendChild(toks[i].rPr.cloneNode(true));
      r.appendChild(mkT(doc, "."));
      let after = last.node;
      if (last.run) {
        splitRun(last.run, last.node.nextSibling);
        after = last.run;
      }
      after.parentNode.insertBefore(r, after.nextSibling);
    }
  }
  return { xml: finish(st) };
}
