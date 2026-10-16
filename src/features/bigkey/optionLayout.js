// Dàn phương án A–D (thuần OOXML, test được ngoài Word).
// Input: flat-OPC package của 1 khối đoạn chứa đủ A., B., C., D. (có thể nằm trên 1–4 đoạn, cách nhau bởi tab/xuống dòng).
// Output: package mới với các đoạn phương án được dàn lại:
//   - cả 4 ngắn  -> 1 dòng: A \t B \t C \t D (tab 5,25 / 9,5 / 13,75 cm)
//   - vừa        -> 2 dòng: A \t B / C \t D (tab 9,5 cm)
//   - dài        -> 4 dòng
// Giữ nguyên nội dung giàu định dạng (công thức OMML, MathType, hình) của từng phương án.
// Nhãn: "A." đậm xanh; chữ cái đáp án đúng (đang gạch chân) đậm đỏ gạch chân.
// Thay cho bkDoRong của BigKey (đo vị trí thật trên trang — Office JS không có API đo).

import { parseXml, documentBody, bodyParagraphs } from "../dgnl/removeBreakEng.js";
import { standardPPr, setFont, sentinelParagraph } from "./normalizeXml.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const CM = 567; // twips / cm
const BLUE = "0000FF";
const RED = "FF0000";
// Ngưỡng độ rộng (pt) cho 1 phương án gồm nhãn — cột 4,25 cm ≈ 120 pt, cột 8,5 cm ≈ 241 pt
export const FOUR_PER_LINE = 110;
export const TWO_PER_LINE = 235;
const CHAR_PT = 5.6; // ước lượng độ rộng trung bình 1 ký tự Times New Roman 12
const MATH_CHAR_PT = 6.5;

const child = (el, ns, name) => {
  for (let c = el && el.firstElementChild; c; c = c.nextElementSibling) if (c.namespaceURI === ns && c.localName === name) return c;
  return null;
};

/**
 * Tách các đoạn thành chuỗi token: ký tự (kèm rPr), tab, br, hoặc node nguyên khối (hình, công thức).
 * Tham chiếu về node gốc (cho optionTools sửa tại chỗ): run, node (w:t/w:tab/w:br/… gốc), idx (offset UTF-16 trong w:t);
 * br nối 2 đoạn không có run/node.
 */
export function tokenize(paras) {
  const toks = [];
  paras.forEach((p, pi) => {
    if (pi > 0) toks.push({ t: "br" });
    for (let n = p.firstElementChild; n; n = n.nextElementSibling) {
      if (n.namespaceURI === W && n.localName === "pPr") continue;
      if (n.namespaceURI === W && n.localName === "r") {
        const rPr = child(n, W, "rPr");
        for (let c = n.firstElementChild; c; c = c.nextElementSibling) {
          if (c.namespaceURI !== W) continue;
          if (c.localName === "t") {
            let idx = 0;
            for (const ch of c.textContent) {
              toks.push({ t: "ch", ch, rPr, run: n, node: c, idx });
              idx += ch.length;
            }
          } else if (c.localName === "tab") toks.push({ t: "tab", run: n, node: c });
          else if (c.localName === "br" && c.getAttributeNS(W, "type") !== "page") toks.push({ t: "br", run: n, node: c });
          else if (c.localName !== "rPr" && !(c.localName === "br")) {
            // drawing / object / pict / sym… : giữ nguyên cả run (chỉ phần này)
            const r = n.cloneNode(false);
            if (rPr) r.appendChild(rPr.cloneNode(true));
            r.appendChild(c.cloneNode(true));
            toks.push({ t: "node", el: r, w: nodeWidth(c), run: n, node: c });
          }
        }
      } else if (n.namespaceURI === W && /^(bookmarkStart|bookmarkEnd|proofErr|lastRenderedPageBreak)$/.test(n.localName)) {
        continue;
      } else {
        // m:oMath, w:hyperlink, w:fldSimple, w:ins… : giữ nguyên
        toks.push({ t: "node", el: n.cloneNode(true), w: nodeWidth(n), node: n });
      }
    }
  });
  return toks;
}

/** Ước lượng độ rộng (pt) của 1 node không phải chữ. */
function nodeWidth(n) {
  const ext = n.getElementsByTagNameNS(WP, "extent")[0];
  if (ext) return (+ext.getAttribute("cx") || 0) / 12700;
  const shape = n.getElementsByTagName("v:shape")[0];
  if (shape) {
    const m = /width:\s*([\d.]+)pt/.exec(shape.getAttribute("style") || "");
    if (m) return +m[1];
  }
  const mt = Array.from(n.getElementsByTagNameNS(M, "t")).map((t) => t.textContent).join("");
  if (mt) return mt.length * MATH_CHAR_PT;
  const tt = Array.from(n.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join("");
  return tt.length * CHAR_PT;
}

export const isUnderlined = (rPr) => {
  const u = child(rPr, W, "u");
  if (!u) return false;
  const v = u.getAttributeNS(W, "val");
  return !!v && v !== "none";
};

/**
 * Tìm 4 nhãn A,B,C,D theo đúng thứ tự (chữ cái + "." hoặc ")").
 * strict: nhãn B–D phải đứng sau tab/xuống dòng (tránh nhầm "điểm B." trong nội dung);
 * không strict: chấp nhận cả sau dấu cách (phương án viết liền "A. 1 B. 2 C. 3 D. 4").
 * puncts: các dấu sau chữ cái được coi là nhãn (optionTools thêm ":").
 */
export function findLabels(toks, strict, puncts = ".)") {
  const labels = [];
  let want = 0;
  for (let i = 0; i < toks.length && want < 4; i++) {
    const tk = toks[i];
    if (tk.t !== "ch" || tk.ch !== "ABCD"[want]) continue;
    // trước nhãn: đầu khối / tab / br (/ khoảng trắng nếu không strict)
    let k = i - 1;
    while (strict && want > 0 && toks[k] && toks[k].t === "ch" && toks[k].ch === " ") k--; // "⇥ B." vẫn hợp lệ
    const prev = toks[k];
    const okPrev = !prev || prev.t === "tab" || prev.t === "br" || (!(strict && want > 0) && prev.t === "ch" && /\s/.test(prev.ch));
    if (!okPrev) continue;
    // sau chữ cái: (khoảng trắng)? rồi "." hoặc ")"
    let j = i + 1;
    while (toks[j] && toks[j].t === "ch" && toks[j].ch === " ") j++;
    if (!toks[j] || toks[j].t !== "ch" || puncts.indexOf(toks[j].ch) < 0) continue;
    labels.push({ letter: tk.ch, start: i, end: j + 1, correct: isUnderlined(tk.rPr) });
    want++;
  }
  return labels.length === 4 ? labels : null;
}

export const isBlank = (tk) => tk.t === "tab" || tk.t === "br" || (tk.t === "ch" && /\s/.test(tk.ch));

/** Dựng các run OOXML từ đoạn token nội dung (gộp ký tự liền nhau cùng rPr). */
function buildRuns(doc, toks) {
  const out = [];
  let cur = null;
  let curPr = undefined;
  const flush = () => {
    if (!cur) return;
    const r = doc.createElementNS(W, "w:r");
    if (curPr) r.appendChild(curPr.cloneNode(true));
    const t = doc.createElementNS(W, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = cur;
    r.appendChild(t);
    out.push(r);
    cur = null;
  };
  for (const tk of toks) {
    if (tk.t === "ch") {
      const ch = tk.ch === "\t" ? " " : tk.ch;
      if (cur !== null && curPr === tk.rPr) cur += ch;
      else {
        flush();
        cur = ch;
        curPr = tk.rPr;
      }
    } else if (tk.t === "node") {
      flush();
      out.push(tk.el);
    } else {
      // tab / br bên trong nội dung phương án -> 1 khoảng trắng
      if (cur !== null) cur += " ";
      else {
        cur = " ";
        curPr = undefined;
      }
    }
  }
  flush();
  return out;
}

function labelRuns(doc, letter, correct) {
  const mk = (text, color, underline) => {
    const r = doc.createElementNS(W, "w:r");
    const pr = doc.createElementNS(W, "w:rPr");
    pr.appendChild(doc.createElementNS(W, "w:b"));
    const c = doc.createElementNS(W, "w:color");
    c.setAttributeNS(W, "w:val", color);
    pr.appendChild(c);
    if (underline) {
      const u = doc.createElementNS(W, "w:u");
      u.setAttributeNS(W, "w:val", "single");
      pr.appendChild(u);
    }
    r.appendChild(pr);
    setFont(doc, r); // nhãn luôn Times New Roman 12, không theo style Normal của tài liệu
    const t = doc.createElementNS(W, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = text;
    r.appendChild(t);
    return r;
  };
  return correct ? [mk(letter, RED, true), mk(". ", BLUE, false)] : [mk(letter + ". ", BLUE, false)];
}

function tabRun(doc) {
  const r = doc.createElementNS(W, "w:r");
  r.appendChild(doc.createElementNS(W, "w:tab"));
  return r;
}

function paraWithTabs(doc, tabsCm) {
  const p = doc.createElementNS(W, "w:p");
  const pPr = standardPPr(doc, { tabsTwip: tabsCm.map((cm) => Math.round(cm * CM)) });
  setFont(doc, pPr); // dấu ¶ cũng Times New Roman 12 (chiều cao dòng trống)
  p.appendChild(pPr);
  return p;
}

/**
 * Dàn lại khối phương án. Trả về { xml, layout, widths } hoặc null nếu không tìm đủ A–D theo thứ tự
 * hoặc có nội dung trước nhãn A (không phải khối phương án thuần).
 */
export function rebuildOptionsXml(pkgXml, forceCorrect = "") {
  const doc = parseXml(pkgXml);
  const body = documentBody(doc);
  const paras = bodyParagraphs(body).filter((p) => p.parentNode === body);
  if (!paras.length) return null;
  const toks = tokenize(paras);
  const labels = findLabels(toks, true) || findLabels(toks, false);
  if (!labels) return null;
  if (toks.slice(0, labels[0].start).some((tk) => !isBlank(tk))) return null;

  const opts = labels.map((lb, k) => {
    let a = lb.end;
    let b = k + 1 < 4 ? labels[k + 1].start : toks.length;
    while (a < b && isBlank(toks[a])) a++;
    while (b > a && isBlank(toks[b - 1])) b--;
    const content = toks.slice(a, b);
    const w = 18 + content.reduce((s, tk) => s + (tk.t === "ch" ? CHAR_PT : tk.t === "node" ? tk.w : CHAR_PT), 0);
    return { ...lb, content, w };
  });
  // Chưa có phương án nào gạch chân nhưng lời giải ghi "Chọn X" -> đánh dấu X là đáp án đúng
  if (forceCorrect && !opts.some((o) => o.correct)) opts.forEach((o) => (o.correct = o.letter === forceCorrect));
  const maxW = Math.max(...opts.map((o) => o.w));
  const layout = maxW <= FOUR_PER_LINE ? 4 : maxW <= TWO_PER_LINE ? 2 : 1;

  const rows = layout === 4 ? [[0, 1, 2, 3]] : layout === 2 ? [[0, 1], [2, 3]] : [[0], [1], [2], [3]];
  const tabs = layout === 4 ? [5.25, 9.5, 13.75] : layout === 2 ? [9.5] : [];
  const newParas = rows.map((row) => {
    const p = paraWithTabs(doc, tabs);
    row.forEach((k, idx) => {
      if (idx > 0) p.appendChild(tabRun(doc));
      for (const r of labelRuns(doc, opts[k].letter, opts[k].correct)) p.appendChild(r);
      for (const r of buildRuns(doc, opts[k].content)) p.appendChild(r);
    });
    return p;
  });
  // thay các đoạn cũ bằng đoạn mới (giữ sectPr nếu có)
  const first = paras[0];
  for (const p of newParas) body.insertBefore(p, first);
  body.insertBefore(sentinelParagraph(doc), first);
  for (const p of paras) body.removeChild(p);
  return { xml: new XMLSerializer().serializeToString(doc), layout, widths: opts.map((o) => Math.round(o.w)), correct: opts.filter((o) => o.correct).map((o) => o.letter) };
}
