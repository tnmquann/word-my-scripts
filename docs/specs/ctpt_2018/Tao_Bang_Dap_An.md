Script giờ kiểm tra số thứ tự câu trong từng phần ngay sau khi đọc đề, trước khi ghi gì vào file. Tôi đã thử logic kiểm tra trên các mẫu số (trùng, đảo lộn, thiếu số, đủ số, tự luận không đánh số) và chạy `dryRun` trên mẫu đang mở. Phần popup chỉ thử bằng hàm giả, chưa thấy popup thật hiện ra trong Word.

**Các trường hợp:**
- **Trùng số là lỗi, script dừng hẳn:**
  - Popup chỉ có nút "Đóng". Chưa ghi gì vào file.
  - Nội dung nêu rõ phần nào, số nào trùng, câu thứ mấy trong phần, kèm 40 ký tự đầu của từng câu.
  - Ví dụ: `Phần II (trả lời ngắn): có 3 câu cùng đánh số "Câu 2" (câu thứ 2, 3, 4 trong phần) — "…"; "…"; "…". Hãy đánh lại số thứ tự câu…`
- **Đảo lộn hoặc thiếu số là cảnh báo, vẫn cho chạy tiếp:**
  - Popup có hai nút "Tiếp tục" và "Dừng lại".
  - Điều kiện cảnh báo: số câu trong một phần không tăng liên tục, mỗi câu hơn câu trước đúng 1.
  - Ví dụ: `1 → 4 → 5 → 2 → 8. Chỗ lệch: Câu 1 → Câu 4; Câu 5 → Câu 2; Câu 2 → Câu 8`.
- **Cách ghi bảng khi đảo lộn:** nếu chọn "Tiếp tục", bảng đáp án ghi đúng số trong đề. Bảng tạo mới xếp tăng dần theo số câu. Bảng có sẵn thì vẫn điền theo nhãn của nó.
- **Phần tự luận không đánh "Câu N." (dùng thẻ `[TL]`):** script bỏ qua kiểm tra. Nếu một câu trong phần có số mà câu khác không đọc được số, script chỉ cảnh báo và đánh theo thứ tự xuất hiện.

Tôi không đòi số câu bắt đầu từ 1, vì có đề đánh số nối tiếp giữa các phần.

**Popup:** script ưu tiên khung hiện ngay trong khung chạy (Script Lab hoặc task pane). Nếu không dựng được thì dùng `confirm`/`alert`, cuối cùng chỉ ghi vào console. Kết quả luôn có trong báo cáo ở `errors` và `numberingWarnings`. Tắt popup bằng `popup: false`.

Trên mẫu đang mở, script báo đúng cảnh báo Phần III: 1 → 2 → 4 (thiếu Câu 3).

```javascript
// ===================== ANSWER-KEY FILLER BUNDLE v7 =====================
const DEFAULTS = {
  answerLabel: /(Đáp án|Đáp số|Trả lời)\s*:/i,
  solutionMarker: /^(Lời giải|Hướng dẫn giải)\s*:?\s*$/i,
  endMarker: /^([\s(\-–—*]*HẾT\b|BẢNG ĐÁP ÁN)/,
  tableAnchor: /^BẢNG ĐÁP ÁN/i,
  tlTag: /^\[TL\]/,
  bold: false, alignment: "Centered",
  createMissingTables: true, ownPage: true, fitTables: true,
  minColWidth: 40, maxColWidth: 80, maxTableWidth: 430,   // point
  popup: true,            // hiện popup khi có lỗi / cảnh báo số thứ tự câu
  notify: null,           // (tuỳ chọn) hàm thay popup: async ({title, lines, buttons}) => tên nút
  dryRun: false,
};
const ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"];
const KIND_VN = { MC: "trắc nghiệm nhiều lựa chọn", TF: "đúng/sai", SA: "trả lời ngắn", ESSAY: "tự luận" };
const toRoman = s => /^\d$/.test(s) ? ROMAN[+s - 1] : s.toUpperCase();
const kindFromHeading = t =>
  /đúng\s*[-/]?\s*sai|đúng hay sai/i.test(t) ? "TF" :
  /trả lời ngắn/i.test(t) ? "SA" :
  /tự luận/i.test(t) ? "ESSAY" :
  /nhiều phương án|nhiều lựa chọn|trắc nghiệm nhiều/i.test(t) ? "MC" : null;
const norm = s => (s || "").toLowerCase().replace(/câu|[\s.\-:()\[\]]/g, "");
const bodyXml = o => (o.match(/<w:body>([\s\S]*?)(?:<w:sectPr|<\/w:body>)/) || [])[1] || "";
const runsOf = x => x.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
const runText = r => (r.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) || []).map(x => x.replace(/<[^>]+>/g, "")).join("");
const hasEquation = x => /<w:object|<m:oMath|<w:fldChar|<w:drawing|<w:pict/.test(x);

// ----- popup: modal trong trang (Script Lab / task pane); dự phòng confirm/alert; cuối cùng chỉ ghi console -----
function showPopup({ title, lines, buttons }) {
  return new Promise(resolve => {
    try {
      if (typeof document !== "undefined" && document.body) {
        const ov = document.createElement("div");
        ov.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:2147483647;display:flex;align-items:center;justify-content:center;font-family:Segoe UI,Arial,sans-serif";
        const box = document.createElement("div");
        box.style.cssText = "background:#fff;max-width:560px;max-height:80vh;overflow:auto;padding:16px 20px;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.35);font-size:13px;line-height:1.45";
        const h = document.createElement("div"); h.textContent = title; h.style.cssText = "font-weight:700;font-size:15px;margin-bottom:8px;color:#a4262c";
        box.appendChild(h);
        lines.forEach(l => { const p = document.createElement("div"); p.textContent = l; p.style.cssText = "margin:6px 0"; box.appendChild(p); });
        const row = document.createElement("div"); row.style.cssText = "margin-top:12px;text-align:right";
        buttons.forEach(name => {
          const b = document.createElement("button"); b.textContent = name; b.style.cssText = "margin-left:8px;padding:5px 14px;cursor:pointer";
          b.onclick = () => { ov.remove(); resolve(name); }; row.appendChild(b);
        });
        box.appendChild(row); ov.appendChild(box); document.body.appendChild(ov); return;
      }
      if (typeof confirm === "function" && buttons.length > 1) {
        resolve(confirm(title + "\n\n" + lines.join("\n\n") + `\n\nOK = ${buttons[0]}, Cancel = ${buttons[1]}`) ? buttons[0] : buttons[1]); return;
      }
      if (typeof alert === "function") { alert(title + "\n\n" + lines.join("\n\n")); resolve(buttons[0]); return; }
    } catch (e) { /* bỏ qua, rơi xuống dự phòng */ }
    console.warn(title, lines); resolve(buttons[0]);
  });
}

// ----- kiểm tra số thứ tự câu trong từng phần -----
// Lỗi: trùng số (vd 3 câu cùng "Câu 2"). Cảnh báo: không tăng liên tục cách nhau 1 (đảo lộn / thiếu số).
function checkNumbering(groups) {
  const errors = [], warns = [];
  for (const g of groups) {
    const head = `Phần ${g.id} (${KIND_VN[g.kind] || g.kind})`;
    const qs = g.qs;
    if (g.kind === "ESSAY" && qs.every(q => !q.num)) continue;          // tự luận không đánh "Câu N." → bỏ qua
    const missing = qs.filter(q => !q.num);
    if (missing.length) { warns.push(`${head}: ${missing.length} câu không đọc được số "Câu N." — bảng đáp án sẽ đánh theo thứ tự xuất hiện.`); continue; }
    const byNum = new Map();
    qs.forEach((q, i) => { if (!byNum.has(q.num)) byNum.set(q.num, []); byNum.get(q.num).push({ pos: i + 1, snip: q.snip }); });
    let dup = false;
    for (const [n, arr] of byNum) if (arr.length > 1) {
      dup = true;
      errors.push(`${head}: có ${arr.length} câu cùng đánh số "Câu ${n}" (câu thứ ${arr.map(a => a.pos).join(", ")} trong phần) — ${arr.map(a => `"${a.snip}"`).join("; ")}.`);
    }
    if (dup) continue;
    const nums = qs.map(q => q.num), breaks = [];
    for (let i = 0; i + 1 < nums.length; i++) if (nums[i + 1] !== nums[i] + 1) breaks.push(`Câu ${nums[i]} → Câu ${nums[i + 1]}`);
    if (breaks.length) warns.push(`${head}: số câu không liên tục tăng dần (cách nhau 1): ${nums.join(" → ")}. Chỗ lệch: ${breaks.join("; ")}.`);
  }
  return { errors, warns };
}

async function ensureJSZip() {
  if (typeof JSZip !== "undefined") return JSZip;
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = res; s.onerror = () => rej(new Error("Cannot load JSZip"));
    document.head.appendChild(s);
  });
  return JSZip;
}
async function buildAnswerDocx(flatOoxml, answerLabel) {
  const Z = await ensureJSZip();
  let flat = flatOoxml.replace(/(<w:body>)(<w:p[ >][\s\S]*?<\/w:p>)/, (m, b, p) => {
    let buf = "", done = false; const kept = [];
    for (const r of runsOf(p)) {
      if (!done) {
        buf += runText(r);
        if (answerLabel.test(buf)) {
          done = true;
          const rest = buf.replace(new RegExp("^[\\s\\S]*?" + answerLabel.source, "i"), "").trim();
          if (rest && !/<w:object/.test(r))
            kept.push(r.replace(/(<w:t(?:\s[^>]*)?>)[^<]*(<\/w:t>)/, `<w:t xml:space="preserve">${rest}</w:t>`));
        }
        continue;
      }
      kept.push(r);
    }
    while (kept.length) {
      const last = kept[kept.length - 1];
      if (!hasEquation(last) && /^[\s.]*$/.test(runText(last))) kept.pop(); else break;
    }
    return `${b}<w:p>${kept.join("")}</w:p>`;
  });
  flat = flat.replace(/(<w:body><w:p>[\s\S]*?<\/w:p>)<w:p[ >][^]*?(?:\/>|<\/w:p>)/, "$1");
  const zip = new Z();
  let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
  for (const part of flat.match(/<pkg:part [\s\S]*?<\/pkg:part>/g) || []) {
    const name = part.match(/pkg:name="([^"]+)"/)[1], type = part.match(/pkg:contentType="([^"]+)"/)[1];
    const bin = part.match(/<pkg:binaryData>([\s\S]*?)<\/pkg:binaryData>/);
    if (bin) zip.file(name.slice(1), bin[1].replace(/\s+/g, ""), { base64: true });
    else zip.file(name.slice(1), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + part.match(/<pkg:xmlData>([\s\S]*)<\/pkg:xmlData>/)[1]);
    if (!/\.rels$/.test(name)) ct += `<Override PartName="${name}" ContentType="${type}"/>`;
  }
  zip.file("[Content_Types].xml", ct + "</Types>");
  return zip.generateAsync({ type: "base64" });
}

// ----- nhận dạng phương án được đánh dấu: gạch chân / highlight / tô nền -----
const runTextT = r => { let s = ""; const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g; let m; while ((m = re.exec(r))) s += m[1] !== undefined ? m[1] : "\t"; return s; };
const isMarked = r => {
  const u = r.match(/<w:u\b([^>]*)>/); if (u && !/w:val="none"/.test(u[1])) return true;
  if (/<w:highlight\b[^>]*w:val="(?!none")[^"]+"/.test(r)) return true;
  const sh = r.match(/<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/); return !!(sh && !/^f{6}$/i.test(sh[1]));
};
function markedLetters(xml, kind) {          // kind "MC": A-D ; "TF": a-d
  const mc = kind === "MC";
  const reL = mc ? /(?:^|[\s\t])([A-D])\./g : /(?:^|[\s\t])([a-d])\)/g;
  const startRe = mc ? /^\s*([A-D])(?:\.|\s|$)/ : /^\s*([a-d])(?:\)|\s|$)/;
  const bareRe = mc ? /^\s*([A-D])\s*$/ : /^\s*([a-d])\s*$/;
  let cur = ""; const found = [];
  for (const run of runsOf(xml)) {
    const txt = runTextT(run);
    if (isMarked(run) && (txt.trim() !== "" || hasEquation(run))) {
      const letter = (txt.match(startRe) || [])[1] || cur;
      if (letter) found.push(letter);
    }
    let m; reL.lastIndex = 0; while ((m = reL.exec(txt))) cur = m[1];
    const bare = txt.match(bareRe); if (bare) cur = bare[1];
  }
  return [...new Set(found)];
}

function analyzeTable(values, labelSet) {
  const seen = new Set(), labels = [];
  values.forEach((row, r) => row.forEach((v, c) => {
    const k = norm(v);
    if (k && labelSet.has(k) && !seen.has(k)) { seen.add(k); labels.push({ r, c, k }); }
  }));
  const vertical = labels.length > 1 && new Set(labels.map(x => x.c)).size === 1;
  labels.forEach(l => {
    if (vertical) { l.ar = l.r; l.ac = l.c + 1; }
    else if (l.r + 1 < values.length) { l.ar = l.r + 1; l.ac = l.c; }
    else { l.ar = l.r; l.ac = l.c + 1; }
  });
  return { labels: labels.filter(l => l.ar < values.length && l.ac < values[l.ar].length), vertical };
}
const baseOf = (q, mode) => mode === "num" ? q.num : q.order;
const makeLab = (g, mode) => {
  const L = {};
  g.qs.forEach(q => Object.entries(q.items).forEach(([suf, it]) => L[String(baseOf(q, mode)) + suf] = it));
  return L;
};
const orderedLabels = (g, mode) =>      // nhãn xếp tăng dần theo số câu (kể cả khi đề đánh số bị đảo)
  [...g.qs].sort((a, b) => baseOf(a, mode) - baseOf(b, mode)).flatMap(q => Object.keys(q.items).map(s => String(baseOf(q, mode)) + s));

async function fillAnswerTable(context, userOpts = {}) {
  const O = { ...DEFAULTS, ...userOpts };
  const report = { groups: [], errors: [], numberingWarnings: [], answers: {}, sources: {}, tables: {}, pageBreak: "", filled: {}, createdTables: [], tableWidths: {}, warnings: [] };
  const body = context.document.body;
  const ask = async (title, lines, buttons) => {
    if (!O.popup) return buttons[0];
    return O.notify ? O.notify({ title, lines, buttons }) : showPopup({ title, lines, buttons });
  };

  const paras = body.paragraphs; paras.load("items/text,items/isListItem");
  await context.sync();
  const P = paras.items;
  const lis = P.map(p => p.isListItem ? p.listItemOrNullObject : null);
  lis.forEach(l => l && l.load("listString,isNullObject"));
  await context.sync();

  // ---------- câu hỏi, gom theo "PHẦN X." ----------
  const isQStart = i => {
    const t = P[i].text.trim();
    return O.tlTag.test(t) || /^Câu\s*\d+\s*[.:]/i.test(t) ||
      !!(lis[i] && !lis[i].isNullObject && /^Câu/i.test(lis[i].listString || ""));
  };
  let stop = P.findIndex(p => O.endMarker.test(p.text.trim()));
  if (stop < 0) stop = P.length;
  const starts = [], headIdAt = [], headKind = {};
  let curId = null;
  for (let i = 0; i < stop; i++) {
    const h = P[i].text.trim().match(/^PHẦN\s+(IV|VIII|VII|VI|V|III|II|I|\d)\b/i);
    if (h && !isQStart(i)) { curId = toRoman(h[1]); headKind[curId] = kindFromHeading(P[i].text); }
    headIdAt[i] = curId;
    if (isQStart(i)) starts.push(i);
  }
  const numOf = i => {
    let m = P[i].text.trim().match(/^Câu\s*(\d+)/i);
    if (!m && lis[i] && !lis[i].isNullObject) m = (lis[i].listString || "").match(/(\d+)/);
    return m ? +m[1] : null;
  };
  const blocks = starts.map((s, k) => {
    const e = k + 1 < starts.length ? starts[k + 1] : stop;
    let sol = -1; for (let i = s; i < e; i++) if (O.solutionMarker.test(P[i].text.trim())) { sol = i; break; }
    const stemEnd = sol >= 0 ? sol : e;
    const stem = P.slice(s, stemEnd).map(p => p.text), joined = stem.join("\n");
    const has  = ch => new RegExp("(^|\\n|\\t)\\s*" + ch + "\\.").test(joined);
    const hasP = ch => new RegExp("(^|\\n|\\t)\\s*" + ch + "\\)").test(joined);
    let fmt;
    if (O.tlTag.test(joined.trim())) fmt = "ESSAY";
    else if (stem.some(t => /^\s*A\.\s/.test(t)) && has("B") && has("C") && has("D")) fmt = "MC";
    else if (hasP("a") && hasP("b") && hasP("c") && hasP("d")) fmt = "TF";
    else fmt = "SA";
    const snip = P[s].text.replace(/^\s*(\[TL\]|Câu\s*\d+\s*[.:])\s*/i, "").trim().slice(0, 40) + "…";
    return { s, e, sol, stemEnd, fmt, id: headIdAt[s], num: numOf(s), snip };
  });
  const groups = [];
  for (const b of blocks) {
    let g = groups[groups.length - 1];
    const same = g && (b.id ? g.id === b.id : (!g.id && g.blocks[g.blocks.length - 1].fmt === b.fmt));
    if (!same) groups.push(g = { id: b.id, blocks: [] });
    g.blocks.push(b);
  }
  groups.forEach(g => {
    const cnt = {}; g.blocks.forEach(b => cnt[b.fmt] = (cnt[b.fmt] || 0) + 1);
    const major = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    g.kind = (g.id && headKind[g.id]) || major;
    g.blocks.forEach(b => {
      if (b.fmt !== g.kind && !(g.kind === "ESSAY" && b.fmt === "SA"))
        report.warnings.push(`Phần ${g.id || "?"}: câu "${P[b.s].text.trim().slice(0, 35)}…" có dạng ${b.fmt} nhưng phần này là ${g.kind}`);
    });
  });
  groups.forEach((g, i) => {
    if (!g.id) { let n = i; while (groups.some(o => o.id === ROMAN[n]) && n < ROMAN.length - 1) n++; g.id = ROMAN[n]; }
  });
  const order = groups.map(g => g.id);
  if (!groups.length) { report.warnings.push('Không tìm thấy câu hỏi nào (dạng "Câu N.")'); return report; }

  // ---------- kiểm tra số thứ tự câu TRƯỚC khi làm gì khác ----------
  groups.forEach(g => g.qs = g.blocks.map((b, qi) => ({ order: qi + 1, num: b.num, snip: b.snip, items: {} })));
  const chk = checkNumbering(groups);
  if (chk.errors.length) {
    report.errors = chk.errors;
    await ask("Lỗi: trùng số thứ tự câu", [...chk.errors, "Hãy đánh lại số thứ tự câu trong đề (mỗi số chỉ dùng một lần trong cùng một phần) rồi chạy lại. Script đã dừng, chưa ghi gì vào file."], ["Đóng"]);
    return report;
  }
  if (chk.warns.length) {
    report.numberingWarnings = chk.warns;
    const choice = await ask("Cảnh báo: số thứ tự câu không liên tục", [...chk.warns, "Script vẫn nhận số câu đúng như trong đề (bảng đáp án ghi theo số đó, xếp tăng dần). Bạn có muốn tiếp tục không?"], ["Tiếp tục", "Dừng lại"]);
    if (choice === "Dừng lại") { report.stoppedByUser = true; return report; }
  }

  // ---------- đáp án ----------
  const tfSolution = b => {                        // "a) Đúng/Sai" trong lời giải
    const res = {};
    for (let i = b.sol + 1; b.sol >= 0 && i < b.e; i++) {
      const m = P[i].text.trim().match(/^([a-d])\)\s*(Đúng|Sai)/i);
      if (m && !res[m[1].toLowerCase()]) res[m[1].toLowerCase()] = m[2][0].toUpperCase() + m[2].slice(1).toLowerCase();
    }
    return res;
  };
  const need = new Map();
  const want = i => { if (!need.has(i)) need.set(i, P[i].getOoxml()); };
  for (const g of groups) for (const b of g.blocks) {
    if (g.kind === "MC") for (let i = b.s; i < b.stemEnd; i++) if (/(^|\t)\s*[A-D]\./.test(P[i].text)) want(i);
    if (g.kind === "TF") {
      b.tf = tfSolution(b);
      if ("abcd".split("").some(k => !b.tf[k]))     // lời giải thiếu Đúng/Sai → xem ý nào được đánh dấu trong đề
        for (let i = b.s; i < b.stemEnd; i++) if (/(^|[\s\t])[a-d]\)/.test(P[i].text)) want(i);
    }
    if (g.kind === "SA" || g.kind === "ESSAY")
      for (let i = Math.max(b.sol, b.s) + 1; i < b.e; i++) if (O.answerLabel.test(P[i].text)) { want(i); break; }
  }
  await context.sync();
  const X = i => bodyXml(need.get(i).value);

  for (const g of groups) {
    const src = { "đánh dấu": 0, "Chọn X": 0, "lời giải": 0, "thiếu": 0 };
    g.blocks.forEach((b, qi) => {
      const q = g.qs[qi], tag = `Phần ${g.id} câu ${b.num || q.order}`;
      if (g.kind === "MC") {
        const marks = [];
        for (let i = b.s; i < b.stemEnd; i++) if (need.has(i)) marks.push(...markedLetters(X(i), "MC"));
        const mk = [...new Set(marks)];
        let chon = "";
        for (let i = Math.max(b.sol, b.s) + 1; i < b.e && !chon; i++) {
          const m = P[i].text.match(/(?:Đáp án|Chọn)\s*:?\s*([A-D])\b/i); if (m) chon = m[1].toUpperCase();
        }
        if (mk.length > 1) report.warnings.push(`${tag}: có ${mk.length} phương án được đánh dấu (${mk.join(",")}) — lấy ${mk[0]}`);
        if (mk.length && chon && mk[0] !== chon) report.warnings.push(`${tag}: đánh dấu = ${mk[0]} nhưng lời giải ghi "Chọn ${chon}" — lấy ${mk[0]}`);
        const found = mk[0] || chon || "";
        if (mk.length) src["đánh dấu"]++; else if (chon) src["Chọn X"]++; else { src["thiếu"]++; report.warnings.push(`${tag}: không thấy đáp án (gạch chân/highlight/"Chọn X")`); }
        q.items[""] = { text: found };
      } else if (g.kind === "TF") {
        let marks = null;
        if ("abcd".split("").some(k => !b.tf[k])) {
          const arr = [];
          for (let i = b.s; i < b.stemEnd; i++) if (need.has(i)) arr.push(...markedLetters(X(i), "TF"));
          marks = new Set(arr);
        }
        for (const k of "abcd") {
          let v = b.tf[k];
          if (v) src["lời giải"]++;
          else if (marks && marks.size) { v = marks.has(k) ? "Đúng" : "Sai"; src["đánh dấu"]++; }   // ý được đánh dấu = Đúng
          else { v = ""; src["thiếu"]++; report.warnings.push(`${tag}${k}: không có "Đúng/Sai" trong lời giải và không thấy ý nào được đánh dấu`); }
          q.items[k] = { text: v };
        }
      } else {
        let pi = -1;
        for (let i = Math.max(b.sol, b.s) + 1; i < b.e; i++) if (O.answerLabel.test(P[i].text)) { pi = i; break; }
        if (pi < 0) { report.warnings.push(`${tag}: không thấy "Đáp án/Đáp số:"`); q.items[""] = { text: "" }; src["thiếu"]++; return; }
        src["lời giải"]++;
        if (hasEquation(X(pi))) q.items[""] = { eqPara: pi, text: "[equation]" };
        else {
          const t = P[pi].text, m = t.match(O.answerLabel);
          q.items[""] = { text: t.slice(t.search(O.answerLabel) + m[0].length).trim().replace(/[.\s]+$/, "") };
        }
      }
    });
    const nums = g.qs.map(q => q.num);
    g.numValid = nums.every(n => n) && new Set(nums).size === nums.length;     // đảo lộn vẫn hợp lệ, trùng thì không
    report.groups.push({ id: g.id, kind: g.kind, questions: g.qs.length, numbers: nums.join(",") });
    report.sources[g.id] = Object.entries(src).filter(([, v]) => v).map(([k, v]) => `${k}: ${v}`).join(", ");
    report.answers[g.id] = g.qs.map(q => `${q.num || q.order}:${Object.values(q.items).map(x => x.text).join("")}`).join(" | ");
  }

  // ---------- tìm bảng đáp án ----------
  const anchorPara = (() => { const i = P.findIndex(p => O.tableAnchor.test(p.text.trim())); return i >= 0 ? P[i] : null; })();
  let anchorRef = anchorPara;
  const refPara = anchorPara || (stop < P.length ? P[stop] : null);
  const tables = body.tables; tables.load("items"); await context.sync();
  const firstQ = P[starts[0]].getRange("Whole");
  const cmp = refPara ? tables.items.map(t => t.getRange("Whole").compareLocationWith(refPara.getRange("Whole"))) : [];
  const cmpFirst = tables.items.map(t => t.getRange("Whole").compareLocationWith(firstQ));
  const befores = tables.items.map(t => { const p = t.getParagraphBeforeOrNullObject(); p.load("text,isNullObject"); return p; });
  tables.items.forEach(t => t.load("values,rowCount"));
  await context.sync();
  const headRe = /^Phần\s*(IV|VIII|VII|VI|V|III|II|I|\d)\s*[.:]?\s*(?:[(\-–:].{0,40})?$/i;
  const headMatch = i => !befores[i].isNullObject && befores[i].text.trim().length < 60 ? befores[i].text.trim().match(headRe) : null;
  const cand = tables.items.map((t, i) => ({ t, i })).filter(x => {
    if (/Before/.test(cmpFirst[x.i].value)) {
      if (headMatch(x.i)) report.warnings.push(`Bỏ qua bảng "${befores[x.i].text.trim()}" nằm trước câu hỏi đầu tiên (có thể là bảng thừa)`);
      return false;
    }
    return refPara ? /After/.test(cmp[x.i].value) : !!headMatch(x.i);
  });
  const tbl = {}, headOf = {}, unmatched = [];
  for (const { t, i } of cand) {
    const m = headMatch(i), id = m && toRoman(m[1]);
    if (id && !tbl[id] && order.includes(id)) { tbl[id] = t; headOf[id] = befores[i]; } else unmatched.push({ t, i });
  }
  for (const { t } of unmatched) { const id = order.find(k => !tbl[k]); if (id) tbl[id] = t; }

  // nhãn theo số "Câu N" trong đề (nếu hợp lệ) hoặc theo thứ tự 1..n — chọn cách khớp bảng đang có
  const lab = {};
  for (const g of groups) {
    const t = tbl[g.id];
    let mode = g.numValid ? "num" : "order";
    if (t) {
      const sc = m => analyzeTable(t.values, new Set(Object.keys(makeLab(g, m)))).labels.length;
      if (g.numValid && sc("order") > sc("num")) mode = "order";
    }
    g.mode = mode; lab[g.id] = makeLab(g, mode);
    const modeTxt = mode === "num" ? "số câu trong đề" : "thứ tự";
    if (!t) { report.tables[g.id] = `chưa có → sẽ tạo (nhãn theo ${modeTxt})`; continue; }
    const a = analyzeTable(t.values, new Set(Object.keys(lab[g.id])));
    report.tables[g.id] = `${t.rowCount}×${Math.max(...t.values.map(r => r.length))}, ${a.vertical ? "dọc" : "ngang"}, khớp ${a.labels.length}/${Object.keys(lab[g.id]).length} nhãn (nhãn theo ${modeTxt})`;
  }

  // ---------- ngắt trang trước vùng đáp án ----------
  const breakTarget = anchorPara || (cand.length ? befores[cand[0].i] : null);
  let breakExists = false;
  if (breakTarget) {
    const prev = breakTarget.getPreviousOrNullObject(); prev.load("text,isNullObject");
    breakTarget.load("text"); await context.sync();
    breakExists = /\f/.test(breakTarget.text) || (!prev.isNullObject && /\f/.test(prev.text));
  }
  report.pageBreak = !O.ownPage ? "bỏ qua" : breakExists ? "đã có ngắt trang"
    : breakTarget ? "sẽ chèn ngắt trang trước vùng đáp án" : "sẽ tạo BẢNG ĐÁP ÁN mới ở cuối file (có ngắt trang)";
  if (O.dryRun) return report;

  if (O.ownPage && breakTarget && !breakExists) { breakTarget.insertBreak("Page", "Before"); await context.sync(); }
  const fmtHead = Object.values(headOf)[0] || null;
  if (fmtHead) fmtHead.load("font/name,font/size,alignment,spaceAfter,spaceBefore,styleBuiltIn");
  await context.sync();

  // ---------- tạo bảng còn thiếu (đúng thứ tự phần trong đề) ----------
  if (O.createMissingTables) {
    for (const g of groups) {
      const part = g.id; if (tbl[part]) continue;
      const labels = orderedLabels(g, g.mode);
      const perRow = g.kind === "TF" ? 8 : 6, cols = Math.min(perRow, labels.length);
      const vals = [];
      for (let r = 0; r < labels.length; r += cols) {
        const row = labels.slice(r, r + cols); while (row.length < cols) row.push("");
        vals.push(row, row.map(() => ""));
      }
      const idx = order.indexOf(part);
      const prevKey = order.slice(0, idx).reverse().find(k => tbl[k]);
      const nextKey = order.slice(idx + 1).find(k => tbl[k] && headOf[k]);
      let head, trailingSpacer = false;
      if (prevKey) { const sp = tbl[prevKey].insertParagraph("", "After"); sp.styleBuiltIn = "Normal"; head = sp.insertParagraph(`Phần ${part}.`, "After"); }
      else if (nextKey) { head = headOf[nextKey].insertParagraph(`Phần ${part}.`, "Before"); trailingSpacer = true; }
      else if (anchorRef) head = anchorRef.insertParagraph(`Phần ${part}.`, "After");
      else {
        const a = body.insertParagraph("BẢNG ĐÁP ÁN", "End"); a.styleBuiltIn = "Normal"; a.font.bold = true; a.alignment = "Centered";
        if (O.ownPage) a.insertBreak("Page", "Before");
        anchorRef = a;
        head = a.insertParagraph(`Phần ${part}.`, "After");
      }
      head.styleBuiltIn = fmtHead ? fmtHead.styleBuiltIn : "Normal";
      if (fmtHead) { head.font.name = fmtHead.font.name; head.font.size = fmtHead.font.size; head.alignment = fmtHead.alignment; head.spaceAfter = fmtHead.spaceAfter; head.spaceBefore = fmtHead.spaceBefore; }
      head.font.bold = true;
      const t = head.insertTable(vals.length, cols, "After", vals);
      t.style = "Table Grid"; t.font.bold = true; t.alignment = "Centered";
      if (fmtHead) { t.font.name = fmtHead.font.name; t.font.size = fmtHead.font.size; }
      if (trailingSpacer) { const sp2 = t.insertParagraph("", "After"); sp2.styleBuiltIn = "Normal"; }
      await context.sync();
      tbl[part] = t; headOf[part] = head; report.createdTables.push(part);
    }
  }

  // ---------- điền đáp án ----------
  const eqCache = {};
  for (const g of groups) {
    const part = g.id, L = lab[part], t = tbl[part];
    if (!t) { report.warnings.push(`Không có bảng đáp án cho Phần ${part}`); continue; }
    t.load("values,rowCount"); await context.sync();
    const { labels } = analyzeTable(t.values, new Set(Object.keys(L)));
    const used = new Set();
    for (const l of labels) {
      const item = L[l.k], cell = t.getCell(l.ar, l.ac);
      cell.body.clear(); await context.sync();
      if (item.eqPara !== undefined) {
        if (!eqCache[item.eqPara]) eqCache[item.eqPara] = await buildAnswerDocx(need.get(item.eqPara).value, O.answerLabel);
        await cell.body.insertFileFromBase64(eqCache[item.eqPara], "End");
      } else if (item.text) cell.body.paragraphs.getFirst().insertText(item.text, "Start");
      cell.body.font.bold = O.bold;
      cell.body.paragraphs.getFirst().alignment = O.alignment;
      await context.sync();
      used.add(l.k);
    }
    report.filled[part] = `${used.size}/${Object.keys(L).length}`;
    Object.keys(L).filter(k => !used.has(k)).forEach(k => report.warnings.push(`Phần ${part}: bảng không có ô nhãn "${k}"`));
  }

  // ---------- thu gọn độ rộng (chỉ các bảng đáp án) ----------
  if (O.fitTables) {
    for (const g of groups) {
      const t = tbl[g.id]; if (!t) continue;
      t.load("rowCount,values"); await context.sync();
      const maxChars = Math.max(1,
        ...Object.values(lab[g.id]).map(it => it.eqPara !== undefined ? 8 : (it.text || "").length),
        ...t.values.flat().map(v => v.trim().length));
      const cols = Math.max(...t.values.map(r => r.length));
      let w = Math.round(Math.min(O.maxColWidth, Math.max(O.minColWidth, maxChars * 7.5 + 20)));
      w = Math.min(w, Math.floor(O.maxTableWidth / cols));
      for (let r = 0; r < t.rowCount; r++)
        for (let c = 0; c < t.values[r].length; c++) t.getCell(r, c).columnWidth = w;
      t.alignment = "Centered";
      await context.sync();
      try {
        t.load("width"); await context.sync();
        if (t.width > cols * w + 6) { t.width = cols * w; await context.sync(); }
      } catch (e) { report.warnings.push(`Phần ${g.id}: không ép được độ rộng bảng (${e.message})`); }
      report.tableWidths[g.id] = `${cols} cột × ${w}pt`;
    }
  }
  return report;
}

// ---- chạy: nên chạy dryRun: true trước, đọc "groups" / "errors" / "numberingWarnings" / "tables" / "warnings" rồi mới chạy thật ----
Word.run(ctx => fillAnswerTable(ctx, { dryRun: true }))
  .then(r => console.log(JSON.stringify(r, null, 2)))
  .catch(e => console.error(e));
```

Môi trường chạy cần có JSZip nếu đề có đáp án MathType. Nếu thiếu, script thử tải từ CDN và có thể bị chặn, nên bạn hãy nạp JSZip sẵn.