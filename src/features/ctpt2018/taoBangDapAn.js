// Tạo bảng đáp án (CTPT 2018)  (spec: docs/specs/ctpt_2018/Tao_Bang_Dap_An.md)
// Port của "ANSWER-KEY FILLER BUNDLE v7" (bản Script Lab đã test). Thay đổi khi đưa vào add-in:
//   - Popup DOM / confirm / alert (không hiện được trong ribbon command) -> askUser (hộp thoại Office)
//   - JSZip được đóng gói sẵn (npm) thay vì tải từ CDN
//   - Kết quả gom thành popup tóm tắt; lỗi trùng số câu chỉ hiện 1 popup lỗi
// Phần nhận dạng câu hỏi / đáp án / bảng giữ nguyên logic v7.

import JSZip from "jszip";
import { askUser } from "../runtime.js";

const DEFAULTS = {
  answerLabel: /(Đáp án|Đáp số|Trả lời)\s*:/i,
  solutionMarker: /^(Lời giải|Hướng dẫn giải)\s*:?\s*$/i,
  endMarker: /^([\s(\-–—*]*HẾT\b|BẢNG ĐÁP ÁN)/,
  tableAnchor: /^BẢNG ĐÁP ÁN/i,
  tlTag: /^\[TL\]/,
  bold: false,
  alignment: "Centered",
  createMissingTables: true, // bảng còn thiếu được tạo tại vị trí con trỏ (không ngắt trang)
  fitTables: true,
  minColWidth: 40,
  maxColWidth: 80,
  maxTableWidth: 430, // point
  popup: true, // hỏi khi số thứ tự câu không liên tục
  dryRun: false,
};
const KIND_VN = { MC: "trắc nghiệm nhiều lựa chọn", TF: "đúng/sai", SA: "trả lời ngắn", ESSAY: "tự luận" };

// ----- số thứ tự "Phần": La Mã (I, IV, IX, XVI, …) hoặc số thường (1, 2, 11, …) -----
// Mỗi đề dùng thống nhất 1 kiểu; số được quy về số nguyên để so khớp, rồi hiển thị lại đúng kiểu của đề.
// Sau số phần phải là hết dòng / khoảng trắng / dấu câu — không dùng \b vì JS coi "ở", "đ"… là ký tự
// không thuộc chữ, khiến "PHẦN Mở đầu" bị hiểu thành Phần M (1000).
const PART_TOKEN = "([IVXLCDM]+|\\d+)(?=$|[\\s.:,;)(\\-–—])";
const MAX_PART = 50; // đề thực tế không có quá 50 phần; chặn "Phần C" (=100)…
const ROMAN_VALID = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
const ROMAN_VAL = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
/** "IV" | "11" -> { n: 4 | 11, style: "roman" | "arabic" } ; không hợp lệ (vd "IIII", "C" trong "cuối") -> null */
function parsePart(tok) {
  if (/^\d+$/.test(tok)) return +tok > 0 && +tok <= MAX_PART ? { n: +tok, style: "arabic" } : null;
  const t = tok.toUpperCase();
  if (!t || !ROMAN_VALID.test(t)) return null;
  let n = 0;
  for (let i = 0; i < t.length; i++) {
    const v = ROMAN_VAL[t[i]];
    n += i + 1 < t.length && v < ROMAN_VAL[t[i + 1]] ? -v : v;
  }
  return n <= MAX_PART ? { n, style: "roman" } : null;
}
function toRomanNum(n) {
  const map = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let s = "";
  for (const [v, r] of map) while (n >= v) (s += r), (n -= v);
  return s;
}
const fmtPart = (n, style) => (style === "arabic" ? String(n) : toRomanNum(n));
const kindFromHeading = (t) =>
  /đúng\s*[-/]?\s*sai|đúng hay sai/i.test(t)
    ? "TF"
    : /trả lời ngắn/i.test(t)
      ? "SA"
      : /tự luận/i.test(t)
        ? "ESSAY"
        : /nhiều phương án|nhiều lựa chọn|trắc nghiệm nhiều/i.test(t)
          ? "MC"
          : null;
const norm = (s) => (s || "").toLowerCase().replace(/câu|[\s.\-:()\[\]]/g, "");
const bodyXml = (o) => (o.match(/<w:body>([\s\S]*?)(?:<w:sectPr|<\/w:body>)/) || [])[1] || "";
const runsOf = (x) => x.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
const runText = (r) =>
  (r.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("");
const hasEquation = (x) => /<w:object|<m:oMath|<w:fldChar|<w:drawing|<w:pict/.test(x);

// ----- kiểm tra số thứ tự câu trong từng phần -----
// Lỗi: trùng số (vd 3 câu cùng "Câu 2"). Cảnh báo: không tăng liên tục cách nhau 1 (đảo lộn / thiếu số).
function checkNumbering(groups) {
  const errors = [];
  const warns = [];
  for (const g of groups) {
    const head = `Phần ${g.id} (${KIND_VN[g.kind] || g.kind})`;
    const qs = g.qs;
    if (g.kind === "ESSAY" && qs.every((q) => !q.num)) continue; // tự luận không đánh "Câu N." -> bỏ qua
    const missing = qs.filter((q) => !q.num);
    if (missing.length) {
      warns.push(`${head}: ${missing.length} câu không đọc được số "Câu N." — bảng đáp án sẽ đánh theo thứ tự xuất hiện.`);
      continue;
    }
    const byNum = new Map();
    qs.forEach((q, i) => {
      if (!byNum.has(q.num)) byNum.set(q.num, []);
      byNum.get(q.num).push({ pos: i + 1, snip: q.snip });
    });
    let dup = false;
    for (const [n, arr] of byNum)
      if (arr.length > 1) {
        dup = true;
        errors.push(
          `${head}: có ${arr.length} câu cùng đánh số "Câu ${n}" (câu thứ ${arr.map((a) => a.pos).join(", ")} trong phần) — ${arr
            .map((a) => `"${a.snip}"`)
            .join("; ")}.`
        );
      }
    if (dup) continue;
    const nums = qs.map((q) => q.num);
    const breaks = [];
    for (let i = 0; i + 1 < nums.length; i++) if (nums[i + 1] !== nums[i] + 1) breaks.push(`Câu ${nums[i]} → Câu ${nums[i + 1]}`);
    if (breaks.length) warns.push(`${head}: số câu không liên tục tăng dần (cách nhau 1): ${nums.join(" → ")}. Chỗ lệch: ${breaks.join("; ")}.`);
  }
  return { errors, warns };
}

// ----- đáp án dạng công thức (MathType/OMML): dựng .docx chứa đoạn đáp án để chèn vào ô bảng -----
async function buildAnswerDocx(flatOoxml, answerLabel) {
  let flat = flatOoxml.replace(/(<w:body>)(<w:p[ >][\s\S]*?<\/w:p>)/, (m, b, p) => {
    let buf = "";
    let done = false;
    const kept = [];
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
      if (!hasEquation(last) && /^[\s.]*$/.test(runText(last))) kept.pop();
      else break;
    }
    return `${b}<w:p>${kept.join("")}</w:p>`;
  });
  flat = flat.replace(/(<w:body><w:p>[\s\S]*?<\/w:p>)<w:p[ >][^]*?(?:\/>|<\/w:p>)/, "$1");
  const zip = new JSZip();
  let ct =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
  for (const part of flat.match(/<pkg:part [\s\S]*?<\/pkg:part>/g) || []) {
    const name = part.match(/pkg:name="([^"]+)"/)[1];
    const type = part.match(/pkg:contentType="([^"]+)"/)[1];
    const bin = part.match(/<pkg:binaryData>([\s\S]*?)<\/pkg:binaryData>/);
    if (bin) zip.file(name.slice(1), bin[1].replace(/\s+/g, ""), { base64: true });
    else
      zip.file(
        name.slice(1),
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + part.match(/<pkg:xmlData>([\s\S]*)<\/pkg:xmlData>/)[1]
      );
    if (!/\.rels$/.test(name)) ct += `<Override PartName="${name}" ContentType="${type}"/>`;
  }
  zip.file("[Content_Types].xml", ct + "</Types>");
  return zip.generateAsync({ type: "base64" });
}

// ----- nhận dạng phương án được đánh dấu: gạch chân / highlight / tô nền -----
const runTextT = (r) => {
  let s = "";
  const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g;
  let m;
  while ((m = re.exec(r))) s += m[1] !== undefined ? m[1] : "\t";
  return s;
};
const isMarked = (r) => {
  const u = r.match(/<w:u\b([^>]*)>/);
  if (u && !/w:val="none"/.test(u[1])) return true;
  if (/<w:highlight\b[^>]*w:val="(?!none")[^"]+"/.test(r)) return true;
  const sh = r.match(/<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/);
  return !!(sh && !/^f{6}$/i.test(sh[1]));
};
function markedLetters(xml, kind) {
  // kind "MC": A-D ; "TF": a-d
  const mc = kind === "MC";
  const reL = mc ? /(?:^|[\s\t])([A-D])\./g : /(?:^|[\s\t])([a-d])\)/g;
  const startRe = mc ? /^\s*([A-D])(?:\.|\s|$)/ : /^\s*([a-d])(?:\)|\s|$)/;
  const bareRe = mc ? /^\s*([A-D])\s*$/ : /^\s*([a-d])\s*$/;
  let cur = "";
  const found = [];
  for (const run of runsOf(xml)) {
    const txt = runTextT(run);
    if (isMarked(run) && (txt.trim() !== "" || hasEquation(run))) {
      const letter = (txt.match(startRe) || [])[1] || cur;
      if (letter) found.push(letter);
    }
    let m;
    reL.lastIndex = 0;
    while ((m = reL.exec(txt))) cur = m[1];
    const bare = txt.match(bareRe);
    if (bare) cur = bare[1];
  }
  return [...new Set(found)];
}

function analyzeTable(values, labelSet) {
  const seen = new Set();
  const labels = [];
  values.forEach((row, r) =>
    row.forEach((v, c) => {
      const k = norm(v);
      if (k && labelSet.has(k) && !seen.has(k)) {
        seen.add(k);
        labels.push({ r, c, k });
      }
    })
  );
  const vertical = labels.length > 1 && new Set(labels.map((x) => x.c)).size === 1;
  labels.forEach((l) => {
    if (vertical) {
      l.ar = l.r;
      l.ac = l.c + 1;
    } else if (l.r + 1 < values.length) {
      l.ar = l.r + 1;
      l.ac = l.c;
    } else {
      l.ar = l.r;
      l.ac = l.c + 1;
    }
  });
  return { labels: labels.filter((l) => l.ar < values.length && l.ac < values[l.ar].length), vertical };
}
const baseOf = (q, mode) => (mode === "num" ? q.num : q.order);
const makeLab = (g, mode) => {
  const L = {};
  g.qs.forEach((q) => Object.entries(q.items).forEach(([suf, it]) => (L[String(baseOf(q, mode)) + suf] = it)));
  return L;
};
const orderedLabels = (
  g,
  mode // nhãn xếp tăng dần theo số câu (kể cả khi đề đánh số bị đảo)
) =>
  [...g.qs]
    .sort((a, b) => baseOf(a, mode) - baseOf(b, mode))
    .flatMap((q) => Object.keys(q.items).map((s) => String(baseOf(q, mode)) + s));

export async function fillAnswerTable(context, userOpts = {}) {
  const O = { ...DEFAULTS, ...userOpts };
  const report = {
    groups: [],
    errors: [],
    numberingWarnings: [],
    answers: {},
    sources: {},
    tables: {},
    filled: {},
    createdTables: [],
    tableWidths: {},
    warnings: [],
  };
  const body = context.document.body;

  const paras = body.paragraphs;
  paras.load("items/text,items/isListItem");
  await context.sync();
  const P = paras.items;
  const lis = P.map((p) => (p.isListItem ? p.listItemOrNullObject : null));
  lis.forEach((l) => l && l.load("listString,isNullObject"));
  await context.sync();

  // ---------- câu hỏi, gom theo "PHẦN X." ----------
  const isQStart = (i) => {
    const t = P[i].text.trim();
    return (
      O.tlTag.test(t) || /^Câu\s*\d+\s*[.:]/i.test(t) || !!(lis[i] && !lis[i].isNullObject && /^Câu/i.test(lis[i].listString || ""))
    );
  };
  let stop = P.findIndex((p) => O.endMarker.test(p.text.trim()));
  if (stop < 0) stop = P.length;
  const starts = [];
  const headIdAt = [];
  const headKind = {};
  // Tiêu đề phần: "PHẦN I.", "Phần IV:", "PHẦN 2.", "Phần 11 - …" (số La Mã hợp lệ hoặc số thường)
  const headingRe = new RegExp("^PHẦN\\s+" + PART_TOKEN, "i");
  const headingAt = (i) => {
    if (isQStart(i)) return null;
    const h = P[i].text.trim().match(headingRe);
    return h ? parsePart(h[1]) : null;
  };
  // Kiểu đánh số của đề (thống nhất): theo tiêu đề phần đầu tiên; mặc định La Mã
  const styles = [];
  for (let i = 0; i < stop; i++) {
    const p = headingAt(i);
    if (p) styles.push(p.style);
  }
  const partStyle = styles[0] || "roman";
  if (new Set(styles).size > 1)
    report.warnings.push(`Đề dùng lẫn số La Mã và số thường cho "Phần"; bảng mới ghi theo kiểu ${partStyle === "roman" ? "La Mã" : "số thường"}.`);
  report.partStyle = partStyle;
  const label = (n) => fmtPart(n, partStyle);

  let curId = null;
  for (let i = 0; i < stop; i++) {
    const p = headingAt(i);
    if (p) {
      curId = label(p.n);
      headKind[curId] = kindFromHeading(P[i].text);
    }
    headIdAt[i] = curId;
    if (isQStart(i)) starts.push(i);
  }
  const numOf = (i) => {
    let m = P[i].text.trim().match(/^Câu\s*(\d+)/i);
    if (!m && lis[i] && !lis[i].isNullObject) m = (lis[i].listString || "").match(/(\d+)/);
    return m ? +m[1] : null;
  };
  const blocks = starts.map((s, k) => {
    const e = k + 1 < starts.length ? starts[k + 1] : stop;
    let sol = -1;
    for (let i = s; i < e; i++)
      if (O.solutionMarker.test(P[i].text.trim())) {
        sol = i;
        break;
      }
    const stemEnd = sol >= 0 ? sol : e;
    const stem = P.slice(s, stemEnd).map((p) => p.text);
    const joined = stem.join("\n");
    const has = (ch) => new RegExp("(^|\\n|\\t)\\s*" + ch + "\\.").test(joined);
    const hasP = (ch) => new RegExp("(^|\\n|\\t)\\s*" + ch + "\\)").test(joined);
    let fmt;
    if (O.tlTag.test(joined.trim())) fmt = "ESSAY";
    else if (stem.some((t) => /^\s*A\.\s/.test(t)) && has("B") && has("C") && has("D")) fmt = "MC";
    else if (hasP("a") && hasP("b") && hasP("c") && hasP("d")) fmt = "TF";
    else fmt = "SA";
    const snip =
      P[s].text
        .replace(/^\s*(\[TL\]|Câu\s*\d+\s*[.:])\s*/i, "")
        .trim()
        .slice(0, 40) + "…";
    return { s, e, sol, stemEnd, fmt, id: headIdAt[s], num: numOf(s), snip };
  });
  const groups = [];
  for (const b of blocks) {
    let g = groups[groups.length - 1];
    const same = g && (b.id ? g.id === b.id : !g.id && g.blocks[g.blocks.length - 1].fmt === b.fmt);
    if (!same) groups.push((g = { id: b.id, blocks: [] }));
    g.blocks.push(b);
  }
  groups.forEach((g) => {
    const cnt = {};
    g.blocks.forEach((b) => (cnt[b.fmt] = (cnt[b.fmt] || 0) + 1));
    const major = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    g.kind = (g.id && headKind[g.id]) || major;
    g.blocks.forEach((b) => {
      if (b.fmt !== g.kind && !(g.kind === "ESSAY" && b.fmt === "SA"))
        report.warnings.push(`Phần ${g.id || "?"}: câu "${P[b.s].text.trim().slice(0, 35)}…" có dạng ${b.fmt} nhưng phần này là ${g.kind}`);
    });
  });
  // Nhóm không có tiêu đề "Phần": gán số kế tiếp chưa dùng (theo kiểu đánh số của đề)
  groups.forEach((g, i) => {
    if (!g.id) {
      let n = i + 1;
      while (groups.some((o) => o.id === label(n))) n++;
      g.id = label(n);
    }
  });
  const order = groups.map((g) => g.id);
  if (!groups.length) {
    report.warnings.push('Không tìm thấy câu hỏi nào (dạng "Câu N.")');
    return report;
  }

  // ---------- kiểm tra số thứ tự câu TRƯỚC khi làm gì khác ----------
  groups.forEach((g) => (g.qs = g.blocks.map((b, qi) => ({ order: qi + 1, num: b.num, snip: b.snip, items: {} }))));
  const chk = checkNumbering(groups);
  if (chk.errors.length) {
    report.errors = chk.errors; // dừng, chưa ghi gì; popup lỗi do taoBangDapAn hiển thị
    return report;
  }
  if (chk.warns.length) {
    report.numberingWarnings = chk.warns;
    if (O.popup && !O.dryRun) {
      const choice = await askUser({
        title: "Số thứ tự câu không liên tục",
        lines: [
          ...chk.warns,
          "Bảng đáp án sẽ ghi đúng số câu trong đề (bảng mới xếp tăng dần). Bạn có muốn tiếp tục không?",
        ],
        buttons: ["Tiếp tục", "Dừng lại"],
      });
      if (choice !== "Tiếp tục") {
        report.stoppedByUser = true;
        return report;
      }
    }
  }

  // ---------- đáp án ----------
  const tfSolution = (b) => {
    // "a) Đúng/Sai" trong lời giải
    const res = {};
    for (let i = b.sol + 1; b.sol >= 0 && i < b.e; i++) {
      const m = P[i].text.trim().match(/^([a-d])\)\s*(Đúng|Sai)/i);
      if (m && !res[m[1].toLowerCase()]) res[m[1].toLowerCase()] = m[2][0].toUpperCase() + m[2].slice(1).toLowerCase();
    }
    return res;
  };
  const need = new Map();
  const want = (i) => {
    if (!need.has(i)) need.set(i, P[i].getOoxml());
  };
  for (const g of groups)
    for (const b of g.blocks) {
      if (g.kind === "MC") for (let i = b.s; i < b.stemEnd; i++) if (/(^|\t)\s*[A-D]\./.test(P[i].text)) want(i);
      if (g.kind === "TF") {
        b.tf = tfSolution(b);
        if ("abcd".split("").some((k) => !b.tf[k]))
          // lời giải thiếu Đúng/Sai -> xem ý nào được đánh dấu trong đề
          for (let i = b.s; i < b.stemEnd; i++) if (/(^|[\s\t])[a-d]\)/.test(P[i].text)) want(i);
      }
      if (g.kind === "SA" || g.kind === "ESSAY")
        for (let i = Math.max(b.sol, b.s) + 1; i < b.e; i++)
          if (O.answerLabel.test(P[i].text)) {
            want(i);
            break;
          }
    }
  await context.sync();
  const X = (i) => bodyXml(need.get(i).value);

  for (const g of groups) {
    const src = { "đánh dấu": 0, "Chọn X": 0, "lời giải": 0, thiếu: 0 };
    g.blocks.forEach((b, qi) => {
      const q = g.qs[qi];
      const tag = `Phần ${g.id} câu ${b.num || q.order}`;
      if (g.kind === "MC") {
        const marks = [];
        for (let i = b.s; i < b.stemEnd; i++) if (need.has(i)) marks.push(...markedLetters(X(i), "MC"));
        const mk = [...new Set(marks)];
        let chon = "";
        for (let i = Math.max(b.sol, b.s) + 1; i < b.e && !chon; i++) {
          const m = P[i].text.match(/(?:Đáp án|Chọn)\s*:?\s*([A-D])\b/i);
          if (m) chon = m[1].toUpperCase();
        }
        if (mk.length > 1) report.warnings.push(`${tag}: có ${mk.length} phương án được đánh dấu (${mk.join(",")}) — lấy ${mk[0]}`);
        if (mk.length && chon && mk[0] !== chon)
          report.warnings.push(`${tag}: đánh dấu = ${mk[0]} nhưng lời giải ghi "Chọn ${chon}" — lấy ${mk[0]}`);
        const found = mk[0] || chon || "";
        if (mk.length) src["đánh dấu"]++;
        else if (chon) src["Chọn X"]++;
        else {
          src["thiếu"]++;
          report.warnings.push(`${tag}: không thấy đáp án (gạch chân/highlight/"Chọn X")`);
        }
        q.items[""] = { text: found };
      } else if (g.kind === "TF") {
        let marks = null;
        if ("abcd".split("").some((k) => !b.tf[k])) {
          const arr = [];
          for (let i = b.s; i < b.stemEnd; i++) if (need.has(i)) arr.push(...markedLetters(X(i), "TF"));
          marks = new Set(arr);
        }
        for (const k of "abcd") {
          let v = b.tf[k];
          if (v) src["lời giải"]++;
          else if (marks && marks.size) {
            v = marks.has(k) ? "Đúng" : "Sai"; // ý được đánh dấu = Đúng
            src["đánh dấu"]++;
          } else {
            v = "";
            src["thiếu"]++;
            report.warnings.push(`${tag}${k}: không có "Đúng/Sai" trong lời giải và không thấy ý nào được đánh dấu`);
          }
          q.items[k] = { text: v };
        }
      } else {
        let pi = -1;
        for (let i = Math.max(b.sol, b.s) + 1; i < b.e; i++)
          if (O.answerLabel.test(P[i].text)) {
            pi = i;
            break;
          }
        if (pi < 0) {
          report.warnings.push(`${tag}: không thấy "Đáp án/Đáp số:"`);
          q.items[""] = { text: "" };
          src["thiếu"]++;
          return;
        }
        src["lời giải"]++;
        if (hasEquation(X(pi))) q.items[""] = { eqPara: pi, text: "[equation]" };
        else {
          const t = P[pi].text;
          const m = t.match(O.answerLabel);
          q.items[""] = { text: t.slice(t.search(O.answerLabel) + m[0].length).trim().replace(/[.\s]+$/, "") };
        }
      }
    });
    const nums = g.qs.map((q) => q.num);
    g.numValid = nums.every((n) => n) && new Set(nums).size === nums.length; // đảo lộn vẫn hợp lệ, trùng thì không
    report.groups.push({ id: g.id, kind: g.kind, questions: g.qs.length, numbers: nums.join(",") });
    report.sources[g.id] = Object.entries(src)
      .filter(([, v]) => v)
      .map(([k, v]) => `${k}: ${v}`)
      .join(", ");
    report.answers[g.id] = g.qs
      .map(
        (q) =>
          `${q.num || q.order}:${Object.values(q.items)
            .map((x) => x.text)
            .join("")}`
      )
      .join(" | ");
  }

  // ---------- tìm bảng đáp án ----------
  const anchorPara = (() => {
    const i = P.findIndex((p) => O.tableAnchor.test(p.text.trim()));
    return i >= 0 ? P[i] : null;
  })();

  const refPara = anchorPara || (stop < P.length ? P[stop] : null);
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  const firstQ = P[starts[0]].getRange("Whole");
  const cmp = refPara ? tables.items.map((t) => t.getRange("Whole").compareLocationWith(refPara.getRange("Whole"))) : [];
  const cmpFirst = tables.items.map((t) => t.getRange("Whole").compareLocationWith(firstQ));
  const befores = tables.items.map((t) => {
    const p = t.getParagraphBeforeOrNullObject();
    p.load("text,isNullObject");
    return p;
  });
  tables.items.forEach((t) => t.load("values,rowCount"));
  await context.sync();
  // Tiêu đề ngay trên bảng: "Phần I.", "Phần 2:", "Phần XVI (…)" -> số phần (null nếu không phải)
  const headRe = new RegExp("^Phần\\s*" + PART_TOKEN + "\\s*[.:]?\\s*(?:[(\\-–:].{0,40})?$", "i");
  const headMatch = (i) => {
    if (befores[i].isNullObject || befores[i].text.trim().length >= 60) return null;
    const m = befores[i].text.trim().match(headRe);
    return m ? parsePart(m[1]) : null;
  };
  const cand = tables.items
    .map((t, i) => ({ t, i }))
    .filter((x) => {
      if (/Before/.test(cmpFirst[x.i].value)) {
        if (headMatch(x.i)) report.warnings.push(`Bỏ qua bảng "${befores[x.i].text.trim()}" nằm trước câu hỏi đầu tiên (có thể là bảng thừa)`);
        return false;
      }
      return refPara ? /After/.test(cmp[x.i].value) : !!headMatch(x.i);
    });
  const tbl = {};
  const headOf = {};
  const unmatched = [];
  for (const { t, i } of cand) {
    const p = headMatch(i);
    const id = p && label(p.n); // khớp theo số, kể cả bảng ghi "Phần 2" còn đề ghi "PHẦN II"
    if (id && !tbl[id] && order.includes(id)) {
      tbl[id] = t;
      headOf[id] = befores[i];
    } else unmatched.push({ t, i });
  }
  for (const { t } of unmatched) {
    const id = order.find((k) => !tbl[k]);
    if (id) tbl[id] = t;
  }

  // nhãn theo số "Câu N" trong đề (nếu hợp lệ) hoặc theo thứ tự 1..n — chọn cách khớp bảng đang có
  const lab = {};
  for (const g of groups) {
    const t = tbl[g.id];
    let mode = g.numValid ? "num" : "order";
    if (t) {
      const sc = (m) => analyzeTable(t.values, new Set(Object.keys(makeLab(g, m)))).labels.length;
      if (g.numValid && sc("order") > sc("num")) mode = "order";
    }
    g.mode = mode;
    lab[g.id] = makeLab(g, mode);
    const modeTxt = mode === "num" ? "số câu trong đề" : "thứ tự";
    if (!t) {
      report.tables[g.id] = `chưa có → sẽ tạo (nhãn theo ${modeTxt})`;
      continue;
    }
    const a = analyzeTable(t.values, new Set(Object.keys(lab[g.id])));
    report.tables[g.id] = `${t.rowCount}×${Math.max(...t.values.map((r) => r.length))}, ${a.vertical ? "dọc" : "ngang"}, khớp ${
      a.labels.length
    }/${Object.keys(lab[g.id]).length} nhãn (nhãn theo ${modeTxt})`;
  }

  // ---------- vị trí tạo bảng còn thiếu: con trỏ của người dùng (không ngắt trang) ----------
  const missing = O.createMissingTables ? order.filter((k) => !tbl[k]) : [];
  let cursorPara = null;
  let cursorEmpty = false;
  if (missing.length) {
    const sel = context.document.getSelection();
    const selTable = sel.parentTableOrNullObject;
    selTable.load("isNullObject");
    cursorPara = sel.paragraphs.getFirst();
    cursorPara.load("text");
    await context.sync();
    if (!selTable.isNullObject) {
      // trong bảng: chèn sẽ rơi vào ô -> dừng trước khi ghi gì
      report.cursorInTable = true;
      return report;
    }
    cursorEmpty = cursorPara.text.trim() === "";
    report.insertAt = cursorEmpty ? "dòng trống tại con trỏ" : "ngay sau đoạn có con trỏ";
  }
  if (O.dryRun) return report;

  const fmtHead = Object.values(headOf)[0] || null;
  if (fmtHead) fmtHead.load("font/name,font/size,alignment,spaceAfter,spaceBefore,styleBuiltIn");
  await context.sync();

  // ---------- tạo bảng còn thiếu tại con trỏ (đúng thứ tự phần trong đề) ----------
  // Bảng đầu tiên: thay dòng trống tại con trỏ / hoặc ngay sau đoạn có con trỏ; các bảng sau nối tiếp bảng trước.
  let lastCreated = null;
  const trailingSpacer = missing.length > 0 && !cursorEmpty;
  if (missing.length) {
    for (const g of groups) {
      const part = g.id;
      if (tbl[part]) continue;
      const labels = orderedLabels(g, g.mode);
      const perRow = g.kind === "TF" ? 8 : 6;
      const cols = Math.min(perRow, labels.length);
      const vals = [];
      for (let r = 0; r < labels.length; r += cols) {
        const row = labels.slice(r, r + cols);
        while (row.length < cols) row.push("");
        vals.push(row, row.map(() => ""));
      }
      let head;
      if (lastCreated) {
        const sp = lastCreated.insertParagraph("", "After");
        sp.styleBuiltIn = "Normal";
        head = sp.insertParagraph(`Phần ${part}.`, "After");
      } else if (cursorEmpty) {
        head = cursorPara.insertParagraph(`Phần ${part}.`, "Before"); // dòng trống ở con trỏ giữ lại làm khoảng cách
      } else {
        head = cursorPara.insertParagraph(`Phần ${part}.`, "After");
      }
      head.styleBuiltIn = fmtHead ? fmtHead.styleBuiltIn : "Normal";
      if (fmtHead) {
        head.font.name = fmtHead.font.name;
        head.font.size = fmtHead.font.size;
        head.alignment = fmtHead.alignment;
        head.spaceAfter = fmtHead.spaceAfter;
        head.spaceBefore = fmtHead.spaceBefore;
      }
      head.font.bold = true;
      const t = head.insertTable(vals.length, cols, "After", vals);
      t.style = "Table Grid";
      t.font.bold = true;
      t.alignment = "Centered";
      if (fmtHead) {
        t.font.name = fmtHead.font.name;
        t.font.size = fmtHead.font.size;
      }
      await context.sync();
      tbl[part] = t;
      headOf[part] = head;
      lastCreated = t;
      report.createdTables.push(part);
    }
    // con trỏ ở giữa nội dung -> thêm 1 dòng trống ngăn bảng cuối với nội dung phía sau
    if (trailingSpacer && lastCreated) {
      const sp2 = lastCreated.insertParagraph("", "After");
      sp2.styleBuiltIn = "Normal";
      await context.sync();
    }
  }

  // ---------- điền đáp án ----------
  const eqCache = {};
  for (const g of groups) {
    const part = g.id;
    const L = lab[part];
    const t = tbl[part];
    if (!t) {
      report.warnings.push(`Không có bảng đáp án cho Phần ${part}`);
      continue;
    }
    t.load("values,rowCount");
    await context.sync();
    const { labels } = analyzeTable(t.values, new Set(Object.keys(L)));
    const used = new Set();
    for (const l of labels) {
      const item = L[l.k];
      const cell = t.getCell(l.ar, l.ac);
      cell.body.clear();
      await context.sync();
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
    Object.keys(L)
      .filter((k) => !used.has(k))
      .forEach((k) => report.warnings.push(`Phần ${part}: bảng không có ô nhãn "${k}"`));
  }

  // ---------- thu gọn độ rộng (chỉ các bảng đáp án) ----------
  if (O.fitTables) {
    for (const g of groups) {
      const t = tbl[g.id];
      if (!t) continue;
      t.load("rowCount,values");
      await context.sync();
      const maxChars = Math.max(
        1,
        ...Object.values(lab[g.id]).map((it) => (it.eqPara !== undefined ? 8 : (it.text || "").length)),
        ...t.values.flat().map((v) => v.trim().length)
      );
      const cols = Math.max(...t.values.map((r) => r.length));
      let w = Math.round(Math.min(O.maxColWidth, Math.max(O.minColWidth, maxChars * 7.5 + 20)));
      w = Math.min(w, Math.floor(O.maxTableWidth / cols));
      for (let r = 0; r < t.rowCount; r++) for (let c = 0; c < t.values[r].length; c++) t.getCell(r, c).columnWidth = w;
      t.alignment = "Centered";
      await context.sync();
      try {
        t.load("width");
        await context.sync();
        if (t.width > cols * w + 6) {
          t.width = cols * w;
          await context.sync();
        }
      } catch (e) {
        report.warnings.push(`Phần ${g.id}: không ép được độ rộng bảng (${e.message})`);
      }
      report.tableWidths[g.id] = `${cols} cột × ${w}pt`;
    }
  }
  return report;
}

// ---------- ribbon command: chạy + tóm tắt kết quả cho popup ----------
const short = (s, n = 110) => (s.length > n ? s.slice(0, n) + "…" : s);

export function summarize(report) {
  if (report.errors.length) {
    return {
      ok: false,
      summary: ["Trùng số thứ tự câu — đã dừng, chưa ghi gì vào file.", ...report.errors.map((e) => short(e, 160)), "Hãy đánh lại số câu (mỗi số chỉ dùng một lần trong một phần) rồi chạy lại."].join("\n"),
      details: report.errors.join("\n"),
    };
  }
  if (report.stoppedByUser) return "Đã dừng theo lựa chọn của bạn. Chưa ghi gì vào file.";
  if (report.cursorInTable)
    return {
      ok: false,
      summary: "Con trỏ đang nằm trong một bảng — chưa ghi gì vào file.\nHãy đặt con trỏ ở dòng trống nơi muốn tạo bảng đáp án rồi chạy lại.",
      details: "",
    };
  if (!report.groups.length) return { ok: false, summary: report.warnings[0] || "Không tìm thấy câu hỏi nào.", details: "" };

  const lines = report.groups.map((g) => {
    const filled = report.filled[g.id] ? ` — điền ${report.filled[g.id]} ô` : "";
    return `Phần ${g.id} (${KIND_VN[g.kind] || g.kind}): ${g.questions} câu${filled}`;
  });
  if (report.createdTables.length) lines.push(`Tạo mới bảng tại con trỏ: Phần ${report.createdTables.join(", ")}`);
  const warns = [...report.numberingWarnings, ...report.warnings];
  if (warns.length) {
    lines.push(`⚠ ${warns.length} cảnh báo:`);
    warns.slice(0, 3).forEach((w) => lines.push("• " + short(w)));
    if (warns.length > 3) lines.push(`… và ${warns.length - 3} cảnh báo khác (xem Chi tiết).`);
  }
  const details = [
    ...warns,
    ...Object.entries(report.answers).map(([id, a]) => `Đáp án Phần ${id}: ${a}`),
    ...Object.entries(report.tables).map(([id, t]) => `Bảng Phần ${id}: ${t}`),
  ].join("\n");
  return { ok: true, summary: lines.join("\n"), details };
}

export async function taoBangDapAn(context) {
  return summarize(await fillAnswerTable(context));
}
