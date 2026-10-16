// End-to-end Chuẩn hoá BigKey (cả đề) on a real exported exam, against a mock Word that models how Office JS
// insertOoxml("Replace") behaves on whole paragraphs: the LAST inserted paragraph takes over the destination's
// final paragraph mark and keeps its OLD paragraph properties (field report 2.0.0-beta.4: option tab stops lost).
//   node tests/bigkey-pipeline.mjs <export.json> [--src <dir>]     (--src: run another source tree, e.g. an old release)
// Asserts: one run lays out the options of EVERY question (the beta.5 bug: the last question needed a second run),
// no empty paragraph is left, and a second run changes nothing.
import fs from "node:fs";
import { loadModule, check, done } from "./helpers.mjs";

const args = process.argv.slice(2);
const file = args.find((a, i) => !a.startsWith("--") && args[i - 1] !== "--src");
const srcArg = args.includes("--src") ? args[args.indexOf("--src") + 1] : undefined;
if (!file || !fs.existsSync(file)) {
  console.log("SKIP bigkey-pipeline: no exported document (pass a JSON from tests/tools/export-docx.ps1)");
  process.exit(0);
}
const CH = await loadModule("features/bigkey/chuanHoaBigKey.js", srcArg);
const RB = await loadModule("features/dgnl/removeBreakEng.js");
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
globalThis.__desktop = false; // no page setup in the mock
const data = JSON.parse(fs.readFileSync(file, "utf8").replace(/^﻿/, ""));

// ---------------------------------------------------------------- mock Word
function mockWord(bodyXml) {
  const doc = RB.parseXml(bodyXml);
  const body = RB.documentBody(doc);
  const paras = () => RB.bodyParagraphs(body);
  const inTable = (el) => { for (let n = el.parentNode; n; n = n.parentNode) if (n.localName === "tc") return true; return false; };
  const textOf = (el) => [...el.getElementsByTagNameNS(W, "*")].filter((n) => n.localName === "t" || n.localName === "tab")
    .map((n) => (n.localName === "tab" ? (n.parentNode.localName === "tabs" ? "" : "\t") : n.textContent)).join("");
  const ser = new XMLSerializer();
  const pkgOf = (els) => {
    const d = RB.parseXml(bodyXml);
    const b = RB.documentBody(d);
    for (const n of [...b.childNodes]) if (n.nodeType === 1 && n.localName !== "sectPr") b.removeChild(n);
    const sect = [...b.childNodes].find((n) => n.nodeType === 1);
    for (const el of els) b.insertBefore(d.importNode(el, true), sect || null);
    return ser.serializeToString(d);
  };
  const font = () => ({});
  const between = (a, b) => { const ps = paras(); return ps.slice(ps.indexOf(a), ps.indexOf(b) + 1); };
  // text range inside one paragraph: replace characters across w:t nodes
  function textRange(el, start, end) {
    return {
      font: font(),
      insertText(t) {
        let pos = 0, put = false;
        for (const n of [...el.getElementsByTagNameNS(W, "t")]) {
          const s = n.textContent, a = pos, b = pos + s.length;
          pos = b;
          if (b <= start || a >= end) continue;
          const keepL = s.slice(0, Math.max(0, start - a)), keepR = s.slice(Math.max(0, end - a));
          n.textContent = keepL + (put ? "" : t) + keepR;
          put = true;
        }
        return { font: font() };
      },
    };
  }
  function range(from, to) {
    return {
      font: font(),
      expandTo: (o) => range(from, o._to),
      _to: to,
      getOoxml: () => ({ value: pkgOf(between(from, to)) }),
      insertOoxml(xml) {
        const old = between(from, to);
        const newPs = RB.bodyParagraphs(RB.documentBody(RB.parseXml(xml))).map((n) => doc.importNode(n, true));
        const last = newPs[newPs.length - 1];
        // Office JS: the last inserted paragraph merges into the destination's final ¶ and keeps its OLD pPr
        const oldPPr = [...old[old.length - 1].childNodes].find((n) => n.localName === "pPr");
        const lastPPr = [...last.childNodes].find((n) => n.localName === "pPr");
        if (lastPPr) last.removeChild(lastPPr);
        if (oldPPr) last.insertBefore(oldPPr.cloneNode(true), last.firstChild);
        for (const p of newPs) from.parentNode.insertBefore(p, from);
        for (const p of old) p.parentNode.removeChild(p);
      },
      insertText(t, where) {
        if (where === "Before") { const n = from.getElementsByTagNameNS(W, "t")[0]; if (n) n.textContent = t + n.textContent; return { font: font() }; }
        return textRange(from, 0, textOf(from).length).insertText(t);
      },
    };
  }
  const proxies = new Map();
  function para(el) {
    if (proxies.has(el)) return proxies.get(el);
    const p = {
      get text() { return textOf(el); },
      isListItem: false,
      get tableNestingLevel() { return inTable(el) ? 1 : 0; },
      listItemOrNullObject: { load() {}, isNullObject: true },
      font: font(),
      getRange: () => range(el, el),
      getOoxml: () => ({ value: pkgOf([el]) }),
      delete: () => el.parentNode && el.parentNode.removeChild(el),
      search(s) {
        const t = textOf(el), items = [];
        for (let i = t.indexOf(s); i >= 0; i = t.indexOf(s, i + s.length)) items.push(textRange(el, i, i + s.length));
        return { load() {}, items };
      },
      detachFromList() {},
    };
    proxies.set(el, p);
    return p;
  }
  const context = {
    sync: async () => {},
    document: {
      body: {
        paragraphs: { load() { this.items = paras().map(para); } },
        search: () => ({ load() {}, items: [] }),
      },
    },
  };
  return { context, paras, textOf, inTable, ser, body };
}

// ---------------------------------------------------------------- checks
function state(m) {
  const ps = m.paras();
  const P = ps.map((el) => ({ text: m.textOf(el), inTable: m.inTable(el), list: "" }));
  const plan = CH.planBlocks(P);
  const optOk = plan.filter((b) => b.optFrom >= 0).map((b) => {
    const xml = ps.slice(b.optFrom, b.optTo + 1).map((el) => m.ser.serializeToString(el)).join("");
    // rebuilt by optionLayout: blue bold labels + standard single-spaced pPr (+ tab stops when 2 or 4 per line)
    const laid = /w:color w:val="0000FF"/.test(xml) && /w:line="240"/.test(xml);
    const tabsOk = b.optTo > b.optFrom + 2 || /<w:tabs>/.test(xml); // 1–2 paragraphs means 4 or 2 per line: needs tabs
    return { num: b.num, ok: laid && tabsOk };
  });
  const empties = P.slice(0, -1).filter((x, i) => !x.inTable && !x.text.trim() && !/drawing|pict|object|oMath/.test(m.ser.serializeToString(ps[i]))).length;
  return { plan, optOk, empties, snapshot: ps.map((el) => m.ser.serializeToString(el)).join("\n") };
}

const m = mockWord(data.bodyXml);
const expected = state(m).optOk.length; // option blocks in the ORIGINAL document
const res1 = await CH.chuanHoaBigKeyDe(m.context);
const s1 = state(m);
const bad1 = s1.optOk.filter((o) => !o.ok).map((o) => o.num);
const name = file.split(/[\\/]/).pop();
console.log(`  run 1: ${res1.summary.replace(/\n/g, " | ")}`);
check(expected > 0 && s1.optOk.length === expected && !bad1.length, `${name}: ONE run lays out the options of all ${expected} questions`, `${s1.optOk.length} option blocks found; not laid out: Câu ${bad1.join(", ")}`);
check(s1.empties === 0, `${name}: no empty paragraph left after one run`, String(s1.empties));
await CH.chuanHoaBigKeyDe(m.context);
const s2 = state(m);
check(s2.snapshot === s1.snapshot, `${name}: a second run changes nothing`);
done();
