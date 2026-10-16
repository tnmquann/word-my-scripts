// Chuyển STT Text sang Auto (src/features/ctpt2018/sttTextToAuto.js): label search helper + flow on a mock Word
// (popup start number, labels deleted, one list "Câu n." from the chosen number, auto-numbered questions skipped, Huỷ).
import { loadModule, check, done } from "./helpers.mjs";

const M = await loadModule("features/ctpt2018/sttTextToAuto.js");
check(M.typedLabelSearch("Câu 31:\tThe project") === "Câu 31:^t", "label search: tab after label -> ^t");
check(M.typedLabelSearch("  Câu 2. Hỏi") === "Câu 2. ", "label search: leading spaces dropped, trailing space kept");
check(M.typedLabelSearch("Ví dụ 1. x") === null, "label search: not a question -> null");

function mock(texts) {
  const lists = [];
  const paras = texts.map((t) => {
    const p = {
      text: t.text ?? t, isListItem: !!t.list, tableNestingLevel: 0, deleted: [], list: null,
      listItemOrNullObject: { load() {}, isNullObject: !t.list, listString: t.list || "" },
      search(s) { const hit = p.text.replace(/\t/g, "^t").startsWith(s) ? [{ delete: () => p.deleted.push(s) }] : []; return { load() {}, items: hit }; },
      startNewList() { const l = { id: lists.length + 1, members: [p], level: null, start: null, load() {}, setLevelNumbering(lv, kind, fmt) { this.level = fmt; }, setLevelStartingNumber(lv, n) { this.start = n; } }; lists.push(l); p.list = l; return l; },
      attachToList(id) { const l = lists.find((x) => x.id === id); l.members.push(p); p.list = l; },
      get listItem() { return { load() {}, get listString() { return p.list ? `Câu ${p.list.start + p.list.members.indexOf(p)}.` : ""; } }; },
    };
    return p;
  });
  return { paras, lists, context: { sync: async () => {}, document: { body: { paragraphs: { load() { this.items = paras; } } } } } };
}
const doc = ["Câu 31:\tHỏi một?", "A. 1\tB. 2\tC. 3\tD. 4", "Câu 32. Hỏi hai?", "A. x", { text: "Hỏi ba (tự động)?", list: "Câu 1." }, "Câu 33: Hỏi bốn?"];

globalThis.__choice = { button: "Chuyển", value: "5" };
let m = mock(doc);
let res = await M.sttTextToAuto(m.context);
check(m.lists.length === 1 && m.lists[0].start === 5 && JSON.stringify(m.lists[0].level) === '["Câu ",0,"."]', "one list 'Câu n.' starting at the chosen number 5", JSON.stringify(m.lists.map((l) => [l.start, l.level])));
check(m.lists[0].members.length === 3 && m.paras[4].list === null, "3 typed questions attached; auto-numbered question untouched");
check(m.paras[0].deleted[0] === "Câu 31:^t" && m.paras[2].deleted[0] === "Câu 32. " && m.paras[5].deleted[0] === "Câu 33: ", "typed labels deleted with their space/tab");
check(m.paras[0].leftIndent === 28.35 && m.paras[0].firstLineIndent === -28.35, "hanging indent 1 cm");
check(/Câu 5\. → Câu 7\./.test(res) && !/CHƯA/.test(res), "summary 5 → 7, numbering verified", res);

globalThis.__choice = { button: "Huỷ", value: "1" };
m = mock(doc);
res = await M.sttTextToAuto(m.context);
check(m.lists.length === 0 && m.paras.every((p) => !p.deleted.length) && /huỷ/i.test(res), "Huỷ: nothing changed", res);

globalThis.__choice = { button: "Chuyển", value: "0" };
m = mock(doc);
res = await M.sttTextToAuto(m.context);
check(res.ok === false && m.lists.length === 0, "invalid start number rejected before any change");
globalThis.__choice = undefined;
done();
