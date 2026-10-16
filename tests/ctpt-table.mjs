// Canh giữa Table (src/features/ctpt2018/canhGiuaTable.js) against a mock Word: every table centred,
// only top-level tables wider than the text area are fitted to the margins, nothing fitted on Word web.
import { loadModule, check, done } from "./helpers.mjs";

const { canhGiuaTable } = await loadModule("features/ctpt2018/canhGiuaTable.js");
function mock(tables) {
  const items = tables.map((t) => ({
    ...t,
    alignment: "Left",
    fitted: false,
    autoFitWindow() { this.fitted = true; },
    getRange: () => ({ sections: { getFirst: () => ({ pageSetup: { pageWidth: 595, leftMargin: 42.5, rightMargin: 42.5, gutter: 0, load() {} } }) } }),
  }));
  return { items, context: { sync: async () => {}, document: { body: { tables: { items, load() {} } } } } };
}
const spec = [{ nestingLevel: 1, width: 400 }, { nestingLevel: 1, width: 600 }, { nestingLevel: 2, width: 900 }];

globalThis.__desktop = true;
let m = mock(spec);
let res = await canhGiuaTable(m.context);
check(m.items.every((t) => t.alignment === "Centered"), "desktop: every table centred (nested too)");
check(!m.items[0].fitted && m.items[1].fitted && !m.items[2].fitted, "desktop: only the top-level table wider than 510 pt is fitted");
check(/3 bảng/.test(res) && /Co 1 bảng/.test(res), "desktop: summary", res);

globalThis.__desktop = false;
m = mock(spec);
res = await canhGiuaTable(m.context);
check(m.items.every((t) => t.alignment === "Centered" && !t.fitted), "web: centred, nothing fitted");
check(/Word Web/.test(res), "web: warning shown");

m = mock([]);
check((await canhGiuaTable(m.context)) === "Tài liệu không có bảng nào.", "no table: message");
done();
