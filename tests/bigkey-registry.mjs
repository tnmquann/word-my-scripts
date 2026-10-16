// Chuẩn hoá BigKey registry: every command in src/features/bigkey/index.js is a function, every manifest
// FunctionName exists in some group registry (or is a system command), and the BigKey pure helpers behave.
import fs from "node:fs";
import path from "node:path";
import { root, loadModule, check, done } from "./helpers.mjs";

const reg = (await loadModule("features/bigkey/index.js")).default;
const bad = Object.entries(reg).filter(([, c]) => typeof c.run !== "function" || !c.label);
check(Object.keys(reg).length === 3 && !bad.length, `bigkey registry: ${Object.keys(reg).length} ribbon commands, all runnable`, bad.map(([k]) => k).join(","));

// task pane "Công cụ BigKey": every data-act button maps to a function of tools.js (or xoaLoiGiai)
const paneHtml = fs.readFileSync(path.join(root, "src/ui/bigkey/bigkey.html"), "utf8");
const acts = [...paneHtml.matchAll(/data-act="(\w+)"/g)].map((m) => m[1]);
const tools = await loadModule("features/bigkey/tools.js");
const plugins = await loadModule("features/bigkey/plugins.js");
const xoa = await loadModule("features/xoa/xoaCanThiet.js");
const missingActs = acts.filter((a) => typeof tools[a] !== "function" && typeof plugins[a] !== "function" && typeof xoa[a] !== "function");
check(acts.length >= 20 && !missingActs.length, `pane: ${acts.length} buttons map to tools.js / plugins.js / xoaCanThiet.js`, missingActs.join(","));

// manifest FunctionNames -> must be registered in some group (duplicates across groups are fatal at startup)
const groups = ["dgnl", "ctpt2018", "equations", "bigkey", "xoa"];
const all = {};
const dup = [];
for (const g of groups) for (const k of Object.keys((await loadModule(`features/${g}/index.js`)).default)) (all[k] ? dup.push(k) : (all[k] = g));
check(!dup.length, "no FunctionName registered twice", dup.join(","));
const manifest = fs.readFileSync(path.join(root, "manifest/manifest.template.xml"), "utf8");
const fns = [...manifest.matchAll(/<FunctionName>(\w+)<\/FunctionName>/g)].map((m) => m[1]);
const system = ["checkUpdate"];
const unknown = [...new Set(fns)].filter((f) => !all[f] && !system.includes(f));
check(!unknown.length, `every manifest FunctionName is registered (${new Set(fns).size})`, unknown.join(","));

const T = await loadModule("features/bigkey/tools.js");
check(JSON.stringify(T.idTagsBK("Câu 1. [2D1-3] x [1] y [0;1] [2D1-3] [a b] [0H2-1.3-2]")) === '["[2D1-3]","[1]","[0H2-1.3-2]"]', "idTagsBK: IDs only, not intervals");
const tbl = (vals) => ({ values: vals });
check(T.keyTable([tbl([["MÃ ĐỀ: 001"], ["1.A", "2.B", "3.C", "4.D", "5.A", "", "", "", "", ""]])]) !== null, "keyTable: BigKey key detected");
check(T.keyTable([tbl([["a", "b"], ["1.A", "2.B"]])]) === null, "keyTable: other tables ignored");
// planBlocks: an empty paragraph inside an option block (left by insertOoxml) must not hide the options —
// regression for "the last question needs Chuẩn hoá twice".
const CH = await loadModule("features/bigkey/chuanHoaBigKey.js");
const para = (text) => ({ text, inTable: false, list: "" });
const P = ["Câu 1. Hỏi?", "A. 1", "B. 2", "", "C. 3", "D. 4", "Lời giải", "Chọn A", "Câu 2. Hỏi?", "A. x", "", "B. y", "", "C. z", "D. t"].map(para);
const plan = CH.planBlocks(P);
check(plan.length === 2 && plan[0].optFrom === 1 && plan[0].optTo === 5 && plan[1].optFrom === 9 && plan[1].optTo === 14,
  "planBlocks: empty paragraphs inside option blocks are skipped (incl. last question)", JSON.stringify(plan.map((b) => [b.optFrom, b.optTo])));
// Check lỗi đề: two underlined answers -> "thừa đáp án" (the BigKey 2024 "Quét lỗi" check)
const CK = await loadModule("features/bigkey/checkLoiDe.js");
const Wn = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const u = (t) => `<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
const r = (t) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const optPkg = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${Wn}"><w:body><w:p>${u("A")}${r(". 1")}<w:r><w:tab/></w:r>${u("B")}${r(". 2")}<w:r><w:tab/></w:r>${r("C. 3")}<w:r><w:tab/></w:r>${r("D. 4")}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const P3 = ["Câu 1. Hỏi?", "A. 1\tB. 2\tC. 3\tD. 4"].map(para);
const errs3 = CK.findErrors(P3, CH.planBlocks(P3), [optPkg], [[]]).map((e) => e.msg);
check(errs3.some((m) => /thừa đáp án/.test(m)), "check lỗi đề: two underlined answers reported", JSON.stringify(errs3));
done();
