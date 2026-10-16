// Canh giữa Table  (spec: docs/specs/ctpt_2018/Canh_Giua_Table.md)
// Mọi bảng trong tài liệu: canh giữa giữa hai lề (bảng lồng: giữa ô chứa nó). Bảng cấp ngoài rộng hơn vùng chữ
// (khổ giấy − lề của section chứa bảng) được co vừa lề (autoFitWindow) để không tràn lề — chỉ Word desktop.

import { isDesktop } from "../runtime.js";

export async function canhGiuaTable(context) {
  const tables = context.document.body.tables;
  tables.load("items/nestingLevel,items/width,items/alignment");
  await context.sync();
  if (!tables.items.length) return "Tài liệu không có bảng nào.";
  const desktop = isDesktop();
  const outer = tables.items.filter((t) => t.nestingLevel === 1);
  const setups = desktop ? outer.map((t) => t.getRange("Whole").sections.getFirst().pageSetup) : [];
  setups.forEach((ps) => ps.load("pageWidth,leftMargin,rightMargin,gutter"));
  await context.sync();
  let fitted = 0;
  outer.forEach((t, k) => {
    if (!desktop) return;
    const ps = setups[k];
    const avail = ps.pageWidth - ps.leftMargin - ps.rightMargin - (ps.gutter || 0);
    if (t.width > avail + 0.5) (t.autoFitWindow(), fitted++);
  });
  tables.items.forEach((t) => (t.alignment = "Centered"));
  await context.sync();
  return (
    `Đã canh giữa ${tables.items.length} bảng.` +
    (fitted ? ` Co ${fitted} bảng rộng hơn lề cho vừa trang.` : "") +
    (desktop ? "" : "\n⚠ Word Web: chỉ canh giữa, không kiểm tra bảng tràn lề (cần Word desktop).")
  );
}
