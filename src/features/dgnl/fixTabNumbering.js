// Fix lỗi tab câu hỏi  (spec: docs/specs/dgnl/Fix_Tab_Numbering.md)
// - Gộp mọi dòng "Câu N." (list item ngoài bảng) vào danh sách của câu đầu tiên
// - Canh đều hai bên + thụt lề treo
// - Đánh số liên tục bắt đầu từ 1

const IND = 49.6; // thụt lề treo (pt) = 1,75 cm

export async function fixTabNumbering(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items/isListItem,items/tableNestingLevel");
  await context.sync();

  // Lấy các dòng đề "Câu N." (list item ngoài bảng, bỏ qua list kiểu "(1)")
  const items = paras.items.filter((p) => p.isListItem && p.tableNestingLevel === 0);
  items.forEach((p) => {
    p.listItem.load("listString");
    p.list.load("id");
  });
  await context.sync();
  const cau = items.filter((p) => /^\D+\d+\.$/.test(p.listItem.listString.trim()));
  if (!cau.length) return "Không tìm thấy dòng \"Câu N.\" nào.";

  // 1) Gộp tất cả vào danh sách của câu đầu tiên
  const mainList = cau[0].list;
  const mainId = cau[0].list.id;
  const move = cau.filter((p) => p.list.id !== mainId);
  if (move.length) {
    move.forEach((p) => p.detachFromList());
    await context.sync();
    move.forEach((p) => p.attachToList(mainId, 0));
    await context.sync();
  }

  // 2) Canh đều hai bên + thụt lề treo
  cau.forEach((p) => {
    p.alignment = "Justified";
    p.leftIndent = IND;
    p.firstLineIndent = -IND;
  });

  // 3) Đánh số liên tục bắt đầu từ 1
  mainList.setLevelStartingNumber(0, 1);
  await context.sync();

  // 4) Kiểm tra
  cau.forEach((p) => p.listItem.load("listString"));
  await context.sync();
  const nums = cau.map((p) => +p.listItem.listString.replace(/\D/g, ""));
  const ok = nums.every((n, k) => n === k + 1);

  const msg = `Đã xử lý ${cau.length} câu (gộp ${move.length} câu vào danh sách chính).`;
  return ok
    ? `${msg}\nSố câu liên tục 1 → ${nums.length}.`
    : `${msg}\n⚠ Số câu CHƯA liên tục, kiểm tra lại các câu: ${nums
        .map((n, k) => (n !== k + 1 ? `#${k + 1}→${n}` : null))
        .filter(Boolean)
        .slice(0, 10)
        .join(", ")}`;
}
