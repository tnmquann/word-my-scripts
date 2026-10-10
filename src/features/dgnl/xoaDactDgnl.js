// Xoá ĐACT ĐGNL  (spec: docs/specs/dgnl/Xoa_DACT_DGNL.md)
// Xoá các khối đáp án chi tiết bắt đầu bằng dòng "Lời giải".
// Khối kết thúc ngay trước: câu hỏi mới (thụt treo), đoạn dẫn/tiêu đề (lề trái 0)
// hoặc dòng "Lời giải" kế tiếp. Bảng trong khối được xoá riêng trước (tránh GeneralException).
//
// Tối ưu so với bản Script Lab: xoá tất cả khối trong vài lần sync.
// Nếu cách nhanh lỗi -> quét lại và xoá từng khối như bản gốc.

// ====== CẤU HÌNH ======
const MARKER = "Lời giải"; // dòng đánh dấu bắt đầu đáp án
const HANGING_MAX = -20; // firstLineIndent < -20pt = đầu câu hỏi mới
const LEFT_MIN = 10; // lề trái < 10pt = đoạn dẫn / tiêu đề phần
// =======================

async function scan(context) {
  const ps = context.document.body.paragraphs;
  ps.load("items/text,items/leftIndent,items/firstLineIndent,items/tableNestingLevel");
  await context.sync();
  const it = ps.items;

  const blocks = [];
  for (let i = 0; i < it.length; i++) {
    if (it[i].text.trim() !== MARKER || it[i].tableNestingLevel !== 0) continue;
    let j = i + 1;
    while (
      j < it.length &&
      !(
        it[j].tableNestingLevel === 0 &&
        (it[j].firstLineIndent < HANGING_MAX || it[j].leftIndent < LEFT_MIN || it[j].text.trim() === MARKER)
      )
    )
      j++;
    blocks.push([i, j - 1]);
  }
  return { it, blocks };
}

const blockRange = (it, [a, b]) => it[a].getRange("Whole").expandTo(it[b].getRange("Whole"));

async function deleteFast(context, it, blocks) {
  // Từ cuối lên đầu
  const ranges = blocks
    .slice()
    .reverse()
    .map((blk) => blockRange(it, blk));
  const tbls = ranges.map((r) => {
    const t = r.tables;
    t.load("items");
    return t;
  });
  await context.sync();

  let nTables = 0;
  tbls.forEach((t) =>
    t.items.forEach((x) => {
      x.delete();
      nTables++;
    })
  );
  if (nTables) await context.sync();

  ranges.forEach((r) => r.delete());
  await context.sync();
  return nTables;
}

async function deleteSlow(context, it, blocks) {
  let ok = 0;
  const fail = [];
  for (let n = blocks.length - 1; n >= 0; n--) {
    const rng = blockRange(it, blocks[n]);
    try {
      const tbs = rng.tables;
      tbs.load("items");
      await context.sync();
      if (tbs.items.length) {
        tbs.items.forEach((t) => t.delete());
        await context.sync();
      }
      rng.delete();
      await context.sync();
      ok++;
    } catch (e) {
      fail.push(blocks[n]);
    }
  }
  return { ok, fail };
}

export async function xoaDactDgnl(context) {
  const first = await scan(context);
  if (!first.blocks.length) return `Không tìm thấy dòng "${MARKER}" nào.`;

  const lines = [`Tìm thấy ${first.blocks.length} khối đáp án.`];
  try {
    const nTables = await deleteFast(context, first.it, first.blocks);
    lines.push(`Đã xoá ${first.blocks.length} khối` + (nTables ? ` (gồm ${nTables} bảng).` : "."));
  } catch (e) {
    // Cách nhanh lỗi giữa chừng -> quét lại những khối còn sót và xoá từng khối
    console.warn("deleteFast failed, fallback to slow path", e);
    const again = await scan(context);
    const { ok, fail } = await deleteSlow(context, again.it, again.blocks);
    lines.push(`Xoá từng khối: thành công ${ok}, lỗi ${fail.length}.`);
    if (fail.length) lines.push(`Khối lỗi (chỉ số đoạn): ${JSON.stringify(fail.slice(0, 10))}`);
  }

  // Kiểm tra lại
  const p2 = context.document.body.paragraphs;
  p2.load("items/text");
  await context.sync();
  const left = p2.items.filter((p) => p.text.trim() === MARKER).length;
  lines.push(`Đoạn còn lại: ${p2.items.length} | Còn dòng "${MARKER}": ${left}`);
  return lines.join("\n");
}
