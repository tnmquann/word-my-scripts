// Chuẩn hoá Toán ĐGNL  (spec: docs/specs/dgnl/Chuan_Hoa_Toan.md)
// - Dòng dẫn "Dựa vào ... trả lời câu hỏi X-Y"
//   -> "Dựa vào thông tin dưới đây để trả lời từ câu X đến câu Y" (in đậm)
// - Indent left 0 cho dòng dẫn và đoạn dữ kiện (đến trước "Câu N.")
// - Canh giữa hình ảnh / bảng trong đoạn dữ kiện
// Chạy lại nhiều lần không làm hỏng kết quả cũ.

// Khoảng câu, chấp nhận các dạng:
//   "câu hỏi 73-74", "câu 73 – 74", "câu 73 đến câu 74" (đã chuẩn),
//   "các câu hỏi từ 72 đến 74", "câu hỏi từ 72 - 74"
const reRange = /câu(?:\s+hỏi)?\s*(?:từ\s*)?(\d+)\s*(?:[-–]|đến)\s*(?:câu\s*)?(\d+)/i;
const reStart = /^\s*(Dựa vào|Trả lời)/i;
// Dòng dẫn đã ở dạng chuẩn (sau bước 1, hoặc vốn đã đúng)
const reCanonical = /^\s*Dựa vào thông tin dưới đây để trả lời từ câu \d+ đến câu \d+/i;

export async function chuanHoaToan(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items/text");
  await context.sync();
  const items = paras.items;

  const lis = items.map((p) => {
    const l = p.listItemOrNullObject;
    l.load("isNullObject,listString");
    return l;
  });
  const pics = items.map((p) => {
    const c = p.inlinePictures;
    c.load("items");
    return c;
  });
  const tbls = items.map((p) => {
    const t = p.parentTableOrNullObject;
    t.load("isNullObject");
    return t;
  });
  await context.sync();

  const isQuestion = (j) => !lis[j].isNullObject && /^Câu/i.test(lis[j].listString || "");
  const getRange = (j) => {
    const t = items[j].text;
    if (isQuestion(j) || !reStart.test(t) || t.length > 120) return null;
    return t.match(reRange);
  };

  let headers = 0;
  let changed = 0;
  for (let i = 0; i < items.length; i++) {
    const m = getRange(i);
    if (!m) continue;
    headers++;

    // 1) Viết lại dòng dẫn (giữ in đậm)
    const newText = `Dựa vào thông tin dưới đây để trả lời từ câu ${m[1]} đến câu ${m[2]}`;
    if (items[i].text !== newText) {
      items[i].insertText(newText, "Replace");
      changed++;
    }
    items[i].font.bold = true;

    // 2) Dòng dẫn + đoạn dữ kiện: indent 0; hình/bảng canh giữa
    for (let j = i; j < items.length; j++) {
      if (j > i && (isQuestion(j) || getRange(j))) break;
      items[j].leftIndent = 0;
      items[j].firstLineIndent = 0;
      if (pics[j].items.length > 0) items[j].alignment = "Centered";
      if (!tbls[j].isNullObject) tbls[j].alignment = "Centered";
    }
  }
  await context.sync();

  // 3) Sau khi chuẩn hoá: mọi dòng dẫn đúng mẫu (kể cả dòng vốn đã đúng, không cần replace)
  //    + toàn bộ đoạn ngay bên dưới, đến trước dòng "Câu N." (auto numbering) -> indent 0.
  //    Chạy sau sync để áp lên đoạn đã được insertText("Replace") ghi lại.
  const after = context.document.body.paragraphs;
  after.load("items/text");
  await context.sync();
  const ps = after.items;
  const lis2 = ps.map((p) => {
    const l = p.listItemOrNullObject;
    l.load("isNullObject,listString");
    return l;
  });
  await context.sync();
  const isQ = (j) => !lis2[j].isNullObject && /^Câu/i.test(lis2[j].listString || "");

  let blocks = 0;
  let indented = 0;
  for (let i = 0; i < ps.length; i++) {
    if (!reCanonical.test(ps[i].text)) continue;
    blocks++;
    for (let j = i; j < ps.length; j++) {
      if (j > i && (isQ(j) || reCanonical.test(ps[j].text))) break;
      ps[j].leftIndent = 0;
      ps[j].firstLineIndent = 0;
      indented++;
    }
  }
  if (indented) await context.sync();

  if (!headers && !blocks) return "Không tìm thấy dòng dẫn \"Dựa vào … câu X-Y\" nào.";
  return [
    `Đã xử lý ${headers} dòng dẫn (viết lại ${changed} dòng).`,
    `Đưa ${indented} đoạn (dòng dẫn + dữ kiện) về lề trái 0.`,
  ].join("\n");
}
