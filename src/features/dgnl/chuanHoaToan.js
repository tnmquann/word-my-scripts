// Chuẩn hoá Toán ĐGNL  (spec: docs/specs/dgnl/Chuan_Hoa_Toan.md)
// - Dòng dẫn "Dựa vào ... trả lời câu hỏi X-Y"
//   -> "Dựa vào thông tin dưới đây để trả lời từ câu X đến câu Y" (in đậm)
// - Indent left 0 cho dòng dẫn và đoạn dữ kiện (đến trước "Câu N.")
// - Canh giữa hình ảnh / bảng trong đoạn dữ kiện
// Chạy lại nhiều lần không làm hỏng kết quả cũ.

const reRange = /câu(?:\s+hỏi)?\s*(\d+)\s*[-–]\s*(\d+)/i; // "câu hỏi 73-74"
const reDone = /câu\s*(\d+)\s*đến\s*câu\s*(\d+)/i; // đã chuẩn hoá rồi
const reStart = /^\s*(Dựa vào|Trả lời)/i;

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
    return t.match(reRange) || t.match(reDone);
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

  if (!headers) return "Không tìm thấy dòng dẫn \"Dựa vào … câu X-Y\" nào.";
  return `Đã xử lý ${headers} dòng dẫn (viết lại ${changed} dòng).`;
}
