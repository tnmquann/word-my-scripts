Đây là script Office.js, chạy bằng add-in **Script Lab** trong Word (Insert → Script Lab → Code → dán vào tab Script → Run). Tôi không gửi được macro VBA từ đây.

```javascript
// Chuẩn hóa dòng dẫn "Dựa vào ... trả lời câu hỏi X-Y"
// -> "Dựa vào thông tin dưới đây để trả lời từ câu X đến câu Y"
// + indent left 0 cho dòng dẫn và đoạn dữ kiện (đến trước "Câu N.")
// + canh giữa hình ảnh / bảng trong đoạn dữ kiện

const reRange = /câu(?:\s+hỏi)?\s*(\d+)\s*[-–]\s*(\d+)/i;   // "câu hỏi 73-74"
const reDone  = /câu\s*(\d+)\s*đến\s*câu\s*(\d+)/i;          // đã chuẩn hóa rồi
const reStart = /^\s*(Dựa vào|Trả lời)/i;

async function run() {
  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const items = paras.items;

    const lis = items.map(p => { const l = p.listItemOrNullObject; l.load("isNullObject,listString"); return l; });
    const pics = items.map(p => { const c = p.inlinePictures; c.load("items"); return c; });
    const tbls = items.map(p => { const t = p.parentTableOrNullObject; t.load("isNullObject"); return t; });
    await ctx.sync();

    const isQuestion = (j) => !lis[j].isNullObject && /^Câu/i.test(lis[j].listString || "");
    const getRange = (j) => {
      const t = items[j].text;
      if (isQuestion(j) || !reStart.test(t) || t.length > 120) return null;
      return t.match(reRange) || t.match(reDone);
    };

    let headers = 0;
    for (let i = 0; i < items.length; i++) {
      const m = getRange(i);
      if (!m) continue;
      headers++;

      // 1) Viết lại dòng dẫn (giữ in đậm)
      const newText = `Dựa vào thông tin dưới đây để trả lời từ câu ${m[1]} đến câu ${m[2]}`;
      if (items[i].text !== newText) items[i].insertText(newText, "Replace");
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
    await ctx.sync();
    console.log(`Đã xử lý ${headers} dòng dẫn.`);
  });
}

run().catch(e => console.error(e));
```

Cách script hoạt động:
- **Nhận diện dòng dẫn:** Script chọn đoạn bắt đầu bằng "Dựa vào" hoặc "Trả lời" có dạng "câu hỏi 73-74" hoặc "câu 73 - 74". Nhờ vậy câu như "Dựa vào BBT suy ra…" không bị đụng tới.
- **Phạm vi đoạn dữ kiện:** Đoạn dữ kiện kéo dài đến trước đoạn "Câu N." tiếp theo.
- **Chạy lại:** Script chạy lại được nhiều lần mà không làm hỏng kết quả cũ.