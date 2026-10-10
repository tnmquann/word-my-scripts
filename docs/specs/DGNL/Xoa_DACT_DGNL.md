Dưới đây là script gộp từ các lần chạy trước. Tôi chưa chạy nguyên khối này một lần, nên bạn hãy thử trên bản sao của file trước.

## Cách dùng

1. Mở file Word, vào **Insert → Get Add-ins**, tìm và cài **Script Lab** (của Microsoft, miễn phí).
2. Vào **Home → Script Lab → Code** để mở trình soạn thảo.
3. Tạo script mới (**New**), chọn tab **Script**, xoá hết nội dung mẫu rồi dán đoạn dưới vào.
4. **Lưu một bản sao file trước khi chạy.** Script xoá thẳng, và nếu tắt Track Changes thì khó hoàn tác từng bước.
5. Nhấn **Run**. Kết quả (số khối đã xoá, số đoạn còn lại) hiện ở cửa sổ **Console**.

```javascript
document.getElementById("run").addEventListener("click", () => tryCatch(run));

async function run() {
  // ====== CẤU HÌNH ======
  const MARKER = "Lời giải";   // dòng đánh dấu bắt đầu đáp án
  const HANGING_MAX = -20;     // firstLineIndent < -20pt = đầu câu hỏi mới
  const LEFT_MIN = 10;         // lề trái < 10pt = đoạn dẫn / tiêu đề phần
  // =======================

  await Word.run(async (context) => {
    const ps = context.document.body.paragraphs;
    ps.load("items/text,items/leftIndent,items/firstLineIndent,items/tableNestingLevel");
    await context.sync();
    const it = ps.items;

    // 1) Tìm các khối đáp án
    const blocks = [];
    for (let i = 0; i < it.length; i++) {
      if (it[i].text.trim() === MARKER && it[i].tableNestingLevel === 0) {
        let j = i + 1;
        while (
          j < it.length &&
          !(
            it[j].tableNestingLevel === 0 &&
            (it[j].firstLineIndent < HANGING_MAX ||
              it[j].leftIndent < LEFT_MIN ||
              it[j].text.trim() === MARKER)
          )
        ) j++;
        blocks.push([i, j - 1]);
      }
    }
    console.log("Số khối đáp án tìm thấy: " + blocks.length);

    // 2) Xoá từ cuối lên đầu
    let ok = 0, fail = [];
    for (let n = blocks.length - 1; n >= 0; n--) {
      const [a, b] = blocks[n];
      const rng = it[a].getRange("Whole").expandTo(it[b].getRange("Whole"));
      try {
        const tbs = rng.tables;
        tbs.load("items");
        await context.sync();
        tbs.items.forEach((t) => t.delete());
        await context.sync();

        rng.delete();
        await context.sync();
        ok++;
      } catch (e) {
        fail.push([a, b]);
      }
    }
    console.log("Đã xoá: " + ok + " | Lỗi: " + JSON.stringify(fail));

    // 3) Kiểm tra lại
    const p2 = context.document.body.paragraphs;
    p2.load("items/text");
    await context.sync();
    console.log("Đoạn còn lại: " + p2.items.length +
      " | Còn dòng '" + MARKER + "': " +
      p2.items.filter((p) => p.text.trim() === MARKER).length);
  });
}

async function tryCatch(callback) {
  try { await callback(); } catch (error) { console.error(error); }
}
```

Trong tab **HTML** của Script Lab, bạn cần một nút có `id="run"`:

```html
<button id="run">Xoá đáp án</button>
```

## Logic của script

- **Điểm bắt đầu:** đoạn có nội dung đúng bằng "Lời giải".
- **Điểm kết thúc:** đoạn liền trước một trong ba loại sau:
  - Đoạn bắt đầu câu hỏi mới, có thụt treo (`firstLineIndent` nhỏ hơn -20pt).
  - Đoạn dẫn hoặc tiêu đề phần, có lề trái bằng 0.
  - Dòng "Lời giải" kế tiếp.
- **Vì sao không dựa vào lề 1,75 cm (49,6pt):** câu hỏi và lời giải đều thụt lề này, nên lề trái không giúp phân biệt.
- **Bảng và hình trong đáp án:** script xoá bảng riêng trước, vì xoá cả khối cùng lúc sẽ báo lỗi `GeneralException`. Hình gắn trong các đoạn đáp án bị xoá theo.

## Khi áp dụng cho file khác

- Nếu dòng đánh dấu khác (ví dụ "Hướng dẫn giải" hay "Đáp án"), đổi `MARKER`.
- Nếu câu hỏi trong file không dùng thụt treo, hoặc đoạn dẫn không có lề 0, chỉnh `HANGING_MAX` và `LEFT_MIN`. Nếu không, script có thể xoá thừa hoặc thiếu.
- Với file mới, hãy chạy thử trước. Đổi `rng.delete()` thành `console.log(it[a].text, a, b)` (và bỏ phần xoá bảng) để xem danh sách khối trước khi xoá thật.