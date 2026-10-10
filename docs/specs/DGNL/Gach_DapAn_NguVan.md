Script Lab báo lỗi vì code chạy ở cấp cao nhất, mà `await` chỉ dùng được trong hàm `async`. Tôi đã bọc toàn bộ vào một hàm `async` và gọi hàm đó khi Office sẵn sàng. Dán thay toàn bộ nội dung ô Script:

```javascript
async function main() {
  await Word.run(async (context) => {
    const RED = "#EE0000";
    const norm = (s) => s.toLowerCase().replace(/[\s.,;:!?"'“”‘’()\-–—…]+/g, "");

    const ps = context.document.body.paragraphs;
    ps.load("items/text,items/isListItem");
    await context.sync();
    const items = ps.items;
    const lis = items.map((p) => (p.isListItem ? p.listItemOrNullObject : null));
    lis.forEach((l) => { if (l) l.load("listString"); });
    await context.sync();

    const isQuestion = (i) => lis[i] && !lis[i].isNullObject && /^Câu/.test(lis[i].listString);
    const isIntro = (i) => /^\s*Dựa vào thông tin/.test(items[i].text);
    const isLoiGiai = (i) => items[i].text.trim() === "Lời giải";

    // Thu thập các câu: dòng "Lời giải" + dòng đáp án + dòng phương án đúng
    const sections = [];
    for (let i = 0; i < items.length - 1; i++) {
      if (!isLoiGiai(i)) continue;
      const m = items[i + 1].text.match(/^\s*([A-D])\./);
      if (!m) continue; // đã xử lý rồi thì bỏ qua
      const letter = m[1];
      let opt = null;
      for (let j = i - 1; j >= Math.max(0, i - 8); j--) {
        if (new RegExp("(^|\\t)\\s*" + letter + "\\.").test(items[j].text)) { opt = j; break; }
      }
      if (opt === null) continue;
      sections.push({ i: i, ans: i + 1, letter: letter, opt: opt });
    }

    // Step 1: gạch chân đáp án (chữ cái: đậm + gạch chân + đỏ)
    const dotHits = sections.map((s) => {
      const h = items[s.opt].search(s.letter + ".", { matchCase: true });
      h.load("items");
      return h;
    });
    await context.sync();
    const letterHits = dotHits.map((h, k) => {
      if (!h.items.length) return null;
      const r = h.items[0].search(sections[k].letter, { matchCase: true });
      r.load("items");
      return r;
    });
    await context.sync();
    letterHits.forEach((r) => {
      if (r && r.items.length) {
        const f = r.items[0].font;
        f.bold = true;
        f.underline = "Single";
        f.color = RED;
      }
    });
    await context.sync();

    // Step 2: xóa dòng đáp án, phần giải thích lặp lại đề, và "Lời giải" thừa
    const toDelete = [];
    for (const s of sections) {
      toDelete.push(s.ans);
      const expl = [];
      for (let j = s.ans + 1; j < items.length && !isQuestion(j) && !isIntro(j) && !isLoiGiai(j); j++) {
        if (items[j].text.trim() !== "") expl.push(j);
      }
      const ctx = norm(items.slice(Math.max(0, s.i - 12), s.i).map((p) => p.text).join(" "));
      let removed = 0;
      for (const j of expl) {
        const n = norm(items[j].text);
        if (n.length >= 15 && ctx.indexOf(n) >= 0) { toDelete.push(j); removed++; } else break;
      }
      if (expl.length - removed === 0) toDelete.push(s.i); // hết nội dung -> xóa luôn "Lời giải"
    }
    const unique = Array.from(new Set(toDelete)).sort((a, b) => b - a);
    unique.forEach((k) => items[k].delete());
    await context.sync();

    console.log("Số câu xử lý: " + sections.length + " - Số đoạn đã xóa: " + unique.length);
  });
}

Office.onReady(() => {
  main().catch((e) => console.error(e));
});
```

**Những chỗ tôi đã sửa so với bản trước:**
- Code nằm trong hàm `async main()` nên `await` không còn báo lỗi.
- Tôi bỏ regex dùng cờ `u` và `\p{L}`, và bỏ cách viết `[...new Set()]`, vì Script Lab có thể không biên dịch được hai kiểu này.

**Lưu ý khi chạy:**
- Tôi chưa chạy được script này trong Script Lab. Trong Word, tôi mới kiểm tra riêng các phần quan trọng, như nhận diện câu và chuẩn hóa chữ. Bản trước thì tôi đã chạy thử đầy đủ trên bản sao có cấu trúc tương tự.
- Bản này gần như giống bản trước, chỉ đổi cách viết. Nếu Script Lab còn báo lỗi thì bạn gửi lại nội dung lỗi cho tôi.
- File hiện tại đã xử lý xong, nên chạy lại sẽ không đổi gì. Dòng log sẽ báo 0 câu. Script chỉ có tác dụng với file gốc còn dòng đáp án dưới "Lời giải".