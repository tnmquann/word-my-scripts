// Chuyển STT Text sang Auto  (spec: docs/specs/ctpt_2018/STT_Text_Sang_Auto.md)
// Ngược với "Chuyển STT Auto sang Text": bỏ nhãn "Câu N." / "Câu N:" gõ tay ở mọi câu hỏi ngoài bảng, gắn các câu vào
// MỘT danh sách đánh số tự động "Câu 1." bắt đầu từ số người dùng chọn (popup), thụt lề treo 1 cm.
// Câu đã đánh số tự động giữ nguyên.

import { askUser } from "../runtime.js";
import { loadParas, planBlocks } from "../bigkey/chuanHoaBigKey.js";

const CM = 28.35; // pt

/**
 * Chuỗi tìm (cú pháp Word search) cho nhãn gõ tay đầu đoạn kèm khoảng trắng sau nó: "Câu 31:\t" -> "Câu 31:^t".
 * null nếu đoạn không bắt đầu bằng "Câu N." / "Câu N:". Thuần.
 */
export function typedLabelSearch(text) {
  const m = /^\s*(Câu\s*\d+\s*[.:][ \t]*)/i.exec(text);
  return m ? m[1].replace(/\^/g, "^^").replace(/\t/g, "^t") : null;
}

export async function sttTextToAuto(context) {
  const P = await loadParas(context);
  const qs = planBlocks(P).filter((b) => !P[b.q].list && typedLabelSearch(P[b.q].text));
  if (!qs.length) return 'Không có câu hỏi "Câu N." gõ tay nào (có thể đã là đánh số tự động).';

  // 0) hỏi số bắt đầu (trước khi sửa tài liệu: Huỷ = không đổi gì)
  const ans = await askUser({
    title: "Chuyển STT Text sang Auto",
    lines: [`Tìm thấy ${qs.length} câu "Câu N." gõ tay (câu đầu hiện là số ${qs[0].num}).`, "Đánh số tự động bắt đầu từ số:"],
    buttons: ["Chuyển", "Huỷ"],
    input: { label: "Bắt đầu từ", value: 1 },
  });
  // runtime cũ (chưa khởi động lại Word) trả về chuỗi tên nút, không có ô nhập -> bắt đầu từ 1
  const button = typeof ans === "string" ? ans : ans.button;
  if (button !== "Chuyển") return "Đã huỷ, tài liệu không thay đổi.";
  const start = typeof ans === "string" ? 1 : Number(ans.value);
  if (!Number.isInteger(start) || start < 1 || start > 9999)
    return { ok: false, summary: `Số bắt đầu không hợp lệ: "${ans.value}". Nhập số nguyên từ 1 đến 9999.`, details: "" };

  // 1) xoá nhãn gõ tay
  const hits = qs.map((b) => {
    const s = P[b.q].p.search(typedLabelSearch(P[b.q].text), { matchCase: false });
    s.load("items");
    return s;
  });
  await context.sync();
  const done = qs.filter((b, k) => hits[k].items[0]);
  const miss = qs.filter((b, k) => !hits[k].items[0]).map((b) => b.num);
  hits.forEach((s) => s.items[0] && s.items[0].delete());
  await context.sync();
  if (!done.length) return { ok: false, summary: "Không xoá được nhãn câu nào.", details: "" };

  // 2) một danh sách "Câu n." cho mọi câu, bắt đầu từ start
  const list = P[done[0].q].p.startNewList();
  list.load("id");
  await context.sync();
  list.setLevelNumbering(0, "Arabic", ["Câu ", 0, "."]);
  list.setLevelStartingNumber(0, start);
  done.slice(1).forEach((b) => P[b.q].p.attachToList(list.id, 0));
  await context.sync();

  // 3) thụt lề treo 1 cm (gắn vào danh sách đổi lề)
  done.forEach((b) => {
    P[b.q].p.leftIndent = CM;
    P[b.q].p.firstLineIndent = -CM;
  });
  await context.sync();

  // 4) kiểm tra như fixTabNumbering: đọc lại số của danh sách
  done.forEach((b) => P[b.q].p.listItem.load("listString"));
  await context.sync();
  const nums = done.map((b) => +P[b.q].p.listItem.listString.replace(/\D/g, ""));
  const ok = nums.every((n, k) => n === start + k);
  const msg = `Đã chuyển ${done.length} câu sang đánh số tự động: Câu ${start}. → Câu ${start + done.length - 1}.`;
  const warn = miss.length ? `\n⚠ Không tìm thấy nhãn ở câu ${miss.join(", ")} (giữ nguyên).` : "";
  return ok
    ? msg + warn
    : `${msg}${warn}\n⚠ Số câu CHƯA liên tục, kiểm tra lại: ${nums
        .map((n, k) => (n !== start + k ? `#${start + k}→${n}` : null))
        .filter(Boolean)
        .slice(0, 10)
        .join(", ")}`;
}
