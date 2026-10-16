// Task pane "Xoá cần thiết" — thay form HO_tro của NHOM-WORD-BIEN-SOAN.dotm.
// Pane được nạp lại mỗi lần mở nên gọi thẳng logic trong features (không qua hot reload).
import * as X from "../../features/xoa/xoaCanThiet.js";
import { isDesktop } from "../../features/runtime.js";

const $ = (s) => document.querySelector(s);
const status = $("#status");

// Mỗi action đọc + kiểm tra input TRƯỚC Word.run (input sai thì không chạm vào tài liệu), trả về hàm (ctx) => …
const ACTIONS = {
  mucDo: () => {
    const lv = [...document.querySelectorAll('input[name="lv"]:checked')].map((c) => +c.value);
    if (!lv.length) throw new Error("Chưa chọn mức nào.");
    return (ctx) => X.xoaTheoMucDo(ctx, lv);
  },
  loiGiai: () => X.xoaLoiGiai,
  suyLuan: () => X.xoaHuongSuyLuan,
  id: () => X.xoaID,
  dots: () => {
    const n = Math.round(+$("#nDots").value);
    if (!(n >= 1 && n <= 30)) throw new Error("Số dòng chấm phải từ 1 đến 30.");
    return (ctx) => X.themDongCham(ctx, n);
  },
  khung: () => {
    const opts = { style: $("#style").value, borderColor: $("#color").value, fill: $("#fill").value };
    return (ctx) => X.taoKhung(ctx, opts);
  },
  huyKhung: () => X.huyKhung,
};

function show(text, cls) {
  status.textContent = text;
  status.className = cls || "";
}

async function run(act) {
  const buttons = document.querySelectorAll("button");
  buttons.forEach((b) => (b.disabled = true));
  show("Đang chạy…");
  const t0 = Date.now();
  try {
    const job = ACTIONS[act]();
    const res = await Word.run(job);
    const sec = ((Date.now() - t0) / 1000).toFixed(1);
    if (res && typeof res === "object") show(`${res.summary}${res.details ? "\n" + res.details : ""}`, res.ok === false ? "err" : "ok");
    else show(`${res} (${sec} s)`, "ok");
  } catch (e) {
    show("Lỗi: " + (e && e.message ? e.message : e), "err");
    console.error(e);
  } finally {
    buttons.forEach((b) => (b.disabled = false));
    if (!isDesktop()) document.querySelectorAll(".desk").forEach((b) => (b.disabled = true));
  }
}

Office.onReady(() => {
  $("#style").innerHTML = Object.entries(X.BORDER_STYLES).map(([k, v]) => `<option value="${v}">${k}</option>`).join("");
  if (!isDesktop()) {
    $("#webWarn").hidden = false;
    document.querySelectorAll(".desk").forEach((b) => (b.disabled = true));
  }
  document.querySelectorAll("button[data-act]").forEach((b) => b.addEventListener("click", () => run(b.dataset.act)));
});
