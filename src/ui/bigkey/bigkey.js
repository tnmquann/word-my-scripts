// Task pane "Công cụ BigKey": các menu con của "Chuẩn hóa" (BigKeyUI.dotm) dưới dạng nút.
// Pane được nạp lại mỗi lần mở nên gọi thẳng logic trong features (không qua hot reload).
import * as T from "../../features/bigkey/tools.js";
import { xoaLoiGiai } from "../../features/xoa/xoaCanThiet.js";
import { isDesktop } from "../../features/runtime.js";

const $ = (s) => document.querySelector(s);
const status = $("#status");
const ACTIONS = { ...T, xoaLoiGiai };

/** Chạy job(ctx) trong Word.run, khoá các nút, hiện kết quả. */
async function run(job) {
  const buttons = document.querySelectorAll("button");
  buttons.forEach((b) => (b.disabled = true));
  status.textContent = "Đang chạy…";
  status.className = "";
  try {
    const res = await Word.run(job);
    const bad = res && typeof res === "object" && res.ok === false;
    status.textContent = res && typeof res === "object" ? res.summary + (res.details ? "\n" + res.details : "") : String(res);
    status.className = bad ? "err" : "ok";
  } catch (e) {
    status.textContent = "Lỗi: " + (e && e.message ? e.message : e);
    status.className = "err";
    console.error(e);
  } finally {
    buttons.forEach((b) => (b.disabled = false));
    if (!isDesktop()) document.querySelectorAll(".desk").forEach((b) => (b.disabled = true));
  }
}

function job(act) {
  if (act === "danhSoThuTu") {
    const n = Number($("#start").value);
    return (ctx) => T.danhSoThuTu(ctx, n);
  }
  return (ctx) => ACTIONS[act](ctx);
}

Office.onReady(() => {
  const g = $("#groups");
  for (let i = 1; i <= 15; i++) {
    const b = document.createElement("button");
    b.dataset.tag = `[${i}]`;
    b.textContent = String(i);
    b.title = `Nhóm ${i}`;
    g.appendChild(b);
  }
  if (!isDesktop()) document.querySelectorAll(".desk").forEach((b) => (b.disabled = true));
  document.querySelectorAll("button[data-act]").forEach((b) => b.addEventListener("click", () => run(job(b.dataset.act))));
  document.querySelectorAll("button[data-tag]").forEach((b) => b.addEventListener("click", () => run((ctx) => T.ganIDNhom(ctx, b.dataset.tag))));
});
