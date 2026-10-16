// Task pane "Công cụ BigKey": các menu con của "Chuẩn hóa" (BigKeyUI.dotm) dưới dạng nút.
// Pane được nạp lại mỗi lần mở nên gọi thẳng logic trong features (không qua hot reload).
import * as T from "../../features/bigkey/tools.js";
import * as PL from "../../features/bigkey/plugins.js";
import { xoaLoiGiai, xoaDongVietLG } from "../../features/xoa/xoaCanThiet.js";
import { isDesktop } from "../../features/runtime.js";

const $ = (s) => document.querySelector(s);
const status = $("#status");
const ACTIONS = { ...T, tachLG: PL.tachLG, chemFormatter: PL.chemFormatter, xoaLoiGiai, xoaDongVietLG };
const ZG_KEY = "myscripts.zg50.template"; // mẫu ans50zg.docx (base64) chỉ lưu trên máy này

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

// ---- Tô phiếu ZG50: mẫu chọn một lần, nhớ trong localStorage (có thể không dùng được -> chọn lại mỗi lần)
function loadTemplate() {
  try {
    const b64 = localStorage.getItem(ZG_KEY);
    if (!b64) return null;
    const bin = atob(b64);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    return buf.buffer;
  } catch (e) {
    return null;
  }
}
let zgTemplate = null;
function showTemplate() {
  $("#zgInfo").textContent = zgTemplate ? "Đã có mẫu phiếu." : "Chưa chọn mẫu phiếu.";
}
async function pickTemplate(file) {
  zgTemplate = await file.arrayBuffer();
  try {
    let bin = "";
    const u8 = new Uint8Array(zgTemplate);
    for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    localStorage.setItem(ZG_KEY, btoa(bin));
  } catch (e) {
    /* không lưu được: vẫn dùng cho phiên này */
  }
  showTemplate();
}

Office.onReady(() => {
  zgTemplate = loadTemplate();
  showTemplate();
  $("#zgFile").addEventListener("change", (e) => e.target.files[0] && pickTemplate(e.target.files[0]));
  $("#zgRun").addEventListener("click", () => run((ctx) => PL.toPhieuZG50(ctx, $("#zgMade").value, zgTemplate)));
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
