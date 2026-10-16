// Task pane "Công cụ BigKey": các menu con của "Chuẩn hóa" (BigKeyUI.dotm) dưới dạng nút.
// Pane được nạp lại mỗi lần mở nên gọi thẳng logic trong features (không qua hot reload).
import * as T from "../../features/bigkey/tools.js";
import * as PL from "../../features/bigkey/plugins.js";
import { xoaLoiGiai, xoaDongVietLG } from "../../features/xoa/xoaCanThiet.js";
import { isDesktop } from "../../features/runtime.js";

const $ = (s) => document.querySelector(s);
const status = $("#status");
const ACTIONS = { ...T, tachLG: PL.tachLG, chuyenLoiGiaiXuongCuoi: PL.chuyenLoiGiaiXuongCuoi, chemFormatter: PL.chemFormatter, xoaLoiGiai, xoaDongVietLG };
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
/** Mẫu mặc định đi kèm add-in (templates/ans50zg.docx trên cùng trang). */
async function defaultTemplate() {
  const res = await fetch(new URL("templates/ans50zg.docx", location.href).href);
  if (!res.ok) throw new Error("Không tải được mẫu phiếu mặc định (" + res.status + ").");
  return res.arrayBuffer();
}
function showTemplate() {
  $("#zgInfo").textContent = zgTemplate ? "Đang dùng mẫu phiếu bạn đã chọn." : "Đang dùng mẫu phiếu mặc định (ans50zg.docx).";
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

// ---- Icon kiểu BigKey: đường nét 1 màu #2B579A, nét 2px trên lưới 32 (như scripts/make-icons.mjs)
const S = 'fill="none" stroke="#2B579A" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"';
const page = `<path d="M9 4h10l6 6v18H9z" ${S}/><path d="M19 4v6h6" ${S}/>`;
const ICONS = {
  chuanHoa: `<path d="M5 8h22" ${S}/><path d="M5 15h4M13 15h4M21 15h4M5 22h4M13 22h4M21 22h4" ${S}/>`,
  palette: `<path d="M16 4a12 12 0 1 0 0 24c2 0 2.5-1.5 2-3s0-3 2-3h3a5 5 0 0 0 5-5C28 9 22.5 4 16 4z" ${S}/><circle cx="10" cy="13" r="1.5" ${S}/><circle cx="15" cy="9" r="1.5" ${S}/><circle cx="21" cy="11" r="1.5" ${S}/>`,
  plusDoc: `${page}<path d="M17 15v8M13 19h8" ${S}/>`,
  minusDoc: `${page}<path d="M13 19h8" ${S}/>`,
  tag: `<path d="M5 6h12l10 10-10 10H5z" ${S}/><circle cx="11" cy="16" r="2" ${S}/>`,
  eye: `<path d="M3 16s5-8 13-8 13 8 13 8-5 8-13 8S3 16 3 16z" ${S}/><circle cx="16" cy="16" r="3.5" ${S}/>`,
  trash: `<path d="M6 9h20M12 9V6h8v3M9 9l1.5 17h11L23 9" ${S}/>`,
  rows: `<circle cx="7" cy="8" r="1.5" ${S}/><circle cx="7" cy="16" r="1.5" ${S}/><circle cx="7" cy="24" r="1.5" ${S}/><path d="M12 8h14M12 16h14M12 24h14" ${S}/>`,
  split: `<path d="M5 8h22M5 16h10M5 24h10" ${S}/><path d="M20 14v8h6M23 19l3 3-3 3" ${S}/>`,
  dot: `<path d="M5 12h16M5 20h12" ${S}/><circle cx="24" cy="20" r="1.8" ${S}/>`,
  check: `<circle cx="16" cy="16" r="11" ${S}/><path d="M10.5 16.5l3.5 3.5 7.5-8" ${S}/>`,
  table: `<rect x="5" y="7" width="22" height="18" rx="2" ${S}/><path d="M5 13h22M5 19h22M12 7v18M19 7v18" ${S}/>`,
  underline: `<path d="M10 6v8a6 6 0 0 0 12 0V6" ${S}/><path d="M8 26h16" ${S}/>`,
  pen: `${page}<path d="M12 23l7-7 2 2-7 7h-2z" ${S}/>`,
  list: `<path d="M6 7v4M5 21h3l-3 4h3" ${S}/><path d="M13 9h13M13 23h13" ${S}/>`,
  picture: `<rect x="4" y="6" width="24" height="20" rx="2" ${S}/><path d="M4 22l7-7 6 6 4-4 7 7" ${S}/><circle cx="21" cy="12" r="2" ${S}/>`,
  center: `<path d="M4 5v22M28 5v22" ${S}/><rect x="9" y="11" width="14" height="10" rx="2" ${S}/>`,
  puzzle: `<path d="M6 10h6a3 3 0 1 1 6 0h6v6a3 3 0 1 0 0 6v6h-6a3 3 0 1 0-6 0H6v-6a3 3 0 1 1 0-6z" ${S}/>`,
  scissors: `<circle cx="9" cy="22" r="4" ${S}/><circle cx="9" cy="10" r="4" ${S}/><path d="M12 12l14 12M12 20l14-12" ${S}/>`,
  arrowDown: `${page}<path d="M17 13v10M13 19l4 4 4-4" ${S}/>`,
  flask: `<path d="M13 4v9L6 25a2 2 0 0 0 2 3h16a2 2 0 0 0 2-3l-7-12V4M11 4h10M9 20h14" ${S}/>`,
  sheet: `<rect x="6" y="4" width="20" height="24" rx="2" ${S}/><circle cx="12" cy="11" r="2" ${S}/><circle cx="20" cy="11" r="2" ${S}/><circle cx="12" cy="19" r="2" ${S}/><path d="M18 19h4" ${S}/>`,
};
// icon theo lệnh (data-act / id) và theo mục (thứ tự <summary>)
const ICON_OF = {
  chuanHoaMauPA: "palette", chenCau0: "plusDoc", xoaCau0: "minusDoc",
  anHienID: "eye", toMauID: "palette", xoaIDBK: "trash", xoaNguonBK: "trash",
  xuongDongPA: "split", chamCuoiPA: "dot", xoaChamCuoiPA: "dot", chuyenDAThanhPA: "rows",
  taoBangKeyBK: "table", ghiDapAnVaoLoiGiai: "pen", gachChanTuBangKey: "underline", gachChanTuLoiGiai: "underline", gachChanHighlightDo: "underline",
  chenLoiGiai: "pen", anHienLoiGiai: "eye", xoaLoiGiai: "trash",
  danhSoThuTu: "list", sttAutoToText: "list", pictureInline: "picture", canGiuaHinh: "center",
  tachLG: "scissors", chuyenLoiGiaiXuongCuoi: "arrowDown", xoaDongVietLG: "trash", chemFormatter: "flask", zgRun: "sheet",
};
const SECTION_ICONS = ["chuanHoa", "tag", "rows", "check", "pen", "list", "picture", "puzzle"];
const svg = (name) => `<svg viewBox="0 0 32 32" aria-hidden="true">${ICONS[name] || ""}</svg>`;

/** Nút kiểu BigKey: icon + tên đậm + mô tả xám (lấy từ title). */
function decorate() {
  document.querySelectorAll("summary").forEach((s, i) => s.insertAdjacentHTML("afterbegin", svg(SECTION_ICONS[i])));
  document.querySelectorAll("button[data-act], #zgRun").forEach((b) => {
    const label = b.textContent;
    const desc = b.title;
    b.classList.add("tool");
    b.innerHTML = svg(ICON_OF[b.dataset.act || b.id]) + "<span><b></b>" + (desc ? "<small></small>" : "") + "</span>";
    b.querySelector("b").textContent = label;
    if (desc) b.querySelector("small").textContent = desc;
  });
}

Office.onReady(() => {
  decorate();
  zgTemplate = loadTemplate();
  showTemplate();
  $("#zgFile").addEventListener("change", (e) => e.target.files[0] && pickTemplate(e.target.files[0]));
  $("#zgRun").addEventListener("click", () => run(async (ctx) => PL.toPhieuZG50(ctx, $("#zgMade").value, zgTemplate || (await defaultTemplate()))));
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
