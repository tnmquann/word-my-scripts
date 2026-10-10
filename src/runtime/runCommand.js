// Bọc một function Word thành ribbon command (ExecuteFunction):
// chạy trong Word.run, đo thời gian, hiện hộp thoại kết quả rồi gọi event.completed().
// Hộp thoại KHÔNG tự đóng — người dùng bấm "Đóng" (hoặc nút X) khi đã đọc xong.

import { getUpdateInfo, launchUpdater } from "./update.js";

const MAX_SUMMARY = 600; // tóm tắt hiển thị
const MAX_DETAILS = 400; // chi tiết kỹ thuật (thu gọn mặc định)

let busy = false;

export function makeCommand(title, fn) {
  return async (event) => {
    if (busy) {
      event.completed();
      return;
    }
    busy = true;
    const t0 = performance.now();
    let ok = true;
    let summary;
    let details = "";
    try {
      summary = await Word.run((context) => fn(context));
    } catch (e) {
      ok = false;
      ({ summary, details } = describeError(e));
      console.error(e);
    }
    const seconds = ((performance.now() - t0) / 1000).toFixed(1);
    try {
      const upd = await getUpdateInfo();
      await showResult({ title, summary: summary || "Xong.", details, ok, seconds, update: upd.available ? upd.latest : null });
    } catch (e) {
      console.error(e);
    } finally {
      busy = false;
      event.completed();
    }
  };
}

/** Lỗi -> câu ngắn dễ hiểu + chi tiết kỹ thuật (ẩn sau "Chi tiết") */
function describeError(e) {
  if (!e) return { summary: "Đã xảy ra lỗi không xác định.", details: "" };
  const raw = String(e.message || e).trim();
  const firstLine = raw.split("\n")[0];
  const known = {
    GeneralException: "Word không thực hiện được thao tác trên đoạn này của tài liệu.",
    InvalidArgument: "Tài liệu có cấu trúc chưa được hỗ trợ.",
    ItemNotFound: "Không tìm thấy nội dung cần xử lý.",
    AccessDenied: "Tài liệu đang bị khoá hoặc chỉ cho phép đọc.",
    ApiNotFound: "Phiên bản Word này chưa hỗ trợ chức năng.",
  };
  const summary = (e.code && known[e.code]) || (firstLine.length > 160 ? firstLine.slice(0, 160) + "…" : firstLine);
  const parts = [];
  if (e.code) parts.push(`Mã: ${e.code}`);
  const loc = e.debugInfo && e.debugInfo.errorLocation;
  if (loc) parts.push(`Vị trí: ${loc}`);
  if (raw !== summary) parts.push(raw);
  return { summary, details: parts.join("\n") };
}

const clip = (s, n) => (s.length > n ? s.slice(0, n) + "…" : s);

// Kích thước hộp thoại (% màn hình) — giữ nhỏ gọn
const DIALOG_SIZE = { height: 24, width: 22, displayInIframe: true };

/**
 * Mở dialog.html KHÔNG kèm query string (thanh tiêu đề Office luôn hiện URL — giữ URL ngắn),
 * rồi gửi nội dung qua messageChild khi dialog báo "ready" (DialogApi 1.2).
 * Word cũ không có DialogApi 1.2 -> fallback truyền qua query string.
 */
function showResult({ title, summary, details, ok, seconds, update }) {
  const payload = {
    t: title,
    m: clip(summary, MAX_SUMMARY),
    x: details ? clip(details, MAX_DETAILS) : "",
    s: ok ? "ok" : "err",
    d: seconds,
    u: update || "",
  };
  const canMessageChild = Office.context.requirements.isSetSupported("DialogApi", "1.2");
  const url = new URL("dialog.html", window.location.href);
  if (!canMessageChild) {
    for (const [k, v] of Object.entries(payload)) if (v) url.searchParams.set(k, v);
  }

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(url.toString(), DIALOG_SIZE, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("displayDialogAsync failed", res.error);
        resolve();
        return;
      }
      const dialog = res.value;
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        try {
          dialog.close();
        } catch (_) {
          /* đã đóng */
        }
        resolve();
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const msg = arg && arg.message;
        if (msg === "ready") {
          if (canMessageChild) dialog.messageChild(JSON.stringify(payload));
          return;
        }
        if (msg === "update") launchUpdater();
        finish();
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, finish); // người dùng bấm X
    });
  });
}

