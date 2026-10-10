// Bọc một function Word thành ribbon command (ExecuteFunction):
// chạy trong Word.run, đo thời gian, hiện hộp thoại kết quả rồi gọi event.completed().
// Nếu có bản cập nhật, hộp thoại hiện thêm nút "Cập nhật" (không tự đóng).

import { getUpdateInfo, launchUpdater } from "./update.js";

const AUTO_CLOSE_MS = 3500; // hộp thoại thành công tự đóng
const MAX_WAIT_MS = 60 * 1000; // hộp thoại lỗi chờ tối đa 60 giây
const MAX_MSG = 1200; // giới hạn độ dài nội dung truyền qua URL của dialog

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
    let message;
    try {
      message = await Word.run((context) => fn(context));
    } catch (e) {
      ok = false;
      message = formatError(e);
      console.error(e);
    }
    const seconds = ((performance.now() - t0) / 1000).toFixed(1);
    try {
      const upd = await getUpdateInfo();
      await showResult({ title, message: message || "Xong.", ok, seconds, update: upd.available ? upd.latest : null });
    } catch (e) {
      console.error(e);
    } finally {
      busy = false;
      event.completed();
    }
  };
}

function formatError(e) {
  if (!e) return "Lỗi không xác định.";
  const parts = [e.message || String(e)];
  if (e.code) parts.push(`Mã lỗi: ${e.code}`);
  const loc = e.debugInfo && e.debugInfo.errorLocation;
  if (loc) parts.push(`Vị trí: ${loc}`);
  return parts.join("\n");
}

function showResult({ title, message, ok, seconds, update }) {
  const url = new URL("dialog.html", window.location.href);
  url.searchParams.set("t", title);
  url.searchParams.set("m", message.length > MAX_MSG ? message.slice(0, MAX_MSG) + "\n…" : message);
  url.searchParams.set("s", ok ? "ok" : "err");
  url.searchParams.set("d", seconds);
  if (update) url.searchParams.set("u", update);

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(url.toString(), { height: 30, width: 28, displayInIframe: true }, (res) => {
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
        clearTimeout(timer);
        try {
          dialog.close();
        } catch (_) {
          /* đã đóng */
        }
        resolve();
      };
      const timer = setTimeout(finish, ok && !update ? AUTO_CLOSE_MS : MAX_WAIT_MS);
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg && arg.message === "update") launchUpdater();
        finish();
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, finish);
    });
  });
}
