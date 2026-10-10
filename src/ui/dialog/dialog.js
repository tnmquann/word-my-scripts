// Hộp thoại kết quả. Dữ liệu { t, m, x, s, d, u }:
//   t = tiêu đề, m = tóm tắt (mỗi dòng 1 ý), x = chi tiết kỹ thuật (thu gọn),
//   s = ok | err, d = số giây, u = phiên bản mới
// Nhận qua messageChild (DialogApi 1.2) sau khi gửi "ready"; fallback: query string.
// Không tự đóng: người dùng bấm "Đóng" hoặc nút X.

const $ = (id) => document.getElementById(id);

const ICON_OK = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
const ICON_ERR = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M10 5.5v5.5"/><path d="M10 14.5v.01"/></svg>';

function render(data) {
  const ok = data.s !== "err";
  $("root").className = ok ? "ok" : "err";
  $("icon").innerHTML = ok ? ICON_OK : ICON_ERR;
  $("title").textContent = data.t || "MY_SCRIPTS";
  $("status").textContent = (ok ? "Hoàn tất" : "Không thành công") + (data.d ? ` · ${data.d} giây` : "");

  $("lines").textContent = "";
  for (const line of String(data.m || "").split("\n")) {
    if (!line.trim()) continue;
    const li = document.createElement("li");
    li.textContent = line.trim();
    $("lines").appendChild(li);
  }
  if (data.x) {
    $("details").textContent = data.x;
    $("more").hidden = false;
  }
  if (data.u) {
    $("upd").style.display = "flex";
    $("updText").textContent = `Có phiên bản mới v${data.u}`;
  }
  $("root").style.visibility = "visible";
}

const q = new URLSearchParams(location.search);
if (q.has("t")) render(Object.fromEntries(q.entries()));

Office.onReady(() => {
  $("close").addEventListener("click", () => Office.context.ui.messageParent("close"));
  $("updBtn").addEventListener("click", () => Office.context.ui.messageParent("update"));
  if (!q.has("t")) {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      try {
        render(JSON.parse(arg.message));
      } catch (e) {
        render({ t: "MY_SCRIPTS", m: "Không đọc được kết quả.", s: "err" });
      }
    });
    Office.context.ui.messageParent("ready");
  }
  $("close").focus();
});
