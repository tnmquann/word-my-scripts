// Hộp thoại kết quả. Query string:
//   t = tiêu đề, m = tóm tắt (mỗi dòng 1 ý), x = chi tiết kỹ thuật (tuỳ chọn, thu gọn),
//   s = ok | err, d = số giây, u = phiên bản mới (tuỳ chọn)
// Không tự đóng: người dùng bấm "Đóng" hoặc nút X.

const q = new URLSearchParams(location.search);
const ok = q.get("s") !== "err";
const $ = (id) => document.getElementById(id);

const ICON_OK = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
const ICON_ERR = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M10 5.5v5.5"/><path d="M10 14.5v.01"/></svg>';

$("root").className = ok ? "ok" : "err";
$("icon").innerHTML = ok ? ICON_OK : ICON_ERR;
$("title").textContent = q.get("t") || "MY_SCRIPTS";
$("status").textContent = (ok ? "Hoàn tất" : "Không thành công") + (q.get("d") ? ` · ${q.get("d")} giây` : "");

for (const line of (q.get("m") || "").split("\n")) {
  if (!line.trim()) continue;
  const li = document.createElement("li");
  li.textContent = line.trim();
  $("lines").appendChild(li);
}

const details = q.get("x");
if (details) {
  $("details").textContent = details;
  $("more").hidden = false;
}

const update = q.get("u");
if (update) {
  $("upd").style.display = "flex";
  $("updText").textContent = `Đã có phiên bản mới v${update}.`;
}

Office.onReady(() => {
  $("close").addEventListener("click", () => Office.context.ui.messageParent("close"));
  $("updBtn").addEventListener("click", () => Office.context.ui.messageParent("update"));
  $("close").focus();
});
