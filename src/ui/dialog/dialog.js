// Hộp thoại kết quả: nội dung truyền qua query string (?t=title&m=message&s=ok|err&d=seconds&u=latestVersion)

const q = new URLSearchParams(location.search);
const ok = q.get("s") !== "err";

document.getElementById("root").className = ok ? "ok" : "err";
document.getElementById("badge").textContent = ok ? "✓" : "!";
document.getElementById("title").textContent = q.get("t") || "MY_SCRIPTS";
document.getElementById("msg").textContent = q.get("m") || "";
document.getElementById("meta").textContent = (ok ? "Hoàn tất" : "Lỗi") + (q.get("d") ? ` · ${q.get("d")} giây` : "");

const update = q.get("u");
if (update) {
  document.getElementById("upd").style.display = "block";
  document.getElementById("updText").textContent =
    `Có bản mới v${update}. Trình cập nhật sẽ hỏi xác nhận và yêu cầu đóng Word để cài.`;
}

Office.onReady(() => {
  document.getElementById("close").addEventListener("click", () => Office.context.ui.messageParent("close"));
  document.getElementById("updBtn").addEventListener("click", () => Office.context.ui.messageParent("update"));
});
