// Hộp thoại kết quả: nội dung truyền qua query string (?t=title&m=message&s=ok|err&d=seconds)

const q = new URLSearchParams(location.search);
const ok = q.get("s") !== "err";

document.getElementById("root").className = ok ? "ok" : "err";
document.getElementById("badge").textContent = ok ? "✓" : "!";
document.getElementById("title").textContent = q.get("t") || "MY_SCRIPTS";
document.getElementById("msg").textContent = q.get("m") || "";
document.getElementById("meta").textContent = (ok ? "Hoàn tất" : "Lỗi") + (q.get("d") ? ` · ${q.get("d")} giây` : "");

Office.onReady(() => {
  document.getElementById("close").addEventListener("click", () => Office.context.ui.messageParent("close"));
});
