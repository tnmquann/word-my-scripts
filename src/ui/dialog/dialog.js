// Hộp thoại kết quả. Dữ liệu { t, m, x, s, d, u }:
//   t = tiêu đề, m = tóm tắt (mỗi dòng 1 ý), x = chi tiết kỹ thuật (thu gọn),
//   s = ok | err, d = số giây, u = phiên bản mới
// Nhận qua messageChild (DialogApi 1.2) sau khi gửi "ready"; fallback: query string.
// Không tự đóng: người dùng bấm "Đóng" hoặc nút X.

const $ = (id) => document.getElementById(id);

const ICON_OK = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4.5 10.5l3.5 3.5 7.5-8"/></svg>';
const ICON_ERR = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M10 5.5v5.5"/><path d="M10 14.5v.01"/></svg>';

function render(data) {
  const state = data.s === "err" ? "err" : data.s === "warn" ? "warn" : "ok";
  $("root").className = state;
  $("icon").innerHTML = state === "ok" ? ICON_OK : ICON_ERR;
  $("title").textContent = data.t || "MY_SCRIPTS";
  $("status").textContent =
    state === "warn" ? "Cần bạn xác nhận" : (state === "ok" ? "Hoàn tất" : "Không thành công") + (data.d ? ` · ${data.d} giây` : "");

  // Câu hỏi (askUser): thay nút "Đóng" bằng các nút được yêu cầu; nút đầu tiên là nút chính
  const buttons = Array.isArray(data.b) ? data.b : data.b ? String(data.b).split("|") : null;
  // Ô nhập số (askUser có input): giá trị gửi kèm sau tên nút: "btn:<nút>\n<giá trị>"
  let field = null;
  if (data.i && typeof data.i === "object") {
    const wrap = document.createElement("label");
    wrap.id = "ask";
    wrap.textContent = data.i.label || "";
    field = document.createElement("input");
    field.type = "number";
    field.min = "1";
    field.value = data.i.value != null ? String(data.i.value) : "";
    wrap.appendChild(field);
    $("body").appendChild(wrap);
  }
  if (buttons && buttons.length) {
    const footer = $("close").parentNode;
    footer.textContent = "";
    buttons.forEach((name, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = name;
      if (i === 0) btn.className = "primary";
      btn.addEventListener("click", () => Office.context.ui.messageParent("btn:" + name + (field ? "\n" + field.value.trim() : "")));
      footer.appendChild(btn);
    });
    if (field) field.addEventListener("keydown", (e) => e.key === "Enter" && footer.firstChild.click());
  }

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
  const first = field || document.querySelector("footer button");
  if (first) first.focus();
  if (field) field.select();
}

const q = new URLSearchParams(location.search);
if (q.has("t")) render(Object.fromEntries(q.entries()));

Office.onReady(() => {
  const closeBtn = $("close"); // có thể đã được thay bằng nút của askUser
  if (closeBtn) closeBtn.addEventListener("click", () => Office.context.ui.messageParent("close"));
  $("updBtn").addEventListener("click", () => {
    $("updBtn").disabled = true;
    $("updBtn").textContent = "Đã mở";
    $("updText").textContent = "Đã mở trang cập nhật trong trình duyệt. Làm theo hướng dẫn ở đó; trình cập nhật sẽ báo khi xong.";
    Office.context.ui.messageParent("update");
  });
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
});
