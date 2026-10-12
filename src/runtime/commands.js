// Entry của FunctionFile: đăng ký mọi ribbon command (ExecuteFunction).
// - Lệnh hệ thống (nhóm "Thông tin") nằm trong runtime.
// - Lệnh của các nhóm (src/features/<group>/) nằm trong bundle features riêng, được nạp lại
//   tự động khi có bản mới (src/runtime/loader.js) -> bản vá có hiệu lực ngay, không cần mở lại Word.
// Key phải khớp <FunctionName> trong manifest/manifest.template.xml và không trùng giữa các nhóm.

import { makeCommand, askUser } from "./runCommand.js";

// API runtime cho bundle features (features không import runtime để có thể hot-reload độc lập)
window.__MY_SCRIPTS_RUNTIME__ = { askUser };
import { checkUpdate } from "./update.js";
import { initFeatures, ensureLatestFeatures } from "./loader.js";

// Command hệ thống (nhóm "Thông tin" trên ribbon)
const system = {
  checkUpdate: { label: "Kiểm tra cập nhật", run: checkUpdate },
};

/** Handler cho 1 lệnh feature: luôn tra registry mới nhất lúc bấm nút. */
function featureCommand(name, fallbackLabel) {
  return async (event) => {
    let groups;
    try {
      groups = await ensureLatestFeatures();
    } catch (e) {
      groups = (window.__MY_SCRIPTS_FEATURES__ || { groups: {} }).groups;
    }
    for (const commands of Object.values(groups)) {
      if (commands[name]) return makeCommand(commands[name].label, commands[name].run)(event);
    }
    return makeCommand(fallbackLabel, async () => {
      throw new Error(`Không tìm thấy lệnh "${name}". Hãy kiểm tra cập nhật.`);
    })(event);
  };
}

// commands.html nạp features-<hash>.js bằng <script> đồng bộ TRƯỚC file này,
// nên registry đã sẵn sàng -> đăng ký action đồng bộ (tránh lần bấm đầu tiên bị lỡ).
Office.onReady(() => {
  for (const [name, { label, run }] of Object.entries(system)) Office.actions.associate(name, makeCommand(label, run));

  initFeatures(document.querySelector("meta[name=features]").getAttribute("content"));
  const seen = new Set(Object.keys(system));
  for (const [group, commands] of Object.entries(window.__MY_SCRIPTS_FEATURES__.groups)) {
    for (const [name, { label }] of Object.entries(commands)) {
      if (seen.has(name)) throw new Error(`Duplicate command "${name}" in group "${group}"`);
      seen.add(name);
      Office.actions.associate(name, featureCommand(name, label));
    }
  }
});
