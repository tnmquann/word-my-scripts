// Cầu nối tới API của runtime (src/runtime/commands.js gán window.__MY_SCRIPTS_RUNTIME__).
// Features không import trực tiếp runtime để bundle features có thể hot-reload độc lập.

/**
 * Hỏi người dùng bằng hộp thoại Office (thay cho confirm/alert, vốn không hiện trong ribbon command).
 * Không có runtime (vd. chạy thử ngoài add-in) -> trả về nút đầu tiên.
 * @param {{title: string, lines: string[], buttons: string[]}} opts
 * @returns {Promise<string>}
 */
export function askUser(opts) {
  const rt = window.__MY_SCRIPTS_RUNTIME__;
  if (rt && typeof rt.askUser === "function") return rt.askUser(opts);
  console.warn(opts.title, opts.lines);
  return Promise.resolve(opts.buttons[0]);
}
