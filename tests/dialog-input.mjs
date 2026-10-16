// Result/ask dialog (src/ui/dialog): askUser with a number input renders the field and sends "btn:<name>\n<value>".
import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";
import { root, loadModule, check, done } from "./helpers.mjs";

const html = fs.readFileSync(path.join(root, "src/ui/dialog/dialog.html"), "utf8").replace(/<script[^>]*><\/script>/g, "");
const dom = new JSDOM(html, { url: "https://example.test/dialog.html" });
Object.assign(globalThis, { window: dom.window, document: dom.window.document, location: dom.window.location, URLSearchParams: dom.window.URLSearchParams });
const sent = [];
let onMsg;
globalThis.Office = {
  EventType: { DialogParentMessageReceived: "m" },
  onReady: (cb) => cb(),
  context: { ui: { messageParent: (m) => sent.push(m), addHandlerAsync: (_, h) => (onMsg = h) } },
};
await loadModule("ui/dialog/dialog.js");
onMsg({ message: JSON.stringify({ t: "Fix lỗi tab câu hỏi", m: "Tìm thấy 3 câu", s: "warn", b: ["Đánh số", "Huỷ"], i: { label: "Bắt đầu từ", value: 1 } }) });

const $ = (s) => document.querySelector(s);
check($("#ask input") && $("#ask input").value === "1", "number field rendered with default 1");
$("#ask input").value = "31";
$("footer button").click();
check(sent[sent.length - 1] === "btn:Đánh số\n31", "primary button sends name + value", JSON.stringify(sent));
$("#ask input").value = "7";
$("#ask input").dispatchEvent(new dom.window.KeyboardEvent("keydown", { key: "Enter" }));
check(sent[sent.length - 1] === "btn:Đánh số\n7", "Enter submits the primary button", JSON.stringify(sent));
done();
