// Smoke test of the built runtime (dist/) against a fake Office/Word:
//  1. every command is registered,
//  2. hot reload: when version.json points to a new features bundle, the next click runs the new code,
//  3. askUser: dialog buttons round-trip, closing with X returns the last (safe) button.
// Run after `npm run build`: node tests/runtime-smoke.cjs
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const dist = path.join(__dirname, "..", "dist");
const commandsJs = fs.readdirSync(dist).find((f) => /^commands-.*\.js$/.test(f));
const featuresJs = fs.readdirSync(dist).find((f) => /^features-.*\.js$/.test(f));

const NEW_FILE = "features-NEWBUILD.js";
const newBundle = `window.__MY_SCRIPTS_FEATURES__ = { groups: { dgnl: {
  chuanHoaToan: { label: "Chuẩn hoá Toán ĐGNL", run: async () => "RAN_NEW_BUNDLE" } } } };`;
let versionFeatures = featuresJs;
let reply = null; // what the fake user clicks in the next dialog: "btn:<name>" | "close" | null (X)
const handlers = {};
const payloads = [];
const ctx = {
  console, setTimeout, clearTimeout, URL, URLSearchParams, performance, AbortController,
  location: { href: "https://x.test/word-my-scripts/commands.html?mv=1.0.0.0", search: "?mv=1.0.0.0" },
  fetch: async () => ({ ok: true, json: async () => ({ version: "1.0.0", features: versionFeatures }) }),
  Word: { run: async (fn) => fn({}) },
  Office: {
    onReady: (cb) => cb(),
    actions: { associate: (name, fn) => (handlers[name] = fn) },
    AsyncResultStatus: { Succeeded: "succeeded" },
    EventType: { DialogMessageReceived: "m", DialogEventReceived: "e" },
    context: {
      requirements: { isSetSupported: () => true },
      ui: {
        displayDialogAsync: (url, opts, cb) => {
          const d = {
            h: {}, addEventHandler(t, fn) { this.h[t] = fn; }, close() {},
            messageChild(m) {
              payloads.push(JSON.parse(m));
              setTimeout(() => (reply === null ? this.h.e({ error: 12006 }) : this.h.m({ message: reply })), 0);
            },
          };
          cb({ status: "succeeded", value: d });
          setTimeout(() => d.h.m({ message: "ready" }), 0);
        },
      },
    },
  },
};
ctx.window = ctx;
ctx.document = {
  querySelector: () => ({ getAttribute: () => featuresJs }),
  createElement: () => ({}),
  head: {
    appendChild: (s) => {
      const file = s.src.split("/").pop();
      vm.runInContext(file === NEW_FILE ? newBundle : fs.readFileSync(path.join(dist, file), "utf8"), ctx);
      setTimeout(() => s.onload(), 0);
    },
  },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(dist, featuresJs), "utf8"), ctx);
vm.runInContext(fs.readFileSync(path.join(dist, commandsJs), "utf8"), ctx);

(async () => {
  const names = Object.keys(handlers);
  console.log("registered:", names.join(", "));
  for (const n of ["checkUpdate", "chuanHoaToan", "taoBangDapAn", "latexToEquation", "equationToLatex"]) assert(names.includes(n), "missing " + n);

  // hot reload
  reply = "close";
  versionFeatures = NEW_FILE;
  await handlers.chuanHoaToan({ completed() {} });
  assert.strictEqual(payloads[payloads.length - 1].m, "RAN_NEW_BUNDLE", "hot reload did not run the new bundle");
  console.log("hot reload: PASS");

  // askUser
  const ask = ctx.__MY_SCRIPTS_RUNTIME__.askUser;
  const opts = { title: "t", lines: ["l"], buttons: ["Tiếp tục", "Dừng lại"] };
  reply = "btn:Tiếp tục";
  assert.strictEqual(await ask(opts), "Tiếp tục");
  reply = null;
  assert.strictEqual(await ask(opts), "Dừng lại", "X must return the last button");
  // number input: "btn:<name>\n<value>" -> { button, value }; X -> last button + default value
  const optsIn = { ...opts, input: { label: "Bắt đầu từ", value: 1 } };
  reply = "btn:Tiếp tục\n31";
  assert.strictEqual(JSON.stringify(await ask(optsIn)), JSON.stringify({ button: "Tiếp tục", value: "31" }));
  reply = null;
  assert.strictEqual(JSON.stringify(await ask(optsIn)), JSON.stringify({ button: "Dừng lại", value: "1" }), "X with input");
  console.log("askUser: PASS");
})().catch((e) => {
  console.error("FAIL:", e.message);
  process.exit(1);
});
