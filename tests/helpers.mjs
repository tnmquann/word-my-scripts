// Shared test helpers: XML DOM (jsdom) as globals, and loading a src module through esbuild.
// Features import ../runtime.js (askUser); tests stub it so a module can run outside Office.
import { createRequire } from "node:module";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";

export const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const { build } = createRequire(root + "/package.json")("esbuild");

const win = new JSDOM("").window;
globalThis.DOMParser = win.DOMParser;
globalThis.XMLSerializer = win.XMLSerializer;

/** Bundle src/<rel> and import it. askUser answers from globalThis.__choice (or the first button). */
export async function loadModule(rel) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "myscripts-"));
  const stub = path.join(tmp, "runtime-stub.js");
  fs.writeFileSync(stub, "export const askUser=async o=>globalThis.__choice||o.buttons[0];export const isDesktop=()=>globalThis.__desktop!==false;");
  const out = await build({
    entryPoints: [path.join(root, "src", rel)], bundle: true, format: "esm", write: false, platform: "node",
    plugins: [{ name: "stub-runtime", setup(b) { b.onResolve({ filter: /features[\\/]runtime\.js$|^\.\.\/runtime\.js$/ }, () => ({ path: stub })); } }],
  });
  return import("data:text/javascript;base64," + Buffer.from(out.outputFiles[0].text).toString("base64"));
}

let failures = 0;
export function check(ok, name, info = "") {
  if (!ok) failures++;
  console.log(ok ? "PASS" : "FAIL", name, info);
}
export function done() {
  if (failures) process.exitCode = 1;
}
