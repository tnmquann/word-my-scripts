// Build the add-in.
//   node scripts/build.mjs                 -> dist/      (production, GitHub Pages URL, hashed + minified)
//   node scripts/build.mjs --dev           -> dist-dev/  (localhost URL, separate add-in ID, "(DEV)" labels)
//   node scripts/build.mjs --dev --serve   -> dist-dev/ + watch + HTTPS dev server
import * as esbuild from "esbuild";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const r = (...p) => path.join(ROOT, ...p);

const args = new Set(process.argv.slice(2));
const DEV = args.has("--dev");
const SERVE = args.has("--serve");

const cfg = JSON.parse(await fs.readFile(r("addin.config.json"), "utf8"));
const pkg = JSON.parse(await fs.readFile(r("package.json"), "utf8"));
const env = DEV ? cfg.dev : cfg.prod;
const OUT = r(DEV ? "dist-dev" : "dist");
const BASE_URL = env.baseUrl.endsWith("/") ? env.baseUrl : env.baseUrl + "/";

const ENTRIES = {
  commands: "src/runtime/commands.js",
  features: "src/features/index.js", // hot-reloaded per click via version.json "features"
  dialog: "src/ui/dialog/dialog.js",
  xoa: "src/ui/xoa/xoa.js", // task pane "Xoá mức độ…" (nạp lại mỗi lần mở pane)
  bigkey: "src/ui/bigkey/bigkey.js", // task pane "Công cụ BigKey"
};
const PAGES = {
  "commands.html": "src/runtime/commands.html",
  "dialog.html": "src/ui/dialog/dialog.html",
  "start.html": "src/ui/start/start.html",
  "author.html": "src/ui/author/author.html",
  "update.html": "src/ui/update/update.html",
  "xoa.html": "src/ui/xoa/xoa.html",
  "bigkey.html": "src/ui/bigkey/bigkey.html",
};

const esbuildOptions = {
  absWorkingDir: ROOT,
  entryPoints: ENTRIES,
  outdir: OUT,
  bundle: true,
  format: "iife",
  target: ["es2017"],
  minify: !DEV,
  sourcemap: DEV ? "inline" : false,
  entryNames: DEV ? "[name]" : "[name]-[hash]",
  metafile: true,
  logLevel: "info",
  legalComments: "none",
};

/** map entry name -> emitted file name (handles hashed names) */
function outputNames(metafile) {
  const names = {};
  for (const [file, info] of Object.entries(metafile.outputs)) {
    if (!info.entryPoint) continue;
    const entry = Object.keys(ENTRIES).find((k) => path.normalize(ENTRIES[k]) === path.normalize(info.entryPoint));
    if (entry) names[`${entry}.js`] = path.basename(file);
  }
  return names;
}

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch]);

const fill = (text, vars) => text.replace(/\{\{\s*([\w.]+)\s*\}\}/g, (m, k) => (k in vars ? vars[k] : m));

async function writeStatic(names) {
  const vars = {
    ...names,
    version: pkg.version,
    ID: env.id,
    VERSION: `${pkg.version}.0`,
    BASE_URL,
    ORIGIN: new URL(BASE_URL).origin,
    DISPLAY_NAME: cfg.displayName + (DEV ? env.labelSuffix : ""),
    PROVIDER: cfg.providerName,
    TAB_LABEL: cfg.tabLabel + (DEV ? env.labelSuffix : ""),
    ...Object.fromEntries(Object.entries(cfg.author || {}).map(([k, v]) => [`author.${k}`, escapeHtml(v)])),
  };
  for (const [out, src] of Object.entries(PAGES)) {
    await fs.writeFile(path.join(OUT, out), fill(await fs.readFile(r(src), "utf8"), vars));
  }
  // Icon URLs get ?v=<version>: Office caches ribbon images by URL, so new icons only show when the URL changes.
  const manifest = fill(await fs.readFile(r("manifest/manifest.template.xml"), "utf8"), vars).replace(
    /(assets\/[\w-]+\.png)"/g,
    `$1?v=${pkg.version}"`
  );
  const left = manifest.match(/\{\{\s*[\w.]+\s*\}\}/g);
  if (left) throw new Error(`Unfilled manifest placeholders: ${left.join(", ")}`);
  await fs.writeFile(path.join(OUT, "manifest.xml"), manifest);

  // Mẫu tài liệu dùng chung (vd. phiếu ZipGrade 50 câu ans50zg.docx — của maintainer)
  await fs.mkdir(path.join(OUT, "templates"), { recursive: true });
  for (const f of await fs.readdir(r("assets/templates"))) await fs.copyFile(r("assets/templates", f), path.join(OUT, "templates", f));
  await fs.mkdir(path.join(OUT, "assets"), { recursive: true });
  for (const f of await fs.readdir(r("assets/icons"))) {
    if (f.endsWith(".png")) await fs.copyFile(r("assets/icons", f), path.join(OUT, "assets", f));
  }
  // Hosted installer: `irm <BASE_URL>install.ps1 | iex`
  // PowerShell scripts.
  // - install.ps1 / uninstall.ps1 are run via `irm <url> | iex`: irm keeps a BOM as U+FEFF in the string,
  //   which breaks parsing -> must be ASCII-only and written WITHOUT BOM.
  // - updater.ps1 contains Vietnamese and is run with -File on Windows PowerShell 5.1 -> UTF-8 WITH BOM.
  for (const [ps1, bom] of [["install.ps1", false], ["uninstall.ps1", false], ["updater.ps1", true]]) {
    const text = fill(await fs.readFile(r("installer", ps1), "utf8"), vars).replace(/^\uFEFF/, "");
    if (!bom && /[^\x00-\x7f]/.test(text)) throw new Error(`${ps1} must be ASCII-only (it is run via irm | iex)`);
    await fs.writeFile(path.join(OUT, ps1), (bom ? "\uFEFF" : "") + text);
  }
  // macOS installers (unverified): `curl -fsSL <BASE_URL>install.sh | bash` -> LF line endings, ASCII only
  for (const sh of ["install.sh", "uninstall.sh"]) {
    const text = fill(await fs.readFile(r("installer", sh), "utf8"), vars).replace(/\r\n/g, "\n");
    if (/[^\x00-\x7f]/.test(text)) throw new Error(`${sh} must be ASCII-only`);
    await fs.writeFile(path.join(OUT, sh), text);
  }
  // Companion VBA template (MathType tools), built on Windows by scripts/build-companion.ps1 and committed
  await fs.copyFile(r("installer", "MY_SCRIPTS_Desktop.dotm"), path.join(OUT, "MY_SCRIPTS_Desktop.dotm"));
  await fs.writeFile(path.join(OUT, ".nojekyll"), "");
  await fs.writeFile(
    path.join(OUT, "version.json"),
    // "features": current features bundle; the runtime loads it on the next click if it differs
    JSON.stringify({ version: pkg.version, features: names["features.js"], builtAt: new Date().toISOString() }, null, 2)
  );
}

await fs.rm(OUT, { recursive: true, force: true });
await fs.mkdir(OUT, { recursive: true });

if (!SERVE) {
  const result = await esbuild.build(esbuildOptions);
  await writeStatic(outputNames(result.metafile));
  console.log(`\nBuilt ${DEV ? "DEV" : "PROD"} v${pkg.version} -> ${path.relative(ROOT, OUT)} (base ${BASE_URL})`);
} else {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const keyfile = path.join(certDir, "localhost.key");
  const certfile = path.join(certDir, "localhost.crt");
  try {
    await fs.access(keyfile);
    await fs.access(certfile);
  } catch {
    console.error("Dev HTTPS certificate not found. Run `npm run certs` once, then retry.");
    process.exit(1);
  }
  const ctx = await esbuild.context({
    ...esbuildOptions,
    plugins: [
      {
        name: "static-pages",
        setup(build) {
          build.onEnd(async (res) => {
            if (res.metafile) await writeStatic(outputNames(res.metafile));
          });
        },
      },
    ],
  });
  await ctx.watch();
  const { port } = await ctx.serve({ servedir: OUT, port: env.port, keyfile, certfile });
  console.log(`\nDev server: https://localhost:${port}/  (watching src/; Ctrl+C to stop)`);
  console.log("Sideload once with: npm run sideload:dev  (then restart Word)");
}
