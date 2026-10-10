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
  dialog: "src/ui/dialog/dialog.js",
};
const PAGES = {
  "commands.html": "src/runtime/commands.html",
  "dialog.html": "src/ui/dialog/dialog.html",
  "about.html": "src/ui/about/about.html",
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
  const manifest = fill(await fs.readFile(r("manifest/manifest.template.xml"), "utf8"), vars);
  const left = manifest.match(/\{\{\s*[\w.]+\s*\}\}/g);
  if (left) throw new Error(`Unfilled manifest placeholders: ${left.join(", ")}`);
  await fs.writeFile(path.join(OUT, "manifest.xml"), manifest);

  await fs.mkdir(path.join(OUT, "assets"), { recursive: true });
  for (const f of await fs.readdir(r("assets/icons"))) {
    if (f.endsWith(".png")) await fs.copyFile(r("assets/icons", f), path.join(OUT, "assets", f));
  }
  // Hosted installer: `irm <BASE_URL>install.ps1 | iex`
  await fs.writeFile(path.join(OUT, "install.ps1"), fill(await fs.readFile(r("installer/install.ps1"), "utf8"), vars));
  await fs.writeFile(path.join(OUT, "uninstall.ps1"), fill(await fs.readFile(r("installer/uninstall.ps1"), "utf8"), vars));
  await fs.writeFile(path.join(OUT, ".nojekyll"), "");
  await fs.writeFile(
    path.join(OUT, "version.json"),
    JSON.stringify({ version: pkg.version, builtAt: new Date().toISOString() }, null, 2)
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
