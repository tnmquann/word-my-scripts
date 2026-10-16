// Package a portable installer from dist/ (run `npm run release`, which builds first).
// Output: release/MY_SCRIPTS-v<version>/  and  release/MY_SCRIPTS-v<version>.zip
import { execFileSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const r = (...p) => path.join(ROOT, ...p);
const pkg = JSON.parse(await fs.readFile(r("package.json"), "utf8"));
const cfg = JSON.parse(await fs.readFile(r("addin.config.json"), "utf8"));

const name = `MY_SCRIPTS-${process.env.RELEASE_TAG || "v" + pkg.version}`;
const dir = r("release", name);
const zip = r("release", `${name}.zip`);

await fs.rm(dir, { recursive: true, force: true });
await fs.rm(zip, { force: true });
await fs.mkdir(dir, { recursive: true });

const crlf = (s) => s.replace(/\r?\n/g, "\r\n");
await fs.copyFile(r("dist", "manifest.xml"), path.join(dir, "manifest.xml"));
await fs.copyFile(r("dist", "install.ps1"), path.join(dir, "install.ps1"));
await fs.copyFile(r("dist", "uninstall.ps1"), path.join(dir, "uninstall.ps1"));
await fs.copyFile(r("dist", "updater.ps1"), path.join(dir, "updater.ps1"));
await fs.copyFile(r("dist", "MY_SCRIPTS_Desktop.dotm"), path.join(dir, "MY_SCRIPTS_Desktop.dotm"));
for (const f of ["install.sh", "uninstall.sh"]) await fs.writeFile(path.join(dir, f), (await fs.readFile(r("dist", f), "utf8")).replace(/\r\n/g, "\n"));
await fs.writeFile(path.join(dir, "LICENSE.txt"), crlf(await fs.readFile(r("LICENSE"), "utf8")));
for (const f of ["Install.cmd", "Uninstall.cmd"]) {
  await fs.writeFile(path.join(dir, f), crlf(await fs.readFile(r("installer", f), "utf8")));
}
await fs.writeFile(
  path.join(dir, "README.txt"),
  crlf(`MY_SCRIPTS Word add-in v${pkg.version}
=====================================

Install
  1. Close Word.
  2. Double-click Install.cmd.
  3. Open Word -> tab "${cfg.tabLabel}".
     First time only, if the tab is missing: Home > Add-ins > Advanced
     (My Add-ins) > DEVELOPER ADD-INS > MY_SCRIPTS > Add.

Uninstall
  Double-click Uninstall.cmd, then restart Word.

Updates
  The installer also sets up an auto-updater: at Windows logon (and from the
  "Cập nhật" button in Word) it checks for a new version and, if you accept,
  reinstalls it. Run updater.ps1 to check manually.

Notes
  - Requires Microsoft 365 / Word 2019+ on Windows and an internet connection
    (the add-in code is served from ${cfg.prod.baseUrl}).
  - Code updates are picked up automatically; re-run Install.cmd only when
    the manifest changes (new buttons, labels, icons).
  - If the tab does not update after a release, run:
      powershell -ExecutionPolicy Bypass -File install.ps1 -ClearCache
    with Word closed.
`)
);

execFileSync(
  "powershell",
  ["-NoProfile", "-Command", `Compress-Archive -Path '${dir}\\*' -DestinationPath '${zip}' -Force`],
  { stdio: "inherit" }
);
console.log(`Release package: ${path.relative(ROOT, dir)}\nZip: ${path.relative(ROOT, zip)}`);
