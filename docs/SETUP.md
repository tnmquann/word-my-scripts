# Setup & operations guide

This guide covers:

1. The architecture.
2. First-time publishing to GitHub Pages.
3. Installing on a user machine.
4. The development loop.
5. Releasing.
6. Adding a new command.
7. Troubleshooting.

---

## 1. Architecture

```
 Word (desktop)                              GitHub Pages (HTTPS)
 ┌────────────────────────────┐              ┌─────────────────────────────┐
 │ Ribbon: MY_SCRIPTS tab     │  loads       │ commands.html + commands-*.js│
 │  └ ĐGNL menu → 5 items  ───┼─────────────▶│ dialog.html  + dialog-*.js  │
 │                            │              │ start/author.html, assets/* │
 │ HKCU\…\WEF\Developer       │              │ manifest.xml, install.ps1   │
 │   {add-in id} = manifest   │              └─────────────────────────────┘
 └────────────────────────────┘
```

- **Manifest** (`manifest/manifest.template.xml`): an add-in-only XML manifest that defines a `CustomTab` → `Group` → `Menu` with five `ExecuteFunction` items. `scripts/build.mjs` fills in the placeholders from `addin.config.json` and `package.json`.
- **FunctionFile** (`commands.html`): a page without UI. Word loads it when a ribbon button is clicked. The page loads two bundles:

- **features** (`src/features/index.js` → `features-<hash>.js`): the registries of all feature groups, exposed as `window.__MY_SCRIPTS_FEATURES__`.
- **runtime** (`src/runtime/commands.js`): binds every `<FunctionName>` with `Office.actions.associate`. Duplicate names across groups throw an error at load time.
- **Hot reload**: Word keeps this page in memory for the whole session. Before each feature command, `src/runtime/loader.js` therefore fetches `version.json` (no-store, 3 s timeout). If its `features` file name differs from the loaded one, it injects the new bundle and then runs the command from the newest registry. Feature fixes apply on the next click without a Word restart. Changes to the runtime, the dialog or the manifest still require a restart or reinstall.
- **runCommand** (`src/runtime/runCommand.js`): wraps every handler. It does four things:
  - Runs the handler inside `Word.run` and measures how long it takes.
  - Turns `OfficeExtension.Error` into readable text (message, code, `debugInfo.errorLocation`).
  - Opens `dialog.html` with `displayDialogAsync` (24% × 22% of the screen) **without a query string**, so the Office caption shows only a short URL. When the dialog reports `ready`, it sends the result via `dialog.messageChild` (DialogApi 1.2). On older builds it falls back to the query string. The payload is: a summary capped at 600 characters, plus optional technical details capped at 400 characters and shown collapsed. Known error codes are mapped to short Vietnamese messages. The dialog stays open until the user closes it.
  - Always calls `event.completed()`.

  A module-level `busy` flag drops re-entrant clicks while a command is still running.
- **Author pane**: the *Thông tin → Tác giả* button (`ShowTaskpane`) opens `author.html`. It is the only place the author card is shown. The manifest's default `SourceLocation` is `start.html`, a minimal welcome page without author information, because Word may open that page right after the add-in is added. Its author card is filled at build time from the `author` block in `addin.config.json`, with values HTML-escaped. That block mirrors the public GitHub profile of `tnmquann`. To refresh it, run `gh api users/tnmquann`, update the block, then rebuild and push. The avatar is loaded from `avatars.githubusercontent.com`.
- **Registration**: the installer writes a per-user registry value:

  ```
  HKCU\Software\Microsoft\Office\16.0\WEF\Developer
      <add-in GUID> = %LOCALAPPDATA%\MY_SCRIPTS\manifests\<add-in GUID>.xml
  ```

  Word reads this key at startup, so the add-in loads for every document. No admin rights, network share or Microsoft 365 admin center is needed.
- **Hosting**: Office add-ins must be served over HTTPS. `dist/` is published to GitHub Pages by `.github/workflows/deploy.yml`. Changes to code alone go live on the next push. Users only need to re-run the installer when the manifest changes, for example new buttons, labels or icons.
- **Two identities**: prod and dev use different GUIDs (`addin.config.json`). The dev build shows as **MY_SCRIPTS (DEV)** and points to `https://localhost:3000/`, so it can be installed next to the prod add-in.

### Build outputs

| Path | Produced by | Contents |
| --- | --- | --- |
| `dist/` | `npm run build` | Minified bundles with hashed names, HTML, icons, `manifest.xml`, hosted `install.ps1`, `.nojekyll`, `version.json` |
| `dist-dev/` | `npm run build:dev` / `npm run dev` | Unminified bundles with inline source maps, manifest pointing to localhost |
| `release/MY_SCRIPTS-vX.Y.Z[.zip]` | `npm run release` | Portable installer: `manifest.xml`, `install.ps1`, `uninstall.ps1`, `Install.cmd`, `Uninstall.cmd`, `README.txt` |

Bundles are built with esbuild (`target: es2017`, IIFE). Hashed file names prevent stale JS after a deploy. Only the HTML is cached by name, and GitHub Pages serves it with `max-age=600`.

---

## 2. First-time publishing (once per repository)

Prerequisites: Node.js 20 or later, Git, and the GitHub CLI (`gh auth login`).

1. Check `addin.config.json`. `prod.baseUrl` must match the Pages URL, which follows the pattern `https://<user>.github.io/<repo>/`. If you rename the repository, update the URL, rebuild, and reinstall on every machine, because the URL is stored in the manifest.
2. Create the repository and push it:

   ```powershell
   gh repo create word-my-scripts --public --source . --push
   ```

   GitHub Pages on a private repository needs a paid plan.
3. Enable Pages with GitHub Actions as the source:

   ```powershell
   gh api -X POST repos/{owner}/word-my-scripts/pages -f build_type=workflow
   ```

   Or go to **Settings → Pages → Build and deployment → Source: GitHub Actions**.
4. Wait for the **Deploy to GitHub Pages** workflow to finish (`gh run watch`), then check `https://<user>.github.io/word-my-scripts/version.json`.

---

## 3. Installing on a user machine

Close Word first. Choose one of the options below.

**A. One-liner** (downloads the published manifest):

```powershell
irm https://tnmquann.github.io/word-my-scripts/install.ps1 | iex
```

**B. Portable package**: copy `MY_SCRIPTS-vX.Y.Z.zip` from `release/` or from the GitHub Release assets, extract it, and double-click `Install.cmd`.

**C. From a clone of this repository**:

```powershell
npm ci; npm run build
powershell -ExecutionPolicy Bypass -File installer\install.ps1 -ManifestPath dist\manifest.xml
```

Open Word. The **MY_SCRIPTS** tab appears on the ribbon.

**First run on recent Word builds** (seen on 16.0.20606): Word registers the add-in but may not activate it on its own. If the tab is missing, go to **Home → Add-ins → Advanced** (or *My Add-ins*), open **DEVELOPER ADD-INS**, select **MY_SCRIPTS** and click **Add**. You only need to do this once per machine. The tab then persists across restarts and updates.

`install.ps1` takes these switches:

| Parameter | Meaning |
| --- | --- |
| `-ManifestPath <file>` | Register a specific manifest file |
| `-ManifestUrl <url>` | Download the manifest from this URL (defaults to the prod Pages URL) |
| `-ClearCache` | Clear the Office add-in web cache `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Word must be closed) |
| `-NoUpdater` | Do not install the auto-updater (logon check + `myscripts:` protocol) |

### Auto-update

The installer also sets up an updater. This is needed because code changes go live automatically, but **manifest changes** (new buttons, labels, icons) only reach a machine when the manifest is reinstalled.

| Piece | Location |
| --- | --- |
| Updater script | `%LOCALAPPDATA%\MY_SCRIPTS\updater.ps1` (from `installer/updater.ps1`, baked with the prod URL and ID, saved as UTF-8 with BOM) |
| Logon check | `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\MY_SCRIPTS Updater` → `updater.ps1 -AtLogon` |
| Protocol handler | `HKCU\Software\Classes\myscripts` → `updater.ps1 "%1"` (opened as `myscripts:update`) |

Flow:

1. The updater compares the installed manifest `<Version>` (`%LOCALAPPDATA%\MY_SCRIPTS\manifests\<id>.xml`) with `<baseUrl>version.json`.
2. If a newer version exists, it asks **Yes/No**. If you answer No, nothing changes. At logon it says nothing when you are up to date or offline.
3. If you answer Yes, it waits until every Word window is closed (OK/Cancel prompt).
4. It downloads the published `install.ps1` and runs it with `-ClearCache`, which re-registers the manifest and refreshes the updater.
5. It confirms when the update is done.

Inside Word:

- The manifest passes its version to the FunctionFile as `commands.html?mv=<Version>`. `src/runtime/update.js` compares that value with `version.json` once per session.
- When a newer version exists, every result dialog shows a **Cập nhật** banner and stays open. The button calls `Office.context.ui.openBrowserWindow("myscripts:update")`, and Windows asks to open the updater.
- **Thông tin → Kiểm tra cập nhật** runs the same check on demand.

To install without the updater, use `install.ps1 -NoUpdater`. To check manually, run `powershell -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\MY_SCRIPTS\updater.ps1"`.

### Uninstalling

Close Word first, then use any one of the options below. Each option removes the per-user registry value and the copied manifest. Restart Word afterwards and the **MY_SCRIPTS** tab is gone.

| Option | Command |
| --- | --- |
| One-liner (prod add-in) | `irm https://tnmquann.github.io/word-my-scripts/uninstall.ps1 \| iex` |
| Portable package | double-click `Uninstall.cmd` in the extracted release folder |
| From the repository, prod | `powershell -ExecutionPolicy Bypass -File installer\uninstall.ps1` |
| From the repository, dev build | `npm run unload:dev` (same as `uninstall.ps1 -Dev`) |
| Everything this installer registered (prod + dev) | `uninstall.ps1 -All` |
| A specific add-in ID | `uninstall.ps1 -Id <guid>` |

What gets removed:

- `HKCU\Software\Microsoft\Office\16.0\WEF\Developer\<add-in GUID>`
- `%LOCALAPPDATA%\MY_SCRIPTS\manifests\<add-in GUID>.xml`
- When no MY_SCRIPTS manifest remains: the auto-updater (the `MY_SCRIPTS Updater` Run value, the `HKCU\Software\Classes\myscripts` protocol and the `%LOCALAPPDATA%\MY_SCRIPTS` folder)

Manual removal, if the scripts cannot run: delete the registry value above in `regedit`, delete the `%LOCALAPPDATA%\MY_SCRIPTS` folder, and restart Word. To also discard cached web files, delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` while Word is closed. This clears the cache of all web add-ins, which re-download on next use.

Uninstalling does not touch your documents, and the GitHub Pages site stays online.

---

## 4. Development loop

```powershell
npm install
npm run certs          # once: trust the localhost dev certificate (accept the Windows prompt)
npm run dev            # esbuild watch + HTTPS server on https://localhost:3000
npm run sideload:dev   # once: register dist-dev/manifest.xml, then restart Word
```

- Edit files under `src/`. esbuild rebuilds `dist-dev/` when you save. To load the new code, close the result dialog and click the button again. You may need to restart Word because the FunctionFile runtime is cached.
- Debugging: open **MY_SCRIPTS (DEV)**, then attach the **Microsoft Edge DevTools Preview** / WebView2 DevTools to the `commands.html` runtime. Output from `console.warn` and `console.error` (fallback paths, count mismatches) appears there.
- When you are done: `npm run unload:dev`.

Recommended test procedure: keep a set of copies of real exam files, and never test on originals. For each command, run it twice and check that the second run reports no changes, because the commands are designed to be idempotent. Check the checklist in each spec.

---

## 5. Releasing

1. Bump the version: `npm version patch|minor|major --no-git-tag-version`, and add a `## [X.Y.Z] - YYYY-MM-DD` section to [CHANGELOG.md](../CHANGELOG.md). The manifest `<Version>` is taken from `package.json` and gets a trailing `.0`.
2. Commit, then push to `main`. Pages redeploys automatically.
3. Optional: publish a portable package.

   ```powershell
   git tag v1.1.0; git push origin v1.1.0
   ```

   A tag with a suffix (for example `v1.1.0-beta.1`) is published as a **pre-release**. The release notes come from the matching `## [1.1.0]` section of `CHANGELOG.md`, and the workflow fails if that section is missing. Publish every release under a new tag instead of overwriting an existing one.
   `.github/workflows/release.yml` builds `MY_SCRIPTS-v1.1.0.zip` on `windows-latest` and attaches it to a GitHub Release. To build the package locally, run `npm run release`.
4. Users need to re-run the installer only if the manifest changed.

---

## 6. Adding a new command

Commands are organized into **feature groups**. Each group maps to one ribbon group on the MY_SCRIPTS tab and lives in `src/features/<group>/`. Its specs live in `docs/specs/<group>/`.

```
src/features/<group>/
├── index.js         # registry: { <FunctionName>: { label, run } }
└── <command>.js     # one module per command
```

**Adding a command to an existing group**

1. Create `src/features/<group>/<name>.js` exporting `async function <name>(context)`. It should:
   - Receive a `Word.RequestContext`.
   - Return a short summary string, which is shown in the dialog.
   - Throw an exception on hard failure.
2. Register the command in `src/features/<group>/index.js`: `<name>: { label: "<dialog title>", run: <name> }`.
3. Add an `<Item>` to the group's menu in `manifest/manifest.template.xml`, together with the matching `bt:Image`, `ShortStrings` and `LongStrings` resources. `FunctionName` must equal the key from step 2. Keys must be unique across all groups.

**Adding a new group**

1. Create `src/features/<group>/index.js` exporting the registry, and add it to `GROUPS` in `src/features/index.js`.
2. In the manifest, add a new `<Group>` inside `<CustomTab>`, containing a `<Control xsi:type="Menu">` (or `Button` controls), plus its resources.
3. Create `docs/specs/<group>/` and list the group in `README.md`.

**For every new command or group**, also:
- Add an icon design to `scripts/make-icons.mjs` and run `npm run icons`.
- Write `docs/specs/<group>/<Name>.md`.
- Run `npm run build && npm run validate`, then test with the dev add-in.

Guidelines for command code:
- Batch the work. Queue `load` calls and writes for all paragraphs, then `context.sync()` once per phase. Never call `sync` inside a per-paragraph loop unless you need to.
- When deleting by index, delete from the end of the document towards the start.
- Use `*OrNullObject` getters instead of try/catch for optional objects.
- Keep commands idempotent. A second run should be a no-op.

---

## 7. Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| Tab does not appear | First check **Home → Add-ins → Advanced → DEVELOPER ADD-INS**. If MY_SCRIPTS is listed there, click **Add** (a one-time activation needed on recent Word builds). Otherwise, Word was open during the install. Close all `WINWORD.exe` processes and reopen Word. Check that `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` contains the GUID. |
| Tab appears but buttons do nothing or show an error | The Pages site is unreachable (offline, or Pages not deployed yet). Open `<baseUrl>commands.html` in a browser. |
| Old behaviour after a deploy | Office has cached the HTML. Close Word, then run `install.ps1 -ClearCache`. |
| Manifest changes are not picked up | Re-run the installer, which copies the new manifest, then restart Word. If that doesn't help, also use `-ClearCache`. |
| `npm run dev` says that the certificate is missing | Run `npm run certs` and accept the prompt. |
| Dialog does not show | Another add-in dialog is already open (error 12007). The command still ran. |
| Error `GeneralException` with `errorLocation` | The document contains structures the command did not expect, for example nested tables or content controls. Send the dialog text along with a sample document. |

Undo: every `context.sync()` is recorded in Word's undo stack, so **Ctrl+Z** works, but one command may need several undo steps. Keep a copy of the file before you run any destructive command, such as *Xoá ĐACT ĐGNL* or *Gạch chân đáp án Ngữ Văn*.
