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
 │                            │              │ about.html, assets/*.png    │
 │ HKCU\…\WEF\Developer       │              │ manifest.xml, install.ps1   │
 │   {add-in id} = manifest   │              └─────────────────────────────┘
 └────────────────────────────┘
```

- **Manifest** (`manifest/manifest.template.xml`): an add-in-only XML manifest that defines a `CustomTab` → `Group` → `Menu` with five `ExecuteFunction` items. `scripts/build.mjs` fills in the placeholders from `addin.config.json` and `package.json`.
- **FunctionFile** (`commands.html`): a page without UI. Word loads it when a ribbon button is clicked. `src/commands/commands.js` maps each `<FunctionName>` to a handler with `Office.actions.associate`.
- **runCommand** (`src/core/runCommand.js`): wraps every handler. It does four things:
  - Runs the handler inside `Word.run` and measures how long it takes.
  - Turns `OfficeExtension.Error` into readable text (message, code, `debugInfo.errorLocation`).
  - Opens `dialog.html` with `displayDialogAsync`, passing the result in the query string.
  - Always calls `event.completed()`.

  A module-level `busy` flag drops re-entrant clicks while a command is still running.
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

`install.ps1` takes these switches:

| Parameter | Meaning |
| --- | --- |
| `-ManifestPath <file>` | Register a specific manifest file |
| `-ManifestUrl <url>` | Download the manifest from this URL (defaults to the prod Pages URL) |
| `-ClearCache` | Clear the Office add-in web cache `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Word must be closed) |

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

1. Bump the version: `npm version patch|minor|major --no-git-tag-version`. The manifest `<Version>` is taken from `package.json` and gets a trailing `.0`.
2. Commit, then push to `main`. Pages redeploys automatically.
3. Optional: publish a portable package.

   ```powershell
   git tag v1.1.0; git push origin v1.1.0
   ```

   A tag with a suffix (for example `v1.1.0-beta.1`) is published as a **pre-release**.
   `.github/workflows/release.yml` builds `MY_SCRIPTS-v1.1.0.zip` on `windows-latest` and attaches it to a GitHub Release. To build the package locally, run `npm run release`.
4. Users need to re-run the installer only if the manifest changed.

---

## 6. Adding a new command

1. Create `src/functions/<group>/<name>.js` exporting `async function <name>(context)`. It should:
   - Receive a `Word.RequestContext`.
   - Return a short summary string, which is shown in the dialog.
   - Throw an exception on hard failure.
2. Register the function in `src/commands/commands.js`: `<name>: makeCommand("<ribbon label>", <name>)`.
3. Add the ribbon entry in `manifest/manifest.template.xml`. For a new menu item, add an `<Item>`. For a new menu, add a new `<Group>` + `<Control xsi:type="Menu">`. Then add the matching `bt:Image`, `ShortStrings` and `LongStrings` resources. `FunctionName` must equal the key you registered in step 2.
4. Add an icon design to `scripts/make-icons.mjs` and run `npm run icons`.
5. Write `docs/specs/<group>/<Name>.md`.
6. Run `npm run build && npm run validate`, then test with the dev add-in.

Guidelines for command code:
- Batch the work. Queue `load` calls and writes for all paragraphs, then `context.sync()` once per phase. Never call `sync` inside a per-paragraph loop unless you need to.
- When deleting by index, delete from the end of the document towards the start.
- Use `*OrNullObject` getters instead of try/catch for optional objects.
- Keep commands idempotent. A second run should be a no-op.

---

## 7. Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| Tab does not appear | Word was open during the install. Close all `WINWORD.exe` processes and reopen Word. Check that `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` contains the GUID. |
| Tab appears but buttons do nothing or show an error | The Pages site is unreachable (offline, or Pages not deployed yet). Open `<baseUrl>commands.html` in a browser. |
| Old behaviour after a deploy | Office has cached the HTML. Close Word, then run `install.ps1 -ClearCache`. |
| Manifest changes are not picked up | Re-run the installer, which copies the new manifest, then restart Word. If that doesn't help, also use `-ClearCache`. |
| `npm run dev` says that the certificate is missing | Run `npm run certs` and accept the prompt. |
| Dialog does not show | Another add-in dialog is already open (error 12007). The command still ran. |
| Error `GeneralException` with `errorLocation` | The document contains structures the command did not expect, for example nested tables or content controls. Send the dialog text along with a sample document. |

Undo: every `context.sync()` is recorded in Word's undo stack, so **Ctrl+Z** works, but one command may need several undo steps. Keep a copy of the file before you run any destructive command, such as *Xoá ĐACT ĐGNL* or *Gạch chân đáp án Ngữ Văn*.
