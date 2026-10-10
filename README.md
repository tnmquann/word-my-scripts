# MY_SCRIPTS — Word add-in (Office JS)

MY_SCRIPTS is a personal collection of Word automation commands, delivered as an Office JS add-in. It adds a **MY_SCRIPTS** ribbon tab. Each **feature group** on that tab bundles related commands, and new groups are added over time. Every command runs as an Office JS ribbon command (`ExecuteFunction`), with no task pane and no VBA.

## Feature groups

| Group (ribbon) | Scope | Code | Specs |
| --- | --- | --- | --- |
| **ĐGNL** | Clean-up tools for exam documents of the Vietnam National University Ho Chi Minh City Academic Competency Test (V-ACT). *ĐGNL* is the Vietnamese abbreviation for this test. | [src/features/dgnl/](src/features/dgnl/) | [docs/specs/dgnl/](docs/specs/dgnl/) |
| **Thông tin** | *Tác giả* button: opens a task pane with author information from the GitHub profile, the add-in version and the list of tools. | [src/ui/about/](src/ui/about/) | — |

```
MY_SCRIPTS (tab)
└── ĐGNL (group)
    └── ĐGNL ▾ (menu button)
        ├── Chuẩn hoá Toán ĐGNL        → chuanHoaToan
        ├── Fix lỗi tab câu hỏi        → fixTabNumbering
        ├── Gạch chân đáp án Ngữ Văn   → gachDapAnNguVan
        ├── Fix lỗi break dòng T.Anh   → removeBreakEng
        └── Xoá ĐACT ĐGNL              → xoaDactDgnl
└── Thông tin (group)
    └── Tác giả (button)               → task pane: author card (GitHub profile), version, tool list
```

Ribbon labels and result messages are in Vietnamese. To add a command or a whole new group, see [docs/SETUP.md § 6](docs/SETUP.md#6-adding-a-new-command).

Each command runs inside a single `Word.run` batch. When it finishes, a small dialog shows the result. On success the dialog closes by itself after 3.5 s. On failure it stays open until you close it.

## Quick start

| Task | Command |
| --- | --- |
| Install on a new PC (no Node needed) | `irm https://tnmquann.github.io/word-my-scripts/install.ps1 \| iex` |
| Or install from the release zip | unzip `MY_SCRIPTS-vX.Y.Z.zip` → double-click `Install.cmd` |
| Uninstall (no Node needed) | `irm https://tnmquann.github.io/word-my-scripts/uninstall.ps1 \| iex`, or `Uninstall.cmd` from the release zip; then restart Word |
| Production build | `npm run build` → `dist/` |
| Local dev server (HTTPS, watch) | `npm run certs` (once), `npm run dev`, `npm run sideload:dev` |
| Portable installer package | `npm run release` → `release/` |
| Validate manifest | `npm run validate` |

For the full setup, release and troubleshooting guide, see **[docs/SETUP.md](docs/SETUP.md)**.

## Repository layout

```
.
├── addin.config.json          # add-in IDs, display name, prod/dev base URLs, author profile
├── manifest/
│   └── manifest.template.xml  # ribbon definition (tab → groups → menus → items)
├── src/
│   ├── runtime/               # FunctionFile (commands.html), command registration, runCommand wrapper
│   ├── ui/                    # result dialog + about page (manifest SourceLocation)
│   └── features/<group>/      # one folder per ribbon group: index.js registry + one module per command
├── assets/icons/              # generated PNG icons (16/32/80 px)
├── installer/                 # install.ps1 / uninstall.ps1 / *.cmd wrappers
├── scripts/                   # build.mjs (esbuild), make-icons.mjs, release.mjs
├── docs/
│   ├── SETUP.md               # setup, deployment, operations guide
│   └── specs/<group>/         # technical spec for each command, per group
└── .github/workflows/         # Pages deploy (push to main) + release (tag v*)
```

Generated folders, which git ignores: `dist/` (prod build), `dist-dev/` (dev build), `release/` (installer packages).

## Command specs

### ĐGNL (V-ACT)

- [Chuan_Hoa_Toan.md](docs/specs/dgnl/Chuan_Hoa_Toan.md): normalizes the lead-in line of shared-stimulus question groups (Math).
- [Fix_Tab_Numbering.md](docs/specs/dgnl/Fix_Tab_Numbering.md): merges question lists, applies hanging indent, renumbers questions from 1.
- [Gach_DapAn_NguVan.md](docs/specs/dgnl/Gach_DapAn_NguVan.md): marks the correct option and removes redundant answer lines (Literature).
- [Remove_Break_Eng.md](docs/specs/dgnl/Remove_Break_Eng.md): removes extra *Keep with next* flags that cause large page gaps (English).
- [Xoa_DACT_DGNL.md](docs/specs/dgnl/Xoa_DACT_DGNL.md): deletes the detailed-solution blocks.

## Requirements

- Word for Windows (Microsoft 365, or Word 2019 and later). The manifest requires `WordApi 1.3`.
- Internet access at runtime, because the add-in's web assets are served from GitHub Pages.
- Development only: Node.js 20 or later and npm.
