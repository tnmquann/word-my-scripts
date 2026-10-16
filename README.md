# MY_SCRIPTS — Word add-in (Office JS)

MY_SCRIPTS is a personal collection of Word automation commands, delivered as an Office JS add-in. It adds a **MY_SCRIPTS** ribbon tab. Each **feature group** on that tab bundles related commands, and new groups are added over time. Commands run as Office JS ribbon commands (`ExecuteFunction`); tools that had a VBA form use a task pane instead. The MathType tools, which Office JS cannot drive, ship as a small companion VBA template with its own **MY_SCRIPTS+** tab (Word desktop only).

## Feature groups

| Group (ribbon) | Scope | Code | Specs |
| --- | --- | --- | --- |
| **ĐGNL** | Clean-up tools for exam documents of the Vietnam National University Ho Chi Minh City Academic Competency Test (V-ACT). *ĐGNL* is the Vietnamese abbreviation for this test. | [src/features/dgnl/](src/features/dgnl/) | [docs/specs/dgnl/](docs/specs/dgnl/) |
| **CTPT 2018** | Tools for exams that follow the 2018 General Education Curriculum (*Chương trình GDPT 2018*): parts for multiple choice, true/false, short answer and essay. | [src/features/ctpt2018/](src/features/ctpt2018/) | [docs/specs/ctpt_2018/](docs/specs/ctpt_2018/) |
| **Equations** | Convert between LaTeX (`$…$`, `$$…$$`, `\[…\]`) and native Word equations (OMML), in either the selection or the whole document. | [src/features/equations/](src/features/equations/) | [docs/specs/equations/](docs/specs/equations/) |
| **Chuẩn hoá BigKey** | Formats multiple-choice exams in the BigKey house style (options laid out 1/2/4 per line by width, numbering, solutions) and checks an exam for structural errors. Tables are left unchanged. A task pane, *Công cụ BigKey*, holds the old sub-menus: answer colours, ID tags, options, answer key, solutions, numbering, pictures. Ported from `BigKeyUI.dotm` (2020 + 2024). | [src/features/bigkey/](src/features/bigkey/) | [docs/specs/bigkey/](docs/specs/bigkey/) |
| **Xoá cần thiết** | Deletes sources, solutions and author lines; a task pane deletes questions by level and ID tags, replaces solutions with dotted lines, and frames questions. Ported from `NHOM-WORD-BIEN-SOAN.dotm`. | [src/features/xoa/](src/features/xoa/), [src/ui/xoa/](src/ui/xoa/) | [docs/specs/xoa/](docs/specs/xoa/) |
| **MY_SCRIPTS+** (VBA, desktop) | *Công cụ MathType*: LaTeX → MathType via Toggle TeX, MathType and picture baseline fixes, Cambria Math for equations. Needs MathType. | [companion/](companion/) | [docs/specs/companion/](docs/specs/companion/) |
| **Thông tin** | *Tác giả*: task pane with author information from the GitHub profile, the version and the tool list. *Kiểm tra cập nhật*: checks for a newer version and launches the updater. | [src/ui/author/](src/ui/author/), [src/runtime/update.js](src/runtime/update.js) | [docs/SETUP.md, Auto-update](docs/SETUP.md#auto-update) |

```
MY_SCRIPTS (tab)
└── ĐGNL (group)
    └── ĐGNL ▾ (menu button)
        ├── Chuẩn hoá Toán ĐGNL        → chuanHoaToan
        ├── Fix lỗi tab câu hỏi        → fixTabNumbering
        ├── Gạch chân đáp án Ngữ Văn   → gachDapAnNguVan
        ├── Fix lỗi break dòng T.Anh   → removeBreakEng
        └── Xoá ĐACT ĐGNL              → xoaDactDgnl
├── CTPT 2018 (group)
│   └── CTPT 2018 ▾ (menu button)
│       ├── Tạo bảng đáp án (CTPT 2018) → taoBangDapAn
│       └── Canh giữa Table             → canhGiuaTable
├── Equations (group)
│   └── Equations ▾ (menu button)       each asks: Vùng chọn / Toàn bộ tài liệu / Huỷ
│       ├── LaTeX → Equation            → latexToEquation
│       └── Equation → LaTeX            → equationToLatex
├── Chuẩn hoá BigKey (group)
│   └── Chuẩn hoá BigKey ▾ (menu button)
│       ├── Chuẩn hoá BigKey (cả đề)     → chuanHoaBigKeyDe
│       ├── Chuẩn hoá BigKey (vùng chọn) → chuanHoaBigKeyBoiDen
│       ├── Check lỗi đề                 → checkLoiDe
│       └── Công cụ BigKey…              → task pane bigkey.html (ID, phương án, đáp án, lời giải, STT, hình ảnh)
│       └── Check lỗi đề                 → checkLoiDe
├── Xoá cần thiết (group)
│   └── Xoá cần thiết ▾ (menu button)
│       ├── Xoá trích nguồn              → xoaTrichNguon
│       ├── Xoá lời giải                 → xoaLoiGiai
│       ├── Xoá tác giả                  → xoaTacGia
│       └── Xoá mức độ, khung…           → task pane xoa.html (levels, ID, dotted lines, frames)
└── Thông tin (group)
    ├── Tác giả (button)               → task pane: author card (GitHub profile), version, tool list
    └── Kiểm tra cập nhật (button)     → checkUpdate: compare installed vs published version, offer update

MY_SCRIPTS+ (tab, companion template MY_SCRIPTS_Desktop.dotm, Word desktop)
└── Công cụ MathType: MathType hoá · Lệch MathType · Lỗi Equation · Lệch Pic-Math
```

Ribbon labels and result messages are in Vietnamese. To add a command or a whole new group, see [docs/SETUP.md § 6](docs/SETUP.md#6-adding-a-new-command).

Each command runs inside a single `Word.run` batch. When it finishes, a small dialog shows the result. The dialog never closes on its own: you close it when you are done reading. It shows a short summary, with technical error details collapsed under *Chi tiết kỹ thuật*, plus a **Cập nhật** banner when a newer version exists. Text uses Aptos, with Arial as the fallback.

## Quick start

| Task | Command |
| --- | --- |
| Install on a new PC (no Node needed) | `irm https://tnmquann.github.io/word-my-scripts/install.ps1 \| iex` |
| Install on macOS (**unverified**) | `curl -fsSL https://tnmquann.github.io/word-my-scripts/install.sh \| bash` |
| Or install from the release zip | unzip `MY_SCRIPTS-vX.Y.Z.zip` → double-click `Install.cmd` |
| Uninstall (no Node needed) | `irm https://tnmquann.github.io/word-my-scripts/uninstall.ps1 \| iex`, or `Uninstall.cmd` from the release zip; then restart Word |
| Production build | `npm run build` → `dist/` |
| Local dev server (HTTPS, watch) | `npm run certs` (once), `npm run dev`, `npm run sideload:dev` |
| Portable installer package | `npm run release` → `release/` |
| Validate manifest | `npm run validate` |
| Rebuild the companion template (Windows + Word) | `npm run companion` |

For the full setup, release and troubleshooting guide, see **[docs/SETUP.md](docs/SETUP.md)**. For the release history, see **[CHANGELOG.md](CHANGELOG.md)**.

## Repository layout

```
.
├── addin.config.json          # add-in IDs, display name, prod/dev base URLs, author profile
├── manifest/
│   └── manifest.template.xml  # ribbon definition (tab → groups → menus → items)
├── src/
│   ├── runtime/               # FunctionFile (commands.html), command registration, runCommand wrapper, update check
│   ├── ui/                    # result dialog, start page (SourceLocation), author pane
│   └── features/<group>/      # one folder per ribbon group: index.js registry + one module per command
├── assets/icons/              # generated PNG icons (16/32/80 px)
├── companion/                 # VBA source + ribbon XML of MY_SCRIPTS_Desktop.dotm (MathType tools)
├── installer/                 # install.ps1 / uninstall.ps1 / updater.ps1 / *.cmd, install.sh / uninstall.sh (macOS), MY_SCRIPTS_Desktop.dotm
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

### CTPT 2018

- [Canh_Giua_Table.md](docs/specs/ctpt_2018/Canh_Giua_Table.md): centres every table between the margins (whole table, not the cell text); too-wide tables are fitted.
- [Tao_Bang_Dap_An.md](docs/specs/ctpt_2018/Tao_Bang_Dap_An.md): extracts the answers of every part (multiple choice, true/false, short answer, essay), validates the question numbering, and fills the answer-key tables, creating missing tables at the cursor.

### Equations

- [LaTeX_Equations.md](docs/specs/equations/LaTeX_Equations.md): LaTeX ⇄ Word equation conversion, with a scope dialog, batched Office JS calls and a sequential fallback.

### Chuẩn hoá BigKey

- [Chuan_Hoa_BigKey.md](docs/specs/bigkey/Chuan_Hoa_BigKey.md): BigKey formatting (option layout by estimated width, numbering, solutions; tables untouched) and *Check lỗi đề*.
- [BigKey_Tools.md](docs/specs/bigkey/BigKey_Tools.md): the *Công cụ BigKey* task pane (the old *Chuẩn hóa* sub-menus).

### Xoá cần thiết

- [Xoa_Can_Thiet.md](docs/specs/xoa/Xoa_Can_Thiet.md): deleting sources, solutions, authors, questions by level and ID tags; dotted lines and question frames (task pane).

### MY_SCRIPTS+ (companion template)

- [MathType_Tools.md](docs/specs/companion/MathType_Tools.md): MathType tools in VBA, build and install of `MY_SCRIPTS_Desktop.dotm`.

## Tests

`npm test` builds the add-in and runs the smoke tests in [tests/](tests/):
- `runtime-smoke.cjs`: command registration, hot reload of the features bundle, and the `askUser` dialog;
- `equations-batch.mjs`: the batched LaTeX → Equation driver against a mock Word API;
- `equations-roundtrip.mjs`: LaTeX → OMML → LaTeX (jsdom);
- `bigkey-options.mjs`: option layout on synthetic blocks;
- `bigkey-normalize.mjs`, `bigkey-optiontools.mjs`, `bigkey-registry.mjs`, `bigkey-pane.mjs`: BigKey layout reset, option-block edits, command/pane wiring;
- `bigkey-real.mjs`: Chuẩn hoá BigKey and Check lỗi đề on a real exam exported by `tests/tools/export-docx.ps1` (skipped without an export);
- `xoa-helpers.mjs`: block finding and tag helpers of Xoá cần thiết.

Windows-only harnesses in `tests/tools/`: `export-docx.ps1` (read-only export of a real document through Word COM) and `companion-smoke.ps1` (the companion template against MathType).

Hovering over any ribbon button shows its full name, the exam format it targets, what it does, and what the document needs for it to work.

## Requirements

- Word for Windows or Mac (Microsoft 365, or Office 2024 and later). The manifest requires `WordApi 1.4`: Office 2024 or later, Microsoft 365 recommended. Word on the web runs most commands; page setup and question frames need Word desktop and show a warning online. The macOS installer has not been tested yet.
- MY_SCRIPTS+ (MathType tools): Word desktop with macros, and MathType.
- Internet access at runtime, because the add-in's web assets are served from GitHub Pages.
- Development only: Node.js 20 or later and npm.

## License

Copyright (C) 2026 TON NGOC MINH QUAN.

This program is free software: you can redistribute it and/or modify it under the terms of the **GNU General Public License version 3**, or (at your option) any later version, as published by the Free Software Foundation. It is distributed in the hope that it will be useful, but **without any warranty**. See [LICENSE](LICENSE) for the full text. The SPDX identifier is `GPL-3.0-or-later`.
