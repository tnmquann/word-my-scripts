# Changelog

All notable changes to MY_SCRIPTS are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/). The manifest `<Version>` is the package version followed by `.0`.

Releases are published as GitHub pre-releases, tagged `vX.Y.Z-beta.N`. The release workflow uses the matching `## [X.Y.Z]` section of this file as the release notes.

## [2.2.0-beta.2] - 2026-10-17

No manifest change: applies on the next click / next opening of the pane.

### Added
- **Công cụ BigKey → Plugin BigKey → Chuyển lời giải xuống cuối** (plugin TACHLOIGIAI_BIGKEY): moves every solution to the end of the document under `LỜI GIẢI CHI TIẾT`; questions without a solution get a red `THIẾU LỜI GIẢI Ở ĐÂY` line and are listed.

## [2.2.0] - 2026-10-17

New ribbon item, so the manifest changes: **re-install once** (the **Cập nhật** button or the install command), then restart Word.

### Added
- **CTPT 2018 → Canh giữa Table**: centres every table in the document between the margins (the table itself, not the text in its cells). Nested tables are centred in their cell; a table wider than the margins is fitted to the page (Word desktop).

## [2.1.0-beta.2] - 2026-10-17

No manifest change: the commands apply on the next click and the pane on its next opening.

### Added
- **Công cụ BigKey → Plugin BigKey** (from `D:\BIGKEY\Apps\Bigkey\Plugins`): **Tách LG**, **Xoá dòng viết LG**, **Chemistry Formatter**, **Tô phiếu ZG50** (ZipGrade 50, optional mã đề A–E; pick the template `ans50zg.docx` once).
- Check lỗi đề reports questions with two underlined answers (*thừa đáp án*), the remaining check of BigKey's *Quét lỗi*.

### Fixed
- Chuẩn hoá BigKey: **the real cause of "the last question needs a second run"**: the empty-line clean-up treated every paragraph as containing a table (the OOXML package includes the styles part, which always has `<w:tblPr>`), so leftover empty paragraphs stayed between options. Now only the paragraph body is checked; one run finishes every question (verified end to end on both test exams).
- Xoá lời giải / Ẩn-Hiện lời giải / Xoá dòng viết LG: a block also ends at a `HẾT` line and at a paragraph indented less than the solution (a passage at the margin), as Xoá ĐACT ĐGNL does.
- Tô phiếu ZG50 with a mã đề no longer fails (the VBA used an undeclared variable).

## [2.1.0] - 2026-10-17

The manifest changes (new pane item, WordApi 1.4), so **re-install once** (the **Cập nhật** button or the install command), then restart Word. The companion template also changes: the installer replaces it (close Word first).

### Added
- **Công cụ BigKey…** (Chuẩn hoá BigKey ▾): a task pane with the old BigKey *Chuẩn hóa* sub-menus (2020 + 2024 versions):
  - Chuẩn hoá: Chuẩn hoá màu PA-ĐA-LG, Chèn / Xoá "Câu 0:".
  - ID: Gắn ID mức độ - nhóm (NB/TH/VD/VDC, nhóm 1–15; jumps to the next question), Ẩn/Hiện ID, Tô màu ID, Xoá ID, Xoá nguồn tác giả.
  - Phương án: Xuống dòng các phương án, Chấm cuối / Xoá chấm cuối phương án, Chuyển đáp án thành phương án.
  - Đáp án: Tạo bảng đáp án (10 columns, "MÃ ĐỀ: 001"), Ghi đáp án vào lời giải, Gạch chân đáp án từ bảng đáp án / từ lời giải / highlight đỏ.
  - Lời giải: Chèn lời giải, Ẩn/Hiện lời giải, Xoá lời giải. Số thứ tự: Đánh số thứ tự (start number field), Chuyển STT Auto sang Text. Hình ảnh: Picture Inline WithText, Căn giữa hình ảnh.
- **MathType Color** (MY_SCRIPTS+ tab, needs MathType): Xanh, Đỏ, Tím, Xanh lá, Hồng, Đen (remove colour), Trắng — on the selected equations, or all of them.

### Changed
- Minimum requirement: WordApi 1.4 (Office 2024 or later; Microsoft 365 recommended).
- Answer tools edit option blocks in place (layout, tab stops and content kept). Unlike the VBA, *Chuyển đáp án thành phương án* no longer removes every underline in the document, and *Gạch chân đáp án highlight đỏ* only looks inside options.

### Fixed
- Chuẩn hoá BigKey: the last question (or any question split across two chunks) sometimes needed a second run. Chunks are now cut only at question starts, the document's last paragraph is no longer replaced through OOXML, empty leftovers are removed right after each OOXML step, and option detection skips empty paragraphs.

### Not ported
- *Gắn ID6* and *Gắn Nguồn tác giả* need BigKey's Bank data (`%APPDATA%\BigKey`). *Quét lỗi* is covered by *Check lỗi đề*.

## [2.0.0-beta.5] - 2026-10-16

Code-only fix for **Chuẩn hoá BigKey**. It applies on the next click.

### Fixed
- Options laid out on one line (or two) had no tab stops, so B/C/D landed on Word's default 1.27 cm tabs (reported on `test_eng.docx`). When `insertOoxml` replaces paragraphs, Word merges the last inserted paragraph into the old paragraph mark and keeps the old paragraph formatting. Each inserted package now ends with an empty paragraph that absorbs the old mark (it is removed in the clean-up step), so the fixed tab stops apply: 4 per line at 5.25 / 9.5 / 13.75 cm, 2 per line at 9.5 cm, or 1 per line. The same fix applies to the layout reset (the last paragraph of every chunk kept its old formatting).

## [2.0.0-beta.4] - 2026-10-16

Code-only hotfix for **Chuẩn hoá BigKey**. It applies on the next click.

### Fixed
- `GeneralException` at `Document.defaultTabStop` (reported on `test_eng.docx`): setting the default tab stop needs WordApiDesktop 1.4, which some Word builds reject. It is now optional (own step, skipped on error), like turning off line numbering. If a beta.3 run stopped with this error, undo it (Ctrl+Z) or reopen the file, then run the command again.

## [2.0.0-beta.3] - 2026-10-16

Code-only patch for **Chuẩn hoá BigKey**. It applies on the next click, with no re-install. (The start-number popup needs the Word restart from beta.2.)

### Fixed
- The layout now matches the old BigKey (`bkBonDong`), checked side by side on a real exam (same 7 pages, same line breaks):
  - every paragraph outside tables loses its direct paragraph formatting: alignment, 1.15/1.08 line spacing, custom tab stops and *Keep with next* are reset; left indent 1 cm, single spacing, no space before/after;
  - every text is Times New Roman 12, including option labels and tables (equations keep Cambria Math). Bold, italic, underline and colours are kept;
  - the tab after `Câu N.` becomes a space; manual page breaks are removed; default tab stop 1.5 cm, gutter 0, no line numbering (Word desktop).
- Option paragraphs no longer inherit the document's Normal style (font, line spacing).
- `Chọn X` is no longer justified (the old BigKey keeps it left-aligned).

### Changed
- Numbering starts from 1 again, like the old BigKey. When an exam starts at another number (e.g. Câu 31), a popup asks for the start number.

## [2.0.0-beta.2] - 2026-10-16

Patch with no manifest change, so no re-install is needed. The new popup field is part of the runtime: **close Word and open it again** once to get it. Until then, the command numbers from 1 as before.

### Changed
- **Fix lỗi tab câu hỏi** asks which number to start from before it changes anything. The popup shows how many questions it found and the current first number; the default is 1. **Huỷ** leaves the document unchanged.
- Popups can now contain a number field (`askUser` with `input`). Enter confirms.

## [2.0.0] - 2026-10-16

**v2 beta**: first batch of tools migrated from the old VBA templates (`BigKeyUI.dotm`, `NHOM-WORD-BIEN-SOAN.dotm`). The manifest changes (two new groups and a task pane), so **re-install once** (the **Cập nhật** button or the install command), then restart Word. The remaining tools of `TonNgocMinhQuanProj.dotm` follow in later betas.

### Added
- Group **Chuẩn hoá BigKey** (from `BigKeyUI.dotm` → *Chuẩn hóa*):
  - **Chuẩn hoá BigKey (cả đề / vùng chọn)**: numbering to text, Times New Roman 12, A–D options laid out 4/2/1 per line by estimated width (MathType objects and pictures kept), correct option red and underlined, `Câu i:` bold blue, *Lời giải* bold blue, `Chọn X` highlighted, empty lines and double spaces removed. Page setup (A4, 1.5 cm) on Word desktop only.
  - **Check lỗi đề**: lists every problem at once (duplicate or skipped numbers, missing/duplicate/out-of-order options, missing answer, underline ≠ `Chọn X`, stray pictures) and selects the first faulty question. Text before the first question is a warning only.
- Group **Xoá cần thiết** (from `NHOM-WORD-BIEN-SOAN.dotm`): **Xoá trích nguồn**, **Xoá lời giải**, **Xoá tác giả**, and **Xoá mức độ, khung…**, a task pane that replaces the old *Xóa mức độ* form: delete questions by level 1–4 (then renumber), delete ID tags and *Phân tích hướng dẫn giải*, replace solutions with n dotted lines, and add or remove question frames (8 border styles, 2 colours, 3 fills; Word desktop only).
- **MY_SCRIPTS+** tab (companion template `MY_SCRIPTS_Desktop.dotm`, Word desktop, needs MathType): **MathType hoá**, **Lệch MathType**, **Lỗi Equation**, **Lệch Pic-Math**. The installer copies it into the Word Startup folder; the uninstaller removes it.
- macOS installer and uninstaller (`install.sh`, `uninstall.sh`), published on the site and in the release zip. **Not tested on a Mac yet.**
- Tests: option layout, real-exam checks from a read-only Word export (`tests/tools/export-docx.ps1`), Xoá helpers, and a companion smoke test against MathType.

### Changed
- Compared with the VBA versions:
  - Chuẩn hoá BigKey **no longer converts tables to text** (the VBA crashed on exams with tables). Tables are left as they are.
  - Question numbers continue from the first question's number instead of restarting at 1 (exams such as `Câu 31–60`).
  - Xoá lời giải / Xoá mức độ stop at group lead-ins (`Questions 46-52: …`, `Đọc đoạn văn…`), so reading passages are kept.
  - MathType hoá uses MathType's Toggle TeX instead of `SendKeys`, and no longer removes the spaces inside a formula.
  - Lệch MathType and Lệch Pic-Math no longer reset the formatting of the surrounding text; Lệch Pic-Math adjusts each picture by its own height.
- The equations round-trip test runs in jsdom instead of Edge headless.

## [1.4.6] - 2026-10-15

This release adds a new ribbon group, so the manifest changes: **re-install once** (the **Cập nhật** button or the install command), then restart Word.

### Added
- New group **Equations** with **LaTeX → Equation** and **Equation → LaTeX**. Each command asks for the scope in a MathType-style dialog: **Vùng chọn** (selection), **Toàn bộ tài liệu** (whole document) or **Huỷ** (cancel).
  - Supports `$…$`, `$$…$$` and `\[…\]`: fractions, sums and integrals, roots, matrices, `cases`, accents, `\mathbb`/`\mathcal`/`\mathbf`, `\text`, and more. `$5 and $6` is not treated as a formula.
  - `\frown` / `\smile` were added, for arc notation `\overset{\frown}{AB}`.
- `npm test` with smoke tests kept in the repository (`tests/`): runtime and hot reload, the equations batch driver, and the LaTeX round trip.

### Changed
- Compared with the original script, conversion is batched. LaTeX → Equation now uses 4 Word round-trips in total, where it used to need about 3 per formula. Equation → LaTeX (whole document) reads the document once and rewrites only the paragraphs that contain equations. If a batch fails, the command falls back to the original one-by-one method.

## [1.4.5-beta.4] - 2026-10-12

Code-only patch for **Tạo bảng đáp án (CTPT 2018)**. It applies on the next click, with no need to close Word.

### Changed
- A summary answer key created at the cursor now starts with a bold, centered **`BẢNG ĐÁP ÁN`** title, followed by each `Phần X.` heading and its table. The title is added only when the whole key is created. When only some parts are missing, their tables are appended without a second title.

## [1.4.5-beta.3] - 2026-10-12

Code-only patch for **Tạo bảng đáp án (CTPT 2018)**. It applies on the next click, with **no need to close Word**.

### Changed
- Answer tables pre-built inside each part (`PHẦN X` → `BẢNG ĐÁP ÁN` → table) are now **skipped**: they are neither filled nor modified. The command writes a **summary answer key for all parts at the cursor** instead, and the result dialog lists the skipped parts.

## [1.4.5-beta.2] - 2026-10-12

Code-only patch for **Tạo bảng đáp án (CTPT 2018)**. It applies on the next click and needs no reinstall.

### Fixed
- Exams that place each part's answer table directly under the part heading (`PHẦN X` → `BẢNG ĐÁP ÁN` → table) reported "Không tìm thấy câu hỏi nào". The first `BẢNG ĐÁP ÁN` line was treated as the end of the questions. It now ends the question area only when no question follows it, and a table right under `BẢNG ĐÁP ÁN` inside a part belongs to that part.
- Part headings produced by automatic numbering (`PHẦN I.` as a list number) were not recognized.
- A question no longer extends past the next part heading or `BẢNG ĐÁP ÁN` line.

### Added
- Combined label + answer cells: `1.A` (multiple choice / short answer) and true/false columns headed `Câu N` with cells `a) Đ`. Answers are written in the table's own style (`Đ`/`S` or `Đúng`/`Sai`).

### Verified
- On a real CTPT 2018 Math mock exam, the extracted answers matched the existing answer tables 100% (12 + 16 + 6 answers), and the simulated fill reproduced tables I and II exactly.

## [1.4.5] - 2026-10-12

This release changes the manifest (new group, button and tooltips) and the runtime. To get it, **re-install once**: use the **Cập nhật** button or the install command, then restart Word.

### Added
- New ribbon group **CTPT 2018** with **Tạo bảng đáp án (CTPT 2018)**. The command reads the answers of every part (multiple choice, true/false, short answer, essay) and fills the answer-key tables. Missing tables are created **at the cursor**, with no page break. Before writing, it checks the question numbering: a duplicate number stops the command, and a gap or out-of-order number asks **Tiếp tục / Dừng lại**. Equation answers are copied into cells (JSZip is bundled; no CDN needed).
- Part headings accept any valid Roman numeral or integer (`PHẦN IX`, `PHẦN XVI`, `PHẦN 11`, …). New table headings follow the document's numbering style.
- Hover tooltips on every button: full name, the exam format it targets, what it does, and what the document needs.
- Runtime: commands can ask the user mid-run through an Office dialog with custom buttons (replaces `confirm`/`alert`, which do not work in ribbon commands).

## [1.4.4-beta.3] - 2026-10-10

### Fixed
- Code updates did not apply until Word was fully restarted, because Word keeps the add-in's command runtime in memory for the whole session.

### Changed
- The commands are now a separate bundle (`features-<hash>.js`). On every button click, the runtime checks `version.json` (uncached, about 3 s timeout). If a newer bundle is deployed, it loads it before running the command. **From this version on, patches apply on the next click, without restarting Word.**
- This one update still needs a single full Word restart (close every Word window), because it replaces the runtime itself.

## [1.4.4-beta.2] - 2026-10-10

Small patch on top of 1.4.4. The add-in version stays 1.4.4, because this is a code-only change and installed machines pick it up automatically.

### Fixed
- **Chuẩn hoá Toán ĐGNL**: after rewriting a lead-in, the paragraph below was pulled up onto the same line (for example `…từ 53 - 56Một nhóm bạn…`). `Paragraph.insertText(…, "Replace")` also replaced the paragraph mark. Only the paragraph content is replaced now, and the paragraph mark is kept.

### Changed
- **Chuẩn hoá Toán ĐGNL**: the canonical lead-in is now `Dựa vào thông tin dưới đây để trả lời các câu hỏi từ X đến Y`. Lines in the previous form (`…từ câu X đến câu Y`) are still recognized and are rewritten to the new form.

## [1.4.4] - 2026-10-10

### Changed
- **Chuẩn hoá Toán ĐGNL**: the indent pass now also resets the **whole stimulus below each lead-in** to left indent 0 and first-line indent 0. A stimulus is every paragraph up to the next auto-numbered `Câu N.` question or the next lead-in.

### Fixed
- Lead-ins worded as `…trả lời các câu hỏi từ 72 đến 74` were not recognized, so they were neither normalized nor un-indented. The range pattern now accepts `(các) câu (hỏi) (từ) X –/-/đến (câu) Y`.

## [1.4.3] - 2026-10-10

### Changed
- **Chuẩn hoá Toán ĐGNL**: new indent pass after normalization. Every lead-in in the canonical form (`Dựa vào thông tin dưới đây để trả lời từ câu X đến câu Y`) gets left indent 0 and first-line indent 0. This includes lines whose text already matched, which were therefore not replaced. The result dialog reports how many lines were re-indented.

### Fixed
- Lead-ins that were rewritten with `insertText("Replace")` could keep their old indent, because the indent was applied in the same batch as the rewrite.

## [1.4.2] - 2026-10-10

### Fixed
- New ribbon icons did not appear after updating. Icon URLs in the manifest now carry `?v=<version>`, so Office reloads them on every version change.

## [1.4.1] - 2026-10-10

### Fixed
- The **Cập nhật** button closed the popup without doing anything. It now opens an update page in the browser, which launches the updater and shows a fallback command.
- Updater prompts could be hidden behind Word. They are now always on top.

### Added
- Updater log file at `%LOCALAPPDATA%\MY_SCRIPTS\updater.log`.

## [1.4.0] - 2026-10-10

### Changed
- New minimal line icons and app logo (single colour, no text glyphs).
- Result popup: smaller (about 22% × 24% of the screen) with a tighter layout. Content is sent via `messageChild`, so the title bar no longer shows a long URL.

### Fixed
- Author information no longer appears when Word opens or restores the pane without the **Tác giả** button.

## [1.3.0] - 2026-10-10

### Changed
- Result popups never close on their own.
- Popup redesign: Aptos/Arial font, one line per result, and technical details collapsed under *Chi tiết kỹ thuật*. Known errors are shown as short Vietnamese messages.
- The default add-in page is now a minimal welcome page. Author information moved to its own page, opened only by **Thông tin → Tác giả**.

## [1.2.3] - 2026-10-10

### Documentation
- Installer, updater and docs explain the one-time activation on recent Word builds: Home → Add-ins → Advanced → DEVELOPER ADD-INS → MY_SCRIPTS → **Add**.

## [1.2.2] - 2026-10-10

### Fixed
- `irm …/install.ps1 | iex` failed with *Unexpected token '-ManifestPath'*. The installer scripts are now published without a BOM.

## [1.2.1] - 2026-10-10

### Added
- GNU GPL v3 (`GPL-3.0-or-later`) license.

## [1.2.0] - 2026-10-10

### Added
- Auto-update. The updater checks for a new version at Windows logon and on demand, asks before installing, waits for Word to close, then reinstalls.
- **Thông tin → Kiểm tra cập nhật** button. Result popups show an update banner when a newer version exists.

## [1.1.0] - 2026-10-10

### Added
- **Thông tin → Tác giả** button: a task pane with author information from the GitHub profile.

## [1.0.0] - 2026-10-10

### Added
- MY_SCRIPTS ribbon tab with the **ĐGNL** group: Chuẩn hoá Toán ĐGNL, Fix lỗi tab câu hỏi, Gạch chân đáp án Ngữ Văn, Fix lỗi break dòng T.Anh, Xoá ĐACT ĐGNL.
- Per-user installer and uninstaller, GitHub Pages hosting, and a portable release zip.
