# Changelog

All notable changes to MY_SCRIPTS are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/). The manifest `<Version>` is the package version followed by `.0`.

Releases are published as GitHub pre-releases, tagged `vX.Y.Z-beta.N`. The release workflow uses the matching `## [X.Y.Z]` section of this file as the release notes.

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
