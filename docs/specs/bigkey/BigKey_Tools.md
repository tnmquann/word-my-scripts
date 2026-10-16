# Chuẩn hoá BigKey — Công cụ BigKey (task pane)

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → Chuẩn hoá BigKey ▾ → **Công cụ BigKey…** (opens the task pane on the right) |
| Modules | `src/features/bigkey/tools.js` (Office JS commands), `optionTools.js` (pure OOXML edits of option blocks), task pane `src/ui/bigkey/bigkey.html` + `bigkey.js` |
| Origin | The sub-menus of *Chuẩn hóa* in `BigKeyUI.dotm`: the 2020 version (Word STARTUP folder) and the 2024 version (`D:\BIGKEY\Apps\Bigkey\Plugins\BigkeyUI.dotm`), merged |
| API sets | WordApi 1.4 (minimum since 2.1.0). *Picture Inline WithText* needs WordApiDesktop 1.2 (Word desktop). |
| Destructive | The Xoá commands delete content; the others rewrite formatting. Ctrl+Z restores. |

Office JS has no nested menus, so the old sub-menus are collapsible sections of one task pane (maintainer's suggestion, 2026-10-17): Chuẩn hoá, ID, Phương án, Đáp án, Lời giải, Số thứ tự, Hình ảnh. The ribbon keeps one ▾ per group. The pane calls `tools.js` directly and shows the result of the last click at the bottom.

## How option blocks are edited

The answer tools work on each question's option block (`planBlocks`: from the paragraph starting with `A.` to the one containing `D.`, up to 4 non-empty paragraphs). `optionTools.js` edits the block's OOXML **in place**: paragraphs, tab stops and content runs are kept, only the labels or the described characters change. Each package ends with an empty sentinel paragraph (see `normalizeXml.sentinelParagraph`); the leftover empty paragraph right after a block is deleted unless it holds a picture, object or equation.

The VBA did the same with Find/Replace and temporary marker characters (`#`, `~`), which broke documents that already contained them.

## Commands

| Section | Command | Behaviour | Change from the VBA |
| --- | --- | --- | --- |
| Chuẩn hoá | Chuẩn hoá màu PA-ĐA-LG | Labels bold blue, the underlined answer letter red + underlined, `A)`/`A:` → `A.`, *Lời giải* bold blue | Layout unchanged; only option blocks are touched |
| Số thứ tự | Đánh số thứ tự | Start number from the field next to the button (default 1), `Câu i:` bold blue, renumbers the BigKey key table too | Typed numbers only; automatic numbering is left to Word |
| | Chuyển STT Auto sang Text | Automatic numbering → typed text (tables included) | — |
| | Chèn / Xoá "Câu 0:" | Inserts `Câu 0: ` at the start of the paragraph at the cursor / deletes every `Câu 0:` | — |
| ID | Gắn ID mức độ - nhóm | Buttons NB/TH/VD/VDC (`[1]`–`[4]`) and groups 1–15; writes `[k]` after `Câu N:` of the question at the cursor (replaces an existing `[k]`), then selects the next question | Replaces the form `Bkf_GanIDNhom` |
| | Ẩn/Hiện ID | Hides every visible ID tag; when none is visible, un-hides all hidden text | Same toggle; state read from the document instead of `%APPDATA%\BigKey\showhide.dat` |
| | Tô màu ID | ID tags bold pink | Only real IDs: `[x…]` without spaces or `;` (not `[0;1]`) |
| | Xoá ID | Deletes ID tags, collapses double spaces | The VBA deleted only pink tags |
| | Xoá nguồn tác giả | Deletes `(…)` groups coloured green (BigKey *Gắn Nguồn*) | — |
| Phương án | Xuống dòng các phương án | Each option on its own paragraph (split before labels preceded by a tab/line break; new paragraphs copy the paragraph format without tab stops) | Formatting kept |
| | Chấm cuối / Xoá chấm cuối phương án | Adds a `.` at the end of every option (also after an equation or object), never `..` / removes one trailing `.` | — |
| | Chuyển đáp án thành phương án | Removes the answer mark: every label plain bold blue | The VBA removed **all** underline in the document |
| Đáp án | Tạo bảng đáp án | 10-column key at the end: merged green header `MÃ ĐỀ: 001`, cells `n.X`; an existing key is overwritten. Answers from the underlined label, else `Chọn X`; missing → `?` + warning | — |
| | Ghi đáp án vào lời giải | Questions without *Lời giải*: inserts `Lời giải` + `Chọn X`; questions with one: fixes or adds `Chọn X` | Structure-based, no page/line positions |
| | Gạch chân đáp án từ bảng đáp án / từ lời giải | Marks the answer from the key table / from `Chọn X` | Other labels lose a stale underline |
| | Gạch chân đáp án highlight đỏ | Highlighted text inside an option → red, no highlight; that option becomes the answer | Only inside option blocks (the VBA turned every highlight red, e.g. `Chọn X`) |
| Lời giải | Chèn lời giải | Like *Ghi đáp án*, without touching existing solutions; essay questions get `Lời giải` + `Nội dung lời giải` | — |
| | Ẩn/Hiện lời giải | Hides every solution block (stops at group lead-ins); when none is visible, un-hides all hidden text | — |
| | Xoá lời giải | Same command as *Xoá cần thiết → Xoá lời giải* | — |
| Plugin BigKey | Tách LG | Copies the solutions of the **selected** questions to the end of the document under a centred bold `LỜI GIẢI CHI TIẾT`: each becomes `Câu N.` (bold blue) + the content after *Lời giải* (tables and pictures kept; stems, options, the *Lời giải* line and passages dropped). The exam itself is unchanged. | One OOXML copy instead of clipboard + a temporary document; no section break carried over |
| | Chuyển lời giải xuống cuối (TACHLOIGIAI_BIGKEY) | Moves every solution (whole document) to the end under `LỜI GIẢI CHI TIẾT`, each as `Câu N.` + content; questions without a solution get a red `THIẾU LỜI GIẢI Ở ĐÂY` line in order and are listed in the result. Uses the same block boundaries as Xoá lời giải. | The VBA added a fake "Câu cuối" and searched `Câu n` by running number (wrong when numbers are not consecutive) |
| | Xoá dòng viết LG | Per question, deletes from *Lời giải* (if any), else from the paragraph after option D (multiple choice) or from the first dotted line (essay), to the end of the question. Stops at the next question, group lead-ins, `HẾT` lines and less-indented passages. | The VBA inserted fake *Lời giải* lines and deleted with marker characters |
| | Chemistry Formatter | On the selection: digits after a letter/`)`/`]` subscript (H₂SO₄, Al₂(SO₄)₃); charges superscript with their 1–2 digits (Fe³⁺, SO₄²⁻; NO₃⁻ keeps 3 subscript; H-Cl untouched); `-->` → →, `*` → ·; `6.02E23` → 6.02 × 10²³. Equations untouched. | Same rules (hdpnChemistryFormatter), applied in OOXML; shapes/SmartArt text not handled |
| | Tô phiếu ZG50 | Reads each question's answer (underlined label, else `Chọn X`) for questions 1–50, fills the ZipGrade 50 template (bubble filled and moved to the A/B/C/D column; optional mã đề A–E), opens it in a new window to save or print as PDF. The default template `assets/templates/ans50zg.docx` (owned by the maintainer) ships with the add-in (`templates/ans50zg.docx` on Pages); the pane also lets the user pick another template of the same structure, kept in local storage. | The VBA crashed when a mã đề was given (undeclared `arMade`) and saved a PDF next to the exam |
| Hình ảnh | Picture Inline WithText | Floating pictures → in line with text (Word desktop) | — |
| | Căn giữa hình ảnh | Centres paragraphs that contain only a picture | The VBA also centred text lines that contained a formula picture |

**Not ported** (need BigKey's Bank data in `%APPDATA%\BigKey`, which this machine does not have; maintainer decision 2026-10-16): *Gắn ID6*, *Gắn Nguồn tác giả*. *Quét lỗi* (2024) is covered by *Check lỗi đề*, which also reports two underlined answers (*thừa đáp án*). Plugin *Trộn đề theo mức độ* is not ported (maintainer decision 2026-10-17). *MathType Color* is in the companion template (MY_SCRIPTS+). The other BigKey groups (Bank, Trộn đề, T-H-F, Chấm bài, Plugins) are outside the *Chuẩn hóa* group.

## Test checklist

Run on a **copy** of a real exam.

- [ ] Chuẩn hoá màu PA-ĐA-LG on an exam with `A)` labels: labels become `A.` bold blue, the underlined answer red; layout unchanged.
- [ ] Xuống dòng: one-line options become 4 paragraphs; formatting kept. Chấm cuối then Xoá chấm cuối returns the original text.
- [ ] Chuyển đáp án thành phương án: no label stays underlined/red; underlined words in the stem stay.
- [ ] Tạo bảng đáp án: key at the end with a merged green header; run again → same table updated, not a second table. Missing answers are listed.
- [ ] Gạch chân từ bảng đáp án (after editing a cell) and từ lời giải (`Chọn C`) move the answer mark.
- [ ] Gạch chân highlight đỏ on an exam marked by highlight; `Chọn X` lines stay highlighted.
- [ ] Ghi đáp án vào lời giải / Chèn lời giải on an exam with some solutions missing; essay questions get the placeholder line.
- [ ] Ẩn/Hiện ID and Ẩn/Hiện lời giải: click twice → back to the original.
- [ ] Tô màu ID / Xoá ID do not touch `[0;1]` intervals.
- [ ] Gắn ID mức độ - nhóm: click VD → `[3]` after `Câu N:` and the next question is selected; clicking NB on the same question replaces `[3]`.
- [ ] Đánh số thứ tự from 31 also renumbers the BigKey key table.
- [ ] Picture Inline / Căn giữa hình ảnh on an exam with a floating picture.
- [ ] Tách LG on 3 selected questions: `LỜI GIẢI CHI TIẾT` + `Câu N.` + solution at the end; the exam is unchanged.
- [ ] Xoá dòng viết LG on an exam with dotted writing lines and a `HẾT` line: the lines go, `HẾT` and passages stay.
- [ ] Chemistry Formatter on `H2SO4, Fe3+, SO42-, NO3-, CuSO4*5H2O, 2H2 + O2 --> 2H2O, 6.02E23`.
- [ ] Tô phiếu ZG50: pick `ans50zg.docx` once; the filled sheet opens in a new window with the right bubbles and mã đề.

Automated: `tests/bigkey-optiontools.mjs [export.json…]` (every transform on synthetic and real option blocks), `tests/bigkey-registry.mjs` (ribbon commands, manifest FunctionNames, every pane button maps to `tools.js`, ID/key-table helpers, planBlocks with empty paragraphs), `tests/bigkey-pane.mjs` (pane renders and runs every button).
