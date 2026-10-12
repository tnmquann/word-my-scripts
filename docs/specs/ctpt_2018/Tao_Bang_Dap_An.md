# Tạo bảng đáp án (CTPT 2018) — build and fill the answer-key tables

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → CTPT 2018 → **Tạo bảng đáp án (CTPT 2018)** |
| Action | `taoBangDapAn` |
| Module | `src/features/ctpt2018/taoBangDapAn.js` (port of the user-tested Script Lab "answer-key filler v7") |
| API sets | WordApi 1.3 (`Table.values`, `insertTable`, `compareLocationWith`, `getParagraphBeforeOrNullObject`, `parentTableOrNullObject`) and WordApi 1.1 (`getOoxml`, `getSelection`, `Body.insertFileFromBase64` for equation answers) |
| Bundled dependency | JSZip 3.10.1 (rebuilds a small `.docx` so that MathType/OMML answers can be copied into a table cell) |
| Destructive | It writes into answer tables and creates missing tables. Question content is never modified. |
| Idempotent | Yes. Cells are cleared before they are filled, so re-running refreshes the tables. |

## Purpose

Exams that follow the 2018 General Education Curriculum (*Chương trình GDPT 2018*, CTPT 2018) are split into numbered parts with fixed answer formats. The command reads every question, extracts the correct answer from the document itself, and writes it into one answer-key table per part. Tables that do not exist yet are created **at the user's cursor**.

| Part type (`kind`) | Heading keywords | Answer format in the table | Where the answer is read from (priority order) |
| --- | --- | --- | --- |
| `MC`: multiple choice | *nhiều phương án / nhiều lựa chọn / trắc nghiệm nhiều* | one letter `A`–`D` per question | 1. the option letter that is **marked** (underline, highlight or cell shading) · 2. `Chọn X` / `Đáp án: X` in the solution |
| `TF`: true/false | *đúng sai / đúng hay sai* | `Đúng` / `Sai` per statement `a)`–`d)` (label `1a`, `1b`, …) | 1. `a) Đúng/Sai` lines in the solution · 2. marked statements in the stem (marked = `Đúng`) |
| `SA`: short answer | *trả lời ngắn* | free text or an equation | the line containing `Đáp án:` / `Đáp số:` / `Trả lời:` after the solution marker |
| `ESSAY` | *tự luận* | same as `SA` | same as `SA`. Questions may be tagged `[TL]` instead of `Câu N.` |

## Document model

**Questions.** A question starts at a paragraph that matches any of the following:
- typed `Câu N.` or `Câu N:`;
- an auto-numbered list item whose list string starts with `Câu`;
- a paragraph starting with the essay tag `[TL]`.

A question ends where the next question starts, or at the end marker. The *stem* runs up to the solution marker `Lời giải` / `Hướng dẫn giải`.

**End of the question area.** This is the first paragraph that matches `HẾT` or `BẢNG ĐÁP ÁN`. If there is none, it is the end of the document.

**Parts.** A part starts at a heading that begins with `PHẦN <n>` and is not itself a question.
- `<n>` is either a **valid Roman numeral** (`I`, `IV`, `IX`, `XVI`, …) or a **positive integer** (`1`, `2`, `11`, …), in the range 1–50.
- The numeral must be followed by the end of the line, whitespace or punctuation. So `PHẦN Mở đầu`, `Phần cuối` and `Phần IIII` are **not** part headings.
- Part numbers are normalized to integers for matching. A table labelled `Phần 2` therefore matches the part headed `PHẦN II`.
- The document's numbering style is taken from the first part heading (Roman by default). Every label the command writes, such as headings of new tables and messages, uses that same style. If the document mixes both styles, a warning is shown.
- Consecutive questions without a heading are grouped by detected format and given the next unused part number.

**The part type** comes from the heading keywords. If the heading has no keywords, it is the majority format of the part's questions. A question whose detected format differs from its part type produces a warning.

## Processing pipeline

1. **Scan** all paragraphs and their list strings (2 round-trips). Build parts and questions.
2. **Validate the question numbering** inside each part, **before anything is written**:

   | Finding | Severity | Behaviour |
   | --- | --- | --- |
   | Two or more questions with the same `Câu N` in one part | **Error** | Stops. Nothing is written. The result dialog lists the part, the duplicated number, the positions and the first 40 characters of each question. |
   | Numbers not strictly consecutive (gaps or out of order, e.g. `1 → 4 → 5 → 2`) | Warning | Shows a dialog with **Tiếp tục** / **Dừng lại**. Closing the dialog counts as *Dừng lại*. If you continue, the table uses the numbers as written. New tables are sorted in ascending order. |
   | Some questions in a part have no readable number | Warning | The table is labelled by order of appearance. |
   | Essay part without `Câu N.` numbering (`[TL]` only) | none | The check is skipped for that part. |

   The numbering does not have to start at 1, because some exams continue the numbering across parts.
3. **Extract the answers.** The command fetches the OOXML of the relevant paragraphs in one batch and reads the marks from the run properties (`w:u`, `w:highlight`, `w:shd`).
4. **Locate the existing answer tables.** These are tables after `BẢNG ĐÁP ÁN` (or after the end marker), whose preceding paragraph is a `Phần <n>` heading.
   - Tables located before the first question are ignored, with a warning.
   - Unlabelled tables are assigned to the remaining parts in order.
   - Table labels are matched against the question numbers, or against the order 1…n, whichever matches more cells.
   - Both horizontal layouts (label row, then answer row) and vertical layouts (label column, then answer column) are supported.
5. **Create the missing tables at the cursor.** This needs no page break and no `BẢNG ĐÁP ÁN` heading.
   - **The cursor is on an empty line:** the tables are inserted at that line. The empty line is kept as spacing after them.
   - **The cursor is inside text:** the tables are inserted right after that paragraph, followed by one empty spacer line.
   - Every new table gets a `Phần <n>.` heading in the document's numbering style. The table is styled *Table Grid* and centered, with labels in bold: 6 columns per row, or 8 for true/false. The font of the heading above an existing answer table is reused.
   - **The cursor is inside a table:** the command stops before writing anything and asks you to move the cursor.
6. **Fill the tables.** Each matched cell is cleared, then filled.
   - Equation answers are copied as a mini `.docx` with `insertFileFromBase64`.
   - Cell text is not bold (`bold: false`) and is centered.
7. **Fit the column widths** of the answer tables only. The width is between 40 and 80 pt per column, depending on the longest value, and the table is at most 430 pt wide.

## Configuration (`DEFAULTS` at the top of the module)

| Option | Default | Meaning |
| --- | --- | --- |
| `answerLabel` | `/(Đáp án\|Đáp số\|Trả lời)\s*:/i` | marker of the answer line for short-answer and essay questions |
| `solutionMarker` | `/^(Lời giải\|Hướng dẫn giải)\s*:?\s*$/i` | marks the end of the stem |
| `endMarker` | `/^([\s(\-–—*]*HẾT\b\|BẢNG ĐÁP ÁN)/` | end of the question area |
| `tableAnchor` | `/^BẢNG ĐÁP ÁN/i` | existing answer tables are searched after this line |
| `tlTag` | `/^\[TL\]/` | essay question tag |
| `createMissingTables` | `true` | create tables for parts that have none (at the cursor) |
| `fitTables` | `true` | adjust column widths of the answer tables |
| `minColWidth` / `maxColWidth` / `maxTableWidth` | `40` / `80` / `430` pt | width limits |
| `bold` / `alignment` | `false` / `"Centered"` | formatting of filled cells |
| `popup` | `true` | ask before continuing when the numbering is not consecutive |
| `dryRun` | `false` | analyse only and return the report without writing (developer use) |

## Output

The result dialog shows:

- one line per part: `Phần I (trắc nghiệm nhiều lựa chọn): 12 câu — điền 12/12 ô`;
- `Tạo mới bảng tại con trỏ: Phần …`, when tables were created;
- `⚠ N cảnh báo:` with the first three warnings. Expand **Chi tiết kỹ thuật** to see every warning, the extracted answers per part and the table match statistics.

The command stops without writing anything, and says so, when:
- the numbering has duplicates (error dialog);
- the user chooses **Dừng lại**;
- the cursor is inside a table while tables need to be created.

## Requirements and limitations

- Questions must be numbered `Câu N.` (typed or auto-numbered), or tagged `[TL]` for essays.
- The correct multiple-choice option must be marked (underline, highlight or shading) or stated as `Chọn X`. True/false answers need `a) Đúng/Sai` lines or marked statements.
- **Place the cursor** where new tables should go (preferably on an empty line) **before** running the command. If the cursor is placed above the first question, those tables will be ignored on the next run (see step 4).
- Equation answers need `insertFileFromBase64` (Word desktop, recent builds).
- Filling is cell by cell (one sync per cell), so large tables take a few seconds.

## Test checklist (run on a copy)

- [ ] An exam with PHẦN I / II / III and no tables: all tables are created at the cursor in part order, headed `Phần I.`, `Phần II.`, `Phần III.`, with no page break.
- [ ] The same exam numbered `PHẦN 1 / 2 / 3`: the new headings read `Phần 1.`, `Phần 2.`, `Phần 3.`.
- [ ] Headings `PHẦN IX`, `PHẦN XVI` and `PHẦN 11` are recognized. `PHẦN Mở đầu` is not treated as a part.
- [ ] An existing table labelled `Phần 2` is filled for the part headed `PHẦN II`.
- [ ] Duplicate `Câu 2` in one part: an error dialog appears and the document is unchanged.
- [ ] Numbering `1 → 2 → 4`: a warning dialog appears. **Dừng lại** leaves the document unchanged, and **Tiếp tục** fills a table labelled 1, 2, 4.
- [ ] Cursor inside a table: the command stops with the message asking to move the cursor.
- [ ] A short-answer value written as an equation appears as an equation in the cell.
- [ ] A second run refills the same cells without duplicating tables.
