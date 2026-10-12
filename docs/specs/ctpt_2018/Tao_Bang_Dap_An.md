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

**Auto numbering.** Both `Câu N.` and `PHẦN X.` may be typed or produced by Word's automatic numbering. The list string (`ListItem.listString`) is read and prepended to the paragraph text before matching. Paragraphs inside tables are never treated as questions or part headings.

**End of the question area.** The area ends at the first `HẾT` line. It also ends at a `BẢNG ĐÁP ÁN` line that has **no question after it**, which is the end-of-document answer-key layout. A `BẢNG ĐÁP ÁN` line that appears *between* questions does not end the area. That is the per-part layout described below.

**Question boundaries.** A question ends at the next question, the next part heading, or the next `BẢNG ĐÁP ÁN` line, whichever comes first. A question therefore never extends into the next part or its answer table.

**Answer-table layouts.** The command recognizes two layouts. Only the end-of-document (summary) tables are filled:

| Layout | Shape | How the table is matched to a part |
| --- | --- | --- |
| Per part (inline) | `PHẦN X …` → `BẢNG ĐÁP ÁN` → *table* → questions of part X | **Skipped** (since 1.4.5-beta.3, at the maintainer's request). These pre-built tables are neither filled nor modified. The answers go into the summary answer key instead, which is created at the cursor when it does not exist yet. The `BẢNG ĐÁP ÁN` line still acts as a question boundary. |
| End of document | questions … → `BẢNG ĐÁP ÁN` / `HẾT` → `Phần X.` → *table* … | the `Phần X` heading right above each table; unlabelled tables go to the remaining parts in order |

**Cell styles.** Each table is analysed in two ways, and the one that matches more labels is used:

| Style | Example | Written as |
| --- | --- | --- |
| Separate label and answer cells | label row `1 2 3 …` with an answer row below, or a label column with answers to the right | the answer alone in the answer cell |
| Combined: number + answer (MC, SA) | `1.A`, `2.`, `Câu 3: 98` | `1.` + answer → `1.A` (the existing separator is kept) |
| Combined: true/false by column | header `Câu 1`, cells below `a) Đ`, `b) S` (or to the right) | `a) ` + `Đ`/`S`. Abbreviated if the table already uses `Đ`/`S` or is empty, and `Đúng`/`Sai` if the table already uses the full words. |

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
4. **Locate the existing answer tables** (see *Answer-table layouts* above).
   - **Per part:** a table right under a `BẢNG ĐÁP ÁN` line inside part X is **skipped**, and listed in the result as *Bỏ qua bảng đáp án dựng sẵn trong từng phần*.
   - **End of document:** a table after the final `BẢNG ĐÁP ÁN` line or after `HẾT` belongs to the part whose `Phần <n>` heading sits right above it. Unlabelled tables in that area are assigned to the remaining parts in order.
   - A table preceded by a `Phần <n>` heading but located before the first question is ignored, with a warning.
   - Labels are matched against the question numbers, or against the order 1…n, whichever matches more cells. Each table is read with both cell styles (separate or combined), and the style with more matches is used.
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

## Validation on a real exam (2026-10-12)

The command was dry-run on `test.docx`, a CTPT 2018 Math mock exam. The run used a read-only COM export of the document fed into the module, and simulated the table writes.

- Detected parts: I (12 multiple-choice questions), II (4 true/false), III (6 short answer). All questions are auto-numbered, and every answer table sits directly under the heading of its part.
- The extracted answers matched the answer tables already in the document **100%**: 12/12 multiple-choice letters (all read from underlined options), 16/16 true/false statements (read from the solutions), and 6/6 short answers.
- 1.4.5-beta.3, simulated run: the three per-part tables were left **unchanged**. A summary answer key was created at the cursor:
  - `Phần I.` 6 × 2 grid with A A B D A B / C D C A D C;
  - `Phần II.` 8 labels per row, `1a … 4d` with Đúng/Sai;
  - `Phần III.` 1025, 56,2, 98, 6, and the MathType equations for questions 5 and 6, because the solution's `Đáp số:` line is an equation.

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
- [ ] Per-part layout (`PHẦN X` → `BẢNG ĐÁP ÁN` → table): the per-part tables stay unchanged, the questions are detected after the first `BẢNG ĐÁP ÁN`, and a summary answer key for all parts is created at the cursor.
- [ ] Running the command again with the cursor elsewhere refills the summary key created earlier (found by its `Phần X.` headings, as long as it is below the first question) instead of creating a second one.
- [ ] Combined cells: `1.A` cells are rewritten as `<n>.<answer>`, and `a) Đ` columns are rewritten with `Đ`/`S`.
- [ ] `Câu N.` and `PHẦN X.` produced by automatic numbering are recognized exactly like typed ones.
