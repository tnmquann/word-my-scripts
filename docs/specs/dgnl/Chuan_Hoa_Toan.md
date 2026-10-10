# Chuẩn hoá Toán ĐGNL — normalize shared-stimulus lead-ins (Math)

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → ĐGNL → **Chuẩn hoá Toán ĐGNL** |
| Action | `chuanHoaToan` |
| Module | `src/features/dgnl/chuanHoaToan.js` |
| API sets | WordApi 1.3 (`listItemOrNullObject`, `parentTableOrNullObject`, `inlinePictures`) |
| Destructive | No. It rewrites text and formatting only. |
| Idempotent | Yes |

## Purpose

In the Math section of Vietnam National University Ho Chi Minh City Academic Competency Test (V-ACT) exams, a group of questions shares one stimulus (text, figure or table). The group starts with a lead-in line whose wording varies, for example `Dựa vào thông tin sau trả lời câu hỏi 73-74`. This command does three things:

1. Rewrites every lead-in to the canonical form

   `Dựa vào thông tin dưới đây để trả lời từ câu X đến câu Y`

   and makes it bold.
2. Removes the indentation from the lead-in and from the stimulus paragraphs that follow it.
3. Centers the pictures and tables inside the stimulus.

## Detection

A paragraph is a **lead-in** when all of the following hold:

| Rule | Pattern |
| --- | --- |
| Starts with "Dựa vào" or "Trả lời" (case-insensitive) | `/^\s*(Dựa vào\|Trả lời)/i` |
| Is not a question line (a list item whose `listString` starts with `Câu`) | `listItemOrNullObject.listString` |
| Length ≤ 120 characters, which excludes solution sentences such as "Dựa vào BBT suy ra…" | `text.length` |
| Contains a question range | `/câu(?:\s+hỏi)?\s*(?:từ\s*)?(\d+)\s*(?:[-–]\|đến)\s*(?:câu\s*)?(\d+)/i`. This matches `câu hỏi 73-74`, `câu 73 – 74`, `các câu hỏi từ 72 đến 74` and the canonical `từ câu 73 đến câu 74`. |

The **stimulus scope** runs from the lead-in up to, but not including, the next question line (`Câu N.`) or the next lead-in.

## Operations

For each lead-in `i` and every paragraph `j` in its scope:

| Target | Change |
| --- | --- |
| Lead-in text | `insertText(canonical, "Replace")`, but only when the text differs |
| Lead-in font | `bold = true` |
| `j` | `leftIndent = 0`, `firstLineIndent = 0` |
| `j` containing ≥ 1 inline picture | `alignment = "Centered"` |
| `j` inside a table | parent `table.alignment = "Centered"` |

**Indent pass (since 1.4.3; stimulus included since 1.4.4).** After the writes above are synced, the command reloads all paragraphs and their list items. For **every** paragraph that matches the canonical form `/^\s*Dựa vào thông tin dưới đây để trả lời từ câu \d+ đến câu \d+/i`, it sets `leftIndent = 0` and `firstLineIndent = 0` on that lead-in and on **every paragraph below it**. The pass stops at the next auto-numbered question (a list item whose `listString` starts with `Câu`) or at the next canonical lead-in. This covers lines that were just rewritten and lines that already matched (so were not replaced). Running it as a separate phase guarantees that the indent applies to the paragraph *after* `insertText("Replace")` has rewritten it.

## Execution profile

The command needs 6 round-trips in total:

1. Load all paragraph texts.
2. Load the list item, inline pictures and parent table of every paragraph.
3. Commit the rewrites and the formatting.
4. Reload the paragraph texts (indent pass).
5. Load the list item of every paragraph, to detect `Câu N.` boundaries.
6. Commit the indent pass.

## Output

`Đã xử lý N dòng dẫn (viết lại M dòng).` followed by `Đưa K đoạn (dòng dẫn + dữ kiện) về lề trái 0.` If no lead-in is found, it reports that instead.

## Limitations

- Floating (non-inline) pictures and shapes are not centered.
- Question lines must be real Word list items with a `Câu …` list string. Questions numbered as typed text are not recognized as scope boundaries.
- `insertText(…, "Replace")` applies the formatting of the first character to the whole line. Mixed formatting inside a lead-in is therefore lost (it is made bold anyway).

## Test checklist

- [ ] `…câu hỏi 73-74`, `…câu 73 – 74` and `…câu 73 đến câu 74` are all rewritten to the canonical form.
- [ ] A solution line such as `Dựa vào BBT suy ra…` is left untouched.
- [ ] A figure or table in the stimulus is centered, and the question lines keep their hanging indent.
- [ ] A second run reports `viết lại 0 dòng`.
- [ ] `…trả lời các câu hỏi từ 72 đến 74` is rewritten to `…trả lời từ câu 72 đến câu 74`. The stimulus below it (for example `Cho hàm số y = …`) moves to left indent 0, and the following auto-numbered `Câu 72.` keeps its hanging indent.
- [ ] Every canonical lead-in ends up at left indent 0 and first-line indent 0. This includes a line that already had the canonical text but was indented, which is fixed even though its text is not replaced.
