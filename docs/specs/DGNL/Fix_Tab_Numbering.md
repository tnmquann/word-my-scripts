# Fix lỗi tab câu hỏi — unify question numbering and hanging indent

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → ĐGNL → **Fix lỗi tab câu hỏi** |
| Action | `fixTabNumbering` |
| Module | `src/functions/dgnl/fixTabNumbering.js` |
| API sets | WordApi 1.3 (`listItem`, `list`, `detachFromList`, `attachToList`, `setLevelStartingNumber`) |
| Destructive | No. It changes list membership and paragraph formatting only. |
| Idempotent | Yes |

## Purpose

Exam files are often assembled from several sources. The `Câu N.` question lines then belong to several different Word lists. This causes:

- Numbering that restarts or skips.
- Mismatched tab positions between the label and the text.
- Inconsistent alignment.

This command merges every question line into one list and numbers the questions 1…N continuously. It also applies a uniform hanging-indent layout.

## Detection

A paragraph is a **question line** when:

- `isListItem === true`;
- `tableNestingLevel === 0` (it is not inside a table);
- `listItem.listString.trim()` matches `/^\D+\d+\.$/`. For example `Câu 12.` matches, while `(1)`, `a)` and `1.` do not.

## Operations

1. **Merge.** The *main list* is the list of the first question line. Every question line in another list is moved into it. All of them are detached first, the batch is synced, and then each is attached with `attachToList(mainId, 0)` at level 0.
2. **Layout.** Every question line is set to:

   | Property | Value |
   | --- | --- |
   | `alignment` | `Justified` |
   | `leftIndent` | `49.6 pt` (1.75 cm) |
   | `firstLineIndent` | `-49.6 pt` (hanging indent) |

3. **Renumber.** `mainList.setLevelStartingNumber(0, 1)`.
4. **Verify.** The command reloads `listString` for every question line and checks that the numbers are exactly `1, 2, …, N`.

## Configuration

| Constant | Default | Meaning |
| --- | --- | --- |
| `IND` | `49.6` | Hanging indent in points |

## Output

`Đã xử lý N câu (gộp M câu vào danh sách chính).` followed by either `Số câu liên tục 1 → N.` or a warning that lists up to 10 positions where the numbering breaks, shown as `#position→actual`.

## Limitations

- Only level 0 is renumbered. Multi-level question lists are flattened to level 0.
- A list item with a `Câu`-style label that is not a question, for example an example numbered `Ví dụ 1.`, matches the pattern too. Such lines are merged into the question list.

## Test checklist

- [ ] A file whose numbering restarts at 1 in the middle ends up numbered continuously 1…N.
- [ ] The tab between `Câu N.` and the text has the same width on every line.
- [ ] A `(1)`/`(2)` list inside a question keeps its own numbering.
- [ ] A second run reports `gộp 0 câu` and `Số câu liên tục`.
