# Xoá ĐACT ĐGNL — delete detailed-solution blocks

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → ĐGNL → **Xoá ĐACT ĐGNL** |
| Action | `xoaDactDgnl` |
| Module | `src/functions/dgnl/xoaDactDgnl.js` |
| API sets | WordApi 1.3 (`getRange`, `expandTo`, `Range.tables`) |
| Destructive | **Yes**. It deletes content. Run it on a copy. |
| Idempotent | Yes. After the first run there are no markers left. |

## Purpose

This command produces the student version of an exam by removing every detailed-solution block (ĐACT, *đáp án chi tiết*). That includes the answer, explanation text, tables and inline images.

## Block detection

A block **starts** at each paragraph where `text.trim() === "Lời giải"` and `tableNestingLevel === 0`.

The block **ends** right before the first following paragraph that is outside a table (`tableNestingLevel === 0`) and meets one of these conditions:

| Condition | Meaning |
| --- | --- |
| `firstLineIndent < HANGING_MAX` (‑20 pt) | the start of the next question (hanging indent) |
| `leftIndent < LEFT_MIN` (10 pt) | a lead-in or section header (no indent) |
| `text.trim() === "Lời giải"` | the next solution block |

Paragraphs inside tables never end a block, so a table in the solution is deleted with it. Left indent alone cannot separate questions from solutions, because both use 1.75 cm (49.6 pt). That is why the hanging indent is used.

## Deletion strategy

Deleting a range that contains a table in a single operation can fail with `GeneralException`. Tables are therefore deleted first.

**Fast path** (default, 3 round-trips no matter how many blocks there are):

1. Build `Range = first.getRange("Whole").expandTo(last.getRange("Whole"))` for every block, from the last block to the first, and load `range.tables`.
2. Delete all tables. This step is skipped when there are none.
3. Delete all block ranges.

**Fallback**: if any step of the fast path throws, the command:

1. Re-scans the document. Blocks that were already deleted are gone, so only the remaining blocks are found.
2. Deletes the remaining blocks one at a time, from last to first, with the same table-first order and a try/catch around each block. This is the behaviour of the original Script Lab script.
3. Lists the paragraph indices of any blocks that still failed.

**Verify**: the command reloads the paragraphs and counts the remaining `Lời giải` markers.

## Configuration (top of the module)

| Constant | Default | Change when |
| --- | --- | --- |
| `MARKER` | `"Lời giải"` | the file uses `Hướng dẫn giải`, `Đáp án`, … |
| `HANGING_MAX` | `-20` | questions use a different hanging indent, or none |
| `LEFT_MIN` | `10` | lead-ins or headers are indented |

If you use the command on a file with a different layout, preview the blocks first: temporarily log `blocks` in `scan()` and return before deleting.

## Output

`Tìm thấy N khối đáp án.`, followed by `Đã xoá N khối (gồm T bảng).` (or the fallback statistics), followed by `Đoạn còn lại: P | Còn dòng "Lời giải": 0`.

## Limitations

- Questions that are not formatted with a hanging indent will not end a block, so content may be over-deleted. Check the configuration before you run it.
- Floating shapes anchored inside a block are deleted together with their anchor paragraph.

## Test checklist (run on a copy)

- [ ] Every `Lời giải` block is removed, including its tables and images, and the next question and lead-in stay intact.
- [ ] The dialog reports `Còn dòng "Lời giải": 0`.
- [ ] On a large file (≥ 100 blocks), the command finishes in a few seconds. The original script took about 3 round-trips per block.
