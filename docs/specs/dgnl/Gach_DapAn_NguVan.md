# Gạch chân đáp án Ngữ Văn — mark correct options, remove redundant solution lines (Literature)

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → ĐGNL → **Gạch chân đáp án Ngữ Văn** |
| Action | `gachDapAnNguVan` |
| Module | `src/features/dgnl/gachDapAnNguVan.js` |
| API sets | WordApi 1.3 (`listItemOrNullObject`, `Paragraph.search`, `Range.search`) |
| Destructive | **Yes**. It deletes paragraphs. |
| Idempotent | Yes. A processed section no longer matches, so a second run reports 0. |

## Purpose

Source Literature files carry a solution block under each question:

```
Câu 5. <stem>
A. …    B. …
C. …    D. …
Lời giải
C. <repeated text of option C>          ← answer line
<explanation that repeats the stem…>    ← redundant
<real explanation>
```

The command marks the correct option letter in the options. It then deletes the answer line and any explanation paragraphs that only repeat the question.

## Detection

A **section** is found for every paragraph `i` that meets all of these conditions:

1. `text.trim() === "Lời giải"`.
2. Paragraph `i+1` (the *answer line*) matches `/^\s*([A-D])\./`. This gives the correct letter `L`.
3. Within the 8 paragraphs above `i`, the nearest paragraph that contains option `L` is the *option line*. It matches `(^|\t)\s*L\.`, which works whether the options are on one line separated by tabs or on separate lines.

If any condition fails, the section is skipped. A section that has already been processed fails condition 2.

## Operations

**Step 1: mark the option** (3 round-trips in total, batched across all sections):

1. Search the option line for `L.` with `matchCase`.
2. Inside the first hit, search for `L`.
3. Apply `bold`, underline `Single` and color `#EE0000` to that letter only.

**Step 2: delete redundant paragraphs** (one batch, deleted from the last index to the first):

- The answer line (`i+1`).
- The *explanation* is the run of non-empty paragraphs after the answer line, up to the next question line (a list item starting with `Câu`), the next `Dựa vào thông tin…` lead-in, or the next `Lời giải`. Each explanation paragraph is checked in order:
  - It is normalized: lower-cased, with whitespace and punctuation `.,;:!?"'“”‘’()-–—…` removed.
  - If the normalized text has at least 15 characters **and** is a substring of the normalized text of the 12 paragraphs above `Lời giải`, it is deleted as a repeat of the stem.
  - The check stops at the first paragraph that is not a repeat.
- If every explanation paragraph was deleted, the now-empty `Lời giải` heading is deleted as well.

## Output

`Số câu xử lý: N (gạch chân K).` and `Số đoạn đã xoá: M.`

## Limitations

- The command expects the exact marker `Lời giải`, with no trailing colon and not in a table.
- If the option line is more than 8 paragraphs above the marker, the section is skipped.
- Duplicate detection is a substring test. A short genuine explanation (under 15 normalized characters) is never removed. A real explanation that quotes the stem verbatim **is** removed.

## Test checklist

- [ ] Options in a single tab-separated line and options on separate lines are both marked.
- [ ] Only the letter is marked, not the period or the option text.
- [ ] Explanations that only repeat the stem are removed, and the real explanation is kept.
- [ ] A second run reports `Không tìm thấy câu nào cần xử lý`.
