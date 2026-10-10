# Fix lỗi break dòng T.Anh — remove spurious "Keep with next" (English)

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → ĐGNL → **Fix lỗi break dòng T.Anh** |
| Action | `removeBreakEng` |
| Module | `src/features/dgnl/removeBreakEng.js` |
| API sets | WordApi 1.1 (`getOoxml`, `insertOoxml`), WordApi 1.3 (`getNextOrNullObject`), WordApi 1.5 optional (`getStyles`, `paragraphFormat.keepWithNext`) |
| Destructive | No. It changes pagination flags only, but it rewrites the affected paragraphs via OOXML. |
| Idempotent | Yes |
| Status | **New implementation.** It replaces a manual procedure and has not yet been run against production files. |

## Problem

In English exam files, the answer options (A–D), the `Lời giải` line, the explanation lines and the reading passages often have **Keep with next** (`<w:keepNext/>`) set as direct paragraph formatting. Word then has to keep the whole chain on one page. When the chain does not fit at the bottom of a page, Word moves all of it to the next page and leaves a large blank area under the question stem.

## Policy

| Paragraph | Keep with next |
| --- | --- |
| Question line `Câu N.`, or any other list item | unchanged (the question stays with its first option line) |
| Section header matching `/^s*Questions?/i` | unchanged (stays with the first question) |
| Option line `A.`–`D.` that is directly followed by another option line | unchanged (the options stay together) |
| Paragraph inside a table | unchanged |
| Every other paragraph whose *effective* keep-with-next is on (last option line, `Lời giải`, explanations, passage paragraphs, empty paragraphs) | **forced off** by writing `<w:keepNext w:val="0"/>` |

The effective value is the direct value when one is set on the paragraph. Otherwise it is the value of the paragraph style (`Style.paragraphFormat.keepWithNext`, which needs WordApi 1.5; on older builds only the direct formatting is considered). The command writes an explicit `val="0"` rather than deleting the element, so a style-level flag cannot come back. It never turns the flag on.

## Implementation

Office JS has no paragraph-level `keepWithNext`, so the flag is read and written through OOXML.

1. **Load** the paragraph collection (`text`, `style`, `isListItem`, `tableNestingLevel`) and `body.getOoxml()` in one round-trip. Then load the style table (`nameLocal`, `paragraphFormat.keepWithNext`).
2. **Select candidates** by structure, following the policy table above.
3. **Map OOXML to paragraphs.**
   - The command parses the body package with `DOMParser`, takes the `/word/document.xml` part, and collects the `<w:p>` elements under `<w:body>` in document order. Paragraphs inside `w:txbxContent` or `mc:Fallback` are excluded, and a trailing empty paragraph is dropped.
   - The mapping is accepted only when the paragraph counts are equal **and** every candidate's text signature matches. The signature is the first 40 letters and digits of `w:t` + `m:t`, compared with the first 40 of `Paragraph.text`.
   - Otherwise the command falls back to calling `getOoxml()` on each candidate. If there are more than 600 candidates it aborts with a message instead.
4. **Safety filter.** Paragraphs that contain `w:sectPr`, `w:fldChar`, `w:fldSimple`, `w:object`, `w:ins`, `w:del` or `w:commentRange*` are **skipped** and counted. These structures do not survive a single-paragraph OOXML round-trip reliably: section breaks, fields, OLE/MathType objects, tracked changes and comments.
5. **Rewrite**, in chunks of 25 paragraphs and from the end of the document to the start. For each chunk:
   - Call `getOoxml()` on each paragraph, set `keepNext val=0` (inserted after `w:pStyle` to keep the schema order), and write it back with `insertOoxml(xml, "Replace")`. Each chunk is wrapped in try/catch, and failures are counted and reported.
   - **Cleanup:** `insertOoxml` can leave an extra empty paragraph behind. An empty paragraph is deleted if it is the last paragraph of the inserted range while the original paragraph was not empty, or if it directly follows the range while the original next paragraph was not empty.
6. **Verify** that the total paragraph count is unchanged. If it changed, a warning is shown.

## Output

`Đã tắt "Keep with next" ở N đoạn.` It is followed, when relevant, by:

- the number of paragraphs skipped by the safety filter;
- the number of failed paragraphs;
- a warning if the paragraph count changed.

## Known limitations

- Each rewritten paragraph is round-tripped through OOXML. Text runs, inline images and OMML equations survive, but the paragraph receives a new `uniqueLocalId`. Bookmarks on the paragraph boundary may move.
- Skipped paragraphs (see step 4) must be fixed by hand: **Paragraph → Line and Page Breaks → Keep with next**.
- Option lines are recognized only when they start with `A.`–`D.` as typed text. Options that are auto-numbered list items count as list items and are left unchanged.
- If Word does not repaginate immediately afterwards, press `Ctrl+End` then `Ctrl+Home`, or switch to Draft view and back.

## Test checklist (run on a copy)

- [ ] The blank area under the affected questions disappears, and the stem plus its first option line still stay together.
- [ ] The A–D lines of one question are never split across pages.
- [ ] Equations and images in the explanations are intact.
- [ ] The paragraph count is unchanged (no warning in the dialog).
- [ ] A second run reports that there is nothing to fix.
