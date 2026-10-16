# Chuyển STT Text sang Auto — typed question numbers to automatic numbering

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → CTPT 2018 → **Chuyển STT Text sang Auto** |
| Action | `sttTextToAuto` |
| Module | `src/features/ctpt2018/sttTextToAuto.js` |
| API sets | WordApi 1.3 (`startNewList`, `List.setLevelNumbering`, `setLevelStartingNumber`, `attachToList`) |
| Destructive | No content is lost: the typed label is replaced by an automatic number. Ctrl+Z restores. |
| Idempotent | Yes. Questions that are already auto-numbered are skipped, so a second run finds nothing. |

The reverse of *Công cụ BigKey → Số thứ tự → Chuyển STT Auto sang Text*.

## Behaviour

1. Finds every question paragraph outside tables whose text starts with a **typed** label `Câu N.` or `Câu N:` (`planBlocks`). Auto-numbered questions are left alone.
2. Asks for the start number in a popup, like *Fix lỗi tab câu hỏi*: a number field "Bắt đầu từ" (default 1), **Chuyển** or **Huỷ**. Huỷ, or closing the popup, changes nothing. A value outside 1–9999 is rejected before any change. Without a Word restart since 2.0.0-beta.2 the popup has no field and numbering starts at 1.
3. Deletes each typed label together with the space or tab after it.
4. Starts **one** list on the first question with the level-0 format `Câu n.` (Arabic) and the chosen start number, and attaches every other question to it.
5. Applies a 1 cm hanging indent (attaching to a list changes the indents).
6. Verifies like *Fix lỗi tab câu hỏi*: it reloads the list numbers and warns if they are not consecutive from the start number.

The automatic number uses the paragraph mark's font, so a bold blue typed label becomes a plain number in the paragraph's font.

## Test checklist

Run on a **copy** of a real exam.

- [ ] Exam with `Câu 1.` … `Câu 30.` typed: start 1 → automatic `Câu 1.` … `Câu 30.`, question text unchanged, no leftover space or tab.
- [ ] Start 31 → `Câu 31.` … `Câu 60.`.
- [ ] Huỷ → the document is unchanged.
- [ ] Mixed exam (some questions already auto-numbered): only the typed ones change.
- [ ] Run again → "Không có câu hỏi "Câu N." gõ tay nào".

Automated: `tests/ctpt-stt.mjs` (label search helper; mock Word: popup, list format and start, labels deleted, Huỷ, invalid start).
