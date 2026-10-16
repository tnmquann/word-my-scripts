# Chuẩn hoá BigKey and Check lỗi đề

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → Chuẩn hoá BigKey → **Chuẩn hoá BigKey (cả đề)** / **Chuẩn hoá BigKey (vùng chọn)** / **Check lỗi đề** |
| Actions | `chuanHoaBigKeyDe`, `chuanHoaBigKeyBoiDen`, `checkLoiDe` |
| Modules | `src/features/bigkey/chuanHoaBigKey.js`, `optionLayout.js`, `checkLoiDe.js` |
| Origin | The *Chuẩn hóa* group of `BigKeyUI.dotm` (VBA): `bkAutoTab` → `bkBonDong`, `bkDoRong`, `bkThayA`, `bkSoTTCau`, `DapanBK`; check: `bkCheckLoiDe` |
| API sets | WordApi 1.3. Page setup (A4, 1.5 cm margins) needs WordApiDesktop 1.3 and is skipped with a warning on Word on the web. |
| Destructive | Chuẩn hoá rewrites formatting, numbering and option layout. Ctrl+Z restores. Check lỗi đề only selects text. |
| Idempotent | Yes. A second run produces the same layout. |

## Purpose

Bring a multiple-choice exam (questions `Câu N.` with options `A.`–`D.`) to the BigKey house format, and find structural errors before formatting.

## Chuẩn hoá BigKey

Scope: the whole document, or (vùng chọn) the paragraphs of the selection. An empty selection stops with a message.

1. **Numbering to text**: automatic list numbers outside tables become literal text (`listString` + space), then the paragraph leaves the list.
2. **Base format** (paragraphs outside tables): style Normal, Times New Roman 12, left indent 1 cm, no space before/after. Whole-document mode on Word desktop also sets every section to A4 portrait, margins 1.5 cm.
3. **Options** (`optionLayout.js`): each A–D block (the paragraph starting with `A.` up to the paragraph containing `D.`, at most 4 paragraphs, before *Lời giải*, outside tables) is rebuilt from its OOXML:
   - The block is tokenized into characters (with their run properties), tabs, breaks and inline objects (drawings, VML, OLE/MathType, `m:oMath`).
   - Labels are found strictly (B–D after a tab or line break), with a lenient fallback.
   - The width of each option is estimated without rendering: 5.6 pt per Times New Roman 12 character, 6.5 pt per math character, drawing width from `wp:extent`, VML width from `v:shape`.
   - Layout: every option ≤ 110 pt → 4 per line (tabs at 5.25 / 9.5 / 13.75 cm); ≤ 235 pt → 2 per line (tab at 9.5 cm); otherwise 1 per line.
   - Labels become bold blue `A. `. The correct option (underlined in the source, or the letter of `Chọn X` when nothing is underlined) gets a bold, underlined red letter.
   - Blocks that do not contain A–D in order are skipped and listed in the result.
4. **Questions**: `Câu N.` becomes `Câu i:` in bold blue, numbered consecutively **from the number of the first question** (the VBA always restarted at 1, which broke exams such as `Câu 31–60` with lead-ins `Questions 31-35`). Question paragraphs get a 1 cm hanging indent. *Lời giải* / *Hướng dẫn giải* become bold blue. `Chọn X` is normalized, highlighted yellow and justified.
5. **Clean-up**: empty paragraphs are deleted unless they hold a drawing, picture, object, equation or table (the last paragraph of the document is kept). Double spaces are collapsed (up to 5 passes).

**Tables are never converted or changed.** The VBA converted every table to text and crashed on exams that contain tables; this port leaves paragraphs in tables untouched and only reports how many there were.

## Check lỗi đề

Reads the document and reports every problem at once (the VBA stopped at the first one), then selects the first faulty question. It never edits the document.

| Check | Message |
| --- | --- |
| Duplicate question number | `Trùng số câu: "Câu N" xuất hiện k lần.` |
| Numbering gap or reorder | `Số câu không liên tục: Câu a → Câu b.` |
| Missing / duplicate option label | `Câu N: thiếu phương án …` / `trùng nhãn phương án …` |
| Option labels out of order | `Câu N: thứ tự phương án bị sai.` |
| No answer | `Câu N: thiếu đáp án (không gạch chân, không có "Chọn X").` |
| Underline ≠ `Chọn X` | `Câu N: gạch chân B nhưng lời giải ghi "Chọn C".` |
| Picture or object between the options and *Lời giải* | `Câu N: có hình nằm giữa phương án và "Lời giải" …` |

**Warnings** (shown with ⚠, not counted as errors): text before the first question, such as a title or a lead-in `Questions 31-35: …`. Tables add an informational line.

## Test checklist

Run on a **copy** of a real exam.

- [ ] An exam with a table (inside the stem or the solution): Chuẩn hoá finishes, the table is unchanged.
- [ ] Short options with MathType objects → 4 per line; medium → 2 per line; sentences → 1 per line. Objects are kept.
- [ ] The correct option is red, bold and underlined: from the source underline, or from `Chọn X` when nothing is underlined.
- [ ] An exam that starts at `Câu 31` keeps 31 as its first number; `Câu 1…` exams are renumbered 1, 2, 3, …
- [ ] Automatic numbering `Câu 1.` (list) is converted to text before renumbering.
- [ ] Vùng chọn: only the selected questions change; page setup is not touched.
- [ ] Word on the web: Chuẩn hoá runs and warns that page setup was skipped.
- [ ] Check lỗi đề on a clean exam: "Không phát hiện lỗi"; a title before Câu 1 only adds a ⚠ line.
- [ ] Check lỗi đề finds a removed `C.`, a duplicated `B.`, a skipped number, and underline ≠ `Chọn X`; the first faulty question is selected.

Automated: `tests/bigkey-options.mjs` (synthetic layouts) and `tests/bigkey-real.mjs <export.json> [--has-errors]` (real exams exported with `tests/tools/export-docx.ps1`).
