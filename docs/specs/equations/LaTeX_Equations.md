# Equations — convert between LaTeX and Word equations

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → Equations → **LaTeX → Equation** / **Equation → LaTeX** |
| Actions | `latexToEquation`, `equationToLatex` |
| Module | `src/features/equations/latexEquations.js` (port of the user-tested Script Lab converter; original draft in commit `a15f81e`) |
| API sets | WordApi 1.1 (`getSelection`, `search`, `getOoxml`, `insertOoxml`), WordApi 1.3 (`compareLocationWith`, `getRange("Start"/"End")`, `expandTo`) |
| Destructive | It replaces text with equations, or equations with text. Ctrl+Z restores the previous state. |
| Idempotent | Yes. Converted items no longer match, so a second run only converts what is left. |

## Purpose

These commands convert mathematical content in both directions between LaTeX source written as text and native Word equations (OMML, the format of *Insert → Equation*). This is useful when importing material written in LaTeX, and when exporting Word exams to LaTeX-based tools.

## User flow

Each ribbon item opens a scope dialog, in the style of MathType's *Convert Equations*:

| Button | Behaviour |
| --- | --- |
| **Vùng chọn** | Convert only inside the current selection. If nothing is selected, the command stops and asks you to select text. |
| **Toàn bộ tài liệu** | Convert the whole document body. |
| **Huỷ** / closing the dialog | Stop without changing anything. |

The primary (highlighted) button is **Vùng chọn** when text is selected, and **Toàn bộ tài liệu** otherwise.

## LaTeX → Equation

**Recognized delimiters**

| Source | Result |
| --- | --- |
| `$$ … $$` and `\[ … \]` | display equation (`m:oMathPara`) |
| `$ … $` | inline equation (`m:oMath`) |

To avoid matching currency, an inline `$` must not be followed by whitespace, the closing `$` must not be preceded by whitespace or followed by a digit, and `\$` is ignored. So `$5 and $6` is not treated as an equation.

**Supported LaTeX** (the `conv` parser builds OMML directly, with no external service):
- fractions `\frac \dfrac \tfrac \cfrac` and binomials `\binom`;
- sub- and superscripts and primes; roots `\sqrt[n]{}`;
- n-ary operators `\sum \prod \int \iint \oint \bigcup …`, with `\limits` / `\nolimits`;
- functions `\sin \log \lim \max …` (limits rendered below), and `\operatorname{}`;
- delimiters `\left … \right`, `\lVert \rVert`, `\langle \rangle`, floor and ceiling;
- environments `matrix pmatrix bmatrix Bmatrix vmatrix Vmatrix cases aligned`;
- accents `\hat \bar \vec \dot \tilde …`, and `\overline \underline \overset \underset \overbrace \underbrace`;
- fonts `\mathbb \mathcal \mathfrak \mathbf \mathrm \mathit \mathsf \mathtt`, and `\text{}`;
- Greek letters, relations, arrows and set symbols, including `\frown` / `\smile`, which were added for arc notation such as `\overset{\frown}{AB}`.

A command that is not supported is written as upright text and listed in the result dialog (*Lệnh LaTeX chưa hỗ trợ*).

## Equation → LaTeX

Every `m:oMath` becomes `$…$`, and every `m:oMathPara` (an equation alone in its paragraph) becomes `$$…$$`. The `cv` walker maps:
- fractions;
- scripts;
- n-ary operators;
- functions;
- delimiters, with matrix, `cases` and `binom` detection;
- radicals and accents;
- bars, group characters and `limLow`/`limUpp`;
- matrices and equation arrays;
- the math alphabets (double-struck, script, fraktur) and normal-text runs (`\text{}`).

The LaTeX output is normalized: `\le` becomes `\leq`, and the spacing around operators may differ from the original source.

**Scope details**
- **Selection**: one `getOoxml` and one `insertOoxml` of the selection.
- **Whole document**: paragraphs that contain drawings, pictures, fields or OLE objects (`w:drawing`, `w:pict`, `w:fldChar`, `w:fldSimple`, `w:object`) are **skipped** and listed in the details. Their whole paragraph would have to be rewritten through OOXML, which these elements do not survive reliably.
- **MathType objects** (OLE `w:object`) are not OMML, so they are not converted.

## Performance design

The original script made 2–3 `context.sync()` round-trips **per formula**: search, compare, insert. The add-in batches each phase.

| Direction | Original | Add-in |
| --- | --- | --- |
| LaTeX → Equation | ≈ 3 × N syncs | 4 syncs total. (1) Load the paragraph texts. (2) Run every `search` (long formulas over 250 chars use a head + tail search, because Word Search is limited to 255 chars). (3) Compare locations with the selection. (4) Run every `insertOoxml`. Search ranges are tracked by Word, so all replacements can be queued in a single batch. |
| Equation → LaTeX (document) | `getOoxml` of **every** paragraph, plus one sync per insert | One `body.getOoxml()` locates the paragraphs that contain `m:oMath`, by index mapping (reusing `parseXml` / `documentBody` / `bodyParagraphs` from `removeBreakEng`). Then `getOoxml` runs only for those paragraphs (1 sync), and every insert runs in one sync. If the paragraph counts do not match, the original per-paragraph path is used. |

**Fallback.** If a batched phase throws, `convert()` re-runs the original sequential driver (`latex2eqSeq` / `eq2latexSeq`). This is safe because items that were already converted no longer match the scan.

## Output

Typical results:
- `Đã chuyển N công thức (LaTeX → equation) trong vùng chọn.`, or the same with *trong toàn tài liệu* for the whole document;
- `⚠ Bỏ qua K công thức (xem Chi tiết).`;
- `⚠ Lệnh LaTeX chưa hỗ trợ: …`.

When nothing is found, the dialog says *Không tìm thấy công thức nào để chuyển …*.

## Limitations

- A formula whose text appears more than once in a paragraph is located by occurrence count. This assumes that Word's paragraph text and the search hits agree, which is true for plain text.
- Display detection is structural: an equation alone in its paragraph becomes `$$…$$`, even if it was originally written as `$…$`.
- Hidden text, tracked changes and content controls inside a formula's range are not handled specially.

## Tests

`npm test` builds the add-in and runs these checks:
- `tests/equations-roundtrip.mjs`: LaTeX → OMML → LaTeX in Edge headless (a real `DOMParser`) for 18 representative formulas. Every formula must convert, and none may fall back to `\mathrm{<unknown>}`.
- `tests/equations-batch.mjs`: the batched LaTeX → Equation driver against a mock Word API. It checks duplicate formulas in one paragraph, a formula longer than 250 characters, currency (`$5 and $6`), display versus inline, **at most 4 syncs**, *Huỷ*, and *Vùng chọn* with an empty selection.

## Test checklist (in Word, on a copy)

- [ ] **LaTeX → Equation, Toàn bộ tài liệu**: every `$…$`, `$$…$$` and `\[…\]` becomes an equation, and `$5 and $6` stays text.
- [ ] **LaTeX → Equation, Vùng chọn**: only the formulas inside the selection are converted.
- [ ] **Equation → LaTeX, Toàn bộ tài liệu**: equations become `$…$` / `$$…$$`, and paragraphs with images or fields are reported as skipped.
- [ ] Running *Equation → LaTeX* and then *LaTeX → Equation* restores the equations, with the same paragraph count.
- [ ] *Huỷ* changes nothing, and *Vùng chọn* without a selection shows the "Chưa bôi đen" message.
- [ ] A document with about 100 formulas converts in a few seconds.
