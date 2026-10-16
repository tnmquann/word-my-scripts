# MY_SCRIPTS Desktop — MathType tools (companion VBA template)

| | |
| --- | --- |
| Ribbon | Tab **MY_SCRIPTS+** → *Công cụ MathType* → **MathType hoá** / **Lệch MathType** / **Lỗi Equation** / **Lệch Pic-Math** |
| File | `installer/MY_SCRIPTS_Desktop.dotm`, built from `companion/MathTools.bas` + `companion/customUI14.xml` by `scripts/build-companion.ps1` |
| Origin | *Công cụ MathType* group of `NHOM-WORD-BIEN-SOAN.dotm` (`S_MATH.bas`) |
| Platforms | Word desktop for Windows (tested) and Mac (not tested). Not available on Word on the web (no VBA). |
| Requires | MathType installed and loaded in Word (for MathType hoá and Lệch MathType). |

## Why a VBA template

Office JS cannot call MathType: MathType is a COM/VBA add-in, and OLE equation objects cannot be created from Office JS. These four tools therefore ship as a small macro-enabled template that the installers copy into Word's Startup folder, which Word trusts by default. It adds its own tab: a VBA ribbon cannot add buttons to the Office JS tab.

## Commands

| Command | Behaviour | Change from the VBA |
| --- | --- | --- |
| MathType hoá | Runs MathType's **Toggle TeX** (`MTCommand_TeXToggle`) on the selection: every `$…$`, `$$…$$`, `\[…\]`, `\(…\)` becomes a MathType object (`Equation.DSMT4`). A selection without delimiters is trimmed and wrapped in `$…$` first. With no selection it asks, then converts the whole document. | The VBA removed every space in the selection (`\alpha x` → `\alphax`), cut it and drove the MathType window with `SendKeys`. Toggle TeX needs no keystrokes, so it also works on Mac. |
| Lệch MathType | Resets the raised/lowered position (`Font.Position`) of MathType objects in the selection (or the whole document) to 0, then opens MathType's *Format Equations* dialog. | The VBA called `Selection.ClearFormatting`, which also removed bold, colours and paragraph formatting. |
| Lỗi Equation | Sets the font of every Word equation (OMML) to Cambria Math and reports how many were changed. | Same behaviour, plus a count. |
| Lệch Pic-Math | Lowers each picture or MathType object in the selection by `-height/2 + 3.5 pt` (height ≥ 14.4 pt) or `-height/2 + 3 pt`, so it sits on the text baseline. | The VBA used the height of the first picture for the whole selection and reset all character formatting of the selection. |

If MathType cannot be called, the command shows *Không gọi được MathType…* instead of failing silently.

## Build and source rules

- Edit `companion/MathTools.bas` (UTF-8) and `companion/customUI14.xml`, then run `npm run companion` (Windows, Word, *Trust access to the VBA project object model*). Commit the rebuilt `.dotm`: CI has no Word and only copies it.
- The build converts non-ASCII string literals to `ChrW()` (the VBA editor stores code in the ANSI code page), strips diacritics from comments, checks every `imageMso` with `CommandBars.GetLabelMso`, and injects the ribbon part.
- `Const` cannot hold `ChrW()`, so Vietnamese constants are written as functions.

## Install locations

| OS | Folder |
| --- | --- |
| Windows | Word option `STARTUP-PATH`, default `%APPDATA%\Microsoft\Word\STARTUP` (`install.ps1`; skip with `-NoDesktopTools`) |
| macOS (unverified) | `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word` (or the non-localized names) (`install.sh`; skip with `--no-desktop-tools`) |

Word must be closed while installing: it keeps Startup templates open.

## Test checklist

- [ ] Select `\alpha x^2 + 1` (no `$`) → one MathType object; the space after the selection stays.
- [ ] Select a paragraph with `$a$`, `$b^2$` and `\[\sqrt{x}\]` → three MathType objects.
- [ ] No selection → confirmation prompt → every delimited formula in the document is converted.
- [ ] Lệch MathType on a raised MathType object → position 0, then the Format Equations dialog opens; bold text around it keeps its formatting.
- [ ] Lỗi Equation on equations pasted with another font → Cambria Math.
- [ ] Lệch Pic-Math on a line with a formula picture → the picture sits on the baseline; the text formatting is unchanged.
- [ ] Without MathType: MathType hoá shows the *Không gọi được MathType* message.
- [ ] Mac: the MY_SCRIPTS+ tab appears after installing with `install.sh` (unverified).

Automated (Windows + Word + MathType): `pwsh -File tests/tools/companion-smoke.ps1`.
