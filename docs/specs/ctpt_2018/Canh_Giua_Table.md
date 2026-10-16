# Canh giữa Table — centre every table on the page

| | |
| --- | --- |
| Ribbon | MY_SCRIPTS → CTPT 2018 → **Canh giữa Table** |
| Action | `canhGiuaTable` |
| Module | `src/features/ctpt2018/canhGiuaTable.js` |
| API sets | WordApi 1.3 (`Table.alignment`, `Table.width`, `autoFitWindow`); page setup for the margin check needs WordApiDesktop 1.3 |
| Destructive | No. It changes table position (and the width of too-wide tables). Ctrl+Z restores. |
| Idempotent | Yes |

## Behaviour

- Every table in the document gets **table alignment = Center** (`Table.alignment`, Word's *Table Properties → Table → Alignment*). This positions the **whole table** between the left and right margins. The text alignment inside the cells is **not** changed (maintainer's requirement, 2026-10-17). Word ignores the table's left indent when it is centred.
- Nested tables are centred inside the cell that contains them.
- A top-level table **wider than the text area** of its section (page width − left/right margins − gutter) is shrunk with *AutoFit to window*, so the centred table never crosses the margins. Word desktop only: on Word on the web the command only centres and says so.

Floating tables (text wrapping *Around*) are positioned by their own anchor settings and may not move.

## Test checklist

Run on a **copy** of a real document.

- [ ] A document with three tables of different widths: all are centred; text inside the cells keeps its alignment (left/right/centre as before).
- [ ] A table wider than the margins: after the command it fits between the margins and is centred.
- [ ] A table nested in a cell: centred inside its cell.
- [ ] Running the command twice changes nothing more.

Automated: `tests/ctpt-table.mjs` (mock Word: every table centred, only the too-wide top-level table fitted, Word web warning).
