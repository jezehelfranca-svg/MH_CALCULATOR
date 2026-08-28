# Detail Engineering and Proposal M/H Calculators

Two single-file browser calculators built from the 실행사업 Rev.3 workbooks:

| App | Built from | Reference figures (internal / GEC / outsourced) |
| --- | --- | --- |
| `DE_MH_Calculator.html` | `MH_Calculator_Detail Engineering.xlsx` | 21,120.8 / 0 / 13,863 |
| `Proposal_MH_Calculator.html` | `MH_Calculator_Propoal.xlsx` | 27,815.8 / 0 / 8,962 |

Each is one self-contained file — styles, the Excel engine, the workbook's formula graph and
the logo are all inside it. Opening the page makes exactly one request: the file itself.

## Why two apps rather than one with a switch

The proposal workbook is **not** the execution workbook with a setting flipped. It carries 49
hand-edited formulas the execution book does not (26 in `OP1_CI`, 23 in `OP1_TEL`) — for
example `OP1_CI!J36 =ROUND((G36*I36),0)+308` and `OP1_TEL!J36 =SUM(Input_TEL!W8:W20)+ROUND(U36*G36,0)*2`,
and `OP1_TEL!G25` deleted outright. No toggle turns one book into the other, so each ships as
its own calculator, with its own title and its own saved-input key.

## It runs the workbook, it does not copy it

The app does not re-implement the calculation. It carries the workbook's own formula graph —
9,989 cells, 4,187 formulas — and evaluates it in the browser through a small Excel engine
(`src/xlengine.js`): tokenizer, parser and evaluator for `IF`, `IFERROR`, `ROUND`, `ROUNDUP`,
`ROUNDDOWN`, `SUM`, `COUNT`, `COUNTA`, `MIN`, `MAX`, arithmetic, comparisons, `&`, references
and ranges.

That means the figures hold for **any** quantities entered, not only the sample project, and
the rows that do not follow `unit × quantity` — the tiered 3D-modelling bands, the FAT
allowances — behave as the workbook makes them behave.

Excel's own arithmetic quirks are reproduced, because otherwise the numbers drift:

- **15 significant digits.** Excel carries 15 decimal digits, not a double's 17. Without this,
  `85*(1-0.3)*3` is `178.49999999999997` and rounds down to 178, where Excel gives 179.
- **Half-up rounding**, not banker's rounding.
- **An empty cell reads as 0**, as does an omitted trailing argument — `=ROUNDUP(AA39*1.5,)`.
- **Text comparison is case-insensitive**: `IF(AG23="yes", …)` matches a cell holding `Yes`.

## Verification

Both books were recomputed cell by cell and diffed against the values Excel had cached in them:

| Check | Cells | Mismatched |
| --- | --- | --- |
| Python engine, `MH_Calculator_Detail Engineering.xlsx` | 7,105 | 0 |
| Python engine, `MH_Calculator_Propoal.xlsx` | 7,098 | 0 |
| Python engine, `Detail Engineering MH Estimation.xlsx` | 7,147 | 0 |
| JavaScript engine, DE bundle | 3,108 | 0 |
| JavaScript engine, Proposal bundle | 3,101 | 0 |

A full recompute in the browser takes about 45 ms, so every figure on screen updates on each
keystroke-commit.

```bash
python3 tools/verify_de.py "MH_Calculator_Detail Engineering.xlsx"
node tools/verify_de_js.js DE_MH_Calculator.bundle.json DE_MH_Calculator.bundle.expect.json
```

## Tabs

`Summary` · `OP1_CI` · `OP1_TEL` · `OP2-1` · `OP2-2` · `OP2-3` · `Input_CI` · `Input_TEL` ·
`Standards_CI` · `Standards_TEL` · `Guide / Help`

Master Control on the left holds the Project name, the Case per discipline and the Outsourcing
Minimization switch per discipline.

### Three Cases, and GEC as its own column

- **Case 1** — 외주-단종
- **Case 2** — GEC + 외주-단종
- **Case 3** — GEC + 외주-종합

GEC is reported separately from internal and outsourced, as the workbook reports it. The GEC
and 외주-종합 ratios are per-activity inputs on the Input tabs, not one fixed rate.

### Outsourcing Minimization

Every standards cell stores both unit-M/H sets and picks between them, as
`=IF(A5="일반 Ver.", <standard>, <minimized>)`, driven by `Input_CI!E3` / `Input_TEL!E3`.
Switching it in Master Control swaps the rates exactly as the workbook does, per discipline.
Of the 77 C&I activities, 25 differ between the two versions; none of the 106 Telecom ones do.

## The Input tabs

Each Input sheet is three tables side by side, and is drawn that way — the block titles and
column headers are read out of the sheet's own rows 6 and 7 rather than hard-coded:

| Block | `Input_CI` | `Input_TEL` |
| --- | --- | --- |
| **A** Material & General | columns B–S | columns B–L |
| **B** Design & Deliverables | columns T–AC | columns M–X |
| **C** Project Conditions | columns AD–AK | columns Y–AJ |

- **White boxes are the workbook's literal cells** and are yours to change. **Shaded italic
  cells are calculated** by the workbook and shown for reference; they cannot be typed into, so
  the figures cannot drift from Excel.
- **Project Condition selectors carry the workbook's own dropdown lists**, taken from its data
  validations — `1)`…`8)`, `Yes`/`No`, `직접수행`/`외부수행`. Where a workbook validation points
  at a broken source range, the choices are recovered from the rows of the group itself, or
  from the answers its sibling rows give.
- **The ratio columns hold fractions as the sheet stores them** — `0.3` means 30%.
- **The Row column gives the Excel row**, so any figure can be traced back to the workbook.
- Cells you change are highlighted, and stay highlighted on a printout.
- `Filter rows…` and `Editable rows only` cut a long block down; `Hidden columns` reveals the
  columns the workbook hides (`Input_CI` M, O, AI; `Input_TEL` AD, AH).

Editing a cell patches the sheet in place rather than redrawing it, so the box keeps the caret
while every other tab is marked stale and redrawn when next opened.

## Saving inputs

`Save inputs` keeps edits in browser localStorage under **this** calculator's own key
(`DE_MH_Calculator_State` / `Proposal_MH_Calculator_State`), so the two apps never mix.
`Export inputs` writes only the cells that differ from the workbook, as `{"Sheet!Ref": value}`.
`Reset to workbook values` puts every cell back to the figure the source workbook shipped with.

## Rebuilding

```bash
python3 tools/build_de_app.py "MH_Calculator_Detail Engineering.xlsx" DE_MH_Calculator.html
python3 tools/build_de_app.py "MH_Calculator_Propoal.xlsx"            Proposal_MH_Calculator.html
```

`tools/build_de_app.py` calls `tools/build_de_bundle.py` for the formula graph and the sheet
metadata, then inlines `src/app.css`, `src/xlengine.js` and `src/de_app.js` into one file. The
app's name, subtitle and storage key are derived from the workbook, so the proposal build
titles itself as such.

| File | Role |
| --- | --- |
| `tools/build_de_app.py` | assembles the single-file app |
| `tools/build_de_bundle.py` | formula graph + row metadata + the values to check against |
| `tools/inputmeta.py` | Input-sheet blocks, column roles, editable cells, dropdown lists |
| `tools/xlformula.py` | the Excel evaluator in Python |
| `tools/verify_de.py` | recomputes a workbook and diffs it against its cached values |
| `tools/verify_de_js.js` | the same check against the JavaScript engine |
| `src/xlengine.js` | the Excel evaluator in JavaScript, plus the memoising cell graph |
| `src/de_app.js` | tabs, rendering, input handling, saved state |
| `src/app.css` | styles, shared with `MH_Calculator.html` |

`.bundle.json` / `.bundle.expect.json` are build intermediates and are not committed.

## Relationship to `MH_Calculator.html`

`MH_Calculator.html` is a port of the desktop Python program (see `WEB_APP.md`). These two are
ports of the Excel workbooks. The desktop program collapsed the workbook's three Cases into
one and hard-coded `gec_ratio = 0.30`; these apps keep the workbook's three Cases and its
per-activity GEC ratios, because they evaluate the workbook itself.
