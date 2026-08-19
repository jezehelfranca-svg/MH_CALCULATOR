# FEED M/H Calculator — single-file web app

`FEED_MH_Calculator.html` is a browser version of the desktop program
(`FEED_MH_Calculator_V0_4_OPExcelGrid.py`) that produces **the same calculation results**.

It is one self-contained file. Styles, the calculation model, the report generator, the base
data and the logo all live inside it — no external stylesheets, scripts, fonts, images or
network calls of any kind. Loading the page makes exactly one request: the file itself.

## Using it

Open `FEED_MH_Calculator.html` in a browser — double-click it, or drag it onto a browser
window. It also works from a USB stick, an email attachment, a network share, or any static
web server. Nothing to install, no Python, no server.

## Tabs (same as the desktop program)

`Edit Inputs` · `Guide / Help` · `Summary` · `Output_CI` · `Output_TEL` ·
`OP1` · `OP2-Single` · `OP2-Comprehensive` · `Standards_CI` · `Standards_TEL`

The Master Control panel on the left drives every tab: Part, Case, ratio basis, Outsourcing
Minimization, Project, Base M/H, design duration and the two outsourcing unit rates.

## Result parity

The Python `Model` and the JavaScript model **extracted from the shipped HTML file** were
cross-checked over:

- Part (All / CI / TEL) × Case (Single / Comprehensive) × ratio basis (original /
  outsourcing-minimized / user-entered) × 3 FEED difficulty levels = **57 scenarios**
- Every Activity row of every scenario — quantity, difficulty, internal/outsourced Unit M/H,
  internal/outsourced/total M/H — plus the totals and ratios → all identical
- Number formatting matches Python's `format(v, ',.0f')` / `',.1f'` / `',.2f'`, including
  banker's rounding on exact `.5` values

## Saving and exchanging inputs

Inputs and any edits to the calculation standards are saved to browser localStorage
automatically and restored on the next visit. `Export inputs` writes a JSON file in the same
format as the desktop program's `FEED_MH_Calculator_Last_Input.json`, so the two programs can
load each other's files.

## Word report

`Generate Word report` downloads a `.doc` file that Word opens natively, laid out like
`report_utils.py`: cover, Executive Summary, cost estimate, composition charts, per-discipline
Activity detail, representative calculation standards and review comments. Charts are drawn as
table-based bars rather than images, which Word renders reliably.

## Language

The interface and all data strings are in English. Three things stay in Korean **inside the
data on purpose**, because they are calculation keys rather than display text:

- the 상 / 중 / 하 / SPI keys of each standard's `int` / `ext` Unit M/H maps,
- the dictionary keys of the standards tables,
- the stored value of a computed difficulty input.

Keeping them means the calculation stays byte-identical to the Python model and an exported
state file can still be read by the desktop program — its load path replaces whole `int`/`ext`
maps, so renamed keys would zero out every Unit M/H lookup. `DIFF_LABEL` / `diffLabel()` in the
model block render them as High / Medium / Low / SPI-Internal / SPI-External.

## Differences from the desktop program

| Item | Desktop | Web |
| --- | --- | --- |
| Saving inputs | `FEED_MH_Calculator_Last_Input.json` beside the program | Browser localStorage, plus JSON export/import |
| Word report | `.docx` via `python-docx` | `.doc` (HTML) download; charts become table-based bars |
| Editing a standard's value | Double-click the cell, then type | Click the cell and type |
| Version label on the standards tabs | Hard-coded `일반 Ver.` | Follows the Outsourcing Minimization setting, like every other tab |
| OP / standards grids | Drawn on a Tkinter canvas | HTML tables — same columns, colours and header structure, with horizontal scrolling |

## Editing the file

The file is organised in labelled sections, in this order:

| Section | Contents |
| --- | --- |
| `<style>` | Screen styles and Excel-like table formatting |
| `<body>` markup | Master Control panel and the 10 tab panels |
| GENERATED DATA | Input items, Output activities, calculation standards, logo |
| `MH` | Port of the `Model` class — difficulty derivation, auto quantities, internal/outsourced split, Case logic |
| `MHReport` | Word report generation |
| app | Tab rendering, input handling, state save/restore |

Everything except the GENERATED DATA block is edited by hand.

## Regenerating the data

If the base data inside `FEED_MH_Calculator_V0_4_OPExcelGrid.py` changes, rebuild just the
data block:

```bash
python3 tools/gen_web_data.py
```

It rewrites only the region between the `BEGIN GENERATED DATA` / `END GENERATED DATA` markers
and leaves the rest of the file untouched. Korean display strings are translated through
`tools/ko_en.json`; anything not in that dictionary is reported on stderr and left as-is, so
new text is easy to spot.
