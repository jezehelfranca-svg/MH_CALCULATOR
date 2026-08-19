# FEED M/H Calculator — HTML Web App

A browser version of `FEED_MH_Calculator_V0_4_OPExcelGrid.py` (the Tkinter desktop program)
that produces **the same calculation results**. No install, no Python, no server.

## Running it

Open `web/index.html` in a browser — double-click it, or drag it onto a browser window.

To host it internally, serve the `web` folder as static files:

```
python -m http.server 8000 --directory web    # example
```

## Layout

| File | Contents |
| --- | --- |
| `index.html` | Page skeleton (Master Control + 10 tabs) |
| `css/styles.css` | The desktop program's colour palette and Excel-style table formatting |
| `js/data.js` | Base data extracted from the Python source (96 input rows, 54 output rows, 32 CI / 17 TEL calculation standards) plus the logo |
| `js/model.js` | Port of the `Model` class (difficulty derivation, auto quantities, internal/outsourced split, Case logic) |
| `js/app.js` | Tab rendering, input handling, state save/restore |
| `js/report.js` | Word report generation (same structure as `report_utils.py`) |

## Tabs (same as the desktop program)

`Edit Inputs` · `Guide / Help` · `Summary` · `Output_CI` · `Output_TEL` ·
`OP1` · `OP2-Single` · `OP2-Comprehensive` · `Standards_CI` · `Standards_TEL`

## Result parity

The Python `Model` and the JavaScript `Model` were cross-checked over:

- Part (All / CI / TEL) × Case (Single / Comprehensive) × ratio basis (original / outsourcing-minimized / user) × 3 FEED difficulty levels = **57 scenarios**
- Every Activity row of every scenario (quantity, difficulty, internal/outsourced Unit M/H, internal/outsourced/total M/H) plus the totals and ratios → all identical
- Number formatting matches Python's `format(v, ',.0f')` / `',.1f'` / `',.2f'`, including banker's rounding on exact `.5` values

## Language

The interface and all data strings are in English. Three things stay in Korean **inside the
data on purpose**, because they are calculation keys rather than display text:

- the 상 / 중 / 하 / SPI keys of each standard's `int` / `ext` Unit M/H maps,
- the dictionary keys of the standards tables,
- the stored value of a computed difficulty input.

Keeping them means the calculation stays byte-identical to the Python model and an exported
state file can still be read by the desktop program. `DIFF_LABEL` / `diffLabel()` in
`js/model.js` render them as High / Medium / Low / SPI-Internal / SPI-External.

Data strings are translated when `js/data.js` is generated, using `tools/ko_en.json`.

## Differences from the desktop program

| Item | Desktop | Web |
| --- | --- | --- |
| Saving inputs | `FEED_MH_Calculator_Last_Input.json` beside the program | Auto-saved to browser localStorage, plus JSON export/import |
| Word report | `.docx` via `python-docx` | `.doc` (HTML) download that Word opens natively; charts become table-based bar graphics |
| Editing a standard's value | Double-click the cell, then type | Click the cell and type |
| Version label on the standards tabs | Hard-coded `일반 Ver.` | Follows the Outsourcing Minimization setting, like every other tab |
| OP / standards grids | Drawn on a Tkinter canvas | HTML tables (same columns, colours and header structure, with horizontal scrolling) |

The exported JSON uses the same format as the desktop program's
`FEED_MH_Calculator_Last_Input.json`, so the two can load each other's files.

## Regenerating the data

If the base data in `FEED_MH_Calculator_V0_4_OPExcelGrid.py` changes, rebuild `js/data.js`:

```bash
python3 tools/gen_web_data.py
```

Any Korean string that is not yet in `tools/ko_en.json` is reported on stderr and left
untranslated, so new text is easy to spot.
