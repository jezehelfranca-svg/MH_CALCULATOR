# M/H Calculator — single-file web app

`MH_Calculator.html` is a browser version of the desktop program
(`FEED_MH_Calculator_V0_4_OPExcelGrid.py`) that produces **the same calculation results**.

It is one self-contained file. Styles, the calculation model, the report generator, the base
data and the logo all live inside it — no external stylesheets, scripts, fonts, images or
network calls of any kind. Loading the page makes exactly one request: the file itself.

## Using it

Open `MH_Calculator.html` in a browser — double-click it, or drag it onto a browser
window. It also works from a USB stick, an email attachment, a network share, or any static
web server. Nothing to install, no Python, no server.

## Tabs

`Edit Inputs` · **`Phase Split`** · `Guide / Help` · `Summary` · `Output_CI` · `Output_TEL` ·
`OP1` · `OP2-Single` · `OP2-Comprehensive` · `Standards_CI` · `Standards_TEL`

All except `Phase Split` mirror the desktop program. The Master Control panel on the left
drives every tab: Part, Case, ratio basis, Outsourcing Minimization, Project, Base M/H, the two
phase durations, phase view and the two outsourcing unit rates.

## Splitting a job into two phases

A job may run in two stages with only part of the scope delivered in the first — a FEED then
detail engineering, a bid stage then execution, or any other pair. **The phases carry no fixed
stage name**: this calculator is used for FEED work, for detail engineering and for bid
estimates, so you name the two phases yourself in Master Control (defaults: `Phase 1` /
`Phase 2`). The names you type appear on every tab and in the Word report.

The `Phase Split` tab assigns each of the 54 activities a **phase 1 share** from 0 to 100%:

- **100** — produced entirely in phase 1
- **0** — deferred entirely to phase 2
- **anything between** — started in phase 1, completed in phase 2 (e.g. a layout at 40 / 60)

Set them one row at a time (`P1` / `50` / `P2` buttons, or type a percentage), a whole section at
once, or every activity at once. Activities are addressed **by row position, not by code** —
four C&I codes (A17–A20) appear twice in the output list.

**What the second-phase figure means.** It is the deferred remainder priced with the same Unit
M/H, so the two phases always reconcile to the whole-programme total — nothing is lost or
invented. The standards tables carry one set of Unit M/H, so treat it as a statement of deferred
scope rather than a separately estimated stage.

**Durations.** Master Control holds a duration per phase. M/M is always Total M/H ÷ Base M/H;
average manpower divides that by the duration of the phase being viewed, so each phase gets its
own headcount. Under `All phases`, manpower is averaged over both durations combined.

**Phase view.** `All phases` / `<phase 1> only` / `<phase 2> only` filters `Output_CI`,
`Output_TEL` and the OP sheets, so a single-phase calculation sheet can be printed. A yellow
banner marks any sheet limited to one phase, and each row's Remarks column states its share —
so a reviewer checking `Unit M/H × Qty` can see why a row does not multiply out. Section 3 of
the Summary always shows both phases and the whole programme side by side, whatever the view.

Leaving every activity at 100 reproduces the calculation exactly as it behaves without phases.

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

Inputs, the phase split and any edits to the calculation standards are saved to browser
localStorage automatically and restored on the next visit. `Export inputs` writes a JSON file in
the same format as the desktop program's `FEED_MH_Calculator_Last_Input.json`, so the two
programs can load each other's files.

The phase fields (`phase_split`, `phase_names`, `months_2`, `phase_view`) are extra keys the
desktop program ignores, so it reads such a file and reports the whole-programme figures. A file
written by the desktop program has no split, so the whole scope stays in phase 1 and the numbers
are unchanged.

## Word report

`Generate Word report` downloads a `.doc` file that Word opens natively, laid out like
`report_utils.py`: cover, Executive Summary, cost estimate, composition charts, per-discipline
Activity detail, representative calculation standards and review comments. Charts are drawn as
table-based bars rather than images, which Word renders reliably.

## Language and naming

The interface and all data strings are in English. Nothing in the app's own voice names a
project stage — titles read `Project M/H Calculation Sheet`, not `FEED …` — because the same
calculator is used for FEED work, detail engineering and bid estimates.

The word FEED does still appear in the **Project Condition difficulty criteria**, e.g.
*"3) FEED projects where we take part in FEED only, with no EPC execution planned"*. Those
sentences are the standard's own wording and hinge on the FEED-vs-EPC distinction; rewording
them would stop the grades matching the Rev.3 standard a reviewer audits against.

Three things stay in Korean **inside the data on purpose**, because they are calculation keys
rather than display text:

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
| Two-phase split | Not available | `Phase Split` tab, user-named phases, per-phase durations and a phase view |
| OP / standards grids | Drawn on a Tkinter canvas | HTML tables — same columns, colours and header structure, with horizontal scrolling |

## Editing the file

The file is organised in labelled sections, in this order:

| Section | Contents |
| --- | --- |
| `<style>` | Screen styles and Excel-like table formatting |
| `<body>` markup | Master Control panel and the 11 tab panels |
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
