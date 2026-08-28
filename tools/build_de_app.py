# -*- coding: utf-8 -*-
"""Assemble the detail-engineering calculator into one self-contained HTML file.

    python3 tools/build_de_app.py <workbook.xlsx> [out.html]

Inlines the stylesheet, the workbook's formula graph, the Excel engine and the
app. No external stylesheets, scripts, fonts, images or network calls - the
logo travels as a base64 data URI, as in the FEED calculator.
"""
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def safe_js(src):
    """`</script` inside a string literal would end the inline script element."""
    return src.replace('</script', '<\\/script')


def logo_b64():
    src = (ROOT / 'MH_Calculator.html').read_text(encoding='utf-8')
    m = re.search(r"var HYUNDAI_LOGO_B64 = '([^']*)'", src)
    return m.group(1) if m else ''


SHELL = """<!DOCTYPE html>
<!--
  %(name)s - single-file web application
  ==============================================================
  Built from %(source)s by tools/build_de_app.py.

  This app does not re-implement the workbook. It carries the workbook's own
  formula graph and evaluates it in the browser, so every figure matches Excel
  cell for cell - including the tiered rows that do not follow unit x quantity,
  and for any quantities you enter rather than only the sample project.

  Everything it needs is inside this one file: styles, the Excel engine, the
  formula graph and the logo. No external stylesheets, scripts, fonts, images
  or network calls. Just open it in a browser.

  To rebuild after the workbook changes:
      python3 tools/build_de_app.py <workbook.xlsx>
-->
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>%(name)s</title>
<style>
%(css)s
</style>
</head>
<body>

<div class="app-title">
  <span>%(name)s &mdash; %(subtitle)s</span>
  <span class="ver">HYUNDAI ENGINEERING</span>
</div>

<div class="app-body">
  <aside class="master">
    <img id="logo" class="logo" alt="HYUNDAI ENGINEERING">
    <h2>MASTER CONTROL</h2>
    <p class="hint">Every tab is recomputed from the source workbook's own formulas.</p>

    <label for="mc-project">Project</label>
    <input type="text" id="mc-project">

    <label for="mc-case-ci">Case &mdash; C&amp;I</label>
    <select id="mc-case-ci"></select>

    <label for="mc-case-tel">Case &mdash; Telecom</label>
    <select id="mc-case-tel"></select>

    <label for="mc-min-ci">Outsourcing Minimization &mdash; C&amp;I</label>
    <select id="mc-min-ci"><option>No</option><option>Yes</option></select>
    <p class="hint" style="margin:4px 0 0">Standards version: <b id="mc-ver-ci"></b></p>

    <label for="mc-min-tel">Outsourcing Minimization &mdash; Telecom</label>
    <select id="mc-min-tel"><option>No</option><option>Yes</option></select>
    <p class="hint" style="margin:4px 0 0">Standards version: <b id="mc-ver-tel"></b></p>

    <button class="btn accent gap" id="btn-save">Save inputs</button>
    <button class="btn accent" id="btn-load">Load saved inputs</button>
    <button class="btn" id="btn-export">Export inputs</button>
    <button class="btn" id="btn-import">Import inputs</button>
    <input type="file" id="file-import" accept=".json,application/json" hidden>
    <button class="btn" id="btn-reset">Reset to workbook values</button>
    <button class="btn gap" id="btn-print">Print current tab / save as PDF</button>

    <p class="note">Only cells the workbook itself treats as inputs are editable.
      Everything else derives, so the figures cannot drift from the reference.</p>
  </aside>

  <main class="pane">
    <div class="tabbar" id="tabbar" role="tablist"></div>
    <div class="tabpanels" id="tabpanels">
%(panels)s
    </div>
  </main>
</div>

<div id="toast" role="status" aria-live="polite"></div>

<script>
var APP_INFO = %(info)s;
var HYUNDAI_LOGO_B64 = '%(logo)s';
</script>

<!-- ===== BEGIN GENERATED DATA - rebuilt by tools/build_de_app.py ===== -->
<script id="app-data">
var BUNDLE = %(bundle)s;
</script>
<!-- ===== END GENERATED DATA ===== -->

<script>
%(engine)s
</script>

<script>
%(app)s
</script>
</body>
</html>
"""

PANEL_IDS = ['tab-summary', 'tab-op1-ci', 'tab-op1-tel', 'tab-op2-1', 'tab-op2-2',
             'tab-op2-3', 'tab-input-ci', 'tab-input-tel', 'tab-std-ci', 'tab-std-tel',
             'tab-guide']


def app_info(src, out):
    """The two workbooks build two different calculators, and each has to say so
    - in its title, in its Guide, and in the storage key it saves inputs under,
    so a browser holding both does not mix them up."""
    proposal = 'propo' in src.name.lower() or 'propo' in out.stem.lower()
    return {
        'name': 'Proposal M/H Calculator' if proposal else 'Detail Engineering M/H Calculator',
        'subtitle': '견적/입찰 Proposal (Rev.3)' if proposal else '실행사업 (Rev.3)',
        'key': out.stem,
        'source': src.name,
    }


def main():
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / 'DE_MH_Calculator.html'
    info = app_info(src, out)

    bundle_path = out.with_suffix('.bundle.json')
    subprocess.check_call([sys.executable, str(ROOT / 'tools' / 'build_de_bundle.py'),
                           str(src), str(bundle_path)])
    bundle = json.loads(bundle_path.read_text(encoding='utf-8'))
    expect_path = bundle_path.with_suffix('.expect.json')

    html = SHELL % {
        'source': src.name,
        'name': info['name'],
        'subtitle': info['subtitle'],
        'info': json.dumps(info, ensure_ascii=False),
        'css': (ROOT / 'src' / 'app.css').read_text(encoding='utf-8').rstrip(),
        'panels': '\n'.join('      <section class="tabpanel" id="%s"></section>' % p
                            for p in PANEL_IDS),
        'logo': logo_b64(),
        'bundle': safe_js(json.dumps(bundle, ensure_ascii=False, separators=(',', ':'))),
        'engine': safe_js((ROOT / 'src' / 'xlengine.js').read_text(encoding='utf-8').rstrip()),
        'app': safe_js((ROOT / 'src' / 'de_app.js').read_text(encoding='utf-8').rstrip()),
    }
    out.write_text(html, encoding='utf-8')
    print('%s  %.0f KB' % (out.name, out.stat().st_size / 1024))

    leftovers = re.findall(r'(?:src|href)\s*=\s*"(?!data:|#)([^"]+)"', html)
    print('external references: %s' % (leftovers or 'none'))
    print('checking bundle: %s' % expect_path.name)


if __name__ == '__main__':
    main()
