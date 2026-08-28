/* Detail-engineering M/H calculator - screen layout and event handling.
 *
 * The app does not re-implement the workbook. It carries the workbook's formula
 * graph and recomputes it through XL.Graph, so every figure on screen is the
 * one Excel produces. Editing a cell writes into the graph and everything
 * downstream follows. */
(function () {
  'use strict';

  var G = new XL.Graph(BUNDLE.cells);
  var META = BUNDLE.meta;
  /* Which of the two calculators this file is. The proposal and the execution
   * workbook are different books, not one book with a switch, so each ships as
   * its own app and keeps its own saved inputs. */
  var INFO = (typeof APP_INFO !== 'undefined' && APP_INFO) ? APP_INFO
    : { name: 'M/H Calculator', subtitle: '', key: 'MH_Calculator', source: BUNDLE.source };

  var CASES = [
    { key: 'p1', ref: 'Summary!P1', label: 'Case 1 - Outsourcing-Single' },
    { key: 'p2', ref: 'Summary!P2', label: 'Case 2 - GEC + Outsourcing-Single' },
    { key: 'p3', ref: 'Summary!P3', label: 'Case 3 - GEC + Outsourcing-Comprehensive' }
  ];

  // OP1 columns, per Case: internal M/H, GEC M/H, outsourced M/H, total
  var CASE_COLS = {
    p1: { int: 'J', gec: null, ext: 'V', tot: 'AA' },
    p2: { int: 'L', gec: 'P', ext: 'X', tot: 'AB' },
    p3: { int: 'N', gec: 'P', ext: 'Z', tot: 'AC' }
  };

  var TABS = [
    { id: 'tab-summary', label: 'Summary', render: renderSummary },
    { id: 'tab-op1-ci', label: 'OP1_CI', render: function (el) { renderOp1(el, 'ci'); } },
    { id: 'tab-op1-tel', label: 'OP1_TEL', render: function (el) { renderOp1(el, 'tel'); } },
    { id: 'tab-op2-1', label: 'OP2-1', render: function (el) { renderOp2(el, 'p1'); } },
    { id: 'tab-op2-2', label: 'OP2-2', render: function (el) { renderOp2(el, 'p2'); } },
    { id: 'tab-op2-3', label: 'OP2-3', render: function (el) { renderOp2(el, 'p3'); } },
    { id: 'tab-input-ci', label: 'Input_CI', render: function (el) { renderInput(el, 'ci'); } },
    { id: 'tab-input-tel', label: 'Input_TEL', render: function (el) { renderInput(el, 'tel'); } },
    { id: 'tab-std-ci', label: 'Standards_CI', render: function (el) { renderStd(el, 'ci'); } },
    { id: 'tab-std-tel', label: 'Standards_TEL', render: function (el) { renderStd(el, 'tel'); } },
    { id: 'tab-guide', label: 'Guide / Help', render: renderGuide }
  ];

  var activeTab = 'tab-summary';
  var dirty = {};

  /* ------------------------------------------------------------------ utils */
  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function val(ref) { return G.get(ref); }
  function numOf(ref) { var v = G.get(ref); return typeof v === 'number' ? v : 0; }

  function group(s) {
    var parts = String(s).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
  }
  function fmt(v, d) {
    d = d === undefined ? 0 : d;
    if (typeof v !== 'number' || !isFinite(v)) return esc(v === null || v === undefined ? '' : v);
    return group(XL.xlround(v, d).toFixed(d));
  }
  /* M/H is usually whole, but some rows produce a fraction - the tiered 3D
   * modelling bands, and any total carrying one. Show the decimal only when
   * there is one, so a figure like 21,120.8 is never displayed as 21,121. */
  function mh(v) {
    if (typeof v !== 'number' || !isFinite(v)) return esc(v === null || v === undefined ? '' : v);
    var r = XL.xlround(v, 1);
    return Math.abs(r - Math.round(r)) < 1e-9 ? fmt(r, 0) : fmt(r, 1);
  }

  var toastTimer = null;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  /* --------------------------------------------------------- master control */
  function currentCase(part) {
    var sel = val(part === 'ci' ? 'Summary!C2' : 'Summary!C3');
    for (var i = 0; i < CASES.length; i++) {
      if (String(val(CASES[i].ref)) === String(sel)) return CASES[i].key;
    }
    return 'p1';
  }

  function applyControls() {
    G.set('Input_CI!B5', 'PROJECT: ' + $('mc-project').value);
    G.set('Summary!C2', val(caseRefFor($('mc-case-ci').value)));
    G.set('Summary!C3', val(caseRefFor($('mc-case-tel').value)));
    G.set('Input_CI!E3', $('mc-min-ci').value);
    G.set('Input_TEL!E3', $('mc-min-tel').value);
    refreshAll();
    saveState(true);
  }

  function caseRefFor(key) {
    for (var i = 0; i < CASES.length; i++) if (CASES[i].key === key) return CASES[i].ref;
    return CASES[0].ref;
  }

  function writeControls() {
    var p = String(val('Input_CI!B5') || '');
    $('mc-project').value = p.replace(/^PROJECT:\s*/, '');
    $('mc-case-ci').value = currentCase('ci');
    $('mc-case-tel').value = currentCase('tel');
    $('mc-min-ci').value = String(val('Input_CI!E3') || 'No');
    $('mc-min-tel').value = String(val('Input_TEL!E3') || 'No');
    $('mc-ver-ci').textContent = String(val('산출기준_CI!A5') || '');
    $('mc-ver-tel').textContent = String(val('산출기준_TEL!A5') || '');
  }

  /* --------------------------------------------------------------- tabs */
  function buildTabs() {
    $('tabbar').innerHTML = TABS.map(function (t) {
      return '<button type="button" role="tab" data-tab="' + t.id + '" aria-selected="' +
        (t.id === activeTab) + '">' + esc(t.label) + '</button>';
    }).join('');
    $('tabbar').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-tab]');
      if (b) selectTab(b.getAttribute('data-tab'));
    });
  }

  function selectTab(id) {
    activeTab = id;
    Array.prototype.forEach.call($('tabbar').children, function (b) {
      b.setAttribute('aria-selected', String(b.getAttribute('data-tab') === id));
    });
    TABS.forEach(function (t) { $(t.id).classList.toggle('active', t.id === id); });
    renderIfDirty(id);
    $('tabpanels').scrollTop = 0;
  }

  function renderIfDirty(id) {
    if (!dirty[id]) return;
    var tab = TABS.filter(function (t) { return t.id === id; })[0];
    if (tab) { tab.render($(id)); dirty[id] = false; }
  }

  function refreshAll() {
    G.invalidate();
    writeControls();
    TABS.forEach(function (t) { dirty[t.id] = true; });
    renderIfDirty(activeTab);
  }

  /* ------------------------------------------------------------- Summary */
  function renderSummary(el) {
    var h = '<div class="summary-title">' + esc(val('Summary!B5') || 'M/H Summary') + '</div>';
    h += '<div class="case-bar">' +
      '<span class="chip">' + esc(val('Summary!B7') || 'PROJECT:') + '</span>' +
      '<span class="chip">C&amp;I : ' + esc(val('Summary!C2')) + '</span>' +
      '<span class="chip">TEL : ' + esc(val('Summary!C3')) + '</span>' +
      '<span class="chip' + (String(val('산출기준_CI!A5')).indexOf('외주최소화') >= 0 ? ' on' : '') + '">' +
      'C&amp;I ' + esc(val('산출기준_CI!A5')) + '</span>' +
      '<span class="chip' + (String(val('산출기준_TEL!A5')).indexOf('외주최소화') >= 0 ? ' on' : '') + '">' +
      'TEL ' + esc(val('산출기준_TEL!A5')) + '</span>' +
      '</div>';

    h += '<table class="summary"><caption>1. M/H Summary</caption>';
    h += '<tr><th>구분</th><th>내부</th><th class="gec">GEC</th><th>외주</th><th>Total</th></tr>';
    [['M/H', 12, 0], ['M/M', 13, 1], ['투입 인원', 14, 1]].forEach(function (r) {
      var i = numOf('Summary!C' + r[1]), g = numOf('Summary!D' + r[1]), e = numOf('Summary!E' + r[1]);
      var f = r[0] === 'M/H' ? mh : function (x) { return fmt(x, r[2]); };
      h += '<tr><td class="lbl">' + r[0] + '</td><td>' + f(i) + '</td>' +
        '<td class="gec">' + f(g) + '</td><td>' + f(e) + '</td>' +
        '<td class="b">' + f(i + g + e) + '</td></tr>';
    });
    h += '</table>';

    h += '<table class="summary"><caption>2. M/H 상세내역</caption>';
    h += '<tr><th rowspan="2">구분</th><th colspan="3">C&amp;I설계 Part</th>' +
      '<th colspan="3">통신설계 Part</th></tr>';
    h += '<tr><th>내부</th><th class="gec">GEC</th><th>외주</th>' +
      '<th>내부</th><th class="gec">GEC</th><th>외주</th></tr>';
    [['M/H', 21, 0], ['M/M', 22, 1], ['투입 인원', 23, 1]].forEach(function (r) {
      h += '<tr><td class="lbl">' + r[0] + '</td>' +
        ['C', 'D', 'E', 'F', 'G', 'H'].map(function (c, i) {
          var x = numOf('Summary!' + c + r[1]);
          return '<td' + (i === 1 || i === 4 ? ' class="gec"' : '') + '>' +
            (r[0] === 'M/H' ? mh(x) : fmt(x, r[2])) + '</td>';
        }).join('') + '</tr>';
    });
    h += '</table>';

    h += '<div class="summary-note"><b>Note</b> &nbsp; Every figure here is recomputed from the ' +
      'source workbook\'s own formulas, so it matches ' + esc(BUNDLE.source) + ' cell for cell. ' +
      'Change quantities on the Input tabs, or the Case and version in Master Control.</div>';
    el.innerHTML = h;
  }

  /* ----------------------------------------------------------------- OP1 */
  function renderOp1(el, part) {
    var sheet = part === 'ci' ? 'OP1_CI' : 'OP1_TEL';
    var rows = META.activities[part];
    var h = '<div class="sheet-title">' + esc(val(sheet + '!A1') || sheet) + '</div>';
    h += '<div class="sheet-src">' + esc(val(sheet + '!A2') || '') + '</div>';
    h += '<div class="scrollx"><table class="sheet" style="width:100%">';
    h += '<tr><th class="head-basic" colspan="6">기본 정보</th>' +
      '<th class="c1" colspan="2">Case 1</th>' +
      '<th class="c2" colspan="3">Case 2</th>' +
      '<th class="c3" colspan="3">Case 3</th></tr>';
    h += '<tr>' +
      ['No.', 'Category', 'Activity', '단위', '수량', '난이도'].map(function (x) {
        return '<th class="head-basic">' + x + '</th>'; }).join('') +
      '<th class="c1">내부</th><th class="c1">외주</th>' +
      '<th class="c2">내부</th><th class="c2 gec">GEC</th><th class="c2">외주</th>' +
      '<th class="c3">내부</th><th class="c3 gec">GEC</th><th class="c3">외주</th></tr>';

    var current = null, no = 1;
    rows.forEach(function (r) {
      if (r.section && r.section !== current) {
        current = r.section;
        h += '<tr class="section"><td colspan="14">' + esc(current) + '</td></tr>';
      }
      var cell = function (c) { return numOf(sheet + '!' + c + r.row); };
      h += '<tr class="' + (no % 2 ? '' : 'odd') + '">' +
        '<td class="ctr">' + no + '</td>' +
        '<td class="ctr">' + esc(r.category || '') + '</td>' +
        '<td class="act">' + esc(r.activity) + '</td>' +
        '<td class="ctr">' + esc(r.unit) + '</td>' +
        '<td class="num">' + fmt(cell('G'), 1) + '</td>' +
        '<td class="ctr">' + esc(val(sheet + '!H' + r.row)) + '</td>' +
        '<td class="num c1">' + mh(cell('J')) + '</td><td class="num c1">' + mh(cell('V')) + '</td>' +
        '<td class="num c2">' + mh(cell('L')) + '</td><td class="num gec">' + mh(cell('P')) +
        '</td><td class="num c2">' + mh(cell('X')) + '</td>' +
        '<td class="num c3">' + mh(cell('N')) + '</td><td class="num gec">' + mh(cell('P')) +
        '</td><td class="num c3">' + mh(cell('Z')) + '</td></tr>';
      no += 1;
    });
    h += '</table></div>';
    el.innerHTML = h;
  }

  /* ----------------------------------------------------------------- OP2 */
  function renderOp2(el, caseKey) {
    var cols = CASE_COLS[caseKey];
    var label = CASES.filter(function (c) { return c.key === caseKey; })[0].label;
    var h = '<div class="sheet-title">M/H 산출서 &mdash; ' + esc(label) + '</div>';
    h += '<div class="sheet-src">Report view of OP1, for the selected Case.</div>';

    ['ci', 'tel'].forEach(function (part) {
      var sheet = part === 'ci' ? 'OP1_CI' : 'OP1_TEL';
      var rows = META.activities[part];
      var ti = 0, tg = 0, te = 0;
      h += '<div class="tablebanner">' + (part === 'ci' ? 'C&amp;I 설계' : '통신 설계') + '</div>';
      h += '<div class="scrollx"><table class="sheet" style="width:100%"><tr>' +
        '<th class="head-basic">No.</th><th class="head-basic">Category</th>' +
        '<th class="head-basic">Activity</th><th class="head-basic">단위</th>' +
        '<th class="head-basic">수량</th><th class="head-basic">난이도</th>' +
        '<th class="head-mh">내부 M/H</th>' + (cols.gec ? '<th class="gec">GEC M/H</th>' : '') +
        '<th class="head-mh">외주 M/H</th><th class="head-unit">Total</th></tr>';
      var current = null, no = 1;
      rows.forEach(function (r) {
        if (r.section && r.section !== current) {
          current = r.section;
          h += '<tr class="section"><td colspan="' + (cols.gec ? 10 : 9) + '">' + esc(current) + '</td></tr>';
        }
        var iv = numOf(sheet + '!' + cols.int + r.row);
        var gv = cols.gec ? numOf(sheet + '!' + cols.gec + r.row) : 0;
        var ev = numOf(sheet + '!' + cols.ext + r.row);
        ti += iv; tg += gv; te += ev;
        h += '<tr class="' + (no % 2 ? '' : 'odd') + '">' +
          '<td class="ctr">' + no + '</td><td class="ctr">' + esc(r.category || '') + '</td>' +
          '<td class="act">' + esc(r.activity) + '</td><td class="ctr">' + esc(r.unit) + '</td>' +
          '<td class="num">' + fmt(numOf(sheet + '!G' + r.row), 1) + '</td>' +
          '<td class="ctr">' + esc(val(sheet + '!H' + r.row)) + '</td>' +
          '<td class="num ext">' + mh(iv) + '</td>' +
          (cols.gec ? '<td class="num gec">' + mh(gv) + '</td>' : '') +
          '<td class="num ext">' + mh(ev) + '</td>' +
          '<td class="num tot">' + mh(iv + gv + ev) + '</td></tr>';
        no += 1;
      });
      h += '<tr class="total"><td colspan="6">TOTAL</td><td class="num">' + mh(ti) + '</td>' +
        (cols.gec ? '<td class="num">' + mh(tg) + '</td>' : '') +
        '<td class="num">' + mh(te) + '</td><td class="num">' + mh(ti + tg + te) + '</td></tr>';
      h += '</table></div>';
    });
    el.innerHTML = h;
  }

  /* --------------------------------------------------------------- Input */
  /* The Input sheets are three tables side by side: Material & General,
   * Design & Deliverables, and Project Conditions. Each is drawn here as its
   * own table, with the workbook's own titles and headers, so a cell on screen
   * is the cell the workbook has. Only literals are editable - anything the
   * workbook calculates is shown read-only, and cannot drift from Excel. */
  var inputView = { ci: null, tel: null };

  function viewOf(part) {
    if (!inputView[part]) inputView[part] = { block: 0, q: '', editOnly: false, adv: false };
    return inputView[part];
  }

  function blockTag(title) {
    var m = /\(([A-C])\)\s*$/.exec(String(title || ''));
    return m ? m[1] : '';
  }
  function blockName(title) {
    var s = String(title || '').replace(/\s*\([A-C]\)\s*$/, '');
    var i = s.indexOf(' - ');
    return (i >= 0 ? s.slice(i + 3) : s).trim();
  }

  /* A plain number for an input box: no thousands separators to re-parse, and
   * no float dust from the workbook's own arithmetic. */
  function inputNum(v) {
    if (typeof v !== 'number' || !isFinite(v)) return v === null || v === undefined ? '' : String(v);
    return String(XL.xlround(v, 6));
  }
  /* A calculated cell, shown as the workbook shows it: whole where it is
   * whole, with the decimals it actually has where it is not. */
  function autoNum(v) {
    if (typeof v !== 'number' || !isFinite(v)) return esc(v === null || v === undefined ? '' : v);
    return group(String(XL.xlround(v, 3)));
  }

  /* An empty cell reads as 0 through the graph, because that is what Excel
   * does inside a formula. On screen an empty cell is empty. */
  function cellRaw(ref) {
    var c = G.cells[ref];
    if (!c) return '';
    if (c.f !== undefined) return G.get(ref);
    return c.v === undefined ? '' : c.v;
  }

  function rowText(sheet, row, cols) {
    var out = [];
    cols.forEach(function (c) {
      if (c.role === 'sn' || c.role === 'label' || c.role === 'unit' || c.role === 'note') {
        var v = cellRaw(sheet + '!' + c.c + row);
        if (v !== null && v !== undefined && v !== '') out.push(String(v));
      }
    });
    return out.join(' ').toLowerCase();
  }

  function renderInput(el, part) {
    var meta = META.inputs[part];
    var v = viewOf(part);
    if (v.block >= meta.blocks.length) v.block = 0;
    var sheet = meta.sheet;
    var blk = meta.blocks[v.block];

    var h = '<div class="sheet-title">' + esc(sheet) + ' &mdash; ' + esc(meta.title) + '</div>';
    h += '<div class="phase-intro"><b>These are the workbook\'s own input cells.</b> ' +
      'Change one and every sheet recomputes exactly as Excel would. White boxes are ' +
      'yours to set; shaded cells are calculated by the workbook and shown for reference. ' +
      'The ratio columns hold fractions the way the sheet stores them &mdash; 0.3 is 30%.</div>';

    h += '<div class="blockbar">';
    meta.blocks.forEach(function (b, i) {
      h += '<button type="button" data-block="' + i + '" class="' + (i === v.block ? 'on' : '') +
        '"><b>' + esc(blockTag(b.title) || String(i + 1)) + '</b> ' + esc(blockName(b.title)) +
        '</button>';
    });
    h += '<span class="grow"></span>';
    h += '<input type="search" id="in-q" placeholder="Filter rows…" value="' + esc(v.q) + '">';
    h += '<label class="chk"><input type="checkbox" id="in-edit"' + (v.editOnly ? ' checked' : '') +
      '> Editable rows only</label>';
    h += '<label class="chk"><input type="checkbox" id="in-adv"' + (v.adv ? ' checked' : '') +
      '> Hidden columns</label>';
    h += '</div>';

    var cols = blk.cols.filter(function (c) { return v.adv || !c.hidden; });
    var q = v.q.trim().toLowerCase();

    h += '<div class="gridwrap"><table class="inputs grid">';
    h += '<thead><tr><th class="rn">Row</th>';
    cols.forEach(function (c) {
      h += '<th class="r-' + c.role + (c.hidden ? ' hid' : '') + '">' +
        esc(c.h || '') + '<span class="cl">' + c.c + '</span></th>';
    });
    h += '</tr></thead><tbody>';

    var shown = 0;
    blk.rows.forEach(function (r) {
      var edits = r.e || {};
      var editable = Object.keys(edits).some(function (k) { return edits[k].k !== 'calc'; });
      if (v.editOnly && !editable) return;
      if (q && rowText(sheet, r.r, blk.cols).indexOf(q) < 0) return;
      shown += 1;
      h += '<tr><td class="rn">' + r.r + '</td>';
      cols.forEach(function (c) {
        var ref = sheet + '!' + c.c + r.r;
        var raw = cellRaw(ref);
        var kind = edits[c.c];
        var cls = 'r-' + c.role + (c.hidden ? ' hid' : '');
        if (!kind || kind.k === 'calc') {
          var isNum = typeof raw === 'number';
          var body = kind ? autoNum(raw) : esc(raw === null || raw === undefined ? '' : raw);
          h += '<td class="' + cls + (kind ? ' calc' : '') + (isNum ? ' num' : '') + '"' +
            (kind ? ' data-calc="' + esc(ref) + '"' : '') + '>' + body + '</td>';
        } else if (kind.k === 'sel') {
          h += '<td class="' + cls + '"><select class="cell" data-ref="' + esc(ref) + '">' +
            kind.o.map(function (o) {
              return '<option' + (String(raw).trim() === o ? ' selected' : '') + '>' + esc(o) +
                '</option>';
            }).join('') + '</select></td>';
        } else {
          h += '<td class="' + cls + '"><input class="cell num" type="text" inputmode="decimal" ' +
            'data-ref="' + esc(ref) + '" value="' + esc(inputNum(raw)) + '"></td>';
        }
      });
      h += '</tr>';
    });
    h += '</tbody></table></div>';
    if (!shown) h += '<p class="hint" style="padding:10px 2px">No rows match this filter.</p>';
    h += '<p class="hint" style="padding:6px 2px">' + shown + ' of ' + blk.rows.length +
      ' rows &middot; block ' + esc(blk.title) + ' &middot; sheet columns ' +
      blk.cols[0].c + '&ndash;' + blk.cols[blk.cols.length - 1].c + '</p>';
    el.innerHTML = h;
    markEdited(el);
    bindInput(el, part);
  }

  /* Cells that differ from the workbook's own value are flagged, so what you
   * changed is visible at a glance and on a printout. */
  function markEdited(el) {
    el.querySelectorAll('.cell[data-ref]').forEach(function (f) {
      var ref = f.dataset.ref, c = G.cells[ref];
      var now = c && c.v !== undefined ? c.v : null;
      f.parentNode.classList.toggle('mod', BASELINE[ref] !== undefined && BASELINE[ref] !== now);
    });
  }

  function refreshInputCells(el) {
    el.querySelectorAll('td[data-calc]').forEach(function (td) {
      td.innerHTML = autoNum(cellRaw(td.dataset.calc));
    });
    markEdited(el);
  }

  /* An edit patches this sheet in place rather than redrawing it, so the box
   * you are typing in keeps the caret while every other tab is marked stale. */
  function applyEdit(el, ref, value) {
    G.set(ref, value);
    G.invalidate();
    writeControls();
    TABS.forEach(function (t) { dirty[t.id] = true; });
    dirty[activeTab] = false;
    refreshInputCells(el);
    saveState(true);
  }

  function bindInput(el, part) {
    var v = viewOf(part);
    el.querySelectorAll('button[data-block]').forEach(function (b) {
      b.addEventListener('click', function () {
        v.block = parseInt(b.getAttribute('data-block'), 10);
        renderInput(el, part);
      });
    });
    var q = el.querySelector('#in-q');
    var timer = null;
    q.addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(function () {
        v.q = q.value;
        renderInput(el, part);
        var box = el.querySelector('#in-q');
        box.focus();
        box.setSelectionRange(box.value.length, box.value.length);
      }, 220);
    });
    el.querySelector('#in-edit').addEventListener('change', function (e) {
      v.editOnly = e.target.checked; renderInput(el, part);
    });
    el.querySelector('#in-adv').addEventListener('change', function (e) {
      v.adv = e.target.checked; renderInput(el, part);
    });

    el.querySelectorAll('input.cell').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var raw = inp.value.trim();
        var n = raw === '' ? 0 : parseFloat(raw.replace(/,/g, ''));
        if (isNaN(n)) {
          toast('Numbers only.');
          inp.value = inputNum(cellRaw(inp.dataset.ref));
          inp.focus();
          return;
        }
        inp.value = inputNum(n);
        applyEdit(el, inp.dataset.ref, n);
      });
      inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') inp.blur(); });
    });
    el.querySelectorAll('select.cell').forEach(function (sel) {
      sel.addEventListener('change', function () {
        applyEdit(el, sel.dataset.ref, sel.value);
      });
    });
  }

  /* ----------------------------------------------------------- Standards */
  function renderStd(el, part) {
    var sheet = part === 'ci' ? '산출기준_CI' : '산출기준_TEL';
    var h = '<div class="sheet-title">' + esc(val(sheet + '!A1') || sheet) + '</div>';
    h += '<div class="std-topbar"><div class="k">산출기준 적용 Version</div>' +
      '<div class="v">' + esc(val(sheet + '!A5')) + '</div>' +
      '<div class="n">The two versions live in the workbook\'s own cells as ' +
      'IF(A5="일반 Ver.", &lt;standard&gt;, &lt;minimized&gt;), so switching Outsourcing ' +
      'Minimization in Master Control swaps the unit rates exactly as Excel does.</div></div>';
    h += '<div class="scrollx"><table class="std" style="width:100%"><tr>' +
      '<th class="basic">Row</th><th class="basic">Activity</th><th class="basic">단위</th>' +
      '<th colspan="4">내부 Unit M/H</th><th colspan="4">외주 Unit M/H</th></tr>' +
      '<tr><th class="basic"></th><th class="basic"></th><th class="basic"></th>' +
      ['SPI', '상', '중', '하', 'SPI', '상', '중', '하'].map(function (g) {
        return '<th>' + g + '</th>'; }).join('') + '</tr>';
    var maxRow = 200, n = 0;
    for (var r = 10; r <= maxRow; r++) {
      var unit = val(sheet + '!D' + r);
      if (!unit) continue;
      var name = val(sheet + '!C' + r) || val(sheet + '!B' + r);
      var cells = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].map(function (c) {
        var v = val(sheet + '!' + c + r);
        return typeof v === 'number' ? fmt(v, 2) : esc(v || '');
      });
      if (!cells.some(function (x) { return x && x !== '-'; })) continue;
      n += 1;
      h += '<tr class="' + (n % 2 ? '' : 'odd') + '"><td class="ctr">' + r + '</td>' +
        '<td class="act">' + esc(name) + '</td><td class="ctr">' + esc(unit) + '</td>' +
        cells.map(function (x, i) {
          return '<td class="num ' + (i < 4 ? 'vint' : 'vext') + '">' + x + '</td>'; }).join('') +
        '</tr>';
    }
    h += '</table></div>';
    el.innerHTML = h;
  }

  /* --------------------------------------------------------------- Guide */
  function renderGuide(el) {
    function block(head, body) {
      return '<div class="block"><h3>' + esc(head) + '</h3><ul>' +
        body.split('\n').map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('') +
        '</ul></div>';
    }
    var h = '<div class="guide"><h1>' + esc(INFO.name) + '</h1>';
    h += '<div class="sub">Built from ' + esc(BUNDLE.source) + '</div>';
    h += '<h2>How this differs from a re-implementation</h2>';
    h += block('It runs the workbook, it does not copy it',
      'The app carries the workbook\'s formula graph and evaluates it in the browser.\n' +
      'Every figure is what Excel computes, including the tiered rows that do not follow unit x quantity.\n' +
      'That holds for any quantities you enter, not just the sample project.');
    h += block('Three Cases, and GEC as its own column',
      'Case 1 - 외주-단종. Case 2 - GEC + 외주-단종. Case 3 - GEC + 외주-종합.\n' +
      'GEC is reported separately from internal and outsourced, as the workbook does.\n' +
      'The GEC and 외주-종합 ratios are per-activity inputs on the Input tabs, not one fixed rate.');
    h += block('Outsourcing Minimization',
      'The workbook stores both unit-M/H sets in every standards cell and picks between them.\n' +
      'Switching it in Master Control swaps the rates exactly as the workbook does, per discipline.');
    h += '<h2>Using it</h2>';
    h += block('1. Set the Case and version', 'Master Control carries the Case per discipline and the Outsourcing Minimization switch.');
    h += block('2. Enter quantities and conditions',
      'Input_CI and Input_TEL are drawn as the workbook draws them: three blocks side by side.\n' +
      'A - Material & General, B - Design & Deliverables, C - Project Conditions. Pick a block with the buttons above the table.\n' +
      'White boxes are the workbook\'s literal cells and are yours to change; shaded italic cells are calculated and shown for reference.\n' +
      'Project Condition selectors carry the workbook\'s own dropdown lists, so a condition can only be set to a value the sheet accepts.\n' +
      'The ratio columns hold fractions as the sheet stores them - 0.3 means 30%.\n' +
      'Anything you change is highlighted, and the Row column gives the Excel row so a figure can be traced back to the workbook.\n' +
      'Filter rows, or tick Editable rows only, to cut a long block down to what you are working on.');
    h += block('3. Read the result', 'Summary is the reporting view; OP1 shows all three Cases side by side; OP2-1/2/3 are the per-Case reports.');
    h += block('Saved inputs stay with this calculator',
      'Save inputs keeps your edits in this browser under this calculator\'s own key, so the proposal and the execution app never mix.\n' +
      'Export inputs writes the changed cells to a JSON file - only what differs from the workbook, so it stays readable.\n' +
      'Reset to workbook values puts every cell back to the figure the source workbook shipped with.');
    h += '</div>';
    el.innerHTML = h;
  }

  /* --------------------------------------------------------------- state */
  var STORAGE_KEY = INFO.key + '_State';

  function editedCells() {
    var out = {};
    Object.keys(G.cells).forEach(function (ref) {
      var c = G.cells[ref];
      if (c.v !== undefined && BASELINE[ref] !== undefined && BASELINE[ref] !== c.v) out[ref] = c.v;
      else if (c.v !== undefined && BASELINE[ref] === undefined) out[ref] = c.v;
    });
    return out;
  }

  var BASELINE = {};
  function snapshotBaseline() {
    Object.keys(G.cells).forEach(function (ref) {
      var c = G.cells[ref];
      if (c.v !== undefined) BASELINE[ref] = c.v;
    });
  }

  function saveState(silent) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(editedCells()));
      if (!silent) toast('Inputs saved in this browser.');
    } catch (e) { if (!silent) toast('Save failed: ' + e.message); }
  }

  function loadState(showMsg) {
    var raw;
    try { raw = localStorage.getItem(STORAGE_KEY); } catch (e) { raw = null; }
    if (!raw) { if (showMsg) toast('No saved inputs found.'); return false; }
    try {
      var edits = JSON.parse(raw);
      Object.keys(edits).forEach(function (ref) { G.cells[ref] = { v: edits[ref] }; });
      refreshAll();
      if (showMsg) toast('Saved inputs loaded.');
      return true;
    } catch (e) { if (showMsg) toast('Load failed: ' + e.message); return false; }
  }

  function download(name, text, type) {
    var blob = new Blob([text], { type: type });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
  }

  function resetState() {
    if (!confirm('Reset all inputs to the workbook values?')) return;
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
    Object.keys(BASELINE).forEach(function (ref) { G.cells[ref] = { v: BASELINE[ref] }; });
    refreshAll();
    toast('Inputs reset to the workbook values.');
  }

  /* ---------------------------------------------------------------- init */
  function bindControls() {
    ['mc-project', 'mc-case-ci', 'mc-case-tel', 'mc-min-ci', 'mc-min-tel'].forEach(function (id) {
      $(id).addEventListener('change', applyControls);
      $(id).addEventListener('keydown', function (e) { if (e.key === 'Enter') $(id).blur(); });
    });
    $('btn-save').addEventListener('click', function () { saveState(false); });
    $('btn-load').addEventListener('click', function () { loadState(true); });
    $('btn-export').addEventListener('click', function () {
      download(INFO.key + '_Inputs.json', JSON.stringify(editedCells(), null, 2), 'application/json');
      toast('Edited inputs exported.');
    });
    $('btn-import').addEventListener('click', function () { $('file-import').click(); });
    $('file-import').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      if (!f) return;
      var rd = new FileReader();
      rd.onload = function () {
        try {
          var edits = JSON.parse(String(rd.result));
          Object.keys(edits).forEach(function (ref) { G.cells[ref] = { v: edits[ref] }; });
          refreshAll(); saveState(true); toast('Input file loaded.');
        } catch (err) { toast('Import failed: ' + err.message); }
      };
      rd.readAsText(f, 'utf-8');
      e.target.value = '';
    });
    $('btn-reset').addEventListener('click', resetState);
    $('btn-print').addEventListener('click', function () { window.print(); });
  }

  function init() {
    if (typeof HYUNDAI_LOGO_B64 !== 'undefined' && HYUNDAI_LOGO_B64) {
      $('logo').src = 'data:image/png;base64,' + HYUNDAI_LOGO_B64;
    } else { $('logo').style.display = 'none'; }
    $('mc-case-ci').innerHTML = $('mc-case-tel').innerHTML =
      CASES.map(function (c) { return '<option value="' + c.key + '">' + esc(c.label) + '</option>'; }).join('');
    snapshotBaseline();
    buildTabs();
    bindControls();
    loadState(false);
    writeControls();
    refreshAll();
    selectTab(activeTab);
    window.addEventListener('beforeunload', function () { saveState(true); });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
