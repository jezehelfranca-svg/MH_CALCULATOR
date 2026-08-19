/* FEED M/H Calculator - screen layout and event handling
 * A web port of the tab structure and rendering of the desktop App class. */
(function () {
  'use strict';

  var num = MH.num, fmt = MH.fmt, f1 = MH.f1, money = MH.money;
  var SHEET_DATE = '2026-08-11';
  var STORAGE_KEY = 'FEED_MH_Calculator_Last_Input';

  var TABS = [
    { id: 'tab-input', label: 'Edit Inputs', render: renderInput },
    { id: 'tab-guide', label: 'Guide / Help', render: renderGuide },
    { id: 'tab-summary', label: 'Summary', render: renderSummary },
    { id: 'tab-output-ci', label: 'Output_CI', render: function (el) { renderOutput(el, 'ci'); } },
    { id: 'tab-output-tel', label: 'Output_TEL', render: function (el) { renderOutput(el, 'tel'); } },
    { id: 'tab-op1', label: 'OP1', render: function (el) { renderOp(el, 'OP1'); } },
    { id: 'tab-op2-short', label: 'OP2-Single', render: function (el) { renderOp(el, 'OP2-Single'); } },
    { id: 'tab-op2-comp', label: 'OP2-Comprehensive', render: function (el) { renderOp(el, 'OP2-Comprehensive'); } },
    { id: 'tab-std-ci', label: 'Standards_CI', render: function (el) { renderStd(el, 'ci'); } },
    { id: 'tab-std-tel', label: 'Standards_TEL', render: function (el) { renderStd(el, 'tel'); } }
  ];

  var m = new MH.Model();
  var activeTab = 'tab-input';
  var dirty = {};
  var suppressRatioSync = false;

  /* ----------------------------------------------------------------- utils */
  function $(id) { return document.getElementById(id); }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  var toastTimer = null;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  // Compact label for the Part column of the Input / OP tables
  function partShort(p) { return MH.P_SHORT[p] || p; }

  /* ------------------------------------------------------- Master Control */
  function ctlValues() {
    return {
      part: $('mc-part').value,
      case_: $('mc-case').value,
      ratio: $('mc-ratio').value,
      external: $('mc-external').value,
      project: $('mc-project').value,
      base: $('mc-base').value,
      months: $('mc-months').value,
      rate_short: $('mc-rate-short').value,
      rate_comp: $('mc-rate-comp').value,
      ip: $('mc-ip').value,
      ep: $('mc-ep').value
    };
  }

  /* Same behaviour as Python's App.apply_left() */
  function applyLeft() {
    var v = ctlValues();
    m.part = MH.REV[v.part] || 'both';
    m.case_ = v.case_ === 'Outsourcing-Single' ? 'short' : 'comp';
    m.ratio_mode = v.ratio === 'Original basis' ? 'original' : 'custom';
    m.external_min = v.external;
    m.project = v.project;
    m.base_mh = num(v.base);
    m.months = num(v.months);
    m.rate_short = num(v.rate_short);
    m.rate_comp = num(v.rate_comp);
    if (m.external_min === 'Yes') m.ratio_mode = 'custom';
    if (m.ratio_mode === 'custom') {
      m.internal_pct = num(v.ip);
      m.external_pct = 100 - m.internal_pct;
      $('mc-ep').value = f1(m.external_pct);
    }
    m.sync_common();
    m.recalc();
    refreshAll();
    saveState(true);
  }

  /* Same behaviour as Python's App.apply_user_state_to_vars() */
  function writeControls() {
    $('mc-part').value = MH.P[m.part] || 'All';
    $('mc-case').value = m.case_ === 'short' ? 'Outsourcing-Single' : 'Outsourcing-Comprehensive';
    $('mc-ratio').value = m.ratio_mode === 'original' ? 'Original basis' : 'User entered';
    $('mc-external').value = m.external_min;
    $('mc-project').value = String(m.project || '');
    $('mc-base').value = String(m.base_mh);
    $('mc-months').value = String(m.months);
    $('mc-rate-short').value = String(m.rate_short);
    $('mc-rate-comp').value = String(m.rate_comp);
    $('mc-ip').value = f1(m.internal_pct);
    $('mc-ep').value = f1(m.external_pct);
  }

  /* ------------------------------------------------------- tab render mgmt */
  function buildTabs() {
    var bar = $('tabbar');
    bar.innerHTML = TABS.map(function (t) {
      return '<button type="button" role="tab" data-tab="' + t.id + '" aria-selected="' +
        (t.id === activeTab) + '">' + esc(t.label) + '</button>';
    }).join('');
    bar.addEventListener('click', function (e) {
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
    if (!tab) return;
    tab.render($(id));
    dirty[id] = false;
  }

  /* Mark every tab stale, then render only the one currently on screen. */
  function refreshAll() {
    m.sync_common();
    m.recalc();
    if (m.ratio_mode === 'original') {
      suppressRatioSync = true;
      $('mc-ip').value = f1(m.internal_pct);
      $('mc-ep').value = f1(m.external_pct);
      suppressRatioSync = false;
    }
    TABS.forEach(function (t) { dirty[t.id] = true; });
    var focusInfo = captureFocus();
    renderIfDirty(activeTab);
    restoreFocus(focusInfo);
  }

  function captureFocus() {
    var a = document.activeElement;
    if (!a || !a.id || !$(activeTab).contains(a)) return null;
    return { id: a.id, start: a.selectionStart, end: a.selectionEnd };
  }

  function restoreFocus(info) {
    if (!info) return;
    var el = $(info.id);
    if (!el) return;
    el.focus();
    try { if (info.start !== null && el.setSelectionRange) el.setSelectionRange(info.start, info.end); } catch (e) { /* number input */ }
  }

  /* =========================================================== Edit Inputs */
  /* Display-side fallback so a computed difficulty cell is never blank
   * (App.fallback_diff). Returns the raw Korean grade used as a lookup key;
   * MH.diffLabel() turns it into the English label. */
  function fallbackDiff(part, code) {
    var val = m.val(part, code);
    if (val !== null && val !== undefined && val !== '') return val;
    if (!/_DIFF$/.test(code)) return '';
    var base = code.replace('_DIFF', '');
    var s;
    if (part === 'tel') {
      if (base === 'C01' || base === 'C03') {
        s = m.val('tel', base + '_SEL');
        return s === '1)' ? '중' : s === '2)' ? '하' : '중';
      }
      if (base === 'C05') {
        s = m.val('tel', 'C05_SEL');
        return s === '1)' ? '상' : '중';
      }
      return '중';
    }
    var gp = m.val('ci', 'C01_SEL');
    var gd = gp === '1)' ? '상' : gp === '2)' ? '중' : gp === '3)' ? '하' : '중';
    if (base === 'C01') return gd;
    if (['C04', 'C06', 'C08', 'C10', 'C12'].indexOf(base) >= 0) {
      s = m.val('ci', base + '_SEL');
      return (s === '1)' && gd === '상') ? '상' : s === '1)' ? '중' : s === '2)' ? '하' : '중';
    }
    if (['C14', 'C23'].indexOf(base) >= 0) {
      s = m.val('ci', base + '_SEL');
      return s === '1)' ? 'SPI-내부' : s === '2)' ? 'SPI-외부' : gd;
    }
    if (['C17', 'C18', 'C26', 'C27', 'C28', 'C32'].indexOf(base) >= 0) return gd;
    if (['C19', 'C21'].indexOf(base) >= 0) {
      s = m.val('ci', base + '_SEL');
      return (s === '1)' && gd === '상') ? '상' : s === '1)' ? '중' : s === '2)' ? '하' : '중';
    }
    if (base === 'C29') {
      s = m.val('ci', 'C29_SEL');
      return s === '1)' ? '상' : s === '2)' ? '중' : s === '3)' ? '하' : '중';
    }
    return gd;
  }

  function inputKind(x) {
    if (x.type === 'computed') return 'Auto result';
    if (x.code === 'CTRL_EXTERNAL_MIN') return 'Master linked';
    if (x.code.indexOf('CTRL') === 0 || /_QTY$/.test(x.code) || /_VALUE$/.test(x.code)) return 'Required';
    return 'Optional';
  }

  var KIND_CLASS = {
    'Auto result': 'k-auto', 'Master linked': 'k-master',
    'Required': 'k-req', 'Optional': 'k-opt'
  };

  function renderInput(el) {
    var h = '';
    h += '<div class="legend">' +
      '<span><span class="sw" style="background:#FFF2CC"></span>Required input</span>' +
      '<span><span class="sw" style="background:#DDEBF7"></span>Master linked</span>' +
      '<span><span class="sw" style="background:#E2F0D9"></span>Auto result</span>' +
      '<span><span class="sw" style="background:#EFE7FF"></span>Optional input</span>' +
      '<span style="color:#506980">Changing a value recalculates the Summary / Output / OP tabs at once.</span>' +
      '</div>';
    h += '<div class="scrollx"><table class="inputs"><thead><tr>' +
      ['Part', 'Code', 'Type', 'Sub-group', 'Input item', 'Value', 'Unit', 'Note']
        .map(function (t) { return '<th>' + t + '</th>'; }).join('') +
      '</tr></thead><tbody>';

    m.inputs.forEach(function (x, idx) {
      if (x.code === 'CTRL_EXTERNAL_MIN') x.value = m.external_min;
      if (m.part !== 'both' && x.part !== m.part) return;
      var kind = inputKind(x);
      var domId = 'in-' + x.part + '-' + x.code;
      var editor;
      if (x.type === 'computed') {
        var dv = fallbackDiff(x.part, x.code);
        x.value = dv;   // stored raw so the exported file stays desktop-compatible
        editor = '<span class="readonly auto">' + esc(MH.diffLabel(dv)) + '</span>';
      } else if (x.code === 'CTRL_EXTERNAL_MIN') {
        editor = '<span class="readonly master">' + esc(m.external_min) + '</span>';
      } else if (x.type === 'select') {
        editor = '<select class="cell" id="' + domId + '" data-idx="' + idx + '">' +
          (x.options || []).map(function (o) {
            return '<option value="' + esc(o.value) + '"' +
              (String(o.value) === String(x.value) ? ' selected' : '') + '>' + esc(o.text) + '</option>';
          }).join('') + '</select>';
      } else {
        editor = '<input class="cell" type="text" id="' + domId + '" data-idx="' + idx + '"' +
          (x.type === 'number' ? ' inputmode="decimal"' : '') +
          ' value="' + esc(x.value) + '">';
      }
      var note = x.code === 'CTRL_EXTERNAL_MIN'
        ? 'Mirrors the value set in Master Control on the left. This row is not editable.'
        : (x.note || '');
      h += '<tr>' +
        '<td class="ctr">' + esc(partShort(x.part)) + '</td>' +
        '<td class="ctr">' + esc(x.code) + '</td>' +
        '<td class="kind ' + KIND_CLASS[kind] + '">' + kind + '</td>' +
        '<td>' + esc(x.middle || '') + '</td>' +
        '<td class="detail">' + esc(x.detail || '') + '</td>' +
        '<td>' + editor + '</td>' +
        '<td class="ctr">' + esc(x.unit || '') + '</td>' +
        '<td class="note">' + esc(note) + '</td>' +
        '</tr>';
    });
    h += '</tbody></table></div>';
    el.innerHTML = h;

    el.querySelectorAll('select.cell').forEach(function (sel) {
      sel.addEventListener('change', function () { updateInput(m.inputs[+sel.dataset.idx], sel.value); });
    });
    el.querySelectorAll('input.cell').forEach(function (inp) {
      inp.addEventListener('change', function () { updateInput(m.inputs[+inp.dataset.idx], inp.value); });
      inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') inp.blur(); });
    });
  }

  function updateInput(rec, value) {
    if (!rec) return;
    if (String(rec.value) === String(value)) return;
    m.input_changed(rec, value);
    writeControls();
    refreshAll();
    saveState(true);
  }

  /* =============================================================== Summary */
  function renderSummary(el) {
    var ci = m.partTotals('ci');
    var tel = m.partTotals('tel');
    var all = m.partTotals('both');
    var p = m.pct();
    var caseLabel = m.case_ === 'short' ? 'Outsourcing-Single Case' : 'Outsourcing-Comprehensive Case';
    var rate = m.case_ === 'short' ? m.rate_short : m.rate_comp;
    var verLabel = m.external_min === 'Yes' ? 'Outsourcing Minimization Ver.' : 'Standard Ver.';

    function mmOf(v) { return m.base_mh ? v / m.base_mh : 0; }
    function headOf(v) { return (m.base_mh && m.months) ? v / m.base_mh / m.months : 0; }

    var h = '<div class="summary-title">FEED Project M/H Summary (C&amp;I + Telecom)</div>';
    h += '<div class="summary-head">' +
      '<span class="chip">PROJECT : ' + esc(m.project || '-') + '</span>' +
      '<span class="chip' + (m.external_min === 'Yes' ? ' on' : '') + '">Outsourcing Minimization : ' + esc(m.external_min) + '</span>' +
      '<span class="chip">Internal / External ' + f1(p.ip) + ' / ' + f1(p.ep) + '</span>' +
      '<span class="chip">' + esc(caseLabel) + ' &middot; unit rate ' + fmt(rate) + ' KRW</span>' +
      '</div>';

    h += '<table class="summary"><caption>1. M/H Summary</caption>';
    h += '<tr><th rowspan="2">Item</th><th colspan="4">Grand total</th></tr>';
    h += '<tr><th>Internal</th><th>Outsourced</th><th>Total</th><th>Est. cost</th></tr>';
    h += '<tr><td class="lbl">M/H</td><td>' + fmt(all.internal) + '</td><td>' + fmt(all.external) +
      '</td><td>' + fmt(all.total) + '</td><td>' + fmt(all.external * rate) + '</td></tr>';
    h += '<tr><td class="lbl">M/M</td><td>' + f1(mmOf(all.internal)) + '</td><td>' + f1(mmOf(all.external)) +
      '</td><td>' + f1(all.mm) + '</td><td>-</td></tr>';
    h += '<tr><td class="lbl">Manpower</td><td>' + f1(headOf(all.internal)) + '</td><td>' + f1(headOf(all.external)) +
      '</td><td>' + f1(all.avg) + '</td><td>-</td></tr>';
    h += '</table>';

    h += '<table class="summary"><caption>2. M/H breakdown</caption>';
    h += '<tr><th rowspan="3">Item</th><th colspan="4">C&amp;I Design Part (' + esc(caseLabel) + ')</th>' +
      '<th colspan="4">Telecom Design Part (' + esc(caseLabel) + ')</th></tr>';
    h += '<tr><td colspan="4" class="b red">' + verLabel + '</td><td colspan="4" class="b red">' + verLabel + '</td></tr>';
    h += '<tr>' + ['Internal', 'Outsourced', 'Total', 'Est. cost', 'Internal', 'Outsourced', 'Total', 'Est. cost']
      .map(function (t) { return '<th>' + t + '</th>'; }).join('') + '</tr>';

    [['M/H', function (o) { return [fmt(o.internal), fmt(o.external), fmt(o.total), fmt(o.external * rate)]; }],
     ['M/M', function (o) { return [f1(mmOf(o.internal)), f1(mmOf(o.external)), f1(o.mm), '-']; }],
     ['Manpower', function (o) { return [f1(headOf(o.internal)), f1(headOf(o.external)), f1(o.avg), '-']; }]
    ].forEach(function (r) {
      h += '<tr><td class="lbl">' + r[0] + '</td>' +
        r[1](ci).map(function (v) { return '<td>' + v + '</td>'; }).join('') +
        r[1](tel).map(function (v) { return '<td>' + v + '</td>'; }).join('') + '</tr>';
    });
    h += '</table>';

    h += '<div class="summary-note"><b>Note</b> &nbsp; The summary is recalculated immediately from ' +
      'Master Control, the values on Edit Inputs and the Unit M/H on the calculation-standards tabs. ' +
      'The breakdown always shows C&amp;I and Telecom in full, whatever Part is selected.</div>';
    el.innerHTML = h;
  }

  /* ============================================== Output_CI / Output_TEL */
  function renderOutput(el, part) {
    var t = m.partTotals(part);
    var rows = t.rows;
    var title = part === 'ci'
      ? 'C&I Design - Petrochemical FEED Project M/H Calculation Sheet (Rev.3)'
      : 'Telecom Design - Petrochemical FEED Project M/H Calculation Sheet (Rev.3)';
    var ver = m.external_min === 'Yes' ? '[ Outsourcing Minimization Ver. ]' : '[ Standard Ver. ]';
    var rate = m.case_ === 'short' ? m.rate_short : m.rate_comp;
    var base = part === 'ci'
      ? [['Total ICSS (DCS+ESD+F&G)+MMS IO quantity', m.nval('ci', 'A01_QTY'), 'Point'],
         ['Total Instrument quantity', m.nval('ci', 'A02_QTY'), 'Set']]
      : [['Telecom System quantity', m.nval('tel', 'A01_QTY'), 'Set'],
         ['Site Area', m.nval('tel', 'A02_QTY'), 'm2'],
         ['Building quantity', m.nval('tel', 'A03_QTY'), 'Ea']];

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="sheet-ver">' + esc(ver) + '</div>';
    h += '<div style="display:flex;gap:24px;flex-wrap:wrap;align-items:flex-start;">';
    h += '<div><div class="b" style="margin-bottom:4px;">Base Data Summary</div><table class="basebox">';
    base.forEach(function (b) {
      h += '<tr><td class="b">' + esc(b[0]) + '</td><td class="v">' + fmt(b[1]) + '</td><td class="b red">' + esc(b[2]) + '</td></tr>';
    });
    h += '</table></div>';
    h += '<div><div class="b" style="margin-bottom:4px;">Notes</div><div class="notesbox">' +
      'The Outsourcing Minimization value (No/Yes) set in Master Control applies to the whole Output.<br>' +
      'The Case (Outsourcing-Single / Outsourcing-Comprehensive) changes how Unit M/H is split between internal and outsourced work.</div></div>';
    h += '</div>';
    h += '<div class="projbar"><div class="k">PROJECT:</div><div class="v">' + esc(m.project || '') + '</div></div>';
    h += '<div class="tablebanner">MAN-HOUR CALCULATION TABLE</div>';

    h += '<div class="scrollx"><table class="sheet" style="width:100%">';
    h += '<tr>' +
      '<th class="head-basic" rowspan="2">No.</th>' +
      '<th class="head-basic" rowspan="2">Activity</th>' +
      '<th class="head-basic" rowspan="2">Unit</th>' +
      '<th class="head-basic" rowspan="2">Qty</th>' +
      '<th class="head-basic" rowspan="2">Difficulty</th>' +
      '<th colspan="2" style="background:#BFE7F3">Internal M/H</th>' +
      '<th colspan="2" style="background:#C6EFCE">External M/H</th>' +
      '<th colspan="2" style="background:#FFF2CC">Total</th></tr>';
    h += '<tr>' +
      '<th style="background:#BFE7F3">Unit M/H</th><th style="background:#BFE7F3">M/H</th>' +
      '<th style="background:#C6EFCE">Unit M/H</th><th style="background:#C6EFCE">M/H</th>' +
      '<th style="background:#FFF2CC">M/H</th><th style="background:#FFF2CC">Remarks</th></tr>';

    rows.forEach(function (r, i) {
      h += '<tr class="' + (i % 2 ? 'odd' : '') + '">' +
        '<td class="ctr">' + (i + 1) + '</td>' +
        '<td>' + esc(r.activity) + '</td>' +
        '<td class="ctr">' + esc(r.unit) + '</td>' +
        '<td class="num">' + f1(r.qty) + '</td>' +
        '<td class="ctr">' + esc(MH.diffLabel(r.diff)) + '</td>' +
        '<td class="num">' + f1(r.iu) + '</td>' +
        '<td class="num">' + fmt(r.hec) + '</td>' +
        '<td class="num">' + f1(r.eu) + '</td>' +
        '<td class="num">' + fmt(r.ext) + '</td>' +
        '<td class="num b">' + fmt(r.total) + '</td>' +
        '<td></td></tr>';
    });

    h += '<tr class="total"><td colspan="2">Base M/H (1 M/M)</td><td class="num">' + fmt(m.base_mh) +
      '</td><td class="ctr">M/H</td><td class="ctr">Total M/H</td><td class="num" colspan="2">' + fmt(t.internal) +
      '</td><td class="num" colspan="2">' + fmt(t.external) + '</td><td class="num" colspan="2">' + fmt(t.total) + '</td></tr>';
    h += '<tr class="b"><td colspan="2">Design duration</td><td class="num">' + f1(m.months) +
      '</td><td class="ctr">months</td><td class="ctr">Total M/M</td><td class="num" colspan="2">' +
      f1(m.base_mh ? t.internal / m.base_mh : 0) + '</td><td class="num" colspan="2">' +
      f1(m.base_mh ? t.external / m.base_mh : 0) + '</td><td class="num" colspan="2">' + f1(t.mm) + '</td></tr>';
    h += '<tr class="b"><td colspan="2">Outsourcing unit rate</td><td class="num">' + fmt(rate) +
      '</td><td class="ctr">KRW</td><td class="ctr">Average manpower</td><td class="num" colspan="2">' +
      f1((m.base_mh && m.months) ? t.internal / m.base_mh / m.months : 0) + '</td><td class="num" colspan="2">' +
      f1((m.base_mh && m.months) ? t.external / m.base_mh / m.months : 0) + '</td><td class="num" colspan="2">' +
      f1(t.avg) + '</td></tr>';
    h += '<tr><td colspan="4" class="b" style="background:#FFF2CC">Outsourcing estimated cost</td>' +
      '<td colspan="7" class="num b red" style="background:#FFF2CC">' + fmt(t.external * rate) + ' KRW</td></tr>';
    h += '</table></div>';
    el.innerHTML = h;
  }

  /* =============================================== OP1 / OP2 Excel sheets */
  function sectionName(code) {
    var n = parseInt(String(code).replace('A', ''), 10);
    if (isNaN(n)) return 'OTHER';
    if (n <= 6) return '1  GENERAL';
    if (n <= 11) return '2  SPECIFICATION';
    if (n <= 14) return '3  CALCULATION';
    if (n <= 20) return '4  PROCUREMENT / DATA SHEET / MR & TBE';
    if (n <= 31) return '5  DRAWING / INDEX / INTERFACE';
    if (n <= 33) return '6  MTO & INFORM TO OTHER DEPTS.';
    return '7  OTHERS';
  }

  function rowsForSheet(sheet) {
    var c = sheet === 'OP2-Single' ? 'short' : sheet === 'OP2-Comprehensive' ? 'comp' : m.case_;
    return m.withCase(c, function () {
      return m.visible_outputs().map(function (o) { return m.row(o); });
    });
  }

  function renderOp(el, sheet) {
    var rows = rowsForSheet(sheet);
    var isOp1 = sheet === 'OP1';
    var title = m.part === 'ci' ? 'C&I Design - Petrochemical FEED Project M/H Calculation Sheet (Rev.3)'
      : m.part === 'tel' ? 'Telecom Design - Petrochemical FEED Project M/H Calculation Sheet (Rev.3)'
        : 'C&I / Telecom FEED Project M/H Calculation Sheet (Rev.3)';
    var ver = m.external_min === 'Yes' ? '[ Outsourcing Minimization Ver. ]' : '[ Standard Ver. ]';

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="sheet-ver">' + esc(ver) + ' &nbsp;<span style="color:#172033;font-size:13px;">' + esc(sheet) + ' Sheet</span></div>';

    h += '<div style="display:flex;gap:24px;flex-wrap:wrap;align-items:flex-start;">';
    h += '<div><div class="b" style="margin-bottom:4px;">Base Data Summary</div><table class="basebox">';
    if (m.part === 'both' || m.part === 'ci') {
      h += '<tr><td class="b">Total ICSS (DCS+ESD+F&G)+MMS IO quantity</td><td class="v">' +
        fmt(m.nval('ci', 'A01_QTY')) + '</td><td class="red">Enter actual quantity</td></tr>';
      h += '<tr><td class="b">Total Instrument quantity</td><td class="v">' +
        fmt(m.nval('ci', 'A02_QTY')) + '</td><td>F&G Inst. included</td></tr>';
    }
    if (m.part === 'both' || m.part === 'tel') {
      h += '<tr><td class="b">Telecom System quantity</td><td class="v">' +
        fmt(m.nval('tel', 'A01_QTY')) + '</td><td>System Qty</td></tr>';
    }
    h += '</table></div>';
    h += '<div><div class="b" style="margin-bottom:4px;">Notes</div><div class="notesbox">' +
      '1. The FEED calculation basis is the Unit M/H on the calculation-standards tabs together with the Project Conditions.<br>' +
      '2. Activities are grouped into a section tree so the OP1 / OP2 results can be reviewed on one screen.</div></div>';
    h += '</div>';
    h += '<div class="projbar"><div class="k">PROJECT:</div><div class="v">' + esc(m.project || '') + '</div></div>';
    h += '<div class="tablebanner">MAN-HOUR CALCULATION TABLE</div>';

    var cols = isOp1
      ? ['No.', 'Part', 'Activity', 'Unit', 'Qty', 'Difficulty', 'Internal Unit', 'Outsourced Unit',
         'Internal M/H', 'Outsourced M/H', 'Total M/H', 'Remarks']
      : ['No.', 'Part', 'Activity', 'Unit', 'Qty', 'Difficulty', 'Internal M/H', 'Outsourced M/H',
         'Total M/H', 'Remarks'];

    h += '<div class="scrollx"><table class="sheet" style="width:100%">';
    h += '<tr><th class="head-basic" colspan="6">Basic information</th>';
    if (isOp1) {
      h += '<th class="head-unit" colspan="2">Unit M/H</th><th class="head-mh" colspan="2">M/H</th>' +
        '<th class="head-unit" rowspan="2">Total M/H</th><th class="head-rem" rowspan="2">Remarks</th></tr>';
    } else {
      h += '<th class="head-mh" colspan="3">M/H</th><th class="head-rem" rowspan="2">Remarks</th></tr>';
    }
    h += '<tr>';
    cols.forEach(function (name, i) {
      if (i >= cols.length - (isOp1 ? 2 : 1)) return; // already drawn via rowspan
      var cls = i < 6 ? 'head-basic' : (name.indexOf('Unit') >= 0 || name.indexOf('Total') >= 0 ? 'head-unit' : 'head-mh');
      h += '<th class="' + cls + '">' + esc(name) + '</th>';
    });
    h += '</tr>';

    var current = null, no = 1;
    var tot = { hec: 0, ext: 0, total: 0 };
    rows.forEach(function (r) {
      var sec = sectionName(r.code);
      if (sec !== current) {
        current = sec;
        h += '<tr class="section"><td colspan="' + cols.length + '">' + esc(current) + '</td></tr>';
      }
      h += '<tr class="' + (no % 2 ? '' : 'odd') + '">' +
        '<td class="ctr">' + no + '</td>' +
        '<td class="ctr">' + esc(partShort(r.part)) + '</td>' +
        '<td>' + esc(r.activity) + '</td>' +
        '<td class="ctr">' + esc(r.unit) + '</td>' +
        '<td class="num">' + money(r.qty, 1) + '</td>' +
        '<td class="ctr">' + esc(MH.diffLabel(r.diff)) + '</td>';
      if (isOp1) {
        h += '<td class="num int">' + money(r.iu, 2) + '</td><td class="num int">' + money(r.eu, 2) + '</td>';
      }
      h += '<td class="num ext">' + fmt(r.hec) + '</td>' +
        '<td class="num ext">' + fmt(r.ext) + '</td>' +
        '<td class="num tot">' + fmt(r.total) + '</td>' +
        '<td></td></tr>';
      tot.hec += r.hec; tot.ext += r.ext; tot.total += r.total;
      no += 1;
    });

    var lead = isOp1 ? 8 : 6;
    h += '<tr class="total"><td colspan="' + lead + '">SUMMARY</td>' +
      '<td class="num">' + fmt(tot.hec) + '</td><td class="num">' + fmt(tot.ext) + '</td>' +
      '<td class="num">' + fmt(tot.total) + '</td><td></td></tr>';
    h += '</table></div>';

    var mm = m.base_mh ? tot.total / m.base_mh : 0;
    var avg = m.months ? mm / m.months : 0;
    h += '<div class="footline">Base M/H: ' + fmt(m.base_mh) + ' &nbsp;|&nbsp; Total M/M: ' + f1(mm) +
      ' &nbsp;|&nbsp; Design duration: ' + fmt(m.months) + ' months &nbsp;|&nbsp; Average manpower: ' +
      f1(avg) + ' persons/month</div>';
    el.innerHTML = h;
  }

  /* ========================================= Calculation standards CI/TEL */
  function renderStd(el, part) {
    var std = m.std[part];
    var keys = Object.keys(std);
    var isCi = part === 'ci';
    var diffs = isCi ? ['SPI', '상', '중', '하'] : ['상', '중', '하'];
    var title = isCi
      ? 'C&I Design - Petrochemical FEED Project M/H Calculation Standards (Rev.3)'
      : 'Telecom Design - Petrochemical FEED Project M/H Calculation Standards (Rev.3)';
    var version = m.external_min === 'Yes' ? 'Outsourcing Minimization Ver.' : 'Standard Ver.';

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="std-topbar">' +
      '<div class="k">Applied standards version\n(Standard vs Outsourcing Minimization)</div>' +
      '<div class="v">' + esc(version) + '</div>' +
      '<div class="n"><b>Notes</b><br>' +
      '1. M/H is derived from the Project standard difficulty, the Unit M/H difficulty grade and whether SPI is applied.<br>' +
      '2. The yellow / green number cells can be edited directly and feed straight into Summary, Output and OP.</div></div>';

    h += '<div class="scrollx"><table class="std" style="width:100%">';
    h += '<tr><th class="basic" rowspan="3">NO.</th><th class="basic" rowspan="3">Activity</th>' +
      '<th class="basic" rowspan="3">Unit</th>' +
      '<th colspan="' + diffs.length + '">Unit M/H by difficulty</th>' +
      '<th colspan="' + diffs.length + '">Unit M/H by difficulty</th>' +
      '<th class="guide" rowspan="2">MAN-HOUR CALCULATION GUIDELINES</th></tr>';
    h += '<tr><th colspan="' + diffs.length + '">Internal Unit M/H</th>' +
      '<th colspan="' + diffs.length + '">Outsourced Unit M/H</th></tr>';
    h += '<tr>' + diffs.map(function (d) { return '<th>' + esc(MH.diffLabel(d)) + '</th>'; }).join('') +
      diffs.map(function (d) { return '<th>' + esc(MH.diffLabel(d)) + '</th>'; }).join('') +
      '<th class="guide">Guide</th></tr>';

    keys.forEach(function (key, i) {
      var r = std[key];
      h += '<tr class="' + ((i + 1) % 2 ? '' : 'odd') + '">' +
        '<td class="ctr">' + (i + 1) + '</td>' +
        '<td class="act">' + esc(r.activity || '') + '</td>' +
        '<td class="ctr">' + esc(r.unit || '') + '</td>';
      ['int', 'ext'].forEach(function (typ) {
        diffs.forEach(function (d) {
          var v = (r[typ] && r[typ][d] !== undefined) ? r[typ][d] : 0;
          h += '<td class="v' + typ + '"><input class="v" type="text" inputmode="decimal" ' +
            'data-part="' + part + '" data-key="' + esc(key) + '" data-typ="' + typ + '" data-diff="' + esc(d) + '" ' +
            'id="std-' + part + '-' + esc(key) + '-' + typ + '-' + esc(d) + '" value="' + esc(v) + '"></td>';
        });
      });
      h += '<td class="guide">' + esc(r.guide || '') + '</td></tr>';
    });
    h += '</table></div>';
    el.innerHTML = h;

    el.querySelectorAll('input.v').forEach(function (inp) {
      inp.addEventListener('change', function () {
        var raw = inp.value.trim();
        var n;
        if (raw === '') n = 0;
        else {
          n = parseFloat(raw.replace(/,/g, ''));
          if (isNaN(n)) { toast('Numbers only.'); inp.focus(); return; }
        }
        updateStdValue(inp.dataset.part, inp.dataset.key, inp.dataset.typ, inp.dataset.diff, n);
      });
      inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') inp.blur(); });
    });
  }

  /* Editing a standard's Unit M/H also updates the std copy held on each
   * Output row (App.update_std_value). */
  function updateStdValue(part, key, typ, diff, value) {
    var st = m.std[part][key];
    if (!st) return;
    if (!st[typ]) st[typ] = {};
    st[typ][diff] = value;
    var target = st.activity;
    m.outputs.forEach(function (o) {
      if (o.part === part && o.std && o.std.activity === target) {
        if (!o.std[typ]) o.std[typ] = {};
        o.std[typ][diff] = value;
      }
    });
    m.recalc();
    refreshAll();
    saveState(true);
  }

  /* ============================================================ Guide/Help */
  function renderGuide(el) {
    function block(head, body, warn) {
      return '<div class="block' + (warn ? ' warn' : '') + '"><h3>' + esc(head) + '</h3><ul>' +
        body.split('\n').map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('') + '</ul></div>';
    }
    var h = '<div class="guide">';
    h += '<h1>M/H Calculation Guide</h1><div class="sub">Attachment #3-1 &nbsp;|&nbsp; Web Application Guide</div>';
    h += '<h2>I. About each sheet (tab)</h2>';
    [['1. Summary sheet', 'Shows the combined C&I + Telecom M/H summary for the values entered in Master Control.\nPart, Case, Outsourcing Minimization, the internal/external ratio, Base M/H, design duration and unit rates all feed both the Summary and the Output tabs.\nUse it to check the outsourcing estimate, Total M/H, Total M/M and average manpower.'],
     ['2. Standards_CI / Standards_TEL sheets', 'The reference tables of internal and outsourced Unit M/H per difficulty grade.\nUse them to review the activity, unit, difficulty, internal/external Unit M/H and the guidance text.\nEditing a yellow (internal) or green (outsourced) number cell recalculates every result immediately.'],
     ['3. Edit Inputs sheet', 'Where you enter the quantities and Project Conditions the calculation is built on.\nYellow is a required input, blue mirrors Master Control, green is an automatic result and purple is an optional selection.\nSelection items are chosen from a drop-down.'],
     ['4. Output_CI / Output_TEL sheets', 'A per-Part calculation sheet showing activity, unit, quantity, difficulty, internal/external M/H and Total M/H.\nThe Case and Outsourcing Minimization settings from Master Control are applied.'],
     ['5. OP1 / OP2 sheets', 'OP1 shows the internal and external Unit M/H and the resulting M/H for every activity.\nOP2-Single and OP2-Comprehensive are the reporting views, each fixed to its own Case regardless of the Case selected in Master Control.'],
     ['6. Generate Word report', 'Downloads a Word report (.doc) built from the current Master Control values, summary and key calculation results.\nIt is saved to the browser download folder and opens directly in Microsoft Word.']
    ].forEach(function (x) { h += block(x[0], x[1]); });

    h += '<h2>II. About each Case</h2>';
    [['1. Outsourcing-Single Case', 'The externally performed share is priced at the Outsourcing-Single unit rate.\nUse it for ordinary FEED calculations that separate internal review from outsourced execution.'],
     ['2. Outsourcing-Comprehensive Case', 'The externally performed share is priced at the Outsourcing-Comprehensive unit rate.\nUse it to study conditions where one subcontractor performs several disciplines together.']
    ].forEach(function (x) { h += block(x[0], x[1]); });

    h += '<h2>III. About Outsourcing Minimization</h2>';
    [['1. Outsourcing Minimization - Yes', 'The internal/external M/H ratio is fixed at 75/25.\nThe same split applies to the Summary, Output, OP tabs and the Word report.', false],
     ['2. Outsourcing Minimization - No', 'The ratio is derived from the internal and outsourced Unit M/H of the original calculation standards.', false],
     ['3. Original basis vs Outsourcing Minimization', 'The original basis keeps the Unit M/H structure of the source standards.\nTo apply Outsourcing Minimization, use the user-entered ratio.', true]
    ].forEach(function (x) { h += block(x[0], x[1], x[2]); });

    h += '<h2>IV. Order of use</h2>';
    h += block('1. Enter the basic information', 'In Master Control on the left, enter the Project, Part, Case, Base M/H, design duration and outsourcing unit rates.');
    h += block('2. Edit the inputs', 'On the Edit Inputs tab, review the quantities and selection conditions and change whatever needs changing.');
    h += block('3. Check the output and generate the report', 'Review the results on the Summary, Output and OP tabs, then generate the Word report.');

    h += '<h2>V. About saving your inputs</h2>';
    h += block('1. Automatic save', 'Your inputs and any edits to the calculation standards are saved to the browser (localStorage) automatically and restored the next time you open the page.\nUse "Load saved inputs" to return to the last saved state at any time.');
    h += block('2. Export / import a file', '"Export inputs" saves a JSON file.\nIt uses the same format as the desktop program\'s FEED_MH_Calculator_Last_Input.json, so the two can exchange files.');
    h += '</div>';
    el.innerHTML = h;
  }

  /* =============================================================== state */
  function collectState() {
    return {
      project: m.project,
      part: m.part,
      case: m.case_,
      ratio_mode: m.ratio_mode,
      external_min: m.external_min,
      base_mh: m.base_mh,
      months: m.months,
      rate_short: m.rate_short,
      rate_comp: m.rate_comp,
      internal_pct: m.internal_pct,
      external_pct: m.external_pct,
      inputs: m.inputs.map(function (x) { return { part: x.part, code: x.code, value: x.value }; }),
      std: m.std
    };
  }

  function saveState(silent) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(collectState()));
      if (!silent) toast('Inputs saved in this browser.');
    } catch (e) {
      if (!silent) toast('Save failed: ' + e.message);
    }
  }

  /* Same restore logic as Python's App.load_user_state() */
  function applyState(data) {
    m.project = data.project || '';
    m.part = data.part || 'both';
    m.case_ = data.case || 'short';
    m.ratio_mode = data.ratio_mode || 'original';
    m.external_min = data.external_min || 'No';
    m.base_mh = num(data.base_mh === undefined ? 161 : data.base_mh);
    m.months = num(data.months === undefined ? 4 : data.months);
    m.rate_short = num(data.rate_short === undefined ? 32000 : data.rate_short);
    m.rate_comp = num(data.rate_comp === undefined ? 39000 : data.rate_comp);
    m.internal_pct = num(data.internal_pct || 0);
    m.external_pct = num(data.external_pct || 0);

    var mp = {};
    m.inputs.forEach(function (x) { mp[x.part + '|' + x.code] = x; });
    (data.inputs || []).forEach(function (item) {
      var t = mp[item.part + '|' + item.code];
      if (t) t.value = item.value;
    });

    if (data.std && typeof data.std === 'object') {
      ['ci', 'tel'].forEach(function (part) {
        var sd = data.std[part];
        if (!sd || typeof sd !== 'object') return;
        Object.keys(sd).forEach(function (sk) {
          if (m.std[part][sk] && typeof sd[sk] === 'object') {
            Object.keys(sd[sk]).forEach(function (f) { m.std[part][sk][f] = sd[sk][f]; });
          }
        });
      });
      // refresh the std copy held on each Output row as well
      m.outputs.forEach(function (o) {
        if (!o.std) return;
        var table = m.std[o.part] || {};
        var keys = Object.keys(table);
        for (var i = 0; i < keys.length; i++) {
          if (table[keys[i]].activity === o.std.activity) {
            var src = table[keys[i]];
            Object.keys(src).forEach(function (f) { o.std[f] = src[f]; });
            break;
          }
        }
      });
    }
    m.sync_common();
    m.recalc();
    writeControls();
  }

  function loadState(showMsg) {
    var raw;
    try { raw = localStorage.getItem(STORAGE_KEY); } catch (e) { raw = null; }
    if (!raw) {
      if (showMsg) toast('No saved inputs found.');
      return false;
    }
    try {
      applyState(JSON.parse(raw));
      refreshAll();
      if (showMsg) toast('Saved inputs loaded.');
      return true;
    } catch (e) {
      if (showMsg) toast('Load failed: ' + e.message);
      return false;
    }
  }

  function exportState() {
    var blob = new Blob([JSON.stringify(collectState(), null, 2)], { type: 'application/json;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'FEED_MH_Calculator_Input_Backup.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
    toast('Inputs exported as a JSON file.');
  }

  function importState(file) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = JSON.parse(String(reader.result));
        applyState(data);
        refreshAll();
        saveState(true);
        toast('Input file loaded.');
      } catch (e) {
        toast('Import failed: ' + e.message);
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  function resetState() {
    if (!confirm('Reset all inputs?\nThe last saved state will be deleted as well.')) return;
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
    m = new MH.Model();
    writeControls();
    refreshAll();
    toast('Inputs reset.');
  }

  /* ============================================================== Report */
  function reportData() {
    applyLeft();
    var ci = m.partTotals('ci');
    var tel = m.partTotals('tel');
    var total = {
      internal: ci.internal + tel.internal,
      external: ci.external + tel.external,
      total: ci.total + tel.total
    };
    total.mm = m.base_mh ? total.total / m.base_mh : 0;
    total.avg = m.months ? total.mm / m.months : 0;
    return {
      project: m.project || '-', external_min: m.external_min,
      base_mh: m.base_mh, months: m.months,
      rate_short: m.rate_short, rate_comp: m.rate_comp,
      ci: ci, tel: tel, total: total,
      short_cost: total.external * m.rate_short,
      comp_cost: total.external * m.rate_comp,
      std: m.std
    };
  }

  function generateReport() {
    try {
      MHReport.download(reportData(), typeof HYUNDAI_LOGO_B64 !== 'undefined' ? HYUNDAI_LOGO_B64 : null,
        'FEED_MH_Report_Hyundai.doc');
      toast('Word report (.doc) downloaded.');
    } catch (e) {
      toast('Word report generation failed: ' + e.message);
    }
  }

  /* ================================================================ start-up */
  function bindControls() {
    ['mc-part', 'mc-case', 'mc-ratio'].forEach(function (id) {
      $(id).addEventListener('change', applyLeft);
    });
    $('mc-external').addEventListener('change', function () {
      if ($('mc-external').value === 'Yes') {
        $('mc-ratio').value = 'User entered';
        $('mc-ip').value = '75.0';
        $('mc-ep').value = '25.0';
      }
      applyLeft();
    });
    ['mc-project', 'mc-base', 'mc-months', 'mc-rate-short', 'mc-rate-comp'].forEach(function (id) {
      $(id).addEventListener('change', applyLeft);
      $(id).addEventListener('keydown', function (e) { if (e.key === 'Enter') $(id).blur(); });
    });
    $('mc-ip').addEventListener('input', function () {
      if (suppressRatioSync) return;
      var v = parseFloat($('mc-ip').value);
      if (isNaN(v)) return;
      suppressRatioSync = true;
      $('mc-ep').value = f1(100 - v);
      $('mc-ratio').value = 'User entered';
      suppressRatioSync = false;
    });
    $('mc-ep').addEventListener('input', function () {
      if (suppressRatioSync) return;
      var v = parseFloat($('mc-ep').value);
      if (isNaN(v)) return;
      suppressRatioSync = true;
      $('mc-ip').value = f1(100 - v);
      $('mc-ratio').value = 'User entered';
      suppressRatioSync = false;
    });
    $('mc-ip').addEventListener('change', applyLeft);
    $('mc-ep').addEventListener('change', applyLeft);

    $('btn-recalc').addEventListener('click', function () { applyLeft(); toast('Output recalculated.'); });
    $('btn-save').addEventListener('click', function () { saveState(false); });
    $('btn-load').addEventListener('click', function () { loadState(true); });
    $('btn-export').addEventListener('click', exportState);
    $('btn-import').addEventListener('click', function () { $('file-import').click(); });
    $('file-import').addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) importState(e.target.files[0]);
      e.target.value = '';
    });
    $('btn-reset').addEventListener('click', resetState);
    $('btn-report').addEventListener('click', generateReport);
    $('btn-print').addEventListener('click', function () { window.print(); });
  }

  function init() {
    if (typeof HYUNDAI_LOGO_B64 !== 'undefined' && HYUNDAI_LOGO_B64) {
      $('logo').src = 'data:image/png;base64,' + HYUNDAI_LOGO_B64;
    } else {
      $('logo').style.display = 'none';
    }
    buildTabs();
    bindControls();
    writeControls();
    loadState(false);
    refreshAll();
    selectTab(activeTab);
    window.addEventListener('beforeunload', function () { saveState(true); });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
