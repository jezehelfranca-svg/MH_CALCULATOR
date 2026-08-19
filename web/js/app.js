/* FEED M/H Calculator - 화면 구성 및 이벤트 처리
 * 데스크톱 프로그램(App 클래스)의 탭 구성/렌더링을 웹으로 옮긴 것입니다. */
(function () {
  'use strict';

  var num = MH.num, fmt = MH.fmt, f1 = MH.f1, money = MH.money;
  var SHEET_DATE = '2026-08-11';
  var STORAGE_KEY = 'FEED_MH_Calculator_Last_Input';

  var TABS = [
    { id: 'tab-input', label: 'Input 수정', render: renderInput },
    { id: 'tab-guide', label: 'Guide / Help', render: renderGuide },
    { id: 'tab-summary', label: 'Summary', render: renderSummary },
    { id: 'tab-output-ci', label: 'Output_CI', render: function (el) { renderOutput(el, 'ci'); } },
    { id: 'tab-output-tel', label: 'Output_TEL', render: function (el) { renderOutput(el, 'tel'); } },
    { id: 'tab-op1', label: 'OP1', render: function (el) { renderOp(el, 'OP1'); } },
    { id: 'tab-op2-short', label: 'OP2-단종', render: function (el) { renderOp(el, 'OP2-단종'); } },
    { id: 'tab-op2-comp', label: 'OP2-종합', render: function (el) { renderOp(el, 'OP2-종합'); } },
    { id: 'tab-std-ci', label: '산출기준_CI', render: function (el) { renderStd(el, 'ci'); } },
    { id: 'tab-std-tel', label: '산출기준_TEL', render: function (el) { renderStd(el, 'tel'); } }
  ];

  var m = new MH.Model();
  var activeTab = 'tab-input';
  var dirty = {};
  var suppressRatioSync = false;

  /* ------------------------------------------------------------------ 유틸 */
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

  function partLabel(p) { return MH.P[p] || p; }

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

  /* Python App.apply_left() 과 동일한 처리 */
  function applyLeft() {
    var v = ctlValues();
    m.part = MH.REV[v.part] || 'both';
    m.case_ = v.case_ === '외주-단종' ? 'short' : 'comp';
    m.ratio_mode = v.ratio === '원본 기준' ? 'original' : 'custom';
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

  /* Python App.apply_user_state_to_vars() 과 동일 */
  function writeControls() {
    $('mc-part').value = MH.P[m.part] || '전체';
    $('mc-case').value = m.case_ === 'short' ? '외주-단종' : '외주-종합';
    $('mc-ratio').value = m.ratio_mode === 'original' ? '원본 기준' : '사용자 입력';
    $('mc-external').value = m.external_min;
    $('mc-project').value = String(m.project || '');
    $('mc-base').value = String(m.base_mh);
    $('mc-months').value = String(m.months);
    $('mc-rate-short').value = String(m.rate_short);
    $('mc-rate-comp').value = String(m.rate_comp);
    $('mc-ip').value = f1(m.internal_pct);
    $('mc-ep').value = f1(m.external_pct);
  }

  /* --------------------------------------------------------- 탭 렌더 관리 */
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

  /* 모든 탭을 다시 그려야 함을 표시하고, 현재 보이는 탭만 즉시 렌더합니다. */
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

  /* ============================================================ Input 수정 */
  /* 계산된 난이도 셀이 비어 보이지 않도록 하는 표시용 보정 (App.fallback_diff) */
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
    if (x.type === 'computed') return '자동결과';
    if (x.code === 'CTRL_EXTERNAL_MIN') return 'Master연동';
    if (x.code.indexOf('CTRL') === 0 || /_QTY$/.test(x.code) || /_VALUE$/.test(x.code)) return '필수';
    return '선택';
  }

  var KIND_CLASS = { '자동결과': 'k-auto', 'Master연동': 'k-master', '필수': 'k-req', '선택': 'k-opt' };

  function renderInput(el) {
    var h = '';
    h += '<div class="legend">' +
      '<span><span class="sw" style="background:#FFF2CC"></span>필수 입력</span>' +
      '<span><span class="sw" style="background:#DDEBF7"></span>Master 연동</span>' +
      '<span><span class="sw" style="background:#E2F0D9"></span>자동결과</span>' +
      '<span><span class="sw" style="background:#EFE7FF"></span>선택 입력</span>' +
      '<span style="color:#506980">값을 바꾸면 Summary / Output / OP 탭이 즉시 다시 계산됩니다.</span>' +
      '</div>';
    h += '<div class="scrollx"><table class="inputs"><thead><tr>' +
      ['Part', 'Code', '구분', '중분류', '입력 항목', '입력값', '단위', 'Note']
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
        x.value = dv;
        editor = '<span class="readonly auto">' + esc(dv) + '</span>';
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
        ? '왼쪽 Master Control 값이 그대로 표시됩니다. 이 행에서는 선택/수정하지 않습니다.'
        : (x.note || '');
      h += '<tr>' +
        '<td class="ctr">' + esc(partLabel(x.part)) + '</td>' +
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
    var caseLabel = m.case_ === 'short' ? '외주-단종 Case' : '외주-종합 Case';
    var rate = m.case_ === 'short' ? m.rate_short : m.rate_comp;
    var verLabel = m.external_min === 'Yes' ? '외주최소화 Ver.' : '일반 Ver.';

    function mmOf(v) { return m.base_mh ? v / m.base_mh : 0; }
    function headOf(v) { return (m.base_mh && m.months) ? v / m.base_mh / m.months : 0; }

    var h = '<div class="summary-title">FEED사업 M/H Summary (C&amp;I + 통신)</div>';
    h += '<div class="summary-head">' +
      '<span class="chip">PROJECT : ' + esc(m.project || '-') + '</span>' +
      '<span class="chip' + (m.external_min === 'Yes' ? ' on' : '') + '">Master 외주최소화 : ' + esc(m.external_min) + '</span>' +
      '<span class="chip">내부/외부 ' + f1(p.ip) + ' / ' + f1(p.ep) + '</span>' +
      '<span class="chip">' + esc(caseLabel) + ' · 단가 ' + fmt(rate) + ' 원</span>' +
      '</div>';

    h += '<table class="summary"><caption>1. M/H Summary</caption>';
    h += '<tr><th rowspan="2">구분</th><th colspan="4">합계</th></tr>';
    h += '<tr><th>내부</th><th>외주</th><th>Total</th><th>예가</th></tr>';
    h += '<tr><td class="lbl">M/H</td><td>' + fmt(all.internal) + '</td><td>' + fmt(all.external) +
      '</td><td>' + fmt(all.total) + '</td><td>' + fmt(all.external * rate) + '</td></tr>';
    h += '<tr><td class="lbl">M/M</td><td>' + f1(mmOf(all.internal)) + '</td><td>' + f1(mmOf(all.external)) +
      '</td><td>' + f1(all.mm) + '</td><td>-</td></tr>';
    h += '<tr><td class="lbl">투입 인원</td><td>' + f1(headOf(all.internal)) + '</td><td>' + f1(headOf(all.external)) +
      '</td><td>' + f1(all.avg) + '</td><td>-</td></tr>';
    h += '</table>';

    h += '<table class="summary"><caption>2. M/H 상세내역</caption>';
    h += '<tr><th rowspan="3">구분</th><th colspan="4">C&amp;I설계 Part (' + esc(caseLabel) + ')</th>' +
      '<th colspan="4">통신설계 Part (' + esc(caseLabel) + ')</th></tr>';
    h += '<tr><td colspan="4" class="b red">' + verLabel + '</td><td colspan="4" class="b red">' + verLabel + '</td></tr>';
    h += '<tr>' + ['내부', '외주', 'Total', '예가', '내부', '외주', 'Total', '예가']
      .map(function (t) { return '<th>' + t + '</th>'; }).join('') + '</tr>';

    [['M/H', function (o) { return [fmt(o.internal), fmt(o.external), fmt(o.total), fmt(o.external * rate)]; }],
     ['M/M', function (o) { return [f1(mmOf(o.internal)), f1(mmOf(o.external)), f1(o.mm), '-']; }],
     ['투입 인원', function (o) { return [f1(headOf(o.internal)), f1(headOf(o.external)), f1(o.avg), '-']; }]
    ].forEach(function (r) {
      h += '<tr><td class="lbl">' + r[0] + '</td>' +
        r[1](ci).map(function (v) { return '<td>' + v + '</td>'; }).join('') +
        r[1](tel).map(function (v) { return '<td>' + v + '</td>'; }).join('') + '</tr>';
    });
    h += '</table>';

    h += '<div class="summary-note"><b>Note</b> &nbsp; Summary는 왼쪽 Master Control, Input 수정값, ' +
      '산출기준 Unit M/H를 기준으로 즉시 재계산됩니다. 상세내역은 현재 Part 선택과 무관하게 ' +
      'C&amp;I / 통신 전체를 각각 표시합니다.</div>';
    el.innerHTML = h;
  }

  /* ============================================== Output_CI / Output_TEL */
  function renderOutput(el, part) {
    var t = m.partTotals(part);
    var rows = t.rows;
    var title = part === 'ci' ? 'C&I설계 화공 FEED사업 M/H 산출서 (Rev.3)' : '통신설계 화공 FEED사업 M/H 산출서 (Rev.3)';
    var ver = m.external_min === 'Yes' ? '[ 외주최소화 Ver. ]' : '[ 일반 Ver. ]';
    var rate = m.case_ === 'short' ? m.rate_short : m.rate_comp;
    var base = part === 'ci'
      ? [['Total ICSS (DCS+ESD+F&G)+MMS IO 수량', m.nval('ci', 'A01_QTY'), 'Point'],
         ['Total Instrument 수량', m.nval('ci', 'A02_QTY'), 'Set']]
      : [['Telecom System 수량', m.nval('tel', 'A01_QTY'), 'Set'],
         ['Site Area', m.nval('tel', 'A02_QTY'), 'm2'],
         ['Building 수량', m.nval('tel', 'A03_QTY'), 'Ea']];

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="sheet-ver">' + esc(ver) + '</div>';
    h += '<div style="display:flex;gap:24px;flex-wrap:wrap;align-items:flex-start;">';
    h += '<div><div class="b" style="margin-bottom:4px;">산출 Base Data Summary</div><table class="basebox">';
    base.forEach(function (b) {
      h += '<tr><td class="b">' + esc(b[0]) + '</td><td class="v">' + fmt(b[1]) + '</td><td class="b red">' + esc(b[2]) + '</td></tr>';
    });
    h += '</table></div>';
    h += '<div><div class="b" style="margin-bottom:4px;">Notes</div><div class="notesbox">' +
      '왼쪽 Master Control의 외주최소화 값(No/Yes)이 Output 전체에 그대로 자동 반영됩니다.<br>' +
      'Case(외주-단종 / 외주-종합) 선택에 따라 내부/외주 Unit M/H 배분이 달라집니다.</div></div>';
    h += '</div>';
    h += '<div class="projbar"><div class="k">PROJECT:</div><div class="v">' + esc(m.project || '') + '</div></div>';
    h += '<div class="tablebanner">MAN-HOUR 산출 TABLE</div>';

    h += '<div class="scrollx"><table class="sheet" style="width:100%">';
    h += '<tr>' +
      '<th class="head-basic" rowspan="2">No.</th>' +
      '<th class="head-basic" rowspan="2">Activity</th>' +
      '<th class="head-basic" rowspan="2">단위</th>' +
      '<th class="head-basic" rowspan="2">수량</th>' +
      '<th class="head-basic" rowspan="2">난이도</th>' +
      '<th colspan="2" style="background:#BFE7F3">내부 M/H</th>' +
      '<th colspan="2" style="background:#C6EFCE">외부 M/H</th>' +
      '<th colspan="2" style="background:#FFF2CC">Total</th></tr>';
    h += '<tr>' +
      '<th style="background:#BFE7F3">Unit M/H</th><th style="background:#BFE7F3">M/H</th>' +
      '<th style="background:#C6EFCE">Unit M/H</th><th style="background:#C6EFCE">M/H</th>' +
      '<th style="background:#FFF2CC">M/H</th><th style="background:#FFF2CC">비고</th></tr>';

    rows.forEach(function (r, i) {
      h += '<tr class="' + (i % 2 ? 'odd' : '') + '">' +
        '<td class="ctr">' + (i + 1) + '</td>' +
        '<td>' + esc(r.activity) + '</td>' +
        '<td class="ctr">' + esc(r.unit) + '</td>' +
        '<td class="num">' + f1(r.qty) + '</td>' +
        '<td class="ctr">' + esc(r.diff) + '</td>' +
        '<td class="num">' + f1(r.iu) + '</td>' +
        '<td class="num">' + fmt(r.hec) + '</td>' +
        '<td class="num">' + f1(r.eu) + '</td>' +
        '<td class="num">' + fmt(r.ext) + '</td>' +
        '<td class="num b">' + fmt(r.total) + '</td>' +
        '<td></td></tr>';
    });

    h += '<tr class="total"><td colspan="2">Base M/H (1 M/M)</td><td class="num">' + fmt(m.base_mh) +
      '</td><td class="ctr">M/H</td><td class="ctr">합계 M/H</td><td class="num" colspan="2">' + fmt(t.internal) +
      '</td><td class="num" colspan="2">' + fmt(t.external) + '</td><td class="num" colspan="2">' + fmt(t.total) + '</td></tr>';
    h += '<tr class="b"><td colspan="2">설계기간</td><td class="num">' + f1(m.months) +
      '</td><td class="ctr">개월</td><td class="ctr">합계 M/M</td><td class="num" colspan="2">' +
      f1(m.base_mh ? t.internal / m.base_mh : 0) + '</td><td class="num" colspan="2">' +
      f1(m.base_mh ? t.external / m.base_mh : 0) + '</td><td class="num" colspan="2">' + f1(t.mm) + '</td></tr>';
    h += '<tr class="b"><td colspan="2">외주 적용단가</td><td class="num">' + fmt(rate) +
      '</td><td class="ctr">원</td><td class="ctr">평균 투입 인원</td><td class="num" colspan="2">' +
      f1((m.base_mh && m.months) ? t.internal / m.base_mh / m.months : 0) + '</td><td class="num" colspan="2">' +
      f1((m.base_mh && m.months) ? t.external / m.base_mh / m.months : 0) + '</td><td class="num" colspan="2">' +
      f1(t.avg) + '</td></tr>';
    h += '<tr><td colspan="4" class="b" style="background:#FFF2CC">외주 예가</td>' +
      '<td colspan="7" class="num b red" style="background:#FFF2CC">' + fmt(t.external * rate) + ' 원</td></tr>';
    h += '</table></div>';
    el.innerHTML = h;
  }

  /* ================================================= OP1 / OP2 Excel 시트 */
  function sectionName(code) {
    var n = parseInt(String(code).replace('A', ''), 10);
    if (isNaN(n)) return 'OTHER';
    if (n <= 6) return '1  GENERAL';
    if (n <= 11) return '2  SPECIFICATION';
    if (n <= 14) return '3  CALCULATION';
    if (n <= 20) return '4  구매관련 / DATA SHEET / MR & TBE';
    if (n <= 31) return '5  DRAWING / INDEX / INTERFACE';
    if (n <= 33) return '6  MTO & 타부서 INFORM';
    return '7  OTHERS';
  }

  function rowsForSheet(sheet) {
    var c = sheet === 'OP2-단종' ? 'short' : sheet === 'OP2-종합' ? 'comp' : m.case_;
    return m.withCase(c, function () {
      return m.visible_outputs().map(function (o) { return m.row(o); });
    });
  }

  function renderOp(el, sheet) {
    var rows = rowsForSheet(sheet);
    var isOp1 = sheet === 'OP1';
    var title = m.part === 'ci' ? 'C&I설계 화공 FEED사업 M/H 산출서 (Rev.3)'
      : m.part === 'tel' ? '통신설계 화공 FEED사업 M/H 산출서 (Rev.3)'
        : 'C&I / Telecom FEED사업 M/H 산출서 (Rev.3)';
    var ver = m.external_min === 'Yes' ? '[ 외주최소화 Ver. ]' : '[ 일반 Ver. ]';

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="sheet-ver">' + esc(ver) + ' &nbsp;<span style="color:#172033;font-size:13px;">' + esc(sheet) + ' Sheet</span></div>';

    h += '<div style="display:flex;gap:24px;flex-wrap:wrap;align-items:flex-start;">';
    h += '<div><div class="b" style="margin-bottom:4px;">산출 Base Data Summary</div><table class="basebox">';
    if (m.part === 'both' || m.part === 'ci') {
      h += '<tr><td class="b">Total ICSS (DCS+ESD+F&G)+MMS IO 수량</td><td class="v">' +
        fmt(m.nval('ci', 'A01_QTY')) + '</td><td class="red">Actual수량 기입</td></tr>';
      h += '<tr><td class="b">Total Instrument 수량</td><td class="v">' +
        fmt(m.nval('ci', 'A02_QTY')) + '</td><td>F&G Inst. 포함</td></tr>';
    }
    if (m.part === 'both' || m.part === 'tel') {
      h += '<tr><td class="b">Telecom System 수량</td><td class="v">' +
        fmt(m.nval('tel', 'A01_QTY')) + '</td><td>System Qty</td></tr>';
    }
    h += '</table></div>';
    h += '<div><div class="b" style="margin-bottom:4px;">Notes</div><div class="notesbox">' +
      '1. FEED 산출 기준은 산출기준 탭의 Unit M/H와 Project Condition을 기준으로 적용합니다.<br>' +
      '2. 섹션별 트리 구조로 Activity를 구분하고, OP1/OP2 결과를 동일 화면에서 검토합니다.</div></div>';
    h += '</div>';
    h += '<div class="projbar"><div class="k">PROJECT:</div><div class="v">' + esc(m.project || '') + '</div></div>';
    h += '<div class="tablebanner">MAN-HOUR 산출 TABLE</div>';

    var cols = isOp1
      ? ['No.', 'Part', 'Activity', '단위', '수량', '난이도', '내부 Unit', '외주 Unit', '내부 M/H', '외주 M/H', 'Total M/H', 'Remarks']
      : ['No.', 'Part', 'Activity', '단위', '수량', '난이도', '내부 M/H', '외주 M/H', 'Total M/H', 'Remarks'];

    h += '<div class="scrollx"><table class="sheet" style="width:100%">';
    h += '<tr><th class="head-basic" colspan="6">기본 정보</th>';
    if (isOp1) {
      h += '<th class="head-unit" colspan="2">Unit M/H</th><th class="head-mh" colspan="2">M/H</th>' +
        '<th class="head-unit" rowspan="2">Total M/H</th><th class="head-rem" rowspan="2">Remarks</th></tr>';
    } else {
      h += '<th class="head-mh" colspan="3">M/H</th><th class="head-rem" rowspan="2">Remarks</th></tr>';
    }
    h += '<tr>';
    cols.forEach(function (name, i) {
      if (i >= cols.length - (isOp1 ? 2 : 1)) return; // rowspan 으로 이미 그린 열
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
        '<td class="ctr">' + esc(partLabel(r.part)) + '</td>' +
        '<td>' + esc(r.activity) + '</td>' +
        '<td class="ctr">' + esc(r.unit) + '</td>' +
        '<td class="num">' + money(r.qty, 1) + '</td>' +
        '<td class="ctr">' + esc(r.diff) + '</td>';
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
      ' &nbsp;|&nbsp; 설계기간: ' + fmt(m.months) + ' 개월 &nbsp;|&nbsp; 평균 투입인원: ' + f1(avg) + ' 명/month</div>';
    el.innerHTML = h;
  }

  /* ==================================================== 산출기준_CI / TEL */
  function renderStd(el, part) {
    var std = m.std[part];
    var keys = Object.keys(std);
    var isCi = part === 'ci';
    var diffs = isCi ? ['SPI', '상', '중', '하'] : ['상', '중', '하'];
    var title = isCi ? 'C&I설계 화공 FEED사업 M/H 산출기준 (Rev.3)' : '통신설계 화공 FEED사업 M/H 산출기준 (Rev.3)';
    var version = m.external_min === 'Yes' ? '외주최소화 Ver.' : '일반 Ver.';

    var h = '<div><span class="sheet-date">' + SHEET_DATE + '</span><div class="sheet-title">' + esc(title) + '</div></div>';
    h += '<div class="std-topbar">' +
      '<div class="k">산출기준 적용 Version\n(일반 vs 외주최소화 Ver.)</div>' +
      '<div class="v">' + esc(version) + '</div>' +
      '<div class="n"><b>Notes</b><br>' +
      '1. Project 표준 난이도, Unit M/H 난이도 및 SPI 적용 여부를 기준으로 M/H를 산정합니다.<br>' +
      '2. 노란색/초록색 숫자 셀은 직접 수정할 수 있으며 Summary/Output/OP에 즉시 반영됩니다.</div></div>';

    h += '<div class="scrollx"><table class="std" style="width:100%">';
    h += '<tr><th class="basic" rowspan="3">NO.</th><th class="basic" rowspan="3">Activity</th>' +
      '<th class="basic" rowspan="3">단위</th>' +
      '<th colspan="' + diffs.length + '">난이도별 Unit M/H</th>' +
      '<th colspan="' + diffs.length + '">난이도별 Unit M/H</th>' +
      '<th class="guide" rowspan="2">MAN-HOUR 산출 지침</th></tr>';
    h += '<tr><th colspan="' + diffs.length + '">내부 Unit M/H</th>' +
      '<th colspan="' + diffs.length + '">외주 Unit M/H</th></tr>';
    h += '<tr>' + diffs.map(function (d) { return '<th>' + esc(d) + '</th>'; }).join('') +
      diffs.map(function (d) { return '<th>' + esc(d) + '</th>'; }).join('') +
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
          if (isNaN(n)) { toast('숫자만 입력할 수 있습니다.'); inp.focus(); return; }
        }
        updateStdValue(inp.dataset.part, inp.dataset.key, inp.dataset.typ, inp.dataset.diff, n);
      });
      inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') inp.blur(); });
    });
  }

  /* 산출기준 Unit M/H 수정 → Output 행에 복사된 std 도 함께 갱신 (App.update_std_value) */
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
    h += '<h1>M/H 산출 가이드</h1><div class="sub">첨부 #3-1 &nbsp;|&nbsp; Web Application Guide</div>';
    h += '<h2>I. 각 시트(탭) 관련 안내</h2>';
    [['1. Summary 시트', '본 프로그램의 Master Control 입력값 기준으로 C&I + 통신 M/H Summary를 표시합니다.\nPart, Case, 외주최소화, 내부/외부 비율, Base M/H, 설계기간, 단가가 Summary와 Output에 공통 반영됩니다.\n외주 예가, Total M/H, Total M/M, 평균 투입인원 확인용입니다.'],
     ['2. 산출기준_CI / 산출기준_TEL 시트', '내부/외주 Unit M/H 및 난이도별 산출 기준을 확인하는 기준표입니다.\nActivity, 단위, 난이도, 내부/외부 Unit M/H, Guide 내용을 검토하는 용도입니다.\n노란색(내부) / 초록색(외주) 숫자 셀을 직접 수정하면 모든 결과가 즉시 다시 계산됩니다.'],
     ['3. Input 수정 시트', 'M/H 산출을 위한 주요 입력값과 Project Condition을 입력합니다.\n노란색은 필수 입력, 파란색은 Master Control 연동, 초록색은 자동결과, 보라색은 선택 입력입니다.\n선택형 항목은 드롭다운에서 선택할 수 있습니다.'],
     ['4. Output_CI / Output_TEL 시트', 'Part별 산출서 형식으로 Activity, Unit, Qty, 난이도, 내부/외부 M/H, Total M/H를 표시합니다.\n좌측 Master Control의 Case와 외주최소화 조건이 반영됩니다.'],
     ['5. OP1 / OP2 시트', 'OP1은 Activity별 내부/외부 Unit 및 M/H 상세 확인용입니다.\nOP2-단종 및 OP2-종합은 보고용 M/H 내역 확인용이며, 각각 해당 Case 로 고정 계산됩니다.'],
     ['6. Word Report 생성', '현재 화면의 Master Control, Summary, 주요 산출 정보를 기준으로 Word Report(.doc)를 내려받습니다.\n브라우저 다운로드 폴더에 저장되며 Microsoft Word 로 바로 열 수 있습니다.']
    ].forEach(function (x) { h += block(x[0], x[1]); });

    h += '<h2>II. 각 Case 관련 안내</h2>';
    [['1. 외주-단종 Case', '외부 수행분을 외주-단종 단가로 예가 산정합니다.\n내부 검토와 외주 수행을 구분하는 일반 FEED 산출에 사용합니다.'],
     ['2. 외주-종합 Case', '외부 수행분을 외주-종합 단가로 예가 산정합니다.\n외주 업체가 복수 업무를 통합 수행하는 조건 검토에 사용합니다.']
    ].forEach(function (x) { h += block(x[0], x[1]); });

    h += '<h2>III. 외주최소화 적용 안내</h2>';
    [['1. 외주최소화 Yes', '내부/외부 M/H 비율을 75/25 기준으로 적용합니다.\nSummary, Output, OP, Word Report에 동일하게 반영됩니다.', false],
     ['2. 외주최소화 No', '원본 산출기준의 내부/외주 Unit M/H 산정 결과를 기준으로 비율을 계산합니다.', false],
     ['3. 원본 기준과 외주최소화', '원본 기준은 원본 Unit M/H 구조를 유지하는 방식입니다.\n외주최소화를 적용하려면 사용자 입력 비율을 사용합니다.', true]
    ].forEach(function (x) { h += block(x[0], x[1], x[2]); });

    h += '<h2>IV. 사용 순서</h2>';
    h += block('1. 기본 정보 입력', '좌측 Master Control에서 Project, Part, Case, Base M/H, 설계기간, 외주 단가를 입력합니다.');
    h += block('2. Input 수정', 'Input 수정 탭에서 수량 및 선택 조건을 검토하고 필요한 값을 수정합니다.');
    h += block('3. Output 확인 및 Word Report 생성', 'Summary, Output, OP 탭에서 결과를 확인한 후 Word Report를 생성합니다.');

    h += '<h2>V. 입력값 저장 안내</h2>';
    h += block('1. 자동 저장', '입력값과 산출기준 수정값은 브라우저(localStorage)에 자동 저장되며, 다시 접속하면 그대로 복원됩니다.\n"저장값 불러오기"로 언제든 마지막 저장 상태를 다시 불러올 수 있습니다.');
    h += block('2. 파일 내보내기 / 가져오기', '"입력값 내보내기"는 JSON 파일로 저장합니다.\n데스크톱 프로그램의 FEED_MH_Calculator_Last_Input.json 과 같은 형식이므로 서로 주고받을 수 있습니다.');
    h += '</div>';
    el.innerHTML = h;
  }

  /* ============================================================ 상태 저장 */
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
      if (!silent) toast('현재 입력값을 브라우저에 저장했습니다.');
    } catch (e) {
      if (!silent) toast('저장 실패: ' + e.message);
    }
  }

  /* Python App.load_user_state() 와 동일한 복원 로직 */
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
      // Output 행 안에 복사되어 있는 std 도 같이 갱신
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
      if (showMsg) toast('저장된 입력값이 없습니다.');
      return false;
    }
    try {
      applyState(JSON.parse(raw));
      refreshAll();
      if (showMsg) toast('저장된 입력값을 불러왔습니다.');
      return true;
    } catch (e) {
      if (showMsg) toast('불러오기 실패: ' + e.message);
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
    toast('현재 입력값을 JSON 파일로 저장했습니다.');
  }

  function importState(file) {
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = JSON.parse(String(reader.result));
        applyState(data);
        refreshAll();
        saveState(true);
        toast('선택한 입력값 파일을 불러왔습니다.');
      } catch (e) {
        toast('가져오기 실패: ' + e.message);
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  function resetState() {
    if (!confirm('입력값을 초기화하시겠습니까?\n저장된 마지막 입력값도 삭제됩니다.')) return;
    try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
    m = new MH.Model();
    writeControls();
    refreshAll();
    toast('입력값을 초기화했습니다.');
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
      toast('Word Report(.doc)를 내려받았습니다.');
    } catch (e) {
      toast('Word Report 생성 실패: ' + e.message);
    }
  }

  /* ================================================================ 초기화 */
  function bindControls() {
    ['mc-part', 'mc-case', 'mc-ratio'].forEach(function (id) {
      $(id).addEventListener('change', applyLeft);
    });
    $('mc-external').addEventListener('change', function () {
      if ($('mc-external').value === 'Yes') {
        $('mc-ratio').value = '사용자 입력';
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
      $('mc-ratio').value = '사용자 입력';
      suppressRatioSync = false;
    });
    $('mc-ep').addEventListener('input', function () {
      if (suppressRatioSync) return;
      var v = parseFloat($('mc-ep').value);
      if (isNaN(v)) return;
      suppressRatioSync = true;
      $('mc-ip').value = f1(100 - v);
      $('mc-ratio').value = '사용자 입력';
      suppressRatioSync = false;
    });
    $('mc-ip').addEventListener('change', applyLeft);
    $('mc-ep').addEventListener('change', applyLeft);

    $('btn-recalc').addEventListener('click', function () { applyLeft(); toast('Output을 다시 계산했습니다.'); });
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
