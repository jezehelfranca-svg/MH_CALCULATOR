/* FEED M/H Calculator - calculation model
 * A direct port of the Model class in FEED_MH_Calculator_V0_4_OPExcelGrid.py.
 * Python's rounding rules (banker's rounding) are reproduced as well, so the
 * figures shown here are identical to the desktop program's. */
(function (global) {
  'use strict';

  var P = { ci: 'C&I (Instrumentation)', tel: 'TEL (Telecom)', both: 'All' };
  var REV = { 'All': 'both', 'C&I (Instrumentation)': 'ci', 'TEL (Telecom)': 'tel' };
  /* Compact form for the Part column, which repeats on every table row. */
  var P_SHORT = { ci: 'C&I', tel: 'TEL', both: 'All' };

  /* Difficulty grades are kept in Korean inside the data because they are the
   * lookup keys of every std entry's int/ext maps, shared with the desktop
   * program's saved-state file. They are translated only for display. */
  var DIFF_LABEL = {
    '상': 'High', '중': 'Medium', '하': 'Low',
    'SPI': 'SPI', 'SPI-내부': 'SPI-Internal', 'SPI-외부': 'SPI-External',
    '외주-종합': 'Outsourcing-Comprehensive', '외주-단종': 'Outsourcing-Single'
  };

  /* 'SPI-외부 / 외주-종합' -> 'SPI-External / Outsourcing-Comprehensive' */
  function diffLabel(v) {
    if (v === null || v === undefined || v === '') return '';
    return String(v).split(' / ').map(function (part) {
      var t = part.trim();
      return DIFF_LABEL[t] || t;
    }).join(' / ');
  }

  function num(v) {
    if (v === null || v === undefined || v === '' || v === '-') return 0.0;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0.0 : n;
  }

  /* Python's round()/format() break ties to even, but only on an exact .5.
   * A double x is exactly .5 at nd decimals iff x * 2^(nd+1) is an odd integer,
   * and multiplying by a power of two is error-free, so this test is exact.
   * Otherwise toFixed() rounds the true binary value correctly, matching Python. */
  function isTie(v, nd) {
    var t = v * Math.pow(2, nd + 1);
    return isFinite(t) && Number.isInteger(t) && Math.abs(t % 2) === 1;
  }

  // Fixed-point string with nd decimals, rounded the way Python's format() does
  function toFixedPy(x, nd) {
    if (!isFinite(x)) return String(x);
    var neg = x < 0;
    var v = Math.abs(x);
    var s;
    if (isTie(v, nd)) {
      var f = Math.pow(10, nd);
      var fl = Math.floor(v * f);
      var n = (fl % 2 === 0) ? fl : fl + 1;
      s = String(n);
      if (nd > 0) {
        while (s.length <= nd) s = '0' + s;
        s = s.slice(0, s.length - nd) + '.' + s.slice(s.length - nd);
      }
    } else {
      s = v.toFixed(nd);
    }
    return (neg && parseFloat(s) !== 0 ? '-' : '') + s;
  }

  // Equivalent to Python's round(x, nd)
  function pyRound(x, nd) {
    nd = nd || 0;
    if (!isFinite(x)) return x;
    return parseFloat(toFixedPy(x, nd));
  }

  function group(s) {
    var neg = s.charAt(0) === '-';
    if (neg) s = s.slice(1);
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + parts.join('.');
  }

  function money(v, d) {
    d = d === undefined ? 0 : d;
    var n = parseFloat(v);
    if (v === null || v === undefined || v === '' || isNaN(n)) {
      return String(v === null || v === undefined ? '' : v);
    }
    return group(toFixedPy(n, d));
  }

  function fmt(v) { return money(v, 0); }
  function f1(v) { return money(v, 1); }

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function Model() {
    this.inputs = clone(DATA.inputs);
    this.outputs = clone(DATA.outputs);
    this.std = clone(DATA.std);
    this.part = 'both';
    this.case_ = 'short';
    this.ratio_mode = 'original';
    this.external_min = 'No';
    this.project = '';
    this.base_mh = 161;
    this.months = 4;
    this.rate_short = 32000;
    this.rate_comp = 39000;
    this.internal_pct = 0;
    this.external_pct = 0;
    this.sync_common();
    this.recalc();
  }

  Model.prototype.find = function (p, c) {
    for (var i = 0; i < this.inputs.length; i++) {
      if (this.inputs[i].part === p && this.inputs[i].code === c) return this.inputs[i];
    }
    return null;
  };
  Model.prototype.val = function (p, c) {
    var x = this.find(p, c);
    return x ? (x.value === undefined ? '' : x.value) : '';
  };
  Model.prototype.setv = function (p, c, v) {
    var x = this.find(p, c);
    if (x) x.value = v;
  };
  Model.prototype.nval = function (p, c) { return num(this.val(p, c)); };

  Model.prototype.sync_common = function () {
    var self = this;
    ['ci', 'tel'].forEach(function (p) {
      self.setv(p, 'CTRL_PROJECT', self.project);
      self.setv(p, 'CTRL_EXTERNAL_MIN', self.external_min);
    });
    [['ci', 'C35_VALUE', this.base_mh], ['tel', 'C09_VALUE', this.base_mh],
     ['ci', 'C36_VALUE', this.months], ['tel', 'C10_VALUE', this.months],
     ['ci', 'C37_VALUE', this.rate_short], ['tel', 'C11_VALUE', this.rate_short],
     ['ci', 'C38_VALUE', this.rate_comp], ['tel', 'C12_VALUE', this.rate_comp]
    ].forEach(function (t) { self.setv(t[0], t[1], t[2]); });
  };

  Model.prototype.input_changed = function (x, v) {
    x.value = v;
    var c = x.code;
    if (c === 'CTRL_PROJECT') {
      this.project = v;
    } else if (c === 'CTRL_EXTERNAL_MIN') {
      this.external_min = v;
      if (v === 'Yes') { this.ratio_mode = 'custom'; this.internal_pct = 75; this.external_pct = 25; }
      else { this.ratio_mode = 'original'; }
    } else if (c === 'C35_VALUE' || c === 'C09_VALUE') { this.base_mh = num(v); }
    else if (c === 'C36_VALUE' || c === 'C10_VALUE') { this.months = num(v); }
    else if (c === 'C37_VALUE' || c === 'C11_VALUE') { this.rate_short = num(v); }
    else if (c === 'C38_VALUE' || c === 'C12_VALUE') { this.rate_comp = num(v); }
    this.sync_common();
    this.recalc();
  };

  /* Difficulty grades stay in Korean here on purpose - they are the keys of the
   * std int/ext tables (see DIFF_LABEL above). diffLabel() renders them. */
  Model.prototype.compute_diffs = function () {
    var self = this;
    var gp = this.val('ci', 'C01_SEL');
    var gd = gp === '1)' ? '상' : gp === '2)' ? '중' : gp === '3)' ? '하' : '중';
    this.setv('ci', 'C01_DIFF', gd);
    ['C04', 'C06', 'C08', 'C10', 'C12'].forEach(function (c) {
      var s = self.val('ci', c + '_SEL');
      self.setv('ci', c + '_DIFF', (s === '1)' && gd === '상') ? '상' : s === '1)' ? '중' : s === '2)' ? '하' : '중');
    });
    ['C14', 'C23'].forEach(function (code) {
      var s = self.val('ci', code + '_SEL');
      self.setv('ci', code + '_DIFF', s === '1)' ? 'SPI-내부' : s === '2)' ? 'SPI-외부' : gd);
    });
    ['C17', 'C18', 'C26', 'C27', 'C28', 'C32'].forEach(function (c) { self.setv('ci', c + '_DIFF', gd); });
    ['C19', 'C21'].forEach(function (c) {
      var s = self.val('ci', c + '_SEL');
      self.setv('ci', c + '_DIFF', (s === '1)' && gd === '상') ? '상' : s === '1)' ? '중' : s === '2)' ? '하' : '중');
    });
    var m = this.val('ci', 'C29_SEL');
    this.setv('ci', 'C29_DIFF', m === '1)' ? '상' : m === '2)' ? '중' : m === '3)' ? '하' : '중');
    this.setv('tel', 'C01_DIFF', this.val('tel', 'C01_SEL') === '1)' ? '중' : this.val('tel', 'C01_SEL') === '2)' ? '하' : '중');
    this.setv('tel', 'C03_DIFF', this.val('tel', 'C03_SEL') === '1)' ? '중' : this.val('tel', 'C03_SEL') === '2)' ? '하' : '중');
    this.setv('tel', 'C05_DIFF', this.val('tel', 'C05_SEL') === '1)' ? '상' : '중');
  };

  Model.prototype.update_qty = function () {
    var self = this;
    var ci1 = this.nval('ci', 'A01_QTY'), ci2 = this.nval('ci', 'A02_QTY'), ci3 = this.nval('ci', 'A03_QTY');
    var inline = ci3 > 0 ? ci3 : ci2 * 0.25;
    this.setv('ci', 'A15_QTY', pyRound(inline));
    this.setv('ci', 'A16_QTY', Math.max(0, pyRound(ci2 - inline)));
    ['A22', 'A23', 'A24'].forEach(function (c) { self.setv('ci', c + '_QTY', ci1 ? Math.ceil(ci1 / 500) : 0); });
    this.setv('ci', 'A26_QTY', ci2);
    this.setv('ci', 'A27_QTY', pyRound(ci2 / 5));
    this.setv('ci', 'A28_QTY', Math.min(100, ci2 ? Math.ceil(ci2 / 80) : 0));
    var sys = this.nval('tel', 'A01_QTY'), site = this.nval('tel', 'A02_QTY'), bld = this.nval('tel', 'A03_QTY');
    this.setv('tel', 'A09_QTY', sys);
    this.setv('tel', 'A10_QTY', sys);
    this.setv('tel', 'A11_QTY', pyRound(site / 90000 * 2 + bld, 2));
    this.setv('tel', 'A12_QTY', pyRound(site / 90000 + bld / 2, 2));
    this.setv('tel', 'A13_QTY', pyRound(site / 10000, 2));
    this.setv('tel', 'A14_QTY', sys);
    this.outputs.forEach(function (o) { o.qty = self.nval(o.part, o.code + '_QTY'); });
  };

  Model.prototype.diff = function (o) {
    var p = o.part, c = o.code;
    if (p === 'ci') {
      if (['A04', 'A05', 'A06'].indexOf(c) >= 0) return this.val('ci', 'C01_DIFF');
      var mp = { A07: 'C04', A08: 'C06', A09: 'C08', A10: 'C10', A11: 'C12', A15: 'C14', A16: 'C17', A21: 'C19', A25: 'C21', A26: 'C23', A34: 'C29' };
      if (mp[c]) return this.val('ci', mp[c] + '_DIFF');
      if (['A17', 'A18', 'A19', 'A20'].indexOf(c) >= 0) return this.val('ci', 'C18_DIFF');
      if (['A22', 'A23', 'A24', 'A27', 'A28', 'A29', 'A30', 'A31'].indexOf(c) >= 0) return this.val('ci', 'C26_DIFF');
      if (['A32', 'A33'].indexOf(c) >= 0) return this.val('ci', 'C27_DIFF');
      if (['A35', 'A36'].indexOf(c) >= 0) return this.val('ci', 'C32_DIFF');
      return this.val('ci', 'C01_DIFF');
    }
    if (['A04', 'A05', 'A06'].indexOf(c) >= 0) return this.val('tel', 'C01_DIFF');
    if (['A09', 'A10'].indexOf(c) >= 0) return this.val('tel', 'C03_DIFF');
    if (c === 'A12') return this.val('tel', 'C05_DIFF');
    return '중';
  };

  Model.prototype.units = function (o) {
    var d = this.diff(o) || '중';
    var st = o.std;
    if (st) {
      var lk = d.indexOf('SPI') >= 0 ? 'SPI' : d;
      var iu = (st.int && st.int[lk]) || 0;
      var eu = (st.ext && st.ext[lk]) || 0;
      if (iu || eu) return { iu: iu, eu: eu, d: d };
    }
    return { iu: o.baseIntUnit || 0, eu: o.baseExtUnit || 0, d: d };
  };

  Model.prototype.visible_outputs = function () {
    var self = this;
    if (this.part === 'both') return this.outputs;
    return this.outputs.filter(function (o) { return o.part === self.part; });
  };

  Model.prototype.raw_row = function (o) {
    var u = this.units(o);
    var q = num(o.qty);
    return {
      part: o.part, code: o.code, activity: o.activity, unit: o.unit, qty: q,
      diff: u.d, iu: u.iu, eu: u.eu,
      hec: u.iu * q, ext: u.eu * q, total: (u.iu + u.eu) * q
    };
  };

  Model.prototype.raw_sum = function () {
    var h = 0, e = 0, self = this;
    this.visible_outputs().forEach(function (o) {
      var r = self.raw_row(o); h += r.hec; e += r.ext;
    });
    return { hec: h, ext: e, total: h + e };
  };

  Model.prototype.pct = function () {
    if (this.ratio_mode === 'original') {
      var s = this.raw_sum();
      var ip = s.total ? s.hec / s.total * 100 : 0;
      return { ip: ip, ep: 100 - ip };
    }
    return { ip: this.internal_pct, ep: this.external_pct };
  };

  Model.prototype.row = function (o) {
    var r = this.raw_row(o);
    // Outsourcing Minimization / user-entered ratio wins: split Total M/H by the ratio
    if (this.ratio_mode === 'custom') {
      var p = this.pct();
      r.hec = r.total * p.ip / 100;
      r.ext = r.total * p.ep / 100;
      return r;
    }
    // Outsourcing-Comprehensive case: mirrors the Case3 logic of the original Excel OP1.
    // Case3 Internal Unit = internal Unit * (1 - GEC/Internal ratio - comprehensive-conversion ratio)
    // Case3 External Unit = external Unit * (1 - outsourcing-reduction ratio)
    //                       + internal Unit * comprehensive-conversion ratio / allocation divisor
    // GEC Unit = internal Unit * GEC/Internal ratio / allocation divisor
    // There is no separate GEC column on screen, so GEC M/H is folded into the
    // external / comprehensive share to keep the Total right.
    if (this.case_ === 'comp') {
      var q = r.qty, iu = r.iu, eu = r.eu;
      var gec_ratio, comp_ratio, ext_reduce, denom;
      if (r.part === 'ci') { gec_ratio = 0.30; comp_ratio = 0.10; ext_reduce = 0.0; denom = 0.70; }
      else { gec_ratio = 0.00; comp_ratio = 0.10; ext_reduce = 0.0; denom = 0.70; }
      var internal_unit = iu * (1 - gec_ratio - comp_ratio);
      var gec_unit = denom ? iu * gec_ratio / denom : 0;
      var comp_external_unit = denom ? eu * (1 - ext_reduce) + iu * comp_ratio / denom : eu;
      r.iu = internal_unit;
      r.eu = gec_unit + comp_external_unit;
      r.hec = pyRound(internal_unit * q);
      r.ext = pyRound((gec_unit + comp_external_unit) * q);
      r.total = r.hec + r.ext;
      r.diff = String(r.diff === undefined ? '' : r.diff) + ' / 외주-종합';
    }
    return r;
  };

  Model.prototype.summary = function () {
    var h = 0, e = 0, t = 0, self = this;
    this.visible_outputs().forEach(function (o) {
      var r = self.row(o); h += r.hec; e += r.ext; t += r.total;
    });
    return { hec: h, ext: e, total: t };
  };

  Model.prototype.recalc = function () {
    this.compute_diffs();
    this.update_qty();
    if (this.ratio_mode === 'original') {
      var s = this.raw_sum();
      this.internal_pct = s.total ? s.hec / s.total * 100 : 0;
      this.external_pct = 100 - this.internal_pct;
    }
  };

  /* Run fn with the Part filter temporarily switched, then restore it. */
  Model.prototype.withPart = function (part, fn) {
    var old = this.part;
    this.part = part;
    try { return fn(); } finally { this.part = old; }
  };

  Model.prototype.withCase = function (c, fn) {
    var old = this.case_;
    this.case_ = c;
    try { return fn(); } finally { this.case_ = old; }
  };

  /* M/H totals, M/M and average manpower for the given Part */
  Model.prototype.partTotals = function (part) {
    var self = this;
    return this.withPart(part, function () {
      var rows = self.visible_outputs().map(function (o) { return self.row(o); });
      var h = 0, e = 0, t = 0;
      rows.forEach(function (r) { h += r.hec; e += r.ext; t += r.total; });
      var mm = self.base_mh ? t / self.base_mh : 0;
      var avg = self.months ? mm / self.months : 0;
      return { rows: rows, internal: h, external: e, total: t, mm: mm, avg: avg };
    });
  };

  global.MH = {
    P: P, P_SHORT: P_SHORT, REV: REV, num: num, fmt: fmt, f1: f1, money: money, pyRound: pyRound,
    toFixedPy: toFixedPy, DIFF_LABEL: DIFF_LABEL, diffLabel: diffLabel, Model: Model
  };
})(window);
