/* A very small Excel evaluator, and a recomputing cell graph.
 *
 * Port of tools/xlformula.py. The calculator recomputes the source workbook
 * rather than re-implementing it, so this has to agree with Excel exactly -
 * including the two details that decided fidelity in the Python version:
 * an empty cell reads as 0, and Excel carries 15 significant decimal digits,
 * so 85*(1-0.3)*3 rounds to 179 and not 178. */
(function (global) {
  'use strict';

  /* ------------------------------------------------------------- numbers */

  // Excel carries 15 significant decimal digits, not a double's 17.
  function xl15(x) {
    if (!isFinite(x)) return x;
    return Number(x.toPrecision(15));
  }

  // Excel ROUND - half away from zero, applied to the 15-digit value.
  function xlround(x, nd) {
    nd = nd || 0;
    if (!isFinite(x)) return x;
    var f = Math.pow(10, nd);
    var y = xl15(xl15(x) * f);
    var r = y >= 0 ? Math.floor(y + 0.5) : -Math.floor(-y + 0.5);
    return r / f;
  }

  function toNum(v) {
    if (v === true) return 1;
    if (v === false || v === null || v === undefined || v === '') return 0;
    return typeof v === 'number' ? v : (isNaN(parseFloat(v)) ? 0 : parseFloat(v));
  }

  /* ------------------------------------------------------------ tokenizer */

  var TOKEN = new RegExp(
    '(\\s+)' +                                                    // 1 ws
    '|(\\d+(?:\\.\\d+)?(?:[eE][-+]?\\d+)?)' +                     // 2 number
    '|("(?:[^"]|"")*")' +                                         // 3 string
    "|((?:'[^']+'|[A-Za-z_\\u3131-\\uD79D][\\w\\u3131-\\uD79D.]*)!\\$?[A-Z]{1,3}\\$?\\d+" +
    '|\\$?[A-Z]{1,3}\\$?\\d+)' +                                  // 4 reference
    '|([A-Za-z_\\u3131-\\uD79D][\\w\\u3131-\\uD79D.]*)' +         // 5 name
    '|(<>|<=|>=|[-+*/^&<>=(),:%])',                               // 6 operator
    'g');

  function tokenize(src) {
    var out = [], m;
    TOKEN.lastIndex = 0;
    var pos = 0;
    while (pos < src.length) {
      TOKEN.lastIndex = pos;
      m = TOKEN.exec(src);
      if (!m || m.index !== pos) throw new Error('cannot tokenize: ' + src.slice(pos, pos + 20));
      pos = TOKEN.lastIndex;
      if (m[1] !== undefined) continue;                       // whitespace
      if (m[2] !== undefined) out.push(['num', m[2]]);
      else if (m[3] !== undefined) out.push(['str', m[3]]);
      else if (m[4] !== undefined) out.push(['ref', m[4]]);
      else if (m[5] !== undefined) out.push(['name', m[5]]);
      else out.push(['op', m[6]]);
    }
    return out;
  }

  /* --------------------------------------------------------------- parser */

  function Parser(tokens) { this.t = tokens; this.i = 0; }
  Parser.prototype.peek = function () { return this.i < this.t.length ? this.t[this.i] : [null, null]; };
  Parser.prototype.take = function (val) {
    var tk = this.peek();
    if (val !== undefined && tk[1] !== val) throw new Error('expected ' + val + ' got ' + tk[1]);
    this.i += 1;
    return tk;
  };
  Parser.prototype.parse = function () {
    var n = this.comparison();
    if (this.i !== this.t.length) throw new Error('trailing tokens');
    return n;
  };
  Parser.prototype.comparison = function () {
    var n = this.additive(), v;
    while (['=', '<', '>', '<=', '>=', '<>'].indexOf(this.peek()[1]) >= 0) {
      v = this.take()[1];
      n = ['cmp', v, n, this.additive()];
    }
    return n;
  };
  Parser.prototype.additive = function () {
    var n = this.multiplicative(), v;
    while (['+', '-', '&'].indexOf(this.peek()[1]) >= 0) {
      v = this.take()[1];
      n = ['bin', v, n, this.multiplicative()];
    }
    return n;
  };
  Parser.prototype.multiplicative = function () {
    var n = this.unary(), v;
    while (['*', '/'].indexOf(this.peek()[1]) >= 0) {
      v = this.take()[1];
      n = ['bin', v, n, this.unary()];
    }
    return n;
  };
  Parser.prototype.unary = function () {
    if (this.peek()[1] === '-') { this.take(); return ['neg', this.unary()]; }
    if (this.peek()[1] === '+') { this.take(); return this.unary(); }
    return this.power();
  };
  Parser.prototype.power = function () {
    var n = this.primary();
    while (this.peek()[1] === '^') { this.take(); n = ['bin', '^', n, this.unary()]; }
    if (this.peek()[1] === '%') { this.take(); n = ['pct', n]; }
    return n;
  };
  // Excel allows an omitted argument - ROUNDUP(x,) - and reads it as 0.
  Parser.prototype.argument = function () {
    var nxt = this.peek()[1];
    if (nxt === ',' || nxt === ')') return ['num', 0];
    return this.comparison();
  };

  Parser.prototype.primary = function () {
    var tk = this.peek(), kind = tk[0], val = tk[1], args, end;
    if (kind === 'num') { this.take(); return ['num', parseFloat(val)]; }
    if (kind === 'str') { this.take(); return ['str', val.slice(1, -1).replace(/""/g, '"')]; }
    if (val === '(') { this.take('('); var n = this.comparison(); this.take(')'); return n; }
    if (kind === 'ref') {
      this.take();
      if (this.peek()[1] === ':') { this.take(':'); end = this.take()[1]; return ['range', val, end]; }
      return ['ref', val];
    }
    if (kind === 'name') {
      this.take();
      if (this.peek()[1] === '(') {
        this.take('(');
        args = [];
        if (this.peek()[1] !== ')') {
          args.push(this.argument());
          while (this.peek()[1] === ',') { this.take(','); args.push(this.argument()); }
        }
        this.take(')');
        return ['call', val.toUpperCase(), args];
      }
      return ['name', val];        // a bare name, e.g. the workbook's "fault"
    }
    throw new Error('unexpected token ' + val);
  };

  /* ------------------------------------------------------------ reference */

  var CELL = /^(?:('[^']+'|[^!]+)!)?\$?([A-Z]{1,3})\$?(\d+)$/;

  function splitRef(ref) {
    var m = CELL.exec(ref);
    if (!m) throw new Error('bad reference ' + ref);
    var sheet = m[1];
    if (sheet && sheet.charAt(0) === "'") sheet = sheet.slice(1, -1);
    return { sheet: sheet || null, col: m[2], row: parseInt(m[3], 10) };
  }

  function colToI(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function iToCol(i) {
    var s = '', r;
    while (i > 0) { r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = (i - 1 - r) / 26; }
    return s;
  }

  function expandRange(start, end) {
    var a = splitRef(start), b = splitRef(end), out = [], c, r;
    for (c = colToI(a.col); c <= colToI(b.col); c++) {
      for (r = a.row; r <= b.row; r++) {
        out.push((a.sheet ? a.sheet + '!' : '') + iToCol(c) + r);
      }
    }
    return out;
  }

  /* ------------------------------------------------------------ evaluator */

  function evaluate(node, resolve) {
    var k = node[0], i, v;
    if (k === 'num' || k === 'str') return node[1];
    if (k === 'name') throw new Error('unknown name ' + node[1]);
    if (k === 'ref') { v = resolve(node[1]); return (v === null || v === undefined) ? 0 : v; }
    if (k === 'range') {
      var refs = expandRange(node[1], node[2]), out = [];
      for (i = 0; i < refs.length; i++) out.push(resolve(refs[i]));
      return out;
    }
    if (k === 'neg') return -toNum(evaluate(node[1], resolve));
    if (k === 'pct') return toNum(evaluate(node[1], resolve)) / 100;
    if (k === 'cmp') {
      var av = evaluate(node[2], resolve), bv = evaluate(node[3], resolve);
      if (typeof av === 'string' || typeof bv === 'string') {
        if (av === null || av === undefined) av = '';
        if (bv === null || bv === undefined) bv = '';
        if (!(typeof av === 'string' && typeof bv === 'string')) { av = String(av); bv = String(bv); }
        // Excel compares text case-insensitively: "Yes" = "yes" is TRUE.
        av = av.toUpperCase(); bv = bv.toUpperCase();
      } else { av = toNum(av); bv = toNum(bv); }
      switch (node[1]) {
        case '=': return av === bv;
        case '<>': return av !== bv;
        case '<': return av < bv;
        case '>': return av > bv;
        case '<=': return av <= bv;
        case '>=': return av >= bv;
      }
    }
    if (k === 'bin') {
      if (node[1] === '&') return String(evaluate(node[2], resolve)) + String(evaluate(node[3], resolve));
      var x = toNum(evaluate(node[2], resolve)), y = toNum(evaluate(node[3], resolve));
      switch (node[1]) {
        case '+': return x + y;
        case '-': return x - y;
        case '*': return x * y;
        case '/': return y ? x / y : 0;
        case '^': return Math.pow(x, y);
      }
    }
    if (k === 'call') {
      var name = node[1], args = node[2];
      if (name === 'IF') {
        var cond = evaluate(args[0], resolve);
        var truthy = (typeof cond === 'string') ? cond !== '' : (cond === true || toNum(cond) !== 0);
        if (truthy) return evaluate(args[1], resolve);
        return args.length > 2 ? evaluate(args[2], resolve) : 0;
      }
      if (name === 'IFERROR') {
        try { return evaluate(args[0], resolve); } catch (e) { return evaluate(args[1], resolve); }
      }
      var vals = [];
      for (i = 0; i < args.length; i++) {
        v = evaluate(args[i], resolve);
        if (Object.prototype.toString.call(v) === '[object Array]') vals = vals.concat(v);
        else vals.push(v);
      }
      var nums = vals.map(toNum);
      switch (name) {
        case 'SUM': return nums.reduce(function (a, b) { return a + b; }, 0);
        case 'COUNT': return vals.filter(function (x) { return typeof x === 'number'; }).length;
        case 'COUNTA': return vals.filter(function (x) { return x !== null && x !== undefined && x !== ''; }).length;
        case 'MIN': return nums.length ? Math.min.apply(null, nums) : 0;
        case 'MAX': return nums.length ? Math.max.apply(null, nums) : 0;
        case 'ROUND': return xlround(nums[0], nums.length > 1 ? nums[1] : 0);
        case 'ROUNDUP': case 'ROUNDDOWN': {
          var nd = nums.length > 1 ? nums[1] : 0, f = Math.pow(10, nd), t = xl15(nums[0] * f);
          return (name === 'ROUNDUP' ? Math.ceil(t - 1e-9) : Math.floor(t + 1e-9)) / f;
        }
      }
      throw new Error('unsupported function ' + name);
    }
    throw new Error('bad node ' + k);
  }

  var cache = {};
  function compile(text) {
    if (cache[text]) return cache[text];
    var body = text.charAt(0) === '=' ? text.slice(1) : text;
    return (cache[text] = new Parser(tokenize(body)).parse());
  }

  /* ---------------------------------------------------------------- graph */

  /* Cells are {f: formula} or {v: literal}. Formula cells recompute on demand
   * and memoise until something changes; a cell caught in a cycle falls back to
   * its own literal so a bad reference cannot hang the page. */
  function Graph(cells) {
    this.cells = cells;
    this.memo = {};
    this.busy = {};
    this.errors = {};
  }

  Graph.prototype.invalidate = function () { this.memo = {}; this.errors = {}; };

  Graph.prototype.set = function (ref, value) {
    var c = this.cells[ref];
    if (c && c.f) throw new Error('refusing to overwrite a formula at ' + ref);
    this.cells[ref] = { v: value };
    this.invalidate();
  };

  Graph.prototype.get = function (ref, contextSheet) {
    var full = ref.indexOf('!') >= 0 ? ref : (contextSheet ? contextSheet + '!' + ref : ref);
    full = full.replace(/\$/g, '');
    if (Object.prototype.hasOwnProperty.call(this.memo, full)) return this.memo[full];
    var cell = this.cells[full];
    if (!cell) return 0;
    if (cell.v !== undefined) return cell.v;
    if (this.busy[full]) return cell.v === undefined ? 0 : cell.v;
    this.busy[full] = true;
    var sheet = full.split('!')[0];
    var self = this;
    var out;
    try {
      out = evaluate(compile(cell.f), function (r) { return self.get(r, sheet); });
    } catch (e) {
      this.errors[full] = String(e.message || e);
      out = 0;
    } finally {
      delete this.busy[full];
    }
    this.memo[full] = out;
    return out;
  };

  var api = { xl15: xl15, xlround: xlround, toNum: toNum, tokenize: tokenize,
              compile: compile, evaluate: evaluate, splitRef: splitRef,
              expandRange: expandRange, Graph: Graph };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.XL = api;
})(typeof window !== 'undefined' ? window : this);
