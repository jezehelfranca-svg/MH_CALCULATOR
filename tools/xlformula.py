# -*- coding: utf-8 -*-
"""A very small Excel formula evaluator.

Only what the 실행사업 workbook actually uses: IF, ROUND, ROUNDUP, ROUNDDOWN,
SUM, MIN, MAX, arithmetic, comparisons, cell references and ranges. That is
enough to reproduce every OP1 column, including the bespoke tiered rows that
do not follow unit x qty.

Values are resolved through a caller-supplied function so the same evaluator
can read a workbook, or the calculator's own state.
"""
import re
from decimal import Context, Decimal, ROUND_HALF_UP

TOKEN = re.compile(r"""
    (?P<ws>\s+)
  | (?P<num>\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<ref>(?:'[^']+'|[A-Za-z_가-힣][\w가-힣.]*)!\$?[A-Z]{1,3}\$?\d+
            |\$?[A-Z]{1,3}\$?\d+)
  | (?P<name>[A-Za-z_가-힣][\w가-힣.]*)
  | (?P<op><>|<=|>=|[-+*/^&<>=(),:%])
""", re.X)


class Err(Exception):
    pass


# Excel carries 15 significant decimal digits, not the full 17 of a double.
# A product like 85*(1-0.3)*3 lands on 178.49999999999997 in IEEE arithmetic but
# Excel sees 178.500000000000 and rounds it to 179. Normalising to 15 digits
# before rounding reproduces that.
_XL15 = Context(prec=15)


def xl15(x):
    """Value as Excel carries it - 15 significant decimal digits."""
    return float(+_XL15.create_decimal(Decimal(repr(float(x)))))


def xlround(x, nd=0):
    """Excel ROUND - half away from zero, on the 15-digit value."""
    q = Decimal(1).scaleb(-int(nd))
    d = +_XL15.create_decimal(Decimal(repr(float(x))))
    return float(d.quantize(q, rounding=ROUND_HALF_UP))


def tokenize(src):
    pos, out = 0, []
    while pos < len(src):
        m = TOKEN.match(src, pos)
        if not m:
            raise Err('cannot tokenize at %r' % src[pos:pos + 20])
        pos = m.end()
        kind = m.lastgroup
        if kind != 'ws':
            out.append((kind, m.group()))
    return out


class Parser:
    def __init__(self, tokens):
        self.t, self.i = tokens, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self, val=None):
        k, v = self.peek()
        if val is not None and v != val:
            raise Err('expected %r, got %r' % (val, v))
        self.i += 1
        return (k, v)

    def parse(self):
        node = self.comparison()
        if self.i != len(self.t):
            raise Err('trailing tokens: %r' % (self.t[self.i:],))
        return node

    def comparison(self):
        node = self.additive()
        while self.peek()[1] in ('=', '<', '>', '<=', '>=', '<>'):
            op = self.take()[1]
            node = ('cmp', op, node, self.additive())
        return node

    def additive(self):
        node = self.multiplicative()
        while self.peek()[1] in ('+', '-', '&'):
            op = self.take()[1]
            node = ('bin', op, node, self.multiplicative())
        return node

    def multiplicative(self):
        node = self.unary()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            node = ('bin', op, node, self.unary())
        return node

    def unary(self):
        if self.peek()[1] == '-':
            self.take()
            return ('neg', self.unary())
        if self.peek()[1] == '+':
            self.take()
            return self.unary()
        return self.power()

    def power(self):
        node = self.primary()
        while self.peek()[1] == '^':
            self.take()
            node = ('bin', '^', node, self.unary())
        if self.peek()[1] == '%':
            self.take()
            node = ('pct', node)
        return node

    def argument(self):
        # Excel allows an omitted argument - ROUNDUP(x,) - and reads it as 0.
        if self.peek()[1] in (',', ')'):
            return ('num', 0.0)
        return self.comparison()

    def primary(self):
        kind, val = self.peek()
        if kind == 'num':
            self.take()
            return ('num', float(val))
        if kind == 'str':
            self.take()
            return ('str', val[1:-1].replace('""', '"'))
        if val == '(':
            self.take('(')
            node = self.comparison()
            self.take(')')
            return node
        if kind == 'ref':
            self.take()
            if self.peek()[1] == ':':
                self.take(':')
                end = self.take()[1]
                return ('range', val, end)
            return ('ref', val)
        if kind == 'name':
            self.take()
            if self.peek()[1] == '(':
                self.take('(')
                args = []
                if self.peek()[1] != ')':
                    args.append(self.argument())
                    while self.peek()[1] == ',':
                        self.take(',')
                        args.append(self.argument())
                self.take(')')
                return ('call', val.upper(), args)
            return ('name', val)          # a bare name - e.g. the workbook's "fault"
        raise Err('unexpected token %r' % (val,))


def _num(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return 0.0


def evaluate(node, resolve):
    """resolve(ref_text) -> value; for a range, resolve is called per cell."""
    k = node[0]
    if k == 'num':
        return node[1]
    if k == 'str':
        return node[1]
    if k == 'name':
        raise Err('unknown name %r' % node[1])
    if k == 'ref':
        v = resolve(node[1])
        # Excel reads an empty cell as 0.
        return 0.0 if v is None else v
    if k == 'range':
        return [resolve(r) for r in expand_range(node[1], node[2])]
    if k == 'neg':
        return -_num(evaluate(node[1], resolve))
    if k == 'pct':
        return _num(evaluate(node[1], resolve)) / 100.0
    if k == 'cmp':
        _, op, a, b = node
        av, bv = evaluate(a, resolve), evaluate(b, resolve)
        if isinstance(av, str) or isinstance(bv, str):
            av = '' if av is None else av
            bv = '' if bv is None else bv
            if not (isinstance(av, str) and isinstance(bv, str)):
                av, bv = str(av), str(bv)
            # Excel compares text case-insensitively: "Yes" = "yes" is TRUE.
            av, bv = av.upper(), bv.upper()
        else:
            av, bv = _num(av), _num(bv)
        return {'=': av == bv, '<>': av != bv, '<': av < bv,
                '>': av > bv, '<=': av <= bv, '>=': av >= bv}[op]
    if k == 'bin':
        _, op, a, b = node
        if op == '&':
            return '%s%s' % (evaluate(a, resolve), evaluate(b, resolve))
        av, bv = _num(evaluate(a, resolve)), _num(evaluate(b, resolve))
        if op == '+':
            return av + bv
        if op == '-':
            return av - bv
        if op == '*':
            return av * bv
        if op == '/':
            return av / bv if bv else 0.0
        if op == '^':
            return av ** bv
    if k == 'call':
        _, name, args = node
        if name == 'IF':
            cond = evaluate(args[0], resolve)
            if isinstance(cond, str):
                cond = bool(cond)
            branch = args[1] if _num(cond) or cond is True else (args[2] if len(args) > 2 else ('num', 0.0))
            return evaluate(branch, resolve)
        if name == 'IFERROR':
            try:
                return evaluate(args[0], resolve)
            except Exception:
                return evaluate(args[1], resolve)
        vals = []
        for a in args:
            v = evaluate(a, resolve)
            vals.extend(v if isinstance(v, list) else [v])
        nums = [_num(v) for v in vals]
        if name == 'SUM':
            return sum(nums)
        if name == 'COUNT':
            return float(sum(1 for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)))
        if name == 'COUNTA':
            return float(sum(1 for v in vals if v is not None and v != ''))
        if name == 'MIN':
            return min(nums) if nums else 0.0
        if name == 'MAX':
            return max(nums) if nums else 0.0
        if name == 'ROUND':
            return xlround(nums[0], nums[1] if len(nums) > 1 else 0)
        if name in ('ROUNDUP', 'ROUNDDOWN'):
            import math
            f = 10 ** int(nums[1] if len(nums) > 1 else 0)
            x = nums[0] * f
            r = math.ceil(x - 1e-9) if name == 'ROUNDUP' else math.floor(x + 1e-9)
            return r / f
        raise Err('unsupported function %s' % name)
    raise Err('bad node %r' % (k,))


CELL = re.compile(r"^(?:(?P<sheet>'[^']+'|[^!]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)$")


def split_ref(ref):
    m = CELL.match(ref)
    if not m:
        raise Err('bad reference %r' % ref)
    sheet = m.group('sheet')
    if sheet and sheet.startswith("'"):
        sheet = sheet[1:-1]
    return sheet, m.group('col'), int(m.group('row'))


def col_to_i(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def i_to_col(i):
    s = ''
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def expand_range(start, end):
    s_sheet, s_col, s_row = split_ref(start)
    _, e_col, e_row = split_ref(end)
    out = []
    for c in range(col_to_i(s_col), col_to_i(e_col) + 1):
        for r in range(s_row, e_row + 1):
            ref = '%s%d' % (i_to_col(c), r)
            out.append('%s!%s' % (s_sheet, ref) if s_sheet else ref)
    return out


def compile_formula(text):
    if text.startswith('='):
        text = text[1:]
    return Parser(tokenize(text)).parse()
