# -*- coding: utf-8 -*-
"""FEED_MH_Calculator_V0_4_OPExcelGrid.py 에 들어 있는 기준 데이터와 로고를
웹 앱이 쓰는 web/js/data.js 로 다시 뽑아냅니다.

    python3 tools/gen_web_data.py
"""
import ast
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'FEED_MH_Calculator_V0_4_OPExcelGrid.py'
DST = ROOT / 'web' / 'js' / 'data.js'


def main():
    lines = SRC.read_text(encoding='utf-8').split('\n')

    data_line = next(l for l in lines if l.startswith('DATA=json.loads('))
    data = json.loads(ast.literal_eval(data_line[len('DATA=json.loads('):-1]))

    logo_line = next(l for l in lines if l.startswith('HYUNDAI_LOGO_B64='))
    logo = ast.literal_eval(logo_line[len('HYUNDAI_LOGO_B64='):])

    out = [
        '/* Auto-generated from FEED_MH_Calculator_V0_4_OPExcelGrid.py - do not edit by hand. */',
        '/* Base data: Input items, Output activities and 산출기준 (standard Unit M/H) tables. */',
        'var DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';',
        '',
        "var HYUNDAI_LOGO_B64 = '" + logo + "';",
        '',
    ]
    DST.write_text('\n'.join(out), encoding='utf-8')
    print('%s 생성 완료 (inputs %d / outputs %d / std ci %d, tel %d)' % (
        DST, len(data['inputs']), len(data['outputs']),
        len(data['std']['ci']), len(data['std']['tel'])))


if __name__ == '__main__':
    main()
