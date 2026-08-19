# -*- coding: utf-8 -*-
import sys, subprocess
from pathlib import Path
APP=Path(__file__).resolve().parent/'FEED_MH_Calculator_V0_4_OPExcelGrid.py'
UTIL=Path(__file__).resolve().parent/'report_utils.py'
NAME='FEED_MH_Calculator_V0_4_OPExcelGrid'
def run(cmd):
    print('\n> '+' '.join(map(str,cmd)))
    subprocess.check_call(cmd)
if not APP.exists():
    input(f'앱 파일을 찾을 수 없습니다: {APP}\nEnter 키를 누르면 종료합니다...')
    raise SystemExit(1)
try:
    import PyInstaller  # noqa
except Exception:
    print('PyInstaller가 없어 자동 설치를 시도합니다.')
    run([sys.executable,'-m','pip','install','pyinstaller'])
cmd=[sys.executable,'-m','PyInstaller','--onefile','--windowed','--clean','--name',NAME,str(APP)]
ICON=Path(__file__).resolve().parent/'feed_mh_calculator_icon.ico'
if ICON.exists():
    cmd.extend(['--icon', str(ICON)])
if UTIL.exists():
    cmd.extend(['--add-data', f'{UTIL};.'])
run(cmd)
exe=Path(__file__).resolve().parent/'dist'/f'{NAME}.exe'
if exe.exists():
    print('\nEXE 생성 완료:')
    print(exe)
else:
    print('\nEXE 생성이 끝났지만 dist 폴더에서 EXE를 찾지 못했습니다.')
input('\nEnter 키를 누르면 종료합니다...')
