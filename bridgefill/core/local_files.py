"""依程式位置定位攜帶式參考資料與日誌，不依賴目前工作目錄。"""
from pathlib import Path
import sys

APP_DIR = Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else Path(__file__).resolve().parents[2]
RESOURCE_DIR = Path(getattr(sys,'_MEIPASS',APP_DIR)).resolve()
REFERENCE_DIR = APP_DIR / 'data'
LOG_DIR = APP_DIR / 'log'
TEMP_DIR = APP_DIR / 'temp'
DEFAULT_REPAIR_NAME = '劣化說明與維修單價1116.xlsx'


def repair_files(directory=REFERENCE_DIR):
    if not directory.is_dir():
        return []
    return sorted((p for p in directory.iterdir() if p.is_file() and
                   p.suffix.lower() in ('.xlsx', '.xlsm') and not p.name.startswith('~$')),
                  key=lambda p: (p.name != DEFAULT_REPAIR_NAME, p.name.casefold()))


def ensure_directories():
    for name in ('data','temp','document','log','backup'):(APP_DIR/name).mkdir(parents=True,exist_ok=True)


def stored_path(value):
    path=Path(value).resolve()
    try:return '@app/'+path.relative_to(APP_DIR.resolve()).as_posix()
    except ValueError:return str(path)


def restored_path(value):
    value=str(value)
    if not value.startswith('@app/'):return Path(value)
    relative=Path(value[5:])
    if relative.is_absolute() or '..' in relative.parts:raise ValueError('可攜路徑格式不合法')
    return APP_DIR/relative
