"""依橋梁總表歸類前期與基本資料（2026-11-28 使用者需求）。

舊版的前期掃描檔放在 temp\\scan，或 scan\\縣市政府\\pre-brigdedata 但沒有依工作表分資料夾（例：「未分縣市」）。
依總表每座橋（工作表、橋名、二代系統ID）把檔案搬到：
- 前期：scan\\{縣市政府}\\pre-brigdedata\\{工作表}\\{日期}-{橋名}-前期.json
- 基本資料：scan\\{縣市政府}\\pre-brigdeFEdata\\{工作表}\\{ID}-{橋名}-基本資料.json
縣市政府取自檔案內的管理機關／所在地（基本資料取所在縣市），讀不到時用 fallback。
同名橋以二代系統ID 核對（檔案記錄的 ID 與總表不同不搬）；目標已有同名檔：內容相同刪掉多的那份，不同則保留並列出。
本模組不依賴 Qt。
"""
import filecmp
import json
import re
import shutil
from pathlib import Path

from .bridge_photos import government

PRIOR_FOLDER, BASIC_FOLDER = 'pre-brigdedata', 'pre-brigdeFEdata'


def safe(name):
    return ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in str(name or ''))


def read(path):
    try:
        return json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}


def prior_identity(data):
    """前期掃描檔記錄的 (二代系統ID, 縣市政府)。"""
    header = data.get('header') or {}
    ident = next((str(r.get('keys', {}).get('inventory_id') or '') for r in data.get('rows') or []
                  if r.get('keys', {}).get('inventory_id')), '')
    return ident, government(header.get('管理機關') or header.get('所在地') or '')


def candidates(roots, pattern):
    found = []
    for root in roots:
        root = Path(root)
        if root.is_dir():
            found += [p for p in root.rglob(pattern) if p.is_file()]
    return found


def bridge_government(name, ident, store_dir, scan_dir):
    """前期檔沒有管理機關時，找這座橋的縣市政府：基本資料（所在縣市）→ 本期掃描檔（查詢結果的管理機關、所在地）。"""
    if ident:
        for path in candidates([store_dir, scan_dir], f'{safe(ident)}-*-基本資料.json'):
            gov = government(read(path).get('county') or '')
            if gov:
                return gov
    for path in candidates([scan_dir], f'*-{safe(name)}-檢測.json'):
        header = read(path).get('header') or {}
        gov = government(header.get('管理機關') or header.get('所在地') or '')
        if gov:
            return gov
    return ''


def plan(bridges, store_dir, scan_dir, fallback=''):
    """要搬移的檔案 [(來源, 目標)] 與略過的說明。bridges：[{sheet, name, inventory_id}]（load_list_bridges）。"""
    store_dir, scan_dir = Path(store_dir), Path(scan_dir)
    moves, notes, claimed = [], [], set()
    for bridge in bridges:
        name, sheet, ident = bridge.get('name', ''), bridge.get('sheet', ''), str(bridge.get('inventory_id') or '')
        if not name or not sheet:
            continue
        pattern = f'*-{safe(name)}-前期.json'
        for path in candidates([scan_dir, store_dir], pattern):
            if path.name.startswith('0000-00-00') or path in claimed:
                continue                                    # 第一次檢測的空白前期：只供本次操作，留在 temp
            saved_id, gov = prior_identity(read(path))
            if ident and saved_id and saved_id != ident:
                notes.append(f'{path.name}：檔案的二代系統ID {saved_id} 與總表「{sheet} {name}」的 {ident} 不同（同名的其他橋），未搬移')
                continue
            gov = gov or bridge_government(name, ident, store_dir, scan_dir) or fallback or '未分縣市'
            target = store_dir / safe(gov) / PRIOR_FOLDER / safe(sheet) / path.name
            claimed.add(path)
            if path != target:
                moves.append((path, target))
        for path in candidates([scan_dir, store_dir], f'*-{safe(name)}-基本資料.json'):   # 本期掃描的基本資料（日期開頭）
            if path in claimed or not re.match(r'\d{4}-\d{2}-\d{2}-', path.name):
                continue
            values = read(path).get('values') or {}
            gov = government(values.get('county') or '') or bridge_government(name, ident, store_dir, scan_dir) \
                or fallback or '未分縣市'
            target = store_dir / safe(gov) / BASIC_FOLDER / safe(sheet) / path.name
            claimed.add(path)
            if path != target:
                moves.append((path, target))
        if ident:
            for path in candidates([store_dir, scan_dir], f'{safe(ident)}-*-基本資料.json'):
                if path in claimed:
                    continue
                data = read(path)
                gov = government(data.get('county') or '') or fallback or '未分縣市'
                target = store_dir / safe(gov) / BASIC_FOLDER / safe(sheet) / path.name
                claimed.add(path)
                if path != target:
                    moves.append((path, target))
    return moves, notes


def apply(moves):
    """搬移檔案。回傳 (已搬移數, 說明)。目標已有：內容相同刪來源，不同保留兩份並說明。"""
    moved, notes = 0, []
    for source, target in moves:
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            if filecmp.cmp(source, target, shallow=False):
                source.unlink()
                moved += 1
            else:
                notes.append(f'{source.name}：目標資料夾已有內容不同的同名檔，未搬移（{source.parent}）')
            continue
        shutil.move(str(source), str(target))
        moved += 1
    for folder in {s.parent for s, _ in moves}:              # 搬空的舊資料夾（不含 temp\scan 本身）順手刪除
        try:
            if folder.is_dir() and not any(folder.iterdir()) and re.search(PRIOR_FOLDER + '|' + BASIC_FOLDER, str(folder)):
                folder.rmdir()
        except OSError:
            pass
    return moved, notes
