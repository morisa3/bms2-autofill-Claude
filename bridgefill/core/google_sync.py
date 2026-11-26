"""Google 試算表同步（Apps Script Web App）與本機無損匯出標準 38 欄 Excel。

- 多位檢測員可同時在瀏覽器開著《115年北雲林自檢橋梁總表》編輯人工欄位；
- 程式透過 HTTPS JSON 呼叫試算表內建的 Apps Script（附金鑰與 LockService 排隊鎖），
  只精準更新該橋的程式欄位與附註識別碼，絕不鎖檔、不破壞版型、不覆寫人工修改。
- 需要交付正式 Excel 時，呼叫 export_to_xlsx() 由雲端抓取資料並以本機 38 欄標準版型產出 .xlsx。
本模組只使用 Python 標準函式庫與 openpyxl，不依賴 Qt。
"""
import json
import re
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from . import bridge_list as bl
from .list_template import (GUIDE_SHEET, add_formulas, add_guide, add_notes, fit, style_template, template_heads)


class GoogleSyncError(bl.ListError):
    """Google 試算表連線、權限或寫入錯誤。"""


def build_sync_payload(sources, sheet='', preserve_manual=True):
    """將 bridge_list.Sources 轉成傳給 Apps Script 的 JSON 內容。"""
    system, values, notes = bl.build(sources)
    header = sources.header or {}
    return {
        'action': 'sync_bridge',
        'sheet': sheet or '',
        'inventory_id': str(sources.inventory_id or '').strip(),
        'code': str(header.get('橋梁編號', '') or '').strip(),
        'name': str(header.get('橋梁名稱', '') or '').strip(),
        'preserve_manual': bool(preserve_manual),
        'only_basic': bool(sources.only_basic),
        'system': {k: '' if v is None else str(v) for k, v in system.items()},
        'values': {k: '' if v is None else str(v) for k, v in values.items()},
        'column_labels': dict(bl.COLUMNS),
        'column_keys': {k: bl.key(v) for k, v in bl.COLUMNS.items()},
        'ticks': list(bl.TICKS),
        'locked': list(bl.LOCKED),           # 只有前期資料加程式附註鎖定；本期資料每次以最新覆寫（2026-11-29）
        'notes': list(notes),
    }


class GoogleSheetClient:
    """與 Google Apps Script Web App 溝通的輕量用戶端。"""

    def __init__(self, url, token='', timeout=30):
        self.url = (url or '').strip()
        self.token = (token or '').strip()
        self.timeout = timeout

    def _call(self, payload):
        if not self.url:
            raise GoogleSyncError('尚未設定 Google 試算表 Apps Script 網址')
        if not re.match(r'^https?://', self.url, re.I):
            raise GoogleSyncError(f'Apps Script 網址格式不正確：{self.url}')
        body = json.dumps({**payload, 'token': self.token}, ensure_ascii=False).encode('utf-8')
        request = urllib.request.Request(
            self.url, data=body, method='POST',
            headers={'Content-Type': 'application/json; charset=utf-8', 'Accept': 'application/json'},
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                raw = response.read().decode('utf-8', errors='replace')
        except urllib.error.HTTPError as error:
            detail = error.read().decode('utf-8', errors='replace')[:300]
            raise GoogleSyncError(f'Google 試算表回應錯誤（HTTP {error.code}）：{detail or error.reason}') from error
        except urllib.error.URLError as error:
            raise GoogleSyncError(f'無法連線到 Google 試算表：{error.reason}') from error
        except TimeoutError as error:
            raise GoogleSyncError('連線 Google 試算表逾時，請稍後再試') from error
        try:
            data = json.loads(raw)
        except ValueError as error:
            snippet = raw.strip()[:160].replace('\n', ' ')
            raise GoogleSyncError(f'Google 試算表回傳格式異常（請確認 Web App 權限設為「所有人」）：{snippet}') from error
        if not isinstance(data, dict) or not data.get('ok'):
            raise GoogleSyncError(data.get('error') if isinstance(data, dict) else 'Google 試算表回傳未知錯誤')
        return data

    def ping(self):
        """測試連線，回傳 {'ok': True, 'title': 試算表名稱, 'sheets': [工作表名稱…]}。"""
        return self._call({'action': 'ping'})

    def list_bridges(self, sheet=''):
        """讀取總表中的橋梁清單（排程選橋、讀回下橋路徑用）。"""
        return self._call({'action': 'list_bridges', 'sheet': sheet or ''})

    def sync_bridge(self, sources, sheet='', preserve_manual=True):
        """將一座橋的資料精準寫入 Google 試算表對應列；回傳報告文字清單。"""
        payload = build_sync_payload(sources, sheet=sheet, preserve_manual=preserve_manual)
        result = self._call(payload)
        report = list(result.get('report') or [])
        for note in payload['notes']:
            if note not in report:
                report.append(note)
        return report or [f"{result.get('sheet', '')} 第 {result.get('row', '')} 列：已同步"]

    def export_all(self):
        """抓回全部工作表的表頭與列資料，供本機產生無損排版 .xlsx。"""
        return self._call({'action': 'export_all'})

    def export_to_xlsx(self, output_path, year=None):
        """從 Google 試算表讀取最新資料，套用 38 欄標準版型在本機產出不跑版 .xlsx。回傳摘要文字清單。"""
        data = self.export_all()
        return build_xlsx_from_cloud(data, output_path, year=year)


def build_xlsx_from_cloud(data, output_path, year=None):
    """將 export_all 回傳的 JSON 資料套用 38 欄標準版型產出 .xlsx（純本機邏輯，可離線測試）。"""
    year = year or (datetime.now().year - 1911)
    sheets_data = data.get('sheets') or []
    if not sheets_data:
        raise GoogleSyncError('雲端試算表沒有任何資料工作表')
    std_heads = template_heads(year)
    std_keys = [bl.key(h) for h in std_heads]
    book = Workbook()
    first = True
    report = []
    for item in sheets_data:
        title = str(item.get('name') or '').strip()
        if not title or title == GUIDE_SHEET:
            continue
        raw_headers = [str(h or '').strip() for h in (item.get('headers') or [])]
        raw_rows = item.get('rows') or []
        raw_fills = item.get('fills') or []
        raw_notes = item.get('notes') or []
        sheet = book.active if first else book.create_sheet()
        first = False
        sheet.title = title
        # 依表頭對應到 38 欄標準版型；額外未對應欄保留在最後
        src_map = {bl.key(h): i for i, h in enumerate(raw_headers) if h}
        extra_cols = [(i, h) for i, h in enumerate(raw_headers) if h and bl.key(h) not in std_keys]
        final_heads = list(std_heads) + [h for _, h in extra_cols]
        src_indices = [src_map.get(k) for k in std_keys] + [i for i, _ in extra_cols]
        sheet.append(final_heads)
        row_count = 0
        for r_idx, row_vals in enumerate(raw_rows):
            if not any(v not in (None, '') for v in row_vals):
                continue
            row_count += 1
            out_row = row_count + 1
            fills_row = raw_fills[r_idx] if r_idx < len(raw_fills) else []
            notes_row = raw_notes[r_idx] if r_idx < len(raw_notes) else []
            for col, src_i in enumerate(src_indices, 1):
                val = row_vals[src_i] if src_i is not None and src_i < len(row_vals) else None
                cell = sheet.cell(out_row, col, bl.number_or_text(val) if val not in (None, '') else None)
                if src_i is not None and src_i < len(fills_row):
                    bg = str(fills_row[src_i] or '').strip().lstrip('#').upper()
                    if len(bg) == 6 and bg not in ('FFFFFF', '000000'):
                        cell.fill = PatternFill('solid', fgColor='FF' + bg, bgColor='FF' + bg)
                if src_i is not None and src_i < len(notes_row) and str(notes_row[src_i] or '').strip():
                    bl.note_cell(cell, str(notes_row[src_i]))
        style_template(sheet, max(row_count + 1, 2))
        # 套回雲端保留底色（style_template 不蓋掉資料格底色，確保標記維持）
        for r_idx in range(row_count):
            fills_row = raw_fills[r_idx] if r_idx < len(raw_fills) else []
            for col, src_i in enumerate(src_indices, 1):
                if src_i is not None and src_i < len(fills_row):
                    bg = str(fills_row[src_i] or '').strip().lstrip('#').upper()
                    if len(bg) == 6 and bg not in ('FFFFFF', '000000'):
                        sheet.cell(r_idx + 2, col).fill = PatternFill('solid', fgColor='FF' + bg, bgColor='FF' + bg)
        add_formulas(sheet)
        add_notes(sheet)
        fit(sheet)
        report.append(f'{title}：匯出 {row_count} 座')
    if first:
        raise GoogleSyncError('雲端試算表沒有可匯出的資料工作表')
    add_guide(book)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    book.save(output_path)
    return report
