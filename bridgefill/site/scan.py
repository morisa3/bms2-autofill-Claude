"""唯讀掃描流程：從缺失主頁面出發，讀取第一級缺失清單與第二級表單，並核對筆數。

網頁取得交給 fetcher（介面層以內嵌瀏覽器實作；測試以本機檔案實作），本模組不依賴 Qt。
fetcher 只能開啟 is_read_only_url() 允許的網址，送出、刪除、更新類網址一律拒絕。
"""
import re
from dataclasses import asdict, dataclass, field

from .pages import (DEFECT, GOOD, UNABLE, check_count, parse_basic_data, parse_defect_form, parse_defect_list,
                    parse_girders, parse_inspection, parse_spans, url_params)

# 允許開啟的唯讀頁面（GET）；其他一律拒絕
READ_ONLY = (
    r'/inspect/inspect/der_list/',          # 第一級缺失清單
    r'/inspect/damage/view_damage/',        # 第二級檢視
    r'/inspect/damage/edit_damage/',        # 第二級編輯頁（只讀取欄位值，不送出）
    r'/inspect/inspect/result/',            # 歷史定期檢測紀錄
    r'/inspect/inspect/edit/',              # 缺失主頁面
    r'/basic/inventory/edit/\d+/?$',        # 基本資料（只讀取）
    r'/basic/extended/list/',               # 構件／擴建列表
    r'/basic/member/list/',                 # 各構件列表
)
FORBIDDEN = re.compile(r'delete|update|_submit|/add/|insert|remove|save', re.I)
# 允許呼叫的網站列表函式（頁內載入，不修改資料）
READ_ONLY_AJAX = re.compile(r"^(?:basic|baisc)/(?:span|girder)/list/[a-z_]+,\d+/(?:[a-z_]+,\d*/)*$")


def is_read_only_url(url):
    path = re.sub(r'^https?://[^/]+', '', url or '')
    return bool(path) and not FORBIDDEN.search(path) and any(re.search(p, path) for p in READ_ONLY)


def is_read_only_ajax(action):
    return bool(READ_ONLY_AJAX.match(action or '')) and not FORBIDDEN.search(action)


class Stopped(Exception):
    pass


@dataclass
class RowResult:
    component: str
    label: str
    keys: dict
    deru: str
    status: str
    repair: str
    expected: int | None
    columns: dict
    defects: list = field(default_factory=list)     # 每筆：第一級摘要 + 第二級欄位
    check: str = ''
    ok: bool = True


@dataclass
class ScanResult:
    header: dict
    rows: list = field(default_factory=list)
    problems: list = field(default_factory=list)

    @property
    def summary(self):
        defects = sum(len([d for d in r.defects if d['status'] == DEFECT]) for r in self.rows)
        failed = [r for r in self.rows if not r.ok]
        return {'構件列數': len(self.rows), '缺失筆數': defects, '核對異常列數': len(failed),
                '網頁問題': len(self.problems)}

    def to_dict(self):
        return {'header': self.header, 'summary': self.summary, 'problems': self.problems,
                'rows': [asdict(r) for r in self.rows]}


def needs_detail(row):
    """良好或無法檢測且只有 1 筆時，那 1 筆是照片紀錄，不必開；其餘都要讀第一級。
    主頁 DERU 是彙總值，多筆時可能顯示良好但其中仍有缺失，所以 n>1 一律讀取。"""
    return not (row.status in (GOOD, UNABLE) and (row.count or 0) <= 1)


def scan_inspection(main_html, fetch, read_forms=True, progress=None, stopped=None, components=None):
    """掃描一期檢測。fetch(url) 回傳網頁 HTML；progress(i, total, label) 回報進度；stopped() 為 True 時中止。"""
    page = parse_inspection(main_html)
    result = ScanResult(page.header, problems=list(page.problems))
    rows = [r for r in page.rows if not components or r.component in components]
    for index, row in enumerate(rows, 1):
        if stopped and stopped():
            raise Stopped(f'已停止，完成 {index - 1}／{len(rows)} 列')
        if progress:
            progress(index, len(rows), f'{row.component} {row.label}')
        item = RowResult(row.component, row.label, row.keys, row.deru, row.status, row.repair, row.count,
                         row.columns, check=row.problem, ok=not row.problem)
        result.rows.append(item)

        def fail(message, item=item):
            item.ok = False
            item.check = '；'.join(m for m in (item.check, message) if m)

        if not needs_detail(row):
            item.check = item.check or f'{row.status}，未開啟'
            continue
        if not is_read_only_url(row.detail_url):
            fail(f'詳細資料網址不在唯讀清單，未開啟：{row.detail_url}')
            continue
        defects = parse_defect_list(fetch(row.detail_url))
        counted = check_count(row, defects)
        if not counted.ok:
            fail(counted.message)
        for summary in defects:
            record = {'deru': summary.deru, 'status': summary.status, 'deterioration': summary.deterioration,
                      'method_id': summary.method_id, 'der_id': summary.der_id, 'repaired': summary.repaired,
                      'fields': {}, 'options': {}}
            item.defects.append(record)
            if read_forms and summary.status == DEFECT:
                url = summary.edit_url  # 檢視頁的網頁結構尚未取得，先讀編輯頁的欄位值（不送出）
                if not is_read_only_url(url):
                    fail(f'表單網址不在唯讀清單：{url}')
                    continue
                form = parse_defect_form(fetch(url))
                if not form.method_id:
                    fail(f'第二級表單讀取失敗（ID {summary.method_id}），請重掃')
                    continue
                if form.method_id != summary.method_id:
                    fail(f'第二級 ID {form.method_id} 與第一級 {summary.method_id} 不一致')
                    continue
                record['fields'], record['options'] = form.fields, form.options
    return result


def basic_urls(inventory_id):
    return {'basic': f'https://bms2.iot.gov.tw/bms3/basic/inventory/edit/{inventory_id}/',
            'extended': f'https://bms2.iot.gov.tw/bms3/basic/extended/list/inventory_id,{inventory_id}/'}


def extended_ids(extended_html):
    """構件／擴建列表頁：取出各結構的 extended_id（主結構與擴建結構）。"""
    return list(dict.fromkeys(re.findall(r'/basic/member/list/extended_id,(\d+)/', extended_html or '')))


def scan_basic(inventory_id, fetch, call):
    """基本資料：淨寬、跨距、橋下淨高；各結構的橋孔與主梁根數。
    call(page_url, ajax_action, ready_text) 在頁內呼叫網站列表函式並回傳載入後的 HTML。"""
    urls = basic_urls(inventory_id)
    basic = parse_basic_data(fetch(urls['basic']))
    structures = []
    for ext in extended_ids(fetch(urls['extended'])):
        member = f'https://bms2.iot.gov.tw/bms3/basic/member/list/extended_id,{ext}/'
        spans = parse_spans(call(member, f'baisc/span/list/extended_id,{ext}/member_action,/', '橋孔編號'))
        span_rows = []
        for span in spans:
            girders = []
            action = re.search(r"'([^']*girder/list/[^']*)'", span.girder_url or '')
            if action and is_read_only_ajax(action[1]):
                girders = parse_girders(call(member, action[1], '主梁編號'))
            span_rows.append({'code': span.code, 'length': str(span.length or ''),
                              'full_width': str(span.full_width or ''), 'girders': girders})
        structures.append({'extended_id': ext, 'spans': span_rows})
    return {'values': basic.values, 'distribution': [str(s) for s in basic.spans], 'structures': structures}


def inventory_id(result_or_page):
    for row in getattr(result_or_page, 'rows', []):
        keys = row.keys if hasattr(row, 'keys') else {}
        if keys.get('inventory_id'):
            return keys['inventory_id']
    return ''


__all__ = ['scan_inspection', 'scan_basic', 'is_read_only_url', 'is_read_only_ajax', 'Stopped', 'url_params']
