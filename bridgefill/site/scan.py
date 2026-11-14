"""唯讀掃描流程：從缺失主頁面出發，讀取第一級缺失清單與第二級表單，並核對筆數。

網頁取得交給 fetcher（介面層以內嵌瀏覽器實作；測試以本機檔案實作），本模組不依賴 Qt。
fetcher 只能開啟 is_read_only_url() 允許的網址，送出、刪除、更新類網址一律拒絕。
"""
import re
from dataclasses import asdict, dataclass, field

from .pages import (DEFECT, GOOD, UNABLE, check_count, parse_basic_data, parse_defect_form, parse_defect_list,
                    parse_defect_view, parse_girders, parse_history, parse_inspection, parse_spans,
                    previous_inspection, url_params)

# 允許開啟的唯讀頁面（GET）；其他一律拒絕
READ_ONLY = (
    r'/inspect/inspect/der_list/',          # 第一級缺失清單
    r'/inspect/damage/view_damage/',        # 第二級檢視
    r'/inspect/damage/edit_damage/',        # 第二級編輯頁（只讀取欄位值，不送出）
    r'/inspect/inspect/result/',            # 歷史定期檢測紀錄
    r'/inspect/inspect/der_content/',       # 缺失主頁面（只讀取）
    r'/inspect/inspect/view/',              # 整期檢測檢視
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


def scan_inspection(main_html, fetch, read_forms=True, progress=None, stopped=None, components=None, mode='edit'):
    """掃描一期檢測。fetch(url) 回傳網頁 HTML；progress(i, total, label) 回報進度；stopped() 為 True 時中止。
    mode='edit'（本期）：讀編輯頁欄位（不送出），可核對 ID 並取得劣化類型選項。
    mode='view'（前期）：讀純檢視頁；檢視頁沒有 ID，以網址上的 der_id 對應第一級的 ID。"""
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
        try:
            defects = parse_defect_list(fetch(row.detail_url))
        except Stopped:
            raise
        except Exception as error:  # 單列讀取失敗不中斷整次掃描
            fail(f'第一級讀取失敗：{error}')
            continue
        counted = check_count(row, defects)
        if not counted.ok:
            fail(counted.message)
        for summary in defects:
            record = {'deru': summary.deru, 'status': summary.status, 'deterioration': summary.deterioration,
                      'method_id': summary.method_id, 'der_id': summary.der_id, 'repaired': summary.repaired,
                      'fields': {}, 'options': {}}
            item.defects.append(record)
            if read_forms and summary.status == DEFECT:
                url = summary.view_url if mode == 'view' else summary.edit_url
                if not is_read_only_url(url):
                    fail(f'表單網址不在唯讀清單：{url}')
                    continue
                try:
                    form = (parse_defect_view if mode == 'view' else parse_defect_form)(fetch(url))
                except Stopped:
                    raise
                except Exception as error:
                    fail(f'第二級讀取失敗（ID {summary.method_id}）：{error}')
                    continue
                if mode == 'view':
                    if not form.fields:
                        fail(f'第二級檢視頁讀取失敗（ID {summary.method_id}），請重掃')
                        continue
                elif not form.method_id:
                    fail(f'第二級表單讀取失敗（ID {summary.method_id}），請重掃')
                    continue
                elif form.method_id != summary.method_id:
                    fail(f'第二級 ID {form.method_id} 與第一級 {summary.method_id} 不一致')
                    continue
                record['fields'], record['options'] = form.fields, form.options
    return result


def page_ids(page):
    """從缺失主頁面任一列的詳細資料連結取得 inspect_id 與 inventory_id。"""
    for row in page.rows:
        if row.keys.get('inspect_id') and row.keys.get('inventory_id'):
            return row.keys['inspect_id'], row.keys['inventory_id']
    return '', ''


def history_url(inventory_id):
    return f'https://bms2.iot.gov.tw/bms3/inspect/inspect/result/inventory_id,{inventory_id}/from,inventory/'


def scan_prior(current_html, fetch, host=None, **options):
    """由本期缺失主頁面找出前期（歷史定期檢測紀錄中日期較早的最新一筆），以檢視頁唯讀掃描。
    回傳 (前期紀錄, 掃描結果)。"""
    current = parse_inspection(current_html)
    inspect_id, inventory = page_ids(current)
    if not inventory:
        raise ValueError('目前頁面找不到橋梁代號，請開到本期的檢測缺失主頁面')
    history = parse_history(fetch(history_url(inventory)))
    prior = previous_inspection(history, inspect_id, current.header.get('檢測日期', ''))
    if prior is None:
        raise ValueError('歷史定期檢測紀錄中沒有更早的檢測')
    if not prior.edit_url:
        raise ValueError(f'前期 {prior.date}（{prior.unit}）在歷史紀錄中沒有「修改」連結，'
                         '可能是其他單位的檢測、帳號只有檢視權限')
    if not is_read_only_url(prior.edit_url):
        raise PermissionError(f'前期主頁面網址不在唯讀清單：{prior.edit_url}')
    prior_html = fetch(prior.edit_url)
    if host:
        host(prior.edit_url)   # 之後的第一級、第二級以前期主頁面為來源頁讀取
    return prior, scan_inspection(prior_html, fetch, mode='view', **options)


def basic_urls(inventory_id):
    return {'basic': f'https://bms2.iot.gov.tw/bms3/basic/inventory/edit/{inventory_id}/',
            'extended': f'https://bms2.iot.gov.tw/bms3/basic/extended/list/inventory_id,{inventory_id}/'}


def extended_ids(extended_html):
    """構件／擴建列表頁：取出各結構的 extended_id（主結構與擴建結構）。"""
    return list(dict.fromkeys(re.findall(r'/basic/member/list/extended_id,(\d+)/', extended_html or '')))


def scan_basic(inventory_id, fetch, call, expected_spans=None):
    """基本資料：淨寬、跨距、橋下淨高；各結構的橋孔與主梁根數。
    call(page_url, ajax_action, ready_text) 在頁內呼叫網站列表函式並回傳載入後的 HTML。
    任一部分失敗只記在 problems，其餘照常讀取。expected_spans 為主頁面的總橋孔數，用來核對。"""
    urls = basic_urls(inventory_id)
    problems, values, distribution, structures = [], {}, [], []
    try:
        basic = parse_basic_data(fetch(urls['basic']))
        values, distribution = basic.values, basic.spans
        problems += basic.problems
    except Stopped:
        raise
    except Exception as error:
        problems.append(f'基本資料頁讀取失敗：{error}')
    try:
        structure_ids = extended_ids(fetch(urls['extended']))
    except Stopped:
        raise
    except Exception as error:
        structure_ids = []
        problems.append(f'構件／擴建列表讀取失敗：{error}')
    for ext in structure_ids:
        member = f'https://bms2.iot.gov.tw/bms3/basic/member/list/extended_id,{ext}/'
        try:
            spans = parse_spans(call(member, f'baisc/span/list/extended_id,{ext}/member_action,/', '橋孔編號'))
        except Stopped:
            raise
        except Exception as error:
            problems.append(f'橋孔資料表讀取失敗（結構 {ext}）：{error}')
            continue
        span_rows = []
        for span in spans:
            girders = []
            action = re.search(r"'([^']*girder/list/[^']*)'", span.girder_url or '')
            if action and is_read_only_ajax(action[1]):
                try:
                    girders = parse_girders(call(member, action[1], '主梁編號'))
                except Stopped:
                    raise
                except Exception as error:
                    problems.append(f'主梁明細讀取失敗（{span.code}）：{error}')
            span_rows.append({'code': span.code, 'length': str(span.length or ''),
                              'full_width': str(span.full_width or ''), 'girders': girders})
        structures.append({'extended_id': ext, 'spans': span_rows})
    table_count = sum(len(s['spans']) for s in structures)
    counts = {'跨距分配': len(distribution), '橋孔資料表': table_count, '總橋孔數': expected_spans}
    known = {k: v for k, v in counts.items() if v}
    if len(set(known.values())) > 1:
        problems.append('孔數不一致：' + '、'.join(f'{k} {v}' for k, v in known.items()) +
                        '；跨距以橋孔資料表為準')
    return {'values': values, 'distribution': [str(s) for s in distribution], 'structures': structures,
            'problems': problems}


def inventory_id(result_or_page):
    for row in getattr(result_or_page, 'rows', []):
        keys = row.keys if hasattr(row, 'keys') else {}
        if keys.get('inventory_id'):
            return keys['inventory_id']
    return ''


__all__ = ['scan_inspection', 'scan_basic', 'is_read_only_url', 'is_read_only_ajax', 'Stopped', 'url_params']
