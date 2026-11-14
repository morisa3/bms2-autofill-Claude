"""介面顯示用資訊，不改動資料識別。"""
import re


def comparison_metadata(entry):
    context=entry.get('context',{});items=['網站 ID：'+(entry.get('website_id') or '未提供')]
    items.extend(f'{label}：{value}' for label,value in context.items() if not label.startswith('_') and ('位置' in label or '編號' in label) and str(value).strip())
    raw=str(context.get('DERU',''));values=dict(re.findall(r'([DERU])\s*[:：=]\s*([^\s｜]+)',raw))
    if re.fullmatch(r'[0-4_—-]{4}',raw):values=dict(zip('DERU',raw))
    der=entry.get('der',[])
    if len(der)>=2:values.update(D=der[0],E=der[1])
    for name in ('D','E','R','U'):
        if context.get(name):values[name]=context[name]
    items.append('DERU：'+' '.join(f'{name}:{values.get(name) or "—"}' for name in 'DERU'))
    items.append('規則 ID：'+(entry.get('rule_id') or '未匹配'))
    return '｜'.join(items)

ACTION_LABELS={
'開啟系統':'開啟專用瀏覽器','連接系統':'連接目前橋梁','斷開系統':'解除連接',
'讀取 Excel 表單':'匯入維修規則（Excel）','前期檢測資料導入':'匯入前期檢測文件','本期檢測資料導入':'匯入本期下載文件','掃描系統建立本期資料':'從網站讀取本期缺失',
'開始運行':'逐筆確認並填寫','接續掃描':'繼續逐筆處理','檢核模式掃描':'檢查網站檢測資料',
'兩期文件比對並生成':'比對前期與本期資料','導入生成 DOCX 文件':'匯入比對完成檔','進行自動填寫':'自動填寫網站','從中斷處接續自動填寫':'從中斷處繼續填寫','填寫檢測員意見':'編輯檢測員意見'}
ACTION_HINTS={
'本期檢測資料導入':'匯入網站下載的完整本期 Word，提供版型、基本資料與良好列；有缺失的內容由下一步掃描建立。',
'掃描系統建立本期資料':'先回網站缺失主頁；只讀取選定構件的缺失及網站 ID，不送出更新。',
'兩期文件比對並生成':'使用本期掃描資料與前期比對，檢查必填欄位後產生並導入完成檔。',
'導入生成 DOCX 文件':'已有助手產生的完成檔時可直接匯入，必須含有效的網站劣化 ID。',
'進行自動填寫':'依完成檔的網站 ID 逐筆填寫與回讀，不先整批預掃描。',
'開始運行':'選用：逐筆掃描網站，必要時在主頁人工核對；完整自動流程請使用下方比對與完成檔填寫。'}


LOADED_LABELS={
'讀取 Excel 表單':'✓ 已匯入維修規則',
'前期檢測資料導入':'✓ 已匯入前期檢測文件',
'本期檢測資料導入':'✓ 已匯入本期下載文件',
'掃描系統建立本期資料':'✓ 已讀取本期缺失',
'導入生成 DOCX 文件':'✓ 已匯入比對完成檔',
'兩期文件比對並生成':'✓ 已完成兩期比對',
}
WORKFLOW_HINT='全部填寫：匯入文件 → 讀取網站缺失 → 比對並產生填寫檔 → 自動填寫網站。\n局部填寫：先選構件，再使用「逐筆確認並填寫」。'
FILL_NOT_READY='請先完成兩期比對，或匯入比對完成檔。'
