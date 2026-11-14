"""主頁 DERU 字串解析（自 v1.00 navigation.py 抽出，不依賴瀏覽器）。"""
import re
import unicodedata


def deru_records(text):
    """主頁 DERU 可為 2211、1___、含換行的多筆，或 D:2 E:2 R:1 U:1。"""
    value = unicodedata.normalize("NFKC", text or "").upper()
    labelled = re.fullmatch(r"\s*D\s*[:：=]?\s*([0-9_\-—－])\s*E\s*[:：=]?\s*([0-9_\-—－])\s*R\s*[:：=]?\s*([0-9_\-—－])\s*U\s*[:：=]?\s*([0-9_\-—－])\s*", value)
    if labelled:
        records = [''.join(labelled.groups())]
    else:
        value = re.sub(r"\s+", "", value)
        if not re.fullmatch(r"[0-9_\-—－]+", value) or len(value) % 4:
            return None
        records = [value[i:i + 4] for i in range(0, len(value), 4)]
    return records


def deru_status(text):
    records=deru_records(text)
    if records is None:return None
    outcomes = []
    for record in records:
        d, e = record[:2]
        outcomes.append(False if d == '1' or e == '0' else
                        True if d.isdigit() and e.isdigit() else None)
    return True if True in outcomes else None if None in outcomes else False


