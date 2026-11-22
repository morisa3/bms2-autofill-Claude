"""下橋路徑與輔助設備（3.0，2026-11-24 使用者需求）：選項組合成總表兩格的文字，也能從總表文字讀回選項。

等級（總表「下橋路徑(ABCD)」第一行）：
- D：徒步／既有檢修梯下橋；C：架設一／二／三截梯下橋——可多選下橋位置（A1上游側、A1下游側、A2上游側、A2下游側），
  可加水深裝備（長筒雨鞋／青蛙裝）；
- B：需用橋檢車／高空作業車／工作船檢測，同時寫入輔助設備（B）；
- A：需用空拍機／水下攝影機檢測，同時寫入輔助設備（A）；
- 都可加備註（手打）。
例：「C 換行 A1下游側架設二截梯下橋 換行 水深需要長筒雨鞋」；「B 換行 需用橋檢車檢測」＋輔助設備「B 換行 橋檢車」。
本模組不依賴 Qt。
"""
import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

LEVELS = ('A', 'B', 'C', 'D')
POSITIONS = ('A1上游側', 'A1下游側', 'A2上游側', 'A2下游側')
METHODS = {'D': ('徒步', '既有檢修梯'), 'C': ('架設一截梯', '架設二截梯', '架設三截梯')}
EQUIPMENT = {'B': ('橋檢車', '高空作業車', '工作船'), 'A': ('空拍機', '水下攝影機')}
WATER = ('長筒雨鞋', '青蛙裝')
LEVEL_NAMES = {'A': 'A級（間接檢測）：須配合無人飛行器、水下無人機等設備',
               'B': 'B級（直接檢測）：須搭載輔助載具，如橋梁檢測車、高空工作車或工作船舶',
               'C': 'C級（直接檢測）：須搭設移動梯（O截梯）等設備到達橋下',
               'D': 'D級（直接檢測）：檢測人員可徒步到達橋下'}


@dataclass
class Access:
    level: str = ''                                   # A／B／C／D
    positions: list = field(default_factory=list)     # C、D：下橋位置（可多選）
    method: str = ''                                  # C、D：徒步、既有檢修梯、架設 n 截梯
    water: list = field(default_factory=list)         # C、D：水深需要的裝備
    equipment: list = field(default_factory=list)     # A、B：設備（可多選）
    note: str = ''                                    # 備註（手打）

    @property
    def empty(self):
        return not self.level


def path_text(access):
    """總表「下橋路徑(ABCD)」：等級＋下橋方式（或需用設備）＋水深裝備＋備註，各一行。"""
    if access.empty:
        return ''
    lines = [access.level]
    if access.level in ('C', 'D'):
        where = '、'.join(p for p in POSITIONS if p in access.positions)
        method = access.method or ('徒步' if access.level == 'D' and where else '')
        if where or method:
            lines.append(f'{where}{method}下橋')
        if access.water:
            lines.append('水深需要' + '、'.join(w for w in WATER if w in access.water))
    elif access.equipment:
        lines.append('需用' + '／'.join(e for e in EQUIPMENT[access.level] if e in access.equipment) + '檢測')
    if access.note.strip():
        lines.append(access.note.strip())
    return '\n'.join(lines)


def aux_text(access):
    """總表「輔助設備(A/B)」：A、B 級寫「等級 換行 設備」；C、D 級空白。"""
    if access.level not in EQUIPMENT or not access.equipment:
        return ''
    return access.level + '\n' + '、'.join(e for e in EQUIPMENT[access.level] if e in access.equipment)


def parse(path, aux=''):
    """總表兩格文字 → Access（程式組出的寫法可完整讀回；其他寫法盡量辨認，認不出的放備註）。"""
    text = str(path or '').strip()
    extra = str(aux or '').strip()
    if not text and not extra:
        return Access()
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    level = ''
    if lines and re.fullmatch(r'[ABCD]', lines[0].upper()):
        level = lines.pop(0).upper()
    elif re.match(r'^[ABCD](?![A-Za-z0-9])', text.upper()):
        level = text[0].upper()
        lines = [text[1:].strip()] if text[1:].strip() else []
    aux_lines = [line.strip() for line in extra.splitlines() if line.strip()]
    if not level and aux_lines and re.fullmatch(r'[AB]', aux_lines[0].upper()):
        level = aux_lines[0].upper()
    access = Access(level)
    whole = '\n'.join(lines + aux_lines)
    access.positions = [p for p in POSITIONS if p in whole]
    for candidates in METHODS.values():
        access.method = next((m for m in candidates if m in whole), access.method)
    access.water = [w for w in WATER if w in whole]
    for candidates in EQUIPMENT.values():
        access.equipment += [e for e in candidates if e in whole and e not in access.equipment]
    known = {path_line for path_line in path_text(access).splitlines()[1:]}
    access.note = '\n'.join(line for line in lines if line not in known and not composed(line))
    return access


def composed(line):
    """程式組出的行（下橋方式、水深、需用設備）：讀回時不放進備註。"""
    return bool(re.search(r'下橋$|^水深需要|^需用.*檢測$', line))


def access_path(folder, bridge, date):
    safe = ''.join('_' if ch in '\\/:*?"<>|' else ch for ch in f'{date}-{bridge}-下橋路徑.json')
    return Path(folder) / safe


def load(folder, bridge, date):
    try:
        data = json.loads(access_path(folder, bridge, date).read_text(encoding='utf-8'))
        return Access(**{k: data.get(k, Access().__dict__[k]) for k in Access().__dict__})
    except (OSError, ValueError, TypeError):
        return None


def save(folder, bridge, date, access, path_override='', aux_override=''):
    """存檔：選項＋實際輸出的文字（預覽框手動改過時以手動為準）。"""
    data = asdict(access)
    data.update(path=path_override or path_text(access), aux=aux_override if aux_override is not None else aux_text(access))
    target = access_path(folder, bridge, date)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return target


def saved_texts(folder, bridge, date):
    """存檔中實際輸出的兩格文字 (下橋路徑, 輔助設備)；沒有存檔回傳 ('', '')。"""
    try:
        data = json.loads(access_path(folder, bridge, date).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return '', ''
    return data.get('path', ''), data.get('aux', '')
