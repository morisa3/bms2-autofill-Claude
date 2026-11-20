"""記住網站帳號密碼（使用者自行開啟；2026-11-19 使用者要求，2026-11-20 改為可存多組並加標籤）。

- 每組帳號密碼各存一筆 Windows「認證管理員」一般認證（名稱「BridgeFill:bms2:標籤」），由 Windows 以登入帳號加密保存；
  程式資料夾、紀錄檔、版本庫都不會有密碼。預設帳號的標籤存在「BridgeFill:bms2-default」（不含密碼）。
- 舊版只有一組（名稱「BridgeFill:bms2」）：第一次讀取時改存成標籤＝帳號的一組，並設為預設。
- 登入頁開啟時自動填入預設帳號；「我不是機器人」與「確定登入」仍由使用者操作，程式不代為登入。
- 非 Windows（開發、測試環境）沒有認證管理員：功能停用（available() 為 False）；測試可替換 backend。
"""
import ctypes
import sys

PREFIX = 'BridgeFill:bms2:'
LEGACY = 'BridgeFill:bms2'
DEFAULT = 'BridgeFill:bms2-default'


class WindowsVault:
    """Windows 認證管理員（advapi32 CredWriteW／CredReadW／CredDeleteW／CredEnumerateW）。"""
    GENERIC, PERSIST_LOCAL = 1, 2

    def __init__(self):
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [('low', wintypes.DWORD), ('high', wintypes.DWORD)]

        class CREDENTIAL(ctypes.Structure):
            _fields_ = [('Flags', wintypes.DWORD), ('Type', wintypes.DWORD), ('TargetName', wintypes.LPWSTR),
                        ('Comment', wintypes.LPWSTR), ('LastWritten', FILETIME),
                        ('CredentialBlobSize', wintypes.DWORD), ('CredentialBlob', ctypes.POINTER(ctypes.c_ubyte)),
                        ('Persist', wintypes.DWORD), ('AttributeCount', wintypes.DWORD),
                        ('Attributes', ctypes.c_void_p), ('TargetAlias', wintypes.LPWSTR),
                        ('UserName', wintypes.LPWSTR)]
        self.CREDENTIAL = CREDENTIAL
        api = ctypes.WinDLL('advapi32', use_last_error=True)
        self.cred_write = api.CredWriteW
        self.cred_write.argtypes = [ctypes.POINTER(CREDENTIAL), wintypes.DWORD]
        self.cred_write.restype = wintypes.BOOL
        self.cred_read = api.CredReadW
        self.cred_read.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                   ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
        self.cred_read.restype = wintypes.BOOL
        self.cred_delete = api.CredDeleteW
        self.cred_delete.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
        self.cred_delete.restype = wintypes.BOOL
        self.cred_enumerate = api.CredEnumerateW
        self.cred_enumerate.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD),
                                        ctypes.POINTER(ctypes.POINTER(ctypes.POINTER(CREDENTIAL)))]
        self.cred_enumerate.restype = wintypes.BOOL
        self.free = api.CredFree
        self.free.argtypes = [ctypes.c_void_p]

    def write(self, target, user, secret, comment=''):
        blob = secret.encode('utf-16-le')
        buffer = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob) if blob else None
        credential = self.CREDENTIAL()
        credential.Type, credential.Persist = self.GENERIC, self.PERSIST_LOCAL
        credential.TargetName, credential.UserName = target, user
        credential.Comment = comment or '公路橋梁填報助手：bms2 網站帳號密碼（登入頁自動填入）'
        credential.CredentialBlobSize = len(blob)
        credential.CredentialBlob = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)) if buffer else None
        if not self.cred_write(ctypes.byref(credential), 0):
            raise OSError(ctypes.get_last_error(), '無法寫入 Windows 認證管理員')

    def read(self, target):
        pointer = ctypes.POINTER(self.CREDENTIAL)()
        if not self.cred_read(target, self.GENERIC, 0, ctypes.byref(pointer)):
            return None
        try:
            credential = pointer.contents
            blob = ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize)
            return credential.UserName or '', blob.decode('utf-16-le')
        finally:
            self.free(pointer)

    def delete(self, target):
        self.cred_delete(target, self.GENERIC, 0)

    def targets(self, prefix):
        from ctypes import wintypes
        count = wintypes.DWORD()
        items = ctypes.POINTER(ctypes.POINTER(self.CREDENTIAL))()
        if not self.cred_enumerate(prefix + '*', 0, ctypes.byref(count), ctypes.byref(items)):
            return []
        try:
            return [items[i].contents.TargetName for i in range(count.value)]
        finally:
            self.free(items)


class MemoryVault:
    """測試用：存在記憶體。"""

    def __init__(self):
        self.items = {}

    def write(self, target, user, secret, comment=''):
        self.items[target] = (user, secret)

    def read(self, target):
        return self.items.get(target)

    def delete(self, target):
        self.items.pop(target, None)

    def targets(self, prefix):
        return [t for t in self.items if t.startswith(prefix)]


_vault = None


def vault():
    global _vault
    if _vault is None and sys.platform == 'win32':
        try:
            _vault = WindowsVault()
        except (OSError, AttributeError):
            _vault = None
    return _vault


def use(backend):
    """替換保存方式（測試用）。"""
    global _vault
    _vault = backend


def available():
    return vault() is not None


def _migrate():
    """舊版單一帳號（BridgeFill:bms2）改存成標籤＝帳號的一組。"""
    old = vault().read(LEGACY)
    if old:
        label = old[0] or '帳號1'
        vault().write(PREFIX + label, old[0], old[1])
        if not vault().read(DEFAULT):
            vault().write(DEFAULT, label, '', '公路橋梁填報助手：預設帳號的標籤（不含密碼）')
        vault().delete(LEGACY)


def accounts():
    """已記住的帳號：[(標籤, 帳號)]，依標籤排序；不含密碼。"""
    if not available():
        return []
    try:
        _migrate()
        found = []
        for target in vault().targets(PREFIX):
            saved = vault().read(target)
            if saved:
                found.append((target[len(PREFIX):], saved[0]))
        return sorted(found)
    except OSError:
        return []


def default_label():
    """預設帳號的標籤；沒有設定或已刪除時用第一組。"""
    labels = [label for label, _ in accounts()]
    if not labels:
        return ''
    try:
        chosen = vault().read(DEFAULT)
    except OSError:
        chosen = None
    return chosen[0] if chosen and chosen[0] in labels else labels[0]


def set_default(label):
    if available() and label:
        vault().write(DEFAULT, label, '', '公路橋梁填報助手：預設帳號的標籤（不含密碼）')


def save(user, password, label=''):
    """新增或更新一組（標籤空白時用帳號當標籤）；回傳標籤。第一組自動設為預設。"""
    if not available():
        raise OSError('這台電腦沒有 Windows 認證管理員，無法記住帳號密碼')
    label = (label or user).strip()
    if not label or ':' in label or '*' in label:
        raise ValueError('標籤不可空白，也不能含「:」「*」')
    _migrate()
    first = not accounts()
    vault().write(PREFIX + label, user, password)
    if first:
        set_default(label)
    return label


def load(label=None):
    """(帳號, 密碼)；label 省略時讀預設帳號。沒有回傳 None。"""
    if not available():
        return None
    try:
        label = label or default_label()
        return vault().read(PREFIX + label) if label else None
    except OSError:
        return None


def clear(label=None):
    """刪除一組；label 省略時刪除全部。"""
    if not available():
        return
    labels = [label] if label else [name for name, _ in accounts()]
    for name in labels:
        vault().delete(PREFIX + name)
    if not label:
        vault().delete(DEFAULT)
        vault().delete(LEGACY)


LOGIN_FILL_JS = r"""((user, password, force) => {
  const name = document.getElementById('user'), secret = document.getElementById('password');
  if (!name || !secret || !secret.form) return;
  if (force || !name.value) name.value = user;
  if (force || !secret.value) secret.value = password;
  for (const box of [name, secret]) box.dispatchEvent(new Event('input', {bubbles: true}));
})(%s, %s, %s)"""
