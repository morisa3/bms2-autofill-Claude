"""記住網站帳號密碼（使用者自行開啟；2026-11-19 使用者要求）。

- 存在 Windows「認證管理員」（控制台 → 認證管理員 → Windows 認證 → 一般認證，名稱「BridgeFill:bms2」），
  由 Windows 以登入帳號加密保存；程式資料夾、紀錄檔、版本庫都不會有密碼。
- 登入頁開啟時自動填入帳號、密碼；「我不是機器人」與「確定登入」仍由使用者操作，程式不代為登入。
- 非 Windows（開發、測試環境）沒有認證管理員：功能停用（available() 為 False）；測試可替換 backend。
"""
import ctypes
import sys

TARGET = 'BridgeFill:bms2'


class WindowsVault:
    """Windows 認證管理員（advapi32 CredWriteW／CredReadW／CredDeleteW）。"""
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
        self.write = api.CredWriteW
        self.write.argtypes = [ctypes.POINTER(CREDENTIAL), wintypes.DWORD]
        self.write.restype = wintypes.BOOL
        self.read = api.CredReadW
        self.read.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
        self.read.restype = wintypes.BOOL
        self.delete = api.CredDeleteW
        self.delete.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
        self.delete.restype = wintypes.BOOL
        self.free = api.CredFree
        self.free.argtypes = [ctypes.c_void_p]

    def save(self, user, password):
        blob = password.encode('utf-16-le')
        buffer = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob) if blob else None
        credential = self.CREDENTIAL()
        credential.Type, credential.Persist = self.GENERIC, self.PERSIST_LOCAL
        credential.TargetName, credential.UserName = TARGET, user
        credential.Comment = '公路橋梁填報助手：bms2 網站帳號密碼（登入頁自動填入）'
        credential.CredentialBlobSize = len(blob)
        credential.CredentialBlob = ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)) if buffer else None
        if not self.write(ctypes.byref(credential), 0):
            raise OSError(ctypes.get_last_error(), '無法寫入 Windows 認證管理員')

    def load(self):
        pointer = ctypes.POINTER(self.CREDENTIAL)()
        if not self.read(TARGET, self.GENERIC, 0, ctypes.byref(pointer)):
            return None
        try:
            credential = pointer.contents
            blob = ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize)
            return credential.UserName or '', blob.decode('utf-16-le')
        finally:
            self.free(pointer)

    def clear(self):
        self.delete(TARGET, self.GENERIC, 0)


class MemoryVault:
    """測試用：存在記憶體。"""

    def __init__(self):
        self.value = None

    def save(self, user, password):
        self.value = (user, password)

    def load(self):
        return self.value

    def clear(self):
        self.value = None


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


def save(user, password):
    if not available():
        raise OSError('這台電腦沒有 Windows 認證管理員，無法記住帳號密碼')
    vault().save(user, password)


def load():
    if not available():
        return None
    try:
        return vault().load()
    except OSError:
        return None


def clear():
    if available():
        vault().clear()


LOGIN_FILL_JS = r"""((user, password) => {
  const name = document.getElementById('user'), secret = document.getElementById('password');
  if (!name || !secret || !secret.form) return;
  if (!name.value) name.value = user;
  if (!secret.value) secret.value = password;
  for (const box of [name, secret]) box.dispatchEvent(new Event('input', {bubbles: true}));
})(%s, %s)"""
