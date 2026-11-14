"""v1.00 測試中有部分經由舊 Flet 介面（bridgefill.app.App）驗證流程。

新專案不搬舊介面；這些測試在呼叫 ui() 時自動略過，待新 PySide6 介面完成後改寫。
同檔案中的其他純邏輯測試照常執行。
"""
import unittest


def ui():
    raise unittest.SkipTest('舊 Flet 介面測試：待新介面完成後改寫')
