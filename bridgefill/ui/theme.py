"""介面外觀：方案 A「Windows 11 辦公風」（2026-11-19 使用者選定）。

淺灰底、白色內容區；分頁選中時藍色底線；圓角按鈕，主要按鈕（property primary=true）藍色實心；
表格斑馬紋、表頭淺灰。狀態色（待處理黃、可送出綠…）維持各元件自己的設定。只改外觀，不改功能。
"""
FONT = '"Microsoft JhengHei UI", "Microsoft JhengHei", "Noto Sans CJK TC", sans-serif'
ACCENT = '#0067c0'

STYLE = f"""
* {{ font-family: {FONT}; font-size: 10pt; }}
QMainWindow, QDialog {{ background: #f3f3f3; }}
QWidget {{ color: #1b1b1b; }}
QPushButton {{ background: #ffffff; border: 1px solid #d0d0d0; border-radius: 6px; padding: 4px 8px; }}
QPushButton:hover {{ background: #f5f9ff; border-color: {ACCENT}; }}
QPushButton:pressed {{ background: #e5efff; }}
QPushButton:checked {{ background: #e5efff; border-color: {ACCENT}; }}
QPushButton:disabled {{ color: #a0a0a0; background: #f7f7f7; border-color: #e0e0e0; }}
QPushButton[primary="true"] {{ background: {ACCENT}; color: #ffffff; border: 1px solid #005ba8; font-weight: bold; }}
QPushButton[primary="true"]:hover {{ background: #1975c5; }}
QPushButton[primary="true"]:disabled {{ background: #9fc3e6; border-color: #9fc3e6; color: #ffffff; }}
QLineEdit, QComboBox, QTimeEdit, QPlainTextEdit, QTextEdit {{ background: #ffffff; border: 1px solid #d0d0d0;
    border-radius: 6px; padding: 1px 6px; selection-background-color: #cce4f7; selection-color: #000000; }}
QLineEdit:focus, QComboBox:focus, QTimeEdit:focus, QPlainTextEdit:focus, QTextEdit:focus {{ border: 1px solid {ACCENT}; }}
QComboBox QAbstractItemView {{ background: #ffffff; selection-background-color: #cce4f7; selection-color: #000000; }}
QTabWidget::pane {{ border: 1px solid #e0e0e0; border-radius: 8px; background: #ffffff; top: -1px; }}
QTabBar::tab {{ background: transparent; padding: 6px 14px; margin-right: 2px; border: none; color: #555555; font-size: 10.5pt; }}
QTabBar::tab:selected {{ color: {ACCENT}; border-bottom: 3px solid {ACCENT}; font-weight: bold; }}
QTabBar::tab:hover {{ color: {ACCENT}; }}
QTableWidget, QListWidget {{ background: #ffffff; alternate-background-color: #fafafa; gridline-color: #ececec;
    border: 1px solid #e0e0e0; border-radius: 6px; }}
QHeaderView::section {{ background: #f7f7f7; border: none; border-bottom: 1px solid #dcdcdc; border-right: 1px solid #ececec;
    padding: 6px; font-weight: bold; color: #444444; }}
QTableWidget::item:selected, QListWidget::item:selected {{ background: #cce4f7; color: #000000; }}
QScrollArea {{ border: none; background: #ffffff; }}
QSplitter::handle {{ background: #e6e6e6; }}
QStatusBar {{ background: #f3f3f3; color: #666666; }}
QToolButton {{ border: none; padding: 4px 8px; border-radius: 6px; }}
QToolButton:hover {{ background: #e8e8e8; }}
QMenu {{ background: #ffffff; border: 1px solid #d0d0d0; }}
QMenu::item:selected {{ background: #cce4f7; color: #000000; }}
QCheckBox {{ spacing: 6px; }}
QProgressBar {{ border: 1px solid #d0d0d0; border-radius: 6px; background: #ffffff; text-align: center; height: 14px; }}
QProgressBar::chunk {{ background: {ACCENT}; border-radius: 5px; }}
"""


def apply(app):
    """套用到整個程式（含對話框）。"""
    app.setStyleSheet(STYLE)


def primary(*buttons):
    """標示主要按鈕（藍色實心）。"""
    for button in buttons:
        if button is None:
            continue
        button.setProperty('primary', True)
        button.style().unpolish(button)
        button.style().polish(button)
