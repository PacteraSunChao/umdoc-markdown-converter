#!/usr/bin/env python3
"""
UMDoc - 通用文档转 Markdown 桌面工具 (PySide6 + markitdown[all])
自动适配非 Excel 文件，隐藏工作表选择界面
环境由 launcher.py 保证，本文件直接启动 GUI
"""

import sys
import os
import re
import locale
import subprocess
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QListWidget, QListWidgetItem, QTextEdit, QPushButton,
    QLabel, QFileDialog, QMessageBox, QStatusBar, QGroupBox,
    QAbstractItemView, QCheckBox, QButtonGroup
)
from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QAction, QDragEnterEvent, QDropEvent

from markitdown import MarkItDown
import openpyxl


# ===================== 国际化 =====================
TEXTS = {
    "zh": {
        "app_title": "UMDoc → Markdown",
        "file_menu": "文件",
        "open_file": "打开文件...",
        "save_md": "保存 Markdown...",
        "exit": "退出",
        "lang_menu": "语言",
        "help_menu": "帮助",
        "about": "关于",
        "about_msg": "UMDoc - 通用文档转 Markdown 桌面工具",
        "file_panel": "文件",
        "drop_hint": "拖拽文件到此区域\n支持 Office/PDF/图片/音频/EPUB 等",
        "select_btn": "选择文件",
        "sheet_label": "工作表（可多选，仅 Excel 有效）：",
        "select_all": "全选工作表",
        "convert_btn": "转换为 Markdown",
        "preview_panel": "预览",
        "source_btn": "源码",
        "render_btn": "渲染",
        "save_to_btn": "保存到",
        "status_ready": "就绪 | 拖入或选择文件开始",
        "status_loaded": "已加载: {}",
        "status_converting": "正在转换...",
        "status_done": "转换完成",
        "status_saved": "已保存: {}",
        "warn_format": "格式错误",
        "warn_format_msg": "不支持的文件格式，请拖入常见文档格式",
        "info_no_content": "提示",
        "info_no_md": "尚未转换任何内容，请先转换文件",
        "save_file": "保存 Markdown",
        "save_filter": "Markdown 文件 (*.md)",
        "choose_folder": "选择保存文件夹",
        "no_sheet_selected": "未选择工作表",
        "no_sheet_selected_msg": "请至少选择一个工作表再转换",
        "error_save_failed": "保存失败: {}"
    },
    "ja": {
        "app_title": "UMDoc → Markdown",
        "file_menu": "ファイル",
        "open_file": "ファイルを開く...",
        "save_md": "Markdown を保存...",
        "exit": "終了",
        "lang_menu": "言語",
        "help_menu": "ヘルプ",
        "about": "バージョン情報",
        "about_msg": "UMDoc - ユニバーサルドキュメントMarkdown変換ツール",
        "file_panel": "ファイル",
        "drop_hint": "ここにファイルをドロップ\nOffice/PDF/画像/音声/EPUB 対応",
        "select_btn": "ファイルを選択",
        "sheet_label": "ワークシート（複数選択可、Excelのみ）：",
        "select_all": "すべて選択",
        "convert_btn": "Markdown に変換",
        "preview_panel": "プレビュー",
        "source_btn": "ソース",
        "render_btn": "レンダリング",
        "save_to_btn": "保存先",
        "status_ready": "準備完了 | ファイルをドロップまたは選択してください",
        "status_loaded": "読み込み中: {}",
        "status_converting": "変換中...",
        "status_done": "変換完了",
        "status_saved": "保存しました: {}",
        "warn_format": "形式エラー",
        "warn_format_msg": "サポートされていない形式です。一般的なドキュメントをドロップしてください",
        "info_no_content": "情報",
        "info_no_md": "変換された内容がありません。先に変換してください",
        "save_file": "Markdown を保存",
        "save_filter": "Markdown ファイル (*.md)",
        "choose_folder": "保存先フォルダを選択",
        "no_sheet_selected": "シート未選択",
        "no_sheet_selected_msg": "少なくとも1つのシートを選択してください",
        "error_save_failed": "保存失敗: {}"
    },
    "en": {
        "app_title": "UMDoc → Markdown",
        "file_menu": "File",
        "open_file": "Open File...",
        "save_md": "Save Markdown...",
        "exit": "Exit",
        "lang_menu": "Language",
        "help_menu": "Help",
        "about": "About",
        "about_msg": "UMDoc - Universal Document to Markdown Desktop Tool",
        "file_panel": "File",
        "drop_hint": "Drag and drop file here\nSupports Office/PDF/Image/Audio/EPUB etc.",
        "select_btn": "Select File",
        "sheet_label": "Worksheets (multiple selection, Excel only):",
        "select_all": "Select All",
        "convert_btn": "Convert to Markdown",
        "preview_panel": "Preview",
        "source_btn": "Source",
        "render_btn": "Render",
        "save_to_btn": "Save to",
        "status_ready": "Ready | Drop or select a file to start",
        "status_loaded": "Loaded: {}",
        "status_converting": "Converting...",
        "status_done": "Conversion complete",
        "status_saved": "Saved: {}",
        "warn_format": "Format Error",
        "warn_format_msg": "Unsupported file format. Please drop a common document format.",
        "info_no_content": "Info",
        "info_no_md": "No content converted yet. Please convert a file first.",
        "save_file": "Save Markdown",
        "save_filter": "Markdown Files (*.md)",
        "choose_folder": "Select Save Folder",
        "no_sheet_selected": "No Sheet Selected",
        "no_sheet_selected_msg": "Please select at least one worksheet.",
        "error_save_failed": "Save failed: {}"
    }
}


def get_system_language():
    """自动检测系统语言，返回 'zh', 'ja' 或 'en'"""
    try:
        lang, _ = locale.getlocale(category=locale.LC_MESSAGES)
        if lang:
            if lang.startswith("zh"): return "zh"
            if lang.startswith("ja"): return "ja"
    except: pass

    for var in ('LANGUAGE', 'LC_ALL', 'LC_MESSAGES', 'LANG'):
        val = os.environ.get(var)
        if val:
            if val.startswith("zh"): return "zh"
            if val.startswith("ja"): return "ja"
            break

    if sys.platform == "darwin":
        try:
            res = subprocess.run(["defaults", "read", "-g", "AppleLocale"],
                                 capture_output=True, text=True, timeout=5)
            if res.returncode == 0:
                loc = res.stdout.strip()
                if loc.startswith("zh"): return "zh"
                if loc.startswith("ja"): return "ja"
        except: pass

    return "en"


# ===================== 全局样式 =====================
STYLE_SHEET = """
QMainWindow { background-color: #F9FAFB; }
QGroupBox {
    background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px;
    margin-top: 16px; padding: 20px; font-size: 15px; font-weight: 600; color: #1F2937;
}
QGroupBox::title { subcontrol-origin: margin; left: 16px; padding: 0 8px; color: #4B5563; }
QLabel { color: #374151; font-size: 14px; }
QPushButton {
    background-color: #3B82F6; color: #FFFFFF; border: none; border-radius: 8px;
    padding: 10px 20px; font-size: 14px; font-weight: 600;
}
QPushButton:hover { background-color: #2563EB; }
QPushButton:pressed { background-color: #1D4ED8; }
QPushButton:disabled { background-color: #D1D5DB; color: #9CA3AF; }
QPushButton#toolBtn {
    background-color: transparent; color: #3B82F6; border: 1px solid #D1D5DB;
    padding: 6px 14px; border-radius: 6px; font-weight: 500;
}
QPushButton#toolBtn:hover { background-color: #EFF6FF; border-color: #3B82F6; }
QPushButton#toolBtn:checked { background-color: #DBEAFE; border-color: #3B82F6; color: #1E40AF; }
QListWidget {
    background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px;
    padding: 6px; outline: none; font-size: 14px; color: #1F2937;
}
QListWidget::item { padding: 8px 12px; border-radius: 6px; margin: 2px 0; }
QListWidget::item:selected { background-color: #DBEAFE; color: #1E40AF; }
QTextEdit, QPlainTextEdit {
    background-color: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px;
    padding: 10px; font-size: 14px; color: #1F2937; selection-background-color: #DBEAFE;
}
QCheckBox { spacing: 8px; font-size: 14px; color: #374151; }
QCheckBox::indicator {
    width: 20px; height: 20px; border: 2px solid #D1D5DB; border-radius: 4px;
    background-color: #FFFFFF;
}
QCheckBox::indicator:hover { border-color: #3B82F6; }
QCheckBox::indicator:checked { background-color: #3B82F6; border-color: #3B82F6; }
QCheckBox::indicator:checked:hover { background-color: #2563EB; border-color: #2563EB; }
QSplitter::handle { background-color: #E5E7EB; width: 2px; }
QStatusBar { background: #FFFFFF; border-top: 1px solid #E5E7EB; color: #6B7280; font-size: 13px; padding: 4px 12px; }
QMenuBar { background: #FFFFFF; border-bottom: 1px solid #E5E7EB; }
QMenuBar::item:selected { background: #F3F4F6; }
"""


# ===================== 拖拽标签 =====================
class DropLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumHeight(100)
        self.normal = "QLabel { border: 2px dashed #D1D5DB; border-radius: 12px; background: #FFFFFF; color: #6B7280; font-size: 14px; padding: 20px; }"
        self.hover = "QLabel { border: 2px solid #3B82F6; border-radius: 12px; background: #EFF6FF; color: #1F2937; font-size: 14px; padding: 20px; }"
        self.setStyleSheet(self.normal)

    def set_hover(self, h):
        self.setStyleSheet(self.hover if h else self.normal)


# ===================== 转换线程 =====================
class ConvertWorker(QThread):
    finished = Signal(str)

    def __init__(self, file_path: str, selected_sheets=None):
        super().__init__()
        self.file_path = file_path
        self.selected_sheets = selected_sheets  # None 表示全转，空列表才做过滤

    def run(self):
        try:
            md = MarkItDown()
            full_md = md.convert(self.file_path).text_content
            if self.selected_sheets is None or len(self.selected_sheets) == 0:
                self.finished.emit(full_md)
                return
            # 仅 Excel 使用过滤
            parts = re.split(r'(?=^##\s)', full_md, flags=re.MULTILINE)
            filtered = []
            for part in parts:
                m = re.match(r'^##\s+(.+)', part)
                if m and m.group(1).strip() in self.selected_sheets:
                    filtered.append(part)
            self.finished.emit('\n'.join(filtered) if filtered else "No selected sheet content found.")
        except Exception as e:
            self.finished.emit(f"Conversion failed: {e}")


def get_sheet_names(file_path: str) -> list:
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except:
        return []


# ===================== 主窗口 =====================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file = None
        self.raw_markdown = ""
        self.lang = get_system_language()
        self._updating_selection = False
        self.is_excel = False

        self.setAcceptDrops(True)
        screen = QApplication.primaryScreen().availableGeometry()
        w = int(screen.width() * 0.75)
        h = int(screen.height() * 0.85)
        self.resize(w, h)
        self.move((screen.width() - w) // 2, (screen.height() - h) // 2)

        self._create_actions()
        self._create_menus()
        self._create_central()
        self._create_statusbar()
        self.retranslate_ui()
        self._update_language_menu_checks()

    def _create_actions(self):
        self.open_action = QAction(self)
        self.open_action.triggered.connect(self.open_file_dialog)
        self.save_action = QAction(self)
        self.save_action.triggered.connect(self.save_markdown)
        self.exit_action = QAction(self)
        self.exit_action.triggered.connect(self.close)
        self.about_action = QAction(self)
        self.about_action.triggered.connect(lambda: QMessageBox.about(self, self.t("about"), self.t("about_msg")))
        self.lang_zh_action = QAction("中文", self, checkable=True, triggered=lambda: self.set_language("zh"))
        self.lang_ja_action = QAction("日本語", self, checkable=True, triggered=lambda: self.set_language("ja"))
        self.lang_en_action = QAction("English", self, checkable=True, triggered=lambda: self.set_language("en"))

    def _create_menus(self):
        bar = self.menuBar()
        self.file_menu = bar.addMenu("")
        self.file_menu.addAction(self.open_action)
        self.file_menu.addAction(self.save_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)
        self.lang_menu = bar.addMenu("")
        self.lang_menu.addAction(self.lang_zh_action)
        self.lang_menu.addAction(self.lang_ja_action)
        self.lang_menu.addAction(self.lang_en_action)
        self.help_menu = bar.addMenu("")
        self.help_menu.addAction(self.about_action)

    def _create_central(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        self.file_panel = self._create_file_panel()
        self.preview_panel = self._create_preview_panel()

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(2)
        splitter.addWidget(self.file_panel)
        splitter.addWidget(self.preview_panel)
        splitter.setSizes([320, 680])
        layout.addWidget(splitter)

    def _create_file_panel(self):
        card = QGroupBox()
        card.setObjectName("fileGroup")
        layout = QVBoxLayout(card)

        self.drop_label = DropLabel()
        self.select_btn = QPushButton()
        self.select_btn.clicked.connect(self.open_file_dialog)

        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.sheet_list.itemSelectionChanged.connect(self.on_sheet_selection_changed)

        self.select_all_check = QCheckBox()
        self.select_all_check.stateChanged.connect(self.on_select_all_toggled)

        self.convert_btn = QPushButton()
        self.convert_btn.clicked.connect(self.run_full_conversion)
        self.convert_btn.setEnabled(False)

        self.sheet_label = QLabel("")

        layout.addWidget(self.drop_label)
        layout.addWidget(self.select_btn)
        layout.addSpacing(12)
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.sheet_list)
        layout.addWidget(self.select_all_check)
        layout.addWidget(self.convert_btn)
        return card

    def _create_preview_panel(self):
        card = QGroupBox()
        card.setObjectName("previewGroup")
        layout = QVBoxLayout(card)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.source_btn = QPushButton()
        self.source_btn.setCheckable(True)
        self.source_btn.setChecked(True)
        self.source_btn.setObjectName("toolBtn")
        self.render_btn = QPushButton()
        self.render_btn.setCheckable(True)
        self.render_btn.setObjectName("toolBtn")

        grp = QButtonGroup(self)
        grp.addButton(self.source_btn, 0)
        grp.addButton(self.render_btn, 1)
        grp.buttonClicked.connect(self._on_view_changed)

        toolbar.addWidget(self.source_btn)
        toolbar.addWidget(self.render_btn)
        toolbar.addStretch()

        self.save_to_btn = QPushButton()
        self.save_to_btn.setObjectName("toolBtn")
        self.save_to_btn.clicked.connect(self.save_to_folder)
        toolbar.addWidget(self.save_to_btn)

        layout.addLayout(toolbar)

        self.source_view = QTextEdit()
        self.source_view.setReadOnly(True)
        self.rendered_view = QTextEdit()
        self.rendered_view.setReadOnly(True)
        self.rendered_view.hide()

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self.source_view)
        splitter.addWidget(self.rendered_view)
        splitter.setSizes([400, 200])
        layout.addWidget(splitter)
        return card

    def _create_statusbar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def t(self, key, *args):
        text = TEXTS[self.lang].get(key, key)
        return text.format(*args) if args else text

    def retranslate_ui(self):
        self.setWindowTitle(self.t("app_title"))
        self.file_menu.setTitle(self.t("file_menu"))
        self.open_action.setText(self.t("open_file"))
        self.save_action.setText(self.t("save_md"))
        self.exit_action.setText(self.t("exit"))
        self.lang_menu.setTitle(self.t("lang_menu"))
        self.help_menu.setTitle(self.t("help_menu"))
        self.about_action.setText(self.t("about"))
        self.file_panel.setTitle(self.t("file_panel"))
        self.drop_label.setText(self.t("drop_hint"))
        self.select_btn.setText(self.t("select_btn"))
        self.sheet_label.setText(self.t("sheet_label"))
        self.select_all_check.setText(self.t("select_all"))
        self.convert_btn.setText(self.t("convert_btn"))
        self.preview_panel.setTitle(self.t("preview_panel"))
        self.source_btn.setText(self.t("source_btn"))
        self.render_btn.setText(self.t("render_btn"))
        self.save_to_btn.setText(self.t("save_to_btn"))
        if self.current_file:
            self.status_bar.showMessage(self.t("status_loaded", os.path.basename(self.current_file)))
        else:
            self.status_bar.showMessage(self.t("status_ready"))

    def set_language(self, lang):
        if lang not in TEXTS: return
        self.lang = lang
        self._update_language_menu_checks()
        self.retranslate_ui()

    def _update_language_menu_checks(self):
        self.lang_zh_action.setChecked(self.lang == "zh")
        self.lang_ja_action.setChecked(self.lang == "ja")
        self.lang_en_action.setChecked(self.lang == "en")

    def _on_view_changed(self, btn):
        if btn == self.source_btn:
            self.source_view.show()
            self.rendered_view.hide()
        else:
            self.source_view.hide()
            self.rendered_view.show()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_label.set_hover(True)

    def dragLeaveEvent(self, event):
        self.drop_label.set_hover(False)

    def dropEvent(self, event: QDropEvent):
        self.drop_label.set_hover(False)
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            self.load_file(path)
            return
        QMessageBox.warning(self, self.t("warn_format"), self.t("warn_format_msg"))

    def open_file_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, self.t("open_file"), "",
                                              "All Files (*);;Documents (*.docx *.pptx *.xlsx *.xls *.pdf *.jpg *.png *.epub *.html *.csv *.json *.xml);;Excel (*.xlsx *.xls)")
        if path:
            self.load_file(path)

    def load_file(self, path):
        self.current_file = path
        self.status_bar.showMessage(self.t("status_loaded", os.path.basename(path)))
        self.is_excel = path.lower().endswith(('.xlsx', '.xls'))

        # 根据是否为 Excel 显示/隐藏工作表选择区域
        self.sheet_label.setVisible(self.is_excel)
        self.sheet_list.setVisible(self.is_excel)
        self.select_all_check.setVisible(self.is_excel)

        if self.is_excel:
            self.sheet_list.clear()
            sheets = get_sheet_names(path)
            if not sheets:
                sheets = ["Sheet1"]  # fallback
            for name in sheets:
                item = QListWidgetItem(name)
                item.setSelected(True)
                self.sheet_list.addItem(item)
            self._updating_selection = True
            self.select_all_check.setChecked(True)
            self._updating_selection = False
        else:
            self.sheet_list.clear()

        self.convert_btn.setEnabled(True)

    def on_select_all_toggled(self, state):
        if self._updating_selection: return
        self._updating_selection = True
        check = state == Qt.Checked
        for i in range(self.sheet_list.count()):
            self.sheet_list.item(i).setSelected(check)
        self._updating_selection = False

    def on_sheet_selection_changed(self):
        if self._updating_selection: return
        self._updating_selection = True
        count = self.sheet_list.count()
        if count == 0:
            self.select_all_check.setChecked(False)
        else:
            all_sel = all(self.sheet_list.item(i).isSelected() for i in range(count))
            self.select_all_check.setChecked(all_sel)
        self._updating_selection = False

    def get_selected_sheets(self):
        if not self.is_excel:
            return None  # 非 Excel 全量转换
        return [self.sheet_list.item(i).text() for i in range(self.sheet_list.count())
                if self.sheet_list.item(i).isSelected()]

    def run_full_conversion(self):
        if not self.current_file: return
        selected = self.get_selected_sheets()
        # 仅对 Excel 检查是否选择了工作表
        if self.is_excel and not selected:
            QMessageBox.warning(self, self.t("no_sheet_selected"), self.t("no_sheet_selected_msg"))
            return
        self.status_bar.showMessage(self.t("status_converting"))
        self.convert_btn.setEnabled(False)
        self.worker = ConvertWorker(self.current_file, selected)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.start()

    def on_conversion_finished(self, text):
        self.raw_markdown = text
        self.source_view.setPlainText(text)
        self.rendered_view.setMarkdown(text)
        self.status_bar.showMessage(self.t("status_done"))
        self.convert_btn.setEnabled(True)
        self.source_btn.setChecked(True)
        self.source_view.show()
        self.rendered_view.hide()

    def save_markdown(self):
        if not self.raw_markdown:
            QMessageBox.information(self, self.t("info_no_content"), self.t("info_no_md"))
            return
        path, _ = QFileDialog.getSaveFileName(self, self.t("save_file"), "output.md", self.t("save_filter"))
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.raw_markdown)
            self.status_bar.showMessage(self.t("status_saved", path))

    def save_to_folder(self):
        if not self.raw_markdown:
            QMessageBox.information(self, self.t("info_no_content"), self.t("info_no_md"))
            return
        folder = QFileDialog.getExistingDirectory(self, self.t("choose_folder"))
        if not folder: return
        base = Path(self.current_file).stem + ".md" if self.current_file else "output.md"
        path = os.path.join(folder, base)
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.raw_markdown)
            self.status_bar.showMessage(self.t("status_saved", path))
        except Exception as e:
            QMessageBox.critical(self, "Error", self.t("error_save_failed", str(e)))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE_SHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()