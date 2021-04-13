import sys

from PySide2.QtGui import QIcon, QFont
from PySide2.QtWidgets import *

from Views.Modal import PickSheetPopup
from Views.Widgets import TabWidget, ScrollLabel
from XLSX.Parser import XLSXParser


class MainView(QDialog):
    def __init__(self):
        QDialog.__init__(self)
        self.parser = XLSXParser()
        self.sheet = ''
        self.filename = ''
        self.completed_comments = {}
        self.setWindowTitle('DFOEC')
        self.question_shown = False
        self.comment_question_label = ScrollLabel()
        self.layout = QGridLayout()
        self.tabs: TabWidget
        gb = QGroupBox()
        bot_menu = QHBoxLayout()

        self.init_comment_label()
        font_btn, save_btn, open_btn = self.make_buttons()

        self.layout.addWidget(self.comment_question_label, 0, 0)
        self.layout.addWidget(gb, 1, 0)
        bot_menu.addWidget(font_btn)
        bot_menu.addWidget(open_btn)
        bot_menu.addWidget(save_btn)

        self.setLayout(self.layout)
        gb.setLayout(bot_menu)

    def comment_or_question(self):
        self.question_shown = not self.question_shown
        self.comment_question_label.setText(
            self.parser.question if self.question_shown else self.parser.get_current_comment())

    def next_comment(self):
        self.parser.next_comment()
        if not self.question_shown:
            self.comment_question_label.setText(self.parser.get_current_comment())

    def init_comment_label(self):
        self.comment_question_label.setMinimumSize(400, 600)
        self.comment_question_label.setMaximumSize(800, 1200)
        self.comment_question_label.setFont(QFont('Source Code Pro', 16))
        self.comment_question_label.setText('Click the folder button below to open a codeframe')

    def make_tabs(self, code_list):
        self.tabs = TabWidget(code_list, self)
        self.tabs.setMinimumSize(400, 600)
        self.tabs.setMaximumSize(800, 1200)
        self.layout.addWidget(self.tabs, 0, 1, 1, 2)
        self.setLayout(self.layout)
        self.tabs.setFocus()
        self.next_comment()

    def make_buttons(self):
        font_btn = QPushButton()
        save_btn = QPushButton()
        open_btn = QPushButton()

        font_btn.setMaximumSize(100, 50)
        font_btn.setIcon(QIcon('images/font.png'))
        save_btn.setMaximumSize(100, 50)
        save_btn.setIcon(QIcon('images/save.png'))
        open_btn.setMaximumSize(100, 50)
        open_btn.setIcon(QIcon('images/open.png'))

        font_btn.clicked.connect(self.on_click_font)
        save_btn.clicked.connect(self.on_click_save)
        open_btn.clicked.connect(self.on_click_open)
        return font_btn, save_btn, open_btn

    def on_click_font(self):
        font, ok = QFontDialog.getFont()
        if ok:
            self.comment_question_label.setFont(font)

    def on_click_save(self):
        self.parser.save_workbook(self.filename, self.completed_comments)

    def on_click_open(self):
        dlg = QFileDialog()
        dlg.setFileMode(QFileDialog.AnyFile)
        # dlg.setFilter("(*.xlsx)")
        if dlg.exec_():
            self.filename = dlg.selectedFiles()[0]
            if self.filename:
                self.parser.load_workbook(self.filename)
                pop = PickSheetPopup(self.parser.get_sheet_names(), self.parser)
                pop.exec_()
                self.make_tabs(self.parser.get_codes(self.parser.sheet))


app = QApplication(sys.argv)
dialog = MainView()
dialog.show()
app.exec_()
