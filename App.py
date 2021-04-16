import sys

from PySide2.QtGui import QIcon, QFont, Qt
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

        self.number_label = QLabel('')
        self.number_label.setFont(QFont('Source Code Pro', 16))
        self.layout.addWidget(self.number_label, 1, 1)

        self.setLayout(self.layout)
        gb.setLayout(bot_menu)

    def set_number_label(self, value):
        self.number_label.setText(value)

    def comment_or_question(self):
        self.question_shown = not self.question_shown
        self.comment_question_label.setText(
            self.parser.question if self.question_shown else self.parser.get_current_comment())

    def check_comment(self):
        if self.parser.get_current_comment() in self.completed_comments.keys():
            self.tabs.display_loaded_data(self.completed_comments[self.parser.get_current_comment()])
            # Todo show row data with self.tabs

    def set_first_comment(self):
        self.parser.comment_index = self.parser.first_comment_index
        self.check_comment()
        if not self.question_shown:
            self.comment_question_label.setText(self.parser.get_current_comment())

    def next_comment(self):
        self.parser.next_comment()
        self.check_comment()
        if not self.question_shown:
            self.comment_question_label.setText(self.parser.get_current_comment())

    def prev_comment(self):
        self.parser.prev_comment()
        self.check_comment()
        if not self.question_shown:
            self.comment_question_label.setText(self.parser.get_current_comment())

    def init_comment_label(self):
        self.comment_question_label.setMinimumSize(400, 600)
        self.comment_question_label.setMaximumSize(800, 1200)
        self.comment_question_label.setFont(QFont('Source Code Pro', 16))
        self.comment_question_label.setText('Click the folder button below to open a codeframe')

    def make_tabs(self, code_list):
        self.parser.parse_old_data()
        self.tabs = TabWidget(code_list, self)
        self.tabs.setMinimumSize(400, 600)
        self.tabs.setMaximumSize(800, 1200)
        self.layout.addWidget(self.tabs, 0, 1, 1, 2)
        self.setLayout(self.layout)
        self.tabs.setFocus()
        self.set_first_comment()

    def make_buttons(self):
        font_btn = QPushButton('font')
        save_btn = QPushButton('save')
        open_btn = QPushButton('open')

        font_btn.setMaximumSize(100, 50)
        # font_btn.setIcon(QIcon('images/font.png'))
        save_btn.setMaximumSize(100, 50)
        # save_btn.setIcon(QIcon('images/save.png'))
        open_btn.setMaximumSize(100, 50)
        # open_btn.setIcon(QIcon('images/open.png'))

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
