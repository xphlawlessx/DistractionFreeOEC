from PySide2.QtCore import QItemSelectionModel
from PySide2.QtGui import QKeyEvent, QFont, Qt
from PySide2.QtWidgets import QTabWidget, QTabBar, QListWidget, QAbstractItemView, QListWidgetItem, QScrollArea, \
    QWidget, \
    QVBoxLayout, QLabel
from openpyxl.styles.colors import COLOR_INDEX


class TabWidget(QTabWidget):

    def __init__(self, code_list, owner):
        QTabWidget.__init__(self)
        self.CONFIRM = 'press enter again to confirm (backspace to cancel)'
        self.is_confirming = False
        self.owner = owner
        self.num_keys = (
            Qt.Key_0, Qt.Key_1, Qt.Key_2, Qt.Key_3, Qt.Key_4, Qt.Key_5, Qt.Key_6, Qt.Key_7, Qt.Key_8, Qt.Key_9)
        self.number = []
        colors = []
        for code in code_list:
            if code.color not in colors:
                colors.append(code.color)
        color_indexes = [15, 19, 22, 24, 26, 27, 45, 46, 47, 41, 42, 43, 44, 55, 56, 57]
        self.setAutoFillBackground(True)
        self.setTabBar(QTabBar())
        self.list_views = []
        for i in range(len(colors)):
            list_widget = QListWidget(self)
            list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
            list_widget.setFont(QFont('Source Code Pro', 16))
            self.list_views.append(list_widget)
            count = 0
            for code in code_list:
                if code.color != colors[i]:
                    continue
                    # create an item with a caption
                item = QListWidgetItem(f'{str(count)} - {code.code_name}')
                count += 1
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                # Add the item to the model
                # if list.count()
                list_widget.addItem(item)
            self.addTab(list_widget, str(i))
            list_widget.setAutoFillBackground(True)
            color = tuple(int(COLOR_INDEX[color_indexes[i]][j:j + 2], 16) for j in (2, 4, 6))

            list_widget.setStyleSheet("QWidget {{background-color: rgb{0};}}".format(color))
            self.convert_old_data()
            # todo load any completed rows 

    def redraw_number(self):
        self.owner.set_number_label(''.join([str(x) for x in self.number]))

    def set_number_ask_confirm(self):
        self.owner.set_number_label(self.CONFIRM)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Plus:
            self.setCurrentIndex((self.currentIndex() + 1) % self.count())
        elif event.key() == Qt.Key_Minus:
            num = self.currentIndex() - 1
            self.setCurrentIndex(num if num >= 0 else self.count() - 1)
        elif event.key() in self.num_keys:
            self.process_number_input(event.key())
        elif event.key() == Qt.Key_Enter:
            if self.number:
                self.enter_number()
            else:
                if not self.is_confirming:
                    self.set_number_ask_confirm()
                    self.is_confirming = True
                else:
                    self.is_confirming = False
                    self.number = []
                    self.save_comment()
                    self.owner.next_comment()
                    self.clear_codes()
        elif event.key() == Qt.Key_Slash:
            self.owner.comment_or_question()
        elif event.key() == Qt.Key_Backspace:
            if self.is_confirming:
                self.is_confirming = False
            else:
                if self.number:
                    self.delete_number()
                else:
                    self.clear_codes()
                    self.owner.prev_comment()
        if not self.is_confirming:
            self.redraw_number()

    def clear_codes(self):
        for i in range(self.count()):
            lv = self.list_views[i]
            for j in range(lv.count()):
                index = lv.model().index(j, 0)
                lv.selectionModel().select(index,
                                           QItemSelectionModel.Deselect | QItemSelectionModel.Rows)

    def process_number_input(self, key):
        self.number.append(self.num_keys.index(key))

    def convert_old_data(self):
        for cell in self.owner.parser.old_data:
            col = cell[0]
            tab = 0
            for lv in self.list_views:
                if col > lv.count() - 1:
                    col = lv.count() - col
                    tab += 1
            if cell[1] in self.owner.completed_comments.keys():
                self.owner.completed_comments[cell[1]].extend([(tab, col)])
            else:
                self.owner.completed_comments[cell[1]] = [(tab, col)]

    def display_loaded_data(self, loaded_list):
        for loaded in loaded_list:
            if isinstance(loaded[1], str):
                lv = self.list_views[loaded[0]]
                index = lv.model().index(int(loaded[1].split('-')[0]), 0)
                lv.selectionModel().select(index,
                                           QItemSelectionModel.Select | QItemSelectionModel.Rows)
            else:
                lv = self.list_views[loaded[0]]
                index = lv.model().index(int(loaded[1]), 0)
                lv.selectionModel().select(index,
                                           QItemSelectionModel.Select | QItemSelectionModel.Rows)

    def enter_number(self):
        i = self.currentIndex()
        lv = self.list_views[i]
        num = int(''.join([str(n) for n in self.number]))
        index = lv.model().index(num, 0)
        lv.selectionModel().select(index,
                                   QItemSelectionModel.Toggle | QItemSelectionModel.Rows)
        self.number = []

    def delete_number(self):
        self.number.pop()

    def save_comment(self):
        results = []
        for i in range(self.count()):
            lv = self.list_views[i]
            results.extend([
                (i, row.data()) for row in
                lv.selectionModel().selectedRows()])
        if results:
            self.owner.completed_comments[self.owner.comment_question_label.label.text()] = results


# class for scrollable label
class ScrollLabel(QScrollArea):

    def __init__(self, *args, **kwargs):
        QScrollArea.__init__(self, *args, **kwargs)

        self.setWidgetResizable(True)
        content = QWidget(self)
        self.setWidget(content)
        lay = QVBoxLayout(content)
        self.label = QLabel(content)
        self.label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.label.setWordWrap(True)
        lay.addWidget(self.label)

    def setText(self, text):
        self.label.setText(text)
