from PySide2.QtCore import QItemSelectionModel
from PySide2.QtGui import QKeyEvent, QFont, Qt
from PySide2.QtWidgets import QTabWidget, QTabBar, QListWidget, QAbstractItemView, QListWidgetItem, QScrollArea, \
    QWidget, \
    QVBoxLayout, QLabel
from openpyxl.styles.colors import COLOR_INDEX


class TabWidget(QTabWidget):
    def __init__(self, code_list, owner):
        QTabWidget.__init__(self)
        self.owner = owner
        self.num_keys = (
            Qt.Key_0, Qt.Key_1, Qt.Key_2, Qt.Key_3, Qt.Key_4, Qt.Key_5, Qt.Key_6, Qt.Key_7, Qt.Key_8, Qt.Key_9)
        colors = []
        for code in code_list:
            if code.color not in colors:
                colors.append(code.color)
        color_indexes = [15, 19, 22, 24, 26, 27, 45, 46, 47, 41, 42, 43, 44, 55, 56, 57]
        self.setAutoFillBackground(True)
        self.setTabBar(QTabBar())
        self.list_views = []
        color_map = {}
        for i in range(len(colors)):
            list = QListWidget(self)
            list.setSelectionMode(QAbstractItemView.ExtendedSelection)
            list.setFont(QFont('Source Code Pro', 16))
            self.list_views.append(list)
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
                list.addItem(item)
            self.addTab(list, str(i))
            list.setAutoFillBackground(True)
            print(colors[i])
            if i not in color_map.keys():
                #todo map real colors to color indicies
                pass
            color = tuple(int(COLOR_INDEX[color_indexes.pop()][i:i + 2], 16) for i in (2, 4, 6))
            
            list.setStyleSheet("QWidget {{background-color: rgb{0};}}".format(color))
            # palette = list.palette()
            # palette.setColor(list.backgroundRole(), QColor(*colors[i]))
            # list.setPalette(palette)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Plus:
            self.setCurrentIndex((self.currentIndex() + 1) % self.count())
        elif event.key() == Qt.Key_Minus:
            num = self.currentIndex() - 1
            self.setCurrentIndex(num if num >= 0 else self.count() - 1)
        elif event.key() in self.num_keys:
            self.process_number_input(event.key())
        elif event.key() == Qt.Key_Enter:
            self.save_comment()
            self.owner.next_comment()
            self.clear_codes()
        elif event.key() == Qt.Key_Slash:
            self.owner.comment_or_question()

    def clear_codes(self):
        for i in range(self.count()):
            # i = self.currentIndex()
            lv = self.list_views[i]
            for key in self.num_keys:
                index = lv.model().index(self.num_keys.index(key), 0)
                lv.selectionModel().select(index,
                                           QItemSelectionModel.Deselect | QItemSelectionModel.Rows)

    def process_number_input(self, key):
        i = self.currentIndex()
        lv = self.list_views[i]
        index = lv.model().index(self.num_keys.index(key), 0)
        lv.selectionModel().select(index,
                                   QItemSelectionModel.Toggle | QItemSelectionModel.Rows)

    def save_comment(self):
        results = []
        for i in range(self.count()):
            lv = self.list_views[i]
            results.extend([
                row.data() for row in
                lv.selectionModel().selectedRows()])
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

        # making label multi-line
        self.label.setWordWrap(True)

        # adding label to the layout
        lay.addWidget(self.label)

    # the setText method
    def setText(self, text):
        # setting text to the label
        self.label.setText(text)
