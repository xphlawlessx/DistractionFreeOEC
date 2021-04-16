from PySide2.QtWidgets import QDialog, QVBoxLayout, QPushButton, QLabel


class PickSheetPopup(QDialog):
    def __init__(self, sheet_names, owner):
        QDialog.__init__(self)
        self.setModal(True)
        self.setWindowTitle('Sheet')
        sheet_names = sheet_names
        layout = QVBoxLayout()
        layout.addWidget(QLabel('Pick a sheet'))
        for name in sheet_names:
            if str(name).lower() == 'completed by':
                continue
            b = QPushButton(name)
            b.clicked.connect(lambda n=name, o=owner: self.pick_sheet(n, o))
            layout.addWidget(b)
        self.setLayout(layout)

    def pick_sheet(self, name, owner):
        owner.sheet = name
        self.close()
