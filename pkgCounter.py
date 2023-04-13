import pandas as pd
from PyQt6.QtWidgets import QApplication, QLabel, QLineEdit, QPushButton, QMainWindow, QFileDialog
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__(parent=None)
        self.setWindowTitle('PACKAGE COUNTER')
        self.setFixedSize(300, 250)
        self.UIComponents()

    def UIComponents(self):
        self.fileMsg = QLabel(parent=self)
        self.fileMsg.setText("<b>CHOOSE EXCEL FILE</b>")
        self.fileMsg.setGeometry(10, 5, 250, 50)
        self.fileEdit = QLineEdit(self)
        self.fileEdit.setGeometry(10, 45, 200, 25)
        self.fileButton = QPushButton('Go', self)
        self.fileButton.setGeometry(225, 37, 55, 40)
        self.fileButton.clicked.connect(self.chooseFile)

        self.folderMsg = QLabel(parent=self)
        self.folderMsg.setText("<b>CHOOSE FOLDER DESTINATION</b>")
        self.folderMsg.setGeometry(10, 65, 250, 50)
        self.folderEdit = QLineEdit(self)
        self.folderEdit.setGeometry(10, 105, 200, 25)
        self.folderButton = QPushButton('Go', self)
        self.folderButton.setGeometry(225, 97, 55, 40)
        self.folderButton.clicked.connect(self.chooseFolder)

        self.inputMsg = QLabel(parent=self)
        self.inputMsg.setText("<b>NAME YOUR FILE</b>")
        self.inputMsg.setGeometry(10, 125, 250, 50)
        self.inputEdit = QLineEdit(self)
        self.inputEdit.setGeometry(10, 165, 200, 25)

        self.countButton = QPushButton('Count', self)
        self.countButton.setGeometry(125,205, 55,40)
        self.countButton.clicked.connect(self.countBtn_clicked)

    def chooseFile(self):
        file = QFileDialog.getOpenFileName(self, "Open File")
        file = file[0]
        if file:
            self.fileEdit.setText(file)

    def chooseFolder(self):
        folder = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        if folder:
            self.folderEdit.setText(folder)

    def countBtn_clicked(self):
        self.file = self.fileEdit.text()
        self.folder = self.folderEdit.text()
        input = self.inputEdit.text()
        df = pd.read_excel(self.file, usecols='K')
        count = df['CONSIGNEE'].value_counts()
        with open(f"{input}.csv", 'w') as file:
            count.to_csv(rf"{self.folder}/{input}.csv")


def main():
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
