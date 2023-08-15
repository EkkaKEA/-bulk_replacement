import sys
import pandas as pd
from PyQt5.QtWidgets import *


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.table_path = None
        self.setWindowTitle("Массовая замена значений")
        self.setGeometry(100, 100, 400, 200)

        self.table_a_path = None
        self.table_b_path = None

        self.table_a_label = QLabel('Таблица в которой меняем')
        self.table_a_text_edit = QTextEdit()
        self.table_a_text_edit.setReadOnly(True)
        self.table_a_button = QPushButton('Load Table Tag')
        self.table_a_button.clicked.connect(self.load_table_a)

        self.table_label = QLabel('Таблица шаблонов')
        self.table_text_edit = QTextEdit()
        self.table_text_edit.setReadOnly(True)
        self.table_button = QPushButton('Load Table Wo')
        self.table_button.clicked.connect(self.load_table)


        self.label = QLabel("Нажмите кнопку для выполнения замены", self)
        self.label.setGeometry(50, 50, 300, 20)

        self.button = QPushButton("Выполнить замену", self)
        self.button.setGeometry(150, 100, 100, 30)
        self.button.clicked.connect(self.perform_replacement)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.table_a_label)
        layout.addWidget(self.table_a_text_edit)
        layout.addWidget(self.table_a_button)
        layout.addWidget(self.table_label)
        layout.addWidget(self.table_text_edit)
        layout.addWidget(self.table_button)


        layout.addWidget(self.button)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def load_table(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter('Excel files (*.xls*)')
        file_dialog.selectNameFilter('Excel files (*.xls*)')
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        if file_dialog.exec_():
            self.table_path = file_dialog.selectedFiles()[0]
            #self.table_label = QLabel(self.table_path)
            self.table_label.setText(self.table_path)
            #table1 = pd.read_excel(self.table_path)
            #self.table_text_edit.setText(table1.to_string(index=False))

    def load_table_a(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter('Excel files (*.xls*)')
        file_dialog.selectNameFilter('Excel files (*.xls*)')
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        if file_dialog.exec_():
            self.table_a_path = file_dialog.selectedFiles()[0]
            self.table_a_label.setText(self.table_a_path)
            #table = pd.read_excel(self.table_a_path)
            #self.table_a_text_edit.setText(table.to_string(index=False))

    def perform_replacement(self):
        # Загрузка файла "Tag" и чтение данных из него
        df_tag = pd.read_excel(self.table_a_path)

        # Загрузка файла "Wo" и чтение данных из него
        df_wo = pd.read_excel(self.table_path)

        # Создание словаря для хранения образцов и значений замены
        replace_dict = dict(zip(df_wo.iloc[:, 0], df_wo.iloc[:, 1]))

        # Массовая замена значений в файле "Tag" с использованием словаря
        df_tag.replace(replace_dict, inplace=True)

        # Сохранение изменений в файле "Tag"
        df_tag.to_excel('Tag_updated.xlsx', index=False)

        self.label.setText("Замена выполнена. Результаты сохранены в Tag_updated.xlsx")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())