import sys
from PyQt5.QtWidgets import *
import hapaxGUI
import pandas as pd
import os
import re
from tqdm import tqdm

#pyuic5 hapaxGUI.ui -o hapaxGUI.py

class hapaxToBigrammsApp(QMainWindow, hapaxGUI.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.searchButton.clicked.connect(self.searchBigramms)
        self.saveButton.clicked.connect(self.saveFile)

        self.texts = list()
        self.text_names = list()
        self.table = ''



    def downloadFile(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open Files", "",
                                                "Text files (*.txt)")
        files = [os.path.normpath(file) for file in files]

        text_list = list()
        self.text_names = list()

        for file in files:
            text_list.append(open(file, 'r').read().encode('utf-8').decode('utf-8-sig'))
            name = os.path.basename(file)
            self.text_names.append(os.path.splitext(name)[0])

        self.texts = list()
        for i in range(len(text_list)):
            one_text = re.sub('\[|\]|[-\'\"«»]+|\d+|', '', text_list[i])
            self.texts.append(re.sub('[.,!?:;…№()]+|\_+|\–+|\n+|\t+|\/+', '.', one_text))

        self.texts = [re.split('\.', text.lower()) for text in self.texts]
        self.texts = [[line.strip().split() for line in text] for text in self.texts]
        self.texts = [[line for line in text if len(line) > 1] for text in self.texts]

        QMessageBox.about(self, "Message", "Текст загружен и обработан")


    def searchBigramms(self):
        table_path, _ = QFileDialog.getOpenFileName(self, "Open Files", "",
                                                "Excel files (*.xls *.xlsx)")
        table_path = os.path.normpath(table_path)
        hapax_table = pd.read_excel(table_path)
        hapax = hapax_table['Word']

        self.table = pd.DataFrame(columns=['Первое слово', 'Второе слово', 'Фраза из текста', 'Название текста'])

        self.progressBar.setMaximum(len(self.texts))
        for i in tqdm(range(len(self.texts))):
            for j in range(len(self.texts[i])):
                switcher = False
                for n in range(len(hapax) - 1):
                    if hapax[n].lower() in self.texts[i][j]:
                        for m in range(n + 1, len(hapax)):
                            if hapax[m].lower() in self.texts[i][j]:
                                self.table = self.table.append(pd.Series([hapax[n].lower(), hapax[m].lower(),
                                                                self.texts[i][j], self.text_names[i]], index=self.table.columns),
                                                     ignore_index=True)

                                switcher = True
                                break
                    if switcher:
                        break
            self.progressBar.setValue(i + 1)

        QMessageBox.about(self, "Message", "Таблица биграмм составлена")


    def saveFile(self):
        table, _ = QFileDialog.getSaveFileName(self, "Save File", "",
                                                  "Excel Files (*.xls *.xlsx);;All Files (*)")
        table = os.path.normpath(table)
        self.table.to_excel(table)
        QMessageBox.about(self, "Message", "Таблица сохранена")


def main():
    app = QApplication(sys.argv)
    window = hapaxToBigrammsApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
