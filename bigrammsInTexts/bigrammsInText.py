import sys
from PyQt5.QtWidgets import *
import bigrammsGUI
import nltk
from nltk.corpus import stopwords
import pandas as pd
import os

nltk.download('stopwords')

class bigrammsInTextApp(QMainWindow, bigrammsGUI.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.searchButton.clicked.connect(self.searchBigramms)
        self.saveButton.clicked.connect(self.saveFile)

        self.tuple_list = list()
        self.table = ''



    def downloadFile(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open Files", "",
                                                     "Text files (*.txt)")
        files = [os.path.normpath(file) for file in files]

        text = ''
        for file in files:
            text += open(file, 'r').read().encode('utf-8').decode('utf-8-sig')
            text += '\n\n\n'

        stop_words = stopwords.words('russian')

        texts = text.replace(',', '.').replace('!', '.').replace('?', '.').replace(':', '.').replace(';', '.').replace(
            '\n', '.')
        texts = texts.split('.')
        texts = [line.strip().split() for line in texts]
        texts = [tuple([word.strip('«').strip('»').strip('[').strip(']')
                       .strip('(').strip(')').lower() for word in line
                        if word not in stop_words and len(word) > 1]) for line in texts]


        count = 0
        for line in texts:
            if len(line) >= 3:
                count += 1
                self.tuple_list.append([count, line])

        QMessageBox.about(self, "Message", "Текст загружен и обработан")


    def searchBigramms(self):
        table_path, _ = QFileDialog.getOpenFileName(self, "Open Files", "",
                                                "Excel files (*.xls *.xlsx)")
        table_path = os.path.normpath(table_path)
        bigramms = pd.read_excel(table_path)
        bigramms['№'] = range(1, len(bigramms) + 1)

        self.table = pd.DataFrame(columns=['Пары слов', '№ пары слов', 'Фраза из текста', '№ фразы'])

        self.progressBar.setMaximum(len(self.tuple_list))
        for i in range(len(self.tuple_list)):
            for j in range(len(bigramms)):
                if bigramms['Первое слово'][j].lower() in self.tuple_list[i][1]:
                    if bigramms['Второе слово'][j].lower() in self.tuple_list[i][1]:
                        self.table = self.table.append(pd.Series([[bigramms['Первое слово'][j], bigramms['Второе слово'][j]],
                                                        bigramms['№'][j], self.tuple_list[i][1], self.tuple_list[i][0]],
                                                       index=self.table.columns), ignore_index=True)
                        break

            self.progressBar.setValue(i + 1)
        QMessageBox.about(self, "Message", "Поиск завершен")


    def saveFile(self):
        table, _ = QFileDialog.getSaveFileName(self, "Save File", "",
                                                  "Excel Files (*.xls *.xlsx);;All Files (*)")
        table = os.path.normpath(table)
        self.table.to_excel(table)
        QMessageBox.about(self, "Message", "Таблица сохранена")


def main():
    app = QApplication(sys.argv)
    window = bigrammsInTextApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()
