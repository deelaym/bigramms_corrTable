import sys
from PyQt5.QtWidgets import *
import freqWordsGUI
import pandas as pd
import os
import re

# pyuic5 freqWordsGUI.ui -o freqWordsGUI.py

class frequencyOfWordsApp(QMainWindow, freqWordsGUI.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.saveButton.clicked.connect(self.saveFile)

        self.word_text_table = ''


    def downloadFile(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open Files", "",
                                                     "Text files (*.txt)")
        files = [os.path.normpath(file) for file in files]

        text_list = list()
        text_names = list()

        for file in files:
            text_list.append(open(file, 'r').read().encode('utf-8').decode('utf-8-sig'))
            name = os.path.basename(file)
            text_names.append(os.path.splitext(name)[0])

        text_list_wo_punct = list()
        long_text = ''
        for i in range(len(text_list)):
            text_list_wo_punct.append(re.sub('\W+|\d+|\_+', ' ', text_list[i]))
            long_text += text_list_wo_punct[i]

        text_list_wo_punct = [re.split('\s+', text.lower()) for text in text_list_wo_punct]
        long_text = long_text.lower()

        word_list = re.split('\s+', long_text)
        word_list = [word for word in word_list if len(word) > 1]
        word_list = sorted(list(set(word_list)))

        word_freq_list = list()

        self.progressBar.setMaximum(len(word_list))
        for i in range(len(word_list)):
            one_word_freq = list()
            for j in range(len(text_list_wo_punct)):
                freq_in_text = text_list_wo_punct[j].count(word_list[i])
                one_word_freq.append(freq_in_text)
            word_freq_list.append(one_word_freq)
            self.progressBar.setValue(i + 1)

        self.word_text_table = pd.DataFrame(columns=['Word', *text_names])

        self.word_text_table['Word'] = word_list
        self.word_text_table.iloc[:, 1:] = word_freq_list

        sum_col = self.word_text_table.iloc[:, 1:].sum(axis=1)
        mean_col = self.word_text_table.iloc[:, 1:].mean(axis=1)
        std_col = self.word_text_table.iloc[:, 1:].std(axis=1)

        self.word_text_table['sum'] = sum_col
        self.word_text_table['mean'] = mean_col
        self.word_text_table['std'] = std_col
        self.word_text_table['std / mean'] = std_col / mean_col

        QMessageBox.about(self, "Message", "Таблица составлена")



    def saveFile(self):
        table, _ = QFileDialog.getSaveFileName(self, "Save File", "",
                                                  "Excel Files (*.xls *.xlsx);;All Files (*)")
        table = os.path.normpath(table)

        len_cols = len(self.word_text_table.columns)

        with pd.ExcelWriter(table) as writer:
            self.word_text_table.iloc[:, :len_cols - 4].to_excel(writer, sheet_name='Лист1')
            self.word_text_table.to_excel(writer, sheet_name='Лист2')

        QMessageBox.about(self, "Message", "Таблица сохранена")


def main():
    app = QApplication(sys.argv)
    window = frequencyOfWordsApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()

