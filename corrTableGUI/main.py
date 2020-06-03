import sys
from PyQt5.QtWidgets import *
import corrTableGUI
import pandas as pd
import numpy as np
import math
from tqdm import tqdm
import os

#pyuic5 /home/dis/Documents/pasha/bigramms_corrTable/corrTableGUI/corrTableGUI.ui -o /home/dis/Documents/pasha/bigramms_corrTable/corrTableGUI/corrTableGUI.py


class corrTableApp(QMainWindow, corrTableGUI.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.downloadButton.clicked.connect(self.downloadFile)
        self.calcButton.clicked.connect(self.calculateCorr)
        self.saveButton.clicked.connect(self.saveFile)
        self.download_sheet.clicked.connect(self.downloadSheet)

        self.data = ''
        self.corrTable = ''
        self.file = ''


    def downloadFile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file, _ = QFileDialog.getOpenFileName(self, "Open Files", "",
                                                     "Excel Files (*.xls *.xlsx)", options=options)
        self.file = os.path.normpath(file)


    def downloadSheet(self):
        sheet = ''
        self.sheet.append(sheet)
        sheet = self.sheet.toPlainText().strip()

        if sheet == '':
            sheet = "Лист1"
        self.data = pd.read_excel(self.file, sheet_name=sheet)
        QMessageBox.about(self, "Message", "Таблица загружена")


    def calculateCorr(self):
        data_sumPrim_sumSq = self.data.copy()
        datalist = np.array(self.data.select_dtypes(include='int64'))
        num_rows = len(datalist)
        num_columns = len(datalist[0])

        sumPrim_list = list()
        sumSq_list = list()
        sum_Prim = 0
        sum_Sq = 0
        for i in range(0, num_rows):
            for j in range(0, num_columns):
                sum_Prim += datalist[i, j]
                sum_Sq += datalist[i, j] ** 2

            sumPrim_list.append(sum_Prim)
            sumSq_list.append(sum_Sq)
            sum_Prim = 0
            sum_Sq = 0

        data_sumPrim_sumSq['Сумма по строке'] = sumPrim_list
        data_sumPrim_sumSq['Сумма квадратов по строке'] = sumSq_list

        nums = range(1, num_rows + 1)
        data_sumPrim_sumSq['№'] = nums

        strInit = ''
        strFin = ''
        corrCritical = ''

        self.strInit.append(strInit)
        strInit = int(self.strInit.toPlainText().strip())
        self.strFin.append(strFin)
        strFin = int(self.strFin.toPlainText().strip())
        self.corrCritical.append(corrCritical)
        corrCritical = float(self.corrCritical.toPlainText().strip())

        self.corrTable = pd.DataFrame(
            columns=['numBigramm', *self.data.columns[1:], '№ 1-го слова', 'Первое слово', '№ 2-го слова', 'Второе слово',
                     'corr'])

        numBigramm = 0
        sumXY = 0

        self.progressBar.setMaximum(strFin - 1)
        for i1 in tqdm(range(strInit, strFin - 1)):
            for i2 in range(i1 + 1, strFin):
                for j in range(0, num_columns):
                    sumXY += datalist[i1, j] * datalist[i2, j]

                up = num_rows * sumXY - data_sumPrim_sumSq.loc[i1, 'Сумма по строке'] * data_sumPrim_sumSq.loc[
                    i2, 'Сумма по строке']
                downX = num_rows * data_sumPrim_sumSq.loc[i1, 'Сумма квадратов по строке'] - data_sumPrim_sumSq.loc[
                    i1, 'Сумма по строке'] ** 2
                downY = num_rows * data_sumPrim_sumSq.loc[i2, 'Сумма квадратов по строке'] - data_sumPrim_sumSq.loc[
                    i2, 'Сумма по строке'] ** 2
                corr = up / math.sqrt(downX * downY)

                sumXY = 0

                if corr > corrCritical:
                    numBigramm += 1
                    self.corrTable = self.corrTable.append(pd.Series(
                        [numBigramm, *self.data.iloc[i1, 1:].values, data_sumPrim_sumSq.loc[i1, '№'], self.data.iloc[i1, 0],
                         data_sumPrim_sumSq.loc[i2, '№'], self.data.iloc[i2, 0], corr], index=self.corrTable.columns),
                        ignore_index=True)
            self.progressBar.setValue(i1+1)
        QMessageBox.about(self, "Message", "Расчет окончен")

    def saveFile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        table, _ = QFileDialog.getSaveFileName(self, "Save File", "",
                                                  "Excel Files (*.xls *.xlsx);;All Files (*)")
        table = os.path.normpath(table)
        self.corrTable.to_excel(table)
        QMessageBox.about(self, "Message", "Таблица сохранена")


def main():
    app = QApplication(sys.argv)
    window = corrTableApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()

