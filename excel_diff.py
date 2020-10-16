# -*- coding: utf-8 -*-

import sys

import numpy
from PyQt5 import QtWidgets
from PyQt5.Qt import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from compare_method import *
from mainwindow import Ui_MainWindow

'''主程序继承QtWidgets.QMainWindow类和Ui_MainWindow类，主程序继承Ui_MainWindow可以使UI和逻辑剥离'''

global openfile_name_1, openfile_name_2, work_file_1, work_file_2, work_sheet_1, work_sheet_2
global strcount2_row, strcount3_row, strcount2_col, strcount3_col, final_matrix, strItem, rowCompareRes, rowCompareRes_t


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):

    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)  # 创建主界面对象
        Ui_MainWindow.__init__(self)  # 主界面对象初始化
        self.setupUi(self)  # 配置主界面对象

        # 信号与槽函数
        self.pushButton.clicked.connect(self.openFile_1)
        self.pushButton_2.clicked.connect(self.openFile_2)
        self.pushButton_3.clicked.connect(self.tableMarkColor)
        self.comboBox.activated.connect(self.comboboxSelectSheet)

        self.tableWidget_4.itemClicked.connect(self.colLocation)
        self.tableWidget_5.itemClicked.connect(self.rowLocation)
        self.tableWidget_6.itemClicked.connect(self.itemLocation)

        # 设置tableMarkColor()中选中单元格的默认颜色为灰色
        self.tableWidget.setStyleSheet("selection-background-color:rgb(154, 145, 106);")
        self.tableWidget_2.setStyleSheet("selection-background-color:rgb(154, 145, 106);")
        self.tableWidget_4.setStyleSheet("selection-background-color:rgb(154, 145, 106);")
        self.tableWidget_5.setStyleSheet("selection-background-color:rgb(154, 145, 106);")
        self.tableWidget_6.setStyleSheet("selection-background-color:rgb(154, 145, 106);")

        # 初始化tabwidget的名字和属性

        self.tableWidget_4.setRowCount(2)
        self.tableWidget_4.setColumnCount(1)

        newItem = QTableWidgetItem('新增')
        self.tableWidget_4.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('删除')
        self.tableWidget_4.setItem(1, 0, newItem)

        self.tableWidget_5.setRowCount(2)
        self.tableWidget_5.setColumnCount(1)

        newItem = QTableWidgetItem('新增')
        self.tableWidget_5.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('删除')
        self.tableWidget_5.setItem(1, 0, newItem)

        self.tableWidget_6.setRowCount(1)
        self.tableWidget_6.setColumnCount(3)
        newItem = QTableWidgetItem('坐标')
        self.tableWidget_6.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('旧值')
        self.tableWidget_6.setItem(0, 1, newItem)
        newItem = QTableWidgetItem('新值')
        self.tableWidget_6.setItem(0, 2, newItem)

    def openFile_1(self):
        global work_sheet_1, work_sheet_2
        '''打开原始文件1并将内容写入LineEdit中'''
        global openfile_name_1, work_file_1

        openfile_name_1 = QFileDialog.getOpenFileName(self, '选择文件', ' ', 'Excel Files(*.xlsx , *.xls)')[0]
        if openfile_name_1 != "":
            self.lineEdit.setText(openfile_name_1)
        else:
            return

        origin_file = xlrd.open_workbook(openfile_name_1)
        origin_file_names = origin_file.sheet_names()
        work_sheet_1 = origin_file.sheet_by_index(0)
        self.tableWidget.setRowCount(origin_file.sheet_by_index(0).nrows)
        self.tableWidget.setColumnCount(origin_file.sheet_by_index(0).ncols)
        for i in range(origin_file.sheet_by_index(0).nrows):
            for j in range(origin_file.sheet_by_index(0).ncols):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(origin_file.sheet_by_index(0).cell(i, j).value)))
                self.tableWidget.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

        work_file_1 = origin_file

    def openFile_2(self):
        global work_file_1, work_file_2, work_sheet_1, work_sheet_2
        '''打开原始文件2并将内容写入LineEdit_2中'''
        global openfile_name_2, work_file_2

        openfile_name_2 = QFileDialog.getOpenFileName(self, '选择文件', ' ', 'Excel Files(*.xlsx , *.xls)')[0]
        if openfile_name_2 != "":
            self.lineEdit_2.setText(openfile_name_2)
        else:
            return

        origin_file = xlrd.open_workbook(openfile_name_2)
        origin_file_names = origin_file.sheet_names()
        work_sheet_2 = origin_file.sheet_by_index(0)
        self.tableWidget_2.setRowCount(origin_file.sheet_by_index(0).nrows)
        self.tableWidget_2.setColumnCount(origin_file.sheet_by_index(0).ncols)
        for i in range(origin_file.sheet_by_index(0).nrows):
            for j in range(origin_file.sheet_by_index(0).ncols):
                self.tableWidget_2.setItem(i, j, QTableWidgetItem(str(origin_file.sheet_by_index(0).cell(i, j).value)))
                self.tableWidget_2.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

        work_file_2 = origin_file

        self.comboBox.clear()

        sameSheetNames = []
        for sheet in work_file_1.sheet_names():
            if sheet in work_file_2.sheet_names():
                sameSheetNames.append(sheet)
            else:
                continue
        for i in range(len(sameSheetNames)):
            self.comboBox.addItem(sameSheetNames[i])

    def comboboxSelectSheet(self):
        '''设置复选框中内容输出到tablewidget'''

        global openfile_name_1, work_file_1, work_file_2, work_sheet_1, work_sheet_2
        sameSheetNames = []
        for sh in work_file_1.sheet_names():
            if sh in work_file_2.sheet_names():
                sameSheetNames.append(sh)
            else:
                continue

        sheet1Current = work_file_1.sheet_by_name(self.comboBox.currentText())
        sheet2Current = work_file_2.sheet_by_name(self.comboBox.currentText())

        self.tableWidget.clear()
        self.tableWidget.setRowCount(sheet1Current.nrows)
        self.tableWidget.setColumnCount(sheet1Current.ncols)
        self.tableWidget_2.clear()
        self.tableWidget_2.setRowCount(sheet2Current.nrows)
        self.tableWidget_2.setColumnCount(sheet2Current.ncols)

        for i in range(sheet1Current.nrows):
            for j in range(sheet1Current.ncols):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(sheet1Current.cell(i, j).value)))
                self.tableWidget.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

        for i in range(sheet2Current.nrows):
            for j in range(sheet2Current.ncols):
                self.tableWidget_2.setItem(i, j, QTableWidgetItem(str(sheet2Current.cell(i, j).value)))
                self.tableWidget_2.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

    def tableMarkColor(self):

        '''给两个TableWidget的对比结果标记颜色并且将相关行列
        增删单元格改动的信号传递给后面的TableWidget 4 5 和 6'''

        global work_file_1, work_file_2, work_sheet_1, work_sheet_2, final_matrix, strItem

        sheet_old = work_file_1.sheet_by_name(self.comboBox.currentText())
        sheet_new = work_file_2.sheet_by_name(self.comboBox.currentText())
        list_old = sheetToList(sheet_old)
        list_new = sheetToList(sheet_new)

        rowCompareRes = rowCompare(list_old, list_new)[1]
        list_old_copy = list_old
        list_new_copy = list_new
        copy_old, copy_new = listMergeAll(list_old_copy, list_new_copy)
        list_old_t = transform(list_old)
        list_new_t = transform(list_new)
        rowCompareRes_t = rowCompare(list_old_t, list_new_t)[1]

        self.tableWidget.setRowCount(len(copy_old))
        self.tableWidget.setColumnCount(len(copy_old[0]))

        self.tableWidget_2.setRowCount(len(copy_new))
        self.tableWidget_2.setColumnCount(len(copy_new[0]))

        final_matrix = numpy.zeros((len(copy_old), len(copy_old[0])))

        for i in range(len(copy_old)):
            for j in range(len(copy_old[0])):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(copy_old[i][j])))
                self.tableWidget.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

        for i in range(len(copy_new)):
            for j in range(len(copy_new[0])):
                self.tableWidget_2.setItem(i, j, QTableWidgetItem(str(copy_new[i][j])))
                self.tableWidget_2.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j) % 26 + ord('A')))))

        for i in range(len(copy_old)):
            for j in range(len(copy_old[0])):
                if copy_old[i][j] == copy_new[i][j]:
                    final_matrix[i][j] = 0
                else:
                    final_matrix[i][j] = 1

        if len(final_matrix) == 1:
            rowCompareRes_t = []
        if len(final_matrix[0]) == 1:
            rowCompareRes = []

        countIncreaseRow, countDeleteRow, countIncreaseCol, countDeleteCol = 0, 0, 0, 0

        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 3:
                final_matrix[rowCompareRes[i][1], :] = 3
                countDeleteRow += 1
            else:
                final_matrix[rowCompareRes[i][1], :] = 2
                countIncreaseRow += 1

        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 3:
                final_matrix[:, rowCompareRes_t[i][1]] = 3
                countDeleteCol += 1
            else:
                final_matrix[:, rowCompareRes_t[i][1]] = 2
                countIncreaseCol += 1

        # 初始化tabwidget的名字和属性
        self.tableWidget_4.clear()
        self.tableWidget_4.setRowCount(2)
        self.tableWidget_4.setColumnCount(1)

        newItem = QTableWidgetItem('新增')
        self.tableWidget_4.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('删除')
        self.tableWidget_4.setItem(1, 0, newItem)

        self.tableWidget_5.clear()
        self.tableWidget_5.setRowCount(2)
        self.tableWidget_5.setColumnCount(1)

        newItem = QTableWidgetItem('新增')
        self.tableWidget_5.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('删除')
        self.tableWidget_5.setItem(1, 0, newItem)

        self.tableWidget_6.clear()
        self.tableWidget_6.setRowCount(1)
        self.tableWidget_6.setColumnCount(3)
        newItem = QTableWidgetItem('坐标')
        self.tableWidget_6.setItem(0, 0, newItem)
        newItem = QTableWidgetItem('旧值')
        self.tableWidget_6.setItem(0, 1, newItem)
        newItem = QTableWidgetItem('新值')
        self.tableWidget_6.setItem(0, 2, newItem)

        self.tableWidget_4.setRowCount(2)
        self.tableWidget_4.setColumnCount(max(countIncreaseCol, countDeleteCol) + 1)

        self.tableWidget_5.setRowCount(2)
        self.tableWidget_5.setColumnCount(max(countIncreaseRow, countDeleteRow) + 1)
        self.printColor()

        m, n = 0, 0
        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 3:
                m = m + 1
                self.tableWidget_5.setItem(1, (m), QTableWidgetItem(str(rowCompareRes[i][1] + 1)))
            elif rowCompareRes[i][0] == 2:
                n = n + 1
                self.tableWidget_5.setItem(0, (n), QTableWidgetItem(str(rowCompareRes[i][1])))

        m, n = 0, 0
        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 3:
                m = m + 1
                self.tableWidget_4.setItem(1, m, QTableWidgetItem(str(chr((rowCompareRes_t[i][1]) % 26 + ord('A')))))
            elif rowCompareRes_t[i][0] == 2:
                n = n + 1
                self.tableWidget_4.setItem(0, n,
                                           QTableWidgetItem(str(chr((rowCompareRes_t[i][1] - 1) % 26 + ord('A')))))

        if len(rowCompareRes_t) == 16:
            self.tableWidget_4.setRowCount(2)
            self.tableWidget_4.setColumnCount(1)

        m = 0
        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 2:
                m += 1
                for j in range(rowCompareRes_t[i][1], len(final_matrix[0])):
                    self.tableWidget.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j - m) % 26 + ord('A')))))
                self.tableWidget.setHorizontalHeaderItem(rowCompareRes_t[i][1], QTableWidgetItem(''))

        m = 0
        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 3:
                m += 1
                for j in range(rowCompareRes_t[i][1], len(final_matrix[0])):
                    self.tableWidget_2.setHorizontalHeaderItem(j, QTableWidgetItem(str(chr((j - m) % 26 + ord('A')))))
                self.tableWidget_2.setHorizontalHeaderItem(rowCompareRes_t[i][1], QTableWidgetItem(''))

        m = 0
        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 2:
                m += 1
                for j in range(rowCompareRes[i][1], len(final_matrix)):
                    self.tableWidget.setVerticalHeaderItem(j, QTableWidgetItem(str(j - m + 1)))
                self.tableWidget.setVerticalHeaderItem(rowCompareRes[i][1], QTableWidgetItem(''))

        m = 0
        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 3:
                m += 1
                for j in range(rowCompareRes[i][1], len(final_matrix)):
                    self.tableWidget_2.setVerticalHeaderItem(j, QTableWidgetItem(str(j - m + 1)))
                self.tableWidget_2.setVerticalHeaderItem(rowCompareRes[i][1], QTableWidgetItem(''))

        # 标记改动单元格的相关信息，追加到tablewidget_6中
        k = 0
        strItem = ''
        for i in range(len(final_matrix)):
            for j in range(len(final_matrix[0])):
                if final_matrix[i][j] == 1:
                    self.tableWidget.item(i, j).setBackground(QColor(255, 235, 83))
                    self.tableWidget_2.item(i, j).setBackground(QColor(255, 235, 83))
                    k = k + 1
                    self.tableWidget_6.setRowCount(k + 1)
                    m = '[' + str(chr((j) % 26 + ord('A'))) + ',' + str(i + 1) + ']'
                    strItem += str(i + 1)
                    strItem += str(j + 1)
                    self.tableWidget_6.setItem(k, 0, QTableWidgetItem(m))
                    self.tableWidget_6.setItem(k, 1, QTableWidgetItem(str(copy_old[i][j])))
                    self.tableWidget_6.setItem(k, 2, QTableWidgetItem(str(copy_new[i][j])))

        if len(rowCompareRes_t) == 16:
            countIncreaseCol = 0

        # 总览中信息输出，追加到各自label中
        if list_old == list_new:
            self.global_label.setText('文件是否有改动：否')
        else:
            self.global_label.setText('文件是否有改动：是')
        self.global_label_2.setText("单元格改动个数：" + str(k))
        self.global_label_3.setText("列增删：" + '增加 ' + str(countIncreaseCol) + ' 列，删除 ' + str(countDeleteCol) + ' 列')
        self.global_label_4.setText("行增删：" + '增加 ' + str(countIncreaseRow) + ' 行，删除 ' + str(countDeleteRow) + ' 行')

        # 记录统计信息中各行列增删情况
        global strcount2_row, strcount3_row, strcount2_col, strcount3_col
        strcount2_row = ''  # 行增
        strcount3_row = ''  # 行删
        strcount2_col = ''  # 列增
        strcount3_col = ''  # 列删

        # 行增删信息联动函数
        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 3:
                strcount3_row = strcount3_row + str(rowCompareRes[i][1])

        for i in range(len(rowCompareRes)):
            if rowCompareRes[i][0] == 2:
                strcount2_row = strcount2_row + str(rowCompareRes[i][1])

        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 3:
                strcount3_col = strcount3_col + str(rowCompareRes_t[i][1])

        for i in range(len(rowCompareRes_t)):
            if rowCompareRes_t[i][0] == 2:
                strcount2_col = strcount2_col + str(rowCompareRes_t[i][1])

        if len(rowCompareRes) == 0:
            self.tableWidget_4.setItem(0, 1, QTableWidgetItem("A"))

        if len(rowCompareRes_t) == 0:
            self.tableWidget_5.setItem(0, 1, QTableWidgetItem('1'))

    # 行增删信息联动函数
    def rowLocation(self):
        global strcount2_row, strcount3_row, final_matrix
        currentItem = self.tableWidget_5.currentItem()
        if currentItem.row() == 0:
            if currentItem.column() != 0:
                self.printColor()
                col_index = int(ord(strcount2_row[currentItem.column() - 1]) - ord('0'))
            else:
                return
        elif currentItem.row() == 1:
            if currentItem.column() != 0:
                self.printColor()
                col_index = int(ord(strcount3_row[currentItem.column() - 1]) - ord('0'))
            else:
                return

        for i in range(len(final_matrix[0])):
            self.tableWidget.item(col_index, i).setBackground(QColor(154, 145, 106))
            self.tableWidget_2.item(col_index, i).setBackground(QColor(154, 145, 106))

    # 列增删信息联动函数
    def colLocation(self):
        global strcount2_col, strcount3_col, final_matrix
        currentItem = self.tableWidget_4.currentItem()
        if currentItem.row() == 0:
            if currentItem.column() != 0:
                self.printColor()
                col_index = int(ord(strcount2_col[currentItem.column() - 1]) - ord('0'))
            else:
                return
        elif currentItem.row() == 1:
            if currentItem.column() != 0:
                self.printColor()
                col_index = int(ord(strcount3_col[currentItem.column() - 1]) - ord('0'))
            else:
                return

        for i in range(len(final_matrix)):
            self.tableWidget.item(i, col_index).setBackground(QColor(154, 145, 106))
            self.tableWidget_2.item(i, col_index).setBackground(QColor(154, 145, 106))

    # 单元格改动联动函数
    def itemLocation(self):
        global strItem
        currentItem = self.tableWidget_6.currentItem()
        if currentItem.column() == 0:
            if currentItem.row() != 0:
                self.printColor()
                index_row = int(ord(strItem[(currentItem.row() - 1) * 2]) - ord('1'))
                index_col = int(ord(strItem[(currentItem.row() - 1) * 2 + 1]) - ord('1'))
                self.tableWidget.item(index_row, index_col).setBackground(QColor(154, 145, 106))
                self.tableWidget_2.item(index_row, index_col).setBackground(QColor(154, 145, 106))
            else:
                return

    # 标记选定坐标的颜色
    def printColor(self):
        global final_matrix
        for i in range(len(final_matrix)):
            for j in range(len(final_matrix[0])):

                if final_matrix[i][j] == 1:
                    self.tableWidget.item(i, j).setBackground(QColor(255, 235, 83))
                    self.tableWidget_2.item(i, j).setBackground(QColor(255, 235, 83))

                if final_matrix[i][j] == 2:
                    self.tableWidget.item(i, j).setBackground(QColor(67, 101, 255))
                    self.tableWidget_2.item(i, j).setBackground(QColor(67, 101, 255))

                if final_matrix[i][j] == 3:
                    self.tableWidget.item(i, j).setBackground(QColor(255, 37, 95))
                    self.tableWidget_2.item(i, j).setBackground(QColor(255, 37, 95))


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()  # 创建QT对象
    window.show()  # QT对象显示
    sys.exit(app.exec_())
