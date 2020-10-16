# -*-coding:utf-8 -*-
from copy import deepcopy

import numpy
import xlrd


def contrastRowMethod(work_sheet_1, work_sheet_2):
    count = 0  # 记录第几行不同
    list_1 = []
    for j in range(work_sheet_1.nrows):
        count_same_item = 0  # 记录每行中有几个单元格是相同的
        for i in range(len(work_sheet_1.row_values(0))):
            if work_sheet_1.row_values(j)[i] == work_sheet_2.row_values(j)[i]:
                count_same_item = count_same_item + 1
            else:
                continue
        if count_same_item >= len(work_sheet_1.row_values(0)) / 2:
            count = count + 1
        else:
            break
    list_1.append(count)

    flag = '0'
    count_sameitem_del = 0  # 记录删除行中有几个单元格是相同的
    count_sameitem_add = 0
    for i in range(len(work_sheet_1.row_values(0))):
        if work_sheet_1.row_values(count + 1)[i] == work_sheet_2.row_values(count)[i]:
            count_sameitem_del = count_sameitem_del + 1

        if work_sheet_1.row_values(count)[i] == work_sheet_2.row_values(count + 1)[i]:
            count_sameitem_add = count_sameitem_add + 1

    if count_sameitem_del >= len(work_sheet_1.row_values(0)) / 2:
        flag = 'del'
    if count_sameitem_add >= len(work_sheet_1.row_values(0)) / 2:
        flag = 'add'

    return (count, flag)


def contrastColMethod(work_sheet_1, work_sheet_2):
    count = 0  # 记录第几列不同
    for j in range(work_sheet_1.ncols):
        count_same_item = 0  # 记录每列中有几个单元格是相同的
        for i in range(len(work_sheet_1.col_values(0))):
            if work_sheet_1.col_values(j)[i] == work_sheet_2.col_values(j)[i]:
                count_same_item = count_same_item + 1
            else:
                continue
        if count_same_item >= len(work_sheet_1.col_values(0)) / 2:
            count = count + 1
        else:
            break

    flag = '0'
    count_sameitem_del = 0  # 记录删除行中有几个单元格是相同的
    count_sameitem_add = 0
    for i in range(len(work_sheet_1.col_values(0))):
        if work_sheet_1.col_values(count - 1)[i] == work_sheet_2.col_values(count - 1)[i]:
            count_sameitem_del = count_sameitem_del + 1

        if work_sheet_1.col_values(count)[i] == work_sheet_2.col_values(count)[i]:
            count_sameitem_add = count_sameitem_add + 1

    if count_sameitem_del >= len(work_sheet_1.col_values(0)) / 2:
        flag = 'del'
    if count_sameitem_add >= len(work_sheet_1.col_values(0)) / 2:
        flag = 'add'

    return (count, flag)


def lcsMethod(list_origin, list_compare):
    c = [[0 for i in range(len(list_compare) + 1)] for j in range(len(list_origin) + 1)]
    for i in range(1, len(c)):
        for j in range(1, len(c[0])):
            if list_origin[i - 1] == list_compare[j - 1]:
                flag = 1
            else:
                flag = 0
            c[i][j] = max(c[i - 1][j - 1] + flag, c[i - 1][j], c[i][j - 1])

    return c[-1][-1]


# 行增删比较算法
def rowCompare(list_origin, list_compare):
    # 找出非空起始位置
    for i in range(len(list_origin)):
        if list_origin[i] == ['' for _ in range(len(list_origin[i]))]:
            del list_origin[i]
            i = i - 1
        else:
            break
    for i in range(len(list_compare)):
        if list_compare[i] == ['' for _ in range(len(list_compare[i]))]:
            del list_compare[i]
            i = i - 1
        else:
            break
    c = [[[0, []] for i in range(len(list_compare) + 1)] for j in range(len(list_origin) + 1)]

    # 其中0代表未修改，1表示单元格修改，2表示增加，3表示删除
    # c[i][0]代表删除操作
    for i in range(1, len(c)):
        c[i][0][0] = c[i - 1][0][0] + 1
        c[i][0][1] = deepcopy(c[i - 1][0][1])
        c[i][0][1].append((3, i - 1))

    # c[0][j]代表增加操作
    for j in range(1, len(c[0])):
        c[0][j][0] = c[0][j - 1][0] + 1
        c[0][j][1] = deepcopy(c[0][j - 1][1])
        c[0][j][1].append((2, j - 1))

    for i in range(1, len(c)):
        for j in range(1, len(c[0])):
            # 相同值大于等于50%时不做增删操作
            if lcsMethod(list_origin[i - 1], list_compare[j - 1]) >= max(len(list_origin[i - 1]),
                                                                         len(list_compare[j - 1])) * 0.5:
                c[i][j][0] = min(c[i - 1][j - 1][0], c[i][j - 1][0] + 1, c[i - 1][j][0] + 1)
                if c[i][j][0] == c[i - 1][j - 1][0]:
                    c[i][j][1] = deepcopy(c[i - 1][j - 1][1])
                elif c[i][j][0] == c[i][j - 1][0] + 1:  # 增加
                    c[i][i][1] = deepcopy(c[i][j - 1][1])
                    c[i][j][1].append((2, i))
                else:  # 删除
                    c[i][i][1] = deepcopy(c[i - 1][j][1])
                    c[i][j][1].append((3, i - 1))
            else:
                c[i][j][0] = min(c[i][j - 1][0] + 1, c[i - 1][j][0] + 1)
                if c[i][j][0] == c[i][j - 1][0] + 1:
                    c[i][j][1] = deepcopy(c[i][j - 1][1])
                    c[i][j][1].append((2, i))
                else:
                    c[i][j][1] = deepcopy(c[i - 1][j][1])
                    c[i][j][1].append((3, i - 1))
    return c[-1][-1]


# 求列表中行的并集
def listMergeRow(list_origin, list_compare):
    shift = 0
    # 修复空表bug
    if len(list_origin) == 0:
        col_origin = 0
    else:
        col_origin = len(list_origin[0])
    if len(list_compare) == 0:
        col_compare = 0
    else:
        col_compare = len(list_compare[0])
    row_compare = rowCompare(list_origin, list_compare)
    for operator, row in row_compare[1]:
        if operator == 2:  # 行增加，list_origin插入''行
            list_origin.insert(row + shift, ['' for _ in range(col_origin)])
        elif operator == 3:
            list_compare.insert(row + shift, ['' for _ in range(col_compare)])
            shift = shift + 1
    return list_origin, list_compare


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


def compare_method(origin_sheet, compare_sheet):
    origin_sheet = xlrd.open_workbook('a.xls')
    compare_sheet = xlrd.open_workbook('b.xls')

    # 默认比较第一张sheet
    sheet1 = origin_sheet.sheet_by_index(0)
    sheet2 = compare_sheet.sheet_by_index(0)

    sheet_max_row = max(sheet1.nrows, sheet2.nrows)
    sheet_max_col = max(sheet1.ncols, sheet2.ncols)

    sheet_min_row = min(sheet1.nrows, sheet2.nrows)
    sheet_min_col = min(sheet1.ncols, sheet2.ncols)

    final_matrix = numpy.zeros((sheet_max_row, sheet_max_col))

    # 结果存储初始化，0-未改动，1-改动，2-增，3-删

    # 以第一张表格为基准对比第二张表格，先比较边缘处的行增删变化
    if sheet1.nrows > sheet2.nrows:  # 行预处理
        for i2 in range(sheet_min_row, sheet_max_row):
            final_matrix[i2, :] = 3
    else:
        for i2 in range(sheet_min_row, sheet_max_row):
            final_matrix[i2, :] = 2

    # 以第一张表格为基准对比第二张表格，先比较边缘处的列增删变化
    if sheet1.ncols > sheet2.ncols:  # 列预处理
        for i in range(sheet_min_col, sheet_max_col):
            final_matrix[:, i] = 3
    else:
        for i in range(sheet_min_col, sheet_max_col):
            final_matrix[:, i] = 2

    # 除去边缘部分处理内部列情况
    for j in range(0, sheet_min_col):
        sheet1_data_col = sheet1.col_values(j)
        sheet2_data_col = sheet2.col_values(j)

        # 设计一个对照列用来比较列的情况
        contrast_col = []
        for k in range(0, sheet_min_row):
            contrast_col.append('')

        if sheet1_data_col == sheet2_data_col:
            continue
        else:
            if sheet1_data_col == contrast_col:
                final_matrix[:, j] = 2
            elif sheet2_data_col == contrast_col:
                final_matrix[:, j] = 3
            else:
                for m in range(0, sheet_min_row):
                    if sheet1_data_col[m] == sheet2_data_col[m]:
                        continue
                    else:
                        if final_matrix[m][j] < 1:
                            final_matrix[m][j] = 1
                        else:
                            continue

    # 除去边缘部分处理内部行情况
    for j in range(sheet_min_row):
        sheet1_data_row = sheet1.row_values(j)
        sheet2_data_row = sheet2.row_values(j)

        # 设计一个对照列用来比较列的情况
        contrast_row = []
        for k in range(sheet_min_col):
            contrast_row.append('')

        if sheet1_data_row == sheet2_data_row:
            continue
        else:
            if sheet1_data_row == contrast_row:
                final_matrix[j, :] = 2
            elif sheet2_data_row == contrast_row:
                final_matrix[j, :] = 3
            else:
                for m in range(sheet_min_col):
                    if sheet1_data_row[m] == sheet2_data_row[m]:
                        continue
                    else:
                        if final_matrix[j][m] < 1:
                            final_matrix[j][m] = 1
                        else:
                            continue
    return final_matrix


# 转换列表的行和列
def transform(list):
    transform_list = []
    if len(list) == 0:
        return
    else:
        for i in range(len(list[0])):
            list2 = []
            for j in range(len(list)):
                list2.append(list[j][i])
            transform_list.append(list2)
    return transform_list


# 求列表中整个行和列的交并集情况
def listMergeAll(list_origin, list_compare):
    list_origin, list_compare = listMergeRow(list_origin, list_compare)
    col_origin = len(list_origin)
    col_compare = len(list_compare)
    list_origin = transform(list_origin)
    list_compare = transform(list_compare)
    col_diff = rowCompare(list_origin, list_compare)
    shift = 0
    for flag, row in col_diff[1]:
        if flag == 2:  # 行增加，list_origin插入''行
            list_origin.insert(row + shift, ['' for _ in range(col_origin)])
            shift = shift + 1
        elif flag == 3:
            list_compare.insert(row + shift, ['' for _ in range(col_compare)])
    list_origin = transform(list_origin)
    list_compare = transform(list_compare)
    return list_origin, list_compare


# 将sheet的数据转换成list
def sheetToList(sheet):
    newList = []
    for r in range(sheet.nrows):
        newList.append(sheet.row_values(r))
    return newList
