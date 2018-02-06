#!/usr/bin/env python
# -*- coding:utf-8 -*-

letterList = list("ZABCDEFGHIJKLMNOPQRSTUVWXY")


def generateColIndex(num):
    """
    生成整数对应的字母列号，即：第1列对应A，第2列对应B，第27列对应AA
    :param num:
    :return:
    """
    res = ''
    while num != 0:
        res = letterList[num % 26] + res
        num = (num - 1) / 26

    return res


def generateColLabels(num):
    res = ['']
    for i in range(1, num + 1):
        res.append(generateColIndex(i))
    return res


def getSheetData(sheet):
    res = []
    for row in range(sheet.nrows):
        rowData = []
        for col in range(sheet.ncols):
            value = sheet.cell(row, col).value
            if not value:
                continue
            if isinstance(value, float):
                if int(value) == value:
                    value = int(value)
            rowData.append(value)
        res.append(rowData)
    return res

if __name__ == '__main__':
    # for i in range(1, 26 * 27 + 3):
    #    print generateColIndex(i)
    print generateColLabels(28)