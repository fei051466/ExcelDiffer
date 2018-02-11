#!/usr/bin/env python
# -*- coding:utf-8 -*-

from wx import grid
from utils import generateColLabels, generateColIndex, getSheetData


class DiffTable(grid.GridTableBase):
    """
    用户存放需要diff的sheet数据
    """
    def __init__(self, sheet, rowInfo, colInfo, cellInfo, flag):
        grid.GridTableBase.__init__(self)

        self.colLabels = generateColLabels(sheet.ncols)

        self.mapping = {}
        self.mapping['row'] = []
        self.mapping['col'] = []
        self.mapping['cell'] = [None for row in range(sheet.nrows)]
        self.data = []
        self.initData =[]
        for row in range(sheet.nrows):
            rowData = [row + 1]
            rowInitData = []
            for col in range(sheet.ncols):
                value = sheet.cell(row, col).value
                if isinstance(value, float):
                    if int(value) == value:
                        value = int(value)
                rowData.append(value)
                rowInitData.append(value)
            self.data.append(rowData)
            self.initData.append(rowInitData)


        # 增加空白行
        blankRow = ['' for _ in range(self.GetNumberCols() + 1)]
        addCount = 0
        for i in range(len(rowInfo)):
            diffType = rowInfo[i][0:1]
            if (flag == 'B' and diffType == 'a') or \
                    (flag == 'A' and diffType == 'd'):
                # self.mapping.append(len(self.data) - 1)
                # self.data.insert((i + addCount), blankRow)
                self.data.insert(i, blankRow)
                addCount += 1
            if diffType != 's':
                self.mapping['row'].append(i)
            else:
                self.mapping['cell'][i - addCount] = [i for col in range(sheet.ncols)]

        # 行列转置
        self.data = map(list, zip(*self.data))

        # 增加空白列
        blankCol = ['' for _ in range(len(self.data))]
        addCount = 0
        for i in range(len(colInfo)):
            diffType = colInfo[i][0:1]
            if (flag == 'B' and diffType == 'a') or \
                    (flag == 'A' and diffType == 'd'):
                # self.mapping.append(len(self.data) - 1)
                self.data.insert((i + 1), blankCol)
                self.colLabels.insert((i + 1), '     ')
                addCount += 1
            if diffType != 's':
                self.mapping['col'].append(i)
            else:
                for row in self.mapping['cell']:
                    if not row:
                        continue
                    row[i - addCount] = (row[i - addCount], i + 1)

        # 行列转置
        self.data = map(list, zip(*self.data))

    def GetNumberCols(self):
        return len(self.colLabels)

    def GetNumberRows(self):
        return len(self.data)

    def IsEmptyCell(self, row, col):
        try:
            return not self.data[row][col]
        except IndexError:
            return True

    def GetValue(self, row, col):
        try:
            return self.data[row][col]
        except IndexError:
            return ''

    def SetValue(self, row, col, value):
        def innerSetValue(row, col, value):
            try:
                self.data[row][col] = value
            except IndexError:
                self.data.append([' '] * self.GetNumberCols())
                innerSetValue(row, col, value)
                msg = grid.GridTableMessage(self,
                            grid.GRIDTABLE_NOTIFY_ROWS_APPENDED,
                            1)
                self.GetView().ProcessTableMessage(msg)
        innerSetValue(row, col, value)

    def GetColLabelValue(self, col):
        return self.colLabels[col]

    def GetMapping(self):
        return self.mapping


class InfoTable(grid.GridTableBase):
    """
    用于提供diff信息的数据表格
    """
    def __init__(self, info, colLabels):
        grid.GridTableBase.__init__(self)
        self.colLabels = colLabels
        self.data = []
        self.addCount = 0
        self.delCount = 0
        self.info = info
        self.generateData()

    def generateData(self):
        pass

    def GetNumberCols(self):
        return len(self.colLabels)

    def GetNumberRows(self):
        return len(self.data)

    def IsEmptyCell(self, row, col):
        try:
            return not self.data[row][col]
        except IndexError:
            return True

    def GetValue(self, row, col):
        try:
            return self.data[row][col]
        except IndexError:
            return ''

    def SetValue(self, row, col, value):
        def innerSetValue(row, col, value):
            try:
                self.data[row][col] = value
            except IndexError:
                self.data.append([' '] * self.GetNumberCols())
                innerSetValue(row, col, value)
                msg = grid.GridTableMessage(self,
                            grid.GRIDTABLE_NOTIFY_ROWS_APPENDED,
                            1)
                self.GetView().ProcessTableMessage(msg)
        innerSetValue(row, col, value)

    def GetColLabelValue(self, col):
        return self.colLabels[col]

    def getCount(self):
        return self.addCount, self.delCount


class RowOrColInfoTable(InfoTable):
    """
    继承于InfoTable，用于提供行或列diff信息的数据表格
    """
    def generateData(self):
        for i in self.info:
            if i[0:1] == 'a':
                self.addCount += 1
                self.data.append(['新增', int(i[1:]) + 1])
            elif i[0:1] == 'd':
                self.delCount += 1
                self.data.append(['删除', int(i[1:]) + 1])


class CellInfoTable(InfoTable):
    """
    继承于InfoTable，用于提供单元格diff信息的数据表格
    """
    def __init__(self, info, colLabels, dataGridB, dataGridA):
        self.dataGridB = dataGridB
        self.dataGridA = dataGridA
        InfoTable.__init__(self, info, colLabels)

    def generateData(self):
        self.addCount = len(self.info)
        for i in self.info:
            rowBefore, colBefore = i[0]
            valueBefore = self.dataGridB.initData[rowBefore][colBefore]
            colBefore = generateColIndex(colBefore)
            rowAfter, colAfter= i[1]
            valueAfter = self.dataGridA.initData[rowAfter][colAfter]
            colAfter = generateColIndex(colAfter)

            self.data.append(['[%d,%s],[%d,%s]' % (rowBefore + 1, colBefore,
                            rowAfter + 1, colAfter), valueBefore, valueAfter])