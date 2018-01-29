#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid


class SheetGrid(grid.Grid):
    """
    用于展示sheet数据的网格
    """
    def __init__(self, parent, sheet):
        grid.Grid.__init__(self, parent, -1)
        self.rows, self.cols = sheet.nrows, sheet.ncols
        self.CreateGrid(self.rows, self.cols)
        for cell in sheet.merged_cells:
            self.SetCellSize(cell[0], cell[2],
                                 (cell[1] - cell[0]), (cell[3] - cell[2]))
            self.SetCellAlignment(cell[0], cell[2],
                                      wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        self.data = [["" for col in range(self.cols)]
                     for row in range(self.rows)]
        for row in range(self.rows):
            for col in range(self.cols):
                value = sheet.cell(row, col).value
                self.data[row][col] = value
                if not value:
                    continue
                if isinstance(value, float):
                    if int(value) == value:
                        value = int(value)
                value = str(value)
                self.SetCellValue(row, col, value)

    def getData(self):
        return self.data


class InfoGrid(grid.Grid):
    """
    用于展示diff信息的网格
    """
    def __init__(self, parent, info):
        grid.Grid.__init__(self, parent, -1)
        self.table = InfoTable(info)
        self.SetTable(self.table, True)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
        self.AutoSizeColumns(False)

    def getInfoMessage(self):
        return "共计新增%d行，删除%d行" % (self.table.getCount())


class InfoTable(grid.GridTableBase):
    """
    用于提供diff信息的表格
    """
    def __init__(self, info):
        grid.GridTableBase.__init__(self)
        self.colLabels = ['改动', '行号']
        self.data = []
        self.addCount = 0
        self.delCount = 0
        print info
        for i in info:
            if i[0:1] == 'a':
                self.addCount += 1
                self.data.append(['新增', i[1:]])
            elif i[0] == 'd':
                self.delCount += 1
                self.data.append(['删除', i[1:]])
        for d in self.data:
            print d

    def GetNumberCols(self):
        return 2

    def GetNumberRows(self):
        return len(self.data) + 1

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
