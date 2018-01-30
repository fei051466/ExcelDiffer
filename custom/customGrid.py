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


class DiffSheetGrid(grid.Grid):
    """
    用于展示需要diff的sheet数据
    """
    def __init__(self, parent, sheet, info, flag):
        grid.Grid.__init__(self, parent, -1)
        self.table = DiffTable(sheet, info, flag)
        self.SetTable(self.table, True)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
        self.AutoSizeColumns(False)

        attrDel = grid.GridCellAttr()
        attrDel.SetBackgroundColour(wx.RED)
        attrAdd = grid.GridCellAttr()
        attrAdd.SetBackgroundColour(wx.BLUE)

        row = 0
        for i in info:
            type = i[0:1]
            if type == 'd':
                self.SetRowAttr(row, attrDel)
            elif type == 'a':
                self.SetRowAttr(row, attrAdd)
            row += 1


class DiffTable(grid.GridTableBase):
    """
    用户存放需要diff的sheet数据
    """
    def __init__(self, sheet, info, flag):
        grid.GridTableBase.__init__(self)

        # 暂时只考虑行变化
        self.colLabels = " ABCDEFGHIJKLMNOPQRSTUVWXYZ"[0:(sheet.ncols + 1)]

        self.data = []
        rowCount = sheet.nrows
        blankData = ['' for _ in range(self.GetNumberCols() + 1)]
        if flag == 'B':
            for i in info:
                diffType = i[0:1]
                if diffType == 's':
                    row = int(i[1:].split(":")[0])
                else:
                    row = int(i[1:])
                data = ['%d' % (row + 1)] + sheet.row_values(row)
                if diffType == 'd' or diffType == 's':
                    self.data.append(data)
                elif diffType == 'a':
                    self.data.append(blankData)
        elif flag == 'A':
            for i in info:
                diffType = i[0:1]
                if diffType == 's':
                    row = int(i[1:].split(":")[1])
                else:
                    row = int(i[1:])
                data = ['%d' % (row + 1)] + sheet.row_values(row)
                if diffType == 'a' or diffType == 's':
                    self.data.append(data)
                elif diffType == 'd':
                    self.data.append(blankData)

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
                self.data.append(['新增', int(i[1:]) + 1])
            elif i[0:1] == 'd':
                self.delCount += 1
                self.data.append(['删除', int(i[1:]) + 1])
        for d in self.data:
            print d

    def GetNumberCols(self):
        return 2

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
