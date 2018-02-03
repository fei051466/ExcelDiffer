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
    def __init__(self, parent, sheet, rowInfo, colInfo, cellInfo, flag):
        grid.Grid.__init__(self, parent, -1)
        self.table = DiffTable(sheet, rowInfo, colInfo, cellInfo, flag)
        self.SetTable(self.table, True)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
        self.AutoSizeColumns(False)
        self.mapping = self.table.GetMapping()
        self.initData = self.table.initData

        for row in range(len(rowInfo)):
            diffType = rowInfo[row][0:1]
            if diffType == 'd':
                # attr不能复用，每行都需要独自的attr
                attrDel = grid.GridCellAttr()
                attrDel.SetBackgroundColour(wx.RED)
                self.SetRowAttr(row, attrDel)
            elif diffType == 'a':
                attrAdd = grid.GridCellAttr()
                attrAdd.SetBackgroundColour(wx.BLUE)
                self.SetRowAttr(row, attrAdd)
        for row in range(len(colInfo)):
            diffType = colInfo[row][0:1]
            if diffType == 'd':
                # attr不能复用，每行都需要独自的attr
                attrDel = grid.GridCellAttr()
                attrDel.SetBackgroundColour(wx.RED)
                self.SetColAttr(row + 1, attrDel)
            elif diffType == 'a':
                attrAdd = grid.GridCellAttr()
                attrAdd.SetBackgroundColour(wx.BLUE)
                self.SetColAttr(row + 1, attrAdd)

        cellMapping = self.table.mapping['cell']
        print 'cellMapping', cellMapping
        for info in cellInfo:
            if flag == 'B':
                row, col = cellMapping[info[0][0]][info[0][1]]
            elif flag == 'A':
                row, col = cellMapping[info[1][0]][info[1][1]]
            self.SetCellBackgroundColour(row, col, wx.YELLOW)

    def GetMapping(self):
        return self.mapping


class DiffTable(grid.GridTableBase):
    """
    用户存放需要diff的sheet数据
    """
    def __init__(self, sheet, rowInfo, colInfo, cellInfo, flag):
        print '**DiffTable** colInfo', colInfo
        grid.GridTableBase.__init__(self)

        self.colLabels = list(" ABCDEFGHIJKLMNOPQRSTUVWXYZ"[0:(sheet.ncols + 1)])

        self.mapping = {}
        self.mapping['row'] = []
        self.mapping['col'] = []
        self.mapping['cell'] = [None for row in range(sheet.nrows)]
        self.data = [['%d' % (row + 1)] + sheet.row_values(row)
                     for row in range(sheet.nrows)]
        self.initData = [sheet.row_values(row) for row in range(sheet.nrows)]

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
                self.mapping['cell'][i] = [(i + addCount) for col in range(sheet.ncols)]

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
                    row[i] = (row[i], i + addCount + 1)

        # 行列转置
        self.data = map(list, zip(*self.data))

        '''
        print '**DiffTable** cellMapping', cellMapping
        for info in cellInfo:
            mappingInfo = cellMapping[info[0][0]][info[0][1]] + \
                            cellMapping[info[1][0]][info[1][1]]
            self.mapping['cell'].append(mappingInfo)
        '''

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


class InfoGrid(grid.Grid):
    """
    用于展示diff信息的网格
    """
    def __init__(self, parent, info, dataGridB, dataGridA):
        grid.Grid.__init__(self, parent, -1)
        self.info = info
        self.dataGridB = dataGridB
        self.dataGridA = dataGridA
        self.table = None
        self.initTable(info)
        self.SetTable(self.table, True)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
        self.AutoSizeColumns(False)

        self.Bind(grid.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)

    def initTable(self, info):
        pass

    def getInfoMessage(self):
        pass

    def OnCellLeftClick(self, evt):
        pass


class RowInfoGrid(InfoGrid):
    def getInfoMessage(self):
        return "共计新增%d行，删除%d行" % (self.table.getCount())

    def initTable(self, info):
        self.table = RowOrColInfoTable(info, ["改动", "行号"])

    def OnCellLeftClick(self, evt):
        row = evt.GetRow()
        diffType = self.GetCellValue(row, 0)
        if diffType == u'删除':
            hightlightRow = self.dataGridB.GetMapping()['row'][row]
        elif diffType == u'新增':
            hightlightRow = self.dataGridA.GetMapping()['row'][row]

        self.dataGridB.SelectRow(hightlightRow)
        self.dataGridA.SelectRow(hightlightRow)


class ColInfoGrid(InfoGrid):
    def getInfoMessage(self):
        return "共计新增%d列，删除%d列" % (self.table.getCount())

    def initTable(self, info):
        self.table = RowOrColInfoTable(info, ["改动", "列号"])

    def OnCellLeftClick(self, evt):
        row = evt.GetRow()
        diffType = self.GetCellValue(row, 0)
        if diffType == u'删除':
            hightlightCol = self.dataGridB.GetMapping()['col'][row] + 1
        elif diffType == u'新增':
            hightlightCol = self.dataGridA.GetMapping()['col'][row] + 1

        self.dataGridB.SelectCol(hightlightCol)
        self.dataGridA.SelectCol(hightlightCol)


class CellInfoGrid(InfoGrid):
    def getInfoMessage(self):
        return "共计%d个单元格修改过" % self.table.getCount()[0]

    def initTable(self, info):
        self.table = CellInfoTable(info, ["坐标", "旧值", "新值"],
                                   self.dataGridB, self.dataGridA)

    def OnCellLeftClick(self, evt):
        row = evt.GetRow()

        cellMappingBefore = self.dataGridB.table.mapping['cell']
        infoBefore = self.info[row]
        rowBefore, colBefore = cellMappingBefore[infoBefore[0][0]][infoBefore[0][1]]
        self.dataGridB.SelectBlock(rowBefore, colBefore, rowBefore, colBefore)

        cellMappingAfter = self.dataGridA.table.mapping['cell']
        infoAfter = self.info[row]
        rowAfter, colAfter = cellMappingAfter[infoAfter[0][0]][infoAfter[0][1]]
        self.dataGridA.SelectBlock(rowAfter, colAfter, rowAfter, colAfter)


class InfoTable(grid.GridTableBase):
    """
    用于提供diff信息的表格
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

    def generateData(self):
        for i in self.info:
            if i[0:1] == 'a':
                self.addCount += 1
                self.data.append(['新增', int(i[1:]) + 1])
            elif i[0:1] == 'd':
                self.delCount += 1
                self.data.append(['删除', int(i[1:]) + 1])


class CellInfoTable(InfoTable):

    def __init__(self, info, colLabels, dataGridB, dataGridA):
        self.dataGridB = dataGridB
        self.dataGridA = dataGridA
        InfoTable.__init__(self, info, colLabels)

    def generateData(self):
        self.addCount  = len(self.info)
        for i in self.info:
            rowBefore, colBefore = i[0]
            valueBefore = self.dataGridB.initData[rowBefore][colBefore]
            colBefore = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[colBefore]
            rowAfter, colAfter= i[1]
            valueAfter = self.dataGridA.initData[rowAfter][colAfter]
            colAfter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[colAfter]

            self.data.append(['[%d,%s],[%d,%s]' % (rowBefore + 1, colBefore,
                            rowAfter + 1, colAfter), valueBefore, valueAfter])
