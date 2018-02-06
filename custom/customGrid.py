#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid
from customTable import DiffTable, RowOrColInfoTable, CellInfoTable


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
        grid.Grid.__init__(self, parent, -1, size=(450, 200))
        self.table = DiffTable(sheet, rowInfo, colInfo, cellInfo, flag)
        self.SetTable(self.table, True)
        self.SetRowLabelSize(0)
        self.SetMargins(0, 0)
        self.AutoSizeColumns(False)
        self.mapping = self.table.GetMapping()
        self.initData = self.table.initData
        self.flag = flag

        # 逐个设置单元格只读
        # 没找到其他便捷方法
        for row in range(self.table.GetRowsCount()):
            for col in range(self.table.GetColsCount()):
                self.SetReadOnly(row, col)

        for row in range(len(rowInfo)):
            diffType = rowInfo[row][0:1]
            if diffType == 'd':
                # attr不能复用，每行都需要独自的attr
                attrDel = grid.GridCellAttr()
                attrDel.SetBackgroundColour("#CD2990")
                self.SetRowAttr(row, attrDel)
            elif diffType == 'a':
                attrAdd = grid.GridCellAttr()
                attrAdd.SetBackgroundColour("#B0E2FF")
                self.SetRowAttr(row, attrAdd)
        for row in range(len(colInfo)):
            diffType = colInfo[row][0:1]
            if diffType == 'd':
                # attr不能复用，每行都需要独自的attr
                attrDel = grid.GridCellAttr()
                attrDel.SetBackgroundColour("#CD2990")
                self.SetColAttr(row + 1, attrDel)
            elif diffType == 'a':
                attrAdd = grid.GridCellAttr()
                attrAdd.SetBackgroundColour("#B0E2FF")
                self.SetColAttr(row + 1, attrAdd)

        cellMapping = self.table.mapping['cell']
        for info in cellInfo:
            if flag == 'B':
                row, col = cellMapping[info[0][0]][info[0][1]]
            elif flag == 'A':
                row, col = cellMapping[info[1][0]][info[1][1]]
            self.SetCellBackgroundColour(row, col, wx.YELLOW)

        self.Bind(grid.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)
        self.SetSelectionBackground(wx.BLUE)

    def GetMapping(self):
        return self.mapping

    def OnCellLeftClick(self, evt):
        col = evt.GetCol()
        if col < 1:
            return
        row = evt.GetRow()
        self.GetParent().selectCell(row, col)


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

        for row in range(self.table.GetRowsCount()):
            for col in range(self.table.GetColsCount()):
                self.SetReadOnly(row, col)

        self.Bind(grid.EVT_GRID_CELL_LEFT_CLICK, self.OnCellLeftClick)

    def initTable(self, info):
        pass

    def getInfoMessage(self):
        pass

    def OnCellLeftClick(self, evt):
        pass


class RowInfoGrid(InfoGrid):
    """
    继承于InfoGrid，用于展示行diff信息
    """
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
        self.dataGridB.GoToCell(hightlightRow, 0)
        self.dataGridA.SelectRow(hightlightRow)
        self.dataGridA.GoToCell(hightlightRow, 0)


class ColInfoGrid(InfoGrid):
    """
    继承于InfoGrid，用于展示列diff信息
    """
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
        self.dataGridB.GoToCell(0, hightlightCol)
        self.dataGridA.SelectCol(hightlightCol)
        self.dataGridA.GoToCell(0, hightlightCol)


class CellInfoGrid(InfoGrid):
    """
    继承于InfoGrid，用于展示单元格diff信息
    """
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
        self.dataGridB.GoToCell(rowBefore, colBefore)

        cellMappingAfter = self.dataGridA.table.mapping['cell']
        infoAfter = self.info[row]
        rowAfter, colAfter = cellMappingAfter[infoAfter[0][0]][infoAfter[0][1]]
        self.dataGridA.SelectBlock(rowAfter, colAfter, rowAfter, colAfter)
        self.dataGridA.GoToCell(rowAfter, colAfter)
