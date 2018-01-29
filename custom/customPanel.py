#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid


class SheetPanel(wx.Panel):
    """
    用于展示独有sheet的Panel
    """
    def __init__(self, parent, excelFilePath, excelFile, sheets):
        wx.Panel.__init__(self, parent)
        if not sheets:
            noDataText = wx.StaticText(self, -1, "没有相关数据")
            return
        self.excelFilePath = excelFilePath
        self.sheetNB = wx.Notebook(self, -1, pos=(20, 10), size=(550, 200))
        for sheetName in sheets:
            print sheetName
            dataPanel = DataPanel(self.sheetNB, excelFile, sheetName)
            self.sheetNB.AddPage(dataPanel, sheetName)


class SameSheetPanel(wx.Panel):
    """
    用于展示相同sheet与diff结果的panel
    """
    def __init__(self, parent, firstExcel, secondExcel, sheets):
        wx.Panel.__init__(self, parent)
        if not sheets:
            noDataText = wx.StaticText(self, -1, "没有相关数据")
            return
        self.sheetNB = wx.Notebook(self, -1, pos=(20, 10), size=(550, 200))
        for sheetName in sheets:
            print sheetName
            diffDataPanel = DiffDataPanel(self.sheetNB, firstExcel,
                                          secondExcel, sheetName)
            self.sheetNB.AddPage(diffDataPanel, sheetName)


class DiffDataPanel(wx.Panel):
    """
    用于展示diff结果的panel
    """
    def __init__(self, parent, firstExcel, secondExcel, sheetName):
        wx.Panel.__init__(self, parent)
        sheetB = firstExcel.sheet_by_name(sheetName)
        dataGridB = grid.Grid(self, -1)
        dataGridB.CreateGrid(sheetB.nrows, sheetB.ncols)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(dataGridB, 1, wx.EXPAND, 5)
        sheetA = firstExcel.sheet_by_name(sheetName)
        dataGridA = grid.Grid(self, -1)
        dataGridA.CreateGrid(sheetA.nrows, sheetA.ncols)
        sizer.Add(dataGridA, 1, wx.EXPAND, 5)
        sizerOuter = wx.BoxSizer(wx.VERTICAL)
        sizerOuter.Add(sizer, 1, wx.EXPAND, 5)
        staticMessage = wx.StaticText(self, -1, "test")
        sizerOuter.Add(staticMessage, 1, wx.EXPAND, 5)
        self.SetSizer(sizerOuter)


class DataPanel(wx.Panel):
    """
    用于展示独有sheet的数据的panel
    """
    def __init__(self, parent, excelFile, sheetName):
        wx.Panel.__init__(self, parent)
        sheet = excelFile.sheet_by_name(sheetName)
        rows, cols = sheet.nrows, sheet.ncols
        dataGrid = grid.Grid(self, -1)
        dataGrid.CreateGrid(rows, cols)
        mergedCells = sheet.merged_cells
        for cell in mergedCells:
            dataGrid.SetCellSize(cell[0], cell[2],
                                 (cell[1] - cell[0]), (cell[3] - cell[2]))
            dataGrid.SetCellAlignment(cell[0], cell[2],
                                      wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        for row in range(0, rows):
            for col in range(0, cols):
                value = sheet.cell(row, col).value
                if not value:
                    continue
                if isinstance(value, float):
                    if int(value) == value:
                        value = int(value)
                value = str(value)
                dataGrid.SetCellValue(row, col, value)
        sizer = wx.BoxSizer(wx.VERTICAL)  # 添加了sizer才能使得网格正常实现，原因不明
        sizer.Add(dataGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)