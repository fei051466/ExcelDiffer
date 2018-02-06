#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid

from customGrid import SheetGrid, DiffSheetGrid
from customGrid import RowInfoGrid, ColInfoGrid, CellInfoGrid
from diffAlgorithm import diff
from utils import getSheetData


class SheetPanel(wx.Panel):
    """
    用于展示独有sheet的Panel
    """
    def __init__(self, parent, size, excelFilePath, excelFile, sheets):
        wx.Panel.__init__(self, parent, size=size)
        sizer = wx.BoxSizer()
        if not sheets:
            noDataText = wx.StaticText(self, -1, "没有相关数据", size=size)
            sizer.Add(noDataText, 0, wx.ALIGN_CENTER)
            return
        self.excelFilePath = excelFilePath
        self.sheetNB = wx.Notebook(self, -1, size=size)
        for sheetName in sheets:
            dataPanel = DataPanel(self.sheetNB, excelFile, sheetName)
            self.sheetNB.AddPage(dataPanel, sheetName)
        sizer.Add(self.sheetNB, 0, wx.ALIGN_CENTER)
        self.SetSizer(sizer)
        self.Bind(wx.EVT_SIZE, self.OnCPSize)

    def OnCPSize(self, evt):
        self.sheetNB.SetSize(evt.GetSize())


class SameSheetPanel(wx.Panel):
    """
    用于展示相同sheet与diff结果的panel
    """
    def __init__(self, parent, size,
                 firstExcel, secondExcel, firstFile, secondFile, sheets):
        wx.Panel.__init__(self, parent, size=size)
        sizer = wx.BoxSizer()
        if not sheets:
            noDataText = wx.StaticText(self, -1, "没有相关数据")
            sizer.Add(noDataText, 0, wx.ALIGN_CENTER)
            return
        self.sheetNB = wx.Notebook(self, -1)
        for sheetName in sheets:
            diffDataPanel = DiffDataPanel(self.sheetNB, size,
                    firstExcel, secondExcel, firstFile, secondFile, sheetName)
            self.sheetNB.AddPage(diffDataPanel, sheetName)
        sizer.Add(self.sheetNB, 0, wx.ALIGN_CENTER|wx.EXPAND)
        self.SetSizer(sizer)

        self.Bind(wx.EVT_SIZE, self.OnCPSize)

    def OnCPSize(self, evt):
        self.sheetNB.SetSize(evt.GetSize())


class DiffDataPanel(wx.Panel):
    """
    用于展示diff结果的panel
    """
    def __init__(self, parent, size,
                 firstExcel, secondExcel, firstFile, secondFile, sheetName):
        wx.Panel.__init__(self, parent, size=size)
        sheetB = firstExcel.sheet_by_name(sheetName)
        sheetA = secondExcel.sheet_by_name(sheetName)

        dataB = getSheetData(sheetB)
        dataA = getSheetData(sheetA)

        rowInfo, colInfo, cellInfo = diff(dataB, dataA)
        fileNameB = wx.StaticText(self, -1, firstFile)
        fileNameA = wx.StaticText(self, -1, secondFile)

        self.dataGridB = DiffSheetGrid(self, sheetB, rowInfo, colInfo, cellInfo, "B")
        self.dataGridA = DiffSheetGrid(self, sheetA, rowInfo, colInfo, cellInfo, "A")

        sizerFileB = wx.BoxSizer(wx.VERTICAL)
        sizerFileA = wx.BoxSizer(wx.VERTICAL)

        sizerFileB.Add(fileNameB, 1, wx.ALIGN_CENTER)
        sizerFileB.Add(self.dataGridB, 9, wx.ALIGN_CENTER)
        sizerFileA.Add(fileNameA, 1, wx.ALIGN_CENTER)
        sizerFileA.Add(self.dataGridA, 9, wx.ALIGN_CENTER)

        # 使用GridSizer可以使得两个数据展示窗口在缩小时同时被缩小
        sizerUp = wx.GridSizer(1, 2, 2, 2)
        sizerUp.Add(sizerFileB, 0, wx.EXPAND)
        sizerUp.Add(sizerFileA, 0, wx.EXPAND)

        infoNB = wx.Notebook(self, -1)

        rowInfoPanel = wx.Panel(infoNB, -1)
        rowInfoGrid = RowInfoGrid(rowInfoPanel, rowInfo,
                                  self.dataGridB, self.dataGridA)
        rowInfoMessage = wx.StaticText(rowInfoPanel, -1,
                                       rowInfoGrid.getInfoMessage())

        sizerRowInfo = wx.BoxSizer(wx.VERTICAL)
        sizerRowInfo.Add(rowInfoMessage)
        sizerRowInfo.Add(rowInfoGrid, 1, wx.EXPAND)
        rowInfoPanel.SetSizer(sizerRowInfo)

        infoNB.AddPage(rowInfoPanel, "行增删")

        colInfoPanel = wx.Panel(infoNB, -1)
        colInfoGrid = ColInfoGrid(colInfoPanel, colInfo,
                                  self.dataGridB, self.dataGridA)
        colInfoMessage = wx.StaticText(colInfoPanel, -1,
                                       colInfoGrid.getInfoMessage())

        sizerColInfo = wx.BoxSizer(wx.VERTICAL)
        sizerColInfo.Add(colInfoMessage)
        sizerColInfo.Add(colInfoGrid, 1, wx.EXPAND)
        colInfoPanel.SetSizer(sizerColInfo)

        infoNB.AddPage(colInfoPanel, "列增删")

        cellInfoPanel = wx.Panel(infoNB, -1)
        cellInfoGrid = CellInfoGrid(cellInfoPanel, cellInfo,
                                    self.dataGridB, self.dataGridA)
        cellInfoMessage = wx.StaticText(cellInfoPanel, -1,
                                       cellInfoGrid.getInfoMessage())

        sizerCellInfo = wx.BoxSizer(wx.VERTICAL)
        sizerCellInfo.Add(cellInfoMessage)
        sizerCellInfo.Add(cellInfoGrid, 1, wx.EXPAND)
        cellInfoPanel.SetSizer(sizerCellInfo)

        infoNB.AddPage(cellInfoPanel, "单元格修改")

        sizerOuter = wx.BoxSizer(wx.VERTICAL)
        sizerOuter.Add(sizerUp, 6, wx.EXPAND)
        sizerOuter.Add(infoNB, 4, wx.EXPAND)

        self.SetSizer(sizerOuter)

    def selectCell(self, row, col):
        self.dataGridB.SelectBlock(row, col, row, col)
        self.dataGridB.GoToCell(row, col)
        self.dataGridA.SelectBlock(row, col, row, col)
        self.dataGridA.GoToCell(row, col)


class DataPanel(wx.Panel):
    """
    用于展示独有sheet的数据的panel
    """
    def __init__(self, parent, excelFile, sheetName):
        wx.Panel.__init__(self, parent)
        sheet = excelFile.sheet_by_name(sheetName)
        dataGrid = SheetGrid(self, sheet)
        sizer = wx.BoxSizer(wx.VERTICAL)  # 添加了sizer才能使得网格正常实现，原因不明
        sizer.Add(dataGrid, 0, wx.ALIGN_CENTER)
        self.SetSizer(sizer)