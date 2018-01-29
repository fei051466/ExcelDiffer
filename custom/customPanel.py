#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid

from customGrid import SheetGrid, InfoGrid
from diffAlgorithm import diff


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
        self.sheetNB = wx.Notebook(self, -1, pos=(20, 10), size=(550, 300))
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
        dataGridB = SheetGrid(self, sheetB)
        sheetA = secondExcel.sheet_by_name(sheetName)
        dataGridA = SheetGrid(self, sheetA)

        sizerUp= wx.BoxSizer(wx.HORIZONTAL)
        sizerUp.Add(dataGridB, 1, wx.EXPAND, 5)
        sizerUp.Add(dataGridA, 1, wx.EXPAND, 5)

        infoNB = wx.Notebook(self, -1)

        infoPanel = wx.Panel(infoNB, -1)
        diffInfo = diff(dataGridB.getData(), dataGridA.getData())
        infoGrid = InfoGrid(infoPanel, diffInfo)
        infoMessage = wx.StaticText(infoPanel, -1, infoGrid.getInfoMessage())

        sizerInner = wx.BoxSizer(wx.VERTICAL)
        sizerInner.Add(infoMessage)
        sizerInner.Add(infoGrid)
        infoPanel.SetSizer(sizerInner)

        infoNB.AddPage(infoPanel, "行增删")

        sizerOuter = wx.BoxSizer(wx.VERTICAL)
        sizerOuter.Add(sizerUp, 1, wx.EXPAND, 5)
        sizerOuter.Add(infoNB, 1, wx.EXPAND, 5)

        self.SetSizer(sizerOuter)


class DataPanel(wx.Panel):
    """
    用于展示独有sheet的数据的panel
    """
    def __init__(self, parent, excelFile, sheetName):
        wx.Panel.__init__(self, parent)
        sheet = excelFile.sheet_by_name(sheetName)
        dataGrid = SheetGrid(self, sheet)
        sizer = wx.BoxSizer(wx.VERTICAL)  # 添加了sizer才能使得网格正常实现，原因不明
        sizer.Add(dataGrid, 1, wx.EXPAND, 5)
        self.SetSizer(sizer)