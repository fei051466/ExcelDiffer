#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
from wx import grid

from customGrid import SheetGrid, DiffSheetGrid
from customGrid import RowInfoGrid, ColInfoGrid, CellInfoGrid
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
        self.sheetNB = wx.Notebook(self, -1, pos=(20, 10), size=(650, 400))
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
        sheetA = secondExcel.sheet_by_name(sheetName)

        dataB = [[str(sheetB.cell_value(row, col))
                 for col in range(sheetB.ncols)]
                 for row in range(sheetB.nrows)]
        dataA = [[str(sheetA.cell_value(row, col))
                 for col in range(sheetA.ncols)]
                 for row in range(sheetA.nrows)]

        rowInfo, colInfo, cellInfo = diff(dataB, dataA)

        '''
        print 'diffInfo'
        rowDiffInfo = []
        for i in diffInfo:
            diffType = i[0:1]
            if diffType == 's':
                indexB, indexA = i[1:].split(':')
                rowDiffInfo.append(diff(dataB[int(indexB)], dataA[int(indexA)]))
        for i in range(len(dataB)):
            delInfo = "d%d" % i
            allHave = True
            for j in rowDiffInfo:
                if delInfo not in j:
                    allHave = False
                    break
            if allHave:
                print 'delfInfo:', delInfo
        for i in range(len(dataA)):
            addInfo = "a%d" % i
            allHave = True
            for j in rowDiffInfo:
                if addInfo not in j:
                    allHave = False
                    break
            if allHave:
                print 'addInfo:', addInfo
        '''

        dataGridB = DiffSheetGrid(self, sheetB, rowInfo, colInfo, cellInfo, "B")
        dataGridA = DiffSheetGrid(self, sheetA, rowInfo, colInfo, cellInfo, "A")

        sizerUp= wx.BoxSizer(wx.HORIZONTAL)
        sizerUp.Add(dataGridB, 1, wx.EXPAND, 5)
        sizerUp.Add(dataGridA, 1, wx.EXPAND, 5)

        infoNB = wx.Notebook(self, -1)

        rowInfoPanel = wx.Panel(infoNB, -1)
        rowInfoGrid = RowInfoGrid(rowInfoPanel, rowInfo, dataGridB, dataGridA)
        rowInfoMessage = wx.StaticText(rowInfoPanel, -1,
                                       rowInfoGrid.getInfoMessage())

        sizerRowInfo = wx.BoxSizer(wx.VERTICAL)
        sizerRowInfo.Add(rowInfoMessage)
        sizerRowInfo.Add(rowInfoGrid)
        rowInfoPanel.SetSizer(sizerRowInfo)

        infoNB.AddPage(rowInfoPanel, "行增删")

        colInfoPanel = wx.Panel(infoNB, -1)
        colInfoGrid = ColInfoGrid(colInfoPanel, colInfo, dataGridB, dataGridA)
        colInfoMessage = wx.StaticText(colInfoPanel, -1,
                                       colInfoGrid.getInfoMessage())

        sizerColInfo = wx.BoxSizer(wx.VERTICAL)
        sizerColInfo.Add(colInfoMessage)
        sizerColInfo.Add(colInfoGrid)
        colInfoPanel.SetSizer(sizerColInfo)

        infoNB.AddPage(colInfoPanel, "列增删")

        cellInfoPanel = wx.Panel(infoNB, -1)
        cellInfoGrid = CellInfoGrid(cellInfoPanel, cellInfo, dataGridB, dataGridA)
        cellInfoMessage = wx.StaticText(cellInfoPanel, -1,
                                       cellInfoGrid.getInfoMessage())

        sizerCellInfo = wx.BoxSizer(wx.VERTICAL)
        sizerCellInfo.Add(cellInfoMessage)
        sizerCellInfo.Add(cellInfoGrid)
        cellInfoPanel.SetSizer(sizerCellInfo)

        infoNB.AddPage(cellInfoPanel, "单元格修改")

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