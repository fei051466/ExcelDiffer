#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
import wx.lib.filebrowsebutton as filebrowse
import xlrd
from wx import grid


class MainFrame(wx.Frame):
    def __init__(self):
        screenSize = wx.DisplaySize()
        frmWidth = screenSize[0] / 2
        frmHeight = screenSize[1] / 2
        frmPos = (screenSize[0] / 4, screenSize[1] / 4)
        wx.Frame.__init__(self, None, title="ExcelDiffer", pos=frmPos,
                          size=(frmWidth, frmHeight))

        self.excelFilePaths = ['', '']
        self.firstFileButton = filebrowse.FileBrowseButton(self, -1, pos=(20, 20), labelText="第一个Excel文件",
                                buttonText="选择", fileMask="*.xlsx", size=(600, -1), changeCallback=self.ChooseFirst,)
        self.secondFileButton = filebrowse.FileBrowseButton(self, -1, pos=(20, 60), labelText="第二个Excel文件",
                                buttonText="选择", fileMask="*.xlsx", size=(600, -1), changeCallback=self.ChooseSecond)
        self.diffButton = wx.Button(self, -1, "开始比对", (20, 100))
        self.Bind(wx.EVT_BUTTON, self.Diff, self.diffButton)
        self.Show()

    def ChooseExcelFile(self, event, index):
        self.excelFilePaths[index] = event.GetString()

    def ChooseFirst(self, event):
        self.ChooseExcelFile(event, 0)

    def ChooseSecond(self, event):
        self.ChooseExcelFile(event, 1)

    def Diff(self, event):
        '''
        if '' in self.excelFiles:
            dlg = wx.MessageDialog(self, "请选择两个Excel文件", "温馨提示", wx.OK)
            dlg.ShowModal()
            dlg.Destroy()
            return
        '''
        firstExcel = xlrd.open_workbook(self.excelFilePaths[0])
        secondExcel = xlrd.open_workbook(self.excelFilePaths[1])
        self.typeNB = wx.Notebook(self, -1, pos=(20, 120), size=(600, 250))
        sameSheets, diffSheets = self.calcSheet(firstExcel, secondExcel)
        self.typeNB.commonSheets = sameSheets
        self.typeNB.diffSheets = diffSheets
        samePanel = SameSheetPanel(self.typeNB, firstExcel, secondExcel, sameSheets)
        onlyInFirstPanel = SheetPanel(self.typeNB, self.excelFilePaths[0], firstExcel, diffSheets[0])
        onlyInSecondPanel = SheetPanel(self.typeNB, self.excelFilePaths[1], secondExcel, diffSheets[1])
        self.typeNB.AddPage(samePanel, "相同sheet")
        self.typeNB.AddPage(onlyInFirstPanel, "第一个文件独有sheet")
        self.typeNB.AddPage(onlyInSecondPanel, "第二个文件独有sheet")

    def calcSheet(self, firstExcel, secondExcel):
        firstSheets = firstExcel.sheet_names()
        secondSheets = secondExcel.sheet_names()
        commonSheets = []
        onlyInFirst = []
        onlyInSecond = []
        for sheet in firstSheets:
            if sheet in secondSheets:
                commonSheets.append(sheet)
            else:
                onlyInFirst.append(sheet)
        for sheet in secondSheets:
            if sheet not in commonSheets:
                onlyInSecond.append(sheet)
        return commonSheets, [onlyInFirst, onlyInSecond]


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
            diffDataPanel = DiffDataPanel(self.sheetNB, firstExcel, secondExcel, sheetName)
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
            dataGrid.SetCellSize(cell[0], cell[2], (cell[1] - cell[0]), (cell[3] - cell[2]))
            dataGrid.SetCellAlignment(cell[0], cell[2], wx.ALIGN_CENTER, wx.ALIGN_CENTER)
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




if __name__ == '__main__':
    app = wx.App()
    frm = MainFrame()
    app.MainLoop()

