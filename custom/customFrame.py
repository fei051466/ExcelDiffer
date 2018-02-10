#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
import wx.lib.filebrowsebutton as filebrowse
import xlrd

from customPanel import SameSheetPanel, SheetPanel


class MainFrame(wx.Frame):
    """
    程序主窗口
    """
    def __init__(self):
        wx.Frame.__init__(self, None, title="ExcelDiffer", size=(960, 640))

        self.Center()

        self.firstFileButton = filebrowse.FileBrowseButton(self, -1,
                labelText="第一个Excel文件", buttonText="选择")
        self.secondFileButton = filebrowse.FileBrowseButton(self, -1,
                labelText="第二个Excel文件", buttonText="选择")

        fileButtonSizer = wx.BoxSizer(wx.VERTICAL)
        fileButtonSizer.Add(self.firstFileButton, 0, wx.EXPAND)
        fileButtonSizer.Add(self.secondFileButton, 0, wx.EXPAND)

        self.diffButton = wx.Button(self, -1, "开始比对")

        fileSelectSizer = wx.BoxSizer(wx.HORIZONTAL)
        fileSelectSizer.Add(fileButtonSizer, 9, wx.EXPAND)
        fileSelectSizer.Add(self.diffButton, 1, wx.EXPAND)

        self.typeNB = wx.Notebook(self, -1)
        frameSizer = wx.BoxSizer(wx.VERTICAL)
        frameSizer.Add(fileSelectSizer, 1, wx.ALIGN_CENTER)
        frameSizer.Add(self.typeNB, 9, wx.EXPAND)

        self.SetSizer(frameSizer)

        self.Bind(wx.EVT_BUTTON, self.Diff, self.diffButton)
        self.Show()

    def Diff(self, event):
        """
        点击"开始对比"按钮后触发的函数，计算diff数据并在窗口中展示
        :param event:
        :return:
        """
        # self.typeNB.DeleteAllPages()
        for pageNum in range(self.typeNB.GetPageCount()):
            self.typeNB.RemovePage(0)
        firstFile = self.firstFileButton.GetValue()
        secondFile = self.secondFileButton.GetValue()
        if firstFile == '' or secondFile == '':
            dlg = wx.MessageDialog(self, "请选择两个Excel文件", "温馨提示", wx.OK)
            dlg.ShowModal()
            dlg.Destroy()
            return
        if firstFile == secondFile:
            dlg = wx.MessageDialog(self, "所选文件相同，无需对比", "温馨提示", wx.OK)
            dlg.ShowModal()
            dlg.Destroy()
            return
        try:
            firstExcel = xlrd.open_workbook(firstFile)
            secondExcel = xlrd.open_workbook(secondFile)
        except Exception, e:
            dlg = wx.MessageDialog(self, "打开文件出错，请检查所选文件", "温馨提示", wx.OK)
            dlg.ShowModal()
            dlg.Destroy()
            return
        sameSheets, diffSheets = self.calcSheet(firstExcel, secondExcel)
        self.typeNB.commonSheets = sameSheets
        self.typeNB.diffSheets = diffSheets
        size = self.GetSize()
        samePanel = SameSheetPanel(self.typeNB, size, firstExcel,
                                secondExcel, firstFile, secondFile, sameSheets)
        onlyInFirstPanel = SheetPanel(self.typeNB, size,
                                      firstFile, firstExcel, diffSheets[0])
        onlyInSecondPanel = SheetPanel(self.typeNB, size,
                                       secondFile, secondExcel, diffSheets[1])
        self.typeNB.AddPage(samePanel, "相同sheet")
        self.typeNB.AddPage(onlyInFirstPanel, "第一个文件独有sheet")
        self.typeNB.AddPage(onlyInSecondPanel, "第二个文件独有sheet")

    def calcSheet(self, firstExcel, secondExcel):
        """
        计算两个Excel文件中sheet表情况，得出相同名称的sheet表和单独的sheet表
        :param firstExcel:
        :param secondExcel:
        :return:
        """
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