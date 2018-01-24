#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx
import wx.lib.filebrowsebutton as filebrowse


class MainFrame(wx.Frame):
    def __init__(self):
        screenSize = wx.DisplaySize()
        frmWidth = screenSize[0] / 2
        frmHeight = screenSize[1] / 2
        frmPos = (screenSize[0] / 4, screenSize[1] / 4)
        wx.Frame.__init__(self, None, title="ExcelDiffer", pos=frmPos,
                          size=(frmWidth, frmHeight))

        self.excelFiles = ['', '']
        self.firstFileButton = filebrowse.FileBrowseButton(self, -1, pos=(20, 20), labelText="第一个Excel文件",
                                buttonText="选择", fileMask="*.xlsx", size=(600, -1), changeCallback=self.ChooseFirst)
        self.secondFileButton = filebrowse.FileBrowseButton(self, -1, pos=(20, 60), labelText="第二个Excel文件",
                                buttonText="选择", fileMask="*.xlsx", size=(600, -1), changeCallback=self.ChooseSecond)
        self.diffButton = wx.Button(self, -1, "开始比对", (20, 100))
        self.Bind(wx.EVT_BUTTON, self.Diff, self.diffButton)
        self.Show()

    def ChooseExcelFile(self, event, index):
        self.excelFiles[index] = event.GetString()

    def ChooseFirst(self, event):
        self.ChooseExcelFile(event, 0)

    def ChooseSecond(self, event):
        self.ChooseExcelFile(event, 1)

    def Diff(self, event):
        for file in self.excelFiles:
            print file

if __name__ == '__main__':
    app = wx.App()
    frm = MainFrame()
    app.MainLoop()

