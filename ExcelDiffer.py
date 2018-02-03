#!/usr/bin/env python
# -*- coding:utf-8 -*-

import wx

from config import getConfig
from custom.customFrame import MainFrame

mode = 'dev'

if __name__ == '__main__':
    app = wx.App()
    frm = MainFrame()
    if mode == 'dev':
        frm.firstFileButton.SetValue(getConfig('excelFilePath', 'firstFile'))
        frm.secondFileButton.SetValue(getConfig('excelFilePath', 'secondFile'))
        frm.Diff(None)
    app.MainLoop()

