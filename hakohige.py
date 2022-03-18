#!/usr/bin/env python
# coding: utf-8

import tkinter
from tkinter import filedialog
import pandas as pd
import xlwings as xw

tk = tkinter.Tk()
tk.withdraw()
fTyp = [("Files","*.*")]
iDir = r'C:\Users'
titleText = "Select file"
file = filedialog.askopenfilename(filetypes = fTyp, title = titleText, initialdir = iDir)
filename = '\\'.join(file.split('/'))

wb = xw.Book(filename)
sheet = wb.sheets.add('hako-hige')

singleDf = pd.read_excel(filename, index_col=0)
# 箱ひげ図の描画
boxFig = singleDf.plot.box(showmeans=True).get_figure()
sheet.pictures.add(boxFig)