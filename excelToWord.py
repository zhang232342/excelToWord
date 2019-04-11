#!/usr/bin/env python
# -*- coding:utf-8 -*-
# @TIME     :2019/4/10 17:12
# @Author   :CandyZ
# @File     :excelToWord.py
# 导入xlrd模块
import xlrd
from xlutils.copy import copy
from mailmerge import MailMerge

#设置文件名和路径
fname = 'd:/excelToWord/dms数据字典.xlsx'
#打开文件
filename = xlrd.open_workbook(fname)
# 获取当前文档的表(得到的是sheet的个数，一个整数）
sheets = filename.nsheets
sheet = filename.sheets()[0] #通过sheet索引获得sheet对象
# print sheet
#获取行数
nrows = sheet.nrows
# 获取列数
ncols = sheet.ncols
#获取第一行,第一列数据数据
cell_value = sheet.cell_value(1,1)
print(cell_value)

#打开模版
template = 'd:/excelToWord/模版.docx'
document = MailMerge(template)
print("Fields included in {}: {}".format(template, document.get_merge_fields()))
document.merge(test = cell_value)
document.write('d:/完成word.docx')
print('数据写入成功!')
