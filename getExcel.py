#!/usr/bin/env python
# encoding: utf-8
'''
@author: zhanghui
@file: getExcel.py
@time: 2018/9/8 15:49
@desc:
'''

import xlrd,xlwt,time,datetime
from xlutils.copy import copy

class getExcel():
    def __init__(self,filename):
        try:
            with xlrd.open_workbook(filename) as self.fileObj:
                self.worksheet1 = self.fileObj.sheet_by_index(0)
                self.num_rows = self.worksheet1.nrows
                self.num_cols = self.worksheet1.ncols

        except Exception as e:
            pass
        pass
        self.new = self.str2date('2018/9/8')
        #self.new = time.strftime("%Y/%m/%d") #此处可做成用户输入
    # def load(self):
    #     #obj.load('file.xls')加载文件对象
    #     pass
    def get(self,ncol=0):
        #取对应行的数据，返回('ncol列数据',[整行数据])worksheet1.row_values(rown)
        if ncol < self.num_cols:
            row = 0
            while row < self.num_rows:
                value1,rowdata = self.handle(self.worksheet1.row_values(row))
                row += 1
                if value1 and rowdata:
                    yield (value1,rowdata)
    def handle(self,data):
        #['43350.0', '彭海城', '湖南湘大兽药有限公司', '冷远广', 18684878585.0, '高企', '电话', '加资料客户企业经理人有意向要上报总公司']
        new = data[0]
        adminname = data[1]
        custome = data[2]
        contacter = data[3]
        contacternum = data[4]
        project = data[5]
        ways = data[6]
        note = data[7]
        if self.comparedays(self.new) == new:
            try:
                contacternum = int(contacternum)
            except Exception as ee:
                pass
            finally:
                return (new, [new, adminname, custome, contacter, contacternum, project, ways, note])
        else:
            return None,None
        #
    def str2date(self,strs):
        return datetime.date(int(strs.split('/')[0]),int(strs.split('/')[1]),int(strs.split('/')[2]))
    def comparedays(self,new):
        old = datetime.date(1900,1,1)
        if isinstance(self.new,str):
            new = self.str2date(new)
        return (new-old).days + 2
# demo = getExcel('客户统计表_张珲_20180907.xls')
# print(demo.num_rows,demo.num_cols)
#
# for i in demo.get():
#     print(i)
