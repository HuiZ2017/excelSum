# -*- coding: utf-8 -*-
# @Time    : 2018/11/1 21:38
# @Author  : 张珲
# @File    : HandleExcel.py

from xlrd import open_workbook
from xlutils.copy import copy

class HandleExcel():
    def __init__(self,filename):
        self.filename = filename
        ii = 0
        with open_workbook(self.filename,formatting_info=True) as self.fileObj:
            self.Sheet = copy(self.fileObj)
            self.Result = []
            while True:
                try:
                    workSheet = self.fileObj.sheet_by_index(ii)
                    sheet = self.Sheet.get_sheet(ii)
                    sheetName = workSheet.name
                    rightnrow = self.getRightCrows(workSheet.col_values(0))
                    nrow = workSheet.nrows if rightnrow == workSheet.nrows \
                        else rightnrow
                    ncol = workSheet.ncols
                    result = {
                        'name':sheetName,
                        'sheet':sheet,
                        'row':nrow,
                        'col':ncol
                    }
                    self.Result.append(result)
                    ii += 1
                except IndexError as e:
                    break
            self.adminlist = [name['name'] for name in self.Result]
    def getRightCrows(self,colvalues):
        for ii in range(0,len(colvalues)):
            if not colvalues[ii]:
                return ii
        return len(colvalues)

    def writer(self,sheet,startrow,data):
        col = 0
        for ii in data:
            sheet.write(startrow,col,ii)
            col +=1
    def writing(self,sheetname,data):
        if sheetname in self.adminlist:
            for index, item in enumerate(self.adminlist):
                if item == sheetname:
                    for rowvalue in data:
                        self.writer(self.Result[index]['sheet'],
                                    self.Result[index]['row'],
                                    rowvalue)
        else:
            raise Exception
    def save(self):

        self.Sheet.save(self.filename)