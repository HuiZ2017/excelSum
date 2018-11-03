#!/usr/bin/env python
# encoding: utf-8
'''
@author: zhanghui
@file: excel_sum_v0.4.py
@time: 2018/11/3 13:33
'''

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as msg
import tkinter.filedialog as tkFileDialog
import re
from HandleExcel import HandleExcel as excel
class excelsum(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("客户信息统计表日汇总 v0.4")
        self.geometry("500x500")
        self.label = tk.Label(self, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.label.pack(side=tk.BOTTOM,fill=tk.X)
        self.label.config(text='by 张珲 有建议欢迎交流 Tel:+86 17673102113')
        self.file_opt = {}
        self.file_opt['initialfile'] = ''
        self.file_opt['filetypes'] = [('xls', 'xls')]
        self.file_opt['initialdir'] = '.'
        #self.wm_attributes('-topmost', 1)
        self.excelobj = None
        self.result = {}
        self.adminlist = []
        self.frame = tk.Frame(self,height=500,width=500)
        self.frame.pack()
        self.input_text = tk.Text(self.frame,width=500,height=400,font=('楷体', 8))
        self.input_text.insert(0.0,'''Step 1
按如下格式将文本粘贴至此

时间;跟进人;公司名称;联系人;联系电话;意向业务;拜访方式;跟进概况

例：
2018/10/31;张珲;湖南科德信息咨询有限公司;张总;185xxxxxxxx;渠道;面谈;xxxxxx
2018/11/1;张珲;阿凡提信息科技(湖南)股份有限公司;张总;185xxxxxxxx;军工;面谈;xxxxxx

Step 2
按Ctrl + E 格式化待导入数据，并确认

Step 3
选择导入excel，限xls格式
        ''')
        self.input_text.bind("<Button-1>", self.clearInputtext);self.flag = True
        self.input_text.bind("<Control-e>", self.getText)
        self.input_text.pack()
    def clearInputtext(self,event):
        if self.flag:
            self.input_text.delete(0.0,tk.END)
            self.flag = False
        else:
            pass
    def askopenfile(self):
        filename = tkFileDialog.askopenfilename(**self.file_opt)
        if filename:
            try:
                self.excelobj = excel(filename)
            except PermissionError as e:
                msg.showerror('错误','写入权限不够，不要打开待写入excel')
                self.askopenfile()
    def getText(self,event):
        text = self.input_text.get(0.0,tk.END)
        self.loadText(text.split('\n'))
    def loadText(self,textlist):
        self.result = {};self.adminlist = []
        for text in textlist:
            if text:
                text = re.split('[;；]',text)
                if text[1] in self.adminlist:
                    self.result[text[1]].append(text)
                else:
                    self.adminlist.append(text[1])
                    self.result[text[1]] = [text]
        self.showloaded()
    def confirm(self):
        self.root.destroy()
        info = ''
        for ii in self.result.keys():
            info = info + '%s: %s\n' %(ii,len(self.result[ii]))
        if info:
            if not msg.askyesno('确认并导入信息',message=info):
                pass
            else:
                self.start()
    def showloaded(self):
        self.root = tk.Tk()
        self.root.title("已检测到的待导入信息")
        self.root.geometry("1000x600")
        self.root.wm_attributes('-topmost', 1)
        tk.Button(self.root, text="确认", fg="black", bg="white", command=self.confirm).pack()
        self.outbox = ttk.Treeview(self.root)  # 表格
        self.outbox = ttk.Treeview(self.root,
                                   show="headings",
                                   height=100,
                                   columns=("a", "b", "c", "d","e","f","g","h"))
        self.outbox.column("a", width=80, anchor="center");self.outbox.heading("a", text="时间")
        self.outbox.column("b", width=80, anchor="center");self.outbox.heading("b", text="跟进人")
        self.outbox.column("c", width=150, anchor="center");self.outbox.heading("c", text="公司名称")
        self.outbox.column("d", width=80, anchor="center");self.outbox.heading("d", text="联系人")
        self.outbox.column("e", width=80, anchor="center");self.outbox.heading("e", text="联系电话")
        self.outbox.column("f", width=80, anchor="center");self.outbox.heading("f", text="意向业务")
        self.outbox.column("g", width=80, anchor="center");self.outbox.heading("g", text="拜访方式")
        self.outbox.column("h", width=400, anchor="center");self.outbox.heading("h", text="跟进概况")
        self.outbox.pack()
        for key in self.result.keys():
            for item in self.result[key]:
                self.outbox.insert('', 0, text="123", values=([ii for ii in item]))
    def start(self):
        info_ = ''
        if self.excelobj:
            pass
        else:
            self.askopenfile()
        for ii in self.result.keys():
            try:
                self.excelobj.writing(ii, self.result[ii])
                info_ = info_ + '%s 添加 %s 条\n' %(ii,len(self.result[ii]))
            except Exception as ee:
                msg.showerror('错误','%s 未匹配到任何sheet，未添加' %ii)
        try:
            self.excelobj.save()
            msg.showinfo('完成', info_)
        except PermissionError as ee:
            msg.showerror('错误', '写入权限不够，不要打开待写入excel')

if __name__ == "__main__":
    demo1 = excelsum()
    demo1.mainloop()