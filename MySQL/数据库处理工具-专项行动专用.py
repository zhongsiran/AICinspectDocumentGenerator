#-*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl import Workbook
from string import Template
import os


class data:
    def __init__(self):
        self.datacontent = ''
        self.datatpl = Template("('${sp_action_no}','${xingdongming}','${daihao}','${corpname}','${regnum}','${startdate}','${enddate}'),\n")

    def divselect(self):
        self.headtpl = Template('''
insert into `app_shilingaic`.`${div}_zhuan_xiang_xing_dong`
(`sp_action_no`,`sp_action_name`, `sp_action_daihao`, `sp_action_corpname`, `sp_action_regnum`, `sp_action_startdate`, `sp_action_enddate`) VALUES
''')
        self.div = input()
        self.head = self.headtpl.substitute(div=self.div)
        
    def loadworkbook(self):
        try:
            wb = load_workbook('专项行动记录表.xlsx')
            self.ws = wb.worksheets[0]
        except FileNotFoundError:
            print('当前目录没有“专项行动记录表.xlsx”文件')
            exit(0)

        
    def processtosql(self):
        x=daihao=corpname=r=startdate=enddate=''
        rows = self.ws.rows
        for row in rows:
            try:
                r = ''.join(row[4].value.split())
            except AttributeError:
                print('存在注册号为空的情况，请检查后重新运行。')

            try:
                no = ''.join(row[0].value.split())
            except AttributeError:
                print('存在行动代号为空的情况，请检查后重新运行。')
                
            try:
                x = ''.join(row[1].value.split())
            except AttributeError:
                print('存在行动名为空的情况，请检查后重新运行。')
            try:
                daihao = ''.join(row[2].value.split())
            except AttributeError:
                daihao = row[2].value
            try:
                corpname = ''.join(row[3].value.split())
            except AttributeError:
                corpname = ''
            try:
                startdate = ''.join(row[5].value.split())
            except AttributeError:
                startdate = ''
            try:
                enddate = ''.join(row[6].value.split())
            except AttributeError:
                enddate = ''

            assert daihao != ''
            assert r != ''

            self.datacontent += self.datatpl.substitute(sp_action_no=no,xingdongming=x,daihao=daihao,corpname=corpname,regnum=r,startdate=startdate,enddate=enddate)
    def savefile(self):
        f = open(self.div + '_zhuangxiang.sql','wb')
        f.write(self.head.encode('utf8'))
        f.write(self.datacontent[:-2].encode('utf8'))
        f.close()
        
if __name__ == '__main__':
    data = data()
    data.divselect()
    data.loadworkbook()
    data.processtosql()
    data.savefile() 
                    
        
