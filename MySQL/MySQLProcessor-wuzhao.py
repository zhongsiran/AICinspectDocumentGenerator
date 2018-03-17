#-*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl import Workbook
from string import Template
import os


class data:
    def __init__(self):
        self.datacontent = ''
        self.datatpl = Template("('${nickname}','${name}','${zelinghao}'),\n")

    def divselect(self):
        self.headtpl = Template('''
insert into `app_shilingaic`.`${div}_wuzhao`
(`wuzhao_nickname`,`wuzhao_name`, `wuzhao_zelinghao`) VALUES
''')
        self.div = input()
        self.head = self.headtpl.substitute(div=self.div)
        
    def loadworkbook(self):
        try:
            wb = load_workbook('无照责令记录表.xlsx')
            self.ws = wb.worksheets[0]
        except FileNotFoundError:
            print('当前目录没有“无照责令记录表.xlsx”文件')
            exit(0)

    def processtosql(self):
        rows = self.ws.rows
        for row in rows:
            try:
                zlh = ''.join(row[2].value.split())
            except AttributeError:
                print('存在责令号为空的情况，请检查后重新运行。')

            try:
                nickname = ''.join(row[0].value.split())
            except AttributeError:
                print('存在代号为空的情况，请检查后重新运行。')
                
            try:
                name = ''.join(row[1].value.split())
            except AttributeError:
                print('存在行动名为空的情况，请检查后重新运行。')

            assert zlh != ''
            assert nickname != ''
            assert name != ''

            self.datacontent += self.datatpl.substitute(nickname=nickname, name=name, zelinghao=zlh)

    def savefile(self):
        f = open(self.div + '_wuzhao.sql','wb')
        f.write(self.head.encode('utf8'))
        f.write(self.datacontent[:-2].encode('utf8'))
        f.close()
        
if __name__ == '__main__':
    data = data()
    data.divselect()
    data.loadworkbook()
    data.processtosql()
    data.savefile() 
                    
        
