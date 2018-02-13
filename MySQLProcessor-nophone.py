#-*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl import Workbook
from string import Template
import os


class data:
    def __init__(self):
        self.datacontent = ''
        self.datatpl = Template("('${corpname}','${regnum}','${addr}','${repperson}','${contactperson}','${inspection}',' ',NULL,NULL,0,0),\n")

    def divselect(self):
        self.headtpl = Template('''
insert into `app_shilingaic`.`${div}_corp`
(`CorpName`, `RegNum`, `Addr`, `RepPerson`, `ContactPerson`, `InspectionStatus`, `PhoneCallRecord`, `Loca_x`, `Loca_y`, `Active`, `PicNum`) VALUES
''')
        self.div = input()
        self.head = self.headtpl.substitute(div=self.div)
        
    def loadworkbook(self):
        try:
            wb = load_workbook('待导入记录表.xlsx')
            self.ws = wb.worksheets[0]
        except FileNotFoundError:
            print('当前目录没有“待导入记录表.xlsx”文件')
            exit(0)

        
    def processtosql(self):
        c=r=a=p=rp=cp=cph=ins=phcal=''
        rows = self.ws.rows
        for row in rows:
            try:
                r = ''.join(row[1].value.split())
            except AttributeError:
                print('存在注册号为空的情况，请检查后重新运行。')
            try:
                a = ''.join(row[2].value.split())
            except AttributeError:
                a = ''
            try:
                p = ''.join(row[3].value.split())
            except AttributeError:
                p = ''
            try:
                rp = ''.join(row[7].value.split())
            except AttributeError:
                rp = ''
            try:
                cp = ''.join(row[8].value.split())
            except AttributeError:
                cp = ''
            try:
                cph = ''.join(row[9].value.split())
            except AttributeError:
                cph = ''
            try:
                c = ''.join(row[0].value.split())
            except AttributeError:
                c = '(' + str(r) + ')无字号'
            
            try:
                ins = row[10].value 
            except IndexError:
                pass
            try:
                phcal = row[11].value
            except IndexError:
                pass
            assert c != ''
            self.datacontent += self.datatpl.substitute(corpname=c,regnum=r,addr=a,repperson=rp,contactperson=cp,inspection=ins,phonecall=phcal)
    def savefile(self):
        f = open(self.div + '.sql','wb')
        f.write(self.head.encode('utf8'))
        f.write(self.datacontent[:-2].encode('utf8'))
        f.write(b'''
on duplicate key update 
CorpName = Values(CorpName),
Addr=values(addr),
repperson = values(repperson),
contactperson = values(contactperson);''')
        f.close()
        
if __name__ == '__main__':
    data = data()
    data.divselect()
    data.loadworkbook()
    data.processtosql()
    data.savefile() 
                    
        
