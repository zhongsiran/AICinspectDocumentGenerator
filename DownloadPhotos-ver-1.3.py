#-*- coding: utf-8 -*-
# ver1.2：增加分所下载功能

import requests
import os
import ast
import re
from datetime import date

class photolibrary:
    def __init__(self):
        self.rootfolder = os.getcwd()
        self.prefolder = ''
        
    def getlinksdict(self):        
        pagepy=requests.post("https://shilingaic.applinzi.com/public/listfilepy.php",data={'secretkey':'ZhongSiRan1990'})
        print(pagepy)
        #now = date.today()
        #today = "%d-%d-%d" %(now.year,now.month,now.day)
        length = len(pagepy.text) - 1
        pagetext = pagepy.text[:length]
        pagetext="{" + pagetext + "}"
        self.file_dict = ast.literal_eval(pagetext)
        
    def divisionselector(self):
        index = input('''请根据以下列表输入序号：
1、狮岭  2、芙蓉  3、新华
4、裕华  5、花城  6、新雅
7、秀全  8、花山  9、花东
10、炭步 11、赤坭 12、梯面

输入数字选择下载对应所的照片或输入“quit”退出程序:''')
        
        divisionlist=['SL','FR','XH','YH','HC','XY','XQ','HS','HD','TB','CN','TM']
        if(index == 'quit'):
            exit(0)
        else:
            try:
                self.division = (divisionlist[int(index)-1])
                print('''
选择成功，正在下载...
''')
            except(ValueError):
                print('''

**请输入数字序号，如 1 ,然后按回车键。不要输入字母等其他内容**。
''')
                self.divisionselector()
        
    def downloadpic(self):
        for filename, sourceurl in self.file_dict.items():
            if(self.divisionfilter(filename)):
                os.chdir(self.rootfolder)
                corpfolder = re.sub(r'-?' + self.division + '?-\d+.jpg',"",filename) #删除剩下企业名称
                if (self.oldname):
                    corpfolder = re.sub(r'-\d+.jpg',"",filename)
                    
                if(self.prefolder == corpfolder):
                    pass
                else:
                    print('当前企业：' + corpfolder)
                try:
                    os.mkdir(corpfolder)
                    os.chdir(os.getcwd() + '\\' + corpfolder)
                except:
                    os.chdir(os.getcwd() + '\\' + corpfolder)
                    
                LocalImgPath = os.getcwd() + '\\'+ filename
                #SortedImgPath = os.getcwd() +"\\"+ today + '\\'+ filename
                
                if os.path.isfile(LocalImgPath):
                    print('文件已经存在')
                    #print(LocalImgPath + " already exists")
                    pass
                else:
                    LocalImg = open(LocalImgPath, 'wb')
                    NetImg=requests.get(sourceurl)
                    LocalImg.write(NetImg.content)
                    LocalImg.close()
                    print('新增文件' + filename)
                self.oldname = False
                self.prefolder = corpfolder
            else:
                pass
        os.system('pause')
            

    def divisionfilter(self,filename):
        patterndiv = re.compile(r'-' + self.division + '-\d+.jpg',re.A)
        patterngen = re.compile(r'-\w{2}-\d+.jpg',re.A)
        if (re.findall(patterndiv,filename) != []):
            self.oldname = False
            return True
        elif((re.findall(patterngen,filename) != []) and (self.division != 'SL')):
            self.oldname = False
            return False
        elif((re.findall(patterngen,filename) == []) and (self.division == 'SL')):
            self.oldname = True
            return True
        
if __name__ == '__main__':
    photolib = photolibrary()
    photolib.divisionselector()
    photolib.getlinksdict()
    photolib.downloadpic()





