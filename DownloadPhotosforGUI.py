#-*- coding: utf-8 -*-
# ver1.2：增加分所下载功能

import requests
import os
import ast
import re
from datetime import date
from threading import Thread
import threading
import sys

from wx.lib.pubsub import pub
import wx
import math

exist_count = 0
new_count = 0
threadLock = threading.Lock()

class photo_dl_thread(Thread):
    """docstring for photo_dl_thread"""
    def __init__(self, division_index, target_dir):
        Thread.__init__(self)
        self.division_index = division_index
        self.target_dir = target_dir
        self.setDaemon(True)
        self.running = True
        self.start()    # start the thread

    def run(self):
        '''
        下载照片
        '''    
        global exist_count
        global new_count
        pl = photolibrary(self.target_dir)
        pl.divisionselector(self.division_index)
        get_dict_result = pl.getlinksdict()
        if (self.running == True):
            if (get_dict_result[0] == '成功连接云服务器，下载照片中；\n' and self.running == True):
                wx.CallAfter(self.postprogress, get_dict_result[0])
                all_keys = list(get_dict_result[2])
                all_keys.sort()
                all_keys_num = get_dict_result[1]
                the_dict = get_dict_result[3]
                if (all_keys_num > 4 and self.running == True):
                    keys_first_part = all_keys[:math.floor(all_keys_num/2)]
                    self.t1 = photo_download_sub_thread('T1', self.division_index, self.target_dir ,keys_first_part, the_dict)
                    keys_last_part = all_keys[math.floor(all_keys_num/2)-1:]
                    self.t2 = photo_download_sub_thread('T2', self.division_index, self.target_dir ,keys_last_part, the_dict)
                    self.t1.join()
                    self.t2.join()
                elif (self.running == True):
                    result_count_all = pl.downloading(all_keys)
                    exist_count = result_count_all[0]
                    new_count = result_count_all[1]
                    content = ['success',('已存在文件：' + str(exist_count) + '\n新增文件：' + str(new_count) + '个')]
                    wx.CallAfter(self.postfinished, content)
                    exist_count = 0
                    new_count = 0
            elif (self.running == True) :
                    wx.CallAfter(self.postfinished, ['fail', u'下载失败，请检查网络后重试'])

    def stop(self):
        wx.CallAfter(self.postfinished, ['fail', u'下载已被取消'])
        self.running = False
        self.t1.pl_object.stop_download()
        self.t2.pl_object.stop_download()


    def postprogress(self, msg_to_post):
        '''
        Sent msg to GUI
        '''
        pub.sendMessage("update", msg = msg_to_post)

    def postfinished(self, msg_to_post):
        pub.sendMessage("dl_finished", result = msg_to_post)


class photo_download_sub_thread (Thread):
    """docstring for ph"""
    def __init__(self, threadid, division_index, target_dir, dict_keys, the_dict):
        Thread.__init__(self)
        self.division_index = division_index
        self.target_dir = target_dir
        self.keys = dict_keys
        self.file_dict = the_dict
        self.threadid = threadid
        self.running = True
        self.start()
    def run(self):
        global exist_count
        global new_count
        self.pl_object = photolibrary(self.target_dir)
        self.pl_object.divisionselector(self.division_index)
        result_count = self.pl_object.downloadpic(self.threadid, self.keys, self.file_dict)
        exist_count += result_count[0]
        new_count += result_count[1]

class photolibrary:
    def __init__(self, target_dir):
        if(target_dir == ''):
            self.rootfolder = os.getcwd()
            self.target_dir = os.getcwd()
        else:
            self.rootfolder = target_dir
            self.target_dir = target_dir

        self.prefolder = ''
        self.running = True
        
    def getlinksdict(self):
        try:
            pagepy=requests.post("https://shilingaic.applinzi.com/public/listfilepy.php",data={'secretkey':'ZhongSiRan1990'})
        except (ConnectionError):
            result.append('连接网络失败，请检查。')

        result = []
        if(pagepy.status_code == 200):
            result.append('成功连接云服务器，下载照片中；\n')
        else:
            result.append('连接网络失败，请检查。http错误码：' + pagepy.status_code +'\n')
        #now = date.today()
        #today = "%d-%d-%d" %(now.year,now.month,now.day)
        length = len(pagepy.text) - 1
        pagetext = pagepy.text[:length]
        pagetext="{" + pagetext + "}"
        self.file_dict = ast.literal_eval(pagetext)
        result.append(len(self.file_dict.keys()))
        result.append(self.file_dict.keys())
        result.append(self.file_dict)
        return result # 连接情况；字典长度；字典健列表；字典本身
        
    def divisionselector(self,division_index):
        index = division_index
        divisionlist=['SL','FR','XH','YH','HC','XY','XQ','HS','HD','TB','CN','TM']
        if(index == 'quit'):
            exit(0)
        else:
            try:
                self.division = (divisionlist[int(index)])
                #print('''选择成功，正在下载...''')
            except(ValueError):
                #print('''**请输入数字序号，如 1 ,然后按回车键。不要输入字母等其他内容**。''')
                self.divisionselector()
        
    def downloadpic(self,threadid, keys, the_dict):
        exist_count = 0
        new_count = 0
        for filename in keys:
            while (self.running):
                sourceurl = the_dict.get(filename)
                if(self.divisionfilter(filename)):
                    os.chdir(self.rootfolder)
                    corpfolder = re.sub(r'-?' + self.division + '?-\d+(_\d+-\d+-\d+)?.jpg',"",filename) #删除剩下企业名称
                    if (self.oldname):
                        corpfolder = re.sub(r'-\d+.jpg',"",filename)
                        
                    if(self.prefolder == corpfolder):
                        pass
                    else:
                        pass
                        #print('当前企业：' + corpfolder)
                    try:
                        os.mkdir(corpfolder)
                        os.chdir(self.target_dir + '\\' + corpfolder)
                    except (FileExistsError):
                        os.chdir(self.target_dir + '\\' + corpfolder)
                        
                    LocalImgPath = os.getcwd() + '\\'+ filename
                    #SortedImgPath = os.getcwd() +"\\"+ today + '\\'+ filename
                    threadLock.acquire()
                    if os.path.isfile(LocalImgPath):
                        #print('文件已经存在')
                        #print(LocalImgPath + " already exists")
                        exist_count += 1
                        threadLock.release()
                        pass
                    else:
                        LocalImg = open(LocalImgPath, 'wb')
                        threadLock.release()
                        NetImg=requests.get(sourceurl)
                        LocalImg.write(NetImg.content)
                        LocalImg.close()
                        print(threadid + '新增文件' + filename)
                        new_count += 1
                    self.oldname = False
                    self.prefolder = corpfolder                    
                else:
                    pass
        #content = ['success',('已存在文件：' + str(exist_count) + '\n新增文件：' + str(new_count) + '个')]
        result_count = [exist_count, new_count]
        return result_count
    def stop_download(self):
        self.running = False

    def divisionfilter(self,filename):
        patterndiv = re.compile(r'-' + self.division + '-\d+(_\d+-\d+-\d+)?.jpg',re.A)
        patterngen = re.compile(r'-\w{2}-\d+.jpg',re.A) #旧有无分所和日期的文件名 ex. 字号-01.JPG
        if (re.findall(patterndiv,filename) != []):
            self.oldname = False
            return True
        elif((re.findall(patterngen,filename) != []) and (self.division != 'SL')):
            self.oldname = False
            return False
        elif((re.findall(patterngen,filename) == []) and (self.division == 'SL')):
            self.oldname = True
            return True
        
# if __name__ == '__main__':
#     photolib = photolibrary()
#     photolib.divisionselector()
#     photolib.getlinksdict()
#     photolib.downloadpic()





