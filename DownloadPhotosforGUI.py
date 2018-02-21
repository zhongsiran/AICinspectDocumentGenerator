#-*- coding: utf-8 -*-
# done

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
import time
import random

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
        new_count = 0
        exist_count = 0
        pl = photolibrary(self.target_dir)
        pl.divisionselector(self.division_index)
        try:
            get_dict_result = pl.getlinksdict()
            if (get_dict_result[0] == '成功连接云服务器，下载照片中；\n'):
                wx.CallAfter(self.postprogress, get_dict_result[0])
                all_keys = list(get_dict_result[2])
                all_keys.sort()
                all_keys_num = get_dict_result[1]
                the_dict = get_dict_result[3]
                try:
                    del self.t1
                    del self.t2
                except Exception as e:
                    print(e)
                t1 = random.randrange(0, 101, 2)
                t2 = random.randrange(0, 101, 2)
                keys_first_part = all_keys[:math.floor(all_keys_num/2)]
                self.t1 = photo_download_sub_thread(t1, self.division_index, self.target_dir ,keys_first_part, the_dict)
                keys_last_part = all_keys[math.floor(all_keys_num/2)-1:]
                self.t2 = photo_download_sub_thread(t2, self.division_index, self.target_dir ,keys_last_part, the_dict)
                if(self.t1.isAlive()):
                    self.t1.join()    
                if(self.t2.isAlive()):            
                    self.t2.join()
                content = '已有' + str(exist_count) + '张照片，新增' +  str(new_count) + '张照片'
                wx.CallAfter(self.postfinished, ['success', content]) 
            else:
                wx.CallAfter(self.postfinished, ['fail', u'下载失败，请检查网络后重试']) 
        except:
            self.stop('无法连接服务器，请检查外网（互联网）连接')

    def stop(self, msg='正在取消下载'):
        wx.CallAfter(self.postfinished, ['fail', msg])
        self.running = False
        try:
            self.t1.running = False
            self.t2.running = False
            self.t1.pl_object.stop_download()
            self.t2.pl_object.stop_download()
        except AttributeError:
            pass
        return True

    def postprogress(self, msg_to_post):
        '''
        Sent msg to GUI
        '''
        pub.sendMessage("update", msg = msg_to_post)

    def postfinished(self, msg_to_post):
        pub.sendMessage("dl_finished", result = msg_to_post)

########################################################################################
class photo_download_sub_thread (Thread):
    """docstring for ph"""
    def __init__(self, name, division_index, target_dir, dict_keys, the_dict):
        Thread.__init__(self)
        self.division_index = division_index
        self.target_dir = target_dir
        self.keys = dict_keys
        self.file_dict = the_dict
        self.name = name
        self.running = True
        self.start()
    def run(self):
        global exist_count
        global new_count
        if (self.running == True):
            self.pl_object = photolibrary(self.target_dir)
            self.pl_object.divisionselector(self.division_index)
            result_count = self.pl_object.downloadpic(self.name, self.keys, self.file_dict)
        exist_count += result_count[0]
        new_count += result_count[1]
#########################################################################
class photolibrary:
    def __init__(self, target_dir):
        if(target_dir == ''):
            self.target_dir = os.getcwd()
        else:
            self.target_dir = target_dir

        self.prefolder = ''
        self.running = True
        
    def downloadpic(self, name, keys, the_dict):
        exist_count = 0
        new_count = 0
        for filename in keys:
            if (self.running):
                sourceurl = the_dict.get(filename)
                if(self.divisionfilter(filename)):
                    corpfolder = re.sub(r'-?' + self.division + '?-\d+(_\d+-\d+-\d+)?.jpg',"",filename) #删除剩下企业名称
                    if (self.oldname):
                        corpfolder = re.sub(r'-\d+.jpg',"",filename)
                    #corpfolder-字号得到
                    try:
                        os.mkdir(self.target_dir + '\\' + corpfolder)
                    except (FileExistsError):
                        pass
                    finally:
                        target_sub_dir = self.target_dir + '\\' + corpfolder
                    with threadLock:
                        self.post_dl_thread_progress('正在下载到：' + target_sub_dir)
                    
                    #print(self.target_dir + '\\' + corpfolder)

                    local_img_file = target_sub_dir + '\\'+ filename
                    #SortedImgPath = os.getcwd() +"\\"+ today + '\\'+ filename
                    #threadLock.acquire()
                    if os.path.isfile(local_img_file):
                        #print('文件已经存在')
                        #print(local_img_file + " already exists")
                        exist_count += 1
                     #   threadLock.release()
                    else:
                        LocalImg = open(local_img_file, 'wb')
                      #  threadLock.release()
                        try:
                            NetImg=requests.get(sourceurl)
                            LocalImg.write(NetImg.content)
                            LocalImg.close()
                            with threadLock:
                                self.post_dl_thread_progress('新增文件：' + filename)
                            new_count += 1
                        except:
                            self.running == False
                            break

                    self.oldname = False
                else:
                    self.oldname = False
        #content = ['success',('已存在文件：' + str(exist_count) + '\n新增文件：' + str(new_count) + '个')]
        result_count = [exist_count, new_count]
        return result_count

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

    def getlinksdict(self):
        try:
            pagepy=requests.post("https://shilingaic.applinzi.com/public/listfilepy.php",data={'secretkey':'ZhongSiRan1990'})
        except:
            result.append('连接网络失败，请检查。')
            self.running == False
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


    def stop_download(self):
        self.running = False

    def post_dl_thread_progress(self, msg_to_post):
        '''
        Sent msg to GUI
        '''
        pub.sendMessage("update", msg = msg_to_post)

        
# if __name__ == '__main__':
#     photolib = photolibrary()
#     photolib.divisionselector()
#     photolib.getlinksdict()
#     photolib.downloadpic()





