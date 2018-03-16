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
import time
import random

exist_count = 0
new_count = 0
threadLock = threading.Lock()

class photo_dl_thread(Thread):
    """docstring for photo_dl_thread"""
    def __init__(self, division_index, target_dir, division_password):
        Thread.__init__(self)
        self.division_index = division_index
        self.division_password = division_password
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
        pl = PhotoLibrary(self.target_dir)
        pl.divisionselector(self.division_index)
        try:
            get_dict_result = pl.getlinksdict(self.division_password)
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
                if (all_keys_num > 4):
                    keys_first_part = all_keys[:math.floor(all_keys_num/2)]
                    keys_last_part = all_keys[math.floor(all_keys_num/2)-1:]
                else:
                    keys_first_part = all_keys
                    keys_last_part = all_keys
                self.t1 = PhotoDownloadSubThread(t1, self.division_index, self.target_dir, keys_first_part, the_dict)
                self.t2 = PhotoDownloadSubThread(t2, self.division_index, self.target_dir, keys_last_part, the_dict)
                if(self.t1.isAlive()):
                    self.t1.join()
                if(self.t2.isAlive()):            
                    self.t2.join()
                content = '新增' + str(new_count) + '张照片'
                wx.CallAfter(self.postfinished, ['success', content]) 
            elif (get_dict_result[0] == 'newwork_error'):
                wx.CallAfter(self.postfinished, ['fail', u'下载失败，请检查网络后重试']) 
            elif (get_dict_result[0] == 'div_pwd_error'):
                wx.CallAfter(self.postfinished, ['fail', u'下载失败，监管所密码有误，请检查或联系管理员'])
            else:
                wx.CallAfter(self.postfinished, ['fail', u'下载失败，出现未知错误，请稍后重试或联系管理员'])                 
        except Exception as e:
            print (e)
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


class PhotoDownloadSubThread (Thread):
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
        result_count = [0, 0]
        if self.running == True:
            self.pl_object = PhotoLibrary(self.target_dir)
            self.pl_object.divisionselector(self.division_index)
            result_count = self.pl_object.download_picture(self.name, self.keys, self.file_dict)
        exist_count += result_count[0]
        new_count += result_count[1]
#########################################################################


class PhotoLibrary:
    def __init__(self, target_dir):
        if target_dir == '':
            self.target_dir = os.getcwd()
        else:
            self.target_dir = target_dir

        self.pre_folder = ''
        self.file_dict = []
        self.running = True
        
    def download_picture(self, name, keys, the_dict):
        exist_count = 0
        new_count = 0
        for filename in keys:
            if (self.running):
                sourceurl = the_dict.get(filename)
                if(self.division_filter(filename)):
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
                            self.running = False
                            break

                    self.oldname = False
                else:
                    self.oldname = False
        #content = ['success',('已存在文件：' + str(exist_count) + '\n新增文件：' + str(new_count) + '个')]
        result_count = [exist_count, new_count]
        return result_count

    def division_filter(self, filename):
        patterndiv = re.compile(r'-' + self.division + '-\d+(_\d+-\d+-\d+)?.jpg',re.A)
        patterngen = re.compile(r'-\w{2}-\d+.jpg',re.A) #旧有无分所和日期的文件名 ex. 字号-01.JPG
        if re.findall(patterndiv, filename):
            self.oldname = False
            return True
        elif re.findall(patterngen, filename) and self.division != 'SL':
            self.oldname = False
            return False
        elif (not re.findall(patterngen, filename)) and self.division == 'SL':
            self.oldname = True
            return True

    def getlinksdict(self, division_password):
        pagepy=requests.post("https://shilingaic.applinzi.com/mylib/PyToolboxControllers/PhotoDownloaderListFile.php",data={'secretkey':'ZhongSiRan1990', 'div_pwd' : division_password, 'div' : self.division})
        result = []
        web_return_text = pagepy.text
        if web_return_text.find('pwd_error') == 0:
            #print(web_return_text)
            #print(web_return_text.find('pwd_error'))
            result.append('div_pwd_error')
        else:
            if pagepy.status_code == 200:
                #print(web_return_text)
                #print(web_return_text.find('pwd_error'))
                result.append('成功连接云服务器，下载照片中；\n')
                length = len(pagepy.text) - 1
                page_text = pagepy.text[:length]
                page_text="{" + page_text + "}"
                self.file_dict = ast.literal_eval(page_text)
                result.append(len(self.file_dict.keys()))
                result.append(self.file_dict.keys())
                result.append(self.file_dict)
            else:
                result.append('network_error')
        #now = date.today()
        #today = "%d-%d-%d" %(now.year,now.month,now.day)
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
