import requests
import os
import openpyxl


class DownloadInspectRecord:

    def __init__(self):
        self.action_list = []
        self.div = ''
        self.division_list = ['SL', 'YH', 'TB']
        self.choose_division()
        self.get_sp_list()

    def choose_division(self):
        while True:
            print('请选择监管所')
            print('1:狮岭')
            print('2:裕华')
            print('3:炭步')
            division_index = input()
            try:
                division_index = int(division_index) - 1
                self.div = self.division_list[division_index]
                print(self.div)
                break
            except ValueError:
                print('请输入对应的数字')

    def get_sp_list(self):
        params = {'secretkey': 'ZhongSiRan1990', 'div': self.div, 'function': 'get_list'}
        url = "https://shilingaic.applinzi.com/mylib/PyToolboxControllers/Sp_action_inspection_record_downloader.php"
        list_page = requests.post(url, params)
        #try:
        action_list = list_page.json()
        if action_list == '\u65e0\u4e13\u9879\u884c\u52a8\u6570\u636e':
            print('该所无专项行动数据')
            exit(0)
            os.system('pause')
        else:
            self.action_list = action_list
            print('请根据下列的列表选择你想下载的专项行动核查情况记录表。')
            action_no = 0
            for action in action_list:
                for key in action:
                    print('第' + str(action_no + 1) + '号专项行动 - 代号:“' + key + '” -- 名称：“' + action[key] + '”')
                    action_no += 1
            self.get_sp_record()
        # except:
        #     result = list_page.text
        #     print('result:' + result)

    def get_sp_record(self):
        print('请输入行动序号（只需要数字），例如 1 ：')
        input_command = input()
        while input_command != 'exit':
            try:
                selected_action_no = int(input_command)
                selected_action = self.action_list[selected_action_no - 1]
                for key in selected_action.keys():
                    print("即将下载以下行动的核查记录表：" + selected_action[key]+ '，请确认 （Y/N）')
                    # confirm = input()

                input_command = input()
            except ValueError:
                print('请只输入数字')
                input_command = input()
            except IndexError:
                print('无此序号')
                input_command = input()


if __name__ == "__main__":
    download = DownloadInspectRecord()