import requests

class DownloadInspectRecord:

    def __init__(self):
        self.action_list = []
        self.get_sp_list()
        self.get_sp_record()


    def get_sp_list(self):
        params = {'secretkey': 'ZhongSiRan1990', 'div': 'SL', 'function': 'get_list'}
        url = "https://shilingaic.applinzi.com/mylib/PyToolboxControllers/Sp_action_inspection_record_downloader.php"
        list_page = requests.post(url , params)
        try:
            action_list = list_page.json()
            self.action_list = action_list
            print('请根据下列的列表选择你想下载的专项行动核查情况记录表。')
            action_no = 0
            for action in action_list:
                for key in action:
                    print('第' + str(action_no + 1) + '号专项行动 - 代号:“' + key + '” -- 名称：“' + action[key] + '”')
                    action_no += 1
        except:
            result = list_page.text
            print(result)

    def get_sp_record(self):
        print('请输入行动序号（只需要数字），例如 1 ：')
        input_command = input()
        while input_command != 'exit':
            try:
                selected_action_no = int(input_command)
                selected_action = self.action_list[selected_action_no - 1]
                for key in selected_action.keys():
                    print("即将下载以下行动的核查记录表：" + selected_action[key]+ '，请确认 （Y/N）')
                    confirm = input()

                input_command = input()
            except ValueError:
                print('请只输入数字')
                input_command = input()
            except IndexError:
                print('无此序号')
                input_command = input()










if __name__ == "__main__":
    download = DownloadInspectRecord()