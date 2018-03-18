import requests

class DownloadInspectRecord:

    def __init__(self):
        self.get_sp_list()

    def get_sp_list(self):
        params = {'secretkey': 'ZhongSiRan1990', 'div': 'SL', 'function': 'get_list'}
        url = "https://shilingaic.applinzi.com/mylib/PyToolboxControllers/Sp_action_inspection_record_downloader.php"
        list_page = requests.post(url , params)
        try:
            action_list = list_page.json()
        except:
            action_list = list_page.text
        print(action_list)

if __name__ == "__main__":
    download = DownloadInspectRecord()