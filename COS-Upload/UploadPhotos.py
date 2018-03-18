import requests
import os


class Main:

    def __init__(self):
        self.root_folder = os.getcwd()
        # self.prefix = 'SL//2018年第一批吊销（内资）'
        self.prefix = input()

    def walkthrough_and_upload(self):
        for singledir, subdirs, files in os.walk(self.root_folder):
            single_dir = singledir.replace(self.root_folder, '')
            for filename in files:
                if '.db' not in filename.lower() or '.py' not in filename.lower():
                    file_path = single_dir + '\\' + filename  #\corpname\corpname.jpg
                    #print(file_path)
                    self.upload_file(file_path)
        print(self.root_folder)

    def upload_file(self, file_path):
        file_full_path = self.root_folder + '\\' + file_path
        file_obj = open(file_full_path, 'r+b')
        file_data = file_obj.read()
        file_obj.close()

        cos_key = file_path.replace('\\', '//')
        data = {"key": 'CorpImg//' + self.prefix + '//' + cos_key, "file_data": file_data}
        upload = requests.post("https://shilingaic.applinzi.com/mylib/PyToolboxControllers/PhotoUpload.php", data)
        if "Location" in upload.text:
            print(file_full_path)
            print(os.remove(file_full_path))
        else:
            print('not remove')
        del file_data
        del upload


if __name__ == '__main__':
    Uploader = Main()
    Uploader.walkthrough_and_upload()
    os.system('pause')
