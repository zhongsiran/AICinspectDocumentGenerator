import os
import re


class RenameFolder:

    def __init__(self):
        self.root_folder = '\\\\psf\\Home\\Desktop\\nf2'

    def walk_through_folder(self):
        for singledir, subdirs, files in os.walk(self.root_folder):
            if singledir != '\\\\psf\\Home\\Desktop\\nf2':
                single_dir = singledir.replace(self.root_folder, '')
                pattern = re.compile(r'\\SL-\S+-\S+-', re.A)
                single_dir = re.sub(pattern, '', single_dir)
                newpath = os.path.join(self.root_folder, single_dir)
                # print('rename' + singledir + 'to' + newpath)
                try:
                    os.rename(singledir + '\\', newpath + '\\')
                except FileNotFoundError:
                    pass
                except FileExistsError:
                    pass
                # print(single_dir)
        print(self.root_folder)

    def walk_through_files(self):
        for singledir, subdirs, files in os.walk(self.root_folder):
            for single_dir in singledir:
                for each_file in files:
                    print( os.path.join(single_dir, each_file))



if __name__ == "__main__":
    rf = RenameFolder()
    #rf.walk_through_folder()
    rf.walk_through_files()
    os.system('pause')
