import wxglade_out
import DocGeneratorGUI
import DownloadPhotosforGUI


'''
对GUI界面第一面功能选择按键的事件处理器
'''
class FuncPageEventHandlers(wxglade_out.MainFrame):

    def __init__(self, mainframe):
        #super(EventHandlers, self).__init__()
        self.mainframe = mainframe

    def goto_ph_dl_page(self):
        self.mainframe.Entrance.SetSelection(1)

    def goto_ins_dl_page(self):
        self.mainframe.Entrance.SetSelection(3)

    def goto_doc_gen_page(self):
        self.mainframe.Entrance.SetSelection(2)

class InspectionRecordDownloadEventHandlers():
    def __init__(self):
        pass

    def select_download_path(self):
        pass
