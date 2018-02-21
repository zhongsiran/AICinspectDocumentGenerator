#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# generated by wxGlade 0.8.0b3 on Mon Feb 19 14:05:38 2018
#

import wx
from wx.lib.pubsub import pub
import re
# begin wxGlade: dependencies
import gettext
# end wxGlade

# begin wxGlade: extracode
import DocGeneratorGUI as dg
import DownloadPhotosforGUI as dp
import os
# end wxGlade


class MainFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MainFrame.__init__
        kwds["style"] = kwds.get("style", 0) | wx.BORDER_SIMPLE | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((730, 754))
        self.frame_statusbar = self.CreateStatusBar(1, wx.STB_SIZEGRIP)
        self.Entrance = wx.Notebook(self, wx.ID_ANY)
        self.function_index_page = wx.Panel(self.Entrance, wx.ID_ANY)
        self.photo_download_page = wx.Panel(self.Entrance, wx.ID_ANY)
        self.rbox_division = wx.RadioBox(self.photo_download_page, wx.ID_ANY, _(u"\u9009\u62e9\u76d1\u7ba1\u6240"), choices=[_(u"\u72ee\u5cad"), _(u"\u8299\u84c9"), _(u"\u65b0\u534e"), _(u"\u88d5\u534e"), _(u"\u82b1\u57ce"), _(u"\u65b0\u96c5"), _(u"\u79c0\u5168"), _(u"\u82b1\u5c71"), _(u"\u82b1\u4e1c"), _(u"\u70ad\u6b65"), _(u"\u8d64\u576d"), _(u"\u82b1\u4e1c"), _(u"\u68af\u9762")], majorDimension=1, style=wx.RA_SPECIFY_COLS)
        self.photo_download_result_text = wx.TextCtrl(self.photo_download_page, wx.ID_ANY, "", style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_WORDWRAP)
        self.save_to_path_text_area = wx.TextCtrl(self.photo_download_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.choose_download_path = wx.Button(self.photo_download_page, wx.ID_ANY, _(u"\u66f4\u6539\u4e0b\u8f7d\u8def\u5f84\n"), style=wx.BU_AUTODRAW)
        self.begin_download_photo = wx.Button(self.photo_download_page, wx.ID_ANY, _(u"\n\u5f00\u59cb\u4e0b\u8f7d\n"), style=wx.BU_AUTODRAW)
        self.cancel_download_photo = wx.Button(self.photo_download_page, wx.ID_ANY, _(u"\n\u53d6\u6d88\u4e0b\u8f7d\n"), style=wx.BU_AUTODRAW)
        self.doc_gen_page = wx.Panel(self.Entrance, wx.ID_ANY)
        self.doc_gen_ins_tpl_path = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.doc_gen_ph_tpl_path = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.doc_gen_xlsx_path = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.doc_gen_corp_photos_path = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.doc_gen_path = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_LEFT | wx.TE_MULTILINE | wx.TE_WORDWRAP)
        self.doc_gen_choose_inspect_tpl = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u9009\u7b14\u5f55\u6a21\u677f\n"), style=wx.BU_AUTODRAW)
        self.doc_gen_choose_photo_tpl = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u9009\u7167\u7247\u6a21\u677f\n"), style=wx.BU_AUTODRAW)
        self.doc_gen_choose_inspect_xlsx = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u9009\u62e9\u6838\u67e5\u8868\n"), style=wx.BU_AUTODRAW)
        self.doc_gen_choose_corp_photos_path = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u9009\u7167\u7247\u76ee\u5f55\n"), style=wx.BU_AUTODRAW)
        self.label_ins_tpl = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u5df2\u7ecf\u6307\u5b9a\u683c\u5f0f\u4e3aDOCX\u7684\u73b0\u573a\u7b14\u5f55\u6a21\u677f\n"), style=wx.ALIGN_CENTER)
        self.label_ph_tpl = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u5df2\u7ecf\u6307\u5b9a\u683c\u5f0f\u4e3aDOCX\u7684\u8bc1\u636e\u63d0\u53d6\u5355\u6a21\u677f\n"), style=wx.ALIGN_CENTER)
        self.label_xlsx = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u5df2\u7ecf\u6307\u5b9a\u683c\u5f0f\u4e3aXLSX\u7684\u4f01\u4e1a\u4fe1\u606f\u53ca\u6838\u67e5\u8bb0\u5f55\u8868\n"), style=wx.ALIGN_CENTER)
        self.doc_gen_result_text_area = wx.TextCtrl(self.doc_gen_page, wx.ID_ANY, "", style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_WORDWRAP)
        self.doc_gen_preview = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u751f\u6210\u4e00\u6237\u9884\u89c8"), style=wx.BU_AUTODRAW)
        self.doc_gen_choose_path = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u66f4\u6539\u751f\u6210\u8def\u5f84\n"), style=wx.BU_AUTODRAW)
        self.doc_gen_begin = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\u5f00\u59cb\u751f\u6210"), style=wx.BU_AUTODRAW)
        self.doc_gen_stop = wx.Button(self.doc_gen_page, wx.ID_ANY, _(u"\n\u4e2d\u6b62\u751f\u6210\n"), style=wx.BU_AUTODRAW)
        self.digital_hash_page = wx.Panel(self.Entrance, wx.ID_ANY)

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_RADIOBOX, self.rbox_div_select, self.rbox_division)
        self.Bind(wx.EVT_BUTTON, self.choose_download_path_btn, self.choose_download_path)
        self.Bind(wx.EVT_BUTTON, self.begin_download_btn, self.begin_download_photo)
        self.Bind(wx.EVT_BUTTON, self.cancel_download_btn, self.cancel_download_photo)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_ins_tpl_btn, self.doc_gen_choose_inspect_tpl)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_ph_tpl_btn, self.doc_gen_choose_photo_tpl)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_xlsx_btn, self.doc_gen_choose_inspect_xlsx)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_cp_phs_btn, self.doc_gen_choose_corp_photos_path)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_preview_btn, self.doc_gen_preview)
        self.Bind(wx.EVT_BUTTON, self.choose_dg_path_btn, self.doc_gen_choose_path)
        self.Bind(wx.EVT_BUTTON, self.begin_dg_btn, self.doc_gen_begin)
        self.Bind(wx.EVT_BUTTON, self.cancel_dg_btn, self.doc_gen_stop)
        self.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.page_changing, self.Entrance)
        # end wxGlade
        # create a pubsub receiver
        pub.subscribe(self.update_pd_result_text, "update")
        pub.subscribe(self.download_finished, "dl_finished")

    def __set_properties(self):
        # begin wxGlade: MainFrame.__set_properties
        self.SetTitle(_(u"\u5e02\u573a\u76d1\u7ba1\u5de5\u5177\u7bb1"))
        _icon = wx.NullIcon
        _icon.CopyFromBitmap(wx.Bitmap("C:\\py\\aicbadge.ico", wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)
        self.frame_statusbar.SetStatusWidths([-1])
        
        # statusbar fields
        frame_statusbar_fields = [_("frame_statusbar")]
        for i in range(len(frame_statusbar_fields)):
            self.frame_statusbar.SetStatusText(frame_statusbar_fields[i], i)
        self.rbox_division.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        self.rbox_division.SetFocus()
        self.rbox_division.SetSelection(1)
        self.photo_download_result_text.SetMinSize((233, 120))
        self.save_to_path_text_area.SetMinSize((500, 11))
        self.save_to_path_text_area.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.choose_download_path.SetMinSize((120, 40))
        self.choose_download_path.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.begin_download_photo.SetMinSize((120, 40))
        self.begin_download_photo.SetForegroundColour(wx.Colour(0, 0, 0))
        self.begin_download_photo.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.cancel_download_photo.SetMinSize((120, 40))
        self.cancel_download_photo.SetForegroundColour(wx.Colour(255, 0, 0))
        self.cancel_download_photo.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.cancel_download_photo.Enable(False)
        self.doc_gen_ins_tpl_path.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_ph_tpl_path.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_xlsx_path.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_corp_photos_path.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_path.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_choose_inspect_tpl.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_choose_photo_tpl.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_choose_inspect_xlsx.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_choose_corp_photos_path.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.label_ins_tpl.SetBackgroundColour(wx.Colour(255, 0, 0))
        self.label_ins_tpl.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.label_ph_tpl.SetBackgroundColour(wx.Colour(255, 0, 0))
        self.label_ph_tpl.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.label_xlsx.SetBackgroundColour(wx.Colour(255, 0, 0))
        self.label_xlsx.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.doc_gen_result_text_area.SetMinSize((297, 220))
        self.doc_gen_preview.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_choose_path.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_begin.SetForegroundColour(wx.Colour(0, 0, 0))
        self.doc_gen_begin.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_stop.SetForegroundColour(wx.Colour(255, 0, 0))
        self.doc_gen_stop.SetFont(wx.Font(16, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "楷体_GB2312"))
        self.doc_gen_stop.Enable(False)
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MainFrame.__do_layout
        sizer_1 = wx.FlexGridSizer(0, 1, 0, 0)
        doc_gen_root_sizer = wx.StaticBoxSizer(wx.StaticBox(self.doc_gen_page, wx.ID_ANY, _(u"\u6587\u4e66\u751f\u6210\u5668")), wx.HORIZONTAL)
        doc_gen_right_side = wx.BoxSizer(wx.VERTICAL)
        doc_gen_left_side = wx.BoxSizer(wx.VERTICAL)
        sizer_17 = wx.BoxSizer(wx.VERTICAL)
        sizer_18 = wx.BoxSizer(wx.VERTICAL)
        sizer_16 = wx.BoxSizer(wx.VERTICAL)
        sizer_15 = wx.BoxSizer(wx.VERTICAL)
        sizer_14 = wx.BoxSizer(wx.VERTICAL)
        grid_sizer_1 = wx.GridSizer(0, 2, 0, 0)
        photo_download_root_sizer = wx.StaticBoxSizer(wx.StaticBox(self.photo_download_page, wx.ID_ANY, _(u"\u7167\u7247\u4e0b\u8f7d\u5668")), wx.HORIZONTAL)
        photo_download_right_side = wx.BoxSizer(wx.VERTICAL)
        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        photo_download_left_side = wx.BoxSizer(wx.VERTICAL)
        photo_download_icons = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add((0, 0), 0, 0, 0)
        self.function_index_page.SetSizer(sizer_2)
        bitmap_3 = wx.StaticBitmap(self.photo_download_page, wx.ID_ANY, wx.Bitmap("Y:\\Downloads\\31469cae9bafa58e542ded0fdfc74004\\PIC\\AICBADGE copy.jpg", wx.BITMAP_TYPE_ANY))
        bitmap_3.SetMinSize((60, 60))
        photo_download_icons.Add(bitmap_3, 0, 0, 0)
        bitmap_2 = wx.StaticBitmap(self.photo_download_page, wx.ID_ANY, wx.Bitmap("Y:\\Downloads\\31469cae9bafa58e542ded0fdfc74004\\PIC\\QUABADGE copy.jpg", wx.BITMAP_TYPE_ANY))
        photo_download_icons.Add(bitmap_2, 0, 0, 0)
        photo_download_left_side.Add(photo_download_icons, 1, wx.EXPAND, 0)
        photo_download_left_side.Add((120, 30), 0, 0, 0)
        photo_download_left_side.Add(self.rbox_division, 0, wx.ALL | wx.EXPAND, 0)
        photo_download_root_sizer.Add(photo_download_left_side, 1, 0, 0)
        photo_download_right_side.Add((20, 10), 0, 0, 0)
        photo_download_label = wx.StaticText(self.photo_download_page, wx.ID_ANY, _(u"\u76d1\u7ba1\u4e91\n\u7167\u7247\u4e0b\u8f7d\u5668"), style=wx.ALIGN_CENTER)
        photo_download_label.SetMinSize((235, 57))
        photo_download_label.SetFont(wx.Font(20, wx.MODERN, wx.NORMAL, wx.BOLD, 0, "楷体_GB2312"))
        photo_download_right_side.Add(photo_download_label, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        photo_download_right_side.Add((20, 30), 0, 0, 0)
        static_line_3 = wx.StaticLine(self.photo_download_page, wx.ID_ANY)
        photo_download_right_side.Add(static_line_3, 0, wx.EXPAND, 0)
        photo_download_right_side.Add(self.photo_download_result_text, 0, wx.EXPAND, 1)
        label_1 = wx.StaticText(self.photo_download_page, wx.ID_ANY, _(u"\u4e0b\u8f7d\u8def\u5f84:"), style=wx.ALIGN_CENTER)
        label_1.SetFont(wx.Font(10, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_8.Add(label_1, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        sizer_8.Add(self.save_to_path_text_area, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        photo_download_right_side.Add(sizer_8, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_1 = wx.StaticLine(self.photo_download_page, wx.ID_ANY)
        photo_download_right_side.Add(static_line_1, 0, wx.EXPAND, 0)
        photo_download_right_side.Add(self.choose_download_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_2 = wx.StaticLine(self.photo_download_page, wx.ID_ANY)
        photo_download_right_side.Add(static_line_2, 0, wx.EXPAND, 0)
        photo_download_right_side.Add(self.begin_download_photo, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_4 = wx.StaticLine(self.photo_download_page, wx.ID_ANY)
        static_line_4.Enable(False)
        photo_download_right_side.Add(static_line_4, 0, wx.EXPAND, 0)
        photo_download_right_side.Add(self.cancel_download_photo, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        photo_download_root_sizer.Add(photo_download_right_side, 1, wx.EXPAND, 0)
        self.photo_download_page.SetSizer(photo_download_root_sizer)
        bitmap_6 = wx.StaticBitmap(self.doc_gen_page, wx.ID_ANY, wx.Bitmap("Y:\\Downloads\\31469cae9bafa58e542ded0fdfc74004\\PIC\\AICBADGE copy.jpg", wx.BITMAP_TYPE_ANY))
        bitmap_6.SetMinSize((60, 60))
        grid_sizer_1.Add(bitmap_6, 0, wx.ALIGN_RIGHT, 0)
        bitmap_7 = wx.StaticBitmap(self.doc_gen_page, wx.ID_ANY, wx.Bitmap("Y:\\Downloads\\31469cae9bafa58e542ded0fdfc74004\\PIC\\QUABADGE copy.jpg", wx.BITMAP_TYPE_ANY))
        grid_sizer_1.Add(bitmap_7, 0, 0, 0)
        doc_gen_left_side.Add(grid_sizer_1, 1, wx.EXPAND, 0)
        sizer_14.Add((20, 10), 0, 0, 0)
        label_7 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u73b0\u573a\u7b14\u5f55\u6a21\u677f:"), style=wx.ALIGN_CENTER)
        label_7.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_14.Add(label_7, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        sizer_14.Add(self.doc_gen_ins_tpl_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_left_side.Add(sizer_14, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        label_8 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u8bc1\u636e\u63d0\u53d6\u5355\u6a21\u677f:"))
        label_8.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_15.Add(label_8, 0, wx.ALIGN_CENTER, 0)
        sizer_15.Add(self.doc_gen_ph_tpl_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_left_side.Add(sizer_15, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        label_9 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u4f01\u4e1a\u4fe1\u606f\u53ca\u6838\u67e5\u8bb0\u5f55\u8868:"), style=wx.ALIGN_CENTER)
        label_9.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_16.Add(label_9, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        sizer_16.Add(self.doc_gen_xlsx_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_left_side.Add(sizer_16, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        label_12 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u6838\u67e5\u7167\u7247\u5b58\u653e\u76ee\u5f55:"), style=wx.ALIGN_CENTER)
        label_12.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_18.Add(label_12, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        sizer_18.Add(self.doc_gen_corp_photos_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_left_side.Add(sizer_18, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        label_10 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u6587\u4e66\u751f\u6210\u6839\u76ee\u5f55:"), style=wx.ALIGN_CENTER)
        label_10.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        sizer_17.Add(label_10, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        sizer_17.Add(self.doc_gen_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_left_side.Add(sizer_17, 1, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_15 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_left_side.Add(static_line_15, 0, wx.EXPAND, 0)
        doc_gen_left_side.Add(self.doc_gen_choose_inspect_tpl, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_13 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_left_side.Add(static_line_13, 0, wx.EXPAND, 0)
        doc_gen_left_side.Add(self.doc_gen_choose_photo_tpl, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_17 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        static_line_17.Enable(False)
        doc_gen_left_side.Add(static_line_17, 0, wx.EXPAND, 0)
        doc_gen_left_side.Add(self.doc_gen_choose_inspect_xlsx, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_14 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        static_line_14.Enable(False)
        doc_gen_left_side.Add(static_line_14, 0, wx.EXPAND, 0)
        doc_gen_left_side.Add(self.doc_gen_choose_corp_photos_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_root_sizer.Add(doc_gen_left_side, 1, wx.EXPAND, 0)
        doc_gen_right_side.Add((20, 10), 0, 0, 0)
        doc_gen_label_copy = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u76d1\u7ba1\u4e91\n\u6838\u67e5\u6587\u4e66\u751f\u6210\u5668"), style=wx.ALIGN_CENTER)
        doc_gen_label_copy.SetFont(wx.Font(20, wx.MODERN, wx.NORMAL, wx.BOLD, 0, "楷体_GB2312"))
        doc_gen_right_side.Add(doc_gen_label_copy, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_right_side.Add((20, 17), 0, 0, 0)
        label_11 = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u9009\u62e9\u6a21\u677f\u60c5\u51b5:"), style=wx.ALIGN_CENTER)
        label_11.SetFont(wx.Font(12, wx.MODERN, wx.NORMAL, wx.NORMAL, 0, "仿宋_GB2312"))
        doc_gen_right_side.Add(label_11, 0, wx.ALIGN_CENTER | wx.ALL, 0)
        doc_gen_right_side.Add((0, 0), 0, 0, 0)
        static_line_9 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_9, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.label_ins_tpl, 0, wx.EXPAND, 0)
        static_line_6 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_6, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.label_ph_tpl, 0, wx.EXPAND, 0)
        static_line_5 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_5, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.label_xlsx, 0, wx.EXPAND, 0)
        static_line_7 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_7, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add((0, 0), 0, 0, 0)
        label_doc_gen_result = wx.StaticText(self.doc_gen_page, wx.ID_ANY, _(u"\u5904\u7406\u60c5\u51b5"), style=wx.ALIGN_CENTER)
        label_doc_gen_result.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        doc_gen_right_side.Add(label_doc_gen_result, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add((20, 9), 0, 0, 0)
        doc_gen_right_side.Add(self.doc_gen_result_text_area, 0, wx.ALL | wx.EXPAND, 1)
        static_line_10 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_10, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.doc_gen_preview, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_16 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_16, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.doc_gen_choose_path, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_11 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        doc_gen_right_side.Add(static_line_11, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.doc_gen_begin, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        static_line_12 = wx.StaticLine(self.doc_gen_page, wx.ID_ANY)
        static_line_12.Enable(False)
        doc_gen_right_side.Add(static_line_12, 0, wx.EXPAND, 0)
        doc_gen_right_side.Add(self.doc_gen_stop, 0, wx.ALIGN_CENTER | wx.EXPAND, 0)
        doc_gen_root_sizer.Add(doc_gen_right_side, 1, wx.EXPAND, 0)
        self.doc_gen_page.SetSizer(doc_gen_root_sizer)
        self.Entrance.AddPage(self.function_index_page, _(u"\u529f\u80fd\u4ecb\u7ecd"))
        self.Entrance.AddPage(self.photo_download_page, _(u"\u7167\u7247\u4e0b\u8f7d\u5668"))
        self.Entrance.AddPage(self.doc_gen_page, _(u"\u6587\u4e66\u751f\u6210\u5668"))
        self.Entrance.AddPage(self.digital_hash_page, _(u"\u7535\u5b50\u6570\u636e\u6307\u7eb9\u751f\u6210\u5668"))
        sizer_1.Add(self.Entrance, 0, wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade

    '''
    以下为照片下载器用
    '''
    def begin_download_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        start_dl_btn = event.GetEventObject()
        cancel_dl_btn = self.cancel_download_photo
        choose_path_btn = self.choose_download_path

        choose_path_btn.Disable()
        choose_path_btn.SetLabel(u'连接和下载过程中请停止其他操作')
        start_dl_btn.Disable()
        cancel_dl_btn.Enable()

        self.update_pd_result_text('开始接连云服务器\n连接和下载过程中请停止其他操作')

        division_index = self.rbox_division.GetSelection()
        target_dir = re.sub(r'\s' , '', self.save_to_path_text_area.GetValue())
        try:
            os.mkdir(target_dir)
        except(FileExistsError):
            pass
        self.main_dl_thread = dp.photo_dl_thread(division_index, target_dir)
        
    def download_finished(self, result):
        if(isinstance(result, list)):
            if(result[0] == 'success'):
                self.choose_download_path.Enable()
                self.begin_download_photo.SetLabel(u'开始下载')
                self.begin_download_photo.Enable()
                self.update_pd_result_text('下载结果：\n' + result[1]) #打印下载结果
                self.cancel_download_photo.Disable()
                self.choose_download_path.SetLabel(u"\u66f4\u6539\u4e0b\u8f7d\u8def\u5f84\n")
            else:
                self.choose_download_path.Enable()
                self.begin_download_photo.SetLabel(u'开始下载')
                self.begin_download_photo.Enable()
                self.update_pd_result_text(result[1]) #打印错误信息
                self.cancel_download_photo.Disable()
                self.choose_download_path.SetLabel(u"\u66f4\u6539\u4e0b\u8f7d\u8def\u5f84\n")

    def choose_download_path_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.DirDialog(self,u"选择文件夹",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:  
            self.save_to_path_text_area.SetValue(dlg.GetPath()) #文件夹路径  
        dlg.Destroy()

    def rbox_div_select(self, event):  # wxGlade: MainFrame.<event_handler>
        self.update_pd_result_text('你当前选择的监管所：' + self.rbox_division.GetStringSelection())

    def update_pd_result_text(self, msg):
        original_msg = self.photo_download_result_text.GetValue()
        new_msg = msg + u"\n" + '*' * 15 +'\n' + original_msg + u'\n'
        self.photo_download_result_text.SetValue(new_msg)

    def cancel_download_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        self.update_pd_result_text('尝试终止下载。')
        if(self.main_dl_thread.stop()):
            self.cancel_download_photo.Disable()
        else:
            self.update_pd_result_text('尝试终止下载失败，仍在下载中。')
            

    '''
    照片下载器结束
    '''
    '''
    dg- doc_generator
    文书生成器用
    '''

    def choose_dg_xlsx_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.FileDialog(self, u"选择核查表", wildcard="xlsx files (*.xlsx)|*.xlsx",style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:  
            self.doc_gen_xlsx_path.SetValue(dlg.GetPath()) #核查表路径
            self.label_xlsx.SetBackgroundColour(wx.Colour(124, 252, 0))
            self.Refresh()
        dlg.Destroy()
    def choose_dg_ins_tpl_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.FileDialog(self, u"选择现场笔录模版", wildcard="docx files (*.docx)|*.docx",style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:  
            self.doc_gen_ins_tpl_path.SetValue(dlg.GetPath()) #笔录模版路径
            self.label_ins_tpl.SetBackgroundColour(wx.Colour(124, 252, 0))
            self.Refresh()
        dlg.Destroy()
    def choose_dg_ph_tpl_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.FileDialog(self, u"选择证据提取单模版", wildcard="docx files (*.docx)|*.docx",style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:  
            self.doc_gen_ph_tpl_path.SetValue(dlg.GetPath()) #笔录模版路径
            self.label_ph_tpl.SetBackgroundColour(wx.Colour(124, 252, 0))
            self.Refresh()
        dlg.Destroy()
    def choose_dg_path_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.DirDialog(self,u"选择文书生成文件夹",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:  
            self.doc_gen_path.SetValue(dlg.GetPath()) #文件夹路径  
        dlg.Destroy()

    def choose_dg_cp_phs_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        dlg = wx.DirDialog(self,u"选择企业核查照片存放文件夹",style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:  
            self.doc_gen_corp_photos_path.SetValue(dlg.GetPath()) #文件夹路径  
        dlg.Destroy()

    def begin_dg_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        print("Event handler 'begin_dg_btn' not implemented!")
        event.Skip()
    def cancel_dg_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        print("Event handler 'cancel_dg_btn' not implemented!")
        event.Skip()
    def choose_dg_preview_btn(self, event):  # wxGlade: MainFrame.<event_handler>
        print("Event handler 'choose_dg_preview_btn' not implemented!")
        event.Skip()
    
    def update_dg_resulttext(self, msg):
        original_msg = self.doc_gen_result_text_area.GetValue()
        new_msg = msg + u"\n" + '*' * 15 +'\n' + original_msg + u'\n'
        self.doc_gen_result_text_area.SetValue(new_msg)

    #页面转换时变更Frame大小
    def page_changing(self, event):  # wxGlade: MainFrame.<event_handler>
        if (event.GetSelection() == 1):
            self.SetSize((640, 539))
        elif (event.GetSelection() == 2):
            self.SetSize((730, 750))
# end of class MainFrame

class AICToolbox(wx.App):
    def OnInit(self):
        self.frame = MainFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        cwd = os.getcwd()
        '''
        以下为照片下载器用：
        设定当前文件夹为默认路径
        设定结果窗口启动内容
        '''
        self.frame.save_to_path_text_area.SetValue(cwd)
        self.frame.update_pd_result_text(u'成功启动\n请选择监管所；\n默认下载到本程序所在的文件夹。')
        #照片下载器用结束
        '''
        以下为文书生成器用：
        判断当前目录是否有预定的模版文件和核查记录表，名称、格式都需要正确
        '''
        self.frame.doc_gen_corp_photos_path.SetValue(cwd)
        self.frame.doc_gen_path.SetValue(cwd)
        for tpl_name in [u'现场笔录模板.docx', u'证据提取单模板.docx', u'企业信息及核查记录表.xlsx']:
            if os.path.isfile(cwd + '//' + tpl_name):
                if (tpl_name.count('笔录')):
                    self.frame.doc_gen_ins_tpl_path.SetValue(cwd + '\\' + tpl_name)
                    self.frame.label_ins_tpl.SetBackgroundColour(wx.Colour(124, 252, 0))

                elif (tpl_name.count('xlsx')):
                    self.frame.doc_gen_xlsx_path.SetValue(cwd + '\\' + tpl_name)
                    self.frame.label_xlsx.SetBackgroundColour(wx.Colour(124, 252, 0))
                else:
                    self.frame.doc_gen_ph_tpl_path.SetValue(cwd + '\\' + tpl_name)
                    self.frame.label_ph_tpl.SetBackgroundColour(wx.Colour(124, 252, 0))
            else:
                if (tpl_name.count('笔录')):
                    self.frame.doc_gen_ins_tpl_path.SetValue('当前目录下无‘' + tpl_name + '’文件，请点击‘选笔录模版’按键指定模板文件。')
                elif (tpl_name.count('xlsx')):
                    self.frame.doc_gen_xlsx_path.SetValue('当前目录下无‘' + tpl_name + '’文件，请点击‘选择核查表’按键指定核查表文件。')    
                else:
                    self.frame.doc_gen_ph_tpl_path.SetValue('当前目录下无‘' + tpl_name + '’文件，请点击‘选照片模版’按键指定模板文件。')
        #文书生成器用结束
        self.frame.Show()
        return True

# end of class AICToolbox

if __name__ == "__main__":
    gettext.install("AICToolbox") # replace with the appropriate catalog name

    AICToolbox = AICToolbox(0)
    AICToolbox.MainLoop()
