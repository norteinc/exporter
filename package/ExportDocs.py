import json
import os
import platform
import re
import zipfile
from datetime import datetime

import pythoncom
import xlsxwriter
from PyQt5.QtCore import QThread, pyqtSignal
from docx import Document
from win32com import client as win32

from package.config import DIRNAME, DATE_DISTANCE
from package.dict import slovar
from package.service import isConnected, send_mail, add_mark, getvedtype, getshortmarkbyid, parsemark


class ExportDocs(QThread):
    update_progress_bar = pyqtSignal(int)
    lock_ui = pyqtSignal(bool)
    set_progress_bar = pyqtSignal(int, int, bool)
    add_string_to_activity_log = pyqtSignal(str)
    add_string_to_error_log = pyqtSignal(str)

    def __init__(self, selected_dir):
        super().__init__()
        self.selected_dir = selected_dir
        self.work_dir = DIRNAME + "\\.sys\\" + selected_dir
        self.dict_total = {}

    def run(self):
        self.lock_ui.emit(True)
        self.load_dict()
        self.check_dict()
        return

    def check_dict(self):
        result = True
        for group in self.dict_total.keys():
            check_FIO = None
            for j in self.dict_total[group]:
                if not check_FIO:
                    check_FIO = self.dict_total[group][j]["order"]
                    continue
                if check_FIO != self.dict_total[group][j]["order"]:
                    self.add_string_to_activity_log.emit("В ведомостях группы " + group + " ведомости на разное число студентов")
                    result = False
        return result


    def load_dict(self):
        dirs = os.listdir(self.work_dir)
        for dir in dirs:
            cur_dir = os.listdir(self.work_dir + "\\" + dir)
            self.dict_total[dir] = {}
            for file in cur_dir:
                current_file = self.work_dir + "\\" + dir + "\\" + file
                if not os.path.exists(current_file):
                    self.add_string_to_activity_log.emit("Папки " + current_file + " не существует")
                    return
                try:
                    dict_file = open(current_file, "r")
                except:
                    self.add_string_to_activity_log.emit("Не удалось открыть файл")
                    return
                try:
                    self.dict_total[dir][file] = json.load(dict_file)
                except:
                    self.add_string_to_error_log.emit("Файл поврежден")
