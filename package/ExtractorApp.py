import json
import os
import re

import xlrd
from datetime import datetime
from docx import Document

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog, QInputDialog

from package.ExportBills import ExportBills
from package.ExportDocs import ExportDocs
from package.ValidateBills import ValidateBills
from gui import design
from package.config import DIRNAME
from package.dict import slovar_en_ru
from package.service import get_count_todo, get_count_ready, get_count_incorrect, get_count_corrupt, \
    delete_empty_folders


class ExtractorApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.dict_FIO = {}
        self.setupUi(self)
        if self.init_ui() == -1:
            self.deactivate_ui()
        else:
            self.update_status()
        self.btnRefresh.clicked.connect(self.update_status)
        self.btnPrepare.clicked.connect(self.validate_bills_run)
        self.btnExport.clicked.connect(self.export_bills_run)
        self.btnTodo.clicked.connect(self.open_folder_todo)
        self.btnIncorrect.clicked.connect(self.open_folder_incorrect)
        self.btnCorrupt.clicked.connect(self.open_folder_corrupt)
        self.btnReady.clicked.connect(self.open_folder_ready)
        self.calDate.selectionChanged.connect(self.update_current_date)
        self.btnLoadFIO.clicked.connect(self.load_FIO)
        self.btnStats.clicked.connect(self.get_stats)
        self.btnExportDocs.clicked.connect(self.export_docs_prepare)

    def init_ui(self):
        if not (os.path.exists(DIRNAME)):
            self.add_string_to_activity_log("Папка 'C:\Ведомости' не найдена.")
            self.add_string_to_activity_log("Проверьте наличие папки и перезапустите программу.")
            return -1
        self.progress.setVisible(False)
        self.listErrors.setAutoScroll(True)
        self.listLog.setAutoScroll(True)
        self.cbTodo.setEnabled(False)
        self.cbReady.setEnabled(False)
        self.cbIncorrect.setEnabled(False)
        self.cbCorrupt.setEnabled(False)
        self.cbDate.setEnabled(False)
        self.cbLoadFIO.setEnabled(False)
        self.add_string_to_activity_log("Интерфейс успешно загружен")
        self.load_dict()
        QtCore.QCoreApplication.processEvents()
        return 0

    def load_FIO(self):
        """Загружаем файл с ФИО"""
        self.add_string_to_activity_log("Открываем окно для выбора файла")
        file_name = self.openfile_widget()
        if file_name:
            self.add_string_to_activity_log("Файл " + file_name + " выбран для загрузки")
        else:
            self.add_string_to_activity_log("Файл не выбран")
            return
        self.parse_FIO_file(file_name)
        self.save_dict()

    def save_dict(self):
        """Сохранение словаря с ФИО"""
        self.add_string_to_activity_log("Сохраняем словарь для дальнейшего использования")
        sys_dir = DIRNAME + "\\" + ".sys"
        if not os.path.exists(sys_dir):
            os.mkdir(sys_dir)
        try:
            a_file = open(sys_dir + "\\" + "FIO.json", "w")
        except:
            self.add_string_to_error_log("Не удалось сохранить данные в словарь")
            return
        json.dump(self.dict_FIO, a_file, ensure_ascii=False)
        a_file.close()

    def load_dict(self):
        """Загрузка словаря с ФИО"""
        self.add_string_to_activity_log("Ищем словарь для загрузки")
        sys_dir = DIRNAME + "\\" + ".sys"
        if not os.path.exists(sys_dir):
            self.add_string_to_activity_log("Системная папка не найдена")
            return
        try:
            a_file = open(sys_dir + "\\" + "FIO.json", "r")
        except:
            self.add_string_to_activity_log("Словарь c ФИО не найден")
            return
        try:
            self.dict_FIO = json.load(a_file)
        except:
            self.add_string_to_error_log("Словарь поврежден")
        self.add_string_to_activity_log(
            "Из словаря загружено " + str(self.get_dict_size()) + " записей ФИО и зачетных книжек")
        self.cbLoadFIO.setChecked(True)
        a_file.close()

    def parse_FIO_file(self, file_name):
        """Разбор файла с ФИО"""
        self.add_string_to_activity_log("Начинаем обработку файла с ФИО")
        try:
            rb = xlrd.open_workbook(file_name, formatting_info=True)
        except:
            self.add_string_to_error_log("Не удалось открыть файл")
            return
        try:
            sheet = rb.sheet_by_index(0)
        except:
            self.add_string_to_error_log("Файл поврежден")
            return
        if sheet.nrows < 2 or sheet.ncols < 6:
            self.add_string_to_error_log("Выбран неподходящий файл")
        try:
            row = sheet.row_values(1)
            if row[0] != "№" or row[1] != "Фамилия" or row[2] != "Имя" or row[3] != "Отчество" or row[
                4] != "номер зач. книжки":
                raise Exception
        except:
            self.add_string_to_error_log("Файл с ФИО неверно выгружен")
            return
        for rownum in range(2, sheet.nrows):
            row = sheet.row_values(rownum)
            self.dict_FIO[row[4]] = row[1] + ' ' + row[2] + ' ' + row[3]
        self.add_string_to_activity_log(
            "Успешно обновлено " + str(self.get_dict_size()) + " записей ФИО и зачетных книжек")
        self.cbLoadFIO.setChecked(True)

    def get_stats(self):
        """Получение статистики подготовленных ведомостей"""
        if not get_count_ready():
            self.add_string_to_activity_log("Нет подготовленных ведомостей")
            return
        self.add_string_to_activity_log("Пожалуйста, подождите, идет подсчет ведомостей")
        files = os.listdir(DIRNAME + '\\Готовы к выгрузке')
        dict_count = {}
        for cur_file in files:
            if re.findall('~', cur_file):  # проверка на временные файлы
                continue
            if not re.search("docx", cur_file):  # нам интересны только доки
                continue
            result = re.search('[А-Яа-я]\d{1,2}_', cur_file)  # ищем название на русском
            if result:
                kaf = result[0][:-1]  # убираем последний _
            else:
                result = re.search('[A-Za-z]\d{1,2}_', cur_file)  # ищем название на английском
                if result:
                    kaf = result[0][:-1]  # убираем последний _
                    for key in slovar_en_ru:  # заменяем английский на русский
                        if key.isalpha():
                            kaf = kaf.replace(key, slovar_en_ru[key])
                else:  # получаем название кафедры из файла
                    document = Document(DIRNAME + '\\Готовы к выгрузке' + "\\" + cur_file)
                    table = document.tables[0]
                    kaf = table.rows[6].cells[4].text.strip()
            if kaf not in dict_count.keys():
                dict_count[kaf] = 1
            else:
                dict_count[kaf] += 1
        for key in sorted(dict_count.keys()):
            self.add_string_to_activity_log("Кафедра " + key + " - " + str(dict_count[key]) + " шт.")

    def get_dict_size(self):
        """Получение размера словаря"""
        return len(self.dict_FIO.keys())

    def openfile_widget(self):
        """Открытие виджета для выбора файла"""
        return QFileDialog.getOpenFileName(self, 'Загрузка ФИО', DIRNAME, "Выгрузка из ИС УМУ (*.xls)")[0]

    def update_current_date(self):
        """Обновление текущей даты"""
        selected_date = self.calDate.selectedDate().toString("dd.MM.yyyy")
        self.add_string_to_activity_log("День конца сессии изменен на " + selected_date)
        self.cbDate.setChecked(True)
        self.update_status()

    def add_string_to_activity_log(self, add_str):
        """Добавление строки в лог активности"""
        self.listLog.addItem(datetime.strftime(datetime.now(), "%H:%M:%S > ") + add_str)
        self.listLog.repaint()
        self.listLog.scrollToBottom()

    def add_string_to_error_log(self, add_str):
        """Добавление строки в лог ошибок"""
        self.listErrors.addItem(datetime.strftime(datetime.now(), "%H:%M:%S > ") + add_str)
        self.listErrors.repaint()
        self.listErrors.scrollToBottom()

    def deactivate_ui(self):
        """Отключение интерфейса при отсутствии рабочей директории"""
        self.cbTodo.setChecked(False)
        self.cbReady.setChecked(False)
        self.cbIncorrect.setChecked(False)
        self.cbCorrupt.setChecked(False)
        self.btnExport.setEnabled(False)
        self.btnRefresh.setEnabled(False)
        self.btnPrepare.setEnabled(False)

    def update_status(self):
        """Обновление всех счетчиков файлов"""
        self.btnRefresh.setDisabled(True)
        cnt_todo = get_count_todo()
        cnt_ready = get_count_ready()
        cnt_incorrect = get_count_incorrect()
        cnt_corrupt = get_count_corrupt()
        self.cntTodo.setText(str(cnt_todo))
        self.cntReady.setText(str(cnt_ready))
        self.cntIncorrect.setText(str(cnt_incorrect))
        self.cntCorrupt.setText(str(cnt_corrupt))
        self.btnTodo.setDisabled(True)
        self.btnReady.setDisabled(True)
        self.btnCorrupt.setDisabled(True)
        self.btnIncorrect.setDisabled(True)
        activate_button = 0
        if cnt_todo != 0:
            self.cbTodo.setChecked(False)
            self.btnTodo.setDisabled(False)
        else:
            self.cbTodo.setChecked(True)
            activate_button += 1
        if cnt_ready > 0:
            self.cbReady.setChecked(True)
            activate_button += 1
            self.btnReady.setDisabled(False)
        else:
            self.cbReady.setChecked(False)
        if cnt_incorrect == 0:
            self.cbIncorrect.setChecked(True)
            activate_button += 1
        else:
            self.btnIncorrect.setDisabled(False)
            self.cbIncorrect.setChecked(False)
        if cnt_corrupt == 0:
            self.cbCorrupt.setChecked(True)
            activate_button += 1
        else:
            self.btnCorrupt.setDisabled(False)
            self.cbCorrupt.setChecked(False)
        if activate_button == 4 and self.cbDate.isChecked():
            self.btnExport.setEnabled(True)
        else:
            self.btnExport.setEnabled(False)
        self.add_string_to_activity_log("Количество документов - данные обновлены")
        delete_empty_folders()
        self.btnRefresh.setDisabled(False)

    def open_folder_todo(self):
        """Открытие рабочей папки"""
        if os.path.exists(DIRNAME):
            path = os.path.realpath(DIRNAME)
            os.startfile(path)

    def open_folder_ready(self):
        """Открытие готовых файлов"""
        if os.path.exists(DIRNAME + '\\Готовы к выгрузке'):
            path = os.path.realpath(DIRNAME + '\\Готовы к выгрузке')
            os.startfile(path)

    def open_folder_corrupt(self):
        """Открытие поврежеденных файлов"""
        if os.path.exists(DIRNAME + '\\Поврежденные ведомости'):
            path = os.path.realpath(DIRNAME + '\\Поврежденные ведомости')
            os.startfile(path)

    def open_folder_incorrect(self):
        """Открытие некорректных файлов"""
        if os.path.exists(DIRNAME + '\\Некорретные ведомости'):
            path = os.path.realpath(DIRNAME + '\\Некорретные ведомости')
            os.startfile(path)

    def validate_bills_run(self):
        """Создание потока для валидации файлов"""
        self.ValidateBillsWorker = ValidateBills()
        self.ValidateBillsWorker.start()
        self.ValidateBillsWorker.finished.connect(self.unlock_ui)
        self.ValidateBillsWorker.lock_ui.connect(self.lock_buttons)
        self.ValidateBillsWorker.update_progress_bar.connect(self.event_update_progress_bar)
        self.ValidateBillsWorker.set_progress_bar.connect(self.event_set_progress_bar)
        self.ValidateBillsWorker.add_string_to_activity_log.connect(self.event_add_string_to_activity_log)
        self.ValidateBillsWorker.add_string_to_error_log.connect(self.event_add_string_to_error_log)

    def export_bills_run(self):
        """Создание потока для экпорта файлов"""
        self.ExportBillsWorker = ExportBills(self.calDate.selectedDate().toString("dd.MM.yyyy"))
        self.ExportBillsWorker.start()
        self.ExportBillsWorker.finished.connect(self.unlock_ui)
        self.ExportBillsWorker.lock_ui.connect(self.lock_buttons)
        self.ExportBillsWorker.update_progress_bar.connect(self.event_update_progress_bar)
        self.ExportBillsWorker.set_progress_bar.connect(self.event_set_progress_bar)
        self.ExportBillsWorker.add_string_to_activity_log.connect(self.event_add_string_to_activity_log)
        self.ExportBillsWorker.add_string_to_error_log.connect(self.event_add_string_to_error_log)

    def export_docs_prepare(self):
        """Подготовка для экпорта документов"""
        sys_dir = DIRNAME + "\\" + ".sys"
        if not os.path.exists(sys_dir):
            self.add_string_to_activity_log("Нет выгруженных ведомостей.")
            return
        directories = [d for d in os.listdir(sys_dir) if os.path.isdir(os.path.join(sys_dir, d))]
        selected_dir, ok = QInputDialog.getItem(self, "Период сессии", "Выберите сессионный период", directories, 0,
                                                False)
        if not ok:
            return
        self.export_docs_run(selected_dir)

    def export_docs_run(self, val):
        """Создание потока для экспорта документов"""
        self.ExportDocsWorker = ExportDocs(val)
        self.ExportDocsWorker.start()
        self.ExportDocsWorker.finished.connect(self.unlock_ui)
        self.ExportDocsWorker.lock_ui.connect(self.lock_buttons)
        self.ExportDocsWorker.update_progress_bar.connect(self.event_update_progress_bar)
        self.ExportDocsWorker.set_progress_bar.connect(self.event_set_progress_bar)
        self.ExportDocsWorker.add_string_to_activity_log.connect(self.event_add_string_to_activity_log)
        self.ExportDocsWorker.add_string_to_error_log.connect(self.event_add_string_to_error_log)

    def lock_buttons(self, val):
        """Блокировка кнопок на время работы потока"""
        self.btnExport.setDisabled(val)
        self.btnRefresh.setDisabled(val)
        self.btnPrepare.setDisabled(val)

    def unlock_ui(self):
        """Разблокировка интерфейса после работы потока"""
        cnt_todo = get_count_todo()
        cnt_ready = get_count_ready()
        cnt_incorrect = get_count_incorrect()
        cnt_corrupt = get_count_corrupt()
        self.cntTodo.setText(str(cnt_todo))
        self.cntReady.setText(str(cnt_ready))
        self.cntIncorrect.setText(str(cnt_incorrect))
        self.cntCorrupt.setText(str(cnt_corrupt))
        self.btnTodo.setDisabled(True)
        self.btnReady.setDisabled(True)
        self.btnCorrupt.setDisabled(True)
        self.btnIncorrect.setDisabled(True)
        activate_button = 0
        if cnt_todo != 0:
            self.cbTodo.setChecked(False)
            self.btnTodo.setDisabled(False)
        else:
            self.cbTodo.setChecked(True)
            activate_button += 1
        if cnt_ready > 0:
            self.cbReady.setChecked(True)
            activate_button += 1
            self.btnReady.setDisabled(False)
        else:
            self.cbReady.setChecked(False)
        if cnt_incorrect == 0:
            self.cbIncorrect.setChecked(True)
            activate_button += 1
        else:
            self.btnIncorrect.setDisabled(False)
            self.cbIncorrect.setChecked(False)
        if cnt_corrupt == 0:
            self.cbCorrupt.setChecked(True)
            activate_button += 1
        else:
            self.btnCorrupt.setDisabled(False)
            self.cbCorrupt.setChecked(False)
        if activate_button == 4 and self.cbDate.isChecked():
            self.btnExport.setEnabled(True)
        else:
            self.btnExport.setEnabled(False)
        self.btnRefresh.setDisabled(False)
        self.btnPrepare.setDisabled(False)
        delete_empty_folders()

    def event_set_progress_bar(self, val1, val2, val3):
        """Событие для установки пределов прогресс-бара"""
        self.progress.setVisible(val3)
        self.progress.setMinimum(val1)
        self.progress.setMaximum(val2)

    def event_update_progress_bar(self, val):
        """Событие для обнловления прогресс-бара"""
        self.progress.setValue(val)

    def event_add_string_to_activity_log(self, val):
        """Событие для добавления строки в лог активности"""
        self.add_string_to_activity_log(val)

    def event_add_string_to_error_log(self, val):
        """Событие для добавления строки в лог ошибок"""
        self.add_string_to_error_log(val)
