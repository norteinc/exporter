import os
import re
import smtplib
import ssl
from datetime import datetime
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mimetypes import guess_type

from package.config import MY_ADDRESS, TARGET_ADDRESS, SERVER, PORT, PASSWORD, DIRNAME
from package.dict import vedtypes, marks


def isConnected():
    """Проверка подключения к сети Интернет"""
    import urllib
    from urllib import request
    try:
        urllib.request.urlopen('https://yandex.ru')
        return True
    except:
        return False


def send_mail(filename):
    """Отправка письма"""
    message = MIMEMultipart()
    message["From"] = MY_ADDRESS
    message["To"] = TARGET_ADDRESS[0]
    message["Subject"] = "Stats from " + datetime.strftime(datetime.now(), "%d/%m/%Y %H:%M:%S")
    body = "Auto-Generated Stats."
    message.attach(MIMEText(body, "plain"))
    with open(filename, "rb") as attachment:
        mimetype, encoding = guess_type(filename)
        mimetype = mimetype.split('/', 1)
        fp = open(filename, 'rb')
        attachment = MIMEBase(mimetype[0], mimetype[1])
        attachment.set_payload(fp.read())
        fp.close()
        encode_base64(attachment)
        attachment.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(filename))
        message.attach(attachment)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SERVER, PORT, context=context) as server:
        server.login(MY_ADDRESS, PASSWORD)
        server.sendmail(MY_ADDRESS, TARGET_ADDRESS, text)


def add_mark(stats_counter, mark):
    """Добавление оценки в словарь"""
    if mark == 1:
        stats_counter['zach'] += 1
        stats_counter['total'] += 1
    if mark == 8:
        stats_counter['dopusk'] += 1
    if mark == 3 or mark == 13:
        stats_counter['three'] += 1
        stats_counter['total'] += 1
    if mark == 4 or mark == 14:
        stats_counter['four'] += 1
        stats_counter['total'] += 1
    if mark == 5 or mark == 15:
        stats_counter['five'] += 1
        stats_counter['total'] += 1
    if mark == 17 or mark == 18:
        stats_counter['ip'] += 1
        stats_counter['total'] += 1
    if mark == 2 or mark == 10 or mark == 11 or mark == 12 or mark == 16:
        stats_counter['two'] += 1
        stats_counter['total'] += 1


def checktypes():
    """Проверка соответствия ведомости и оценки"""
    for typex in vedtypes:
        for markx in marks:
            for curtype in range(len(typex)):
                if typex[curtype] == markx[2]:
                    print(typex, markx[0])


def getvedtype(vedtype):
    """Полуечние типа ведомости по названию"""
    for cur in range(len(vedtypes)):
        if vedtypes[cur][0] == vedtype:
            return cur
    return -1


def getmarkbyid(markid):
    """Получение текста оценки по номеру"""
    for cur in range(len(marks)):
        if marks[cur][2] == markid:
            return marks[cur][0]
    return -1


def getshortmarkbyid(markid):
    """Получение короткой оценки по номеру"""
    for cur in range(len(marks)):
        if marks[cur][2] == markid:
            return marks[cur][1]
    return -1


def parsemark(markplace, getvedtypeid, pos):  # анализ оценки
    """Проверка бизнес-логики оценок"""
    if getvedtypeid == 0:  # зачетные ведомости
        if (re.findall("НЕ", markplace) and re.findall("ЗАЧ", markplace)) or re.findall('НЕТ', markplace) or (
                re.findall("-", markplace) and len(markplace) == 1) or re.findall("Н([-_ /.]*)З", markplace):
            return 2
        if re.findall("ПЕРЕЗАЧ", markplace):
            return 18
        if re.findall("ЗАЧ", markplace) or re.findall('ДА', markplace) or re.findall('[+]', markplace):
            return 1
        if re.findall("ИНД", markplace) or re.findall("ПЛАН", markplace) or re.findall("ИП", markplace):
            return 17
        if re.findall("Н([-_ /.]*)Д", markplace) or (re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 19
        return -1
    if getvedtypeid == 1:  # дифзачетные ведомости
        if (re.findall("НЕ", markplace) and re.findall("ЗАЧ", markplace)) or re.findall('НЕТ', markplace) or (
                re.findall("-", markplace) and len(markplace) == 1) or re.findall("Н([-_ /.]*)З", markplace):
            return 2
        if (re.findall("ЗАЧ", markplace) and re.findall("УД", markplace)) or re.findall("3", markplace) or re.findall(
                "УД", markplace):
            return 3
        if (re.findall("ЗАЧ", markplace) and re.findall("ХОР", markplace)) or re.findall("4", markplace) or re.findall(
                "ХОР", markplace):
            return 4
        if (re.findall("ЗАЧ", markplace) and re.findall("ОТЛ", markplace)) or re.findall("5", markplace) or re.findall(
                "ОТЛ", markplace):
            return 5
        if re.findall("ИНД", markplace) or re.findall("ПЛАН", markplace) or re.findall("ИП", markplace):
            return 17
        if re.findall("ПЕРЕЗАЧ", markplace):
            return 18
        if re.findall("Н([-_ /.]*)Д", markplace) or (re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 19
        return -1
    if getvedtypeid == 2 and pos == 1:  # сдано / не сдано
        if (re.findall("НЕ", markplace) and re.findall("СД", markplace)) or re.findall('НЕТ', markplace) or (
                re.findall("-", markplace) and len(markplace) == 1) or re.findall("Н([-_ /.]*)С", markplace) or (
                re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 9
        if (re.findall("СД", markplace) and not re.findall('НЕ', markplace)) or re.findall('[+]',
                                                                                           markplace) or re.findall(
            'ДА', markplace) or (re.findall('ДОП', markplace) and not re.findall('НЕ', markplace)):
            return 8
        if re.findall("ИНД", markplace) or re.findall("ПЛАН", markplace) or re.findall("ИП", markplace):
            return 17
        if re.findall("ПЕРЕЗАЧ", markplace):
            return 18
        if re.findall("Н([-_ /.]*)Д", markplace) or (re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 19
        return -1
    if getvedtypeid == 2 and pos == 2:  # экзамен / оценки
        if (re.findall("НЕ", markplace) and re.findall("ДОП", markplace)) or re.search("Н([-_ /.]*)Д", markplace):
            return 10
        if (re.findall("НЕ", markplace) and re.findall("ЯВ", markplace)) or re.search("Н([-_ /.]*)Я", markplace):
            return 11
        if (re.findall("УД", markplace) and re.findall("НЕ", markplace)) or re.findall("2", markplace) or re.findall(
                "ДВА", markplace):
            return 12
        if re.findall("УД", markplace) or re.findall("3", markplace):
            return 13
        if re.findall("ХОР", markplace) or re.findall("4", markplace):
            return 14
        if re.findall("ОТЛ", markplace) or re.findall("5", markplace):
            return 15
        if re.findall("ИНД", markplace) or re.findall("ПЛАН", markplace) or re.findall("ИП", markplace):
            return 17
        if re.findall("ПЕРЕЗАЧ", markplace):
            return 18
        if re.findall("Н([-_ /.]*)Д", markplace) or (re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 19
        return -1
    if getvedtypeid == 3 or getvedtypeid == 4:  # КП КР
        if re.findall("УД", markplace) or re.findall("3", markplace):
            return 13
        if re.findall("ХОР", markplace) or re.findall("4", markplace):
            return 14
        if re.findall("ОТЛ", markplace) or re.findall("5", markplace):
            return 15
        if re.findall("НЕ", markplace) or re.findall("2", markplace):
            return 16
        if re.findall("ИНД", markplace) or re.findall("ПЛАН", markplace) or re.findall("ИП", markplace):
            return 17
        if re.findall("ПЕРЕЗАЧ", markplace):
            return 18
        if re.findall("Н([-_ /.]*)Д", markplace) or (re.findall('ДОП', markplace) and re.findall('НЕ', markplace)):
            return 19
        return -1


def validved(markid1, markid2):
    """Валидация оценки и ведомости"""
    if markid1 == 17 or markid2 == 17 or markid1 == 18 or markid2 == 18 or markid1 == 19 or markid2 == 19:
        return 1
    if markid1 == 9 and markid2 == 10:
        return 1
    if markid1 == 8 and (markid2 == 11 or markid2 == 12 or markid2 == 13 or markid2 == 14 or markid2 == 15):
        return 1
    return 0


def get_count_todo():
    """Получение числа ведомостей для обработки"""
    if not os.path.exists(DIRNAME):
        return 0
    files = os.listdir(DIRNAME)
    count_files = 0
    for cur_file in files:
        if re.search("docx|doc|rtf|odt", cur_file):
            count_files = count_files + 1
    return count_files


def delete_empty_folders():
    """Удаление пустых директорий"""
    resPath = DIRNAME + '\\Поврежденные ведомости'
    if os.path.exists(resPath):
        files = os.listdir(resPath)
        if len(files) == 0:
            os.rmdir(resPath)
    resPath = DIRNAME + '\\Некорретные ведомости'
    if os.path.exists(resPath):
        files = os.listdir(resPath)
        if len(files) == 0:
            os.rmdir(resPath)
    return


def get_count_ready():
    """Получение числа подготовленных ведомостей"""
    resPath = DIRNAME + '\\Готовы к выгрузке'
    if not (os.path.exists(resPath)):
        return 0
    files = os.listdir(resPath)
    count_files = 0
    for cur_file in files:
        if re.search("docx", cur_file):
            count_files = count_files + 1
    return count_files


def get_count_corrupt():
    """Получение числа поврежденных ведомостей"""
    resPath = DIRNAME + '\\Поврежденные ведомости'
    if not (os.path.exists(resPath)):
        # print("Папка 'C:\Ведомости\Поврежденные ведомости' не найдена.")
        return 0
    files = os.listdir(resPath)
    count_files = 0
    for cur_file in files:
        if re.search("docx", cur_file):
            count_files = count_files + 1
    return count_files


def get_count_incorrect():
    """Получение числа некорректных ведомостей"""
    resPath = DIRNAME + '\\Некорретные ведомости'
    if not (os.path.exists(resPath)):
        # print("Папка 'C:\Ведомости\Некорретные ведомости' не найдена.")
        return 0
    files = os.listdir(resPath)
    count_files = 0
    for cur_file in files:
        if re.search("docx", cur_file):
            count_files = count_files + 1
    return count_files