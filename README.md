# Программа для обработки типовых электронных документов

## Установка

python>=3.7

Требуемые библиотеки:
```bash
lxml==4.6.3
Pillow==8.2.0
PyQt5==5.15.4
PyQt5-Qt5==5.15.2
PyQt5-sip==12.8.1
PyQt5-stubs==5.15.2.0
python-docx==0.8.10
pywin32==300
xlrd==2.0.1
XlsxWriter==1.4.0
```
## Использование

В файле package/config.py необходимо указать данные для отправки отчетов по почте и при необходимости исправить рабочую директорию (по умолчанию C:\Ведомости).
Для отправки отчетов необходимо изменить переменную среды ENV на "RELEASE"

В папке example находятся пример ведомостей для обработки.
Их необходимо поместить в рабочую директорию (по умолчанию C:\Ведомости) и запустить main.py

В файле docs/guide.docx краткое руководство пользователя