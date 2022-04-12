"""
Скрипт для опроса сетевых устройств
"""
import openpyxl  # Библиотека для работы с Excel
from prettytable import PrettyTable
from colorama import init
from colorama import Fore, Back, Style
from ping3 import ping, verbose_ping
from tqdm import tqdm
import win10toast
from time import sleep
import os

clear = lambda: os.system('cls')
from tkinter import messagebox

ip_book = openpyxl.load_workbook("ip cam.xlsx")  # Открывает файл
worksheet = ip_book.active  # Делаем его активным
max_row = worksheet.max_row  # Получаем максимальное количество строк
max_col = worksheet.max_column  # Получаем максимальное количество столбцов

# for row in range (1, max_row):
#     if
# print(worksheet.max_row, worksheet.max_column)
# for i in range(0, worksheet.max_row):
# for col in worksheet.iter_cols(1, worksheet.max_column):
#     print(col[2].value, end="\t\t")
#     print('')
# print(worksheet.cell(row=1,column=5).value)
# for row in range(1, max_row):
#     val = worksheet.cell(row=row, column=5).value
#     if val == 'Камера':
#         print(worksheet.cell(row=row, column=1).value)
scan_ip = []
off_ip = []
init()
table_ip = PrettyTable()
table_ip.field_names = ['IP Адрес', 'Статус', 'Объект', 'Комментарий']


for row in range(1, max_row):  # Запускаем цикл по всем строкам
    val = worksheet.cell(row=row, column=5).value
    if val == 'Камера':  # Если устройство камера
        ip_cam = worksheet.cell(row=row, column=1).value
        ip_object = worksheet.cell(row=row, column=2).value
        ip_comment = worksheet.cell(row=row, column=6).value
        scan_ip.append([row, ip_cam, ip_object, ip_comment])


def modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_comment):
    flag_add = True
    for i in range(len(off_ip)):
        if ip_cam == off_ip[i][2]:
            if not flag_ip:
                off_ip[i][0] = 1
                flag_add = False
            else:
                off_ip[i][0] = 2

    if flag_add and not flag_ip:
        off_ip.append([0, row, ip_cam, ip_object, ip_comment])


def tab_ping():
    count_device = len(scan_ip)
    for status in tqdm(range(count_device)):
        row, ip_cam, ip_object , ip_comment = scan_ip[status]
        ip_pin = ping(ip_cam)
        flag_ip = True
        if ip_pin is None or type(ip_pin) is not float:
            flag_ip = False
            # ip_pin = ping(ip_cam)
            # if ip_pin is None or type(ip_pin) is not float:
            #     flag_ip = False

        modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_comment)
        if not flag_ip:
            table_ip.add_row([Fore.RED+ip_cam+Style.RESET_ALL, ip_pin, ip_object, ip_comment])
    print('Всего устройств : ' + str(count_device) + Fore.GREEN + '  На связи : ' + str(count_device-len(off_ip)) +
          Fore.RED + '  Отсутсвуют : ' + str(len(off_ip)) + Style.RESET_ALL)
    if len(off_ip)>0:
        print(table_ip)
    table_ip.clear_rows()
    print(off_ip)
    toaster = win10toast.ToastNotifier()
    off_text = ''
    on_text = ''
    del_index = []
    for flag, _, ip, obj, com in off_ip:
        if not flag:
            off_text += f'{ip[10:]},'
        if flag == 2:
            on_text += f'{ip[10:]},'
            del_index.append(off_ip.index([flag, _, ip, obj, com]))
            # off_ip.pop(off_ip.index([flag, _, ip, obj, com]))
    del_index.reverse()
    for _ in del_index:
        off_ip.pop(_)
    print(off_ip)
    if len(off_text)>0:
        toaster.show_toast("Отсутсвуют", off_text, icon_path='ping.ico', duration=5, threaded=True)
    if len(on_text)>0:
        toaster.show_toast("Восстановление связи", on_text, icon_path='ping.ico', duration=5, threaded=True)
# Надо запустить отдельным потоком
    print('До обновления (сек) :')
    for i in range(10, 0, -1):
        print("%03d" % i, end='')
        sleep(1)
        print('', end='\r')
while True:
    tab_ping()
    clear()
