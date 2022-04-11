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
init()
table_ip = PrettyTable()
table_ip.field_names = ['IP Адрес', 'Статус', 'Комментарий']


for row in range(1, max_row):  # Запускаем цикл по всем строкам
    val = worksheet.cell(row=row, column=5).value
    if val == 'Камера':  # Если устройство камера
        ip_cam = worksheet.cell(row=row, column=1).value
        comment = worksheet.cell(row=row, column=6).value
        scan_ip.append([ip_cam, row, comment])




def tab_ping():

    off_ip = []
    count_device = len(scan_ip)
    for status in tqdm(range(count_device)):
        ip_cam, row, comment = scan_ip[status]
        ip_pin = ping(ip_cam)
        if ip_pin is None:
            ip_pin = ping(ip_cam)
            if ip_pin is None:
                off_ip.append([ip_cam, row, comment])
                table_ip.add_row([Fore.RED+ip_cam+Style.RESET_ALL, ip_pin, comment])
    print('Всего устройств : ' + str(count_device) + Fore.GREEN + '  На связи : ' + str(count_device-len(off_ip)) +
          Fore.RED + '  Отсутсвуют : ' + str(len(off_ip)) + Style.RESET_ALL)
    print(table_ip)
    table_ip.clear_rows()

    toaster = win10toast.ToastNotifier()
    error_text = ''
    for ip, _, com in off_ip:
        error_text +=f'{ip} - {com}\n'
    toaster.show_toast("Отсутсвуют", error_text, duration=5)
# Надо запустить отдельным потоком
    print('До обновления (сек) :')
    for i in range(10, 0, -1):
        print("%03d" % i, end='')
        sleep(1)
        print('', end='\r')
while True:
    tab_ping()
    clear()
