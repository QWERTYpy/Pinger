"""
Скрипт для опроса сетевых устройств
"""
import openpyxl  # Библиотека для работы с Excel
from prettytable import PrettyTable
from colorama import init
from colorama import Fore, Back, Style
from ping3 import ping, verbose_ping
from tqdm import tqdm
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

init()
table_ip = PrettyTable()
table_ip.field_names = ['IP Адрес', 'Статус', 'Комментарий']
scan_ip=[]
off_ip=[]
count_device = 0
for row in range(1, max_row):  # Запускаем цикл по всем строкам
    val = worksheet.cell(row=row, column=5).value
    if val == 'Камера':  # Если устройство камера

        count_device+=1
        ip_cam = worksheet.cell(row=row, column=1).value
        comment = worksheet.cell(row=row, column=6).value
        ip_pin = ping(ip_cam)
        if ip_pin is None:
            ip_pin = ping(ip_cam)
            if ip_pin is None:
                off_ip.append([ip_cam, row, comment])
                table_ip.add_row([Fore.RED+ip_cam+Style.RESET_ALL, ip_pin, comment])
print('Всего устройств : ' + str(count_device) + Fore.GREEN + '  На связи : ' + str(count_device-len(off_ip)) +
      Fore.RED + '  Отсутсвуют : ' + str(len(off_ip)) + Style.RESET_ALL)
print(table_ip)

import win10toast
toaster = win10toast.ToastNotifier()
error_text = ''
for ip, _, com in off_ip:
    error_text +=f'{ip} - {com}\n'
toaster.show_toast("Отсутсвуют", error_text, duration=100)
#messagebox.showwarning('Отсутствуют', off_ip)
# print('Проверка недоступных устройств')
# for ip, row in off_ip:
#     print(Fore.RED+ip+Style.RESET_ALL)
#     verbose_ping(ip)
