"""
Скрипт для опроса сетевых устройств
"""
import sys

import openpyxl  # Библиотека для работы с таблицами
from prettytable import PrettyTable
from colorama import init
from colorama import Fore, Style
from ping3 import ping
from tqdm import tqdm  # Библиотека для рисования шкалы выполнения
import win10toast  # Библиотека для всплывающих сообщений
import time
import os
import msvcrt


ip_book = openpyxl.load_workbook("ip cam.xlsx")  # Открывает файл
worksheet = ip_book.active  # Делаем его активным
max_row = worksheet.max_row  # Получаем максимальное количество строк
max_col = worksheet.max_column  # Получаем максимальное количество столбцов

# Инициализация начальных данных
sec = 600  # По умолчанию перезагрузка через 600 сек
first_start = True  # Флаг первичной загрузки данных
scan_ip = []  # Список всех устройств
off_ip = []  # Список отсутствующих
init()  # Проводим инициализацию для цветного оформления
table_ip = PrettyTable()  # Таблица с устройствами off-line
table_priority = PrettyTable()  # Таблица с приоритетными устройствами
table_error = PrettyTable()  # Таблица с отключенными устройствами
table_ip.field_names = ['IP Адрес', 'Статус', 'Объект', 'Тип', 'Комментарий']
table_priority.field_names = ['IP Адрес', 'Статус', 'Объект', 'Тип', 'Комментарий']
table_error.field_names = ['IP Адрес', 'Объект', 'Тип', 'Комментарий']


def excel_to_list():
    """
    Функция заполняет список устройствами для опроса

    :return:
    """
    for row in range(1, max_row):  # Запускаем цикл по всем строкам
        val = worksheet.cell(row=row, column=5).value
        if val == 'Камера' or val == 'Хранилище' or val == 'ПК' \
                or val == 'Видеорегистратор' or val == 'Wi-Fi' or val == 'Преобразователь':  # Если устройство камера
            ip_cam = worksheet.cell(row=row, column=1).value
            ip_object = worksheet.cell(row=row, column=2).value
            ip_type = worksheet.cell(row=row, column=5).value
            ip_comment = worksheet.cell(row=row, column=6).value
            ip_priority = worksheet.cell(row=row, column=7).value
            ip_active = worksheet.cell(row=row, column=8).value
            scan_ip.append([row, ip_cam, ip_object, ip_type, ip_comment, ip_priority, ip_active])


def modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_type, ip_comment):
    """
    Данная функция изменяет список с отсутствующими камерами, добавляя новые или помечая на удаление старые

    :param flag_ip: bool Ответило устройство на пинг или нет
    :param row: Строка в таблице с устройством
    :param ip_cam: IP устройства
    :param ip_object: Объект на котором установлено устройство
    :param ip_type: Тип устройства
    :param ip_comment: Комментарий
    :return: None
    """
    flag_add = True  # По умолчанию новую запись помечаем на добавление
    for i in range(len(off_ip)):
        if ip_cam == off_ip[i][2]:  # Если устройство присутствует в списке
            if not flag_ip:  # Если оно снова не ответило на пинг отключаем уведомление
                off_ip[i][0] = 1
            else:
                off_ip[i][0] = 2  # Если ответило на пинг, помечаем на удаление из списка
            flag_add = False

    if flag_add and not flag_ip:  # Если устройства нет в списке добавляем
        off_ip.append([0, row, ip_cam, ip_object, ip_type, ip_comment])


def print_time_input(timeout):
    """
    Данная функция выводит обратный таймер и следит за вводом для управления программой

    :param timeout: Значение таймера

    :return:
    """
    global sec
    start = time.monotonic()
    now = time.monotonic()
    count_sec = 0
    while now - start < timeout:
        if now - start > count_sec:
            count_sec += 1
            print_time = timeout-count_sec
            print("\rОсталось %03d c. | u-обновить | t-таймер | q-выход :" % print_time, end='')
        if msvcrt.kbhit():
            c = msvcrt.getwch()
            if c == 'q':
                sys.exit()
            if c == 'u':
                break
            if c == 't':
                sec = int(input('Введите количество секунд :'))
                break

        now = time.monotonic()


def tab_ping():
    """

    :return:
    """
    count_device = len(scan_ip)
    count_error = 0
    global first_start
    first_start = False
    for status in tqdm(range(count_device), desc='PING'):
        row, ip_cam, ip_object, ip_type, ip_comment, ip_priority, ip_active = scan_ip[status]
        if ip_active == 'ON':
            ip_pin = ping(ip_cam)
            flag_ip = True
            if ip_pin is None or type(ip_pin) is not float:
                flag_ip = False

            modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_type, ip_comment)
            if ip_priority == 'True':
                if flag_ip:
                    table_priority.add_row([Fore.GREEN+ip_cam+Style.RESET_ALL,
                                            flag_ip, ip_object, ip_type, ip_comment])
                else:
                    table_priority.add_row([Fore.RED + ip_cam + Style.RESET_ALL,
                                            flag_ip, ip_object, ip_type, ip_comment])
            if not flag_ip:
                table_ip.add_row([Fore.RED+ip_cam+Style.RESET_ALL, ip_pin, ip_object, ip_type, ip_comment])
        if ip_active == 'OFF':
            count_error += 1
            table_error.add_row([Fore.RED+ip_cam+Style.RESET_ALL, ip_object, ip_type, ip_comment])
    print('Приоритетные устройства')
    print(table_priority)
    print('Неисправные устройства')
    print(table_error)
    print('Всего устройств : ' + str(count_device) + Fore.GREEN + '  На связи : ' +
          str(count_device-len(off_ip)-count_error) +
          Fore.RED + '  Отсутствуют : ' + str(len(off_ip)) + Fore.BLUE +
          ' Отключены : ' + str(count_error) + Style.RESET_ALL)
    if len(off_ip) > 0:
        print('Устройства не в сети')
        print(table_ip)
    table_ip.clear_rows()
    table_priority.clear_rows()
    table_error.clear_rows()
    toaster = win10toast.ToastNotifier()
    off_text = ''
    on_text = ''
    del_index = []

    for flag, _, ip, obj, tipe, com in off_ip:
        if not flag:
            off_text += f'{ip[10:]},'
        if flag == 2:
            on_text += f'{ip[10:]},'
            del_index.append(off_ip.index([flag, _, ip, obj, tipe, com]))
            # off_ip.pop(off_ip.index([flag, _, ip, obj, com]))
    del_index.reverse()
    for _ in del_index:
        off_ip.pop(_)
    if len(off_text) > 0:
        toaster.show_toast("Отсутствуют", off_text[:-1], icon_path='ping.ico', duration=5, threaded=True)
    if len(on_text) > 0:
        toaster.show_toast("Восстановление связи", on_text[:-1], icon_path='ping.ico', duration=5, threaded=True)


while True:
    if first_start:
        excel_to_list()
    tab_ping()
    print_time_input(sec)
    os.system('cls')
