"""
Скрипт для опроса сетевых устройств
"""
import sys

import openpyxl  # Библиотека для работы с таблицами
from ping3 import ping
import win10toast  # Библиотека для всплывающих сообщений
import time
import os
import msvcrt
from rich.console import Console
from rich.table import Table
from rich.progress import track
import datetime
from collections import deque


ip_book = openpyxl.load_workbook("ip cam.xlsx")  # Открывает файл
worksheet = ip_book.active  # Делаем его активным
max_row = worksheet.max_row  # Получаем максимальное количество строк
max_col = worksheet.max_column  # Получаем максимальное количество столбцов
#file_log = open('log.txt', 'w')
# Инициализация начальных данных
sec = 600  # По умолчанию перезагрузка через 600 сек
scan_ip = []  # Список всех устройств
off_ip = []  # Список отсутствующих

console = Console()

def table_generation(title_table, columns_table, type = 'Main'):
    if type == 'Main':
        table = Table(title=title_table)
    elif type == 'Grid':
        table = Table.grid(expand=True)
    for column in columns_table:
        table.add_column(column)
    return  table


def excel_to_list():
    """
    Функция заполняет список устройствами для опроса

    :return:
    """
    scan_ip = []
    for row in range(1, max_row):  # Запускаем цикл по всем строкам
        val = worksheet.cell(row=row, column=5).value
        if val == 'Камера' or val == 'Хранилище' or val == 'ПК' \
                or val == 'Видеорегистратор' or val == 'Wi-Fi' or val == 'Преобразователь':  # Если устройство камера
            ip_cam = worksheet.cell(row=row, column=1).value  # Адрес устройства
            ip_object = worksheet.cell(row=row, column=2).value  # Объект на котором оно установлена
            ip_type = worksheet.cell(row=row, column=5).value  # Тип устройства
            ip_comment = worksheet.cell(row=row, column=6).value  # Комментарий
            ip_priority = worksheet.cell(row=row, column=7).value  # Важность для отображения
            ip_active = worksheet.cell(row=row, column=8).value  # Работоспособность устройства
            if ip_priority != 'False':  # Если устройство не исключено из наблюдения
                scan_ip.append([row, ip_cam, ip_object, ip_type, ip_comment, ip_priority, ip_active])
    return scan_ip

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
    time_end = 0
    for i in range(len(off_ip)):
        if ip_cam == off_ip[i][2]:  # Если устройство присутствует в списке
            flag_add = False  # Помечаем, что добавлять его не надо
            if flag_ip:
                off_ip[i][0] = 2  # Если ответило на пинг, помечаем на удаление из списка
            else:
                off_ip[i][0] = 1  # Если оно снова не ответило на пинг отключаем уведомление
            time_end = off_ip[i][6]

    if flag_add and not flag_ip:  # Если устройства нет в списке и оно не отвечает - добавляем
        time_end = int(time.time())
        off_ip.append([0, row, ip_cam, ip_object, ip_type, ip_comment, time_end])
    return time_end


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


def tab_ping(scan_ip):
    """

    :return:
    """
    count_device = len(scan_ip)
    count_error = 0
    #  Создаем необходимые таблицы для визуального оформления
    columns_table = ('IP Адрес', 'Статус', 'Объект', 'Тип', 'Комментарий')
    table_priority = table_generation("Приоритетные устройства", columns_table)
    columns_table = ('IP Адрес', 'Статус', 'Объект', 'Тип', 'Комментарий', 'Время OFF-LINE')
    table_ip = table_generation("Устройства не в сети", columns_table)
    columns_table = ('IP Адрес', 'Объект', 'Тип', 'Комментарий')
    table_error = table_generation("Неисправные устройства", columns_table)
    columns_table = ("", 'justify="left"')
    grid_main = table_generation("", columns_table, 'Grid')
    grid_left = table_generation("", ("",), 'Grid')
    grid_left.add_row(table_priority)
    grid_left.add_row(table_error)
    # Создаем таблицу для вывода логов
    table_log = Table(title="История опросов")
    table_log.add_column('IP Адрес - Время')
    grid_main.add_row(grid_left, table_log)

    for row, ip_cam, ip_object, ip_type, ip_comment, ip_priority, ip_active in \
            track(scan_ip, description='[green]Ping'):  # Рисуем трак бар и пробегаем все устройства
        if ip_active == 'ON':  # Если устройство включено
            ip_pin = ping(ip_cam)
            flag_ip = True  # Пинг прошёл
            if ip_pin is None or type(ip_pin) is not float:
                flag_ip = False  # Пинг не прошёл
            # Изменяем список устройств не отвечающих на пинг
            time_end = modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_type, ip_comment)
            if ip_priority == 'True':  # Заполняем таблицу приоритетных устройств
                if flag_ip:
                    table_priority.add_row(f'[green]{ip_cam}[/green]', str(flag_ip), ip_object, ip_type, ip_comment)
                else:
                    table_priority.add_row(f'[red]{ip_cam}[/red]', str(flag_ip), ip_object, ip_type, ip_comment)
            if not flag_ip:  # Заполняем таблицу не ответивших устройств
                time_format = str(datetime.timedelta(seconds=int(time.time())-time_end))
                table_ip.add_row(f'[red]{ip_cam}[/red]', str(ip_pin), ip_object, ip_type, ip_comment, time_format)
        if ip_active == 'OFF':  # Если устройства отключены заполняем таблицу отключенных
            table_error.add_row(f'[red]{ip_cam}[/red]', ip_object, ip_type, ip_comment)
    #  Заполняем строки таблицами
    grid_left.add_row(f'Всего устройств: {str(count_device)} '
                  f'[green] На связи: {count_device-table_ip.row_count-table_error.row_count}[/green]'
                  f'[red] Отсутствуют: {table_ip.row_count}[/red]'
                  f'[blue] Отключены: {table_error.row_count}[/blue]')
    if table_ip.row_count > 0:  # Если таблица не пустая - выводим
        grid_left.add_row(table_ip)

    off_text = ''
    on_text = ''
    del_index = []
    file_log = open('log.txt', 'a')
    for flag, _, ip, obj, _type, com, tm in off_ip:
        if not flag:  # Если устройство впервые не ответило, выводим сообщение и пишем в лог
            off_text += f'{ip[10:]},'
            file_log.write(f'OFF >> {ip} - {time.ctime()}\n')
        if flag == 2:  # Если устройство ответило, выводим сообщение, пишем в лог, удаляем из списка
            on_text += f'{ip[10:]},'
            file_log.write(f'ONN >> {ip} - {time.ctime()}\n')
            del_index.append(off_ip.index([flag, _, ip, obj, _type, com, tm]))
    file_log.close()
    # Удаляем появившиеся на связи устройства
    del_index.reverse()
    for _ in del_index:
        off_ip.pop(_)
    # Выводим всплывающее собщение
    toaster = win10toast.ToastNotifier()
    if len(off_text) > 0:
        toaster.show_toast("Отсутствуют", off_text[:-1], icon_path='ping.ico', duration=5, threaded=True)
    if len(on_text) > 0:
        toaster.show_toast("Восстановление связи", on_text[:-1], icon_path='ping.ico', duration=5, threaded=True)
    # Заполняем таблицу событий
    file_log = open('log.txt', 'r')
    my_stack = deque(maxlen=15)  # Выводим последние 15
    for line_log in file_log:
        my_stack.append(line_log)
    file_log.close()
    for stack_line in my_stack:
        if stack_line[0:3] == 'ONN':
            table_log.add_row(f'[green]{stack_line}[/green]')
        if stack_line[0:3] == 'OFF':
            table_log.add_row(f'[red]{stack_line}[/red]')
    # Выводим таблицы на экран
    console.print(grid_main)


while True:
    if len(scan_ip) == 0:
        scan_ip = excel_to_list()
    tab_ping(scan_ip)
    print_time_input(sec)
    os.system('cls')
