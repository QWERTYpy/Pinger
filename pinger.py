"""
Скрипт для опроса сетевых устройств
"""
import sys
import win10toast  # Библиотека для всплывающих сообщений
import time
import os
import msvcrt
from rich.console import Console
from init_const import dict_const
import threading
import concurrent.futures
from list import ListsIP
from table import TableGeneration


def print_time_input(timeout):
    """
    Данная процедура выводит обратный таймер и следит за вводом для управления программой

    :param timeout: Значение таймера

    :return:
    """
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
            if c == 'q' or c == 'й':
                lists.reboot_ping = True
                sys.exit()
            if c == 'u' or c == 'г':
                break
            if c == 't' or c == 'е':
                dict_const['time_sec'] = int(input('Введите количество секунд :'))
                break
        if lists.reboot_ping:
            break
        now = time.monotonic()


lists = ListsIP()
console = Console()

while True:
    tables = TableGeneration()
    # Создаем пул на определенной в файле настроек количесвто потоков
    with concurrent.futures.ThreadPoolExecutor(max_workers=dict_const['ping_device']) as executor:
        executor.map(lists.thread_ping, lists.scan_ip)

    tables.table_filling(lists.scan_ip, lists.modification_off_ip)
    on_text, off_text = lists.message()
    toaster = win10toast.ToastNotifier()
    if len(off_text) > 0:
        toaster.show_toast("Отсутствуют", off_text[:-1], icon_path='ping.ico', duration=5, threaded=True)
    if len(on_text) > 0:
        toaster.show_toast("Восстановление связи", on_text[:-1], icon_path='ping.ico', duration=5, threaded=True)
    tables.table_logs_filling()
    # Выводим таблицы на экран
    os.system(f"mode con:cols={dict_const['console'][0]} lines={dict_const['console'][1]}")
    console.print(tables.grid_main)

    if len(lists.off_ip) > 0:
        threading.Thread(target=lists.hidden_ping, args=(lists.off_ip,)).start()
    print_time_input(dict_const['time_sec'])
    os.system('cls')
