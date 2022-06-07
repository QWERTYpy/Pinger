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
from rich.progress import Progress
from datetime import datetime


def print_time_input(timeout):
    """
    Данная процедура выводит обратный таймер и следит за вводом для управления программой

    :param timeout: Значение таймера

    :return:
    """
    start = time.monotonic()
    now = time.monotonic()
    count_sec = 0
    global hidden_ping
    while now - start < timeout:
        if now - start > count_sec:
            count_sec += 1
            print_time = timeout-count_sec
            delta = str(datetime.now()-start_programs)[:-7]
            hping = 'ON' if hidden_ping else 'OFF'
            print("\r%s Осталось %03d c. | h - hping %s | u-обновить | t-таймер | q-выход :" % (delta, print_time, hping), end='')
        if msvcrt.kbhit():
            c = msvcrt.getwch()
            if c == 'q' or c == 'й':
                lists.reboot_ping = True
                sys.exit()
            if c == 'h' or c == 'р':
                hidden_ping = not hidden_ping
            if c == 'u' or c == 'г':
                break
            if c == 't' or c == 'е':
                dict_const['time_sec'] = int(input('Введите количество секунд :'))
                break
        if lists.reboot_ping:
            break
        now = time.monotonic()


def progress_bar():
    with Progress() as progress:
        task_ping = progress.add_task("[red]Опрос устройств...", total=lists.progressbar)
        while not progress.finished:
            progress.update(task_ping, completed=lists.progressbar_complit)


lists = ListsIP()
console = Console()
hidden_ping = True
start_programs = datetime.now()
while True:
    tables = TableGeneration()
    # Создаем пул на определенной в файле настроек количесвто потоков
    lists.progressbar_complit = 0

    threading.Thread(target=progress_bar).start()
    with concurrent.futures.ThreadPoolExecutor(max_workers=dict_const['ping_device']) as executor:
        executor.map(lists.thread_ping, lists.scan_ip)

    lists.progressbar_complit = lists.progressbar # После сборки в exe наблюдается глюк. Иногда. Это костыль для завершения Опроса
        # while not progress.finished:
        #     progress.update(task_ping, completed=lists.progressbar_complit)
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

    if len(lists.off_ip) > 0 and hidden_ping:
        threading.Thread(target=lists.hidden_ping, args=(lists.off_ip,)).start()
    print_time_input(dict_const['time_sec'])
    os.system('cls')
