from rich.table import Table
from init_const import dict_const
import datetime
import time
from collections import deque


class TableGeneration:
    def __init__(self):
        #  Создаем необходимые таблицы для визуального оформления
        # Таблица отображения приоритетных устройств
        self.table_priority = self.table_generation(dict_const['table_priority_name'], dict_const['table_priority_field'])
        # Таблица отображения отсутствующих устройств
        self.table_ip = self.table_generation(dict_const['table_ip_name'], dict_const['table_ip_field'])
        # Таблица отображения отключенных устройств
        self.table_error = self.table_generation(dict_const['table_error_name'], dict_const['table_error_field'])
        # Создаем таблицу для вывода логов
        self.table_log = self.table_generation(dict_const['table_history_name'], dict_const['table_history_field'])

        # Невидимая таблица для оформления
        self.grid_main = self.table_generation("", dict_const['console'][2:4], 'Grid')
        self.grid_left = self.table_generation("", ("",), 'Grid')
        self.grid_left.add_row(self.table_priority)
        self.grid_left.add_row(self.table_error)
        self.grid_main.add_row(self.grid_left, self.table_log)

    @staticmethod
    def table_generation(title_table, columns_table, _type='Main'):
        """ Функция создает таблицы для форматирования вывода консоли

        :param title_table: Заголовок таблицы
        :param columns_table: Список заголовков столбцов
        :param _type: Видимая таблица или для зонирования
        :return:
        Указатель на созданную таблицу
        """
        if _type == 'Main':
            table = Table(title=f'{title_table}')
        elif _type == 'Grid':
            table = Table.grid(expand=True)
        for column in columns_table:
            if isinstance(column, str):
                table.add_column(column)
            if isinstance(column, list):
                table.add_column(column[0], width=column[1])
        return table

    def table_filling(self, scan_ip, modification_off_ip):
        count_device = len(scan_ip)
        for flag_ip, row, ip_cam, ip_object, ip_type, ip_comment, ip_priority, ip_active in scan_ip:
            if ip_active == 'ON':  # Если устройство включено
                # Изменяем список устройств не отвечающих на пинг
                time_end = modification_off_ip(flag_ip, row, ip_cam, ip_object, ip_type, ip_comment)
                if ip_priority == dict_const['important']:  # Заполняем таблицу приоритетных устройств
                    if flag_ip:
                        self.table_priority.add_row(f'[green]{ip_cam}[/green]', str(True if flag_ip else False), ip_object,
                                               ip_type, ip_comment)
                    else:
                        self.table_priority.add_row(f'[red]{ip_cam}[/red]', str(True if flag_ip else False), ip_object,
                                               ip_type, ip_comment)
                if not flag_ip:  # Заполняем таблицу не ответивших устройств
                    time_format = str(datetime.timedelta(seconds=int(time.time()) - time_end))
                    self.table_ip.add_row(f'[red]{ip_cam}[/red]', str(True if flag_ip else False), ip_object, ip_type,
                                     ip_comment, time_format)
            if ip_active == 'OFF':  # Если устройства отключены заполняем таблицу отключенных
                self.table_error.add_row(f'[red]{ip_cam}[/red]', ip_object, ip_type, ip_comment)
        #  Заполняем строки таблицами
        self.grid_left.add_row(f'Всего устройств: {str(count_device)} '
                          f'[green] На связи: {count_device - self.table_ip.row_count - self.table_error.row_count}[/green]'
                          f'[red] Отсутствуют: {self.table_ip.row_count}[/red]'
                          f'[blue] Отключены: {self.table_error.row_count}[/blue]')
        if self.table_ip.row_count > 0:  # Если таблица не пустая - выводим
            self.grid_left.add_row(self.table_ip)

    def table_logs_filling(self):
        # Заполняем таблицу событий
        file_log = open(dict_const['log_file'], 'r')
        my_stack = deque(maxlen=dict_const['buffer'])  # Выводим согласно размера буфера
        for line_log in file_log:
            my_stack.append(line_log[:-1])
        file_log.close()
        for stack_line in my_stack:
            if stack_line[0:3] == 'ONN':
                self.table_log.add_row(f'[green]{stack_line}[/green]')
            if stack_line[0:3] == 'OFF':
                self.table_log.add_row(f'[red]{stack_line}[/red]')
