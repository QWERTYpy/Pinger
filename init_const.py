import configparser

# Инициализация основных переменных
config = configparser.ConfigParser()
config.read("ping_config.ini", encoding="utf-8")
dict_const = {}
# Количество сек для ping
dict_const['time_sec'] = int(config["ping_time"]["time_sec"])
dict_const['ping_device'] = int(config["ping_time"]["ping_device"])
# Имя файла
dict_const['name_file'] = str(config["file_name"]["name"])
# Столбец из которого берется IP
dict_const['ip_camera'] = int(config["Columns_table"]["ip_camera"])
# Столбец из которого берется наименование Объекта
dict_const['ip_object'] = int(config["Columns_table"]["ip_object"])
# Столбец из которого берется тип устройства
dict_const['ip_type'] = int(config["Columns_table"]["ip_type"])
# Столбец из которого берется комментарий
dict_const['ip_comment'] = int(config["Columns_table"]["ip_comment"])
# Столбец из которого берется приоритетность отображения
dict_const['ip_priority'] = int(config["Columns_table"]["ip_priority"])
# Столбец из которого берется работоспособность
dict_const['ip_active'] = int(config["Columns_table"]["ip_active"])
# Важные данные
dict_const['important'] = str(config["ip_priority"]["important"])
# Неважные данные
dict_const['not_important'] = str(config["ip_priority"]["not_important"])
# Типы устройств для отображения
dict_const['type_name'] = [_.strip() for _ in str(config["ip_type"]["type_name"]).split(',')]
# Настройка консоли
dict_const['console'] = [int(config["console"]["cols"]), int(config["console"]["lines"]),
                         int(config["console"]["left_grid"]), int(config["console"]["right_grid"])]
# Количество строк в истории
dict_const['buffer'] = int(str(config["history"]["buffer"]))
dict_const['log_file'] = str(config["history"]["log_file"])
# Таблица приоритетных устройств
dict_const['table_priority_name'] = str(config["table"]["table_priority_name"])
dict_const['table_priority_field'] = [_.strip() for _ in str(config["table"]["table_priority_field"]).split(',')]
# Таблица отсутсвующих устройств
dict_const['table_ip_name'] = str(config["table"]["table_ip_name"])
dict_const['table_ip_field'] = [_.strip() for _ in str(config["table"]["table_ip_field"]).split(',')]
# Таблица отключенных устройств
dict_const['table_error_name'] = str(config["table"]["table_error_name"])
dict_const['table_error_field'] = [_.strip() for _ in str(config["table"]["table_error_field"]).split(',')]

