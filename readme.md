Скрипт на python для работы с Excel файлами
***
Работает с двумя таблицами, читает 5 столбцов из таблицы "Директ.xlsx", содержащих дату, номер объявления и
данные о конверсиях, а также столбцы "A" и "G" таблицы "Заявки.xlsx", содержащие дату и ID объявления. Затем он создает
словарь "dict_bboard_id_all_info", где ключ - это комбинация номера объявления и номера строки таблицы
"Заявки.xlsx", а значение - это дата, на которую было создано объявление.

Далее скрипт проходит по всем строкам таблицы "Директ.xlsx" и проверяет, есть ли для этой строки соответствующая запись
в словаре "dict_bboard_id_all_info". Если запись есть, скрипт добавляет данные в новую таблицу "result.xlsx" в
соответствующую строку и столбцы.

Таким образом, скрипт помогает сопоставить данные из двух разных таблиц на основе общих значений в столбцах "F" и "G",
а также даты из столбца "A".