"""
Пример кода для работы с БД СЭД "Дело".
"""
from contextlib import closing
from django.conf import settings
from .delo_utils import Delo


class DeloApi:
    # Устанавливаем связь с БД СЭД "Дело".

    def __init__(self):
        self.DELO_SERVER_NAME = settings.DELO_SERVER_NAME
        # Порт СУБД, где расположена БД «Дело»
        self.DELO_SERVER_PORT = settings.DELO_SERVER_PORT
        # Имя базы данных СЭД «Дело»
        self.DELO_DB_NAME = settings.DELO_DB_NAME
        # Имя пользователя БД «Дело» для подключения
        self.DELO_USERNAME = settings.DELO_USERNAME
        # Пароль пользователя БД «Дело» для подключения
        self.DELO_PASSWORD = settings.DELO_PASSWORD
        # Имя ODBC-драйвера БД
        self.DELO_DB_DRIVER = settings.DELO_DB_DRIVER
        # Размер буфера записей курсора БД СЭД "Дело"
        self.CURSOR_ARRAY_SIZE = 5000
        self.DELO_CONNECTION = ""

    def run(self, sql_string):
        with closing(Delo(self.DELO_SERVER_NAME, self.DELO_DB_NAME, self.DELO_USERNAME, self.DELO_PASSWORD,
                          self.DELO_SERVER_PORT, driver=self.DELO_DB_DRIVER)) as delo:
            if not delo.is_connected:
                # Сигнализируем неудачное завершение команды.
                raise RuntimeError(
                    "Couldn't connect to DELO DB. Wrong settings.DELO_* or a network problem?"
                )

            self.DELO_CONNECTION = delo.connection

            if sql_string:
                """ Функция выполняющая sql запрос и возвращающая результат"""
                with closing(self.DELO_CONNECTION.cursor()) as cur:
                    cur.execute(sql_string)
                    results = cur.fetchall()
                if results:
                    return results
            else:
                return None
