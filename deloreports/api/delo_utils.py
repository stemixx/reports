
'''
Утилиты приложения delo.
'''

import pyodbc


pyodbc.pooling = True


class Delo:
    '''
    База данных СЭД "Дело".
    '''
    def __init__(self, server_name, db_name, username, password, port=1433,
                 autocommit=False, readonly="readonly", driver="SQL Server"):
        '''
        Инициализирует экземпляр класса Delo.
        '''
        self.server_name = server_name
        self.port = port or 1433
        self.db_name = db_name
        self.username = username
        self.password = password
        self.autocommit = autocommit
        self.readonly = readonly
        self.driver = driver

        self._connection = None

    @property
    def connection_string(self):
        """
        Возвращает строку соединения ODBC для базы данных СЭД "Дело".
        """
        return ''.join([
            'DRIVER={%s};APP=ud;ClientCharset=UTF-8;' % self.driver,
            'SERVER=%s;' % self.server_name,
            'DATABASE=%s;' % self.db_name,
            'UID=%s;' % self.username,
            'PWD=%s;' % self.password,
            'Port=%s;' % self.port
        ])

    @property
    def connection(self):
        """
        Возвращает соединение с базой данных СЭД "Дело".
        """
        if not self._connection:
            try:
                # Устанавливаем соединение.
                self._connection = pyodbc.connect(
                    self.connection_string,
                    autocommit=self.autocommit,
                    readonly=self.readonly,
                    unicode_results=True
                )

            except pyodbc.Error:
                pass

        return self._connection

    @property
    def is_connected(self):
        """
        Возвращает True, если соединение с БД СЭД "Дело" успешно установлено,
        и False в остальных случаях.
        """
        return self.connection is not None

    def close(self):
        """
        Освобождает занятые объектом ресурсы, в т.ч. закрывает соединение с БД.
        """
        if self.is_connected:
            self.connection.close()
            self._connection = None

    def get_db_username(self, username):
        """
        Возвращает имя пользователя *базы данных* "Дело", соответствующее
        указанному пользователю СЭД "Дело".
        """
        if self.is_connected:
            cur = self.connection.cursor()
            cur.execute('''
                SELECT
                    [dbo].[USER_CL].[ORACLE_ID]
                FROM
                    [dbo].[USER_CL]
                WHERE
                    [dbo].[USER_CL].[CLASSIF_NAME] LIKE UPPER(?)
            ''', username)
            result = cur.fetchone()
            cur.close()

            if result:
                return result[0]

        return None

    def is_valid_delo_user(self, username, password):
        """
        Возвращает True, если пользователь с указанными именем, паролем имеется
        в СЭД "Дело", и False в остальных случаях.
        """
        if self.is_connected:
            cur = self.connection.cursor()
            cur.execute('''
                SELECT
                    COUNT(*)
                FROM
                    [master].[dbo].[sysxlogins] L
                    INNER JOIN [dbo].[USER_CL] U ON L.[name] = U.[ORACLE_ID]
                WHERE
                    U.[CLASSIF_NAME] LIKE UPPER(?)
                    AND PWDCOMPARE(?, L.[password]) = 1
            ''', username, password)
            result = cur.fetchone()
            cur.close()

            if result:
                return result[0] == 1

        return False


