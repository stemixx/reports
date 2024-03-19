from deloreports.api.delo_start import DeloApi
from dashboard.models import Department, ReportName, Statistic
from dashboard.utils.datetime_functions import *
from collections import OrderedDict


class DeloSqlQuery(DeloApi):
    """
    Класс для работы с запросами в БД Дело.
    """

    def __init__(self):
        DeloApi.__init__(self)

    def get_resolutions_count_of_month(self, dep_due: str, start_date: str = None, end_date: str = None) -> int:
        """Запрос из БД Дело для отчёта:
        Количество документов, по которым введено хотя бы одно поручение.
        :param dep_due: str. Код подразделения из БД Дело.
        :param start_date: str. Месяц.год в формате "%m.%Y" (01.2024)
        :param end_date: str. Месяц.год в формате "%m.%Y" (01.2024)
        :return: int. Количество поручений по одному подразделению.
        Здесь и далее pyodbc значения даты передаёт в БД в виде '2024-03-01'
        В среде Microsoft SQL Server Management Studio запросы даты вида '2024-03-01' не обрабатываются из-за
        ошибки конвертации строкового значения в дату, поэтому там передаём в формате '01/03/2024'
        """
        if start_date and end_date:
            start_date_ymd = first_day_of_month(start_date).strftime('%Y-%m-%d')
            end_date_ymd = last_day_of_month(end_date).strftime('%Y-%m-%d')
        else:
            start_date_ymd = first_day_of_month().strftime('%Y-%m-%d')
            end_date_ymd = last_day_of_month().strftime('%Y-%m-%d')
        sql_string = f'''
            SELECT COUNT (DISTINCT doc.FREE_NUM)
                    FROM delo_db.dbo.RESOLUTION AS RES
                    LEFT JOIN delo_db.dbo.DEPARTMENT AS DEPART ON RES.DUE = DEPART.DUE
                    LEFT JOIN delo_db.dbo.REPLY AS REPL ON RES.ISN_RESOLUTION = REPL.ISN_RESOLUTION
                    LEFT JOIN delo_db.dbo.DEPARTMENT AS DEP_ISPOLNITEL ON DEP_ISPOLNITEL.DUE = REPL.DUE
                    LEFT JOIN delo_db.dbo.DOC_RC AS DOC ON DOC.ISN_DOC = RES.ISN_REF_DOC
                    WHERE 
                        DOC.DOC_DATE BETWEEN '{start_date_ymd}' AND '{end_date_ymd}'
                        AND NOT RES.SEND_DATE IS NULL
                        AND REPL.MAIN_FLAG = 1 
                        AND DEP_ISPOLNITEL.DUE LIKE '{dep_due}%'
        '''

        return self.run(sql_string)[0][0]

    def get_count_of_opened_closed_and_expired_docs(
            self,
            dep_due: str,
            start_date: str = None,
            end_date: str = None
    ) -> tuple[int, int, int]:
        """
        Функция возвращает на sql-запрос количество открытых, закрытых и просроченных документов (по отчёту).
        """
        if start_date and end_date:
            start_date_ymd = first_day_of_month(start_date).strftime('%Y-%m-%d')
            end_date_ymd = last_day_of_month(end_date).strftime('%Y-%m-%d')
        else:
            start_date_ymd = first_day_of_month().strftime('%Y-%m-%d')
            end_date_ymd = last_day_of_month().strftime('%Y-%m-%d')
        sql_string = f'''
            SELECT COUNT(DISTINCT DOC.FREE_NUM)
            FROM delo_db.dbo.DOC_RC AS DOC
            JOIN delo_db.dbo.RESOLUTION as RES on DOC.ISN_DOC = RES.ISN_REF_DOC
            JOIN delo_db.dbo.REPLY as REP on RES.ISN_RESOLUTION = REP.ISN_RESOLUTION
            WHERE 
                RES.PLAN_DATE BETWEEN '{start_date_ymd}' AND '{end_date_ymd}'
                AND DUE_CONTROLLER LIKE '{dep_due}%'
                AND DOC.ISN_DOC NOT IN (
                    SELECT DOC.isn_doc
                    FROM delo_db.dbo.DOC_RC AS DOC
                    JOIN delo_db.dbo.RESOLUTION as RES on DOC.ISN_DOC = RES.ISN_REF_DOC
                    JOIN delo_db.dbo.REPLY as REP on RES.ISN_RESOLUTION = REP.ISN_RESOLUTION
                    WHERE REP.REPLY_DATE IS NOT NULL
                )
             
            UNION ALL   
        
            SELECT COUNT (DISTINCT doc.FREE_NUM)
            FROM delo_db.dbo.DOC_RC AS DOC
            JOIN delo_db.dbo.RESOLUTION AS RES ON doc.ISN_DOC = RES.ISN_REF_DOC
            JOIN delo_db.dbo.REPLY as REP on RES.ISN_RESOLUTION = REP.ISN_RESOLUTION
            WHERE 
                RES.PLAN_DATE BETWEEN '{start_date_ymd}' AND '{end_date_ymd}'
                AND CONVERT(date, REP.REPLY_DATE) <= REP.PLAN_DATE
                AND DUE_CONTROLLER LIKE '{dep_due}%'
                
            UNION ALL
            
            SELECT COUNT (DISTINCT doc.FREE_NUM)
            FROM delo_db.dbo.DOC_RC AS DOC
            JOIN delo_db.dbo.RESOLUTION AS RES ON doc.ISN_DOC = RES.ISN_REF_DOC
            JOIN delo_db.dbo.REPLY as REP on RES.ISN_RESOLUTION = REP.ISN_RESOLUTION
            WHERE 
                RES.PLAN_DATE BETWEEN '{start_date_ymd}' AND '{end_date_ymd}'
                AND CONVERT(date, REP.REPLY_DATE) > REP.PLAN_DATE
                AND DUE_CONTROLLER LIKE '{dep_due}%'
        '''
        opened_closed_and_expired_docs = self.run(sql_string)
        opened_docs, closed_docs, expired_docs = opened_closed_and_expired_docs[0][0], \
                                                 opened_closed_and_expired_docs[1][0], \
                                                 opened_closed_and_expired_docs[2][0]
        return opened_docs, closed_docs, expired_docs


class DjangoSqlQuery:
    """
    Запросы в БД Postgres текущего проекта
    """

    def __init__(self):
        self.delo_query = DeloSqlQuery()
        self.deps_dues = self.get_deps()

    @classmethod
    def get_deps(cls):
        deps = Department.objects.only('short_name', 'due')
        deps_dues = {department.short_name: department.due for department in deps}
        return deps_dues

    def set_and_update_resolutions_count_of_month(self, start_month_year: str = None, end_month_year: str = None):
        """
        Функция для менеджмент-команды django.
        Записывает в БД данные о количестве резолюций за текущий месяц, если такие данные ещё не созданы
        и обновляет столбцы amount, update_date, если записи за текущий месяц уже существуют.
        :param start_month_year: str. Месяц.год в формате "%m.%Y" (01.2024)
        :param end_month_year: str. Месяц.год в формате "%m.%Y" (01.2024)
        :return: get_or_create экземпляра Statistic. Создаёт или обновляет
        строки в БД таблицы Statistic с report_name_id = 1.
        """

        for dep, due in self.deps_dues.items():
            if start_month_year and end_month_year:
                start_date_dt = first_day_of_month(start_month_year)
                end_date_dt = last_day_of_month(end_month_year)
                amount = self.delo_query.get_resolutions_count_of_month(due, start_month_year, end_month_year)
            else:
                start_date_dt = first_day_of_month()
                end_date_dt = last_day_of_month()
                amount = self.delo_query.get_resolutions_count_of_month(due)

            obj, created = Statistic.objects.get_or_create(
                report_name=ReportName.objects.get(name='resolutions_count_of_current_month'),
                department=Department.objects.get(due=due),
                start_date=start_date_dt,
                end_date=end_date_dt,
                defaults={
                    'amount': amount,
                    'update_date': datetime.now()
                },
            )

            if not created:
                obj.amount = amount
                obj.update_date = datetime.now()
                obj.save(update_fields=['amount', 'update_date'])

    @staticmethod
    def get_resolutions_count_per_month(n: int) -> tuple[dict[str, int], datetime]:
        """
        Данные по резолюциям из промежуточной таблицы django для более быстрого отображения и снижения
        нагрузки на сервер БД Дело.
        :param n: int. Число от 1 до 12 согласно функции update_graph(n), ссылающееся на значение "месяц.год" .
        :return: tuple[dict[str, int], datetime].
        Кортеж из:
        Словарь. Ключ - подразделение, значение - количество документов, по которым введена хотя бы одна резолюция.
        Объект datetime. Дата обновления сведений
        """
        try:
            month_year = get_last_12_month_period()[n]
            start_date = first_day_of_month(month_year)
            statistics = Statistic.objects.filter(report_name=1, start_date=start_date).select_related('department')
            # ниже django ORM или особенность драйвера Postgre? При ORM запросе поля timestamptz возвращаются в UTC
            # добавляем +3 часа для нашего часового пояса
            update_date_utc = statistics.values().first()['update_date']
            utc_offset = timedelta(hours=3)
            update_date_utc3 = update_date_utc + utc_offset
            resolutions_count_data = {}
            for stat in statistics:
                resolutions_count_data[stat.department.short_name] = stat.amount

        except Exception:
            resolutions_count_data = {}
            update_date_utc3 = ''

        sorted_resolutions_summary = OrderedDict(sorted(resolutions_count_data.items()))
        return sorted_resolutions_summary, update_date_utc3

    def set_and_update_count_of_open_and_closed_and_expired_docs(self,
                                                                 start_month_year: str = None,
                                                                 end_month_year: str = None):
        """
        Записывает в БД данные о количестве всех, закрытых и истекших документов, если такие данные ещё не созданы
        и обновляет столбцы closed, expired, update_date, если записи за текущий месяц уже существуют.
        :param start_month_year: str. Месяц.год в формате "%m.%Y" (01.2024) плановой даты поручения.
        :param end_month_year: str. Месяц.год в формате "%m.%Y" (01.2024)плановой даты поручения.
        :return: get_or_create экземпляра Statistic. Создаёт или обновляет
        строки в БД таблицы Statistic с report_name_id = 2.
        """
        for dep, due in self.deps_dues.items():
            if start_month_year and end_month_year:
                start_date_dt = first_day_of_month(start_month_year)
                end_date_dt = last_day_of_month(end_month_year)
                opened, closed, expired = self.delo_query.get_count_of_opened_closed_and_expired_docs(
                    due, start_month_year, end_month_year
                )
            else:
                start_date_dt = first_day_of_month()
                end_date_dt = last_day_of_month()
                opened, closed, expired = self.delo_query.get_count_of_opened_closed_and_expired_docs(due)

            obj, created = Statistic.objects.get_or_create(
                report_name=ReportName.objects.get(name='count_of_closed_and_expired_docs'),
                department=Department.objects.get(due=due),
                start_date=start_date_dt,
                end_date=end_date_dt,
                defaults={
                    'opened': opened,
                    'closed': closed,
                    'expired': expired,
                    'update_date': datetime.now()
                },
            )

            if not created:
                obj.amount = opened
                obj.closed = closed
                obj.expired = expired
                obj.update_date = datetime.now()
                obj.save(update_fields=['opened', 'closed', 'expired', 'update_date'])

    @staticmethod
    def get_count_of_all_and_closed_and_expired_docs(n: int) -> tuple[dict[str, int, int, int], datetime]:
        """
        Данные по количеству отработанных поручений из промежуточной таблицы django для более быстрого отображения и
        снижения нагрузки на сервер БД Дело.
        :param n: int. Число от 1 до 12 согласно функции update_graph(n), ссылающееся на значение "месяц.год" .
        :return: tuple[dict[str, int], datetime].
        Кортеж из:
        Словарь. Ключ - подразделение, значение - количество документов, по которым введена хотя бы одна резолюция.
        Объект datetime. Дата обновления сведений
        """
        try:
            month_year = get_last_12_month_period()[n]
            start_date = first_day_of_month(month_year)
            statistics = Statistic.objects.filter(report_name=2, start_date=start_date).select_related('department')
            # ниже django ORM или особенность драйвера Postgre? При ORM запросе поля timestamptz возвращаются в UTC
            # добавляем +3 часа для нашего часового пояса
            update_date_utc = statistics.values().first()['update_date']
            utc_offset = timedelta(hours=3)
            update_date_utc3 = update_date_utc + utc_offset
            opened_and_closed_and_expired_docs = {}
            for stat in statistics:
                opened_and_closed_and_expired_docs[stat.department.short_name] = stat.opened, stat.closed, stat.expired

        except Exception:
            opened_and_closed_and_expired_docs = {}
            update_date_utc3 = ''

        sorted_all_and_closed_and_expired_docs = OrderedDict(sorted(opened_and_closed_and_expired_docs.items()))
        return sorted_all_and_closed_and_expired_docs, update_date_utc3
