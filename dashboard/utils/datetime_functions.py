from datetime import datetime, timedelta
from typing import Dict
import calendar


def first_day_of_month(month_year: str = None) -> datetime:
    """
    Принимает месяц.год в формате "%m.%Y"
    Возвращает первый день указанного месяца
    """
    if month_year:
        month, year = map(int, month_year.split('.'))
        first_day = datetime(year, month, 1)
    else:
        today = datetime.today()
        first_day = datetime(today.year, today.month, 1)
    return first_day


def last_day_of_month(month_year: str = None) -> datetime:
    """
    Принимает месяц.год в формате "%m.%Y"
    Возвращает последний день указанного месяца
    """
    if month_year:
        month, year = map(int, month_year.split('.'))
        _, last_day = calendar.monthrange(year, month)
        last_day = datetime(year, month, last_day)
    else:
        today = datetime.today()
        _, last_day = calendar.monthrange(today.year, today.month)
        last_day = datetime(today.year, today.month, last_day)
    return last_day


def get_last_12_month_period() -> Dict[int, str]:
    """
    Возвращает словарь дат за последний год, где ключ - число от 1 до 12,
    а значение - строка в формате "mm.yyyy"
    Пример:
    {1: '03.2024', 2: '02.2024', ... , 12: '03.2023'}
    """
    current_date = datetime.now()
    past_12_month_year_dict = {1: current_date.strftime("%m.%Y")}
    for month_number in range(2, 13):
        # определяем предыдущий месяц
        prev_month_datetime = current_date - timedelta(days=current_date.day)

        # Преобразуем дату в строку в формате "месяц.год" (например, "01.2022")
        prev_month_year_string = prev_month_datetime.strftime("%m.%Y")

        past_12_month_year_dict[month_number] = prev_month_year_string
        # Уменьшаем месяц на 1
        current_date = prev_month_datetime - timedelta(days=1)

    return past_12_month_year_dict
