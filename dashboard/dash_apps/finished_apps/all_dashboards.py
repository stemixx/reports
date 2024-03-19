from dash import dcc, html
from dash.dependencies import Input, Output
from django_plotly_dash import DjangoDash
from datetime import datetime
from plotly.graph_objs import Figure

from dashboard.sql.db_queries import DeloSqlQuery, DjangoSqlQuery
from dashboard.utils.datetime_functions import get_last_12_month_period

# создаём экземпляры классов для sql-запросов в БД Дело и БД приложения.
delo_query = DeloSqlQuery()
django_query = DjangoSqlQuery()

colors = {
    'background': '#E6F1FF',
    'text': '#7FDBFF',
    'bar': '#1e90a4',
    'opened': dict(color='rgb(50, 146, 149)'),
    'closed': dict(color='rgb(35, 134, 41)'),
    'expired': dict(color='rgb(164, 39, 35)'),
    'line': dict(color='black', width=1),
}


def set_title_text(title: str, date_time: datetime or '') -> str:
    """
    Функция возвращает текст заголовка графика.
    Принимает на вход текст заголовка и дату/время.
    Возвращает строку с заголовком.
    """
    if date_time:
        upd_date = date_time.strftime("%d.%m.%Y %H:%M")
        text = f'<b>{title} по состоянию на {upd_date}</b>'
    else:
        text = ''

    return text


# Cоздаём приложение-график plotly для django-сервера:
# Количество документов, по которым введено хотя бы одно поручение.
# Назначаем параметры слайдера для последующего обновления графика.
month_resolutions_app = DjangoDash('MonthResolutions')
month_resolutions_app.layout = html.Div([
    dcc.Graph(id='bar-chart-month-resolutions'),
    dcc.Slider(
        id='month-resolutions-slider',
        min=1,
        max=12,
        step=1,

        # marks принимает словарь, где ключ - целые числа (шаги),
        # а значения - строки с отображаемыми на слайдере датами.
        marks=get_last_12_month_period(),
        value=1,  # Значение по умолчанию. 1 - текущий месяц из функции get_last_12_month_period().
    ),
])


# Декоратор, определяющий из каких данных (здесь - слайдер c id=month-resolutions-slider)
# будет обновляться график, указанный в Output.
@month_resolutions_app.callback(
    Output('bar-chart-month-resolutions', 'figure'),
    Input('month-resolutions-slider', 'value')
)
def update_graph_month_resolutions(step_of_marks: int) -> dict:
    """
    Функция отображения данных при смещении ползунка слайдера.
    Принимает на вход число - максимальное количество шагов, определённое в объекте Slider
    """
    # Получаем обновленные данные по документам, по которым введено хотя бы одно поручение и дату обновления.
    resolutions_summary, update_date = django_query.get_resolutions_count_per_month(step_of_marks)
    title_text = set_title_text('Количество документов, по которым введено хотя бы одно поручение', update_date)

    figure = {
        'data': [
            {
                'x': list(resolutions_summary.keys()),  # названия подразделений (сокращенные) по оси X.
                'y': list(resolutions_summary.values()),  # количество документов по оси Y.
                'text': list(resolutions_summary.values()),  # отображение количества документов над барами.
                'textposition': 'outside',
                'type': 'bar',
                'marker': {'color': colors['bar'], 'opacity': 0.7, 'line': {'color': 'black', 'width': 1}}
            }
        ],
        'layout': {
            'title': {
                'text': title_text,
                'font': {'size': 16}
            },
            'xaxis': {'title': 'Орган Администрации'},
            'yaxis': {'title': 'Количество документов'},
            'height': 600,
            'plot_bgcolor': colors['background'],
        }
    }
    return figure


# создаём следующее приложение-график plotly 'Статистика по отчётам на поручения с плановой датой в текущем месяце'.
expired_docs_app = DjangoDash('ExpiredDocs')


def get_fig(step_of_marks: int = 1) -> "Figure":
    """
    Получение и отрисовка данных по отчётам на поручения, плановая дата - текущий месяц.
    Принимает на вход число, идентифицирующее месяц, по которому будут получены данные. По умолчанию 1 - текущий месяц.
    """
    import plotly.graph_objs as go
    # Получаем обновленные данные по количеству документов, по которым:
    # 1. Нет отчёта
    # 2. Введён отчёт
    # 3. Просрочено
    opened_and_closed_and_expired_docs_summary, update_date = django_query.get_count_of_all_and_closed_and_expired_docs(
        step_of_marks)
    # сокращенное название подразделений
    dep_short = list(opened_and_closed_and_expired_docs_summary.keys())

    # разбиваем кортеж с количеством документов по статусу поручений для каждого подразделения
    opened = [val[0] for val in opened_and_closed_and_expired_docs_summary.values()]
    closed = [val[1] for val in opened_and_closed_and_expired_docs_summary.values()]
    expired = [val[2] for val in opened_and_closed_and_expired_docs_summary.values()]

    # суммарно по каждому подразделению
    total_docs_values = [sum(x) for x in zip(opened, closed, expired)]

    # каждый объект go.Bar() - один слой данных на все бары. Происходит стекирование снизу вверх
    fig = go.Figure(data=[
        go.Bar(name='Нет отчёта', x=dep_short, y=opened, text=opened,
               textposition='inside', marker=colors['opened'], opacity=0.7),

        go.Bar(name='Введён отчёт', x=dep_short, y=closed, text=closed,
               textposition='inside', marker=colors['closed'], opacity=0.7),

        go.Bar(name='Просрочено', x=dep_short, y=expired, text=expired,
               textposition='inside', marker=colors['expired'], opacity=0.7),

        go.Scatter(name='Всего', x=dep_short, y=total_docs_values, mode='text', text=total_docs_values,
                   textposition='top center', showlegend=False)
    ])

    fig.update_layout(barmode='stack',
                      title=set_title_text(
                          'Статистика по отчётам на поручения с плановой датой в текущем месяце', update_date
                      ),
                      # заголовок графика по центру
                      title_x=0.5,
                      xaxis_title='Орган Администрации',
                      yaxis_title='Количество документов',
                      height=800,
                      plot_bgcolor=colors['background'],
                      )

    return fig


# определяем размещение и параметры слайдера для графика
expired_docs_app.layout = html.Div([
    dcc.Graph(id='bar-chart-month-slider-expired-reports', figure=get_fig()),
    dcc.Slider(
        id='month-slider-expired-reports',
        min=1,
        max=12,
        step=1,

        # marks принимает словарь, где ключ - целые числа (шаги),
        # а значения - строки с отображаемыми на слайдере датами.
        marks=get_last_12_month_period(),
        value=1,  # Значение по умолчанию. 1 - текущий месяц из функции get_last_12_month_period()
    ),
])


# Обновляем график с id=bar-chart-month-slider-expired-reports при смещении ползунка слайдера
@expired_docs_app.callback(
    Output('bar-chart-month-slider-expired-reports', 'figure'),
    Input('month-slider-expired-reports', 'value')
)
def update_graph_month_expired_reports(step_of_marks: int = 1) -> "Figure":
    figure = get_fig(step_of_marks)

    return figure
