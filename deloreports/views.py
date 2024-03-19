import os
import os.path
import json
from io import BytesIO
from collections import Counter
import tempfile
import xlsxwriter
from docxtpl import DocxTemplate
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import permission_required
from django.core.cache import cache
from django.http import HttpResponse, HttpResponseBadRequest, StreamingHttpResponse, JsonResponse, \
    HttpResponseNotAllowed
from django.shortcuts import render, redirect
from deloreports.forms import (DateRangeForm, DateFormWithTOSChoices, LoginForm, DateFormSimpleDep,
                               DateFormSimpleOneDep, PaperFlowDateForm
                               )
from deloreports.functions.control_cases_report_api import *
from deloreports.functions.correspondents_foiv_report_api import *
from deloreports.functions.correspondents_report_api import *
from deloreports.functions.court_documents_report_api import *
from deloreports.functions.district_themes_report_api import *
from deloreports.functions.doc_flow_report_api import *
from deloreports.functions.effectiveness_report_api import *
from deloreports.functions.effectiveness_report_api_betatest import *
from deloreports.functions.gov_app_eight_report_api import *
from deloreports.functions.municipal_legal_act_registration_report_api import *
from deloreports.functions.not_assigned_court_documents_report_api import *
from deloreports.functions.paper_flow_report_api import *
from deloreports.functions.prosecutors_incoming_docs_report_api import *
from deloreports.functions.prosecutors_reaction_act_report_api import *
from deloreports.functions.reports_master_report_api import *
from deloreports.functions.municipal_legal_acts_consideration_report import *
from deloreports.functions.appeals_two_week_period import *
from deloreports.functions.check_documents import *
from deloreports.functions.editing_resolutions_by_mayor_report_api import *
from deloreports.functions.citizens_appeals_json import *
from deloreports.functions.utils import (
    get_department_by_due,
    get_padeg_dep
)


def index(request):
    login_form = LoginForm()
    return render(request, 'deloreports/index.html', {'form': login_form})


def auth_login(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(request, username=username, password=password)
        if user is not None:
            if user.is_active:
                login(request, user)
                return redirect('/deloreports/')

    login_form = LoginForm()
    error_message = "Неверно введены имя пользователя и/или пароль. Попробуйте снова."
    return render(request, 'deloreports/index.html', {'form': login_form, 'error': error_message})


def auth_logout(request):
    logout(request)
    return redirect('/deloreports/')


def get_master_report_cache_key(user_id, from_date, due_date, dep_list):
    return f"master_report-" \
           f"{user_id}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}-" \
           f"{''.join(dep_list)}"


def get_effectiveness_report_cache_key(user_id, from_date, due_date):
    return f"effectiveness_report-" \
           f"{user_id}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}"


def get_effectiveness_report_betatest_cache_key(user_id, from_date, due_date):
    return f"effectiveness_report_betatest-" \
           f"{user_id}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}"


def get_check_documents_cache_key(user_id, from_date, due_date, dep_name):
    return f"check_documents-" \
           f"{user_id}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}-" \
           f"{dep_name}"


def get_cache_key(report_name, user_id, from_date, due_date, *args):
    return f"{report_name}-" \
           f"{user_id}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}-" \
           f"{args}"


def get_cache_key2(report_name, from_date, due_date):
    return f"{report_name}-" \
           f"{from_date.strftime('%d.%m.%Y')}-" \
           f"{due_date.strftime('%d.%m.%Y')}"


def get_period_string(from_date, due_date):
    if (
            from_date.day == 1
            and from_date.month == 1
            and due_date.day == 31
            and due_date.month == 12
            and from_date.year == due_date.year
    ):
        return "за " + str(from_date.year) + " год"
    elif (
            from_date.day == 1
            and from_date.month == 1
            and due_date.day == 31
            and due_date.month == 3
            and from_date.year == due_date.year
    ):
        return "за 1 квартал " + str(from_date.year) + " года"
    elif (
            from_date.day == 1
            and from_date.month == 4
            and due_date.day == 30
            and due_date.month == 6
            and from_date.year == due_date.year
    ):
        return "за 2 квартал " + str(from_date.year) + " года"
    elif (
            from_date.day == 1
            and from_date.month == 7
            and due_date.day == 30
            and due_date.month == 9
            and from_date.year == due_date.year
    ):
        return "за 3 квартал " + str(from_date.year) + " года"
    elif (
            from_date.day == 1
            and from_date.month == 10
            and due_date.day == 31
            and due_date.month == 12
            and from_date.year == due_date.year
    ):
        return "за 4 квартал " + str(from_date.year) + " года"
    elif (
            from_date.day == 1
            and from_date.month == 1
            and due_date.day == 30
            and due_date.month == 6
            and from_date.year == due_date.year
    ):
        return "за 1 полугодие " + str(from_date.year) + " года"
    elif (
            from_date.day == 1
            and from_date.month == 7
            and due_date.day == 31
            and due_date.month == 12
            and from_date.year == due_date.year
    ):
        return "за 2 полугодие " + str(from_date.year) + " года"
    else:
        return "с " + from_date.strftime("%d.%m.%Y") + " по " + due_date.strftime("%d.%m.%Y")


def first_day_of_current_month():
    """
    Получаем дату начала текущего месяца
    """
    today = datetime.date.today()
    first_day_of_month = datetime.date(today.year, today.month, 1)

    return first_day_of_month.strftime('%d.%m.%Y')


def current_day():
    """
    Получаем сегодняшнюю дату
    """
    today = datetime.date.today()

    return today.strftime('%d.%m.%Y')


@permission_required('deloreports.show_reports_master_report', login_url='/deloreports/')
def reports_master_report(request):
    """
    Сведения о поступлении, рассмотрении и характере обращений граждан
    """
    result = []
    result_rubric = []
    dep_name = []
    rows = []
    result_executors = []
    is_only_mayor = False
    reports_master_report = ReportsMasterReportApi()
    from_date = datetime(datetime.now().year, 1, 1).date()
    due_date = datetime.now().date()
    if request.method == 'POST':
        form = DateFormSimpleDep(request.POST)
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            dep_name = form.cleaned_data["dep_name"]

            cache_key = get_master_report_cache_key(request.user.id, from_date, due_date, dep_name)
            cached_results = cache.get(cache_key)
            if cached_results:
                result, rows, result_rubric, result_executors, is_only_mayor = cached_results
            else:
                if from_date <= due_date:
                    result, rows, result_rubric, result_executors, is_only_mayor = reports_master_report.get_report_data(
                        from_date, due_date, dep_name)
                    cache.set(cache_key, [result, rows, result_rubric, result_executors, is_only_mayor], 60 * 3)

    form = DateFormSimpleDep(request.POST)
    context = {
        'report_data': result,
        'rows': rows,
        'result_rubric': result_rubric,
        'cur_year': due_date.year,
        'prev_year': due_date.year - 1,
        'form': form,
        'result_executors': result_executors,
        'is_only_mayor': is_only_mayor
    }
    return render(request, 'deloreports/reports_master_report.html', context)


@permission_required('deloreports.show_reports_master_report', login_url='/deloreports/')
def get_citizen_list(request):
    result = []
    if request.method == 'POST':
        dep_list = request.POST.get('depName')
        from_date = datetime.strptime(request.POST.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.POST.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")

        reports_master_report = ReportsMasterReportApi()
        rows = reports_master_report.get_citizen_list(from_date, due_date, dep_list)
        if rows:
            for row in rows:
                t = [row[1], row[2].strftime('%d.%m.%Y'), row[3], row[4], row[5].strftime('%d.%m.%Y')]
                result.append(t)

    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


# def reports_master_report_export_excel(request, report_data, result_rubric, cur_year, prev_year, result_executors, is_only_mayor) :
#     with closing(StringIO()) as report_file:
#         with closing(xlsxwriter.Workbook(report_file)) as workbook:
#             reportTableHeader = [
#                 u"№",
#                 u"Контролер",
#                 u"Номер документа",
#                 u"Дата регистрации",
#                 u"Ответственный исполнитель",
#                 u"Содержание",
#                 u"Текст резолюции"
#             ]
#             # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
#             worksheet = workbook.add_worksheet(u"Отчет")

#             worksheet.set_column('A:A', 10)
#             worksheet.set_column('B:B', 18)
#             worksheet.set_column('C:C', 18)
#             worksheet.set_column('D:D', 17)
#             worksheet.set_column('E:E', 30)
#             worksheet.set_column('F:F', 60)

#             # Форматирование шапки
#             header = workbook.add_format({'bold': True})
#             header.set_align('left')
#             header.set_align('top')
#             header.set_border(1)
#             header.set_border_color("#000000")

#             # Форматирование тела таблицы
#             body = workbook.add_format()
#             body.set_text_wrap()
#             body.set_align('left')
#             body.set_align('top')
#             body.set_border(1)
#             body.set_border_color("#000000")

#             # вывод значений
#             for i, row in enumerate(result):
#                 if i == 0:
#                     worksheet.write(i, 0, reportTableHeader[0], header)
#                     worksheet.write(i, 1, reportTableHeader[1], header)
#                     worksheet.write(i, 2, reportTableHeader[2], header)
#                     worksheet.write(i, 3, reportTableHeader[3], header)
#                     worksheet.write(i, 4, reportTableHeader[4], header)
#                     worksheet.write(i, 5, reportTableHeader[5], header)
#                     worksheet.write(i, 6, reportTableHeader[6], header)

#                 worksheet.write(i + 1, 0, i + 1, body)
#                 worksheet.write(i + 1, 1, unicode(row[0] if row[0] else "Не указан"), body)
#                 worksheet.write(i + 1, 2, unicode(row[1] if row[1] else "Не указан"), body)
#                 worksheet.write(i + 1, 3, unicode(row[2] if row[2] else "Не указана"), body)
#                 worksheet.write(i + 1, 4, unicode(row[3] if row[3] else "Не указан"), body)
#                 worksheet.write(i + 1, 5, unicode(row[4] if row[4] else "Не указано"), body)
#                 worksheet.write(i + 1, 6, unicode(row[5] if row[5] else "Не указано"), body)

#         report = report_file.getvalue()

#     # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

#     # Имя файла отчета, предложенное пользователю при сохранении отчета.
#     # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
#     report_filename = u'report.xlsx'

#     # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
#     # См. также mimetypes.guess_type из стандартной библиотеки Python:
#     #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
#     #content_type = 'application/ms-excel'
#     content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

#     response = HttpResponse(report, content_type=content_type)
#     response['Content-Length'] = len(report)    # Полезно, но не обязательно.
#     response['Content-Disposition'] = u'attachment; filename=%s' % report_filename

#     return response

@permission_required('deloreports.show_district_themes_report', login_url='/deloreports/')
def district_themes_report(request):
    """
    Информация о тематиках обращений граждан в разрезе округов города Вологды
    """
    result = []
    district_themes_report = DistrictThemesReportApi()
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            result = district_themes_report.get_report_data(from_date, due_date)

    context = {
        'report_data': result["report_data"] if result and result["report_data"] else 0,
        'summary_data': result["summary_data"] if result and result["summary_data"] else 0,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/district_themes_report.html', context)


##### Отчет об эффективности работы органов Администрации города Вологды  #####
@permission_required('deloreports.show_effectiveness_report', login_url='/deloreports/')
def effectiveness_report(request):
    """
    Результаты оценки эффективности деятельности органов Администрации города Вологды
    """
    result = []
    sum_data = {}
    from_date = due_date = None
    form = DateRangeForm(request.POST)
    if request.method == 'POST':
        effectiveness_report = EffectivenessReportApi()
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            cache_key = get_effectiveness_report_cache_key(request.user.id, from_date, due_date)
            cached_results = cache.get(cache_key)

            if cached_results:
                result, sum_data = cached_results
            else:
                result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
                cache.set(cache_key, [result, sum_data], 300)

    context = {
        'report_data': result,
        'sum_data': sum_data,
        'form': form,
        'from_date': from_date,
        'due_date': due_date
    }
    return render(request, 'deloreports/effectiveness_report.html', context)


def effectiveness_report_export_excel(request):
    """
    Выгрузка отчёта Результаты оценки эффективности деятельности органов Администрации города Вологды (временные показатели)
    """
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")

    cache_key = get_effectiveness_report_cache_key(request.user.id, from_date, due_date)
    cached_results = cache.get(cache_key)  # получение данных из хэша
    if cached_results:
        result, sum_data = cached_results
    else:  # cache = none
        effectiveness_report = EffectivenessReportApi()
        result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
        cache.set(cache_key, [result, sum_data], 60 * 10)  # передаю данные в хэш

    with BytesIO() as report_file:
        with xlsxwriter.Workbook(report_file) as workbook:
            worksheet = workbook.add_worksheet("Временные показатели")
            worksheet.fit_to_pages(1, 1)
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.24, right=0, top=0.75, bottom=0.75)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 16})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1,
                 'text_wrap': True})
            header_table_bold_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True})
            header_table_bold_green_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#92d050'})
            header_table_bold_orange_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#fcd5b4'})
            result_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 14,
                 'border': 1})
            point_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'num_format': '0.0', 'border': 1, 'fg_color': '#92d050'})
            percent_format1 = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4', 'num_format': '0.0%'})
            percent_format2 = workbook.add_format(
                {'align': 'center', 'bold': False, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'num_format': '0.00%'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            header_table_bold_format.set_align('vcenter')
            header_table_bold_green_format.set_align('vcenter')
            header_table_bold_orange_format.set_align('vcenter')
            result_format.set_align('vcenter')
            point_format.set_align('vcenter')
            percent_format1.set_align('vcenter')
            percent_format2.set_align('vcenter')

            # Установка овка ширины столбцов
            worksheet.set_column('A:A', 3.86)
            worksheet.set_column('B:B', 21.14)
            worksheet.set_column('C:C', 11)
            worksheet.set_column('D:D', 13.86)
            worksheet.set_column('E:E', 15.43)
            worksheet.set_column('F:F', 8.14)
            worksheet.set_column('G:G', 13)
            worksheet.set_column('H:H', 18.57)
            worksheet.set_column('I:I', 14.71)
            worksheet.set_column('J:J', 8.57)
            worksheet.set_column('K:K', 11.43)
            worksheet.set_column('L:L', 10.86)
            worksheet.set_column('M:M', 11.57)
            worksheet.set_column('N:N', 12)
            worksheet.set_column('O:O', 14.86)
            worksheet.set_column('P:P', 8.14)
            worksheet.set_column('Q:Q', 15.43)
            worksheet.set_column('O:O', 14.86)
            worksheet.set_column('R:R', 17)
            worksheet.set_column('S:S', 12.71)
            worksheet.set_column('T:T', 10.43)
            worksheet.set_column('U:U', 11.71)
            worksheet.set_column('V:V', 12.14)
            # Установка высоты строк 
            worksheet.set_row(0, 33.75)
            worksheet.set_row(1, 16.5)
            worksheet.set_row(2, 33)
            worksheet.set_row(3, 36)
            worksheet.set_row(4, 137.25)

            # Шапка таблицы
            worksheet.merge_range('A1:S1',
                                  'Результаты оценки эффективности деятельности органов Администрации города Вологды ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)

            worksheet.merge_range('A3:A5', '№ п/п', header_table_format)
            worksheet.merge_range('B3:B5', 'Орган Администрации города Вологды', header_table_format)
            worksheet.merge_range('C3:V3', 'Временные показатели оценки эффективности', header_table_bold_format)
            worksheet.merge_range('C4:L4',
                                  'Поручения Мэра города Вологды прямые, с оперативных и тематических совещаний, по правовым актам и служебной корреспонденции',
                                  header_table_format)
            worksheet.merge_range('M4:V4', 'Обращения граждан', header_table_format)
            i = 2
            for headercol_str in ("Всего", "Исполнено", "% выполнения показателя", "Кол-во баллов",
                                  "Всего справок, ходатайств, служебных записок по контролю",
                                  "Число  справок, ходатайств, служебных записок по контролю, представленных с нарушением срока",
                                  "% выполнения показателя",
                                  "Кол-во баллов", "Общий балл", "Средний %", "Всего поручений", "Исполено поручений",
                                  "% выполнения показателя", "Кол-во баллов",
                                  "Всего проектов ответов за подписью Мэра города Вологды",
                                  "Количество проектов ответов, представленных с нарушением срока",
                                  "% выполнения показателя", "Кол-во баллов", "Общий балл", "Средний %"):
                if headercol_str in ("Кол-во баллов", "Общий балл"):
                    worksheet.write(4, i, str(headercol_str), header_table_bold_green_format)
                elif headercol_str == "Средний %":
                    worksheet.write(4, i, str(headercol_str), header_table_bold_orange_format)
                else:
                    worksheet.write(4, i, str(headercol_str), header_table_format)
                i += 1
            i = 5
            for department in result:
                # Установка высоты для строки данных
                worksheet.set_row(i, 33)
                worksheet.write(i, 0, i - 4, result_format)
                worksheet.write(i, 1, department["name"], result_format)
                worksheet.write(i, 2, department["main_resolutions_all"], result_format)
                worksheet.write(i, 3, department["main_resolutions_completed"], result_format)
                worksheet.write(i, 4, department["main_resolutions_completed_percent"] / 100, percent_format2)
                worksheet.write(i, 5, department["main_resolutions_completed_points"], point_format)
                worksheet.write(i, 6, department["control_all"], result_format)
                worksheet.write(i, 7, department["control_neispoln"], result_format)
                worksheet.write(i, 8, department["control_percent"] / 100, percent_format2)
                worksheet.write(i, 9, department["control_points"], point_format)
                worksheet.write(i, 10, department["official_total_points"], point_format)
                worksheet.write(i, 11, department["official_avg_percent"] / 100, percent_format1)
                worksheet.write(i, 12, department["citizen_resolutions_data"], result_format)
                worksheet.write(i, 13, department["citizen_resolutions_completed_data"], result_format)
                worksheet.write(i, 14, department["citizen_resolutions_percent"] / 100, percent_format2)
                worksheet.write(i, 15, department["citizen_resolutions_points"], point_format)
                worksheet.write(i, 16, department["citizen_answers_all_data"], result_format)
                worksheet.write(i, 17, department["citizen_answers_failed_data"], result_format)
                worksheet.write(i, 18, department["citizen_answers_percent"] / 100, percent_format2)
                worksheet.write(i, 19, department["citizen_answers_points"], point_format)
                worksheet.write(i, 20, department["citizen_total_points"], point_format)
                worksheet.write(i, 21, department["citizen_avg_percent"] / 100, percent_format1)
                i += 1
            # Печать последней строки ВСЕГО
            worksheet.set_row(i, 33)
            worksheet.write(i, 0, i - 4, result_format)
            worksheet.write(i, 1, "ВСЕГО", result_format)
            worksheet.write(i, 2, sum_data["main_resolutions_all"], result_format)
            worksheet.write(i, 3, sum_data["main_resolutions_completed"], result_format)
            worksheet.write(i, 4, "-", result_format)
            worksheet.write(i, 5, "", result_format)
            worksheet.write(i, 6, sum_data["control_all"], result_format)
            worksheet.write(i, 7, sum_data["control_neispoln"], result_format)
            worksheet.write(i, 8, "-", percent_format2)
            worksheet.write(i, 9, "", point_format)
            worksheet.write(i, 10, "", point_format)
            worksheet.write(i, 11, "", percent_format1)
            worksheet.write(i, 12, sum_data["citizen_resolutions_data"], result_format)
            worksheet.write(i, 13, sum_data["citizen_resolutions_completed_data"], result_format)
            worksheet.write(i, 14, "-", percent_format2)
            worksheet.write(i, 15, "", point_format)
            worksheet.write(i, 16, sum_data["citizen_answers_all_data"], result_format)
            worksheet.write(i, 17, sum_data["citizen_answers_failed_data"], result_format)
            worksheet.write(i, 18, "-", percent_format2)
            worksheet.write(i, 19, "", point_format)
            worksheet.write(i, 20, "", point_format)
            worksheet.write(i, 21, "", percent_format1)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'effectiveness.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


def effectiveness_report_export_excel_betatest(request):
    """
    Выгрузка отчёта Результаты оценки эффективности деятельности органов Администрации города Вологды (временные
    показатели)(Beta testing)
    """
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")

    cache_key = get_effectiveness_report_betatest_cache_key(request.user.id, from_date, due_date)
    cached_results = cache.get(cache_key)  # получение данных из хэша
    if cached_results:
        result, sum_data = cached_results
    else:  # cache = none
        effectiveness_report = EffectivenessReportBetaApi()
        result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
        cache.set(cache_key, [result, sum_data], 60 * 10)  # передаю данные в хэш

    # with closing(StringIO()) as report_file:
    with BytesIO() as report_file:
        # with closing(xlsxwriter.Workbook(report_file)) as workbook:
        with xlsxwriter.Workbook(report_file) as workbook:
            worksheet = workbook.add_worksheet("Временные показатели")
            worksheet.fit_to_pages(1, 1)
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.24, right=0, top=0.75, bottom=0.75)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 16})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1,
                 'text_wrap': True})
            header_table_bold_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True})
            header_table_bold_green_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#92d050'})
            header_table_bold_orange_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#fcd5b4'})
            result_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1,
                 'text_wrap': True})
            mark_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4'})
            percent_format_orange = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4', 'num_format': '0.0%'})
            percent_format_green = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#92d050', 'num_format': '0.0%'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            header_table_bold_format.set_align('vcenter')
            header_table_bold_green_format.set_align('vcenter')
            header_table_bold_orange_format.set_align('vcenter')
            result_format.set_align('vcenter')
            mark_format.set_align('vcenter')
            percent_format_orange.set_align('vcenter')
            percent_format_green.set_align('vcenter')

            # Установка овка ширины столбцов
            worksheet.set_column('A:A', 3.86)
            worksheet.set_column('B:B', 21.14)
            worksheet.set_column('C:C', 11.14)
            worksheet.set_column('D:D', 11.43)
            worksheet.set_column('E:E', 10.43)
            worksheet.set_column('F:F', 12)
            worksheet.set_column('G:G', 12.71)
            worksheet.set_column('H:H', 10.43)
            worksheet.set_column('I:I', 16)
            worksheet.set_column('J:J', 15.29)
            worksheet.set_column('K:K', 10.43)
            worksheet.set_column('L:L', 10.43)
            worksheet.set_column('M:M', 8.14)
            worksheet.set_column('N:N', 11.14)
            worksheet.set_column('O:O', 11.43)
            worksheet.set_column('P:P', 10.43)
            worksheet.set_column('Q:Q', 11.43)
            worksheet.set_column('R:R', 13.43)
            worksheet.set_column('S:S', 10.43)
            worksheet.set_column('T:T', 16)
            worksheet.set_column('U:U', 15.29)
            worksheet.set_column('V:V', 10.43)
            worksheet.set_column('W:W', 10.43)
            worksheet.set_column('X:X', 8.14)
            worksheet.set_column('Y:Y', 11.71)
            worksheet.set_column('Z:Z', 11.29)
            worksheet.set_column('AA:AA', 8.14)
            # Установка высоты строк 
            worksheet.set_row(0, 33.75)
            worksheet.set_row(1, 16.5)
            worksheet.set_row(2, 33)
            worksheet.set_row(3, 68.25)
            worksheet.set_row(4, 166.5)

            # Шапка таблицы
            worksheet.merge_range('A1:AA1',
                                  'Результаты оценки эффективности деятельности органов Администрации города Вологды ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)

            worksheet.merge_range('A3:A5', '№ п/п', header_table_format)
            worksheet.merge_range('B3:B5', 'Орган Администрации города Вологды', header_table_format)
            worksheet.merge_range('C3:AA3', 'Временные показатели оценки эффективности', header_table_bold_format)
            worksheet.merge_range('C4:M4',
                                  'Своевременность исполнения поручений Президента РФ, Губернатора ВО, Правительства ВО, Главы города Вологды, Мэра города Вологды, служебных документов',
                                  header_table_format)
            worksheet.merge_range('N4:X4', 'Своевременность рассмотрения обращений граждан', header_table_format)
            worksheet.merge_range('Y4:AA4', 'Своевременность рассмотрения обращений граждан в соответствии с 59-ФЗ',
                                  header_table_format)
            i = 2
            for headercol_str in (
                    "Всего поручений", "Исполнено", "% выполнения",
                    "Всего справок, ходатайств, служебных записок по контролю",
                    "Число  справок, ходатайств, служебных записок по контролю, представленных с нарушением срока",
                    "% выполнения", "Всего поручений (соисполнение)",
                    "Число нарушений соисполнителями срока представления информации отв. исполнителю",
                    "% выполнения", "% выполнения показателя", "Балл показателя", "Всего поручений", "Исполено",
                    "% выполнения",
                    "Всего проектов ответов за подписью Мэра города Вологды",
                    "Количество проектов ответов, представленных с нарушением срока",
                    "% выполнения", "Всего поручений (соисполнение)",
                    "Число нарушений соисполнителями срока представления информации отв. исполнителю", "% выполнения",
                    "% выполнения показателя", "Балл показателя",
                    "Число нарушений срока рассмотрения обращений граждан, 59-ФЗ", "% выполнения показателя",
                    "Балл показателя"):
                if headercol_str == "% выполнения":
                    worksheet.write(4, i, str(headercol_str), header_table_bold_green_format)
                elif headercol_str in ("Балл показателя", "% выполнения показателя"):
                    worksheet.write(4, i, str(headercol_str), header_table_bold_orange_format)
                else:
                    worksheet.write(4, i, str(headercol_str), header_table_format)
                i += 1
            i = 5
            for department in result:
                # Установка высоты для строки данных
                worksheet.set_row(i, 36)
                worksheet.write(i, 0, i - 4, result_format)
                worksheet.write(i, 1, department["name"], result_format)
                worksheet.write(i, 2, department["main_resolutions_all"], result_format)
                worksheet.write(i, 3, department["main_resolutions_completed"], result_format)
                worksheet.write(i, 4, department["main_resolutions_completed_percent"] / 100, percent_format_green)
                worksheet.write(i, 5, department["control_all"], result_format)
                worksheet.write(i, 6, department["control_neispoln"], result_format)
                worksheet.write(i, 7, department["control_percent"] / 100, percent_format_green)
                worksheet.write(i, 8, department["resolutions_official_all_irresponsible"], result_format)
                worksheet.write(i, 9, department["resolutions_official_irresponsible"], result_format)
                worksheet.write(i, 10, department["resolutions_official_irresponsible_percent"] / 100,
                                percent_format_green)
                worksheet.write(i, 11, department["official_avg_percent"] / 100, percent_format_orange)
                worksheet.write(i, 12, department["official_total_points"], mark_format)
                worksheet.write(i, 13, department["citizen_resolutions_data"], result_format)
                worksheet.write(i, 14, department["citizen_resolutions_completed_data"], result_format)
                worksheet.write(i, 15, department["citizen_resolutions_percent"] / 100, percent_format_green)
                worksheet.write(i, 16, department["citizen_answers_all_data"], result_format)
                worksheet.write(i, 17, department["citizen_answers_failed_data"], result_format)
                worksheet.write(i, 18, department["citizen_answers_percent"] / 100, percent_format_green)
                worksheet.write(i, 19, department["resolutions_citizen_all_irresponsible"], result_format)
                worksheet.write(i, 20, department["resolutions_citizen_irresponsible"], result_format)
                worksheet.write(i, 21, department["resolutions_citizen_irresponsible_percent"] / 100,
                                percent_format_green)
                worksheet.write(i, 22, department["citizen_avg_percent"] / 100, percent_format_orange)
                worksheet.write(i, 23, department["citizen_total_points"], mark_format)
                worksheet.write(i, 24, department["citizen_resolutions_59fz"], result_format)
                worksheet.write(i, 25, department["citizen_resolutions_59fz_percent"] / 100, percent_format_orange)
                worksheet.write(i, 26, department["total_points_59fz"], mark_format)

                i += 1
            # Печать последней строки ВСЕГО
            worksheet.set_row(i, 36)
            worksheet.write(i, 0, i - 4, result_format)
            worksheet.write(i, 1, "ВСЕГО", result_format)
            worksheet.write(i, 2, sum_data["main_resolutions_all"], result_format)
            worksheet.write(i, 3, sum_data["main_resolutions_completed"], result_format)
            worksheet.write(i, 4, "-", percent_format_green)
            worksheet.write(i, 5, sum_data["control_all"], result_format)
            worksheet.write(i, 6, sum_data["control_neispoln"], result_format)
            worksheet.write(i, 7, "-", percent_format_green)
            worksheet.write(i, 8, sum_data["resolutions_official_all_irresponsible"], result_format)
            worksheet.write(i, 9, sum_data["resolutions_official_irresponsible"], result_format)
            worksheet.write(i, 10, "-", percent_format_green)
            worksheet.write(i, 11, "-", percent_format_orange)
            worksheet.write(i, 12, "", mark_format)
            worksheet.write(i, 13, sum_data["citizen_resolutions_data"], result_format)
            worksheet.write(i, 14, sum_data["citizen_resolutions_completed_data"], result_format)
            worksheet.write(i, 15, "-", percent_format_green)
            worksheet.write(i, 16, sum_data["citizen_answers_all_data"], result_format)
            worksheet.write(i, 17, sum_data["citizen_answers_failed_data"], result_format)
            worksheet.write(i, 18, "-", percent_format_green)
            worksheet.write(i, 19, sum_data["resolutions_citizen_all_irresponsible"], result_format)
            worksheet.write(i, 20, sum_data["resolutions_citizen_irresponsible"], result_format)
            worksheet.write(i, 21, "-", percent_format_green)
            worksheet.write(i, 22, "-", percent_format_orange)
            worksheet.write(i, 23, "", mark_format)
            worksheet.write(i, 24, sum_data["citizen_resolutions_59fz"], result_format)
            worksheet.write(i, 25, "-", percent_format_orange)
            worksheet.write(i, 26, "", mark_format)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'effectiveness.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


##### Бета-тест Отчет об эффективности работы органов Администрации города Вологды  #####
@permission_required('deloreports.show_effectiveness_report', login_url='/deloreports/')
def effectiveness_report_betatest(request):
    """
    Результаты оценки эффективности деятельности органов Администрации города Вологды (Beta testing)
    """
    result = []
    result_dgp = []
    sum_data = {}
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    effectiveness_report = EffectivenessReportBetaApi()
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            cache_key = get_effectiveness_report_betatest_cache_key(request.user.id, from_date, due_date)
            cached_results = cache.get(cache_key)  # получение данных из хэша

            if cached_results:
                result, sum_data = cached_results  # распаковка данных хэша
            else:  # cache = none
                result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
                cache.set(cache_key, [result, sum_data], 60 * 5)  # передаю данные в хэш

    context = {
        'report_data': result,
        'sum_data': sum_data,
        'form': form,
        'from_date': from_date,
        'due_date': due_date
    }
    return render(request, 'deloreports/effectiveness_report_betatest.html', context)


# Отчет по обращениям граждан в разрезе рубрик
@permission_required('deloreports.show_citizen_rubrics_app_eight_report', login_url='/deloreports/')
def gov_app_eight_report(request):
    """
    Информация о рассмотрении поступивших обращений граждан, организаций (юридических лиц) и
    общественных объединений в органы исполнительной государственной власти и органы местного самоуправления области
    """
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            gov_app_eight_report = GovAppEightReportApi()
            result = gov_app_eight_report.get_report_data(from_date, due_date)
    context = {
        'report_data': result["rubrics_data"] if result else None,
        'summary_data': result["summary_data"] if result else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/gov_app_eight_report.html', context)


@permission_required('deloreports.show_court_documents_report', login_url='/deloreports/')
def not_assigned_court_documents(request):
    """
    Сведения о поступиших судебных документах, по которым нет резолюции ПУ
    """
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    not_assigned_court_documents_report_api = NotAssignedCourtDocumentsReportApi()
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            result = not_assigned_court_documents_report_api.get_report_data(from_date, due_date)

    context = {
        'report_data': list(result) if result else 0,
        'form': form,
        'from_date': from_date,
        'due_date': due_date
    }
    return render(request, 'deloreports/not_assigned_court_documents_report.html', context)


@permission_required('deloreports.show_correspondents_report', login_url='/deloreports/')
def correspondents_report(request):
    """
    Сведения о входящей служебной корреспонденции в разрезе корреспондентов
    """
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            correspondents_report = CorrespondentsReportApi()
            result = correspondents_report.get_report_data(from_date, due_date)
    context = {
        'result': result,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/correspondents_report.html', context)


@permission_required('deloreports.show_correspondents_foiv_report', login_url='/deloreports/')
def correspondents_foiv_report(request):
    """
    Перечень документов, поступивших от Федеральных органов исполнительной власти
    """
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            correspondents_foiv_report = CorrespondentsFoivReportApi()
            result = correspondents_foiv_report.get_report_data(from_date, due_date)

    context = {
        'result': result,
        # 'foiv_list' : result["foiv_list"] if result.get("foiv_list") else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }

    return render(request, 'deloreports/correspondents_foiv_report.html', context)


def court_documents_report(request):
    """
    Отчет по судебным документам поступившим в Администрацию города Вологды
    """
    result = []
    from_date = due_date = None
    form = DateRangeForm(request.POST)
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            court_documents_report = CourtDocumentsReportApi()
            result = court_documents_report.get_report_data(from_date, due_date)
    context = {
        'report_data': result["report_data"] if result else None,
        'summary_data': result["summary_data"] if result else None,
        'form': form,
        'is_mistakes': result["is_mistakes"] if result else None,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/court_documents_report.html', context)


def get_docs_list(request):
    """
    Получить списки документов для модального окна в отчетах по эффективности
    """
    result = []
    if request.method == 'POST':
        due = request.POST.get('due')
        from_date = datetime.strptime(request.POST.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.POST.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.POST.get("reportName").strip(' \t\n\r') if request.POST.get("reportName") else None
        if due:
            due = due.strip(' \t\n\r')
            effectiveness_report = EffectivenessReportApi()
            if report_name != None:
                if report_name == "get_report_effectiveness_control_all_list":
                    rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_control_neispoln_list":
                    rows = effectiveness_report.get_report_effectiveness_control_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_list":
                    rows = effectiveness_report.get_report_effectiveness_answers_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_neispoln_list":
                    rows = effectiveness_report.get_report_effectiveness_answers_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
                if report_name == "get_report_effectiveness_control_all_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_control_neispoln_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_control_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_answers_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_neispoln_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_answers_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
            else:
                rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
            if rows:
                for row in rows:
                    t = [row[0], row[1].strftime('%d.%m.%Y'), row[2], row[3]]
                    result.append(t)

    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


def download_list(request):
    """
    Экспорт перечня поручений из модального окна в Эксель в отчёте оценки эффективности
    """
    result = []
    if request.method == 'GET':
        due = request.GET.get('due') if request.GET.get('due') else ""
        from_date = datetime.strptime(request.GET.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.GET.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.GET.get("reportName").strip(' \t\n\r') if request.GET.get("reportName") else None
        if due:
            due = due.strip(' \t\n\r')
            effectiveness_report = EffectivenessReportBetaApi()
            if report_name != None:
                rows = None
                if report_name == "get_main_resolutions_all_list":
                    rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)
                elif report_name == "get_main_resolutions_completed_list":
                    rows = effectiveness_report.get_main_resolutions_completed_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_list":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_list(due, from_date,
                                                                                                  due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_completed_list":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_completed_list(due,
                                                                                                            from_date,
                                                                                                            due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
                elif report_name == "get_resolutions_double_date_change":
                    dep = effectiveness_report.get_dep_on_due(due)
                    rows = effectiveness_report.get_resolutions_double_date_change(
                        {"from_date": from_date, "due_date": due_date, "dep": dep, "resultquery": "list"})
                elif report_name == "get_main_resolutions_all_list_betatest":
                    rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)
                elif report_name == "get_main_resolutions_completed_list_betatest":
                    rows = effectiveness_report.get_main_resolutions_completed_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_list(due, from_date,
                                                                                                  due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_completed_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_completed_list(due,
                                                                                                            from_date,
                                                                                                            due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
                elif report_name == "get_resolutions_irresponsible_executor_all":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor_all(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": False})
                elif report_name == "get_resolutions_irresponsible_executor":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": False})
                elif report_name == "get_resolutions_citizen_irresponsible_executor_all":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor_all(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": True})
                elif report_name == "get_resolutions_citizen_irresponsible_executor":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": True})
                elif report_name == "get_resolutions_59fz":
                    rows = effectiveness_report.get_report_effectiveness_citizen_59fz(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True})
            else:
                rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)
            # print('download_list. здесь должна быть переменная rows: ', rows)
            if rows:
                for row in rows:
                    t = []
                    for val in row:
                        if type(val) == datetime:
                            t.append(val.strftime('%d.%m.%Y'))
                        else:
                            t.append(val)
                    result.append(t)
            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            if report_name != "get_resolutions_irresponsible_executor_all" and report_name != "get_resolutions_irresponsible_executor" and report_name != "get_resolutions_citizen_irresponsible_executor" and report_name != "get_resolutions_citizen_irresponsible_executor_all" and report_name != "get_resolutions_59fz":
                reportTableHeader = [
                    "№",
                    "Номер документа",
                    "Дата регистрации",
                    "Ответственный исполнитель",
                    "Содержание",
                    "Текст резолюции"
                ]
                # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
                worksheet = workbook.add_worksheet("Отчет")

                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 18)
                worksheet.set_column('C:C', 17)
                worksheet.set_column('D:D', 30)
                worksheet.set_column('E:E', 60)
                worksheet.set_column('F:F', 30)

                # Форматирование шапки
                header = workbook.add_format({'bold': True})
                header.set_align('left')
                header.set_align('top')
                header.set_border(1)
                header.set_border_color("#000000")

                # Форматирование тела таблицы
                body = workbook.add_format()
                body.set_text_wrap()
                body.set_align('left')
                body.set_align('top')
                body.set_border(1)
                body.set_border_color("#000000")

                # вывод значений
                for i, row in enumerate(result):
                    if i == 0:
                        worksheet.write(i, 0, reportTableHeader[0], header)
                        worksheet.write(i, 1, reportTableHeader[1], header)
                        worksheet.write(i, 2, reportTableHeader[2], header)
                        worksheet.write(i, 3, reportTableHeader[3], header)
                        worksheet.write(i, 4, reportTableHeader[4], header)
                        worksheet.write(i, 5, reportTableHeader[5], header)

                    worksheet.write(i + 1, 0, i + 1, body)
                    worksheet.write(i + 1, 1, str(row[1] if row[1] else "Не указан"), body)
                    worksheet.write(i + 1, 2, str(row[2] if row[2] else "Не указана"), body)
                    worksheet.write(i + 1, 3, str(row[3] if row[3] else "Не указан"), body)
                    worksheet.write(i + 1, 4, str(row[4] if row[4] else "Не указано"), body)
                    worksheet.write(i + 1, 5, str(row[5] if row[5] else "Не указано"), body)
            else:
                if report_name == "get_resolutions_irresponsible_executor" or report_name == "get_resolutions_citizen_irresponsible_executor":
                    reportTableHeader = [
                        "№",
                        "Группа, №, дата, содержание документа",
                        "ФИО отв. исполнителя, поручение Мэра города Вологды",
                        "Плановая дата",
                        "ФИО соисполнителя",
                        "Информация о соисполнении"
                    ]
                elif report_name == "get_resolutions_irresponsible_executor_all" or report_name == "get_resolutions_citizen_irresponsible_executor_all":
                    reportTableHeader = [
                        "№",
                        "Группа, №, дата, содержание документа",
                        "ФИО отв. исполнителя, поручение Мэра города Вологды",
                        "Плановая дата",
                        "ФИО соисполнителя"
                    ]
                elif report_name == "get_resolutions_59fz":
                    reportTableHeader = [
                        "№",
                        "Группа, №, дата, содержание документа",
                        "ФИО отв. исполнителя, поручение Мэра города Вологды",
                        "Плановая дата",
                        "Фактическая дата",
                        "Нарушение срока (кол-во дней)"
                    ]
                # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
                worksheet = workbook.add_worksheet("Отчет")

                worksheet.set_column('A:A', 10)
                worksheet.set_column('B:B', 50)
                worksheet.set_column('C:C', 30)
                worksheet.set_column('D:D', 17)
                worksheet.set_column('E:E', 17)
                if report_name != "get_resolutions_irresponsible_executor_all" and report_name != "get_resolutions_citizen_irresponsible_executor_all":
                    worksheet.set_column('F:F', 30)

                # Форматирование шапки
                header = workbook.add_format({'bold': True})
                header.set_align('left')
                header.set_align('top')
                header.set_border(1)
                header.set_border_color("#000000")

                # Форматирование тела таблицы
                body = workbook.add_format()
                body.set_text_wrap()
                body.set_align('left')
                body.set_align('top')
                body.set_border(1)
                body.set_border_color("#000000")

                # вывод значений
                for i, row in enumerate(result):
                    if i == 0:
                        worksheet.write(i, 0, reportTableHeader[0], header)
                        worksheet.write(i, 1, reportTableHeader[1], header)
                        worksheet.write(i, 2, reportTableHeader[2], header)
                        worksheet.write(i, 3, reportTableHeader[3], header)
                        worksheet.write(i, 4, reportTableHeader[4], header)
                        if report_name != "get_resolutions_irresponsible_executor_all" and report_name != "get_resolutions_citizen_irresponsible_executor_all":
                            worksheet.write(i, 5, reportTableHeader[5], header)

                    worksheet.write(i + 1, 0, i + 1, body)
                    worksheet.write(i + 1, 1, str(row[0] if row[0] else "Не указаны"), body)
                    worksheet.write(i + 1, 2, str(row[1] if row[1] else "Не указаны"), body)
                    worksheet.write(i + 1, 3, str(row[2] if row[2] else "Не указана"), body)
                    if report_name != "get_resolutions_59fz":
                        worksheet.write(i + 1, 4, str(row[3] if row[3] else "Не указаны"), body)
                    else:
                        worksheet.write(i + 1, 4, str(row[3] if row[3] else "Не исполнено"), body)
                    if report_name != "get_resolutions_irresponsible_executor_all" and report_name != "get_resolutions_citizen_irresponsible_executor_all" and report_name != "get_resolutions_59fz":
                        worksheet.write(i + 1, 5, str(row[4] if row[4] else "Не указана"), body)
                    elif report_name == "get_resolutions_59fz":
                        worksheet.write(i + 1, 5, str(row[4] if row[4] else "-"), body)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'report.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


def get_resolutions_list(request):
    """
    Возвращает всплывающие списки документов поручений в модальном окне для отчетов по эффективности
    """
    result = []
    if request.method == 'POST':
        due = request.POST.get('due')
        from_date = datetime.strptime(request.POST.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.POST.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.POST.get("reportName").strip(' \t\n\r') if request.POST.get("reportName") else None
        if due:
            due = due.strip(' \t\n\r')
            effectiveness_report = EffectivenessReportApi()
            rows = None
            if report_name != None:
                if report_name == "get_main_resolutions_all_list":
                    rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)
                elif report_name == "get_main_resolutions_completed_list":
                    rows = effectiveness_report.get_main_resolutions_completed_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_list":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_list(due, from_date,
                                                                                                  due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_completed_list":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_completed_list(due,
                                                                                                            from_date,
                                                                                                            due_date)
                elif report_name == "get_resolutions_double_date_change":
                    dep = effectiveness_report.get_dep_on_due(due)
                    rows = effectiveness_report.get_resolutions_double_date_change(
                        {"from_date": from_date, "due_date": due_date, "dep": dep, "resultquery": "list"})
                elif report_name == "get_main_resolutions_all_list_betatest":
                    rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)
                elif report_name == "get_main_resolutions_completed_list_betatest":
                    rows = effectiveness_report.get_main_resolutions_completed_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_list(due, from_date,
                                                                                                  due_date)
                elif report_name == "get_report_effectiveness_citizen_resolutions_completed_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_citizen_resolutions_completed_list(due,
                                                                                                            from_date,
                                                                                                            due_date)
                elif report_name == "get_resolutions_irresponsible_executor_all":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor_all(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": False})
                elif report_name == "get_resolutions_irresponsible_executor":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": False})
                elif report_name == "get_resolutions_citizen_irresponsible_executor_all":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor_all(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": True})
                elif report_name == "get_resolutions_citizen_irresponsible_executor":
                    rows = effectiveness_report.get_resolutions_irresponsible_executor(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True, "citizen": True})
                elif report_name == "get_resolutions_59fz":
                    rows = effectiveness_report.get_report_effectiveness_citizen_59fz(
                        {"dep": effectiveness_report.get_department_by_due(due), "from_date": from_date,
                         "due_date": due_date, "list": True})

            else:
                rows = effectiveness_report.get_main_resolutions_all_list(due, from_date, due_date)

            if rows:
                for row in rows:
                    t = []
                    for val in row:
                        if type(val) == datetime:
                            t.append(val.strftime('%d.%m.%Y'))
                        else:
                            t.append(val)
                    result.append(t)

            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )
            # print('get_resolutions_list. здесь должна быть переменная rows: ', rows)
    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


def download_docs_list(request):
    """
    Экспорт из модального окна справок, ходатайств, служебных записок по контролю в отчёте эффективности
    """
    result = []
    if request.method == 'GET':
        due = request.GET.get('due') if request.GET.get('due') else ""
        from_date = datetime.strptime(request.GET.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.GET.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.GET.get("reportName").strip(' \t\n\r') if request.GET.get("reportName") else None
        if due:
            due = due.strip(' \t\n\r')
            effectiveness_report = EffectivenessReportApi()
            rows = None
            if report_name != None:
                if report_name == "get_report_effectiveness_control_all_list":
                    rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_control_neispoln_list":
                    rows = effectiveness_report.get_report_effectiveness_control_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_list":
                    rows = effectiveness_report.get_report_effectiveness_answers_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_neispoln_list":
                    rows = effectiveness_report.get_report_effectiveness_answers_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
                elif report_name == "get_report_effectiveness_control_all_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_control_neispoln_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_control_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_answers_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_answers_neispoln_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_answers_neispoln_list(due, from_date, due_date)
                elif report_name == "get_report_effectiveness_quality_control_negative_list_betatest":
                    rows = effectiveness_report.get_report_effectiveness_quality_control_negative_list(due, from_date,
                                                                                                       due_date)
            else:
                rows = effectiveness_report.get_report_effectiveness_control_all_list(due, from_date, due_date)
            if rows:
                for row in rows:
                    t = []
                    for val in row:
                        if type(val) == datetime:
                            t.append(val.strftime('%d.%m.%Y'))
                        else:
                            t.append(val)
                    result.append(t)
            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            reportTableHeader = [
                "№",
                "Номер документа",
                "Дата регистрации",
                "Исполнитель",
                "Содержание"
            ]
            # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
            worksheet = workbook.add_worksheet("Отчет")

            worksheet.set_column('A:A', 5)
            worksheet.set_column('B:B', 18)
            worksheet.set_column('C:C', 18)
            worksheet.set_column('D:D', 30)
            worksheet.set_column('E:E', 60)

            # Форматирование шапки
            header = workbook.add_format({'bold': True})
            header.set_align('left')
            header.set_align('top')
            header.set_border(1)
            header.set_border_color("#000000")

            # Форматирование тела таблицы
            body = workbook.add_format()
            body.set_text_wrap()
            body.set_align('left')
            body.set_align('top')
            body.set_border(1)
            body.set_border_color("#000000")

            # вывод значений
            for i, row in enumerate(result):
                if i == 0:
                    worksheet.write(i, 0, reportTableHeader[0], header)
                    worksheet.write(i, 1, reportTableHeader[1], header)
                    worksheet.write(i, 2, reportTableHeader[2], header)
                    worksheet.write(i, 3, reportTableHeader[3], header)
                    worksheet.write(i, 4, reportTableHeader[4], header)

                worksheet.write(i + 1, 0, i + 1, body)
                worksheet.write(i + 1, 1, str(row[0] if row[0] else "Не указан"), body)
                worksheet.write(i + 1, 2, str(row[1] if row[1] else "Не указана"), body)
                worksheet.write(i + 1, 3, str(row[2] if row[2] else "Не указан"), body)
                worksheet.write(i + 1, 4, str(row[3] if row[3] else "Не указано"), body)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'report.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


def effectiveness_quality_report_export_excel(request):
    """
    Выгрузка отчёта Результаты оценки эффективности деятельности органов Администрации города Вологды
    (качественные показатели)
    """
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")

    cache_key = get_effectiveness_report_cache_key(request.user.id, from_date, due_date)
    cached_results = cache.get(cache_key)  # получение данных из хэша
    if cached_results:
        result, sum_data = cached_results
    else:  # cache = none
        effectiveness_report = EffectivenessReportApi()
        result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
        cache.set(cache_key, [result, sum_data], 60 * 10)  # передаю данные в хэш

    with BytesIO() as report_file:
        with xlsxwriter.Workbook(report_file) as workbook:
            worksheet = workbook.add_worksheet("Качественные показатели")
            worksheet.fit_to_pages(1, 1)  # печать на одном листе
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.24, right=0, top=0.75, bottom=0.75)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 16})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1,
                 'text_wrap': True})
            header_table_bold_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True})
            header_table_bold_yellow_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#ffff00'})
            header_table_bold_orange_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#fcd5b4'})
            result_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 14,
                 'border': 1})
            point_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'num_format': '0.0', 'border': 1, 'fg_color': '#ffff00'})
            percent_format1 = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4', 'num_format': '0.0%'})
            percent_format2 = workbook.add_format(
                {'align': 'center', 'bold': False, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'num_format': '0.00%'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            header_table_bold_format.set_align('vcenter')
            header_table_bold_yellow_format.set_align('vcenter')
            header_table_bold_orange_format.set_align('vcenter')
            result_format.set_align('vcenter')
            point_format.set_align('vcenter')
            percent_format1.set_align('vcenter')
            percent_format2.set_align('vcenter')

            # Установка овка ширины столбцов
            worksheet.set_column('A:A', 8.43)
            worksheet.set_column('B:B', 21.71)
            worksheet.set_column('C:C', 11.86)
            worksheet.set_column('D:D', 21.86)
            worksheet.set_column('E:E', 13.29)
            worksheet.set_column('F:F', 8.43)
            worksheet.set_column('G:G', 17.86)
            worksheet.set_column('H:H', 13.57)
            worksheet.set_column('I:I', 8.43)
            worksheet.set_column('J:J', 8.86)
            worksheet.set_column('K:K', 10)
            worksheet.set_column('L:L', 12)
            worksheet.set_column('M:M', 16.14)
            worksheet.set_column('N:N', 16)
            worksheet.set_column('O:O', 8.43)
            worksheet.set_column('P:P', 10.43)

            # Установка высоты строк 
            worksheet.set_row(0, 18.75)
            worksheet.set_row(1, 16.5)
            worksheet.set_row(2, 25.5)
            worksheet.set_row(3, 39.75)
            worksheet.set_row(4, 104.25)

            # Шапка таблицы
            worksheet.merge_range('A1:P1',
                                  'Результаты оценки эффективности деятельности органов Администрации города Вологды ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)

            worksheet.merge_range('A3:A5', '№ п/п', header_table_format)
            worksheet.merge_range('B3:B5', 'Орган Администрации города Вологды', header_table_format)
            worksheet.merge_range('C3:P3', 'Качественные показатели оценки эффективности', header_table_bold_format)
            worksheet.merge_range('C4:K4',
                                  'Поручения Мэра города Вологды прямые, с оперативных и тематических совещаний, по правовым актам и служебной корреспонденции',
                                  header_table_format)
            worksheet.merge_range('L4:P4', 'Обращения граждан', header_table_format)
            i = 2
            for headercol_str in ("Всего поручений",
                                  "Количество отрицательно рассмотренных справок, ходатайств, служебных записок по контролю",
                                  "% выполнения показателя", "Кол-во баллов",
                                  "Количество поручений с переносом срока исполнения 2 и более раз",
                                  "% выполнения показателя",
                                  "Кол-во баллов", "Общий балл", "Средний %", "Всего поручений",
                                  "Количество повторных, неоднократных обращений граждан",
                                  "% выполнения показателя", "Кол-во баллов", "Средний %"):
                if headercol_str in ("Кол-во баллов", "Общий балл"):
                    worksheet.write(4, i, str(headercol_str), header_table_bold_yellow_format)
                elif headercol_str == "Средний %":
                    worksheet.write(4, i, str(headercol_str), header_table_bold_orange_format)
                else:
                    worksheet.write(4, i, str(headercol_str), header_table_format)
                i += 1
            i = 5
            for department in result:
                # Установка высоты для строки данных
                worksheet.set_row(i, 33)
                worksheet.write(i, 0, i - 4, result_format)
                worksheet.write(i, 1, department["name"], result_format)
                worksheet.write(i, 2, department["main_resolutions_all"], result_format)
                worksheet.write(i, 3, department["control_negative"], result_format)
                worksheet.write(i, 4, department["main_resolutions_control_negative_percent"] / 100, percent_format2)
                worksheet.write(i, 5, department["main_resolutions_control_negative_points"], point_format)
                worksheet.write(i, 6, department["resolution_double_date_change"], result_format)
                worksheet.write(i, 7, department["resolution_double_date_change_percent"] / 100, percent_format2)
                worksheet.write(i, 8, department["resolution_double_date_change_points"], point_format)
                worksheet.write(i, 9, department["official_total_points_quality"], point_format)
                worksheet.write(i, 10, department["official_avg_percent"] / 100, percent_format1)
                worksheet.write(i, 11, department["citizen_resolutions_data"], result_format)
                worksheet.write(i, 12, department["citizen_resolutions_several_data"], result_format)
                worksheet.write(i, 13, department["citizen_resolutions_percent"] / 100, percent_format2)
                worksheet.write(i, 14, department["citizen_resolutions_points"], point_format)
                worksheet.write(i, 15, department["citizen_avg_percent"] / 100, percent_format1)
                i += 1
            # Печать последней строки ВСЕГО
            worksheet.set_row(i, 33)
            worksheet.write(i, 0, i - 4, result_format)
            worksheet.write(i, 1, "ВСЕГО", result_format)
            worksheet.write(i, 2, sum_data["main_resolutions_all"], result_format)
            worksheet.write(i, 3, sum_data["control_negative"], result_format)
            worksheet.write(i, 4, "-", result_format)
            worksheet.write(i, 5, "", point_format)
            worksheet.write(i, 6, sum_data["resolution_double_date_change"], result_format)
            worksheet.write(i, 7, "-", percent_format2)
            worksheet.write(i, 8, "", point_format)
            worksheet.write(i, 9, "", point_format)
            worksheet.write(i, 10, "", percent_format1)
            worksheet.write(i, 11, sum_data["citizen_resolutions_data"], result_format)
            worksheet.write(i, 12, sum_data["citizen_resolutions_several_data"], result_format)
            worksheet.write(i, 13, "-", percent_format2)
            worksheet.write(i, 14, "", point_format)
            worksheet.write(i, 15, "", percent_format1)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'effectiveness_quality.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


def effectiveness_quality_report_export_excel_betatest(request):
    """
    Эффективность по новому формату, формирует табличку качественных показателей с нужным форматированием
    """
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")

    cache_key = get_effectiveness_report_betatest_cache_key(request.user.id, from_date, due_date)
    cached_results = cache.get(cache_key)  # получение данных из хэша
    if cached_results:
        result, sum_data = cached_results
    else:  # cache = none
        effectiveness_report = EffectivenessReportBetaApi()
        result, sum_data = effectiveness_report.get_report_data(from_date, due_date)
        cache.set(cache_key, [result, sum_data], 60 * 10)  # передаю данные в хэш

    # with closing(StringIO()) as report_file:
    with BytesIO() as report_file:
        # with closing(xlsxwriter.Workbook(report_file)) as workbook:
        with xlsxwriter.Workbook(report_file) as workbook:
            worksheet = workbook.add_worksheet("Качественные показатели")
            worksheet.fit_to_pages(1, 1)  # печать на одном листе
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.24, right=0, top=0.75, bottom=0.75)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 16})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1,
                 'text_wrap': True})
            header_table_bold_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True})
            header_table_bold_yellow_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#ffff00'})
            header_table_bold_orange_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 13, 'border': 1, 'text_wrap': True, 'fg_color': '#fcd5b4'})
            result_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1,
                 'text_wrap': True})
            mark_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4'})
            percent_format_yellow = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#ffff00', 'num_format': '0.0%'})
            percent_format_orange = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 14, 'border': 1, 'fg_color': '#fcd5b4', 'num_format': '0.0%'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            header_table_bold_format.set_align('vcenter')
            header_table_bold_yellow_format.set_align('vcenter')
            header_table_bold_orange_format.set_align('vcenter')
            result_format.set_align('vcenter')
            mark_format.set_align('vcenter')
            percent_format_yellow.set_align('vcenter')
            percent_format_orange.set_align('vcenter')

            # Установка овка ширины столбцов
            worksheet.set_column('A:A', 8.43)
            worksheet.set_column('B:B', 21.71)
            worksheet.set_column('C:C', 15.14)
            worksheet.set_column('D:D', 21.86)
            worksheet.set_column('E:E', 14.71)
            worksheet.set_column('F:F', 15.43)
            worksheet.set_column('G:G', 13.43)
            worksheet.set_column('H:H', 15.14)
            worksheet.set_column('I:I', 21.86)
            worksheet.set_column('J:J', 14.71)
            worksheet.set_column('K:K', 15.43)
            worksheet.set_column('L:L', 13.43)

            # Установка высоты строк 
            worksheet.set_row(0, 18.75)
            worksheet.set_row(1, 16.5)
            worksheet.set_row(2, 25.5)
            worksheet.set_row(3, 39.75)
            worksheet.set_row(4, 104.25)

            # Шапка таблицы
            worksheet.merge_range('A1:L1',
                                  'Результаты оценки эффективности деятельности органов Администрации города Вологды ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)

            worksheet.merge_range('A3:A5', '№ п/п', header_table_format)
            worksheet.merge_range('B3:B5', 'Орган Администрации города Вологды', header_table_format)
            worksheet.merge_range('C3:L3', 'Качественные показатели оценки эффективности', header_table_bold_format)
            worksheet.merge_range('C4:G4',
                                  'Качество исполнения поручений Президента РФ, Губернатора ВО, Правительства ВО, Главы города Вологды, Мэра города Вологды, служебных документов',
                                  header_table_format)
            worksheet.merge_range('H4:L4', 'Качество рассмотрения обращений граждан', header_table_format)
            i = 2
            for headercol_str in ("Всего поручений",
                                  "Количество отрицательно рассмотренных справок, ходатайств, служебных записок по контролю",
                                  "% выполнения", "% выполнения показателя", "Балл показателя",
                                  "Всего поручений", "Количество повторных, неоднократных обращений граждан",
                                  "% выполнения", "% выполнения показателя", "Балл показателя"):
                if headercol_str in ("% выполнения показателя", "Балл показателя"):
                    worksheet.write(4, i, str(headercol_str), header_table_bold_orange_format)
                elif headercol_str == "% выполнения":
                    worksheet.write(4, i, str(headercol_str), header_table_bold_yellow_format)
                else:
                    worksheet.write(4, i, str(headercol_str), header_table_format)
                i += 1
            i = 5
            for department in result:
                # Установка высоты для строки данных
                worksheet.set_row(i, 36)
                worksheet.write(i, 0, i - 4, result_format)
                worksheet.write(i, 1, department["name"], result_format)
                worksheet.write(i, 2, department["main_resolutions_all"], result_format)
                worksheet.write(i, 3, department["control_negative"], result_format)
                worksheet.write(i, 4, department["main_resolutions_control_negative_percent"] / 100,
                                percent_format_yellow)
                worksheet.write(i, 5, department["official_avg_percent"] / 100, percent_format_orange)
                worksheet.write(i, 6, department["official_total_points"], mark_format)
                worksheet.write(i, 7, department["citizen_resolutions_data"], result_format)
                worksheet.write(i, 8, department["citizen_resolutions_several_data"], result_format)
                worksheet.write(i, 9, department["citizen_resolutions_several_data_percent"] / 100,
                                percent_format_yellow)
                worksheet.write(i, 10, department["citizen_avg_percent"] / 100, percent_format_orange)
                worksheet.write(i, 11, department["citizen_total_points"], mark_format)
                i += 1
            # Печать последней строки ВСЕГО
            worksheet.set_row(i, 36)
            worksheet.write(i, 0, i - 4, result_format)
            worksheet.write(i, 1, "ВСЕГО", result_format)
            worksheet.write(i, 2, sum_data["main_resolutions_all"], result_format)
            worksheet.write(i, 3, sum_data["control_negative"], result_format)
            worksheet.write(i, 4, "-", percent_format_yellow)
            worksheet.write(i, 5, "-", percent_format_orange)
            worksheet.write(i, 6, "", mark_format)
            worksheet.write(i, 7, sum_data["citizen_resolutions_data"], result_format)
            worksheet.write(i, 8, sum_data["citizen_resolutions_several_data"], result_format)
            worksheet.write(i, 9, "-", percent_format_yellow)
            worksheet.write(i, 10, "-", percent_format_orange)
            worksheet.write(i, 11, "", mark_format)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'effectiveness_quality.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


##### Отчет "Сведения о документообороте органов местного самоуправления города Вологды"
@permission_required('deloreports.show_docflow_report')
def doc_flow_report(request):
    """
    Сведения о документообороте Администрации города Вологды
    """
    departments_data = {}
    summary_data = {}
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            doc_flow_report = DocFlowReportApi()
            result = doc_flow_report.get_report_data(from_date, due_date)
            departments_data = result["departments_data"]
            summary_data = result["summary_data"]
    context = {
        'result': departments_data if departments_data else None,
        'summary': summary_data if summary_data else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }

    return render(request, 'deloreports/doc_flow_report.html', context)


@permission_required('deloreports.show_paper_flow_report')
def paper_flow_report(request):
    result = []
    departments_data = {}
    summary_data = {}
    form = PaperFlowDateForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            paper_pack_cost = request.POST.get("paper_pack_cost")
            paper_flow_report = PaperFlowReportApi()
            result = paper_flow_report.get_report_data(from_date, due_date, paper_pack_cost)
            departments_data = result["departments_data"]
            summary_data = result["summary_data"]
    context = {
        'result': departments_data if departments_data else None,
        'summary': summary_data if summary_data else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }

    return render(request, 'deloreports/paper_flow_report.html', context)


@permission_required('deloreports.show_municipal_legal_act_registration_report', login_url='/deloreports/')
def municipal_legal_act_registration_report(request):
    """
    МПА, зарегистрированные за день
    """
    municipal_legal_act_registration_report_api = MunicipalLegalActRegistrationReportApi()
    dt = datetime.now()
    result = municipal_legal_act_registration_report_api.get_report_data(dt)
    context = {
        'report_data': result if result else 0,
        'datetime': dt.strftime('%d.%m.%Y %H:%M:%S')
    }
    return render(request, 'deloreports/municipal_legal_act_registration_report.html', context)


@permission_required('deloreports.show_prosecutors_reaction_act_report', login_url='/deloreports/')
def prosecutors_reaction_act_report(request):
    """
    Отчет о количестве актов прокурорского реагирования
    """
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            prosecutors_reaction_act_report_api = ProsecutorsReactionActReportApi()
            result = prosecutors_reaction_act_report_api.get_report_data(from_date, due_date)
    context = {
        'report_data': result["departments_data"] if result and result.get("departments_data") else None,
        'summary_data': result["summary_data"] if result and result.get("summary_data") else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/prosecutors_reaction_act_report.html', context)


@permission_required('deloreports.show_prosecutors_incoming_docs_report', login_url='/deloreports/')
def prosecutors_incoming_docs_report(request):
    """
    Отчет о количестве документов, поступивших из прокуратуры
    """
    # Ежемесячный отчет в прокуратуру
    result = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            prosecutors_incoming_docs_report_api = ProsecutorsIncomingDocsReportApi()
            result = prosecutors_incoming_docs_report_api.get_report_data(from_date, due_date)
    context = {
        'report_data': result["incoming_docs_by_groups_data"] if result and result.get(
            "incoming_docs_by_groups_data") else None,
        'incoming_letters_appearance': result["incoming_letters_appearance"] if result and result.get(
            "incoming_letters_appearance") else None,
        "incoming_docs_all": result["incoming_docs_all"] if result and result.get("incoming_docs_all") else None,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/prosecutors_incoming_docs_report.html', context)


@permission_required('deloreports.show_control_cases_report', login_url='/deloreports/')
def control_cases_report(request):
    """
    Формирование контрольных дел по поручениям Главы города Вологды
    """
    result = []
    control_cases_report = ControlCasesReportApi()
    result = control_cases_report.get_report_data()
    control_cases_count = 0

    if result != [] and result["report_data"]:
        for row in result["report_data"]:
            if row["control_cases_ufolders"]:
                for r in row["control_cases_ufolders"]:
                    control_cases_count = control_cases_count + 1

    context = {
        'report_data': result["report_data"] if result != [] and result["report_data"] else 0,
        'control_cases_count': control_cases_count
    }
    return render(request, 'deloreports/control_cases_report.html', context)


@permission_required('deloreports.show_control_cases_report2', login_url='/deloreports/')
def control_cases_report2(request):
    """
    Формирование контрольных дел по поручениям Главы города Вологды
    """
    # Отчет по контрольным делам Андрей
    result = []
    control_cases_report = ControlCasesReportApi()
    result = control_cases_report.get_report_data2()
    control_cases_count = 0

    if result != [] and result["report_data"]:
        for row in result["report_data"]:
            if row:
                control_cases_count = control_cases_count + 1

    context = {
        'report_data': result["report_data"] if result != [] and result["report_data"] else 0,
        'control_cases_count': control_cases_count
    }
    return render(request, 'deloreports/control_cases_report2.html', context)


@permission_required('deloreports.show_control_cases_report', login_url='/deloreports/')
def get_control_case_docs_list(request):
    """
    Выгрузка документов из модального окна отчёта Формирование контрольных дел по поручениям Главы города Вологды
    (control_cases_report)
    """
    result = []
    if request.method == 'POST':
        username = request.POST.get('username')
        control_case_name = request.POST.get('control_case_name')

        if username and control_case_name:
            username = username.strip(' \t\n\r')
            control_case_name = control_case_name.strip(' \t\n\r')
            control_cases_report = ControlCasesReportApi()
            rows = control_cases_report.get_control_cases_doc_count_by_ufolder_name(username, control_case_name)
            if rows:
                for row in rows:
                    t = [row[0], row[1].strftime('%d.%m.%Y'), row[2], row[3]]
                    result.append(t)

    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


@permission_required('deloreports.show_control_cases_report2', login_url='/deloreports/')
def get_control_case_docs_list2(request):
    result = []
    if request.method == 'POST':
        isn_doc = request.POST.get('isn_doc')
        if isn_doc:
            control_cases_report = ControlCasesReportApi()
            rows = control_cases_report.get_control_cases_doc_count_by_isndoc(isn_doc)
            if rows:
                for row in rows:
                    t = [row[0], row[1].strftime('%d.%m.%Y'), row[2], row[3]]
                    result.append(t)

    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


def download_control_case_docs_list(request):
    # print('вход в функцию download_control_case_docs_list')
    report = []
    result = []

    if request.method == 'GET':
        username = request.GET.get('username') if request.GET.get('username') else ""
        control_case_name = request.GET.get('control_case_name') if request.GET.get('control_case_name') else ""

        if username != "" and control_case_name != "":
            username = username.strip(' \t\n\r')
            control_case_name = control_case_name.strip(' \t\n\r')

            control_case_report = ControlCasesReportApi()
            rows = control_case_report.get_control_cases_doc_count_by_ufolder_name(username, control_case_name)
            if rows:
                for row in rows:
                    t = [row[0], row[1].strftime('%d.%m.%Y'), row[2]]
                    result.append(t)
            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            reportTableHeader = [
                "№",
                "Номер документа",
                "Дата регистрации",
                "Содержание",
            ]
            # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
            worksheet = workbook.add_worksheet("Отчет")

            worksheet.set_column('A:A', 10)
            worksheet.set_column('B:B', 18)
            worksheet.set_column('C:C', 18)
            worksheet.set_column('D:D', 60)

            # Форматирование шапки
            header = workbook.add_format({'bold': True})
            header.set_align('left')
            header.set_align('top')
            header.set_border(1)
            header.set_border_color("#000000")

            # Форматирование тела таблицы
            body = workbook.add_format()
            body.set_text_wrap()
            body.set_align('left')
            body.set_align('top')
            body.set_border(1)
            body.set_border_color("#000000")

            # вывод значений
            for i, row in enumerate(result):
                if i == 0:
                    worksheet.write(i, 0, reportTableHeader[0], header)
                    worksheet.write(i, 1, reportTableHeader[1], header)
                    worksheet.write(i, 2, reportTableHeader[2], header)
                    worksheet.write(i, 3, reportTableHeader[3], header)

                worksheet.write(i + 1, 0, i + 1, body)
                worksheet.write(i + 1, 1, str(row[0] if row[0] else "Не указан"), body)
                worksheet.write(i + 1, 2, str(row[1] if row[1] else "Не указан"), body)
                worksheet.write(i + 1, 3, str(row[2] if row[2] else "Не указана"), body)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'report.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


@permission_required('deloreports.show_constituency_report', login_url='/deloreports/')
def constituency(request):
    ogd_data_base = Counter()
    lot_data_base = Counter()
    zso_data_base = Counter()
    themes_base = Counter()
    river_district_counter = Counter()
    central_district_counter = Counter()
    east_district_counter = Counter()
    west_district_counter = Counter()
    empty_district_counter = Counter()
    new_themes = {}

    spheres = []
    themes_new = []
    spheres_new = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        # print "__________________________________________________________"

        if form.is_valid():
            from deloreports.functions.constituency import Constituency
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            data_obj = Constituency()
            data_base = data_obj.get_new_statistic(start_date=from_date, end_date=due_date)
            data_delta = data_obj.get_new_delta(start_date=from_date,
                                                end_date=due_date)  # delta_date=data_base["max_date"])

            # test = Counter()

            spheres = [sphere["name"]
                       for sphere in data_obj.sphere_list]

            # for data in data_base:
            #     # print type(data)
            #     # test.update(data["themes"])
            #     # print dat
            # print data_base

            for zam in spheres:
                new_themes[zam] = {
                    "themes": Counter(),
                    "river_district": Counter(),
                    "central_district": Counter(),
                    "east_district": Counter(),
                    "west_district": Counter(),
                    "empty_district": Counter(),
                }

            if data_base or data_delta:
                # print data_base, data_base
                for data in (data_base, data_delta):
                    ogd_data_base.update(data["ogd_data"])
                    lot_data_base.update(data["lot_data"])
                    zso_data_base.update(data["zso_data"])
                    themes_base.update(data["themes"])
                    river_district_counter.update(data["river_district"])
                    central_district_counter.update(data["central_district"])
                    east_district_counter.update(data["east_district"])
                    west_district_counter.update(data["west_district"])
                    empty_district_counter.update(data["empty_district"])
                    for zam in list(new_themes.keys()):
                        # print zam
                        # print new_themes[zam]["themes"]
                        # print data["new_themes"][zam]["themes"]
                        # 1/0
                        new_themes[zam]["themes"].update(data["new_themes"][zam]["themes"]),
                        new_themes[zam]["river_district"].update(data["new_themes"][zam]["river_district"]),
                        new_themes[zam]["central_district"].update(data["new_themes"][zam]["central_district"]),
                        new_themes[zam]["east_district"].update(data["new_themes"][zam]["east_district"]),
                        new_themes[zam]["west_district"].update(data["new_themes"][zam]["west_district"]),
                        new_themes[zam]["empty_district"].update(data["new_themes"][zam]["empty_district"]),

            # spheres = {
            #     n: Counter()
            #     for n in data_obj.get_sphere_list()}

    ogd_data = [
        {"num": k, "count": v}
        for k, v in list(ogd_data_base.items())
    ]
    ogd_data.sort(key=lambda item: item["count"], reverse=True)

    lot_data = [
        {"num": k, "count": v}
        for k, v in list(lot_data_base.items())
    ]

    lot_data.sort(key=lambda item: item["count"], reverse=True)

    zso_data = [
        {"name": k, "count": v}
        for k, v in list(zso_data_base.items())
    ]

    zso_data.sort(key=lambda item: item["count"], reverse=True)

    themes = [
        {
            "name": k,
            "count": v,
            "river_district": river_district_counter[k],
            "central_district": central_district_counter[k],
            "east_district": east_district_counter[k],
            "west_district": west_district_counter[k],
            "empty_district": empty_district_counter[k],
            "sphere": data_obj.get_sphere(k)
        }
        for k, v in list(themes_base.items())
    ]

    themes.sort(key=lambda item: item["count"], reverse=True)

    # themes_new = [
    #     {
    #         "name": k, 
    #         "count": v, 
    #         "river_district": river_district_counter[k],
    #         "central_district": central_district_counter[k],
    #         "east_district": east_district_counter[k],
    #         "west_district": west_district_counter[k],
    #         "empty_district": empty_district_counter[k],
    #         "sphere": data_obj.get_sphere(k)
    #     }
    #     for k, v in themes_base.items() 
    # ]

    # считаем данные по сферам(замам) - старый вариант#
    spheres_all = []
    for i, spher in enumerate(spheres):
        one_sphere = {
            "id": i,
            "name": spher,
            "count": 0,
            "river_district": 0,
            "central_district": 0,
            "east_district": 0,
            "west_district": 0,
            "empty_district": 0,
        }

        for theme in themes:
            if theme["sphere"] == spher:
                one_sphere["count"] = one_sphere["count"] + theme["count"]
                one_sphere["river_district"] = one_sphere["river_district"] + theme["river_district"]
                one_sphere["central_district"] = one_sphere["central_district"] + theme["central_district"]
                one_sphere["east_district"] = one_sphere["east_district"] + theme["east_district"]
                one_sphere["west_district"] = one_sphere["west_district"] + theme["west_district"]
                one_sphere["empty_district"] = one_sphere["empty_district"] + theme["empty_district"]

                # print one_sphere["count"], theme["count"]
        # print one_sphere["name"]
        if one_sphere["count"] != 0:
            spheres_all.append(one_sphere)
            # print one_sphere["name"]

    ## формируем 2 списка данных для новой версии отчета (сумма по замам и полный список всех комбинаций зам + тема с подсчетами)
    spheres_all_new = []
    for i, spher in enumerate(spheres):
        one_sphere = {
            "id": i,
            "name": spher,
            "count": 0,
            "river_district": 0,
            "central_district": 0,
            "east_district": 0,
            "west_district": 0,
            "empty_district": 0,
        }
        for theme in list(new_themes[spher]["themes"].keys()):
            # print 
            # if theme["sphere"] == spher:
            one_sphere["count"] = one_sphere["count"] + new_themes[spher]["themes"][theme]
            one_sphere["river_district"] = one_sphere["river_district"] + new_themes[spher]["river_district"][theme]
            one_sphere["central_district"] = one_sphere["central_district"] + new_themes[spher]["central_district"][
                theme]
            one_sphere["east_district"] = one_sphere["east_district"] + new_themes[spher]["east_district"][theme]
            one_sphere["west_district"] = one_sphere["west_district"] + new_themes[spher]["west_district"][theme]
            one_sphere["empty_district"] = one_sphere["empty_district"] + new_themes[spher]["empty_district"][theme]

            themes_new.append({
                "name": theme,
                "count": new_themes[spher]["themes"][theme],
                "river_district": new_themes[spher]["river_district"][theme],
                "central_district": new_themes[spher]["central_district"][theme],
                "east_district": new_themes[spher]["east_district"][theme],
                "west_district": new_themes[spher]["west_district"][theme],
                "empty_district": new_themes[spher]["empty_district"][theme],
                "sphere": spher,
            })

        if one_sphere["count"] != 0:
            spheres_all_new.append(one_sphere)

    themes_new.sort(key=lambda item: item["count"], reverse=True)

    return render(request, 'deloreports/constituency.html', {
        "ogd_data": ogd_data,
        "lot_data": lot_data,
        "zso_data": zso_data,
        "themes": themes,
        "form": form,
        "spheres": spheres_all,
        "themes_new": themes_new,
        "spheres_new": spheres_all_new,
        "from_date": from_date,
        "due_date": due_date,
    })


@permission_required('deloreports.show_SSTU_report', login_url='/deloreports/')
def sstu_export(request, download=False):
    form = DateRangeForm(request.POST)
    if request.method == 'POST':
        if form.is_valid():
            from deloreports.functions.sstu_export import Sstu_export
            from_date = form.cleaned_data["from_date"]  # {date} 2023-11-09
            due_date = form.cleaned_data["due_date"] + timedelta(days=1)
            start = Sstu_export()
            start.collector(start_date=from_date, end_date=due_date)
            return download_zip(start.big_file_temp)

    context = {
        'form': form,
    }
    return render(request, 'deloreports/sstu_export.html', context)


@permission_required('deloreports.show_SSTU_protocol', login_url='/deloreports/')
def sstu_get_protocol(request):
    """
    Выгрузка протокола ошибок при регистрации обращений граждан в ЕСЭД ОМСУ для портала ССТУ
    """
    result = []
    err_count = -1
    form = DateFormSimpleDep(request.POST)
    from_date = due_date = None
    if request.method == 'POST':
        if form.is_valid():
            from deloreports.functions.sstu_get_protocol import Sstu_export_protocol
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            dep_name = form.cleaned_data["dep_name"]
            start = Sstu_export_protocol()
            result, err_count = start.collector(start_date=from_date, end_date=due_date, dep_list=dep_name)
            # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    context = {
        'report_data': result,
        'form': form,
        'err_count': err_count,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/sstu_get_protocol.html', context)


def download_zip(cont_file):
    '''
    Отдача файла без загрузки в оперативную память.
    '''
    # cont_file.seek(0)
    file_size = cont_file.tell()
    if file_size == 0:
        return HttpResponseBadRequest('За указанный период данных нет')
    file_name = f"{datetime.now().strftime('%Y-%m-%m-%H-%M-%S')}.zip"
    # print file_size
    cont_file.seek(0)
    response = StreamingHttpResponse(cont_file, content_type='application/zip')
    response['Content-Disposition'] = f'attachment; filename={file_name}'
    response['Content-Length'] = file_size
    # print "tell_result", response['Content-Length']
    return response


def download_arc(file_zip):
    cont_file = file_zip.getvalue()
    report_filename = 'arc.zip'
    content_type = 'application/zip'
    response = HttpResponse(cont_file, content_type=content_type)
    response['Content-Length'] = len(cont_file)
    response['Content-Disposition'] = f'attachment; filename={report_filename}'

    return response


@permission_required('deloreports.show_tos_appeals_report', login_url='/deloreports/')
def tos_appeals_report(request):
    """
    Обращения граждан в разрезе территориальных общественных самоуправлений
    """

    result = []
    tos_list = []
    from_date = due_date = None
    form = DateFormWithTOSChoices(request.POST)
    if request.method == 'POST':
        from deloreports.functions.tos_appeals_report_api import TosAppealsReportApi
        tos_appeals_report = TosAppealsReportApi()
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            tos_list = form.cleaned_data["tos"]
            str_tos_list = str(tos_list).replace(' ', '')
            cache_key_tos_appeals = get_cache_key('tos_appeals', request.user.id, from_date, due_date, str_tos_list)
            cached_results = cache.get(cache_key_tos_appeals)

            if cached_results:
                result = cached_results
            else:
                result = sorted(tos_appeals_report.get_report_data(from_date, due_date, tos_list),
                                key=lambda x: x['tos'])
                cache.set(cache_key_tos_appeals, result, 60 * 5)

    context = {
        'report_data': result,
        'form': form,
        'tos_list': tos_list,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/tos_appeals_report.html', context)


def tos_appeals_export_excel(request):
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    tos_str = request.GET.get("tos_list", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")
    if tos_str == "[]":
        tos_list = []
    else:
        import ast
        tos_list = ast.literal_eval(tos_str)
    str_tos_list = str(tos_list).replace(' ', '')
    cache_key_tos_appeals = get_cache_key('tos_appeals', request.user.id, from_date, due_date, str_tos_list)
    cached_results = cache.get(cache_key_tos_appeals)
    if cached_results:
        result = cached_results
    else:
        from deloreports.functions.tos_appeals_report_api import TosAppealsReportApi
        tos_appeals_report = TosAppealsReportApi()
        result = sorted(tos_appeals_report.get_report_data(from_date, due_date, tos_list),
                        key=lambda x: x['tos'])
        cache.set(cache_key_tos_appeals, result, 60 * 5)

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            worksheet = workbook.add_worksheet("Обращения граждан")
            worksheet.fit_to_pages(1, 1)
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.2, right=0.2, top=0.4, bottom=0.4)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 20, 'text_wrap': True})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True, 'fg_color': '#bfbfbf'})
            table_format = workbook.add_format(
                {'align': 'left', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True})
            table_format_time = workbook.add_format(
                {'align': 'left', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True, 'num_format': 'dd.mm.yyyy'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            table_format.set_align('top')
            table_format_time.set_align('top')

            # Установка ширины столбцов
            worksheet.set_column('A:A', 7)
            worksheet.set_column('B:B', 10)
            worksheet.set_column('C:C', 13.14)
            worksheet.set_column('D:D', 22.14)
            worksheet.set_column('E:E', 45)
            worksheet.set_column('F:F', 30)
            worksheet.set_column('G:G', 30)
            worksheet.set_column('H:H', 30)
            worksheet.set_column('I:I', 20)

            # Установка высоты строк
            worksheet.set_row(0, 74.25)
            worksheet.set_row(1, 58.5)

            # Шапка таблицы
            worksheet.merge_range(
                'A1:I1',
                f"Обращения граждан в разрезе территориальных общественных самоуправлений, "
                f"период: {get_period_string(from_date, due_date)} по состоянию на "
                f"{str(datetime.now().strftime('%d.%m.%Y'))}", header_format
            )

            worksheet.write('A2', '№ п/п', header_table_format)
            worksheet.write('B2', 'Рег. №', header_table_format)
            worksheet.write('C2', 'Рег. дата', header_table_format)
            worksheet.write('D2', 'Заявитель', header_table_format)
            worksheet.write('E2', 'Содержание обращения', header_table_format)
            worksheet.write('F2', 'Тема/Рубрика', header_table_format)
            worksheet.write('G2', 'Группа документов', header_table_format)
            worksheet.write('H2', 'Адрес проблемы', header_table_format)
            worksheet.write('I2', 'ТОС', header_table_format)

            i = 2
            for appeal in result:
                # Установка высоты для строки данных
                # worksheet.set_row(i, 33)
                worksheet.write(i, 0, i - 1, table_format)
                worksheet.write(i, 1, appeal['regnum'], table_format)
                worksheet.write(i, 2, appeal['regdate'], table_format_time)
                worksheet.write(i, 3, appeal['citizen'], table_format)
                worksheet.write(i, 4, appeal['annotat'], table_format)
                worksheet.write(i, 5, appeal['rubr'], table_format_time)
                worksheet.write(i, 6, appeal['docgroup'], table_format_time)
                worksheet.write(i, 7, appeal['address'], table_format)
                worksheet.write(i, 8, appeal['tos'], table_format)
                i += 1

        report = report_file.getvalue()

    report_filename = f'tos_appeals_report_{from_date.strftime("%Y.%m.%d")}-{due_date.strftime("%Y.%m.%d")}.xlsx'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


@permission_required('deloreports.show_municipal_legal_acts_consideration_report', login_url='/deloreports/')
def municipal_legal_acts_consideration_report(request):
    """
    Информация о сроках и качестве рассмотрения проектов муниципальных правовых актов в органах Администрации города Вологды
    """
    municipal_legal_acts_consideration_report = MunicipalLegalActsConsiderationReportApi()
    result = []
    dep_name_list = []
    qual_data_list = []
    from_date = due_date = None
    form = DateRangeForm(request.POST)
    if request.method == 'POST':

        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            if from_date <= due_date:
                result, dep_name_list, qual_data_list = municipal_legal_acts_consideration_report.get_report_data(
                    from_date, due_date)
            else:
                return HttpResponse("Дата начала больше даты окончания")
                # return HttpResponseBadRequest()

    form = DateFormWithTOSChoices(request.POST)
    context = {
        'report_data': result,
        'dep_name_list': dep_name_list,
        'qual_data_list': qual_data_list,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
    }
    return render(request, 'deloreports/municipal_legal_acts_consideration_report.html', context)


def municipal_legal_acts_consideration_report_export_excel(request):
    municipal_legal_acts_consideration_report = MunicipalLegalActsConsiderationReportApi()
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")

    result, dep_name_list, qual_data_list = municipal_legal_acts_consideration_report.get_report_data(from_date,
                                                                                                      due_date)
    count_dep = len(dep_name_list)
    count_row = len(result)

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            # **************************Печать таблицы 1 на 1 листе книги Excel**********************
            worksheet = workbook.add_worksheet("Сроки согласования")
            worksheet.fit_to_pages(1, 1)
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.70, right=0.70, top=0.75, bottom=0.75)  # поля в коде в дюймах, в excel в см

            header_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': False, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 16})
            header_table_format = workbook.add_format(
                {'align': 'center', 'valign': 'bottom', 'bold': True, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1, 'text_wrap': True,
                 'fg_color': '#d9d9d9'})
            result_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': False, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1})
            result_sum_itog = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1, 'fg_color': '#d9d9d9'})
            result_sum_num = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 14, 'border': 1, 'num_format': '0.00',
                 'fg_color': '#d9d9d9'})

            # header_format.set_align('vcenter')
            num_col = 0
            # Установкановка ширины столбцов
            for dep in dep_name_list:
                if dep != 'Зам. Мэра по соц. вопросам':
                    worksheet.set_column(num_col, num_col, 10.43)
                    worksheet.set_column(num_col + 1, num_col + 1, 10)
                    num_col += 2
                else:
                    worksheet.set_column(num_col, num_col, 11)
                    worksheet.set_column(num_col + 1, num_col + 1, 12)
                    num_col += 2

            # Установка высоты строк 
            worksheet.set_row(0, 90.75)
            worksheet.set_row(1, 83.25)
            worksheet.set_row(2, 117)
            num_row = 3  # вывод данных идет с 4 строки
            for i in range(3, count_row + 3):
                worksheet.set_row(i, 18.75)

                # Шапка таблицы
            worksheet.merge_range(0, 0, 0, count_dep * 2 - 1,
                                  'Таблица сроков согласования проектов муниципальных правовых актов в органах Администрации города Вологды ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)
            num_col = 0
            for dep in dep_name_list:
                worksheet.merge_range(1, num_col, 1, num_col + 1, str(dep), header_table_format)
                worksheet.write(2, num_col, 'Кол-во дней на согласовании', header_table_format)
                worksheet.write(2, num_col + 1, 'Кол-во документов с учетом всех версий', header_table_format)
                num_col += 2

            num_row = 3

            for input_str in result:
                num_col = 0
                for i in range(count_dep):
                    if input_str[i].days == "Вес:":
                        worksheet.write(num_row, num_col, input_str[i].days, result_sum_itog)
                        worksheet.write(num_row, num_col + 1, input_str[i].mark, result_sum_num)
                    elif input_str[i].days >= 0:
                        worksheet.write(num_row, num_col, input_str[i].days, result_format)
                        worksheet.write(num_row, num_col + 1, input_str[i].mark, result_format)

                    num_col += 2
                num_row += 1
            # **************************Печать таблицы 2 на 2 листе книги Excel****************************************
            worksheet2 = workbook.add_worksheet("Качественные критерии")

            worksheet2.fit_to_pages(1, 1)
            worksheet2.set_landscape()  # установка альбомной ориентации страницы
            worksheet2.set_margins(left=0.70, right=0.70, top=0.75, bottom=0.75)  # поля в коде в дюймах, в excel в см

            header_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 14, 'text_wrap': True})
            header_table_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1, 'text_wrap': True,
                 'fg_color': '#d9d9d9'})
            result_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': False, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1, 'text_wrap': True})
            result_oagv_format = workbook.add_format(
                {'align': 'center', 'valign': 'vcenter', 'bold': False, 'font_color': 'black',
                 'font_name': 'Times New Roman', 'font_size': 13, 'border': 1, 'fg_color': '#d9d9d9',
                 'text_wrap': True})

            # Установкановка ширины столбцов
            worksheet2.set_column(0, 0, 27.57)
            worksheet2.set_column(1, 1, 26.86)
            worksheet2.set_column(2, 2, 33.86)
            worksheet2.set_column(3, 3, 22.86)
            worksheet2.set_column(4, 4, 21.57)
            worksheet2.set_column(5, 5, 20.29)
            worksheet2.set_column(6, 6, 24.86)

            # Установка высоты строк 
            worksheet2.set_row(0, 83.25)
            worksheet2.set_row(1, 87)
            worksheet2.set_row(2, 156)
            # worksheet2.set_row(3, 16.5)

            # Шапка таблицы
            worksheet2.merge_range(0, 0, 0, 6,
                                   'Качество подготовки и рассмотрения проектов муниципальных правовых актов в Администрации города Вологды ' + get_period_string(
                                       from_date, due_date) + ' по состоянию на ' + str(
                                       datetime.now().strftime('%d.%m.%Y')), header_format)
            worksheet2.merge_range(1, 0, 2, 0, 'Органы Администрации города Вологды', header_table_format)
            worksheet2.merge_range(1, 1, 1, 2,
                                   'Качество рассмотрения проектов МПА (органы Администрации рассматриваются в качестве визирующих проектов)',
                                   header_table_format)
            worksheet2.merge_range(1, 3, 1, 6,
                                   'Качество подготовки проектов МПА, разработанных в отчетном периоде и утвержденных (органы Администрации рассматриваются в качестве исполнителей проектов)',
                                   header_table_format)
            worksheet2.write(2, 1, 'Количество проектов МПА, рассмотренных органом АГВ, всего', header_table_format)
            worksheet2.write(2, 2,
                             'Количество проектов МПА с  отметкой "Не согласовано"  или  "Согласовано с замечаниями" при условии полного отсутствия комментариев от органа АГВ (ни один работник органа АГВ не указал в проекте МПА комментарий)',
                             header_table_format)
            worksheet2.write(2, 3, 'Количество проектов МПА, разработанных в отчетном периоде и утвержденных',
                             header_table_format)
            worksheet2.write(2, 4, 'Среднее число версий МПА', header_table_format)
            worksheet2.write(2, 5, 'Количество проектов МПА с числом версий более 3', header_table_format)
            worksheet2.write(2, 6, 'Количество проектов МПА с общей продолжительностью согласованиия более 60 дней',
                             header_table_format)
            num_row = 3  # вывод данных идет с 4 строки
            for dep in qual_data_list:
                if dep.dep == "Итого:":
                    not_agreed = float(dep.not_agreed)
                    versions_more_3 = float(dep.versions_more_3)
                    considered_more_60_days = float(dep.considered_more_60_days)
                    considered_prj = float(dep.considered_prj)
                    created_prj = float(dep.created_prj)
                    if not_agreed == 0:
                        not_agreed = "0 (0,00%)"
                    else:
                        not_agreed = str(int(not_agreed)) + " (" + str(
                            round(not_agreed / considered_prj * 100, 2)) + "%)"
                    if versions_more_3 == 0:
                        versions_more_3 = "0 (0,00%)"
                    else:
                        versions_more_3 = str(int(versions_more_3)) + " (" + str(
                            round(versions_more_3 / created_prj * 100, 2)) + "%)"
                    if considered_more_60_days == 0:
                        considered_more_60_days = "0 (0,00%)"
                    else:
                        considered_more_60_days = str(int(considered_more_60_days)) + " (" + str(
                            round(considered_more_60_days / created_prj * 100, 2)) + "%)"
                    worksheet2.write(num_row, 0, dep.dep, header_table_format)
                    worksheet2.write(num_row, 1, dep.considered_prj, header_table_format)
                    worksheet2.write(num_row, 2, not_agreed, header_table_format)
                    worksheet2.write(num_row, 3, dep.created_prj, header_table_format)
                    worksheet2.write(num_row, 4, dep.quantity_of_versions, header_table_format)
                    worksheet2.write(num_row, 5, versions_more_3, header_table_format)
                    worksheet2.write(num_row, 6, considered_more_60_days, header_table_format)
                else:
                    not_agreed = float(dep.not_agreed)
                    versions_more_3 = float(dep.versions_more_3)
                    considered_more_60_days = float(dep.considered_more_60_days)
                    considered_prj = float(dep.considered_prj)
                    created_prj = float(dep.created_prj)
                    if not_agreed == 0:
                        not_agreed = "0 (0,00%)"
                    else:
                        not_agreed = str(int(not_agreed)) + " (" + str(
                            round(not_agreed / considered_prj * 100, 2)) + "%)"
                    if versions_more_3 == 0:
                        versions_more_3 = "0 (0,00%)"
                    else:
                        versions_more_3 = str(int(versions_more_3)) + " (" + str(
                            round(versions_more_3 / created_prj * 100, 2)) + "%)"
                    if considered_more_60_days == 0:
                        considered_more_60_days = "0 (0,00%)"
                    else:
                        considered_more_60_days = str(int(considered_more_60_days)) + " (" + str(
                            round(considered_more_60_days / created_prj * 100, 2)) + "%)"
                    worksheet2.write(num_row, 0, dep.dep, result_oagv_format)
                    worksheet2.write(num_row, 1, dep.considered_prj, result_format)
                    worksheet2.write(num_row, 2, not_agreed, result_format)
                    worksheet2.write(num_row, 3, dep.created_prj, result_format)
                    worksheet2.write(num_row, 4, dep.quantity_of_versions, result_format)
                    worksheet2.write(num_row, 5, versions_more_3, result_format)
                    worksheet2.write(num_row, 6, considered_more_60_days, result_format)
                num_row += 1

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'AnalysisOfLegalActs.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


# отчет "Аналитика по проектам МПА", получение списка РКПД для последующего отображения в модальном окне
def get_docs_list_municipal_legal_acts_consideration_report(request):
    result = []
    if request.method == 'POST':
        due_dep = request.POST.get('due')
        from_date = datetime.strptime(request.POST.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.POST.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.POST.get("reportName").strip(' \t\n\r') if request.POST.get("reportName") else None
        days_count = int(request.POST.get("days_count") if request.POST.get("days_count") else 0)
        if due_dep:
            municipal_legal_acts_consideration_report = MunicipalLegalActsConsiderationReportApi()
            due_dep = due_dep.strip(' \t\n\r')
            if report_name != None:
                if report_name == "get_prj_not_agreed_modal":
                    rows = municipal_legal_acts_consideration_report.get_prj_not_agreed_modal(due_dep, from_date,
                                                                                              due_date)
                elif report_name == "get_versions_more_3_modal":
                    rows = municipal_legal_acts_consideration_report.get_versions_more_3_modal(due_dep, from_date,
                                                                                               due_date)
                elif report_name == "get_considered_more_60_days_modal":
                    rows = municipal_legal_acts_consideration_report.get_considered_more_60_days_modal(due_dep,
                                                                                                       from_date,
                                                                                                       due_date)
                elif report_name == "get_considered_prj_modal":
                    rows = municipal_legal_acts_consideration_report.get_considered_prj_modal(due_dep, from_date,
                                                                                              due_date)
                elif report_name == "get_count_prj_modal_on_days":
                    rows = municipal_legal_acts_consideration_report.get_count_prj_modal_on_days(due_dep, from_date,
                                                                                                 due_date, days_count)
                elif report_name == "get_created_prj":
                    rows = municipal_legal_acts_consideration_report.get_created_prj_modal(due_dep, from_date, due_date)

            if rows:
                for row in rows:
                    t = []
                    for val in row:
                        if type(val) == datetime:
                            t.append(val.strftime('%d.%m.%Y'))
                        else:
                            t.append(val)
                    result.append(t)
            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )

    return HttpResponse(
        json.dumps({"data": result}),
        content_type="application/json"
    )


# отчет "Аналитика по проектам МПА", экспорт перечня поручений из модального окна в Эксель
def download_list_municipal_legal_acts_consideration_report(request):
    # print('вход в функцию download_list_municipal_legal_acts_consideration_report')
    report = []
    result = []
    if request.method == 'GET':
        due_dep = request.GET.get('due') if request.GET.get('due') else ""
        from_date = datetime.strptime(request.GET.get('fromDate', "1970-01-01")[:10], "%Y-%m-%d")
        due_date = datetime.strptime(request.GET.get('dueDate', "2200-01-01")[:10], "%Y-%m-%d")
        report_name = request.GET.get("reportName").strip(' \t\n\r') if request.GET.get("reportName") else None
        days_count = int(request.GET.get("days_count") if request.GET.get("days_count") else 0)

        if due_dep:
            due_dep = due_dep.strip(' \t\n\r')
            municipal_legal_acts_consideration_report = MunicipalLegalActsConsiderationReportApi()
            if report_name != None:
                if report_name == "get_prj_not_agreed_modal":
                    rows = municipal_legal_acts_consideration_report.get_prj_not_agreed_modal(due_dep,
                                                                                              from_date,
                                                                                              due_date)
                elif report_name == "get_versions_more_3_modal":
                    rows = municipal_legal_acts_consideration_report.get_versions_more_3_modal(due_dep,
                                                                                               from_date,
                                                                                               due_date)
                elif report_name == "get_considered_more_60_days_modal":
                    rows = municipal_legal_acts_consideration_report.get_considered_more_60_days_modal(due_dep,
                                                                                                       from_date,
                                                                                                       due_date)
                elif report_name == "get_considered_prj_modal":
                    rows = municipal_legal_acts_consideration_report.get_considered_prj_modal(due_dep,
                                                                                              from_date,
                                                                                              due_date)
                elif report_name == "get_count_prj_modal_on_days":
                    rows = municipal_legal_acts_consideration_report.get_count_prj_modal_on_days(due_dep,
                                                                                                 from_date,
                                                                                                 due_date,
                                                                                                 days_count)
                elif report_name == "get_created_prj":
                    rows = municipal_legal_acts_consideration_report.get_created_prj_modal(due_dep, from_date, due_date)
            if rows:
                for row in rows:
                    t = []
                    for val in row:
                        if type(val) == datetime:
                            t.append(val.strftime('%d.%m.%Y'))
                        else:
                            t.append(val)
                    result.append(t)
            else:
                return HttpResponse(
                    json.dumps({"error": "Неверно введены данные, или нарушено сетевое соединение"}),
                    content_type="application/json"
                )

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            if report_name == "get_prj_not_agreed_modal":
                reportTableHeader = [
                    "№",
                    "Рег.номер РКПД",
                    "Дата регистрации РКПД",
                    "Содержание",
                    "Визирующий",
                    "Тип визы",
                    "Дата визирования"
                ]
                # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
                worksheet = workbook.add_worksheet("Отчет")

                worksheet.set_column('A:A', 6)
                worksheet.set_column('B:B', 23)
                worksheet.set_column('C:C', 23)
                worksheet.set_column('D:D', 54)
                worksheet.set_column('E:E', 43)
                worksheet.set_column('F:F', 32)
                worksheet.set_column('G:G', 23)

                # Форматирование шапки
                header = workbook.add_format({'bold': True})
                header.set_align('left')
                header.set_align('top')
                header.set_border(1)
                header.set_border_color("#000000")

                # Форматирование тела таблицы
                body = workbook.add_format()
                body.set_text_wrap()
                body.set_align('left')
                body.set_align('top')
                body.set_border(1)
                body.set_border_color("#000000")

                # вывод значений
                for i, row in enumerate(result):
                    if i == 0:
                        worksheet.write(i, 0, reportTableHeader[0], header)
                        worksheet.write(i, 1, reportTableHeader[1], header)
                        worksheet.write(i, 2, reportTableHeader[2], header)
                        worksheet.write(i, 3, reportTableHeader[3], header)
                        worksheet.write(i, 4, reportTableHeader[4], header)
                        worksheet.write(i, 5, reportTableHeader[5], header)
                        worksheet.write(i, 6, reportTableHeader[6], header)

                    worksheet.write(i + 1, 0, i + 1, body)
                    worksheet.write(i + 1, 1, str(row[1] if row[1] else "Не указан"), body)
                    worksheet.write(i + 1, 2, str(row[2] if row[2] else "Не указана"), body)
                    worksheet.write(i + 1, 3, str(row[3] if row[3] else "Не указано"), body)
                    worksheet.write(i + 1, 4, str(row[4] if row[4] else "Не указан"), body)
                    worksheet.write(i + 1, 5, str(row[5] if row[5] else "Не указана"), body)
                    worksheet.write(i + 1, 6, str(row[6] if row[6] else "Не указана"), body)

            elif report_name == "get_versions_more_3_modal" or report_name == "get_considered_more_60_days_modal" or report_name == "get_considered_prj_modal" or report_name == "get_created_prj":
                reportTableHeader = [
                    "№",
                    "Рег.номер РКПД",
                    "Дата регистрации РКПД",
                    "Содержание",
                    "Исполнитель РКПД"
                ]
                # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
                worksheet = workbook.add_worksheet("Отчет")

                worksheet.set_column('A:A', 6)
                worksheet.set_column('B:B', 23)
                worksheet.set_column('C:C', 23)
                worksheet.set_column('D:D', 54)
                worksheet.set_column('E:E', 43)

                # Форматирование шапки
                header = workbook.add_format({'bold': True})
                header.set_align('left')
                header.set_align('top')
                header.set_border(1)
                header.set_border_color("#000000")

                # Форматирование тела таблицы
                body = workbook.add_format()
                body.set_text_wrap()
                body.set_align('left')
                body.set_align('top')
                body.set_border(1)
                body.set_border_color("#000000")

                # вывод значений
                for i, row in enumerate(result):
                    if i == 0:
                        worksheet.write(i, 0, reportTableHeader[0], header)
                        worksheet.write(i, 1, reportTableHeader[1], header)
                        worksheet.write(i, 2, reportTableHeader[2], header)
                        worksheet.write(i, 3, reportTableHeader[3], header)
                        worksheet.write(i, 4, reportTableHeader[4], header)

                    worksheet.write(i + 1, 0, i + 1, body)
                    worksheet.write(i + 1, 1, str(row[1] if row[1] else "Не указан"), body)
                    worksheet.write(i + 1, 2, str(row[2] if row[2] else "Не указана"), body)
                    worksheet.write(i + 1, 3, str(row[3] if row[3] else "Не указано"), body)
                    worksheet.write(i + 1, 4, str(row[4] if row[4] else "Не указан"), body)

            elif report_name == "get_count_prj_modal_on_days":
                reportTableHeader = [
                    "№",
                    "Рег.номер РКПД",
                    "Дата регистрации РКПД",
                    "Содержание",
                    "Срок визирования (дни)",
                    "Дата визирования",
                    "Дата направления на визирование",
                    "Визирующий",
                ]
                # Наполняем workbook, см. документацию https://xlsxwriter.readthedocs.org/
                worksheet = workbook.add_worksheet("Отчет")

                worksheet.set_column('A:A', 6)
                worksheet.set_column('B:B', 23)
                worksheet.set_column('C:C', 23)
                worksheet.set_column('D:D', 54)
                worksheet.set_column('E:E', 43)
                worksheet.set_column('F:F', 43)
                worksheet.set_column('G:G', 43)
                worksheet.set_column('H:H', 54)

                # Форматирование шапки
                header = workbook.add_format({'bold': True})
                header.set_align('left')
                header.set_align('top')
                header.set_border(1)
                header.set_border_color("#000000")

                # Форматирование тела таблицы
                body = workbook.add_format()
                body.set_text_wrap()
                body.set_align('left')
                body.set_align('top')
                body.set_border(1)
                body.set_border_color("#000000")

                # вывод значений
                for i, row in enumerate(result):
                    if i == 0:
                        worksheet.write(i, 0, reportTableHeader[0], header)
                        worksheet.write(i, 1, reportTableHeader[1], header)
                        worksheet.write(i, 2, reportTableHeader[2], header)
                        worksheet.write(i, 3, reportTableHeader[3], header)
                        worksheet.write(i, 4, reportTableHeader[4], header)
                        worksheet.write(i, 5, reportTableHeader[5], header)
                        worksheet.write(i, 6, reportTableHeader[6], header)
                        worksheet.write(i, 7, reportTableHeader[7], header)
                    worksheet.write(i + 1, 0, i + 1, body)
                    worksheet.write(i + 1, 1, str(row[0] if row[0] else "Не указан"), body)
                    worksheet.write(i + 1, 2, str(row[1] if row[1] else "Не указана"), body)
                    worksheet.write(i + 1, 3, str(row[2] if row[2] else "Не указано"), body)
                    worksheet.write(i + 1, 4, str(row[3] if row[3] else "Не указан"), body)
                    worksheet.write(i + 1, 5, str(row[4] if row[4] else "Не указана"), body)
                    worksheet.write(i + 1, 6, str(row[5] if row[5] else "Не указана"), body)
                    worksheet.write(i + 1, 7, str(row[6] if row[6] else "Не указан"), body)

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'report.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


@permission_required('deloreports.show_appeals_two_week_period', login_url='/deloreports/')
def appeals_two_week_period(request, download=False):
    result = []
    from_date = due_date = dep_list = None
    form = DateFormSimpleDep(request.POST)
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            dep_list = form.cleaned_data["dep_name"]
            appeals_two_week_period = AppealsTwoWeekPeriodApi()
            result = appeals_two_week_period.get_report_data(from_date, due_date, dep_list)

    context = {
        'report_data': result,
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
        'dep_list': dep_list
    }

    return render(request, 'deloreports/appeals_two_week_period.html', context)


# экспорт в Эксель
def appeals_two_week_period_export_excel(request):
    from_date = request.GET.get("from_date", None)
    due_date = request.GET.get("due_date", None)
    dep_str = request.GET.get("dep_list", None)
    if not from_date or not due_date:
        return HttpResponseBadRequest()
    from_date = datetime.strptime(from_date, "%Y-%m-%d")
    due_date = datetime.strptime(due_date, "%Y-%m-%d")
    if dep_str == "[]":
        dep_list = []
    else:
        dep_list = dep_str[1:len(dep_str) - 1].split(",")

    appeals_two_week_period = AppealsTwoWeekPeriodApi()
    result = appeals_two_week_period.get_report_data(from_date, due_date, dep_list)

    with closing(BytesIO()) as report_file:
        with closing(xlsxwriter.Workbook(report_file)) as workbook:
            worksheet = workbook.add_worksheet("Обращения граждан")
            worksheet.fit_to_pages(1, 1)
            worksheet.set_landscape()  # установка альбомной ориентации страницы
            worksheet.set_margins(left=0.2, right=0.2, top=0.4, bottom=0.4)  # в коде в дюймах, в excel в см
            header_format = workbook.add_format(
                {'align': 'center', 'bold': True, 'font_color': 'black', 'font_name': 'Times New Roman',
                 'font_size': 20, 'text_wrap': True})
            header_table_format = workbook.add_format(
                {'align': 'center', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True, 'fg_color': '#bfbfbf'})
            table_format = workbook.add_format(
                {'align': 'left', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True})
            table_format_time = workbook.add_format(
                {'align': 'left', 'font_color': 'black', 'font_name': 'Times New Roman', 'font_size': 15, 'border': 1,
                 'text_wrap': True, 'num_format': 'dd.mm.yyyy'})

            header_format.set_align('vcenter')
            header_table_format.set_align('vcenter')
            table_format.set_align('top')
            table_format_time.set_align('top')

            # Установка ширины столбцов
            worksheet.set_column('A:A', 5.43)
            worksheet.set_column('B:B', 10)
            worksheet.set_column('C:C', 13.14)
            worksheet.set_column('D:D', 22.14)
            worksheet.set_column('E:E', 56.43)
            worksheet.set_column('F:F', 13.57)
            worksheet.set_column('G:G', 13.57)
            worksheet.set_column('H:H', 31.29)
            worksheet.set_column('I:I', 26.71)
            worksheet.set_column('J:J', 25)

            # Установка высоты строк 
            worksheet.set_row(0, 74.25)
            worksheet.set_row(1, 58.5)

            # Шапка таблицы
            worksheet.merge_range('A1:J1',
                                  'Обращения граждан, поступившие в Администрацию города Вологды на имя Мэра города Вологды, срок исполнения которых составляет не более 14 дней, ' + get_period_string(
                                      from_date, due_date) + ' по состоянию на ' + str(
                                      datetime.now().strftime('%d.%m.%Y')), header_format)

            worksheet.write('A2', '№ п/п', header_table_format)
            worksheet.write('B2', 'Рег. №', header_table_format)
            worksheet.write('C2', 'Дата регистрации', header_table_format)
            worksheet.write('D2', 'ФИО заявителя', header_table_format)
            worksheet.write('E2', 'Содержание', header_table_format)
            worksheet.write('F2', 'План', header_table_format)
            worksheet.write('G2', 'Факт', header_table_format)
            worksheet.write('H2', 'Отв. исполнитель', header_table_format)
            worksheet.write('I2', 'Ответ заявителю', header_table_format)
            worksheet.write('J2', 'Отчет ответственного исполнителя', header_table_format)
            i = 2
            for app in result:
                # Установка высоты для строки данных
                # worksheet.set_row(i, 33)  
                worksheet.write(i, 0, i - 1, table_format)
                worksheet.write(i, 1, app.doc_num, table_format)
                worksheet.write(i, 2, app.doc_date, table_format_time)
                worksheet.write(i, 3, app.doc_cit, table_format)
                worksheet.write(i, 4, app.doc_annotat, table_format)
                worksheet.write(i, 5, app.res_plan, table_format_time)
                if app.res_fact:
                    worksheet.write(i, 6, app.res_fact, table_format_time)
                else:
                    worksheet.write(i, 6, "не исполнено", table_format_time)
                worksheet.write(i, 7, app.res_exec, table_format)
                worksheet.write(i, 8, app.link_text, table_format)
                worksheet.write(i, 9, app.res_reptext, table_format)

                i += 1

        report = report_file.getvalue()

    # 2. ВЫДАЧА ОТЧЕТА КЛИЕНТУ...

    # Имя файла отчета, предложенное пользователю при сохранении отчета.
    # Желательно, чтобы имя не включало ничего «экзотического», в т.ч. русских букв.
    report_filename = 'citizensappeals.xlsx'

    # В примере указан MIME type для файлов xlsx; для xls должен быть 'application/ms-excel'.
    # См. также mimetypes.guess_type из стандартной библиотеки Python:
    #    https://docs.python.org/2/library/mimetypes.html#mimetypes.guess_type
    # content_type = 'application/ms-excel'
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    response = HttpResponse(report, content_type=content_type)
    response['Content-Length'] = len(report)  # Полезно, но не обязательно.
    response['Content-Disposition'] = 'attachment; filename=%s' % report_filename

    return response


##### Проверка по делопроизводству #####
@permission_required('deloreports.show_check_documents', login_url='/deloreports/')
def check_documents(request):
    """
    Проверка состояния делопроизводства в органах Администрации города Вологды
    Формирует таблицы с незаполненными/некорректно заполненными реквизитами
    """
    result = {
        "incoming_documents_without_correspondent_registration_data": [],
        'incoming_documents_incorrect_rubric': [],
        'incoming_documents_without_files': [],
        'outgoing_documents_incorrect_rubric': [],
        'outgoing_documents_no_links': [],
        'outgoing_documents_without_files': [],
        'internal_documents_incorrect_rubric': [],
        'internal_documents_no_eds': [],
        'citizens_appeals_incorrect_rubric': []
    }
    from_date = due_date = dep_name = None
    form = DateFormSimpleOneDep(request.POST)
    if request.method == 'POST':
        check_documents = CheckDocumentsApi()
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            due_date = form.cleaned_data["due_date"]
            dep_name = form.cleaned_data["dep_name"]
            # dep = Department.objects.get(due=dep_name)
            # print ('dep from check doc', dep)
            # print ('oda_ud', oda_ud)
            # if not request.user.has_perm("deloreports.get_report_for", dep):
            #     return HttpResponseBadRequest("Нет права получать отчет по подразделению")
            if from_date <= due_date:
                cache_key = get_check_documents_cache_key(request.user.id, from_date, due_date, dep_name)
                cached_results = cache.get(cache_key)  # получение данных из хэша

                if cached_results:
                    result = cached_results  # распаковка данных хэша
                else:  # cache = none
                    result = check_documents.get_report_data(from_date, due_date, dep_name)
                    cache.set(cache_key, result, 60 * 5)  # передаю данные в хэш
            else:
                return HttpResponseBadRequest("Некорректно задан отчетный период")

    context = {
        'report_data': result,
        'report_data_incoming_documents_without_correspondent_registration_data': result[
            "incoming_documents_without_correspondent_registration_data"],
        'incoming_documents_incorrect_rubric': result["incoming_documents_incorrect_rubric"],
        'incoming_documents_without_files': result["incoming_documents_without_files"],
        'outgoing_documents_incorrect_rubric': result["outgoing_documents_incorrect_rubric"],
        'outgoing_documents_no_links': result["outgoing_documents_no_links"],
        'outgoing_documents_without_files': result["outgoing_documents_without_files"],
        'internal_documents_incorrect_rubric': result["internal_documents_incorrect_rubric"],
        'internal_documents_no_eds': result["internal_documents_no_eds"],
        'citizens_appeals_incorrect_rubric': result["citizens_appeals_incorrect_rubric"],
        'form': form,
        'from_date': from_date,
        'due_date': due_date,
        'dep_name': dep_name,
    }
    return render(request, 'deloreports/check_documents.html', context)


@permission_required('deloreports.show_check_documents', login_url='/deloreports/')
def check_documents_export_word(request):
    """
    Выгрузка текстовой части отчета проверок делопроизводства в ОАГВ на основе шаблона shablon_check_doc.docx
    с помощью модуля docxtpl.
    Возвращает объект HttpResponse с данными о тех РК, где некорректно заполнены реквизиты
    """
    check_documents = CheckDocumentsApi()
    from_date_str = request.GET.get("from_date", None)
    due_date_str = request.GET.get("due_date", None)
    dep_due = request.GET.get("dep_due", None)
    if not from_date_str or not due_date_str:
        return HttpResponseBadRequest("Некорректно задан отчетный период")
    from_date = datetime.strptime(from_date_str, "%Y-%m-%d")
    due_date = datetime.strptime(due_date_str, "%Y-%m-%d")

    if from_date <= due_date:
        cache_key = get_check_documents_cache_key(request.user.id, from_date, due_date, dep_due)
        cached_results = cache.get(cache_key)  # получение данных из хэша

        if cached_results:
            result = cached_results  # распаковка данных хэша
        else:  # cache = none
            result = check_documents.get_report_data(from_date, due_date, dep_due)
            cache.set(cache_key, result, 60 * 5)  # передаю данные в хэш
    else:
        return HttpResponseBadRequest("Некорректно задан отчетный период")

    report_filename = f'Check_documents.{from_date_str}-{due_date_str}.docx'
    from ud.settings import BASE_DIR
    # путь к шаблону на сервере на основе BASE_DIR
    if os.name == 'posix':
        template_filepath = rf"{BASE_DIR}/deloreports/templates/deloreports/shablon_check_doc.docx"
    # путь к шаблону в Windows для тестовой разработки
    else:
        template_filepath = f"{BASE_DIR}\\deloreports\\templates\\deloreports\\shablon_check_doc.docx"
    dep_name_without_shortname = get_department_by_due(dep_due, "True", "True")
    department_loct_padeg = get_padeg_dep(dep_name_without_shortname)
    context = {
        'start_date': from_date.strftime("%d.%m.%Y"),
        'now_date': datetime.now().strftime("%d.%m.%Y %H:%M"),
        'end_date': due_date.strftime("%d.%m.%Y"),
        'department_loct_padeg': department_loct_padeg,  # структурное подразделение в предложном падеже
        'incoming_documents_without_correspondent_registration_data': len(
            result["incoming_documents_without_correspondent_registration_data"]),
        'incoming_documents_incorrect_rubric': len(result["incoming_documents_incorrect_rubric"]),
        'incoming_documents_without_files': len(result["incoming_documents_without_files"]),
        'outgoing_documents_incorrect_rubric': len(result["outgoing_documents_incorrect_rubric"]),
        'outgoing_documents_no_links': len(result["outgoing_documents_no_links"]),
        'outgoing_documents_without_files': len(result["outgoing_documents_without_files"]),
        'internal_documents_incorrect_rubric': len(result["internal_documents_incorrect_rubric"]),
        'internal_documents_no_eds': len(result["internal_documents_no_eds"]),
        'citizens_appeals_incorrect_rubric': len(result["citizens_appeals_incorrect_rubric"]),
        'sum_errors': len(result["incoming_documents_without_correspondent_registration_data"]) + len(
            result["incoming_documents_incorrect_rubric"]) + len(result["incoming_documents_without_files"]) + len(
            result["outgoing_documents_incorrect_rubric"]) + len(result["outgoing_documents_no_links"]) + len(
            result["outgoing_documents_without_files"]) + len(result["internal_documents_incorrect_rubric"]) + len(
            result["internal_documents_no_eds"]) + len(result["citizens_appeals_incorrect_rubric"]),
        'list_incoming_documents_without_correspondent_registration_data':
            result["incoming_documents_without_correspondent_registration_data"],
        'list_incoming_documents_incorrect_rubric': result["incoming_documents_incorrect_rubric"],
        'list_incoming_documents_without_files': result["incoming_documents_without_files"],
        'list_outgoing_documents_incorrect_rubric': result["outgoing_documents_incorrect_rubric"],
        'list_outgoing_documents_no_links': result["outgoing_documents_no_links"],
        'list_outgoing_documents_without_files': result["outgoing_documents_without_files"],
        'list_internal_documents_incorrect_rubric': result["internal_documents_incorrect_rubric"],
        'list_internal_documents_no_eds': result["internal_documents_no_eds"],
        'list_citizens_appeals_incorrect_rubric': result["citizens_appeals_incorrect_rubric"],
    }

    # Получаем имя временного файла, куда будет сформирован отчет.
    fd, tmp_file = tempfile.mkstemp(suffix=report_filename)
    with open(tmp_file, 'bw') as tmp:
        docx_template = DocxTemplate(template_filepath)
        docx_template.render(context)
        docx_template.save(tmp)

        response = HttpResponse(open(tmp_file, 'rb').read(), content_type='application/msword')
        response['Content-Disposition'] = 'attachment; filename=%s' % report_filename
        # закрываем дескриптор файла
        os.close(fd)
        return response


##### Перечень документов, по которым Мэром города Вологды были произведены редакционные правки при рассмотрении проектов поручений в электронном виде   #####
@permission_required('deloreports.show_editing_resolutions_by_mayor_report', login_url='/deloreports/')
def editing_resolutions_by_mayor_report(request):
    """
    Перечень документов, по которым Мэром города Вологды были произведены редакционные правки при рассмотрении проектов поручений в электронном виде
    """
    resol_list = []
    doc_list = []
    form = DateRangeForm(request.POST)
    from_date = due_date = None
    editing_resolutions_by_mayor = EditingResolutionsByMayorApi()
    if request.method == 'POST':
        if form.is_valid():
            from_date = form.cleaned_data["from_date"]
            # прибавим день к конечной дате, поскольку почему-то в sql запросе к таблице PROT не работает не строгое <= включение конечной даты в заданный период, будем использовать подход с прибалением дня и строгим знаком <
            due_date = form.cleaned_data["due_date"] + timedelta(days=1)
            if from_date <= due_date:
                resol_list, doc_list = editing_resolutions_by_mayor.get_report_data(from_date, due_date)
            else:
                return HttpResponseBadRequest("Некорректно задан отчетный период")

    context = {
        'resol_list': resol_list,
        'doc_list': doc_list,
        'form': form,
        'from_date': from_date,
        'due_date': due_date
    }
    return render(request, 'deloreports/editing_resolutions_by_mayor_report.html', context)


def editing_resolutions_by_mayor_report_export_excel(request):
    pass


def citizens_appeals_json(request):
    """
    Возвращает информацию в формате json об обращениях граждан за период.
    http://10.16.1.72:8001/deloreports/citizens_appeals_json/?start=2020-05-01&end=2020-05-31
    http://10.16.1.72:8001/deloreports/citizens_appeals_list_json/?day=2020-05-01
    http://10.16.1.72:8001/deloreports/citizens_appeals_rubric_json/?start=2020-05-01&end=2020-05-31
    http://10.16.1.72:8001/deloreports/citizens_appeals_executor_json/?start=2020-05-01&end=2020-05-31
    http://10.16.1.72:8001/deloreports/citizens_appeals_address_json/?start=2020-05-01&end=2020-05-31
    
    """
    path_name = request.META.get("PATH_INFO").split("/")[2]
    citizens_appeals_json = CitizensAppealsJsonApi()
    if request.method != 'GET':
        return HttpResponseNotAllowed(['GET'], 'Метод не разрешён')

    result = []
    if path_name == 'citizens_appeals_list_json':
        from_date = request.GET.get("day", None)
        due_date = request.GET.get("day", None)
    else:
        from_date = request.GET.get("start", None)
        due_date = request.GET.get("end", None)

    try:
        from_date = datetime.strptime(from_date, "%Y-%m-%d")
        due_date = datetime.strptime(due_date, "%Y-%m-%d")
    # перехват некорректной даты или значения None
    except (ValueError, TypeError):
        return JsonResponse(result, safe=False, json_dumps_params={'ensure_ascii': False})

    # Пытаемся выдать кешированный результат.
    if from_date <= due_date:
        result = citizens_appeals_json.get_report_data(from_date, due_date, path_name)
        # cache_key = get_cache_key2(path_name, from_date, due_date)
        # cached_results = cache.get(cache_key) #получение данных из хэша

        # if cached_results:
        #     result = cached_results #распаковка данных хэша
        # else: # cache = none
        #     result = citizens_appeals_json.get_report_data(from_date, due_date, path_name)
        #     cache.set(cache_key, result, 60 * 5) #передаю данные в хэш
    else:
        return JsonResponse(result, safe=False, json_dumps_params={'ensure_ascii': False})

    # if result:
    #     for row in result:
    #         data.append({
    #             'reg_date': row[0].strftime('%Y-%m-%d'),
    #             'count': row[1]
    #         })

    return JsonResponse(result, safe=False, json_dumps_params={'ensure_ascii': False})

# def daterange(request):
#     if request.method == 'POST':
#         form = DateRangeForm(request.POST)
#         if form.is_valid():
#             start_date = form.cleaned_data['start_date']
#             end_date = form.cleaned_data['end_date']
#             print('start_date', start_date)
#             print('end_date', end_date)
#     else:
#         form = DateRangeForm()
#
#     context = {'form': form}
#
#     return render(request, 'deloreports/daterange.html', context)
