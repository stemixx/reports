from datetime import date, datetime
from django.http import JsonResponse
from django.views.generic.edit import CreateView
from django import forms
from bootstrap_daterangepicker import widgets, fields
# from django.forms.extras.widgets import SelectDateWidget, Select
from django.forms.widgets import Select, SelectDateWidget
from deloreports.models import DocData
from deloreports.functions.structure import *

#tos = (("1", "один"), ("2", "два"))

YEAR_CHOICES = list(range(2007, date.today().year + 1))


def get_first_date():
    return "01.01." + str(date.today().year)


def get_last_date():
    return datetime.now().strftime("%d.%m.%Y")


def get_tos_list():
    tos_query_set = DocData.objects.exclude(tos="").values_list("tos", flat=True).order_by().distinct("tos")
    result_list = []
    for tos in tos_query_set:
        result_list.append((tos, f"ТОС {tos}"))
    return tuple(result_list)


def get_dep_list():  # используется только для мастера отчетов
    result = []
    agv_structure = AgvStructure()
    departments_list = agv_structure.get_departments_list()
    for dep in departments_list:
        result.append((dep["due"], dep["name"]))
    result.sort(key=lambda item: item[1], reverse=False)
    return tuple(result)


TOS_CHOICES = get_tos_list()
DEPARTMENT = get_dep_list()


class DateRangeForm(forms.Form):
    from_date = forms.DateField()
    due_date = forms.DateField()


class DateFormWithTOSChoices(forms.Form):
    from_date = forms.DateField()
    due_date = forms.DateField()
    tos = forms.MultipleChoiceField(choices=TOS_CHOICES, required=False)


class DateFormSimpleDep(forms.Form):
    # from_date = forms.DateField(widget=forms.TextInput(attrs={"class": "form-control", "value": get_first_date()}))
    # due_date = forms.DateField(widget=forms.TextInput(attrs={"class": "form-control", "value": get_last_date()}))
    from_date = forms.DateField()
    due_date = forms.DateField()
    # text = forms.CharField(max_length=256, required=False)
    dep_name = forms.MultipleChoiceField(choices=DEPARTMENT, required=False)


class DateFormSimpleOneDep(forms.Form):
    from_date = forms.DateField(widget=forms.TextInput(attrs={"class": "form-control", "value": get_first_date()}))
    due_date = forms.DateField(widget=forms.TextInput(attrs={"class": "form-control", "value": get_last_date()}))
    # text = forms.CharField(max_length=256, required=False)
    dep_name = forms.ChoiceField(choices=DEPARTMENT, required=False)


class LoginForm(forms.Form):
    username = forms.CharField(label="Имя пользователя", max_length=256, required=True)
    password = forms.CharField(label="Пароль", max_length=256, widget=forms.PasswordInput, required=True)


class TextForm(forms.Form):
    text = forms.CharField(label="Наименование ТОС", max_length=256, required=True)


class TextFormDep(forms.Form):
    dep_name = forms.CharField(label="Структурное подразделение Администрации города Вологды:", max_length=256,
                               required=True)


class DateMainForm(forms.Form):
    # Date Picker Fields
    date_single_normal = fields.DateField()
    date_single_with_format = fields.DateField(
        input_formats=['%d/%m/%Y'],
        widget=widgets.DatePickerWidget(format='%d/%m/%Y')
    )
    date_single_clearable = fields.DateField(required=False)

    # Date Range Fields
    date_range_normal = fields.DateRangeField()
    date_range_with_format = fields.DateRangeField(
        input_formats=['%d/%m/%Y'],
        widget=widgets.DateRangeWidget(
            format='%d/%m/%Y'
        )
    )
    date_range_clearable = fields.DateRangeField(required=False)

    # DateTime Range Fields
    datetime_range_normal = fields.DateTimeRangeField()
    datetime_range_with_format = fields.DateTimeRangeField(
        input_formats=['%d/%m/%Y (%I:%M:%S)'],
        widget=widgets.DateTimeRangeWidget(
            format='%d/%m/%Y (%I:%M:%S)'
        )
    )
    datetime_range_clearable = fields.DateTimeRangeField(required=False)


class DateForm(forms.Form):
    # from_date = forms.DateTimeField(
    #     input_formats=['%d/%m/%Y %H:%M'],
    #     widget=forms.DateTimeInput(attrs={
    #         'class': 'form-control datetimepicker-input',
    #         'data-target': '#datetimepicker1'
    #     })
    # )
    #
    # date_range = DateRangeField(
    #     widget=BootstrapDateRangePickerInput(),
    #     label='Выберите диапазон дат:',
    #     required=False,
    # )

    from_date = forms.DateField(
        widget=SelectDateWidget(years=YEAR_CHOICES, attrs=({'class': 'form-control form-date-field'})),
        label='Birthday', required=False)
    due_date = forms.DateField(
        widget=SelectDateWidget(years=YEAR_CHOICES, attrs=({'class': 'form-control form-date-field'})),
        label='Birthday', required=False)


class DumaSessionDateForm(DateForm):
    is_extraordinary = forms.BooleanField(required=False)


class PaperFlowDateForm(DateRangeForm):
    paper_pack_cost = forms.IntegerField(required=False)
