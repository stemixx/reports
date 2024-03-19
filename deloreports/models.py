from django.db import models
from datetime import datetime
# from django.contrib.auth.models import User


def getnow():
    return datetime.now()


# Create your models here.
class Assistant(models.Model):
    class Meta:
        permissions = (
            ("show_effectiveness_report", "Can see effectiveness report"),
            ("show_effectiveness_quality_report", "Can see effectiveness quality report"),
            ("show_docflow_report", "Can see docflow report"),
            ("show_court_documents_report", "Can see court documents report"),
            ("show_citizen_rubrics_app_eight_report", "Can see App.8 citizen guidelines"),
            ("show_municipal_legal_act_registration_report", "Can see MPA registration report"),
            ("show_reports_master_report", "Can see Reports Master report"),
            ("show_prosecutors_reaction_act_report", "Can see prosecutors reaction acts report"),
            ("show_control_cases_report", "Can see control cases report"),
            ("show_control_cases_report2", "Can see control cases report2"),
            ("show_prosecutors_incoming_docs_report", "Can see prosecutors incoming docs report"),
            ("show_paper_flow_report", "Can see paper flow report"),
            ("show_correspondents_report", "Can see correspondents report"),
            ("show_correspondents_foiv_report", "Can see correspondents foiv report"),
            ("show_district_themes_report", "Can see distinct themes report"),
            ("show_SSTU_report", "Can see SSTU report"),
            ("show_constituency_report", "Can see constistuency report"),
            ("show_tos_appeals_report", "Can see tos appeals report"),
            ("show_editing_resolutions_by_mayor_report", "Can see editing resolutions by mayor report"),
            ("show_check_documents", "Can see check documents report"),
        )


class DocData(models.Model):
    isn_doc = models.IntegerField("ISN обращения гражданина в СЭД ДЕЛО", db_index=True)
    elect_lot_number = models.CharField("Номер ИУ", max_length=200, blank=True, default="")
    ogd_number = models.CharField("Номер избирательного округа ВГД", max_length=200, blank=True, default="")
    ozs_number = models.CharField("Номер избирательного округа ЗСО ВО", max_length=200, blank=True, default="")
    # theme_code = models.CharField(max_length=24)
    address = models.CharField("Адрес проблемы с СЭД ДЕЛО", max_length=500)
    create_date = models.DateTimeField("Дата создания записи", auto_now_add=True)
    update_date = models.DateTimeField("Дата обновления записи", auto_now=True)
    tos = models.CharField("ТОС", max_length=200, blank=True, default="")

    class Meta:
        verbose_name = "ИУ адреса проблемы из обращения"
        verbose_name_plural = "ИУ адресов проблем из обращений"
        ordering = ["create_date"]


class Citizens_Appeals_By_Executors(models.Model):
    doc_num = models.CharField("Регистрационный номер", max_length=50)
    doc_date = models.DateField("Дата регистрации", auto_now=False, auto_now_add=False)
    annotat = models.TextField("Содержание", blank=True)
    res = models.CharField("Результат", max_length=60, blank=True)
    rubr_list = models.TextField("Список рубрик", blank=True)
    surname_list = models.TextField("Список заявителей", blank=True)
    executor_list = models.TextField("Список исполнителей", blank=True)
    plan_date_list = models.TextField("Список плановых дат", blank=True)
    isn_doc = models.PositiveIntegerField("ISN документа")
    kind = models.PositiveIntegerField("Тип документа")
    update_date = models.DateTimeField("Дата обновления записи", auto_now=True)

    class Meta:
        verbose_name = "Обращение граждан(ина) с исполнителем"
        verbose_name_plural = "Обращения граждан в разрезе исполнителей"
        ordering = ["doc_date", "doc_num", "isn_doc"]


class Citizens_Appeals_With_Address(models.Model):
    doc_num = models.CharField("Регистрационный номер", max_length=50)
    doc_date = models.DateField("Дата регистрации", auto_now=False, auto_now_add=False)
    annotat = models.TextField("Содержание", blank=True)
    res = models.CharField("Результат", max_length=60, blank=True)
    rubr_list = models.TextField("Список рубрик", blank=True)
    surname_list = models.TextField("Список заявителей", blank=True)
    problem_address_list = models.TextField("Адрес проблемы", blank=True)
    executor_list = models.TextField("Список исполнителей", blank=True)
    plan_date_list = models.TextField("Список плановых дат", blank=True)
    isn_doc = models.PositiveIntegerField("ISN документа")
    kind = models.PositiveIntegerField("Тип документа")
    update_date = models.DateTimeField("Дата обновления записи", auto_now=True)

    class Meta:
        verbose_name = "Обращение гр. с адресом проблемы"
        verbose_name_plural = "Обращения граждан с адресами проблем"
        ordering = ["doc_date", "doc_num", "isn_doc"]


class DocGroup(models.Model):
    """
    Группа документов
    """
    name = models.CharField("наименование", max_length=64)
    due = models.CharField("код", max_length=48, unique=True, db_index=True)

    class Meta:
        verbose_name = "группа документов"
        verbose_name_plural = "группы документов"
        ordering = ["name"]

    def __str__(self):
        return self.name


class Department(models.Model):
    """
    Справочник структурных подразделений Администрации города Вологды
    """
    name = models.CharField("наименование", max_length=255)
    short_name = models.CharField("сокращенное наименование", max_length=50)
    due = models.CharField("код подразделения", max_length=50)
    # id_deleted = models.BooleanField(default=False)

    inbound_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="входящие документы", blank=True, related_name="%(class)s_inbound"
    )
    citizen_docgroups_spoken = models.ManyToManyField(
        DocGroup, verbose_name="устные обращения граждан", blank=True, related_name="%(class)s_citizen_spoken"
    )
    citizen_docgroups_written = models.ManyToManyField(
        DocGroup, verbose_name="письменные обращения граждан", blank=True, related_name="%(class)s_citizen_written"
    )
    submission_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="представления (акты прокурорского реагирования)", blank=True,
        related_name="%(class)s_submission"
    )
    protest_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="протесты (акты прокурорского реагирования)", blank=True,
        related_name="%(class)s_protest"
    )
    demand_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="требования (акты прокурорского реагирования)", blank=True,
        related_name="%(class)s_demand"
    )
    court_inbound_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="входящие судебные документы", blank=True, related_name="%(class)s_court_inbound"
    )
    inbound_docgroups_official = models.ManyToManyField(
        DocGroup, verbose_name="входящие официальные документы", blank=True, related_name="%(class)s_official_inbound"
    )
    internal_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="внутренние документы", blank=True, related_name="%(class)s_internal"
    )
    protocol_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="протоколы совещаний", blank=True, related_name="%(class)s_protocol"
    )
    memorandum_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="служебные записки", blank=True, related_name="%(class)s_memorandum"
    )
    control_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="внутренние документы по контролю", blank=True, related_name="%(class)s_control"
    )
    order_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="приказы", blank=True, related_name="%(class)s_order"
    )
    conclusion_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="заключения", blank=True, related_name="%(class)s_conclusion"
    )
    ruling_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="постановления", blank=True, related_name="%(class)s_ruling"
    )
    disposal_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="распоряжения", blank=True, related_name="%(class)s_disposal"
    )
    assignment_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="поручения", blank=True, related_name="%(class)s_assignment"
    )
    outbound_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="исходящие документы", blank=True, related_name="%(class)s_outbound"
    )
    court_outbound_docgroups = models.ManyToManyField(
        DocGroup, verbose_name="исходящие судебные документы", blank=True, related_name="%(class)s_court_outbound"
    )

    class Meta:
        verbose_name = "подразделение"
        verbose_name_plural = "подразделения"
        ordering = ["name"]

        permissions = (
            ("get_report_for", "Может получать отчет для текущего подразделения"),
        )

    def __str__(self):
        return self.name
