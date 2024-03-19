from django.db import models
from deloreports.models import Department


class ReportName(models.Model):
    """
    Название отчёта
    """
    name = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.name


class Statistic(models.Model):
    """
    Статистика документооборота для дашборда
    """
    report_name = models.ForeignKey(ReportName, on_delete=models.CASCADE)
    department = models.ForeignKey(Department, on_delete=models.CASCADE)
    start_date = models.DateField()
    end_date = models.DateField()
    amount = models.IntegerField(blank=True, default=0)
    opened = models.IntegerField(blank=True, default=0)
    closed = models.IntegerField(blank=True, default=0)
    expired = models.IntegerField(blank=True, default=0)
    update_date = models.DateTimeField(auto_now=True)
