from django.urls import path
from . import views

# этот импорт ниже необходим для библиотеки django_plotly_dash
from dashboard.dash_apps.finished_apps import all_dashboards

app_name = 'dashboard'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
]
