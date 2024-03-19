from django.core.management import BaseCommand
from dashboard.sql.db_queries import DjangoSqlQuery


class Command(BaseCommand):
    def add_arguments(self, parser):
        parser.add_argument('--start', type=str, default=None, help='Start month_year date in format "%m.%Y"')
        parser.add_argument('--end', type=str, default=None, help='End month_year date in format "%m.%Y"')

    def handle(self, *args, **options):
        start = options['start']
        end = options['end']

        if start:
            DjangoSqlQuery().set_and_update_resolutions_count_of_month(start, end)
        else:
            DjangoSqlQuery().set_and_update_resolutions_count_of_month()
