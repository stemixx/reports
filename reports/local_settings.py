# Build paths inside the project like this: os.path.join(BASE_DIR, ...)
from pathlib import Path
# from settings import INSTALLED_APPS


BASE_DIR = Path(__file__).resolve().parent.parent

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = True

TEMPLATE_DEBUG = True

COMPRESS_ENABLED = False

# Application definition

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'channels',
    'compressor',
    'guardian',
    'user_visit',
    'deloreports',
    'dashboard.apps.DashboardConfig',
    'django_plotly_dash.apps.DjangoPlotlyDashConfig',

    'debug_toolbar',
]

INTERNAL_IPS = [
    '127.0.0.1',
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'debug_toolbar.middleware.DebugToolbarMiddleware',
    'user_visit.middleware.UserVisitMiddleware',
]

# DATABASES = {
#     'default': {
#         'ENGINE': 'django.db.backends.postgresql_psycopg2',
#         'NAME': 'ud_test',
#         'USER': 'ud',
#         'PASSWORD': 'ud3214',
#         'HOST': '10.16.8.38',
#         'PORT': '5434',
#     }
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql',
        'NAME': 'ud_test_backup',
        'USER': 'ud',
        'PASSWORD': 'ud3214',
        'HOST': '127.0.0.1',
        'PORT': '5432',
    }
    # 'default': {
    #     'ENGINE': 'django.db.backends.sqlite3',
    #     'NAME': os.path.join(BASE_DIR, 'db.sqlite3'),
    # }
}

# ODBC-драйвер для подключения к СУБД
DELO_DB_DRIVER = 'SQL Server'
