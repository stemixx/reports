"""
Django settings for reports project.

Generated by 'django-admin startproject' using Django 4.2.2.

For more information on this file, see
https://docs.djangoproject.com/en/4.2/topics/settings/

For the full list of settings and their values, see
https://docs.djangoproject.com/en/4.2/ref/settings/
"""
import os
from pathlib import Path
import ldap
from django_auth_ldap.config import LDAPSearch, LDAPSearchUnion

# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent

AD_DOMAIN = "XXXX"
AD_USERNAME = "XXXX"
AD_PASSWORD = "XXXXXX"

# Baseline configuration.
AUTH_LDAP_SERVER_URI = "ldap://XX.XX.X.X"

#AUTH_LDAP_START_TLS = True

AUTH_LDAP_BIND_DN = r"%s\%s" % (AD_DOMAIN, AD_USERNAME)
AUTH_LDAP_BIND_PASSWORD = AD_PASSWORD

AUTH_LDAP_USER_SEARCH = LDAPSearchUnion(
    LDAPSearch("ou=Пользователи,dc=%s,dc=local" % AD_DOMAIN,
               ldap.SCOPE_SUBTREE, "(sAMAccountName=%(user)s)"),
    LDAPSearch("cn=users,dc=%s,dc=local" % AD_DOMAIN,
               ldap.SCOPE_SUBTREE, "(sAMAccountName=%(user)s)"),
    LDAPSearch("ou=Управляемые пользователи,dc=%s,dc=local" % AD_DOMAIN,
               ldap.SCOPE_SUBTREE, "(sAMAccountName=%(user)s)"),
    LDAPSearch("ou=Администраторы,dc=%s,dc=local" % AD_DOMAIN,
               ldap.SCOPE_SUBTREE, "(sAMAccountName=%(user)s)")
)

# or perhaps:
#AUTH_LDAP_USER_DN_TEMPLATE = AD_DOMAIN + r"\%(user)s"

# Populate the Django user from the LDAP directory.
AUTH_LDAP_USER_ATTR_MAP = {
    "first_name": "givenName",
    "last_name": "sn",
    "email": "mail"
}

AUTH_LDAP_PROFILE_ATTR_MAP = {
    "employee_number": "employeeNumber"
}

# This is the default, but I like to be explicit.
AUTH_LDAP_ALWAYS_UPDATE_USER = True

# Use LDAP group membership to calculate group permissions.
AUTH_LDAP_FIND_GROUP_PERMS = False

# Keep ModelBackend around for per-user permissions and maybe a local
# superuser.
AUTHENTICATION_BACKENDS = (
    'django_auth_ldap.backend.LDAPBackend',
    'django.contrib.auth.backends.ModelBackend',
    'guardian.backends.ObjectPermissionBackend',
)

RAVEN_CONFIG = {
    'dsn': 'http://8c3e1f89f36341ed9e8a82cd0e5083ef:ce9f334541d442f8bbb14e41ca19dfc1@10.16.1.42:9000/7',
}

SECRET_KEY = 'django-insecure-fk=+ivnh+zxh8y0^&fq7a1u)0tp2e@4v^48-4561ri6gb5a7qz'

# SECURITY WARNING: don't run with debug turned on in production!
DEBUG = False

ALLOWED_HOSTS = ['*']

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
]

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'user_visit.middleware.UserVisitMiddleware',
]

ROOT_URLCONF = 'reports.urls'
SECURE_CROSS_ORIGIN_OPENER_POLICY = None
SESSION_COOKIE_SECURE = False
TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [os.path.join(BASE_DIR, 'deloreports\\templates\\'),
                 os.path.join(BASE_DIR, 'dashboard\\templates\\'),
                 ],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'reports.wsgi.application'

# Database
# https://docs.djangoproject.com/en/4.2/ref/settings/#databases

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.postgresql_psycopg2',
        'NAME': 'reports',
        'USER': 'reports',
        'PASSWORD': 'password',
        'HOST': '127.0.0.1',
        'PORT': '5432',
    }
}

# Password validation
# https://docs.djangoproject.com/en/4.2/ref/settings/#auth-password-validators

AUTH_PASSWORD_VALIDATORS = [
    {
        'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',
    },
    {
        'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',
    },
]

# Internationalization
# https://docs.djangoproject.com/en/4.2/topics/i18n/

LANGUAGE_CODE = 'ru-RU'

TIME_ZONE = 'Europe/Moscow'

USE_I18N = True

USE_TZ = True

# Static files (CSS, JavaScript, Images)
# https://docs.djangoproject.com/en/4.2/howto/static-files/

STATIC_URL = 'static/'
STATIC_ROOT = os.path.join(BASE_DIR, 'static/')

STATICFILES_FINDERS = (
    'django.contrib.staticfiles.finders.FileSystemFinder',
    'django.contrib.staticfiles.finders.AppDirectoriesFinder',
    # other finders..
    'compressor.finders.CompressorFinder',
    'django_plotly_dash.finders.DashAssetFinder',
    'django_plotly_dash.finders.DashComponentFinder',
)

PLOTLY_COMPONENTS = [
    'dash_core_components',
    'dash_html_components',
    'dpd_components'
]

X_FRAME_OPTIONS = 'SAMEORIGIN'

# Параметры подключения к базе данных СЭД «Дело»
# Адрес сервера СУБД, где расположена БД «Дело»
DELO_SERVER_NAME = 'XX.XX.X.XXX'
# Порт СУБД, где расположена БД «Дело»
DELO_SERVER_PORT = 1433
# Имя базы данных СЭД «Дело»
DELO_DB_NAME = 'DELO_DB'
# Имя пользователя БД «Дело» для подключения
DELO_USERNAME = 'XXXXXX'
# Пароль пользователя БД «Дело» для подключения
DELO_PASSWORD = 'XXXXXX'
# ODBC-драйвер для подключения к СУБД
DELO_DB_DRIVER = 'ODBC Driver 17 for SQL Server'

# Default primary key field type
# https://docs.djangoproject.com/en/4.2/ref/settings/#default-auto-field

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

try:
    from .local_settings import *
except ImportError:
    pass
