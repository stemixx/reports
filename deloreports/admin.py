from django.contrib import admin
from deloreports.models import Assistant, Department, DocGroup

admin.register(Assistant)
admin.site.register(DocGroup)


@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ['name', 'due']
    # list_filter = ['is_deleted']
    search_fields = ['name']
