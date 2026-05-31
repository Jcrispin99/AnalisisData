from django.contrib import admin
from django.urls import path

from pacientes.views import reporte_atenciones_view

urlpatterns = [
    path(
        "admin/reportes/atenciones/",
        admin.site.admin_view(reporte_atenciones_view),
        name="reportes_atenciones",
    ),
    path("admin/", admin.site.urls),
]
