from django.contrib import admin
from django.shortcuts import render

from . import reportes


def reporte_atenciones_view(request):
    year = request.GET.get("year") or None
    convenio = request.GET.get("convenio") or None
    area = request.GET.get("area") or None
    sexo = request.GET.get("sexo") or None
    if sexo not in ("M", "F"):
        sexo = None

    if year:
        try:
            year = int(year)
        except ValueError:
            year = None

    data = reportes.reporte_atenciones(year=year, convenio=convenio, area=area, sexo=sexo)
    filtros = reportes.filtros_disponibles()

    # Aplana las series por canvas id para que el template las exponga vía json_script.
    chart_data = {}
    for p in data["pestañas"]:
        for fila in p["filas"]:
            chart_data[f"chart-{p['id']}-{fila['clave']}"] = fila["serie"]

    context = {
        **admin.site.each_context(request),
        "title": "Reporte clínico de atenciones",
        "data": data,
        "kpis": data.get("kpis", []),
        "insights": data.get("insights", []),
        "chart_data": chart_data,
        "filtros": filtros,
        "años": reportes.años_disponibles(),
        "selected_year": year,
        "selected_convenio": convenio or "",
        "selected_area": area or "",
        "selected_sexo": sexo or "",
    }
    return render(request, "admin/pacientes/reporte_atenciones.html", context)
