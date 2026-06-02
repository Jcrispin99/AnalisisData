from django.contrib import admin
from django.shortcuts import render

from . import reportes


def reporte_atenciones_view(request):
    year = request.GET.get("year") or None
    convenio = request.GET.get("convenio") or None
    area = request.GET.get("area") or None
    altitud_banda = request.GET.get("altitud") or None
    departamento = request.GET.get("departamento") or None

    if year:
        try:
            year = int(year)
        except ValueError:
            year = None

    if altitud_banda not in reportes.ALTITUD_BANDAS_DICT:
        altitud_banda = None

    data = reportes.reporte_atenciones(
        year=year,
        convenio=convenio,
        area=area,
        altitud_banda=altitud_banda,
        departamento=departamento,
    )
    filtros = reportes.filtros_disponibles()

    # Aplana las series por canvas id. Para filas by_sex generamos 3 canvas:
    # uno general (toda la población), uno M y uno F. El frontend muestra
    # u oculta los M/F con flechas (swap con la tabla opuesta).
    chart_data = {}
    for p in data["pestañas"]:
        for fila in p["filas"]:
            cid = f"chart-{p['id']}-{fila['clave']}"
            if fila.get("by_sex"):
                chart_data[f"{cid}-general"] = fila["serie"]
                chart_data[f"{cid}-M"] = fila["por_sexo"]["M"]["serie"]
                chart_data[f"{cid}-F"] = fila["por_sexo"]["F"]["serie"]
            else:
                chart_data[cid] = fila["serie"]

    context = {
        **admin.site.each_context(request),
        "title": "Reporte clínico de atenciones",
        "data": data,
        "insights": data.get("insights", []),
        "chart_data": chart_data,
        "filtros": filtros,
        "años": reportes.años_disponibles(),
        "altitud_bandas": reportes.ALTITUD_BANDAS,
        "selected_year": year,
        "selected_convenio": convenio or "",
        "selected_area": area or "",
        "selected_altitud": altitud_banda or "",
        "selected_departamento": (departamento or "").upper(),
    }
    return render(request, "admin/pacientes/reporte_atenciones.html", context)
