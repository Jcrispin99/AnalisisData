"""Agregaciones para el reporte clínico basado en CRITERIOS DE ESTUDIO.xlsx.

`reporte_atenciones` devuelve un dict listo para renderizar:
- `pestañas`: una por sección (Dislipidemias, Eritrocitosis, …). Cada pestaña
  tiene `filas` con la tabla de distribución y una serie temporal por categoría
  (para gráfico de área apilada).
- Si hay año filtrado → bins por mes. Si no → bins por año.

No optimiza con SQL: itera el queryset filtrado en Python porque las
clasificaciones son @property que dependen de campos del paciente.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date

from django.db.models import Count
from django.db.models.functions import ExtractYear

from . import umbrales as u
from .models import Atencion, AtencionNutricion


CATEGORIAS = {
    "dislipidemias": [
        "Deseable",
        "Dislipidemia leve",
        "Dislipidemia moderada",
        "Dislipidemia severa",
    ],
    "dislipidemia_global": [
        "Normal",
        "HDL bajo aislado",
        "Hipercolesterolemia",
        "Hipertrigliceridemia",
        "Dislipidemia mixta",
        "Dislipidemia aterogénica",
    ],
    "colesterol_no_hdl": [
        "Deseable",
        "Limítrofe alto",
        "Riesgo aumentado",
        "Riesgo muy alto",
    ],
    "eritrocitosis": [
        "Normal",
        "Eritrocitosis leve",
        "Eritrocitosis moderada",
        "Eritrocitosis severa",
    ],
    "obesidad_abdominal": [
        "Normal",
        "Grado 1 (riesgo aumentado)",
        "Grado 2 (riesgo significativo)",
    ],
    "diabetes": ["Normal", "Prediabetes", "Diabetes"],
    "hta": [
        "Normal",
        "Presión elevada",
        "Hipertensión grado 1",
        "Hipertensión grado 2",
        "Crisis hipertensiva",
    ],
    "imc": [
        "Bajo peso",
        "Normal",
        "Sobrepeso",
        "Obesidad grado I",
        "Obesidad grado II",
        "Obesidad grado III",
    ],
    "grasa_corporal": ["Bajo", "Saludable", "Aceptable", "Obesidad"],
    "grasa_visceral": ["Saludable", "Alta"],
    "masa_muscular": ["Bajo", "Normal", "Alto", "Muy alto"],
}

# (clave, etiqueta, property, categorías)
FILAS = [
    ("dislipidemia_global", "Diagnóstico de dislipidemia", "clasif_dislipidemia_tipo", CATEGORIAS["dislipidemia_global"]),
    ("colesterol_total", "Colesterol Total", "clasif_colesterol_total", CATEGORIAS["dislipidemias"]),
    ("hdl", "HDL Colesterol", "clasif_hdl", CATEGORIAS["dislipidemias"]),
    ("ldl", "LDL Colesterol", "clasif_ldl", CATEGORIAS["dislipidemias"]),
    ("trigliceridos", "Triglicéridos", "clasif_trigliceridos", CATEGORIAS["dislipidemias"]),
    ("colesterol_no_hdl", "Colesterol no-HDL", "clasif_colesterol_no_hdl", CATEGORIAS["colesterol_no_hdl"]),
    ("eritrocitosis", "Eritrocitosis (Hb)", "clasif_eritrocitosis", CATEGORIAS["eritrocitosis"]),
    ("obesidad_abdominal", "Obesidad abdominal", "clasif_obesidad_abdominal", CATEGORIAS["obesidad_abdominal"]),
    ("glicemia", "Glicemia en ayunas", "clasif_glicemia_ayunas", CATEGORIAS["diabetes"]),
    ("hb_a1c", "Hemoglobina glicosilada", "clasif_hb_a1c", CATEGORIAS["diabetes"]),
    ("presion", "Presión arterial", "clasif_presion", CATEGORIAS["hta"]),
]

SECCIONES = [
    ("dislipidemias", "Dislipidemias", [
        "dislipidemia_global", "colesterol_total", "hdl", "ldl",
        "trigliceridos", "colesterol_no_hdl",
    ]),
    ("eritrocitosis", "Eritrocitosis", ["eritrocitosis"]),
    ("obesidad_abdominal", "Obesidad abdominal", ["obesidad_abdominal"]),
    ("diabetes", "Diabetes Mellitus", ["glicemia", "hb_a1c"]),
    ("hta", "Hipertensión arterial", ["presion"]),
]

FILAS_NUTRICION = [
    ("imc", "IMC", "clasif_imc", CATEGORIAS["imc"]),
    ("grasa_corporal", "% Grasa corporal", "clasif_grasa_corporal", CATEGORIAS["grasa_corporal"]),
    ("grasa_visceral", "Grasa visceral", "clasif_grasa_visceral", CATEGORIAS["grasa_visceral"]),
    ("masa_muscular", "% Masa muscular", "clasif_masa_muscular", CATEGORIAS["masa_muscular"]),
]

_MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Paleta alineada con la del chart (verde/amarillo/naranja/rojo/morado/azul).
_PALETA = ["#22c55e", "#eab308", "#f97316", "#ef4444", "#a855f7", "#3b82f6"]


def _color_para(cats, i):
    """Color por posición en la lista de categorías.

    Para listas de 3 elementos donde la última categoría es 'severa'
    (diabetes, obesidad abdominal), saltamos del amarillo al rojo
    para reflejar la severidad correctamente.
    """
    if len(cats) == 3:
        return [_PALETA[0], _PALETA[1], _PALETA[3]][i]
    return _PALETA[i % len(_PALETA)]


# Tooltips de criterios que NO son simples cortes (combinador, multi-variable, o
# con formato compacto). Usan f-strings que leen las MISMAS constantes que la
# función clasificadora, así no pueden desincronizarse.
_p = u.PRESION
_TOOLTIP_DISLIPIDEMIA_GLOBAL = (
    "<p class='font-semibold mb-1'>Diagnóstico combinado del perfil lipídico</p>"
    "<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>Combina CT, LDL, HDL y TG:</p>"
    "<ul class='list-disc pl-4 space-y-0.5'>"
    "<li><b>Normal:</b> los 4 valores en rango deseable</li>"
    "<li><b>HDL bajo aislado:</b> solo HDL bajo</li>"
    "<li><b>Hipercolesterolemia:</b> CT/LDL alto, TG normal</li>"
    "<li><b>Hipertrigliceridemia:</b> solo TG alto</li>"
    "<li><b>Dislipidemia mixta:</b> colesterol + TG altos</li>"
    "<li><b>Aterogénica:</b> TG altos + HDL bajo</li>"
    "</ul>"
    "<p class='mt-1.5 text-font-subtle-light dark:text-font-subtle-dark'>"
    f"Severidad = peor componente alterado. Bandera ⚠ si LDL ≥{u.LDL_UMBRAL_HF} (sospecha "
    "hipercolesterolemia familiar — ACC/AHA)."
    "</p>"
)

_TOOLTIP_PRESION = (
    "<p class='font-semibold mb-1'>Presión arterial</p>"
    "<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>"
    "Fuente: CRITERIOS DE ESTUDIO. Se toma la categoría más severa entre sistólica y diastólica."
    "</p>"
    "<ul class='list-disc pl-4 space-y-0.5'>"
    f"<li>Normal: &lt;{_p['elevada_s_min']} <b>y</b> &lt;{_p['elevada_d_max']}</li>"
    f"<li>Elevada: {_p['elevada_s_min']}-{_p['elevada_s_max']} <b>y</b> &lt;{_p['elevada_d_max']}</li>"
    f"<li>Grado 1: {_p['g1_s_min']}-{_p['g1_s_max']} <b>ó</b> {_p['g1_d_min']}-{_p['g1_d_max']}</li>"
    f"<li>Grado 2: ≥{_p['g2_s']} <b>ó</b> ≥{_p['g2_d']}</li>"
    f"<li>Crisis: ≥{_p['crisis_s']} <b>y/o</b> ≥{_p['crisis_d']}</li>"
    "</ul>"
)

_TOOLTIP_ERITROCITOSIS = (
    "<p class='font-semibold mb-1'>Eritrocitosis — Hb por sexo + altitud</p>"
    "<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>Fuente: CRITERIOS DE ESTUDIO</p>"
    "<p class='mb-0.5'><b>Varones &lt;2500 msnm:</b></p>"
    f"<p class='pl-2 mb-1'>{u._eritro_linea(u.ERITROCITOSIS[('M', False)])}</p>"
    "<p class='mb-0.5'><b>Varones ≥2500 msnm:</b></p>"
    f"<p class='pl-2 mb-1'>{u._eritro_linea(u.ERITROCITOSIS[('M', True)])}</p>"
    "<p class='mb-0.5'><b>Mujeres &lt;2500 msnm:</b></p>"
    f"<p class='pl-2 mb-1'>{u._eritro_linea(u.ERITROCITOSIS[('F', False)])}</p>"
    "<p class='mb-0.5'><b>Mujeres ≥2500 msnm:</b></p>"
    f"<p class='pl-2 mb-1'>{u._eritro_linea(u.ERITROCITOSIS[('F', True)])}</p>"
    "<p class='mt-1 text-font-subtle-light dark:text-font-subtle-dark'>"
    "Altitud se estima del distrito/provincia/departamento del paciente."
    "</p>"
)

# HTML mostrado en el tooltip "?" al lado del título de cada tabla.
# Los criterios simples se renderizan desde umbrales.py (single source of truth
# con la lógica de clasificación). Los complejos están definidos arriba.
TOOLTIPS = {
    "dislipidemia_global": _TOOLTIP_DISLIPIDEMIA_GLOBAL,
    "colesterol_total":    u.render_tooltip("Colesterol Total", u.COLESTEROL_TOTAL),
    "hdl":                 u.render_tooltip_por_sexo("HDL Colesterol", u.HDL),
    "ldl":                 u.render_tooltip("LDL Colesterol", u.LDL),
    "trigliceridos":       u.render_tooltip("Triglicéridos en ayunas", u.TRIGLICERIDOS),
    "colesterol_no_hdl":   u.render_tooltip("Colesterol no-HDL", u.COLESTEROL_NO_HDL),
    "eritrocitosis":       _TOOLTIP_ERITROCITOSIS,
    "obesidad_abdominal":  u.render_tooltip_por_sexo("Perímetro abdominal", u.OBESIDAD_ABDOMINAL),
    "glicemia":            u.render_tooltip("Glicemia en ayunas", u.GLICEMIA_AYUNAS),
    "hb_a1c":              u.render_tooltip("Hemoglobina glicosilada (HbA1c)", u.HB_A1C),
    "presion":             _TOOLTIP_PRESION,
    "imc":                 u.render_tooltip("Índice de Masa Corporal (IMC)", u.IMC),
    "grasa_corporal":      u.render_tooltip_por_sexo("% Grasa corporal", u.GRASA_CORPORAL,
                                                     fuente_html="Estándar ACSM. El xlsx no define cortes."),
    "grasa_visceral":      u.render_tooltip("Grasa visceral (rating Tanita 1-59)", u.GRASA_VISCERAL),
    "masa_muscular":       u.render_tooltip_por_sexo("% Masa muscular", u.MASA_MUSCULAR,
                                                     fuente_html="Estándar Tanita. El xlsx no define cortes."),
}


def años_disponibles():
    años_clinico = Atencion.objects.annotate(y=ExtractYear("fecha")).values_list("y", flat=True)
    años_nutri = AtencionNutricion.objects.annotate(y=ExtractYear("fecha")).values_list("y", flat=True)
    return sorted({a for a in list(años_clinico) + list(años_nutri) if a}, reverse=True)


def filtros_disponibles():
    convenios = set()
    areas = set()
    for model in (Atencion, AtencionNutricion):
        convenios.update(
            model.objects.exclude(paciente__convenio="").values_list("paciente__convenio", flat=True)
        )
        areas.update(
            model.objects.exclude(paciente__nombre_area="").values_list("paciente__nombre_area", flat=True)
        )
    return {"convenios": sorted(convenios), "areas": sorted(areas)}


def _aplicar_filtros(qs, year, convenio, area, sexo=None):
    if year:
        qs = qs.filter(fecha__year=year)
    if convenio:
        qs = qs.filter(paciente__convenio=convenio)
    if area:
        qs = qs.filter(paciente__nombre_area=area)
    if sexo in ("M", "F"):
        qs = qs.filter(paciente__sexo=sexo)
    return qs


def _bins_y_keyfn(year, fechas):
    """Devuelve (bins, labels, key_fn) para el eje temporal.

    - Si hay year filtrado: 12 meses de ese año, key=mes.
    - Si no, y todas las fechas son del mismo año: 12 meses, key=mes.
    - Si no, y hay varios años: rango (año, mes) de min a max, key=(año, mes),
      labels "May 26".
    Siempre se rellenan los huecos para que el gráfico no se vea degenerado.
    """
    if year:
        return list(range(1, 13)), _MESES, lambda at: at.fecha.month

    años = {f.year for f in fechas if f}
    if not años or len(años) == 1:
        return list(range(1, 13)), _MESES, lambda at: at.fecha.month

    año_min, año_max = min(años), max(años)
    bins = [(a, m) for a in range(año_min, año_max + 1) for m in range(1, 13)]
    labels = [f"{_MESES[m - 1]} {str(a)[-2:]}" for a, m in bins]
    return bins, labels, lambda at: (at.fecha.year, at.fecha.month)


def _resumen(qs, filas_def, year):
    """Cuenta atenciones por categoría y por bin temporal.

    Devuelve (total, filas_dict). Cada fila incluye `celdas` (tabla) y `serie`
    (datasets para Chart.js).
    """
    atenciones = list(qs)
    total = len(atenciones)
    fechas = [at.fecha for at in atenciones]
    bins, bin_labels, key_fn = _bins_y_keyfn(year, fechas)

    # counters[clave][cat] = total ; series[clave][cat][bin] = n
    counters = {clave: Counter() for clave, *_ in filas_def}
    series = {clave: defaultdict(lambda: defaultdict(int)) for clave, *_ in filas_def}

    for at in atenciones:
        b = key_fn(at)
        for clave, _label, prop, _cats in filas_def:
            valor = getattr(at, prop)
            if valor:
                counters[clave][valor] += 1
                series[clave][valor][b] += 1

    filas = {}
    for clave, label, _prop, cats in filas_def:
        counter = counters[clave]
        total_con_dato = sum(counter.values())
        celdas = []
        for i, cat in enumerate(cats):
            n = counter.get(cat, 0)
            pct = (n / total_con_dato * 100) if total_con_dato else 0
            celdas.append({
                "categoria": cat,
                "n": n,
                "pct": pct,
                "color": _color_para(cats, i),
            })

        datasets = [
            {
                "label": cat,
                "data": [series[clave][cat].get(b, 0) for b in bins],
                "color": _color_para(cats, i),
            }
            for i, cat in enumerate(cats)
        ]
        filas[clave] = {
            "clave": clave,
            "label": label,
            "categorias": cats,
            "celdas": celdas,
            "total_con_dato": total_con_dato,
            "sin_dato": total - total_con_dato,
            "serie": {"labels": bin_labels, "datasets": datasets},
            "tooltip": TOOLTIPS.get(clave, ""),
        }
    return total, filas


def reporte_atenciones(year=None, convenio=None, area=None, sexo=None):
    qs_clinico = _aplicar_filtros(
        Atencion.objects.select_related("paciente"), year, convenio, area, sexo
    )
    total_clinico, filas_clinico = _resumen(qs_clinico, FILAS, year)

    qs_nutri = _aplicar_filtros(
        AtencionNutricion.objects.select_related("paciente"), year, convenio, area, sexo
    )
    total_nutri, filas_nutri = _resumen(qs_nutri, FILAS_NUTRICION, year)

    pestañas = [
        {
            "id": pid,
            "titulo": titulo,
            "filas": [filas_clinico[c] for c in claves],
        }
        for pid, titulo, claves in SECCIONES
    ]
    if total_nutri:
        pestañas.append(
            {
                "id": "composicion",
                "titulo": "Composición corporal",
                "filas": [filas_nutri[c] for c, *_ in FILAS_NUTRICION],
            }
        )

    return {"pestañas": pestañas}


def distribucion_atenciones_por_año():
    return list(
        Atencion.objects.annotate(y=ExtractYear("fecha"))
        .values("y")
        .annotate(n=Count("id"))
        .order_by("-y")
        .values_list("y", "n")
    )
