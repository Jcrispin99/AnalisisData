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
from typing import NamedTuple

from django.db.models import Count
from django.db.models.functions import ExtractYear

from . import umbrales as u
from .models import Atencion, AtencionNutricion


class FilaDef(NamedTuple):
    clave: str
    label: str
    prop: str
    cats: list
    by_sex: bool = False  # True si el criterio tiene umbrales distintos por sexo


# Categorías consideradas "saludables" por criterio (lo demás se cuenta como alterado).
CATEGORIAS_BUENAS = {
    "dislipidemia_global": {"Normal"},
    "colesterol_total": {"Deseable"},
    "hdl": {"Deseable"},
    "ldl": {"Deseable"},
    "trigliceridos": {"Deseable"},
    "colesterol_no_hdl": {"Deseable"},
    "eritrocitosis": {"Normal"},
    "obesidad_abdominal": {"Normal"},
    "glicemia": {"Normal"},
    "hb_a1c": {"Normal"},
    "presion": {"Normal Alta"},
    "imc": {"Normal"},
    "grasa_corporal": {"Saludable", "Aceptable"},
    "grasa_visceral": {"Saludable"},
    "masa_muscular": {"Normal", "Alto", "Muy alto"},
}


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
        "Normal Alta",
        "Elevada",
        "Hipertensión Grado 1",
        "Hipertensión Grado 2",
        "Crisis Hipertensiva",
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

# by_sex=True cuando el criterio tiene umbrales distintos para varones y mujeres
# (ver umbrales.py — CriterioPorSexo).
FILAS = [
    FilaDef("dislipidemia_global", "Diagnóstico de dislipidemia", "clasif_dislipidemia_tipo", CATEGORIAS["dislipidemia_global"]),
    FilaDef("colesterol_total", "Colesterol Total", "clasif_colesterol_total", CATEGORIAS["dislipidemias"]),
    FilaDef("hdl", "HDL Colesterol", "clasif_hdl", CATEGORIAS["dislipidemias"], by_sex=True),
    FilaDef("ldl", "LDL Colesterol", "clasif_ldl", CATEGORIAS["dislipidemias"]),
    FilaDef("trigliceridos", "Triglicéridos", "clasif_trigliceridos", CATEGORIAS["dislipidemias"]),
    FilaDef("colesterol_no_hdl", "Colesterol no-HDL", "clasif_colesterol_no_hdl", CATEGORIAS["colesterol_no_hdl"]),
    FilaDef("eritrocitosis", "Eritrocitosis (Hb)", "clasif_eritrocitosis", CATEGORIAS["eritrocitosis"], by_sex=True),
    FilaDef("obesidad_abdominal", "Obesidad abdominal", "clasif_obesidad_abdominal", CATEGORIAS["obesidad_abdominal"], by_sex=True),
    FilaDef("glicemia", "Glicemia en ayunas", "clasif_glicemia_ayunas", CATEGORIAS["diabetes"]),
    FilaDef("hb_a1c", "Hemoglobina glicosilada", "clasif_hb_a1c", CATEGORIAS["diabetes"]),
    FilaDef("presion", "Presión arterial", "clasif_presion", CATEGORIAS["hta"]),
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
    FilaDef("imc", "IMC", "clasif_imc", CATEGORIAS["imc"]),
    FilaDef("grasa_corporal", "% Grasa corporal", "clasif_grasa_corporal", CATEGORIAS["grasa_corporal"], by_sex=True),
    FilaDef("grasa_visceral", "Grasa visceral", "clasif_grasa_visceral", CATEGORIAS["grasa_visceral"]),
    FilaDef("masa_muscular", "% Masa muscular", "clasif_masa_muscular", CATEGORIAS["masa_muscular"], by_sex=True),
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
    "<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>"
    "Categorías <b>mutuamente excluyentes</b> (cada paciente cae en una sola). Combina CT, LDL, HDL y TG:"
    "</p>"
    "<ul class='list-disc pl-4 space-y-0.5'>"
    "<li><b>Normal:</b> los 4 valores en rango deseable</li>"
    "<li><b>HDL bajo aislado:</b> solo HDL bajo (CT, LDL y TG normales)</li>"
    "<li><b>Hipercolesterolemia:</b> CT y/o LDL alto, con TG normal</li>"
    "<li><b>Hipertrigliceridemia:</b> TG alto <b>aislado</b> (CT, LDL y HDL normales)</li>"
    "<li><b>Dislipidemia mixta:</b> CT/LDL altos <b>+</b> TG altos (HDL normal)</li>"
    "<li><b>Aterogénica:</b> TG altos <b>+</b> HDL bajo (prevalece sobre mixta)</li>"
    "</ul>"
    "<p class='mt-1.5 text-font-subtle-light dark:text-font-subtle-dark'>"
    "Por eso un paciente con TG alto puede no contar como “Hipertrigliceridemia”: "
    "si además tiene CT/LDL alto cuenta como mixta, y si tiene HDL bajo cuenta como aterogénica."
    "</p>"
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
    f"<li>Normal Alta: &lt;{_p['elevada_s_min']} <b>y</b> &lt;{_p['elevada_d_max']}</li>"
    f"<li>Elevada: {_p['elevada_s_min']}-{_p['elevada_s_max']} <b>y</b> &lt;{_p['elevada_d_max']}</li>"
    f"<li>Hipertensión Grado 1: {_p['g1_s_min']}-{_p['g1_s_max']} <b>ó</b> {_p['g1_d_min']}-{_p['g1_d_max']}</li>"
    f"<li>Hipertensión Grado 2: ≥{_p['g2_s']} <b>ó</b> ≥{_p['g2_d']}</li>"
    f"<li>Crisis Hipertensiva: ≥{_p['crisis_s']} <b>y/o</b> ≥{_p['crisis_d']}</li>"
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


def _construir_celdas(counter, cats):
    """Genera celdas (categoria/n/pct/color) y total con dato a partir de un Counter."""
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
    return celdas, total_con_dato


def _pct_alterado(celdas, total_con_dato, buenas):
    """% de categoría 'alterada' (no en el set de categorías buenas)."""
    if not total_con_dato:
        return 0
    alt = sum(c["n"] for c in celdas if c["categoria"] not in buenas)
    return alt / total_con_dato * 100


def _resumen(qs, filas_def, year):
    """Cuenta atenciones por categoría y por bin temporal.

    Devuelve (total, filas_dict). Cada fila incluye `celdas` (tabla), `serie`
    (datasets para Chart.js) y `pct_alterado`. Si `by_sex=True`, además expone
    `por_sexo` con sub-tablas separadas para varones y mujeres.
    """
    atenciones = list(qs)
    total = len(atenciones)
    fechas = [at.fecha for at in atenciones]
    bins, bin_labels, key_fn = _bins_y_keyfn(year, fechas)

    # counters[clave][cat] = total ; series[clave][cat][bin] = n
    counters = {f.clave: Counter() for f in filas_def}
    series = {f.clave: defaultdict(lambda: defaultdict(int)) for f in filas_def}
    # counters_sex[clave][sexo][cat] solo se llena para filas by_sex
    counters_sex = {
        f.clave: {"M": Counter(), "F": Counter()} for f in filas_def if f.by_sex
    }
    sin_sexo_count = 0  # atenciones cuyo paciente no tiene sexo registrado

    for at in atenciones:
        b = key_fn(at)
        sexo = at.paciente.sexo if at.paciente.sexo in ("M", "F") else None
        if sexo is None:
            sin_sexo_count += 1
        for f in filas_def:
            valor = getattr(at, f.prop)
            if valor:
                counters[f.clave][valor] += 1
                series[f.clave][valor][b] += 1
                if f.by_sex and sexo:
                    counters_sex[f.clave][sexo][valor] += 1

    filas = {}
    for f in filas_def:
        celdas, total_con_dato = _construir_celdas(counters[f.clave], f.cats)
        buenas = CATEGORIAS_BUENAS.get(f.clave, set())
        pct_alterado = _pct_alterado(celdas, total_con_dato, buenas)

        datasets = [
            {
                "label": cat,
                "data": [series[f.clave][cat].get(b, 0) for b in bins],
                "color": _color_para(f.cats, i),
            }
            for i, cat in enumerate(f.cats)
        ]
        fila = {
            "clave": f.clave,
            "label": f.label,
            "categorias": f.cats,
            "celdas": celdas,
            "total_con_dato": total_con_dato,
            "sin_dato": total - total_con_dato,
            "pct_alterado": pct_alterado,
            "serie": {"labels": bin_labels, "datasets": datasets},
            "tooltip": TOOLTIPS.get(f.clave, ""),
            "by_sex": f.by_sex,
        }

        if f.by_sex:
            por_sexo = {}
            for sx in ("M", "F"):
                celdas_sx, total_sx = _construir_celdas(counters_sex[f.clave][sx], f.cats)
                por_sexo[sx] = {
                    "celdas": celdas_sx,
                    "total_con_dato": total_sx,
                    "pct_alterado": _pct_alterado(celdas_sx, total_sx, buenas),
                }
            fila["por_sexo"] = por_sexo
            fila["sin_sexo"] = sin_sexo_count

        filas[f.clave] = fila
    return total, filas


def _calcular_kpis(filas_clinico, total_clinico):
    """KPIs gerenciales — sintetizados de las filas ya calculadas para evitar
    re-iterar atenciones. Cada KPI: clave, label, n, pct, hint."""
    def pct(n):
        return (n / total_clinico * 100) if total_clinico else 0

    def alterados(clave):
        fila = filas_clinico.get(clave)
        if not fila:
            return 0
        buenas = CATEGORIAS_BUENAS.get(clave, set())
        return sum(c["n"] for c in fila["celdas"] if c["categoria"] not in buenas)

    hta_alt = alterados("presion") - sum(
        c["n"] for c in filas_clinico["presion"]["celdas"]
        if c["categoria"] == "Elevada"
    )  # excluye "Elevada" para HTA Grado 1+
    diab = sum(
        c["n"] for c in filas_clinico.get("glicemia", {}).get("celdas", [])
        if c["categoria"] == "Diabetes"
    ) + sum(
        c["n"] for c in filas_clinico.get("hb_a1c", {}).get("celdas", [])
        if c["categoria"] == "Diabetes"
    )
    prediab = sum(
        c["n"] for c in filas_clinico.get("glicemia", {}).get("celdas", [])
        if c["categoria"] == "Prediabetes"
    )

    return [
        {"clave": "total", "label": "Atenciones del periodo", "n": total_clinico,
         "pct": None, "es_total": True, "hint": "Población base del reporte"},
        {"clave": "dislipidemia", "label": "Con dislipidemia",
         "n": alterados("dislipidemia_global"),
         "pct": pct(alterados("dislipidemia_global")),
         "hint": "Cualquier perfil lipídico distinto de Normal"},
        {"clave": "hta", "label": "HTA Grado 1 o más",
         "n": hta_alt, "pct": pct(hta_alt),
         "hint": "Hipertensión confirmada (excluye «Elevada»)"},
        {"clave": "obesidad", "label": "Obesidad abdominal",
         "n": alterados("obesidad_abdominal"),
         "pct": pct(alterados("obesidad_abdominal")),
         "hint": "Perímetro abdominal en Grado 1 o 2"},
        {"clave": "diabetes", "label": "Diabetes o prediabetes",
         "n": diab + prediab, "pct": pct(diab + prediab),
         "hint": f"Diabetes: {diab} · Prediabetes: {prediab}"},
    ]


def _generar_insights(filas_clinico, filas_nutri):
    """Detecta diferencias notables M vs F en criterios sex-dependientes.

    Criterio: ambos sexos con n >= 20 y diferencia en % alterado >= 10pp.
    Devuelve hasta los 3 insights más fuertes.
    """
    insights = []
    for source in (filas_clinico, filas_nutri):
        for clave, fila in source.items():
            if not fila.get("by_sex"):
                continue
            ps = fila["por_sexo"]
            n_m, n_f = ps["M"]["total_con_dato"], ps["F"]["total_con_dato"]
            if n_m < 20 or n_f < 20:
                continue
            pct_m, pct_f = ps["M"]["pct_alterado"], ps["F"]["pct_alterado"]
            diff = abs(pct_m - pct_f)
            if diff < 10:
                continue
            mayor = "mujeres" if pct_f > pct_m else "varones"
            menor = "varones" if pct_f > pct_m else "mujeres"
            insights.append({
                "criterio": fila["label"],
                "mayor": mayor, "menor": menor,
                "pct_mayor": max(pct_m, pct_f),
                "pct_menor": min(pct_m, pct_f),
                "diff": diff,
            })
    insights.sort(key=lambda x: -x["diff"])
    return insights[:3]


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

    return {
        "pestañas": pestañas,
        "kpis": _calcular_kpis(filas_clinico, total_clinico),
        "insights": _generar_insights(filas_clinico, filas_nutri),
        "total_clinico": total_clinico,
    }


def distribucion_atenciones_por_año():
    return list(
        Atencion.objects.annotate(y=ExtractYear("fecha"))
        .values("y")
        .annotate(n=Count("id"))
        .order_by("-y")
        .values_list("y", "n")
    )
