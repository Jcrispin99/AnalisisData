"""Clasificaciones clínicas según CRITERIOS DE ESTUDIO.xlsx.

La lógica de cortes vive en `umbrales.py` (single source of truth para
clasificación + tooltips). Aquí solo se exponen funciones con nombre
estable que delegan al módulo de umbrales.
"""
from __future__ import annotations

from typing import Optional

from . import umbrales as u
from .umbrales import Numeric, _f, clasificar_por_bandas, clasificar_por_sexo


# ----- DISLIPIDEMIAS -----

def clasificar_colesterol_total(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.COLESTEROL_TOTAL)


def clasificar_hdl(valor: Numeric, sexo: Optional[str]) -> Optional[str]:
    return clasificar_por_sexo(valor, sexo, u.HDL)


def clasificar_ldl(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.LDL)


def clasificar_trigliceridos(valor: Numeric) -> Optional[str]:
    """Asume valor en ayunas (la hoja distingue post-prandial pero el modelo no lo guarda)."""
    return clasificar_por_bandas(valor, u.TRIGLICERIDOS)


def clasificar_colesterol_no_hdl(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.COLESTEROL_NO_HDL)


def sospecha_hipercolesterolemia_familiar(ldl: Numeric) -> bool:
    v = _f(ldl)
    return v is not None and v >= u.LDL_UMBRAL_HF


_ORDEN_SEVERIDAD = {
    "Deseable": 0,
    "Dislipidemia leve": 1,
    "Dislipidemia moderada": 2,
    "Dislipidemia severa": 3,
}
_SEVERIDAD_LABEL = {1: "leve", 2: "moderada", 3: "severa"}


def analizar_dislipidemia(
    colesterol_total: Numeric,
    ldl: Numeric,
    hdl: Numeric,
    trigliceridos: Numeric,
    sexo: Optional[str],
) -> Optional[dict]:
    """Diagnóstico combinado del perfil lipídico.

    Devuelve dict con `tipo`, `severidad`, `label`, `sospecha_familiar` o None
    si no hay ningún componente con dato.

    Lógica de tipos (referencia: MSD/Merck, ACC/AHA 2018/2026):
      - Perfil normal: todo deseable.
      - HDL bajo aislado: solo HDL bajo.
      - Hipercolesterolemia: CT o LDL alterado, TG normal.
      - Hipertrigliceridemia: TG alterado, CT y LDL normales.
      - Dislipidemia mixta: CT/LDL alterado + TG alterado.
      - Dislipidemia aterogénica: TG alterado + HDL bajo (típico síndrome metabólico).
    """
    ct = clasificar_colesterol_total(colesterol_total)
    ldl_c = clasificar_ldl(ldl)
    hdl_c = clasificar_hdl(hdl, sexo)
    tg = clasificar_trigliceridos(trigliceridos)

    if all(c is None for c in (ct, ldl_c, hdl_c, tg)):
        return None

    col_alterado = (ct and ct != "Deseable") or (ldl_c and ldl_c != "Deseable")
    tg_alterado = bool(tg and tg != "Deseable")
    hdl_bajo = bool(hdl_c and hdl_c != "Deseable")

    niveles = [
        _ORDEN_SEVERIDAD.get(c, 0)
        for c in (ct, ldl_c, hdl_c, tg)
        if c is not None
    ]
    nivel = max(niveles) if niveles else 0
    severidad = _SEVERIDAD_LABEL.get(nivel)

    sospecha_familiar = sospecha_hipercolesterolemia_familiar(ldl)

    if not col_alterado and not tg_alterado and not hdl_bajo:
        tipo = "Normal"
        label = "Perfil lipídico normal"
        severidad = None
    elif tg_alterado and hdl_bajo:
        # Aterogénica prevalece sobre mixta: TG alto + HDL bajo es el patrón distintivo.
        tipo = "Dislipidemia aterogénica"
        label = "Dislipidemia aterogénica"
        severidad = None
    elif col_alterado and tg_alterado:
        tipo = "Dislipidemia mixta"
        label = f"Dislipidemia mixta {severidad}" if severidad else tipo
    elif col_alterado:
        tipo = "Hipercolesterolemia"
        label = f"Hipercolesterolemia {severidad}" if severidad else tipo
    elif tg_alterado:
        tipo = "Hipertrigliceridemia"
        label = f"Hipertrigliceridemia {severidad}" if severidad else tipo
    else:
        # Solo HDL bajo.
        tipo = "HDL bajo aislado"
        label = "HDL bajo aislado"
        severidad = None

    return {
        "tipo": tipo,
        "severidad": severidad,
        "label": label,
        "sospecha_familiar": sospecha_familiar,
    }


# ----- ERITROCITOSIS -----

def clasificar_eritrocitosis(
    hemoglobina: Numeric, sexo: Optional[str], es_altura: bool
) -> Optional[str]:
    """Clasifica según Hb, sexo y si el paciente vive a >2500 msnm."""
    v = _f(hemoglobina)
    if v is None or sexo not in ("M", "F"):
        return None
    t = u.ERITROCITOSIS[(sexo, bool(es_altura))]
    if v < t["leve_min"]:
        return "Normal"
    if v <= t["leve_max"]:
        return "Eritrocitosis leve"
    if v <= t["moderada_max"]:
        return "Eritrocitosis moderada"
    return "Eritrocitosis severa"


# ----- OBESIDAD ABDOMINAL -----

def clasificar_obesidad_abdominal(
    perimetro_cm: Numeric, sexo: Optional[str]
) -> Optional[str]:
    return clasificar_por_sexo(perimetro_cm, sexo, u.OBESIDAD_ABDOMINAL)


# ----- DIABETES MELLITUS -----

def clasificar_glicemia_ayunas(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.GLICEMIA_AYUNAS)


def clasificar_hb_a1c(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.HB_A1C)


# ----- HIPERTENSIÓN ARTERIAL -----

def clasificar_presion(sistolica: Numeric, diastolica: Numeric) -> Optional[str]:
    """Devuelve la categoría más severa entre sistólica y diastólica."""
    s = _f(sistolica)
    d = _f(diastolica)
    if s is None and d is None:
        return None
    p = u.PRESION

    if (s is not None and s >= p["crisis_s"]) or (d is not None and d >= p["crisis_d"]):
        return "Crisis Hipertensiva"
    if (s is not None and s >= p["g2_s"]) or (d is not None and d >= p["g2_d"]):
        return "Hipertensión Grado 2"
    if (s is not None and p["g1_s_min"] <= s <= p["g1_s_max"]) or \
       (d is not None and p["g1_d_min"] <= d <= p["g1_d_max"]):
        return "Hipertensión Grado 1"
    if s is not None and p["elevada_s_min"] <= s <= p["elevada_s_max"] and \
       (d is None or d < p["elevada_d_max"]):
        return "Elevada"
    return "Normal Alta"


# ----- COMPOSICIÓN CORPORAL (Opción A: referencias estándar) -----
# Las celdas del xlsx vinieron vacías; usamos:
#   IMC: OMS.
#   % Grasa corporal: ACSM por sexo (sin ajuste por edad).
#   % Grasa visceral: rating tipo Tanita (1-59).
#   % Masa muscular: rangos Tanita por sexo (sin ajuste por edad).

def clasificar_imc(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.IMC)


def clasificar_grasa_corporal(valor: Numeric, sexo: Optional[str]) -> Optional[str]:
    return clasificar_por_sexo(valor, sexo, u.GRASA_CORPORAL)


def clasificar_grasa_visceral(valor: Numeric) -> Optional[str]:
    return clasificar_por_bandas(valor, u.GRASA_VISCERAL)


def clasificar_masa_muscular(valor: Numeric, sexo: Optional[str]) -> Optional[str]:
    return clasificar_por_sexo(valor, sexo, u.MASA_MUSCULAR)


# Categorías "saludables" por cada clasificación que aporta al SM. Si la
# clasificación cae fuera de este set, el componente cuenta como alterado.
_SM_BUENAS = {
    "obesidad_abdominal": {"Normal"},
    "hdl": {"Deseable"},
    "trigliceridos": {"Deseable"},
    "glicemia": {"Normal"},
    "presion": {"Normal Alta"},
}


def clasificar_sindrome_metabolico(
    abdominal: Numeric,
    sexo: Optional[str],
    hdl: Numeric,
    trigliceridos: Numeric,
    glucosa: Numeric,
    sistolica: Numeric,
    diastolica: Numeric,
) -> Optional[str]:
    """Diagnóstico simple de síndrome metabólico reutilizando las
    clasificaciones existentes. Cuenta cuántos de los 5 componentes son
    "no saludables" (cualquier categoría fuera de _SM_BUENAS).

    Retorna:
      - "Con SM"  si ≥3 componentes alterados
      - "Sin SM"  si <3 alterados y al menos 3 componentes medidos
      - None      si <3 componentes medidos (cae al contador "sin dato")
    """
    componentes = [
        ("obesidad_abdominal", clasificar_obesidad_abdominal(abdominal, sexo)),
        ("hdl", clasificar_hdl(hdl, sexo)),
        ("trigliceridos", clasificar_trigliceridos(trigliceridos)),
        ("glicemia", clasificar_glicemia_ayunas(glucosa)),
        ("presion", clasificar_presion(sistolica, diastolica)),
    ]
    medidos = 0
    alterados = 0
    for clave, valor in componentes:
        if valor is None:
            continue
        medidos += 1
        if valor not in _SM_BUENAS[clave]:
            alterados += 1
    if medidos < 3:
        return None
    return "Con SM" if alterados >= 3 else "Sin SM"
