"""Estimación de altitud (msnm) para ubicaciones del Perú.

Sólo se necesita determinar si el paciente reside a más o menos de 2500 msnm
(umbral para criterios de eritrocitosis), no la altitud exacta.

Fuente: capitales departamentales / distritos comunes según INEI y referencias
geográficas estándar. Si una ubicación no aparece aquí, se asume <2500 msnm.
"""
from __future__ import annotations

import unicodedata


def _norm(s: str | None) -> str:
    if not s:
        return ""
    s = " ".join(str(s).split())
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()


# Departamentos predominantemente de sierra (>2500 msnm en su mayoría).
# Se usa como fallback si el distrito/provincia no está en el dataset.
_DEPARTAMENTOS_SIERRA = {
    "APURIMAC",
    "AYACUCHO",
    "CUSCO",
    "HUANCAVELICA",
    "PASCO",
    "PUNO",
}

# Distritos/provincias específicos con altitud aproximada (msnm).
# Lista no exhaustiva, ampliable según los datos reales que entren al sistema.
_LOCATION_ALTITUDES: dict[str, int] = {
    # Costa (referencia)
    "LIMA": 154,
    "CALLAO": 16,
    "TRUJILLO": 34,
    "CHICLAYO": 27,
    "PIURA": 29,
    "TUMBES": 5,
    "ICA": 406,
    "CHIMBOTE": 4,
    "TACNA": 562,
    "MOQUEGUA": 1410,
    # Selva
    "IQUITOS": 106,
    "PUCALLPA": 154,
    "TARAPOTO": 356,
    "MADRE DE DIOS": 186,
    "PUERTO MALDONADO": 186,
    # Sierra (>2500 mayormente)
    "AREQUIPA": 2335,
    "CUSCO": 3399,
    "HUANCAYO": 3271,
    "HUARAZ": 3052,
    "CAJAMARCA": 2750,
    "PUNO": 3827,
    "JULIACA": 3825,
    "AYACUCHO": 2761,
    "HUANCAVELICA": 3676,
    "ABANCAY": 2378,
    "CERRO DE PASCO": 4330,
    "HUANUCO": 1894,
    "TINGO MARIA": 660,
    "CHACHAPOYAS": 2335,
    "JAEN": 729,
    "JULI": 3870,
    "LA OROYA": 3745,
}


def altitud_estimada(distrito: str | None = None,
                     provincia: str | None = None,
                     departamento: str | None = None) -> int | None:
    """Devuelve una altitud aproximada en msnm, o None si no se puede estimar.

    Busca en orden: distrito → provincia → capital del departamento → fallback
    por departamento de sierra (devuelve 3000 si está en sierra, sino None).
    """
    for clave in (distrito, provincia, departamento):
        key = _norm(clave)
        if key and key in _LOCATION_ALTITUDES:
            return _LOCATION_ALTITUDES[key]

    if _norm(departamento) in _DEPARTAMENTOS_SIERRA:
        return 3000  # promedio sierra alta, suficiente para superar 2500
    return None


def es_altura(altitud_msnm: int | None) -> bool:
    """True si la altitud supera el umbral de 2500 msnm."""
    return altitud_msnm is not None and altitud_msnm > 2500
