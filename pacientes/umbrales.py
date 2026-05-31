"""Single source of truth para cortes clínicos y tooltips del reporte.

Modelo
------
Cada criterio se describe como una lista ordenada de `Banda(label, op, valor)`:
- `op` es uno de `<`, `<=`, `>`, `>=`. Significa "si v op valor → label".
- La última banda usa `op=None` (catch-all).

A partir de esa lista se derivan:
- la clasificación (`clasificar_por_bandas`),
- el HTML del tooltip (`render_tooltip` / `render_tooltip_por_sexo`).

Para los criterios que no son simples cortes (presión arterial, eritrocitosis,
diagnóstico de dislipidemia) las constantes se exponen como dicts y los tooltips
se construyen con f-strings que leen esas mismas constantes, así no pueden
desincronizarse.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional, Tuple, Union

Numeric = Union[int, float, Decimal, None]


@dataclass(frozen=True)
class Banda:
    label: str
    op: Optional[str] = None  # "<", "<=", ">", ">=", o None para catch-all
    valor: Optional[float] = None


@dataclass(frozen=True)
class Criterio:
    """Criterio de clasificación basado en una lista de bandas.

    `estilo` controla cómo se imprime el límite inferior de las bandas medias
    cuando la banda previa fue inclusiva (`<=` o `>=`):
    - "overlap": muestra el mismo número (e.g. CT: "200-240, 240-300").
    - "gap": suma/resta `step` para mostrar un valor sin solapamiento
      (e.g. eritrocitosis: "Leve 18.1-19.5, Moderada 19.6-21").
    """
    bandas: Tuple[Banda, ...]
    unidad: str = ""  # sufijo del primer umbral, e.g. " mg/dL" o "%"
    step: float = 1   # incremento de display para bandas con op "<" o ">"
    estilo: str = "overlap"
    fuente_html: str = "Fuente: CRITERIOS DE ESTUDIO"
    nota_pre: str = ""
    nota_post: str = ""
    notas_banda: dict = field(default_factory=dict)
    display_labels: dict = field(default_factory=dict)
    lower_finite: Optional[float] = None  # mostrado como inicio del primer rango


@dataclass(frozen=True)
class CriterioPorSexo:
    m: Criterio
    f: Criterio


def _f(v: Numeric) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def clasificar_por_bandas(valor: Numeric, criterio: Criterio) -> Optional[str]:
    v = _f(valor)
    if v is None:
        return None
    for b in criterio.bandas:
        if b.op is None:
            return b.label
        if b.op == "<" and v < b.valor:
            return b.label
        if b.op == "<=" and v <= b.valor:
            return b.label
        if b.op == ">" and v > b.valor:
            return b.label
        if b.op == ">=" and v >= b.valor:
            return b.label
    return None


def clasificar_por_sexo(valor: Numeric, sexo: Optional[str], criterio: CriterioPorSexo) -> Optional[str]:
    if sexo == "M":
        return clasificar_por_bandas(valor, criterio.m)
    if sexo == "F":
        return clasificar_por_bandas(valor, criterio.f)
    return None


# ----- Render de tooltips -----

_OP_SYM = {"<": "<", "<=": "≤", ">": ">", ">=": "≥"}


def _fmt(n: float, step: float) -> str:
    if step >= 1 and float(n).is_integer():
        return str(int(n))
    # Recorta ceros sobrantes: 19.0 → "19", 19.5 → "19.5".
    return f"{n:g}"


def _direccion(criterio: Criterio) -> str:
    """'asc' si las bandas usan '<' / '<=', 'desc' si usan '>' / '>='."""
    for b in criterio.bandas:
        if b.op in ("<", "<="):
            return "asc"
        if b.op in (">", ">="):
            return "desc"
    return "asc"


def _linea_banda(criterio: Criterio, i: int, direccion: str) -> str:
    """Devuelve el texto plano de una banda, e.g. '<200 mg/dL — Deseable'."""
    b = criterio.bandas[i]
    is_first = i == 0
    is_catch = b.op is None
    prev = criterio.bandas[i - 1] if i > 0 else None
    label = criterio.display_labels.get(b.label, b.label)
    nota = criterio.notas_banda.get(b.label, "")
    step = criterio.step
    estilo = criterio.estilo

    if is_catch:
        if direccion == "asc":
            if prev.op == "<":
                sym, val = "≥", prev.valor
            elif estilo == "gap":
                sym, val = "≥", prev.valor + step
            else:
                sym, val = ">", prev.valor
        else:
            if prev.op == ">":
                sym, val = "≤", prev.valor
            elif estilo == "gap":
                sym, val = "≤", prev.valor - step
            else:
                sym, val = "<", prev.valor
        return f"{sym}{_fmt(val, step)} — {label}{nota}"

    if is_first:
        unidad = criterio.unidad
        if criterio.lower_finite is not None and direccion == "asc":
            upper = b.valor if b.op == "<=" else b.valor - step
            return (
                f"{_fmt(criterio.lower_finite, step)}-{_fmt(upper, step)}"
                f"{unidad} — {label}{nota}"
            )
        return f"{_OP_SYM[b.op]}{_fmt(b.valor, step)}{unidad} — {label}{nota}"

    # Banda media: rango "lower-upper".
    if direccion == "asc":
        if estilo == "gap" and prev.op == "<=":
            lower = prev.valor + step
        else:
            lower = prev.valor
        upper = b.valor if b.op == "<=" else b.valor - step
    else:
        if estilo == "gap" and prev.op == ">=":
            upper = prev.valor - step
        else:
            upper = prev.valor
        lower = b.valor if b.op == ">=" else b.valor + step

    return f"{_fmt(lower, step)}-{_fmt(upper, step)} — {label}{nota}"


def _escape_op(linea: str) -> str:
    """Convierte '<X' / '>X' al inicio de la línea a entities HTML."""
    if linea.startswith("<") and not linea.startswith("<p") and not linea.startswith("<l"):
        return "&lt;" + linea[1:]
    if linea.startswith(">"):
        return "&gt;" + linea[1:]
    return linea


def _items(criterio: Criterio) -> str:
    direccion = _direccion(criterio)
    out = []
    for i in range(len(criterio.bandas)):
        out.append(f"<li>{_escape_op(_linea_banda(criterio, i, direccion))}</li>")
    return "".join(out)


def render_tooltip(titulo: str, criterio: Criterio) -> str:
    partes = [f"<p class='font-semibold mb-1'>{titulo}</p>"]
    if criterio.nota_pre:
        partes.append(criterio.nota_pre)
    partes.append(
        f"<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>"
        f"{criterio.fuente_html}</p>"
    )
    partes.append("<ul class='list-disc pl-4 space-y-0.5'>")
    partes.append(_items(criterio))
    partes.append("</ul>")
    if criterio.nota_post:
        partes.append(criterio.nota_post)
    return "".join(partes)


def render_tooltip_por_sexo(titulo: str, criterio: CriterioPorSexo,
                            *, fuente_html: Optional[str] = None,
                            nota_post: str = "") -> str:
    fuente = fuente_html if fuente_html is not None else criterio.m.fuente_html
    partes = [
        f"<p class='font-semibold mb-1'>{titulo} (depende del sexo)</p>",
        f"<p class='text-font-subtle-light dark:text-font-subtle-dark mb-1.5'>"
        f"{fuente}</p>",
        "<p class='mb-1'><b>Varones:</b></p>",
        "<ul class='list-disc pl-4 space-y-0.5 mb-1.5'>",
        _items(criterio.m),
        "</ul>",
        "<p class='mb-1'><b>Mujeres:</b></p>",
        "<ul class='list-disc pl-4 space-y-0.5'>",
        _items(criterio.f),
        "</ul>",
    ]
    if nota_post:
        partes.append(nota_post)
    return "".join(partes)


# =========================================================================
# Definiciones de criterios
# =========================================================================

# ----- Dislipidemias -----

_DISLIPIDEMIA_LABELS = {
    "Deseable": "Deseable",
    "Dislipidemia leve": "Leve",
    "Dislipidemia moderada": "Moderada",
    "Dislipidemia severa": "Severa",
}

COLESTEROL_TOTAL = Criterio(
    bandas=(
        Banda("Deseable", "<", 200),
        Banda("Dislipidemia leve", "<=", 240),
        Banda("Dislipidemia moderada", "<=", 300),
        Banda("Dislipidemia severa"),
    ),
    unidad=" mg/dL",
    display_labels=_DISLIPIDEMIA_LABELS,
)

LDL = Criterio(
    bandas=(
        Banda("Deseable", "<", 100),
        Banda("Dislipidemia leve", "<=", 150),
        Banda("Dislipidemia moderada", "<=", 190),
        Banda("Dislipidemia severa"),
    ),
    unidad=" mg/dL",
    display_labels=_DISLIPIDEMIA_LABELS,
    notas_banda={"Dislipidemia severa": " ⚠ sospecha hipercolesterolemia familiar"},
)

TRIGLICERIDOS = Criterio(
    bandas=(
        Banda("Deseable", "<", 150),
        Banda("Dislipidemia leve", "<=", 300),
        Banda("Dislipidemia moderada", "<=", 500),
        Banda("Dislipidemia severa"),
    ),
    unidad=" mg/dL",
    display_labels=_DISLIPIDEMIA_LABELS,
)

HDL = CriterioPorSexo(
    m=Criterio(
        bandas=(
            Banda("Deseable", ">", 40),
            Banda("Dislipidemia leve", ">=", 35),
            Banda("Dislipidemia moderada", ">=", 30),
            Banda("Dislipidemia severa"),
        ),
        unidad=" mg/dL",
        display_labels=_DISLIPIDEMIA_LABELS,
    ),
    f=Criterio(
        bandas=(
            Banda("Deseable", ">", 50),
            Banda("Dislipidemia leve", ">=", 45),
            Banda("Dislipidemia moderada", ">=", 40),
            Banda("Dislipidemia severa"),
        ),
        unidad=" mg/dL",
        display_labels=_DISLIPIDEMIA_LABELS,
    ),
)

COLESTEROL_NO_HDL = Criterio(
    bandas=(
        Banda("Deseable", "<", 130),
        Banda("Limítrofe alto", "<", 160),
        Banda("Riesgo aumentado", "<", 190),
        Banda("Riesgo muy alto"),
    ),
    unidad=" mg/dL",
    fuente_html=(
        "Estándar ATP III / AHA (LDL target + 30 mg/dL). No está en el xlsx."
    ),
    nota_pre="<p class='mb-1.5'>Fórmula: <b>Colesterol Total − HDL</b></p>",
)

# Umbral usado por sospecha_hipercolesterolemia_familiar (ACC/AHA).
LDL_UMBRAL_HF = 190


# ----- Eritrocitosis -----
# Hb en g/dL. La banda Normal es implícita (< leve_min).
# Tabla literal del xlsx: rangos por sexo × altitud.
ERITROCITOSIS = {
    ("F", False): {"leve_min": 16.1, "leve_max": 17.5, "moderada_max": 19.0},
    ("F", True):  {"leve_min": 17.5, "leve_max": 19.0, "moderada_max": 20.5},
    ("M", False): {"leve_min": 18.1, "leve_max": 19.5, "moderada_max": 21.0},
    ("M", True):  {"leve_min": 19.5, "leve_max": 21.0, "moderada_max": 22.5},
}
ERITROCITOSIS_STEP = 0.1


def _eritro_linea(t):
    """'Leve 18.1–19.5 · Moderada 19.6–21 · Severa &gt;21' a partir del dict."""
    s = ERITROCITOSIS_STEP
    leve = f"{_fmt(t['leve_min'], s)}–{_fmt(t['leve_max'], s)}"
    mod_lower = t["leve_max"] + s
    moderada = f"{_fmt(mod_lower, s)}–{_fmt(t['moderada_max'], s)}"
    severa = f"&gt;{_fmt(t['moderada_max'], s)}"
    return f"Leve {leve} · Moderada {moderada} · Severa {severa}"


# ----- Obesidad abdominal -----

OBESIDAD_ABDOMINAL = CriterioPorSexo(
    m=Criterio(
        bandas=(
            Banda("Normal", "<", 94),
            Banda("Grado 1 (riesgo aumentado)", "<=", 102),
            Banda("Grado 2 (riesgo significativo)"),
        ),
        unidad=" cm",
        display_labels={
            "Grado 1 (riesgo aumentado)": "Grado 1 (riesgo aumentado)",
            "Grado 2 (riesgo significativo)": "Grado 2 (riesgo significativo)",
        },
    ),
    f=Criterio(
        bandas=(
            Banda("Normal", "<", 88),
            Banda("Grado 1 (riesgo aumentado)", "<=", 90),
            Banda("Grado 2 (riesgo significativo)"),
        ),
        unidad=" cm",
    ),
)


# ----- Diabetes Mellitus -----

GLICEMIA_AYUNAS = Criterio(
    bandas=(
        Banda("Normal", "<", 100),
        Banda("Prediabetes", "<", 126),
        Banda("Diabetes"),
    ),
    unidad=" mg/dL",
)

HB_A1C = Criterio(
    bandas=(
        Banda("Normal", "<", 5.6),
        Banda("Prediabetes", "<", 6.5),
        Banda("Diabetes"),
    ),
    unidad="%",
    step=0.1,
)


# ----- Hipertensión arterial -----
# Decisión multi-variable (sistólica/diastólica con AND/OR). No es banda.
PRESION = {
    "elevada_s_min": 120, "elevada_s_max": 129, "elevada_d_max": 80,  # d < 80
    "g1_s_min": 130, "g1_s_max": 139, "g1_d_min": 80, "g1_d_max": 89,
    "g2_s": 140, "g2_d": 90,
    "crisis_s": 180, "crisis_d": 120,
}


# ----- IMC y composición corporal -----

IMC = Criterio(
    bandas=(
        Banda("Bajo peso", "<", 18.5),
        Banda("Normal", "<", 25),
        Banda("Sobrepeso", "<", 30),
        Banda("Obesidad grado I", "<", 35),
        Banda("Obesidad grado II", "<", 40),
        Banda("Obesidad grado III"),
    ),
    step=0.1,
    fuente_html="Estándar OMS. El xlsx no define cortes para composición corporal.",
    nota_pre="<p class='mb-1.5'>Fórmula: <b>peso (kg) / talla² (m)</b></p>",
    display_labels={
        "Obesidad grado I": "Obesidad I",
        "Obesidad grado II": "Obesidad II",
        "Obesidad grado III": "Obesidad III",
    },
)

GRASA_CORPORAL = CriterioPorSexo(
    m=Criterio(
        bandas=(
            Banda("Bajo", "<", 6),
            Banda("Saludable", "<", 18),
            Banda("Aceptable", "<", 25),
            Banda("Obesidad"),
        ),
        unidad="%",
    ),
    f=Criterio(
        bandas=(
            Banda("Bajo", "<", 14),
            Banda("Saludable", "<", 25),
            Banda("Aceptable", "<", 32),
            Banda("Obesidad"),
        ),
        unidad="%",
    ),
)

GRASA_VISCERAL = Criterio(
    bandas=(
        Banda("Saludable", "<=", 12),
        Banda("Alta"),
    ),
    step=1,
    estilo="gap",
    lower_finite=1,
    fuente_html="Estándar Tanita. El xlsx no define cortes.",
)

MASA_MUSCULAR = CriterioPorSexo(
    m=Criterio(
        bandas=(
            Banda("Bajo", "<", 33),
            Banda("Normal", "<", 40),
            Banda("Alto", "<", 45),
            Banda("Muy alto"),
        ),
        unidad="%",
    ),
    f=Criterio(
        bandas=(
            Banda("Bajo", "<", 24),
            Banda("Normal", "<", 31),
            Banda("Alto", "<", 36),
            Banda("Muy alto"),
        ),
        unidad="%",
    ),
)
