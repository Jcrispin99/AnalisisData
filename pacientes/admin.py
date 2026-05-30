import io
import unicodedata
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from django import forms
from django.contrib import admin, messages
from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render
from django.urls import reverse

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from unfold.admin import ModelAdmin
from unfold.decorators import action

from .models import Atencion, AtencionNutricion, Paciente


EXCEL_COLUMNS = [
    ("id_rh", "Id RH"),
    ("nombre_completo", "Nombre Completo"),
    ("fec_ingreso", "Fec. Ingreso"),
    ("nro_documento", "Nº de documento"),
    ("convenio", "Convenio"),
    ("nombre_posicion", "Nombre posición"),
    ("nombre_area", "Nombre de Area"),
    ("nombre_unidad_org", "Nombre unidad org."),
    ("fec_nacimiento", "Fec. Nacimiento"),
    ("provincia", "Provincia"),
    ("distrito", "Distrito"),
    ("direccion", "Dirección"),
]

DATE_FIELDS = {"fec_ingreso", "fec_nacimiento"}


ATENCION_COLUMNS = [
    ("nombre_completo", "Nombre Completo"),
    ("dni", "DNI Paciente"),
    ("fecha", "Fecha"),
    ("peso", "Peso (kg)"),
    ("talla", "Talla (cm)"),
    ("imc", "IMC"),
    ("sistolica", "Sistólica"),
    ("diastolica", "Diastólica"),
    ("abdominal", "Abdominal (cm)"),
    ("icc", "ICC"),
    ("colesterol_total", "Colesterol Total"),
    ("hdl", "HDL Colesterol"),
    ("ldl", "LDL Colesterol"),
    ("vldl", "VLDL Colesterol"),
    ("trigliceridos", "Triglicéridos"),
    ("hemoglobina", "Hemoglobina"),
    ("hematocrito", "Hematocrito"),
    ("glucosa", "Glucosa"),
    ("hb_a1c", "HB A1c"),
]

ATENCION_NUTRICION_COLUMNS = [
    ("nombre_completo", "Nombre Completo"),
    ("dni", "DNI Paciente"),
    ("fecha", "Fecha"),
    ("peso", "Peso (kg)"),
    ("talla", "Talla (cm)"),
    ("imc", "IMC"),
    ("grasa_corporal", "% Grasa corporal"),
    ("grasa_visceral", "% Grasa visceral"),
    ("masa_muscular", "% Masa muscular"),
]

PACIENTE_LOOKUP_FIELDS = {"dni", "nombre_completo"}


class ExcelUploadForm(forms.Form):
    archivo = forms.FileField(
        label="Archivo Excel (.xlsx)",
        help_text="Use la plantilla descargada para conservar los nombres de columna.",
    )


def _parse_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError(f"fecha no reconocida: {value!r}")


def _parse_decimal(value):
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value).strip().replace(",", "."))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"número no reconocido: {value!r}") from exc


def _build_template_response(columns, sheet_name, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")

    for col_idx, (_, label) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = header_font
        cell.fill = header_fill
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(label) + 4, 18)

    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _normalize_name(s):
    """Normaliza un nombre para comparación: colapsa espacios, quita acentos, MAYÚSCULAS.

    "  José  García  " → "JOSE GARCIA"
    """
    if not s:
        return ""
    s = " ".join(str(s).split())
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()


def _build_pacientes_index():
    """Devuelve dict {nombre_normalizado: [pacientes]} para lookup acento-insensible."""
    index = {}
    for p in Paciente.objects.only("id", "nro_documento", "nombre_completo"):
        index.setdefault(_normalize_name(p.nombre_completo), []).append(p)
    return index


def _buscar_paciente(dni, nombre, pacientes_by_norm):
    """Resuelve un Paciente por DNI (prioridad) o Nombre Completo.

    Reglas:
    - Si hay DNI: debe coincidir exactamente (no hay fallback silencioso al nombre).
    - Si no hay DNI pero sí nombre: comparación ignora acentos, mayúsculas y espacios extra.
      Falla si hay 0 o >1 coincidencias.

    Devuelve (paciente_o_None, mensaje_de_error_o_None).
    """
    if dni:
        try:
            return Paciente.objects.get(nro_documento=dni), None
        except Paciente.DoesNotExist:
            return None, f"paciente con DNI {dni!r} no existe."
        except Paciente.MultipleObjectsReturned:
            return None, f"DNI {dni!r} corresponde a más de un paciente."

    if not nombre:
        return None, "se requiere 'Nombre Completo' o 'DNI Paciente'."

    matches = pacientes_by_norm.get(_normalize_name(nombre), [])
    nombre_visible = " ".join(nombre.split())
    if not matches:
        return None, f"paciente con nombre {nombre_visible!r} no existe."
    if len(matches) > 1:
        return None, (
            f"nombre {nombre_visible!r} es ambiguo: {len(matches)} pacientes coinciden. "
            "Agregue el DNI para desambiguar."
        )
    return matches[0], None


def _procesar_excel_atenciones(archivo, model, columns):
    """Upsert masivo de atenciones a partir de un xlsx con la plantilla dada.

    Vincula al paciente por Nombre Completo o DNI (ver `_buscar_paciente`).
    Clave natural del upsert: (paciente, fecha). Devuelve (creados, actualizados, errores).
    """
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    headers = next(rows, None)
    if not headers:
        raise ValueError("El archivo está vacío.")

    label_to_field = {label: field for field, label in columns}
    try:
        field_order = [label_to_field[h] for h in headers if h is not None]
    except KeyError as exc:
        raise ValueError(
            f"Columna desconocida: {exc.args[0]!r}. Use la plantilla descargada."
        ) from exc

    decimal_fields = {f for f, _ in columns} - PACIENTE_LOOKUP_FIELDS - {"fecha"}

    pacientes_by_norm = _build_pacientes_index()

    creados = actualizados = 0
    errores = []

    for row_num, row in enumerate(rows, start=2):
        if not row or not any(cell not in (None, "") for cell in row):
            continue

        data = {}
        row_ok = True
        for field_name, value in zip(field_order, row):
            if field_name == "fecha":
                try:
                    data["fecha"] = _parse_date(value)
                except ValueError as exc:
                    errores.append(f"Fila {row_num} (Fecha): {exc}")
                    row_ok = False
                    break
            elif field_name in PACIENTE_LOOKUP_FIELDS:
                data[field_name] = "" if value is None else str(value).strip()
            elif field_name in decimal_fields:
                try:
                    data[field_name] = _parse_decimal(value)
                except ValueError as exc:
                    errores.append(f"Fila {row_num} ({field_name}): {exc}")
                    row_ok = False
                    break

        if not row_ok:
            continue

        dni = data.pop("dni", "")
        nombre = data.pop("nombre_completo", "")
        paciente, error = _buscar_paciente(dni, nombre, pacientes_by_norm)
        if error:
            errores.append(f"Fila {row_num}: {error}")
            continue

        fecha = data.pop("fecha", None)
        if not fecha:
            errores.append(f"Fila {row_num}: 'Fecha' vacía, fila omitida.")
            continue

        _, created = model.objects.update_or_create(
            paciente=paciente, fecha=fecha, defaults=data
        )
        if created:
            creados += 1
        else:
            actualizados += 1

    return creados, actualizados, errores


@admin.register(Paciente)
class PacienteAdmin(ModelAdmin):
    list_display = (
        "id_rh",
        "nombre_completo",
        "nro_documento",
        "convenio",
        "nombre_area",
        "fec_ingreso",
    )
    list_filter = ("convenio", "nombre_area", "provincia", "distrito")
    search_fields = ("id_rh", "nombre_completo", "nro_documento")
    ordering = ("nombre_completo",)
    fieldsets = (
        ("Identificación", {
            "fields": (
                "id_rh",
                "nombre_completo",
                "nro_documento",
                "fec_nacimiento",
                "sexo",
                "instruccion",
            ),
        }),
        ("Información laboral", {
            "fields": (
                "fec_ingreso",
                "convenio",
                "nombre_posicion",
                "nombre_area",
                "nombre_unidad_org",
            ),
        }),
        ("Ubicación", {
            "fields": ("provincia", "distrito", "direccion"),
        }),
    )

    actions_list = ("descargar_plantilla", "subir_excel")

    @action(description="Descargar plantilla", url_path="descargar-plantilla")
    def descargar_plantilla(self, request):
        return _build_template_response(
            EXCEL_COLUMNS, "Pacientes", "plantilla_pacientes.xlsx"
        )

    @action(description="Subir Excel", url_path="subir-excel")
    def subir_excel(self, request):
        cancel_url = reverse("admin:pacientes_paciente_changelist")
        if request.method == "POST":
            form = ExcelUploadForm(request.POST, request.FILES)
            if form.is_valid():
                try:
                    creados, actualizados, errores = self._procesar_excel(
                        form.cleaned_data["archivo"]
                    )
                except Exception as exc:
                    messages.error(request, f"Error procesando el archivo: {exc}")
                    return HttpResponseRedirect(request.path)

                for err in errores[:10]:
                    messages.warning(request, err)
                if len(errores) > 10:
                    messages.warning(request, f"... y {len(errores) - 10} errores más.")

                messages.success(
                    request,
                    f"Carga completa. Creados: {creados}, actualizados: {actualizados}, "
                    f"con errores: {len(errores)}.",
                )
                return HttpResponseRedirect(cancel_url)
        else:
            form = ExcelUploadForm()

        context = {
            **self.admin_site.each_context(request),
            "form": form,
            "title": "Subir Pacientes desde Excel",
            "opts": self.model._meta,
            "columnas": [label for _, label in EXCEL_COLUMNS],
            "cancel_url": cancel_url,
            "key_fields_help": (
                "El campo <strong>Id RH</strong> es obligatorio y se usa para "
                "identificar al registro (si ya existe, se actualiza)."
            ),
        }
        return render(request, "admin/pacientes/subir_excel.html", context)

    def _procesar_excel(self, archivo):
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active

        rows = ws.iter_rows(values_only=True)
        headers = next(rows, None)
        if not headers:
            raise ValueError("El archivo está vacío.")

        label_to_field = {label: field for field, label in EXCEL_COLUMNS}
        try:
            field_order = [label_to_field[h] for h in headers if h is not None]
        except KeyError as exc:
            raise ValueError(
                f"Columna desconocida: {exc.args[0]!r}. Use la plantilla descargada."
            ) from exc

        creados = actualizados = 0
        errores = []

        for row_num, row in enumerate(rows, start=2):
            if not row or not any(cell not in (None, "") for cell in row):
                continue

            data = {}
            row_ok = True
            for field_name, value in zip(field_order, row):
                if field_name in DATE_FIELDS:
                    try:
                        data[field_name] = _parse_date(value)
                    except ValueError as exc:
                        errores.append(f"Fila {row_num} ({field_name}): {exc}")
                        row_ok = False
                        break
                else:
                    data[field_name] = "" if value is None else str(value).strip()

            if not row_ok:
                continue

            id_rh = data.get("id_rh", "").strip()
            if not id_rh:
                errores.append(f"Fila {row_num}: 'Id RH' vacío, fila omitida.")
                continue

            data.pop("id_rh")
            _, created = Paciente.objects.update_or_create(
                id_rh=id_rh, defaults=data
            )
            if created:
                creados += 1
            else:
                actualizados += 1

        return creados, actualizados, errores


class _AtencionImportExportMixin:
    """Acciones compartidas para descargar plantilla y subir Excel.

    Subclases deben definir: `import_columns`, `import_sheet_name`,
    `import_filename`, `import_title`.
    """

    import_columns = ()
    import_sheet_name = ""
    import_filename = ""
    import_title = ""

    actions_list = ("descargar_plantilla", "subir_excel")

    def _changelist_url(self):
        opts = self.model._meta
        return reverse(f"admin:{opts.app_label}_{opts.model_name}_changelist")

    @action(description="Descargar plantilla", url_path="descargar-plantilla")
    def descargar_plantilla(self, request):
        return _build_template_response(
            self.import_columns, self.import_sheet_name, self.import_filename
        )

    @action(description="Subir Excel", url_path="subir-excel")
    def subir_excel(self, request):
        cancel_url = self._changelist_url()
        if request.method == "POST":
            form = ExcelUploadForm(request.POST, request.FILES)
            if form.is_valid():
                try:
                    creados, actualizados, errores = _procesar_excel_atenciones(
                        form.cleaned_data["archivo"], self.model, self.import_columns
                    )
                except Exception as exc:
                    messages.error(request, f"Error procesando el archivo: {exc}")
                    return HttpResponseRedirect(request.path)

                for err in errores[:10]:
                    messages.warning(request, err)
                if len(errores) > 10:
                    messages.warning(request, f"... y {len(errores) - 10} errores más.")

                messages.success(
                    request,
                    f"Carga completa. Creados: {creados}, actualizados: {actualizados}, "
                    f"con errores: {len(errores)}.",
                )
                return HttpResponseRedirect(cancel_url)
        else:
            form = ExcelUploadForm()

        context = {
            **self.admin_site.each_context(request),
            "form": form,
            "title": self.import_title,
            "opts": self.model._meta,
            "columnas": [label for _, label in self.import_columns],
            "cancel_url": cancel_url,
            "key_fields_help": (
                "El paciente se identifica por <strong>Nombre Completo</strong> "
                "(la comparación ignora acentos, mayúsculas y espacios extra: "
                "<code>José Pérez</code> = <code>JOSE PEREZ</code>) o por "
                "<strong>DNI Paciente</strong>. Si llenas DNI, se usa como prioridad. "
                "Si el nombre coincide con más de un paciente, agrega el DNI para "
                "desambiguar. <strong>Fecha</strong> es obligatoria y forma la clave "
                "única del registro junto con el paciente (si ya existe una atención "
                "de ese paciente en esa fecha, se actualiza)."
            ),
        }
        return render(request, "admin/pacientes/subir_excel.html", context)


@admin.register(Atencion)
class AtencionAdmin(_AtencionImportExportMixin, ModelAdmin):
    autocomplete_fields = ("paciente",)
    date_hierarchy = "fecha"
    list_display = (
        "fecha",
        "paciente_dni",
        "paciente_nombre",
        "peso",
        "talla",
        "imc",
        "sistolica",
        "diastolica",
        "glucosa",
        "hb_a1c",
    )
    list_filter = ("fecha", "paciente__convenio", "paciente__nombre_area")
    search_fields = (
        "paciente__nro_documento",
        "paciente__nombre_completo",
        "paciente__id_rh",
    )
    ordering = ("-fecha",)
    fieldsets = (
        ("Paciente", {
            "fields": ("paciente", "fecha"),
        }),
        ("Antropometría", {
            "fields": ("peso", "talla", "imc", "abdominal", "icc"),
        }),
        ("Presión arterial", {
            "fields": ("sistolica", "diastolica"),
        }),
        ("Perfil lipídico", {
            "fields": ("colesterol_total", "hdl", "ldl", "vldl", "trigliceridos"),
        }),
        ("Hematología y glucosa", {
            "fields": ("hemoglobina", "hematocrito", "glucosa", "hb_a1c"),
        }),
    )

    import_columns = ATENCION_COLUMNS
    import_sheet_name = "Atenciones"
    import_filename = "plantilla_atenciones.xlsx"
    import_title = "Subir Atenciones desde Excel"

    @admin.display(description="DNI", ordering="paciente__nro_documento")
    def paciente_dni(self, obj):
        return obj.paciente.nro_documento

    @admin.display(description="Paciente", ordering="paciente__nombre_completo")
    def paciente_nombre(self, obj):
        return obj.paciente.nombre_completo


@admin.register(AtencionNutricion)
class AtencionNutricionAdmin(_AtencionImportExportMixin, ModelAdmin):
    autocomplete_fields = ("paciente",)
    date_hierarchy = "fecha"
    list_display = (
        "fecha",
        "paciente_dni",
        "paciente_nombre",
        "peso",
        "talla",
        "imc",
        "grasa_corporal",
        "grasa_visceral",
        "masa_muscular",
    )
    list_filter = ("fecha", "paciente__convenio", "paciente__nombre_area")
    search_fields = (
        "paciente__nro_documento",
        "paciente__nombre_completo",
        "paciente__id_rh",
    )
    ordering = ("-fecha",)
    fieldsets = (
        ("Paciente", {
            "fields": ("paciente", "fecha"),
        }),
        ("Antropometría", {
            "fields": ("peso", "talla", "imc"),
        }),
        ("Composición corporal", {
            "fields": ("grasa_corporal", "grasa_visceral", "masa_muscular"),
        }),
    )

    import_columns = ATENCION_NUTRICION_COLUMNS
    import_sheet_name = "Nutrición"
    import_filename = "plantilla_atenciones_nutricion.xlsx"
    import_title = "Subir Atenciones de Nutrición desde Excel"

    @admin.display(description="DNI", ordering="paciente__nro_documento")
    def paciente_dni(self, obj):
        return obj.paciente.nro_documento

    @admin.display(description="Paciente", ordering="paciente__nombre_completo")
    def paciente_nombre(self, obj):
        return obj.paciente.nombre_completo

