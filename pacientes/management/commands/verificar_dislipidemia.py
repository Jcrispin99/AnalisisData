"""Verifica la consistencia del conteo de diagnósticos de dislipidemia.

Imprime en consola el desglose paso a paso de la población filtrada:
1. Sanity check de la fila "Triglicéridos" del reporte.
2. Desglose de los pacientes con TG alterado por diagnóstico combinado.
3. Auditoría de calidad de data en los pacientes clasificados como
   "Hipertrigliceridemia" (cuántos tienen valores faltantes que podrían
   estar inflando o desinflando ese conteo).
4. Universo completo del diagnóstico combinado.

Acepta los mismos filtros que la vista del reporte
(`/admin/pacientes/atencion/reporte/`).
"""
from collections import Counter

from django.core.management.base import BaseCommand

from pacientes.models import Atencion
from pacientes.reportes import _aplicar_filtros


class Command(BaseCommand):
    help = "Verifica el conteo de diagnósticos de dislipidemia y muestra el desglose."

    def add_arguments(self, parser):
        parser.add_argument("--year", type=int, default=None)
        parser.add_argument("--convenio", type=str, default=None)
        parser.add_argument("--area", type=str, default=None)
        parser.add_argument("--sexo", type=str, default=None, choices=["M", "F"])

    def handle(self, *args, **opts):
        qs = _aplicar_filtros(
            Atencion.objects.select_related("paciente"),
            opts["year"],
            opts["convenio"],
            opts["area"],
            opts["sexo"],
        )
        atenciones = list(qs)
        total = len(atenciones)

        filtros = {k: opts[k] for k in ("year", "convenio", "area", "sexo") if opts[k]}
        self.stdout.write(self.style.MIGRATE_HEADING("\n=== Verificación de dislipidemia ==="))
        self.stdout.write(f"Filtros aplicados: {filtros or 'ninguno'}")
        self.stdout.write(f"Total atenciones: {total}\n")

        self._bloque1_trigliceridos(atenciones)
        tg_alterados = [a for a in atenciones if self._tg_alterado(a)]
        self._bloque2_desglose_tg_alterado(tg_alterados)
        hipertri = [a for a in tg_alterados if a.clasif_dislipidemia_tipo == "Hipertrigliceridemia"]
        self._bloque3_auditoria_hipertri(hipertri)
        self._bloque4_universo(atenciones)

    # ---------- helpers ----------

    @staticmethod
    def _tg_alterado(at):
        c = at.clasif_trigliceridos
        return c is not None and c != "Deseable"

    def _linea(self, etiqueta, n, total=None, ancho=32):
        if total:
            pct = f"  ({n / total * 100:.1f}%)" if total else ""
            self.stdout.write(f"  {etiqueta.ljust(ancho)} {n:>6}{pct}")
        else:
            self.stdout.write(f"  {etiqueta.ljust(ancho)} {n:>6}")

    # ---------- bloques ----------

    def _bloque1_trigliceridos(self, atenciones):
        self.stdout.write(self.style.HTTP_INFO("Bloque 1 — Tabla Triglicéridos"))
        con_dato = sum(1 for a in atenciones if a.clasif_trigliceridos is not None)
        sin_dato = len(atenciones) - con_dato
        bandas = Counter(
            a.clasif_trigliceridos for a in atenciones if a.clasif_trigliceridos
        )
        self._linea("Con dato de TG", con_dato)
        self._linea("Sin dato de TG", sin_dato)
        for cat in ("Deseable", "Dislipidemia leve", "Dislipidemia moderada", "Dislipidemia severa"):
            self._linea(cat, bandas.get(cat, 0), con_dato)

        tg_alterado = sum(bandas[c] for c in bandas if c != "Deseable")
        suma_alterados = (
            bandas.get("Dislipidemia leve", 0)
            + bandas.get("Dislipidemia moderada", 0)
            + bandas.get("Dislipidemia severa", 0)
        )
        self.stdout.write("")
        self._linea("TG alterado (leve+mod+sev)", tg_alterado)
        ok = "✓" if tg_alterado == suma_alterados else "✗ INCONSISTENCIA"
        self.stdout.write(f"  Sanity check suma manual: {suma_alterados} {ok}\n")

    def _bloque2_desglose_tg_alterado(self, tg_alterados):
        self.stdout.write(self.style.HTTP_INFO("Bloque 2 — Diagnóstico combinado para los pacientes con TG alterado"))
        total = len(tg_alterados)
        if total == 0:
            self.stdout.write("  (ningún paciente con TG alterado)\n")
            return

        buckets = Counter(a.clasif_dislipidemia_tipo for a in tg_alterados)
        esperados = (
            "Hipertrigliceridemia",
            "Dislipidemia mixta",
            "Dislipidemia aterogénica",
        )
        for cat in esperados:
            self._linea(cat, buckets.get(cat, 0), total)

        otros = {k: v for k, v in buckets.items() if k not in esperados}
        if otros:
            self.stdout.write(self.style.WARNING("  Otros (no esperados con TG alterado):"))
            for k, v in otros.items():
                self._linea(f"  · {k}", v, total)

        suma = sum(buckets.values())
        ok = "✓" if suma == total else "✗ INCONSISTENCIA"
        self.stdout.write("")
        self._linea("Suma buckets", suma)
        self._linea("Total TG alterado", total)
        self.stdout.write(f"  {ok}\n")

    def _bloque3_auditoria_hipertri(self, hipertri):
        self.stdout.write(self.style.HTTP_INFO('Bloque 3 — Calidad de data en "Hipertrigliceridemia"'))
        total = len(hipertri)
        self._linea("Pacientes clasificados", total)
        if total == 0:
            self.stdout.write("")
            return

        ct_none = sum(1 for a in hipertri if a.colesterol_total is None)
        ldl_none = sum(1 for a in hipertri if a.ldl is None)
        hdl_none = sum(1 for a in hipertri if a.hdl is None)
        sexo_vacio = sum(1 for a in hipertri if a.paciente.sexo not in ("M", "F"))

        self._linea("CT medido", total - ct_none, total)
        self._linea("CT no medido (None)", ct_none, total)
        self._linea("LDL medido", total - ldl_none, total)
        self._linea("LDL no medido (None)", ldl_none, total)
        self._linea("HDL medido", total - hdl_none, total)
        self._linea("HDL no medido (None)", hdl_none, total)
        self._linea("Sexo registrado", total - sexo_vacio, total)
        self._linea("Sexo en blanco", sexo_vacio, total)

        completos = sum(
            1 for a in hipertri
            if a.colesterol_total is not None
            and a.ldl is not None
            and a.hdl is not None
            and a.paciente.sexo in ("M", "F")
        )
        self.stdout.write("")
        self._linea("Hipertrigliceridemia con perfil completo", completos, total)
        incompletos = total - completos
        self._linea("Hipertrigliceridemia con algún dato faltante", incompletos, total)
        if incompletos:
            self.stdout.write(
                self.style.WARNING(
                    "  ↑ estos casos no se pueden confirmar como hipertrigliceridemia "
                    "pura: podrían ser mixta o aterogénica si tuvieran el dato faltante."
                )
            )
        self.stdout.write("")

    def _bloque4_universo(self, atenciones):
        self.stdout.write(self.style.HTTP_INFO("Bloque 4 — Universo del diagnóstico combinado"))
        total = len(atenciones)
        buckets = Counter(a.clasif_dislipidemia_tipo for a in atenciones)
        sin_diag = buckets.pop(None, 0)
        orden = (
            "Normal",
            "HDL bajo aislado",
            "Hipercolesterolemia",
            "Hipertrigliceridemia",
            "Dislipidemia mixta",
            "Dislipidemia aterogénica",
        )
        con_diag = total - sin_diag
        for cat in orden:
            self._linea(cat, buckets.get(cat, 0), con_diag)
        extras = {k: v for k, v in buckets.items() if k not in orden}
        for k, v in extras.items():
            self._linea(f"(no esperado) {k}", v, con_diag)

        self.stdout.write("")
        self._linea("Con diagnóstico", con_diag)
        self._linea("Sin diagnóstico (sin datos lipídicos)", sin_diag)
        suma = sum(buckets.values()) + sin_diag
        ok = "✓" if suma == total else "✗ INCONSISTENCIA"
        self._linea("Suma total", suma)
        self.stdout.write(f"  {ok}\n")
