"""Tests de criterios + coherencia con umbrales.

Dos objetivos:
1. Asegurar que las funciones `criterios.*` clasifican exactamente igual a
   como lo hacían antes del refactor a `umbrales.py`.
2. Garantizar que los tooltips de los criterios que se construyen a mano
   (presión, dislipidemia global) referencian los MISMOS números que usa
   la lógica clasificadora — así un cambio futuro de umbrales no puede
   desincronizarlos.
"""
from django.test import SimpleTestCase

from . import criterios as c
from . import umbrales as u
from .reportes import TOOLTIPS


class ClasificacionesTests(SimpleTestCase):
    """Valores de borde de cada criterio. Si alguno se rompe, el refactor
    cambió comportamiento clínico observable."""

    def test_colesterol_total(self):
        self.assertEqual(c.clasificar_colesterol_total(199), "Deseable")
        self.assertEqual(c.clasificar_colesterol_total(200), "Dislipidemia leve")
        self.assertEqual(c.clasificar_colesterol_total(240), "Dislipidemia leve")
        self.assertEqual(c.clasificar_colesterol_total(241), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_colesterol_total(300), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_colesterol_total(301), "Dislipidemia severa")
        self.assertIsNone(c.clasificar_colesterol_total(None))

    def test_ldl(self):
        self.assertEqual(c.clasificar_ldl(99), "Deseable")
        self.assertEqual(c.clasificar_ldl(100), "Dislipidemia leve")
        self.assertEqual(c.clasificar_ldl(150), "Dislipidemia leve")
        self.assertEqual(c.clasificar_ldl(151), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_ldl(190), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_ldl(191), "Dislipidemia severa")

    def test_hdl_por_sexo(self):
        self.assertEqual(c.clasificar_hdl(41, "M"), "Deseable")
        self.assertEqual(c.clasificar_hdl(40, "M"), "Dislipidemia leve")
        self.assertEqual(c.clasificar_hdl(35, "M"), "Dislipidemia leve")
        self.assertEqual(c.clasificar_hdl(34, "M"), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_hdl(30, "M"), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_hdl(29, "M"), "Dislipidemia severa")
        self.assertEqual(c.clasificar_hdl(51, "F"), "Deseable")
        self.assertEqual(c.clasificar_hdl(50, "F"), "Dislipidemia leve")
        self.assertEqual(c.clasificar_hdl(40, "F"), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_hdl(39, "F"), "Dislipidemia severa")
        self.assertIsNone(c.clasificar_hdl(40, None))

    def test_trigliceridos(self):
        self.assertEqual(c.clasificar_trigliceridos(149), "Deseable")
        self.assertEqual(c.clasificar_trigliceridos(150), "Dislipidemia leve")
        self.assertEqual(c.clasificar_trigliceridos(300), "Dislipidemia leve")
        self.assertEqual(c.clasificar_trigliceridos(500), "Dislipidemia moderada")
        self.assertEqual(c.clasificar_trigliceridos(501), "Dislipidemia severa")

    def test_colesterol_no_hdl(self):
        self.assertEqual(c.clasificar_colesterol_no_hdl(129), "Deseable")
        self.assertEqual(c.clasificar_colesterol_no_hdl(130), "Limítrofe alto")
        self.assertEqual(c.clasificar_colesterol_no_hdl(159), "Limítrofe alto")
        self.assertEqual(c.clasificar_colesterol_no_hdl(160), "Riesgo aumentado")
        self.assertEqual(c.clasificar_colesterol_no_hdl(189), "Riesgo aumentado")
        self.assertEqual(c.clasificar_colesterol_no_hdl(190), "Riesgo muy alto")

    def test_sospecha_hipercolesterolemia_familiar(self):
        self.assertFalse(c.sospecha_hipercolesterolemia_familiar(189))
        self.assertTrue(c.sospecha_hipercolesterolemia_familiar(190))
        self.assertFalse(c.sospecha_hipercolesterolemia_familiar(None))

    def test_eritrocitosis(self):
        # F sin altura. leve_min=16.1, leve_max=17.5, moderada_max=19.0
        self.assertEqual(c.clasificar_eritrocitosis(16.0, "F", False), "Normal")
        self.assertEqual(c.clasificar_eritrocitosis(16.1, "F", False), "Eritrocitosis leve")
        self.assertEqual(c.clasificar_eritrocitosis(17.5, "F", False), "Eritrocitosis leve")
        self.assertEqual(c.clasificar_eritrocitosis(17.6, "F", False), "Eritrocitosis moderada")
        self.assertEqual(c.clasificar_eritrocitosis(19.0, "F", False), "Eritrocitosis moderada")
        self.assertEqual(c.clasificar_eritrocitosis(19.01, "F", False), "Eritrocitosis severa")
        # F con altura
        self.assertEqual(c.clasificar_eritrocitosis(20.5, "F", True), "Eritrocitosis moderada")
        self.assertEqual(c.clasificar_eritrocitosis(20.6, "F", True), "Eritrocitosis severa")
        # M sin altura
        self.assertEqual(c.clasificar_eritrocitosis(18.1, "M", False), "Eritrocitosis leve")
        self.assertEqual(c.clasificar_eritrocitosis(21.0, "M", False), "Eritrocitosis moderada")
        self.assertEqual(c.clasificar_eritrocitosis(21.01, "M", False), "Eritrocitosis severa")
        # M con altura
        self.assertEqual(c.clasificar_eritrocitosis(22.5, "M", True), "Eritrocitosis moderada")
        self.assertEqual(c.clasificar_eritrocitosis(22.51, "M", True), "Eritrocitosis severa")
        self.assertIsNone(c.clasificar_eritrocitosis(15, None, False))

    def test_obesidad_abdominal(self):
        self.assertEqual(c.clasificar_obesidad_abdominal(93, "M"), "Normal")
        self.assertEqual(c.clasificar_obesidad_abdominal(94, "M"), "Grado 1 (riesgo aumentado)")
        self.assertEqual(c.clasificar_obesidad_abdominal(102, "M"), "Grado 1 (riesgo aumentado)")
        self.assertEqual(c.clasificar_obesidad_abdominal(103, "M"), "Grado 2 (riesgo significativo)")
        self.assertEqual(c.clasificar_obesidad_abdominal(87, "F"), "Normal")
        self.assertEqual(c.clasificar_obesidad_abdominal(88, "F"), "Grado 1 (riesgo aumentado)")
        self.assertEqual(c.clasificar_obesidad_abdominal(90, "F"), "Grado 1 (riesgo aumentado)")
        self.assertEqual(c.clasificar_obesidad_abdominal(91, "F"), "Grado 2 (riesgo significativo)")

    def test_glicemia(self):
        self.assertEqual(c.clasificar_glicemia_ayunas(99), "Normal")
        self.assertEqual(c.clasificar_glicemia_ayunas(100), "Prediabetes")
        self.assertEqual(c.clasificar_glicemia_ayunas(125), "Prediabetes")
        self.assertEqual(c.clasificar_glicemia_ayunas(126), "Diabetes")

    def test_hb_a1c(self):
        self.assertEqual(c.clasificar_hb_a1c(5.5), "Normal")
        self.assertEqual(c.clasificar_hb_a1c(5.6), "Prediabetes")
        self.assertEqual(c.clasificar_hb_a1c(6.4), "Prediabetes")
        self.assertEqual(c.clasificar_hb_a1c(6.5), "Diabetes")

    def test_presion(self):
        # Etiquetas literales del xlsx CRITERIOS DE ESTUDIO.
        self.assertEqual(c.clasificar_presion(110, 70), "Normal Alta")
        self.assertEqual(c.clasificar_presion(125, 70), "Elevada")
        self.assertEqual(c.clasificar_presion(125, 80), "Hipertensión Grado 1")
        self.assertEqual(c.clasificar_presion(135, 85), "Hipertensión Grado 1")
        self.assertEqual(c.clasificar_presion(110, 85), "Hipertensión Grado 1")
        self.assertEqual(c.clasificar_presion(145, 95), "Hipertensión Grado 2")
        self.assertEqual(c.clasificar_presion(110, 95), "Hipertensión Grado 2")
        self.assertEqual(c.clasificar_presion(185, 95), "Crisis Hipertensiva")
        self.assertEqual(c.clasificar_presion(110, 125), "Crisis Hipertensiva")
        self.assertIsNone(c.clasificar_presion(None, None))

    def test_imc(self):
        self.assertEqual(c.clasificar_imc(18.4), "Bajo peso")
        self.assertEqual(c.clasificar_imc(18.5), "Normal")
        self.assertEqual(c.clasificar_imc(24.9), "Normal")
        self.assertEqual(c.clasificar_imc(25), "Sobrepeso")
        self.assertEqual(c.clasificar_imc(35), "Obesidad grado II")
        self.assertEqual(c.clasificar_imc(40), "Obesidad grado III")

    def test_grasa_corporal(self):
        self.assertEqual(c.clasificar_grasa_corporal(5, "M"), "Bajo")
        self.assertEqual(c.clasificar_grasa_corporal(6, "M"), "Saludable")
        self.assertEqual(c.clasificar_grasa_corporal(17, "M"), "Saludable")
        self.assertEqual(c.clasificar_grasa_corporal(18, "M"), "Aceptable")
        self.assertEqual(c.clasificar_grasa_corporal(25, "M"), "Obesidad")
        self.assertEqual(c.clasificar_grasa_corporal(13, "F"), "Bajo")
        self.assertEqual(c.clasificar_grasa_corporal(14, "F"), "Saludable")
        self.assertEqual(c.clasificar_grasa_corporal(32, "F"), "Obesidad")

    def test_grasa_visceral(self):
        self.assertEqual(c.clasificar_grasa_visceral(12), "Saludable")
        self.assertEqual(c.clasificar_grasa_visceral(13), "Alta")

    def test_masa_muscular(self):
        self.assertEqual(c.clasificar_masa_muscular(32, "M"), "Bajo")
        self.assertEqual(c.clasificar_masa_muscular(33, "M"), "Normal")
        self.assertEqual(c.clasificar_masa_muscular(44, "M"), "Alto")
        self.assertEqual(c.clasificar_masa_muscular(45, "M"), "Muy alto")
        self.assertEqual(c.clasificar_masa_muscular(23, "F"), "Bajo")
        self.assertEqual(c.clasificar_masa_muscular(36, "F"), "Muy alto")


class DislipidemiaCombinadaTests(SimpleTestCase):

    def test_perfil_normal(self):
        d = c.analizar_dislipidemia(180, 90, 50, 100, "M")
        self.assertEqual(d["tipo"], "Normal")
        self.assertIsNone(d["severidad"])

    def test_hipercolesterolemia(self):
        d = c.analizar_dislipidemia(210, 90, 50, 100, "M")
        self.assertEqual(d["tipo"], "Hipercolesterolemia")
        self.assertEqual(d["severidad"], "leve")

    def test_hipertrigliceridemia(self):
        d = c.analizar_dislipidemia(180, 90, 50, 200, "M")
        self.assertEqual(d["tipo"], "Hipertrigliceridemia")

    def test_mixta(self):
        d = c.analizar_dislipidemia(210, 90, 50, 200, "M")
        self.assertEqual(d["tipo"], "Dislipidemia mixta")

    def test_aterogenica_prevalece_sobre_mixta(self):
        d = c.analizar_dislipidemia(210, 90, 35, 200, "M")
        self.assertEqual(d["tipo"], "Dislipidemia aterogénica")

    def test_hdl_bajo_aislado(self):
        d = c.analizar_dislipidemia(180, 90, 35, 100, "M")
        self.assertEqual(d["tipo"], "HDL bajo aislado")

    def test_sospecha_hf_bandera(self):
        d = c.analizar_dislipidemia(180, 200, 50, 100, "M")
        self.assertTrue(d["sospecha_familiar"])

    def test_sin_datos(self):
        self.assertIsNone(c.analizar_dislipidemia(None, None, None, None, "M"))


class CoherenciaTooltipsTests(SimpleTestCase):
    """Para los tooltips que se construyen a mano (no auto-generados desde
    bandas), valida que los números que aparecen en el HTML son exactamente
    los que usa la función clasificadora."""

    def test_presion_tooltip_contiene_umbrales(self):
        html = TOOLTIPS["presion"]
        p = u.PRESION
        # Cada umbral debe aparecer literalmente en el HTML.
        for key in ("elevada_s_min", "elevada_s_max", "elevada_d_max",
                    "g1_s_min", "g1_s_max", "g1_d_min", "g1_d_max",
                    "g2_s", "g2_d", "crisis_s", "crisis_d"):
            self.assertIn(str(p[key]), html,
                          f"PRESION['{key}']={p[key]} no aparece en el tooltip")

    def test_dislipidemia_global_tooltip_referencia_umbral_hf(self):
        html = TOOLTIPS["dislipidemia_global"]
        self.assertIn(str(u.LDL_UMBRAL_HF), html)

    def test_eritrocitosis_tooltip_contiene_umbrales(self):
        html = TOOLTIPS["eritrocitosis"]
        for (sexo, altura), t in u.ERITROCITOSIS.items():
            for v in (t["leve_min"], t["leve_max"], t["moderada_max"]):
                # _eritro_linea formatea con :g; los enteros pierden el .0.
                rendered = f"{v:g}"
                self.assertIn(rendered, html,
                              f"Umbral {sexo}/{altura} valor {v} no aparece")
