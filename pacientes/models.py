from datetime import date

from django.db import models

from . import criterios
from .peru_altitudes import altitud_estimada, es_altura


class Paciente(models.Model):
    SEXO_CHOICES = [
        ("M", "Masculino"),
        ("F", "Femenino"),
    ]
    INSTRUCCION_CHOICES = [
        ("SIN_INSTRUCCION", "Sin instrucción"),
        ("PRIMARIA", "Primaria"),
        ("SECUNDARIA", "Secundaria"),
        ("TECNICA", "Técnica"),
        ("SUPERIOR", "Superior"),
        ("POSGRADO", "Posgrado"),
    ]

    id_rh = models.CharField("Id RH", max_length=50, unique=True)
    nombre_completo = models.CharField("Nombre Completo", max_length=200)
    fec_ingreso = models.DateField("Fec. Ingreso", null=True, blank=True)
    nro_documento = models.CharField("Nº de documento", max_length=20)
    convenio = models.CharField("Convenio", max_length=150, blank=True)
    nombre_posicion = models.CharField("Nombre posición", max_length=150, blank=True)
    nombre_area = models.CharField("Nombre de Area", max_length=150, blank=True)
    nombre_unidad_org = models.CharField("Nombre unidad org.", max_length=150, blank=True)
    fec_nacimiento = models.DateField("Fec. Nacimiento", null=True, blank=True)
    sexo = models.CharField("Sexo", max_length=1, choices=SEXO_CHOICES, blank=True)
    instruccion = models.CharField(
        "Instrucción", max_length=20, choices=INSTRUCCION_CHOICES, blank=True
    )
    departamento = models.CharField("Departamento", max_length=100, blank=True)
    provincia = models.CharField("Provincia", max_length=100, blank=True)
    distrito = models.CharField("Distrito", max_length=100, blank=True)
    direccion = models.CharField("Dirección", max_length=255, blank=True)
    altitud_msnm = models.IntegerField(
        "Altitud (msnm)",
        null=True,
        blank=True,
        help_text="Si se deja vacío se estima a partir de departamento/distrito.",
    )

    class Meta:
        verbose_name = "Paciente"
        verbose_name_plural = "Pacientes"
        ordering = ["nombre_completo"]

    def __str__(self):
        return f"{self.nro_documento} - {self.nombre_completo}"

    @property
    def edad(self):
        if not self.fec_nacimiento:
            return None
        hoy = date.today()
        return (
            hoy.year
            - self.fec_nacimiento.year
            - ((hoy.month, hoy.day) < (self.fec_nacimiento.month, self.fec_nacimiento.day))
        )

    @property
    def altitud_efectiva(self):
        if self.altitud_msnm is not None:
            return self.altitud_msnm
        return altitud_estimada(self.distrito, self.provincia, self.departamento)

    @property
    def es_altura(self):
        return es_altura(self.altitud_efectiva)


class Atencion(models.Model):
    paciente = models.ForeignKey(
        Paciente,
        on_delete=models.PROTECT,
        related_name="atenciones",
        verbose_name="Paciente",
    )
    fecha = models.DateField("Fecha de atención", default=date.today)

    peso = models.DecimalField("Peso (kg)", max_digits=5, decimal_places=2, null=True, blank=True)
    talla = models.DecimalField("Talla (cm)", max_digits=5, decimal_places=2, null=True, blank=True)
    imc = models.DecimalField("IMC", max_digits=5, decimal_places=2, null=True, blank=True)
    sistolica = models.DecimalField("Sistólica", max_digits=5, decimal_places=1, null=True, blank=True)
    diastolica = models.DecimalField("Diastólica", max_digits=5, decimal_places=1, null=True, blank=True)
    abdominal = models.DecimalField("Abdominal (cm)", max_digits=5, decimal_places=2, null=True, blank=True)
    icc = models.DecimalField("ICC", max_digits=4, decimal_places=2, null=True, blank=True)

    colesterol_total = models.DecimalField("Colesterol Total", max_digits=6, decimal_places=2, null=True, blank=True)
    hdl = models.DecimalField("HDL Colesterol", max_digits=6, decimal_places=2, null=True, blank=True)
    ldl = models.DecimalField("LDL Colesterol", max_digits=6, decimal_places=2, null=True, blank=True)
    vldl = models.DecimalField("VLDL Colesterol", max_digits=6, decimal_places=2, null=True, blank=True)
    trigliceridos = models.DecimalField("Triglicéridos", max_digits=6, decimal_places=2, null=True, blank=True)

    hemoglobina = models.DecimalField("Hemoglobina", max_digits=5, decimal_places=2, null=True, blank=True)
    hematocrito = models.DecimalField("Hematocrito", max_digits=5, decimal_places=2, null=True, blank=True)
    glucosa = models.DecimalField("Glucosa", max_digits=6, decimal_places=2, null=True, blank=True)
    hb_a1c = models.DecimalField("HB A1c", max_digits=5, decimal_places=2, null=True, blank=True)

    class Meta:
        verbose_name = "Atención"
        verbose_name_plural = "Atenciones"
        ordering = ["-fecha", "paciente__nombre_completo"]

    def __str__(self):
        return f"{self.paciente.nro_documento} - {self.fecha}"

    # ----- Clasificaciones clínicas (calculadas, no se guardan en BD) -----

    @property
    def clasif_colesterol_total(self):
        return criterios.clasificar_colesterol_total(self.colesterol_total)

    @property
    def clasif_hdl(self):
        return criterios.clasificar_hdl(self.hdl, self.paciente.sexo)

    @property
    def clasif_ldl(self):
        return criterios.clasificar_ldl(self.ldl)

    @property
    def clasif_trigliceridos(self):
        return criterios.clasificar_trigliceridos(self.trigliceridos)

    @property
    def colesterol_no_hdl(self):
        if self.colesterol_total is None or self.hdl is None:
            return None
        return self.colesterol_total - self.hdl

    @property
    def clasif_colesterol_no_hdl(self):
        return criterios.clasificar_colesterol_no_hdl(self.colesterol_no_hdl)

    @property
    def _dislipidemia_analisis(self):
        return criterios.analizar_dislipidemia(
            self.colesterol_total,
            self.ldl,
            self.hdl,
            self.trigliceridos,
            self.paciente.sexo,
        )

    @property
    def clasif_dislipidemia(self):
        a = self._dislipidemia_analisis
        return a["label"] if a else None

    @property
    def clasif_dislipidemia_tipo(self):
        a = self._dislipidemia_analisis
        return a["tipo"] if a else None

    @property
    def sospecha_hipercolesterolemia_familiar(self):
        return criterios.sospecha_hipercolesterolemia_familiar(self.ldl)

    @property
    def clasif_eritrocitosis(self):
        return criterios.clasificar_eritrocitosis(
            self.hemoglobina, self.paciente.sexo, self.paciente.es_altura
        )

    @property
    def clasif_obesidad_abdominal(self):
        return criterios.clasificar_obesidad_abdominal(
            self.abdominal, self.paciente.sexo
        )

    @property
    def clasif_glicemia_ayunas(self):
        return criterios.clasificar_glicemia_ayunas(self.glucosa)

    @property
    def clasif_hb_a1c(self):
        return criterios.clasificar_hb_a1c(self.hb_a1c)

    @property
    def clasif_presion(self):
        return criterios.clasificar_presion(self.sistolica, self.diastolica)


class AtencionNutricion(models.Model):
    paciente = models.ForeignKey(
        Paciente,
        on_delete=models.PROTECT,
        related_name="atenciones_nutricion",
        verbose_name="Paciente",
    )
    fecha = models.DateField("Fecha de atención", default=date.today)

    peso = models.DecimalField("Peso (kg)", max_digits=5, decimal_places=2, null=True, blank=True)
    talla = models.DecimalField("Talla (cm)", max_digits=5, decimal_places=2, null=True, blank=True)
    imc = models.DecimalField("IMC", max_digits=5, decimal_places=2, null=True, blank=True)

    grasa_corporal = models.DecimalField(
        "% Grasa corporal", max_digits=5, decimal_places=2, null=True, blank=True
    )
    grasa_visceral = models.DecimalField(
        "% Grasa visceral", max_digits=5, decimal_places=2, null=True, blank=True
    )
    masa_muscular = models.DecimalField(
        "% Masa muscular", max_digits=5, decimal_places=2, null=True, blank=True
    )

    class Meta:
        verbose_name = "Atención de nutrición"
        verbose_name_plural = "Atenciones de nutrición"
        ordering = ["-fecha", "paciente__nombre_completo"]

    def __str__(self):
        return f"{self.paciente.nro_documento} - {self.fecha}"

    @property
    def clasif_imc(self):
        return criterios.clasificar_imc(self.imc)

    @property
    def clasif_grasa_corporal(self):
        return criterios.clasificar_grasa_corporal(self.grasa_corporal, self.paciente.sexo)

    @property
    def clasif_grasa_visceral(self):
        return criterios.clasificar_grasa_visceral(self.grasa_visceral)

    @property
    def clasif_masa_muscular(self):
        return criterios.clasificar_masa_muscular(self.masa_muscular, self.paciente.sexo)
