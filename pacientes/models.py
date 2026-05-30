from datetime import date

from django.db import models


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
