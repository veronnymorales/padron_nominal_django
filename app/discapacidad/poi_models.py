from django.db import models

class ActividadPOI(models.Model):
    TIPO_PRESUPUESTO = [('PIA', 'PIA'), ('PIM', 'PIM')]
    TIPO_CATEGORIA = [('PpR', 'PpR'), ('APNOP', 'APNOP')]
    TIPO_PRODUCTO = [('Producto', 'Producto'), ('Proyecto', 'Proyecto')]
    UNIDAD_MEDIDA = [('Informes', 'Informes'), ('Reportes', 'Reportes')]

    ano = models.IntegerField(verbose_name="Año")
    tipo_presupuesto = models.CharField(max_length=3, choices=TIPO_PRESUPUESTO)
    fecha_registro = models.DateField()

    # Estructura Orgánica
    pliego = models.CharField(max_length=255)
    unidad_ejecutora = models.CharField(max_length=255)

    # Plan Estratégico
    objetivo_sectorial = models.CharField(max_length=255, verbose_name="Objetivo Sectorial/PESEM",null=True, blank=True)
    objetivo_institucional = models.CharField(max_length=255, verbose_name="Objetivo Institucional/PEI",null=True, blank=True)
    accion_estrategica = models.CharField(max_length=255, verbose_name="Acción Estratégica/PEI",null=True, blank=True)

    # Planificación
    tipo_categoria = models.CharField(max_length=10, choices=TIPO_CATEGORIA,null=True, blank=True)
    categoria_presupuestal = models.CharField(max_length=255,null=True, blank=True)
    tipo_producto_proyecto = models.CharField(max_length=20, choices=TIPO_PRODUCTO,null=True, blank=True)
    producto_presupuestal = models.CharField(max_length=255,null=True, blank=True)
    tipo_actividad_obra = models.CharField(max_length=255,null=True, blank=True)
    actividad_presupuestal = models.CharField(max_length=255,null=True, blank=True)
    funcion = models.CharField(max_length=255,null=True, blank=True)
    division_funcional = models.CharField(max_length=255,null=True, blank=True)
    grupo_funcional = models.CharField(max_length=255,null=True, blank=True)

    # Actividad Operativa
    actividad_operativa = models.CharField(max_length=255,null=True, blank=True)
    actividad_poi = models.CharField(max_length=255,null=True, blank=True)
    unidad_medida = models.CharField(max_length=50, choices=UNIDAD_MEDIDA,null=True, blank=True)
    total_meta_fisica = models.IntegerField(verbose_name="Total Meta Física",null=True, blank=True)
    meta_presupuestal = models.DecimalField(max_digits=12, decimal_places=2,null=True, blank=True)
    
    # Meta Fisiaca
    meta_fisica = models.PositiveIntegerField(default=0)
    meta_programada = models.PositiveIntegerField(default=0,null=True, blank=True)
    meta_anual = models.PositiveIntegerField(default=0,null=True, blank=True)
    enero = models.PositiveIntegerField(default=0,null=True, blank=True)
    febrero = models.PositiveIntegerField(default=0,null=True, blank=True)
    marzo = models.PositiveIntegerField(default=0,null=True, blank=True)
    abril = models.PositiveIntegerField(default=0,null=True, blank=True)
    mayo = models.PositiveIntegerField(default=0,null=True, blank=True)
    junio = models.PositiveIntegerField(default=0,null=True, blank=True)
    julio = models.PositiveIntegerField(default=0,null=True, blank=True)
    agosto = models.PositiveIntegerField(default=0,null=True, blank=True)
    setiembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    octubre = models.PositiveIntegerField(default=0,null=True, blank=True)
    noviembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    diciembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    enero_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    febrero_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    marzo_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    abril_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    mayo_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    junio_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    julio_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    agosto_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    setiembre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    octubre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    noviembre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    diciembre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    total_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    primer_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    segundo_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    primer_semestre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    segundo_semestre_e = models.PositiveIntegerField(default=0,null=True, blank=True)
    porcentaje = models.CharField(max_length=50,null=True, blank=True)
    
    class Meta:
            db_table = 'POI_META_FISICA'
    
    def __str__(self):
        return self.actividad_presupuestal



class ProgramacionMensual(models.Model):
    MESES = [
        ('Enero', 'Enero'), ('Febrero', 'Febrero'), ('Marzo', 'Marzo'),
        ('Abril', 'Abril'), ('Mayo', 'Mayo'), ('Junio', 'Junio'),
        ('Julio', 'Julio'), ('Agosto', 'Agosto'), ('Septiembre', 'Septiembre'),
        ('Octubre', 'Octubre'), ('Noviembre', 'Noviembre'), ('Diciembre', 'Diciembre')
        ]

    actividad_poi = models.ForeignKey(ActividadPOI, on_delete=models.CASCADE)
    mes = models.CharField(max_length=20, choices=MESES,null=True, blank=True)
    meta_fisica = models.IntegerField()
    meta_presupuestal = models.DecimalField(max_digits=12, decimal_places=2,null=True, blank=True)
    
class ProgramacionMetaFisica(models.Model):
    MESES = (
        ('Enero', 'Enero'),
        ('Febrero', 'Febrero'),
        ('Marzo', 'Marzo'),
        ('Abril', 'Abril'),
        ('Mayo', 'Mayo'),
        ('Junio', 'Junio'),
    # ... otros meses
    )
    actividad_poi = models.ForeignKey(ActividadPOI, on_delete=models.CASCADE)
    mes = models.CharField(max_length=10, choices=MESES,null=True, blank=True)
    meta_fisica = models.PositiveIntegerField(default=0)
    meta_programada = models.PositiveIntegerField(default=0,null=True, blank=True)
    meta_anual = models.PositiveIntegerField(default=0,null=True, blank=True)
    enero = models.PositiveIntegerField(default=0,null=True, blank=True)
    febrero = models.PositiveIntegerField(default=0,null=True, blank=True)
    marzo = models.PositiveIntegerField(default=0,null=True, blank=True)
    abril = models.PositiveIntegerField(default=0,null=True, blank=True)
    mayo = models.PositiveIntegerField(default=0,null=True, blank=True)
    junio = models.PositiveIntegerField(default=0,null=True, blank=True)
    julio = models.PositiveIntegerField(default=0,null=True, blank=True)
    agosto = models.PositiveIntegerField(default=0,null=True, blank=True)
    setiembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    octubre = models.PositiveIntegerField(default=0,null=True, blank=True)
    noviembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    diciembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    total = models.PositiveIntegerField(default=0,null=True, blank=True)
    primer_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    segundo_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    porcentaje = models.CharField(max_length=50,null=True, blank=True)
    
    class Meta:
        db_table = 'META_FISICA'
    
    def __str__(self):
        return self.actividad_poi

class ProgramacionMensualMetaFinanciera(models.Model):
    actividad_poi = models.ForeignKey(ActividadPOI, on_delete=models.CASCADE,default=1)
    meta_programada = models.PositiveIntegerField(default=0,null=True, blank=True)
    meta_anual = models.PositiveIntegerField(default=0,null=True, blank=True)
    fuente_f = models.CharField(max_length=100,null=True, blank=True)
    generico_gasto = models.CharField(max_length=100,null=True, blank=True)
    enero = models.PositiveIntegerField(default=0,null=True, blank=True)
    febrero = models.PositiveIntegerField(default=0,null=True, blank=True)
    marzo = models.PositiveIntegerField(default=0,null=True, blank=True)
    abril = models.PositiveIntegerField(default=0,null=True, blank=True)
    mayo = models.PositiveIntegerField(default=0,null=True, blank=True)
    junio = models.PositiveIntegerField(default=0,null=True, blank=True)
    julio = models.PositiveIntegerField(default=0,null=True, blank=True)
    agosto = models.PositiveIntegerField(default=0,null=True, blank=True)
    setiembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    octubre = models.PositiveIntegerField(default=0,null=True, blank=True)
    noviembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    diciembre = models.PositiveIntegerField(default=0,null=True, blank=True)
    total = models.PositiveIntegerField(default=0,null=True, blank=True)
    primer_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    segundo_semestre = models.PositiveIntegerField(default=0,null=True, blank=True)
    porcentaje = models.CharField(max_length=50,null=True, blank=True)
    
    def __str__(self):
        return f"{self.actividad_poi.actividad_operativa} - {self.mes} - Meta Física: {self.meta_programada}"
    
    
class ProgramacionMensualMetaFisica(models.Model):
    MESES = [
        ('Enero', 'Enero'), ('Febrero', 'Febrero'), ('Marzo', 'Marzo'),
        ('Abril', 'Abril'), ('Mayo', 'Mayo'), ('Junio', 'Junio'),
        ('Julio', 'Julio'), ('Agosto', 'Agosto'), ('Septiembre', 'Septiembre'),
        ('Octubre', 'Octubre'), ('Noviembre', 'Noviembre'), ('Diciembre', 'Diciembre')
    ]

    actividad_poi = models.ForeignKey(ActividadPOI, on_delete=models.CASCADE)
    mes = models.CharField(max_length=20, choices=MESES,null=True, blank=True)
    meta_fisica = models.IntegerField()
    meta_presupuestal = models.DecimalField(max_digits=12, decimal_places=2,null=True, blank=True)