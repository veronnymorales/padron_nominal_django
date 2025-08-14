from django.db import models

class TramaBaseDiscapacidadRpt02FisicaNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_02_FISICA_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita
    

class TramaBaseDiscapacidadRpt01CapacitacionMedicinaNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    casos = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_01_CAPACITACION_MEDICINA_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita
    
class TramaBaseDiscapacidadRpt03SensorialNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_03_SENSORIAL_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita
    
class TramaBaseDiscapacidadRpt04CertificadoNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_04_CERTIFICADO_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita

class TramaBaseDiscapacidadRpt05RBCNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    casos = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_05_RBC_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita
    
class TramaBaseDiscapacidadRpt06CapacitacionAgenteNominalAT(models.Model):
    id_cita = models.CharField(max_length=50, primary_key=True)
    renaes = models.CharField(max_length=9)
    id_persona = models.CharField(max_length=50, null=True)
    periodo = models.CharField(max_length=8, null=True)
    fichafam = models.CharField(max_length=50, null=True)
    ubigeo = models.IntegerField(null=True)
    edad = models.IntegerField(null=True)
    sexo = models.CharField(max_length=1, null=True)
    et = models.CharField(max_length=2, null=True)
    fi = models.CharField(max_length=2, null=True)
    id_profesional = models.CharField(max_length=50, null=True)
    Categoria = models.IntegerField(null=True)
    gedad = models.IntegerField(null=True)
    pais = models.CharField(max_length=3, null=True)
    id_ups = models.CharField(max_length=6, null=True)
    ate = models.CharField(max_length=6, null=True)
    atc = models.CharField(max_length=6, null=True)

    class Meta:
        db_table = 'TRAMA_BASE_DISCAPACIDAD_RPT_06_CAPACITACION_AGENTE_NOMINAL_ATE_ATC'

    def __str__(self):
        return self.id_cita