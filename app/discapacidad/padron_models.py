from django.db import models
from django.contrib.auth.models import User

class Directorio_municipio(models.Model):
    TIPO_DOCUMENTO = [
                ('DNI', 'DNI'),
                ('Carnet de Extranjeria', 'Carnet de Extranjeria'),
                ('Pasaporte', 'Pasaporte'),
                ('Cedula de Identidad', 'Cedula de Identidad'),
                ('Carnet de solicitante de refugio', 'Carnet de solicitante de refugio'),
                ('Sin Documento', 'Sin Documento'),
            ]
    
    CARGO = [
                ('1', 'Responsable PN'),
                ('2', 'Responsable SELLO'),
                ('3', 'Responsable COMPROMISO 1'),
                ('4', 'Gerente desarrollo social'),
                ('5', 'Consulta Municipio'),
            ]
    
    PERFIL = [
                ('1', 'Consultor'),
                ('2', 'Registrador'),
            ]

    CONDICION = [
                    ('1', 'Alta'),
                    ('0', 'Baja'),
                ]

    CUENTA_USUARIO = [
                    ('0', 'No'),
                    ('1', 'Si'),
                    ('3', 'Espera respuesta MINSA/RENIEC'),
                ]

    ESTADO_USUARIO = [
                    ('Nuevo', 'Nuevo'),
                    ('Continuador', 'Continuador'),
                    ('Espera respuesta MINSA', 'Espera respuesta MINSA'),
                ]

    SITUACION_USUARIO = [
                    ('Tengo Usuario','Tengo Usuario'),
                    ('No llega el correo con la contraseña temporal','No llega el correo con la contraseña temporal'),
                    ('Usuario aparece bloqueado','Usuario aparece bloqueado'),
                    ('Solicitud de ALTA TEMPORAL','Solicitud de ALTA TEMPORAL'),
                    ('Solicitud de USUARIO NUEVO','Solicitud de USUARIO NUEVO'),
                    ('Distrito asignado no corresponde','Distrito asignado no corresponde'),
                    ('Usuario no pertenece a un grupo válido','Usuario no pertenece a un grupo válido'),
                    ('Otros','Otros'),
                ]
    
    TIPO_EMPLEADO = [
                ('Nombrado', 'Nombrado'),
                ('Contrato Plazo Fijo', 'Contrato Plazo Fijo'),
                ('Contrato Plazo Indet.', 'Contrato Plazo Indet.'),
                ('Contrato-CAS', 'Contrato-CAS'),
                ('Destacado Externo', 'Destacado Externo'),
                ('Contrato P.S./CAS Asistencial', 'Contrato P.S./CAS Asistencial'),   
                ('Tercero', 'Tercero'),          
            ]
    
    ESTADO_CHOICES = [
                ('0', 'Pendiente'),
                ('1', 'Aprobado'),
                ('2', 'Proceso'),
                ('3', 'Observado'),
            ]
    
    tipo_documento = models.CharField(choices=TIPO_DOCUMENTO, max_length=100, null=True, blank=True)
    documento_identidad = models.CharField(max_length=100,null=True, blank=True)
    apellido_paterno = models.CharField(max_length=100,null=True, blank=True)
    apellido_materno = models.CharField(max_length=200,null=True, blank=True)
    nombres = models.CharField(max_length=200,null=True, blank=True)
    nombre_completo = models.CharField(max_length=200,null=True, blank=True)
    telefono = models.CharField(max_length=200,null=True, blank=True)
    correo_electronico = models.CharField(max_length=200,null=True, blank=True)
    provincia= models.CharField(max_length=100,null=True, blank=True)
    distrito = models.CharField(max_length=100,null=True, blank=True)
    ubigueo = models.CharField(max_length=100,null=True, blank=True)
    nombre_municipio = models.CharField(max_length=100,null=True, blank=True)
    
    cargo = models.CharField(choices=CARGO, max_length=100, null=True, blank=True)
    perfil = models.CharField(choices=PERFIL,max_length=100,null=True, blank=True)
    condicion = models.CharField(choices=CONDICION,max_length=100,null=True, blank=True)
    cuenta_usuario = models.CharField(choices=CUENTA_USUARIO,max_length=100,null=True, blank=True)
    estado_usuario = models.CharField(choices=ESTADO_USUARIO,max_length=100,null=True, blank=True)
    situacion_usuario = models.CharField(choices=SITUACION_USUARIO,max_length=100,null=True, blank=True)
    condicion_laboral = models.CharField(choices=TIPO_EMPLEADO,max_length=200,null=True, blank=True)
    estado_auditoria = models.CharField(max_length=50,choices=ESTADO_CHOICES,null=True, blank=True)
    
    user = models.ForeignKey(User, on_delete=models.CASCADE,null=True, blank=True)  
    
    req_oficio = models.FileField(upload_to="discapacidad/static/oficio",null=True, blank=True)
    dateTimeOfUpload_req_oficio = models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_resolucion = models.FileField(upload_to="discapacidad/static/resolucion",null=True, blank=True)
    dateTimeOfUpload_req_resolucion= models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_formato_alta = models.FileField(upload_to="discapacidad/static/formato_alta",null=True, blank=True)
    dateTimeOfUpload_req_formato_alta = models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_formato_excel = models.FileField(upload_to="discapacidad/static/formato_excel",null=True, blank=True)
    dateTimeOfUpload_req_formato_excel = models.DateTimeField(auto_now = True,null=True, blank=True)
        
    def __str__(self):
        return self.documento_identidad 

class Directorio_salud(models.Model):
    TIPO_DOCUMENTO = [
                ('DNI', 'DNI'),
                ('Carnet de Extranjeria', 'Carnet de Extranjeria'),
                ('Pasaporte', 'Pasaporte'),
                ('Cedula de Identidad', 'Cedula de Identidad'),
                ('Carnet de solicitante de refugio', 'Carnet de solicitante de refugio'),
                ('Sin Documento', 'Sin Documento'),
            ]
    
    CARGO = [
                ('1', 'Responsable PN'),
                ('2', 'Responsable PROMSA'),
                ('3', 'Responsable IMNUMIZACIONES'),
                ('4', 'Otros'),
            ]
    
    PERFIL = [
                ('1', 'Consultor'),
                ('2', 'Registrador'),
            ]

    CONDICION = [
                    ('1', 'Alta'),
                    ('0', 'Baja'),
                ]

    CUENTA_USUARIO = [
                    ('0', 'No'),
                    ('1', 'Si'),
                    ('3', 'Espera respuesta MINSA/RENIEC'),
                ]

    ESTADO_USUARIO = [
                    ('Nuevo', 'Nuevo'),
                    ('Continuador', 'Continuador'),
                    ('Espera respuesta MINSA', 'Espera respuesta MINSA'),
                ]

    SITUACION_USUARIO = [
                    ('Tengo Usuario','Tengo Usuario'),
                    ('No llega el correo con la contraseña temporal','No llega el correo con la contraseña temporal'),
                    ('Usuario aparece bloqueado','Usuario aparece bloqueado'),
                    ('Solicitud de ALTA TEMPORAL','Solicitud de ALTA TEMPORAL'),
                    ('Solicitud de USUARIO NUEVO','Solicitud de USUARIO NUEVO'),
                    ('Distrito asignado no corresponde','Distrito asignado no corresponde'),
                    ('Usuario no pertenece a un grupo válido','Usuario no pertenece a un grupo válido'),
                    ('Otros','Otros'),
                ]
    
    TIPO_EMPLEADO = [
                ('Nombrado', 'Nombrado'),
                ('Contrato Plazo Fijo', 'Contrato Plazo Fijo'),
                ('Contrato Plazo Indet.', 'Contrato Plazo Indet.'),
                ('Contrato-CAS', 'Contrato-CAS'),
                ('Destacado Externo', 'Destacado Externo'),
                ('Contrato P.S./CAS Asistencial', 'Contrato P.S./CAS Asistencial'),   
                ('Tercero', 'Tercero'),          
            ]
    
    ESTADO_CHOICES = [
                ('0', 'Pendiente'),
                ('1', 'Aprobado'),
                ('2', 'Proceso'),
                ('3', 'Observado'),
            ]
    
    tipo_documento = models.CharField(choices=TIPO_DOCUMENTO, max_length=100, null=True, blank=True)
    documento_identidad = models.CharField(max_length=100,null=True, blank=True)
    apellido_paterno = models.CharField(max_length=100,null=True, blank=True)
    apellido_materno = models.CharField(max_length=200,null=True, blank=True)
    nombres = models.CharField(max_length=200,null=True, blank=True)
    nombre_completo = models.CharField(max_length=200,null=True, blank=True)
    telefono = models.CharField(max_length=200,null=True, blank=True)
    correo_electronico = models.CharField(max_length=200,null=True, blank=True)
    provincia= models.CharField(max_length=100,null=True, blank=True)
    distrito = models.CharField(max_length=100,null=True, blank=True)
    ubigueo = models.CharField(max_length=100,null=True, blank=True)
    red = models.CharField(max_length=100,null=True, blank=True)
    microred = models.CharField(max_length=100,null=True, blank=True)
    establecimiento = models.CharField(max_length=100,null=True, blank=True)
    
    cargo = models.CharField(choices=CARGO, max_length=100, null=True, blank=True)
    perfil = models.CharField(choices=PERFIL,max_length=100,null=True, blank=True)
    condicion = models.CharField(choices=CONDICION,max_length=100,null=True, blank=True)
    cuenta_usuario = models.CharField(choices=CUENTA_USUARIO,max_length=100,null=True, blank=True)
    estado_usuario = models.CharField(choices=ESTADO_USUARIO,max_length=100,null=True, blank=True)
    situacion_usuario = models.CharField(choices=SITUACION_USUARIO,max_length=100,null=True, blank=True)
    condicion_laboral = models.CharField(choices=TIPO_EMPLEADO,max_length=200,null=True, blank=True)
    estado_auditoria = models.CharField(max_length=50,choices=ESTADO_CHOICES,null=True, blank=True)
    
    user = models.ForeignKey(User, on_delete=models.CASCADE,null=True, blank=True)  
    
    req_oficio = models.FileField(upload_to="discapacidad/static/oficio",null=True, blank=True)
    dateTimeOfUpload_req_oficio = models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_resolucion = models.FileField(upload_to="discapacidad/static/resolucion",null=True, blank=True)
    dateTimeOfUpload_req_resolucion= models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_formato_alta = models.FileField(upload_to="discapacidad/static/formato_alta",null=True, blank=True)
    dateTimeOfUpload_req_formato_alta = models.DateTimeField(auto_now = True,null=True, blank=True)
    
    req_formato_excel = models.FileField(upload_to="discapacidad/static/formato_excel",null=True, blank=True)
    dateTimeOfUpload_req_formato_excel = models.DateTimeField(auto_now = True,null=True, blank=True)
        
    def __str__(self):
        return self.documento_identidad 
