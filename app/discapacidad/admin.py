from import_export import resources
from import_export.admin import ImportExportModelAdmin
from django.contrib import admin
from django.contrib.auth.models import Permission
# Register your models here.
from django.contrib.auth.models import Permission

from .poi_models import ActividadPOI
from .padron_models import Directorio_municipio, Directorio_salud
from base.models import Actualizacion

admin.site.register(Permission)

## admin.site.register(ActividadPOI)

#------------Red---------------------------------
class POIResources(resources.ModelResource):
    class Meta:
        model = ActividadPOI

@admin.register(ActividadPOI)
class POIAdmin(ImportExportModelAdmin,admin.ModelAdmin):
    resource_class = POIResources
    list_display = (
        'id',
        'unidad_ejecutora',
        'categoria_presupuestal',
        'producto_presupuestal',
        'tipo_actividad_obra',
        'actividad_presupuestal',
        'actividad_operativa',
        'total_meta_fisica',
    )
    search_fields = ('unidad_ejecutora','categoria_presupuestal','producto_presupuestal','tipo_actividad_obra','actividad_presupuestal','actividad_operativa','total_meta_fisica',)

#------------Red---------------------------------
class Directorio_municipioResources(resources.ModelResource):
    class Meta:
        model = Directorio_municipio

@admin.register(Directorio_municipio)
class Directorio_municipioAdmin(ImportExportModelAdmin,admin.ModelAdmin):
    resource_class = Directorio_municipioResources
    list_display = (
                'id',
                'tipo_documento', 
                'documento_identidad', 
                'apellido_paterno', 
                'apellido_materno', 
                'nombres', 
                'telefono', 
                'correo_electronico', 
                'provincia', 
                'distrito', 
                'ubigueo', 
                'nombre_municipio', 
                'cargo', 
                'perfil', 
                'condicion', 
                'cuenta_usuario', 
                'estado_usuario', 
                'situacion_usuario', 
                'condicion_laboral', 
                'estado_auditoria',
    )
    search_fields = (                'id',
                'tipo_documento', 
                'documento_identidad', 
                'apellido_paterno', 
                'apellido_materno', 
                'nombres', 
                'telefono', 
                'correo_electronico', 
                'provincia', 
                'distrito', 
                'ubigueo', 
                'nombre_municipio', 
                'cargo', 
                'perfil', 
                'condicion', 
                'cuenta_usuario', 
                'estado_usuario', 
                'situacion_usuario', 
                'condicion_laboral', 
                'estado_auditoria',)
    
    
#----------- DISCAPACIDAD ---------------------------------
class ActualizacionResources(resources.ModelResource):
    class Meta:
        model = Actualizacion

@admin.register(Actualizacion)
class ActualizacionAdmin(ImportExportModelAdmin,admin.ModelAdmin):
    resource_class = ActualizacionResources
    list_display = (
                'id',
                'fecha', 
                'hora', 
                'Descripcion', 

    )
    search_fields = (             
                'id',
                'fecha', 
                'hora', 
                'Descripcion', 
                )    
    
#------------Red---------------------------------
class Directorio_saludResources(resources.ModelResource):
    class Meta:
        model = Directorio_salud

@admin.register(Directorio_salud)
class Directorio_saludAdmin(ImportExportModelAdmin,admin.ModelAdmin):
    resource_class = Directorio_saludResources
    list_display = (
                'id',
                'tipo_documento', 
                'documento_identidad', 
                'apellido_paterno', 
                'apellido_materno', 
                'nombres', 
                'telefono', 
                'correo_electronico', 
                'provincia', 
                'distrito', 
                'ubigueo', 
                'red', 
                'microred', 
                'establecimiento', 
                'cargo', 
                'perfil', 
                'condicion', 
                'cuenta_usuario', 
                'estado_usuario', 
                'situacion_usuario', 
                'condicion_laboral', 
                'estado_auditoria',
    )
    search_fields = (                
                'id',
                'tipo_documento', 
                'documento_identidad', 
                'apellido_paterno', 
                'apellido_materno', 
                'nombres', 
                'telefono', 
                'correo_electronico', 
                'provincia', 
                'distrito', 
                'ubigueo', 
                'red', 
                'microred', 
                'establecimiento', 
                'cargo', 
                'perfil', 
                'condicion', 
                'cuenta_usuario', 
                'estado_usuario', 
                'situacion_usuario', 
                'condicion_laboral', 
                'estado_auditoria',)