from django import forms
from django.forms import modelformset_factory
from .poi_models import ActividadPOI, ProgramacionMensual, ProgramacionMensualMetaFinanciera,ProgramacionMetaFisica
from .padron_models import Directorio_municipio, Directorio_salud
class ActividadPOIForm(forms.ModelForm):
    class Meta:
       model =  ActividadPOI
       exclude = ['ano','tipo_presupuesto','fecha_registro','pliego','unidad_ejecutora','objetivo_sectorial','objetivo_institucional','accion_estrategica','tipo_categoria']       
       fields = [
                'categoria_presupuestal',
                'tipo_producto_proyecto',
                'producto_presupuestal',
                'tipo_actividad_obra',
                'actividad_presupuestal',
                'funcion',
                'division_funcional',
                'grupo_funcional',
                'actividad_operativa',
                'unidad_medida',
                'total_meta_fisica',
                'meta_presupuestal',
                'meta_fisica',
                'meta_programada',
                'meta_anual',
                'enero',
                'febrero',
                'marzo',
                'abril',
                'mayo',
                'junio',
                'julio',
                'agosto',
                'setiembre',
                'octubre',
                'noviembre',
                'diciembre',
                'enero_e',
                'febrero_e',
                'marzo_e',
                'abril_e',
                'mayo_e',
                'junio_e',
                'julio_e',
                'agosto_e',
                'setiembre_e',
                'octubre_e',
                'noviembre_e',
                'diciembre_e',
                'total_e',
                'primer_semestre',
                'segundo_semestre',
                'primer_semestre_e',
                'segundo_semestre_e',
                'porcentaje',
                ]     
       widgets = {
                'categoria_presupuestal': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'tipo_producto_proyecto': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'producto_presupuestal': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'tipo_actividad_obra': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'actividad_presupuestal': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'funcion': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'division_funcional': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'grupo_funcional': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'actividad_operativa': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'unidad_medida': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'total_meta_fisica': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'meta_presupuestal': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'meta_fisica': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'meta_programada': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'meta_anual': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'enero': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'febrero': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'marzo': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'abril': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'mayo': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'junio': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'julio': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'agosto': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'setiembre': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'octubre': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'noviembre': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'diciembre': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'enero_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'febrero_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'marzo_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'abril_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'mayo_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'junio_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'julio_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'agosto_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'setiembre_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'octubre_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'noviembre_e' : forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'diciembre_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'total_e' : forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'primer_semestre' : forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'segundo_semestre' : forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'primer_semestre_e' : forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'segundo_semestre_e': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'porcentaje': forms.TextInput(attrs={'class':'form-control','style': 'border-color: silver; color: silver;','readonly':'readonly'}),          
       }
       labels = {
            'categoria_presupuestal':'Categoria Presupuestal',
            
        }


ProgramacionMensualFormSet = modelformset_factory(
    ProgramacionMensual,
    fields=['mes', 'meta_fisica', 'meta_presupuestal'],
    extra=0, max_num=12, validate_max=True,
    widgets={
        'mes': forms.Select(attrs={'class': 'form-control'}),
        'meta_fisica': forms.NumberInput(attrs={'class': 'form-control'}),
        'meta_presupuestal': forms.NumberInput(attrs={'class': 'form-control', 'step': '0.01'}),
    }
)

class RegistroTareaForm(forms.ModelForm):
    class Meta:
        model = ActividadPOI
        fields = '__all__'

class ProgramacionMensualMetaFisicaForm(forms.ModelForm):
    class Meta:
        model = ProgramacionMetaFisica
        fields = '__all__'

class ProgramacionMensualMetaFinancieraForm(forms.ModelForm):
    class Meta:
        model = ProgramacionMensualMetaFinanciera
        fields = '__all__'
        
class Directorio_MunicipioForm(forms.ModelForm):
    class Meta:
        model =  Directorio_municipio
        exclude = ['nombre_completo','dateTimeOfUpload_req_oficio','dateTimeOfUpload_req_resolucion','dateTimeOfUpload_req_formato_alta','dateTimeOfUpload_req_formato_excel']       
        fields = [
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
                'condicion_laboral',
                'situacion_usuario',
                'user',
                'req_oficio',
                'req_resolucion',
                'req_formato_alta',
                'req_formato_excel',
        ]     
        widgets = {
                'tipo_documento' : forms.Select(attrs={'class':'form-control','required': True, 'tabindex': '1'}),
                'documento_identidad' : forms.NumberInput(attrs={'class':'form-control','required': True,'tabindex': '2','placeholder': 'Ingrese solo números'}),
                'apellido_paterno' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '3'}),
                'apellido_materno' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '4'}),
                'nombres' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '5'}),
                'telefono' : forms.NumberInput(attrs={'class':'form-control','required': True,'tabindex': '6','placeholder': 'Ingrese solo números'}),
                'correo_electronico' : forms.EmailInput(attrs={'class':'form-control','required': True,'tabindex': '7'}),
                'provincia' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '8','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'distrito' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '9','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'ubigueo' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '10','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'nombre_municipio' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '11','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'condicion_laboral' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '12'}),
                'cargo' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '13'}),
                'perfil' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '14'}),
                'condicion' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '15'}),
                'user' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'cuenta_usuario' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'estado_usuario' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'nombre_completo' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'estado_auditoria' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                #####################
                'situacion_usuario': forms.Select(attrs={'class':'form-control','required': True,'tabindex': '16'}),
                'req_oficio': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '17'}),
                'req_resolucion': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '18'}),
                'req_formato_alta': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '19'}),
                'req_formato_excel': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '20'}),
                'dateTimeOfUpload_req_oficio': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'dateTimeOfUpload_req_resolucion': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}), 
                'dateTimeOfUpload_req_formato_alta':forms.TextInput(attrs={'class':'form-control','style': 'display: none'}), 
                'dateTimeOfUpload_req_formato_excel': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
        }
        labels = {
            'tipo_documento':'Tipo Documento',
            'documento_identidad':'Numero de Documento',
            'correo_electronico':'Correo electronico',
            'nombre_municipio':'Nombre de Municipio',
            'req_oficio': 'Oficio',
            'req_resolucion':'Resolucion',
            'req_formato_alta':'Formato Alta/Baja',
            'req_formato_excel':'Formato Excel',            
            'ubigueo': '',
            'estado_usuario': '',
            'nombre_completo': '',
            'situacion_usuario': 'Situacion',
        }
        
class Directorio_SaludForm(forms.ModelForm):
    class Meta:
        model =  Directorio_salud
        exclude = ['nombre_completo','dateTimeOfUpload_req_oficio','dateTimeOfUpload_req_resolucion','dateTimeOfUpload_req_formato_alta','dateTimeOfUpload_req_formato_excel']       
        fields = [
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
                'condicion_laboral',
                'situacion_usuario',
                'user',
                'req_oficio',
                'req_resolucion',
                'req_formato_alta',
                'req_formato_excel',
        ]     
        widgets = {
                'tipo_documento' : forms.Select(attrs={'class':'form-control','required': True, 'tabindex': '1'}),
                'documento_identidad' : forms.NumberInput(attrs={'class':'form-control','required': True,'tabindex': '2','placeholder': 'Ingrese solo números'}),
                'apellido_paterno' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '3'}),
                'apellido_materno' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '4'}),
                'nombres' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '5'}),
                'telefono' : forms.NumberInput(attrs={'class':'form-control','required': True,'tabindex': '6','placeholder': 'Ingrese solo números'}),
                'correo_electronico' : forms.EmailInput(attrs={'class':'form-control','required': True,'tabindex': '7'}),
                'provincia' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '8','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'distrito' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '9','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'ubigueo' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '10','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'red' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '11','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'microred' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '11','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'establecimiento' : forms.TextInput(attrs={'class':'form-control','required': True,'tabindex': '11','style': 'border-color: silver; color: silver;','readonly':'readonly'}),
                'condicion_laboral' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '12'}),
                'cargo' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '13'}),
                'perfil' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '14'}),
                'condicion' : forms.Select(attrs={'class':'form-control','required': True,'tabindex': '15'}),
                'user' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'cuenta_usuario' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'estado_usuario' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'nombre_completo' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'estado_auditoria' : forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                #####################
                'situacion_usuario': forms.Select(attrs={'class':'form-control','required': True,'tabindex': '16'}),
                'req_oficio': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '17'}),
                'req_resolucion': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '18'}),
                'req_formato_alta': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '19'}),
                'req_formato_excel': forms.ClearableFileInput(attrs={'class':'form-control', 'tabindex': '20'}),
                'dateTimeOfUpload_req_oficio': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
                'dateTimeOfUpload_req_resolucion': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}), 
                'dateTimeOfUpload_req_formato_alta':forms.TextInput(attrs={'class':'form-control','style': 'display: none'}), 
                'dateTimeOfUpload_req_formato_excel': forms.TextInput(attrs={'class':'form-control','style': 'display: none'}),
        }
        labels = {
            'tipo_documento':'Tipo Documento',
            'documento_identidad':'Numero de Documento',
            'correo_electronico':'Correo electronico',
            'red':'Red de Salud',
            'microred':'Microred',
            'establecimiento':'Establecimiento',
            'req_oficio': 'Oficio',
            'req_resolucion':'Resolucion',
            'req_formato_alta':'Formato Alta/Baja',
            'req_formato_excel':'Formato Excel',            
            'ubigueo': '',
            'estado_usuario': '',
            'nombre_completo': '',
            'situacion_usuario': 'Situacion',
        }