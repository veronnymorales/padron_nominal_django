from django.shortcuts import render

# TABLERO PAQUETE GESTANTE 
from django.db import connection
from django.http import JsonResponse
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo, Actualizacion
from django.db.models.functions import Substr
import logging

from .queries import (obtener_avance_paquete_compromiso, obtener_variables_paquete_compromiso, 
                    obtener_avance_regional_mensual_paquete_compromiso, obtener_avance_cobertura_paquete_compromiso, obtener_cobertura_por_edad, 
                    obtener_cobertura_por_red, obtener_cobertura_por_microred, obtener_cobertura_por_establecimiento, 
                    obtener_seguimiento_paquete_compromiso_red, obtener_seguimiento_paquete_compromiso_microred, obtener_seguimiento_paquete_compromiso_establecimiento)

# report excel
from django.http.response import HttpResponse
from django.views.generic.base import TemplateView
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import openpyxl
from openpyxl.utils import get_column_letter

from django.db.models.functions import Substr

from datetime import datetime
import locale

from django.db.models import IntegerField  # Importar IntegerField
from django.db.models.functions import Cast, Substr  # Importar Cast y Substr
# linea de border 
from openpyxl.utils import column_index_from_string

# Reporte excel
from datetime import datetime
import getpass  # Para obtener el nombre del usuario
from django.contrib.auth.models import User  # O tu modelo de usuario personalizado
from django.http import HttpResponse
from io import BytesIO
from django.contrib.auth import get_user_model
from django.contrib.auth.decorators import login_required

User = get_user_model()

from django.db.models import IntegerField             # Importar IntegerField
from django.db.models.functions import Cast, Substr     # Importar Cast y Substr

logger = logging.getLogger(__name__)


def BASE(request):
    actualizacion = Actualizacion.objects.all()
    return render(request,'paquete_compromiso/index_paquete_compromiso.html', {"actualizacion": actualizacion})
############################
## COMPONENTES Y GRAFICOS 
############################
def process_avance_por_region(resultados_avance_por_region):
    """
    Procesa los resultados del query de avance por región.
    Garantiza que siempre retorne un registro válido.
    
    Args:
        resultados_avance_por_region: Lista de diccionarios con keys: num, den, avance
    
    Returns:
        Dict con arrays r_numerador_resumen, r_denominador_resumen, r_avance_resumen (cada uno con un elemento)
    """
    data = {    
        'r_numerador_resumen': [],
        'r_denominador_resumen': [],
        'r_avance_resumen': []
    }
    
    # Si no hay resultados, retornar valores por defecto
    if not resultados_avance_por_region:
        data['r_numerador_resumen'] = [0]
        data['r_denominador_resumen'] = [0]
        data['r_avance_resumen'] = [0.0]
        print(f"[PROCESS] Sin datos de entrada, usando valores por defecto: {data}")
        return data
    
    # Procesar el primer (y único) registro
    row = resultados_avance_por_region[0]
    try:
        # Mapear las claves correctas del resultado
        numerador = row.get('num', 0)
        denominador = row.get('den', 0) 
        avance = row.get('avance', 0.0)

        # Asegurar que los valores no sean None
        numerador = numerador if numerador is not None else 0
        denominador = denominador if denominador is not None else 0
        avance = avance if avance is not None else 0.0

        data['r_numerador_resumen'].append(int(numerador))
        data['r_denominador_resumen'].append(int(denominador))
        data['r_avance_resumen'].append(float(avance))
        
        print(f"[PROCESS] Datos procesados: Num={numerador}, Den={denominador}, Avance={avance}%")
        
    except (KeyError, ValueError, TypeError) as e:
        logger.error(f"Error procesando datos: {e}, Row: {row}")
        # Valores por defecto en caso de error
        data['r_numerador_resumen'] = [0]
        data['r_denominador_resumen'] = [0]
        data['r_avance_resumen'] = [0.0]
    
    print(f"[PROCESS] Resultado final: {data}")
    return data

def process_variables_por_region(resultados_variables_por_region):
    data = {    
        'v_den_variable': [],
        'v_num_cred': [],
        'v_avance_cred': [],
        'v_num_cred_rn': [],
        'v_avance_cred_rn': [],
        'v_num_cred_mensual': [],    
        'v_avance_cred_mensual': [],
        'v_num_vac':[], 
        'v_avance_vac':[],
        'v_num_vac_antineumococica': [],
        'v_avance_vac_antineumococica': [],
        'v_num_vac_antipolio': [],
        'v_avance_vac_antipolio': [],
        'v_num_vac_pentavalente': [],
        'v_avance_vac_pentavalente': [],
        'v_num_vac_rotavirus': [],
        'v_avance_vac_rotavirus': [],
        'v_num_esq': [],
        'v_avance_esq': [],
        'v_num_esq4M': [],
        'v_avance_esq4M': [],
        'v_num_esq6M': [],
        'v_avance_esq6M': [],
        'v_num_esq6M_trat': [],
        'v_avance_esq6M_trat': [],
        'v_num_esq6M_multi': [],
        'v_avance_esq6M_multi': [],
        'v_num_dosaje_Hb': [],
        'v_avance_num_dosaje_Hb': [],
        'v_num_DNIemision': [],
        'v_avance_DNIemision': []
    }
    
    # Si no hay resultados, retornar valores por defecto
    if not resultados_variables_por_region:
        data['v_den_variable'] = [0]
        data['v_num_cred'] = [0]
        data['v_avance_cred'] = [0.0]
        data['v_num_cred_rn'] = [0]
        data['v_avance_cred_rn'] = [0.0]
        data['v_num_cred_mensual'] = [0]
        data['v_avance_cred_mensual'] = [0.0]
        data['v_num_vac'] = [0]
        data['v_avance_vac'] = [0.0]
        data['v_num_vac_antineumococica'] = [0]
        data['v_avance_vac_antineumococica'] = [0.0]
        data['v_num_vac_antipolio'] = [0]
        data['v_avance_vac_antipolio'] = [0.0]
        data['v_num_vac_pentavalente'] = [0]
        data['v_avance_vac_pentavalente'] = [0.0]
        data['v_num_vac_rotavirus'] = [0]
        data['v_avance_vac_rotavirus'] = [0.0]
        data['v_num_esq'] = [0]
        data['v_avance_esq'] = [0.0]
        data['v_num_esq4M'] = [0]
        data['v_avance_esq4M'] = [0.0]
        data['v_num_esq6M'] = [0]
        data['v_avance_esq6M'] = [0.0]
        data['v_num_esq6M_trat'] = [0]
        data['v_avance_esq6M_trat'] = [0.0]
        data['v_num_esq6M_multi'] = [0]
        data['v_avance_esq6M_multi'] = [0.0]
        data['v_num_dosaje_Hb'] = [0]
        data['v_avance_num_dosaje_Hb'] = [0.0]
        data['v_num_DNIemision'] = [0]
        data['v_avance_DNIemision'] = [0.0]
        return data
    
    # Procesar el primer (y único) registro
    row = resultados_variables_por_region[0]
    try:
        # Mapear las claves correctas del resultado        
        den_variable = row.get('den_variable', 0)
        num_cred = row.get('num_cred', 0)
        avance_cred = row.get('avance_cred', 0.0)
        num_cred_rn = row.get('num_cred_rn', 0)
        avance_cred_rn = row.get('avance_cred_rn', 0.0)
        num_cred_mensual = row.get('num_cred_mensual', 0)
        avance_cred_mensual = row.get('avance_cred_mensual', 0.0)
        num_vac = row.get('num_vac', 0)     
        avance_vac = row.get('avance_vac', 0.0)
        num_vac_antineumococica = row.get('num_vac_antineumococica', 0)
        avance_vac_antineumococica = row.get('avance_vac_antineumococica', 0.0)
        num_vac_antipolio = row.get('num_vac_antipolio', 0)
        avance_vac_antipolio = row.get('avance_vac_antipolio', 0.0)
        num_vac_pentavalente = row.get('num_vac_pentavalente', 0)
        avance_vac_pentavalente = row.get('avance_vac_pentavalente', 0.0)
        num_vac_rotavirus = row.get('num_vac_rotavirus', 0)
        avance_vac_rotavirus = row.get('avance_vac_rotavirus', 0.0)
        num_esq = row.get('num_esq', 0)
        avance_esq = row.get('avance_esq', 0.0)
        num_esq4M = row.get('num_esq4M', 0)
        avance_esq4M = row.get('avance_esq4M', 0.0)
        num_esq6M = row.get('num_esq6M', 0)
        avance_esq6M = row.get('avance_esq6M', 0.0)
        num_esq6M_trat = row.get('num_esq6M_trat', 0)
        avance_esq6M_trat = row.get('avance_esq6M_trat', 0.0)
        num_esq6M_multi = row.get('num_esq6M_multi', 0)
        avance_esq6M_multi = row.get('avance_esq6M_multi', 0.0)
        num_dosaje_Hb = row.get('num_dosaje_Hb', 0)
        avance_num_dosaje_Hb = row.get('avance_num_dosaje_Hb', 0.0)
        num_DNIemision = row.get('num_DNIemision', 0)
        avance_DNIemision = row.get('avance_DNIemision', 0.0)


        # Asegurar que los valores no sean None        
        den_variable = den_variable if den_variable is not None else 0
        num_cred = num_cred if num_cred is not None else 0
        avance_cred = avance_cred if avance_cred is not None else 0.0
        num_cred_rn = num_cred_rn if num_cred_rn is not None else 0
        avance_cred_rn = avance_cred_rn if avance_cred_rn is not None else 0.0
        num_cred_mensual = num_cred_mensual if num_cred_mensual is not None else 0
        avance_cred_mensual = avance_cred_mensual if avance_cred_mensual is not None else 0.0
        num_vac = num_vac if num_vac is not None else 0
        avance_vac = avance_vac if avance_vac is not None else 0.0
        num_vac_antineumococica = num_vac_antineumococica if num_vac_antineumococica is not None else 0
        avance_vac_antineumococica = avance_vac_antineumococica if avance_vac_antineumococica is not None else 0.0
        num_vac_antipolio = num_vac_antipolio if num_vac_antipolio is not None else 0
        avance_vac_antipolio = avance_vac_antipolio if avance_vac_antipolio is not None else 0.0
        num_vac_pentavalente = num_vac_pentavalente if num_vac_pentavalente is not None else 0
        avance_vac_pentavalente = avance_vac_pentavalente if avance_vac_pentavalente is not None else 0.0
        num_vac_rotavirus = num_vac_rotavirus if num_vac_rotavirus is not None else 0
        avance_vac_rotavirus = avance_vac_rotavirus if avance_vac_rotavirus is not None else 0.0
        num_esq = num_esq if num_esq is not None else 0
        avance_esq = avance_esq if avance_esq is not None else 0.0
        num_esq4M = num_esq4M if num_esq4M is not None else 0
        avance_esq4M = avance_esq4M if avance_esq4M is not None else 0.0
        num_esq6M = num_esq6M if num_esq6M is not None else 0
        avance_esq6M = avance_esq6M if avance_esq6M is not None else 0.0
        num_esq6M_trat = num_esq6M_trat if num_esq6M_trat is not None else 0
        avance_esq6M_trat = avance_esq6M_trat if avance_esq6M_trat is not None else 0.0
        num_esq6M_multi = num_esq6M_multi if num_esq6M_multi is not None else 0
        avance_esq6M_multi = avance_esq6M_multi if avance_esq6M_multi is not None else 0.0
        num_dosaje_Hb = num_dosaje_Hb if num_dosaje_Hb is not None else 0
        avance_num_dosaje_Hb = avance_num_dosaje_Hb if avance_num_dosaje_Hb is not None else 0.0
        num_DNIemision = num_DNIemision if num_DNIemision is not None else 0
        avance_DNIemision = avance_DNIemision if avance_DNIemision is not None else 0.0

        data['v_den_variable'].append(int(den_variable))
        data['v_num_cred'].append(int(num_cred))
        data['v_avance_cred'].append(float(avance_cred))
        data['v_num_cred_rn'].append(int(num_cred_rn))
        data['v_avance_cred_rn'].append(float(avance_cred_rn))
        data['v_num_cred_mensual'].append(int(num_cred_mensual))
        data['v_avance_cred_mensual'].append(float(avance_cred_mensual))
        data['v_num_vac'].append(int(num_vac))
        data['v_avance_vac'].append(float(avance_vac))
        data['v_num_vac_antineumococica'].append(int(num_vac_antineumococica))
        data['v_avance_vac_antineumococica'].append(float(avance_vac_antineumococica))
        data['v_num_vac_antipolio'].append(int(num_vac_antipolio))
        data['v_avance_vac_antipolio'].append(float(avance_vac_antipolio))
        data['v_num_vac_pentavalente'].append(int(num_vac_pentavalente))
        data['v_avance_vac_pentavalente'].append(float(avance_vac_pentavalente))
        data['v_num_vac_rotavirus'].append(int(num_vac_rotavirus))
        data['v_avance_vac_rotavirus'].append(float(avance_vac_rotavirus))
        data['v_num_esq'].append(int(num_esq))
        data['v_avance_esq'].append(float(avance_esq))
        data['v_num_esq4M'].append(int(num_esq4M))
        data['v_avance_esq4M'].append(float(avance_esq4M))
        data['v_num_esq6M'].append(int(num_esq6M))
        data['v_avance_esq6M'].append(float(avance_esq6M))
        data['v_num_esq6M_trat'].append(int(num_esq6M_trat))
        data['v_avance_esq6M_trat'].append(float(avance_esq6M_trat))
        data['v_num_esq6M_multi'].append(int(num_esq6M_multi))
        data['v_avance_esq6M_multi'].append(float(avance_esq6M_multi))
        data['v_num_dosaje_Hb'].append(int(num_dosaje_Hb))
        data['v_avance_num_dosaje_Hb'].append(float(avance_num_dosaje_Hb))
        data['v_num_DNIemision'].append(int(num_DNIemision))
        data['v_avance_DNIemision'].append(float(avance_DNIemision))
        
    except (KeyError, ValueError, TypeError) as e:
        logger.error(f"Error procesando datos: {e}, Row: {row}")
        # Valores por defecto en caso de error
        data['v_den_variable'] = [0]
        data['v_num_cred'] = [0]
        data['v_avance_cred'] = [0.0]
        data['v_num_cred_rn'] = [0]
        data['v_avance_cred_rn'] = [0.0]
        data['v_num_cred_mensual'] = [0]
        data['v_avance_cred_mensual'] = [0.0]
        data['v_num_vac'] = [0]
        data['v_avance_vac'] = [0.0]
        data['v_num_vac_antineumococica'] = [0]
        data['v_avance_vac_antineumococica'] = [0.0]
        data['v_num_vac_antipolio'] = [0]
        data['v_avance_vac_antipolio'] = [0.0]
        data['v_num_vac_pentavalente'] = [0]
        data['v_avance_vac_pentavalente'] = [0.0]
        data['v_num_vac_rotavirus'] = [0]
        data['v_avance_vac_rotavirus'] = [0.0]
        data['v_num_esq'] = [0]
        data['v_avance_esq'] = [0.0]
        data['v_num_esq4M'] = [0]
        data['v_avance_esq4M'] = [0.0]
        data['v_num_esq6M'] = [0]
        data['v_avance_esq6M'] = [0.0]
        data['v_num_esq6M_trat'] = [0]
        data['v_avance_esq6M_trat'] = [0.0]
        data['v_num_esq6M_multi'] = [0]
        data['v_avance_esq6M_multi'] = [0.0]
        data['v_num_dosaje_Hb'] = [0]
        data['v_avance_num_dosaje_Hb'] = [0.0]
        data['v_num_DNIemision'] = [0]
        data['v_avance_DNIemision'] = [0.0]

    print(f"[PROCESS] Resultado final: {data}")
    return data

def process_avance_regional_mensual_paquete_compromiso(resultados_avance_regional_mensual_paquete_compromiso):
    """Procesa los resultados del graficos"""
    data = {
        'num_1': [],
        'den_1': [],
        'cob_1': [],
        'num_2': [],
        'den_2': [],
        'cob_2': [],
        'num_3': [],
        'den_3': [],
        'cob_3': [],
        'num_4': [],
        'den_4': [],
        'cob_4': [],
        'num_5': [],
        'den_5': [],
        'cob_5': [],
        'num_6': [],
        'den_6': [],
        'cob_6': [],
        'num_7': [],
        'den_7': [],
        'cob_7': [],
        'num_8': [],
        'den_8': [],
        'cob_8': [],                
        'num_9': [],
        'den_9': [],
        'cob_9': [],
        'num_10': [],
        'den_10': [],
        'cob_10': [],
        'num_11': [],
        'den_11': [],
        'cob_11': [],
        'num_12': [],
        'den_12': [],
        'cob_12': [],
    }
    for index, row in enumerate(resultados_avance_regional_mensual_paquete_compromiso):
        try:
            # Verifica que el diccionario tenga las claves necesarias
            required_keys = {'num_1','den_1','cob_1','num_2','den_2','cob_2','num_3','den_3','cob_3','num_4','den_4','cob_4','num_5','den_5','cob_5','num_6','den_6','cob_6','num_7','den_7','cob_7','num_8','den_8','cob_8','num_9','den_9','cob_9','num_10','den_10','cob_10','num_11','den_11','cob_11','num_12','den_12','cob_12'}
            
            if not required_keys.issubset(row.keys()):
                raise ValueError(f"La fila {index} no tiene las claves necesarias: {row}")
            # Extraer cada valor, convirtiendo a float
            num_1_value = float(row.get('num_1', 0.0))
            den_1_value = float(row.get('den_1', 0.0))
            cob_1_value = float(row.get('cob_1', 0.0))
            num_2_value = float(row.get('num_2', 0.0))
            den_2_value = float(row.get('den_2', 0.0))
            cob_2_value = float(row.get('cob_2', 0.0))
            num_3_value = float(row.get('num_3', 0.0))
            den_3_value = float(row.get('den_3', 0.0))
            cob_3_value = float(row.get('cob_3', 0.0))
            num_4_value = float(row.get('num_4', 0.0))
            den_4_value = float(row.get('den_4', 0.0))
            cob_4_value = float(row.get('cob_4', 0.0))
            num_5_value = float(row.get('num_5', 0.0))
            den_5_value = float(row.get('den_5', 0.0))
            cob_5_value = float(row.get('cob_5', 0.0))
            num_6_value = float(row.get('num_6', 0.0))
            den_6_value = float(row.get('den_6', 0.0))
            cob_6_value = float(row.get('cob_6', 0.0))
            num_7_value = float(row.get('num_7', 0.0))
            den_7_value = float(row.get('den_7', 0.0))
            cob_7_value = float(row.get('cob_7', 0.0))
            num_8_value = float(row.get('num_8', 0.0))
            den_8_value = float(row.get('den_8', 0.0))
            cob_8_value = float(row.get('cob_8', 0.0))
            num_9_value = float(row.get('num_9', 0.0))
            den_9_value = float(row.get('den_9', 0.0))
            cob_9_value = float(row.get('cob_9', 0.0))
            num_10_value = float(row.get('num_10', 0.0))
            den_10_value = float(row.get('den_10', 0.0))
            cob_10_value = float(row.get('cob_10', 0.0))
            num_11_value = float(row.get('num_11', 0.0))
            den_11_value = float(row.get('den_11', 0.0))
            cob_11_value = float(row.get('cob_11', 0.0))
            num_12_value = float(row.get('num_12', 0.0))
            den_12_value = float(row.get('den_12', 0.0))
            cob_12_value = float(row.get('cob_12', 0.0))
            
            data['num_1'].append(num_1_value)
            data['den_1'].append(den_1_value)
            data['cob_1'].append(cob_1_value)
            data['num_2'].append(num_2_value)
            data['den_2'].append(den_2_value)
            data['cob_2'].append(cob_2_value)
            data['num_3'].append(num_3_value)
            data['den_3'].append(den_3_value)
            data['cob_3'].append(cob_3_value)
            data['num_4'].append(num_4_value)
            data['den_4'].append(den_4_value)
            data['cob_4'].append(cob_4_value)
            data['num_5'].append(num_5_value)
            data['den_5'].append(den_5_value)
            data['cob_5'].append(cob_5_value)
            data['num_6'].append(num_6_value)
            data['den_6'].append(den_6_value)
            data['cob_6'].append(cob_6_value)
            data['num_7'].append(num_7_value)
            data['den_7'].append(den_7_value)
            data['cob_7'].append(cob_7_value)
            data['num_8'].append(num_8_value)
            data['den_8'].append(den_8_value)
            data['cob_8'].append(cob_8_value)
            data['num_9'].append(num_9_value)
            data['den_9'].append(den_9_value)
            data['cob_9'].append(cob_9_value)
            data['num_10'].append(num_10_value)
            data['den_10'].append(den_10_value)
            data['cob_10'].append(cob_10_value)
            data['num_11'].append(num_11_value)
            data['den_11'].append(den_11_value)
            data['cob_11'].append(cob_11_value)
            data['num_12'].append(num_12_value)
            data['den_12'].append(den_12_value)
            data['cob_12'].append(cob_12_value)

        except Exception as e:
            logger.error(f"Error procesando la fila {index}: {str(e)}")
    return data

def process_avance_cobertura_paquete_compromiso(resultados_avance_cobertura_paquete_compromiso):
    """Procesa los resultados del graficos"""
    data = {
            'anio': [],
            'mes': [],
            'Ubigueo_Establecimiento':[],
            'Distrito': [],
            'Provincia': [],
            'Codigo_Red': [],
            'red': [],
            'Codigo_MicroRed': [],
            'microred': [],
            'Codigo_Unico': [],
            'Nombre_Establecimiento': [],
            'grupo_edad': [],
            'total_denominador': [],
            'total_numerador': [],
            'total_brecha': [],
            'cobertura_porcentaje': [],
    }   
    for row in resultados_avance_cobertura_paquete_compromiso:
        try:
            data['anio'].append(row['anio'])
            data['mes'].append(row['mes'])
            data['Ubigueo_Establecimiento'].append(row['Ubigueo_Establecimiento'])
            data['Distrito'].append(row['Distrito'])
            data['Provincia'].append(row['Provincia'])
            data['Codigo_Red'].append(row['Codigo_Red'])
            data['red'].append(row['red'])
            data['Codigo_MicroRed'].append(row['Codigo_MicroRed'])
            data['microred'].append(row['microred'])
            data['Codigo_Unico'].append(row['Codigo_Unico'])
            data['Nombre_Establecimiento'].append(row['Nombre_Establecimiento'])
            data['grupo_edad'].append(row['grupo_edad'])

            # Cambia null (None) a 0
            data['total_denominador'].append(row['total_denominador'] if row['total_denominador'] is not None else 0)
            data['total_numerador'].append(row['total_numerador'] if row['total_numerador'] is not None else 0)
            data['total_brecha'].append(row['total_brecha'] if row['total_brecha'] is not None else 0)
            data['cobertura_porcentaje'].append(row['cobertura_porcentaje'] if row['cobertura_porcentaje'] is not None else 0)
        except KeyError as e:
            logger.warning(f"Fila con estructura inválida (clave faltante: {e}): {row}")

    return data

def process_cobertura_por_edad(resultados_cobertura_por_edad):
    """Procesa los resultados del graficos"""
    data = {    
            'c_grupo_edad': [],
            'c_denominador': [],
            'c_numerador': [],
            'c_brecha': [],
            'c_cobertura': [],
    }
    for row in resultados_cobertura_por_edad:
        try:
            data['c_grupo_edad'].append(row['c_grupo_edad'])

            # Cambia null (None) a 0
            data['c_denominador'].append(row['c_denominador'] if row['c_denominador'] is not None else 0)
            data['c_numerador'].append(row['c_numerador'] if row['c_numerador'] is not None else 0)
            data['c_brecha'].append(row['c_brecha'] if row['c_brecha'] is not None else 0)
            data['c_cobertura'].append(row['c_cobertura'] if row['c_cobertura'] is not None else 0)
        except KeyError as e:
            logger.warning(f"Fila con estructura inválida (clave faltante: {e}): {row}")

    return data

def process_cobertura_por_red(resultados_cobertura_por_red):
    """Procesa los resultados del graficos"""
    data = {
            'r_red': [],
            'r_denominador': [],
            'r_numerador': [],
            'r_brecha': [],
            'r_cobertura': [],
    }   
    for row in resultados_cobertura_por_red:
        try:
            data['r_red'].append(row['r_red'])

            # Cambia null (None) a 0
            data['r_denominador'].append(row['r_denominador'] if row['r_denominador'] is not None else 0)
            data['r_numerador'].append(row['r_numerador'] if row['r_numerador'] is not None else 0)
            data['r_brecha'].append(row['r_brecha'] if row['r_brecha'] is not None else 0)
            data['r_cobertura'].append(row['r_cobertura'] if row['r_cobertura'] is not None else 0)
        except KeyError as e:
            logger.warning(f"Fila con estructura inválida (clave faltante: {e}): {row}")

    return data

def process_cobertura_por_microred(resultados_cobertura_por_microred):
    """Procesa los resultados del graficos"""
    data = {
            'm_microred': [],
            'm_denominador': [],
            'm_numerador': [],
            'm_brecha': [],
            'm_cobertura': [],
    }   
    for row in resultados_cobertura_por_microred:
        try:
            data['m_microred'].append(row['m_microred'])

            # Cambia null (None) a 0
            data['m_denominador'].append(row['m_denominador'] if row['m_denominador'] is not None else 0)
            data['m_numerador'].append(row['m_numerador'] if row['m_numerador'] is not None else 0)
            data['m_brecha'].append(row['m_brecha'] if row['m_brecha'] is not None else 0)
            data['m_cobertura'].append(row['m_cobertura'] if row['m_cobertura'] is not None else 0)
        except KeyError as e:
            logger.warning(f"Fila con estructura inválida (clave faltante: {e}): {row}")

    return data

def process_cobertura_por_establecimiento(resultados_cobertura_por_establecimiento):
    """Procesa los resultados del graficos"""
    data = {
            'e_establecimiento': [],
            'e_denominador': [],
            'e_numerador': [],
            'e_brecha': [],
            'e_cobertura': [],
    }   
    for row in resultados_cobertura_por_establecimiento:
        try:
            data['e_establecimiento'].append(row['e_establecimiento'])

            # Cambia null (None) a 0
            data['e_denominador'].append(row['e_denominador'] if row['e_denominador'] is not None else 0)
            data['e_numerador'].append(row['e_numerador'] if row['e_numerador'] is not None else 0)
            data['e_brecha'].append(row['e_brecha'] if row['e_brecha'] is not None else 0)
            data['e_cobertura'].append(row['e_cobertura'] if row['e_cobertura'] is not None else 0)
        except KeyError as e:
            logger.warning(f"Fila con estructura inválida (clave faltante: {e}): {row}")

    return data

#######################
## PANTALLA PRINCIPAL
#######################
def index_paquete_compromiso(request):
    actualizacion = Actualizacion.objects.all()

    # Capturamos el año que viene por GET
    anio = request.GET.get('anio', '2025')
    mes = request.GET.get('mes', '')
    provincia = request.GET.get('provincia', '')
    distrito = request.GET.get('distrito', '')

    if anio not in ['2024', '2025' , '2026']:
        anio = '2025'

    redes_h = (
        MAESTRO_HIS_ESTABLECIMIENTO
        .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL', Disa='JUNIN')
        .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
        .values('Red', 'codigo_red_filtrado')
        .distinct()
        .order_by('Red')
    )
    
    provincias_h = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )

    mes_seleccionado_inicio = request.GET.get('mes_inicio')
    mes_seleccionado_fin = request.GET.get('mes_fin')
    
    provincia_seleccionada = request.GET.get('provincia_h')
    distrito_seleccionado = request.GET.get('distrito_h')
    
    red_seleccionada = request.GET.get('red_h')
    microred_seleccionada = request.GET.get('p_microredes_establec_h')
    establecimiento_seleccionado = request.GET.get('p_establecimiento_h')

    if request.headers.get('x-requested-with') == 'XMLHttpRequest':
        try:
            print(f"Parámetros recibidos - Año: {anio}, Mes inicio: {mes_seleccionado_inicio}, Mes fin: {mes_seleccionado_fin}, Provincia: {provincia_seleccionada}, Distrito: {distrito_seleccionado}")
            resultados_avance_por_region = obtener_avance_paquete_compromiso(anio, mes_seleccionado_inicio, mes_seleccionado_fin, provincia_seleccionada, distrito_seleccionado)
            resultados_variables_por_region = obtener_variables_paquete_compromiso(anio, mes_seleccionado_inicio, mes_seleccionado_fin, provincia_seleccionada, distrito_seleccionado)
            
            
            
            resultados_avance_regional_mensual_paquete_compromiso = obtener_avance_regional_mensual_paquete_compromiso(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado)
            resultados_avance_cobertura_paquete_compromiso = obtener_avance_cobertura_paquete_compromiso(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado, provincia, distrito)
            resultados_cobertura_por_edad = obtener_cobertura_por_edad(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado, provincia, distrito)
            resultados_cobertura_por_red = obtener_cobertura_por_red(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado, provincia, distrito)
            resultados_cobertura_por_microred = obtener_cobertura_por_microred(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado, provincia, distrito)
            resultados_cobertura_por_establecimiento = obtener_cobertura_por_establecimiento(anio, mes, red_seleccionada, microred_seleccionada, establecimiento_seleccionado, provincia, distrito)
            print("resultados_avance_por_region:", resultados_avance_por_region)
            
            # Check for empty results and handle them
            if not resultados_avance_regional_mensual_paquete_compromiso:
                # You can return a specific message or an empty structure
                return JsonResponse({'error': 'No se encontraron datos para el ranking.'}, status=404)

            data = {
                **process_avance_por_region(resultados_avance_por_region),
                **process_variables_por_region(resultados_variables_por_region),
                
                **process_avance_regional_mensual_paquete_compromiso(resultados_avance_regional_mensual_paquete_compromiso),
                **process_avance_cobertura_paquete_compromiso(resultados_avance_cobertura_paquete_compromiso),
                **process_cobertura_por_edad(resultados_cobertura_por_edad),
                **process_cobertura_por_red(resultados_cobertura_por_red),
                **process_cobertura_por_microred(resultados_cobertura_por_microred),
                **process_cobertura_por_establecimiento(resultados_cobertura_por_establecimiento)
            }

            logger.info(f"Datos enviados como JSON: {data}")
            return JsonResponse(data)

        except Exception as e:
            logger.error(f"Error al obtener datos: {str(e)}")
            # Devolver una respuesta JSON con detalles del error
            return JsonResponse({'error': f"Error al obtener datos: {str(e)}"}, status=500)  # Incluye el error en la respuesta

    return render(request, 'paquete_compromiso/index_paquete_compromiso.html', {
        'mes_seleccionado_inicio': mes_seleccionado_inicio,
        'mes_seleccionado_fin': mes_seleccionado_fin,
        'actualizacion': actualizacion,
        'provincia_seleccionada': provincia_seleccionada,
        'distrito_seleccionado': distrito_seleccionado,
        'provincias_h': provincias_h,
    })

#####################################
## FILTROS HORIZONTAL
#####################################
##----------------------------------
## FILTROS HORIZONTAL POR SALUD
##----------------------------------
def get_establecimientos_paquete_compromiso_h(request,establecimiento_id):
    redes_h = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL',Disa='JUNIN')
                .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
                .values('Red','codigo_red_filtrado')
                .distinct()
                .order_by('Red')
    )
    provincias_h = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'redes_h': redes_h,
                'provincias_h': provincias_h,
                'mes_inicio':mes_inicio
    }
    return render(request,'paquete_compromiso/establecimientos_h.html', context)

def p_microredes_establec_paquete_compromiso_h(request):
    redes_param = request.GET.get('red_h') 
    microredes = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Codigo_Red=redes_param, Descripcion_Sector='GOBIERNO REGIONAL',Disa='JUNIN').values('Codigo_MicroRed','MicroRed').distinct()
    context = {
        'microredes': microredes,
        'is_htmx': True
    }
    return render(request, 'paquete_compromiso/partials/p_microredes_establec_h.html', context)

def p_establecimientos_paquete_compromiso_h(request):
    microredes = request.GET.get('p_microredes_establec_h', '')    
    codigo_red = request.GET.get('red_h', '')
    
    # Construir el filtro dinámicamente
    filtros = {
        'Descripcion_Sector': 'GOBIERNO REGIONAL',
        'Disa': 'JUNIN'
    }
    
    if microredes:
        filtros['Codigo_MicroRed'] = microredes
    if codigo_red:
        filtros['Codigo_Red__startswith'] = codigo_red
    
    establec = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(**filtros).values('Codigo_Unico','Nombre_Establecimiento').distinct()
    
    # Debug: contar establecimientos encontrados
    establec_list = list(establec)
    
    context= {
        'establec': establec_list
    }
    return render(request, 'paquete_compromiso/partials/p_establecimientos_h.html', context)

##-----------------------------------
## FILTROS HORIZONTAL POR MUNICIPIO
##-----------------------------------
def p_distritos_paquete_compromiso_h(request):
    provincias_h = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )
    
    provincia_param = request.GET.get('provincia_h')
    
    distritos = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(
        Ubigueo_Establecimiento__startswith=provincia_param,
        Descripcion_Sector='GOBIERNO REGIONAL'
    ).values('Ubigueo_Establecimiento', 'Distrito').distinct().order_by('Distrito')
    context = {
        'distritos': distritos,
        'provincias_h' :  provincias_h
    }
    return render(request, 'paquete_compromiso/partials/p_distritos.html', context)

###############################
## SEGUIMIENTO NOMINAL FILTROS
###############################

##---------------------------
## AMBITO DE SALUD
##---------------------------

## SEGUIMIENTO POR REDES
def get_redes_paquete_compromiso(request,redes_id):
    redes = (
            MAESTRO_HIS_ESTABLECIMIENTO
            .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL',Departamento='JUNIN')
            .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
            .values('Red','codigo_red_filtrado')
            .distinct()
            .order_by('Red')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    mes_fin = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'redes': redes,
                'mes_inicio':mes_inicio,
                'mes_fin':mes_fin,
    }
    
    return render(request, 'paquete_compromiso/components/salud/redes.html', context)

## SEGUIMIENTO POR MICRO-REDES
def get_microredes_paquete_compromiso(request, microredes_id):
    redes = (
            MAESTRO_HIS_ESTABLECIMIENTO
            .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL',Departamento='JUNIN')
            .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
            .values('Red','codigo_red_filtrado')
            .distinct()
            .order_by('Red')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    mes_fin = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'redes': redes,
                'mes_inicio':mes_inicio,
                'mes_fin':mes_fin,
    }
    
    return render(request, 'paquete_compromiso/components/salud/microredes.html', context)

def p_microredes_paquete_compromiso(request):
    redes_param = request.GET.get('red', '')
    

    
    # Consulta principal
    microredes = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(
        Codigo_Red__startswith=redes_param, 
        Descripcion_Sector='GOBIERNO REGIONAL', 
        Disa='JUNIN'
    ).values('Codigo_MicroRed','MicroRed').distinct()
    
    microredes_list = list(microredes)
    
    context = {
        'redes_param': redes_param,
        'microredes': microredes_list
    }
    
    # Si es una petición AJAX, devolver solo las opciones
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return render(request, 'paquete_compromiso/partials/p_microredes_options.html', context)
    
    return render(request, 'paquete_compromiso/partials/p_microredes.html', context)

## REPORTE POR ESTABLECIMIENTO
def get_establecimientos_paquete_compromiso(request,establecimiento_id):
    redes = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL',Disa='JUNIN')
                .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
                .values('Red','codigo_red_filtrado')
                .distinct()
                .order_by('Red')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    mes_fin = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'redes': redes,
                'mes_inicio':mes_inicio,
                'mes_fin':mes_fin,
    }
    return render(request,'paquete_compromiso/components/salud/establecimientos.html', context)

def p_microredes_establec_paquete_compromiso(request):
    redes_param = request.GET.get('red') 
    

    
    # Usar startswith en lugar de igualdad exacta
    microredes = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(
        Codigo_Red__startswith=redes_param, 
        Descripcion_Sector='GOBIERNO REGIONAL',
        Disa='JUNIN'
    ).values('Codigo_MicroRed','MicroRed').distinct()
    
    microredes_list = list(microredes)
    
    context = {
        'microredes': microredes_list,
        'is_htmx': True
    }
    return render(request, 'paquete_compromiso/partials/p_microredes_establec.html', context)

def p_establecimientos_paquete_compromiso(request):
    # Limpiar parámetros - obtener solo el primer valor si hay duplicados
    microredes = request.GET.get('p_microredes_establec', '').strip()
    
    # Para el código de red, obtener todos los valores y usar el primero no vacío
    red_values = request.GET.getlist('red')
    codigo_red = ''
    for red_val in red_values:
        if red_val and red_val.strip():
            codigo_red = red_val.strip()
            break
    
    # Construir el filtro dinámicamente
    filtros = {
        'Descripcion_Sector': 'GOBIERNO REGIONAL',
        'Disa': 'JUNIN'
    }
    
    # Agregar filtro de red si está disponible
    if codigo_red:
        filtros['Codigo_Red__startswith'] = codigo_red
    
    # Agregar filtro de microred si está presente
    if microredes:
        filtros['Codigo_MicroRed'] = microredes
        
        # Si hay microred pero no hay red, intentar obtener la red de la microred
        if not codigo_red:
            try:
                red_desde_microred = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(
                    Codigo_MicroRed=microredes,
                    Descripcion_Sector='GOBIERNO REGIONAL',
                    Disa='JUNIN'
                ).values('Codigo_Red').first()
                
                if red_desde_microred:
                    codigo_red_obtenido = red_desde_microred['Codigo_Red']
                    filtros['Codigo_Red'] = codigo_red_obtenido
            except Exception as e:
                pass
        
    try:
        establec = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(**filtros).values('Codigo_Unico','Nombre_Establecimiento').distinct()
        establec_list = list(establec)
    except Exception as e:
        establec_list = []
    
    context= {
        'establec': establec_list
    }
    
    # Si es una petición AJAX, devolver solo las opciones
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return render(request, 'paquete_compromiso/partials/p_establecimientos_options.html', context)
    
    return render(request, 'paquete_compromiso/partials/p_establecimientos.html', context)

##-------------------------------
## AMBITO DE MUNICIPIO
##-------------------------------

## SEGUIMIENTO POR PROVINCIA
def get_provincias_paquete_compromiso(request, provincia_id):
    provincias = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )
    context = {
                'provincias': provincias,
            }
    
    return render(request, 'paquete_compromiso/components/municipio/provincias.html', context)

## SEGUIMIENTO POR DISTRITOS
def get_distritos_paquete_compromiso(request, distrito_id):
    provincias = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(periodo_filtrado=Substr('Periodo', 1, 6))
                .values('Mes','periodo_filtrado')
                .order_by('NroMes')
                .distinct()
    ) 
    mes_fin = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(periodo_filtrado=Substr('Periodo', 1, 6))
                .values('Mes','periodo_filtrado')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'provincias': provincias,
                'mes_inicio':mes_inicio,
                'mes_fin':mes_fin,
    }
    return render(request, 'paquete_compromiso/components/municipio/distritos.html', context)

def p_distrito_paquete_compromiso(request):
    provincia_param = request.GET.get('provincia', '')

    # Filtra los establecimientos por sector "GOBIERNO REGIONAL"
    establecimientos = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')

    # Filtra los establecimientos por el código de la provincia
    if provincia_param:
        establecimientos = establecimientos.filter(Ubigueo_Establecimiento__startswith=provincia_param[:4])
    # Selecciona el distrito y el código Ubigueo
    distritos = establecimientos.values('Distrito', 'Ubigueo_Establecimiento').distinct().order_by('Distrito')
    
    context = {
        'provincia': provincia_param,
        'distritos': distritos
    }
    return render(request, 'paquete_compromiso/partials/p_distritos.html', context)


########################################
## SEGUIMIENTO REPORTE EXCEL FILTROS
#######################################

## REPORTE DE EXCEL
class RptPnPoblacionRed(TemplateView):
    def get(self, request, *args, **kwargs):
        # Variables ingresadas
        p_departamento = 'JUNIN'
        p_red = request.GET.get('red', '')
        p_microred = ''
        p_establec = ''
        p_edades = request.GET.get('edades', '')
        p_cumple = request.GET.get('cumple', '') 

        # Creación de la consulta
        resultado_seguimiento = obtener_seguimiento_paquete_compromiso_red(p_departamento, p_red, p_edades, p_cumple)
                
        wb = Workbook()
        
        consultas = [
                ('Seguimiento', resultado_seguimiento)
        ]
        
        for index, (sheet_name, results) in enumerate(consultas):
            if index == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)
        
            fill_worksheet_paquete_compromiso(ws, results)
        
        ##########################################################################          
        # Establecer el nombre del archivo
        nombre_archivo = "rpt_paquete_compromiso_red.xlsx"
        # Definir el tipo de respuesta que se va a dar
        response = HttpResponse(content_type="application/ms-excel")
        contenido = "attachment; filename={}".format(nombre_archivo)
        response["Content-Disposition"] = contenido
        wb.save(response)

        return response

class RptPnPoblacionMicroRed(TemplateView):
    def get(self, request, *args, **kwargs):
        # Variables ingresadas
        p_departamento = 'JUNIN'
        p_red = request.GET.get('red', '')
        p_microred = request.GET.get('microredes', '')
        p_establec = ''
        p_edades =  request.GET.get('edades','')
        p_cumple = request.GET.get('cumple', '') 
        # Creación de la consulta
        resultado_seguimiento_microred = obtener_seguimiento_paquete_compromiso_microred(p_departamento, p_red, p_microred, p_edades, p_cumple)

        wb = Workbook()
        
        consultas = [
                ('Seguimiento', resultado_seguimiento_microred)
        ]
        
        for index, (sheet_name, results) in enumerate(consultas):
            if index == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)
        
            fill_worksheet_paquete_compromiso(ws, results)
        
        ##########################################################################          
        # Establecer el nombre del archivo
        nombre_archivo = "rpt_paquete_compromiso_microred.xlsx"
        # Definir el tipo de respuesta que se va a dar
        response = HttpResponse(content_type="application/ms-excel")
        contenido = "attachment; filename={}".format(nombre_archivo)
        response["Content-Disposition"] = contenido
        wb.save(response)

        return response

class RptPnPoblacionEstablec(TemplateView):
    def get(self, request, *args, **kwargs):
        # Variables ingresadas
        p_departamento = 'JUNIN'
        p_red = request.GET.get('red','')
        p_microred = request.GET.get('p_microredes_establec','')  # Corregido
        p_establec = request.GET.get('p_establecimiento','')
        p_mes = request.GET.get('mes', '')
        p_edades = request.GET.get('edades', '')
        # Manejo seguro de fechas - usar valores por defecto si no están presentes
        p_cumple = request.GET.get('cumple', '')
        
        # Creación de la consulta
        resultado_seguimiento = obtener_seguimiento_paquete_compromiso_establecimiento(p_departamento,p_establec,p_edades,p_cumple)
                
        wb = Workbook()
        
        consultas = [
                ('Seguimiento', resultado_seguimiento)
        ]
        
        for index, (sheet_name, results) in enumerate(consultas):
            if index == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)
        
            fill_worksheet_paquete_compromiso(ws, results)
        
        ##########################################################################          
        # Establecer el nombre del archivo
        nombre_archivo = "rpt_paquete_compromiso_establecimiento.xlsx"
        # Definir el tipo de respuesta que se va a dar
        response = HttpResponse(content_type="application/ms-excel")
        contenido = "attachment; filename={}".format(nombre_archivo)
        response["Content-Disposition"] = contenido
        wb.save(response)

        return response


def fill_worksheet_paquete_compromiso(ws, results): 
    # cambia el alto de la columna
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 12
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 25
    
    # cambia el ancho de la columna
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 9
    ws.column_dimensions['D'].width = 6
    ws.column_dimensions['E'].width = 9
    ws.column_dimensions['F'].width = 9
    ws.column_dimensions['G'].width = 9
    ws.column_dimensions['H'].width = 17
    ws.column_dimensions['I'].width = 9
    ws.column_dimensions['J'].width = 35
    ws.column_dimensions['K'].width = 6
    ws.column_dimensions['L'].width = 9
    ws.column_dimensions['M'].width = 9
    ws.column_dimensions['N'].width = 4
    ws.column_dimensions['O'].width = 4
    ws.column_dimensions['P'].width = 4
    ws.column_dimensions['Q'].width = 22
    ws.column_dimensions['R'].width = 25
    ws.column_dimensions['S'].width = 35
    ws.column_dimensions['T'].width = 25
    ws.column_dimensions['U'].width = 7
    ws.column_dimensions['V'].width = 9
    ws.column_dimensions['W'].width = 15
    ws.column_dimensions['X'].width = 15
    ws.column_dimensions['Y'].width = 9
    ws.column_dimensions['Z'].width = 25    
    ws.column_dimensions['AA'].width = 9
    ws.column_dimensions['AB'].width = 9       
    ws.column_dimensions['AC'].width = 9
    ws.column_dimensions['AD'].width = 9    
    ws.column_dimensions['AE'].width = 9
    ws.column_dimensions['AF'].width = 9   
    ws.column_dimensions['AG'].width = 9    
    ws.column_dimensions['AH'].width = 9   
    ws.column_dimensions['AI'].width = 25
    ws.column_dimensions['AJ'].width = 9    
    ws.column_dimensions['AK'].width = 25
    ws.column_dimensions['AL'].width = 18    
    ws.column_dimensions['AM'].width = 9
    ws.column_dimensions['AN'].width = 9    
    ws.column_dimensions['AO'].width = 25
    ws.column_dimensions['AP'].width = 9    
    ws.column_dimensions['AQ'].width = 9
    ws.column_dimensions['AR'].width = 9    
    ws.column_dimensions['AS'].width = 25
    ws.column_dimensions['AT'].width = 9   
    ws.column_dimensions['AU'].width = 9
    ws.column_dimensions['AV'].width = 18
    ws.column_dimensions['AW'].width = 9
    ws.column_dimensions['AX'].width = 9
    ws.column_dimensions['AY'].width = 13
    ws.column_dimensions['AZ'].width = 9
    ws.column_dimensions['BA'].width = 13
    ws.column_dimensions['BB'].width = 10
    ws.column_dimensions['BC'].width = 15
    ws.column_dimensions['BD'].width = 9
    ws.column_dimensions['BE'].width = 11
    ws.column_dimensions['BF'].width = 5
    ws.column_dimensions['BG'].width = 9
    ws.column_dimensions['BH'].width = 6
    ws.column_dimensions['BI'].width = 6
    ws.column_dimensions['BJ'].width = 25
    ws.column_dimensions['BK'].width = 7
    ws.column_dimensions['BL'].width = 5 
    ws.column_dimensions['BM'].width = 9
    ws.column_dimensions['BN'].width = 5 
    ws.column_dimensions['BO'].width = 18
    ws.column_dimensions['BP'].width = 5
    ws.column_dimensions['BQ'].width = 18
    ws.column_dimensions['BR'].width = 9
    ws.column_dimensions['BS'].width = 5
    ws.column_dimensions['BT'].width = 18
    ws.column_dimensions['BU'].width = 18
    ws.column_dimensions['BV'].width = 18
    ws.column_dimensions['BW'].width = 5
    
    # linea de division
    ws.freeze_panes = 'H7'
    # Configuración del fondo y el borde
    # Definir el color usando formato aRGB (opacidad completa 'FF' + color RGB)
    fill = PatternFill(start_color='FF60D7E0', end_color='FF60D7E0', fill_type='solid')
    # Definir el color anaranjado usando formato aRGB
    orange_fill = PatternFill(start_color='FFE0A960', end_color='FFE0A960', fill_type='solid')
    # Definir los estilos para gris
    gray_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
    # Definir el estilo de color verde
    green_fill = PatternFill(start_color='FF60E0B3', end_color='FF60E0B3', fill_type='solid')
    # Definir el estilo de color amarillo
    yellow_fill = PatternFill(start_color='FFE0DE60', end_color='FFE0DE60', fill_type='solid')
    # Definir el estilo de color azul
    blue_fill = PatternFill(start_color='FF60A2E0', end_color='FF60A2E0', fill_type='solid')
    # Definir el estilo de color verde 2
    green_fill_2 = PatternFill(start_color='FF60E07E', end_color='FF60E07E', fill_type='solid')
    
    green_font = Font(name='Arial', size=8, color='00FF00')  # Verde
    red_font = Font(name='Arial', size=8, color='FF0000')    # Rojo
    
    
    border = Border(left=Side(style='thin', color='00B0F0'),
                    right=Side(style='thin', color='00B0F0'),
                    top=Side(style='thin', color='00B0F0'),
                    bottom=Side(style='thin', color='00B0F0'))
    borde_plomo = Border(left=Side(style='thin', color='A9A9A9'), # Plomo
                    right=Side(style='thin', color='A9A9A9'), # Plomo
                    top=Side(style='thin', color='A9A9A9'), # Plomo
                    bottom=Side(style='thin', color='A9A9A9')) # Plomo
    
        # Configuración del fondo y el borde
    # Definir el color usando formato aRGB (opacidad completa 'FF' + color RGB)
    fill = PatternFill(start_color='FF60D7E0', end_color='FF60D7E0', fill_type='solid')
    # Definir el color anaranjado usando formato aRGB
    orange_fill = PatternFill(start_color='FFE0A960', end_color='FFE0A960', fill_type='solid')
    # Definir los estilos para gris
    gray_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
    # Definir el estilo de color verde
    green_fill = PatternFill(start_color='FF60E0B3', end_color='FF60E0B3', fill_type='solid')
    # Definir el estilo de color amarillo
    yellow_fill = PatternFill(start_color='FFE0DE60', end_color='FFE0DE60', fill_type='solid')
    # Definir el estilo de color azul
    blue_fill = PatternFill(start_color='FF60A2E0', end_color='FF60A2E0', fill_type='solid')
    # Definir el estilo de color verde 2
    green_fill_2 = PatternFill(start_color='FF60E07E', end_color='FF60E07E', fill_type='solid')   
    # Definir el estilo de relleno celeste
    celeste_fill = PatternFill(start_color='FF87CEEB', end_color='FF87CEEB', fill_type='solid')
    # Morado más claro
    morado_claro_fill = PatternFill(start_color='FFE9D8FF', end_color='FFE9D8FF', fill_type='solid')
    # Plomo más claro
    plomo_claro_fill = PatternFill(start_color='FFEDEDED', end_color='FFEDEDED', fill_type='solid')
    # Azul más claro
    azul_claro_fill = PatternFill(start_color='FFD8EFFA', end_color='FFD8EFFA', fill_type='solid')
    # Naranja más claro
    naranja_claro_fill = PatternFill(start_color='FFFFEBD8', end_color='FFFFEBD8', fill_type='solid')
    # Verde más claro
    verde_claro_fill = PatternFill(start_color='FFBDF7BD', end_color='FFBDF7BD', fill_type='solid')
    # Guinda (bordó / burdeos)
    guinda_claro_fill = PatternFill(start_color='FFE8A8A6', end_color='FFE8A8A6', fill_type='solid')

        
    green_font = Font(name='Arial', size=8, color='00FF00')  # Verde
    red_font = Font(name='Arial', size=8, color='FF0000')    # Rojo
    
    border = Border(left=Side(style='thin', color='00B0F0'),
                    right=Side(style='thin', color='00B0F0'),
                    top=Side(style='thin', color='00B0F0'),
                    bottom=Side(style='thin', color='00B0F0'))
    borde_plomo = Border(left=Side(style='thin', color='A9A9A9'), # Plomo
                    right=Side(style='thin', color='A9A9A9'), # Plomo
                    top=Side(style='thin', color='A9A9A9'), # Plomo
                    bottom=Side(style='thin', color='A9A9A9')) # Plomo
    
    borde_plomo = Border(left=Side(style='thin', color='A9A9A9'), # Plomo
                    right=Side(style='thin', color='A9A9A9'), # Plomo
                    top=Side(style='thin', color='A9A9A9'), # Plomo
                    bottom=Side(style='thin', color='A9A9A9')) # Plomo
    
    border_negro = Border(left=Side(style='thin', color='000000'), # negro
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'), 
            bottom=Side(style='thin', color='000000')) 
    
    # Merge cells 
    # numerador y denominador
    ws.merge_cells('B5:Q5') 
    ws.merge_cells('R5:AB5')
    ws.merge_cells('AC5:AG5')
    ws.merge_cells('AH5:AO5')
    ws.merge_cells('AQ5:AV5')
    ws.merge_cells('AW5:BD5')
    ws.merge_cells('BF5:BW5')

    # Combina cela
    ws['B5'] = 'DATOS DEL MENOR - VARIABLES DE IDENTIFICACION'
    ws['R5'] = 'DIRECCION COMPLETA DEL MENOR'
    ws['AC5'] = 'VISITAS DOMICILARIAS'
    ws['AH5'] = 'VARIBALES CON INFORMACION DE SALUD'
    ws['AQ5'] = 'DATOS DE LA MADRE'
    ws['AW5'] = 'AUDITORIA DE LOS REGISTROS'
    ws['BF5'] = 'INFORMACION HIS MINSA - ULTIMA ATENCION REGIONAL'

    ### numerador y denominador 
    
    ws['B5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['B5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['B5'].fill = fill
    ws['B5'].border = border_negro
    
    ws['R5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['R5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['R5'].fill = gray_fill
    ws['R5'].border = border_negro
    
    ### intervalo 
    ws['AC5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AC5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['AC5'].fill = green_fill
    ws['AC5'].border = border_negro
    
    ws['AH5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AH5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['AH5'].fill = blue_fill
    ws['AH5'].border = border_negro

    ws['AQ5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AQ5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['AQ5'].fill = yellow_fill
    ws['AQ5'].border = border_negro
    
    ws['AW5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AW5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['AW5'].fill = green_fill_2
    ws['AW5'].border = border_negro
    
    ws['BF5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BF5'].font = Font(name = 'Arial', size= 10, bold = True)
    ws['BF5'].fill = celeste_fill
    ws['BF5'].border = border_negro
    
    ### BORDE DE CELDAS CONBINADAS
    
    # NUM y DEN
    inicio_columna = 'B'
    fin_columna = 'BW'
    fila = 5
    from openpyxl.utils import column_index_from_string
    # Convertir letras de columna a índices numéricos
    indice_inicio = column_index_from_string(inicio_columna)
    indice_fin = column_index_from_string(fin_columna)
    # Iterar sobre las columnas en el rango especificado
    for col in range(indice_inicio, indice_fin + 1):
        celda = ws.cell(row=fila, column=col)
        celda.border = border_negro
    
    # NUM y DEN
    inicio_columna = 'B'
    fin_columna = 'BW'
    fila = 6
    from openpyxl.utils import column_index_from_string
    # Convertir letras de columna a índices numéricos
    indice_inicio = column_index_from_string(inicio_columna)
    indice_fin = column_index_from_string(fin_columna)
    # Iterar sobre las columnas en el rango especificado
    for col in range(indice_inicio, indice_fin + 1):
        celda = ws.cell(row=fila, column=col)
        celda.border = border_negro
            
    ##### imprimer fecha y hora del reporte
    fecha_hora_actual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    nombre_usuario = getpass.getuser()

    # Obtener el usuario actualmente autenticado
    try:
        user = User.objects.get(is_active=True)
    except User.DoesNotExist:
        user = None
    except User.MultipleObjectsReturned:
        # Manejar el caso donde hay múltiples usuarios activos
        user = User.objects.filter(is_active=True).first()  # Por ejemplo, obtener el primero
    # Asignar fecha y hora a la celda A1
    ws['V1'].value = 'Fecha y Hora:'
    ws['W1'].value = fecha_hora_actual

    # Asignar nombre de usuario a la celda A2
    ws['V2'].value = 'Usuario:'
    ws['W2'].value = nombre_usuario
    
    # Formatear las etiquetas en negrita
    etiqueta_font = Font(name='Arial', size=8)
    ws['V1'].font = etiqueta_font
    ws['W1'].font = etiqueta_font
    ws['V2'].font = etiqueta_font
    ws['W2'].font = etiqueta_font

    # Alinear el texto
    ws['V1'].alignment = Alignment(horizontal="right", vertical="center")
    ws['W1'].alignment = Alignment(horizontal="left", vertical="center")
    ws['V2'].alignment = Alignment(horizontal="right", vertical="center")
    ws['W2'].alignment = Alignment(horizontal="left", vertical="center")
    
    ## crea titulo del reporte
    ws['B1'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B1'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B1'] = 'OFICINA DE TECNOLOGIAS DE LA INFORMACION'
    
    ws['B2'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B2'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B2'] = 'DIRECCION REGIONAL DE SALUD JUNIN'
    
    ws['B4'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B4'].font = Font(name = 'Arial', size= 12, bold = True)
    ws['B4'] = 'SEGUIMIENTO NOMINAL DEL PADRON NOMINAL DE LA REGION JUNIN'
    
    ws['B3'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B3'].font = Font(name = 'Arial', size= 7, color='0000CC')
    ws['B3'] ='El usuario se compromete a mantener la confidencialidad de los datos personales que conozca como resultado del reporte realizado, cumpliendo con lo establecido en la Ley N° 29733 - Ley de Protección de Datos Personales y sus normas complementarias.'
        
    ws['AP5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AP5'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AP5'].fill = celeste_fill
    
    ws['BE5'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BE5'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BE5'].fill = orange_fill
    
    ws['B6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['B6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['B6'].fill = fill
    ws['B6'].border = border_negro
    ws['B6'] = 'ID'
    
    ws['C6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['C6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['C6'].fill = fill
    ws['C6'].border = border_negro 
    ws['C6'] = 'COD PADRON'
    
    ws['D6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['D6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['D6'].fill = fill
    ws['D6'].border = border
    ws['D6'] = 'TIP DOC'      
    
    ws['E6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['E6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['E6'].fill = fill
    ws['E6'].border = border
    ws['E6'] = 'CNV' 
    
    ws['F6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['F6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['F6'].fill = fill
    ws['F6'].border = border
    ws['F6'] = 'CUI'     
    
    ws['G6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['G6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['G6'].fill = fill
    ws['G6'].border = border
    ws['G6'] = 'DNI'    
    
    ws['H6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['H6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['H6'].fill = fill
    ws['H6'].border = border
    ws['H6'] = 'ESTADO TRAMITE'    
    
    ws['I6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['I6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['I6'].fill = fill
    ws['I6'].border = border
    ws['I6'] = 'IND DNI'    
    
    ws['J6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['J6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['J6'].fill = fill
    ws['J6'].border = border
    ws['J6'] = 'NOMBRE DE NIÑO/A'  
    
    ws['K6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['K6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['K6'].fill = fill
    ws['K6'].border = border
    ws['K6'] = 'SEXO'  
    
    ws['L6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['L6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['L6'].fill = fill
    ws['L6'].border = border
    ws['L6'] = 'SEGURO'  
    
    ws['M6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['M6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['M6'].fill = fill
    ws['M6'].border = border
    ws['M6'] = 'FECHA NAC'  
    
    ws['N6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['N6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['N6'].fill = fill
    ws['N6'].border = border
    ws['N6'] = 'ED A'  
    
    ws['O6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['O6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['O6'].fill = fill
    ws['O6'].border = border
    ws['O6'] = 'ED M' 
    
    ws['P6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['P6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['P6'].fill = fill
    ws['P6'].border = border
    ws['P6'] = 'ED D'  
    
    ws['Q6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Q6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Q6'].fill = fill
    ws['Q6'].border = border
    ws['Q6'] = 'EDAD ACTUAL' 
    
    ws['R6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['R6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['R6'].fill = gray_fill
    ws['R6'].border = border
    ws['R6'] = 'EJE VIAL'  
    
    ws['S6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['S6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['S6'].fill = gray_fill
    ws['S6'].border = border
    ws['S6'] = 'DIRECCION' 
    
    ws['T6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['T6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['T6'].fill = gray_fill
    ws['T6'].border = border
    ws['T6'] = 'REFERENCIA'  
    
    ws['U6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['U6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['U6'].fill = gray_fill
    ws['U6'].border = border
    ws['U6'] = 'UBIGUEO' 
    
    ws['V6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['V6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['V6'].fill = gray_fill
    ws['V6'].border = border
    ws['V6'] = 'DEP'  
    
    ws['W6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['W6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['W6'].fill = gray_fill
    ws['W6'].border = border
    ws['W6'] = 'PROVINCIA' 
        
    ws['X6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['X6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['X6'].fill = gray_fill
    ws['X6'].border = border
    ws['X6'] = 'DISTRITO' 

    ws['Y6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Y6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Y6'].fill = gray_fill
    ws['Y6'].border = border
    ws['Y6'] = 'COD CP'  
    
    ws['Z6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Z6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Z6'].fill = gray_fill
    ws['Z6'].border = border
    ws['Z6'] = 'CENTRO POBLADO' 

    ws['AA6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AA6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AA6'].fill = gray_fill
    ws['AA6'].border = border
    ws['AA6'] = 'AREA'  
    
    ws['AB6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AB6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AB6'].fill = gray_fill
    ws['AB6'].border = border
    ws['AB6'] = 'IND DIR' 
    
    ws['AC6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AC6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AC6'].fill = green_fill
    ws['AC6'].border = border
    ws['AC6'] = 'MENOR VISITADO'  
    
    ws['AD6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AD6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AD6'].fill = green_fill
    ws['AD6'].border = border
    ws['AD6'] = 'MENOR ENCON' 
    
    ws['AE6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AE6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AE6'].fill = green_fill
    ws['AE6'].border = border
    ws['AE6'] = 'FECHA VISITA'  
    
    ws['AF6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AF6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AF6'].fill = green_fill
    ws['AF6'].border = border
    ws['AF6'] = 'IND VIS' 
    
    ws['AG6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AG6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AG6'].fill = green_fill
    ws['AG6'].border = border
    ws['AG6'] = 'TRANSITO'  
    
    ws['AH6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AH6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AH6'].fill = blue_fill
    ws['AH6'].border = border
    ws['AH6'] = 'COD NAC' 
    
    ws['AI6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AI6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AI6'].fill = blue_fill
    ws['AI6'].border = border
    ws['AI6'] = 'ESTABLECIMIENTO DE NACIMIENTO'  
    
    ws['AJ6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AJ6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AJ6'].fill = blue_fill
    ws['AJ6'].border = border
    ws['AJ6'] = 'COD EESS' 
    
    ws['AK6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AK6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AK6'].fill = blue_fill
    ws['AK6'].border = border
    ws['AK6'] = 'ESTABLECIMIENTO DE ATENCION' 
    
    ws['AL6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AL6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AL6'].fill = blue_fill
    ws['AL6'].border = border
    ws['AL6'] = 'FRECUENCIA ATENCION'  
    
    ws['AM6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AM6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AM6'].fill = blue_fill
    ws['AM6'].border = border
    ws['AM6'] = 'IND SALUD' 
    
    ws['AN6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AN6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AN6'].fill = blue_fill
    ws['AN6'].border = border
    ws['AN6'] = 'COD ADS'  
    
    ws['AO6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AO6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AO6'].fill = blue_fill
    ws['AO6'].border = border
    ws['AO6'] = 'ESTABLECIMIENTO DE ADSCRIPCION (SIS)' 
    
    ws['AP6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AP6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AP6'].fill = celeste_fill
    ws['AP6'].border = border_negro
    ws['AP6'] = 'PROG. SOCIAL'  
    
    ws['AQ6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AQ6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AQ6'].fill = yellow_fill
    ws['AQ6'].border = border
    ws['AQ6'] = 'TIPO DOC' 
    
    ws['AR6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AR6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AR6'].fill = yellow_fill
    ws['AR6'].border = border
    ws['AR6'] = 'NUM DOC'  
    
    ws['AS6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AS6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AS6'].fill = yellow_fill
    ws['AS6'].border = border
    ws['AS6'] = 'NOMBRE DE LA MADRE' 
    
    ws['AT6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AT6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AT6'].fill = yellow_fill
    ws['AT6'].border = border
    ws['AT6'] = 'CELULAR'  
    
    ws['AU6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AU6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AU6'].fill = yellow_fill
    ws['AU6'].border = border
    ws['AU6'] = 'IND MADRE' 
    
    ws['AV6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AV6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AV6'].fill = yellow_fill
    ws['AV6'].border = border
    ws['AV6'] = 'CORREO ELECTRONICO' 
    
    ws['AW6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AW6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AW6'].fill = green_fill_2
    ws['AW6'].border = border
    ws['AW6'] = 'ESTADO REGISTRO' 
    
    ws['AX6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AX6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AX6'].fill = green_fill_2
    ws['AX6'].border = border
    ws['AX6'] = 'FECHA REGISTRO' 
    
    ws['AY6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AY6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AY6'].fill = green_fill_2
    ws['AY6'].border = border
    ws['AY6'] = 'USUARIO'  
    
    ws['AZ6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AZ6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AZ6'].fill = green_fill_2
    ws['AZ6'].border = border
    ws['AZ6'] = 'FECHA MODIF'       
    
    ws['BA6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BA6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BA6'].fill = green_fill_2
    ws['BA6'].border = border
    ws['BA6'] = 'USUARIO MODIF' 
    
    ws['BB6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BB6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BB6'].fill = green_fill_2
    ws['BB6'].border = border
    ws['BB6'] = 'ENTIDAD'  
    
    ws['BC6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BC6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BC6'].fill = green_fill_2
    ws['BC6'].border = border
    ws['BC6'] = 'TIPO REGISTRO'  
    
    ws['BD6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BD6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BD6'].fill = green_fill_2
    ws['BD6'].border = border
    ws['BD6'] = 'FECHA CORTE'  
    
    ws['BE6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BE6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BE6'].fill = orange_fill
    ws['BE6'].border = border
    ws['BE6'] = 'INDICADOR' 
    
    ws['BF6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BF6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BF6'].fill = celeste_fill
    ws['BF6'].border = border
    ws['BF6'] = 'DEN'  
    
    ws['BG6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BG6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BG6'].fill = celeste_fill
    ws['BG6'].border = border
    ws['BG6'] = 'FECHA ULT ATE'       
    
    ws['BH6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BH6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BH6'].fill = celeste_fill
    ws['BH6'].border = border
    ws['BH6'] = 'RENAES' 
    
    ws['BI6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BI6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BI6'].fill = celeste_fill
    ws['BI6'].border = border
    ws['BI6'] = 'ID EST'  
    
    ws['BJ6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BJ6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BJ6'].fill = celeste_fill
    ws['BJ6'].border = border
    ws['BJ6'] = 'NOMBRE ESTABLECIMIENTO DE ATENCION'  
    
    ws['BK6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BK6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BK6'].fill = celeste_fill
    ws['BK6'].border = border
    ws['BK6'] = 'UBIG ESTA'  
    
    ws['BL6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BL6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BL6'].fill = celeste_fill
    ws['BL6'].border = border
    ws['BL6'] = 'COD DISA'  
    
    ws['BM6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BM6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BM6'].fill = celeste_fill
    ws['BM6'].border = border
    ws['BM6'] = 'DISA'  
    
    ws['BN6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BN6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BN6'].fill = celeste_fill
    ws['BN6'].border = border
    ws['BN6'] = 'COD RED'       
    
    ws['BO6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BO6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BO6'].fill = celeste_fill
    ws['BO6'].border = border
    ws['BO6'] = 'RED DE SALUD' 
    
    ws['BP6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BP6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BP6'].fill = celeste_fill
    ws['BP6'].border = border
    ws['BP6'] = 'COD MICRO'  
    
    ws['BQ6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BQ6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BQ6'].fill = celeste_fill
    ws['BQ6'].border = border
    ws['BQ6'] = 'NOMBRE DE MICRORED'  
    
    ws['BR6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BR6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BR6'].fill = celeste_fill
    ws['BR6'].border = border
    ws['BR6'] = 'COD UNICO' 
    
    ws['BS6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BS6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BS6'].fill = celeste_fill
    ws['BS6'].border = border
    ws['BS6'] = 'COD SEC'   
    
    ws['BT6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BT6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BT6'].fill = celeste_fill
    ws['BT6'].border = border
    ws['BT6'] = 'SECTOR'  
    
    ws['BU6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BU6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BU6'].fill = celeste_fill
    ws['BU6'].border = border
    ws['BU6'] = 'PROVINCIA'  
    
    ws['BV6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BV6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BV6'].fill = celeste_fill
    ws['BV6'].border = border
    ws['BV6'] = 'DISTRITO'  
    
    ws['BW6'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BW6'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BW6'].fill = celeste_fill
    ws['BW6'].border = border
    ws['BW6'] = 'CAT EST'  
    
        
        # Define styles
    promo_fill = PatternFill(patternType='solid', fgColor='FFD966')  # Yellow fill for promo
    font_normal = Font(name='Arial', size=8)
    font_bold_white = Font(name='Arial', size=7, bold=True, color='FFFFFF')
    font_red_bold = Font(name='Arial', size=7, bold=True, color='FF0000')
    font_green_bold = Font(name='Arial', size=7, bold=True, color='00FF00')
    font_red = Font(name='Arial', size=7, color='FF0000')
    font_green = Font(name='Arial', size=7, color='00B050')
    font_check = Font(name='Arial', size=10, color='00B050')
    font_x = Font(name='Arial', size=10, color='FF0000')
    plomo_claro_font = Font(name='Arial', size=7, color='FFEDEDED', bold=False)
    
    # Definir estilos
    header_font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    centered_alignment = Alignment(horizontal='center')
    border = Border(left=Side(style='thin', color='A9A9A9'),
            right=Side(style='thin', color='A9A9A9'),
            top=Side(style='thin', color='A9A9A9'),
            bottom=Side(style='thin', color='A9A9A9'))
    header_fill = PatternFill(patternType='solid', fgColor='00B0F0')
    
    # Definir los caracteres especiales de check y X
    check_mark = '✓'  # Unicode para check
    x_mark = '✗'  # Unicode para X
    sub_cumple = 'CUMPLE'
    sub_no_cumple = 'NO CUMPLE'
    
    # Escribir datos
    for row, record in enumerate(results, start=7):
        for col, value in enumerate(record, start=2):
            cell = ws.cell(row=row, column=col, value=value)

            # Alinear a la izquierda solo en las columnas 6,14,15,16
            if col in [10, 17, 18, 19, 20,24,26,35, 37, 41,45, 62, 67, 69, 74 ]:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='center')

            # Aplicar color en la columna 27
            if col == 57:
                if value == 0:
                    cell.value = sub_no_cumple  # Insertar check
                    cell.fill = PatternFill(patternType='solid', fgColor='FF0000')  # Fondo rojo
                    cell.font = Font(name='Arial', size=7,  bold = True, color='FFFFFF')  # Letra blanca
                elif value == 1:
                    cell.value = sub_cumple  
                    cell.fill = PatternFill(patternType='solid', fgColor='00FF00')  # Fondo verde
                    cell.font = Font(name='Arial', size=7,  bold = True, color='FFFFFF')  # Letra blanca
                else:
                    cell.font = Font(name='Arial', size=7)
            
            # Aplicar color de letra SUB INDICADORES
            elif col in [9, 32, 39,47]:
                if value == 0:
                    cell.value = sub_no_cumple  # Insertar check
                    cell.font = Font(name='Arial', size=7, color="FF0000")  # Letra roja
                elif value == 1:
                    cell.value = sub_cumple # Insertar check
                    cell.font = Font(name='Arial', size=7, color="00B050")  # Letra verde
                else:
                        cell.font = Font(name='Arial', size=7)
                        
            elif col in [28,39,47]:
                if value == '0':
                    cell.value = sub_no_cumple  # Insertar check
                    cell.font = Font(name='Arial', size=7, color="FF0000")  # Letra roja
                elif value == '1':
                    cell.value = sub_cumple # Insertar check
                    cell.font = Font(name='Arial', size=7, color="00B050")  # Letra verde
                else:
                    cell.font = Font(name='Arial', size=7)
                        
            # Aplicar color de letra SUB INDICADORES
            elif col in [9, 28, 33,39,47]:
                cell.font = Font(name='Arial', size=8, color="FF000033")

            # Fuente normal para otras columnas
            else:
                cell.font = Font(name='Arial', size=8)  # Fuente normal para otras columnas

            # Aplicar caracteres especiales check y X
            if col in [33,58]:
                if value == 1:
                    cell.value = check_mark  # Insertar check
                    cell.font = Font(name='Arial', size=10, color='00B050')  # Letra verde
                elif value == 0:
                    cell.value = x_mark  # Insertar X
                    cell.font = Font(name='Arial', size=10, color='FF0000')  # Letra roja
                else:
                    cell.font = Font(name='Arial', size=8)  # Fuente normal si no es 1 o 0
            
            if col in [33,58]:
                if value == '1':
                    cell.value = check_mark  # Insertar check
                    cell.font = Font(name='Arial', size=10, color='00B050')  # Letra verde
                elif value == '0':
                    cell.value = x_mark  # Insertar X
                    cell.font = Font(name='Arial', size=10, color='FF0000')  # Letra roja
                else:
                    cell.font = Font(name='Arial', size=8) 
            
                        
            cell.border = border