import logging
from django.shortcuts import render
from django.http import JsonResponse
from .queries import (obtener_tabla_acta,obtener_distritos, obtener_grafico_regional_acta,obtener_ranking_acta, obtener_detalle_acta)

from base.models import MAESTRO_HIS_ESTABLECIMIENTO, Actualizacion

from django.db.models.functions import Substr

from django.db import connection
from django.http import JsonResponse
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo, Actualizacion
from django.db.models.functions import Substr

# REPORTE EXCEL
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


def index_acta_padron(request):
    actualizacion = Actualizacion.objects.all()

    # Provincias para el primer <select>
    provincias = (MAESTRO_HIS_ESTABLECIMIENTO.objects
                    .values_list('Provincia', flat=True)
                    .distinct()
                    .order_by('Provincia'))
    
    # Obtener parámetros
    departamento_selecionado = request.GET.get('departamento')
    provincia_seleccionada = request.GET.get('provincia')
    distrito_seleccionado = request.GET.get('distrito')

    # -- Manejo de distritos via HTMX (retorna template parcial) --
    if 'get_distritos' in request.GET:
        if provincia_seleccionada:
            distritos = obtener_distritos(provincia_seleccionada)
        else:
            distritos = []
        
        return render(request, "pn_situacion_actual/partials/_distritos_options.html", {
            "distritos": distritos
        })

    # -- Si es una solicitud AJAX, devolvemos JsonResponse con la data --
    if request.headers.get('x-requested-with') == 'XMLHttpRequest':
        # Intentamos procesar la data sin mostrar errores
        try:
            # Llamada redundante a get_distritos (por si la plantilla antigua lo usara)
            if 'get_distritos' in request.GET:
                distritos = obtener_distritos(provincia_seleccionada)
                return JsonResponse(distritos, safe=False)
            
            # AVANCE MATRIZ ACTA
            resultados_tabla_acta = obtener_tabla_acta(
                departamento_selecionado, provincia_seleccionada, distrito_seleccionado
            )
            
            # AVANCE GRAFICO REGIONAL MENSUALIZADO
            resultados_grafico_regional = obtener_grafico_regional_acta()           
            
            resultados_ranking_acta = obtener_ranking_acta()
            
            # OBTENER DETALLE
            resultados_detalle_acta = obtener_detalle_acta()
            #print("resultados_detalle_acta:", resultados_detalle_acta)
            
            
            # Estructura de datos inicial
            data = {
                # DATA MATRIZ
                'departamento': [],
                'provincia_matriz': [],
                'distrito_matriz': [],
                'municipio': [],
                'mes_enero': [],
                'mes_febrero': [],
                'mes_marzo': [],
                'mes_abril': [],
                'mes_mayo': [],
                'mes_junio': [],
                'mes_julio': [],
                'mes_agosto': [],
                'mes_septiembre': [],
                'mes_octubre': [],
                'mes_noviembre': [],
                'mes_diciembre': [],
                
                # DATA GRAFICO REGIONAL
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
                
                # DATA RANKING ACTA
                'provincia_ranking': [],
                'distrito_ranking': [],
                'den_ranking': [],
                'num_ranking': [],
                'avance_ranking': [],
                'estado_ranking': [],
                
                # DATA DETALLE ACTA
                'provincia_detalle': [],
                'distrito_detalle': [],
                'municipio_detalle': [],
                'mes_detalle': [],
                'fecha_inicial_detalle': [],
                'fecha_final_detalle': [],
                'fecha_envio_detalle': [], 
                'dni_detalle': [],
                'primer_apellido_detalle': [], 
                'segundo_apellido_detalle': [], 
                'nombres_detalle': [],
            }

            # ----------------------------------------------------------------------------
            # MATRIZ DETALLADA
            # ----------------------------------------------------------------------------
            for row in resultados_tabla_acta:
                # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 16:
                    data['departamento'].append(row[0])
                    data['provincia_matriz'].append(row[1])
                    data['distrito_matriz'].append(row[2])
                    data['municipio'].append(row[3])
                    
                    # Convertir fechas a string y manejar None
                    data['mes_enero'].append(str(row[4]) if row[4] else None)
                    data['mes_febrero'].append(str(row[5]) if row[5] else None)
                    data['mes_marzo'].append(str(row[6]) if row[6] else None)
                    data['mes_abril'].append(str(row[7]) if row[7] else None)
                    data['mes_mayo'].append(str(row[8]) if row[8] else None)
                    data['mes_junio'].append(str(row[9]) if row[9] else None)
                    data['mes_julio'].append(str(row[10]) if row[10] else None)
                    data['mes_agosto'].append(str(row[11]) if row[11] else None)
                    data['mes_septiembre'].append(str(row[12]) if row[12] else None)
                    data['mes_octubre'].append(str(row[13]) if row[13] else None)
                    data['mes_noviembre'].append(str(row[14]) if row[14] else None)
                    data['mes_diciembre'].append(str(row[15]) if row[15] else None)
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            # --------------------------------------------------------
            # 1. Procesar DATOS MENSUALIZADOS (filtrados por año)
            # ---------------------------------------------------------
            for index, row in enumerate(resultados_grafico_regional):
                try:
                    if len(row) != 36:
                        raise ValueError(f"La fila {index} no tiene 36 columnas: {row}")

                    for i in range(12):
                        data[f'num_{i+1}'].append(row[i*3] if row[i*3] is not None else 0)
                        data[f'den_{i+1}'].append(row[i*3+1] if row[i*3+1] is not None else 0)
                        # Convertir Decimal a float
                        cob_value = row[i*3+2]
                        if cob_value is None:
                            data[f'cob_{i+1}'].append(0.0)
                        elif isinstance(cob_value, float):
                            data[f'cob_{i+1}'].append(cob_value)
                        else:
                            data[f'cob_{i+1}'].append(float(cob_value))
                except Exception as e:
                    logger.error(f"Error procesando la fila {index}: {str(e)}")
            
            # ----------------------------------------------------------------------------
            # 2. Procesar RANKING ACTA
            # ----------------------------------------------------------------------------
            for row in resultados_ranking_acta:
                # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 6:
                    data['provincia_ranking'].append(row[0])
                    data['distrito_ranking'].append(row[1])
                    # Cambia null (None) a 0
                    data['den_ranking'].append(str(row[2]) if row[2] is not None else '0')
                    data['num_ranking'].append(str(row[3]) if row[3] is not None else '0')
                    data['avance_ranking'].append(str(row[4]) if row[4] is not None else '0')
                    data['estado_ranking'].append(str(row[5]) if row[5] is not None else 'RIESGO')
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            # ----------------------------------------------------------------------------
            # 3. DETALLE DE ACTA
            # ----------------------------------------------------------------------------
            for row in resultados_detalle_acta:
                # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 11:
                    data['provincia_detalle'].append(row[0])
                    data['distrito_detalle'].append(row[1])
                    data['municipio_detalle'].append(row[2])
                    data['mes_detalle'].append(row[3])
                    # Cambia null (None) a 0
                    data['fecha_inicial_detalle'].append(str(row[4]) if row[4] is not None else '')
                    data['fecha_final_detalle'].append(str(row[5]) if row[5] is not None else '')
                    data['fecha_envio_detalle'].append(str(row[6]) if row[6] is not None else '')
                    data['dni_detalle'].append(str(row[7]) if row[7] is not None else '0')
                    data['primer_apellido_detalle'].append(str(row[8]) if row[8] is not None else '0')
                    data['segundo_apellido_detalle'].append(str(row[9]) if row[9] is not None else '0')
                    data['nombres_detalle'].append(str(row[10]) if row[10] is not None else '0')
                
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            return JsonResponse(data)
        
        except:
            # Si ocurre alguna excepción global, la silenciamos (no mostramos nada)
            return JsonResponse({}, status=200)
        

    # -- Si no es AJAX, render normal de la plantilla --
    return render(request, 'pn_acta_homologacion/index_pn_acta_homologacion.html', {
        'actualizacion': actualizacion,
        'provincias': provincias,
    })
