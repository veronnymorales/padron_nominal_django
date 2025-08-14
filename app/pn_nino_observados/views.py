import logging
from django.shortcuts import render
from django.http import JsonResponse
from .queries import (obtener_distritos,obtener_cards_observados, obtener_grafico_barras, obtener_ranking_observados, obtener_seguimiento_observados )

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


def index_observados_padron(request):
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

            # CARDS
            resultados_obtener_cards_observados = obtener_cards_observados()           
            
            # OBTENER GRAFICO EDAD
            resultados_obtener_grafico_barras = obtener_grafico_barras()
            
            # OBTENER RANKIG
            resultados_ranking_observados = obtener_ranking_observados()
            #print("resultados_detalle_acta:", resultados_detalle_acta)
            
            # OBTENER RANKIG
            resultados_seguimiento_observados = obtener_seguimiento_observados()
            #print("resultados_detalle_acta:", resultados_detalle_acta)
            
            # Estructura de datos inicial
            data = {
                # DATA MATRIZ
                'estado': [],
                'cantidad': [],
                
                ## DATA GRAFICO BARRAS
                'edad': [],
                'estado_barras': [],
                'cantidad_barras': [],
                
                # DATA RANKING ACTA
                'provincia_ranking': [],
                'distrito_ranking': [],
                'edad_ranking': [],
                'estado_ranking': [],
                'cantidad_ranking': [],
                
                # DATA SEGUIMIENTO ACTA
                'provincia_ranking': [],
                'distrito_ranking': [],
                'edad_ranking': [],
                'estado_ranking': [],
                'cantidad_ranking': [],
                
                # DATA SEGUIMIENTO
                'cod_pad': [],
                'tipo_doc': [],
                'cnv': [],
                'cui': [],
                'dni': [],
                'nombre_completo_nino': [],
                'sexo_letra': [],
                'seguro': [],
                'fecha_nacimiento_date': [],
                'edad_letras': [],
                'descripcion': [],
                'provincia_seguimiento': [],
                'distrito_seguimiento': [],
                'menor_visitado': [],
                'menor_encontrado': [],
                'codigo_eess': [],
                'nombre_eess': [],
                'frecuencia_atencion': [],
                'dni_madre': [],
                'nombre_completo_madre': [],
                'numero_celular': [],
                'estado_registro': [],
                'renaes': [],
                'nombre_establecimiento': [],
                'ubigueo_establecimiento': [],
                'codigo_red': [],
                'red': [],
                'codigo_microred': [],
                'microred': [],
            }

            # ----------------------------------------------------------------------------
            # 1. MATRIZ DETALLADA
            # ----------------------------------------------------------------------------
            for row in resultados_obtener_cards_observados:
                # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 2:
                    data['estado'].append(row[0])
                    data['cantidad'].append(row[1])

                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            # --------------------------------------------------------
            # 2. DATA GRAFICO BARRAS
            # ---------------------------------------------------------
            for row in resultados_obtener_grafico_barras:
                    # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 3:
                    data['edad'].append(row[0])
                    data['estado_barras'].append(row[1])
                    data['cantidad_barras'].append(row[2])
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            # --------------------------------------------------------
            # 3. RANKING
            # ---------------------------------------------------------
            for row in resultados_ranking_observados:
                    # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 5:
                    data['provincia_ranking'].append(row[0])
                    data['distrito_ranking'].append(row[1])
                    # Cambia null (None) a 0
                    data['edad_ranking'].append(str(row[2]) if row[2] is not None else '0')
                    data['estado_ranking'].append(str(row[3]) if row[3] is not None else '0')
                    data['cantidad_ranking'].append(str(row[4]) if row[4] is not None else '0')
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
            # --------------------------------------------------------
            # 4. SEGUIMIENTO
            # ---------------------------------------------------------
            for row in resultados_seguimiento_observados:
                    # En lugar de lanzar error, checamos si la tupla es la longitud esperada:
                if len(row) == 29:
                    data['cod_pad'].append(row[0])
                    data['tipo_doc'].append(row[1])
                    # Cambia null (None) a 0
                    data['cnv'].append(str(row[2]) if row[2] is not None else '0')
                    data['cui'].append(str(row[3]) if row[3] is not None else '0')
                    data['dni'].append(str(row[4]) if row[4] is not None else '0')
                    data['nombre_completo_nino'].append(str(row[5]) if row[5] is not None else '0')
                    data['sexo_letra'].append(str(row[6]) if row[6] is not None else '0')
                    data['seguro'].append(str(row[7]) if row[7] is not None else '0')
                    data['fecha_nacimiento_date'].append(str(row[8]) if row[8] is not None else '0')
                    data['edad_letras'].append(str(row[9]) if row[9] is not None else '0')
                    data['descripcion'].append(str(row[10]) if row[10] is not None else '0')
                    data['provincia_seguimiento'].append(str(row[11]) if row[11] is not None else '0')
                    data['distrito_seguimiento'].append(str(row[12]) if row[12] is not None else '0')
                    data['menor_visitado'].append(str(row[13]) if row[13] is not None else '0')
                    data['menor_encontrado'].append(str(row[14]) if row[14] is not None else '0')
                    data['codigo_eess'].append(str(row[15]) if row[15] is not None else '0')
                    data['nombre_eess'].append(str(row[16]) if row[16] is not None else '0')
                    data['frecuencia_atencion'].append(str(row[17]) if row[17] is not None else '0')
                    data['dni_madre'].append(str(row[18]) if row[18] is not None else '0')
                    data['nombre_completo_madre'].append(str(row[19]) if row[19] is not None else '0')
                    data['numero_celular'].append(str(row[20]) if row[20] is not None else '0')
                    data['estado_registro'].append(str(row[21]) if row[21] is not None else '0')
                    data['renaes'].append(str(row[22]) if row[22] is not None else '0')
                    data['nombre_establecimiento'].append(str(row[23]) if row[23] is not None else '0')
                    data['ubigueo_establecimiento'].append(str(row[24]) if row[24] is not None else '0')
                    data['codigo_red'].append(str(row[25]) if row[25] is not None else '0')
                    data['red'].append(str(row[26]) if row[26] is not None else '0')
                    data['codigo_microred'].append(str(row[27]) if row[27] is not None else '0')
                    data['microred'].append(str(row[28]) if row[28] is not None else '0')
        
                else:
                    logger.warning(f"Fila con estructura inválida: {row}")
            
                
            return JsonResponse(data)
        
        except:
            # Si ocurre alguna excepción global, la silenciamos (no mostramos nada)
            return JsonResponse({}, status=200)
        

    # -- Si no es AJAX, render normal de la plantilla --
    return render(request, 'pn_nino_observados/index_observados_padron.html', {
        'actualizacion': actualizacion,
        'provincias': provincias,
    })
