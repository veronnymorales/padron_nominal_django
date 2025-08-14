from django.http import JsonResponse, HttpResponse
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo
from django.db.models.functions import Substr
from django.db.models import IntegerField
from django.db.models.functions import Cast

from django.db import connection

# ===========================================================
# Funciones auxiliares de obtención de datos
# ===========================================================
def obtener_provincias(request):
    provincias = (
                MAESTRO_HIS_ESTABLECIMIENTO
                .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL')
                .annotate(ubigueo_filtrado=Substr('Ubigueo_Establecimiento', 1, 4))
                .values('Provincia','ubigueo_filtrado')
                .distinct()
                .order_by('Provincia')
    )    
    return list(provincias)

def obtener_distritos(provincia):
    distritos = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Provincia=provincia).values('Distrito').distinct().order_by('Distrito')
    return list(distritos)

def obtener_cards_observados():
    with connection.cursor() as cursor:       
        # Ejecutar el query con crosstab y parámetros dinámicos
        cursor.execute(
            '''
                SELECT
                    estado,
                    SUM(cant_registros) AS cantidad
                FROM dbo.OBSERVADOS
                GROUP BY estado
            ''',
        )
        return cursor.fetchall()

def obtener_grafico_barras():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT 
                    edad AS edad_ranking,
                    SUM(CAST(estado AS INT)) AS estado_ranking,
                    SUM(cant_registros) AS cantidad_ranking
                FROM 
                    dbo.OBSERVADOS
                WHERE 
                    estado IN ('2','3','4')
                GROUP BY 
                    edad
            ''',
        )
        return cursor.fetchall()


def obtener_ranking_observados():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT 
                    provincia AS provincia_ranking,
                    distrito AS distrito_ranking,
                    edad AS edad_ranking,
                    SUM(CAST(estado AS INT)) AS estado_ranking,
                    SUM(cant_registros) AS cantidad_ranking
                FROM 
                    dbo.OBSERVADOS
                WHERE 
                    estado IN ('2','3','4')
                GROUP BY 
                    provincia,
                    distrito,
                    edad
                ORDER BY 
                    cantidad_ranking DESC
            ''',
        )
        return cursor.fetchall()



def obtener_detalle_acta():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT  
                    provincia AS provincia_detalle, 
                    distrito AS distrito_detalle, 
                    municipio AS municipio_detalle, 
                    mes AS mes_detalle, 
                    fecha_inicial AS fecha_inicial_detalle, 
                    fecha_final AS fecha_final_detalle, 
                    fecha_envio AS fecha_envio_detalle, 
                    dni AS dni_detalle, 
                    primer_apellido AS primer_apellido_detalle, 
                    segundo_apellido AS segundo_apellido_detalle, 
                    nombres AS nombres_detalle
                FROM dbo.indicador_acta
                WHERE fecha_inicial IS NOT NULL
            ''',
        )
        return cursor.fetchall()


# ===========================================================
# Funciones para el seguimiento
# ===========================================================   
def obtener_seguimiento_observados():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT 
                [COD_PAD] AS cod_pad, 
                [TIPO_DOC] AS tipo_doc, 
                [CNV] AS cnv, 
                [CUI] AS cui, 
                [DNI] AS dni, 
                [NOMBRE_COMPLETO_NINO] AS nombre_completo_nino, 
                [SEXO_LETRA] AS sexo_letra, 
                [SEGURO] AS seguro, 
                [FECHA_NACIMIENTO_DATE] AS fecha_nacimiento_date, 
                [EDAD_LETRAS] AS edad_letras, 
                [DESCRIPCION] AS descripcion, 
                [PROVINCIA] AS provincia_seguimiento, 
                [DISTRITO] AS distrito_seguimiento, 
                [MENOR_VISITADO] AS menor_visitado, 
                [MENOR_ENCONTRADO] AS menor_encontrado, 
                [CODIGO_EESS] AS codigo_eess, 
                [NOMBRE_EESS] AS nombre_eess, 
                [FRECUENCIA_ATENCION] AS frecuencia_atencion, 
                [DNI_MADRE] AS dni_madre, 
                [NOMBRE_COMPLETO_MADRE] AS nombre_completo_madre, 
                [NUMERO_CELULAR] AS numero_celular, 
                [ESTADO_REGISTRO] AS estado_registro, 
                renaes, 
                [Nombre_Establecimiento] AS nombre_establecimiento, 
                [Ubigueo_Establecimiento] AS ubigueo_establecimiento,  
                [Codigo_Red] AS codigo_red, 
                [Red] AS red, 
                [Codigo_MicroRed] AS codigo_microred, 
                [MicroRed] AS microred
            FROM 
                dbo.SEGUIMIENTO_SITUACION_PADRON
            WHERE 
                [ESTADO_REGISTRO] IN ('2','3','4','5','6')
            ''',
        )
        return cursor.fetchall()