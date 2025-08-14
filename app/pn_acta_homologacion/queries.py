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

def obtener_tabla_acta(departamento, provincia, distrito):
    with connection.cursor() as cursor:       
        # Ejecutar el query con crosstab y parámetros dinámicos
        cursor.execute(
            '''
            SELECT departamento, provincia, distrito, municipio,
                    mes_enero, mes_febrero, mes_marzo, mes_abril, mes_mayo, mes_junio,
                    mes_julio, mes_agosto, mes_septiembre, mes_octubre, mes_noviembre, mes_diciembre  
            FROM 
                matriz_acta
            WHERE 
                (COALESCE(%s, '') = '' OR "departamento" = %s) AND
                (COALESCE(%s, '') = '' OR "provincia" = %s) AND
                (COALESCE(%s, '') = '' OR "distrito" = %s);
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_grafico_regional_acta():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT
                -- ENERO
                SUM(CASE WHEN mes = 1 THEN CAST(num AS INT) ELSE 0 END) AS NUM_1,
                SUM(CASE WHEN mes = 1 THEN CAST(den AS INT) ELSE 0 END) AS DEN_1,
                CASE 
                    WHEN SUM(CASE WHEN mes = 1 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 1 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 1 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_1,
                -- FEBRERO
                SUM(CASE WHEN mes = 2 THEN CAST(num AS INT) ELSE 0 END) AS NUM_2,
                SUM(CASE WHEN mes = 2 THEN CAST(den AS INT) ELSE 0 END) AS DEN_2,
                CASE 
                    WHEN SUM(CASE WHEN mes = 2 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 2 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 2 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_2,
                -- MARZO
                SUM(CASE WHEN mes = 3 THEN CAST(num AS INT) ELSE 0 END) AS NUM_3,
                SUM(CASE WHEN mes = 3 THEN CAST(den AS INT) ELSE 0 END) AS DEN_3,
                CASE 
                    WHEN SUM(CASE WHEN mes = 3 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 3 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 3 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_3,
                -- ABRIL
                SUM(CASE WHEN mes = 4 THEN CAST(num AS INT) ELSE 0 END) AS NUM_4,
                SUM(CASE WHEN mes = 4 THEN CAST(den AS INT) ELSE 0 END) AS DEN_4,
                CASE 
                    WHEN SUM(CASE WHEN mes = 4 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 4 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 4 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_4,
                -- MAYO
                SUM(CASE WHEN mes = 5 THEN CAST(num AS INT) ELSE 0 END) AS NUM_5,
                SUM(CASE WHEN mes = 5 THEN CAST(den AS INT) ELSE 0 END) AS DEN_5,
                CASE 
                    WHEN SUM(CASE WHEN mes = 5 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 5 THEN CAST(num AS INT) ELSE 0 END) * 1.0
                            / NULLIF(SUM(CASE WHEN mes = 5 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_5,
                -- JUNIO
                SUM(CASE WHEN mes = 6 THEN CAST(num AS INT) ELSE 0 END) AS NUM_6,
                SUM(CASE WHEN mes = 6 THEN CAST(den AS INT) ELSE 0 END) AS DEN_6,
                CASE 
                    WHEN SUM(CASE WHEN mes = 6 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 6 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 6 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_6,
                -- JULIO
                SUM(CASE WHEN mes = 7 THEN CAST(num AS INT) ELSE 0 END) AS NUM_7,
                SUM(CASE WHEN mes = 7 THEN CAST(den AS INT) ELSE 0 END) AS DEN_7,
                CASE 
                    WHEN SUM(CASE WHEN mes = 7 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 7 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 7 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_7,
                -- AGOSTO
                SUM(CASE WHEN mes = 8 THEN CAST(num AS INT) ELSE 0 END) AS NUM_8,
                SUM(CASE WHEN mes = 8 THEN CAST(den AS INT) ELSE 0 END) AS DEN_8,
                CASE 
                    WHEN SUM(CASE WHEN mes = 8 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 8 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 8 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_8,
                -- SETIEMBRE
                SUM(CASE WHEN mes = 9 THEN CAST(num AS INT) ELSE 0 END) AS NUM_9,
                SUM(CASE WHEN mes = 9 THEN CAST(den AS INT) ELSE 0 END) AS DEN_9,
                CASE 
                    WHEN SUM(CASE WHEN mes = 9 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 9 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 9 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2)
                END AS COB_9,
                -- OCTUBRE
                SUM(CASE WHEN mes = 10 THEN CAST(num AS INT) ELSE 0 END) AS NUM_10,
                SUM(CASE WHEN mes = 10 THEN CAST(den AS INT) ELSE 0 END) AS DEN_10,
                CASE 
                    WHEN SUM(CASE WHEN mes = 10 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 10 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 10 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_10,
                -- NOVIEMBRE
                SUM(CASE WHEN mes = 11 THEN CAST(num AS INT) ELSE 0 END) AS NUM_11,
                SUM(CASE WHEN mes = 11 THEN CAST(den AS INT) ELSE 0 END) AS DEN_11,
                CASE 
                    WHEN SUM(CASE WHEN mes = 11 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 11 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 11 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_11,
                -- DICIEMBRE
                SUM(CASE WHEN mes = 12 THEN CAST(num AS INT) ELSE 0 END) AS NUM_12,
                SUM(CASE WHEN mes = 12 THEN CAST(den AS INT) ELSE 0 END) AS DEN_12,
                CASE 
                    WHEN SUM(CASE WHEN mes = 12 THEN CAST(den AS INT) ELSE 0 END) = 0 
                    THEN 0 
                    ELSE ROUND(
                        (
                            SUM(CASE WHEN mes = 12 THEN CAST(num AS INT) ELSE 0 END) * 1.0 
                            / NULLIF(SUM(CASE WHEN mes = 12 THEN CAST(den AS INT) ELSE 0 END), 0)
                        ) * 100
                    , 2) 
                END AS COB_12
                FROM 
                indicador_acta
            ''',
        )
        return cursor.fetchall()

def obtener_ranking_acta():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT 
                provincia AS provincia_ranking,
                distrito AS distrito_ranking,
                SUM(den) AS den_ranking,
                SUM(num) AS num_ranking,
                COALESCE(
                    CAST(ROUND(SUM(num) * 100.0 / NULLIF(SUM(den), 0), 2) AS DECIMAL(10,2)),
                    0
                ) AS avance_ranking,
                CASE 
                    WHEN COALESCE(ROUND(SUM(num) * 100.0 / NULLIF(SUM(den), 0), 2), 0) BETWEEN 0 AND 8 THEN 'RIESGO'
                    WHEN COALESCE(ROUND(SUM(num) * 100.0 / NULLIF(SUM(den), 0), 2), 0) BETWEEN 8 AND 24 THEN 'PROCESO'
                    WHEN COALESCE(ROUND(SUM(num) * 100.0 / NULLIF(SUM(den), 0), 2), 0) BETWEEN 25 AND 100 THEN 'CUMPLE'
                    ELSE 'RIESGO'
                END AS estado_ranking
            FROM 
                indicador_acta
            WHERE
                distrito IS NOT NULL
            GROUP BY 
                provincia,
                distrito
            ''',
        )
        return cursor.fetchall()

def obtener_detalle_acta():
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT  
                    provincia as provincia_detalle, 
                    distrito as distrito_detalle, 
                    municipio as municipio_detalle, 
                    mes as mes_detalle, 
                    fecha_inicial as fecha_inicial_detalle, 
                    fecha_final as fecha_final_detalle, 
                    fecha_envio as fecha_envio_detalle, 
                    dni as dni_detalle, 
                    primer_apellido as primer_apellido_detalle, 
                    segundo_apellido as segundo_apellido_detalle, 
                    nombres as nombres_detalle
                FROM 
                    indicador_acta
                WHERE fecha_inicial is not null
            ''',
        )
        return cursor.fetchall()


# ===========================================================
# Funciones para el seguimiento
# ===========================================================
def obtener_seguimiento_situacion_padron_old(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT 
            * 
            FROM 
            TRAMA_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR "DEPARTAMENTO" = %s) AND
                (COALESCE(%s, '') = '' OR "PROVINCIA" = %s) AND
                (COALESCE(%s, '') = '' OR "DISTRITO" = %s);
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()
    
def obtener_seguimiento_situacion_padron(departamento, provincia, edad, cumple):
    """
    Función para obtener datos del seguimiento del padrón nominal filtrados por ubicación, edad y cumplimiento.
    
    Parámetros:
        - departamento (str): Departamento a filtrar.
        - provincia (str): Provincia a filtrar.
        - distrito (str): Distrito a filtrar.
        - edad (str): Filtro por categoría de edad ('N28_dias', 'N0a5meses', etc.).
        - cumple (str): Filtro por cumplimiento ('1', '0', '').
    
    Retorna:
        - Listado de tuplas con los resultados de la consulta.
    """
    # Mapeo de las categorías de edad a condiciones SQL
    edad_conditions = {
        "": "1=1",  # Todos los registros si no se selecciona ninguna edad
        "N28_dias": "(edad_anio2 = 0 AND edad_mes2 = 0 AND edad_dias2 < 28)",  
        "N0a5meses": "(edad_anio2 = 0 AND edad_mes2 BETWEEN 0 AND 5)",
        "N6a11meses": "(edad_anio2 = 0 AND edad_mes2 BETWEEN 6 AND 11)",
        "cero_anios": "(edad_anio2 = 0 OR (edad_anio2 = 1 AND edad_mes2 = 0))",
        "un_anios": "(edad_anio2 = 1)",
        "dos_anios": "(edad_anio2 = 2)",
        "tres_anios": "(edad_anio2 = 3)",
        "cuatro_anio": "(edad_anio2 = 4)",
        "cinco_anios": "(edad_anio2 = 5)"
    }

    # Obtener la condición SQL para la edad seleccionada
    edad_condition = edad_conditions.get(edad, "1=1")  # Por defecto, incluir todos los registros

    with connection.cursor() as cursor:
        # Consulta SQL con parámetros dinámicos
        query = f'''
            SELECT 
                "NRO", "COD_PAD", "TIPO_DOC", "CNV", "CUI", "DNI", "ESTADO_DE_TRAMITE_DNI", "CUMPLE_CUI_DNI", 
                "NOMBRE_COMPLETO_NINO", "SEXO_LETRA", "SEGURO", "FECHA_NACIMIENTO_DATE", 
                "edad_anio2", "edad_mes2", "edad_dias2", "EDAD_LETRAS", "EJE_VIAL", "DESCRIPCION", 
                "REFERENCIA_DIRECCION", "COD_UBIGEO", "DEPARTAMENTO", "PROVINCIA", "DISTRITO", "CODIGO_CENTRO_POBLADO", 
                "CENTRO_POBLADO", "AREA_CENTRO_POBLADO", "DIRECCION_COMPLETA", "MENOR_VISITADO", "MENOR_ENCONTRADO", 
                "FECHA_VISITA", "CUMPLE_FECHA_VISITA", "VISITADO_NO_ENCONTRADO", "CODIGO_NACIMIENTO", "NOMBRE_NACIMIENTO", 
                "CODIGO_EESS", "NOMBRE_EESS", "FRECUENCIA_ATENCION", "CUMPLE_FRECUENCIA_ATENCION", "CODIGO_EESS_ADSCRIPCION", 
                "NOMBRE_EESS_ADSCRIPCION", "PROGRAMAS_SOCIALES", "TIPO_DE_DOCUMENTO_DE_LA_MADRE", "DNI_MADRE", "NOMBRE_COMPLETO_MADRE", 
                "NUMERO_CELULAR", "CUMPLE_CELULAR", "CORREO_ELECTRONICO", "ESTADO_REGISTRO", "FECHA_CREACION", "USUARIO_CREA", 
                "FECHA_MODIFICACION", "USUARIO_MODIFICA", "ENTIDAD", "TIPO_REGISTRO", "FECHA_CORTE", "NUM", "DEN", 
                ultimo_periodo, renaes, "Id_Establecimiento", "Nombre_Establecimiento", "Ubigueo_Establecimiento", "Codigo_Disa", 
                "Disa", "Codigo_Red", "Red", "Codigo_MicroRed", "MicroRed", "Codigo_Unico", "Codigo_Sector", "Descripcion_Sector", 
                "PROV", "DIST", "Categoria_Establecimiento"
	        FROM SEGUIMIENTO_SITUACION_PADRON
            WHERE
                -- Filtrar por ubicación geográfica
                ("DEPARTAMENTO" = %s OR %s = '')
                AND (LEFT("COD_UBIGEO", 4)= %s OR %s = '')

                -- Filtrar por edad
                AND {edad_condition}

                -- Filtrar por cumplimiento
                AND (
                    %s = ''
                    OR (%s = '1' AND "NUM" = 1)
                    OR (%s = '0' AND "NUM" = 0)
                )
            ORDER BY "edad_anio2", "edad_mes2", "edad_dias2"
        '''
        
        # Ejecutar la consulta con los parámetros
        cursor.execute(
            query,
            [
                departamento, departamento,
                provincia, provincia,
                cumple, cumple, cumple
            ]
        )
        
        # Obtener los resultados
        return cursor.fetchall()
    
def obtener_seguimiento_situacion_padron_distrito(departamento, provincia, distrito, edad, cumple):
    # Mapeo de las categorías de edad a condiciones SQL
    edad_conditions = {
        "": "1=1",  # Todos los registros si no se selecciona ninguna edad
        "N28_dias": "(edad_anio2 = 0 AND edad_mes2 = 0 AND edad_dias2 < 28)",  
        "N0a5meses": "(edad_anio2 = 0 AND edad_mes2 BETWEEN 0 AND 5)",
        "N6a11meses": "(edad_anio2 = 0 AND edad_mes2 BETWEEN 6 AND 11)",
        "cero_anios": "(edad_anio2 = 0 OR (edad_anio2 = 1 AND edad_mes2 = 0))",
        "un_anios": "(edad_anio2 = 1)",
        "dos_anios": "(edad_anio2 = 2)",
        "tres_anios": "(edad_anio2 = 3)",
        "cuatro_anio": "(edad_anio2 = 4)",
        "cinco_anios": "(edad_anio2 = 5)"
    }

    # Obtener la condición SQL para la edad seleccionada
    edad_condition = edad_conditions.get(edad, "1=1")  # Por defecto, incluir todos los registros

    with connection.cursor() as cursor:
        # Consulta SQL con parámetros dinámicos
        query = f'''
            SELECT 
                "NRO", "COD_PAD", "TIPO_DOC", "CNV", "CUI", "DNI", "ESTADO_DE_TRAMITE_DNI", "CUMPLE_CUI_DNI", 
                "NOMBRE_COMPLETO_NINO", "SEXO_LETRA", "SEGURO", "FECHA_NACIMIENTO_DATE", 
                "edad_anio2", "edad_mes2", "edad_dias2", "EDAD_LETRAS", "EJE_VIAL", "DESCRIPCION", 
                "REFERENCIA_DIRECCION", "COD_UBIGEO", "DEPARTAMENTO", "PROVINCIA", "DISTRITO", "CODIGO_CENTRO_POBLADO", 
                "CENTRO_POBLADO", "AREA_CENTRO_POBLADO", "DIRECCION_COMPLETA", "MENOR_VISITADO", "MENOR_ENCONTRADO", 
                "FECHA_VISITA", "CUMPLE_FECHA_VISITA", "VISITADO_NO_ENCONTRADO", "CODIGO_NACIMIENTO", "NOMBRE_NACIMIENTO", 
                "CODIGO_EESS", "NOMBRE_EESS", "FRECUENCIA_ATENCION", "CUMPLE_FRECUENCIA_ATENCION", "CODIGO_EESS_ADSCRIPCION", 
                "NOMBRE_EESS_ADSCRIPCION", "PROGRAMAS_SOCIALES", "TIPO_DE_DOCUMENTO_DE_LA_MADRE", "DNI_MADRE", "NOMBRE_COMPLETO_MADRE", 
                "NUMERO_CELULAR", "CUMPLE_CELULAR", "CORREO_ELECTRONICO", "ESTADO_REGISTRO", "FECHA_CREACION", "USUARIO_CREA", 
                "FECHA_MODIFICACION", "USUARIO_MODIFICA", "ENTIDAD", "TIPO_REGISTRO", "FECHA_CORTE", "NUM", "DEN", 
                ultimo_periodo, renaes, "Id_Establecimiento", "Nombre_Establecimiento", "Ubigueo_Establecimiento", "Codigo_Disa", 
                "Disa", "Codigo_Red", "Red", "Codigo_MicroRed", "MicroRed", "Codigo_Unico", "Codigo_Sector", "Descripcion_Sector", 
                "PROV", "DIST", "Categoria_Establecimiento"
	        FROM public."SEGUIMIENTO_SITUACION_PADRON"
            WHERE
                -- Filtrar por ubicación geográfica
                ("DEPARTAMENTO" = %s)
                AND (LEFT("COD_UBIGEO", 4) = %s)
                AND (LEFT("COD_UBIGEO", 6) = %s)
                -- Filtrar por edad
                AND {edad_condition}

                -- Filtrar por cumplimiento
                AND (
                    %s = ''
                    OR (%s = '1' AND "NUM" = 1)
                    OR (%s = '0' AND "NUM" = 0)
                )
            ORDER BY "edad_anio2", "edad_mes2", "edad_dias2"
        '''
        
        # Ejecutar la consulta con los parámetros
        cursor.execute(
            query,
            [
                departamento, 
                provincia, 
                distrito,
                cumple, cumple, cumple
            ]
        )
        
        # Obtener los resultados
        return cursor.fetchall()