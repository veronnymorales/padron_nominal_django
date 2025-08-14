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

def obtener_avance_situacion_padron(departamento,provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT 
                SUM(N28_dias) AS N28_dias,
                SUM(N0a5meses) AS N0a5meses,
                SUM(N6a11meses) AS N6a11meses,
                SUM(cero_anios) AS cero_anios,
                SUM(un_anios) AS un_anios,
                SUM(dos_anios) AS dos_anios,
                SUM(tres_anios) AS tres_anios,
                SUM(cuatro_anio) AS cuatro_anio,
                SUM(cinco_anios) AS cinco_anios,
                SUM(total_den) AS total_den
            FROM SITUACION_PADRON 
            WHERE
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [   
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_dni(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_dni) AS total_cumple_dni,
                SUM(brecha_dni) AS brecha_dni,
                SUM(cob_dni) AS cob_dni
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_cnv(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_cnv) AS total_cumple_cnv,
                SUM(brecha_cumple_cnv) AS brecha_cumple_cnv,
                SUM(cob_cnv) AS cob_cnv
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_eje_vial(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_eje_vial) AS total_cumple_eje_vial,
                SUM(brecha_eje_vial) AS brecha_eje_vial,
                SUM(cob_eje_vial) AS cob_eje_vial
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_direccion(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_direccion) AS total_cumple_direccion,
                SUM(brecha_direccion) AS brecha_direccion,
                SUM(cob_direccion) AS cob_direccion
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_referencia(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT  
                    SUM(total_cumple_referencia) AS total_cumple_referencia,
                    SUM(brecha_referencia) AS brecha_referencia,
                    SUM(cob_referencia) AS cob_referencia
                FROM dbo.SITUACION_PADRON
                WHERE 
                    (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                    (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                    (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_visitado(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_visitado) AS total_cumple_visitado,
                SUM(brecha_visitado) AS brecha_visitado,
                SUM(cob_visitado) AS cob_visitado
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_encontrado(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_encontrado) AS total_cumple_encontrado,
                SUM(brecha_encontrado) AS brecha_encontrado,
                SUM(cob_encontrado) AS cob_encontrado
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_celular(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_celular) AS total_cumple_celular,
                SUM(brecha_celular) AS brecha_celular,
                SUM(cob_celular) AS cob_celular
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_sexo(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_sexo_masculino) AS total_cumple_sexo_masculino,
                SUM(total_cumple_sexo_femenino) AS total_cumple_sexo_femenino,
                SUM(cob_sexo) AS cob_sexo
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_seguro(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_cumple_seguro) AS total_cumple_seguro,
                SUM(brecha_seguro) AS brecha_seguro,
                SUM(cob_seguro) AS cob_seguro
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_eess(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_eess) AS total_eess,
                SUM(brecha_eess) AS brecha_eess,
                SUM(cob_eess) AS cob_eess
            FROM SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_frecuencia(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_frecuencia) AS total_frecuencia,
                SUM(brecha_frecuencia) AS brecha_frecuencia,
                SUM(cob_frecuencia) AS cob_frecuencia
            FROM dbo.SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_direccion_completa(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
                SELECT  
                    SUM(total_direccion_completa) AS total_direccion_completa,
                    SUM(brecha_direccion_completa) AS brecha_direccion_completa,
                    SUM(cob_direccion_completa) AS cob_direccion_completa
                FROM dbo.SITUACION_PADRON
                WHERE 
                    (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                    (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                    (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()

def obtener_cumple_situacion_visitado_no_encontrado(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT  
                SUM(total_visitado_no_encontrado) AS total_visitado_no_encontrado,
                SUM(brecha_visitado_no_encontrado) AS brecha_visitado_no_encontrado,
                SUM(cob_visitado_no_encontrado) AS cob_visitado_no_encontrado
            FROM dbo.SITUACION_PADRON
            WHERE 
                (COALESCE(%s, '') = '' OR DEPARTAMENTO = %s) AND
                (COALESCE(%s, '') = '' OR PROVINCIA = %s) AND
                (COALESCE(%s, '') = '' OR DISTRITO = %s)
            ''',
            [
                departamento, departamento,
                provincia, provincia,
                distrito, distrito
            ]
        )
        return cursor.fetchall()


# ===========================================================
# Funciones para el seguimiento
# ===========================================================
def obtener_seguimiento_situacion_padron_old(departamento, provincia, distrito):
    with connection.cursor() as cursor:
        cursor.execute(
            '''
            SELECT * FROM public."TRAMA_PADRON"
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
                    NRO, COD_PAD, TIPO_DOC, CNV, CUI, DNI, ESTADO_DE_TRAMITE_DNI, CUMPLE_CUI_DNI, 
                    NOMBRE_COMPLETO_NINO, SEXO_LETRA, SEGURO, FECHA_NACIMIENTO_DATE, 
                    edad_anio2, edad_mes2, edad_dias2, EDAD_LETRAS, EJE_VIAL, DESCRIPCION, 
                    REFERENCIA_DIRECCION, COD_UBIGEO, DEPARTAMENTO, PROVINCIA, DISTRITO, CODIGO_CENTRO_POBLADO, 
                    CENTRO_POBLADO, AREA_CENTRO_POBLADO, DIRECCION_COMPLETA, MENOR_VISITADO, MENOR_ENCONTRADO, 
                    FECHA_VISITA, CUMPLE_FECHA_VISITA, VISITADO_NO_ENCONTRADO, CODIGO_NACIMIENTO, NOMBRE_NACIMIENTO, 
                    CODIGO_EESS, NOMBRE_EESS, FRECUENCIA_ATENCION, CUMPLE_FRECUENCIA_ATENCION, CODIGO_EESS_ADSCRIPCION, 
                    NOMBRE_EESS_ADSCRIPCION, PROGRAMAS_SOCIALES, TIPO_DE_DOCUMENTO_DE_LA_MADRE, DNI_MADRE, NOMBRE_COMPLETO_MADRE, 
                    NUMERO_CELULAR, CUMPLE_CELULAR, CORREO_ELECTRONICO, ESTADO_REGISTRO, FECHA_CREACION, USUARIO_CREA, 
                    FECHA_MODIFICACION, USUARIO_MODIFICA, ENTIDAD, TIPO_REGISTRO, FECHA_CORTE, NUM, DEN, 
                    ultimo_periodo, renaes, Id_Establecimiento, Nombre_Establecimiento, Ubigueo_Establecimiento, Codigo_Disa, 
                    Disa, Codigo_Red, Red, Codigo_MicroRed, MicroRed, Codigo_Unico, Codigo_Sector, Descripcion_Sector, 
                    PROV, DIST, Categoria_Establecimiento
                FROM dbo.SEGUIMIENTO_SITUACION_PADRON
                WHERE
                    -- Filtrar por ubicación geográfica
                    (DEPARTAMENTO = %s OR %s = '')
                    AND (LEFT(COD_UBIGEO, 4) = %s OR %s = '')
                    -- Filtrar por edad
                    AND {edad_condition}
                    -- Filtrar por cumplimiento
                    AND (
                        %s = ''
                        OR (%s = '1' AND NUM = 1)
                        OR (%s = '0' AND NUM = 0)
                    )
                ORDER BY edad_anio2, edad_mes2, edad_dias2
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
                        NRO, COD_PAD, TIPO_DOC, CNV, CUI, DNI, ESTADO_DE_TRAMITE_DNI, CUMPLE_CUI_DNI, 
                        NOMBRE_COMPLETO_NINO, SEXO_LETRA, SEGURO, FECHA_NACIMIENTO_DATE, 
                        edad_anio2, edad_mes2, edad_dias2, EDAD_LETRAS, EJE_VIAL, DESCRIPCION, 
                        REFERENCIA_DIRECCION, COD_UBIGEO, DEPARTAMENTO, PROVINCIA, DISTRITO, CODIGO_CENTRO_POBLADO, 
                        CENTRO_POBLADO, AREA_CENTRO_POBLADO, DIRECCION_COMPLETA, MENOR_VISITADO, MENOR_ENCONTRADO, 
                        FECHA_VISITA, CUMPLE_FECHA_VISITA, VISITADO_NO_ENCONTRADO, CODIGO_NACIMIENTO, NOMBRE_NACIMIENTO, 
                        CODIGO_EESS, NOMBRE_EESS, FRECUENCIA_ATENCION, CUMPLE_FRECUENCIA_ATENCION, CODIGO_EESS_ADSCRIPCION, 
                        NOMBRE_EESS_ADSCRIPCION, PROGRAMAS_SOCIALES, TIPO_DE_DOCUMENTO_DE_LA_MADRE, DNI_MADRE, NOMBRE_COMPLETO_MADRE, 
                        NUMERO_CELULAR, CUMPLE_CELULAR, CORREO_ELECTRONICO, ESTADO_REGISTRO, FECHA_CREACION, USUARIO_CREA, 
                        FECHA_MODIFICACION, USUARIO_MODIFICA, ENTIDAD, TIPO_REGISTRO, FECHA_CORTE, NUM, DEN, 
                        ultimo_periodo, renaes, Id_Establecimiento, Nombre_Establecimiento, Ubigueo_Establecimiento, Codigo_Disa, 
                        Disa, Codigo_Red, Red, Codigo_MicroRed, MicroRed, Codigo_Unico, Codigo_Sector, Descripcion_Sector, 
                        PROV, DIST, Categoria_Establecimiento
                    FROM dbo.SEGUIMIENTO_SITUACION_PADRON
                    WHERE
                        -- Filtrar por ubicación geográfica
                        (DEPARTAMENTO = %s)
                        AND (LEFT(COD_UBIGEO, 4) = %s)
                        AND (LEFT(COD_UBIGEO, 6) = %s)
                        -- Filtrar por edad
                        AND {edad_condition}
                        -- Filtrar por cumplimiento
                        AND (
                            %s = ''
                            OR (%s = '1' AND NUM = 1)
                            OR (%s = '0' AND NUM = 0)
                        )
                    ORDER BY edad_anio2, edad_mes2, edad_dias2
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