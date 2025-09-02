from django.http import JsonResponse, HttpResponse
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo
from django.db.models.functions import Substr
from django.db.models import IntegerField
from django.db.models.functions import Cast

from django.db import connection

# Create your views here.
def obtener_distritos(provincia):
    distritos = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Provincia=provincia).values('Distrito').distinct().order_by('Distrito')
    return list(distritos)

## VELOCIMETRO DETALLADOS
def obtener_avance_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            #print(f"[QUERY] Parámetros - Año: {anio}, Mes: {mes_inicio}-{mes_fin}, Provincia: {provincia}, Distrito: {distrito}")
            
            sql_query = '''
                SELECT 
                    SUM(ISNULL(CAST([numerador] AS INT), 0)) AS num,
                    SUM(ISNULL(CAST([denominador] AS INT), 0)) AS den,
                    CASE 
                        WHEN SUM(ISNULL(CAST([denominador] AS INT), 0)) = 0 THEN 0.0
                        ELSE ROUND(
                            (SUM(ISNULL(CAST([numerador] AS INT), 0)) * 100.0) / 
                            SUM(ISNULL(CAST([denominador] AS INT), 0)), 2
                        )
                    END AS avance
                FROM 
                    Compromiso_1.dbo.PAQUETE_COMPROMISO
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
                #print(f"Filtro provincia aplicado: LEFT(Ubigeo, 4) = {provincia}")
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)
                #print(f"Filtro distrito aplicado: Ubigeo = {distrito}")
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            #print(f"[QUERY] SQL: {sql_query.strip()}")
            #print(f"[QUERY] Parámetros: {params}")
            
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
                
        return datos
    except Exception as e:
        print(f"[ERROR] Error al obtener el avance regional: {e}")
        return []

## RESUMEN NUMERADOR Y DENOMINADOR 
def obtener_resumen_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito):
    """
    Obtiene un resumen detallado del paquete compromiso con información adicional
    """
    datos_base = obtener_avance_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito)
    
    if not datos_base:
        return None
    
    resultado = datos_base[0]
    num = resultado.get('num', 0)
    den = resultado.get('den', 0)
    avance = resultado.get('avance', 0.0)
    
    # Calcular métricas adicionales
    brecha = den - num
    porcentaje_brecha = (brecha / den * 100) if den > 0 else 0
    
    # Determinar clasificación
    if avance >= 67:
        clasificacion = "CUMPLE"
        color = "success"
        icono = "check-circle"
    elif avance >= 33:
        clasificacion = "EN PROCESO"
        color = "warning"
        icono = "clock"
    else:
        clasificacion = "EN RIESGO"
        color = "danger"
        icono = "exclamation-triangle"
    
    resumen = {
        'numerador': num,
        'denominador': den,
        'avance': round(avance, 2),
        'brecha': brecha,
        'porcentaje_brecha': round(porcentaje_brecha, 2),
        'clasificacion': clasificacion,
        'color': color,
        'icono': icono,
        'ambito': 'NACIONAL' if not provincia else ('PROVINCIA' if not distrito else 'DISTRITO')
    }
    
    return resumen

## AVANCE REGIONAL MENSUALIZADO
def obtener_avance_regional_mensual_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                    SELECT
                        -- ENERO
                        SUM(CASE WHEN mes = 1 THEN CAST(numerador AS INT) ELSE 0 END) AS num_1,
                        SUM(CASE WHEN mes = 1 THEN CAST(denominador AS INT) ELSE 0 END) AS den_1,
                        CASE
                            WHEN SUM(CASE WHEN mes = 1 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 1 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 1 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_1,
                        -- FEBRERO
                        SUM(CASE WHEN mes = 2 THEN CAST(numerador AS INT) ELSE 0 END) AS num_2,
                        SUM(CASE WHEN mes = 2 THEN CAST(denominador AS INT) ELSE 0 END) AS den_2,
                        CASE
                            WHEN SUM(CASE WHEN mes = 2 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 2 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 2 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_2,
                        -- MARZO
                        SUM(CASE WHEN mes = 3 THEN CAST(numerador AS INT) ELSE 0 END) AS num_3,
                        SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END) AS den_3,
                        CASE
                            WHEN SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 3 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_3,
                        -- ABRIL
                        SUM(CASE WHEN mes = 4 THEN CAST(numerador AS INT) ELSE 0 END) AS num_4,
                        SUM(CASE WHEN mes = 4 THEN CAST(denominador AS INT) ELSE 0 END) AS den_4,
                        CASE
                            WHEN SUM(CASE WHEN mes = 4 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 4 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 4 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_4,
                        -- MAYO
                        SUM(CASE WHEN mes = 5 THEN CAST(numerador AS INT) ELSE 0 END) AS num_5,
                        SUM(CASE WHEN mes = 5 THEN CAST(denominador AS INT) ELSE 0 END) AS den_5,
                        CASE
                            WHEN SUM(CASE WHEN mes = 5 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 5 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 5 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_5,
                        -- JUNIO
                        SUM(CASE WHEN mes = 6 THEN CAST(numerador AS INT) ELSE 0 END) AS num_6,
                        SUM(CASE WHEN mes = 6 THEN CAST(denominador AS INT) ELSE 0 END) AS den_6,
                        CASE
                            WHEN SUM(CASE WHEN mes = 6 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 6 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 6 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_6,
                        -- JULIO
                        SUM(CASE WHEN mes = 7 THEN CAST(numerador AS INT) ELSE 0 END) AS num_7,
                        SUM(CASE WHEN mes = 7 THEN CAST(denominador AS INT) ELSE 0 END) AS den_7,
                        CASE
                            WHEN SUM(CASE WHEN mes = 7 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 7 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 7 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_7,
                        -- AGOSTO
                        SUM(CASE WHEN mes = 8 THEN CAST(numerador AS INT) ELSE 0 END) AS num_8,
                        SUM(CASE WHEN mes = 8 THEN CAST(denominador AS INT) ELSE 0 END) AS den_8,
                        CASE
                            WHEN SUM(CASE WHEN mes = 8 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 8 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 8 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_8,
                        -- SETIEMBRE
                        SUM(CASE WHEN mes = 9 THEN CAST(numerador AS INT) ELSE 0 END) AS num_9,
                        SUM(CASE WHEN mes = 9 THEN CAST(denominador AS INT) ELSE 0 END) AS den_9,
                        CASE
                            WHEN SUM(CASE WHEN mes = 9 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 9 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 9 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_9,
                        -- OCTUBRE
                        SUM(CASE WHEN mes = 10 THEN CAST(numerador AS INT) ELSE 0 END) AS num_10,
                        SUM(CASE WHEN mes = 10 THEN CAST(denominador AS INT) ELSE 0 END) AS den_10,
                        CASE
                            WHEN SUM(CASE WHEN mes = 10 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 10 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 10 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_10,
                        -- NOVIEMBRE
                        SUM(CASE WHEN mes = 11 THEN CAST(numerador AS INT) ELSE 0 END) AS num_11,
                        SUM(CASE WHEN mes = 11 THEN CAST(denominador AS INT) ELSE 0 END) AS den_11,
                        CASE
                            WHEN SUM(CASE WHEN mes = 11 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 11 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 11 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_11,
                        -- DICIEMBRE
                        SUM(CASE WHEN mes = 12 THEN CAST(numerador AS INT) ELSE 0 END) AS num_12,
                        SUM(CASE WHEN mes = 12 THEN CAST(denominador AS INT) ELSE 0 END) AS den_12,
                        CASE
                            WHEN SUM(CASE WHEN mes = 12 THEN CAST(denominador AS INT) ELSE 0 END) = 0
                            THEN 0
                            ELSE ROUND(
                                (
                                    SUM(CASE WHEN mes = 12 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0
                                    / NULLIF(SUM(CASE WHEN mes = 12 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                                ) * 100
                            , 2)
                        END AS cob_12
                    FROM Compromiso_1.dbo.PAQUETE_COMPROMISO
            '''
            params = []
            conditions = []

            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)

            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
                
        return datos
    except Exception as e:
        return []

## VARIABLES DETALLADOS
def obtener_variables_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            
            sql_query = '''
                    SELECT 
                        -- ind cred
                        SUM(ISNULL(CAST(denominador AS INT), 0)) AS den_variable,
                        SUM(ISNULL(CAST(num_cred AS INT), 0)) AS num_cred,
                    	    CASE 
                    	    WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	    ELSE ROUND(
                    	        (SUM(ISNULL(CAST(num_cred AS INT), 0)) * 100.0) / 
                    	        SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	    )
                    	END AS avance_cred,
                    	-- cred rn
                        SUM(ISNULL(CAST(num_cred_rn AS INT), 0)) AS num_cred_rn,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_cred_rn AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_cred_rn,
                    	-- cred mensual
                        SUM(ISNULL(CAST(num_cred_mensual AS INT), 0)) AS num_cred_mensual,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_cred_mensual AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_cred_mensual,
                    	-- ind vac
                    	SUM(ISNULL(CAST(num_vac AS INT), 0)) AS num_vac,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_vac AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_vac,
                    	-- vac antineumococica
                    	SUM(ISNULL(CAST(num_vac_antineumococica AS INT), 0)) AS num_vac_antineumococica,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_vac_antineumococica AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_vac_antineumococica,
                    	-- vac antipolio
                    	SUM(ISNULL(CAST(num_vac_antipolio AS INT), 0)) AS num_vac_antipolio,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_vac_antipolio AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_vac_antipolio,
                    	-- vac pentavalente
                    	SUM(ISNULL(CAST(num_vac_pentavalente AS INT), 0)) AS num_vac_pentavalente,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_vac_pentavalente AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_vac_pentavalente,
                    	-- vac rotavirus
                    	SUM(ISNULL(CAST(num_vac_rotavirus AS INT), 0)) AS num_vac_rotavirus,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_vac_rotavirus AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_vac_rotavirus,
                    	-- esq suplementacion 
                    	SUM(ISNULL(CAST(num_esq AS INT), 0)) AS num_esq,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_esq AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_esq,
                    	-- esq4M 
                    	SUM(ISNULL(CAST(num_esq4M AS INT), 0)) AS num_esq4M,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_esq4M AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_esq4M,
                    	-- esq6M
                    	SUM(ISNULL(CAST(num_esq6M AS INT), 0)) AS num_esq6M,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_esq6M AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_esq6M,
                    	-- num_esq6M_trat
                    	SUM(ISNULL(CAST(num_esq6M_trat AS INT), 0)) AS num_esq6M_trat,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_esq6M_trat AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_esq6M_trat,
                    	-- num_esq6M_multi
                    	SUM(ISNULL(CAST(num_esq6M_multi AS INT), 0)) AS num_esq6M_multi,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_esq6M_multi AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_esq6M_multi,
                    	-- num_dosaje_Hb
                    	SUM(ISNULL(CAST(num_dosaje_Hb AS INT), 0)) AS num_dosaje_Hb,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_dosaje_Hb AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_num_dosaje_Hb,
                    	-- num_DNIemision
                    	SUM(ISNULL(CAST(num_DNIemision AS INT), 0)) AS num_DNIemision,
                    	CASE 
                    	   WHEN SUM(ISNULL(CAST(denominador AS INT), 0)) = 0 THEN 0.0
                    	   ELSE ROUND(
                    	       (SUM(ISNULL(CAST(num_DNIemision AS INT), 0)) * 100.0) / 
                    	       SUM(ISNULL(CAST(denominador AS INT), 0)), 2
                    	   )
                    	END AS avance_DNIemision
                    FROM 
                        Compromiso_1.dbo.PAQUETE_COMPROMISO
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
        return datos
    except Exception as e:
        return []



def obtener_ranking_paquete_compromiso(anio, mes, red, microred, establecimiento, provincia, distrito):
    with connection.cursor() as cursor:
        # Base query with aggregation
        sql_query = """
            SELECT
                red,
                microred,
                Nombre_Establecimiento,
                SUM(ISNULL(CAST(denominador AS INT), 0)) as total_denominador,
                SUM(ISNULL(CAST(numerador AS INT), 0)) as total_numerador,
                SUM(ISNULL(CAST(denominador AS INT), 0) - ISNULL(CAST(numerador AS INT), 0)) as total_brecha
            FROM 
                Indicadores_FED.dbo.MC01_PaqueteGestante_Combinado
            WHERE 1=1
        """
        params = []

        # Appending filters
        if red and red != '':
            sql_query += " AND Codigo_Red = %s"
            params.append(red)
        if microred and microred != '':
            sql_query += " AND Codigo_MicroRed = %s"
            params.append(microred)
        if establecimiento and establecimiento != '':
            sql_query += " AND Codigo_Unico = %s"
            params.append(establecimiento)
        if provincia and provincia != '':
            sql_query += " AND Provincia = %s"
            params.append(provincia)
        if distrito and distrito != '':
            sql_query += " AND Distrito = %s"
            params.append(distrito)

        # Grouping and ordering
        sql_query += """
            GROUP BY
                red,
                microred,
                Nombre_Establecimiento
            ORDER BY
                red,
                microred,
                Nombre_Establecimiento;
        """

        cursor.execute(sql_query, params)
        result = cursor.fetchall()
        return result

## AVANCE REGIONAL
def obtener_avance_regional_paquete_compromiso():
    try:
        # Asegúrate de que la conexión a la base de datos está establecida
        with connection.cursor() as cursor:
            cursor.execute(
                '''
                    SELECT
                    -- ENERO
                    SUM(CASE WHEN mes = 3 THEN CAST(numerador AS INT) ELSE 0 END) AS num,
                    SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END) AS den,
                    CASE 
                        WHEN SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END) = 0 								
                        THEN 0 
                        ELSE ROUND(
                            (
                                SUM(CASE WHEN mes = 3 THEN CAST(numerador AS INT) ELSE 0 END) * 1.0 
                                / NULLIF(SUM(CASE WHEN mes = 3 THEN CAST(denominador AS INT) ELSE 0 END), 0)
                            ) * 100
                        , 2) 
                    END AS cob
                    FROM MC01_PaqueteGestante_Combinado
					WHERE "año" = '2025'
                ''',
            )
            resultados = cursor.fetchall()
            
            # Obtener los nombres de las columnas
            column_names = [desc[0] for desc in cursor.description]
            
            # Convertir cada fila en un diccionario
            datos = [dict(zip(column_names, fila)) for fila in resultados]
        
        return datos
    except Exception as e:
        #print(f"Error al obtener el avance regional: {e}")
        return []

#############################
## Cobertura con filtros
#############################
## COBERTURA POR ZONA
def obtener_cobertura_por_zona(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                        SELECT
                            ZONA as z_zona,
                            SUM(ISNULL(denominador, 0)) as z_den,
                            SUM(ISNULL(numerador, 0)) as z_num,
                            SUM(ISNULL(denominador - numerador, 0)) as z_brecha,
                            CASE 
                                WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                                ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                            END as z_cob  
                        FROM [Compromiso_1].[dbo].[PAQUETE_COMPROMISO]
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            ZONA      
                        ORDER BY
                            ZONA ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
        return datos
    except Exception as e:
        return []

## COBERTURA POR PROVINCIA
def obtener_cobertura_por_provincia(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                    SELECT
                        Provincia as p_provincia,
                        SUM(ISNULL(denominador, 0)) as p_den,
                        SUM(ISNULL(numerador, 0)) as p_num,
                        SUM(ISNULL(denominador - numerador, 0)) as p_brecha,
                        CASE 
                            WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                            ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                        END as p_cob  
                    FROM [Compromiso_1].[dbo].[PAQUETE_COMPROMISO]
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            Provincia      
                        ORDER BY
                            Provincia ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
        return datos
    except Exception as e:
        return []

## COBERTURA POR DISTRITO
def obtener_cobertura_por_distrito(anio, mes_inicio, mes_fin, provincia, distrito):
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                    SELECT
                        Distrito as d_distrito,
                        SUM(ISNULL(denominador, 0)) as d_den,
                        SUM(ISNULL(numerador, 0)) as d_num,
                        SUM(ISNULL(denominador - numerador, 0)) as d_brecha,
                        CASE 
                            WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                            ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                        END as d_cob  
                    FROM [Compromiso_1].[dbo].[PAQUETE_COMPROMISO]
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(ubigeo, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("ubigeo = %s")
                params.append(distrito)
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            Distrito      
                        ORDER BY
                            Distrito ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
        return datos
    except Exception as e:
        return []


def dictfetchall(cursor):
    """
    Devuelve todas las filas de un cursor como una lista de diccionarios.
    """
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def obtener_avance_cobertura_paquete_compromiso(anio, mes, red_h, p_microredes_establec_h, p_establecimiento_h, provincia, distrito):
    """
    Obtiene los datos de cobertura de población, agrupados por red, microred y establecimiento,
    con cálculos agregados y el porcentaje de cobertura para cada grupo.
    
    Parámetros:
    - anio: Año de consulta
    - mes: Mes de consulta
    - red_h: Código de red
    - p_microredes_establec_h: Código de microred
    - p_establecimiento_h: Código de establecimiento
    """
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                        SELECT
                            anio,
                            mes,
                            Ubigueo_Establecimiento,
                            Distrito,
                            Provincia,
                            Codigo_Red,
                            red,
                            Codigo_MicroRed,
                            microred,
                            Codigo_Unico,
                            Nombre_Establecimiento,
                            grupo_edad,
                            SUM(ISNULL(denominador, 0)) as total_denominador,
                            SUM(ISNULL(numerador, 0)) as total_numerador,
                            SUM(ISNULL(denominador, 0) - ISNULL(numerador, 0)) as total_brecha,
                            cobertura_porcentaje
                        FROM [Padron_Nominal].[dbo].[cobertura_poblacion]
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año y mes (siempre necesarios)
            if anio:
                conditions.append("anio = %s")
                params.append(anio)
            if mes:
                conditions.append("mes = %s")
                params.append(mes)
            
            # Construir condiciones dinámicamente para filtros adicionales
            if red_h and red_h != '':
                conditions.append("Codigo_Red = %s")
                params.append(red_h)
            if p_microredes_establec_h and p_microredes_establec_h != '':
                conditions.append("Codigo_MicroRed = %s")
                params.append(p_microredes_establec_h)
            if p_establecimiento_h and p_establecimiento_h != '':
                conditions.append("Codigo_Unico = %s")
                params.append(p_establecimiento_h)
            if provincia and provincia != '':
                conditions.append("Provincia = %s")
                params.append(provincia)
            if distrito and distrito != '':
                conditions.append("Distrito = %s")
                params.append(distrito)
            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            anio, 
                            mes,
                            Ubigueo_Establecimiento,
                            Distrito,
                            Provincia,  
                            Codigo_Red,
                            red,
                            Codigo_MicroRed,
                            microred,
                            Codigo_Unico,
                            Nombre_Establecimiento,
                            grupo_edad,
                            cobertura_porcentaje
                        ORDER BY
                            grupo_edad
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
        return datos
    except Exception as e:
        #print(f"Error al obtener el avance regional: {e}")
        return []


def obtener_cobertura_por_red(anio, mes, red_h, p_microredes_establec_h, p_establecimiento_h, provincia, distrito):
    """
    Obtiene los datos de cobertura de población, agrupados por RED,
    para mostrar EN GRAFICO DE BARRAS.
    
    Parámetros:
    - anio: Año de consulta
    - mes: Mes de consulta
    - red_h: Código de red
    - p_microredes_establec_h: Código de microred
    - p_establecimiento_h: Código de establecimiento
    - provincia: Provincia
    - distrito: Distrito
    """
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                        SELECT
                            red as r_red,
                            SUM(ISNULL(denominador, 0)) as r_denominador,
                            SUM(ISNULL(numerador, 0)) as r_numerador,
                            SUM(ISNULL(brecha, 0)) as r_brecha,
                            CASE 
                                WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                                ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                            END as r_cobertura
                            
                        FROM [Padron_Nominal].[dbo].[cobertura_poblacion]
            '''
            params = []
            conditions = []
            
            if anio:
                conditions.append("anio = %s")
                params.append(anio)
            if mes:
                conditions.append("mes = %s")
                params.append(mes)
            
            if red_h and red_h != '':
                conditions.append("Codigo_Red = %s")
                params.append(red_h)
            if p_microredes_establec_h and p_microredes_establec_h != '':
                conditions.append("Codigo_MicroRed = %s")
                params.append(p_microredes_establec_h)
            if p_establecimiento_h and p_establecimiento_h != '':
                conditions.append("Codigo_Unico = %s")
                params.append(p_establecimiento_h)
            if provincia and provincia != '':
                conditions.append("Provincia = %s")
                params.append(provincia)
            if distrito and distrito != '':
                conditions.append("Distrito = %s")
                params.append(distrito)
            
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            red
                        ORDER BY
                            red ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            return datos
    except Exception as e:
        #print(f"Error al obtener cobertura por red: {e}")
        return []    


def obtener_cobertura_por_microred(anio, mes, red_h, p_microredes_establec_h, p_establecimiento_h, provincia, distrito):
    """
    Obtiene los datos de cobertura de población, agrupados por RED,
    para mostrar EN GRAFICO DE BARRAS.
    
    Parámetros:
    - anio: Año de consulta
    - mes: Mes de consulta
    - red_h: Código de red
    - p_microredes_establec_h: Código de microred
    - p_establecimiento_h: Código de establecimiento
    - provincia: Provincia
    - distrito: Distrito
    """
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                        SELECT
                            microred as m_microred,
                            SUM(ISNULL(denominador, 0)) as m_denominador,
                            SUM(ISNULL(numerador, 0)) as m_numerador,
                            SUM(ISNULL(brecha, 0)) as m_brecha,
                            CASE 
                                WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                                ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                            END as m_cobertura
                            
                        FROM [Padron_Nominal].[dbo].[cobertura_poblacion]
            '''
            params = []
            conditions = []
            
            if anio:
                conditions.append("anio = %s")
                params.append(anio)
            if mes:
                conditions.append("mes = %s")
                params.append(mes)
            
            if red_h and red_h != '':
                conditions.append("Codigo_Red = %s")
                params.append(red_h)
            if p_microredes_establec_h and p_microredes_establec_h != '':
                conditions.append("Codigo_MicroRed = %s")
                params.append(p_microredes_establec_h)
            if p_establecimiento_h and p_establecimiento_h != '':
                conditions.append("Codigo_Unico = %s")
                params.append(p_establecimiento_h)
            if provincia and provincia != '':
                conditions.append("Provincia = %s")
                params.append(provincia)
            if distrito and distrito != '':
                conditions.append("Distrito = %s")
                params.append(distrito)
            
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            microred
                        ORDER BY
                            microred ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            return datos
    except Exception as e:
        #print(f"Error al obtener cobertura por red: {e}")
        return []    


def obtener_cobertura_por_establecimiento(anio, mes, red_h, p_microredes_establec_h, p_establecimiento_h, provincia, distrito):
    """
    Obtiene los datos de cobertura de población, agrupados por RED,
    para mostrar EN GRAFICO DE BARRAS.
    
    Parámetros:
    - anio: Año de consulta
    - mes: Mes de consulta
    - red_h: Código de red
    - p_microredes_establec_h: Código de microred
    - p_establecimiento_h: Código de establecimiento
    - provincia: Provincia
    - distrito: Distrito
    """
    try:
        with connection.cursor() as cursor:
            sql_query = '''
                        SELECT
                            Nombre_Establecimiento as e_establecimiento,
                            SUM(ISNULL(denominador, 0)) as e_denominador,
                            SUM(ISNULL(numerador, 0)) as e_numerador,
                            SUM(ISNULL(brecha, 0)) as e_brecha,
                            CASE 
                                WHEN SUM(ISNULL(denominador, 0)) = 0 THEN 0
                                ELSE ROUND((SUM(ISNULL(numerador, 0)) * 100.0) / SUM(ISNULL(denominador, 0)), 2)
                            END as e_cobertura
                            
                        FROM [Padron_Nominal].[dbo].[cobertura_poblacion]
            '''
            params = []
            conditions = []
            
            if anio:
                conditions.append("anio = %s")
                params.append(anio)
            if mes:
                conditions.append("mes = %s")
                params.append(mes)
            
            if red_h and red_h != '':
                conditions.append("Codigo_Red = %s")
                params.append(red_h)
            if p_microredes_establec_h and p_microredes_establec_h != '':
                conditions.append("Codigo_MicroRed = %s")
                params.append(p_microredes_establec_h)
            if p_establecimiento_h and p_establecimiento_h != '':
                conditions.append("Codigo_Unico = %s")
                params.append(p_establecimiento_h)
            if provincia and provincia != '':
                conditions.append("Provincia = %s")
                params.append(provincia)
            if distrito and distrito != '':
                conditions.append("Distrito = %s")
                params.append(distrito)
            
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            sql_query += '''
                        GROUP BY
                            Nombre_Establecimiento
                        ORDER BY
                            Nombre_Establecimiento ASC
            '''
    
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            return datos
    except Exception as e:
        #print(f"Error al obtener cobertura por red: {e}")
        return []    

#############################
## REPORTES EN EXCEL EN SALUD
#############################

def obtener_seguimiento_paquete_compromiso(anio, mes_inicio, mes_fin, provincia, distrito, p_red, p_microredes, p_establecimiento, p_cumple):
    try:
        with connection.cursor() as cursor:
            
            sql_query = '''
                SELECT
                    tipo_doc,
                    num_doc,
                    CONVERT(VARCHAR(10), fecha_nac, 103) AS fecha_nac,
                    sexo,
                    seguro,
                    edad_dias,
                    edad_mes,
                    flag_cnv,
                    peso_cnv,
                    flag_BPN,
                    Semana_gest_cnv,
                    flag_prematuro,
                    flag_BPN_Prematuro,
                    flag_indicador,
                    numerador_sinDNI,
                    num_cred,
                    num_cred_rn,
                    CONVERT(VARCHAR(10), fecha_cred_rn1, 103) AS fecha_cred_rn1,
                    num_cred_rn1,
                    CONVERT(VARCHAR(10), fecha_cred_rn2, 103) AS fecha_cred_rn2,
                    num_cred_rn2,
                    CONVERT(VARCHAR(10), fecha_cred_rn3, 103) AS fecha_cred_rn3,
                    num_cred_rn3,
                    CONVERT(VARCHAR(10), fecha_cred_rn4, 103) AS fecha_cred_rn4,
                    num_cred_rn4,
                    num_cred_mensual,
                    CONVERT(VARCHAR(10), fecha_cred_mes1, 103) AS fecha_cred_mes1,
                    num_cred_mes1,
                    CONVERT(VARCHAR(10), fecha_cred_mes2, 103) AS fecha_cred_mes2,
                    num_cred_mes2,
                    CONVERT(VARCHAR(10), fecha_cred_mes3, 103) AS fecha_cred_mes3,
                    num_cred_mes3,
                    CONVERT(VARCHAR(10), fecha_cred_mes4, 103) AS fecha_cred_mes4,
                    num_cred_mes4,
                    CONVERT(VARCHAR(10), fecha_cred_mes5, 103) AS fecha_cred_mes5,
                    num_cred_mes5,
                    CONVERT(VARCHAR(10), fecha_cred_mes6, 103) AS fecha_cred_mes6,
                    num_cred_mes6,
                    CONVERT(VARCHAR(10), fecha_cred_mes7, 103) AS fecha_cred_mes7,
                    num_cred_mes7,
                    CONVERT(VARCHAR(10), fecha_cred_mes8, 103) AS fecha_cred_mes8,
                    num_cred_mes8,
                    CONVERT(VARCHAR(10), fecha_cred_mes9, 103) AS fecha_cred_mes9,
                    num_cred_mes9,
                    CONVERT(VARCHAR(10), fecha_cred_mes10, 103) AS fecha_cred_mes10,
                    num_cred_mes10,
                    CONVERT(VARCHAR(10), fecha_cred_mes11, 103) AS fecha_cred_mes11,
                    num_cred_mes11,
                    num_vac,
                    num_vac_antineumococica,
                    CONVERT(VARCHAR(10), fecha_vac_antineumococica1, 103) AS fecha_vac_antineumococica1,
                    num_vac_antineumococica1,
                    CONVERT(VARCHAR(10), fecha_vac_antineumococica2, 103) AS fecha_vac_antineumococica2,
                    num_vac_antineumococica2,
                    num_vac_antipolio,
                    CONVERT(VARCHAR(10), fecha_vac_antipolio1, 103) AS fecha_vac_antipolio1,
                    num_vac_antipolio1,
                    CONVERT(VARCHAR(10), fecha_vac_antipolio2, 103) AS fecha_vac_antipolio2,
                    num_vac_antipolio2,
                    CONVERT(VARCHAR(10), fecha_vac_antipolio3, 103) AS fecha_vac_antipolio3,
                    num_vac_antipolio3,
                    num_vac_pentavalente,
                    CONVERT(VARCHAR(10), fecha_vac_pentavalente1, 103) AS fecha_vac_pentavalente1,
                    num_vac_pentavalente1,
                    CONVERT(VARCHAR(10), fecha_vac_pentavalente2, 103) AS fecha_vac_pentavalente2,
                    num_vac_pentavalente2,
                    CONVERT(VARCHAR(10), fecha_vac_pentavalente3, 103) AS fecha_vac_pentavalente3,
                    num_vac_pentavalente3,
                    num_vac_rotavirus,
                    CONVERT(VARCHAR(10), fecha_vac_rotavirus1, 103) AS fecha_vac_rotavirus1,
                    num_vac_rotavirus1,
                    CONVERT(VARCHAR(10), fecha_vac_rotavirus2, 103) AS fecha_vac_rotavirus2,
                    num_vac_rotavirus2,
                    num_esq,
                    num_esq4M,
                    CONVERT(VARCHAR(10), fecha_Esq4m_sup_E1, 103) AS fecha_Esq4m_sup_E1,
                    num_Esq4m_sup_E1,
                    num_esq6M,
                    num_esq6M_sup,
                    CONVERT(VARCHAR(10), fecha_Esq6m_sup_E1, 103) AS fecha_Esq6m_sup_E1,
                    num_Esq6m_sup_E1,
                    CONVERT(VARCHAR(10), fecha_Esq6m_sup_E2, 103) AS fecha_Esq6m_sup_E2,
                    num_Esq6m_sup_E2,
                    num_esq6M_trat,
                    CONVERT(VARCHAR(10), fecha_Esq6m_trat_E1, 103) AS fecha_Esq6m_trat_E1,
                    num_Esq6m_trat_E1,
                    CONVERT(VARCHAR(10), fecha_Esq6m_trat_E2, 103) AS fecha_Esq6m_trat_E2,
                    num_Esq6m_trat_E2,
                    CONVERT(VARCHAR(10), fecha_Esq6m_trat_E3, 103) AS fecha_Esq6m_trat_E3,
                    num_Esq6m_trat_E3,
                    num_esq6M_multi,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E1, 103) AS fecha_Esq6m_multi_E1,
                    num_Esq6m_multi_E1,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E2, 103) AS fecha_Esq6m_multi_E2,
                    num_Esq6m_multi_E2,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E3, 103) AS fecha_Esq6m_multi_E3,
                    num_Esq6m_multi_E3,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E4, 103) AS fecha_Esq6m_multi_E4,
                    num_Esq6m_multi_E4,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E5, 103) AS fecha_Esq6m_multi_E5,
                    num_Esq6m_multi_E5,
                    CONVERT(VARCHAR(10), fecha_Esq6m_multi_E6, 103) AS fecha_Esq6m_multi_E6,
                    num_Esq6m_multi_E6,
                    num_dosaje_Hb,
                    CONVERT(VARCHAR(10), fecha_Hb, 103) AS fecha_Hb,
                    num_Hb,
                    num_DNIemision,
                    CONVERT(VARCHAR(10), fecha_DNIemision, 103) AS fecha_DNIemision,
                    num_DNIemision_30d,
                    num_DNIemision_60d,
                    CASE mes
                        WHEN 1 THEN 'ENERO' WHEN 2 THEN 'FEBRERO' WHEN 3 THEN 'MARZO' WHEN 4 THEN 'ABRIL'
                        WHEN 5 THEN 'MAYO' WHEN 6 THEN 'JUNIO' WHEN 7 THEN 'JULIO' WHEN 8 THEN 'AGOSTO'
                        WHEN 9 THEN 'SETIEMBRE' WHEN 10 THEN 'OCTUBRE' WHEN 11 THEN 'NOVIEMBRE' WHEN 12 THEN 'DICIEMBRE'
                    END AS mes_nombre,
                    CASE WHEN numerador = '1' THEN 'CUMPLE' ELSE 'NO CUMPLE' END AS IND,
                    Ubigueo_Establecimiento,
                    Provincia,
                    Distrito,
                    Red,
                    MicroRed,
                    Id_Establecimiento,
                    Nombre_Establecimiento
            FROM 
                Compromiso_1.DBO.PAQUETE_COMPROMISO
            '''
            params = []
            conditions = []
            
            # Agregar filtros de año
            if anio:
                conditions.append("año = %s")
                params.append(anio)

            # Agregar filtro de mes con BETWEEN
            if mes_inicio and mes_fin:
                conditions.append("mes BETWEEN %s AND %s")
                params.append(mes_inicio)
                params.append(mes_fin)
            elif mes_inicio:
                conditions.append("mes = %s")
                params.append(mes_inicio)
            
            # Filtros de ubicación geográfica - usando LIKE para códigos de ubigeo
            if provincia and provincia != '':
                conditions.append("LEFT(Ubigueo_Establecimiento, 4) = %s")
                params.append(provincia)
            
            if distrito and distrito != '':
                conditions.append("Ubigueo_Establecimiento = %s")
                params.append(distrito)
                
            # Filtros de salud - usando LIKE para códigos de ubigeo    
            if p_red and p_red != '':
                conditions.append("LEFT(Codigo_Red, 2) = %s")
                params.append(p_red)
            
            if p_microredes and p_microredes != '':
                conditions.append("LEFT(Codigo_MicroRed, 4) = %s")
                params.append(p_microredes)
            
            if p_establecimiento and p_establecimiento != '':
                conditions.append("Id_Establecimiento = %s")
                params.append(p_establecimiento)
            
            # Agregar filtro de cumplimiento del indicador
            if p_cumple == '1':
                conditions.append("ts.numerador = '1'")
            elif p_cumple == '0':
                conditions.append("ts.numerador <> '1'")
            # Si p_cumple = '' no se agrega filtro (todos los registros)            
            # Agregar WHERE solo si hay condiciones
            if conditions:
                sql_query += " WHERE " + " AND ".join(conditions)
            
            cursor.execute(sql_query, params)
            resultados = cursor.fetchall()
            column_names = [desc[0] for desc in cursor.description]
            datos = [dict(zip(column_names, fila)) for fila in resultados]
            
            print(f"[QUERY] SQL: {sql_query.strip()}")
            print(f"[QUERY] Parámetros: {params}")
            
        return datos
    except Exception as e:
        return []


### --- 
def obtener_seguimiento_paquete_compromiso_red(departamento, red, edad, cumple):
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
                FROM Padron_Nominal_web.dbo.PN_FED_ESTABLECIMIENTOS
                WHERE
                    -- Filtrar por ubicación geográfica
                    (DEPARTAMENTO = %s OR %s = '')
                    AND (LEFT(Codigo_Red, 4) = %s OR %s = '')
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
                red, red,
                cumple, cumple, cumple
            ]
        )
        
        # Obtener los resultados
        return cursor.fetchall()

def obtener_seguimiento_paquete_compromiso_microred(departamento, red, microred, edad, cumple):
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
                FROM Padron_Nominal_web.dbo.PN_FED_ESTABLECIMIENTOS
                WHERE
                    -- Filtrar por ubicación geográfica
                    (DEPARTAMENTO = %s OR %s = '')
                    AND (LEFT(Codigo_Red_Backup, 2) = %s OR %s = '')
                    AND (LEFT(Codigo_MicroRed_Backup, 2) = %s OR %s = '')
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
                red, red,
                microred, microred,
                cumple, cumple, cumple
            ]
        )
        
        # Obtener los resultados
        return cursor.fetchall()

def obtener_seguimiento_paquete_compromiso_establecimiento(departamento, establecimiento, edad, cumple):
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
                    PROV, DIST, Categoria_Establecimiento, CODIGO_EESS_ACTUALIZADO
                FROM Padron_Nominal_web.dbo.PN_FED_ESTABLECIMIENTOS
                WHERE
                    -- Filtrar por ubicación geográfica
                    (DEPARTAMENTO = %s OR %s = '')
                    AND (
                        %s = '' OR
                        CODIGO_EESS_ACTUALIZADO = %s OR
                        LEFT(CODIGO_EESS_ACTUALIZADO,9) = %s OR
                        Codigo_Unico = %s
                    )
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
                establecimiento, establecimiento, establecimiento, establecimiento,
                cumple, cumple, cumple
            ]
        )
        
        # Obtener los resultados
        return cursor.fetchall()