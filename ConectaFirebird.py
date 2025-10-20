import datetime
import fdb

# Configura los detalles de la conexión
'''Si estás intentando conectarte a un servidor remoto, reemplaza "localhost" con la dirección IP o el nombre de host del servidor Firebird.'''
lista = []
# Intentar conectarse a la base de datos
def coneccion(host,database, user,password,consulta):
    try:
        conn = fdb.connect(
            host = host,
            database = database,
            user = user,
            password = password,
            charset='UTF8'
            )
        print("Conexión exitosa a Firebird!")
        cursor = conn.cursor()
        cursor.execute(consulta)
        rows = cursor.fetchall()
        lista.clear()  # Limpiar la lista antes de obtener nuevos datos
        for row in rows:
            lista.append(row)
        print(lista)
        conn.close()
    except Exception as e:
        print("Error al conectar a Firebird:", e)

host = '192.168.128.51'
database='E:\\Sys\\Microsip datos\\COPAC 2024.FDB'
user='SYSDBA'
password = 'masterkey'
#MG_ORDENES_PROD.FOLIO, MG_ORDENES_PROD.SOLICITANTE, MG_ORDENES_PROD.OBSERVACIONES 
#MG_ORDENES_PROD.OBSERVACIONES,
#consulta = "SELECT DOCTOS_IN_DET.UNIDADES, DOCTOS_IN_DET.FECHA, DOCTOS_IN.USUARIO_CREADOR, DOCTOS_IN.FECHA_HORA_ULT_MODIF, MG_ORDENES_PROD.FOLIO, MG_ORDENES_PROD.SOLICITANTE, MG_ORDENES_PROD.OBSERVACIONES FROM DOCTOS_IN_DET INNER JOIN DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID INNER JOIN MG_ORDENES_PROD ON CAST(SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 1 FOR CHAR_LENGTH(DOCTOS_IN.DESCRIPCION) - POSITION(':' IN DOCTOS_IN.DESCRIPCION)) AS INT) = MG_ORDENES_PROD.FOLIO WHERE DOCTOS_IN_DET.FECHA = '2024-06-30' AND DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3D';"
#consulta= "SELECT DOCTOS_IN.USUARIO_CREADOR FROM DOCTOS_IN where DOCTOS_IN.FECHA = '2024-06-29' AND (DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR13' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3B' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3C' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3D' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3E' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3F' OR DOCTOS_IN.USUARIO_CREADOR = 'IMPRESION3' OR DOCTOS_IN.USUARIO_CREADOR = 'INOCUIDAD03');"
consulta2 = """
SELECT 
    DOCTOS_IN_DET.UNIDADES, 
    DOCTOS_IN_DET.FECHA, 
    DOCTOS_IN.USUARIO_CREADOR, 
    DOCTOS_IN.FECHA_HORA_ULT_MODIF,  
    MG_ORDENES_PROD.FOLIO, 
    MG_ORDENES_PROD.SOLICITANTE, 
    MG_ORDENES_PROD.OBSERVACIONES 
FROM 
    DOCTOS_IN_DET 
    INNER JOIN DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
    INNER JOIN MG_ORDENES_PROD ON CAST(SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 1 FOR CHAR_LENGTH(DOCTOS_IN.DESCRIPCION) - POSITION(':' IN DOCTOS_IN.DESCRIPCION)) AS INT) = MG_ORDENES_PROD.FOLIO 
WHERE 
    DOCTOS_IN.FECHA_HORA_ULT_MODIF BETWEEN CAST('2024-06-28 00:00:00.000' AS TIMESTAMP) AND CAST('2024-06-28 23:59:59.998' AS  TIMESTAMP)
    AND (DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3A' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3B' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3C' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3E' OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3F' OR DOCTOS_IN.USUARIO_CREADOR = 'IMPRESION3' OR DOCTOS_IN.USUARIO_CREADOR = 'INOCUIDAD03');
"""
consulta= """
SELECT 
    MG_ORDENES_PROD.FOLIO, 
    MG_ORDENES_PROD.SOLICITANTE, 
    MG_ORDENES_PROD.OBSERVACIONES,
    DOCTOS_IN_DET.UNIDADES
FROM 
    DOCTOS_IN_DET 
    INNER JOIN DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
    INNER JOIN MG_ORDENES_PROD ON CAST(SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 1 FOR CHAR_LENGTH(DOCTOS_IN.DESCRIPCION) - POSITION(':' IN DOCTOS_IN.DESCRIPCION)) AS INT) = MG_ORDENES_PROD.FOLIO 
WHERE 
    DOCTOS_IN.FECHA_HORA_ULT_MODIF BETWEEN CAST('2024-06-28 00:00:00.000' AS TIMESTAMP) AND CAST('2024-06-28 23:59:59.998' AS  TIMESTAMP)
    AND (DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3A' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3B' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3C' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3E' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'SUPERVISOR3F' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'IMPRESION3' 
    OR DOCTOS_IN.USUARIO_CREADOR = 'INOCUIDAD03');
"""

fecha=datetime.date.today()
hora_actual = datetime.datetime.now()
fecha_inicio_consulta = '2024-07-08 22:00:00.000'
fecha_fin_consulta = '2024-07-09 06:00:00.000'

consulta_diaria = f"""
SELECT
    MG_ORDENES_PROD.FOLIO, 
    MG_ORDENES_PROD.SOLICITANTE, 
    MG_ORDENES_PROD.OBSERVACIONES,
    DOCTOS_IN_DET.UNIDADES
FROM 
    DOCTOS_IN_DET 
    INNER JOIN DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
    INNER JOIN MG_ORDENES_PROD ON CAST(SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 2 FOR CHAR_LENGTH(DOCTOS_IN.DESCRIPCION) - POSITION(':' IN DOCTOS_IN.DESCRIPCION)) AS INT) = MG_ORDENES_PROD.FOLIO 
WHERE 
    DOCTOS_IN.FECHA_HORA_ULT_MODIF BETWEEN {fecha_inicio_consulta}' AS TIMESTAMP) AND CAST('{fecha_fin_consulta}' AS  TIMESTAMP);
"""

coneccion(host,database,user,password,consulta_diaria)
"""
def sumar_campos(lista_tuplas):
    acumulador = {}
    detalles = {}
    # Itera sobre cada tupla en la lista
    for folio,solicitante,observaciones,unidades in lista_tuplas:
        # Si la clave ya está en el diccionario, suma el valor
        if folio in acumulador:
            acumulador[folio] += unidades
        # Si la clave no está en el diccionario, inicializa con el valor
        else:
            acumulador[folio] = unidades
        detalles[folio] = (folio,solicitante, observaciones)
    # Convierte el diccionario acumulador de nuevo en una lista de tuplas
    resultado = [
        ( detalles[folio][0], detalles[folio][1],detalles[folio][2],acumulador[folio]) 
        for folio in acumulador
    ]
    return resultado

# Ejemplo de uso
resultado = sumar_campos(lista)
print(resultado)

"""