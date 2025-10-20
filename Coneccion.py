# Coneccion.py

import fdb

# Lista para almacenar los datos de conexión
lista = []

# Función para la conexión y otras operaciones
def coneccion(host, database, user, password, accion, consulta="", datos=None):
    global lista  # Acceder a la lista global dentro de la función
    lista = []
    try:
        conn = fdb.connect(
            host=host,
            database=database,
            user=user,
            password=password,
            charset='ISO8859_1'
        )
        res = "Conexión exitosa a Firebird!"
        cursor = conn.cursor()
        if accion == "consulta":
            if consulta:
                if datos:
                    cursor.execute(consulta, datos)  # Ejecutar la consulta con los datos proporcionados
                else:
                    cursor.execute(consulta)
                rows = cursor.fetchall()
                lista.clear()  # Limpiar la lista antes de obtener nuevos datos
                for row in rows:
                    lista.append(row)
        elif accion == "insertar":
            if consulta and datos:
                cursor.execute(consulta, datos)
                conn.commit()
        cursor.close()
        conn.close()
        return res
    except fdb.fbcore.DatabaseError as db_err:
        print(f"Error de base de datos: {db_err}")
        return f"Error de base de datos: {db_err}"
    except Exception as e:
        return f"Error al conectar a Firebird: {e}"
