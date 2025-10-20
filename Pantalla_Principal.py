import datetime
from decimal import Decimal
from email.mime import image
import threading
import time
import tkinter as tk
from tkinter import Canvas, ttk
from tkinter import font
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import Coneccion
from tkinter import messagebox as mb

user = ""
pasw = ""
#consulta = "SELECT MG_ORDENES_PROD.FOLIO, MG_ORDENES_PROD.SOLICITANTE,MG_ORDENES_PROD.OBSERVACIONES,  MG_ORDENES_PROD_ARTS.CANTIDAD_POR_PROD,MG_ORDENES_PROD_ARTS.CANTIDAD_PRODUCIDA FROM MG_ORDENES_PROD INNER JOIN MG_ORDENES_PROD_ARTS ON MG_ORDENES_PROD_ARTS.ORDEN_PROD_ID = MG_ORDENES_PROD.ORDEN_PROD_ID AND  MG_ORDENES_PROD.ESTATUS= 'P'"
ban=True
hora_actual=""
turnoActu="" 
turnoAnte=""
lista=[]
resultado=[]
fecha_inicio_consulta=''
fecha_fin_consulta=''
fecha_inicio_consulta_ta=''
lista_t_anterior=[]
resultado_t_anterior=[]
fecha_actual= ''
consulta_diaria2 = """
WITH CTE_DESCRIPCION AS (
    SELECT 
        DOCTO_IN_ID, 
        CAST(SUBSTRING(DESCRIPCION FROM POSITION(':' IN DESCRIPCION) + 2) AS INT) AS FOLIO
    FROM DOCTOS_IN
    WHERE 
        POSITION(':' IN DESCRIPCION) > 0
        AND NOT SUBSTRING(DESCRIPCION FROM POSITION(':' IN DESCRIPCION) + 2) SIMILAR TO '%[^0-9]%'
)
SELECT 
    MG_ORDENES_PROD.FOLIO, 
    MG_ORDENES_PROD.SOLICITANTE, 
    MG_ORDENES_PROD.OBSERVACIONES, 
    DOCTOS_IN_DET.UNIDADES 
FROM 
    DOCTOS_IN_DET 
INNER JOIN 
    DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
INNER JOIN 
    CTE_DESCRIPCION ON DOCTOS_IN.DOCTO_IN_ID = CTE_DESCRIPCION.DOCTO_IN_ID
INNER JOIN 
    MG_ORDENES_PROD ON CTE_DESCRIPCION.FOLIO = MG_ORDENES_PROD.FOLIO 
WHERE 
    DOCTOS_IN.FECHA_HORA_CREACION BETWEEN CAST(? AS TIMESTAMP) AND CAST(? AS TIMESTAMP)
    AND DOCTOS_IN.NATURALEZA_CONCEPTO = 'E';
"""



"""
SELECT 
    MG_ORDENES_PROD.FOLIO, 
    MG_ORDENES_PROD.SOLICITANTE, 
    MG_ORDENES_PROD.OBSERVACIONES, 
    DOCTOS_IN_DET.UNIDADES 
FROM 
    DOCTOS_IN_DET 
INNER JOIN 
    DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
INNER JOIN 
    MG_ORDENES_PROD ON 
    (
        SELECT CAST(SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 2) AS INT) 
        FROM DOCTOS_IN 
        WHERE 
            DOCTOS_IN.DOCTO_IN_ID = DOCTOS_IN_DET.DOCTO_IN_ID 
            AND POSITION(':' IN DOCTOS_IN.DESCRIPCION) > 0
            AND NOT SUBSTRING(DOCTOS_IN.DESCRIPCION FROM POSITION(':' IN DOCTOS_IN.DESCRIPCION) + 2) SIMILAR TO '%[^0-9]%'
    ) = MG_ORDENES_PROD.FOLIO 
WHERE 
    DOCTOS_IN.FECHA_HORA_CREACION BETWEEN CAST(? AS TIMESTAMP) AND CAST(? AS TIMESTAMP)
    AND DOCTOS_IN.NATURALEZA_CONCEPTO = 'E';
"""
opcion_activa = False

root = tk.Tk()
main_frame = tk.Frame(root)
main_frame.pack(expand=True, fill="both")

content_frame = tk.Frame(main_frame)

# Crear el control de pestañas
tab_control = ttk.Notebook(content_frame)
    # Crear las pestañas
tab1 = ttk.Frame(tab_control, style="MyFrame.TFrame")
tab2 = ttk.Frame(tab_control, style="MyFrame.TFrame")
tab3 = ttk.Frame(tab_control, style="MyFrame.TFrame")
# Crear la figura de matplotlib
fig = Figure(figsize=(6, 5), dpi=100)
#fig.bar()
ax = fig.add_subplot(111)

canvasc = Canvas(bg='white', master=tab2)
scrollbar = ttk.Scrollbar(tab2, orient="vertical", command=canvasc.yview)
scrollable_frame = ttk.Frame(canvasc)

# Usar grid para layout
canvasc.grid(row=0, column=0, sticky='nsew')
scrollbar.grid(row=0, column=1, sticky='ns')

canvas2c = Canvas(bg='white', master=scrollable_frame)
canvas2c.grid(row=0, column=0, sticky='nsew')

tab2.grid_rowconfigure(0, weight=1)
tab2.grid_columnconfigure(0, weight=1)

# Crear el canvas de matplotlib
canvas = FigureCanvasTkAgg(fig, master=scrollable_frame)
canvas.draw()
canvas.get_tk_widget().grid(row=0, column=4, sticky='nsew')

scrollable_frame.grid_rowconfigure(0, weight=1)
scrollable_frame.grid_columnconfigure(0, weight=1)

# Crear la figura de matplotlib
fig2 = Figure(figsize=(6, 5), dpi=100)
ax2 = fig2.add_subplot(111)

canvascc = Canvas(bg='white', master=tab3)
scrollbar2 = ttk.Scrollbar(tab3, orient="vertical", command=canvascc.yview)
scrollable_frame2 = ttk.Frame(canvascc)

# Usar grid para layout
canvascc.grid(row=0, column=0, sticky='nsew')
scrollbar2.grid(row=0, column=1, sticky='ns')

canvas2cc = Canvas(bg='white', master=scrollable_frame2)
canvas2cc.grid(row=0, column=0, sticky='nsew')

tab3.grid_rowconfigure(0, weight=1)
tab3.grid_columnconfigure(0, weight=1)

# Crear el canvas de matplotlib
canvas2 = FigureCanvasTkAgg(fig2, master=scrollable_frame2)
canvas2.draw()
canvas2.get_tk_widget().grid(row=0, column=4, sticky='nsew')

scrollable_frame2.grid_rowconfigure(0, weight=1)
scrollable_frame2.grid_columnconfigure(0, weight=1)

canvas_prod_ant = Canvas(bg='white', master=scrollable_frame)
canvas_prod_ant.grid(row=0, column=5, sticky='nsew')

canvas_prod_ant2 = Canvas(bg='white', master=scrollable_frame2)
canvas_prod_ant2.grid(row=0, column=5, sticky='nsew')


canvas_tab1 = Canvas(bg='white', master=tab1)
scrollbar2_tab1 = ttk.Scrollbar(tab1, orient="vertical", command=canvas_tab1.yview)
scrollable_frame_tab1 = ttk.Frame(canvas_tab1)

# Usar grid para layout
canvas_tab1.grid(row=0, column=0, sticky='nsew')
scrollbar2_tab1.grid(row=0, column=1, sticky='ns')

canvas2_tab1 = Canvas(bg='white', master=scrollable_frame_tab1)
canvas2_tab1.grid(row=0, column=0, sticky='nsew')

tab1.grid_rowconfigure(0, weight=1)
tab1.grid_columnconfigure(0, weight=1)

scrollable_frame_tab1.grid_rowconfigure(0, weight=1)
scrollable_frame_tab1.grid_columnconfigure(0, weight=1)


def cambiar_pestana():
    global opcion_activa
    if opcion_activa:
        current_tab = tab_control.index("current")
        next_tab = (current_tab + 1) % tab_control.index("end")  # Obtener el índice de la siguiente pestaña
        tab_control.select(next_tab)  # Cambiar a la siguiente pestaña
        root.after(6400, cambiar_pestana)  # Llamar a cambiar_pestana nuevamente después de 1 minuto (60000 milisegundos)

def toggle_opcion():
    global opcion_activa
    opcion_activa = not opcion_activa
    if opcion_activa:
        cambiar_pestana()

def turno(num):
     # Obtener la hora actual
    hora_actual = datetime.datetime.now().time()
    if num != 0:
        hora_actual =datetime.datetime.combine(datetime.date.today(), hora_actual) - datetime.timedelta(hours=8)

    # Verificar el turno actual
    if hora_actual.hour >= 6 and hora_actual.hour <= 13:
        return "Matutino"
    elif hora_actual.hour >= 14 and hora_actual.hour <= 22:
        return "Vespertino"
    else:
        return "Nocturno"    
    
def actualiza(usuario, contraseña):
    global user, pasw, hora_actual, turnoActu, turnoAnte, lista, fecha_inicio_consulta, fecha_fin_consulta, fecha_inicio_consulta_ta, lista_t_anterior, fecha_actual
    user = usuario
    pasw = contraseña
    Coneccion.lista.clear()  # Limpiar la lista antes de obtener nuevos datos
    
    turnoActu = turno(0)
    turnoAnte = turno(1)
    fecha_actual = datetime.datetime.now()
    fecha = datetime.date.today()
    
    if turnoActu == "Matutino":
        fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(6, 0, 0, 0))
        if fecha_actual.weekday() == 0:
            resta_dias = datetime.timedelta(days=2)
        else:
            resta_dias = datetime.timedelta(days=1)
        fecha = fecha - resta_dias
        fecha_inicio_consulta_ta = datetime.datetime.combine(fecha, datetime.time(22, 0, 0, 0))
    elif turnoActu == "Vespertino":
        fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(14, 0, 0, 0))
        fecha_inicio_consulta_ta = datetime.datetime.combine(fecha, datetime.time(6, 0, 0, 0))
    else:
        if datetime.time(0, 0) <= datetime.datetime.now().time() <= datetime.time(6, 0):
            un_dia = datetime.timedelta(days=1)
            fecha = fecha - un_dia
        fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(22, 0, 0, 0))
        fecha_inicio_consulta_ta = datetime.datetime.combine(fecha, datetime.time(14, 0, 0, 0))
    
    fecha_inicio_consulta = fecha_inicio_consulta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    fecha_fin_consulta = fecha_actual.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    fecha_inicio_consulta_ta = fecha_inicio_consulta_ta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    
    params = (fecha_inicio_consulta, fecha_fin_consulta)
    params2 = (fecha_inicio_consulta_ta, fecha_inicio_consulta)
    

    # Verificar si alguno de los parámetros es una cadena vacía
    if not all(params) or not all(params2):
        mb.showerror("Error!!", "Parámetros de fecha no válidos.")
        return
    
    # Realizar la conexión y ejecutar la consulta
    try:
        connect = Coneccion.coneccion(host='192.168.128.51', database='E:\\Sys\\Microsip datos\\COPAC 2024.FDB', user=usuario, password=contraseña, accion="consulta", consulta=consulta_diaria2, datos=params)
        if connect == "Conexión exitosa a Firebird!":
            if ban:  
                lista.clear()
                resultado.clear()
                lista = Coneccion.lista
                try:
                    connect2 = Coneccion.coneccion(host='192.168.128.51', database='E:\\Sys\\Microsip datos\\COPAC 2024.FDB', user=usuario, password=contraseña, accion="consulta", consulta=consulta_diaria2, datos=params2)
                    if connect2 == "Conexión exitosa a Firebird!":
                        lista_t_anterior.clear()
                        resultado_t_anterior.clear()
                        lista_t_anterior = Coneccion.lista 
                    else:
                        mb.showerror("¡¡Error al traer datos anteriores!!", connect2)  
                except Exception as e:
                    mb.showerror("¡¡Error al traer datos anteriores!!", f"Error general: {e}") 
                
                Pantalla()
        else:
            mb.showerror("Error!!", connect) 
    except Exception as e:
        mb.showerror("Error!!", f"Error general: {e}")
        
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
def tabla_timer():
    global resultado, resultado_t_anterior
    resultado=[]
    resultado_t_anterior=[]
    while True:
        
        
        #-------------------------------------------------------------------------------------------------------------------------------------------
        
        resultado =sumar_campos(lista)
        resultado_t_anterior =sumar_campos(lista_t_anterior)
        
        # Limpiar la tabla antes de insertar nuevos datos
        #tree.after(0, actualizar_tabla(tree))
        """if tree.get_children() !=():
            tree.delete(*tree.get_children())
        #-------------------------------------------------------------------------------------------------------------------------------------------
        # Actualizar los datos y agregarlos a la tabla
        for item in resultado:
            #producido = Decimal(item[3]) - Decimal(item[4])
            tree.insert('', tk.END, values=(item[0], item[1], item[2], '{:,.2f}'.format(round(float(item[3]), 2))))
                
        # Ajustar el ancho de las columnas según el contenido
        for col in tree["columns"]:
            width = max(font.Font().measure(col.title()), len(col) * 6)  # Ajusta el ancho mínimo de la columna
            tree.column(col, width=width)
"""
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvasc.configure(
                scrollregion=canvasc.bbox("all")
            )
        )
        scrollable_frame2.bind(
            "<Configure>",
            lambda e: canvascc.configure(
                scrollregion=canvascc.bbox("all")
            )
        )
        scrollable_frame_tab1.bind(
            "<Configure>",
            lambda e: canvas_tab1.configure(
                scrollregion=canvas_tab1.bbox("all")
            )
        )
        
        canvas_tab1.create_window((0, 0), window=scrollable_frame_tab1, anchor="nw")
        canvas_tab1.configure(yscrollcommand=scrollbar.set)
        
        # Añadir encabezados
        headers_tab1 = ["Folio", "Medida","Máquina", "Producción"]
        for j, header in enumerate(headers_tab1):
            header_label = ttk.Label(canvas2_tab1, text=header, font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
            header_label.grid(row=0, column=j, padx=1, pady=1, sticky="nsew")
        # Añadir datos a la tabla 
        total_rows_tab1 = len(resultado)

        for i in range(total_rows_tab1):
            # Columna de folio
            folio_label = ttk.Label(canvas2_tab1, background="#FFFFFF", text=f"{resultado[i][0]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            folio_label.grid(row=i+1, column=0, padx=1, pady=1, sticky="nsew")
            
            # Columna de nombre
            etiqueta_label = ttk.Label(canvas2_tab1, background="#FFFFFF", text=f"{resultado[i][1]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
                
            # Columna de etiquetas
            etiqueta_label = ttk.Label(canvas2_tab1, background="#FFFFFF", text=resultado[i][2], font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")

            # Columna de cantidades faltantes
            cantidad_label = ttk.Label(canvas2_tab1, background="#FFFFFF", text=resultado[i][3], font=('Helvetica', 9, 'bold'), borderwidth=1, relief="solid")
            cantidad_label.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")

        # Ajustar las columnas para que se expandan uniformemente
        for col in range(len(headers_tab1)):
            canvas2_tab1.grid_columnconfigure(col, weight=1)
        #-------------------------------------------------------------------------------------------------------------------------------------------
        folio = [str(item[0]) for item in resultado if "ext" not in (item[2] or "").lower()]
        nombre = [str(item[1]) for item in resultado if "ext" not in (item[2] or "").lower()]
        etiquetas = [item[2] if item[2] is not None else "" for item in resultado if "ext" not in (item[2] or "").lower()]
        cantidades_faltantes = [round(float(item[3]), 2) for item in resultado if "ext" not in (item[2] or "").lower()]
         #-------------------------------------------------------------------------------------------------------------------------------------------
        # Graficar los datos
        ax.clear()
        ax.bar(folio, cantidades_faltantes, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
            
        # Dibujar las líneas indicadoras de cantidad detrás de las barras
        for i, n in enumerate(cantidades_faltantes):
            ax.text(i, n + 5, (i+1), ha='center', va='bottom')

        ax.set_title("Produccion bolseo del Turno: " + turno(0))
        ax.set_xticklabels([])
        ax.set_facecolor(color="#EBFAF5")
        ax.set_yticklabels([])
        ax.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        # Actualizar la gráfica
        canvas.draw() 
        color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28']

        canvasc.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvasc.configure(yscrollcommand=scrollbar.set)
        
        
        #-------------------------------------------------------------------------------------------------------------------------------------------    
        # Añadir encabezados
        headers = ["Color","Folio", "Medida","Máquina", "Producción"]
        for j, header in enumerate(headers):
            header_label = ttk.Label(canvas2c, text=header, font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
            header_label.grid(row=0, column=j, padx=1, pady=1, sticky="nsew")
        # Añadir datos a la tabla con colores
        total_rows = len(etiquetas)
        total_bolseo=0
        total_extrusion=0
        total_bolseo_ant=0
        total_extrusion_ant=0
        
        for i in range(total_rows):
            # Columna de colores
            color_label = tk.Label(canvas2c, width=3, bg=color[i % len(color)],text=f"{i + 1}", borderwidth=1, relief="solid")
            color_label.grid(row=i+1, column=0, padx=1, pady=1, sticky="nsew")

            # Columna de folio
            folio_label = ttk.Label(canvas2c, background="#FFFFFF", text=f"{folio[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            folio_label.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
            
            # Columna de nombre
            etiqueta_label = ttk.Label(canvas2c, background="#FFFFFF", text=f"{nombre[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")
                
            # Columna de etiquetas
            etiqueta_label = ttk.Label(canvas2c, background="#FFFFFF", text=etiquetas[i], font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")

            # Columna de cantidades faltantes
            cantidad_label = ttk.Label(canvas2c, background="#FFFFFF", text=cantidades_faltantes[i], font=('Helvetica', 9, 'bold'), borderwidth=1, relief="solid")
            cantidad_label.grid(row=i+1, column=4, padx=1, pady=1, sticky="nsew")

            total_bolseo+= cantidades_faltantes[i]
        # Ajustar las columnas para que se expandan uniformemente
        for col in range(len(headers)):
            canvas2c.grid_columnconfigure(col, weight=1)
            
            
        label_total_bolseo = ttk.Label(scrollable_frame, text=f"PRODUCCIÓN TOTAL DE BOLSEO:\t {total_bolseo}", font=("Helvetica", 12,"bold underline"), anchor="center")
        label_total_bolseo .grid(row=1, column=0, padx=1, pady=1, sticky="nsew")
        label_total_bolseo .configure(background='#FFFFFF') 
        
        
        #-------------------------------------------------------------------------------------------------------------------------------------------    
        folio2 = [str(item[0])for item in resultado if "ext" in (item[2] or "").lower()]
        nombre2 = [str(item[1]) for item in resultado if "ext" in (item[2] or "").lower()]
        etiquetas2 = [item[2] if item[2] is not None else "" for item in resultado if "ext" in (item[2] or "").lower()]
        cantidades_faltantes2 = [round(float(item[3]), 2) for item in resultado if "ext" in (item[2] or "").lower()]
        
        #-------------------------------------------------------------------------------------------------------------------------------------------    
        # Graficar los datos
        ax2.clear()
        ax2.bar(folio2, cantidades_faltantes2, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
            
        # Dibujar las líneas indicadoras de cantidad detrás de las barras
        for i, n in enumerate(cantidades_faltantes2):
            ax2.text(i, n + 10, (i+1), ha='center', va='bottom')

        ax2.set_title("Produccion extrusion del Turno: " + turno(0))
        ax2.set_xticklabels([])
        ax2.set_facecolor(color="#EBFAF5")
        ax2.set_yticklabels([])
        ax2.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        # Actualizar la gráfica
        canvas2.draw()
            
        color2=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845']
        canvascc.create_window((0, 0), window=scrollable_frame2, anchor="nw")
        canvascc.configure(yscrollcommand=scrollbar.set)
        #-------------------------------------------------------------------------------------------------------------------------------------------    
        # Añadir encabezados
        headers = ["Color", "Folio", "Medida","Máquina", "Producción"]
        for j, header in enumerate(headers):
            header_label = ttk.Label(canvas2cc, text=header, font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
            header_label.grid(row=0, column=j, padx=1, pady=1, sticky="nsew")

        # Añadir datos a la tabla con colores
        total_rows = len(etiquetas2)

        for i in range(total_rows):
            # Columna de colores
            color_label = tk.Label(canvas2cc, width=3, bg=color2[i % len(color2)],text=f"{i + 1}", borderwidth=1, relief="solid")
            color_label.grid(row=i+1, column=0, padx=1, pady=1, sticky="nsew")

            # Columna de nombre
            folio_label = ttk.Label(canvas2cc, background="#FFFFFF", text=f"{folio2[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            folio_label.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
            # Columna de nombre
            etiqueta_label = ttk.Label(canvas2cc, background="#FFFFFF", text=f"{nombre2[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")
                
            # Columna de etiquetas
            etiqueta_label = ttk.Label(canvas2cc, background="#FFFFFF", text=etiquetas2[i], font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")

            # Columna de cantidades faltantes
            cantidad_label = ttk.Label(canvas2cc, background="#FFFFFF", text=cantidades_faltantes2[i], font=('Helvetica', 9, 'bold'), borderwidth=1, relief="solid")
            cantidad_label.grid(row=i+1, column=4, padx=1, pady=1, sticky="nsew")
            
            total_extrusion+= cantidades_faltantes2[i]

        # Ajustar las columnas para que se expandan uniformemente
        for col in range(len(headers)):
            canvas2cc.grid_columnconfigure(col, weight=1)
            
        label_total_extrusion = ttk.Label(scrollable_frame2, text=f"PRODUCCIÓN TOTAL DE EXTRUSIÓN:\t {total_extrusion}", font=("Helvetica", 12,"bold underline"), anchor="center")
        label_total_extrusion .grid(row=1, column=0, padx=1, pady=1, sticky="nsew")
        label_total_extrusion .configure(background='#FFFFFF') 
        #-------------------------------------------------------------------------------------------------------------------------------------------  
        
        folio_ant = [str(item[0])for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
        nombre_ant = [str(item[1]) for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
        etiquetas_ant = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
        cantidades_faltantes_ant = [round(float(item[3]), 2) for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
        #print("nombre_ant: " ,nombre_ant)
        # Añadir encabezados
        headers_ant = ["Folio", "Medida","Máquina", "Producción"]
        for j, header_ant in enumerate(headers_ant):
            header_label = ttk.Label(canvas_prod_ant, text=header_ant, font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
            header_label.grid(row=1, column=j, padx=1, pady=1, sticky="nsew")
        # Añadir datos a la tabla con colores
        total_rows_ant = len(etiquetas_ant)

        for i in range(total_rows_ant):
            # Columna de colores
            color_label1 = tk.Label(canvas_prod_ant, background="#FFFFFF", text=f"{folio_ant[i]}", borderwidth=1, relief="solid")
            color_label1.grid(row=i+2, column=0, padx=1, pady=1, sticky="nsew")

            # Columna de nombre
            etiqueta_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=f"{nombre_ant[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label1.grid(row=i+2, column=1, padx=1, pady=1, sticky="nsew")
                
            # Columna de etiquetas
            etiqueta_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=etiquetas_ant[i], font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label1.grid(row=i+2, column=2, padx=1, pady=1, sticky="nsew")

            # Columna de cantidades faltantes
            cantidad_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=cantidades_faltantes_ant[i], font=('Helvetica', 9, 'bold'), borderwidth=1, relief="solid")
            cantidad_label1.grid(row=i+2, column=3, padx=1, pady=1, sticky="nsew")
            
            total_bolseo_ant+= cantidades_faltantes_ant[i]

        # Ajustar las columnas para que se expandan uniformemente
        for col in range(len(headers_ant)):
            canvas_prod_ant.grid_columnconfigure(col, weight=1)
            
        label_total_bolseo_ant = ttk.Label(scrollable_frame, text=f"PRODUCCIÓN ANTERIOR: \t {total_bolseo_ant}", font=("Helvetica", 12,"bold underline"), anchor="center")
        label_total_bolseo_ant .grid(row=1, column=5, padx=1, pady=1, sticky="nsew")
        label_total_bolseo_ant .configure(background='#FFFFFF') 
        #-------------------------------------------------------------------------------------------------------------------------------------------     
        #-------------------------------------------------------------------------------------------------------------------------------------------  
        
        folio2_anterior  = [str(item[0]) for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
        nombre2_anterior  = [str(item[1]) for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
        etiquetas2_anterior  = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
        cantidades_faltantes2_anterior  = [round(float(item[3]), 2) for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
        #print("nombre_ant: " ,nombre_ant)
        # Añadir encabezados
        headers_ant = ["Folio", "Medida","Máquina", "Producción"]
        for j, header_ant in enumerate(headers_ant):
            header_label = ttk.Label(canvas_prod_ant2, text=header_ant, font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
            header_label.grid(row=1, column=j, padx=1, pady=1, sticky="nsew")
        # Añadir datos a la tabla con colores
        total_rows_ant2 = len(etiquetas2_anterior)

        for i in range(total_rows_ant2):
            # Columna de colores
            color_label1 = tk.Label(canvas_prod_ant2, background="#FFFFFF", text=f"{folio2_anterior[i]}", borderwidth=1, relief="solid")
            color_label1.grid(row=i+2, column=0, padx=1, pady=1, sticky="nsew")

            # Columna de nombre
            etiqueta_label1 = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=f"{nombre2_anterior[i]}", font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label1.grid(row=i+2, column=1, padx=1, pady=1, sticky="nsew")
                
            # Columna de etiquetas
            etiqueta_label1 = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=etiquetas2_anterior[i], font=('Helvetica', 9), borderwidth=1, relief="solid")
            etiqueta_label1.grid(row=i+2, column=2, padx=1, pady=1, sticky="nsew")

            # Columna de cantidades faltantes
            cantidad_label1 = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=cantidades_faltantes2_anterior[i], font=('Helvetica', 9, 'bold'), borderwidth=1, relief="solid")
            cantidad_label1.grid(row=i+2, column=3, padx=1, pady=1, sticky="nsew")
            
            total_extrusion_ant+= cantidades_faltantes2_anterior[i]

        # Ajustar las columnas para que se expandan uniformemente
        for col in range(len(headers_ant)):
            canvas_prod_ant2.grid_columnconfigure(col, weight=1)
            
        label_total_extrusion_ant = ttk.Label(scrollable_frame2, text=f"PRODUCCIÓN ANTERIOR: \t {total_extrusion_ant}", font=("Helvetica", 12,"bold underline"), anchor="center")
        label_total_extrusion_ant .grid(row=1, column=5, padx=1, pady=1, sticky="nsew")
        label_total_extrusion_ant .configure(background='#FFFFFF') 
        #-------------------------------------------------------------------------------------------------------------------------------------------    
        time.sleep(500)  # Esperar 3200 segundos (aproximadamente 53 minutos)
        #time.sleep(3600)  # Esperar 3600 segundos (aproximadamente 60 minutos)

def generar_pdf():
    global resultado, fecha_inicio_consulta, fecha_fin_consulta, fecha_inicio_consulta_ta
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    import matplotlib.pyplot as plt
    from tkinter import Tk, filedialog
    from tkinter import messagebox
    # Datos
    
    # Función para seleccionar la ubicación de guardado
    def get_save_location():
        root = Tk()
        root.withdraw()  # Ocultar la ventana principal de Tkinter
        save_location = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Guardar PDF como"
        )
        root.destroy()
        return save_location

    # Obtener la ubicación de guardado del archivo
    save_location = get_save_location()
    if not save_location:
        messagebox.showwarning("Alerta", "No se seleccionó ninguna ubicación de guardado. ¿Desea salir?")
        return
    # Crear el archivo PDF
    pdf = SimpleDocTemplate(save_location, pagesize=letter,
                            rightMargin=72, leftMargin=72, topMargin=20, bottomMargin=18)
    elements = []

    # Añadir la imagen
    logo = "logo_ecopac.png"
    im = Image(logo, 200, 100)
    elements.append(im)

    # Añadir un pequeño espacio después de la imagen
    elements.append(Spacer(1, 10))

    # Título de la tabla
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='SubTitle', fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name='NormalText', fontSize=12, leading=14, spaceAfter=5))
    
    title = Paragraph("Reporte de producción", styles['Title'])
    elements.append(title)

    turnox = Paragraph("Turno: " + turno(0), styles['SubTitle'])
    elements.append(turnox)

    fechax = Paragraph("Fecha de inicio del turno: " + fecha_inicio_consulta + "\n Ultima fecha del registro: " + fecha_fin_consulta, styles['NormalText'])
    elements.append(fechax)
    # Añadir un espacio después del título
    elements.append(Spacer(1, 20))

    # Añadir los datos a la tabla con los encabezados de las columnas
    data = [('Folio', 'Solicitante', 'Observaciones','KG producidos')] + resultado
        # Crear la tabla
    table = Table(data)

    # Añadir estilo a la tabla
    style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), '#009688'),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
    ('FONTSIZE', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    table.setStyle(style)
    elements.append(table)
    # Añadir un salto de página
    elements.append(PageBreak())
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------
    # Título de la gráfica
    graph_title = Paragraph("Produccion de bolseo ", styles['Title'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    
    # Crear la gráfica de barras
    folio = [str(item[0]) for item in resultado if "ext" not in (item[2] or "").lower()]
    nombre = [item[1] for item in resultado if "ext" not in (item[2] or "").lower()]
    etiquetas = [item[2] if item[2] is not None else "" for item in resultado if "ext" not in (item[2] or "").lower()]
    cantidades_faltantes = [round(float(item[3]), 2) for item in resultado if "ext" not in (item[2] or "").lower()]
    
    plt.figure(figsize=(6, 4))
    plt.bar(folio, cantidades_faltantes, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
    
    for i, n in enumerate(cantidades_faltantes):
        plt.text(i, n + 10, (i+1), ha='center', va='bottom')

    plt.title("Turno: " + turno(0))
    plt.xticks([])
    plt.yticks([])
    plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
    plt.savefig('grafica.png')
    plt.close()

    # Añadir la gráfica al PDF
    graph_image = Image('grafica.png', 400, 300)
    elements.append(graph_image)
    
    numeros = list(range(1, (len(nombre)+1)))
    combinacion = list(zip(numeros,folio, nombre, etiquetas, cantidades_faltantes))
    # Añadir los datos a la tabla con los encabezados de las columnas
    data2 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinacion
        # Crear la tabla
    table2 = Table(data2)
    table2.setStyle(style)
    elements.append(table2)
    
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------

    # Crear la gráfica de barras
    folio2 = [item[0] for item in resultado if "ext" in (item[2] or "").lower()]
    nombre2 = [item[1] for item in resultado if "ext" in (item[2] or "").lower()]
    etiquetas2 = [item[2] if item[2] is not None else "" for item in resultado if "ext" in (item[2] or "").lower()]
    cantidades_faltantes2 = [round(float(item[3]), 2) for item in resultado if "ext" in (item[2] or "").lower()]
    if nombre2!=[]:
        # Título de la gráfica
        # Añadir un salto de página
        elements.append(PageBreak())
        graph_title = Paragraph("Produccion de extrusión ", styles['Title'])
        elements.append(graph_title)
        elements.append(Spacer(1, 20))
        
        plt.figure(figsize=(6, 4))
        plt.bar(nombre2, cantidades_faltantes2, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        for i, n in enumerate(cantidades_faltantes2):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("Turno: " + turno(0))
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('grafica1.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_image1 = Image('grafica1.png', 400, 300)
        elements.append(graph_image1)
        
        numeros = list(range(1, (len(nombre2)+1)))
        combinacion2 = list(zip(numeros,folio2, nombre2, etiquetas2, cantidades_faltantes2))
        # Añadir los datos a la tabla con los encabezados de las columnas
        data2 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinacion2
            # Crear la tabla
        table3 = Table(data2)
        table3.setStyle(style)
        elements.append(table3)
    
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------
    # Título de la gráfica
    
    elements.append(PageBreak())
    graph_titlet = Paragraph("Produccion del turno anterior ", styles['Title'])
    elements.append(graph_titlet)
    elements.append(Spacer(1, 20))
    turnot = Paragraph("Turno: " + turno(1), styles['SubTitle'])
    elements.append(turnot)

    fechat = Paragraph("Fecha de inicio del turno: " + fecha_inicio_consulta_ta, styles['NormalText'])
    elements.append(fechat)
    fechat = Paragraph("Fecha de fin del turno: " + fecha_inicio_consulta, styles['NormalText'])
    elements.append(fechat)
    elements.append(Spacer(1, 20))
    graph_title = Paragraph("Produccion de bolseo ", styles['Title'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    # Crear la gráfica de barras
    nombret = [item[1] for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
    etiquetast = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]
    cantidades_faltantest = [round(float(item[3]), 2) for item in resultado_t_anterior if "ext" not in (item[2] or "").lower()]

    if nombret!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(nombret, cantidades_faltantest, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
        for i, n in enumerate(cantidades_faltantest):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("Turno: " + turno(1))
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficat.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_imaget = Image('graficat.png', 400, 300)
        elements.append(graph_imaget)
        
        numerost = list(range(1, (len(nombret)+1)))
        combinaciont = list(zip(numerost,nombret, etiquetast, cantidades_faltantest))
        # Añadir los datos a la tabla con los encabezados de las columnas
        datat = [('No.', 'Medida', 'Maquina', 'Produccion')] + combinaciont
            # Crear la tabla
        table_t = Table(datat)
        table_t.setStyle(style)
        elements.append(table_t)
    else:
        fechatt = Paragraph("No hay datos, parece que a ocurrido un error inesperado, intente mas tarde", styles['NormalText'])
        elements.append(fechatt)
    
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    # Crear la gráfica de barras
    nombre2t = [item[1] for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
    etiquetas2t = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
    cantidades_faltantes2t = [round(float(item[3]), 2) for item in resultado_t_anterior if "ext" in (item[2] or "").lower()]
    
        # Título de la gráfica
        # Añadir un salto de página
    elements.append(PageBreak())
    graph_title = Paragraph("Produccion de extrusión ", styles['Title'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    if nombre2t!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(nombre2t, cantidades_faltantes2t, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        for i, n in enumerate(cantidades_faltantes2t):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("Turno: " + turno(1))
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('grafica1t.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_image1t = Image('grafica1t.png', 400, 300)
        elements.append(graph_image1t)
        
        numerostt = list(range(1, (len(nombre2t)+1)))
        combinacion2t = list(zip(numerostt,nombre2t, etiquetas2t, cantidades_faltantes2t))
        # Añadir los datos a la tabla con los encabezados de las columnas
        data2tt = [('No.', 'Medida', 'Maquina', 'Produccion')] + combinacion2t
            # Crear la tabla
        table3t = Table(data2tt)
        table3t.setStyle(style)
        elements.append(table3t)
    else:
        fechatt = Paragraph("No hay datos, parece que a ocurrido un error inesperado o no tienen registrada ninguna entrada, intente mas tarde", styles['NormalText'])
        elements.append(fechatt)
    #-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    # Construir el archivo PDF
    pdf.build(elements)

    #print(f"PDF creado exitosamente en {save_location}.")
    messagebox.showinfo(message=f"PDF creado exitosamente en {save_location}.", title="PDF creado")

def generar_pdf1(tipo,fechas):
    global fecha_actual, user,pasw
    lista_pdf= []
    fecha_inicio_consulta =""
    fecha_fin_consulta=""
    consulta_diaria2 = """
    WITH CTE_DESCRIPCION AS (
        SELECT 
            DOCTO_IN_ID, 
            CAST(SUBSTRING(DESCRIPCION FROM POSITION(':' IN DESCRIPCION) + 2) AS INT) AS FOLIO
        FROM DOCTOS_IN
        WHERE 
            POSITION(':' IN DESCRIPCION) > 0
            AND NOT SUBSTRING(DESCRIPCION FROM POSITION(':' IN DESCRIPCION) + 2) SIMILAR TO '%[^0-9]%'
    )
    SELECT 
        MG_ORDENES_PROD.FOLIO, 
        MG_ORDENES_PROD.SOLICITANTE, 
        MG_ORDENES_PROD.OBSERVACIONES, 
        DOCTOS_IN_DET.UNIDADES,
        DOCTOS_IN.USUARIO_CREADOR
    FROM 
        DOCTOS_IN_DET 
    INNER JOIN 
        DOCTOS_IN ON DOCTOS_IN_DET.DOCTO_IN_ID = DOCTOS_IN.DOCTO_IN_ID 
    INNER JOIN 
        CTE_DESCRIPCION ON DOCTOS_IN.DOCTO_IN_ID = CTE_DESCRIPCION.DOCTO_IN_ID
    INNER JOIN 
        MG_ORDENES_PROD ON CTE_DESCRIPCION.FOLIO = MG_ORDENES_PROD.FOLIO 
    WHERE 
        DOCTOS_IN.FECHA_HORA_CREACION BETWEEN CAST(? AS TIMESTAMP) AND CAST(? AS TIMESTAMP)
        AND DOCTOS_IN.NATURALEZA_CONCEPTO = 'E';
    """
    if tipo == "diario":
        if fecha_actual.weekday() == 0:
            resta_dias = datetime.timedelta(days=2)
            resta_dias1 = datetime.timedelta(days=1)
            fecha_actual = fecha_actual - resta_dias1
        else:
            resta_dias = datetime.timedelta(days=1)
        fecha = fecha_actual - resta_dias
        
        fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(6, 0, 0, 0))
        fecha_fin_consulta = datetime.datetime.combine(fecha_actual, datetime.time(6, 0, 0, 0))
    if tipo == "especial":
        fecha_inicio_consulta = datetime.datetime.combine(fechas[0], datetime.time(6, 0, 0, 0))
        fecha_fin_consulta = datetime.datetime.combine(fechas[1], datetime.time(6, 0, 0, 0))
    fecha_inicio_consulta = fecha_inicio_consulta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    fecha_fin_consulta = fecha_fin_consulta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
     
    params = (fecha_inicio_consulta, fecha_fin_consulta)
    #params = ('2024-07-11 06:00:00.000', '2024-07-12 06:00:00.000')
    try:
        connect2 = Coneccion.coneccion(host='192.168.128.51', database='E:\\Sys\\Microsip datos\\COPAC 2024.FDB', user=user, password=pasw, accion="consulta", consulta=consulta_diaria2, datos=params)
        if connect2 == "Conexión exitosa a Firebird!":
            lista_pdf.clear()
            lista_pdf = Coneccion.lista 
        else:
            mb.showerror("¡¡Error al traer datos anteriores!!", connect2)  
    except Exception as e:
        mb.showerror("¡¡Error al traer datos anteriores!!", f"Error general: {e}") 
    
    supervisor_1 = []
    supervisor_2 = []
    supervisor_3 = []
    sin_numero = []

    for item in lista_pdf:
        supervisor = item[-1]
        if '1' in supervisor:
            supervisor_1.append(item)
        elif '2' in supervisor:
            supervisor_2.append(item)
        elif '3' in supervisor:
            supervisor_3.append(item)
        else:
            sin_numero.append(item)
    
    def sumar_campos1(lista_tuplas):
        acumulador = {}
        detalles = {}
        # Itera sobre cada tupla en la lista
        for folio,solicitante,observaciones,unidades,sup in lista_tuplas:
            # Si la clave ya está en el diccionario, suma el valor
            if folio in acumulador:
                acumulador[folio] += unidades
            # Si la clave no está en el diccionario, inicializa con el valor
            else:
                acumulador[folio] = unidades
            detalles[folio] = (folio,solicitante, observaciones,sup)
            # Convierte el diccionario acumulador de nuevo en una lista de tuplas
        resultado = [
            ( detalles[folio][0], detalles[folio][1],detalles[folio][2],detalles[folio][3],acumulador[folio]) 
            for folio in acumulador
        ]
        return resultado
    
    resultadoS1=[]
    resultadoS2=[]
    resultadoS3=[]
    sin_numero1=[]
    resultadoS1 =sumar_campos1(supervisor_1)
    resultadoS2 =sumar_campos1(supervisor_2)
    resultadoS3 =sumar_campos1(supervisor_3)
    sin_numero1 =sumar_campos1(sin_numero)
    
    
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    import matplotlib.pyplot as plt
    from tkinter import Tk, filedialog
    from tkinter import messagebox
    
    style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), '#009688'),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
    ('FONTSIZE', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    # Función para seleccionar la ubicación de guardado
    def get_save_location():
        root = Tk()
        root.withdraw()  # Ocultar la ventana principal de Tkinter
        save_location = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Guardar PDF como"
        )
        root.destroy()
        return save_location

    # Obtener la ubicación de guardado del archivo
    save_location = get_save_location()
    if not save_location:
        messagebox.showwarning("Alerta", "No se seleccionó ninguna ubicación de guardado. ¿Desea salir?")
        return
    # Crear el archivo PDF
    pdf = SimpleDocTemplate(save_location, pagesize=letter,
                            rightMargin=72, leftMargin=72, topMargin=20, bottomMargin=18)
    elements = []

    # Añadir la imagen
    logo = "logo_ecopac.png"
    im = Image(logo, 400, 200)
    elements.append(im)

    # Añadir un pequeño espacio después de la imagen
    elements.append(Spacer(1, 10))

    # Título de la tabla
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Title1', fontSize=20, leading=30, spaceAfter=20))
    styles.add(ParagraphStyle(name='SubTitle', fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name='NormalText', fontSize=12, leading=14, spaceAfter=5))
    
    title = Paragraph("Reporte diario de producción", styles['Title1'])
    elements.append(title)

    turnox = Paragraph("Fecha de inicio: " + fecha_inicio_consulta, styles['SubTitle'])
    elements.append(turnox)

    fechax = Paragraph("Fecha de fin: " + fecha_fin_consulta, styles['SubTitle'])
    elements.append(fechax)
    # Añadir un espacio después del título
    elements.append(Spacer(1, 20))
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    elements.append(PageBreak())
    graph_sup = Paragraph("SUPERVISOR 1", styles['Title'])
    elements.append(graph_sup)
    
    graph_title = Paragraph("Produccion de bolseo", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    
    # Crear la gráfica de barras
    #folio = [item[0] for item in resultadoS1 if "ext" not in (item[2] or "").lower()]
    folio = [str(item[0]) for item in resultadoS1 if "ext" not in (item[2] or "").lower()]
    medida = [item[1] for item in resultadoS1 if "ext" not in (item[2] or "").lower()]
    maquina = [item[2] if item[2] is not None else "" for item in resultadoS1 if "ext" not in (item[2] or "").lower()]
    produccion = [round(float(item[4]), 2) for item in resultadoS1 if "ext" not in (item[2] or "").lower()]

    if medida!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(folio, produccion, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
        for i, n in enumerate(produccion):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 1: Produccion de bolseo")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS1B.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_imaget = Image('graficaS1B.png', 400, 300)
        elements.append(graph_imaget)
        
        numerost = list(range(1, (len(medida)+1)))
        combinaciont = list(zip(numerost,folio, medida, maquina, produccion))
        # Añadir los datos a la tabla con los encabezados de las columnas
        datat = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinaciont
            # Crear la tabla
        table_t = Table(datat)
        table_t.setStyle(style)
        elements.append(table_t)
    else:
        fechatt = Paragraph("No hay datos, parece que el SUPERVISOR 1 no ha generado datos de bolseo aún", styles['NormalText'])
        elements.append(fechatt)
    
    elements.append(PageBreak())
    graph_sup = Paragraph("SUPERVISOR 1", styles['Title'])
    elements.append(graph_sup)
    graph_title = Paragraph("Produccion de extrusión", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    # Crear la gráfica de barras
    folioe = [str(item[0]) for item in resultadoS1 if "ext" in (item[2] or "").lower()]
    medida3e = [item[1] for item in resultadoS1 if "ext" in (item[2] or "").lower()]
    maquinae = [item[2] if item[2] is not None else "" for item in resultadoS1 if "ext" in (item[2] or "").lower()]
    produccione = [round(float(item[4]), 2) for item in resultadoS1 if "ext" in (item[2] or "").lower()]
    
    if medida3e!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(folioe, produccione, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        for i, n in enumerate(produccione):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 1: Produccion de extrusión")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS1E.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_image1t = Image('graficaS1E.png', 400, 300)
        elements.append(graph_image1t)
        
        numerostt = list(range(1, (len(medida3e)+1)))
        combinacion2t = list(zip(numerostt,folioe,medida3e, maquinae, produccione))
        # Añadir los datos a la tabla con los encabezados de las columnas
        data2tt = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinacion2t
            # Crear la tabla
        table3t = Table(data2tt)
        table3t.setStyle(style)
        elements.append(table3t)
    else:
        fechatt =  Paragraph("No hay datos, parece que el SUPERVISOR 1 no ha generado datos de extrusión aún", styles['NormalText'])
        elements.append(fechatt)
    
    elements.append(PageBreak())
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    graph_sup = Paragraph("SUPERVISOR 2", styles['Title'])
    elements.append(graph_sup)
    
    graph_title = Paragraph("Produccion de bolseo", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    
    # Crear la gráfica de barras
    folio2 = [str(item[0]) for item in resultadoS2 if "ext" not in (item[2] or "").lower()]
    medida2 = [item[1] for item in resultadoS2 if "ext" not in (item[2] or "").lower()]
    maquina2 = [item[2] if item[2] is not None else "" for item in resultadoS2 if "ext" not in (item[2] or "").lower()]
    produccion2 = [round(float(item[4]), 2) for item in resultadoS2 if "ext" not in (item[2] or "").lower()]

    if medida2!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(folio2, produccion2, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
        for i, n in enumerate(produccion2):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 2: Produccion de bolseo")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS2B.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_imaget2 = Image('graficaS2B.png', 400, 300)
        elements.append(graph_imaget2)
        
        numerost2 = list(range(1, (len(medida2)+1)))
        combinaciont2 = list(zip(numerost2,folio2, medida2, maquina2, produccion2))
        # Añadir los datos a la tabla con los encabezados de las columnas
        datat2 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinaciont2
            # Crear la tabla
        table_t2 = Table(datat2)
        table_t2.setStyle(style)
        elements.append(table_t2)
    else:
        fechatt = Paragraph("No hay datos, parece que el SUPERVISOR 2 no ha generado datos de bolseo aún", styles['NormalText'])
        elements.append(fechatt)
    
    elements.append(PageBreak())
    graph_sup = Paragraph("SUPERVISOR 2", styles['Title'])
    elements.append(graph_sup)
    graph_title = Paragraph("Produccion de extrusión", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    
    # Crear la gráfica de barras
    folioe2 = [str(item[0]) for item in resultadoS2 if "ext" in (item[2] or "").lower()]
    medida3e2 = [item[1] for item in resultadoS2 if "ext" in (item[2] or "").lower()]
    maquinae2 = [item[2] if item[2] is not None else "" for item in resultadoS2 if "ext" in (item[2] or "").lower()]
    produccione2 = [round(float(item[4]), 2) for item in resultadoS2 if "ext" in (item[2] or "").lower()]
    
    if medida3e2!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(folioe2, produccione2, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        for i, n in enumerate(produccione2):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 2: Produccion de extrusión")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS2E.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_image1t2 = Image('graficaS2E.png', 400, 300)
        elements.append(graph_image1t2)
        
        numerostt2 = list(range(1, (len(medida3e2)+1)))
        combinacion2t2 = list(zip(numerostt2,folioe2,medida3e2, maquinae2, produccione2))
        # Añadir los datos a la tabla con los encabezados de las columnas
        data2tt2 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinacion2t2
            # Crear la tabla
        table3t2 = Table(data2tt2)
        table3t2.setStyle(style)
        elements.append(table3t2)
    else:
        fechatt =  Paragraph("No hay datos, parece que el SUPERVISOR 2 no ha generado datos de extrusión aún", styles['NormalText'])
        elements.append(fechatt)
    
    elements.append(PageBreak())
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    graph_sup = Paragraph("SUPERVISOR 3", styles['Title'])
    elements.append(graph_sup)
    
    graph_title = Paragraph("Produccion de bolseo", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    
    # Crear la gráfica de barras
    folio3 = [str(item[0]) for item in resultadoS3 if "ext" not in (item[2] or "").lower()]
    medida3 = [item[1] for item in resultadoS3 if "ext" not in (item[2] or "").lower()]
    maquina3 = [item[2] if item[2] is not None else "" for item in resultadoS3 if "ext" not in (item[2] or "").lower()]
    produccion3 = [round(float(item[4]), 2) for item in resultadoS3 if "ext" not in (item[2] or "").lower()]

    if medida!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(folio3, produccion3, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
        for i, n in enumerate(produccion3):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 3: Produccion de bolseo")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS3B.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_imaget = Image('graficaS3B.png', 400, 300)
        elements.append(graph_imaget)
        
        numerost3 = list(range(1, (len(medida3)+1)))
        combinaciont3 = list(zip(numerost3,folio3, medida3, maquina3, produccion3))
        # Añadir los datos a la tabla con los encabezados de las columnas
        datat3 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinaciont3
            # Crear la tabla
        table_t3 = Table(datat3)
        table_t3.setStyle(style)
        elements.append(table_t3)
    else:
        fechatt = Paragraph("No hay datos, parece que el SUPERVISOR 3 no ha generado datos de bolseo aún", styles['NormalText'])
        elements.append(fechatt)
        
    elements.append(PageBreak())
    graph_sup = Paragraph("SUPERVISOR 3", styles['Title'])
    elements.append(graph_sup)
    graph_title = Paragraph("Produccion de extrusión", styles['SubTitle'])
    elements.append(graph_title)
    elements.append(Spacer(1, 20))
    # Crear la gráfica de barras
    folioe3 = [str(item[0]) for item in resultadoS3 if "ext" in (item[2] or "").lower()]
    medida3e3 = [item[1] for item in resultadoS3 if "ext" in (item[2] or "").lower()]
    maquinae3 = [item[2] if item[2] is not None else "" for item in resultadoS3 if "ext" in (item[2] or "").lower()]
    produccione3 = [round(float(item[4]), 2) for item in resultadoS3 if "ext" in (item[2] or "").lower()]
    

    if medida3e3!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(folioe3, produccione3, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        for i, n in enumerate(produccione3):
            plt.text(i, n + 10, (i+1), ha='center', va='bottom')

        plt.title("SUPERVISOR 3: Produccion de extrusión")
        plt.xticks([])
        plt.yticks([])
        plt.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.savefig('graficaS3E.png')
        plt.close()

        # Añadir la gráfica al PDF
        graph_image1t3 = Image('graficaS3E.png', 400, 300)
        elements.append(graph_image1t3)
        
        numerostt3 = list(range(1, (len(medida3e3)+1)))
        combinacion2t3 = list(zip(numerostt3,folioe3,medida3e3, maquinae3, produccione3))
        # Añadir los datos a la tabla con los encabezados de las columnas
        data2tt3 = [('No.','Folio', 'Medida', 'Maquina', 'Produccion')] + combinacion2t3
            # Crear la tabla
        table3t3 = Table(data2tt3)
        table3t3.setStyle(style)
        elements.append(table3t3)
    else:
        fechatt =  Paragraph("No hay datos, parece que el SUPERVISOR 3 no ha generado datos de extrusión aún", styles['NormalText'])
        elements.append(fechatt)
    elements.append(PageBreak())
    #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    extras =  Paragraph("Produccion de extra", styles['SubTitle'])
    elements.append(extras)
    # Añadir un espacio después del título
    elements.append(Spacer(1, 20))

    # Añadir los datos a la tabla con los encabezados de las columnas
    data = [('Folio', 'Medida', 'Maquina','Usuario','KG producidos')] + sin_numero1
        # Crear la tabla
    table = Table(data)

    table.setStyle(style)
    elements.append(table)

    suma_total1 = sum(produccion)
    suma_totalE1 = sum(produccione)
    suma_total2 = sum(produccion2)
    suma_totalE2 = sum(produccione2)
    suma_total3 = sum(produccion3)
    suma_totalE3 = sum(produccione3)
    
    elements.append(Spacer(1, 30))
    suma =  Paragraph("Producción total de bolseo por supervisor", styles['Title'])
    elements.append(suma)
    elements.append(Spacer(1, 20))
    suma1 =  Paragraph(f"SUPERVISOR 1: {suma_total1}", styles['SubTitle'])
    elements.append(suma1)
    suma2 =  Paragraph(f"SUPERVISOR 2: {suma_total2}", styles['SubTitle'])
    elements.append(suma2)
    suma3 =  Paragraph(f"SUPERVISOR 3: {suma_total3}", styles['SubTitle'])
    elements.append(suma3)
    # Añadir un espacio después del título
    elements.append(Spacer(1, 20))
    sumae =  Paragraph("Producción total de extrusión por supervisor", styles['Title'])
    elements.append(sumae)
    elements.append(Spacer(1, 20))
    suma4 =  Paragraph(f"SUPERVISOR 1: {suma_totalE1}", styles['SubTitle'])
    elements.append(suma4)
    suma5 =  Paragraph(f"SUPERVISOR 2: {suma_totalE2}", styles['SubTitle'])
    elements.append(suma5)
    suma6 =  Paragraph(f"SUPERVISOR 3: {suma_totalE3}", styles['SubTitle'])
    elements.append(suma6)
    
#--------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Construir el archivo PDF
    pdf.build(elements)

    #print(f"PDF creado exitosamente en {save_location}.")
    messagebox.showinfo(message=f"PDF creado exitosamente en {save_location}.", title="PDF creado")
    return save_location
   
def reporte():
    global root
    from reporte import mostrar_reporte
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal mientras se muestra el reporte
    resultados = mostrar_reporte(tk.Toplevel(root))
    root.destroy()  # Destruye la ventana principal después de cerrar el reporte
    print(resultados[0])
    
    destinatarios = ["LV_Ale@outlook.com", "operaciones@copacsa.com"]

    pdf_path = generar_pdf1("especial",resultados)
    if pdf_path:
        enviar_correo(pdf_path, destinatarios)
    
def enviar_correo(con_adjunto, destinatarios):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    from email.utils import formatdate

    remitente = "rechumanos@copacsa.com"
    password = "Alejandra2024*!"
    
    # Crear el objeto MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = "Reporte Diario de Producción"
    msg['Date'] = formatdate(localtime=True)

    # Adjuntar el cuerpo del mensaje
    cuerpo = "Adjunto encontrarás el reporte diario de producción."
    msg.attach(MIMEText(cuerpo, 'plain'))

    # Adjuntar el archivo PDF
    archivo_adjunto = con_adjunto
    with open(archivo_adjunto, "rb") as adjunto:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(adjunto.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename="{archivo_adjunto}"')
        msg.attach(parte)

    # Enviar el correo
    try:
        servidor = smtplib.SMTP('mail.copacsa.com', 587)
        servidor.starttls()
        servidor.login(remitente, password)
        servidor.send_message(msg)
        servidor.quit()
        print("Correo enviado exitosamente")
    except Exception as e:
        print(f"No se pudo enviar el correo. Error: {e}")  
def c_email():
    global root
    from enviar_email import mostrar_configuracion_email
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal mientras se muestra el reporte
    mostrar_configuracion_email(tk.Toplevel(root))
    root.destroy()  # Destruye la ventana principal después de cerrar el reporte     
    
def Pantalla():
    global ban
    # Crear ventana
    root.title('ECOPAC, Bienvenido')
    root.iconbitmap('icono.ico')
    ban=False
    an = root.winfo_screenwidth()-2
    al = root.winfo_screenheight()-80

    # Establecer las dimensiones de la ventana
    root.geometry("%dx%d+0+0" % (an, al)) 
    root.resizable(0, 1)  
    root.config(bg='silver', width='1220', height='200')

    # Se crea el menú de la ventana
    menu = tk.Menu()
    menu_Producción = tk.Menu(menu, tearoff=0)
    menu_Configuración = tk.Menu(menu, tearoff=0)

    # Agregar las opciones principales al menú
    menu.add_cascade(label="Producción", menu=menu_Producción)
    menu.add_cascade(label="Configuración", menu=menu_Configuración)
    menu.add_cascade(label="Salir", command=root.quit)

    menu_Producción.add_command(label="Reporte actual", command=lambda: generar_pdf())
    menu_Producción.add_command(label="Reporte diario", command=lambda: generar_pdf1("diario",None))
    menu_Producción.add_command(label="Reporte especial", command=lambda: reporte())


    cambiar_var = tk.BooleanVar()
    cambiar_var.set(False)
    menu_Configuración.add_command(label="Configuración de e-mail", command=lambda: c_email())
    menu_Configuración.add_checkbutton(label="Cambiar pestaña automáticamente", variable=cambiar_var)
    menu_Configuración.entryconfig("Cambiar pestaña automáticamente", command=toggle_opcion)

    # Se muestra la barra de menú en la ventana principal
    root.config(menu=menu)
    root.configure(bg='#F2F7F7')

    # Crear estilo para las pestañas
    style = ttk.Style()
    # Configurar color de fondo para el Frame
    style.configure('MyFrame.TFrame', background='#F2F7F7')

    # Frame para el título
    title_frame = tk.Frame(main_frame, bg='#F2F7F7')
    title_frame.pack(side="top", fill="x")

    # Imagen
    image = tk.PhotoImage(file="logo_ecopac.png")
    image = image.subsample(5)  # Ajustar la imagen a la mitad del tamaño original

    label_image = ttk.Label(title_frame, image=image)
    label_image.pack(side="left", padx=10, pady=4)
    label_image.configure(background='#F2F7F7') 

    # Etiqueta de turno
    label_turno = ttk.Label(title_frame, text=" Turno: "+turno(0), font=("Helvetica", 30, "bold underline"), anchor="center")
    label_turno.pack(side="left", padx=10, pady=4)
    label_turno.configure(background='#F2F7F7') 

    # Etiqueta de título
    label_monitor = ttk.Label(title_frame, text="Monitor de producción\t", font=("Helvetica", 50), anchor="center")
    label_monitor.pack(side="right", padx=10, pady=4)
    label_monitor.configure(background='#F2F7F7') 
    content_frame.pack(expand=True, fill="both")
    tab_control.pack(expand=True, fill="both")
    tab_control.add(tab1, text='Producción')
    tab_control.add(tab2, text='Bolseo')
    tab_control.add(tab3, text='Extrusión')

    # Contenido de la pestaña 1
    """tree_frame = tk.Frame(tab1, borderwidth=2, relief="sunken")
    tree_frame.place(relx=0.01, rely=0.02, relwidth=0.6, relheight=0.96)
    tree = ttk.Treeview(tree_frame, show='headings', selectmode='browse', style="Treeview")
    tree.pack(side='left', fill='both', expand=True)
    scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.pack(side='right', fill='y')
    """

    ax.bar([], [],align='center', color=['blue','red','green','white','brown'], edgecolor='none')
    canvas.draw()
    
    # Contenido de la pestaña 2
    label_turno_anterior = ttk.Label(canvas_prod_ant, text="  Turno: "+turno(1)+"   \n", font=("Helvetica", 20,"bold underline"), anchor="center")
    label_turno_anterior .grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
    label_turno_anterior .configure(background='#FFFFFF') 
     # Contenido de la pestaña 3
    label_turno_anterior2 = ttk.Label(canvas_prod_ant2, text="  Turno: "+turno(1)+"   \n", font=("Helvetica", 20,"bold underline"), anchor="center")
    label_turno_anterior2 .grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
    label_turno_anterior2 .configure(background='#FFFFFF')   
    
    # Graficar los datos
    ax2.bar([], [])
    canvas2.draw()
 
    # Configurar el scrollbar para que controle el Treeview
    """tree.configure(yscrollcommand=scrollbar.set)
    style = ttk.Style()
    style.theme_use("default")
    tree.config(columns=("Folio", "Solicitante", "Observaciones", "Kg producidos"), show="headings")
    for heading in ("Folio", "Solicitante", "Observaciones", "Kg producidos"):
        tree.heading(heading, text=heading)"""
    t = threading.Thread(target=tabla_timer, args=())
    t.daemon = True
    t.start()
    root.mainloop()
    
    
Pantalla.is_alive = threading.Event()  # Utilizar un evento para verificar si la pantalla está en ejecución
