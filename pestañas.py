#Desarrollado por Alejandra López Vega
#Para Comercializadora de Plásticos de Alta Calidad S.A. de C.V.
#Version 1.0 lanzada en Octubre de 2025
import datetime
from decimal import Decimal
from email.mime import image
import itertools
import json
import os
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
import schedule
from resource_path import resource_path as rp
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
#LABELS
label_total_bolseo_ant = None
label_total_extrusion_ant= None
label_total_extrusion= None
label_total_bolseo= None
label_turno = None
label_turno_anterior = None
label_turno_anterior2 = None
folio_label_1= None 
etiqueta_label_1= None
etiqueta_label_2= None
cantidad_label_1= None
color_label1= None
cantidad_label1= None
etiqueta_label1= None
color_label= None
folio_label= None
medida_label= None
cantidad_label= None
label_actua_hora  = None

color_label_ext= None
folio_label_ext= None
medida_label_ext= None
maquina_label_ext= None
cantidad_label_ext= None
nombre_label1= None
color_label2= None
nombre_label2= None
cantidad_label2= None

et_label= None
label_dife=None  
label_difb=None 
label_c_bola=None  
label_c_bol=None  
label_c_exta=None 
label_c_ext=None 
label_ant=None  
label_act=None  

labels_creados = []
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
    ARTICULOS.NOMBRE,
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
INNER JOIN 
    ARTICULOS ON DOCTOS_IN_DET.ARTICULO_ID = ARTICULOS.ARTICULO_ID 
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


opcion_activa = True

root = tk.Tk()
main_frame = tk.Frame(root)
main_frame.pack(expand=True, fill="both")

# Frame para el título
title_frame = tk.Frame(main_frame, bg='#F2F7F7')
title_frame.pack(side="top", fill="x")

content_frame = tk.Frame(main_frame)
content_frame.pack(expand=True, fill="both")

# Crear el control de pestañas
tab_control = ttk.Notebook(content_frame)
tab_control.pack(expand=True, fill="both")
    # Crear las pestañas
tab1 = ttk.Frame(tab_control, style="MyFrame.TFrame")
tab2 = ttk.Frame(tab_control, style="MyFrame.TFrame")
tab3 = ttk.Frame(tab_control, style="MyFrame.TFrame")
tab4 = ttk.Frame(tab_control, style="MyFrame.TFrame")
# Crear la figura de matplotlib
fig = Figure(figsize=(6, 5), dpi=110)
#fig.bar()
ax = fig.add_subplot(111)
# Canvas para el gráfico en la segunda pestaña (tab2)

canvasc = Canvas( master=tab2)
canvasc.pack(fill="both", expand=True)
scrollbar = ttk.Scrollbar(canvasc, orient="vertical", command=canvasc.yview)
scrollbar.pack(side="right", fill="y")
scrollable_frame = ttk.Frame(canvasc)
scrollable_frame.pack(fill="both", expand=True)

canvas_frame = tk.Frame(master=scrollable_frame)
canvas_frame.pack(side="left", fill="both", expand=True, padx=150, pady=30)

canvas2c = Canvas( master=canvas_frame)
canvas2c.pack(side="top", fill="both", expand=True)

canvas_prod_ant = Canvas( master=canvas_frame)
canvas_prod_ant.pack(side="left", fill="both", expand=True)
#canvas_prod_ant.grid(row=1, column=0, sticky='nsew')

canvas_frame1 = tk.Frame(master=scrollable_frame)
canvas_frame1.pack(side="right", fill="both", expand=True)

canvas = FigureCanvasTkAgg(fig, master=canvas_frame1)
canvas.draw()
canvas.get_tk_widget().pack(side="top", fill="both", expand=True, pady=30)

totalprod = Canvas( master=canvas_frame1)
totalprod.pack(side="left", fill="both", expand=True, pady=30)


# Crear la figura de matplotlib
fig2 = Figure(figsize=(6, 5), dpi=110)
ax2 = fig2.add_subplot(111)
canvascc = Canvas(master=tab3)
canvascc.pack(fill="both", expand=True)
scrollbar2 = ttk.Scrollbar(canvascc, orient="vertical", command=canvascc.yview)
scrollbar2.pack(side="right", fill="y")
scrollable_frame2 = ttk.Frame(canvascc)
scrollable_frame2.pack(fill="both", expand=True)


canvas_framec = tk.Frame(scrollable_frame2)
canvas_framec.pack(side="left", fill="both", expand=True, padx=150, pady=30)

canvas2cc = Canvas( master=canvas_framec)
canvas2cc.pack(side="top", fill="both", expand=True)

canvas_prod_ant2 = Canvas(master=canvas_framec)
canvas_prod_ant2.pack(side="left", fill="both", expand=True)

canvas_frame1c = tk.Frame(scrollable_frame2)
canvas_frame1c.pack(side="right", fill="both", expand=True)

# Crear el canvas de matplotlib
canvas2 = FigureCanvasTkAgg(fig2, master=canvas_frame1c)
canvas2.draw()
canvas2.get_tk_widget().pack(side="top", fill="both", expand=True, pady=30)

totalprod2 = Canvas(master=canvas_frame1c)
totalprod2.pack(side="left", fill="both", expand=True, pady=30)


canvas_tab1 = Canvas(master=tab1)
canvas_tab1.grid(row=0, column=0, sticky='nsew')


scrollbar2_tab1 = ttk.Scrollbar(tab1, orient="vertical", command=canvas_tab1.yview)
scrollbar2_tab1.grid(row=0, column=1, sticky='ns')

scrollable_frame_tab1 = ttk.Frame(canvas_tab1)


canvas2_tab1 = Canvas(master=scrollable_frame_tab1)
canvas2_tab1.pack(side="left", fill="both", expand=True,padx=150, pady=30)

canvas2_tab2 = Canvas(master=scrollable_frame_tab1)
canvas2_tab2.pack(side="right", fill="both", expand=True,padx=0, pady=30)


tab1.grid_rowconfigure(0, weight=1)
tab1.grid_columnconfigure(0, weight=1)

scrollable_frame_tab1.grid_rowconfigure(0, weight=1)
scrollable_frame_tab1.grid_columnconfigure(0, weight=1)



canvas_tab4 = Canvas(master=tab4)
canvas_tab4.pack(fill="both", expand=True)
scrollbar_tab4 = ttk.Scrollbar(canvas_tab4, orient="vertical", command=canvas_tab4.yview)
scrollbar_tab4.pack(side="right", fill="y")
scrollable_frame_tab4 = ttk.Frame(canvas_tab4)
scrollable_frame_tab4.pack(fill="both", expand=True)

canvas_frame4 = tk.Frame(master=scrollable_frame_tab4)
canvas_frame4.pack(side="left", fill="both", expand=True, padx=150, pady=30)

canvas_hora = Canvas( master=canvas_frame4)
canvas_hora.pack(side="top", fill="both", expand=True)
def cambiar_pestana():
    global opcion_activa
    if opcion_activa:
        current_tab = tab_control.index("current")
        next_tab = (current_tab + 1) % tab_control.index("end")  # Obtener el índice de la siguiente pestaña
        tab_control.select(next_tab)  # Cambiar a la siguiente pestaña
        root.after(60000, cambiar_pestana)  # Llamar a cambiar_pestana nuevamente después de 1 minuto (60000 milisegundos)

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

    if hora_actual.hour >= 6 and hora_actual.hour <= 13:
        return "Matutino"
    elif hora_actual.hour >= 14 and hora_actual.hour <= 22: #14-22
        return "Vespertino"
    else:
        return "Nocturno"    
    
def actualiza(usuario, contraseña):
    global user, pasw, hora_actual, turnoActu, turnoAnte, lista, fecha_inicio_consulta, fecha_fin_consulta, fecha_inicio_consulta_ta, lista_t_anterior, fecha_actual
    user = usuario
    pasw = contraseña
    Coneccion.lista.clear()  # Limpiar la lista antes de obtener nuevos datos
    
    turnoActu = turno(0)
    
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
            lista.clear()
            resultado.clear()
            lista = Coneccion.lista
            #print(lista)  
            #print (F"params {params}...params2 {params2}")
            try:
                connect2 = Coneccion.coneccion(host='192.168.128.51', database='E:\\Sys\\Microsip datos\\COPAC 2024.FDB', user=usuario, password=contraseña, accion="consulta", consulta=consulta_diaria2, datos=params2)
                if connect2 == "Conexión exitosa a Firebird!":
                    lista_t_anterior.clear()
                    resultado_t_anterior.clear()
                    lista_t_anterior = Coneccion.lista 
                    #print(lista_t_anterior)
                else:
                    mb.showerror("¡¡Error al traer datos anteriores!!", connect2)  
            except Exception as e:
                mb.showerror("¡¡Error al traer datos anteriores!!", f"Error general: {e}") 
            turnoAnte = turnoActu    
            if ban:    
                #print("consulta de inicio...")
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
        if solicitante  in acumulador:
            acumulador[solicitante ] += unidades
        # Si la clave no está en el diccionario, inicializa con el valor
        else:
            acumulador[solicitante ] = unidades
        detalles[solicitante ] = (folio,solicitante, observaciones)
        # Convierte el diccionario acumulador de nuevo en una lista de tuplas
    resultado = [
        ( detalles[solicitante ][0], detalles[solicitante ][1],detalles[solicitante ][2],acumulador[solicitante ]) 
        for solicitante  in acumulador
    ]
    return resultado

def actualizar_base_datos():
    global resultado, resultado_t_anterior, user, pasw, ban, lista, lista_t_anterior, label_total_bolseo_ant,label_total_extrusion_ant,label_total_extrusion,label_total_bolseo,label_turno,label_turno_anterior, et_label
    global label_turno_anterior2,folio_label_1, etiqueta_label_1,etiqueta_label_2,cantidad_label_1,color_label1,cantidad_label1,etiqueta_label1,color_label, folio_label, medida_label,cantidad_label, label_actua_hora, labels_creados
    global label_dife,label_difb,label_c_bola,label_c_bol,label_c_exta,label_c_ext,label_ant,label_act, color_label_ext, folio_label_ext,medida_label_ext,maquina_label_ext,cantidad_label_ext,nombre_label1,color_label2,nombre_label2,cantidad_label2
    
    from PIL import Image, ImageTk
    from PIL import Image, ImageTk
    from datetime import datetime

    resultado=[]
    resultado_t_anterior=[]
    if not ban: 
        #print("actualiza consulta...")
        actualiza(user,pasw)
    
    if label_turno is not None:
        label_turno.destroy()   
    # Etiqueta de turno
    label_turno = ttk.Label(title_frame, text=" Turno: "+turno(0), font=("Helvetica", 30, "bold underline"), anchor="center")
    label_turno.pack(side="left", padx=10, pady=4)
    label_turno.configure(background='#F2F7F7') 
    
    if label_turno_anterior is not None:
        label_turno_anterior.destroy() 
    # Contenido de la pestaña 2
    label_turno_anterior = ttk.Label(canvas_prod_ant, text="  Turno: "+turno(1)+"   \n", font=("Helvetica", 20,"bold underline"), anchor="center")
    label_turno_anterior .grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
    #label_turno_anterior .configure(background='#FFFFFF') 
    # Contenido de la pestaña 3
    if label_turno_anterior2 is not None:
        label_turno_anterior2.destroy() 
    label_turno_anterior2 = ttk.Label(canvas_prod_ant2, text="  Turno: "+turno(1)+"   \n", font=("Helvetica", 20,"bold underline"), anchor="center")
    label_turno_anterior2 .grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
    #label_turno_anterior2 .configure(background='#FFFFFF')   
    
    resultado =sumar_campos(lista)
    resultado_t_anterior =sumar_campos(lista_t_anterior)
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
        header_label1 = ttk.Label(canvas2_tab1, text=header, font=('Helvetica', 12, 'bold'), borderwidth=1, relief="solid")
        header_label1.grid(row=0, column=j, padx=0, pady=0, sticky="nsew")
    # Añadir datos a la tabla 
    
    for label in labels_creados:
        label.destroy()
    labels_creados = []
  
    total_rows_tab1 = len(resultado)
    
    for i in range(total_rows_tab1):
        # Columna de folio
        folio_label_1 = ttk.Label(canvas2_tab1, background="#FFFFFF", text=f"{resultado[i][0]}", font=('Helvetica', 10), borderwidth=1, relief="solid")
        folio_label_1.grid(row=i+1, column=0, padx=1, pady=1, sticky="nsew")
        labels_creados.append(folio_label_1)    
        # Columna de nombre
        etiqueta_label_1 = ttk.Label(canvas2_tab1, background="#FFFFFF", text=f"{resultado[i][1]}", font=('Helvetica', 10), borderwidth=1, relief="solid")
        etiqueta_label_1.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
        labels_creados.append(etiqueta_label_1)          
        # Columna de etiquetas
        etiqueta_label_2 = ttk.Label(canvas2_tab1, background="#FFFFFF", text=resultado[i][2], font=('Helvetica', 10), borderwidth=1, relief="solid")
        etiqueta_label_2.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")
        labels_creados.append(etiqueta_label_2)  
        # Columna de cantidades
        cantidad_label_1 = ttk.Label(canvas2_tab1, background="#FFFFFF", text=resultado[i][3], font=('Helvetica', 10, 'bold'), borderwidth=1, relief="solid")
        cantidad_label_1.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")
        labels_creados.append(cantidad_label_1)  

    # Ajustar las columnas para que se expandan uniformemente
    for col in range(len(headers_tab1)):
        canvas2_tab1.grid_columnconfigure(col, weight=1)
    #-------------------------------------------------------------------------------------------------------------------------------------------
    folio = [str(item[0]) for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    nombre = [str(item[1]) for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    etiquetas = [item[2] if item[2] is not None else "" for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    cantidades_faltantes = [round(float(item[3]), 2) for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    #-------------------------------------------------------------------------------------------------------------------------------------------
    # Graficar los datos
    ax.clear()
    ax.bar(nombre, cantidades_faltantes, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
            
    # Dibujar las líneas indicadoras de cantidad detrás de las barras
    for i, n in enumerate(cantidades_faltantes):
        ax.text(i, n + 5, (i+1), ha='center', va='bottom')

    ax.set_title("Producción bolseo del Turno: " + turno(0))
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
        header_label2 = ttk.Label(canvas2c, text=header, font=('Helvetica', 14, 'bold'), borderwidth=1, relief="solid")
        header_label2.grid(row=0, column=j, padx=1, pady=1, sticky="nsew")
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
        labels_creados.append(color_label) 
        # Columna de folio
        folio_label = ttk.Label(canvas2c, background="#FFFFFF", text=f"{folio[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        folio_label.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
        labels_creados.append(folio_label)      
        # Columna de medida
        medida_label = ttk.Label(canvas2c, background="#FFFFFF", text=f"{nombre[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        medida_label.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")
        labels_creados.append(medida_label)         
        # Columna de maquina
        maquina_label = ttk.Label(canvas2c, background="#FFFFFF", text=etiquetas[i], font=('Helvetica', 12), borderwidth=1, relief="solid")
        maquina_label.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")
        labels_creados.append(maquina_label) 
        # Columna de cantidades 
        cantidad_label = ttk.Label(canvas2c, background="#FFFFFF", text=cantidades_faltantes[i], font=('Helvetica', 12, 'bold'), borderwidth=1, relief="solid")
        cantidad_label.grid(row=i+1, column=4, padx=1, pady=1, sticky="nsew")
        labels_creados.append(cantidad_label) 
        
        total_bolseo+= cantidades_faltantes[i]
    # Ajustar las columnas para que se expandan uniformemente
    for col in range(len(headers)):
        canvas2c.grid_columnconfigure(col, weight=2)
            
    if label_total_bolseo is not None:
        label_total_bolseo.destroy()          
    label_total_bolseo = ttk.Label(totalprod, text=f"Total de bolseo {turno(0)}:   \t\t{round(total_bolseo, 2):,.2f}", font=("Helvetica", 15,"bold underline"), anchor="center")
    label_total_bolseo .grid(row=0, column=0, padx=1, pady=1, sticky="e")
    #label_total_bolseo .configure(background='#FFFFFF') 
        
        
    #-------------------------------------------------------------------------------------------------------------------------------------------    
    folio2 = [str(item[0])for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    nombre2 = [str(item[1]) for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    etiquetas2 = [item[2] if item[2] is not None else "" for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    cantidades_faltantes2 = [round(float(item[3]), 2) for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
        
    #-------------------------------------------------------------------------------------------------------------------------------------------    
    # Graficar los datos
    ax2.clear()
    ax2.bar(nombre2, cantidades_faltantes2, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
            
    # Dibujar las líneas indicadoras de cantidad detrás de las barras
    for i, n in enumerate(cantidades_faltantes2):
        ax2.text(i, n + 10, (i+1), ha='center', va='bottom')

    ax2.set_title("Producción extrusión del Turno: " + turno(0))
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
    for j, header in enumerate(headers):
        header_label3 = ttk.Label(canvas2cc, text=header, font=('Helvetica', 14, 'bold'), borderwidth=1, relief="solid")
        header_label3.grid(row=0, column=j, padx=1, pady=1, sticky="nsew")

    # Añadir datos a la tabla con colores
    total_rows = len(etiquetas2)
    
    for i in range(total_rows):
        # Columna de colores
        color_label_ext = tk.Label(canvas2cc, width=3, bg=color2[i % len(color2)],text=f"{i + 1}", borderwidth=1, relief="solid")
        color_label_ext.grid(row=i+1, column=0, padx=1, pady=1, sticky="nsew")
        labels_creados.append(color_label_ext) 
        # Columna de folio
        folio_label_ext = ttk.Label(canvas2cc, background="#FFFFFF", text=f"{folio2[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        folio_label_ext.grid(row=i+1, column=1, padx=1, pady=1, sticky="nsew")
        labels_creados.append(folio_label_ext) 
        # Columna de medida
        medida_label_ext = ttk.Label(canvas2cc, background="#FFFFFF", text=f"{nombre2[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        medida_label_ext.grid(row=i+1, column=2, padx=1, pady=1, sticky="nsew")
        labels_creados.append(medida_label_ext)         
        # Columna de maquina
        maquina_label_ext = ttk.Label(canvas2cc, background="#FFFFFF", text=etiquetas2[i], font=('Helvetica', 12), borderwidth=1, relief="solid")
        maquina_label_ext.grid(row=i+1, column=3, padx=1, pady=1, sticky="nsew")
        labels_creados.append(maquina_label_ext) 
        # Columna de cantidades faltantes
        cantidad_label_ext = ttk.Label(canvas2cc, background="#FFFFFF", text=cantidades_faltantes2[i], font=('Helvetica', 12, 'bold'), borderwidth=1, relief="solid")
        cantidad_label_ext.grid(row=i+1, column=4, padx=1, pady=1, sticky="nsew")
        labels_creados.append(cantidad_label_ext) 
            
        total_extrusion+= cantidades_faltantes2[i]

    # Ajustar las columnas para que se expandan uniformemente
    for col in range(len(headers)):
        canvas2cc.grid_columnconfigure(col, weight=1)
    if label_total_extrusion is not None:
        label_total_extrusion.destroy()         
    label_total_extrusion = ttk.Label(totalprod2, text=f"Total de extrusión {turno(0)}:\t {round(total_extrusion, 2):,.2f}", font=("Helvetica", 15,"bold underline"), anchor="center")
    label_total_extrusion .grid(row=0, column=0, padx=1, pady=1, sticky="e")
    #label_total_extrusion .configure(background='#FFFFFF') 
    #-------------------------------------------------------------------------------------------------------------------------------------------  
        
    folio_ant = [str(item[0])for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    nombre_ant = [str(item[1]) for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    etiquetas_ant = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    cantidades_faltantes_ant = [round(float(item[3]), 2) for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    #print("nombre_ant: " ,nombre_ant)
    # Añadir encabezados
    for j, header_ant in enumerate(headers_tab1):
        header_label4 = ttk.Label(canvas_prod_ant, text=header_ant, font=('Helvetica', 14, 'bold'), borderwidth=1, relief="solid")
        header_label4.grid(row=1, column=j, padx=1, pady=1, sticky="nsew")
    # Añadir datos a la tabla con colores
    total_rows_ant = len(etiquetas_ant)
  
    for i in range(total_rows_ant):
        # Columna de colores
        color_label1 = tk.Label(canvas_prod_ant, background="#FFFFFF", text=f"{folio_ant[i]}", borderwidth=1, relief="solid")
        color_label1.grid(row=i+2, column=0, padx=1, pady=1, sticky="nsew")
        labels_creados.append(color_label1)
        # Columna de nombre
        nombre_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=f"{nombre_ant[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        nombre_label1.grid(row=i+2, column=1, padx=1, pady=1, sticky="nsew")
        labels_creados.append(nombre_label1)        
        # Columna de etiquetas
        etiqueta_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=etiquetas_ant[i], font=('Helvetica', 12), borderwidth=1, relief="solid")
        etiqueta_label1.grid(row=i+2, column=2, padx=1, pady=1, sticky="nsew")
        labels_creados.append(etiqueta_label1)
        # Columna de cantidades faltantes
        cantidad_label1 = ttk.Label(canvas_prod_ant, background="#FFFFFF", text=cantidades_faltantes_ant[i], font=('Helvetica', 12, 'bold'), borderwidth=1, relief="solid")
        cantidad_label1.grid(row=i+2, column=3, padx=1, pady=1, sticky="nsew")
        labels_creados.append(cantidad_label1)    
        total_bolseo_ant+= cantidades_faltantes_ant[i]

    # Ajustar las columnas para que se expandan uniformemente
    for col in range(len(headers_tab1)):
        canvas_prod_ant.grid_columnconfigure(col, weight=1)
     
    if label_total_bolseo_ant is not None:
        label_total_bolseo_ant.destroy()
    label_total_bolseo_ant = ttk.Label(totalprod, text=f"Total de bolseo {turno(1)}: \t {round(total_bolseo_ant, 2):,.2f}", font=("Helvetica", 15,"bold underline"), anchor="center")
    label_total_bolseo_ant .grid(row=1, column=0, padx=1, pady=1, sticky="e")
    #label_total_bolseo_ant .configure(background='#FFFFFF') 
    #-------------------------------------------------------------------------------------------------------------------------------------------     
    #-------------------------------------------------------------------------------------------------------------------------------------------  
        
    folio2_anterior  = [str(item[0]) for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    nombre2_anterior  = [str(item[1]) for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    etiquetas2_anterior  = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    cantidades_faltantes2_anterior  = [round(float(item[3]), 2) for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    #print("nombre_ant: " ,nombre_ant)
    # Añadir encabezados    
    for j, header_ant in enumerate(headers_tab1):
        header_label5 = ttk.Label(canvas_prod_ant2, text=header_ant, font=('Helvetica', 14, 'bold'), borderwidth=1, relief="solid")
        header_label5.grid(row=1, column=j, padx=1, pady=1, sticky="nsew")
    # Añadir datos a la tabla con colores
    total_rows_ant2 = len(etiquetas2_anterior)
  
    for i in range(total_rows_ant2):
        # Columna de colores
        color_label2 = tk.Label(canvas_prod_ant2, background="#FFFFFF", text=f"{folio2_anterior[i]}", borderwidth=1, relief="solid")
        color_label2.grid(row=i+2, column=0, padx=1, pady=1, sticky="nsew")
        labels_creados.append(color_label2)
        # Columna de nombre
        nombre_label2 = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=f"{nombre2_anterior[i]}", font=('Helvetica', 12), borderwidth=1, relief="solid")
        nombre_label2.grid(row=i+2, column=1, padx=1, pady=1, sticky="nsew")
        labels_creados.append(nombre_label2)        
        # Columna de etiquetas
        et_label = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=etiquetas2_anterior[i], font=('Helvetica', 12), borderwidth=1, relief="solid")
        et_label.grid(row=i+2, column=2, padx=1, pady=1, sticky="nsew")
        labels_creados.append(et_label)
        # Columna de cantidades faltantes
        cantidad_label2 = ttk.Label(canvas_prod_ant2, background="#FFFFFF", text=cantidades_faltantes2_anterior[i], font=('Helvetica', 12, 'bold'), borderwidth=1, relief="solid")
        cantidad_label2.grid(row=i+2, column=3, padx=1, pady=1, sticky="nsew")
        labels_creados.append(cantidad_label2)    
        total_extrusion_ant+= cantidades_faltantes2_anterior[i]

    # Ajustar las columnas para que se expandan uniformemente
    for col in range(len(headers_tab1)):
        canvas_prod_ant2.grid_columnconfigure(col, weight=1)
    if label_total_extrusion_ant is not None:
        label_total_extrusion_ant.destroy()        
    label_total_extrusion_ant = ttk.Label(totalprod2, text=f"Total de extrusión {turno(1)}: \t {round(total_extrusion_ant, 2):,.2f}", font=("Helvetica", 15,"bold underline"), anchor="center")
    label_total_extrusion_ant .grid(row=1, column=0, padx=1, pady=1, sticky="e")
    #label_total_extrusion_ant .configure(background='#FFFFFF') 
    
    if label_act is not None:
        label_act.destroy()  
    if label_ant is not None:
        label_ant.destroy()  
    if label_c_ext is not None:
        label_c_ext.destroy() 
    if label_c_exta is not None:
        label_c_exta.destroy()   
    if label_c_bol is not None:
        label_c_bol.destroy()  
    if label_c_bola is not None:
        label_c_bola.destroy()      
    if label_dife is not None:
        label_dife.destroy()     
    if label_difb is not None:
        label_difb.destroy()   
       
    
    label_ext = ttk.Label(canvas2_tab2,background="#FFFFFF", text=f"EXTRUSIÓN", font=("Helvetica", 18,"bold"), borderwidth=1, relief="solid")
    label_ext.grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
  
    label_bol = tk.Label(canvas2_tab2,background="#FFFFFF",text=f"BOLSEO", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_bol.grid(row=0, column=2, padx=1, pady=1, sticky="nsew")

    label_act = ttk.Label(canvas2_tab2, text=turno(0), font=('Helvetica', 14, 'bold'))
    label_act.grid(row=1, column=0, padx=1, pady=1, sticky="nsew")      
    labels_creados.append(label_act)  
      
    label_ant = ttk.Label(canvas2_tab2, text=turno(1), font=('Helvetica', 14, 'bold'))
    label_ant.grid(row=2, column=0, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_ant)  
    
    label_dif = ttk.Label(canvas2_tab2, text=f"Diferencia", font=('Helvetica', 14, 'bold'))
    label_dif.grid(row=3, column=0, padx=1, pady=1, sticky="nsew")  
    
    label_c_ext = ttk.Label(canvas2_tab2, text=f" {round(total_extrusion, 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_c_ext.grid(row=1, column=1, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_c_ext)  
    
    label_c_exta = ttk.Label(canvas2_tab2, text=f" {round(total_extrusion_ant, 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_c_exta.grid(row=2, column=1, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_c_exta)  
    
    label_c_bol = ttk.Label(canvas2_tab2, text=f" {round(total_bolseo, 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_c_bol.grid(row=1, column=2, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_c_bol)  
    
    label_c_bola = ttk.Label(canvas2_tab2, text=f" {round(total_bolseo_ant, 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_c_bola.grid(row=2, column=2, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_c_bola)  
    
    label_dife = ttk.Label(canvas2_tab2, text=f" {round((total_extrusion-total_extrusion_ant ), 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_dife.grid(row=3, column=1, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_dife)  
    
    label_difb = ttk.Label(canvas2_tab2, text=f" {round((total_bolseo-total_bolseo_ant), 2):,.2f}", font=('Helvetica', 18, 'bold'), borderwidth=1, relief="solid")
    label_difb.grid(row=3, column=2, padx=1, pady=1, sticky="nsew")
    labels_creados.append(label_difb)  
    
    
    if (total_extrusion_ant - total_extrusion) <= 0:
        # Abrir la imagen correspondiente si la condición es verdadera
        
        image = Image.open(rp("recursos/xsuperadaext.jpg"))
    else:
        # Verificar si la hora actual es 13:30 PM
        now = datetime.now()
        if (now.hour == 5 or now.hour == 13 or now.hour == 21) and now.minute >= 30:
            image = Image.open(rp("recursos/xnosuperadaext.jpg"))
        else:
            image = Image.open(rp("recursos/xextvamos.jpg"))

    image = image.resize((int(image.width / 3), int(image.height /3)))
    photo = ImageTk.PhotoImage(image)

    label = tk.Label(canvas2_tab2, image=photo, bg="#F2F7F7", borderwidth=1, relief="solid")
    label.image = photo  # Mantener una referencia
    label.grid(row=4, column=1, padx=1, pady=1, sticky="nsew")

    if (total_bolseo_ant-total_bolseo) <= 0:
        # Abrir la imagen correspondiente si la condición es verdadera
        image2 = Image.open(rp("recursos/xsuperada.jpg"))
    else:
        # Verificar si la hora actual es 13:30 PM
        now = datetime.now()
        if (now.hour == 5 or now.hour == 13 or now.hour == 21) and now.minute >= 30:
            image2 = Image.open(rp("recursos/xnosuperada.jpg"))
        else:
            image2 = Image.open(rp("recursos/xvamos.jpg"))

    image2 = image2.resize((int(image2.width / 3), int(image2.height /3)))
    photo2 = ImageTk.PhotoImage(image2)

    label2 = tk.Label(canvas2_tab2, image=photo2, bg="#F2F7F7", borderwidth=1, relief="solid")
    label2.image2 = photo2  # Mantener una referencia
    label2.grid(row=4, column=2, padx=1, pady=1, sticky="nsew")

    #print(total_bolseo,total_extrusion,total_bolseo_ant,total_extrusion_ant)
    ban=False
    #root.after(1800000, actualizar_base_datos)
    
def generar_pdf(locacion):
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
    if locacion:
        save_location = get_save_location()
        if not save_location:
            messagebox.showwarning("Alerta", "No se seleccionó ninguna ubicación de guardado. ¿Desea salir?")
            return
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Especifica el nombre del archivo que deseas guardar en el mismo directorio
        file_name = "produccion.pdf"

        # Construye la ruta completa donde se guardará el archivo
        save_location = os.path.join(current_dir, file_name)

        # Verifica si el directorio existe (aunque esto no debería fallar si estás usando el directorio actual)
        if not os.path.exists(current_dir):
            messagebox.showwarning("Alerta", "La ubicación de guardado no existe. ¿Desea salir?")
            
    # Crear el archivo PDF
    pdf = SimpleDocTemplate(save_location, pagesize=letter,
                            rightMargin=72, leftMargin=72, topMargin=20, bottomMargin=18)
    elements = []

    # Añadir la imagen
    logo = rp("recursos/logo_ecopac.png")
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
    folio = [str(item[0]) for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    nombre = [item[1] for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    etiquetas = [item[2] if item[2] is not None else "" for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    cantidades_faltantes = [round(float(item[3]), 2) for item in resultado if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    
    plt.figure(figsize=(6, 4))
    plt.bar(nombre, cantidades_faltantes, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
    
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
    folio2 = [item[0] for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    nombre2 = [item[1] for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    etiquetas2 = [item[2] if item[2] is not None else "" for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    cantidades_faltantes2 = [round(float(item[3]), 2) for item in resultado if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
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
    nombret = [item[1] for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    etiquetast = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    cantidades_faltantest = [round(float(item[3]), 2) for item in resultado_t_anterior if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    print
    if nombret!=[]:  
        plt.figure(figsize=(6, 4))
        color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28']
        colores_repetidos = list(itertools.islice(itertools.cycle(color), len(nombret)))
        plt.bar(nombret, cantidades_faltantest, align='center',color=colores_repetidos , edgecolor='none')
        
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
    nombre2t = [item[1] for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    etiquetas2t = [item[2] if item[2] is not None else "" for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    cantidades_faltantes2t = [round(float(item[3]), 2) for item in resultado_t_anterior if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    
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
    if locacion:
        messagebox.showinfo(message=f"PDF creado exitosamente en {save_location}.", title="PDF creado")
    return save_location

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
        ARTICULOS.NOMBRE,
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
    INNER JOIN 
        ARTICULOS ON DOCTOS_IN_DET.ARTICULO_ID = ARTICULOS.ARTICULO_ID 
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
            #print(lista_pdf)
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
            if solicitante in acumulador:
                acumulador[solicitante] += unidades
            # Si la clave no está en el diccionario, inicializa con el valor
            else:
                acumulador[solicitante] = unidades
            detalles[solicitante] = (folio,solicitante, observaciones,sup)
            # Convierte el diccionario acumulador de nuevo en una lista de tuplas
        resultado = [
            ( detalles[solicitante][0], detalles[solicitante][1],detalles[solicitante][2],detalles[solicitante][3],acumulador[solicitante]) 
            for solicitante in acumulador
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
    
    #print(supervisor_2)
    
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
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),  # encabezado
        ('FONTSIZE', (0, 1), (-1, -1), 9),  # resto de la tabla
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
    logo = rp("recursos/logo_ecopac.png")
    im = Image(logo, 400, 200)
    elements.append(im)

    # Añadir un pequeño espacio después de la imagen
    elements.append(Spacer(1, 10))

    # Título de la tabla
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Title1', fontSize=20, leading=30, spaceAfter=20))
    styles.add(ParagraphStyle(name='SubTitle', fontSize=14, leading=16, spaceAfter=10))
    styles.add(ParagraphStyle(name='NormalText', fontSize=10, leading=14, spaceAfter=5))
    
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
    folio = [str(item[0]) for item in resultadoS1 if "bobina"  not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    medida = [item[1] for item in resultadoS1 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    maquina = [item[2] if item[2] is not None else "" for item in resultadoS1 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    produccion = [round(float(item[4]), 2) for item in resultadoS1 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]

    if medida!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(medida, produccion, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
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
    folioe = [str(item[0]) for item in resultadoS1 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    medida3e = [item[1] for item in resultadoS1 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    maquinae = [item[2] if item[2] is not None else "" for item in resultadoS1 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    produccione = [round(float(item[4]), 2) for item in resultadoS1 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    
    if medida3e!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(medida3e, produccione, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
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
    folio2 = [str(item[0]) for item in resultadoS2 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    medida2 = [item[1] for item in resultadoS2 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    maquina2 = [item[2] if item[2] is not None else "" for item in resultadoS2 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    produccion2 = [round(float(item[4]), 2) for item in resultadoS2 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]

    if medida2!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(medida2, produccion2, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
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
    folioe2 = [str(item[0]) for item in resultadoS2 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    medida3e2 = [item[1] for item in resultadoS2 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    maquinae2 = [item[2] if item[2] is not None else "" for item in resultadoS2 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    produccione2 = [round(float(item[4]), 2) for item in resultadoS2 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    
    if medida3e2!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(medida3e2, produccione2, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
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
    folio3 = [str(item[0]) for item in resultadoS3 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    medida3 = [item[1] for item in resultadoS3 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    maquina3 = [item[2] if item[2] is not None else "" for item in resultadoS3 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]
    produccion3 = [round(float(item[4]), 2) for item in resultadoS3 if "bobina" not in (item[1] or "").lower() and "lamina" not in (item[1] or "").lower()]

    if medida!=[]:  
        plt.figure(figsize=(6, 4))
        plt.bar(medida3, produccion3, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
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
    folioe3 = [str(item[0]) for item in resultadoS3 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    medida3e3 = [item[1] for item in resultadoS3 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    maquinae3 = [item[2] if item[2] is not None else "" for item in resultadoS3 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    produccione3 = [round(float(item[4]), 2) for item in resultadoS3 if "bobina" in (item[1] or "").lower() or "lamina" in (item[1] or "").lower()]
    

    if medida3e3!=[]:    
        plt.figure(figsize=(6, 4))
        plt.bar(medida3e3, produccione3, align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
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
    if sin_numero1!=[]: 
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
    root.configure(bg='#F2F7F7')
    root.resizable(False, False)
    root.title('Reporte especial de producción')
    root.iconbitmap(rp('recursos/icono.ico'))
    resultados = mostrar_reporte(tk.Toplevel(root))
    root.destroy()  
    
    try:
        import sys
    
        # Detecta si está ejecutándose desde el exe o desde Python
        if getattr(sys, 'frozen', False):
            # Carpeta donde está el .exe
            carpeta_base = os.path.dirname(sys.executable)
        else:
            # Carpeta donde está el script .py
            carpeta_base = os.path.dirname(os.path.abspath(__file__))

        # Ruta completa del archivo JSON
        ruta_json = os.path.join(carpeta_base, "recipients.json")
        with open(ruta_json, 'r') as f:
            destinatarios = json.load(f)  # Cargar la lista de destinatarios
            
        pdf_path = generar_pdf1("especial",resultados)
        if pdf_path:
            enviar_correo(pdf_path, destinatarios,"Reporte Especial de Producción","Adjunto encontrarás el reporte Especial de producción.")

    except FileNotFoundError:
        print("El archivo de destinatarios no se encontró.")
    except json.JSONDecodeError as e:
        print(f"Error al cargar destinatarios: {e}")
         
    
def enviar_correo(con_adjunto, destinatarios,titulo,cuerpo):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    from email.utils import formatdate

    remitente = "monitor@copacsa.com"
    password = "1Mqazxsw2!"
    
    # Crear el objeto MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = titulo #"Reporte Diario de Producción"
    msg['Date'] = formatdate(localtime=True)

    # Adjuntar el cuerpo del mensaje
    #"Adjunto encontrarás el reporte diario de producción."
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
    import re
    import sys
    
        # Detecta si está ejecutándose desde el exe o desde Python
    if getattr(sys, 'frozen', False):
        # Carpeta donde está el .exe
        carpeta_base = os.path.dirname(sys.executable)
    else:
        # Carpeta donde está el script .py
        carpeta_base = os.path.dirname(os.path.abspath(__file__))

    # Ruta completa del archivo JSON
    ruta_json = os.path.join(carpeta_base, "recipients.json")
    
    def validar_destinatarios():
        contenido = entry_destinatarios.get("1.0", "end-1c")
        if not contenido.strip():
            mb.showerror("Error", "El campo de destinatarios no puede estar vacío.")
            root.lift()
            return False
            
        destinatarios = [email.strip() for email in contenido.split(',') if email.strip()]
        for email in destinatarios:
            if not validar_email(email):
                mb.showerror("Error", f"'{email}' no es un correo electrónico válido.")
                root.lift()
                return False
                    
        return True

    def validar_email(email):
        # Expresión regular para validar el formato de un correo electrónico
        regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(regex, email) is not None
        
    def save_recipients():
        if validar_destinatarios():
            recipients = [email.strip() for email in entry_destinatarios.get("1.0", "end-1c").split(',') if email.strip()]
            with open(ruta_json, 'w') as f:
                json.dump(recipients, f)  # Always overwrite with a fresh array
            label_status.config(text="Destinatarios guardados correctamente.", fg="green")
            if mb.askyesno("", "Destinatarios guardados correctamente. \n¿Desea salir?"):
                root.destroy()

    def load_recipients():
        try:
            with open(ruta_json, 'r') as f:
                content = f.read()
                # Print the content for debugging
                #print("Loaded content:", content)  
                recipients = json.loads(content)
                #print(content)
            entry_destinatarios.delete('1.0', tk.END)  # Clear the field before loading
            entry_destinatarios.insert('1.0', ', '.join(recipients))
        except FileNotFoundError:
            pass
        except json.JSONDecodeError as e:
            mb.showerror("Error", f"Error al cargar destinatarios: {e}")

    root = tk.Tk()
    root.title("Gestión de Destinatarios")
    root.geometry("600x450")  # Tamaño inicial de la ventana
    root.configure(bg="#F2F7F7")
    root.iconbitmap(rp('recursos/icono.ico'))
    root.resizable(False, False)
   
    # Frame para el título e imagen
    frame_top = tk.Frame(root, bg="#F2F7F7", pady=20)
    frame_top.pack(fill=tk.X)

    label_title = tk.Label(frame_top, text="Gestión de Destinatarios", font=("Arial", 16, "bold"), bg="#F2F7F7")
    label_title.pack(side=tk.TOP)

    # Frame para el contenido principal
    frame_content = tk.Frame(root, bg="#F2F7F7")
    frame_content.pack(padx=10, pady=1, fill=tk.BOTH, expand=True)

    canvas_frame = tk.Frame(frame_content, bg='#F2F7F7')
    canvas_frame.pack(fill="both", expand=True)
    # Mensaje para los destinatarios

    entry_destinatarios = tk.Text(canvas_frame, font=("Helvetica", 12), width=50)
    entry_destinatarios.config(width=50, height=10, font=("Consolas", 12), 
                                padx=15, pady=15, selectbackground="red")
    entry_destinatarios.pack(pady=10)
    
    label_info = tk.Label(canvas_frame, text="Por favor, separa cada destinatario con una coma.  \t  \t   ", font=("Arial", 12), bg="#F2F7F7", anchor="w")
    label_info.pack(pady=0,side=tk.TOP)
    # Botón para guardar
    save_button = tk.Button(frame_content, text="Guardar Destinatarios", command=save_recipients, font=("Arial", 12), bg="#009688", fg="white", bd=0, padx=10, pady=5, relief="raised")
    save_button.pack(pady=10)

    # Etiqueta de estado
    label_status = tk.Label(frame_content, text="", font=("Arial", 12), bg="#F2F7F7")
    label_status.pack(pady=10)
    # Cargar destinatarios si existen
    load_recipients()

    root.mainloop()

def Pantalla():
    from datetime import datetime
    # Crear ventana
    root.title('ECOPAC, Bienvenido')
    root.iconbitmap(rp('recursos/icono.ico'))
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
    menu.add_cascade(label="Reporte", menu=menu_Producción)
    menu.add_cascade(label="Configuración", menu=menu_Configuración)
    menu.add_cascade(label="Salir", command=cerrar_app)

    menu_Producción.add_command(label="Actual", command=lambda: generar_pdf(True))
    menu_Producción.add_command(label="Diario", command=lambda: generar_pdf1("diario",None))
    menu_Producción.add_command(label="Especial", command=lambda: reporte())


    cambiar_var = tk.BooleanVar()
    cambiar_var.set(True)
    menu_Configuración.add_command(label="E-mail", command=lambda: c_email())
    menu_Configuración.add_checkbutton(label="Pestaña automática", variable=cambiar_var)
    menu_Configuración.entryconfig("Pestaña automática", command=toggle_opcion)
    root.after(3200, cambiar_pestana)
    # Se muestra la barra de menú en la ventana principal
    root.config(menu=menu)
    root.configure(bg='#F2F7F7')

    # Crear estilo para las pestañas
    style = ttk.Style()
    # Configurar color de fondo para el Frame
    style.configure('MyFrame.TFrame', background='#F2F7F7')

    # Imagen
    image = tk.PhotoImage(file=rp("recursos/logo_ecopac.png"))
    image = image.subsample(5)  # Ajustar la imagen a la mitad del tamaño original

    label_image = ttk.Label(title_frame, image=image)
    label_image.pack(side="left", padx=10, pady=4)
    label_image.configure(background='#F2F7F7') 

    # Etiqueta de título
    label_monitor = ttk.Label(title_frame, text="Monitor de producción\t", font=("Helvetica", 50), anchor="center")
    label_monitor.pack(side="right", padx=10, pady=4)
    label_monitor.configure(background='#F2F7F7') 
    content_frame.pack(expand=True, fill="both")
    tab_control.pack(expand=True, fill="both")
    tab_control.add(tab1, text='Producción')
    tab_control.add(tab2, text='Bolseo')
    tab_control.add(tab3, text='Extrusión')
    tab_control.add(tab4, text='Hora')
    
    ax.bar([], [],align='center', color=['blue','red','green','white','brown'], edgecolor='none')
    canvas.draw()
    
    # Graficar los datos
    ax2.bar([], [])
    canvas2.draw()

    # Programar el envío de correos a horas específicas
    schedule.every().day.at("05:58").do(enviar_correox)  # 6 AM
    schedule.every().day.at("13:58").do(enviar_correox)  # 12 PM
    schedule.every().day.at("21:58").do(enviar_correox)  # 10 PM
    

    # Actualizaciones matutino
    schedule.every().day.at("06:00").do(actualizar_base_datos)  #
    schedule.every().day.at("07:00").do(actualizar_base_datos)  #
    schedule.every().day.at("08:00").do(actualizar_base_datos)  #
    schedule.every().day.at("09:00").do(actualizar_base_datos)  #
    schedule.every().day.at("10:00").do(actualizar_base_datos)  #
    schedule.every().day.at("11:00").do(actualizar_base_datos)  #
    schedule.every().day.at("12:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("13:00").do(actualizar_base_datos)  #
    schedule.every().day.at("13:55").do(actualizar_base_datos)  #
    
    # Actualizaciones vespertino
    schedule.every().day.at("14:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("15:00").do(actualizar_base_datos)  #
    schedule.every().day.at("16:00").do(actualizar_base_datos)  #
    schedule.every().day.at("17:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("18:00").do(actualizar_base_datos)  #
    schedule.every().day.at("19:00").do(actualizar_base_datos)  #
    schedule.every().day.at("20:00").do(actualizar_base_datos)  
    schedule.every().day.at("21:00").do(actualizar_base_datos)  #
    schedule.every().day.at("21:55").do(actualizar_base_datos)  # 
    
     # Actualizaciones nocturno
    schedule.every().day.at("22:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("23:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("00:00").do(actualizar_base_datos)  #
    schedule.every().day.at("01:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("02:00").do(actualizar_base_datos)  #
    schedule.every().day.at("03:00").do(actualizar_base_datos)  #
    schedule.every().day.at("04:00").do(actualizar_base_datos)  #
    schedule.every().day.at("05:00").do(actualizar_base_datos)  # 
    schedule.every().day.at("05:55").do(actualizar_base_datos)  #
    
    # Programar envío el primer día de cada semana
    #schedule.every().day.at("05:50").do(enviar_correo_especial("semanal"))
    # Programar envío el primer día de cada mes
    #schedule.every().day.at("08:02").do(enviar_correo_especial("mensual"))
    # Crear el hilo para correr el schedule en paralelo
    t = threading.Thread(target=ejecutar_schedule)
    t.daemon = True
    t.start()
  

    def current_date_format(date):
        months = ("Enero", "Febrero", "Marzo", "Abri", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
        day = date.day
        month = months[date.month - 1]
        year = date.year
        messsage = "{} de {} del {}".format(day, month, year)

        return messsage

    now = datetime.now()
    label_fecha = tk.Label(canvas_hora, text=current_date_format(now),font=("Arial", 80))
    label_fecha.grid(row=0, column=0, pady=20)
    label_hora = tk.Label(canvas_hora, font=("Arial", 290))
    label_hora.grid(row=1, column=0, pady=20) 

    def actualizar_hora():
        global label_actua_hora
        if not label_hora.winfo_exists():
            return  # si ya no existe, salimos
        hora_actual = time.strftime("%H:%M:%S")
        label_hora.config(text=hora_actual)
        if label_actua_hora is not None:
            label_actua_hora.destroy() 
    
        def actualizacion():
            if turno(0)=="Matutino":
                messsage = " 06:00, 07:00, 08:00, 09:00, 10:00, 11:00, 12:00, 13:00 y 13:55hrs"
            else:
                if turno(0)=="Vespertino":
                    messsage = "14:00, 15:00, 16:00, 17:00, 18:00, 19:00, 20:00, 21:00 y 21:55hrs"
                else:
                    messsage = "22:00, 23:00, 00:00, 01:00, 02:00, 03:00, 04:00, 05:00  y 05:55hrs"
            return messsage

        label_actua_hora  = tk.Label(canvas_hora,text=" Las horas de actualización son: "+actualizacion(), font=("Arial", 25))
        label_actua_hora.grid(row=2, column=0, pady=20)
        root.after(1000, actualizar_hora)

    actualizar_hora()
    
    
    actualizar_base_datos()    
    #root.after(60000, actualizar_base_datos)  # Comienza la actualización cada 1 minuto
    root.protocol("WM_DELETE_WINDOW", cerrar_app)
    root.mainloop()
    
def cerrar_app():
    root.destroy()   # Cierra todas las ventanas de Tkinter
    os._exit(0) 
def enviar_correo_especial(tipo):
    today = datetime.date.today().day
    if today == 1:
        # Último día del mes anterior
        last_day_of_prev_month = today - datetime.timedelta(days=1)
            
        # Primer día del mes anterior
        first_day_of_prev_month = last_day_of_prev_month.replace(day=1)
        resultados =  first_day_of_prev_month,last_day_of_prev_month
        try:
            import sys
            
                # Detecta si está ejecutándose desde el exe o desde Python
            if getattr(sys, 'frozen', False):
                # Carpeta donde está el .exe
                carpeta_base = os.path.dirname(sys.executable)
            else:
                # Carpeta donde está el script .py
                carpeta_base = os.path.dirname(os.path.abspath(__file__))

            # Ruta completa del archivo JSON
            ruta_json = os.path.join(carpeta_base, "recipients.json")
            with open(ruta_json, 'r') as f:
                destinatarios = json.load(f)  # Cargar la lista de destinatarios
                    
            pdf_path = generar_pdf1("especial",resultados)
            if pdf_path:
                enviar_correo(pdf_path, destinatarios,"Reporte Mensual de Producción","Adjunto encontrarás el reporte Mensual de producción.")

        except FileNotFoundError:
            print("El archivo de destinatarios no se encontró.")
        except json.JSONDecodeError as e:
            print(f"Error al cargar destinatarios: {e}")         
         

def enviar_correox():
    try:
        with open(rp('recursos/recipients.json'), 'r') as f:
            destinatarios = json.load(f)  # Cargar la lista de destinatarios
            
        pdf_path = generar_pdf(False)  #función genera el PDF y devuelve la ruta
        if pdf_path:
            print("Enviando correo...")
            enviar_correo(pdf_path, destinatarios,"Reporte Diario de Producción","Adjunto encontrarás el reporte Diario de producción.")

    except FileNotFoundError:
        print("El archivo de destinatarios no se encontró.")
    except json.JSONDecodeError as e:
        print(f"Error al cargar destinatarios: {e}")
         
def ejecutar_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)       

