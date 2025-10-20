#Desarrollado por Alejandra López Vega
#Para Comercializadora de Plásticos de Alta Calidad S.A. de C.V.
#Version 1.0 lanzada en Octubre de 2025
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
from resource_path import resource_path as rp

def mostrar_reporte(ventana):
    def validar_fechas():
        fecha_inicio = cal_inicio.get_date()
        fecha_fin = cal_fin.get_date()
        hoy = datetime.now().date()
        
        if fecha_inicio > hoy:
            messagebox.showerror("Error de fecha", "La fecha de inicio no puede ser mayor que la fecha de hoy.")
            cal_inicio.set_date(hoy)
            ventana.lift()  # Eleva la ventana principal después de cerrar el messagebox
            
        if fecha_fin > hoy:
            messagebox.showerror("Error de fecha", "La fecha de fin no puede ser mayor que la fecha de hoy.")
            cal_fin.set_date(hoy)
            ventana.lift()  # Eleva la ventana principal después de cerrar el messagebox
            
        if fecha_fin < fecha_inicio:
            messagebox.showerror("Error de fecha", "La fecha de fin no puede ser menor que la fecha de inicio.")
            cal_fin.set_date(fecha_inicio)
            ventana.lift()  # Eleva la ventana principal después de cerrar el messagebox

    # Crear el título
    
    ventana.configure(bg='#F2F7F7')
    ventana.resizable(False, False)
    ventana.iconbitmap(rp("recursos/icono.ico"))
    
    titulo = tk.Label(ventana, text="Reporte especial de producción",  bg="#F2F7F7",font=("Arial", 16, "bold"))
    titulo.pack(pady=20)

    # Frame con borde negro para la selección de fechas
    frame_fecha_seccion = tk.LabelFrame(ventana, text="Por Fecha", bg="#F2F7F7", font=("Arial", 12))
    frame_fecha_seccion.pack(padx=10, pady=10, fill="x")

    # Frame para la selección de fechas dentro del frame con borde negro
    frame_fecha = tk.Frame(frame_fecha_seccion, bg="#F2F7F7")
    frame_fecha.pack()

    label_fecha_inicio = tk.Label(frame_fecha, text="Fecha de inicio:", bg="#F2F7F7", font=("Arial", 12))
    label_fecha_inicio.grid(row=0, column=0, padx=5, pady=5)
    cal_inicio = DateEntry(frame_fecha, selectmode='day', date_pattern='yyyy-mm-dd', maxdate=datetime.now().date())
    cal_inicio.grid(row=0, column=1, padx=5, pady=5)

    label_fecha_fin = tk.Label(frame_fecha, text="Fecha de fin:", bg="#F2F7F7", font=("Arial", 12))
    label_fecha_fin.grid(row=1, column=0, padx=5, pady=5)
    cal_fin = DateEntry(frame_fecha, selectmode='day', date_pattern='yyyy-mm-dd', maxdate=datetime.now().date())
    cal_fin.grid(row=1, column=1, padx=5, pady=5)

    cal_inicio.bind("<<DateEntrySelected>>", lambda e: validar_fechas())
    cal_fin.bind("<<DateEntrySelected>>", lambda e: validar_fechas())

    # Botón para obtener los resultados
    def obtener_resultados():
        ventana.quit()
        
    btn_obtener = tk.Button(ventana, text="Obtener Resultados", command=obtener_resultados,bg="#009688", fg="white", font=("Arial", 12), bd=0, padx=10, pady=5, relief="raised")
    btn_obtener.pack(side="bottom", pady=10)

    # Nota al final de la pantalla
    nota = tk.Label(ventana, text="Nota: Recuerda que la fecha de inicio comenzará en el turno de la mañana y terminará en el turno nocturno, "
                                  "es importante que tu fecha final sea la fecha de fin del turno nocturno.", bg="#F2F7F7", font=("Arial", 10), wraplength=480)
    nota.pack(side="bottom", pady=10)

    # Configurar el estilo del frame con borde negro
    style = ttk.Style()
    style.configure("Custom.TLabelframe.Label", font=("Arial", 12, "bold"))
    style.configure("Custom.TLabelframe", borderwidth=2, relief="solid")

    # Iniciar el bucle principal de la ventana
    ventana.mainloop()

    fecha_inicio = cal_inicio.get_date()
    fecha_fin = cal_fin.get_date()
    return fecha_inicio, fecha_fin
