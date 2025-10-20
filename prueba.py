from decimal import Decimal
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import smtplib
import threading
import time
import tkinter as tk
from tkinter import BOTH, YES, Canvas, Label, ttk
from tkinter import font
import datetime
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import base64
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Se crea la ventana del programa
root = tk.Tk()

user = ""
pasw = ""
consulta = "SELECT MG_ORDENES_PROD.FOLIO, MG_ORDENES_PROD.SOLICITANTE,MG_ORDENES_PROD.OBSERVACIONES,  MG_ORDENES_PROD_ARTS.CANTIDAD_POR_PROD,MG_ORDENES_PROD_ARTS.CANTIDAD_PRODUCIDA FROM MG_ORDENES_PROD INNER JOIN MG_ORDENES_PROD_ARTS ON MG_ORDENES_PROD_ARTS.ORDEN_PROD_ID = MG_ORDENES_PROD.ORDEN_PROD_ID AND  MG_ORDENES_PROD.ESTATUS= 'P'"
ban=True
lista=[(31, '20X30 MILLAR', '20X30 MILLAR', Decimal('1000000'), Decimal('3363.1')), (215, 'Camiseta Mediana poliponch azul', 'Camiseta Mediana poliponch azul', Decimal('900000'), Decimal('573126.306')), (255, '12x20c150 comercial', '12x20c150 comercial', Decimal('900000'), Decimal('147712.45275')), (262, '20x30c150 Comercial', '20x30c150 Comercial', Decimal('900000'), Decimal('617423.41')), (279, '60x90c200 Negro Comercial', '60x90c200 Negro Comercial', Decimal('900000'), Decimal('475307.22')), (296, '40x60c150 Comercial', '40x60c150 Comercial', Decimal('900000'), Decimal('438460.9005')), (312, '25x35c35 millar', '25x35c35 millar', Decimal('900000'), Decimal('833566.414')), (667, '35x45c100 rollos comercial', '35x45c100 rollos comercial', Decimal('999000'), Decimal('392837.515')), (673, '60x90c70 rollo comercial', '60x70c70 rollo comercial', Decimal('999000'), Decimal('223003.829')), (1193, 'CAMISETA DE BAJA GRANDE NEGRA', 'CAMISETA DE BAJA GRANDE NEGRA', Decimal('999999'), Decimal('382968.1035')), (10635, 'ROLLODEBAJA25X35C100NATURAL', None, Decimal('999999'), Decimal('297208.84')), (11333, 'BOBINA95C100BLANCOSIN PERFE', None, Decimal('100000'), Decimal('78099.62')), (12924, '60X90C300NAT', None, Decimal('10000'), Decimal('2849.75')), (13148, '90X60C320AZUL', None, Decimal('2000'), Decimal('724.2'))]

def turno(num):
    # Obtener la hora actual
    hora_actual = datetime.datetime.now().time()
    if num != 0:
        hora_actual =datetime.datetime.combine(datetime.date.today(), hora_actual) - datetime.timedelta(hours=8)

    # Verificar el turno actual
    if hora_actual.hour >= 6 and hora_actual.hour <= 13:
        return "Turno: Matutino"
    elif hora_actual.hour >= 14 and hora_actual.hour <= 19:
        return "Turno: Vespertino"
    else:
        return "Turno: Nocturno"    
def tabla_timer(tree):
    while True:
        global lista
        # Limpiar la tabla antes de insertar nuevos datos
        tree.delete(*tree.get_children())
        
        # Actualizar los datos y agregarlos a la tabla
        for item in lista:
            producido = Decimal(item[3]) - Decimal(item[4])
            tree.insert('', tk.END, values=(item[0], item[1], item[2], '{:,.2f}'.format(round(float(item[3]), 2)), '{:,.2f}'.format(round(float(item[4]), 2)), '{:,.2f}'.format(round(float(producido), 2))))
            
        # Ajustar el ancho de las columnas según el contenido
        for col in tree["columns"]:
            width = max(font.Font().measure(col.title()), len(col) * 6)  # Ajusta el ancho mínimo de la columna
            tree.column(col, width=width)
            
        # Extraer los datos relevantes para la gráfica
        # Extraer los datos relevantes para la gráfica
        etiquetas = [item[1] for item in lista if item[1] is not None and not "bobina" in item[1].lower()]
        cantidades_faltantes = [round(float(item[4]), 2) for item in lista if item[4] is not None and not "bobina" in item[1].lower()]

        # Graficar los datos
        ax.clear()
        
        
        ax.bar(etiquetas, cantidades_faltantes, align='center', color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28'], edgecolor='none')
        
        # Dibujar las líneas indicadoras de cantidad detrás de las barras
        for i, n in enumerate(cantidades_faltantes):
            ax.text(i, n + 10, '{:,.2f}'.format(n), ha='center', va='bottom')

        #ax.set_ylabel('Cantidad Faltante')
        ax.set_title("Produccion bolseo del " + turno(0))
        numeros_etiqueta = [str(i + 1) for i in range(len(etiquetas))]  # Crear una lista de números de etiqueta
        ax.set_xlabel('Observaciones')
        ax.set_xticks(range(len(etiquetas)))
        ax.set_xticklabels(numeros_etiqueta, ha='right')
        ax.set_facecolor(color="#EBFAF5")
        ax.set_yticklabels([])
        ax.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
        plt.axis([-1,5,0,1200])
        
        color=['#009688', '#388E3C', '#4CAF50', '#7CB342', '#C0CA33', '#FFCA28']
        y=0
        cont =0
        yl=0
        scrollbar1 = ttk.Scrollbar(orient=tk.VERTICAL, command=canvasc.yview)
        canvasc.config(yscrollcommand=scrollbar1.set)
        scrollbar1.place(relx=0.34, rely=0.5, relheight=0.49)
        scrollbar2 = ttk.Scrollbar(orient=tk.HORIZONTAL, command=canvasc.xview)
        canvasc.config(xscrollcommand=scrollbar2.set)
        scrollbar2.place(relx=0.25, rely=0.985, relwidth=0.09)
        # Dibujar cuadrados
        for i in range(len(etiquetas)):
            if i%len(color)==0:
                cont+=1
            a=i-(6*cont)
            numero_etiqueta = str(i + 1)
            canvasc.create_rectangle(7, (7+y), 30, (30+y), width=1, fill=color[a])
            label1 = ttk.Label(root, text=numero_etiqueta+". "+etiquetas[i], font=("Helvetica", 9,), anchor="w")
            label1.place(relx=0.27, rely=(0.5075+yl), relwidth=0.065, relheight=0.0255)
            #label1.configure(background='white') 
            y+=33
            yl+=0.0335
           
      
        # Actualizar la gráfica
        canvas.draw()   
        
        
        # Extraer los datos relevantes para la gráfica
        etiquetas2 = [item[1] for item in lista if item[1] is not None and "bobina" in item[1].lower()]
        cantidades_faltantes2 = [item[4] for item in lista if item[4] is not None and "bobina" in item[1].lower()]
        
        # Graficar los datos
        ax2.clear()
        ax2.bar(etiquetas2, cantidades_faltantes2,align='center', color=['#FFB300', '#FFC300', '#FF5733', '#C70039', '#900C3F', '#581845'], edgecolor='none')
        
        # Dibujar las líneas indicadoras de cantidad detrás de las barras
        for i, n in enumerate(cantidades_faltantes2):
            ax2.text(i, n + 10, '{:,.2f}'.format(n), ha='center', va='bottom')
            
        
        #ax2.set_ylabel('Cantidad Faltante')
        ax2.set_title("Produccion bobina del "+turno(0))
        ax2.set_xlabel('Observaciones')
        ax2.set_xticks(range(len(etiquetas2)))
        ax2.set_facecolor(color="#EBFAF5")
        ax2.set_xticklabels(etiquetas2, rotation=20, ha='right')
        ax2.set_yticklabels([])
        ax2.grid(axis='y', linestyle='dashed', color='gray', alpha=0.6)   
       
        
        # Actualizar la gráfica
        canvas2.draw()
        
        time.sleep(5)  # Esperar 5 segundos

def cerrar_ventana():
    root.quit()
    root.destroy()  

def generar_pdf():
    # Crear una nueva figura de Matplotlib
    '''fig = Figure(figsize=(8, 6), dpi=100)
    ax = fig.add_subplot(111)
    ax.plot([1, 2, 3, 4], [1, 4, 9, 16], 'b-')
    ax.set_xlabel('X Label')
    ax.set_ylabel('Y Label')
    ax.set_title('Title')'''
    fig_table = plt.figure(figsize=(8, 6))
    ax_table = fig_table.add_subplot(111)
    ax_table.axis('off')  # Desactivar los ejes para la tabla
    #ax_table.table(cellText=tree.get_string().splitlines(), colLabels=[col.title() for col in tree['columns']], loc='center')
    # Guardar la figura en un archivo PDF temporal
    with PdfPages("informe_produccion.pdf") as pdf:
        pdf.savefig(fig)
        pdf.savefig(fig2)
        pdf.savefig(fig3)
        pdf.savefig(fig4)
        #pdf.savefig(fig_table)
    return "informe_produccion.pdf"



def enviar_correo(pdf_file):
    remitente = "hirecursos351@gmail.com"
    destinatario = "rechumanos@copacsa.com"
    asunto = "Informe de Producción"
    cuerpo = "Adjunto encontrarás el informe de producción."

    # Crear el mensaje
    mensaje = MIMEMultipart()
    mensaje["From"] = remitente
    mensaje["To"] = destinatario
    mensaje["Subject"] = asunto

    # Adjuntar el archivo PDF al mensaje
    adjunto = MIMEBase("application", "octet-stream")
    adjunto.set_payload(open(pdf_file, "rb").read())
    encoders.encode_base64(adjunto)
    adjunto.add_header("Content-Disposition", f"attachment; filename={pdf_file}")
    mensaje.attach(adjunto)

    # Conectar al servidor SMTP y enviar el correo electrónico
    servidor_smtp = "smtp.gmail.com"
    puerto_smtp = 587
    usuario_smtp = "hirecursos351@gmail.com"
    contrasena_smtp = "Fran-rh2023"

    with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
        servidor.starttls()
        servidor.login(usuario_smtp, contrasena_smtp)
        servidor.send_message(mensaje)
        

def create_message_with_attachment(sender, to, subject, message_text, file_path):
    message = MIMEMultipart()
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    # Adjuntar el archivo PDF
    content_type, encoding = mimetypes.guess_type(file_path)
    if content_type is None or encoding is not None:
        content_type = "application/octet-stream"
    main_type, sub_type = content_type.split("/", 1)
    with open(file_path, "rb") as file:
        part = MIMEBase(main_type, sub_type)
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
    message.attach(part)

    return {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        print("Mensaje enviado:", message["id"])
    except Exception as e:
        print("Error al enviar el mensaje:", e)

    
# Configurar la ventana principal
root.title('ECOPAC, Bienvenido')
root.iconbitmap('icono.ico')
ban=False
an = root.winfo_screenwidth()-2
al = root.winfo_screenheight()-80

# Establecer las dimensiones de la ventana
root.geometry("%dx%d+0+0" % (an, al))  # El +0+0 asegura que la ventana esté en la esquina superior izquierda
root.resizable(0, 1)  # Evitar que la ventana sea redimensionable
root.config(bg='silver', width='1220', height='200')

# Se crea el menú de la ventana
menu = tk.Menu()

# Se crean las opciones principales
menu_Producción = tk.Menu(menu, tearoff=0)
menu_Configuración = tk.Menu(menu, tearoff=0)

# Agregar las opciones principales al menú
menu.add_cascade(label="Producción", menu=menu_Producción)
menu.add_cascade(label="Configuración", menu=menu_Configuración)
menu.add_cascade(label="Salir", command=root.quit)

# Se crean las subopciones al menú "Producción"
menu_Producción.add_command(label="Descargar PDF", command=lambda: generar_pdf())
menu_Producción.add_command(label="Enviar e-mail", command=lambda: enviar_correo(generar_pdf()))

# Se crean las subopciones para "Archivo > Preferencias"
menu_preferencias = tk.Menu(menu_Producción, tearoff=0)
menu_preferencias.add_command(label="Arial")
menu_preferencias.add_command(label="Calibri")
menu_preferencias.add_command(label="Times New Roman")

# Se crean las subopciones al menú "Configuración"
menu_Configuración.add_command(label="Editar remitente")
# Se crea la cascada" al menú "Configuración"
menu_Configuración.add_cascade(label="Preferencias", menu=menu_preferencias)

# Se muestra la barra de menú en la ventana principal
root.config(menu=menu)

root.configure(bg='#F2F7F7')

# Cargar la imagen y ajustar sus dimensiones
image = tk.PhotoImage(file="logo_ecopac.png")
image = image.subsample(5)  # Ajustar la imagen a la mitad del tamaño original

# Mostrar la imagen en un Label
label = ttk.Label(root, image=image)
label.place(relx=0.01,rely=0.01,relwidth=0.13,relheight=0.11)
label.configure(background='#F2F7F7') 


label_turno = ttk.Label(root, text=turno(0), font=("Helvetica", 30,"bold underline"), anchor="center")
label_turno.place(relx=0.15, rely=0.01, relwidth=0.18, relheight=0.11)
label_turno.configure(background='#F2F7F7') 

label_monitor = ttk.Label(root, text="Monitor de producción", font=("Helvetica", 50), anchor="center",)
label_monitor.place(relx=0.39, rely=0.01, relwidth=0.38, relheight=0.11)
label_monitor.configure(background='#F2F7F7') 

tree_frame = tk.Frame(root)
tree_frame = tk.Frame(root, borderwidth=2, relief="sunken")
tree_frame.place(relx=0.01, rely=0.14, relwidth=0.625, relheight=0.35)
tree = ttk.Treeview(tree_frame, show='headings', selectmode='browse', style="Treeview")
tree.pack(side='left', fill='both', expand=True)

scrollbar = ttk.Scrollbar(orient=tk.VERTICAL, command=tree.yview)
scrollbar.place(relx=0.63, rely=0.14, relheight=0.35)

style = ttk.Style()
style.theme_use("default")
tree.config(columns=("Folio", "Solicitante", "Observaciones", "Cantidad por producir", "Cantidad producida", "Cantidad faltante"), show="headings")
for heading in ("Folio", "Solicitante", "Observaciones", "Cantidad por producir", "Cantidad producida", "Cantidad faltante"):
    tree.heading(heading, text=heading)

t = threading.Thread(target=tabla_timer, args=(tree,))
t.daemon = True
t.start()

# Crear la figura de matplotlib
fig = Figure(figsize=(6, 5), dpi=100)
#fig.bar()
ax = fig.add_subplot(111)

# Graficar los datos
ax.bar([], [],align='center', color=['blue','red','green','white','brown'], edgecolor='none')

# Crear el lienzo de matplotlib dentro de Tkinter
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.draw()
canvas.get_tk_widget().place(relx=0.01, rely=0.50, relwidth=0.26, relheight=0.49)

# Crear la figura de matplotlib
fig2 = Figure(figsize=(6, 5), dpi=100)
ax2 = fig2.add_subplot(111)

# Graficar los datos
ax2.bar([], [])

# Crear el lienzo de matplotlib dentro de Tkinter
canvas2 = FigureCanvasTkAgg(fig2, master=root)
canvas2.draw()
canvas2.get_tk_widget().place(relx=0.32, rely=0.50, relwidth=0.26, relheight=0.49)

canvasc = Canvas( bg='white')
canvasc.place(relx=0.25, rely=0.50, relwidth=0.09, relheight=0.49)
#canvasc.pack(expand=YES, fill=BOTH)

label_turno_anterior = ttk.Label(root, text="   "+turno(1)+"   ", font=("Helvetica", 20,"bold underline"), anchor="center")
label_turno_anterior .place(relx=0.65, rely=0.11, relwidth=0.34, relheight=0.11)
label_turno_anterior .configure(background='#F2F7F7') 



# Crear la figura de matplotlib
fig3 = Figure(figsize=(6, 5), dpi=100)
#fig.bar()
ax3 = fig3.add_subplot(111)

# Graficar los datos
ax3.bar([], [],align='center', color=['blue','red','green','white','brown'], edgecolor='none')

# Añadir etiquetas y título
ax3.set_xlabel('Observaciones')
ax3.set_ylabel('Cantidad Faltante')
ax3.set_title("Produccion bolseo del "+turno(1))

# Crear el lienzo de matplotlib dentro de Tkinter
canvas3 = FigureCanvasTkAgg(fig3, master=root)
canvas3.draw()
canvas3.get_tk_widget().place(relx=0.695, rely=0.20, relwidth=0.25, relheight=0.30)

# Crear la figura de matplotlib
fig4 = Figure(figsize=(6, 5), dpi=100)
ax4 = fig4.add_subplot(111)

# Graficar los datos
ax4.bar([], [])

# Añadir etiquetas y título
ax4.set_xlabel('Observaciones')
ax4.set_ylabel('Cantidad Faltante')
ax4.set_title("Produccion bobina del "+turno(1))

# Crear el lienzo de matplotlib dentro de Tkinter
canvas4 = FigureCanvasTkAgg(fig4, master=root)
canvas4.draw()
canvas4.get_tk_widget().place(relx=0.695, rely=0.52, relwidth=0.25, relheight=0.30)
# Configurar el controlador de eventos para cerrar la ventana

produccion_total= 6000
produccion_anterior=50000
resto= produccion_total-produccion_anterior

produccion_actual = ttk.Label(root, text="Produccion total del "+turno(0)+ " \t" + '{:,.2f}'.format(produccion_total), font=("Helvetica", 14), anchor="e")
produccion_actual .place(relx=0.695, rely=0.84, relwidth=0.25, relheight=0.05)
produccion_actual .configure(background='#F2F7F7') 

produccion_atras = ttk.Label(root, text="Produccion total del "+turno(1)+ "\t" + '{:,.2f}'.format(produccion_anterior), font=("Helvetica", 14,"underline" ), anchor="e")
produccion_atras.place(relx=0.695, rely=0.89, relwidth=0.25, relheight=0.05)
produccion_atras .configure(background='#F2F7F7') 

diferencia = ttk.Label(root, text="Diferencia: \t"+'{:,.2f}'.format(resto), font=("Helvetica", 14,"bold "), anchor="e")
diferencia.place(relx=0.695, rely=0.94, relwidth=0.25, relheight=0.05)
diferencia .configure(background='#F2F7F7') 

if resto<=0:
    produccion = ttk.Label(root, text="!!Animo!!, faltan "+'{:,.2f}'.format(resto*-1)+"Kg \n para alcanzar al \n "+turno(1), font=("Helvetica", 14,"bold "), anchor="center")
    produccion.place(relx=0.84, rely=0.01, relwidth=0.15, relheight=0.1)
    produccion .configure(background='#87CEEB') 
else:
    produccion = ttk.Label(root, text="!!FELICIDADES \n PRODUCCION \n SUPERADA!!", font=("Helvetica", 18,"bold "), anchor="center")
    produccion.place(relx=0.84, rely=0.01, relwidth=0.15, relheight=0.1)
    produccion .configure(background='#87CEEB') 
    
root.protocol("WM_DELETE_WINDOW", cerrar_ventana)

# Bucle de ejecución del programa
root.mainloop()
