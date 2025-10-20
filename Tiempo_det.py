"""import datetime
import threading
import time
# Tarea a ejecutarse cada determinado tiempo.
def timer():
    while True:
        fecha_actual = datetime.datetime.now()
        fecha=datetime.date.today()
        turnoActu = "Nocturno"
        if turnoActu == "Matutino":
            fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(6, 0, 0, 0))
            if fecha_actual.weekday() == 0:
                resta_dias = datetime.timedelta(days=2)
            else:
                resta_dias = datetime.timedelta(days=1)
            fecha = fecha - resta_dias
            fecha_inicio_consulta_ta=datetime.datetime.combine(fecha, datetime.time(22, 0, 0, 0))
            
        elif turnoActu == "Vespertino":
            fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(14, 0, 0, 0))
            fecha_inicio_consulta_ta = datetime.datetime.combine(fecha, datetime.time(6, 0, 0, 0))
        else:
            if datetime.time(0, 0) <= datetime.datetime.now().time() > datetime.time(6, 0):
                un_dia = datetime.timedelta(days=1)
                fecha = fecha - un_dia
            fecha_inicio_consulta = datetime.datetime.combine(fecha, datetime.time(22, 0, 0, 0))
            fecha_inicio_consulta_ta = datetime.datetime.combine(fecha, datetime.time(14, 0, 0, 0))
        
        fecha_inicio_consulta = fecha_inicio_consulta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        fecha_fin_consulta = fecha_actual.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        fecha_inicio_consulta_ta=fecha_inicio_consulta_ta.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        print("Fecha de inicio del turno actual",fecha_inicio_consulta)
        print("Fecha de termino del turno actual",fecha_fin_consulta)
        print("Fecha de inicio del turno anterior",fecha_inicio_consulta_ta)
        print("Fecha de termino del turno anterior",fecha_inicio_consulta)
        
        time.sleep(20)   # 3 segundos.
# Iniciar la ejecución en segundo plano.
t = threading.Thread(target=timer)
t.start()
print("Esto se ejecuta antes que la función f().")
    
    """
import re
import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox

def validar_destinatarios():
    contenido = entry_destinatarios.get("1.0", "end-1c")
    if not contenido.strip():
        messagebox.showerror("Error", "El campo de destinatarios no puede estar vacío.")
        root.lift()
        return False
        
    destinatarios = [email.strip() for email in contenido.split(',') if email.strip()]
    for email in destinatarios:
        if not validar_email(email):
            messagebox.showerror("Error", f"'{email}' no es un correo electrónico válido.")
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
        with open('recipients.json', 'w') as f:
            json.dump(recipients, f)  # Always overwrite with a fresh array
        label_status.config(text="Destinatarios guardados correctamente.", fg="green")
        if messagebox.askyesno("", "Destinatarios guardados correctamente. \n¿Desea salir?"):
            root.quit()

def load_recipients():
    try:
        with open('recipients.json', 'r') as f:
            content = f.read()
            # Print the content for debugging
            #print("Loaded content:", content)  
            recipients = json.loads(content)
            print(content)
        entry_destinatarios.delete('1.0', tk.END)  # Clear the field before loading
        entry_destinatarios.insert('1.0', ', '.join(recipients))
    except FileNotFoundError:
        pass
    except json.JSONDecodeError as e:
        messagebox.showerror("Error", f"Error al cargar destinatarios: {e}")

root = tk.Tk()
root.title("Gestión de Destinatarios")
root.geometry("600x450")  # Tamaño inicial de la ventana
root.configure(bg="#F2F7F7")
root.iconbitmap('icono.ico')
root.resizable(False, False)
# Cargar la imagen
img = tk.PhotoImage(file="logo_ecopac.png")
img = img.subsample(7)  # Ajustar tamaño de la imagen

# Frame para el título e imagen
frame_top = tk.Frame(root, bg="#F2F7F7", pady=20)
frame_top.pack(fill=tk.X)

# Título e Imagen en un diseño horizontal
label_img = tk.Label(frame_top, image=img, bg="#F2F7F7")
label_img.pack(side=tk.LEFT, padx=(5, 10))

label_title = tk.Label(frame_top, text="Gestión de Destinatarios", font=("Arial", 16, "bold"), bg="#F2F7F7")
label_title.pack(side=tk.LEFT)

# Frame para el contenido principal
frame_content = tk.Frame(root, bg="#F2F7F7")
frame_content.pack(padx=10, pady=1, fill=tk.BOTH, expand=True)


# Mensaje para los destinatarios
label_info = tk.Label(frame_content, text="Por favor, separa cada destinatario con una coma.", font=("Arial", 12), bg="#F2F7F7")
label_info.pack(pady=0)

entry_destinatarios = tk.Text(frame_content, font=("Helvetica", 12), width=50)
entry_destinatarios.config(width=50, height=10, font=("Consolas", 12), 
                               padx=15, pady=15, selectbackground="red")
entry_destinatarios.pack(pady=10)
# Botón para guardar
save_button = tk.Button(frame_content, text="Guardar Destinatarios", command=save_recipients, font=("Arial", 12), bg="#009688", fg="white", bd=0, padx=10, pady=5, relief="raised")
save_button.pack(pady=10)

# Etiqueta de estado
label_status = tk.Label(frame_content, text="", font=("Arial", 12), bg="#F2F7F7")
label_status.pack(pady=10)
# Cargar destinatarios si existen
load_recipients()

root.mainloop()
