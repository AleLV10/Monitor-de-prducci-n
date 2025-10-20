#Desarrollado por Alejandra López Vega
#Para Comercializadora de Plásticos de Alta Calidad S.A. de C.V.
#Version 1.0 lanzada en Octubre de 2025
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as mb
from resource_path import resource_path as rp
import Coneccion
usuario=""
contaseña=""
def cerrar_ventana(ventana):
    ventana.destroy()

def abrir_nueva_ventana(usu, cont):
    #import python
    import pestañas
    pestañas.actualiza(usu, cont)
    #import Pantalla_Principal
    #Pantalla_Principal.actualiza(usu, cont)


def iniciar_sesion():
    # Aquí puedes agregar la lógica para iniciar sesión
    if entrada_usuario.get()!="":
        usuario=entrada_usuario.get()
        contaseña=entrada_contrasena.get()
        conn=Coneccion.coneccion(host = 'IP DEL SERVIDOR', database='RUTA BD EN SERVIDOR', user=usuario, password=contaseña,accion="consulta",consulta="",datos={})
        if conn=="Conexión exitosa a Firebird!":
            cerrar_ventana(root)
            abrir_nueva_ventana(usuario,contaseña)
        else:
            mb.showerror("Error!!",conn)
    else: 
        if entrada_usuario.get()=="":
             mb.showwarning("Cuidado!!","No puede dejar el campo vacío")  
     
root = tk.Tk()
root.title("Bienvenido")
root.geometry("400x240+760+300")
root.iconbitmap(rp("recursos/icono.ico"))
root.configure(bg="#F2F7F7")

# Bloquear la ventana para que no se pueda redimensionar
root.resizable(False, False)

# Crear un estilo para el botón
style = ttk.Style()
style.configure('Custom.TButton', background='lime green')

# Cargar la imagen y ajustar sus dimensiones
image = tk.PhotoImage(file=rp("recursos/logo_ecopac.png"))
image = image.subsample(5)  # Ajustar la imagen a la mitad del tamaño original

# Mostrar la imagen en un Label
label = tk.Label(root, image=image, bg="#F2F7F7")
label.place(relx=0.5, rely=0, anchor=tk.N)

# Etiqueta y entrada de texto para el usuario
label_usuario = tk.Label(root, text="Usuario:", bg="#F2F7F7", font=("Arial", 12))
label_usuario.place(relx=0.1, rely=0.55, anchor=tk.W)
texto_por_defecto = tk.StringVar(value="Usuario")
entrada_usuario = tk.Entry(root, width=30, textvariable=texto_por_defecto, font=("Arial", 12), bd=2, )
entrada_usuario.place(relx=0.28, rely=0.5)

# Etiqueta y entrada de texto para la contraseña
label_contrasena = tk.Label(root, text="Contraseña:", bg="#F2F7F7", font=("Arial", 12))
label_contrasena.place(relx=0.034, rely=0.68, anchor=tk.W)
texto_por_defecto1 = tk.StringVar(value="Contraseña")
entrada_contrasena = tk.Entry(root, width=30, show="*", textvariable=texto_por_defecto1, font=("Arial", 12), bd=2,)
entrada_contrasena.place(relx=0.28, rely=0.63)

# Botón de iniciar sesión con el estilo personalizado
boton_iniciar_sesion = tk.Button(root, text="Iniciar sesión", command=iniciar_sesion, bg="#009688", fg="white", font=("Arial", 12), bd=0, padx=10, pady=5, relief="raised")
boton_iniciar_sesion.place(relx=0.965, rely=0.95, anchor=tk.SE)
root.mainloop()
