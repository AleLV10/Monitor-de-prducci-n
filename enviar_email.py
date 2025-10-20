import re
import tkinter as tk
from tkinter import ttk
import json
from tkinter import messagebox

def mostrar_configuracion_email():
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
            with open(r'C:\\Users\\RH1\\Downloads\\Pantalla\\pantalla\\recipients.json', 'w') as f:
                json.dump(recipients, f)  # Always overwrite with a fresh array
            label_status.config(text="Destinatarios guardados correctamente.", fg="green")
            if messagebox.askyesno("", "Destinatarios guardados correctamente. \n¿Desea salir?"):
                root.quit()

    def load_recipients():
        try:
            with open(r'C:\\Users\\RH1\\Downloads\\Pantalla\\pantalla\\recipients.json', 'r') as f:
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
    root.iconbitmap(r'C:\\Users\\RH1\\Downloads\\Pantalla\\pantalla\\icono.ico')
    root.resizable(False, False)
    # Cargar la imagen
    img = tk.PhotoImage(file=r"C:\\Users\\RH1\\Downloads\\Pantalla\\pantalla\\logo_ecopac_copy.png")
    img = img.subsample(7)  # Ajustar tamaño de la imagen

    # Frame para el título e imagen
    frame_top = tk.Frame(root, bg="#F2F7F7", pady=20)
    frame_top.pack(fill=tk.X)

    frame_top.image_names = img  # Esto evita que la imagen sea eliminada

    # Título e Imagen en un diseño horizontal
    label_img = tk.Label(frame_top, image=frame_top.image_names, bg="#F2F7F7")
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

# Ejemplo de uso
#ventana = tk.Tk()
mostrar_configuracion_email()

