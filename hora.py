#Desarrollado por Alejandra López Vega
#Para Comercializadora de Plásticos de Alta Calidad S.A. de C.V.
#Version 1.0 lanzada en Octubre de 2025
import tkinter as tk
import time
from threading import Thread

class ClockThread(Thread):
    def __init__(self, label):
        super().__init__()
        self.label = label
        self.daemon = True  # Hilo demonio para que termine con la aplicación
        self.start()

    def run(self):
        while True:
            current_time = time.strftime('%H:%M:%S')
            self.label.config(text=current_time)
            time.sleep(1)  # Actualiza cada segundo

class InfoThread(Thread):
    def __init__(self, label):
        super().__init__()
        self.label = label
        self.daemon = True
        self.start()

    def run(self):
        while True:
            current_hour = time.strftime('%H')
            if current_hour == '03':  # Mostrar información cada hora
                self.label.config(text="Información actualizada cada hora")
                time.sleep(60)  # Espera 1 hora
            else:
                time.sleep(60)  # Verifica cada minuto si es hora en punto

def main():
    root = tk.Tk()
    root.title("Reloj con Información")
    
    clock_label = tk.Label(root, font=('calibri', 40, 'bold'), bg='white')
    clock_label.pack(pady=20)

    info_label = tk.Label(root, font=('calibri', 14), bg='lightyellow', wraplength=400, justify='left')
    info_label.pack(pady=10, padx=10)

    clock_thread = ClockThread(clock_label)
    info_thread = InfoThread(info_label)

    root.mainloop()

if __name__ == "__main__":
    main()