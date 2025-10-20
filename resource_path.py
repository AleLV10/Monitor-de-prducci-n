
import os, sys
def resource_path(relative_path):
    """Devuelve la ruta absoluta, compatible con .py y .exe"""
    try:
        base_path = sys._MEIPASS  # Carpeta temporal creada por PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
