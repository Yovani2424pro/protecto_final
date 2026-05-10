import tkinter as tk
from tkinter import ttk

class DataAnalyzerApp:
    """Clase principal de la aplicación - Orientada a Objetos"""
    
    CSV_FILE = 'titanic(1).csv' 
    REPORTES_DIR = 'reportes'
    MAX_ROWS_TABLE = 100

    def __init__(self, root):
        self.root = root
        self.root.title("Analizador de Datos CSV")
        self.root.geometry("1100x700")
        self.root.configure(bg='#2b2b2b')
        self.root.minsize(900, 600)
        
        # Variables de datos
        self.df_original = None
        self.df_filtrado = None
        
        # Aplicar tema oscuro a ttk
        self.setup_theme()

    def setup_theme(self):
        """Configura el tema oscuro para los widgets ttk"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configuración general
        style.configure('.', background='#2b2b2b', foreground='white', fieldbackground='#3c3f41', borderwidth=0)
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TLabel', background='#2b2b2b', foreground='white', font=('Segoe UI', 10))
        style.configure('TButton', background='#3c3f41', foreground='white', font=('Segoe UI', 10, 'bold'), padding=5)
        style.map('TButton', background=[('active', '#505355')])

if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()