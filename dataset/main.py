# ============================================================
# SISTEMA DE ANALISIS DE TERREMOTOS
# PYTHON + TKINTER + PANDAS + MATPLOTLIB
# ============================================================

# ============================================================
# LIBRERIAS
# ============================================================

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# ============================================================
# CLASE PRINCIPAL
# ============================================================

class SistemaTerremotos:

    # ========================================================
    # INICIO
    # ========================================================

    def __init__(self, root):

        self.root = root
        self.root.title("Sistema de Analisis de Terremotos")
        self.root.geometry("1500x850")
        self.root.configure(bg="#0f172a")

        # ====================================================
        # TEMA OSCURO
        # ====================================================

        estilo = ttk.Style()
        estilo.theme_use("clam")

        estilo.configure(
            "Treeview",
            background="#1e293b",
            foreground="white",
            fieldbackground="#1e293b",
            rowheight=30,
            font=("Segoe UI", 10)
        )

        estilo.configure(
            "Treeview.Heading",
            background="#111827",
            foreground="white",
            font=("Segoe UI", 10, "bold")
        )

        estilo.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8
        )

        # ====================================================
        # CREAR CARPETA REPORTES
        # ====================================================

        if not os.path.exists("reportes"):
            os.makedirs("reportes")

        # ====================================================
        # VARIABLES
        # ====================================================

        self.df = None
        self.df_filtrado = None

        # ====================================================
        # PANTALLA DE INICIO
        # ====================================================

        self.pantalla_inicio()

    # ========================================================
    # PANTALLA DE INICIO
    # ========================================================

    def pantalla_inicio(self):

        self.frame_inicio = tk.Frame(
            self.root,
            bg="#0f172a"
        )

        self.frame_inicio.pack(fill="both", expand=True)

        titulo = tk.Label(
            self.frame_inicio,
            text="🌎 SISTEMA DE TERREMOTOS",
            bg="#0f172a",
            fg="#38bdf8",
            font=("Segoe UI", 32, "bold")
        )

        titulo.pack(pady=120)

        subtitulo = tk.Label(
            self.frame_inicio,
            text="Analisis y visualizacion de datos sismicos",
            bg="#0f172a",
            fg="white",
            font=("Segoe UI", 16)
        )

        subtitulo.pack()

        boton = tk.Button(
            self.frame_inicio,
            text="INGRESAR",
            bg="#38bdf8",
            fg="black",
            font=("Segoe UI", 14, "bold"),
            width=20,
            height=2,
            command=self.iniciar_dashboard
        )

        boton.pack(pady=50)

    # ========================================================
    # INICIAR DASHBOARD
    # ========================================================

    def iniciar_dashboard(self):

        self.frame_inicio.destroy()

        self.crear_menu()
        self.crear_interfaz()
        self.cargar_csv()

    # ========================================================
    # MENU
    # ========================================================

    def crear_menu(self):

        barra_menu = tk.Menu(self.root)

        # ====================================================
        # ARCHIVO
        # ====================================================

        menu_archivo = tk.Menu(
            barra_menu,
            tearoff=0
        )

        menu_archivo.add_command(
            label="Exportar Excel",
            command=self.exportar_excel
        )

        menu_archivo.add_separator()

        menu_archivo.add_command(
            label="Salir",
            command=self.root.quit
        )

        barra_menu.add_cascade(
            label="Archivo",
            menu=menu_archivo
        )

        # ====================================================
        # HERRAMIENTAS
        # ====================================================

        menu_herramientas = tk.Menu(
            barra_menu,
            tearoff=0
        )

        menu_herramientas.add_command(
            label="Limpiar filtros",
            command=self.limpiar_filtros
        )

        barra_menu.add_cascade(
            label="Herramientas",
            menu=menu_herramientas
        )

        # ====================================================
        # AYUDA
        # ====================================================

        menu_ayuda = tk.Menu(
            barra_menu,
            tearoff=0
        )

        menu_ayuda.add_command(
            label="Acerca de",
            command=lambda: messagebox.showinfo(
                "Informacion",
                "Sistema de Analisis de Terremotos\nPython + Tkinter"
            )
        )

        barra_menu.add_cascade(
            label="Ayuda",
            menu=menu_ayuda
        )

        self.root.config(menu=barra_menu)

    # ========================================================
    # INTERFAZ
    # ========================================================

    def crear_interfaz(self):

        # ====================================================
        # PANEL IZQUIERDO
        # ====================================================

        self.panel_izquierdo = tk.Frame(
            self.root,
            bg="#111827",
            width=300
        )

        self.panel_izquierdo.pack(
            side="left",
            fill="y"
        )

        titulo = tk.Label(
            self.panel_izquierdo,
            text="📊 PANEL DE CONTROL",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 18, "bold")
        )

        titulo.pack(pady=20)

        # ====================================================
        # FILTRO CONTINENTE
        # ====================================================

        tk.Label(
            self.panel_izquierdo,
            text="Continente",
            bg="#111827",
            fg="white",
            font=("Segoe UI", 11)
        ).pack(pady=5)

        self.combo_continente = ttk.Combobox(
            self.panel_izquierdo,
            width=25,
            state="readonly"
        )

        self.combo_continente.pack(pady=5)

        # ====================================================
        # FILTRO MAGNITUD
        # ====================================================

        tk.Label(
            self.panel_izquierdo,
            text="Magnitud minima",
            bg="#111827",
            fg="white",
            font=("Segoe UI", 11)
        ).pack(pady=5)

        self.entry_magnitud = ttk.Entry(
            self.panel_izquierdo,
            width=28
        )

        self.entry_magnitud.pack(pady=5)

        # ====================================================
        # BOTONES
        # ====================================================

        ttk.Button(
            self.panel_izquierdo,
            text="Aplicar filtros",
            command=self.aplicar_filtros
        ).pack(pady=15)

        ttk.Button(
            self.panel_izquierdo,
            text="Limpiar filtros",
            command=self.limpiar_filtros
        ).pack(pady=5)

        ttk.Button(
            self.panel_izquierdo,
            text="Exportar Excel",
            command=self.exportar_excel
        ).pack(pady=5)

        # ====================================================
        # PANEL DERECHO
        # ====================================================

        self.panel_derecho = tk.Frame(
            self.root,
            bg="#0f172a"
        )

        self.panel_derecho.pack(
            side="right",
            fill="both",
            expand=True
        )

        # ====================================================
        # TARJETAS
        # ====================================================

        frame_tarjetas = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        frame_tarjetas.pack(fill="x", pady=10)

        self.label_total = self.crear_tarjeta(
            frame_tarjetas,
            "Total registros"
        )

        self.label_promedio = self.crear_tarjeta(
            frame_tarjetas,
            "Promedio magnitud"
        )

        self.label_maximo = self.crear_tarjeta(
            frame_tarjetas,
            "Magnitud maxima"
        )

        # ====================================================
        # TABLA
        # ====================================================

        frame_tabla = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        frame_tabla.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        columnas = (
            "Lugar",
            "Magnitud",
            "Profundidad",
            "Pais"
        )

        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings"
        )

        for col in columnas:

            self.tabla.heading(col, text=col)

            self.tabla.column(col, width=220)

        scrollbar_y = ttk.Scrollbar(
            frame_tabla,
            orient="vertical",
            command=self.tabla.yview
        )

        self.tabla.configure(
            yscrollcommand=scrollbar_y.set
        )

        self.tabla.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar_y.pack(
            side="right",
            fill="y"
        )

        # ====================================================
        # FRAME GRAFICAS
        # ====================================================

        self.frame_graficas = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        self.frame_graficas.pack(
            fill="both",
            expand=True
        )

    # ========================================================
    # TARJETAS
    # ========================================================

    def crear_tarjeta(self, parent, titulo):

        frame = tk.Frame(
            parent,
            bg="#1e293b",
            width=250,
            height=90
        )

        frame.pack(side="left", padx=10)
        frame.pack_propagate(False)

        tk.Label(
            frame,
            text=titulo,
            bg="#1e293b",
            fg="white",
            font=("Segoe UI", 11)
        ).pack(pady=10)

        valor = tk.Label(
            frame,
            text="0",
            bg="#1e293b",
            fg="#38bdf8",
            font=("Segoe UI", 20, "bold")
        )

        valor.pack()

        return valor

    # ========================================================
    # CARGAR CSV
    # ========================================================

    def cargar_csv(self):

        try:

            ruta = os.path.join(
                os.path.dirname(__file__),
                "earthquake_1995-2023.csv"
            )

            print("Buscando archivo en:")
            print(ruta)

            if not os.path.exists(ruta):

                messagebox.showerror(
                    "Error",
                    "No se encontro el archivo CSV\n\n"
                    "Asegurate de poner:\n"
                    "earthquake_1995-2023.csv\n\n"
                    "en la misma carpeta de main.py"
                )

                return

            # =================================================
            # LEER CSV
            # =================================================

            self.df = pd.read_csv(ruta)

            # =================================================
            # ELIMINAR NULOS
            # =================================================

            self.df.dropna(inplace=True)

            # =================================================
            # COPIA
            # =================================================

            self.df_filtrado = self.df.copy()

            # =================================================
            # COLUMNAS
            # =================================================

            columnas = self.df.columns.tolist()

            print(columnas)

            self.col_titulo = columnas[0]
            self.col_magnitud = columnas[1]
            self.col_profundidad = columnas[2]
            self.col_pais = columnas[3]

            # =================================================
            # CONTINENTES
            # =================================================

            if "continent" in self.df.columns:

                continentes = sorted(
                    self.df["continent"].unique()
                )

                self.combo_continente["values"] = (
                    ["Todos"] + continentes
                )

                self.combo_continente.current(0)

            # =================================================
            # ACTUALIZAR
            # =================================================

            self.actualizar_tarjetas()
            self.cargar_tabla()
            self.actualizar_graficas()

            messagebox.showinfo(
                "Exito",
                "CSV cargado correctamente"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo abrir el CSV\n\n{e}"
            )

    # ========================================================
    # TARJETAS
    # ========================================================

    def actualizar_tarjetas(self):

        total = len(self.df_filtrado)

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        promedio = round(
            magnitudes.mean(),
            2
        )

        maximo = magnitudes.max()

        self.label_total.config(text=str(total))
        self.label_promedio.config(text=str(promedio))
        self.label_maximo.config(text=str(maximo))

    # ========================================================
    # TABLA
    # ========================================================

    def cargar_tabla(self):

        for item in self.tabla.get_children():
            self.tabla.delete(item)

        datos = self.df_filtrado.head(100)

        for _, row in datos.iterrows():

            self.tabla.insert(
                "",
                "end",
                values=(
                    row[self.col_titulo],
                    row[self.col_magnitud],
                    row[self.col_profundidad],
                    row[self.col_pais]
                )
            )

    # ========================================================
    # FILTROS
    # ========================================================

    def aplicar_filtros(self):

        try:

            self.df_filtrado = self.df.copy()

            continente = self.combo_continente.get()

            magnitud = self.entry_magnitud.get()

            if (
                "continent" in self.df.columns
                and continente != ""
                and continente != "Todos"
            ):

                self.df_filtrado = self.df_filtrado[
                    self.df_filtrado["continent"] == continente
                ]

            if magnitud != "":

                self.df_filtrado = self.df_filtrado[
                    pd.to_numeric(
                        self.df_filtrado[self.col_magnitud],
                        errors="coerce"
                    ) >= float(magnitud)
                ]

            self.actualizar_tarjetas()
            self.cargar_tabla()
            self.actualizar_graficas()

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"Error en filtros\n\n{e}"
            )

    # ========================================================
    # LIMPIAR FILTROS
    # ========================================================

    def limpiar_filtros(self):

        self.entry_magnitud.delete(0, tk.END)

        self.df_filtrado = self.df.copy()

        self.actualizar_tarjetas()
        self.cargar_tabla()
        self.actualizar_graficas()

    # ========================================================
    # EXPORTAR EXCEL
    # ========================================================

    def exportar_excel(self):

        try:

            ruta_excel = os.path.join(
                "reportes",
                "reporte_terremotos.xlsx"
            )

            self.df_filtrado.to_excel(
                ruta_excel,
                index=False,
                engine="openpyxl"
            )

            messagebox.showinfo(
                "Exportado",
                f"Excel guardado en:\n{ruta_excel}"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo exportar\n\n{e}"
            )

    # ========================================================
    # GRAFICAS
    # ========================================================

    def actualizar_graficas(self):

        for widget in self.frame_graficas.winfo_children():
            widget.destroy()

        fig, axes = plt.subplots(
            1,
            3,
            figsize=(15, 4)
        )

        # ====================================================
        # PIE CHART
        # ====================================================

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        categorias = pd.cut(
            magnitudes,
            bins=[0, 4, 6, 10],
            labels=["Baja", "Media", "Alta"]
        )

        categorias.value_counts().plot.pie(
            ax=axes[0],
            autopct="%1.1f%%"
        )

        axes[0].set_title("Categorias")

        # ====================================================
        # GRAFICA BARRAS
        # ====================================================

        magnitudes.head(10).plot.bar(
            ax=axes[1]
        )

        axes[1].set_title("Magnitudes")

        # ====================================================
        # SCATTER
        # ====================================================

        if (
            "longitude" in self.df.columns
            and "latitude" in self.df.columns
        ):

            axes[2].scatter(
                self.df_filtrado["longitude"],
                self.df_filtrado["latitude"],
                alpha=0.5
            )

            axes[2].set_title(
                "Latitud y Longitud"
            )

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(
            fig,
            master=self.frame_graficas
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = SistemaTerremotos(root)

    root.mainloop()