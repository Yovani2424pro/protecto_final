# =====================================================
# IMPORTACIONES
# =====================================================

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# =====================================================
# CLASE PRINCIPAL
# =====================================================

class SistemaTerremotos:

    def __init__(self, root):

        self.root = root
        self.root.title("🌎 Sistema Profesional de Terremotos")
        self.root.geometry("1600x920")
        self.root.configure(bg="#0f172a")

        # =====================================================
        # ESTILOS
        # =====================================================

        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#111827",
            foreground="white",
            fieldbackground="#111827",
            rowheight=28,
            font=("Segoe UI", 9)
        )

        style.configure(
            "Treeview.Heading",
            background="#0f172a",
            foreground="#38bdf8",
            font=("Segoe UI", 10, "bold")
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 9, "bold"),
            padding=5
        )

        self.df = None
        self.df_filtrado = None

        # =====================================================
        # CREAR CARPETA REPORTES
        # =====================================================

        if not os.path.exists("reportes"):
            os.makedirs("reportes")

        self.pantalla_inicio()

    # =====================================================
    # PANTALLA INICIO
    # =====================================================

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
            font=("Segoe UI", 34, "bold")
        )

        titulo.pack(pady=100)

        descripcion = tk.Label(
            self.frame_inicio,
            text="Análisis sísmico profesional con Python",
            bg="#0f172a",
            fg="white",
            font=("Segoe UI", 16)
        )

        descripcion.pack(pady=10)

        boton = tk.Button(
            self.frame_inicio,
            text="INGRESAR AL SISTEMA",
            bg="#2563eb",
            fg="white",
            font=("Segoe UI", 14, "bold"),
            width=25,
            height=2,
            relief="flat",
            command=self.iniciar_sistema
        )

        boton.pack(pady=40)

    # =====================================================
    # INICIAR SISTEMA
    # =====================================================

    def iniciar_sistema(self):

        self.frame_inicio.destroy()

        self.crear_menu()
        self.crear_interfaz()
        self.cargar_csv()

    # =====================================================
    # MENU SUPERIOR
    # =====================================================

    def crear_menu(self):

        barra = tk.Menu(self.root)

        # =====================================================
        # MENU ARCHIVO
        # =====================================================

        archivo = tk.Menu(barra, tearoff=0)

        archivo.add_command(
            label="Exportar Excel",
            command=self.exportar_excel
        )

        archivo.add_separator()

        archivo.add_command(
            label="Salir",
            command=self.root.quit
        )

        barra.add_cascade(
            label="Archivo",
            menu=archivo
        )

        # =====================================================
        # MENU AYUDA
        # =====================================================

        ayuda = tk.Menu(barra, tearoff=0)

        ayuda.add_command(
            label="Acerca de",
            command=self.mostrar_acerca_de
        )

        barra.add_cascade(
            label="Ayuda",
            menu=ayuda
        )

        self.root.config(menu=barra)

    # =====================================================
    # VENTANA ACERCA DE
    # =====================================================

    def mostrar_acerca_de(self):

        ventana = tk.Toplevel(self.root)

        ventana.title("Acerca de")
        ventana.geometry("620x520")
        ventana.configure(bg="#111827")
        ventana.resizable(False, False)

        # =====================================================
        # TITULO
        # =====================================================

        titulo = tk.Label(
            ventana,
            text="🌎 SISTEMA PROFESIONAL DE TERREMOTOS",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 18, "bold")
        )

        titulo.pack(pady=10)

        # =====================================================
        # FRAME PRINCIPAL
        # =====================================================

        frame_principal = tk.Frame(
            ventana,
            bg="#111827"
        )

        frame_principal.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        # =====================================================
        # SCROLLBAR
        # =====================================================

        scrollbar = tk.Scrollbar(frame_principal)

        scrollbar.pack(
            side="right",
            fill="y"
        )

        # =====================================================
        # AREA DE TEXTO
        # =====================================================

        texto = tk.Text(
            frame_principal,
            wrap="word",
            yscrollcommand=scrollbar.set,
            bg="#1e293b",
            fg="white",
            font=("Segoe UI", 10),
            relief="flat",
            padx=15,
            pady=15
        )

        texto.pack(
            fill="both",
            expand=True
        )

        scrollbar.config(command=texto.yview)

        # =====================================================
        # INFORMACION
        # =====================================================

        informacion = """
👨‍💻 DESARROLLADOR

Nombre:
Yovani Andres Canche Robertos


🎓 CARRERA

Ciencia de Datos


🏫 ESCUELA

CBTIS 72


📊 DESCRIPCIÓN DEL PROYECTO

Este sistema profesional de terremotos
fue desarrollado utilizando Python,
Tkinter, Pandas y Matplotlib para el
análisis y visualización de datos sísmicos.

La aplicación permite cargar archivos CSV
con información sobre terremotos ocurridos
en distintas partes del mundo.

Además, el sistema cuenta con filtros
interactivos y gráficas dinámicas que
facilitan el análisis estadístico de
la información.


⚙ FUNCIONES PRINCIPALES

✔ Cargar archivos CSV automáticamente

✔ Eliminar datos nulos automáticamente

✔ Analizar datos sísmicos

✔ Filtrar terremotos por magnitud

✔ Mostrar tablas dinámicas

✔ Generar gráficas profesionales

✔ Visualizar mapas de terremotos

✔ Exportar reportes a Excel

✔ Mostrar estadísticas en tiempo real

✔ Interfaz moderna con tema oscuro


📈 GRÁFICAS DISPONIBLES

✔ Gráfica circular de magnitudes

✔ Top 10 terremotos más fuertes

✔ Mapa de dispersión (Scatter)

✔ Histograma de magnitudes


🖥 TECNOLOGÍAS UTILIZADAS

✔ Python

✔ Tkinter

✔ Pandas

✔ Matplotlib

✔ Openpyxl


🌎 OBJETIVO DEL SISTEMA

Facilitar el análisis y visualización
de datos sísmicos mediante una interfaz
gráfica moderna, intuitiva y fácil de usar.

El sistema fue diseñado para apoyar
proyectos académicos relacionados con
ciencia de datos y análisis estadístico.


📚 APLICACIONES

✔ Análisis de terremotos

✔ Visualización de datos

✔ Estadística descriptiva

✔ Ciencia de datos

✔ Proyectos escolares

✔ Investigación académica
"""

        texto.insert("1.0", informacion)

        texto.config(state="disabled")

        # =====================================================
        # BOTON CERRAR
        # =====================================================

        boton = tk.Button(
            ventana,
            text="Cerrar",
            bg="#2563eb",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            width=18,
            relief="flat",
            command=ventana.destroy
        )

        boton.pack(pady=10)

    # =====================================================
    # INTERFAZ PRINCIPAL
    # =====================================================

    def crear_interfaz(self):

        self.panel_izquierdo = tk.Frame(
            self.root,
            bg="#111827",
            width=220
        )

        self.panel_izquierdo.pack(
            side="left",
            fill="y",
            padx=5,
            pady=5
        )

        self.panel_izquierdo.pack_propagate(False)

        titulo_panel = tk.Label(
            self.panel_izquierdo,
            text="📊 PANEL",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 14, "bold")
        )

        titulo_panel.pack(pady=15)

        # =====================================================
        # FILTRO
        # =====================================================

        tk.Label(
            self.panel_izquierdo,
            text="Magnitud mínima",
            bg="#111827",
            fg="white",
            font=("Segoe UI", 10)
        ).pack(pady=3)

        self.entry_magnitud = ttk.Entry(
            self.panel_izquierdo,
            width=18
        )

        self.entry_magnitud.pack(pady=4)

        tk.Label(
            self.panel_izquierdo,
            text="Ejemplo: 5 o 6.5",
            bg="#111827",
            fg="gray",
            font=("Segoe UI", 8)
        ).pack()

        ttk.Button(
            self.panel_izquierdo,
            text="Aplicar",
            command=self.aplicar_filtros
        ).pack(pady=4)

        ttk.Button(
            self.panel_izquierdo,
            text="Limpiar",
            command=self.limpiar_filtros
        ).pack(pady=4)

        ttk.Button(
            self.panel_izquierdo,
            text="Exportar Excel",
            command=self.exportar_excel
        ).pack(pady=4)

        # =====================================================
        # BOTONES GRAFICAS
        # =====================================================

        tk.Label(
            self.panel_izquierdo,
            text="📈 GRÁFICAS",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 11, "bold")
        ).pack(pady=18)

        ttk.Button(
            self.panel_izquierdo,
            text="Gráfica Circular",
            command=self.grafica_pie
        ).pack(pady=3)

        ttk.Button(
            self.panel_izquierdo,
            text="Top 10 Terremotos",
            command=self.grafica_barras
        ).pack(pady=3)

        ttk.Button(
            self.panel_izquierdo,
            text="Mapa Scatter",
            command=self.grafica_scatter
        ).pack(pady=3)

        ttk.Button(
            self.panel_izquierdo,
            text="Histograma",
            command=self.grafica_histograma
        ).pack(pady=3)

        # =====================================================
        # PANEL DERECHO
        # =====================================================

        self.panel_derecho = tk.Frame(
            self.root,
            bg="#0f172a"
        )

        self.panel_derecho.pack(
            side="left",
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

        # =====================================================
        # TARJETAS
        # =====================================================

        tarjetas = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        tarjetas.pack(fill="x", pady=5)

        self.total_label = self.crear_tarjeta(
            tarjetas,
            "TOTAL"
        )

        self.max_label = self.crear_tarjeta(
            tarjetas,
            "MÁXIMA"
        )

        self.promedio_label = self.crear_tarjeta(
            tarjetas,
            "PROMEDIO"
        )

        # =====================================================
        # TABLA
        # =====================================================

        frame_tabla = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        frame_tabla.pack(
            fill="x",
            padx=5,
            pady=5
        )

        columnas = (
            "#",
            "Lugar",
            "Magnitud",
            "Fecha"
        )

        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings",
            height=7
        )

        anchos = [40, 520, 100, 180]

        for col, ancho in zip(columnas, anchos):

            self.tabla.heading(col, text=col)

            self.tabla.column(
                col,
                width=ancho
            )

        scrollbar = ttk.Scrollbar(
            frame_tabla,
            orient="vertical",
            command=self.tabla.yview
        )

        self.tabla.configure(
            yscrollcommand=scrollbar.set
        )

        self.tabla.pack(
            side="left",
            fill="x",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        # =====================================================
        # FRAME GRAFICAS
        # =====================================================

        self.frame_graficas = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        self.frame_graficas.pack(
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

    # =====================================================
    # TARJETAS
    # =====================================================

    def crear_tarjeta(self, parent, titulo):

        frame = tk.Frame(
            parent,
            bg="#111827",
            width=250,
            height=80
        )

        frame.pack(side="left", padx=8)

        frame.pack_propagate(False)

        tk.Label(
            frame,
            text=titulo,
            bg="#111827",
            fg="white",
            font=("Segoe UI", 10, "bold")
        ).pack(pady=6)

        valor = tk.Label(
            frame,
            text="0",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 20, "bold")
        )

        valor.pack()

        return valor

    # =====================================================
    # CARGAR CSV
    # =====================================================

    def cargar_csv(self):

        try:

            ruta = os.path.join(
                os.path.dirname(__file__),
                "earthquake_1995-2023.csv"
            )

            if not os.path.exists(ruta):

                messagebox.showerror(
                    "Error",
                    "No se encontró el archivo CSV"
                )

                return

            self.df = pd.read_csv(ruta)

            self.df.dropna(inplace=True)

            self.df_filtrado = self.df.copy()

            columnas = self.df.columns.tolist()

            self.col_lugar = columnas[0]
            self.col_magnitud = columnas[1]
            self.col_fecha = columnas[2]

            self.actualizar_tarjetas()
            self.cargar_tabla()

            messagebox.showinfo(
                "Éxito",
                "CSV cargado correctamente"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # =====================================================
    # ACTUALIZAR TARJETAS
    # =====================================================

    def actualizar_tarjetas(self):

        total = len(self.df_filtrado)

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        promedio = round(magnitudes.mean(), 2)
        maximo = round(magnitudes.max(), 2)

        self.total_label.config(text=str(total))
        self.promedio_label.config(text=str(promedio))
        self.max_label.config(text=str(maximo))

    # =====================================================
    # CARGAR TABLA
    # =====================================================

    def cargar_tabla(self):

        for item in self.tabla.get_children():
            self.tabla.delete(item)

        datos = self.df_filtrado.head(100)

        for i, (_, row) in enumerate(datos.iterrows(), start=1):

            self.tabla.insert(
                "",
                "end",
                values=(
                    i,
                    row[self.col_lugar],
                    row[self.col_magnitud],
                    row[self.col_fecha]
                )
            )

    # =====================================================
    # FILTROS
    # =====================================================

    def aplicar_filtros(self):

        try:

            self.df_filtrado = self.df.copy()

            magnitud = self.entry_magnitud.get()

            if magnitud != "":

                self.df_filtrado = self.df_filtrado[
                    pd.to_numeric(
                        self.df_filtrado[self.col_magnitud],
                        errors="coerce"
                    ) >= float(magnitud)
                ]

            self.actualizar_tarjetas()
            self.cargar_tabla()

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # =====================================================
    # LIMPIAR FILTROS
    # =====================================================

    def limpiar_filtros(self):

        self.entry_magnitud.delete(0, tk.END)

        self.df_filtrado = self.df.copy()

        self.actualizar_tarjetas()
        self.cargar_tabla()

    # =====================================================
    # EXPORTAR EXCEL
    # =====================================================

    def exportar_excel(self):

        try:

            ruta = "reportes/reporte_terremotos.xlsx"

            self.df_filtrado.to_excel(
                ruta,
                index=False,
                engine="openpyxl"
            )

            messagebox.showinfo(
                "Excel",
                "Reporte exportado correctamente"
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # =====================================================
    # LIMPIAR GRAFICAS
    # =====================================================

    def limpiar_graficas(self):

        for widget in self.frame_graficas.winfo_children():
            widget.destroy()

    # =====================================================
    # GRAFICA PIE
    # =====================================================

    def grafica_pie(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(12, 8))

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        categorias = pd.cut(
            magnitudes,
            bins=[0, 5, 7, 10],
            labels=["Moderados", "Fuertes", "Muy fuertes"]
        )

        datos = categorias.value_counts()

        ax.pie(
            datos,
            labels=datos.index,
            autopct="%1.1f%%",
            startangle=90,
            textprops={"fontsize": 15}
        )

        ax.set_title(
            "PORCENTAJE DE TERREMOTOS",
            fontsize=22,
            fontweight="bold"
        )

        canvas = FigureCanvasTkAgg(
            fig,
            master=self.frame_graficas
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

    # =====================================================
    # GRAFICA BARRAS
    # =====================================================

    def grafica_barras(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(16, 9))

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        datos = self.df_filtrado.copy()

        datos["magnitud_num"] = magnitudes

        top10 = datos.sort_values(
            by="magnitud_num",
            ascending=False
        ).head(10)

        nombres = []

        for nombre in top10[self.col_lugar]:

            texto = str(nombre)

            if len(texto) > 25:
                texto = texto[:25] + "..."

            nombres.append(texto)

        valores = top10["magnitud_num"]

        barras = ax.bar(
            nombres,
            valores,
            color="#2563eb",
            edgecolor="white",
            linewidth=2
        )

        for barra in barras:

            altura = barra.get_height()

            ax.text(
                barra.get_x() + barra.get_width()/2,
                altura + 0.05,
                f"{altura:.1f}",
                ha="center",
                fontsize=13,
                fontweight="bold"
            )

        ax.set_title(
            "TOP 10 TERREMOTOS MÁS FUERTES",
            fontsize=24,
            fontweight="bold"
        )

        ax.set_xlabel(
            "Lugar del terremoto",
            fontsize=15
        )

        ax.set_ylabel(
            "Magnitud",
            fontsize=15
        )

        plt.xticks(
            rotation=10,
            ha="right",
            fontsize=11
        )

        ax.grid(True, linestyle="--", alpha=0.3)

        plt.subplots_adjust(bottom=0.25)

        canvas = FigureCanvasTkAgg(
            fig,
            master=self.frame_graficas
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

    # =====================================================
    # GRAFICA SCATTER
    # =====================================================

    def grafica_scatter(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(15, 8))

        if "longitude" in self.df.columns and "latitude" in self.df.columns:

            scatter = ax.scatter(
                self.df_filtrado["longitude"],
                self.df_filtrado["latitude"],
                c=pd.to_numeric(
                    self.df_filtrado[self.col_magnitud],
                    errors="coerce"
                ),
                cmap="coolwarm",
                s=140,
                alpha=0.8,
                edgecolors="white"
            )

            barra = fig.colorbar(scatter)

            barra.set_label(
                "Magnitud",
                fontsize=13
            )

            ax.set_title(
                "MAPA DE TERREMOTOS",
                fontsize=24,
                fontweight="bold"
            )

            ax.set_xlabel(
                "Longitud",
                fontsize=15
            )

            ax.set_ylabel(
                "Latitud",
                fontsize=15
            )

            ax.grid(True)

        canvas = FigureCanvasTkAgg(
            fig,
            master=self.frame_graficas
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

    # =====================================================
    # HISTOGRAMA
    # =====================================================

    def grafica_histograma(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(15, 8))

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        ax.hist(
            magnitudes,
            bins=12,
            color="#22c55e",
            edgecolor="white"
        )

        ax.set_title(
            "DISTRIBUCIÓN DE MAGNITUDES",
            fontsize=24,
            fontweight="bold"
        )

        ax.set_xlabel(
            "Magnitud",
            fontsize=15
        )

        ax.set_ylabel(
            "Cantidad de terremotos",
            fontsize=15
        )

        ax.grid(True)

        canvas = FigureCanvasTkAgg(
            fig,
            master=self.frame_graficas
        )

        canvas.draw()

        canvas.get_tk_widget().pack(
            fill="both",
            expand=True
        )

# =====================================================
# EJECUTAR
# =====================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = SistemaTerremotos(root)

    root.mainloop()