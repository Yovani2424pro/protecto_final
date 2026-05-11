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
        self.root.title("Sistema Profesional de Terremotos")
        self.root.geometry("1450x850")
        self.root.configure(bg="#0f172a")

        # =====================================================
        # ESTILOS
        # =====================================================

        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#1e293b",
            foreground="white",
            fieldbackground="#1e293b",
            rowheight=28,
            font=("Segoe UI", 10)
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
    # PANTALLA DE INICIO
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
            font=("Segoe UI", 30, "bold")
        )

        titulo.pack(pady=100)

        descripcion = tk.Label(
            self.frame_inicio,
            text="Análisis de terremotos con Python",
            bg="#0f172a",
            fg="white",
            font=("Segoe UI", 15)
        )

        descripcion.pack(pady=10)

        boton = tk.Button(
            self.frame_inicio,
            text="INGRESAR AL SISTEMA",
            bg="#38bdf8",
            fg="black",
            font=("Segoe UI", 14, "bold"),
            width=22,
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
    # MENU
    # =====================================================

    def crear_menu(self):

        barra = tk.Menu(self.root)

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

        ayuda = tk.Menu(barra, tearoff=0)

        ayuda.add_command(
            label="Acerca de",
            command=lambda: messagebox.showinfo(
                "Información",
                "Sistema Profesional de Terremotos\nPython + Tkinter"
            )
        )

        barra.add_cascade(
            label="Ayuda",
            menu=ayuda
        )

        self.root.config(menu=barra)

    # =====================================================
    # INTERFAZ
    # =====================================================

    def crear_interfaz(self):

        # =====================================================
        # PANEL IZQUIERDO
        # =====================================================

        self.panel_izquierdo = tk.Frame(
            self.root,
            bg="#111827",
            width=240
        )

        self.panel_izquierdo.pack(
            side="left",
            fill="y"
        )

        self.panel_izquierdo.pack_propagate(False)

        titulo_panel = tk.Label(
            self.panel_izquierdo,
            text="📊 PANEL DE CONTROL",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 14, "bold")
        )

        titulo_panel.pack(pady=20)

        # =====================================================
        # FILTRO
        # =====================================================

        tk.Label(
            self.panel_izquierdo,
            text="Magnitud mínima",
            bg="#111827",
            fg="white",
            font=("Segoe UI", 11)
        ).pack(pady=5)

        self.entry_magnitud = ttk.Entry(
            self.panel_izquierdo,
            width=20
        )

        self.entry_magnitud.pack(pady=5)

        tk.Label(
            self.panel_izquierdo,
            text="Ejemplo: 5 o 6.5",
            bg="#111827",
            fg="gray",
            font=("Segoe UI", 9)
        ).pack()

        ttk.Button(
            self.panel_izquierdo,
            text="Aplicar Filtro",
            command=self.aplicar_filtros
        ).pack(pady=5)

        ttk.Button(
            self.panel_izquierdo,
            text="Limpiar Filtros",
            command=self.limpiar_filtros
        ).pack(pady=5)

        ttk.Button(
            self.panel_izquierdo,
            text="Exportar Excel",
            command=self.exportar_excel
        ).pack(pady=5)

        # =====================================================
        # BOTONES GRAFICAS
        # =====================================================

        tk.Label(
            self.panel_izquierdo,
            text="📈 GENERAR GRÁFICAS",
            bg="#111827",
            fg="#38bdf8",
            font=("Segoe UI", 12, "bold")
        ).pack(pady=20)

        ttk.Button(
            self.panel_izquierdo,
            text="Gráfica Circular",
            command=self.grafica_pie
        ).pack(pady=4)

        ttk.Button(
            self.panel_izquierdo,
            text="Top 10 Terremotos",
            command=self.grafica_barras
        ).pack(pady=4)

        ttk.Button(
            self.panel_izquierdo,
            text="Mapa Scatter",
            command=self.grafica_scatter
        ).pack(pady=4)

        ttk.Button(
            self.panel_izquierdo,
            text="Histograma",
            command=self.grafica_histograma
        ).pack(pady=4)

        # =====================================================
        # PANEL DERECHO
        # =====================================================

        self.panel_derecho = tk.Frame(
            self.root,
            bg="#0f172a"
        )

        self.panel_derecho.pack(
            side="right",
            fill="both",
            expand=True
        )

        # =====================================================
        # TARJETAS
        # =====================================================

        tarjetas = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        tarjetas.pack(fill="x", pady=10)

        self.total_label = self.crear_tarjeta(
            tarjetas,
            "Total Registros"
        )

        self.max_label = self.crear_tarjeta(
            tarjetas,
            "Magnitud Máxima"
        )

        self.promedio_label = self.crear_tarjeta(
            tarjetas,
            "Promedio"
        )

        # =====================================================
        # TABLA
        # =====================================================

        frame_tabla = tk.Frame(
            self.panel_derecho,
            bg="#0f172a"
        )

        frame_tabla.pack(
            fill="both",
            padx=10,
            pady=5
        )

        columnas = (
            "Lugar",
            "Magnitud",
            "Fecha",
            "Profundidad"
        )

        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=columnas,
            show="headings",
            height=6
        )

        for col in columnas:

            self.tabla.heading(col, text=col)

            self.tabla.column(
                col,
                width=240
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
            fill="both",
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
            bg="#0f172a",
            height=500
        )

        self.frame_graficas.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        self.frame_graficas.pack_propagate(False)

    # =====================================================
    # TARJETAS
    # =====================================================

    def crear_tarjeta(self, parent, titulo):

        frame = tk.Frame(
            parent,
            bg="#1e293b",
            width=220,
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
        ).pack(pady=8)

        valor = tk.Label(
            frame,
            text="0",
            bg="#1e293b",
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
            self.col_profundidad = columnas[3]

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
    # TARJETAS
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
    # TABLA
    # =====================================================

    def cargar_tabla(self):

        for item in self.tabla.get_children():
            self.tabla.delete(item)

        datos = self.df_filtrado.head(100)

        for _, row in datos.iterrows():

            self.tabla.insert(
                "",
                "end",
                values=(
                    row[self.col_lugar],
                    row[self.col_magnitud],
                    row[self.col_fecha],
                    row[self.col_profundidad]
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

        fig, ax = plt.subplots(figsize=(11, 5))

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        categorias = pd.cut(
            magnitudes,
            bins=[0, 5, 7, 10],
            labels=[
                "Moderados",
                "Fuertes",
                "Muy fuertes"
            ]
        )

        datos = categorias.value_counts()

        colores = [
            "#22c55e",
            "#facc15",
            "#ef4444"
        ]

        ax.pie(
            datos,
            labels=datos.index,
            autopct="%1.1f%%",
            colors=colores,
            startangle=90
        )

        ax.set_title(
            "Porcentaje de Terremotos",
            fontsize=18,
            fontweight="bold"
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

    # =====================================================
    # GRAFICA BARRAS MEJORADA
    # =====================================================

    def grafica_barras(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(13, 6))

        fig.patch.set_facecolor("#0f172a")

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

        nombres = top10[self.col_lugar]

        valores = top10["magnitud_num"]

        barras = ax.bar(
            nombres,
            valores,
            color="#38bdf8",
            edgecolor="white",
            linewidth=2
        )

        for barra in barras:

            altura = barra.get_height()

            ax.text(
                barra.get_x() + barra.get_width()/2,
                altura + 0.1,
                f"{altura:.1f}",
                ha="center",
                fontsize=10,
                color="white",
                fontweight="bold"
            )

        ax.set_title(
            "TOP 10 TERREMOTOS MÁS FUERTES",
            fontsize=20,
            fontweight="bold",
            pad=20
        )

        ax.set_xlabel(
            "Lugar del terremoto",
            fontsize=13
        )

        ax.set_ylabel(
            "Magnitud",
            fontsize=13
        )

        plt.xticks(
            rotation=25,
            ha="right",
            fontsize=9
        )

        ax.grid(
            True,
            linestyle="--",
            alpha=0.3
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

    # =====================================================
    # GRAFICA SCATTER
    # =====================================================

    def grafica_scatter(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(12, 5))

        if "longitude" in self.df.columns and "latitude" in self.df.columns:

            scatter = ax.scatter(
                self.df_filtrado["longitude"],
                self.df_filtrado["latitude"],
                c=pd.to_numeric(
                    self.df_filtrado[self.col_magnitud],
                    errors="coerce"
                ),
                cmap="coolwarm",
                s=90,
                alpha=0.8,
                edgecolors="white"
            )

            barra = fig.colorbar(scatter)

            barra.set_label("Magnitud")

            ax.set_title(
                "Mapa Mundial de Terremotos",
                fontsize=18,
                fontweight="bold"
            )

            ax.set_xlabel("Longitud")
            ax.set_ylabel("Latitud")

            ax.grid(True)

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

    # =====================================================
    # HISTOGRAMA
    # =====================================================

    def grafica_histograma(self):

        self.limpiar_graficas()

        plt.style.use("dark_background")

        fig, ax = plt.subplots(figsize=(11, 5))

        magnitudes = pd.to_numeric(
            self.df_filtrado[self.col_magnitud],
            errors="coerce"
        )

        ax.hist(
            magnitudes,
            bins=12,
            color="#38bdf8",
            edgecolor="white"
        )

        ax.set_title(
            "Distribución de Magnitudes",
            fontsize=18,
            fontweight="bold"
        )

        ax.set_xlabel("Magnitud")
        ax.set_ylabel("Cantidad")

        ax.grid(True)

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

# =====================================================
# EJECUTAR PROGRAMA
# =====================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = SistemaTerremotos(root)

    root.mainloop()