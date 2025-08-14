import unicodedata
import customtkinter as ctk
import json
import os
import customtkinter as ctk
import tkinter as tk
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from tkinter import Tk, Toplevel, filedialog, messagebox, simpledialog

# ==== CONFIGURACIÓN ====
CONFIG_FILE = "config.json"
LOGO_PATH = "logoecolite.jpg"  # Logo opcional para los gráficos

ventana_cargando = None

def mostrar_cargando():
    global ventana_cargando
    ventana_cargando = ctk.CTkToplevel(app)
    ventana_cargando.title("Cargando...")
    ventana_cargando.geometry("300x100")

    # Obtener dimensiones y posición de la ventana principal
    app.update_idletasks()
    app_width = app.winfo_width()
    app_height = app.winfo_height()
    app_x = app.winfo_x()
    app_y = app.winfo_y()

    # Calcular coordenadas para centrar
    pos_x = app_x + (app_width // 2) - (300 // 2)
    pos_y = app_y + (app_height // 2) - (100 // 2)

    ventana_cargando.geometry(f"300x100+{pos_x}+{pos_y}")

    ventana_cargando.transient(app)
    ventana_cargando.grab_set()

    label = ctk.CTkLabel(ventana_cargando, text="Generando reporte, por favor espere...")
    label.pack(pady=10)

    progress = ctk.CTkProgressBar(ventana_cargando, mode="indeterminate")
    progress.pack(pady=10, fill="x", padx=20)
    progress.start()

    ventana_cargando.update()

def cerrar_cargando():
    global ventana_cargando
    if ventana_cargando is not None:
        ventana_cargando.destroy()
        ventana_cargando = None

# ==== Funciones para manejar configuración ====
def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"excel_path": ""}

def guardar_config(ruta_excel):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"excel_path": ruta_excel}, f)

# ==== Función para subir/cambiar archivo ====
def subir_excel():
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo de Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if ruta:
        guardar_config(ruta)
        config["excel_path"] = ruta
        label_archivo.configure(text=f"📂 {os.path.basename(ruta)}")
        messagebox.showinfo("Archivo cargado", "Archivo Excel guardado correctamente.")

# ==== Función para mostrar descripción ====
def mostrar_descripcion(accion, descripcion):
    global accion_seleccionada
    accion_seleccionada = accion

    for widget in frame_der.winfo_children():
        widget.destroy()

    titulo = ctk.CTkLabel(frame_der, text="Acción seleccionada", font=("Segoe UI", 24, "bold"), text_color="#333")
    titulo.pack(pady=(40, 10))

    desc = ctk.CTkLabel(frame_der, text=descripcion, font=("Segoe UI", 14), wraplength=450, justify="left", text_color="#555")
    desc.pack(pady=10)

    btn_ejecutar = ctk.CTkButton(
        frame_der, 
        text="▶ Ejecutar acción", 
        height=45, 
        width=220, 
        font=("Segoe UI", 15, "bold"),
        fg_color="#0A84FF",
        hover_color="#0066CC",
        text_color="white",
        command=ejecutar_accion
    )
    btn_ejecutar.pack(pady=40)

# ==== Función para ejecutar acción seleccionada ====
def ejecutar_accion():
    if not config["excel_path"]:
        messagebox.showwarning("Archivo no cargado", "Por favor, sube un archivo Excel antes de ejecutar acciones.")
        return
    
    if accion_seleccionada == "volumen_vendedores":
        generar_vendedores(config["excel_path"])
        
    elif accion_seleccionada == "margen_vendedores":
        generar_margen_vendedores(config["excel_path"])
        
    elif accion_seleccionada == "comparativo_vendedor":
        generar_comparativo_vendedor(config["excel_path"])

    elif accion_seleccionada == "departamentos_vendedor":
        generar_departamentos_vendedor(config["excel_path"])

    elif accion_seleccionada == "cuidades_vendedor":
        generar_ciudades_vendedor(config["excel_path"])

    elif accion_seleccionada == "comparativo_cuidad":
        generar_comparativo_ciudad(config["excel_path"])

    elif accion_seleccionada == "comparativo_departamento":
        generar_comparativo_departamento(config["excel_path"])

    elif accion_seleccionada == "margen_productos":
        generar_margen_productos(config["excel_path"])

    elif accion_seleccionada == "producto_volumen_margen":
        generar_producto_volumen_margen(config["excel_path"])

    elif accion_seleccionada == "reporte_cuidades":
        generar_ciudades(config["excel_path"])

    elif accion_seleccionada == "reporte_departamentos":
        generar_departamentos(config["excel_path"])

    elif accion_seleccionada == "comparativo_linea":
        generar_comparativo_linea(config["excel_path"])

    elif accion_seleccionada == "rotacion_inventario":
        generar_rotacion_inventario(config["excel_path"])

    elif accion_seleccionada == "ventas_semana":
        generar_ventas_semana(config["excel_path"])

    elif accion_seleccionada == "presupuesto_año":
        generar_presupuesto_año(config["excel_path"])

    else:
        messagebox.showinfo("Acción no implementada", f"La acción '{accion_seleccionada}' todavía no tiene lógica asociada.")

def generar_vendedores(archivo_excel):
    try:
        # Pedir año con parent=app
        anio = simpledialog.askstring(
            "Filtrar por año",
            "¿Qué año deseas filtrar? (Ejemplo: 2025):",
            parent=app
        )
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.", parent=app)
            return
            
        anio = int(anio)

        mostrar_cargando()

        # Vendedores a excluir
        excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]

        # Cargar datos
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['NETO'] = pd.to_numeric(df['NETO'], errors='coerce')

        # Filtrar
        df_filtrado = df[df['FECHA'].dt.year == anio]
        df_filtrado = df_filtrado[~df_filtrado['VENDEDOR'].isin(excluir)]

        if df_filtrado.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.", parent=app)
            return

        # Agrupar y ordenar
        reporte = df_filtrado.groupby('VENDEDOR').agg(NETO=('NETO', 'sum')).reset_index()
        reporte = reporte.sort_values(by='NETO', ascending=False)

        # Calcular total
        total_ventas = reporte["NETO"].sum()

        # === Crear ventana hija para el gráfico ===
        top = tk.Toplevel(app)
        top.title(f"Ventas por Vendedor - {anio}")
        top.attributes("-topmost", True)  # Mantener al frente
        top.focus_force()

        # Graficar en esa ventana 
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(reporte['VENDEDOR'], reporte['NETO'], color="lightsteelblue")
        ax.invert_yaxis()
        ax.set_title(f"Ventas por Vendedor - {anio} | Total: ${total_ventas:,.0f}", 
                     fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor vendido ($)")

        for i, v in enumerate(reporte["NETO"]):
            ax.text(v + (max(reporte["NETO"]) * 0.01), i, f"${v:,.0f}", va="center", fontsize=9, color="black")
 
        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.0, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(0, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # Incrustar matplotlib en Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte:\n{e}", parent=app)
        
    cerrar_cargando()

def generar_margen_vendedores(archivo_excel):
    try:
        # Pedir año
        anio = simpledialog.askstring(
            "Filtrar por año",
            "¿Qué año deseas filtrar? (Ejemplo: 2025):",
            parent=app
        )
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.", parent=app)
            return
        anio = int(anio)
        mostrar_cargando()
        # Vendedores a excluir
        excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor"]

        # Cargar datos
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['NETO'] = pd.to_numeric(df['NETO'], errors='coerce')
        df['COST'] = pd.to_numeric(df['COST'], errors='coerce')
        df['QTYSHIP'] = pd.to_numeric(df['QTYSHIP'], errors='coerce')

        # Filtrar
        df_filtrado = df[df['FECHA'].dt.year == anio]
        df_filtrado = df_filtrado[~df_filtrado['VENDEDOR'].isin(excluir)]
        df_filtrado['COSTO_TOTAL'] = df_filtrado['COST'] * df_filtrado['QTYSHIP']

        if df_filtrado.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.", parent=app)
            return

        # Agrupar y calcular margen
        reporte = (
            df_filtrado.groupby('VENDEDOR')
            .agg(
                NETO=('NETO', 'sum'),
                COSTO_TOTAL=('COSTO_TOTAL', 'sum')
            )
            .reset_index()
        )
        reporte['MARGEN_%'] = ((reporte['NETO'] - reporte['COSTO_TOTAL']) / reporte['NETO']) * 100
        reporte = reporte.sort_values(by='MARGEN_%', ascending=False)

        # === Crear ventana hija para el gráfico ===
        top = tk.Toplevel(app)
        top.title(f"Margen por Vendedor (%) - {anio}")
        top.attributes("-topmost", True)  # Mantener al frente
        top.focus_force()

        # Graficar
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(reporte['VENDEDOR'], reporte['MARGEN_%'], color="orange")
        ax.invert_yaxis()
        ax.set_title(f"Margen por Vendedor (%) - {anio}", fontsize=14, fontweight="bold")
        ax.set_xlabel("Margen (%)")

        # Etiquetas
        for i, v in enumerate(reporte["MARGEN_%"]):
            ax.text(v + 0.5, i, f"{v:.1f}%", va="center", fontsize=9, color="black")

        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.0, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(0, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # Incrustar matplotlib en Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte de margen:\n{e}", parent=app)

def generar_departamentos_vendedor(archivo_excel):
    try:
        anio = simpledialog.askstring("Año", "Ingrese el año (Ejemplo: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        top_n = simpledialog.askstring("Top", "Ingrese la cantidad de vendedores en el top:", parent=app)
        if not top_n or not top_n.isdigit():
            messagebox.showerror("Error", "Cantidad inválida.")
            return
        top_n = int(top_n)
        mostrar_cargando()

        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        df_filtro = df[df["FECHA"].dt.year == anio]
        if df_filtro.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.")
            return

        resumen = df_filtro.groupby(["DEPARTAMENTO", "VENDEDOR"])["NETO"].sum().reset_index()
        resumen = resumen.sort_values(by="NETO", ascending=False).head(top_n)

        total_top = resumen["NETO"].sum()
        titulo = f"Top {top_n} vendedores por departamento - {anio} (${total_top:,.0f})"

        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(resumen["VENDEDOR"] + " - " + resumen["DEPARTAMENTO"], resumen["NETO"], color="skyblue")
        ax.invert_yaxis()
        ax.set_title(titulo, fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor Total ($)")

        for bar in bars:
            width = bar.get_width()
            ax.text(width + (width * 0.01), bar.get_y() + bar.get_height()/2, f"${width:,.0f}", va="center", fontsize=9)

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.02, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        top = tk.Toplevel(app)
        top.title(titulo)
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte por departamento:\n{e}")

def generar_ciudades_vendedor(archivo_excel):
    try:
        anio = simpledialog.askstring("Año", "Ingrese el año (Ejemplo: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        top_n = simpledialog.askstring("Top", "Ingrese la cantidad de vendedores en el top:", parent=app)
        if not top_n or not top_n.isdigit():
            messagebox.showerror("Error", "Cantidad inválida.")
            return
        top_n = int(top_n)
        mostrar_cargando()

        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        df_filtro = df[df["FECHA"].dt.year == anio]
        if df_filtro.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.")
            return

        resumen = df_filtro.groupby(["CITY", "VENDEDOR"])["NETO"].sum().reset_index()
        resumen = resumen.sort_values(by="NETO", ascending=False).head(top_n)

        total_top = resumen["NETO"].sum()
        titulo = f"Top {top_n} vendedores por ciudad - {anio} (${total_top:,.0f})"

        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(resumen["VENDEDOR"] + " - " + resumen["CITY"], resumen["NETO"], color="lightgreen")
        ax.invert_yaxis()
        ax.set_title(titulo, fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor Total ($)")

        top_win = tk.Toplevel(app)
        top_win.title(titulo)
        top_win.attributes("-topmost", True)
        top_win.lift()
        top_win.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top_win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        
        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top_win)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        for bar in bars:
            width = bar.get_width()
            ax.text(width + (width * 0.01), bar.get_y() + bar.get_height()/2, f"${width:,.0f}", va="center", fontsize=9)

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.02, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte por ciudad:\n{e}")

def generar_comparativo_vendedor(archivo_excel):
    try:
        mostrar_cargando()

        # === Fechas YTD ===
        hoy = datetime.today()
        fecha_inicio_actual = datetime(hoy.year, 1, 1)
        fecha_fin_actual = hoy

        fecha_inicio_anterior = fecha_inicio_actual.replace(year=fecha_inicio_actual.year - 1)
        fecha_fin_anterior = fecha_fin_actual.replace(year=fecha_fin_actual.year - 1)

        # === Lista de vendedores a excluir ===
        excluir = ["ARANGO JULIO CESAR", "LOPEZ GAITAN JORGE HERNAN", "Sin vendedor", "INACTIVO",
                   "ORDOÑEZ RIVERA ANDRES FELIPE", "AVENDAÑO ANDRES"]

        # === Cargar datos ===
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        # === Excluir vendedores ===
        df = df[~df["VENDEDOR"].isin(excluir)]

        # === Filtrar solo dos rangos YTD ===
        df_actual = df[(df["FECHA"] >= fecha_inicio_actual) & (df["FECHA"] <= fecha_fin_actual)]
        df_anterior = df[(df["FECHA"] >= fecha_inicio_anterior) & (df["FECHA"] <= fecha_fin_anterior)]

        # === Agrupar por vendedor ===
        ventas_actual = df_actual.groupby("VENDEDOR")["NETO"].sum().reset_index()
        ventas_anterior = df_anterior.groupby("VENDEDOR")["NETO"].sum().reset_index()

        # === Unir y calcular diferencia ===
        comparativo = pd.merge(
            ventas_anterior,
            ventas_actual,
            on="VENDEDOR",
            how="outer",
            suffixes=(f"_{fecha_inicio_anterior.year}", f"_{fecha_inicio_actual.year}")
        ).fillna(0)

        comparativo["DIFERENCIA"] = (
            comparativo[f"NETO_{fecha_inicio_actual.year}"] -
            comparativo[f"NETO_{fecha_inicio_anterior.year}"]
        )

        # === Ordenar para gráfico (menor a mayor para que el barh quede bien visualmente) ===
        comparativo = comparativo.sort_values(by=f"NETO_{fecha_inicio_actual.year}", ascending=True)

        # === Gráfico horizontal ===
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(12, 7))

        y_pos = range(len(comparativo))
        bar_height = 0.4

        ax.barh([y + bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_anterior.year}"],
                height=bar_height, label=str(fecha_inicio_anterior.year), color="#4F81BD")
        ax.barh([y - bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_actual.year}"],
                height=bar_height, label=str(fecha_inicio_actual.year), color="#F79646")

        # Etiquetas con valores
        for i, (val_ant, val_act) in enumerate(zip(
            comparativo[f"NETO_{fecha_inicio_anterior.year}"],
            comparativo[f"NETO_{fecha_inicio_actual.year}"]
        )):
            ax.text(val_ant + (val_ant * 0.01), i + bar_height/2, f"${val_ant:,.0f}",
                    va="center", fontsize=8)
            ax.text(val_act + (val_act * 0.01), i - bar_height/2, f"${val_act:,.0f}",
                    va="center", fontsize=8)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(comparativo["VENDEDOR"])
        ax.set_xlabel("Ventas ($)")
        ax.set_title(f"Comparativo YTD de Ventas por Vendedor\n"
                     f"{fecha_inicio_anterior.date()} a {fecha_fin_anterior.date()} vs "
                     f"{fecha_inicio_actual.date()} a {fecha_fin_actual.date()}",
                     fontsize=14, fontweight="bold")
        ax.legend()

        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (1.05, 1.05), frameon=False,
                                xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # === Mostrar en ventana ===
        top = tk.Toplevel(app)
        top.title("Comparativo YTD de Ventas por Vendedor")
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)


        cerrar_cargando()

        # === Guardar Excel ordenado de MAYOR a menor ===
        ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Archivos Excel", "*.xlsx")],
                                                    title="Guardar comparativo como")
        if ruta_guardado:
            comparativo_excel = comparativo.sort_values(
                by=f"NETO_{fecha_inicio_actual.year}", ascending=False
            )
            comparativo_excel.to_excel(ruta_guardado, index=False)

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error generando el comparativo:\n{e}")

def generar_comparativo_ciudad(archivo_excel):
    try:
        mostrar_cargando()

        # === Fechas YTD ===
        hoy = datetime.today()
        fecha_inicio_actual = datetime(hoy.year, 1, 1)
        fecha_fin_actual = hoy

        fecha_inicio_anterior = fecha_inicio_actual.replace(year=fecha_inicio_actual.year - 1)
        fecha_fin_anterior = fecha_fin_actual.replace(year=fecha_fin_actual.year - 1)

        # === Cargar datos ===
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        # === Filtrar solo dos rangos YTD ===
        df_actual = df[(df["FECHA"] >= fecha_inicio_actual) & (df["FECHA"] <= fecha_fin_actual)]
        df_anterior = df[(df["FECHA"] >= fecha_inicio_anterior) & (df["FECHA"] <= fecha_fin_anterior)]

        # === Agrupar por ciudad ===
        ventas_actual = df_actual.groupby("CITY")["NETO"].sum().reset_index()
        ventas_anterior = df_anterior.groupby("CITY")["NETO"].sum().reset_index()

        # === Unir y calcular diferencia ===
        comparativo = pd.merge(
            ventas_anterior,
            ventas_actual,
            on="CITY",
            how="outer",
            suffixes=(f"_{fecha_inicio_anterior.year}", f"_{fecha_inicio_actual.year}")
        ).fillna(0)

        # === Quitar ciudades con valor 0 en ambos años ===
        comparativo = comparativo[
            (comparativo[f"NETO_{fecha_inicio_actual.year}"] > 0) |
            (comparativo[f"NETO_{fecha_inicio_anterior.year}"] > 0)
        ]

        # === Pareto 70% sobre el año actual ===
        comparativo = comparativo.sort_values(by=f"NETO_{fecha_inicio_actual.year}", ascending=False)
        comparativo["ACUM_PCT"] = (comparativo[f"NETO_{fecha_inicio_actual.year}"].cumsum() /
                                   comparativo[f"NETO_{fecha_inicio_actual.year}"].sum())
        comparativo = comparativo[comparativo["ACUM_PCT"] <= 0.7]

        # === Calcular diferencia ===
        comparativo["DIFERENCIA"] = (
            comparativo[f"NETO_{fecha_inicio_actual.year}"] -
            comparativo[f"NETO_{fecha_inicio_anterior.year}"]
        )

        # === Ordenar por ventas año actual (descendente) ===
        comparativo = comparativo.sort_values(by=f"NETO_{fecha_inicio_actual.year}", ascending=False)

        # === Gráfico horizontal ===
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(12, 7))

        y_pos = range(len(comparativo))
        bar_height = 0.4

        ax.barh([y + bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_anterior.year}"],
                height=bar_height, label=str(fecha_inicio_anterior.year), color="#4F81BD")
        ax.barh([y - bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_actual.year}"],
                height=bar_height, label=str(fecha_inicio_actual.year), color="#F79646")

        for i, (val_ant, val_act) in enumerate(zip(
            comparativo[f"NETO_{fecha_inicio_anterior.year}"],
            comparativo[f"NETO_{fecha_inicio_actual.year}"]
        )):
            ax.text(val_ant + (val_ant * 0.01), i + bar_height/2, f"${val_ant:,.0f}",
                    va="center", fontsize=8)
            ax.text(val_act + (val_act * 0.01), i - bar_height/2, f"${val_act:,.0f}",
                    va="center", fontsize=8)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(comparativo["CITY"])
        ax.set_xlabel("Ventas ($)")
        ax.set_title(f"Comparativo YTD de Ventas por Ciudad (Pareto 70%)\n"
                     f"{fecha_inicio_anterior.date()} a {fecha_fin_anterior.date()} vs "
                     f"{fecha_inicio_actual.date()} a {fecha_fin_actual.date()}",
                     fontsize=14, fontweight="bold")
        ax.legend()
        ax.invert_yaxis()  # Para que el mayor quede arriba

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (1.05, 1.05), frameon=False,
                                xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        top = tk.Toplevel(app)
        top.title("Comparativo YTD de Ventas por Ciudad (Pareto 70%)")
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

        ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Archivos Excel", "*.xlsx")],
                                                    title="Guardar comparativo como")
        if ruta_guardado:
            comparativo.to_excel(ruta_guardado, index=False)

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error generando el comparativo:\n{e}")

def generar_comparativo_departamento(archivo_excel):
    try:
        mostrar_cargando()

        hoy = datetime.today()
        fecha_inicio_actual = datetime(hoy.year, 1, 1)
        fecha_fin_actual = hoy

        fecha_inicio_anterior = fecha_inicio_actual.replace(year=fecha_inicio_actual.year - 1)
        fecha_fin_anterior = fecha_fin_actual.replace(year=fecha_fin_actual.year - 1)

        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        df_actual = df[(df["FECHA"] >= fecha_inicio_actual) & (df["FECHA"] <= fecha_fin_actual)]
        df_anterior = df[(df["FECHA"] >= fecha_inicio_anterior) & (df["FECHA"] <= fecha_fin_anterior)]

        ventas_actual = df_actual.groupby("DEPARTAMENTO")["NETO"].sum().reset_index()
        ventas_anterior = df_anterior.groupby("DEPARTAMENTO")["NETO"].sum().reset_index()

        comparativo = pd.merge(
            ventas_anterior,
            ventas_actual,
            on="DEPARTAMENTO",
            how="outer",
            suffixes=(f"_{fecha_inicio_anterior.year}", f"_{fecha_inicio_actual.year}")
        ).fillna(0)

        # Filtrar: eliminar filas donde cualquiera de los dos años sea 0
        col_ant = f"NETO_{fecha_inicio_anterior.year}"
        col_act = f"NETO_{fecha_inicio_actual.year}"
        comparativo = comparativo[(comparativo[col_ant] != 0) & (comparativo[col_act] != 0)]

        # Calcular diferencia
        comparativo["DIFERENCIA"] = comparativo[col_act] - comparativo[col_ant]

        # Ordenar por ventas actuales de mayor a menor
        comparativo = comparativo.sort_values(by=col_act, ascending=False)

        # Gráfico
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(12, 7))

        y_pos = range(len(comparativo))
        bar_height = 0.4

        ax.barh([y + bar_height/2 for y in y_pos], comparativo[col_ant],
                height=bar_height, label=str(fecha_inicio_anterior.year), color="#4F81BD")
        ax.barh([y - bar_height/2 for y in y_pos], comparativo[col_act],
                height=bar_height, label=str(fecha_inicio_actual.year), color="#F79646")

        # Etiquetas de valores
        for i, (val_ant, val_act) in enumerate(zip(comparativo[col_ant], comparativo[col_act])):
            ax.text(val_ant + (val_ant * 0.01), i + bar_height/2, f"${val_ant:,.0f}",
                    va="center", fontsize=8)
            ax.text(val_act + (val_act * 0.01), i - bar_height/2, f"${val_act:,.0f}",
                    va="center", fontsize=8)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(comparativo["DEPARTAMENTO"])
        ax.invert_yaxis()  # Mantener mayor arriba
        ax.set_xlabel("Ventas ($)")
        ax.set_title(f"Comparativo YTD de Ventas por Departamento\n"
                     f"{fecha_inicio_anterior.date()} a {fecha_fin_anterior.date()} vs "
                     f"{fecha_inicio_actual.date()} a {fecha_fin_actual.date()}",
                     fontsize=14, fontweight="bold")
        ax.legend()

        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (1.05, 1.05), frameon=False,
                                xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # Mostrar en ventana
        top = tk.Toplevel(app)
        top.title("Comparativo YTD de Ventas por Departamento")
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

        # Guardar archivo
        ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Archivos Excel", "*.xlsx")],
                                                    title="Guardar comparativo como")
        if ruta_guardado:
            comparativo.to_excel(ruta_guardado, index=False)

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error generando el comparativo:\n{e}")

def generar_margen_productos(archivo_excel):
    try:
        # === Pedir datos al usuario ===
        anio = simpledialog.askstring("Año", "Ingrese el año (Ejemplo: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        mes_inicio = simpledialog.askstring("Mes inicio", "Ingrese el mes de inicio (1-12):", parent=app)
        if not mes_inicio or not mes_inicio.isdigit():
            messagebox.showerror("Error", "Mes de inicio inválido.")
            return
        mes_inicio = int(mes_inicio)

        dia_inicio = simpledialog.askstring("Día inicio", "Ingrese el día de inicio (1-31):", parent=app)
        if not dia_inicio or not dia_inicio.isdigit():
            messagebox.showerror("Error", "Día de inicio inválido.")
            return
        dia_inicio = int(dia_inicio)

        mes_fin = simpledialog.askstring("Mes fin", "Ingrese el mes de fin (1-12):", parent=app)
        if not mes_fin or not mes_fin.isdigit():
            messagebox.showerror("Error", "Mes de fin inválido.")
            return
        mes_fin = int(mes_fin)

        dia_fin = simpledialog.askstring("Día fin", "Ingrese el día de fin (1-31):", parent=app)
        if not dia_fin or not dia_fin.isdigit():
            messagebox.showerror("Error", "Día de fin inválido.")
            return
        dia_fin = int(dia_fin)

        top_n = simpledialog.askstring("Cantidad", "Ingrese la cantidad de productos en el top:", parent=app)
        if not top_n or not top_n.isdigit():
            messagebox.showerror("Error", "Cantidad inválida.")
            return
        top_n = int(top_n)
        mostrar_cargando()

        fecha_inicio = datetime(anio, mes_inicio, dia_inicio)
        fecha_fin = datetime(anio, mes_fin, dia_fin)

        # === Cargar datos ===
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")
        df["COST"] = pd.to_numeric(df["COST"], errors="coerce")
        df["QTYSHIP"] = pd.to_numeric(df["QTYSHIP"], errors="coerce")

        # === Filtrar fechas ===
        df_filtro = df[(df["FECHA"] >= fecha_inicio) & (df["FECHA"] <= fecha_fin)].copy()
        df_filtro["ITEM"] = df_filtro["ITEM"].astype(str).str.strip().str.upper()

        # === Excluir palabras clave ===
        excluir_keywords = ["FLETE", "MANTENIMIENTO", "DF-DISEÑO", "DISEÑO"]
        patron_excluir = "|".join(excluir_keywords)
        df_filtro = df_filtro[~df_filtro["ITEM"].str.contains(patron_excluir, case=False, na=False, regex=True)]

        if df_filtro.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el rango {fecha_inicio.date()} a {fecha_fin.date()}.")
            return

        # === Calcular margen ===
        df_filtro["COSTO_TOTAL"] = df_filtro["COST"] * df_filtro["QTYSHIP"]
        df_filtro["MARGEN_%"] = ((df_filtro["NETO"] - df_filtro["COSTO_TOTAL"]) / df_filtro["NETO"]) * 100

        # === Agrupar por producto ===
        top_df = df_filtro.groupby("ITEM")["MARGEN_%"].mean().reset_index()
        top_df = top_df.sort_values(by="MARGEN_%", ascending=False).head(top_n)

        # === Graficar ===
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(top_df["ITEM"], top_df["MARGEN_%"], color="steelblue")
        ax.invert_yaxis()
        ax.set_title(f"Top {top_n} productos por rentabilidad (%)\n{fecha_inicio.date()} a {fecha_fin.date()}",
                     fontsize=14, fontweight="bold")
        ax.set_xlabel("Margen (%)")

        # Etiquetas
        for bar in bars:
            width = bar.get_width()
            ax.text(width + 0.5, bar.get_y() + bar.get_height()/2, f"{width:.2f}%", va="center", fontsize=12)

        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.02, 1.10), frameon=False,
                                 xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # === Crear ventana hija para el gráfico ===
        top = tk.Toplevel(app)
        top.title(f"Top {top_n} productos por rentabilidad (%) - {fecha_inicio.date()} a {fecha_fin.date()}")
        top.attributes("-topmost", True)
        top.lift()
        top.focus_force()
        top.after_idle(lambda: top.attributes("-topmost", True))

        # Incrustar matplotlib en Tkinter
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte de margen por producto:\n{e}")
        generar_producto_volumen_margen(config["excel_path"])

def generar_producto_volumen_margen(archivo_excel):
    try:
        anio = simpledialog.askstring("Año", "Ingrese el año (Ej: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        mes_inicio = simpledialog.askstring("Mes inicio", "Ingrese el mes de inicio (1-12):", parent=app)
        dia_inicio = simpledialog.askstring("Día inicio", "Ingrese el día de inicio (1-31):", parent=app)
        mes_fin = simpledialog.askstring("Mes fin", "Ingrese el mes de fin (1-12):", parent=app)
        dia_fin = simpledialog.askstring("Día fin", "Ingrese el día de fin (1-31):", parent=app)
        top_n = simpledialog.askstring("Top N", "Ingrese el número de productos a mostrar:", parent=app)

        if not all([mes_inicio, dia_inicio, mes_fin, dia_fin, top_n]):
            messagebox.showerror("Error", "Todos los campos son obligatorios.")
            return
        mostrar_cargando()

        mes_inicio, dia_inicio, mes_fin, dia_fin, top_n = map(int, [mes_inicio, dia_inicio, mes_fin, dia_fin, top_n])
        fecha_inicio = datetime(anio, mes_inicio, dia_inicio)
        fecha_fin = datetime(anio, mes_fin, dia_fin)

        df_fact = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df_fact.columns = df_fact.columns.str.strip().str.upper()

        df_filtro = df_fact[(df_fact["FECHA"] >= fecha_inicio) & (df_fact["FECHA"] <= fecha_fin)].copy()
        df_filtro["ITEM"] = df_filtro["ITEM"].astype(str).str.strip().str.upper()

        excluir_keywords = ["FLETE", "MANTENIMIENTO", "DF-DISEÑO", "DISEÑO"]
        patron_excluir = "|".join(excluir_keywords)
        df_filtro = df_filtro[~df_filtro["ITEM"].str.contains(patron_excluir, case=False, na=False, regex=True)]

        df_filtro["NETO"] = pd.to_numeric(df_filtro["NETO"], errors="coerce")
        df_filtro["COST"] = pd.to_numeric(df_filtro["COST"], errors="coerce")
        df_filtro["QTYSHIP"] = pd.to_numeric(df_filtro["QTYSHIP"], errors="coerce")

        df_filtro["COSTO_TOTAL"] = df_filtro["COST"] * df_filtro["QTYSHIP"]
        df_filtro["MARGEN_%"] = ((df_filtro["NETO"] - df_filtro["COSTO_TOTAL"]) / df_filtro["NETO"]) * 100

        top_df = df_filtro.groupby("ITEM").agg({
            "NETO": "sum",
            "MARGEN_%": "mean"
        }).reset_index()
        top_df = top_df.sort_values(by="NETO", ascending=False).head(top_n)

        total_top = top_df["NETO"].sum()
        titulo = f"Top {top_n} productos por Valor vendido y Margen\n{fecha_inicio.date()} a {fecha_fin.date()} (${total_top:,.0f})"

        margen_escalado = (top_df["MARGEN_%"] / 100) * top_df["NETO"]
        margen_escalado = margen_escalado.clip(upper=top_df["NETO"] * 0.9)

        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(top_df["ITEM"], top_df["NETO"], color="lightsteelblue", label="Valor vendido ($)")
        ax.barh(top_df["ITEM"], margen_escalado, color="orange", height=0.4, label="Margen (%)")

        ax.invert_yaxis()
        ax.set_title(titulo, fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor vendido ($)")
        ax.legend()

        for i, (valor, margen, margen_px) in enumerate(zip(top_df["NETO"], top_df["MARGEN_%"], margen_escalado)):
            ax.text(valor + (max(top_df["NETO"]) * 0.01), i, f"${valor:,.0f}", va="center", fontsize=9)
            ax.text(margen_px + (max(top_df["NETO"]) * 0.01), i, f"{margen:.1f}%", va="center", fontsize=9)

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.0, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(0, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        top = tk.Toplevel(app)
        top.title(titulo)
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte de productos:\n{e}")

def generar_ciudades(archivo_excel):
    try:
        root = Tk()
        root.withdraw()

        anio = simpledialog.askstring("Filtrar por año", "¿Qué año deseas filtrar? (Ejemplo: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        top_n = simpledialog.askstring("Top ciudades", "¿Cuántas ciudades deseas mostrar en el TOP?", parent=app)
        if not top_n or not top_n.isdigit() or int(top_n) <= 0:
            messagebox.showerror("Error", "Número de TOP inválido.")
            return
        top_n = int(top_n)
        mostrar_cargando()

        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)

        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['QTYSHIP'] = pd.to_numeric(df['QTYSHIP'], errors='coerce')
        df['NETO'] = pd.to_numeric(df['NETO'], errors='coerce')

        df_filtrado = df[df['FECHA'].dt.year == anio]
        if df_filtrado.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.")
            return

        reporte = (
            df_filtrado.groupby('CITY')
            .agg(
                TOTAL_FACTURACION=('NETO', 'sum'),
                UNIDADES_VENDIDAS=('QTYSHIP', 'sum'),
                TRANSACCIONES=('ITEM', 'count'),
                CLIENTES_UNICOS=('ID_N', pd.Series.nunique)
            )
            .sort_values(by='TOTAL_FACTURACION', ascending=False)
            .head(top_n)
        )

        total_top = reporte['TOTAL_FACTURACION'].sum()
        titulo = f"Top {top_n} Ciudades - Facturación {anio} (${total_top:,.0f})"

        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(reporte.index, reporte['TOTAL_FACTURACION'], color="skyblue")
        ax.invert_yaxis()
        ax.set_title(titulo, fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor facturado ($)")

        for i, v in enumerate(reporte['TOTAL_FACTURACION']):
            ax.text(v + (max(reporte['TOTAL_FACTURACION']) * 0.01), i, f"${v:,.0f}", va="center", fontsize=9)

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.0, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(0, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        top = tk.Toplevel(app)
        top.title(titulo)
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte:\n{e}")

def generar_departamentos(archivo_excel):
    try:
        root = Tk()
        root.withdraw()

        anio = simpledialog.askstring("Filtrar por año", "¿Qué año deseas filtrar? (Ejemplo: 2025):", parent=app)
        if not anio or not anio.isdigit():
            messagebox.showerror("Error", "Año inválido.")
            return
        anio = int(anio)

        top_n = simpledialog.askstring("Top departamentos", "¿Cuántos departamentos deseas mostrar en el TOP?", parent=app)
        if not top_n or not top_n.isdigit() or int(top_n) <= 0:
            messagebox.showerror("Error", "Número de TOP inválido.")
            return
        top_n = int(top_n)
        mostrar_cargando()

        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)

        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['QTYSHIP'] = pd.to_numeric(df['QTYSHIP'], errors='coerce')
        df['NETO'] = pd.to_numeric(df['NETO'], errors='coerce')

        df_filtrado = df[df['FECHA'].dt.year == anio]
        if df_filtrado.empty:
            messagebox.showinfo("Sin datos", f"No hay datos para el año {anio}.")
            return

        reporte = (
            df_filtrado.groupby('DEPARTAMENTO')
            .agg(
                TOTAL_FACTURACION=('NETO', 'sum'),
                UNIDADES_VENDIDAS=('QTYSHIP', 'sum'),
                TRANSACCIONES=('ITEM', 'count'),
                CLIENTES_UNICOS=('ID_N', pd.Series.nunique)
            )
            .sort_values(by='TOTAL_FACTURACION', ascending=False)
            .head(top_n)
        )

        total_top = reporte['TOTAL_FACTURACION'].sum()
        titulo = f"Top {top_n} Departamentos - Facturación {anio} (${total_top:,.0f})"

        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(reporte.index, reporte['TOTAL_FACTURACION'], color="lightgreen")
        ax.invert_yaxis()
        ax.set_title(titulo, fontsize=14, fontweight="bold")
        ax.set_xlabel("Valor facturado ($)")

        for i, v in enumerate(reporte['TOTAL_FACTURACION']):
            ax.text(v + (max(reporte['TOTAL_FACTURACION']) * 0.01), i, f"${v:,.0f}", va="center", fontsize=9)

        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (0.0, 1.10), frameon=False, xycoords='axes fraction', box_alignment=(0, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        top = Toplevel(app)
        top.title(titulo)
        top.geometry("900x700")
        try:
            top.attributes('-topmost', 1)
        except:
            pass

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error generando el reporte:\n{e}")

def generar_comparativo_linea(archivo_excel):
    try:
        mostrar_cargando()

        # === Fechas YTD ===
        hoy = datetime.today()

        # Inicio de año actual (00:00:00) y fin hasta hoy (23:59:59)
        fecha_inicio_actual = datetime(hoy.year, 1, 1, 0, 0, 0)
        fecha_fin_actual = datetime(hoy.year, hoy.month, hoy.day, 23, 59, 59)

        # Mismo rango pero para el año anterior
        fecha_inicio_anterior = fecha_inicio_actual.replace(year=fecha_inicio_actual.year - 1)
        fecha_fin_anterior = fecha_fin_actual.replace(year=fecha_fin_actual.year - 1)

        # === Cargar datos ===
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")

        # Normalizar columnas
        df.columns = df.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)

        # Asegurar formato correcto
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"] = pd.to_numeric(df["NETO"], errors="coerce")

        # Limpiar filas con fecha o neto inválido
        df = df.dropna(subset=["FECHA", "NETO", "DESCLINEA"])

        # === Filtrar solo dos rangos YTD ===
        df_actual = df[(df["FECHA"] >= fecha_inicio_actual) & (df["FECHA"] <= fecha_fin_actual)]
        df_anterior = df[(df["FECHA"] >= fecha_inicio_anterior) & (df["FECHA"] <= fecha_fin_anterior)]

        # === Agrupar por línea de producto ===
        ventas_actual = df_actual.groupby("DESCLINEA")["NETO"].sum().reset_index()
        ventas_anterior = df_anterior.groupby("DESCLINEA")["NETO"].sum().reset_index()

        # === Unir y calcular diferencia ===
        comparativo = pd.merge(
            ventas_anterior,
            ventas_actual,
            on="DESCLINEA",
            how="outer",
            suffixes=(f"_{fecha_inicio_anterior.year}", f"_{fecha_inicio_actual.year}")
        ).fillna(0)

        comparativo["DIFERENCIA"] = (
            comparativo[f"NETO_{fecha_inicio_actual.year}"] -
            comparativo[f"NETO_{fecha_inicio_anterior.year}"]
        )

        # === Ordenar para gráfico ===
        comparativo = comparativo.sort_values(by=f"NETO_{fecha_inicio_actual.year}", ascending=True)

        # === Gráfico horizontal ===
        plt.style.use("seaborn-v0_8-whitegrid")
        fig, ax = plt.subplots(figsize=(12, 7))

        y_pos = range(len(comparativo))
        bar_height = 0.4

        ax.barh([y + bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_anterior.year}"],
                height=bar_height, label=str(fecha_inicio_anterior.year), color="#4F81BD")
        ax.barh([y - bar_height/2 for y in y_pos], comparativo[f"NETO_{fecha_inicio_actual.year}"],
                height=bar_height, label=str(fecha_inicio_actual.year), color="#F79646")

        # Etiquetas con valores
        for i, (val_ant, val_act) in enumerate(zip(
            comparativo[f"NETO_{fecha_inicio_anterior.year}"],
            comparativo[f"NETO_{fecha_inicio_actual.year}"]
        )):
            ax.text(val_ant + (val_ant * 0.01), i + bar_height/2, f"${val_ant:,.0f}",
                    va="center", fontsize=8)
            ax.text(val_act + (val_act * 0.01), i - bar_height/2, f"${val_act:,.0f}",
                    va="center", fontsize=8)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(comparativo["DESCLINEA"])
        ax.set_xlabel("Ventas ($)")
        ax.set_title(f"Comparativo YTD de Ventas por Línea\n"
                     f"{fecha_inicio_anterior.date()} a {fecha_fin_anterior.date()} vs "
                     f"{fecha_inicio_actual.date()} a {fecha_fin_actual.date()}",
                     fontsize=14, fontweight="bold")
        ax.legend()

        # Logo
        try:
            logo_img = plt.imread(LOGO_PATH)
            imagebox = OffsetImage(logo_img, zoom=0.6)
            ab = AnnotationBbox(imagebox, (1.05, 1.05), frameon=False,
                                xycoords='axes fraction', box_alignment=(1, 1))
            ax.add_artist(ab)
        except FileNotFoundError:
            pass

        plt.tight_layout()

        # === Mostrar en ventana ===
        top = tk.Toplevel(app)
        top.title("Comparativo YTD de Ventas por Línea")
        top.attributes("-topmost", True)
        top.focus_force()

        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        canvas = FigureCanvasTkAgg(fig, master=top)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Barra de herramientas
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, top)
        toolbar.update()

        canvas.get_tk_widget().pack(fill="both", expand=True)

        cerrar_cargando()

        # === Guardar Excel ordenado de MAYOR a menor ===
        ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                    filetypes=[("Archivos Excel", "*.xlsx")],
                                                    title="Guardar comparativo como")
        if ruta_guardado:
            comparativo_excel = comparativo.sort_values(
                by=f"NETO_{fecha_inicio_actual.year}", ascending=False
            )
            comparativo_excel.to_excel(ruta_guardado, index=False)

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error generando el comparativo:\n{e}")

def generar_rotacion_inventario(archivo_excel):
    try:
        mostrar_cargando()

        # === 1. FACTURACIÓN ===
        df_fac = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df_fac.columns = df_fac.columns.str.strip().str.upper().str.replace('\xa0', '', regex=True)

        columnas_fact = {"ITEM", "DESCRIPCION", "FECHA", "QTYSHIP"}
        if columnas_fact - set(df_fac.columns):
            raise ValueError(f"Faltan columnas en Facturacion: {columnas_fact - set(df_fac.columns)}")

        df_fac = df_fac.dropna(subset=["FECHA", "QTYSHIP", "ITEM"]).copy()
        df_fac["FECHA"] = pd.to_datetime(df_fac["FECHA"], errors="coerce")
        df_fac["QTYSHIP"] = pd.to_numeric(df_fac["QTYSHIP"], errors="coerce").fillna(0)
        df_fac["ITEM"] = df_fac["ITEM"].astype(str).str.strip()
        df_fac["AÑO_MES"] = df_fac["FECHA"].dt.to_period("M")

        meses_activos = (
            df_fac[df_fac["QTYSHIP"] > 0]
            .groupby(["ITEM", "DESCRIPCION"])["AÑO_MES"]
            .nunique()
            .reset_index()
            .rename(columns={"AÑO_MES": "MESES"})
        )

        ventas_totales = (
            df_fac.groupby(["ITEM", "DESCRIPCION"], as_index=False)["QTYSHIP"].sum()
            .rename(columns={"QTYSHIP": "VENTA_TOTAL"})
        )

        resumen = pd.merge(ventas_totales, meses_activos, on=["ITEM", "DESCRIPCION"], how="left").fillna(0)
        resumen["PROMEDIO_MES"] = resumen.apply(
            lambda row: row["VENTA_TOTAL"] / row["MESES"] if row["MESES"] > 0 else 0, axis=1
        )

        # === 2. INVENTARIO (simulando tu BUSCARV) ===
        df_inv = pd.read_excel(archivo_excel, sheet_name="Inventario SAI", header=None)

        # Rango equivalente a C7:E1200
        df_inv = df_inv.iloc[6:1200, 2:5]  # fila 7 → índice 6, columnas C:E → 2:5
        df_inv.columns = ["ITEM", "OTRA_COL", "INVENTARIO"]

        df_inv["ITEM"] = df_inv["ITEM"].astype(str).str.strip()
        df_inv["INVENTARIO"] = pd.to_numeric(df_inv["INVENTARIO"], errors="coerce").fillna(0)

        # Unión tipo BUSCARV
        resumen = pd.merge(resumen, df_inv[["ITEM", "INVENTARIO"]], on="ITEM", how="left").fillna(0)

        # Meses división
        resumen["MESES_DIVISION"] = resumen.apply(
            lambda row: row["INVENTARIO"] / row["PROMEDIO_MES"] if row["PROMEDIO_MES"] != 0 else 0,
            axis=1
        )

        for col in resumen.select_dtypes(include=["float", "int"]).columns:
            resumen[col] = resumen[col].astype(int)

        # Reordenar columnas
        resumen = resumen[["ITEM", "DESCRIPCION", "MESES", "VENTA_TOTAL",
                           "PROMEDIO_MES", "INVENTARIO", "MESES_DIVISION"]]

        # Guardar archivo
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar rotación Inventario como"
        )
        if ruta_guardado:
            resumen.to_excel(ruta_guardado, index=False)

        cerrar_cargando()
        messagebox.showinfo("Éxito", "Archivo de rotación de Inventario generado correctamente.")

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error generando la rotación:\n{e}")

def generar_ventas_semana(archivo_excel):
    try:

        mostrar_cargando()
        # === 1. Cargar y preparar datos ===
        df = pd.read_excel(archivo_excel, sheet_name="Facturacion")
        df.columns = df.columns.str.strip().str.upper()
        df = df.dropna(subset=["FECHA"]).copy()
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df["NETO"]  = pd.to_numeric(df["NETO"], errors="coerce").fillna(0)

        hoy          = datetime.today()
        anyo_full    = 2024
        anyo_actual  = hoy.year

        mask_full    = df["FECHA"].dt.year == anyo_full
        mask_actual  = (df["FECHA"].dt.year == anyo_actual) & (df["FECHA"] <= hoy)
        df = df[mask_full | mask_actual]

        # Mapeo de meses completos
        meses_map = {
            1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
            5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
            9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
        }
        df["MES"] = df["FECHA"].dt.month.map(meses_map)

        # Semana del mes
        def semana_del_mes(fecha):
            dia_semana_inicio = fecha.replace(day=1).weekday()
            desplaz = fecha.day + dia_semana_inicio - 1
            return (desplaz // 7) + 1

        df["SEMANA_MES"] = df["FECHA"].apply(semana_del_mes)
        df["AÑO"] = df["FECHA"].dt.year

        # === 2. Crear pivot por año ===
        meses_ord = list(meses_map.values())
        max_semana = df["SEMANA_MES"].max()
        semanas = list(range(1, int(max_semana) + 1))

        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Guardar ventas por semana"
        )
        if not ruta:
            return

        with pd.ExcelWriter(ruta, engine="xlsxwriter") as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Reporte")
            writer.sheets["Reporte"] = worksheet

            formato_moneda = workbook.add_format({"num_format": '"$"#,##0', "align": "right"})
            formato_header = workbook.add_format({"bold": True, "bg_color": "#92D050", "align": "center"})
            formato_titulo = workbook.add_format({"bold": True, "align": "left"})

            fila_inicio = 0
            for anyo in [anyo_actual, anyo_full]:
                df_a = df[df["AÑO"] == anyo]
                if df_a.empty:
                    continue

                pivot = (
                    df_a
                    .pivot_table(index="SEMANA_MES", columns="MES", values="NETO", aggfunc="sum", fill_value=0)
                    .reindex(index=semanas, columns=meses_ord, fill_value=0)
                )
                pivot["Total"] = pivot.sum(axis=1)

                # Título del año
                worksheet.write(fila_inicio, 0, anyo, formato_titulo)

                # Encabezados
                encabezados = ["Semana"] + meses_ord + ["Total"]
                for col, nombre in enumerate(encabezados):
                    worksheet.write(fila_inicio + 1, col, nombre, formato_header)

                # Datos
                for i, semana in enumerate(semanas):
                    worksheet.write(fila_inicio + 2 + i, 0, semana)
                    for j, mes in enumerate(meses_ord):
                        worksheet.write(fila_inicio + 2 + i, j + 1, pivot.iloc[i, j], formato_moneda)
                    worksheet.write(fila_inicio + 2 + i, len(meses_ord) + 1, pivot.iloc[i, -1], formato_moneda)

                # Fila total
                fila_total = fila_inicio + 2 + len(semanas)
                worksheet.write(fila_total, 0, "")
                for j, mes in enumerate(meses_ord):
                    worksheet.write(fila_total, j + 1, pivot[mes].sum(), formato_moneda)
                worksheet.write(fila_total, len(meses_ord) + 1, pivot["Total"].sum(), formato_moneda)

                fila_inicio = fila_total + 3  # Espacio entre bloques

        cerrar_cargando()
        messagebox.showinfo("Éxito", "Reporte generado correctamente.")

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

def generar_presupuesto_año(archivo_excel):
    try:

        mostrar_cargando()
        # === 1. leer hoja ppto año ===
        df_ppto_año = pd.read_excel(archivo_excel, sheet_name="Ppto Año", header=None)

        # tomar solo las filas de meses (fila 6 a 17 en tu archivo)
        df_ppto_año = df_ppto_año.iloc[6:18, [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]]
        df_ppto_año.columns = [
            "Mes", "2024", "Acum 2024",
            "Ppto Mes 2025", "Ppto Acum 2025",
            "Real 2025", "Acum 2025",
            "% Cumpl Mes", "Acum 2025 vs Acum 2024",
            "Acum 2025 vs Ppto Acum 2025"
        ]

        # limpiar y convertir datos
        for col in ["2024", "Acum 2024", "Ppto Mes 2025", "Ppto Acum 2025", "Real 2025", "Acum 2025"]:
            df_ppto_año[col] = pd.to_numeric(df_ppto_año[col], errors="coerce").fillna(0).astype(int)

        for col in ["% Cumpl Mes", "Acum 2025 vs Acum 2024", "Acum 2025 vs Ppto Acum 2025"]:
            df_ppto_año[col] = pd.to_numeric(df_ppto_año[col], errors="coerce")

        # === 2. guardar excel con formato ===
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Guardar reporte anual"
        )
        if not ruta:
            return

        with pd.ExcelWriter(ruta, engine="xlsxwriter") as writer:
            df_ppto_año.to_excel(writer, index=False, sheet_name="Reporte")
            wb = writer.book
            ws = writer.sheets["Reporte"]

            formato_moneda     = wb.add_format({"num_format": '"$"#,##0'})
            formato_porcentaje = wb.add_format({"num_format": "0.0%"})
            formato_header     = wb.add_format({"bold": True, "bg_color": "#92D050"})

            # formato columnas
            for col in range(1, 7):
                ws.set_column(col, col, 15, formato_moneda)
            for col in range(7, 10):
                ws.set_column(col, col, 20, formato_porcentaje)
            for col, nombre in enumerate(df_ppto_año.columns):
                ws.write(0, col, nombre, formato_header)

        cerrar_cargando()
        messagebox.showinfo("Éxito", "Reporte generado correctamente")

    except Exception as e:
        cerrar_cargando()
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# ==== Interfaz ====
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")
app = ctk.CTk()
app.title("🔦 EcoLite")
app.geometry("950x560")
app.resizable(True, True)

# Panel izquierdo
frame_izq = ctk.CTkFrame(app, width=300, fg_color="white", corner_radius=0)
frame_izq.pack(side="left", fill="y")

# Encabezado fijo (Logo, botón de subida y nombre de archivo)
titulo_app = ctk.CTkLabel(frame_izq, text="🔦 EcoLite", font=("Segoe UI", 20, "bold"), text_color="#222")
titulo_app.pack(pady=15)

btn_subir = ctk.CTkButton(
    frame_izq, text="📤 Subir Excel", command=subir_excel,
    height=40, fg_color="#0A84FF", hover_color="#0066CC", text_color="white"
)
btn_subir.pack(pady=(5, 10), padx=15, fill="x")

label_archivo = ctk.CTkLabel(frame_izq, text="📂 Ningún archivo cargado", font=("Segoe UI", 13), text_color="#777")
label_archivo.pack(pady=(0, 20), padx=15)

# ==== Scrollable Frame para los módulos ====
scroll_modulos = ctk.CTkScrollableFrame(
    frame_izq,
    width=260, height=700,       
    fg_color="white",
    corner_radius=0
)
scroll_modulos.pack(fill="both", expand=True, padx=15, pady=(0, 15))


# Lista de módulos y acciones
modulos = {
    "📈 Vendedores": [
        ("Volumen de ventas", "volumen_vendedores", 
         "Reporte anual que presenta el volumen total de ventas por vendedor, ordenado de mayor a menor, excluyendo vendedores inactivos o no comerciales."),

        ("Margen de ventas", "margen_vendedores", 
         "Análisis anual del margen de ganancia por vendedor, expresado en porcentaje, para identificar la rentabilidad individual."),

        ("Comparativo", "comparativo_vendedor", 
         "Comparación Year-To-Date entre el año actual y el anterior, mostrando el crecimiento o disminución en ventas por vendedor. Genera el archivo en excel en donde muestra organizado de mayor a menor si hay crecimiento o no."),

        ("Departamentos", "departamentos_vendedor", 
         "Clasificación del Top de vendedores segmentados por departamento en el año seleccionado."),

        ("Ciudades", "cuidades_vendedor", 
         "Clasificación del Top de vendedores segmentados por ciudad en el año seleccionado."),

        ("Comparativo Ciudades", "comparativo_cuidad", 
         "Comparativo YTD de ventas por ciudad entre el año actual y el anterior, aplicando análisis Pareto para identificar el 70% de la facturación. Genera el archivo en excel en donde muestra organizado de mayor a menor si hay crecimiento o no."),

        ("Comparativo Departamentos", "comparativo_departamento", 
         "Comparativo YTD de ventas por departamento entre el año actual y el anterior, destacando variaciones absolutas y porcentuales. Genera el archivo en excel en donde muestra organizado de mayor a menor si hay crecimiento o no.")
    ],

    "📦 Producto": [
        ("Margen de ventas", "margen_productos", 
         "Análisis de rentabilidad de productos en un rango de fechas personalizado, calculando el margen promedio y excluyendo artículos no comerciales."),

        ("Volumen y Margen", "producto_volumen_margen", 
         "Reporte combinado que presenta el Top de productos con mayor volumen de ventas y su margen de ganancia en un periodo definido.")
    ],

    "🌍 Otros": [
        ("Ciudades", "reporte_cuidades", 
         "Top de ciudades con mayor facturación anual, incluyendo métricas adicionales: unidades vendidas, número de transacciones y clientes únicos."),

        ("Departamentos", "reporte_departamentos", 
         "Top de departamentos con mayor facturación anual, incluyendo métricas adicionales: unidades vendidas, número de transacciones y clientes únicos."),

        ("Comparativo Líneas", "comparativo_linea", 
        "Comparativo YTD de ventas por línea de producto, mostrando valores del año actual y el anterior, diferencia absoluta, y visualización gráfica de las variaciones."),
        
        ("Rotación Inventario", "rotacion_inventario", 
        "Comparativo YTD de ventas por línea de producto, mostrando valores del año actual y el anterior, diferencia absoluta, y visualización gráfica de las variaciones."),
        
        ("Ventas Semana", "ventas_semana", 
        "Comparativo YTD de ventas por línea de producto, mostrando valores del año actual y el anterior, diferencia absoluta, y visualización gráfica de las variaciones."),
        
        ("Presupuesto Año", "presupuesto_año", 
        "Comparativo YTD de ventas por línea de producto, mostrando valores del año actual y el anterior, diferencia absoluta, y visualización gráfica de las variaciones.")


    ]
}

for nombre_modulo, acciones in modulos.items():
    bloque = ctk.CTkFrame(scroll_modulos, fg_color="#F8F9FA", corner_radius=8)
    bloque.pack(fill="x", pady=6)
    
    sub = ctk.CTkLabel(
        bloque, text=nombre_modulo, font=("Segoe UI", 15, "bold"), text_color="#333"
    )
    sub.pack(anchor="w", padx=10, pady=(8, 4))
    
    for texto_accion, accion_id, descripcion in acciones:
        btn = ctk.CTkButton(
            bloque,
            text=texto_accion,
            anchor="w",
            fg_color="#E9ECEF",
            text_color="#333",
            hover_color="#DEE2E6",
            corner_radius=6,
            height=32,
            font=("Segoe UI", 13),
            command=lambda a=accion_id, d=descripcion: mostrar_descripcion(a, d)
        )
        btn.pack(fill="x", padx=10, pady=3)

# Panel derecho
frame_der = ctk.CTkFrame(app, fg_color="white", corner_radius=0)
frame_der.pack(side="right", fill="both", expand=True)

# Cargar configuración inicial
config = cargar_config()
if config["excel_path"]:
    label_archivo.configure(text=f"📂 {os.path.basename(config['excel_path'])}")

accion_seleccionada = None

app.mainloop()