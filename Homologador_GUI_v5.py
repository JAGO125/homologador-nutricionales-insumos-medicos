import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import threading
import os
import sys
import re
import subprocess
import pyodbc
from rapidfuzz import fuzz
from PIL import Image, ImageTk


# ================= CONFIGURACIÓN BASE DE DATOS =================

servidor = "192.168.10.10"
base_datos = "BDProcedimiento"
usuario = "AlejandraME"
contraseña = "@leja_Mora"

# ===============================================================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ================= NORMALIZACIÓN =================

def normalizar(txt):
    if not isinstance(txt, str):
        return ""

    txt = txt.replace(",", ".").lower()

    txt = re.sub(r'(\d+(\.\d+)?)([a-z%]+)', r'\1 \3', txt)
    txt = re.sub(r'([a-z%]+)(\d+)', r'\1 \2', txt)

    patrones_unidades = {
        r'\bmiligramo(s)?\b': 'mg',
        r'\bgramo(s)?\b': 'g',
        r'\bkilogramo(s)?\b': 'kg',
        r'\bmililitro(s)?\b': 'ml',
        r'\bcc\b': 'ml',
        r'\blitro(s)?\b': 'l',
        r'\bmicrogramo(s)?\b': 'mcg',
        r'\bu\s*i\b': 'ui',
        r'\b%\b': 'porciento'
    }

    for patron, reemplazo in patrones_unidades.items():
        txt = re.sub(patron, reemplazo, txt)

    txt = re.sub(r'[^a-z0-9\s\.]', ' ', txt)
    txt = re.sub(r'\s+', ' ', txt).strip()

    return txt


def marca_principal(texto):
    palabras = texto.split()
    if palabras:
        return palabras[0][:5]
    return ""


def extraer_volumen(texto):
    match = re.search(r'(\d+(\.\d+)?)\s*(ml|l)', texto)
    if match:
        valor = float(match.group(1))
        unidad = match.group(3)

        if unidad == "l":
            valor = valor * 1000

        return int(valor)
    return None


# ================= NUEVAS FUNCIONES (AJUSTE MOTOR) =================

def extraer_referencia(texto):
    match = re.search(r'\b(s[\s\-]?\d{3,4})\b', texto)
    if match:
        return match.group(1).replace(" ", "").replace("-", "")
    return None


def extraer_tipo_insumo(texto):
    if "equipo" in texto:
        return "equipo"
    if "tornillo" in texto:
        return "tornillo"
    if "clavo" in texto:
        return "clavo"
    return None


# ================= DESCARGA DESDE SQL =================

def descargar_desde_sql(tipo):

    try:
        ruta_actual = os.getcwd()
        carpeta_base = os.path.join(ruta_actual, "Base_Maestra_SQL")
        os.makedirs(carpeta_base, exist_ok=True)

        ruta_archivo = os.path.join(carpeta_base, "Base_Maestra_SQL.xlsx")

        try:
            conexion = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={servidor};"
                f"DATABASE={base_datos};"
                f"UID={usuario};"
                f"PWD={contraseña};"
                "TrustServerCertificate=yes;"
                "Connection Timeout=5;"
            )
        except pyodbc.Error as err:
            print("ERROR REAL:", err)
            messagebox.showerror(
                "Error de conexión",
                "No se pudo conectar al servidor SQL Server.\n\n"
                "Verifique:\n"
                "• Servidor activo\n"
                "• Red disponible\n"
                "• Credenciales\n"
                "• Driver ODBC instalado"
            )
            return

        escritor = pd.ExcelWriter(ruta_archivo, engine="openpyxl")

        if tipo in ["Medicamentos", "Todo"]:
            query = """
                SELECT Expediente_Consecutivo AS codigo,
                       descripcion
                FROM VCon_Cum_Vigentes_Activos_No_Muestra_Medica
                WHERE EstadoRegistro IN ('VIGENTE','EN TRAMITE RENOVACION')
                AND EstadoCum = 'Activo'
                AND MuestraMedica = 'NO'
            """
            df = pd.read_sql(query, conexion)
            df.to_excel(escritor, sheet_name="MEDICAMENTOS", index=False)

        if tipo in ["Insumos", "Todo"]:
            query = """
                SELECT CodTecnologiaUnica AS codigo,
                       DesTecnologia AS descripcion
                FROM VCon_Insumos
            """
            df = pd.read_sql(query, conexion)
            df.to_excel(escritor, sheet_name="INSUMOS", index=False)

        if tipo in ["Nutricionales", "Todo"]:
            query = """
                SELECT CodTecnologiaUnica AS codigo,
                       DesTecnologia AS descripcion
                FROM VCon_Nutricionales
            """
            df = pd.read_sql(query, conexion)
            df.to_excel(escritor, sheet_name="NUTRICIONALES", index=False)

        escritor.close()
        conexion.close()

        messagebox.showinfo(
            "Descarga Exitosa",
            f"Base Maestra generada correctamente en:\n{ruta_archivo}"
        )

        return ruta_archivo

    except Exception as e:
        messagebox.showerror("Error inesperado", str(e))


# ================= HOMOLOGACIÓN AJUSTADA =================

def homologar_y_guardar(path_base, path_ref, umbral, progress_callback):

    dic_hojas = pd.read_excel(path_base, sheet_name=None)
    df_ref = pd.read_excel(path_ref)
    col_desc_ref = df_ref.columns[0]

    coincidencias = []
    no_coinc = []
    total = len(df_ref)

    bases_norm = {}

    for nombre_hoja, df in dic_hojas.items():
        nombre_hoja = nombre_hoja.strip().upper()

        if nombre_hoja not in ["INSUMOS", "NUTRICIONALES", "MEDICAMENTOS"]:
            continue

        df = df.iloc[:, :2]
        df.columns = ["CODIGO", "DESCRIPCION"]
        df["desc_norm"] = df["DESCRIPCION"].apply(normalizar)

        bases_norm[nombre_hoja] = df

    for i, row in df_ref.iterrows():

        desc_original = row[col_desc_ref]
        desc_norm = normalizar(desc_original)

        marca_ref = marca_principal(desc_norm)
        ref_ref = extraer_referencia(desc_norm)
        tipo_ref = extraer_tipo_insumo(desc_norm)
        vol_ref = extraer_volumen(desc_norm)

        mejor_puntaje = 0
        mejor_fila = None
        mejor_categoria = None

        for categoria, df_base in bases_norm.items():

            for _, fila_base in df_base.iterrows():

                marca_base = marca_principal(fila_base["desc_norm"])
                if marca_ref != marca_base:
                    continue

                ref_base = extraer_referencia(fila_base["desc_norm"])
                tipo_base = extraer_tipo_insumo(fila_base["desc_norm"])
                vol_base = extraer_volumen(fila_base["desc_norm"])

                if categoria == "INSUMOS":
                    if ref_ref and ref_base and ref_ref != ref_base:
                        continue
                    if tipo_ref and tipo_base and tipo_ref != tipo_base:
                        continue

                if vol_ref and vol_base and vol_ref != vol_base:
                    continue

                puntaje_raw = fuzz.token_set_ratio(desc_norm, fila_base["desc_norm"])

                long_diff = abs(len(desc_norm) - len(fila_base["desc_norm"])) / max(
                    len(desc_norm), len(fila_base["desc_norm"])
                )

                penalizacion = 1 - (long_diff * 0.4)
                puntaje_final = puntaje_raw * penalizacion

                if puntaje_final > mejor_puntaje:
                    mejor_puntaje = puntaje_final
                    mejor_fila = fila_base
                    mejor_categoria = categoria

        if mejor_puntaje >= umbral and mejor_fila is not None:
            coincidencias.append({
                "Descripción a Homologar": desc_original,
                "Categoría": mejor_categoria,
                "Código Homologado": mejor_fila["CODIGO"],
                "Descripción Homologada": mejor_fila["DESCRIPCION"],
                "Similitud (%)": round(mejor_puntaje, 2)
            })
        else:
            no_coinc.append({
                "Descripción": desc_original,
                "Motivo": f"Sin coincidencia relevante (máx {round(mejor_puntaje,2)}%)"
            })

        progress_callback(int((i + 1) / total * 100))

    os.makedirs("Resultados", exist_ok=True)

    pd.DataFrame(coincidencias).to_excel("Resultados/Coincidencias.xlsx", index=False)
    pd.DataFrame(no_coinc).to_excel("Resultados/No_Coincidencias.xlsx", index=False)

    return "Resultados", len(coincidencias), len(no_coinc)


# ================= INTERFAZ =================

class HomologadorApp:

    def __init__(self, root):
        self.root = root
        root.title("Homologador de Tecnologías e Insumos Médicos")
        root.geometry("760x600")
        root.resizable(False, False)
        root.configure(bg="#f2f2f2")

        # 🔵 HEADER CON LOGO
        header_frame = tk.Frame(root, bg="#f2f2f2")
        header_frame.pack(fill="x", pady=(10, 0))

        titulo = tk.Label(header_frame,
                          text="HOMOLOGADOR DE NUTRICIONALES E INSUMOS MÉDICOS",
                          font=("Arial", 14, "bold"),
                          fg="#003366",
                          bg="#f2f2f2")
        titulo.pack(side="left", padx=20)

        try:
            logo_path = resource_path("6a31b2e0-d8e8-43b4-895d-c03bf5fc56dd.png")
            img = Image.open(logo_path).resize((140, 70))
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(header_frame, image=self.logo_img, bg="#f2f2f2").pack(side="right", padx=7)
        except:
            pass

        tk.Canvas(root, height=2, bg="#0078D7", highlightthickness=0).pack(fill="x", pady=(5, 15))

        frame = tk.Frame(root, bg="#f2f2f2")
        frame.pack(pady=5)

        ttk.Label(frame, text="Archivo Base (Maestra):").grid(row=0, column=0, sticky="w")
        self.path_base = tk.StringVar()
        ttk.Entry(frame, textvariable=self.path_base, width=80).grid(row=1, column=0, padx=5)
        ttk.Button(frame, text="Descargar Base SQL", command=self.descargar_base).grid(row=1, column=1, padx=10)

        ttk.Label(frame, text="Archivo a Homologar:").grid(row=2, column=0, sticky="w", pady=(15, 0))
        self.path_ref = tk.StringVar()
        ttk.Entry(frame, textvariable=self.path_ref, width=80).grid(row=3, column=0, padx=5)
        ttk.Button(frame, text="Seleccionar", command=self.cargar_ref).grid(row=3, column=1)

        ttk.Label(root, text="Umbral de Similitud (%):", background="#f0f0f0").pack()
        self.umbral = tk.IntVar(value=50)
        ttk.Entry(root, textvariable=self.umbral, width=10, justify="center").pack(pady=4)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
        self.progress.pack(pady=20)

        self.progress_label = ttk.Label(root, text="Progreso: 0%", background="#f2f2f2")
        self.progress_label.pack()

        botones_frame = tk.Frame(root, bg="#f2f2f2")
        botones_frame.pack(pady=20)

        self.btn_iniciar = ttk.Button(botones_frame, text="Iniciar Homologación", command=self.iniciar)
        self.btn_iniciar.grid(row=0, column=0, padx=10)

        

        self.result_label = ttk.Label(root, text="", foreground="green", wraplength=700, background="#f2f2f2")
        self.result_label.pack(pady=10)

        tk.Label(root, text="Creado por Ing. John Alejandro Gómez Hernández",
                 font=("Arial", 10), bg="#f2f2f2").pack(side="left", anchor="sw", padx=10, pady=10)

    def cargar_base(self):
        archivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if archivo:
            self.path_base.set(archivo)

    def cargar_ref(self):
        archivo = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if archivo:
            self.path_ref.set(archivo)

    def descargar_base(self):
        ventana = tk.Toplevel(self.root)
        ventana.title("Seleccione Base a Descargar")
        ventana.geometry("320x260")
        ventana.grab_set()

        opcion = tk.StringVar(value="Insumos")

        ttk.Label(ventana, text="¿Qué base desea descargar?",
                  font=("Arial", 10, "bold")).pack(pady=10)

        opciones = ["Insumos", "Nutricionales", "Medicamentos", "Todo"]

        for op in opciones:
            ttk.Radiobutton(ventana, text=op, variable=opcion, value=op).pack(anchor="w", padx=30, pady=5)

        def confirmar():
            tipo = opcion.get()
            archivo = descargar_desde_sql(tipo)
            self.path_base.set(archivo)
            messagebox.showinfo("Éxito", f"{tipo} descargado correctamente")
            ventana.destroy()

        ttk.Button(ventana, text="Descargar", command=confirmar).pack(pady=15)

    def actualizar_progreso(self, val):
        self.progress["value"] = val
        self.progress_label.config(text=f"Progreso: {val}%")
        self.root.update_idletasks()

    def iniciar(self):
        if not self.path_base.get() or not self.path_ref.get():
            messagebox.showwarning("Advertencia", "Debe seleccionar ambos archivos.")
            return

        self.btn_iniciar.config(state="disabled", text="Procesando...")
        threading.Thread(target=self.procesar, daemon=True).start()

    def procesar(self):
        try:
            carpeta, n_coinc, n_noco = homologar_y_guardar(
                self.path_base.get(),
                self.path_ref.get(),
                self.umbral.get(),
                self.actualizar_progreso
            )

            msg = f"Proceso completado.\nCoincidencias: {n_coinc}\nNo coincidencias: {n_noco}"
            self.result_label.config(text=msg)
            messagebox.showinfo("Finalizado", msg)

            if os.name == "nt":
                os.startfile(carpeta)
            else:
                subprocess.Popen(["xdg-open", carpeta])

        except Exception as e:
            messagebox.showerror("Error", str(e))

        finally:
            self.btn_iniciar.config(state="normal", text="Iniciar Homologación")


# ================= LANZADOR =================

if __name__ == "__main__":
    root = tk.Tk()
    app = HomologadorApp(root)
    root.mainloop()