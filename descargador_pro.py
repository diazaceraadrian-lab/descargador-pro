import customtkinter as ctk
import pandas as pd
import requests
import os
import io
import re
import json
import queue
import threading
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
from rembg import remove
from tkinter import filedialog, messagebox

# Formatos de imagen que sabemos guardar con su extensión correcta
EXTENSIONES = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp', 'GIF': '.gif'}


class DescargadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Turbo Descargador Inteligente v3.2")
        self.geometry("600x700")

        self.ruta_excel = ctk.StringVar()
        self.lista_errores = []

        # Buzón entre el hilo de descarga y la ventana.
        # IMPORTANTE: el hilo de descarga NO puede tocar la ventana directamente
        # (Tkinter no lo permite y provoca cuelgues raros). Deja aquí sus
        # mensajes y la ventana los recoge ella misma cada 100 milisegundos.
        self.buzon = queue.Queue()

        # --- DICCIONARIO DE SINÓNIMOS AMPLIADO ---
        self.SINONIMOS_SKU = ['SKU', 'REFERENCIA', 'ID', 'NOMBRE', 'COD', 'CODIGO', 'NOMBRE_FOTO', 'NOMBRE FOTO']
        self.SINONIMOS_URL = ['URLS', 'URL', 'LINK', 'IMAGEN', 'IMAGENES', 'LINKS', 'PHOTO', 'ENLACE']
        self.SINONIMOS_FONDO = ['FONDO', 'QUITAR FONDO', 'BORRAR FONDO', 'REMOVE BG', 'REMBG', 'SIN FONDO']

        # Interfaz Gráfica
        self.label_titulo = ctk.CTkLabel(self, text="Descargador de Imágenes Pro", font=("Arial", 22, "bold"))
        self.label_titulo.pack(pady=20)

        # Frame de selección de archivo
        self.frame_archivo = ctk.CTkFrame(self)
        self.frame_archivo.pack(pady=10, padx=20, fill="x")
        self.btn_archivo = ctk.CTkButton(self.frame_archivo, text="Seleccionar Excel o CSV", command=self.seleccionar_archivo)
        self.btn_archivo.pack(pady=10)
        self.lbl_ruta = ctk.CTkLabel(self.frame_archivo, text="Archivo no cargado", font=("Arial", 11, "italic"))
        self.lbl_ruta.pack(pady=(0, 10))

        # Frame de ajustes
        self.frame_config = ctk.CTkFrame(self)
        self.frame_config.pack(pady=10, padx=20, fill="x")

        self.label_hilos = ctk.CTkLabel(self.frame_config, text="Velocidad (Hilos simultáneos):")
        self.label_hilos.pack(pady=(10, 0))

        self.slider_hilos = ctk.CTkSlider(self.frame_config, from_=1, to=50, number_of_steps=49)
        self.slider_hilos.set(20)
        self.slider_hilos.pack(pady=5, padx=20)

        self.lbl_hilos_val = ctk.CTkLabel(self.frame_config, text="20", font=("Arial", 12, "bold"))
        self.lbl_hilos_val.pack()
        self.slider_hilos.configure(command=lambda v: self.lbl_hilos_val.configure(text=str(int(v))))

        self.switch_fondo = ctk.CTkSwitch(self.frame_config, text="Procesar borrado de fondos automáticamente")
        self.switch_fondo.select()
        self.switch_fondo.pack(pady=15)

        # Consola de texto
        self.txt_consola = ctk.CTkTextbox(self, width=540, height=160, font=("Consolas", 12))
        self.txt_consola.pack(pady=10, padx=20)

        # Botón de acción principal
        self.btn_iniciar = ctk.CTkButton(self, text="INICIAR DESCARGA", fg_color="#1f6aa5", font=("Arial", 16, "bold"), height=45, command=self.lanzar_hilo_proceso)
        self.btn_iniciar.pack(pady=10)

        # Botones de carpetas
        self.frame_carpetas = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_carpetas.pack(pady=10)
        ctk.CTkButton(self.frame_carpetas, text="📂 Ver Originales", width=140, command=lambda: self.abrir_carpeta("originales")).grid(row=0, column=0, padx=10)
        ctk.CTkButton(self.frame_carpetas, text="✨ Ver Sin Fondo", width=140, command=lambda: self.abrir_carpeta("sin_fondo")).grid(row=0, column=1, padx=10)

        self.vaciar_buzon()

    def abrir_carpeta(self, nombre):
        ruta = os.path.join(os.getcwd(), nombre)
        if not os.path.exists(ruta): os.makedirs(ruta)
        os.startfile(ruta)

    def seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if ruta:
            self.ruta_excel.set(ruta)
            self.lbl_ruta.configure(text=os.path.basename(ruta))

    def log(self, mensaje):
        """Lo puede llamar cualquier hilo: el mensaje va al buzón, no a la ventana."""
        self.buzon.put(f"> {mensaje}\n")

    def en_ventana(self, funcion):
        """Encarga a la ventana que ejecute algo ELLA misma (tocar botones, avisos)."""
        self.buzon.put(funcion)

    def vaciar_buzon(self):
        """La ventana recoge lo que ha dejado el hilo de descarga. Se llama a sí
        misma cada 100 milisegundos; como la ejecuta la ventana, es seguro."""
        try:
            while True:
                cosa = self.buzon.get_nowait()
                if callable(cosa):
                    cosa()
                else:
                    self.txt_consola.insert("end", cosa)
                    self.txt_consola.see("end")
        except queue.Empty:
            pass
        self.after(100, self.vaciar_buzon)

    def encontrar_columna(self, columnas, lista_sinonimos):
        for col in columnas:
            if str(col).strip().upper() in lista_sinonimos:
                return col
        return None

    def leer_listado(self, ruta):
        if not ruta.lower().endswith('.csv'):
            return pd.read_excel(ruta)
        # sep=None adivina el separador: los CSV de Excel en España usan ';'
        # y los de otros programas ','. 'utf-8-sig' quita los caracteres
        # invisibles que Excel pone al principio del archivo.
        try:
            return pd.read_csv(ruta, sep=None, engine='python', encoding='utf-8-sig')
        except UnicodeDecodeError:
            return pd.read_csv(ruta, sep=None, engine='python', encoding='latin-1')

    def ejecutar_descarga(self):
        self.lista_errores = []

        try:
            df = self.leer_listado(self.ruta_excel.get())

            col_sku = self.encontrar_columna(df.columns, self.SINONIMOS_SKU)
            col_url = self.encontrar_columna(df.columns, self.SINONIMOS_URL)
            col_fondo = self.encontrar_columna(df.columns, self.SINONIMOS_FONDO)

            if not col_sku or not col_url:
                self.log("❌ ERROR: No se detectó columna de Nombre/SKU o de URLs.")
                return

            self.log(f"🔎 Mapeo: ID='{col_sku}' | Link='{col_url}'")

            tareas = []
            nombres_usados = set()
            repetidos = 0
            for _, fila in df.iterrows():
                sku = str(fila[col_sku]).strip()
                urls_raw = str(fila[col_url]).strip()
                urls = self.limpiar_urls(urls_raw)

                quitar_fondo = False
                if col_fondo:
                    quitar_fondo = str(fila[col_fondo]).lower().strip() in ("si", "sí")

                for i, url in enumerate(urls):
                    nombre_final = sku if i == 0 else f"{sku}_{i}"
                    # Si el SKU se repite en otra fila, no pisamos la foto anterior:
                    # le añadimos _b, _c... y lo contamos para avisar.
                    if nombre_final in nombres_usados:
                        repetidos += 1
                        letra = ord('b')
                        while f"{nombre_final}_{chr(letra)}" in nombres_usados:
                            letra += 1
                        nombre_final = f"{nombre_final}_{chr(letra)}"
                    nombres_usados.add(nombre_final)
                    tareas.append((url, nombre_final, quitar_fondo))

            if repetidos:
                self.log(f"⚠️ {repetidos} fotos con SKU repetido: se guardan con _b, _c...")

            # Las carpetas se crean UNA vez aquí, antes de lanzar los hilos.
            # Si las creara cada hilo, dos a la vez podrían chocar y dar error.
            os.makedirs("originales", exist_ok=True)
            os.makedirs("sin_fondo", exist_ok=True)

            hilos = int(self.slider_hilos.get())
            self.log(f"🚀 Procesando {len(tareas)} imágenes a {hilos} hilos...")

            with ThreadPoolExecutor(max_workers=hilos) as executor:
                list(executor.map(self.procesar_tarea, tareas))

            if self.lista_errores:
                pd.DataFrame(self.lista_errores).to_excel("errores_recientes.xlsx", index=False)
                self.log(f"⚠️ {len(self.lista_errores)} errores. Ver 'errores_recientes.xlsx'")

            self.log("✅ ¡Proceso finalizado!")
            self.en_ventana(lambda: messagebox.showinfo("Completado", "Todas las imágenes han sido procesadas."))

        except Exception as e:
            self.log(f"❌ Error crítico: {e}")

        finally:
            self.en_ventana(lambda: self.btn_iniciar.configure(state="normal", text="INICIAR DESCARGA"))

    def limpiar_urls(self, texto):
        if texto.startswith('[') and texto.endswith(']'):
            try:
                u = json.loads(texto)
                return [i.replace(r'\/', '/') for i in u] if isinstance(u, list) else []
            except: pass
        t = texto.replace('"', '').replace(r'\/', '/')
        # Cortamos por la coma que va JUSTO ANTES de "http", pero sin comernos
        # el "http": (?=...) solo mira lo que viene detrás, no lo quita.
        partes = re.split(r',\s*(?=https?://)', t)
        return [p.strip() for p in partes if p.strip().startswith('http')]

    def procesar_tarea(self, tarea):
        url, nombre, fondo = tarea
        try:
            r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
            if r.status_code != 200:
                self.lista_errores.append({'SKU_Nombre': nombre, 'URL': url, 'Motivo': f"La web respondió {r.status_code}"})
                return False

            # Guardamos la foto con la extensión de su formato real, no siempre .jpg
            formato = Image.open(io.BytesIO(r.content)).format
            ruta_original = os.path.join("originales", f"{nombre}{EXTENSIONES.get(formato, '.jpg')}")
            with open(ruta_original, 'wb') as f: f.write(r.content)

            if fondo and self.switch_fondo.get():
                ruta_png = os.path.join("sin_fondo", f"{nombre}.png")
                res = remove(r.content)
                with open(ruta_png, 'wb') as f: f.write(res)
            return True
        except Exception as e:
            self.lista_errores.append({'SKU_Nombre': nombre, 'URL': url, 'Motivo': str(e)})
            return False

    def lanzar_hilo_proceso(self):
        if not self.ruta_excel.get():
            messagebox.showwarning("Falta archivo", "Por favor, selecciona un archivo Excel primero.")
            return
        self.btn_iniciar.configure(state="disabled", text="TRABAJANDO...")
        threading.Thread(target=self.ejecutar_descarga, daemon=True).start()

if __name__ == "__main__":
    app = DescargadorApp()
    app.mainloop()
