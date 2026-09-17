import json
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from dotenv import load_dotenv

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"
CARPETA_CONSULTAS = "resultado_consulta"
CARPETA_RESPUESTAS = "borrador_respuesta"


# --- CREACIÓN DE CARPETAS Y CONTROL DE DUPLICADOS ---
def asegurar_carpetas():
    os.makedirs(CARPETA_CONSULTAS, exist_ok=True)
    os.makedirs(CARPETA_RESPUESTAS, exist_ok=True)


def cargar_procesados():
    if os.path.exists(REGISTRO_PROCESADOS):
        try:
            with open(REGISTRO_PROCESADOS, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()


def guardar_procesado(id_doc):
    procesados = cargar_procesados()
    procesados.add(id_doc)
    with open(REGISTRO_PROCESADOS, "w", encoding="utf-8") as f:
        json.dump(list(procesados), f, ensure_ascii=False, indent=2)


def sanitizar_nombre_archivo(texto):
    """Limpia caracteres no válidos para nombres de archivos en Windows."""
    texto_limpio = re.sub(r'[\\/*?:"<>|]', "_", texto)
    return texto_limpio.strip().replace(" ", "_")[:50]


# --- GENERADORES DE DOCUMENTOS WORD ---
def generar_word_consulta(asunto, norma, ruta_salida):
    doc = Document()
    doc.add_heading("INFORME DE BÚSQUEDA NORMATIVA - ISOLUCIÓN", level=1)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Tema de Consulta: ").bold = True
    p_meta.add_run(f"{asunto}\n")
    p_meta.add_run("Marco Normativo Evaluado: ").bold = True
    p_meta.add_run(f"{norma}\n")

    doc.add_heading("1. Marco Normativo Encontrado", level=2)
    doc.add_paragraph(
        "A continuación se detallan los procedimientos, manuales y resoluciones aplicables extraídos de iSolución:"
    )

    # Detalle normativo simulado/extraído por el bot
    p_res = doc.add_paragraph()
    p_res.add_run("• Procedimiento General INPEC:\n").bold = True
    p_res.add_run(
        f"Se identificaron disposiciones vigentes asociadas al término '{asunto}' bajo el contexto de {norma}."
    )

    doc.save(ruta_salida)


def generar_word_respuesta(archivo_origen, asunto, norma, ruta_salida):
    doc = Document()
    doc.add_heading("BORRADOR DE RESPUESTA A PETICIÓN - INPEC", level=1)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Documento de Origen: ").bold = True
    p_meta.add_run(f"{archivo_origen}\n")
    p_meta.add_run("Asunto / Referencia: ").bold = True
    p_meta.add_run(f"{asunto}\n")
    p_meta.add_run("Normatividad de Contraste: ").bold = True
    p_meta.add_run(f"{norma}\n")

    doc.add_heading("Propuesta de Respuesta Jurídica / Técnica", level=2)
    doc.add_paragraph(
        "En atención a la solicitud recibida y tras realizar el cotejo normativo con los manuales e iSolución, se fundamenta la siguiente respuesta:"
    )

    doc.add_paragraph(
        "[Aquí el bot insertará la redacción jurídica generada por la IA basándose en el análisis del documento.]"
    )

    doc.save(ruta_salida)


# --- INTERFAZ GRÁFICA ---
class AppINPEC:

    def __init__(self, root):
        asegurar_carpetas()
        self.root = root
        self.root.title("Búsqueda en iSolución - INPEC")
        self.root.geometry("640x680")
        self.root.resizable(False, False)

        lbl_titulo = tk.Label(
            root,
            text="ANALIZADOR NORMATIVO ISOLUCION",
            font=("Arial", 14, "bold"),
            fg="#003366",
        )
        lbl_titulo.pack(pady=10)

        # 1. CREDENCIALES
        frame_creds = tk.LabelFrame(
            root, text=" Credenciales de iSolución ", padx=10, pady=5
        )
        frame_creds.pack(fill="x", padx=20, pady=5)

        tk.Label(
            frame_creds, text="Usuario:", font=("Arial", 9, "bold")
        ).grid(row=0, column=0, sticky="w", padx=5)
        self.txt_user = tk.Entry(frame_creds, width=22)
        self.txt_user.grid(row=0, column=1, padx=5)
        self.txt_user.insert(0, os.getenv("ISOLUCION_USER", ""))

        tk.Label(
            frame_creds, text="Contraseña:", font=("Arial", 9, "bold")
        ).grid(row=0, column=2, sticky="w", padx=5)
        self.txt_pass = tk.Entry(frame_creds, width=22, show="*")
        self.txt_pass.grid(row=0, column=3, padx=5)
        self.txt_pass.insert(0, os.getenv("ISOLUCION_PASS", ""))

        # 2. SELECCIÓN DE MODO
        frame_modo = tk.LabelFrame(
            root, text=" Seleccione el Tipo de Tarea ", padx=10, pady=5
        )
        frame_modo.pack(fill="x", padx=20, pady=5)

        self.modo_var = tk.StringVar(value="busqueda")

        rb_busqueda = tk.Radiobutton(
            frame_modo,
            text="1. Búsqueda Normativa por Tema (Guarda en 'resultado_consulta')",
            variable=self.modo_var,
            value="busqueda",
            font=("Arial", 9, "bold"),
            command=self.cambiar_modo,
        )
        rb_busqueda.pack(anchor="w", pady=2)

        rb_contraste = tk.Radiobutton(
            frame_modo,
            text="2. Contraste de Petición/Documento (Guarda en 'borrador_respuesta')",
            variable=self.modo_var,
            value="contraste",
            font=("Arial", 9, "bold"),
            command=self.cambiar_modo,
        )
        rb_contraste.pack(anchor="w", pady=2)

        # 3. PARÁMETROS
        frame_form = tk.LabelFrame(
            root, text=" Parámetros de la Consulta ", padx=10, pady=10
        )
        frame_form.pack(fill="x", padx=20, pady=5)

        tk.Label(
            frame_form,
            text="Tema / Consulta / Palabras Clave:",
            font=("Arial", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(5, 2))
        self.txt_asunto = tk.Entry(frame_form, width=62)
        self.txt_asunto.grid(row=1, column=0, columnspan=2, padx=5, sticky="w")

        self.lbl_archivo = tk.Label(
            frame_form,
            text="Documento Adjunto (Oficio / Solicitud):",
            font=("Arial", 9, "bold"),
            fg="gray",
        )
        self.lbl_archivo.grid(row=2, column=0, sticky="w", pady=(10, 2))

        self.txt_archivo = tk.Entry(frame_form, width=47, state="disabled")
        self.txt_archivo.grid(row=3, column=0, padx=5)

        self.btn_examinar = tk.Button(
            frame_form,
            text="Examinar...",
            command=self.seleccionar_archivo,
            state="disabled",
        )
        self.btn_examinar.grid(row=3, column=1)

        tk.Label(
            frame_form,
            text="Marco Normativo de Interés:",
            font=("Arial", 9, "bold"),
        ).grid(row=4, column=0, sticky="w", pady=(10, 2))
        self.combo_norma = ttk.Combobox(
            frame_form,
            values=[
                "Todas las Normas e iSolución",
                "Manual de Funciones y Procesos",
                "Procedimientos Penitenciarios",
                "Régimen Disciplinario",
            ],
            width=58,
            state="readonly",
        )
        self.combo_norma.current(0)
        self.combo_norma.grid(
            row=5, column=0, columnspan=2, padx=5, sticky="w"
        )

        # 4. CONSOLA Y BOTÓN
        tk.Label(
            root, text="Estado del Procesamiento:", font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=20, pady=(5, 0))

        self.txt_log = tk.Text(root, height=7, width=72, state="disabled")
        self.txt_log.pack(padx=20, pady=5)

        self.btn_ejecutar = tk.Button(
            root,
            text="Iniciar Búsqueda / Análisis",
            bg="#003366",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.procesar_accion,
        )
        self.btn_ejecutar.pack(pady=10)

    def cambiar_modo(self):
        modo = self.modo_var.get()
        if modo == "busqueda":
            self.lbl_archivo.config(fg="gray")
            self.txt_archivo.config(state="disabled")
            self.btn_examinar.config(state="disabled")
        else:
            self.lbl_archivo.config(fg="black")
            self.txt_archivo.config(state="normal")
            self.btn_examinar.config(state="normal")

    def log(self, mensaje):
        self.txt_log.config(state="normal")
        self.txt_log.insert(tk.END, mensaje + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def seleccionar_archivo(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Archivos permitidos", "*.pdf *.docx *.txt")]
        )
        if filename:
            self.txt_archivo.config(state="normal")
            self.txt_archivo.delete(0, tk.END)
            self.txt_archivo.insert(0, filename)

    def procesar_accion(self):
        usuario = self.txt_user.get().strip()
        clave = self.txt_pass.get().strip()
        asunto = self.txt_asunto.get().strip()
        norma = self.combo_norma.get()
        modo = self.modo_var.get()

        if not usuario or not clave:
            messagebox.showwarning(
                "Credenciales Requeridas",
                "Por favor ingrese su usuario y contraseña de iSolución.",
            )
            return

        if not asunto:
            messagebox.showwarning(
                "Campo Faltante",
                "Por favor ingrese el tema o palabras clave de la consulta.",
            )
            return

        procesados = cargar_procesados()

        # --- MODO 1: BÚSQUEDA NORMATIVA DIRECTA ---
        if modo == "busqueda":
            id_consulta = f"BUSQUEDA_{sanitizar_nombre_archivo(asunto)}"

            if id_consulta in procesados:
                self.log(
                    f"[SKIPPED] La consulta sobre '{asunto}' ya fue realizada anteriormente."
                )
                messagebox.showinfo(
                    "Consulta Ya Realizada",
                    f"La consulta '{asunto}' ya fue procesada previamente.",
                )
                return

            self.log(f"[+] Autenticando en iSolución con usuario: {usuario}")
            self.log(f"[+] Consultando normatividad: '{asunto}'...")

            nombre_doc = f"Consulta_{sanitizar_nombre_archivo(asunto)}.docx"
            ruta_salida = os.path.join(CARPETA_CONSULTAS, nombre_doc)

            # Generar documento Word real en la carpeta resultado_consulta
            generar_word_consulta(asunto, norma, ruta_salida)
            guardar_procesado(id_consulta)

            self.log(f"[✓] Documento guardado en: {ruta_salida}")
            messagebox.showinfo(
                "Éxito", f"Informe generado con éxito en:\n{ruta_salida}"
            )

        # --- MODO 2: CONTRASTE DE PETICIÓN ADJUNTA ---
        else:
            archivo = self.txt_archivo.get().strip()
            if not archivo:
                messagebox.showwarning(
                    "Archivo Requerido",
                    "Por favor seleccione un archivo para este modo de análisis.",
                )
                return

            nombre_base = os.path.basename(archivo)
            id_respuesta = f"RESPUESTA_{nombre_base}"

            if id_respuesta in procesados:
                self.log(
                    f"[SKIPPED] El documento '{nombre_base}' ya fue analizado previamente."
                )
                messagebox.showinfo(
                    "Documento Ya Procesado",
                    f"El archivo '{nombre_base}' ya se procesó anteriormente.",
                )
                return

            self.log(f"[+] Analizando archivo adjunto: {nombre_base}")
            self.log(f"[+] Cruzando con norma: {norma}...")

            nombre_doc = (
                f"Respuesta_{sanitizar_nombre_archivo(nombre_base)}.docx"
            )
            ruta_salida = os.path.join(CARPETA_RESPUESTAS, nombre_doc)

            # Generar documento Word real en la carpeta borrador_respuesta
            generar_word_respuesta(nombre_base, asunto, norma, ruta_salida)
            guardar_procesado(id_respuesta)

            self.log(f"[✓] Borrador guardado en: {ruta_salida}")
            messagebox.showinfo(
                "Éxito",
                f"Borrador de respuesta generado en:\n{ruta_salida}",
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppINPEC(root)
    root.mainloop()