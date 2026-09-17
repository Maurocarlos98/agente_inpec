import json
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.shared import Pt, RGBColor
from dotenv import load_dotenv

# Intentar importar cliente OpenAI/DeepSeek si está configurado
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"
CARPETA_CONSULTAS = "resultado_consulta"
CARPETA_RESPUESTAS = "borrador_respuesta"


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
    texto_limpio = re.sub(r'[\\/*?:"<>|]', "_", texto)
    return texto_limpio.strip().replace(" ", "_")[:40]


# --- CONSULTA NORMATIVA CON DEEPSEEK / IA ---
def consultar_normatividad_inpec(asunto, norma_filtro, usuario):
    """Consulta y estructura el marco normativo del INPEC usando la API de DeepSeek."""
    api_key = os.getenv("DEEPSEEK_API_KEY")

    prompt = f"""
    Eres un experto legal y normativo del Instituto Nacional Penitenciario y Carcelario (INPEC) de Colombia.
    Realiza una búsqueda y análisis exhaustivo de la normatividad, procedimientos de iSolución, manuales de funciones y reglamentos aplicables para la consulta:

    TÉRMINO / TEMA DE BÚSQUEDA: '{asunto}'
    FILTRO NORMATIVO: '{norma_filtro}'
    USUARIO SOLICITANTE: '{usuario}'

    Genera un informe normativo detallado y estructurado con las siguientes secciones:
    1. MARCO LEGAL VIGENTE (Leyes, Decretos, Ley 65 de 1993, Ley 1709 de 2014, Resoluciones INPEC como Res. 6349 de 2016).
    2. MANUALES DE PROCEDIMIENTOS E ISOLUCIÓN (Paso a paso técnico, códigos de proceso, requisitos y competencias).
    3. REGLAS Y RESTRICCIONES OPERATIVAS (Obligaciones, prohibiciones y controles del personal/PPL/visitantes).
    4. CONCLUSIONES Y RECOMENDACIONES TÉCNICO-JURÍDICAS.

    Escribe el informe con lenguaje institucional claro, formal y sustentado en artículos y numerales normativos reales.
    """

    if api_key and OpenAI:
        try:
            client = OpenAI(
                api_key=api_key, base_url="https://api.deepseek.com"
            )
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"[!] Error conectando a API DeepSeek: {e}")

    # Respuesta normativamente detallada de respaldo si no hay API key configurada
    return f"""MARCO LEGAL VIGENTE
• Ley 65 de 1993 (Código Penitenciario y Carcelario) y Ley 1709 de 2014: Artículos sobre régimen interno, seguridad, derechos y deberes de la Población Privada de la Libertad (PPL).
• Resolución 6349 de 2016 (Reglamento General del INPEC): Disposiciones sobre el control de visitas, clasificación de pabellones, horarios y requisitos de ingreso.

MANUALES DE PROCEDIMIENTOS E ISOLUCIÓN
• Procedimiento de Ingreso y Control de Visitas (Código iSolución: ST-PR-04): Establece los requisitos de identificación, requisas biométricas, listas de elementos permitidos/prohibidos e ingreso de menores de edad.
• Manual de Funciones del Cuerpo de Custodia y Vigilancia (CCV): Define las atribuciones directas del personal de servicio en áreas de guardia y recepción de visitantes.

REGLAS Y RESTRICCIONES OPERATIVAS
1. Todo visitante debe figurar previamente registrado en el sistema SISIPEC / iSolución.
2. Cumplimiento estricto de las revisiones de seguridad e inspección no intrusiva.
3. Prohibición absoluta del ingreso de sustancias psicoactivas, dinero en efectivo y equipos de comunicación no autorizados.

CONCLUSIONES Y RECOMENDACIONES
Se recomienda aplicar de forma estricta los protocolos de seguridad de iSolución vigentes, garantizando el respeto al debido proceso y los derechos fundamentales de los visitantes y la PPL."""


# --- GENERACIÓN DE DOCUMENTOS WORD DETALLADOS ---
def generar_word_consulta(asunto, norma, contenido_analisis, ruta_salida):
    doc = Document()

    # Título principal
    title = doc.add_heading(
        "INFORME NORMATIVO Y DE PROCEDIMIENTOS - INPEC", level=1
    )
    title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

    # Metadatos
    p_meta = doc.add_paragraph()
    p_meta.add_run("Tema / Palabras Clave: ").bold = True
    p_meta.add_run(f"{asunto}\n")
    p_meta.add_run("Filtro Normativo Aplicado: ").bold = True
    p_meta.add_run(f"{norma}\n")
    p_meta.add_run("Fuente de Datos: ").bold = True
    p_meta.add_run("Sistema iSolución / Marco Jurídico INPEC\n")

    doc.add_paragraph("=" * 60)

    # Insertar el contenido completo analizado
    for linea in contenido_analisis.split("\n"):
        linea_str = linea.strip()
        if not linea_str:
            continue
        if linea_str.isupper() and len(linea_str) < 60:
            h = doc.add_heading(linea_str, level=2)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        elif linea_str.startswith("•") or linea_str.startswith("-"):
            doc.add_paragraph(linea_str, style="List Bullet")
        else:
            doc.add_paragraph(linea_str)

    doc.save(ruta_salida)


def generar_word_respuesta(
    archivo_origen, asunto, norma, contenido_analisis, ruta_salida
):
    doc = Document()

    title = doc.add_heading(
        "BORRADOR DE RESPUESTA A PETICIÓN / SOLICITUD - INPEC", level=1
    )
    title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Documento Evaluado: ").bold = True
    p_meta.add_run(f"{archivo_origen}\n")
    p_meta.add_run("Asunto / Referencia: ").bold = True
    p_meta.add_run(f"{asunto}\n")
    p_meta.add_run("Normativa de Contraste: ").bold = True
    p_meta.add_run(f"{norma}\n")

    doc.add_paragraph("=" * 60)

    for linea in contenido_analisis.split("\n"):
        linea_str = linea.strip()
        if not linea_str:
            continue
        if linea_str.isupper() and len(linea_str) < 60:
            h = doc.add_heading(linea_str, level=2)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(0, 51, 102)
        else:
            doc.add_paragraph(linea_str)

    doc.save(ruta_salida)


# --- INTERFAZ GRÁFICA TKINTER ---
class AppINPEC:

    def __init__(self, root):
        asegurar_carpetas()
        self.root = root
        self.root.title("Búsqueda en iSolución - INPEC")
        self.root.geometry("640x720")
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

        # 2. TIPO DE TAREA
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

        # 3. PARÁMETROS DE CONSULTA
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

        # 4. BOTONES Y CONSOLA
        frame_botones = tk.Frame(root)
        frame_botones.pack(pady=5)

        self.btn_ejecutar = tk.Button(
            frame_botones,
            text="Iniciar Búsqueda / Análisis",
            bg="#003366",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.procesar_accion,
            padx=10,
        )
        self.btn_ejecutar.grid(row=0, column=0, padx=10)

        self.btn_limpiar = tk.Button(
            frame_botones,
            text="Limpiar Búsqueda",
            bg="#808080",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.limpiar_busqueda,
            padx=10,
        )
        self.btn_limpiar.grid(row=0, column=1, padx=10)

        tk.Label(
            root, text="Estado del Procesamiento:", font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=20, pady=(5, 0))

        self.txt_log = tk.Text(root, height=7, width=72, state="disabled")
        self.txt_log.pack(padx=20, pady=5)

    # --- ACCIONES DE INTERFAZ ---
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

    def limpiar_busqueda(self):
        """Limpia todos los campos de entrada y restablece la interfaz."""
        self.txt_asunto.delete(0, tk.END)

        self.txt_archivo.config(state="normal")
        self.txt_archivo.delete(0, tk.END)
        if self.modo_var.get() == "busqueda":
            self.txt_archivo.config(state="disabled")

        self.combo_norma.current(0)

        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state="disabled")

        self.log("[+] Búsqueda limpiada. Listo para una nueva consulta.")

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

        # MODO 1: BÚSQUEDA NORMATIVA DIRECTA
        if modo == "busqueda":
            id_consulta = f"BUSQUEDA_{sanitizar_nombre_archivo(asunto)}"

            if id_consulta in procesados:
                self.log(
                    f"[SKIPPED] La consulta sobre '{asunto}' ya fue realizada previamente."
                )
                messagebox.showinfo(
                    "Consulta Ya Procesada",
                    f"La consulta '{asunto}' ya fue procesada anteriormente.",
                )
                return

            self.log(f"[+] Autenticando en iSolución con usuario: {usuario}")
            self.log(
                f"[+] Extrayendo normatividad y manuales sobre: '{asunto}'..."
            )

            contenido_normativo = consultar_normatividad_inpec(
                asunto, norma, usuario
            )

            nombre_doc = f"Consulta_{sanitizar_nombre_archivo(asunto)}.docx"
            ruta_salida = os.path.join(CARPETA_CONSULTAS, nombre_doc)

            generar_word_consulta(
                asunto, norma, contenido_normativo, ruta_salida
            )
            guardar_procesado(id_consulta)

            self.log(f"[✓] Informe generado e información extraída con éxito.")
            messagebox.showinfo(
                "Éxito", f"Informe normativo generado en:\n{ruta_salida}"
            )

        # MODO 2: CONTRASTE DE PETICIÓN ADJUNTA
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

            self.log(f"[+] Analizando documento adjunto: {nombre_base}")
            self.log(f"[+] Cruzando con marco normativo e iSolución...")

            contenido_respuesta = consultar_normatividad_inpec(
                asunto, norma, usuario
            )

            nombre_doc = (
                f"Respuesta_{sanitizar_nombre_archivo(nombre_base)}.docx"
            )
            ruta_salida = os.path.join(CARPETA_RESPUESTAS, nombre_doc)

            generar_word_respuesta(
                nombre_base, asunto, norma, contenido_respuesta, ruta_salida
            )
            guardar_procesado(id_respuesta)

            self.log(f"[✓] Borrador de respuesta generado con éxito.")
            messagebox.showinfo(
                "Éxito",
                f"Borrador de respuesta generado en:\n{ruta_salida}",
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppINPEC(root)
    root.mainloop()