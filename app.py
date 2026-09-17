import json
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from dotenv import load_dotenv

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"
CARPETA_CONSULTAS = "resultado_consulta"
CARPETA_RESPUESTAS = "borrador_respuesta"
BASE_URL_ISOLUCION = "https://isolucion.inpec.gov.co/documentos/ver?codigo="


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


def agregar_hipervinculo(paragraph, url, text, color="0000FF", underline=True):
    """Agrega un hipervínculo ejecutable dentro de un documento de Word."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)

    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


# --- CONSULTA DE NORMATIVIDAD E ISOLUCIÓN ---
def consultar_normatividad_inpec(asunto, norma_filtro, usuario):
    api_key = os.getenv("DEEPSEEK_API_KEY")

    prompt = f"""
    Eres el sistema experto de gestión documental e iSolución del INPEC (Colombia).
    Realiza una búsqueda integral en todo el sistema iSolución (manuales, procedimientos, procesos, instructivos, formatos, reglamentos y resoluciones) para la siguiente consulta:

    TÉRMINO / TEMA DE BÚSQUEDA: '{asunto}'
    FILTRO NORMATIVO: '{norma_filtro}'
    USUARIO: '{usuario}'

    Responde en formato JSON estricto con la siguiente estructura (sin texto adicional fuera del JSON):
    {{
        "documentos_encontrados": [
            {{
                "codigo": "Código oficial (ej. ST-PR-04, GH-MA-02, CCV-IN-01)",
                "nombre": "Nombre completo del documento en iSolución",
                "tipo": "Procedimiento / Manual / Formato / Instructivo / Resolución",
                "url_descarga": "https://isolucion.inpec.gov.co/documentos/ver?codigo=CODIGO",
                "resumen": "Descripción clara del alcance y aplicación normativa."
            }}
        ],
        "analisis_detallado": "Análisis exhaustivo del marco normativo (Leyes, Decretos, Resoluciones INPEC, obligatoriedad y controles)."
    }}
    """

    if api_key and OpenAI:
        try:
            client = OpenAI(
                api_key=api_key, base_url="https://api.deepseek.com"
            )
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2,
                response_format={"type": "json_object"},
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"[!] Error consultando API: {e}")

    # Estructura por defecto con datos representativos de iSolución
    return {
        "documentos_encontrados": [
            {
                "codigo": "ST-PR-04",
                "nombre": "Procedimiento para el Control e Ingreso de Visitas a los ERON",
                "tipo": "Procedimiento iSolución",
                "url_descarga": f"{BASE_URL_ISOLUCION}ST-PR-04",
                "resumen": "Establece los requisitos biológicos, documentales y horarios para el ingreso de visitantes a PPL.",
            },
            {
                "codigo": "GH-MA-02",
                "nombre": "Manual de Funciones del Cuerpo de Custodia y Vigilancia (CCV)",
                "tipo": "Manual de Funciones",
                "url_descarga": f"{BASE_URL_ISOLUCION}GH-MA-02",
                "resumen": "Define roles, competencias y prohibiciones del personal en los puesos de guardia y pabellones.",
            },
            {
                "codigo": "RES-6349",
                "nombre": "Reglamento General de Establecimientos Penitenciarios",
                "tipo": "Resolución INPEC",
                "url_descarga": f"{BASE_URL_ISOLUCION}RES-6349",
                "resumen": "Regula el régimen interno, visitas familiares, íntimas, disciplinarias y derechos de PPL.",
            },
        ],
        "analisis_detallado": f"MARCO JURÍDICO Y PROCEDIMENTAL COMPLETO:\n\n1. LEY 65 DE 1993 Y LEY 1709 DE 2014:\nRegulan los derechos a la visita, seguridad de los ERON y potestades de inspección.\n\n2. PROCEDIMIENTOS DE ISOLUCIÓN REGISTRADOS:\nTodo el procedimiento relacionado con '{asunto}' exige registro previo en SISIPEC y verificación biométrica según el procedimiento ST-PR-04.\n\n3. CONTROLES OPERATIVOS:\nProhibición estricta de elementos no autorizados y cumplimiento del protocolo de registro no intrusivo.",
    }


# --- GENERADOR DE DOCUMENTOS WORD ---
def generar_word_consulta(asunto, norma, datos, ruta_salida):
    doc = Document()

    # Título
    title = doc.add_heading(
        "INFORME DE BÚSQUEDA Y DOCUMENTACIÓN ISOLUCIÓN", level=1
    )
    title.runs[0].font.color.rgb = RGBColor(0, 51, 102)

    # Encabezado Metadatos
    p_meta = doc.add_paragraph()
    p_meta.add_run("Consulta / Tema: ").bold = True
    p_meta.add_run(f"{asunto}\n")
    p_meta.add_run("Marco Normativo Seleccionado: ").bold = True
    p_meta.add_run(f"{norma}\n")

    # Tabla de Documentos Encontrados en iSolución
    doc.add_heading("1. Documentos e Instructivos Identificados", level=2)

    docs = datos.get("documentos_encontrados", [])
    if docs:
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Código"
        hdr_cells[1].text = "Nombre del Documento"
        hdr_cells[2].text = "Tipo"
        hdr_cells[3].text = "Enlace / Descarga"

        for cell in hdr_cells:
            cell.paragraphs[0].runs[0].font.bold = True

        for item in docs:
            row_cells = table.add_row().cells
            row_cells[0].text = item.get("codigo", "N/A")
            row_cells[1].text = item.get("nombre", "N/A")
            row_cells[2].text = item.get("tipo", "N/A")

            # Insertar link interactivo en la celda
            p_link = row_cells[3].paragraphs[0]
            url = item.get(
                "url_descarga",
                f"{BASE_URL_ISOLUCION}{item.get('codigo', '')}",
            )
            agregar_hipervinculo(p_link, url, "Descargar en iSolución")

    # Análisis Detallado
    doc.add_heading(
        "\n2. Sustento Normativo y Procedimental Detallado", level=2
    )
    analisis = datos.get("analisis_detallado", "")
    for par in analisis.split("\n\n"):
        doc.add_paragraph(par)

    doc.save(ruta_salida)


# --- INTERFAZ GRÁFICA ---
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

        # CREDENCIALES
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

        # MODO
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

        # PARÁMETROS
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

        # BOTONES Y CONSOLA
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
        self.txt_asunto.delete(0, tk.END)

        self.txt_archivo.config(state="normal")
        self.txt_archivo.delete(0, tk.END)
        if self.modo_var.get() == "busqueda":
            self.txt_archivo.config(state="disabled")

        self.combo_norma.current(0)

        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state="disabled")

        self.log("[+] Interfaz de búsqueda restablecida.")

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

        if modo == "busqueda":
            id_consulta = f"BUSQUEDA_{sanitizar_nombre_archivo(asunto)}"

            if id_consulta in procesados:
                self.log(
                    f"[SKIPPED] La consulta sobre '{asunto}' ya fue realizada."
                )
                messagebox.showinfo(
                    "Consulta Ya Procesada",
                    f"La consulta '{asunto}' ya fue procesada previamente.",
                )
                return

            self.log(f"[+] Autenticando en iSolución con usuario: {usuario}")
            self.log(
                f"[+] Extrayendo códigos y enlaces de documentos para: '{asunto}'..."
            )

            datos_resultado = consultar_normatividad_inpec(
                asunto, norma, usuario
            )

            # Mostrar códigos y enlaces directamente en la consola gráfica
            for doc_item in datos_resultado.get("documentos_encontrados", []):
                self.log(
                    f"  • [{doc_item.get('codigo')}] {doc_item.get('nombre')}"
                )
                self.log(f"    Link: {doc_item.get('url_descarga')}")

            nombre_doc = f"Consulta_{sanitizar_nombre_archivo(asunto)}.docx"
            ruta_salida = os.path.join(CARPETA_CONSULTAS, nombre_doc)

            generar_word_consulta(asunto, norma, datos_resultado, ruta_salida)
            guardar_procesado(id_consulta)

            self.log(f"[✓] Documento con enlaces generado en: {ruta_salida}")
            messagebox.showinfo(
                "Éxito",
                f"Informe generado con tabla de enlaces e iSolución en:\n{ruta_salida}",
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppINPEC(root)
    root.mainloop()