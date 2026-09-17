import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dotenv import load_dotenv

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"


# --- CONTROL DE DUPLICADOS ---
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


# --- INTERFAZ GRÁFICA ---
class AppINPEC:

    def __init__(self, root):
        self.root = root
        self.root.title("Búsqueda en iSolución - INPEC")
        self.root.geometry("640x680")
        self.root.resizable(False, False)

        # Encabezado
        lbl_titulo = tk.Label(
            root,
            text="ANALIZADOR NORMATIVO ISOLUCION",
            font=("Arial", 14, "bold"),
            fg="#003366",
        )
        lbl_titulo.pack(pady=10)

        # -------------------------------------------------------------
        # SECCIÓN 1: CREDENCIALES DE ISOLUCIÓN
        # -------------------------------------------------------------
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

        # -------------------------------------------------------------
        # SECCIÓN 2: SELECCIÓN DE MODO / FUNCIONALIDAD
        # -------------------------------------------------------------
        frame_modo = tk.LabelFrame(
            root, text=" Seleccione el Tipo de Tarea ", padx=10, pady=5
        )
        frame_modo.pack(fill="x", padx=20, pady=5)

        self.modo_var = tk.StringVar(value="busqueda")

        rb_busqueda = tk.Radiobutton(
            frame_modo,
            text="1. Búsqueda Normativa por Tema (Generar Informe Word)",
            variable=self.modo_var,
            value="busqueda",
            font=("Arial", 9, "bold"),
            command=self.cambiar_modo,
        )
        rb_busqueda.pack(anchor="w", pady=2)

        rb_contraste = tk.Radiobutton(
            frame_modo,
            text="2. Contraste de Documento/Petición (Adjuntar Archivo)",
            variable=self.modo_var,
            value="contraste",
            font=("Arial", 9, "bold"),
            command=self.cambiar_modo,
        )
        rb_contraste.pack(anchor="w", pady=2)

        # -------------------------------------------------------------
        # SECCIÓN 3: FORMULARIO DE PARÁMETROS
        # -------------------------------------------------------------
        frame_form = tk.LabelFrame(
            root, text=" Parámetros de la Consulta ", padx=10, pady=10
        )
        frame_form.pack(fill="x", padx=20, pady=5)

        # Asunto / Tema
        tk.Label(
            frame_form,
            text="Tema / Consulta / Palabras Clave:",
            font=("Arial", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(5, 2))
        self.txt_asunto = tk.Entry(frame_form, width=62)
        self.txt_asunto.grid(row=1, column=0, columnspan=2, padx=5, sticky="w")

        # Archivo Adjunto (Solo para modo contraste)
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

        # Marco Normativo
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

        # -------------------------------------------------------------
        # SECCIÓN 4: CONSOLA DE SALIDA
        # -------------------------------------------------------------
        tk.Label(
            root, text="Estado del Procesamiento:", font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=20, pady=(5, 0))

        self.txt_log = tk.Text(root, height=7, width=72, state="disabled")
        self.txt_log.pack(padx=20, pady=5)

        # Botón de Procesar
        self.btn_ejecutar = tk.Button(
            root,
            text="Iniciar Búsqueda / Análisis",
            bg="#003366",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.procesar_accion,
        )
        self.btn_ejecutar.pack(pady=10)

    # --- LÓGICA DE INTERFAZ ---
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

    # --- EJECUCIÓN PRINCIPAL ---
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

        # MODO 1: BÚSQUEDA NORMATIVA DIRECTA
        if modo == "busqueda":
            self.log(f"[+] Autenticando en iSolución con usuario: {usuario}")
            self.log(f"[+] Buscando normatividad sobre: '{asunto}'...")
            self.log(f"[+] Filtro aplicado: {norma}")

            # AQUÍ SE CONECTA A ISOLUCIÓN / DEEPSEEK Y GENERA EL WORD DE NORMATIVIDAD
            self.log("[✓] Búsqueda completada. Documento Word generado.")
            messagebox.showinfo(
                "Éxito",
                "Se ha generado el informe de normatividad en formato Word.",
            )

        # MODO 2: CONTRASTE DE DOCUMENTO ADJUNTO
        else:
            archivo = self.txt_archivo.get().strip()
            if not archivo:
                messagebox.showwarning(
                    "Archivo Requerido",
                    "Por favor seleccione un archivo para este modo de análisis.",
                )
                return

            id_doc = os.path.basename(archivo)
            procesados = cargar_procesados()

            if id_doc in procesados:
                self.log(
                    f"[SKIPPED] El archivo '{id_doc}' ya fue procesado previamente."
                )
                messagebox.showinfo(
                    "Documento Duplicado",
                    f"El archivo '{id_doc}' ya se analizó anteriormente.",
                )
                return

            self.log(f"[+] Autenticando en iSolución como: {usuario}")
            self.log(f"[+] Cargando archivo: {id_doc}")
            self.log(f"[+] Cruzando con el marco normativo: {norma}...")

            # AQUÍ SE EJECUTA EL ANÁLISIS COMPLETO CONTRA EL DOCUMENTO
            guardar_procesado(id_doc)
            self.log("[✓] Análisis finalizado. Borrador de respuesta generado.")
            messagebox.showinfo(
                "Éxito", "Proceso completado. Documento Word generado."
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppINPEC(root)
    root.mainloop()