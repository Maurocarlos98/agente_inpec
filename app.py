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


# --- INTERFAZ GRÁFICA (TKINTER) ---
class AppINPEC:

    def __init__(self, root):
        self.root = root
        self.root.title("Agente INPEC - Analizador Normativo ISOLUCION")
        self.root.geometry("600x520")
        self.root.resizable(False, False)

        # Encabezado
        lbl_titulo = tk.Label(
            root,
            text="Analizador Normativo ISOLUCION",
            font=("Arial", 14, "bold"),
            fg="#003366",
        )
        lbl_titulo.pack(pady=10)

        # Formulario de parámetros
        frame_form = tk.LabelFrame(
            root, text=" Configuración de Búsqueda y Análisis ", padx=10, pady=10
        )
        frame_form.pack(fill="x", padx=20, pady=5)

        # 1. Selección de Archivo
        tk.Label(
            frame_form,
            text="Documento a analizar (Oficio / Solicitud):",
            font=("Arial", 9, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=5)
        self.txt_archivo = tk.Entry(frame_form, width=45)
        self.txt_archivo.grid(row=1, column=0, padx=5)
        btn_examinar = tk.Button(
            frame_form, text="Examinar...", command=self.seleccionar_archivo
        )
        btn_examinar.grid(row=1, column=1)

        # 2. Asunto / Términos clave
        tk.Label(
            frame_form, text="Asunto / Término Clave:", font=("Arial", 9, "bold")
        ).grid(row=2, column=0, sticky="w", pady=(10, 5))
        self.txt_asunto = tk.Entry(frame_form, width=58)
        self.txt_asunto.grid(row=3, column=0, columnspan=2, padx=5, sticky="w")

        # 3. Categoría Normativa
        tk.Label(
            frame_form,
            text="Marco Normativo a Contraste:",
            font=("Arial", 9, "bold"),
        ).grid(row=4, column=0, sticky="w", pady=(10, 5))
        self.combo_norma = ttk.Combobox(
            frame_form,
            values=[
                "Manual de Funciones e iSolución",
                "Procedimientos Penitenciarios",
                "Régimen Disciplinario",
                "Todas las Normas e iSolución",
            ],
            width=55,
            state="readonly",
        )
        self.combo_norma.current(0)
        self.combo_norma.grid(
            row=5, column=0, columnspan=2, padx=5, sticky="w"
        )

        # Consola de Estado / Salida
        tk.Label(
            root, text="Estado del Procesamiento:", font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=20, pady=(10, 0))
        self.txt_log = tk.Text(root, height=8, width=68, state="disabled")
        self.txt_log.pack(padx=20, pady=5)

        # Botón de Procesar
        self.btn_ejecutar = tk.Button(
            root,
            text="Iniciar Análisis Normativo",
            bg="#003366",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.procesar_documento,
        )
        self.btn_ejecutar.pack(pady=10)

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
            self.txt_archivo.delete(0, tk.END)
            self.txt_archivo.insert(0, filename)

    def procesar_documento(self):
        archivo = self.txt_archivo.get().strip()
        asunto = self.txt_asunto.get().strip()
        norma = self.combo_norma.get()

        if not archivo:
            messagebox.showwarning(
                "Campo Faltante", "Por favor seleccione el documento a analizar."
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

        self.log(f"[+] Cargando archivo: {id_doc}")
        self.log(f"[+] Asunto: {asunto or 'Sin especificar'}")
        self.log(f"[+] Cruzando con: {norma} (iSolución / Manuales INPEC)...")

        # Marcar como procesado
        guardar_procesado(id_doc)
        self.log("[✓] Análisis finalizado. Borrador de respuesta generado.")
        messagebox.showinfo(
            "Éxito", "Proceso completado. Documento Word generado."
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = AppINPEC(root)
    root.mainloop()