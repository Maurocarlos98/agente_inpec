import json
import os
from dotenv import load_dotenv

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"


# --- 1. GESTIÓN DE ARCHIVOS PROCESADOS ---
def cargar_procesados():
    if os.path.exists(REGISTRO_PROCESADOS):
        with open(REGISTRO_PROCESADOS, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def guardar_procesado(id_documento):
    procesados = cargar_procesados()
    procesados.add(id_documento)
    with open(REGISTRO_PROCESADOS, "w", encoding="utf-8") as f:
        json.dump(list(procesados), f, ensure_ascii=False, indent=2)


# --- 2. CONFIGURACIÓN DINÁMICA DE BÚSQUEDA ---
def obtener_parametros_usuario():
    print("\n" + "=" * 40)
    print("  CONFIGURACIÓN DE BÚSQUEDA - INPEC")
    print("=" * 40)

    palabra_clave = input("1. Palabra clave / Término de búsqueda: ").strip()
    dependencia = (
        input("2. Dependencia / Área (Opcional, Enter para omitir): ")
        .strip()
        or "TODAS"
    )
    limite = input("3. Cantidad máxima de documentos a procesar [5]: ").strip()

    cant_limite = int(limite) if limite.isdigit() else 5

    return {
        "kw": palabra_clave,
        "dependencia": dependencia,
        "limite": cant_limite,
    }


# --- 3. LÓGICA DE FILTRADO Y PROCESAMIENTO ---
def procesar_documentos(documentos_encontrados, filtro):
    procesados = cargar_procesados()
    contador = 0

    print(f"\n[+] Buscando coincidencias para '{filtro['kw']}'...")

    for doc in documentos_encontrados:
        doc_id = doc.get("id")  # ID único del documento o URL en iSolucion

        # Omitir si ya fue analizado previamente
        if doc_id in procesados:
            print(
                f"[SKIPPED] El documento '{doc.get('nombre')}' ya fue procesado anteriormente."
            )
            continue

        if contador >= filtro["limite"]:
            print(
                f"\n[!] Se alcanzó el límite configurado ({filtro['limite']} documentos)."
            )
            break

        # Ejecutar análisis con DeepSeek
        print(f"\n[-->] Analizando nuevo documento: {doc.get('nombre')}...")
        # ... Aquí va tu llamada a la API de DeepSeek ...

        # Marcar como procesado
        guardar_procesado(doc_id)
        contador += 1


if __name__ == "__main__":
    config = obtener_parametros_usuario()

    # Ejemplo de lista simulada de documentos obtenidos de iSolución
    docs_ejemplo = [
        {"id": "DOC-001", "nombre": "Informe_Disciplinario_2026_01.docx"},
        {"id": "DOC-002", "nombre": "Acta_Seguridad_2026_02.docx"},
        {"id": "DOC-003", "nombre": "Resolucion_Traslado_2026_03.docx"},
    ]

    procesar_documentos(docs_ejemplo, config)