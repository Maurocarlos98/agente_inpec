import json
import os
from dotenv import load_dotenv

load_dotenv()

REGISTRO_PROCESADOS = "procesados.json"


# --- 1. CONTROL DE DUPLICADOS ---
def cargar_procesados():
    """Carga el conjunto de IDs o radicados ya analizados."""
    if os.path.exists(REGISTRO_PROCESADOS):
        try:
            with open(REGISTRO_PROCESADOS, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception as e:
            print(f"[!] Error al leer {REGISTRO_PROCESADOS}: {e}")
            return set()
    return set()


def guardar_procesado(id_documento):
    """Guarda un ID o radicado procesado en el archivo local."""
    procesados = cargar_procesados()
    procesados.add(id_documento)
    with open(REGISTRO_PROCESADOS, "w", encoding="utf-8") as f:
        json.dump(list(procesados), f, ensure_ascii=False, indent=2)


# --- 2. CONFIGURACIÓN DINÁMICA POR USUARIO ---
def solicitar_configuracion_usuario():
    """Solicita los criterios de búsqueda al usuario en cada ejecución."""
    print("\n" + "=" * 50)
    print("      CONFIGURACIÓN DE BÚSQUEDA - AGENTE INPEC")
    print("=" * 50)

    termino = input("1. Ingrese el término o asunto a buscar: ").strip()

    print("\nTipos de documento disponibles:")
    print("  [1] Actas")
    print("  [2] Resoluciones")
    print("  [3] Informes")
    print("  [4] Todos")
    opcion_tipo = input("Seleccione una opción [4]: ").strip() or "4"

    tipos_map = {
        "1": "Actas",
        "2": "Resoluciones",
        "3": "Informes",
        "4": "Todos",
    }
    tipo_doc = tipos_map.get(opcion_tipo, "Todos")

    limite = input("\nCantidad máxima de documentos a procesar [5]: ").strip()
    cant_limite = int(limite) if limite.isdigit() and int(limite) > 0 else 5

    return {
        "termino": termino,
        "tipo_documento": tipo_doc,
        "limite": cant_limite,
    }


# --- 3. PROCESAMIENTO PRINCIPAL ---
def ejecutar_bot():
    config = solicitar_configuracion_usuario()
    procesados = cargar_procesados()

    print("\n" + "-" * 50)
    print(f"  Iniciando búsqueda para: '{config['termino']}'")
    print(f"  Filtro de documento: {config['tipo_documento']}")
    print(f"  Límite configurado: {config['limite']}")
    print("-" * 50 + "\n")

    # TODO: Aquí va la conexión a iSolución / extracción de documentos.
    # Supongamos que obtenemos la lista 'documentos_encontrados' desde iSolución:
    documentos_encontrados = [
        {
            "id": "RAD-2026-001",
            "titulo": "Informe_Disciplinario_Enero.pdf",
            "tipo": "Informes",
        },
        {
            "id": "RAD-2026-002",
            "titulo": "Resolucion_Traslado_042.pdf",
            "tipo": "Resoluciones",
        },
        {
            "id": "RAD-2026-003",
            "titulo": "Acta_Comite_Seguridad.pdf",
            "tipo": "Actas",
        },
    ]

    procesados_en_sesion = 0

    for doc in documentos_encontrados:
        doc_id = doc["id"]
        doc_titulo = doc["titulo"]

        # Verificar si se alcanzó el límite configurado por el usuario
        if procesados_en_sesion >= config["limite"]:
            print(
                f"\n[!] Se alcanzó el límite máximo configurado ({config['limite']} documentos)."
            )
            break

        # Control de descarte: verificar si ya fue procesado
        if doc_id in procesados:
            print(
                f"[SKIPPED] Omitiendo '{doc_titulo}' (Radicado {doc_id} ya procesado)."
            )
            continue

        print(
            f"\n[-->] Procesando documento nuevo: {doc_titulo} (ID: {doc_id})"
        )

        # -------------------------------------------------------------
        # AQUÍ SE EJECUTA LA ANALÍTICA CON DEEPSEEK Y LA GENERACIÓN WORD
        # -------------------------------------------------------------

        # Una vez completado con éxito, guardamos el ID para no repetirlo
        guardar_procesado(doc_id)
        procesados_en_sesion += 1
        print(f"[✓] Documento {doc_id} analizado y registrado correctamente.")

    print(
        f"\n[+] Proceso finalizado. Documentos nuevos procesados en esta sesión: {procesados_en_sesion}"
    )


if __name__ == "__main__":
    ejecutar_bot()