import asyncio
from playwright.async_api import async_playwright
from langchain_openai import ChatOpenAI
from pypdf import PdfReader

# Configuración del modelo DeepSeek
llm = ChatOpenAI(
    base_url="https://api.deepseek.com/v1",
    api_key="TU_API_KEY_DEEPSEEK_AQUI",  # Reemplaza por tu API Key real
    model="deepseek-chat"
)

def leer_oficio_peticion(ruta_pdf):
    """Extrae el texto del oficio que subiste"""
    reader = PdfReader(ruta_pdf)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text()
    return texto

async def ejecutar_analisis():
    # 1. Leer el archivo del oficio subido
    nombre_archivo = "oficio_peticion.pdf"  # Asegúrate de poner el nombre exacto de tu archivo
    print(f"📄 Leyendo el documento: {nombre_archivo}...")
    texto_oficio = leer_oficio_peticion(nombre_archivo)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # 2. Iniciar sesión en ISolución
        print("🌐 Conectando a ISolución INPEC...")
        await page.goto("https://isolucion.inpec.gov.co/Isolucion4Inpec/PaginaLogin.aspx")
        await page.wait_for_load_state("networkidle")

        await page.fill("input[name*='TxtUsuario']", "consulta")
        await page.fill("input[name*='TxtClave']", "123456")
        await page.click("input[type='submit'], button[type='submit'], input[name*='Btn']")
        await page.wait_for_load_state("networkidle")
        print("✅ Sesión iniciada.")

        # 3. Extraer contenido visible del portal sobre seguridad
        print("🔍 Extrayendo el marco normativo de ISolución...")
        await asyncio.sleep(5)
        marco_normativo_html = await page.content()
        await browser.close()

    # 4. Generación del borrador de respuesta con DeepSeek
    print("🤖 Generando borrador de respuesta juridico-técnica con DeepSeek...")
    
    prompt_analisis = f"""
    Actúa como un Asesor Jurídico y Oficial de Seguridad de la Información del INPEC.
    
    CONTEXTO Y SOLICITUD:
    Hemos recibido una petición/oficio con el siguiente texto:
    \"\"\"{texto_oficio}\"\"\"

    REGLAS Y MARCO NORMATIVO DE LA ENTIDAD:
    Las políticas institucionales de seguridad de la información del INPEC prohíben la inclusión de enlaces directos, redirigimientos no autorizados o botones a dominios externos no controlados (como 'dominio.com') por riesgos de Phishing, exfiltración de datos, falta de controles de autenticación y vulneración del Modelo de Seguridad y Privacidad de la Información (MSPI).

    TAREA:
    Redacta un BORRADOR DE OFICIO DE RESPUESTA oficial, formal y riguroso argumentando jurídicamente y técnicamente POR QUÉ NO ES POSIBLE acceder a la petición de colocar el botón hacia 'dominio.com'.

    ESTRUCTURA DEL BORRADOR:
    1. Encabezado institucional y referencia a la petición.
    2. Resumen de la solicitud recibida.
    3. Sustento Técnico-Jurídico (Cita principios de Seguridad de la Información, riesgos cibernéticos, integridad de datos y normatividad del MSPI/INPEC).
    4. Conclusión y decisión de NO viabilidad.
    """

    respuesta = llm.invoke(prompt_analisis)
    
    # 5. Guardar la respuesta en un archivo de texto
    with open("Borrador_Respuesta_INPEC.txt", "w", encoding="utf-8") as f:
        f.write(respuesta.content)

    print("\n✅ ¡Análisis completado! El borrador se guardó en 'Borrador_Respuesta_INPEC.txt'")

if __name__ == "__main__":
    asyncio.run(ejecutar_analisis())