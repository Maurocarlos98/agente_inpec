import os
import asyncio
import docx
from dotenv import load_dotenv
from playwright.async_api import async_playwright
from langchain_openai import ChatOpenAI
from pypdf import PdfReader

# Cargar variables del archivo .env
load_dotenv()

# Configuración del modelo leyendo la clave desde el entorno
llm = ChatOpenAI(
    base_url="https://api.deepseek.com/v1",
    api_key=os.getenv("DEEPSEEK_API_KEY"),
    model="deepseek-chat"
)
import asyncio
import docx
from playwright.async_api import async_playwright
from langchain_openai import ChatOpenAI
from pypdf import PdfReader
# Configuración del modelo DeepSeek
llm = ChatOpenAI(
    base_url="https://api.deepseek.com/v1",
    api_key="sk-c1dcd2d7a77e462d8f3f159fcfed73d3",  # Reemplaza por tu API Key real
    model="deepseek-chat"
)
def leer_oficio_peticion(ruta_pdf):
    """Extrae el texto del oficio en PDF"""
    reader = PdfReader(ruta_pdf)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text() or ""
    return texto

def guardar_en_word(texto_respuesta, nombre_salida="Borrador_Respuesta_INPEC.docx"):
    """Guarda la respuesta analizada en formato Word (.docx)"""
    doc = docx.Document()
    doc.add_heading('BORRADOR DE RESPUESTA OFICIAL - INPEC', 0)
    
    for p in texto_respuesta.split('\n'):
        if p.strip():
            doc.add_paragraph(p.strip())
            
    doc.save(nombre_salida)
    print(f"📄 Documento de Word generado con éxito: {nombre_salida}")

async def ejecutar_analisis():
    nombre_archivo = "oficio_peticion.pdf"
    print(f"📄 1. Leyendo el archivo del oficio: {nombre_archivo}...")
    
    try:
        texto_oficio = leer_oficio_peticion(nombre_archivo)
        print("✅ Texto del oficio cargado correctamente.")
    except Exception as e:
        print(f"❌ Error al leer el PDF: {e}")
        return

    texto_marco_normativo = ""

    async with async_playwright() as p:
        # Abrimos navegador
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()

        print("🌐 2. Conectando a ISolución INPEC...")
        await page.goto("https://isolucion.inpec.gov.co/Isolucion4Inpec/PaginaLogin.aspx", timeout=60000)
        
        # Inicio de sesión (Filtrando solo elementos visibles)
        print("🔑 3. Ingresando credenciales...")
        
        # Selecciona el primer campo de texto VISIBLE para el usuario
        campo_usuario = page.locator("input[type='text']:visible, input[id*='TxtUsuario']:visible").first
        await campo_usuario.wait_for(state="visible", timeout=30000)
        await campo_usuario.fill("consulta")

        # Selecciona el campo de contraseña VISIBLE
        campo_clave = page.locator("input[type='password']:visible, input[id*='TxtClave']:visible").first
        await campo_clave.fill("123456")

        # Clic en el botón ingresar
        bot_ingresar = page.locator("input[type='submit']:visible, button:visible, input[name*='Btn']:visible").first
        await bot_ingresar.click()
        
        print("✅ 4. Sesión iniciada. Esperando carga del panel de control de ISolución...")
        await page.wait_for_timeout(8000) # Espera a que carguen los marcos internos

        print("🔍 5. Analizando interfaz de ISolución...")
        
        # Extraer contenido de todos los marcos activos dentro de ISolución
        for frame in page.frames:
            try:
                contenido_frame = await frame.content()
                if "seguridad" in contenido_frame.lower() or "documento" in contenido_frame.lower():
                    texto_marco_normativo += await frame.inner_text("body") + "\n"
            except Exception:
                continue

        if not texto_marco_normativo.strip():
            texto_marco_normativo = await page.inner_text("body")

        print("✅ Marco normativo e interfaz extraídos con éxito.")
        await browser.close()

    print("🤖 6. Contrastando información y redactando el borrador con DeepSeek...")
    
    prompt_analisis = f"""
    Actúa como Asesor Jurídico y Oficial de Seguridad de la Información del INPEC.

    1. OFICIO RECIBIDO (PETICIÓN A EVALUAR):
    \"\"\"{texto_oficio}\"\"\"

    2. MARCO DE INFORMACIÓN OBTENIDO DE ISOLUCIÓN INPEC:
    \"\"\"{texto_marco_normativo[:4000]}\"\"\"

    3. POLÍTICA Y LINEAMIENTO DE SEGURIDAD (MSPI INPEC):
    Por normatividad estricta de Seguridad de la Información, está PROHIBIDO agregar botones, enlaces o accesos directos hacia dominios externos o no controlados (como 'dominio.com') en plataformas o aplicativos institucionales, debido a riesgos de suplantación (Phishing), redireccionamiento malicioso e incumplimiento de políticas del Modelo de Seguridad y Privacidad de la Información (MSPI).

    TAREA:
    Genera un Borrador de Oficio de Respuesta Oficial formal que argumente técnicamente y jurídicamente la NO VIABILIDAD de la solicitud del usuario de anexar el botón/enlace a dominio.com, basándote en la imposibilidad de vulnerar las políticas de seguridad de la entidad.

    ESTRUCTURA DEL OFICIO:
    - Referencia y asunto institucional.
    - Antecedentes del requerimiento.
    - Fundamento Jurídico y Técnico de Seguridad de la Información (citar riesgos cibernéticos, normatividad de seguridad digital y protección de datos).
    - Conclusión clara sobre la improcedencia de la solicitud.
    """

    respuesta = llm.invoke(prompt_analisis)
    
    # Exportación del documento a Word
    guardar_en_word(respuesta.content, "Borrador_Respuesta_INPEC.docx")
    print("\n🎉 ¡Proceso finalizado con éxito! El archivo Word está listo en la carpeta 'agente_inpec'.")

if __name__ == "__main__":
    asyncio.run(ejecutar_analisis())
