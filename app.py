import streamlit as st
import openai
from google import genai
import os
import tempfile
import base64
import copy
import io
import re
import datetime
import time
from docx import Document
from docx.oxml.ns import qn

def extraer_vista_previa(docx_bytes):
    """Devuelve una lista de líneas de texto (y marcadores de imagen) para
    mostrar como vista previa, sin necesidad de descargar el archivo."""
    doc = Document(io.BytesIO(docx_bytes))
    lineas = []
    for el in doc.element.body:
        if el.tag == qn('w:p'):
            tiene_imagen = len(el.findall('.//' + qn('w:drawing'))) > 0
            texto = ''.join(
                t.text or '' for r in el.findall(qn('w:r')) for t in [r.find(qn('w:t'))] if t is not None
            ).strip()
            if tiene_imagen:
                lineas.append("🖼️ *(imagen incluida aquí)*")
            elif texto:
                lineas.append(texto)
    return lineas


def extraer_solo_texto(docx_bytes):
    """Devuelve todo el texto de un .docx como un solo string, sin marcadores
    de imagen ni nada extra: solo las líneas de texto separadas por saltos de línea."""
    doc = Document(io.BytesIO(docx_bytes))
    lineas = []
    for el in doc.element.body:
        if el.tag == qn('w:p'):
            texto = ''.join(
                t.text or '' for r in el.findall(qn('w:r')) for t in [r.find(qn('w:t'))] if t is not None
            ).strip()
            if texto:
                lineas.append(texto)
    return '\n'.join(lineas)


CARPETA_HISTORIAL = "informes_generados"
os.makedirs(CARPETA_HISTORIAL, exist_ok=True)


def guardar_en_historial(nombre_archivo, contenido_bytes_o_texto):
    """Guarda una copia del informe generado en la carpeta de historial."""
    ruta = os.path.join(CARPETA_HISTORIAL, nombre_archivo)
    modo = "w" if isinstance(contenido_bytes_o_texto, str) else "wb"
    encoding = "utf-8" if modo == "w" else None
    with open(ruta, modo, encoding=encoding) as f:
        f.write(contenido_bytes_o_texto)


def _coincide_busqueda(nombre_archivo, termino_lower):
    """Revisa si un término de búsqueda coincide con el nombre del archivo
    O con el contenido del informe (texto, sin imágenes). Así se puede
    encontrar un informe por el nombre del paciente o cualquier palabra
    clave, aunque no esté en el nombre del archivo."""
    if termino_lower in nombre_archivo.lower():
        return True
    ruta = os.path.join(CARPETA_HISTORIAL, nombre_archivo)
    try:
        if nombre_archivo.endswith(".docx"):
            with open(ruta, "rb") as f:
                texto = extraer_solo_texto(f.read())
        else:
            with open(ruta, "r", encoding="utf-8") as f:
                texto = f.read()
        return termino_lower in texto.lower()
    except Exception:
        return False


def categorizar_archivo(nombre_archivo):
    """Determina a qué tipo de estudio pertenece un archivo del historial,
    según palabras clave presentes en su nombre."""
    nombre_lower = nombre_archivo.lower()
    if 'renal' in nombre_lower:
        return "Gammagrafía Renal"
    if 'ósea' in nombre_lower or 'osea' in nombre_lower:
        return "Gammagrafía Ósea"
    if 'tiroidea' in nombre_lower and ('tc_99m' in nombre_lower or 'tc-99m' in nombre_lower or 'tc99m' in nombre_lower):
        return "Gammagrafía Tiroidea (Tc-99m)"
    if 'iodo' in nombre_lower:
        return "Gammagrafía Tiroidea (Iodo-131)"
    if 'rastreo' in nombre_lower:
        return "Rastreo Corporal Total"
    if 'plantilla_libre' in nombre_lower:
        return "Plantilla Libre"
    return "Otros"


# Secciones esperadas por cada tipo de estudio (sin asteriscos, sin mayúsculas/minúsculas
# porque la comparación se hace ignorando eso). Se usan solo para la verificación automática
# opcional; no afectan la generación del informe.
SECCIONES_ESPERADAS = {
    "Gammagrafía Ósea": [
        "I. DATOS TÉCNICOS", "II. ANTECEDENTES CLÍNICOS",
        "III. HALLAZGOS", "IV. CONCLUSIÓN DIAGNÓSTICO",
        "VISTA ANTERIOR", "VISTA POSTERIOR",
    ],
    "Gammagrafía Tiroidea (Tc-99m)": [
        "I. DATOS TÉCNICOS", "II. ANTECEDENTES CLÍNICOS",
        "III. HALLAZGOS", "IV. IMPRESIÓN DIAGNÓSTICA",
    ],
    "Rastreo Corporal Total": ["HALLAZGOS", "CONCLUSIÓN"],
    "Gammagrafía Renal (DTPA/DMSA)": ["HALLAZGOS", "CONCLUSIÓN"],
}


def verificar_informe_basico(texto, tipo_estudio):
    """Revisa el informe generado con reglas simples y gratuitas (sin usar IA):
    busca marcadores de plantilla sin rellenar, texto sospechosamente corto, y
    secciones esperadas que falten. Devuelve una lista de avisos (vacía si no
    encontró nada). Esto NO reemplaza la revisión clínica de un transcriptor."""
    avisos = []
    texto_normalizado = texto.replace('**', '').upper()

    if '[' in texto and ']' in texto:
        avisos.append("Parece que quedó un marcador de plantilla sin completar (revisa si hay corchetes [ ] en el texto).")

    if len(texto.strip()) < 80:
        avisos.append("El informe generado es muy corto; revisa si el audio se transcribió bien.")

    secciones = SECCIONES_ESPERADAS.get(tipo_estudio, [])
    faltantes = [s for s in secciones if s not in texto_normalizado]
    if faltantes:
        avisos.append("No se encontraron estas secciones esperadas: " + ", ".join(faltantes) + ".")

    return avisos


def verificar_informe_renal_basico(merged_bytes, nombre_detectado):
    """Revisa el documento Word ya unido (Dr. Quijada + plantilla) con reglas
    simples y gratuitas (sin IA): si se detectó el nombre del paciente, si el
    contenido no quedó vacío/muy corto, y si se copió al menos una imagen.
    Devuelve una lista de avisos (vacía si no encontró nada)."""
    avisos = []

    if not nombre_detectado:
        avisos.append("No se detectó automáticamente el nombre del paciente; revisa que el documento final tenga el nombre correcto.")

    lineas_preview = extraer_vista_previa(merged_bytes)
    texto_solo = extraer_solo_texto(merged_bytes)

    if len(texto_solo.strip()) < 80:
        avisos.append("El documento final parece muy corto; revisa que el contenido del Dr. Quijada se haya insertado correctamente.")

    tiene_imagenes = any("🖼️" in linea for linea in lineas_preview)
    if not tiene_imagenes:
        avisos.append("No se detectó ninguna imagen en el documento final; verifica que las imágenes del estudio se hayan copiado.")

    return avisos


# Color y emoji distintivo por cada tipo de estudio, para las etiquetas del historial
# (colores claros/pastel para que resalten sobre fondo oscuro)
ESTILO_CATEGORIA = {
    "Gammagrafía Ósea": ("🦴", "#C4A98A"),
    "Gammagrafía Tiroidea (Tc-99m)": ("🦋", "#5EEAD4"),
    "Gammagrafía Tiroidea (Iodo-131)": ("🦋", "#67E8F9"),
    "Rastreo Corporal Total": ("🔎", "#FCD34D"),
    "Gammagrafía Renal": ("🫘", "#FCA5A5"),
    "Plantilla Libre": ("📝", "#CBD5E1"),
    "Otros": ("📄", "#CBD5E1"),
}


def contar_estudios_por_tipo():
    """Cuenta cuántos archivos hay en el historial por cada tipo de estudio."""
    archivos = os.listdir(CARPETA_HISTORIAL)
    conteo = {}
    for nombre_archivo in archivos:
        categoria = categorizar_archivo(nombre_archivo)
        conteo[categoria] = conteo.get(categoria, 0) + 1
    return conteo, len(archivos)


def _llenar_campo(tpl_body, etiqueta_texto, valor):
    """Escribe un valor justo después de una etiqueta con ':' en la plantilla
    (por ejemplo 'Cedula:', 'Medico Referente:', 'Fecha De Estudio:')."""
    if not valor:
        return
    for el in tpl_body:
        if el.tag == qn('w:p'):
            texto_parrafo = ''.join(el.itertext())
            if etiqueta_texto in texto_parrafo:
                runs = el.findall(qn('w:r'))
                run_colon = None
                for r in runs:
                    t = r.find(qn('w:t'))
                    if t is not None and t.text and ':' in t.text:
                        run_colon = r
                if run_colon is not None:
                    nuevo_run = copy.deepcopy(run_colon)
                    t_nuevo = nuevo_run.find(qn('w:t'))
                    t_nuevo.text = '  ' + str(valor)
                    t_nuevo.set(qn('xml:space'), 'preserve')
                    run_colon.addnext(nuevo_run)
                return


def unir_informe_renal(template_path, contenido_docx_file, fecha_estudio=None, medico_referente=None, cedula=None):
    """Une un documento Word del Dr. Quijada (hallazgos + imágenes) con la
    plantilla institucional de Gammagrama Renal. No usa IA: solo copia
    texto e imágenes, y detecta datos (nombre del paciente, M.S.A.S., C.M.)
    por patrones simples de texto."""
    tpl_doc = Document(template_path)
    content_doc = Document(contenido_docx_file)

    tpl_body = tpl_doc.element.body
    content_body = content_doc.element.body

    def encontrar_indice(body, condicion, desde=0):
        for i in range(desde, len(body)):
            el = body[i]
            if el.tag == qn('w:p') and condicion(''.join(el.itertext())):
                return i
        return None

    paciente_idx = encontrar_indice(content_body, lambda t: t.strip().startswith('Paciente:'))
    if paciente_idx is None:
        raise ValueError("No se encontró la línea 'Paciente:' en el documento del doctor.")
    atentamente_idx = encontrar_indice(content_body, lambda t: t.strip().startswith('Atentamente'), desde=paciente_idx + 1)
    if atentamente_idx is None:
        raise ValueError("No se encontró la línea 'Atentamente' (cierre/firma) en el documento del doctor.")

    insert_elements = [content_body[i] for i in range(paciente_idx + 1, atentamente_idx)]

    rid_map = {}
    for el in insert_elements:
        for blip in el.findall('.//' + qn('a:blip')):
            old_rid = blip.get(qn('r:embed'))
            if old_rid and old_rid not in rid_map:
                image_part = content_doc.part.related_parts[old_rid]
                new_rid, _ = tpl_doc.part.get_or_add_image(io.BytesIO(image_part.blob))
                rid_map[old_rid] = new_rid

    new_elements = []
    for el in insert_elements:
        new_el = copy.deepcopy(el)
        for blip in new_el.findall('.//' + qn('a:blip')):
            old_rid = blip.get(qn('r:embed'))
            if old_rid in rid_map:
                blip.set(qn('r:embed'), rid_map[old_rid])
        new_elements.append(new_el)

    title_idx = encontrar_indice(tpl_body, lambda t: 'ESTUDIO GAMMAGRAMA RENAL' in t)
    if title_idx is None:
        raise ValueError("No se encontró el título 'ESTUDIO GAMMAGRAMA RENAL' en la plantilla.")

    sig_idx = None
    for i in range(title_idx + 1, len(tpl_body)):
        el = tpl_body[i]
        if el.tag == qn('w:p') and el.findall('.//' + qn('w:drawing')):
            sig_idx = i
            break
    if sig_idx is None:
        raise ValueError("No se encontró la imagen de firma en la plantilla.")

    sig_el = tpl_body[sig_idx]

    for el in [tpl_body[i] for i in range(title_idx + 1, sig_idx)]:
        el.getparent().remove(el)

    for new_el in new_elements:
        sig_el.addprevious(new_el)

    nombre_paciente = None
    m_nombre = re.search(r'Paciente:\s*([^.:]+?)\.', ''.join(content_body[paciente_idx].itertext()))
    if m_nombre:
        nombre = m_nombre.group(1).strip()
        nombre_paciente = nombre
        for el in tpl_body:
            if el.tag == qn('w:p') and 'Paciente:' in ''.join(el.itertext()):
                for r in el.findall(qn('w:r')):
                    t = r.find(qn('w:t'))
                    if t is not None and t.text and t.text.strip() == '' and len(t.text) > 5:
                        t.text = ('  ' + nombre).ljust(len(t.text))
                        break
                break

    texto_completo = '\n'.join(''.join(el.itertext()) for el in content_body if el.tag == qn('w:p'))
    credenciales = []
    for patron in [r'M\.S\.A\.S\.?\s*\d+', r'C\.M\.?\s*\d+']:
        m = re.search(patron, texto_completo)
        if m and m.group(0) not in credenciales:
            credenciales.append(m.group(0))

    credential_style_p = None
    for el in tpl_body:
        if el.tag == qn('w:p') and 'Radioterapeuta' in ''.join(el.itertext()):
            credential_style_p = el

    if credential_style_p is not None:
        for cred in credenciales:
            nuevo_p = copy.deepcopy(credential_style_p)
            runs = nuevo_p.findall(qn('w:r'))
            for extra in runs[1:]:
                nuevo_p.remove(extra)
            t_el = runs[0].find(qn('w:t'))
            t_el.text = cred
            t_el.set(qn('xml:space'), 'preserve')
            credential_style_p.addnext(nuevo_p)
            credential_style_p = nuevo_p

    if fecha_estudio:
        _llenar_campo(tpl_body, 'Fecha De Estudio', fecha_estudio)
    if medico_referente:
        _llenar_campo(tpl_body, 'Medico Referente', medico_referente)
    if cedula:
        _llenar_campo(tpl_body, 'Cedula', cedula)

    buffer = io.BytesIO()
    tpl_doc.save(buffer)
    buffer.seek(0)
    return buffer, nombre_paciente


st.set_page_config(
    page_title="Medicina Nuclear - Sistema de Transcripción",
    page_icon="🏥",
    layout="wide"
)

if os.path.exists("logo.png"):
    with open("logo.png", "rb") as image_file:
        logo_base64 = base64.b64encode(image_file.read()).decode()

    st.markdown(f"""
        <style>
        .block-container {{ padding-top: 3.5rem !important; }}
        </style>
        <div style="
            background-color: #ebe20e; 
            width: 100%; 
            padding: 15px 20px; 
            margin-bottom: 25px;
            box-shadow: 0px 4px 6px rgba(0,0,0,0.08);
            display: flex;
            align-items: center;
            border-radius: 5px;
        ">
            <img src="data:image/png;base64,{logo_base64}" width="150">
        </div>
    """, unsafe_allow_html=True)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; }

    .main-title { font-size: 26px; font-weight: 700; color: #F1F5F9; letter-spacing: -0.3px; }
    .sub-title { font-size: 14px; color: #94A3B8; margin-bottom: 28px; }

    /* Encabezados de sección (st.header / st.subheader) */
    h2 {
        color: #F1F5F9 !important;
        font-weight: 700 !important;
        font-size: 20px !important;
        border-bottom: 2px solid #334155;
        padding-bottom: 10px;
        margin-top: 10px !important;
    }
    h3 {
        color: #CBD5E1 !important;
        font-weight: 600 !important;
        font-size: 16px !important;
    }

    /* Botones */
    .stButton button {
        background-color: #C9A227;
        color: #0F172A;
        border-radius: 8px;
        font-weight: 700;
        width: 100%;
        border: none;
        padding: 0.55rem 1rem;
        transition: background-color 0.15s ease;
        letter-spacing: 0.2px;
    }
    .stButton button:hover { background-color: #E0B92E; color: #0F172A; }

    /* Contenedores con borde (cajas de resultados, tarjetas del historial) */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 10px !important;
        border: 1px solid #334155 !important;
        background-color: #1E293B !important;
    }

    /* Métricas (contador de estudios) */
    div[data-testid="stMetric"] {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 14px 10px;
    }
    div[data-testid="stMetricLabel"] { font-size: 11px !important; color: #94A3B8 !important; }
    div[data-testid="stMetricValue"] { color: #F1F5F9 !important; font-weight: 700 !important; }

    /* Menús desplegables (Historial) */
    div[data-testid="stExpander"] {
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
        background-color: #1E293B !important;
    }

    /* Cuadros de texto y áreas de texto */
    .stTextInput input, .stTextArea textarea {
        border-radius: 8px !important;
        border: 1px solid #334155 !important;
        background-color: #1E293B !important;
        color: #F1F5F9 !important;
    }

    /* Zona de arrastrar archivos */
    [data-testid="stFileUploaderDropzone"] {
        border-radius: 10px !important;
        border: 1.5px dashed #C9A227 !important;
        background-color: #1E293B !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title"> Unidad de Medicina Nuclear</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Panel de Transcripción y Estructuración de Informes Médicos</div>', unsafe_allow_html=True)

st.sidebar.header("🔑 Claves de API")
openai_key = st.sidebar.text_input("OpenAI API Key", type="password", value=os.getenv("OPENAI_API_KEY", ""))
gemini_key = st.sidebar.text_input("Gemini API Key", type="password", value=os.getenv("GEMINI_API_KEY", ""))

tipo_estudio = st.sidebar.selectbox("Selecciona el estudio:", [
    "Gammagrafía Ósea", 
    "Gammagrafía Tiroidea (Tc-99m)", 
    "Gammagrafía Tiroidea (Iodo-131)", 
    "Rastreo Corporal Total", 
    "Gammagrafía Renal (DTPA/DMSA)", 
    "Plantilla Libre"
])

PLANTILLAS = {
    "Gammagrafía Ósea": """I. DATOS TÉCNICOS

- Estudio: Rastreo corporal óseo
- Radiofármaco: MDP (Metilendifosfonato) marcado con Tc-99m (Tecnecio)
- Actividad administrada: 20 mCi
- Vía de administración: Endovenosa
- Equipo: Gammacámara Elscint Apex 409 AG
- Proyecciones obtenidas: Anterior y posterior de cuerpo entero
- Calidad técnica del estudio: Sin obstrucciones


II. ANTECEDENTES CLÍNICOS
Historia clínica, gammagramas previos, motivo de estudio


III. HALLAZGOS
Se realizó gammagrafía ósea total, observándose la distribución del radiofármaco en las siguientes regiones:

Vista anterior (proyección ANT)
- Cráneo y macizo facial:
- Escápulas y esternón:
- Parrilla costal anterior:
- Miembros superiores:
- Pelvis:
- Miembros inferiores:

Vista posterior (proyección POST)
- Columna vertebral:
- Pelvis posterior:
- Parrilla costal posterior y escápulas:
- Cráneo posterior:
- Sistema