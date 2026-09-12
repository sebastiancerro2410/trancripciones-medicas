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
- Sistema renal y tejidos blandos:


IV. CONCLUSIÓN DIAGNÓSTICO
""",
    "Gammagrafía Tiroidea (Tc-99m)": """Paciente: [Nombre del Paciente]
Edad: [Edad] Años.
Fecha del estudio: [Fecha].
Médico remitente: [Médico]

I. DATOS TÉCNICOS

- Estudio: Gammagrafía tiroidea
- Radiofármaco: Pertecnetato de sodio marcado con Tc-99m
- Vía de administración: Endovenosa
- Proyecciones obtenidas: Anterior y oblicuas de la región cervical

II. ANTECEDENTES CLÍNICOS

[Motivo de la evaluación / Antecedentes]

III. HALLAZGOS

Se visualiza la glándula tiroidea en su localización anatómica habitual a nivel de la región cervical.

[Descripción de la glándula, tamaño, morfología, distribución y nódulos]

IV. IMPRESIÓN DIAGNÓSTICA

[Impresión diagnóstica final]""",
    "Gammagrafía Tiroidea (Iodo-131)": """REALIZAMOS UN GAMMAGRAMA TIROIDEO UTILIZANDO UNA GAMMACAMARA MARCA ELSCINT APEX 409 AG Y PREVIA ADMINISTRACION POR VIA ORAL CON IODO 131. EL COLIMADOR PARA REALIZAR EL ESTUDIO ES UN PINHOLE.

SE APRECIA GLÁNDULA TIROIDE A NIVEL DE CUELLO, [Descripción de tamaño, distribución heterogénea/homogénea, zonas iso/hipercaptantes, lóbulos]. NO SE APRECIAN IMAGENES DE L.O.E., en tal caso que no haya

PRUEBA DE CAPTACION (V-N  .15-45%)…………………………….. RESULTADO:

CONCLUSIÓN:

[Conclusión diagnóstica, ej. Tiroides aumentada de tamaño, Bocio multinodular, etc.]""",
    "Rastreo Corporal Total": "INFORMACIÓN DEL ESTUDIO: RASTREO CORPORAL TOTAL\n[DATOS DEL PACIENTE]\nINDICACIÓN:\n\nHALLAZGOS:\n- Áreas de hipercaptación fisiológica y patológica:\n\nCONCLUSIÓN:",
    "Gammagrafía Renal (DTPA/DMSA)": "INFORMACIÓN DEL ESTUDIO: GAMMAGRAFÍA RENAL\n[DATOS DEL PACIENTE]\nINDICACIÓN:\n\nHALLAZGOS:\n- Perfusión y función renal izquierda:\n- Perfusión y función renal derecha:\n- Excreción / Función relativa:\n\nCONCLUSIÓN:",
    "Plantilla Libre": "Insertar informe médico estructurado."
}

INSTRUCCIONES_ESTRICTAS_OSEA = """
REGLAS ADICIONALES OBLIGATORIAS PARA ESTE INFORME (Gammagrafía Ósea):
1. Respeta EXACTAMENTE la estructura de la plantilla: I. DATOS TÉCNICOS, II. ANTECEDENTES CLÍNICOS, III. HALLAZGOS (con sus dos subsecciones "Vista anterior (proyección ANT)" y "Vista posterior (proyección POST)", cada una con sus mismos ítems en el mismo orden), y IV. CONCLUSIÓN DIAGNÓSTICO.
2. NO fusiones, elimines, renombres ni reordenes ninguna subsección o ítem de "Vista anterior" ni de "Vista posterior", aunque el dictado los mencione en otro orden o de forma mezclada. Coloca cada hallazgo dictado en el ítem anatómico que le corresponda.
3. En "I. DATOS TÉCNICOS", los valores ya vienen fijos en la plantilla (radiofármaco, actividad, vía, equipo, proyecciones, calidad técnica). Mantenlos tal cual salvo que el dictado mencione explícitamente un valor distinto para ese campo puntual; en ese caso, usa el valor del dictado solo para ese campo.
4. Si un ítem de "Vista anterior" o "Vista posterior" no fue mencionado en el dictado, escribe exactamente "Dentro de límites normales" en ese ítem. No lo dejes vacío y no inventes hallazgos.
5. No agregues secciones, encabezados ni comentarios que no estén en la plantilla original.
"""

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Cargar Dictado en Audio")
    st.caption("Puedes subir más de un audio si el médico envió varios para el mismo paciente (se combinan antes de generar el informe).")
    audio_files = st.file_uploader(
        "Arrastra el audio o los audios (.ogg, .mp3, .opus, .wav, .m4a)",
        type=["ogg", "mp3", "opus", "wav", "m4a"],
        accept_multiple_files=True
    )
    if audio_files:
        for i, af in enumerate(audio_files, start=1):
            st.write(f"Audio {i}: {af.name}")
            st.audio(af)
            if af.size < 20_000:
                st.warning(f"⚠️ El audio {i} ({af.name}) parece muy corto o vacío. Puede que la transcripción salga incompleta.")
    
    altura_texto = 320 if tipo_estudio == "Gammagrafía Ósea" else (360 if "Tc-99m" in tipo_estudio else 220)
    plantilla_actual = st.text_area("Plantilla a completar:", value=PLANTILLAS[tipo_estudio], height=altura_texto)

    verificacion_activada = st.checkbox(
        "🔍 Verificar formato automáticamente al generar (opcional, gratis, no usa IA)",
        value=False,
        help="Revisa campos sin completar, secciones faltantes y texto demasiado corto. No reemplaza la revisión clínica del transcriptor."
    )

with col2:
    caja_resultados = st.container(border=True)
    with caja_resultados:
        st.subheader("2. Informe Final Generado")
        if not audio_files: st.info("Sube uno o más audios a la izquierda para comenzar.")

if audio_files and st.button("🚀 Procesar e Generar Informe"):
    if not openai_key or not gemini_key:
        st.error("⚠️ Ingrese ambas API Keys en la barra lateral.")
    else:
        barra_progreso = st.progress(0, text="Iniciando...")

        transcripciones = []
        client_openai = openai.OpenAI(api_key=openai_key)
        total_audios = len(audio_files)

        for i, af in enumerate(audio_files, start=1):
            porcentaje = int((i - 1) / total_audios * 60)
            barra_progreso.progress(porcentaje, text=f"🎧 Transcribiendo audio {i} de {total_audios}...")

            ext = af.name.split('.')[-1].lower()
            if ext == 'opus': ext = 'ogg'
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
                tmp.write(af.read())
                tmp_path = tmp.name
            try:
                with open(tmp_path, "rb") as f:
                    transcript = client_openai.audio.transcriptions.create(model="whisper-1", file=f, language="es")
                transcripciones.append(transcript.text)
            except Exception as e:
                st.error(f"Error al transcribir el audio {i} ({af.name}): {e}")
            finally:
                os.remove(tmp_path)

        barra_progreso.progress(60, text="🎧 Transcripción completa.")

        if transcripciones:
            texto_transcrito = "\n\n".join(
                f"[Audio {i}]\n{texto}" for i, texto in enumerate(transcripciones, start=1)
            )
            with st.expander("Ver texto crudo (todos los audios)"): st.write(texto_transcrito)

            try:
                client_gemini = genai.Client(api_key=gemini_key)
                instrucciones_extra = INSTRUCCIONES_ESTRICTAS_OSEA if tipo_estudio == "Gammagrafía Ósea" else ""
                nota_multi_audio = (
                    "\nNOTA: El dictado puede venir dividido en varios audios (marcados como [Audio 1], [Audio 2], etc.) "
                    "del mismo paciente. Combina la información de todos ellos en un solo informe coherente, "
                    "sin repetir datos duplicados ni mencionar que venían separados en audios distintos.\n"
                    if len(transcripciones) > 1 else ""
                )
                prompt = (
                    f"Eres un experto en medicina nuclear. Rellena la plantilla con el dictado. "
                    f"IMPORTANTE: no inventes ni infieras ningún dato que no esté explícitamente mencionado en el dictado.\n"
                    f"REGLA ESPECÍFICA PARA RESULTADOS: Si el doctor NO menciona el resultado de la 'PRUEBA DE CAPTACION' o cualquier campo marcado como 'RESULTADO:', déjalo strictly EN BLANCO (es decir, deja 'RESULTADO:' sin añadir nada a continuación). No escribas 'Dentro de límites normales', ni inventes valores o porcentajes.\n"
                    f"Para los demás campos anatómicos o descriptivos no mencionados en el dictado, si la plantilla requiere completarlos, escribe exactamente 'Dentro de límites normales'. Nunca completes con información supuesta. Usa negritas (Markdown **) para los títulos."
                    f"{nota_multi_audio}"
                    f"{instrucciones_extra}\n"
                    f"DICTADO: {texto_transcrito}\n"
                    f"PLANTILLA:\n{plantilla_actual}"
                )

                MAX_INTENTOS = 4
                response = None
                ultimo_error = None
                for intento in range(MAX_INTENTOS):
                    porcentaje = 60 + int((intento / MAX_INTENTOS) * 35)
                    if intento == 0:
                        barra_progreso.progress(porcentaje, text="🤖 Organizando informe con Gemini...")
                    else:
                        barra_progreso.progress(porcentaje, text=f"⏳ Servidor ocupado, reintentando ({intento + 1}/{MAX_INTENTOS})...")
                    try:
                        response = client_gemini.models.generate_content(
                            model="gemini-3.6-flash",
                            contents=prompt
                        )
                        break
                    except Exception as err_intento:
                        ultimo_error = err_intento
                        if "503" in str(err_intento) or "UNAVAILABLE" in str(err_intento):
                            if intento < MAX_INTENTOS - 1:
                                time.sleep(5 * (intento + 1))
                        else:
                            raise

                if response is None:
                    barra_progreso.progress(100, text="❌ No se pudo completar.")
                    raise ultimo_error

                barra_progreso.progress(95, text="💾 Guardando informe...")

                timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                nombre_limpio = tipo_estudio.replace(' ', '_').replace('(', '').replace(')', '').replace('-', '_')
                nombre_hist = f"{timestamp}_{nombre_limpio}.txt"
                guardar_en_historial(nombre_hist, response.text)

                barra_progreso.progress(100, text="✅ ¡Listo!")

                with caja_resultados:
                    st.success("✨ ¡Informe listo! (se guardó una copia en el historial)")
                    st.markdown(response.text)
                    st.markdown("---")
                    st.text_area("Copiar para sistema:", value=response.text, height=200)
                    st.download_button("📥 Descargar Informe (.txt)", data=response.text, file_name="informe_medico.txt", mime="text/plain")

                    if verificacion_activada:
                        st.markdown("---")
                        advertencias = verificar_informe_basico(response.text, tipo_estudio)
                        if advertencias:
                            st.warning("🔍 Verificación automática — revisa esto antes de enviarlo:")
                            for adv in advertencias:
                                st.markdown(f"- {adv}")
                        else:
                            st.success("🔍 Verificación automática: no se detectaron problemas de formato o estructura. (No reemplaza la revisión clínica.)")
            except Exception as e:
                barra_progreso.progress(100, text="❌ Error.")
                st.error(f"Error al estructurar: {e}")

st.markdown("---")
st.header("📎 Unir Informe Renal (Dr. Quijada) con Plantilla")
st.markdown("Sube el documento Word que envía el Dr. Quijada (con sus hallazgos e imágenes) y la app lo une automáticamente con la plantilla institucional. No usa IA, por lo que no necesita las claves de API.")

doc_quijada = st.file_uploader("Documento del Dr. Quijada (.docx)", type=["docx"], key="doc_quijada")

if doc_quijada:
    try:
        _doc_check = Document(doc_quijada)
        doc_quijada.seek(0)
        _texto_check = '\n'.join(''.join(p.itertext()) for p in _doc_check.element.body if p.tag == qn('w:p'))
        if 'Paciente:' not in _texto_check:
            st.warning("⚠️ No se encontró la línea 'Paciente:' en este documento. Revísalo antes de continuar, la unión podría fallar o quedar incompleta.")
        if 'Atentamente' not in _texto_check:
            st.warning("⚠️ No se encontró la palabra 'Atentamente' (cierre de firma) en este documento. Revísalo antes de continuar, la unión podría fallar o quedar incompleta.")
    except Exception:
        st.warning("⚠️ No se pudo leer este archivo como un documento Word válido.")

st.markdown("**Datos adicionales (opcional, solo si faltan en el documento):**")
col_fecha, col_medico, col_cedula = st.columns(3)
with col_fecha:
    fecha_estudio = st.date_input("Fecha de Estudio", value=None, format="DD/MM/YYYY")
with col_medico:
    medico_referente = st.text_input("Médico Referente")
with col_cedula:
    cedula_paciente = st.text_input("Cédula del Paciente")

nombre_para_guardar = st.text_input("Nombre para guardar el archivo (opcional, si lo dejas vacío se usa el nombre del paciente detectado en el documento):")

verificacion_renal_activada = st.checkbox(
    "🔍 Verificar automáticamente al unir (opcional, gratis, no usa IA)",
    value=False,
    help="Revisa que se haya detectado el nombre del paciente, que el documento no quede vacío, y que se hayan copiado imágenes. No reemplaza la revisión clínica."
)

if doc_quijada and st.button("🔗 Unir y Guardar en Historial", type="primary"):
    try:
        fecha_texto = fecha_estudio.strftime("%d/%m/%Y") if fecha_estudio else None
        merged_bytes, nombre_detectado = unir_informe_renal(
            "plantilla_renal.docx", doc_quijada,
            fecha_estudio=fecha_texto,
            medico_referente=medico_referente,
            cedula=cedula_paciente
        )
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nombre_elegido = nombre_para_guardar.strip() if nombre_para_guardar.strip() else nombre_detectado
        nombre_base = nombre_elegido.replace(' ', '_') if nombre_elegido else "SinNombre"
        nombre_hist = f"{timestamp}_Renal_{nombre_base}.docx"
        guardar_en_historial(nombre_hist, merged_bytes.getvalue())
        st.success(f"✅ Guardado como '{nombre_hist}'. Lo encuentras más abajo, en Historial de Informes.")

        if verificacion_renal_activada:
            advertencias_renal = verificar_informe_renal_basico(merged_bytes.getvalue(), nombre_detectado)
            if advertencias_renal:
                st.warning("🔍 Verificación automática — revisa esto antes de enviarlo:")
                for adv in advertencias_renal:
                    st.markdown(f"- {adv}")
            else:
                st.success("🔍 Verificación automática: no se detectaron problemas. (No reemplaza la revisión clínica.)")
    except Exception as e:
        st.error(f"Error al unir los documentos: {e}")

st.markdown("---")
st.header("📊 Estudios Realizados")

conteo_tipos, total_estudios = contar_estudios_por_tipo()

if total_estudios > 0:
    tipos_ordenados = [
        "Gammagrafía Ósea",
        "Gammagrafía Tiroidea (Tc-99m)",
        "Gammagrafía Tiroidea (Iodo-131)",
        "Rastreo Corporal Total",
        "Gammagrafía Renal",
        "Plantilla Libre",
        "Otros",
    ]
    columnas_conteo = st.columns(len(tipos_ordenados) + 1)
    for col, tipo in zip(columnas_conteo, tipos_ordenados):
        with col:
            st.metric(tipo, conteo_tipos.get(tipo, 0))
    with columnas_conteo[-1]:
        st.metric("Total", total_estudios)
else:
    st.info("Todavía no hay estudios registrados.")

st.markdown("---")

archivos_historial_todos = os.listdir(CARPETA_HISTORIAL)

with st.expander(f"📁 Historial de Informes ({len(archivos_historial_todos)})", expanded=False):
    st.markdown("Todos los informes generados (de audio o unidos con el Dr. Quijada) quedan guardados aquí automáticamente. Puedes buscar, renombrar, borrar, o sacar el texto de cualquiera.")

    busqueda = st.text_input("🔍 Buscar por nombre de archivo o paciente")

    archivos_historial = sorted(archivos_historial_todos, reverse=True)
    if busqueda:
        archivos_historial = [a for a in archivos_historial if busqueda.lower() in a.lower()]

    st.caption(f"{len(archivos_historial)} informe(s) encontrados" if busqueda else f"{len(archivos_historial)} informe(s) en total")

    if archivos_historial:
        for nombre_archivo in archivos_historial:
            ruta = os.path.join(CARPETA_HISTORIAL, nombre_archivo)
            es_word = nombre_archivo.endswith(".docx")

            categoria = categorizar_archivo(nombre_archivo)
            emoji_cat, color_cat = ESTILO_CATEGORIA.get(categoria, ("📄", "#CBD5E1"))

            fecha_legible = ""
            partes = nombre_archivo.split("_", 2)
            if len(partes) >= 2:
                try:
                    fecha_legible = datetime.datetime.strptime(f"{partes[0]}_{partes[1]}", "%Y-%m-%d_%H-%M-%S").strftime("%d/%m/%Y · %H:%M")
                except ValueError:
                    pass

            with st.container(border=True):
                col_info, col_desc, col_texto, col_ren, col_del = st.columns([3, 1, 1, 1, 1])

                with col_info:
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
                            <span style="
                                background-color:#0F172A;
                                color:{color_cat};
                                border-left:3px solid {color_cat};
                                border-radius:3px;
                                padding:3px 10px;
                                font-size:11px;
                                font-weight:600;
                                letter-spacing:0.3px;
                                text-transform:uppercase;
                                white-space:nowrap;
                            ">{emoji_cat} {categoria}</span>
                        </div>
                        <div style="font-weight:600; font-size:15px; color:#F1F5F9;">{nombre_archivo}</div>
                        """,
                        unsafe_allow_html=True
                    )
                    if fecha_legible:
                        st.caption(f"🕐 {fecha_legible}")

                with col_desc:
                    with open(ruta, "rb") as f:
                        datos_archivo = f.read()
                    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if es_word else "text/plain"
                    st.download_button("⬇️", data=datos_archivo, file_name=nombre_archivo, mime=mime, key=f"desc_{nombre_archivo}", help="Descargar")

                with col_texto:
                    if st.button("📋", key=f"texto_btn_{nombre_archivo}", help="Ver texto sin imágenes"):
                        st.session_state["viendo_texto"] = nombre_archivo
                        st.session_state.pop("renombrando", None)

                with col_ren:
                    if st.button("✏️", key=f"ren_btn_{nombre_archivo}", help="Renombrar"):
                        st.session_state["renombrando"] = nombre_archivo
                        st.session_state.pop("viendo_texto", None)

                with col_del:
                    if st.button("🗑️", key=f"del_btn_{nombre_archivo}", help="Borrar"):
                        os.remove(ruta)
                        st.session_state.pop("renombrando", None)
                        st.session_state.pop("viendo_texto", None)
                        st.rerun()

                # --- Flujo para ver/copiar el texto sin imágenes ---
                if st.session_state.get("viendo_texto") == nombre_archivo:
                    if es_word:
                        with open(ruta, "rb") as f:
                            texto_solo = extraer_solo_texto(f.read())
                    else:
                        with open(ruta, "r", encoding="utf-8") as f:
                            texto_solo = f.read()
                    st.code(texto_solo, language=None, wrap_lines=True)
                    if st.button("Cerrar", key=f"cerrar_texto_{nombre_archivo}"):
                        st.session_state.pop("viendo_texto", None)
                        st.rerun()

                # --- Flujo de renombrado ---
                if st.session_state.get("renombrando") == nombre_archivo:
                    extension = os.path.splitext(nombre_archivo)[1]
                    nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
                    nuevo_nombre = st.text_input("Nuevo nombre (sin extensión):", value=nombre_sin_ext, key=f"input_ren_{nombre_archivo}")
                    col_ok, col_cancel = st.columns(2)
                    with col_ok:
                        if st.button("✅ Guardar nombre", key=f"ok_ren_{nombre_archivo}"):
                            nueva_ruta = os.path.join(CARPETA_HISTORIAL, nuevo_nombre.strip() + extension)
                            if nuevo_nombre.strip():
                                os.rename(ruta, nueva_ruta)
                            st.session_state.pop("renombrando", None)
                            st.rerun()
                    with col_cancel:
                        if st.button("❌ Cancelar", key=f"cancel_ren_{nombre_archivo}"):
                            st.session_state.pop("renombrando", None)
                            st.rerun()
    else:
        if busqueda:
            st.info("No se encontraron informes que coincidan con la búsqueda.")
        else:
            st.info("Todavía no hay informes generados.")