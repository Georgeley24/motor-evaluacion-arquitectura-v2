import streamlit as st
import pandas as pd
import datetime
import os
import re
import io
import math
import random
import unicodedata

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Circle, FancyArrowPatch

st.set_page_config(page_title="Motor de Evaluación Arquitectónica V2.0", layout="wide")

# ==========================================
# 🔒 CONTROL DE ACCESO MEDIANTE EXCEL
# ==========================================
EXCEL_ACCESOS = "accesos_alumnos.xlsx"

def cargar_padron_accesos():
    padron = {}
    if os.path.exists(EXCEL_ACCESOS):
        try:
            df_acc = pd.read_excel(EXCEL_ACCESOS, sheet_name="Datos")
            cols_norm = {str(c).strip().lower(): c for c in df_acc.columns}
            col_correo = cols_norm.get("correo_electronico")
            col_pass = None
            for candidato in ["contraseña", "contrasena", "password", "clave"]:
                if candidato in cols_norm:
                    col_pass = cols_norm[candidato]
                    break
            col_nombre = cols_norm.get("nombre_alumno")

            if col_correo and col_pass:
                for _, row in df_acc.iterrows():
                    correo_raw = str(row[col_correo]).strip().lower() if pd.notna(row[col_correo]) else ""
                    pass_raw = str(row[col_pass]).strip() if pd.notna(row[col_pass]) else ""
                    nombre_raw = str(row[col_nombre]).strip() if (col_nombre and pd.notna(row[col_nombre])) else ""
                    if correo_raw:
                        padron[correo_raw] = {
                            "contrasena": pass_raw,
                            "nombre": nombre_raw
                        }
            else:
                st.error(
                    f"❌ El archivo '{EXCEL_ACCESOS}' debe incluir las columnas 'correo_electronico' y 'contraseña'."
                )
        except Exception as e:
            st.error(f"Error al leer el archivo de accesos ({EXCEL_ACCESOS}): {e}")
    return padron

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("### 🏛️ Taller de proyectos VII (Inteligencia artificial aplicada a la arquitectura)")
    st.title("🔐 Acceso Restringido - Plataforma de Evaluación")
    st.markdown("Por favor, ingresa tu correo institucional registrado en el padrón del taller y tu contraseña asignada.")
    
    padron_autorizado = cargar_padron_accesos()
    
    if not padron_autorizado:
        st.warning(f"⚠️ Atención Docente: No se encontró el archivo `{EXCEL_ACCESOS}` o no tiene registros válidos en la hoja `Datos`. Verifica que esté en la ruta correspondiente.")

    with st.form("form_login"):
        correo_user = st.text_input("Correo Electrónico").strip().lower()
        clave_user = st.text_input("Contraseña", type="password").strip()
        btn_login = st.form_submit_button("Ingresar al Motor V2.0")
        
        if btn_login:
            if not padron_autorizado:
                st.error("🚫 El sistema no tiene registros de alumnos cargados en el Excel de accesos.")
            elif correo_user in padron_autorizado and clave_user == padron_autorizado[correo_user]["contrasena"]:
                st.session_state.autenticado = True
                st.session_state.usuario_correo = correo_user
                st.session_state.usuario_nombre = padron_autorizado[correo_user]["nombre"]
                st.success("✅ Acceso concedido. Cargando Motor V2.0...")
                st.rerun()
            else:
                st.error("🚫 Acceso denegado: El correo no se encuentra en el padrón autorizado o la contraseña es incorrecta.")
    st.stop()

# ==========================================
# 🏛️ APLICACIÓN PRINCIPAL - MOTOR V2.0
# ==========================================

# Estado persistente escalonado para entregables y resultados
if "entregables_guardados" not in st.session_state:
    st.session_state.entregables_guardados = {}
if "pdfs_normativa_guardados" not in st.session_state:
    st.session_state.pdfs_normativa_guardados = []
if "evaluacion_activa" not in st.session_state:
    st.session_state.evaluacion_activa = False
if "resultado_evaluacion" not in st.session_state:
    st.session_state.resultado_evaluacion = None

col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown("### Taller de proyectos VII (Inteligencia artificial aplicada a la arquitectura)")
with col_head2:
    etiqueta_usuario = st.session_state.get("usuario_nombre") or st.session_state.usuario_correo
    st.caption(f"👤 {etiqueta_usuario}")
    if st.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

st.title("🏛️ Motor de Evaluación Arquitectónica con IA - Versión 2.0")
st.markdown("Plataforma de auditoría paramétrica, normativa y de rigor académico.")
st.markdown("*Creado por: **MGP-Arq. Jorge Alejandro López Kúk (Georgeley)***")
st.markdown("---")

# Sidebar para Datos Generales y Carga Escalonada de Archivos
st.sidebar.header("📋 Datos Generales del Proyecto")
nombre_proyecto = st.sidebar.text_input("Nombre del Proyecto", "Ej. Centro Comunitario Kúk", key="proj_nombre")
tipologia_proyecto = st.sidebar.text_input("Tipología", "Ej. Equipamiento Educativo / Cultural", key="proj_tipologia")
superficie_terreno_input = st.sidebar.text_input("Superficie del Terreno (m²)", "450.0", key="proj_sup_terreno")

st.sidebar.markdown("---")
st.sidebar.markdown("### 👥 Integrantes del Equipo (Hasta 7)")
integrantes = []
for i in range(1, 8):
    integrante = st.sidebar.text_input(f"Integrante {i}", "", key=f"integrante_{i}")
    if integrante.strip():
        integrantes.append(integrante.strip())

fecha_evaluacion = datetime.date.today()
st.sidebar.markdown(f"**Fecha de evaluación:** {fecha_evaluacion.strftime('%d/%m/%Y')}")

st.sidebar.markdown("---")
st.sidebar.header("📂 Carga Escalonada de Entregables (V2.0)")
st.sidebar.caption("Puedes cargar por etapas y evaluar sin borrar lo ya cargado anteriormente.")

deliverables_config = [
    ("01_normativa.xlsx", "01. Marco Normativo", 10),
    ("02_terreno_cos_cus.xlsx", "02. COS, CUS y Terreno", 15),
    ("03_analisis_usuario.xlsx", "03. Análisis de Usuario", 5),
    ("04_programa_arquitectonico.xlsx", "04. Programa Arquitectónico", 20),
    ("05_estacionamiento.xlsx", "05. Estacionamiento", 5),
    ("06_matriz_funcionamientoV2.xlsx", "06. Matriz de Funcionamiento", 10),
    ("07_concepto_volumetria.xlsx", "07. Concepto y Volumetría", 10),
    ("08_caracteristicas_tecnicas.xlsx", "08. Características Técnicas", 5),
    ("09_catalogo_especificaciones.xlsx", "09. Catálogo de Especificaciones", 5),
    ("10_verificacion_propuesta.xlsx", "10. Verificación de Propuesta", 15)
]

# 01. Marco Normativo (Persistencia escalonada)
st.sidebar.subheader("01. Marco Normativo")
file_normativa_xlsx = st.sidebar.file_uploader("Matriz Excel Normativa (01_normativa.xlsx)", type=["xlsx"], key="normativa_xlsx")
file_normativas_pdf = st.sidebar.file_uploader(
    "PDFs de Normativa (Match exacto con Col. B)", 
    type=["pdf"], 
    accept_multiple_files=True, 
    key="normativa_pdf_multi"
)

if file_normativas_pdf:
    st.session_state.pdfs_normativa_guardados = [
        {"name": pdf.name, "bytes": pdf.getvalue()} for pdf in file_normativas_pdf
    ]

if st.session_state.pdfs_normativa_guardados:
    st.sidebar.success(f"📎 {len(st.session_state.pdfs_normativa_guardados)} PDF(s) normativo(s) guardado(s) en sesión.")

if file_normativa_xlsx is not None:
    if file_normativa_xlsx.name != "01_normativa.xlsx":
        st.sidebar.error("❌ Nombre de archivo incorrecto. Debe llamarse exactamente '01_normativa.xlsx'. Corríjalo para continuar.")
    else:
        try:
            df_01_leido = pd.read_excel(io.BytesIO(file_normativa_xlsx.getvalue()), sheet_name="Datos")
            st.session_state.entregables_guardados["01_normativa.xlsx"] = df_01_leido
        except Exception as e:
            st.sidebar.error(f"Error leyendo 01_normativa.xlsx: {e}")

if "01_normativa.xlsx" in st.session_state.entregables_guardados:
    st.sidebar.caption("✅ `01_normativa.xlsx` listo y conservado en memoria.")

# Resto de entregables (02 al 10) con persistencia escalonada
for filename, label, weight in deliverables_config[1:]:
    st.sidebar.subheader(label)
    uploaded_file = st.sidebar.file_uploader(f"Matriz Excel ({filename})", type=["xlsx"], key=filename)
    if uploaded_file is not None:
        if uploaded_file.name != filename:
            st.sidebar.error(f"❌ Nombre incorrecto. Debe llamarse exactamente '{filename}'. Corríjalo para continuar.")
        else:
            try:
                file_bytes = uploaded_file.getvalue()
                if filename == "06_matriz_funcionamientoV2.xlsx":
                    xls_06 = pd.ExcelFile(io.BytesIO(file_bytes))
                    if "Espacios" not in xls_06.sheet_names or "Relaciones" not in xls_06.sheet_names:
                        st.sidebar.error("❌ El archivo '06_matriz_funcionamientoV2.xlsx' debe contener las hojas 'Espacios' y 'Relaciones'.")
                    else:
                        st.session_state.entregables_guardados[filename] = {
                            "Espacios": pd.read_excel(io.BytesIO(file_bytes), sheet_name="Espacios"),
                            "Relaciones": pd.read_excel(io.BytesIO(file_bytes), sheet_name="Relaciones"),
                        }
                else:
                    st.session_state.entregables_guardados[filename] = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Datos")
            except Exception as e:
                st.sidebar.error(f"Error leyendo {filename}: {e}")

    if filename in st.session_state.entregables_guardados:
        st.sidebar.caption(f"✅ `{filename}` conservado en memoria.")

with st.sidebar.expander("🗂️ Entregables guardados en esta sesión", expanded=False):
    cargados_keys = list(st.session_state.entregables_guardados.keys())
    if cargados_keys:
        for k_file in cargados_keys:
            st.write(f"• {k_file}")
    else:
        st.write("Sin archivos cargados aún.")

def extraer_numeros_texto(valor):
    if valor is None or pd.isna(valor):
        return []
    limpio = str(valor).replace(',', '.')
    return [float(m) for m in re.findall(r'-?\d+(?:\.\d+)?', limpio)]

def extraer_entero_estricto(valor):
    if valor is None or pd.isna(valor):
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if float(valor).is_integer():
            return int(valor)
        return None
    txt = str(valor).strip()
    if re.fullmatch(r'-?\d+', txt):
        return int(txt)
    if re.fullmatch(r'-?\d+\.0+', txt):
        return int(float(txt))
    return None

def normalizar_texto_clave(valor):
    t = "" if valor is None or pd.isna(valor) else str(valor).strip().lower()
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")

ZONAS_DIAGRAMA = {
    "publica":  ("Pública",     "#c9732a"),
    "semi":     ("Semipública", "#23847a"),
    "privada":  ("Privada",     "#4a5bc4"),
    "servicio": ("Servicio",    "#6f7680"),
    "exterior": ("Exterior",    "#4a9446"),
}

ALIAS_ZONA_DIAGRAMA = {
    "publica": "publica", "pública": "publica",
    "semipublica": "semi", "semipública": "semi", "semi": "semi",
    "privada": "privada",
    "servicio": "servicio", "servicios": "servicio",
    "exterior": "exterior",
    "administracion": "semi", "administración": "semi",
}

ALIAS_TIPO_DIAGRAMA = {
    "directa": "directa",
    "circulacion": "circ", "circulación": "circ",
    "visual": "visual",
    "indirecta": "indirecta",
}

def generar_diagrama_1920x1080(espacios_data, relaciones_data, titulo_diag="Diagrama de Funcionamiento", semilla=1):
    # Lienzo lógico 16:9 a 160 dpi -> 12.0 x 6.75 in = 1920 x 1080 px exactos
    W_log, H_log = 1200.0, 675.0
    rnd = random.Random(semilla)
    n_esp = len(espacios_data)

    espacios_work = [
        {
            "id": e.get("id", ""),
            "nombre": e.get("nombre", ""),
            "zona": e.get("zona", "semi"),
            "area": float(e.get("area", 10.0) or 10.0),
        }
        for e in espacios_data
    ]
    por_nombre_norm = {normalizar_texto_clave(e["nombre"]): e for e in espacios_work}

    relaciones_work = []
    for r in relaciones_data:
        a = por_nombre_norm.get(normalizar_texto_clave(r.get("origen", "")))
        b = por_nombre_norm.get(normalizar_texto_clave(r.get("destino", "")))
        tipo_raw = normalizar_texto_clave(r.get("tipo", ""))
        tipo_code = ALIAS_TIPO_DIAGRAMA.get(tipo_raw, "directa")
        if a and b and a is not b:
            relaciones_work.append((a, b, tipo_code))

    def radio_nodo(esp):
        base_r = 22.0 + math.sqrt(max(esp["area"], 1.0)) * 6.0
        if n_esp > 18:
            base_r *= 0.80
        elif n_esp > 12:
            base_r *= 0.90
        return max(22.0, min(62.0, base_r))

    def objetivo_zona(z):
        x = {"publica": 0.16, "semi": 0.40, "privada": 0.78, "servicio": 0.56, "exterior": 0.12}.get(z, 0.45) * W_log
        y = {"servicio": 0.78, "exterior": 0.80}.get(z, 0.46) * H_log
        return x, y

    for e in espacios_work:
        tx, ty = objetivo_zona(e["zona"])
        e["x"] = tx + (rnd.random() - 0.5) * 180.0
        e["y"] = ty + (rnd.random() - 0.5) * 180.0
        e["vx"] = e["vy"] = 0.0

    iteraciones = 400
    for it in range(iteraciones):
        k = 1.0 - it / iteraciones
        for e in espacios_work:
            tx, ty = objetivo_zona(e["zona"])
            e["fx"] = (tx - e["x"]) * 0.014
            e["fy"] = (ty - e["y"]) * 0.012
        for i, a in enumerate(espacios_work):
            for b in espacios_work[i + 1:]:
                dx, dy = a["x"] - b["x"], a["y"] - b["y"]
                d = math.hypot(dx, dy) or 1.0
                minimo = radio_nodo(a) + radio_nodo(b) + 24.0
                f = (minimo - d) * 0.55 if d < minimo else 3600.0 / (d * d)
                a["fx"] += dx / d * f; a["fy"] += dy / d * f
                b["fx"] -= dx / d * f; b["fy"] -= dy / d * f
        for a, b, _ in relaciones_work:
            dx, dy = b["x"] - a["x"], b["y"] - a["y"]
            d = math.hypot(dx, dy) or 1.0
            f = (d - (radio_nodo(a) + radio_nodo(b) + 28.0)) * 0.03
            a["fx"] += dx / d * f; a["fy"] += dy / d * f
            b["fx"] -= dx / d * f; b["fy"] -= dy / d * f
        for e in espacios_work:
            e["vx"] = (e["vx"] + e["fx"]) * 0.8
            e["vy"] = (e["vy"] + e["fy"]) * 0.8
            e["x"] += e["vx"] * k * 2.0
            e["y"] += e["vy"] * k * 2.0
            r = radio_nodo(e)
            e["x"] = max(r + 16.0, min(W_log - r - 16.0, e["x"]))
            e["y"] = max(r + 68.0, min(H_log - 76.0, e["y"]))

    fig = plt.figure(figsize=(12.0, 6.75), dpi=160, facecolor="white")
    ax = fig.add_axes([0.02, 0.11, 0.96, 0.79])
    ax.set_xlim(0, W_log)
    ax.set_ylim(H_log, 0)
    ax.set_aspect("equal")
    ax.axis("off")

    ink = "#1d2430"
    for a, b, tipo in relaciones_work:
        dx, dy = b["x"] - a["x"], b["y"] - a["y"]
        d = math.hypot(dx, dy) or 1.0
        ux, uy = dx / d, dy / d
        p1 = (a["x"] + ux * radio_nodo(a), a["y"] + uy * radio_nodo(a))
        p2 = (b["x"] - ux * (radio_nodo(b) + 3.0), b["y"] - uy * (radio_nodo(b) + 3.0))
        if tipo == "circ":
            ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=14,
                                         color=ink, lw=2.0, zorder=1, shrinkA=0, shrinkB=0))
        else:
            estilo = {
                "directa":   dict(lw=4.2, ls="-"),
                "visual":    dict(lw=2.6, ls=(0, (0.1, 2.5)), dash_capstyle="round"),
                "indirecta": dict(lw=2.0, ls=(0, (5, 3))),
            }.get(tipo, dict(lw=2.0, ls="-"))
            ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color=ink, zorder=1, **estilo)

    for e in espacios_work:
        r = radio_nodo(e)
        z_color = ZONAS_DIAGRAMA.get(e["zona"], ("Semipública", "#23847a"))[1]
        ax.add_patch(Circle((e["x"], e["y"]), r, color=z_color, zorder=2))
        palabras = str(e["nombre"]).split()
        lineas, actual = [], ""
        for p_txt in palabras:
            if actual and len(f"{actual} {p_txt}") > 11:
                lineas.append(actual)
                actual = p_txt
            else:
                actual = f"{actual} {p_txt}".strip()
        if actual:
            lineas.append(actual)
        y0 = e["y"] - (len(lineas) - 1) * 6.5
        for idx_l, l_txt in enumerate(lineas):
            ax.text(e["x"], y0 + idx_l * 13.0, l_txt, ha="center", va="center",
                    color="white", fontsize=8.5, fontweight="bold", zorder=3)

    if titulo_diag:
        fig.text(0.02, 0.93, titulo_diag, fontsize=15, fontweight="bold", color=ink, ha="left", va="center")

    usadas = {e["zona"] for e in espacios_work}
    zonas_leg = [Line2D([0], [0], marker="o", color="none", markerfacecolor=c, markersize=9, label=n)
                 for k, (n, c) in ZONAS_DIAGRAMA.items() if k in usadas]
    tipos_leg = [
        Line2D([0], [0], color=ink, lw=4.2, label="Relación directa"),
        Line2D([0], [0], color=ink, lw=2.0, marker=">", markevery=[1], label="Circulación"),
        Line2D([0], [0], color=ink, lw=2.0, ls=(0, (5, 3)), label="Relación indirecta"),
        Line2D([0], [0], color=ink, lw=2.6, ls=(0, (0.1, 2.5)), dash_capstyle="round", label="Relación visual"),
    ]
    fig.legend(handles=zonas_leg + tipos_leg, loc="lower center", bbox_to_anchor=(0.5, 0.015),
               ncol=5, frameon=False, fontsize=8.5)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

def escapar_pdf_texto(texto):
    limpio = "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")
    limpio = limpio.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return "".join(ch if 32 <= ord(ch) <= 126 else " " for ch in limpio)

def generar_pdf_reporte_bytes(res_data):
    """Genera un PDF válido en memoria sin alterar ningún dato de sesión ni carga."""
    lineas_pdf = []
    lineas_pdf.append("MOTOR DE EVALUACION ARQUITECTONICA V2.0 - EVIDENCIA DE EVALUACION")
    lineas_pdf.append("=" * 72)
    lineas_pdf.append(f"EVALUACION FINAL: {res_data['calificacion_final']:,.1f} / 100  |  {escapar_pdf_texto(res_data['tipo_evaluacion'])}")
    lineas_pdf.append(f"Fecha y hora de generacion del reporte: {res_data['fecha_hora_gen']}")
    lineas_pdf.append("-" * 72)
    lineas_pdf.append("DATOS GENERALES DEL PROYECTO")
    lineas_pdf.append(f"Proyecto: {escapar_pdf_texto(res_data['nombre_proyecto'])}")
    lineas_pdf.append(f"Tipologia: {escapar_pdf_texto(res_data['tipologia_proyecto'])}")
    lineas_pdf.append(f"Superficie del Terreno (m2): {escapar_pdf_texto(res_data['superficie_terreno_input'])}")
    lineas_pdf.append(f"Usuario / Alumno sesion: {escapar_pdf_texto(res_data['usuario_sesion'])}")
    integrantes_txt = ", ".join(res_data["integrantes"]) if res_data["integrantes"] else "No especificados"
    lineas_pdf.append(f"Integrantes del equipo: {escapar_pdf_texto(integrantes_txt)}")
    lineas_pdf.append("-" * 72)
    lineas_pdf.append("RESUMEN DE PONDERACION POR ENTREGABLE")
    for ent, peso, calif, est in res_data["detalles_puntuacion"]:
        lineas_pdf.append(f"- {escapar_pdf_texto(ent)} [Peso {peso}%]: {calif:,.1f} / 100 ({escapar_pdf_texto(est)})")

    secciones = [
        ("AUDITORIA 01 - MARCO NORMATIVO", res_data["score_normativa"], res_data["mensajes_normativa"]),
        ("AUDITORIA 02 - COS, CUS Y TERRENO", res_data["score_terreno"], res_data["mensajes_terreno"]),
        ("AUDITORIA 03 - ANALISIS DE USUARIO", res_data["score_usuario"], res_data["mensajes_usuario"]),
        ("AUDITORIA 04 - PROGRAMA ARQUITECTONICO", res_data["score_programa"], res_data["mensajes_programa"]),
        ("AUDITORIA 05 - ESTACIONAMIENTO", res_data["score_estacionamiento"], res_data["mensajes_estacionamiento"]),
        ("AUDITORIA 06 - MATRIZ DE FUNCIONAMIENTO", res_data["score_matriz"], res_data["mensajes_matriz"]),
        ("AUDITORIA 08 - CARACTERISTICAS TECNICAS", res_data["score_tecnicas"], res_data["mensajes_tecnicas"]),
    ]

    for titulo_sec, punt_sec, msgs_sec in secciones:
        lineas_pdf.append("-" * 72)
        lineas_pdf.append(f"{titulo_sec} -> Puntaje: {punt_sec:,.1f} / 100")
        for m in msgs_sec:
            txt_m = escapar_pdf_texto(m)
            while len(txt_m) > 92:
                lineas_pdf.append("  " + txt_m[:92])
                txt_m = txt_m[92:]
            lineas_pdf.append("  " + txt_m)

    # Paginación en páginas tamaño carta (612 x 792 pt)
    max_lineas_pag = 46
    paginas_lineas = [lineas_pdf[i:i + max_lineas_pag] for i in range(0, len(lineas_pdf), max_lineas_pag)]
    if not paginas_lineas:
        paginas_lineas = [["Reporte sin datos."]]

    objects = []
    # Objeto 1: Catalog
    objects.append(b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
    # Objeto 2: Pages (se completa con Kids)
    page_obj_ids = []
    content_obj_ids = []
    next_id = 3
    font_obj_id = None

    for _ in paginas_lineas:
        page_obj_ids.append(next_id)
        next_id += 1
        content_obj_ids.append(next_id)
        next_id += 1

    font_obj_id = next_id
    next_id += 1

    kids_str = " ".join(f"{pid} 0 R" for pid in page_obj_ids)
    pages_obj = f"2 0 obj\n<< /Type /Pages /Kids [ {kids_str} ] /Count {len(page_obj_ids)} >>\nendobj\n".encode("latin-1")
    objects.insert(1, pages_obj)

    for p_idx, p_lines in enumerate(paginas_lineas):
        pid = page_obj_ids[p_idx]
        cid = content_obj_ids[p_idx]
        stream_cmds = ["BT", "/F1 9.5 Tf", "45 752 Td", "14 TL"]
        for l_idx, l_txt in enumerate(p_lines):
            if l_idx == 0:
                stream_cmds.append(f"({l_txt}) Tj")
            else:
                stream_cmds.append(f"T* ({l_txt}) Tj")
        stream_cmds.append("ET")
        stream_bytes = ("\n".join(stream_cmds) + "\n").encode("latin-1")

        page_dict = (
            f"{pid} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            f"/Resources << /Font << /F1 {font_obj_id} 0 R >> >> "
            f"/Contents {cid} 0 R >>\nendobj\n"
        ).encode("latin-1")
        content_dict = (
            f"{cid} 0 obj\n<< /Length {len(stream_bytes)} >>\nstream\n".encode("latin-1")
            + stream_bytes
            + b"endstream\nendobj\n"
        )
        objects.append(page_dict)
        objects.append(content_dict)

    font_dict = f"{font_obj_id} 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n".encode("latin-1")
    objects.append(font_dict)

    pdf_out = io.BytesIO()
    pdf_out.write(b"%PDF-1.4\n")
    offsets = [0]
    for obj_b in objects:
        offsets.append(pdf_out.tell())
        pdf_out.write(obj_b)

    xref_pos = pdf_out.tell()
    pdf_out.write(f"xref\n0 {len(objects) + 1}\n".encode("latin-1"))
    pdf_out.write(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf_out.write(f"{off:010d} 00000 n \n".encode("latin-1"))

    pdf_out.write(
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode("latin-1")
    )
    return pdf_out.getvalue()

def ejecutar_motor_evaluacion():
    uploaded_files = st.session_state.entregables_guardados
    pdfs_guardados = st.session_state.pdfs_normativa_guardados

    missing_files = [d for d in deliverables_config if d[0] not in uploaded_files]
    is_complete = len(missing_files) == 0
    tipo_evaluacion = "Evaluación Final V2.0 🏆" if is_complete else "Evaluación Preliminar V2.0 🔍"

    # ==========================================
    # EVALUACIÓN 01: MARCO NORMATIVO
    # ==========================================
    score_normativa = 0.0
    mensajes_normativa = []
    condicionantes_guardadas = []

    if "01_normativa.xlsx" in uploaded_files:
        df_norm = uploaded_files["01_normativa.xlsx"]
        expected_cols_1 = ['id_norma', 'nombre_normativa', 'ambito', 'articulo', 'descripcion_norma', 'condicionante_diseno', 'parametro_evaluable']
        
        if list(df_norm.columns) != expected_cols_1:
            score_normativa = 0.0
            mensajes_normativa.append("❌ Las columnas del archivo Excel de Normativa no coinciden exactamente con la estructura base. Propuesta rechazada (0 pts).")
        elif df_norm.empty:
            score_normativa = 0.0
            mensajes_normativa.append("❌ El archivo Excel de Normativa está vacío (0 pts).")
        else:
            filas_validas_indices = []
            pdf_names = [p["name"].lower().replace('.pdf', '').strip() for p in pdfs_guardados] if pdfs_guardados else []
            
            for idx, row in df_norm.iterrows():
                art_val = str(row['articulo']).strip() if pd.notna(row['articulo']) else ""
                nombre_norm = str(row['nombre_normativa']).strip() if pd.notna(row['nombre_normativa']) else ""
                
                if not re.match(r'^ART_\d+$', art_val, re.IGNORECASE):
                    mensajes_normativa.append(f"⚠️ Fila {idx+1}: El artículo '{art_val}' no cumple con el formato requerido ('ART_' + número). Fila con puntaje 0.")
                    continue
                
                if pdf_names:
                    match_pdf = any(nombre_norm.lower() == pdf_n for pdf_n in pdf_names)
                    if not match_pdf:
                        mensajes_normativa.append(f"⚠️ Norma '{nombre_norm}' (Art: {art_val}): No coincide exactamente con ningún PDF cargado. Fila no revisada (0 pts).")
                        continue
                
                filas_validas_indices.append(idx)
                condicionantes_guardadas.append({
                    "norma": nombre_norm,
                    "articulo": art_val,
                    "condicionante": row['condicionante_diseno'],
                    "parametro": row['parametro_evaluable']
                })

            df_validas = df_norm.loc[filas_validas_indices]
            total_normas_unicas = df_validas['id_norma'].nunique() if not df_validas.empty else 0

            if total_normas_unicas >= 3:
                score_normativa += 40.0
                mensajes_normativa.append(f"✅ Se detectaron {total_normas_unicas} normas válidas (Cumple mínimo de 3): +40 pts base.")
            else:
                mensajes_normativa.append(f"❌ Solo se detectaron {total_normas_unicas} normas válidas (Se requieren al menos 3 para los 40 pts base): 0 pts base (conserva evaluación de rubros).")

            rubros_pesos = {
                'cos': (15.0, ['cos', 'coeficiente de ocupacion']),
                'cus': (15.0, ['cus', 'coeficiente de utilizacion']),
                'uso_suelo': (5.0, ['uso de suelo', 'compatibilidad', 'destino']),
                'alturas': (5.0, ['altura', 'alturas', 'nivel', 'niveles']),
                'area_verde': (5.0, ['area verde', 'jardin', 'vegetacion']),
                'area_permeable': (5.0, ['permeable', 'absorcion', 'cas'])
            }
            
            rubros_encontrados = set()
            for _, row in df_validas.iterrows():
                texto_analisis = f"{str(row['descripcion_norma']).lower()} {str(row['condicionante_diseno']).lower()} {str(row['parametro_evaluable']).lower()}"
                for rubro, (_, keywords) in rubros_pesos.items():
                    if any(kw in texto_analisis for kw in keywords):
                        rubros_encontrados.add(rubro)

            puntos_rubros = sum(rubros_pesos[r][0] for r in rubros_encontrados)
            score_normativa += puntos_rubros
            
            faltantes_rubros = [r for r in rubros_pesos if r not in rubros_encontrados]
            if faltantes_rubros:
                mensajes_normativa.append(f"⚠️ La propuesta no integró datos de: {faltantes_rubros}. Puntaje por rubros: +{puntos_rubros:.1f}/50 pts.")
            else:
                mensajes_normativa.append("✅ Se integraron los 6 rubros normativos clave (COS=15, CUS=15, uso_suelo=5, alturas=5, area_verde=5, area_permeable=5): +50.0/50 pts.")

            if total_normas_unicas >= 3 and total_normas_unicas > 5:
                score_normativa += 10.0
                mensajes_normativa.append(f"⭐ Se entregaron {total_normas_unicas} normas válidas (>5): +10 pts extra.")
            elif total_normas_unicas > 5 and total_normas_unicas < 3:
                mensajes_normativa.append("ℹ️ No aplica bono de +10 pts extra por tener menos de 3 normas válidas.")

            score_normativa = min(score_normativa, 100.0)
    else:
        score_normativa = 0.0
        mensajes_normativa.append("❌ Entregable 01 no cargado.")

    # ==========================================
    # EVALUACIÓN 02: COS, CUS Y TERRENO
    # ==========================================
    score_terreno = 0.0
    mensajes_terreno = []
    terreno_datos_globales = {}

    if "02_terreno_cos_cus.xlsx" in uploaded_files:
        df_ter = uploaded_files["02_terreno_cos_cus.xlsx"]
        expected_cols_2 = ['superficie_terreno_m2', 'cos_permitido_porcentaje', 'cus_permitido_coeficiente', 'restriccion_frontal_m', 'restriccion_posterior_m', 'restriccion_lateral_m']
        
        if list(df_ter.columns) != expected_cols_2:
            score_terreno = 0.0
            mensajes_terreno.append("❌ Las columnas del archivo Excel de Terreno (02) no coinciden exactamente con la estructura base. Propuesta rechazada (0 pts).")
        elif df_ter.empty:
            score_terreno = 0.0
            mensajes_terreno.append("❌ El archivo Excel de Terreno está vacío (0 pts). Corrige la carga.")
        else:
            row_ter = df_ter.iloc[0]
            sup_raw = row_ter['superficie_terreno_m2']
            cos_raw = row_ter['cos_permitido_porcentaje']
            cus_raw = row_ter['cus_permitido_coeficiente']

            if pd.isna(sup_raw) or pd.isna(cos_raw) or pd.isna(cus_raw):
                score_terreno = 0.0
                mensajes_terreno.append("❌ Carga inválida (0 pts): Faltan datos obligatorios de 'superficie_terreno_m2', 'cos_permitido_porcentaje' o 'cus_permitido_coeficiente'. Por favor corrige la carga.")
            else:
                try:
                    sup_propuesta = float(sup_raw)
                    cos_propuesto = float(cos_raw)
                    cus_propuesto = float(cus_raw)

                    if sup_propuesta <= 0:
                        score_terreno = 0.0
                        mensajes_terreno.append("❌ Carga inválida (0 pts): El dato de 'superficie_terreno_m2' debe ser mayor a 0. Por favor corrige la carga.")
                    else:
                        cos_nums_norm = []
                        cus_nums_norm = []
                        for item in condicionantes_guardadas:
                            param_txt = str(item.get('parametro', '')).strip().lower()
                            cond_val = item.get('condicionante', '')
                            nums_cond = extraer_numeros_texto(cond_val)
                            if re.search(r'\bcos\b', param_txt):
                                cos_nums_norm.extend(nums_cond)
                            if re.search(r'\bcus\b', param_txt):
                                cus_nums_norm.extend(nums_cond)

                        cos_match = any(abs(n - cos_propuesto) <= 0.01 for n in cos_nums_norm)
                        cus_match = any(abs(n - cus_propuesto) <= 0.01 for n in cus_nums_norm)

                        if not cos_match or not cus_match:
                            score_terreno = 0.0
                            mensajes_terreno.append(
                                "❌ Carga rechazada (0 pts): No hay información suficiente o coincidente en la Plantilla 01 para validar COS y/o CUS. "
                                "Por favor corrige tu carga asegurando que en la Plantilla 01 la columna 'parametro_evaluable' indique 'COS' o 'CUS' "
                                "y la columna 'condicionante_diseno' incluya el valor numérico correspondiente (ej. parametro_evaluable: 'COS' / condicionante_diseno: '70' o '0.70'; "
                                "parametro_evaluable: 'CUS' / condicionante_diseno: '2.1')."
                            )
                        else:
                            terreno_datos_globales = {
                                'superficie_terreno_m2': sup_propuesta,
                                'cos_permitido_porcentaje': cos_propuesto,
                                'cus_permitido_coeficiente': cus_propuesto
                            }
                            score_terreno += 70.0
                            mensajes_terreno.append("✅ Superficie de terreno válida y COS/CUS validados contra la Plantilla 01: +70 pts base.")

                            col_d = row_ter['restriccion_frontal_m']
                            col_e = row_ter['restriccion_posterior_m']
                            col_f = row_ter['restriccion_lateral_m']
                            
                            if pd.notna(col_d) and pd.notna(col_e) and pd.notna(col_f):
                                score_terreno += 30.0
                                mensajes_terreno.append("✅ Restricciones en colindancias (D, E, F) completas: +30 pts extra.")
                            else:
                                mensajes_terreno.append("⚠️ Faltan datos en restricciones de colindancias (0 pts extra en este rubro).")

                except Exception as ex:
                    score_terreno = 0.0
                    mensajes_terreno.append(f"❌ Error al procesar valores numéricos de Terreno ({ex}). Carga invalidada (0 pts); por favor corrige la carga.")
    else:
        score_terreno = 0.0
        mensajes_terreno.append("❌ Entregable 02 no cargado.")

    # ==========================================
    # EVALUACIÓN 03: ANÁLISIS DE USUARIO
    # ==========================================
    score_usuario = 0.0
    mensajes_usuario = []
    espacios_usuarios_guardados = []

    if "03_analisis_usuario.xlsx" in uploaded_files:
        df_usu = uploaded_files["03_analisis_usuario.xlsx"]
        expected_cols_3 = ['tipo_usuario', 'nivel_socioeconomico', 'caracteristicas_culturales', 'necesidad_espacial', 'cantidad_estimada']
        
        if list(df_usu.columns) != expected_cols_3:
            score_usuario = 0.0
            mensajes_usuario.append("❌ Las columnas de Análisis de Usuario no coinciden exactamente. Propuesta rechazada (0 pts).")
        elif df_usu.empty:
            score_usuario = 0.0
            mensajes_usuario.append("❌ El archivo de Análisis de Usuario está vacío (0 pts).")
        else:
            usuarios_validos = []
            culturas_vistas = set()
            
            for idx, row in df_usu.iterrows():
                u_raw = str(row['tipo_usuario']).strip() if pd.notna(row['tipo_usuario']) else ""
                u_norm = u_raw.lower()
                cult_raw = str(row['caracteristicas_culturales']).strip() if pd.notna(row['caracteristicas_culturales']) else ""
                nec_raw = str(row['necesidad_espacial']).strip() if pd.notna(row['necesidad_espacial']) else ""
                
                if not u_norm or len(u_norm) < 3 or not re.match(r'^[a-záéíóúüñ\s]+$', u_norm):
                    continue
                if not cult_raw or not nec_raw:
                    continue
                if u_norm in [u['usuario'] for u in usuarios_validos]:
                    continue
                if cult_raw.lower() in culturas_vistas:
                    continue
                    
                culturas_vistas.add(cult_raw.lower())
                usuarios_validos.append({
                    'usuario': u_norm,
                    'necesidad_espacial': nec_raw,
                    'cantidad_estimada': row['cantidad_estimada']
                })
                
                espacios_usuarios_guardados.append({
                    'necesidad_espacial': nec_raw,
                    'cantidad_estimada': row['cantidad_estimada']
                })

            num_usuarios = len(usuarios_validos)
            if num_usuarios == 0:
                score_tipologias = 0.0
            elif num_usuarios < 3:
                penalizacion = (3 - num_usuarios) * 10
                score_tipologias = max(0.0, 40.0 - penalizacion)
            else:
                score_tipologias = 40.0
                if 3 < num_usuarios <= 6:
                    score_tipologias += 20.0
                elif num_usuarios > 6:
                    score_tipologias += 30.0

            score_usuario += score_tipologias

            todas_necesidades_ok = all(pd.notna(row['necesidad_espacial']) and str(row['necesidad_espacial']).strip() != "" for _, row in df_usu.iterrows())
            todas_cantidades_ok = True
            for _, row in df_usu.iterrows():
                try:
                    if pd.isna(row['cantidad_estimada']) or float(row['cantidad_estimada']) < 0:
                        todas_cantidades_ok = False
                except:
                    todas_cantidades_ok = False

            if todas_necesidades_ok:
                score_usuario += 15.0
            if todas_cantidades_ok:
                score_usuario += 15.0

            score_usuario = min(score_usuario, 100.0)
            mensajes_usuario.append(f"✅ Análisis de Usuario evaluado con {num_usuarios} tipología(s) válida(s).")
    else:
        score_usuario = 0.0
        mensajes_usuario.append("❌ Entregable 03 no cargado.")

    # ==========================================
    # EVALUACIÓN 04: PROGRAMA ARQUITECTÓNICO
    # ==========================================
    score_programa = 0.0
    mensajes_programa = []
    programa_04_registros = []
    programa_04_areas_por_nombre = {}

    if "04_programa_arquitectonico.xlsx" in uploaded_files:
        df_prog = uploaded_files["04_programa_arquitectonico.xlsx"]
        expected_cols_4 = ['id_espacio', 'nombre_espacio', 'zona', 'nivel', 'cantidad', 'area_unitaria_m2', 'area_total_m2', 'relacion_directa_ids', 'relacion_indirecta_ids', 'ancho_minimo_m', 'alto_minimo_m']
        
        if list(df_prog.columns) != expected_cols_4:
            score_programa = 0.0
            mensajes_programa.append("❌ Las columnas del Programa Arquitectónico no coinciden exactamente con la estructura base. Propuesta rechazada (0 pts).")
        elif df_prog.empty:
            score_programa = 0.0
            mensajes_programa.append("❌ El archivo de Programa Arquitectónico está vacío (0 pts).")
        else:
            for _, r_p in df_prog.iterrows():
                id_p = str(r_p['id_espacio']).strip() if pd.notna(r_p['id_espacio']) else ""
                nom_p = str(r_p['nombre_espacio']).strip() if pd.notna(r_p['nombre_espacio']) else ""
                zon_p = str(r_p['zona']).strip() if pd.notna(r_p['zona']) else ""
                area_p = float(pd.to_numeric(r_p['area_total_m2'], errors='coerce') or 0.0)
                programa_04_registros.append((
                    normalizar_texto_clave(id_p),
                    normalizar_texto_clave(nom_p),
                    normalizar_texto_clave(zon_p)
                ))
                n_key = normalizar_texto_clave(nom_p)
                if n_key:
                    programa_04_areas_por_nombre[n_key] = programa_04_areas_por_nombre.get(n_key, 0.0) + area_p

            palabras_omitidas = ["PASILLO", "ANDADORES", "CIRCULACIONES", "ESTACIONAMIENTO", "PARKING", "DISCAPACITADOS"]
            incongruencia_encontrada = False

            for idx, row in df_prog.iterrows():
                nom_esp_raw = str(row['nombre_espacio']).strip() if pd.notna(row['nombre_espacio']) else ""
                nom_esp_norm = nom_esp_raw.lower()
                cant_esp = row['cantidad']
                
                es_excepcion = any(exc in nom_esp_raw.upper() for exc in palabras_omitidas)
                if es_excepcion:
                    continue
                
                match_encontrado = False
                for eu in espacios_usuarios_guardados:
                    nec_esp_norm = str(eu['necesidad_espacial']).strip().lower()
                    cant_est = eu['cantidad_estimada']
                    
                    if nom_esp_norm == nec_esp_norm:
                        try:
                            if float(cant_esp) == float(cant_est):
                                match_encontrado = True
                                break
                        except:
                            if str(cant_esp).strip() == str(cant_est).strip():
                                match_encontrado = True
                                break
                
                if not match_encontrado:
                    incongruencia_encontrada = True
                    mensajes_programa.append(
                        f"⚠️ Incongruencia en fila {idx+1} (Espacio: '{nom_esp_raw}', Cantidad: {cant_esp}): "
                        "No coincide de forma idéntica en nombre y cantidad con la Plantilla 03 (Análisis de Usuario). "
                        "Te invitamos a verificar los datos en la plantilla correspondiente."
                    )

            if incongruencia_encontrada:
                score_programa = 0.0
                mensajes_programa.append("❌ Propuesta rechazada en su totalidad (0 pts) por incongruencias respecto al Análisis de Usuario (Plantilla 03). Verifica los datos en la plantilla correspondiente.")
            else:
                score_programa += 50.0
                mensajes_programa.append("✅ Coherencia idéntica con el Análisis de Usuario (Plantilla 03): +50 pts base.")

                try:
                    terreno_val = terreno_datos_globales.get('superficie_terreno_m2', 0.0)
                    cus_coef = terreno_datos_globales.get('cus_permitido_coeficiente', 0.0)
                    cus_maximo_permitido = terreno_val * cus_coef
                    
                    suma_area_total = pd.to_numeric(df_prog['area_total_m2'], errors='coerce').sum()
                    
                    if suma_area_total > cus_maximo_permitido:
                        score_programa = 0.0
                        mensajes_programa.append(f"❌ Propuesta rechazada (0 pts): El área total construida ({suma_area_total:,.2f} m²) excede el CUS máximo permitido ({cus_maximo_permitido:,.2f} m²).")
                    else:
                        score_programa += 20.0
                        mensajes_programa.append(f"✅ Cumple con el CUS ({suma_area_total:,.2f} m² <= {cus_maximo_permitido:,.2f} m²): +20 pts.")

                        cos_pct = terreno_datos_globales.get('cos_permitido_porcentaje', 0.0)
                        cos_factor = (cos_pct / 100.0) if cos_pct > 1.0 else cos_pct
                        cos_maximo_permitido = terreno_val * cos_factor
                        
                        df_pb = df_prog[pd.to_numeric(df_prog['nivel'], errors='coerce') == 0]
                        suma_area_pb = pd.to_numeric(df_pb['area_total_m2'], errors='coerce').sum()
                        
                        if suma_area_pb > cos_maximo_permitido:
                            score_programa = 0.0
                            mensajes_programa.append(f"❌ Propuesta rechazada (0 pts): El desplante en Planta Baja ({suma_area_pb:,.2f} m²) excede el COS máximo permitido ({cos_maximo_permitido:,.2f} m²).")
                        else:
                            score_programa += 20.0
                            mensajes_programa.append(f"✅ Cumple con el COS en Planta Baja ({suma_area_pb:,.2f} m² <= {cos_maximo_permitido:,.2f} m²): +20 pts.")

                            relaciones_completas = all(
                                pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan"
                                for val in df_prog['relacion_directa_ids']
                            )

                            if relaciones_completas:
                                score_programa += 10.0
                                mensajes_programa.append("✅ Columna 'relacion_directa_ids' completa en todas las celdas: +10 pts (Total 100 pts en el rubro).")
                            else:
                                score_programa += 5.0
                                mensajes_programa.append(
                                    "⚠️ La columna 'relacion_directa_ids' tiene celdas vacías: +5 pts (en lugar de 10 pts). "
                                    "Te invitamos a completar todas las celdas; en el caso extraordinario de que el espacio no tenga relación directa con ningún espacio, debes colocar el texto N/A."
                                )

                except Exception as ex_cus:
                    score_programa = 0.0
                    mensajes_programa.append(f"❌ Error al validar CUS/COS con Entregable 02: {ex_cus}. Propuesta rechazada (0 pts).")
    else:
        score_programa = 0.0
        mensajes_programa.append("❌ Entregable 04 no cargado.")

    # ==========================================
    # EVALUACIÓN 05: ESTACIONAMIENTO
    # ==========================================
    score_estacionamiento = 0.0
    mensajes_estacionamiento = []

    if "05_estacionamiento.xlsx" in uploaded_files:
        df_est = uploaded_files["05_estacionamiento.xlsx"]
        if df_est.empty:
            score_estacionamiento = 0.0
            mensajes_estacionamiento.append("❌ El archivo '05_estacionamiento.xlsx' está vacío (0 pts). Corrige la carga.")
        else:
            score_estacionamiento += 50.0
            mensajes_estacionamiento.append("✅ Archivo '05_estacionamiento.xlsx' cargado correctamente sin alteración de nombre: +50 pts base.")

            if len(df_est.columns) > 6:
                suma_m2_05 = pd.to_numeric(df_est.iloc[:, 6], errors='coerce').fillna(0.0).sum()
            else:
                suma_m2_05 = 0.0

            if "04_programa_arquitectonico.xlsx" in uploaded_files:
                df_prog_ref = uploaded_files["04_programa_arquitectonico.xlsx"]
                if 'nombre_espacio' in df_prog_ref.columns and 'area_total_m2' in df_prog_ref.columns:
                    mask_est_04 = df_prog_ref['nombre_espacio'].astype(str).str.strip().str.upper() == "ESTACIONAMIENTO"
                    if mask_est_04.any():
                        suma_m2_04 = pd.to_numeric(df_prog_ref.loc[mask_est_04, 'area_total_m2'], errors='coerce').fillna(0.0).sum()
                        if abs(suma_m2_05 - suma_m2_04) <= 0.01:
                            score_estacionamiento += 30.0
                            mensajes_estacionamiento.append(
                                f"✅ Coherencia de m² validada con Programa Arquitectónico (Plantilla 05: {suma_m2_05:,.2f} m² vs. ESTACIONAMIENTO en Plantilla 04: {suma_m2_04:,.2f} m²): +30 pts (Acumulado: {score_estacionamiento:.0f} pts)."
                            )
                        else:
                            mensajes_estacionamiento.append(
                                f"⚠️ Incongruencia de m² entre Plantilla 05 ({suma_m2_05:,.2f} m²) y 'ESTACIONAMIENTO' en Plantilla 04 ({suma_m2_04:,.2f} m²): +0 pts en este tramo. "
                                "Revisa y corrige la Plantilla 04 y vuelve a cargarla si corresponde (esto no penaliza la calificación de la Plantilla 04)."
                            )
                    else:
                        mensajes_estacionamiento.append(
                            "⚠️ No se encontró la clasificación 'ESTACIONAMIENTO' en la columna 'nombre_espacio' (Col. B) de la Plantilla 04: +0 pts en este tramo. "
                            "Te invitamos a agregar dicho rubro en tu Programa Arquitectónico o reclasificarlo bajo el nombre exacto: `ESTACIONAMIENTO` "
                            "(puedes copiar y pegar `ESTACIONAMIENTO`), corregir la Plantilla 04 y volver a cargarla. Esto no afecta el puntaje de la Plantilla 04."
                        )
                else:
                    mensajes_estacionamiento.append("⚠️ No fue posible leer 'nombre_espacio' o 'area_total_m2' en la Plantilla 04 para validar m² (+0 pts en este tramo).")
            else:
                mensajes_estacionamiento.append(
                    "⚠️ La Plantilla 04 no está cargada para validar los m² de `ESTACIONAMIENTO` (+0 pts en este tramo). "
                    "Si requieres ajustarla, agrega o reclasifica bajo el nombre exacto `ESTACIONAMIENTO` en tu Plantilla 04 y cárgalo (sin penalización a la Plantilla 04)."
                )

            cajones_05 = 0
            tiene_col_cajones = 'cajones_requeridos' in df_est.columns
            if tiene_col_cajones:
                cajones_05 = pd.to_numeric(df_est['cajones_requeridos'], errors='coerce').fillna(0).sum()

            cajones_norm_requeridos = None
            if "01_normativa.xlsx" in uploaded_files:
                df_norm_ref = uploaded_files["01_normativa.xlsx"]
                if 'condicionante_diseno' in df_norm_ref.columns and 'parametro_evaluable' in df_norm_ref.columns:
                    valores_cajones_norm = []
                    for _, r_n in df_norm_ref.iterrows():
                        param_n = str(r_n['parametro_evaluable']).strip().upper() if pd.notna(r_n['parametro_evaluable']) else ""
                        if param_n == "CANTIDAD DE CAJONES":
                            entero_f = extraer_entero_estricto(r_n['condicionante_diseno'])
                            if entero_f is not None:
                                valores_cajones_norm.append(entero_f)
                    if valores_cajones_norm:
                        cajones_norm_requeridos = max(valores_cajones_norm)

            if not tiene_col_cajones:
                mensajes_estacionamiento.append(
                    "⚠️ No se encontró la columna 'cajones_requeridos' (Col. B) en la Plantilla 05 para validar cajones (+0 pts en este tramo)."
                )
            elif cajones_norm_requeridos is None:
                mensajes_estacionamiento.append(
                    "⚠️ No se encontró información suficiente en la Plantilla 01 para validar cajones (+0 pts en este tramo). "
                    "Te invitamos a agregar la normativa correspondiente, capturar el número entero en 'condicionante_diseno' (Col. F) "
                    "y clasificar como `CANTIDAD DE CAJONES` en 'parametro_evaluable' (Col. G). "
                    "Esto no penaliza los puntos de la Plantilla 01; si lo deseas, es opcional volver a cargar únicamente el Excel de la Plantilla 01 con esta clasificación."
                )
            else:
                if cajones_05 >= cajones_norm_requeridos:
                    score_estacionamiento += 20.0
                    mensajes_estacionamiento.append(
                        f"✅ Cantidad de cajones validada ({cajones_05:g} cajones propuestos >= {cajones_norm_requeridos} requeridos en Plantilla 01): +20 pts (Total: {score_estacionamiento:.0f}/100 pts)."
                    )
                else:
                    mensajes_estacionamiento.append(
                        f"❌ Cajones insuficientes ({cajones_05:g} cajones en Plantilla 05 < {cajones_norm_requeridos} requeridos en Plantilla 01): +0 pts en este tramo. "
                        "Esto no penaliza los puntos de la Plantilla 01; puedes revisar tu propuesta o recargar opcionalmente la Plantilla 01 si requiere ajuste de clasificación."
                    )
    else:
        score_estacionamiento = 0.0
        mensajes_estacionamiento.append("❌ Entregable 05 no cargado. El archivo debe llamarse exactamente '05_estacionamiento.xlsx'.")

    # ==========================================
    # EVALUACIÓN 06: MATRIZ DE FUNCIONAMIENTO (V2.2)
    # ==========================================
    score_matriz = 0.0
    mensajes_matriz = []
    espacios_06_visual = []
    relaciones_06_visual = []

    if "06_matriz_funcionamientoV2.xlsx" in uploaded_files:
        payload_06 = uploaded_files["06_matriz_funcionamientoV2.xlsx"]
        df_esp_06 = payload_06["Espacios"]
        df_rel_06 = payload_06["Relaciones"]

        expected_cols_6_esp = ['id_espacio', 'nombre_espacio', 'zona']
        expected_cols_6_rel = ['Espacio A', 'Espacio B', 'Tipo de relación']

        cols_esp_actuales = [str(c).strip() for c in df_esp_06.columns if str(c).strip().lower() != 'notas']
        cols_rel_actuales = [str(c).strip() for c in df_rel_06.columns if str(c).strip().lower() != 'notas']

        if cols_esp_actuales != expected_cols_6_esp or cols_rel_actuales != expected_cols_6_rel:
            score_matriz = 0.0
            mensajes_matriz.append(
                "❌ Las columnas de las hojas 'Espacios' o 'Relaciones' en '06_matriz_funcionamientoV2.xlsx' no coinciden con la estructura esperada. "
                "Propuesta rechazada (0 pts). Verifica los encabezados del archivo."
            )
        elif df_esp_06.dropna(how='all').empty:
            score_matriz = 0.0
            mensajes_matriz.append("❌ La hoja 'Espacios' de '06_matriz_funcionamientoV2.xlsx' está vacía (0 pts). Corrige la carga.")
        else:
            score_matriz = 100.0
            mensajes_matriz.append("✅ Archivo '06_matriz_funcionamientoV2.xlsx', hojas y columnas validados: 100 pts iniciales.")

            espacios_06_nombres_norm = set()

            for _, r_e in df_esp_06.iterrows():
                if r_e.isna().all():
                    continue
                id_e = str(r_e['id_espacio']).strip() if pd.notna(r_e['id_espacio']) else ""
                nom_e = str(r_e['nombre_espacio']).strip() if pd.notna(r_e['nombre_espacio']) else ""
                zon_e = str(r_e['zona']).strip() if pd.notna(r_e['zona']) else ""
                if not nom_e:
                    continue

                nom_k = normalizar_texto_clave(nom_e)
                zon_k = normalizar_texto_clave(zon_e)

                espacios_06_nombres_norm.add(nom_k)

                zona_cod = ALIAS_ZONA_DIAGRAMA.get(zon_k, "semi")
                area_estimada = programa_04_areas_por_nombre.get(nom_k, 12.0)

                espacios_06_visual.append({
                    "id": id_e,
                    "nombre": nom_e,
                    "zona_raw": zon_e,
                    "zona": zona_cod,
                    "area": area_estimada,
                })

            if programa_04_registros:
                espacios_06_tuplas = {
                    (
                        normalizar_texto_clave(item["id"]),
                        normalizar_texto_clave(item["nombre"]),
                        normalizar_texto_clave(item["zona_raw"]),
                    )
                    for item in espacios_06_visual
                }
                set_04 = set(programa_04_registros)
                if espacios_06_tuplas != set_04:
                    score_matriz -= 20.0
                    mensajes_matriz.append(
                        "⚠️ [Penalización -20 pts] Se detectaron diferencias en 'id_espacio', 'nombre_espacio' o 'zona' respecto a la Plantilla 04. "
                        "Te invitamos a corregir estos datos en '06_matriz_funcionamientoV2.xlsx' para no descontar puntos."
                    )
                else:
                    mensajes_matriz.append("✅ Coherencia exacta de 'id_espacio', 'nombre_espacio' y 'zona' con la Plantilla 04.")
            else:
                mensajes_matriz.append("ℹ️ La Plantilla 04 no está cargada para contrastar 'id_espacio', 'nombre_espacio' y 'zona'.")

            tipos_validos_06 = {'directa', 'indirecta', 'circulacion', 'visual'}
            falta_relaciones = False
            relaciones_validas_count = 0
            espacios_conectados = set()
            pares_conectados = []

            for _, r_r in df_rel_06.iterrows():
                if r_r.isna().all():
                    continue
                sp_a = str(r_r['Espacio A']).strip() if pd.notna(r_r['Espacio A']) else ""
                sp_b = str(r_r['Espacio B']).strip() if pd.notna(r_r['Espacio B']) else ""
                t_rel_raw = str(r_r['Tipo de relación']).strip() if pd.notna(r_r['Tipo de relación']) else ""
                t_rel_norm = normalizar_texto_clave(t_rel_raw)

                if not sp_a or not sp_b or not t_rel_norm or t_rel_norm not in tipos_validos_06:
                    falta_relaciones = True
                    continue

                relaciones_validas_count += 1
                relaciones_06_visual.append({
                    "origen": sp_a,
                    "destino": sp_b,
                    "tipo": t_rel_raw,
                })

                a_k = normalizar_texto_clave(sp_a)
                b_k = normalizar_texto_clave(sp_b)
                espacios_conectados.add(a_k)
                espacios_conectados.add(b_k)
                pares_conectados.append((a_k, b_k))

            if relaciones_validas_count == 0 or falta_relaciones:
                score_matriz -= 50.0
                mensajes_matriz.append("❌ [Penalización -50 pts] No existen relaciones válidas o hay filas sin 'Espacio A', 'Espacio B' o 'Tipo de relación' completos.")
            else:
                mensajes_matriz.append(f"✅ Se validaron {relaciones_validas_count} relación(es) completa(s) en la hoja 'Relaciones'.")

            hay_saltos_o_aislados = False
            if espacios_06_nombres_norm:
                aislados_06 = espacios_06_nombres_norm - espacios_conectados
                if aislados_06:
                    hay_saltos_o_aislados = True
                    mensajes_matriz.append(f"⚠️ Existen espacios sin conectarse en las relaciones: {sorted(list(aislados_06))}.")

            if pares_conectados and not hay_saltos_o_aislados:
                adj = {node: set() for node in espacios_conectados}
                for u, v in pares_conectados:
                    adj[u].add(v)
                    adj[v].add(u)
                start_node = next(iter(espacios_conectados))
                visited = {start_node}
                queue = [start_node]
                while queue:
                    curr = queue.pop(0)
                    for nxt in adj[curr]:
                        if nxt not in visited:
                            visited.add(nxt)
                            queue.append(nxt)
                if len(visited) < len(espacios_conectados):
                    hay_saltos_o_aislados = True
                    mensajes_matriz.append("⚠️ El diagrama de relaciones presenta saltos o sub-redes desconectadas.")

            if hay_saltos_o_aislados:
                score_matriz -= 20.0
                mensajes_matriz.append("❌ [Penalización -20 pts] Se detectaron espacios sin conectar o saltos en el flujo.")

            score_matriz = max(10.0, score_matriz)
    else:
        score_matriz = 0.0
        mensajes_matriz.append("❌ Entregable 06 no cargado. El archivo debe llamarse exactamente '06_matriz_funcionamientoV2.xlsx'.")

    # ==========================================
    # EVALUACIÓN 08: CARACTERÍSTICAS TÉCNICAS
    # ==========================================
    score_tecnicas = 0.0
    mensajes_tecnicas = []

    if "08_caracteristicas_tecnicas.xlsx" in uploaded_files:
        df_tec = uploaded_files["08_caracteristicas_tecnicas.xlsx"]
        expected_cols_8 = ['id_caracteristica', 'categoria_tecnica', 'descripcion_propuesta', 'especificacion_tecnica', 'norma_referencia']

        if list(df_tec.columns) != expected_cols_8:
            score_tecnicas = 0.0
            mensajes_tecnicas.append("❌ Las columnas de Características Técnicas no coinciden exactamente con la estructura base. Propuesta rechazada (0 pts).")
        elif df_tec.empty:
            score_tecnicas = 0.0
            mensajes_tecnicas.append("❌ El archivo de Características Técnicas está vacío (0 pts).")
        else:
            categorias_requeridas = {'estructural', 'hidraulica', 'sanitaria', 'electrica', 'sustentabilidad'}
            df_tec['cat_norm'] = df_tec['categoria_tecnica'].dropna().astype(str).str.strip().str.lower()
            cats_presentes = set(df_tec['cat_norm'].unique())
            cats_validas_presentes = cats_presentes.intersection(categorias_requeridas)
            
            puntos_base_tec = len(cats_validas_presentes) * 10.0
            score_tecnicas += puntos_base_tec
            mensajes_tecnicas.append(f"✅ Categorías técnicas presentes ({len(cats_validas_presentes)}/5): +{puntos_base_tec} pts base.")
            
            faltantes_tec = categorias_requeridas - cats_validas_presentes
            if faltantes_tec:
                mensajes_tecnicas.append(f"⚠️ Faltan las siguientes categorías obligatorias (-10 pts c/u): {list(faltantes_tec)}.")

            conteo_cat = df_tec['cat_norm'].value_counts()
            puntos_adicionales_tec = 0.0
            for cat in categorias_requeridas:
                num_prop = conteo_cat.get(cat, 0)
                if num_prop >= 3:
                    puntos_adicionales_tec += 10.0
                    mensajes_tecnicas.append(f"⭐ Categoría '{cat}' cuenta con {num_prop} propuestas (>= 3): +10 pts adicionales.")
                else:
                    mensajes_tecnicas.append(f"ℹ️ Categoría '{cat}' cuenta con {num_prop} propuesta(s) (< 3): 0 pts adicionales.")

            score_tecnicas += puntos_adicionales_tec
            score_tecnicas = min(score_tecnicas, 100.0)
    else:
        score_tecnicas = 0.0
        mensajes_tecnicas.append("❌ Entregable 08 no cargado.")

    # ==========================================
    # CÁLCULO GENERAL DE PONDERACIÓN V2.0
    # ==========================================
    score_total = 0.0
    peso_cargado_total = 0.0
    detalles_puntuacion = []
    
    for filename, label, weight in deliverables_config:
        if filename == "01_normativa.xlsx":
            sec_score = score_normativa
            estado_txt = "Evaluado (V2.0)"
        elif filename == "02_terreno_cos_cus.xlsx":
            sec_score = score_terreno
            estado_txt = "Evaluado (V2.0)"
        elif filename == "03_analisis_usuario.xlsx":
            sec_score = score_usuario
            estado_txt = "Evaluado (V2.0)"
        elif filename == "04_programa_arquitectonico.xlsx":
            sec_score = score_programa
            estado_txt = "Evaluado (V2.0)"
        elif filename == "05_estacionamiento.xlsx":
            sec_score = score_estacionamiento
            estado_txt = "Evaluado (V2.0)" if filename in uploaded_files else "Pendiente de Carga"
        elif filename == "06_matriz_funcionamientoV2.xlsx":
            sec_score = score_matriz
            estado_txt = "Evaluado (V2.2)" if filename in uploaded_files else "Pendiente de Carga"
        elif filename == "07_concepto_volumetria.xlsx" and filename in uploaded_files:
            sec_score = 90.0
            estado_txt = "Evaluado (Estándar)"
        elif filename == "08_caracteristicas_tecnicas.xlsx":
            sec_score = score_tecnicas
            estado_txt = "Evaluado (V2.0)"
        elif filename in uploaded_files:
            sec_score = 85.0
            estado_txt = "Cargado (Pendiente V2.0)"
        else:
            sec_score = 0.0
            estado_txt = "Pendiente de Carga"

        if filename in uploaded_files:
            score_total += sec_score * (weight / 100.0)
            peso_cargado_total += weight
            detalles_puntuacion.append((label, weight, sec_score, estado_txt))
        else:
            detalles_puntuacion.append((label, weight, 0.0, estado_txt))

    calificacion_final = score_total if is_complete else (score_total / (peso_cargado_total / 100.0) if peso_cargado_total > 0 else 0)
    fecha_hora_gen = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    return {
        "missing_files": missing_files,
        "is_complete": is_complete,
        "tipo_evaluacion": tipo_evaluacion,
        "calificacion_final": calificacion_final,
        "fecha_hora_gen": fecha_hora_gen,
        "nombre_proyecto": nombre_proyecto,
        "tipologia_proyecto": tipologia_proyecto,
        "superficie_terreno_input": superficie_terreno_input,
        "usuario_sesion": etiqueta_usuario,
        "integrantes": integrantes,
        "detalles_puntuacion": detalles_puntuacion,
        "score_normativa": score_normativa,
        "mensajes_normativa": mensajes_normativa,
        "score_terreno": score_terreno,
        "mensajes_terreno": mensajes_terreno,
        "score_usuario": score_usuario,
        "mensajes_usuario": mensajes_usuario,
        "score_programa": score_programa,
        "mensajes_programa": mensajes_programa,
        "score_estacionamiento": score_estacionamiento,
        "mensajes_estacionamiento": mensajes_estacionamiento,
        "score_matriz": score_matriz,
        "mensajes_matriz": mensajes_matriz,
        "espacios_06_visual": espacios_06_visual,
        "relaciones_06_visual": relaciones_06_visual,
        "score_tecnicas": score_tecnicas,
        "mensajes_tecnicas": mensajes_tecnicas,
    }

if st.sidebar.button("🚀 Ejecutar Evaluación V2.0", type="primary"):
    st.session_state.resultado_evaluacion = ejecutar_motor_evaluacion()
    st.session_state.evaluacion_activa = True

if st.session_state.evaluacion_activa and st.session_state.resultado_evaluacion is not None:
    res = st.session_state.resultado_evaluacion
    missing_files = res["missing_files"]
    is_complete = res["is_complete"]
    tipo_evaluacion = res["tipo_evaluacion"]
    calificacion_final = res["calificacion_final"]
    detalles_puntuacion = res["detalles_puntuacion"]
    score_normativa = res["score_normativa"]
    mensajes_normativa = res["mensajes_normativa"]
    score_terreno = res["score_terreno"]
    mensajes_terreno = res["mensajes_terreno"]
    score_usuario = res["score_usuario"]
    mensajes_usuario = res["mensajes_usuario"]
    score_programa = res["score_programa"]
    mensajes_programa = res["mensajes_programa"]
    score_estacionamiento = res["score_estacionamiento"]
    mensajes_estacionamiento = res["mensajes_estacionamiento"]
    score_matriz = res["score_matriz"]
    mensajes_matriz = res["mensajes_matriz"]
    espacios_06_visual = res["espacios_06_visual"]
    relaciones_06_visual = res["relaciones_06_visual"]
    score_tecnicas = res["score_tecnicas"]
    mensajes_tecnicas = res["mensajes_tecnicas"]

    st.markdown(f"## Estado del Dictamen: **{tipo_evaluacion}**")
    
    if not is_complete:
        st.warning("⚠️ Faltan entregables por cargar para completar la evaluación final. Puedes seguir cargando plantillas de forma escalonada sin perder las anteriores.")
        with st.expander("📋 Elementos faltantes por cargar:", expanded=True):
            for _, label, weight in missing_files:
                st.write(f"- **{label}** (Ponderación asignada: {weight}%)")
    else:
        st.success("🎉 ¡Todos los entregables han sido cargados!")

    # Visualización de Resultados
    st.markdown("---")
    col_sc1, col_sc2 = st.columns([1, 2])
    with col_sc1:
        st.metric("Calificación Ponderada V2.0", f"{calificacion_final:,.1f} / 100", delta=f"{tipo_evaluacion}")
    with col_sc2:
        st.markdown(f"### Dictamen General V2.0")
        if calificacion_final < 60:
            st.error("😢 Desempeño bajo. No cumple con los mínimos requeridos.")
        elif calificacion_final <= 85:
            st.warning("😐 Desempeño aceptable con observaciones pendientes.")
        else:
            st.success("😃 ¡Excelente desempeño! Alta congruencia técnica.")

    # Botón de Imprimir / Descargar Reporte en PDF (conserva intacto todo lo cargado)
    pdf_bytes = generar_pdf_reporte_bytes(res)
    nombre_pdf_descarga = f"reporte_evaluacion_V2_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    col_imp1, col_imp2 = st.columns([2, 3])
    with col_imp1:
        st.download_button(
            label="🖨️ Imprimir Reporte (Descargar PDF)",
            data=pdf_bytes,
            file_name=nombre_pdf_descarga,
            mime="application/pdf",
            key="btn_imprimir_pdf_v2",
            help="Descarga tu reporte completo en PDF sin borrar ningún archivo ni evaluación previa."
        )
    with col_imp2:
        st.caption(f"Generado: {res['fecha_hora_gen']} • Tu carga y evaluaciones permanecen activas para continuar ajustando.")

    st.markdown("---")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Puntuación V2.0", 
        "📏 Auditoría 01 (Normativa)", 
        "📐 Auditoría 02 (Terreno)", 
        "👥 Auditoría 03 (Usuario)", 
        "🏢 Auditoría 04 (Programa)",
        "🅿️ Auditoría 05 (Estacionamiento)",
        "🔗 Auditoría 06 (Funcionamiento)",
        "⚙️ Auditoría 08 (Técnicas)"
    ])
    
    with tab1:
        st.subheader("Desglose de Ponderación V2.0")
        df_score_table = pd.DataFrame(detalles_puntuacion, columns=["Entregable", "Peso (%)", "Calificación (0-100)", "Estado"])
        st.dataframe(df_score_table, use_container_width=True)

    with tab2:
        st.subheader("Auditoría Detallada - Entregable 01: Marco Normativo")
        st.metric("Calificación Sección 01", f"{score_normativa:,.1f} / 100")
        for msg in mensajes_normativa:
            st.write(msg)

    with tab3:
        st.subheader("Auditoría Detallada - Entregable 02: COS, CUS y Terreno")
        st.metric("Calificación Sección 02", f"{score_terreno:,.1f} / 100")
        for msg in mensajes_terreno:
            st.write(msg)

    with tab4:
        st.subheader("Auditoría Detallada - Entregable 03: Análisis de Usuario")
        st.metric("Calificación Sección 03", f"{score_usuario:,.1f} / 100")
        for msg in mensajes_usuario:
            st.write(msg)

    with tab5:
        st.subheader("Auditoría Detallada - Entregable 04: Programa Arquitectónico")
        st.metric("Calificación Sección 04", f"{score_programa:,.1f} / 100")
        for msg in mensajes_programa:
            st.write(msg)

    with tab6:
        st.subheader("Auditoría Detallada - Entregable 05: Estacionamiento")
        st.metric("Calificación Sección 05", f"{score_estacionamiento:,.1f} / 100")
        for msg in mensajes_estacionamiento:
            st.write(msg)

    with tab7:
        st.subheader("Auditoría Detallada - Entregable 06: Matriz de Funcionamiento")
        st.metric("Calificación Sección 06", f"{score_matriz:,.1f} / 100")
        for msg in mensajes_matriz:
            st.write(msg)

        if "06_matriz_funcionamientoV2.xlsx" in st.session_state.entregables_guardados and espacios_06_visual:
            st.markdown("---")
            st.markdown("### 🖼️ Visualizador Gráfico Opcional del Diagrama de Funcionamiento (1920×1080 px)")
            st.caption("Herramienta opcional e independiente de la evaluación: sirve para validar visualmente los datos capturados y descargar el diagrama en 1920×1080 px para integrar a tu lámina.")

            activar_vis = st.checkbox("Mostrar / activar visualizador gráfico del diagrama (opcional)", value=True, key="activar_vis_06")
            if activar_vis:
                col_v1, col_v2, col_v3 = st.columns([2, 1, 1])
                with col_v1:
                    titulo_v = st.text_input("Título del diagrama", value=nombre_proyecto or "Diagrama de Funcionamiento", key="titulo_diag_06")
                with col_v2:
                    semilla_v = st.number_input("Semilla de acomodo", min_value=1, max_value=9999, value=1, step=1, key="semilla_diag_06")
                with col_v3:
                    zonas_disp = ["Todas"] + sorted(list({e["zona_raw"] for e in espacios_06_visual if e["zona_raw"]}))
                    zona_filtro = st.selectbox("Filtrar por zona", zonas_disp, key="zona_diag_06")

                if zona_filtro != "Todas":
                    z_key_sel = normalizar_texto_clave(zona_filtro)
                    esp_filtrados = [e for e in espacios_06_visual if normalizar_texto_clave(e["zona_raw"]) == z_key_sel]
                    noms_visibles = {normalizar_texto_clave(e["nombre"]) for e in esp_filtrados}
                    rel_filtradas = [
                        r for r in relaciones_06_visual
                        if normalizar_texto_clave(r["origen"]) in noms_visibles and normalizar_texto_clave(r["destino"]) in noms_visibles
                    ]
                else:
                    esp_filtrados = espacios_06_visual
                    rel_filtradas = relaciones_06_visual

                if esp_filtrados:
                    png_1080 = generar_diagrama_1920x1080(
                        esp_filtrados,
                        rel_filtradas,
                        titulo_diag=titulo_v,
                        semilla=int(semilla_v)
                    )
                    st.image(png_1080, caption=f"{titulo_v} — Resolución: 1920×1080 px", use_container_width=True)
                    st.download_button(
                        label="⬇️ Descargar Diagrama de Flujo (PNG 1920×1080 px)",
                        data=png_1080,
                        file_name="diagrama_funcionamiento_1920x1080.png",
                        mime="image/png",
                        key="dl_png_06_1080"
                    )

    with tab8:
        st.subheader("Auditoría Detallada - Entregable 08: Características Técnicas")
        st.metric("Calificación Sección 08", f"{score_tecnicas:,.1f} / 100")
        for msg in mensajes_tecnicas:
            st.write(msg)
else:
    st.info("👈 Llena los datos generales, carga tus archivos de forma escalonada en la barra lateral y haz clic en **Ejecutar Evaluación V2.0**.")