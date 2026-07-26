import json
import io
import sqlite3
import base64
import streamlit as st
import pandas as pd
from PIL import Image
from groq import Groq
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from xhtml2pdf import pisa

# --- BASE DE DATOS (SQLITE) ---
DB_NAME = "presupuestos.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS presupuestos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            destino TEXT,
            fecha_salida TEXT,
            duracion TEXT,
            operador TEXT,
            total_usd REAL,
            datos_json TEXT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def guardar_en_db(datos, total_calc):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        INSERT INTO presupuestos (destino, fecha_salida, duracion, operador, total_usd, datos_json)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (datos['destino'], datos['fecha_salida'], datos['duracion'], datos['operador'], total_calc, json.dumps(datos)))
    conn.commit()
    conn.close()

def cargar_historial():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT id, destino, fecha_salida, operador, total_usd, fecha_registro FROM presupuestos ORDER BY id DESC", conn)
    conn.close()
    return df

# Inicializar Base de Datos
init_db()

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestor de Presupuestos de Viaje", page_icon="✈️", layout="wide")

st.title("✈️ Consolidador de Presupuestos de Viaje")

# Opciones de Navegación
menu = st.sidebar.radio("Navegación", ["🧮 Crear Presupuesto", "📜 Historial Guardado"])

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Configuración")

# DETECCIÓN AUTOMÁTICA DE API KEY DE GROQ
api_key = ""
if "GROQ_API_KEY" in st.secrets:
    api_key = st.secrets["GROQ_API_KEY"]

if api_key:
    st.sidebar.success("🔑 API Key cargada automáticamente")
    user_key = st.sidebar.text_input("Groq API Key (opcional si deseas cambiarla)", value=api_key, type="password")
    api_key = user_key
else:
    api_key = st.sidebar.text_input("Ingresa tu Groq API Key (gsk_...)", type="password", help="Obtenla gratis en console.groq.com")

# Control de versión para refrescar formulario
if "form_version" not in st.session_state:
    st.session_state["form_version"] = 0

if "datos_viaje" not in st.session_state:
    st.session_state["datos_viaje"] = {
        "destino": "Costa Rica",
        "fecha_salida": "04/12/2026",
        "duracion": "9 Días / 8 Noches",
        "operador": "Mediterránea Turismo",
        "precio_base": 2649.00,
        "impuestos_aereo": 444.00,
        "cuotas_cant": 5,
        "cuota_monto": 618.60,
        "equipaje_extra": 100.00,
        "comidas_extra": 280.00,
        "excursiones_extra": 120.00,
        "seguro_medico": 75.00,
        "gastos_personales": 220.00,
        "notas": "Incluye 6N c/desayuno + 2N pensión completa + traslados y tours."
    }

# --- GENERADORES DE ARCHIVOS ---
def generar_excel(datos, total_calc):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Presupuesto Viaje"
    
    font_title = Font(name="Calibri", size=14, bold=True, color="1E4D2B")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    fill_header = PatternFill(start_color="1E4D2B", end_color="1E4D2B", fill_type="solid")
    fill_accent = PatternFill(start_color="E8F0E6", end_color="E8F0E6", fill_type="solid")
    
    ws["A1"] = f"PRESUPUESTO INTEGRADO - {datos['destino'].upper()}"
    ws["A1"].font = font_title
    ws["A2"] = f"Fecha: {datos['fecha_salida']} | Duración: {datos['duracion']} | Operador: {datos['operador']}"
    
    headers = ["Categoría", "Concepto / Ítem", "Fuente", "Costo (USD)"]
    ws.append([])
    ws.append(headers)
    
    for col_num in range(1, 5):
        cell = ws.cell(row=4, column=col_num)
        cell.font = font_header
        cell.fill = fill_header

    filas = [
        ("Transporte", "Paquete Base Terrestre", "Imagen / Promo", datos['precio_base']),
        ("Transporte", "Impuestos Aéreos", "Imagen / Promo", datos['impuestos_aereo']),
        ("Transporte", "Equipaje Facturado Extra", "Manual", datos['equipaje_extra']),
        ("Alimentación", "Comidas No Incluidas", "Manual", datos['comidas_extra']),
        ("Excursiones", "Actividades Extras / Parques", "Manual", datos['excursiones_extra']),
        ("Seguro", "Asistencia Médica Internacional", "Manual", datos['seguro_medico']),
        ("Varios", "Propinas e Imprevistos", "Manual", datos['gastos_personales']),
    ]
    
    for f in filas:
        ws.append(list(f))
        
    tot_row = len(filas) + 5
    ws.cell(row=tot_row, column=1, value="TOTAL ESTIMADO (USD)").font = font_bold
    ws.cell(row=tot_row, column=4, value=f"=SUM(D5:D{tot_row-1})").font = font_bold
    
    for c in range(1, 5):
        ws.cell(row=tot_row, column=c).fill = fill_accent

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def generar_pdf(datos, total_calc):
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Helvetica, Arial, sans-serif; color: #2C3E50; padding: 20px; }}
            .header {{ background-color: #1E4D2B; color: white; padding: 15px; border-radius: 6px; }}
            h1 {{ margin: 0; font-size: 20px; }}
            .card {{ background-color: #F8FAF7; border: 1px solid #D3D3D3; padding: 10px; margin-top: 15px; border-radius: 6px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th {{ background-color: #1E4D2B; color: white; padding: 8px; font-size: 12px; }}
            td {{ padding: 8px; border-bottom: 1px solid #E0E0E0; font-size: 12px; }}
            .total {{ font-weight: bold; background-color: #E8F0E6; }}
            .text-right {{ text-align: right; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>Presupuesto Integrado de Viaje: {datos['destino']}</h1>
            <p>Salida: {datos['fecha_salida']} | Duración: {datos['duracion']} | Operador: {datos['operador']}</p>
        </div>
        
        <div class="card">
            <h3>Resumen Económico Total: USD ${total_calc:,.2f}</h3>
            <p><strong>Financiación Base:</strong> {datos['cuotas_cant']} cuotas de USD ${datos['cuota_monto']:,.2f}</p>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Categoría</th>
                    <th>Concepto</th>
                    <th>Fuente</th>
                    <th class="text-right">Monto (USD)</th>
                </tr>
            </thead>
            <tbody>
                <tr><td>Transporte</td><td>Paquete Base Terrestre</td><td>Imagen / Promo</td><td class="text-right">${datos['precio_base']:,.2f}</td></tr>
                <tr><td>Transporte</td><td>Impuestos Aéreos</td><td>Imagen / Promo</td><td class="text-right">${datos['impuestos_aereo']:,.2f}</td></tr>
                <tr><td>Transporte</td><td>Equipaje Facturado Extra</td><td>Estimación Manual</td><td class="text-right">${datos['equipaje_extra']:,.2f}</td></tr>
                <tr><td>Alimentación</td><td>Comidas No Incluidas</td><td>Estimación Manual</td><td class="text-right">${datos['comidas_extra']:,.2f}</td></tr>
                <tr><td>Excursiones</td><td>Tours Extra / Entradas</td><td>Estimación Manual</td><td class="text-right">${datos['excursiones_extra']:,.2f}</td></tr>
                <tr><td>Seguro</td><td>Asistencia Médica Internacional</td><td>Estimación Manual</td><td class="text-right">${datos['seguro_medico']:,.2f}</td></tr>
                <tr><td>Varios</td><td>Propinas e Imprevistos</td><td>Estimación Manual</td><td class="text-right">${datos['gastos_personales']:,.2f}</td></tr>
                <tr class="total">
                    <td colspan="3"><strong>TOTAL ESTIMADO POR PERSONA</strong></td>
                    <td class="text-right"><strong>${total_calc:,.2f}</strong></td>
                </tr>
            </tbody>
        </table>
        
        <div class="card" style="margin-top:20px;">
            <p><strong>Notas adicionales:</strong> {datos['notas']}</p>
        </div>
    </body>
    </html>
    """
    output = io.BytesIO()
    pisa.CreatePDF(html_content, dest=output)
    return output.getvalue()

# --- VISTA 1: CREAR PRESUPUESTO ---
if menu == "🧮 Crear Presupuesto":
    st.subheader("1. Carga de Documentos y Fuentes de Información")
    col_left, col_right = st.columns(2)

    with col_left:
        uploaded_image = st.file_uploader("🖼️ Imagen / Flyer / Comprobante (JPG, PNG)", type=["jpg", "png", "jpeg"])
        web_url = st.text_input("🌐 Enlace Web (opcional)")

    with col_right:
        email_text = st.text_area("📧 Texto de Correo Electrónico / Notas", height=140, placeholder="Pega aquí el contenido del mail de la agencia...")

    if st.button("🔍 Extraer Datos con IA (Groq)"):
        if not api_key:
            st.error("⚠️ Por favor ingresa tu API Key de Groq (gsk_...) en el panel lateral o en Secrets.")
        elif not (uploaded_image or email_text or web_url):
            st.warning("⚠️ Debes proporcionar al menos una fuente de datos.")
        else:
            with st.spinner("Analizando información con IA (Groq)..."):
                try:
                    client = Groq(api_key=api_key)
                    
                    messages_content = []
                    
                    prompt_text = f"""
                    Analiza la información de viaje y responde ÚNICAMENTE con un objeto JSON sin formato Markdown adicional (sin ```json ...
