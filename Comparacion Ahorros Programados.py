import sys
import io
import requests
import openpyxl
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ==============================================================================
# AUTORUNNER: FUNCIONA TANTO LOCALMENTE COMO EN STREAMLIT CLOUD
# ==============================================================================
if __name__ == "__main__":
    if not st.runtime.exists():
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Comparación Ahorros Programados",
    page_icon="💰",
    layout="wide"
)

# ARCHIVOS DE GOOGLE DRIVE POR AÑO
DRIVE_FILES = {
    "2024": "1ehwdQzNEeu3jhQWTlPhcezUMfCHtNute",
    "2025": "1ke9f5XiX4b5DSuXNYXTNr9LNPUO5bUzz",
    "2026": "1s8H4eqdO3LyIFwPdjGZG0TwHCRZmtnsw"
}

# ==============================================================================
# CARGAR DESDE GOOGLE DRIVE (CON CACHÉ REUTILIZABLE)
# ==============================================================================
@st.cache_data(ttl=300)
def cargar_excel_drive(file_id):
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    session = requests.Session()
    response = session.get(url)
    
    if response.status_code != 200:
        st.error(f"No se pudo descargar el archivo con ID {file_id} desde Google Drive.")
        return None

    excel_bytes = io.BytesIO(response.content)
    wb = openpyxl.load_workbook(excel_bytes, data_only=True)
    return wb

def obtener_valor(sheet, celda_principal, celda_secundaria=None):
    val = sheet[celda_principal].value
    if val is None and celda_secundaria:
        val = sheet[celda_secundaria].value
    return val

def fmt_tna(val):
    if val is None:
        return "-"
    try:
        f = float(val)
        if f <= 1:
            f = f * 100
        return f"{f:.2f}%".replace(".", ",")
    except:
        return str(val)

def fmt_fecha(val):
    if val is None:
        return "-"
    if hasattr(val, 'strftime'):
        return val.strftime('%d/%m/%Y')
    return str(val)

def fmt_moneda(val):
    if val is None:
        return "$0,00"
    try:
        f = float(val)
        return f"${f:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
    except:
        return str(val)

def fmt_porcentaje(val):
    if val is None:
        return "0,00%"
    try:
        f = float(val)
        if f <= 1:
            f = f * 100
        return f"{f:.2f}%".replace(".", ",")
    except:
        return str(val)

# ==============================================================================
# ENCABEZADO PRINCIPAL Y BOTÓN RECARGAR
# ==============================================================================
col_title, col_btn = st.columns([4, 1])
with col_title:
    st.title("💰 Comparación Ahorros Programados")
    st.caption("Sincronizado en tiempo real desde Google Drive")

with col_btn:
    if st.button("🔄 Recargar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.divider()

# ==============================================================================
# LECTURA DE LIBROS Y EXTRACCIÓN DE PARÁMETROS
# ==============================================================================
wbs = {}
for year, file_id in DRIVE_FILES.items():
    wbs[year] = cargar_excel_drive(file_id)

if all(wb is not None for wb in wbs.values()):
    datos = {}
    for year, wb in wbs.items():
        sheet = wb.active
        
        # Extracción y cálculos de parámetros
        tna_raw = obtener_valor(sheet, "F1", "H1")
        tna_num = float(tna_raw) if isinstance(tna_raw, (int, float)) else 0.065
        tna_dec = tna_num if tna_num <= 1 else tna_num / 100.0
        tea_calc = ((1.0 + (tna_dec / 12.0)) ** 12.0) - 1.0

        d2_raw = sheet["D2"].value # Total Interés Ganados
        d6_raw = sheet["D6"].value # Total Ganado Incluido Intereses
        d2_num = float(d2_raw) if isinstance(d2_raw, (int, float)) else 0.0
        d6_num = float(d6_raw) if isinstance(d6_raw, (int, float)) else 0.0
        ratio_ganancia_calc = (d2_num / d6_num) if d6_num > 0 else 0.0

        datos[year] = {
            # Sección 1: Condiciones
            "tna": fmt_tna(tna_raw),
            "fecha_ini": fmt_fecha(obtener_valor(sheet, "F3", "H3")),
            "fecha_fin": fmt_fecha(obtener_valor(sheet, "F5", "H5")),
            "total_dias": str(obtener_valor(sheet, "F6", "H6") or "-"),
            "aporte_mensual": fmt_moneda(obtener_valor(sheet, "F7", "H7")),
            
            # Sección 2: Resultados
            "int_ganados": fmt_moneda(d2_raw),
            "dep_mensuales": fmt_moneda(sheet["D3"].value),
            "dep_personales": fmt_moneda(sheet["D4"].value),
            "dep_sin_int": fmt_moneda(sheet["D5"].value),
            "total_ganado": fmt_moneda(d6_raw),
            
            # Sección 3: Rendimientos
            "tir": fmt_porcentaje(sheet["B7"].value),
            "tea": fmt_porcentaje(tea_calc),
            "roi": fmt_porcentaje(sheet["D7"].value),
            "ratio_ganancia": fmt_porcentaje(ratio_ganancia_calc),
            "int_mensual_pond": fmt_moneda(sheet["D13"].value),
            "saldo_prom_pond": fmt_moneda(sheet["D16"].value),
            "aportes_prom": fmt_moneda(sheet["D19"].value),
            "int_diario_prom": fmt_moneda(sheet["B19"].value)
        }

    # ==============================================================================
    # ESTILOS CSS REUTILIZABLES COMPACTOS Y ADAPTATIVOS A MÓVIL
    # ==============================================================================
    css_tablas = """<style>
.excel-tbl-card {
    background-color: #ffffff;
    padding: 0px;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
    display: block;
    margin-bottom: 25px;
    max-width: 100%;
    overflow-x: auto;
    border: 1px solid #000000;
}
.excel-tbl {
    border-collapse: separate;
    border-spacing: 0;
    font-family: Calibri, 'Segoe UI', Arial, sans-serif;
    font-size: 12px;
    color: #000000;
    width: 100%;
    margin: 0;
}
.excel-tbl th {
    background-color: #1F2937;
    color: #ffffff;
    border-right: 1px solid #000000;
    border-bottom: 1px solid #000000;
    padding: 6px 8px;
    text-align: center;
    font-weight: bold;
    font-size: 13px;
    white-space: nowrap;
    min-width: 90px;
}
.excel-tbl th:first-child {
    position: sticky;
    left: 0;
    z-index: 10;
    background-color: #1F2937;
    text-align: center;
    white-space: normal;
    width: 125px;
    min-width: 125px;
    max-width: 125px;
    font-size: 12px;
}
.excel-tbl td {
    border-right: 1px solid #000000;
    border-bottom: 1px solid #000000;
    padding: 6px 8px;
    vertical-align: middle;
    text-align: center;
    white-space: nowrap;
    min-width: 90px;
}
.lbl-yellow {
    position: sticky;
    left: 0;
    z-index: 5;
    background-color: #FFE600 !important;
    font-weight: bold;
    text-align: center;
    white-space: normal !important;
    word-wrap: break-word;
    line-height: 1.2;
    width: 125px;
    min-width: 125px;
    max-width: 125px;
    font-size: 11px;
    padding: 6px 4px !important;
}
.val-green {
    background-color: #E2EFDA;
    text-align: center;
    font-weight: 600;
    white-space: nowrap;
}
.val-blue {
    background-color: #EFF6FF;
    text-align: center;
    font-weight: 600;
    white-space: nowrap;
}
.val-tan {
    background-color: #FFF2CC;
    text-align: center;
    font-weight: 600;
    white-space: nowrap;
}
.val-purple {
    background-color: #F5F3FF;
    text-align: center;
    font-weight: 600;
    white-space: nowrap;
}
</style>"""

    st.markdown(css_tablas, unsafe_allow_html=True)

    # ==============================================================================
    # 1. TABLA COMPARATIVA: CONDICIONES DEL AHORRO
    # ==============================================================================
    st.subheader("📋 Condiciones del Ahorro Programado")

    html_condiciones = f"""<div class="excel-tbl-card">
<table class="excel-tbl">
<thead>
<tr>
<th>Condiciones</th>
<th>2024</th>
<th>2025</th>
<th>2026</th>
</tr>
</thead>
<tbody>
<tr>
<td class="lbl-yellow">Tasa Nominal Anual</td>
<td class="val-green">{datos["2024"]["tna"]}</td>
<td class="val-green">{datos["2025"]["tna"]}</td>
<td class="val-green">{datos["2026"]["tna"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Fecha inicio</td>
<td class="val-green">{datos["2024"]["fecha_ini"]}</td>
<td class="val-green">{datos["2025"]["fecha_ini"]}</td>
<td class="val-green">{datos["2026"]["fecha_ini"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Fecha Fin</td>
<td class="val-tan">{datos["2024"]["fecha_fin"]}</td>
<td class="val-tan">{datos["2025"]["fecha_fin"]}</td>
<td class="val-tan">{datos["2026"]["fecha_fin"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Total de Días</td>
<td class="val-tan">{datos["2024"]["total_dias"]}</td>
<td class="val-tan">{datos["2025"]["total_dias"]}</td>
<td class="val-tan">{datos["2026"]["total_dias"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Valor Aporte Mensual</td>
<td class="val-green">{datos["2024"]["aporte_mensual"]}</td>
<td class="val-green">{datos["2025"]["aporte_mensual"]}</td>
<td class="val-green">{datos["2026"]["aporte_mensual"]}</td>
</tr>
</tbody>
</table>
</div>"""

    st.markdown(html_condiciones, unsafe_allow_html=True)

    st.divider()

    # ==============================================================================
    # 2. TABLA COMPARATIVA: RESULTADOS
    # ==============================================================================
    st.subheader("📊 Resultados")

    html_resultados = f"""<div class="excel-tbl-card">
<table class="excel-tbl">
<thead>
<tr>
<th>Resultados</th>
<th>2024</th>
<th>2025</th>
<th>2026</th>
</tr>
</thead>
<tbody>
<tr>
<td class="lbl-yellow">Total Interés Ganados</td>
<td class="val-green">{datos["2024"]["int_ganados"]}</td>
<td class="val-green">{datos["2025"]["int_ganados"]}</td>
<td class="val-green">{datos["2026"]["int_ganados"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Total Depósitos Mensuales</td>
<td class="val-blue">{datos["2024"]["dep_mensuales"]}</td>
<td class="val-blue">{datos["2025"]["dep_mensuales"]}</td>
<td class="val-blue">{datos["2026"]["dep_mensuales"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Total Depósitos Personales</td>
<td class="val-tan">{datos["2024"]["dep_personales"]}</td>
<td class="val-tan">{datos["2025"]["dep_personales"]}</td>
<td class="val-tan">{datos["2026"]["dep_personales"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Total Depositado sin Intereses</td>
<td class="val-blue">{datos["2024"]["dep_sin_int"]}</td>
<td class="val-blue">{datos["2025"]["dep_sin_int"]}</td>
<td class="val-blue">{datos["2026"]["dep_sin_int"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Total Ganado Incluido Intereses</td>
<td class="val-green">{datos["2024"]["total_ganado"]}</td>
<td class="val-green">{datos["2025"]["total_ganado"]}</td>
<td class="val-green">{datos["2026"]["total_ganado"]}</td>
</tr>
</tbody>
</table>
</div>"""

    st.markdown(html_resultados, unsafe_allow_html=True)

    st.divider()

    # ==============================================================================
    # 3. TABLA COMPARATIVA: RENDIMIENTOS
    # ==============================================================================
    st.subheader("📈 Rendimientos")

    html_rendimientos = f"""<div class="excel-tbl-card">
<table class="excel-tbl">
<thead>
<tr>
<th>Rendimientos</th>
<th>2024</th>
<th>2025</th>
<th>2026</th>
</tr>
</thead>
<tbody>
<tr>
<td class="lbl-yellow">TIR (Tasa Interna de Retorno)</td>
<td class="val-purple">{datos["2024"]["tir"]}</td>
<td class="val-purple">{datos["2025"]["tir"]}</td>
<td class="val-purple">{datos["2026"]["tir"]}</td>
</tr>
<tr>
<td class="lbl-yellow">TEA (Tasa Efectiva Anual)</td>
<td class="val-purple">{datos["2024"]["tea"]}</td>
<td class="val-purple">{datos["2025"]["tea"]}</td>
<td class="val-purple">{datos["2026"]["tea"]}</td>
</tr>
<tr>
<td class="lbl-yellow">ROI (Rentabilidad Acumulada)</td>
<td class="val-green">{datos["2024"]["roi"]}</td>
<td class="val-green">{datos["2025"]["roi"]}</td>
<td class="val-green">{datos["2026"]["roi"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Ratio de Ganancia</td>
<td class="val-purple">{datos["2024"]["ratio_ganancia"]}</td>
<td class="val-purple">{datos["2025"]["ratio_ganancia"]}</td>
<td class="val-purple">{datos["2026"]["ratio_ganancia"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Interés Mensual Ponderado</td>
<td class="val-purple">{datos["2024"]["int_mensual_pond"]}</td>
<td class="val-purple">{datos["2025"]["int_mensual_pond"]}</td>
<td class="val-purple">{datos["2026"]["int_mensual_pond"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Saldo Fin de Mes Promedio Ponderado</td>
<td class="val-blue">{datos["2024"]["saldo_prom_pond"]}</td>
<td class="val-blue">{datos["2025"]["saldo_prom_pond"]}</td>
<td class="val-blue">{datos["2026"]["saldo_prom_pond"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Aportes Personales Promedio</td>
<td class="val-tan">{datos["2024"]["aportes_prom"]}</td>
<td class="val-tan">{datos["2025"]["aportes_prom"]}</td>
<td class="val-tan">{datos["2026"]["aportes_prom"]}</td>
</tr>
<tr>
<td class="lbl-yellow">Interés Diario Promedio</td>
<td class="val-green">{datos["2024"]["int_diario_prom"]}</td>
<td class="val-green">{datos["2025"]["int_diario_prom"]}</td>
<td class="val-green">{datos["2026"]["int_diario_prom"]}</td>
</tr>
</tbody>
</table>
</div>"""

    st.markdown(html_rendimientos, unsafe_allow_html=True)


    # ==============================================================================
    # 4. GRÁFICOS DINÁMICOS
    # ==============================================================================
    st.divider()
    st.subheader("📊 Análisis Gráfico de Ahorros")

    # Función auxiliar para convertir las cadenas formateadas a números puros para graficar
    def str_to_float(val_str):
        if val_str in ("-", "", None): return 0.0
        clean_str = str(val_str).replace("$", "").replace("%", "").replace(".", "").replace(",", ".")
        try:
            return float(clean_str)
        except:
            return 0.0

    years = list(datos.keys())
    
    cap_depositado = [str_to_float(datos[y]["dep_sin_int"]) for y in years]
    int_ganado = [str_to_float(datos[y]["int_ganados"]) for y in years]
    
    tna_vals = [str_to_float(datos[y]["tna"]) for y in years]
    tea_vals = [str_to_float(datos[y]["tea"]) for y in years]
    tir_vals = [str_to_float(datos[y]["tir"]) for y in years]
    
    int_mensual_pond = [str_to_float(datos[y]["int_mensual_pond"]) for y in years]
    int_diario_prom = [str_to_float(datos[y]["int_diario_prom"]) for y in years]

    tab_comp, tab_tasas, tab_int = st.tabs(["💰 Capital vs Intereses", "📉 Comparativa de Tasas", "📈 Interés Promedio"])
    
    # CONFIGURACIÓN ESTRICTA PARA MÓVIL: Bloquea todo tipo de zoom, paneo y toque en la pantalla
    plotly_config = {
        'displayModeBar': False,
        'staticPlot': True
    }

    with tab_comp:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=years, y=cap_depositado, name="Capital Depositado", marker_color="#1E3A8A", text=[f"${v:,.0f}" for v in cap_depositado], textposition="inside"))
        fig1.add_trace(go.Bar(x=years, y=int_ganado, name="Intereses Ganados", marker_color="#10B981", text=[f"${v:,.0f}" for v in int_ganado], textposition="inside"))
        
        fig1.update_layout(
            barmode='stack',
            title="Composición del Ahorro Final",
            template="plotly_dark",
            yaxis_tickprefix="$",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        fig1.update_xaxes(fixedrange=True)
        fig1.update_yaxes(fixedrange=True)
        st.plotly_chart(fig1, use_container_width=True, config=plotly_config)

    with tab_tasas:
        fig2 = go.Figure()
        # Colores completamente distintos para que no se repitan visualmente
        fig2.add_trace(go.Bar(x=years, y=tna_vals, name="TNA", marker_color="#38BDF8", text=[f"{v:.2f}%" for v in tna_vals], textposition="auto"))
        fig2.add_trace(go.Bar(x=years, y=tea_vals, name="TEA", marker_color="#F43F5E", text=[f"{v:.2f}%" for v in tea_vals], textposition="auto"))
        fig2.add_trace(go.Bar(x=years, y=tir_vals, name="TIR", marker_color="#FBBF24", text=[f"{v:.2f}%" for v in tir_vals], textposition="auto"))
        
        fig2.update_layout(
            barmode='group',
            title="Evolución de Tasas de Rendimiento",
            template="plotly_dark",
            yaxis_ticksuffix="%",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        fig2.update_xaxes(fixedrange=True)
        fig2.update_yaxes(fixedrange=True)
        st.plotly_chart(fig2, use_container_width=True, config=plotly_config)

    with tab_int:
        # Usamos doble eje Y para que la línea de interés diario (pequeña) no se aplaste contra la mensual
        fig3 = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig3.add_trace(go.Scatter(
            x=years, y=int_mensual_pond, mode='lines+markers+text', name="Int. Mensual", 
            line=dict(color="#00D1B2", width=4),
            marker=dict(size=12, color="#00D1B2"),
            text=[f"${v:,.2f}" for v in int_mensual_pond], 
            textposition="top center",
            textfont=dict(size=13, color="#00D1B2")
        ), secondary_y=False)

        fig3.add_trace(go.Scatter(
            x=years, y=int_diario_prom, mode='lines+markers+text', name="Int. Diario", 
            line=dict(color="#A855F7", width=4, dash='dot'),
            marker=dict(size=12, color="#A855F7"),
            text=[f"${v:,.2f}" for v in int_diario_prom], 
            textposition="bottom center",
            textfont=dict(size=13, color="#A855F7")
        ), secondary_y=True)
        
        fig3.update_layout(
            title="Interés Mensual vs Diario",
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        fig3.update_xaxes(fixedrange=True, showgrid=False)
        fig3.update_yaxes(fixedrange=True, secondary_y=False, showgrid=False, range=[0, max(int_mensual_pond) * 1.3])
        fig3.update_yaxes(fixedrange=True, secondary_y=True, showgrid=False, range=[0, max(int_diario_prom) * 1.3], showticklabels=False)

        st.plotly_chart(fig3, use_container_width=True, config=plotly_config)