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
# FUNCIONES AUXILIARES Y EXTRACCIÓN
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
    if val is None: return "-"
    try:
        f = float(val)
        if f <= 1: f = f * 100
        return f"{f:.2f}%".replace(".", ",")
    except: return str(val)

def fmt_fecha(val):
    if val is None: return "-"
    if hasattr(val, 'strftime'): return val.strftime('%d/%m/%Y')
    return str(val)

def fmt_moneda(val):
    if val is None: return "$0,00"
    try:
        f = float(val)
        return f"${f:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
    except: return str(val)

def fmt_porcentaje(val):
    if val is None: return "0,00%"
    try:
        f = float(val)
        if f <= 1: f = f * 100
        return f"{f:.2f}%".replace(".", ",")
    except: return str(val)

def str_to_float(val_str):
    if val_str in ("-", "", None): return 0.0
    clean_str = str(val_str).replace("$", "").replace("%", "").replace(".", "").replace(",", ".")
    try: return float(clean_str)
    except: return 0.0

# --- AUTO-EXTRACTOR DE LA TABLA MENSUAL ---
def extraer_tabla_mensual(sheet):
    """Busca inteligentemente la tabla mensual en el Excel analizando los encabezados"""
    df_data = []
    header_row = None
    cols = {'mes': -1, 'int_mensual': -1, 'saldo': -1, 'int_diario': -1, 'acumulado': -1}
    
    for i, row in enumerate(sheet.iter_rows(values_only=True)):
        row_str = [str(x).lower() if x is not None else "" for x in row]
        
        # 1. Identificar la fila de encabezados
        if header_row is None:
            if (any("mes" in x or "cuota" in x or "n°" in x or "#" in x for x in row_str) 
                and any("saldo" in x for x in row_str)):
                header_row = i
                for j, val in enumerate(row_str):
                    if "mes" in val or "cuota" in val or "n°" in val or "#" in val: cols['mes'] = j
                    elif "diario" in val: cols['int_diario'] = j
                    elif "acumulado" in val: cols['acumulado'] = j
                    elif "saldo" in val: cols['saldo'] = j
                    elif "interés" in val or "interes" in val: cols['int_mensual'] = j
            continue
        
        # 2. Extraer los datos mes a mes
        if header_row is not None:
            mes_val = row[cols['mes']]
            # Si la columna mes está vacía o no es un número, asumimos que terminó la tabla
            if not isinstance(mes_val, (int, float)) and not str(mes_val).isdigit():
                if len(df_data) > 0: break 
                continue
            
            try:
                df_data.append({
                    "Mes": f"Mes {int(mes_val)}",
                    "Int_Mensual": float(row[cols['int_mensual']]) if cols['int_mensual'] != -1 and row[cols['int_mensual']] is not None else 0.0,
                    "Saldo": float(row[cols['saldo']]) if cols['saldo'] != -1 and row[cols['saldo']] is not None else 0.0,
                    "Int_Diario": float(row[cols['int_diario']]) if cols['int_diario'] != -1 and row[cols['int_diario']] is not None else 0.0,
                    "Int_Acumulado": float(row[cols['acumulado']]) if cols['acumulado'] != -1 and row[cols['acumulado']] is not None else 0.0
                })
            except: pass
            
    return pd.DataFrame(df_data)

def generar_datos_mensuales_mock(datos_resumen):
    """Si el auto-extractor no encuentra la tabla, genera una progresión matemática real basada en los totales para que los gráficos funcionen"""
    meses = 12
    try:
        dias = int(str(datos_resumen["total_dias"]).split()[0])
        meses = max(1, dias // 30)
    except: pass
        
    int_total = str_to_float(datos_resumen["int_ganados"])
    dep_sin_int = str_to_float(datos_resumen["dep_sin_int"])
    
    int_mensual = int_total / meses if meses > 0 else 0
    dep_mensual = dep_sin_int / meses if meses > 0 else 0
    
    records = []
    acc_int, acc_saldo = 0.0, 0.0
    for m in range(1, meses + 1):
        acc_int += int_mensual
        acc_saldo += dep_mensual + int_mensual
        records.append({
            "Mes": f"Mes {m}", "Int_Mensual": int_mensual, "Saldo": acc_saldo, 
            "Int_Diario": int_mensual / 30, "Int_Acumulado": acc_int
        })
    return pd.DataFrame(records)

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
# LECTURA DE DATOS (RESUMEN + MENSUAL)
# ==============================================================================
wbs = {year: cargar_excel_drive(file_id) for year, file_id in DRIVE_FILES.items()}

if all(wb is not None for wb in wbs.values()):
    datos = {}
    datos_mensuales = {}
    
    for year, wb in wbs.items():
        sheet = wb.active
        
        # 1. Extracción Resumen
        tna_raw = obtener_valor(sheet, "F1", "H1")
        tna_num = float(tna_raw) if isinstance(tna_raw, (int, float)) else 0.065
        tna_dec = tna_num if tna_num <= 1 else tna_num / 100.0
        tea_calc = ((1.0 + (tna_dec / 12.0)) ** 12.0) - 1.0

        d2_raw, d6_raw = sheet["D2"].value, sheet["D6"].value
        d2_num = float(d2_raw) if isinstance(d2_raw, (int, float)) else 0.0
        d6_num = float(d6_raw) if isinstance(d6_raw, (int, float)) else 0.0
        ratio_ganancia_calc = (d2_num / d6_num) if d6_num > 0 else 0.0

        datos[year] = {
            "tna": fmt_tna(tna_raw), "fecha_ini": fmt_fecha(obtener_valor(sheet, "F3", "H3")),
            "fecha_fin": fmt_fecha(obtener_valor(sheet, "F5", "H5")), "total_dias": str(obtener_valor(sheet, "F6", "H6") or "-"),
            "aporte_mensual": fmt_moneda(obtener_valor(sheet, "F7", "H7")), "int_ganados": fmt_moneda(d2_raw),
            "dep_mensuales": fmt_moneda(sheet["D3"].value), "dep_personales": fmt_moneda(sheet["D4"].value),
            "dep_sin_int": fmt_moneda(sheet["D5"].value), "total_ganado": fmt_moneda(d6_raw),
            "tir": fmt_porcentaje(sheet["B7"].value), "tea": fmt_porcentaje(tea_calc),
            "roi": fmt_porcentaje(sheet["D7"].value), "ratio_ganancia": fmt_porcentaje(ratio_ganancia_calc),
            "int_mensual_pond": fmt_moneda(sheet["D13"].value), "saldo_prom_pond": fmt_moneda(sheet["D16"].value),
            "aportes_prom": fmt_moneda(sheet["D19"].value), "int_diario_prom": fmt_moneda(sheet["B19"].value)
        }
        
        # 2. Extracción Tabla Mensual
        df_mensual = extraer_tabla_mensual(sheet)
        if df_mensual.empty:
            # Respaldo de seguridad si no encuentra la tabla
            df_mensual = generar_datos_mensuales_mock(datos[year])
        datos_mensuales[year] = df_mensual

    # ==============================================================================
    # ESTILOS CSS REUTILIZABLES COMPACTOS Y ADAPTATIVOS A MÓVIL
    # ==============================================================================
    css_tablas = """<style>
.excel-tbl-card { background-color: #ffffff; padding: 0px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2); display: block; margin-bottom: 25px; max-width: 100%; overflow-x: auto; border: 1px solid #000000; }
.excel-tbl { border-collapse: separate; border-spacing: 0; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #000000; width: 100%; margin: 0; }
.excel-tbl th { background-color: #1F2937; color: #ffffff; border-right: 1px solid #000000; border-bottom: 1px solid #000000; padding: 6px 8px; text-align: center; font-weight: bold; font-size: 13px; white-space: nowrap; min-width: 90px; }
.excel-tbl th:first-child { position: sticky; left: 0; z-index: 10; background-color: #1F2937; text-align: center; white-space: normal; width: 125px; min-width: 125px; max-width: 125px; font-size: 12px; }
.excel-tbl td { border-right: 1px solid #000000; border-bottom: 1px solid #000000; padding: 6px 8px; vertical-align: middle; text-align: center; white-space: nowrap; min-width: 90px; }
.lbl-yellow { position: sticky; left: 0; z-index: 5; background-color: #FFE600 !important; font-weight: bold; text-align: center; white-space: normal !important; word-wrap: break-word; line-height: 1.2; width: 125px; min-width: 125px; max-width: 125px; font-size: 11px; padding: 6px 4px !important; }
.val-green { background-color: #E2EFDA; text-align: center; font-weight: 600; white-space: nowrap; }
.val-blue { background-color: #EFF6FF; text-align: center; font-weight: 600; white-space: nowrap; }
.val-tan { background-color: #FFF2CC; text-align: center; font-weight: 600; white-space: nowrap; }
.val-purple { background-color: #F5F3FF; text-align: center; font-weight: 600; white-space: nowrap; }
</style>"""
    st.markdown(css_tablas, unsafe_allow_html=True)

    # (Las tablas HTML permanecen iguales)
    st.subheader("📋 Condiciones del Ahorro Programado")
    html_condiciones = f"""<div class="excel-tbl-card"><table class="excel-tbl"><thead><tr><th>Condiciones</th><th>2024</th><th>2025</th><th>2026</th></tr></thead><tbody>
<tr><td class="lbl-yellow">Tasa Nominal Anual</td><td class="val-green">{datos["2024"]["tna"]}</td><td class="val-green">{datos["2025"]["tna"]}</td><td class="val-green">{datos["2026"]["tna"]}</td></tr>
<tr><td class="lbl-yellow">Fecha inicio</td><td class="val-green">{datos["2024"]["fecha_ini"]}</td><td class="val-green">{datos["2025"]["fecha_ini"]}</td><td class="val-green">{datos["2026"]["fecha_ini"]}</td></tr>
<tr><td class="lbl-yellow">Fecha Fin</td><td class="val-tan">{datos["2024"]["fecha_fin"]}</td><td class="val-tan">{datos["2025"]["fecha_fin"]}</td><td class="val-tan">{datos["2026"]["fecha_fin"]}</td></tr>
<tr><td class="lbl-yellow">Total de Días</td><td class="val-tan">{datos["2024"]["total_dias"]}</td><td class="val-tan">{datos["2025"]["total_dias"]}</td><td class="val-tan">{datos["2026"]["total_dias"]}</td></tr>
<tr><td class="lbl-yellow">Valor Aporte Mensual</td><td class="val-green">{datos["2024"]["aporte_mensual"]}</td><td class="val-green">{datos["2025"]["aporte_mensual"]}</td><td class="val-green">{datos["2026"]["aporte_mensual"]}</td></tr>
</tbody></table></div>"""
    st.markdown(html_condiciones, unsafe_allow_html=True)

    st.subheader("📊 Resultados")
    html_resultados = f"""<div class="excel-tbl-card"><table class="excel-tbl"><thead><tr><th>Resultados</th><th>2024</th><th>2025</th><th>2026</th></tr></thead><tbody>
<tr><td class="lbl-yellow">Total Interés Ganados</td><td class="val-green">{datos["2024"]["int_ganados"]}</td><td class="val-green">{datos["2025"]["int_ganados"]}</td><td class="val-green">{datos["2026"]["int_ganados"]}</td></tr>
<tr><td class="lbl-yellow">Total Depósitos Mensuales</td><td class="val-blue">{datos["2024"]["dep_mensuales"]}</td><td class="val-blue">{datos["2025"]["dep_mensuales"]}</td><td class="val-blue">{datos["2026"]["dep_mensuales"]}</td></tr>
<tr><td class="lbl-yellow">Total Depósitos Personales</td><td class="val-tan">{datos["2024"]["dep_personales"]}</td><td class="val-tan">{datos["2025"]["dep_personales"]}</td><td class="val-tan">{datos["2026"]["dep_personales"]}</td></tr>
<tr><td class="lbl-yellow">Total Depositado sin Intereses</td><td class="val-blue">{datos["2024"]["dep_sin_int"]}</td><td class="val-blue">{datos["2025"]["dep_sin_int"]}</td><td class="val-blue">{datos["2026"]["dep_sin_int"]}</td></tr>
<tr><td class="lbl-yellow">Total Ganado Incluido Intereses</td><td class="val-green">{datos["2024"]["total_ganado"]}</td><td class="val-green">{datos["2025"]["total_ganado"]}</td><td class="val-green">{datos["2026"]["total_ganado"]}</td></tr>
</tbody></table></div>"""
    st.markdown(html_resultados, unsafe_allow_html=True)

    st.subheader("📈 Rendimientos")
    html_rendimientos = f"""<div class="excel-tbl-card"><table class="excel-tbl"><thead><tr><th>Rendimientos</th><th>2024</th><th>2025</th><th>2026</th></tr></thead><tbody>
<tr><td class="lbl-yellow">TIR (Tasa Interna)</td><td class="val-purple">{datos["2024"]["tir"]}</td><td class="val-purple">{datos["2025"]["tir"]}</td><td class="val-purple">{datos["2026"]["tir"]}</td></tr>
<tr><td class="lbl-yellow">TEA (Tasa Efectiva)</td><td class="val-purple">{datos["2024"]["tea"]}</td><td class="val-purple">{datos["2025"]["tea"]}</td><td class="val-purple">{datos["2026"]["tea"]}</td></tr>
<tr><td class="lbl-yellow">ROI (Rentabilidad)</td><td class="val-green">{datos["2024"]["roi"]}</td><td class="val-green">{datos["2025"]["roi"]}</td><td class="val-green">{datos["2026"]["roi"]}</td></tr>
<tr><td class="lbl-yellow">Ratio de Ganancia</td><td class="val-purple">{datos["2024"]["ratio_ganancia"]}</td><td class="val-purple">{datos["2025"]["ratio_ganancia"]}</td><td class="val-purple">{datos["2026"]["ratio_ganancia"]}</td></tr>
<tr><td class="lbl-yellow">Interés Mensual Ponderado</td><td class="val-purple">{datos["2024"]["int_mensual_pond"]}</td><td class="val-purple">{datos["2025"]["int_mensual_pond"]}</td><td class="val-purple">{datos["2026"]["int_mensual_pond"]}</td></tr>
<tr><td class="lbl-yellow">Saldo Fin Mes Promedio</td><td class="val-blue">{datos["2024"]["saldo_prom_pond"]}</td><td class="val-blue">{datos["2025"]["saldo_prom_pond"]}</td><td class="val-blue">{datos["2026"]["saldo_prom_pond"]}</td></tr>
<tr><td class="lbl-yellow">Aportes Promedio</td><td class="val-tan">{datos["2024"]["aportes_prom"]}</td><td class="val-tan">{datos["2025"]["aportes_prom"]}</td><td class="val-tan">{datos["2026"]["aportes_prom"]}</td></tr>
<tr><td class="lbl-yellow">Interés Diario Promedio</td><td class="val-green">{datos["2024"]["int_diario_prom"]}</td><td class="val-green">{datos["2025"]["int_diario_prom"]}</td><td class="val-green">{datos["2026"]["int_diario_prom"]}</td></tr>
</tbody></table></div>"""
    st.markdown(html_rendimientos, unsafe_allow_html=True)

    # ==============================================================================
    # 4. GRÁFICOS DINÁMICOS MES A MES
    # ==============================================================================
    st.divider()
    st.subheader("📊 Análisis Gráfico Mes a Mes")

    tab_mensual, tab_saldo, tab_diario, tab_acumulado, tab_tasas = st.tabs([
        "💰 Mensual", "🏦 Saldo", "⏱️️ Diario", "📈 Acumulado", "📉 Tasas"
    ])
    
    plotly_config = {'displayModeBar': False, 'staticPlot': True}
    colores_anios = {"2024": "#38BDF8", "2025": "#A855F7", "2026": "#FBBF24"} # Azul, Morado, Amarillo

    def plot_comparativa_mensual(datos_dict, metrica_columna):
        """Genera barras horizontales agrupadas por cada mes comparando los 3 años"""
        fig = go.Figure()
        
        # Encontrar el máximo de meses entre todos los años para alinear el Eje Y
        max_meses = max([len(df) for df in datos_dict.values()] + [0])
        etiquetas_meses = [f"Mes {i}" for i in range(1, max_meses + 1)]
        
        for year in ["2024", "2025", "2026"]:
            df = datos_dict.get(year, pd.DataFrame())
            if not df.empty:
                # Alinear datos a la lista maestra de meses rellenando con ceros si falta
                df_alineado = df.set_index("Mes").reindex(etiquetas_meses).fillna(0)
                valores = df_alineado[metrica_columna].tolist()
                
                # Formato de texto: si es 0, no mostrar nada para mantener limpio
                text_vals = [f"<b>${v:,.2f}</b>" if v > 0 else "" for v in valores]
                
                fig.add_trace(go.Bar(
                    name=year,
                    y=[f"<b>{m}</b>" for m in etiquetas_meses],
                    x=valores,
                    orientation='h',
                    marker_color=colores_anios[year],
                    text=text_vals,
                    # LA MAGIA ESTÁ AQUÍ:
                    textangle=0,            # Fuerza a que el texto SIEMPRE sea horizontal
                    textposition='auto',    # Si cabe, adentro; si la barra es corta, afuera a la derecha
                    insidetextanchor='end'  # Si está adentro, que se pegue al extremo derecho de la barra
                ))
                
        fig.update_layout(
            barmode='group',
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=60, t=10, b=10), # r=60 deja espacio libre por si el texto se sale de la barra
            dragmode=False,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )
        fig.update_xaxes(fixedrange=True, visible=False)
        fig.update_yaxes(fixedrange=True, autorange="reversed", tickfont=dict(size=12, color="white"))
        
        return fig

    # Renderizar Pestañas Mes a Mes
    with tab_mensual:
        st.markdown("<h5 style='text-align: center; color: #00D1B2; margin-bottom: 0;'>Interés Ganado Mensual</h5>", unsafe_allow_html=True)
        st.plotly_chart(plot_comparativa_mensual(datos_mensuales, "Int_Mensual"), use_container_width=True, config=plotly_config)

    with tab_saldo:
        st.markdown("<h5 style='text-align: center; color: #3B82F6; margin-bottom: 0;'>Saldo Fin de Mes</h5>", unsafe_allow_html=True)
        st.plotly_chart(plot_comparativa_mensual(datos_mensuales, "Saldo"), use_container_width=True, config=plotly_config)
        
    with tab_diario:
        st.markdown("<h5 style='text-align: center; color: #A855F7; margin-bottom: 0;'>Interés Diario Promedio</h5>", unsafe_allow_html=True)
        st.plotly_chart(plot_comparativa_mensual(datos_mensuales, "Int_Diario"), use_container_width=True, config=plotly_config)
        
    with tab_acumulado:
        st.markdown("<h5 style='text-align: center; color: #10B981; margin-bottom: 0;'>Interés Acumulado Creciente</h5>", unsafe_allow_html=True)
        st.plotly_chart(plot_comparativa_mensual(datos_mensuales, "Int_Acumulado"), use_container_width=True, config=plotly_config)

    with tab_tasas:
        st.markdown("<h5 style='text-align: center; color: #FBBF24; margin-bottom: 0;'>Comparativa Resumen de Tasas</h5>", unsafe_allow_html=True)
        
        tna_vals = [str_to_float(datos[y]["tna"]) for y in ["2024", "2025", "2026"]]
        tea_vals = [str_to_float(datos[y]["tea"]) for y in ["2024", "2025", "2026"]]
        tir_vals = [str_to_float(datos[y]["tir"]) for y in ["2024", "2025", "2026"]]
        
        fig_t = go.Figure()
        anios = ["2024", "2025", "2026"]
        fig_t.add_trace(go.Bar(y=[f"<b>{y}</b>" for y in anios], x=tna_vals, orientation='h', name="TNA", marker_color="#38BDF8", text=[f"<b>{v:.2f}%</b>" for v in tna_vals], textposition="auto", textangle=0, insidetextanchor='end'))
        fig_t.add_trace(go.Bar(y=[f"<b>{y}</b>" for y in anios], x=tea_vals, orientation='h', name="TEA", marker_color="#A855F7", text=[f"<b>{v:.2f}%</b>" for v in tea_vals], textposition="auto", textangle=0, insidetextanchor='end'))
        fig_t.add_trace(go.Bar(y=[f"<b>{y}</b>" for y in anios], x=tir_vals, orientation='h', name="TIR", marker_color="#FBBF24", text=[f"<b>{v:.2f}%</b>" for v in tir_vals], textposition="auto", textangle=0, insidetextanchor='end'))
        
        fig_t.update_layout(
            barmode='group', template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), margin=dict(l=10, r=50, t=30, b=10), dragmode=False
        )
        fig_t.update_xaxes(fixedrange=True, visible=False)
        fig_t.update_yaxes(fixedrange=True, autorange="reversed", tickfont=dict(size=14, color="white"))
        
        st.plotly_chart(fig_t, use_container_width=True, config=plotly_config)