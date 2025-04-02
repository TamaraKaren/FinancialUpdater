import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go # Para gráficos más complejos (ej. con indicadores)
import numpy as np
from pathlib import Path
import datetime

# --- Configuración de la Página ---
st.set_page_config(layout="wide", page_title="Dashboard Financiero Dinámico")

# --- Título del Dashboard ---
st.title("📊 Dashboard Financiero Dinámico")
st.markdown("Visualización de datos financieros actualizados periódicamente.")

# --- Cargar Datos desde Excel ---
# Ruta al archivo Excel generado por financial_updater.py
EXCEL_FILE = Path("./Financial_Data/Dynamic Financial Data.xlsx")

@st.cache_data(ttl=60) # Cachear datos por 60 segundos para balancear frescura y rendimiento
def load_financial_data(file_path):
    """Carga datos desde el archivo Excel generado, manejando hojas faltantes."""
    all_data = {}
    try:
        # Leer TODAS las hojas existentes en el archivo
        excel_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl') # sheet_name=None lee todas

        # Asignar DataFrames si la hoja existe, si no, asignar uno vacío
        all_data['summary'] = excel_sheets.get("Resumen General", pd.DataFrame())
        all_data['history'] = excel_sheets.get("Historial_Adj_Close", pd.DataFrame())
        all_data['live'] = excel_sheets.get("Live Data", pd.DataFrame()) # No fallará si no existe
        all_data['news'] = excel_sheets.get("Noticias Recientes", pd.DataFrame())

        # --- Procesamiento Post-Carga ---
        # Convertir columna 'Fecha' en historial a datetime si no lo es
        if not all_data['history'].empty and 'Fecha' in all_data['history'].columns:
            all_data['history']['Fecha'] = pd.to_datetime(all_data['history']['Fecha'], errors='coerce')
            # Establecer Fecha como índice para facilitar resampling y gráficos
            # Comprobar si 'Fecha' no es ya el índice antes de establecerlo
            if all_data['history'].index.name != 'Fecha':
                all_data['history'].set_index('Fecha', inplace=True, drop=False) # drop=False para mantener la columna Fecha

        # Convertir Timestamp en live a datetime
        if not all_data['live'].empty and 'Timestamp Live' in all_data['live'].columns:
             all_data['live']['Timestamp Live'] = pd.to_datetime(all_data['live']['Timestamp Live'], errors='coerce')

        # Convertir Fecha en news a datetime
        if not all_data['news'].empty and 'Fecha' in all_data['news'].columns:
             all_data['news']['Fecha'] = pd.to_datetime(all_data['news']['Fecha'], errors='coerce')

        return all_data

    except FileNotFoundError:
        st.error(f"❌ No se encontró el archivo Excel: '{file_path}'. Asegúrate de que `financial_updater.py` se haya ejecutado al menos una vez.")
        return None
    except Exception as e:
        st.error(f"Error inesperado al cargar los datos desde Excel: {e}")
        return None

# --- Cargar y Validar Datos ---
financial_data = load_financial_data(EXCEL_FILE)

if financial_data is None:
    st.stop() # Detener si no se cargaron datos

df_summary = financial_data['summary']
df_history = financial_data['history']
df_live = financial_data['live']
df_news = financial_data['news']

if df_summary.empty or df_history.empty:
    st.warning("⚠️ Faltan datos esenciales (Resumen o Historial). El dashboard puede estar incompleto.")
    # Podríamos detenernos aquí o continuar mostrando lo que haya
    # st.stop()

st.success(f"✔️ Datos cargados desde '{EXCEL_FILE.name}'.")

# --- Sidebar para Filtros y Opciones ---
st.sidebar.header("Filtros y Opciones")

# Selección de Tickers para gráficos
# Usar tickers disponibles en el historial como opciones
available_tickers = df_history.columns.tolist() if not df_history.empty else []
# Excluir la columna 'Fecha' si todavía está presente después de set_index(drop=False)
if 'Fecha' in available_tickers:
    available_tickers.remove('Fecha')

selected_tickers = st.sidebar.multiselect(
    "Selecciona Tickers para Gráfico Histórico:",
    options=available_tickers,
    default=available_tickers[:min(5, len(available_tickers))] # Mostrar los primeros 5 por defecto
)

# --- Resumen de Métricas Clave ---
st.header("🚀 Resumen General")

last_update_summary = "N/A"
if not df_summary.empty and 'Última Actualización Info' in df_summary.columns:
     try:
        last_update_summary = pd.to_datetime(df_summary['Última Actualización Info'], errors='coerce').max()
        last_update_summary = last_update_summary.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(last_update_summary) else "N/A"
     except Exception: last_update_summary = "Error fecha"

num_tickers = len(df_summary['Ticker'].unique()) if not df_summary.empty else 0 # Contar tickers únicos
sp500_change = "N/A"
if not df_summary.empty and '^GSPC' in df_summary['Ticker'].values:
    change_val = df_summary.loc[df_summary['Ticker'] == '^GSPC', 'Cambio Hoy (%)'].iloc[0]
    if pd.notna(change_val) and isinstance(change_val, (int, float)): sp500_change = f"{change_val:.2f}%"

col1, col2, col3 = st.columns(3)
col1.metric("Tickers Monitoreados", f"{num_tickers}")
col2.metric("Última Actualización (Full)", last_update_summary)
col3.metric("S&P 500 (% Hoy)", sp500_change)

# --- Datos en Vivo (si existen) ---
if not df_live.empty:
    st.header("🔴 Datos en Vivo (Última Actualización Rápida)")
    df_live_sorted = df_live.sort_values(by='Timestamp Live', ascending=False).reset_index(drop=True)
    st.dataframe(df_live_sorted, use_container_width=True, height=250) # Limitar altura
    last_live_update = df_live_sorted['Timestamp Live'].iloc[0].strftime('%Y-%m-%d %H:%M:%S') if not df_live_sorted.empty and pd.notna(df_live_sorted['Timestamp Live'].iloc[0]) else "N/A"
    st.caption(f"Última actualización de datos en vivo: {last_live_update}")
else:
    st.info("Hoja 'Live Data' no encontrada o vacía. Espera a la próxima ejecución de `job_frequent_update`.")

# --- Visualizaciones ---
st.header("📈 Gráficos de Rendimiento")

# --- Gráfico Histórico Normalizado (Interactivo) ---
st.subheader("📊 Rendimiento Histórico Normalizado (1 Año)")

if not df_history.empty and selected_tickers:
    # Asegurarse que el índice es Datetime
    if not isinstance(df_history.index, pd.DatetimeIndex):
         st.warning("El índice del historial no es Datetime. Intentando convertir...")
         try:
             df_history.index = pd.to_datetime(df_history.index)
         except Exception as e:
             st.error(f"No se pudo convertir el índice a Datetime: {e}")
             df_history = pd.DataFrame() # Resetear para evitar más errores

    if isinstance(df_history.index, pd.DatetimeIndex):
        # Filtrar historial por tickers seleccionados
        df_hist_selected = df_history[selected_tickers]

        # Normalizar datos seleccionados
        try:
            first_valid_idx = df_hist_selected.apply(pd.Series.first_valid_index)
            # Usar .at para obtener el primer valor válido de forma segura
            df_first_values = pd.Series([df_hist_selected.at[idx, col] if pd.notna(idx) else np.nan for col, idx in first_valid_idx.items()], index=df_hist_selected.columns)

            df_normalized = df_hist_selected.apply(lambda col: (col / df_first_values[col.name]) * 100 if col.name in df_first_values and pd.notna(df_first_values[col.name]) and df_first_values[col.name] != 0 else col, axis=0)
            df_normalized.fillna(method='ffill', inplace=True)

            # Resetear índice para que 'Fecha' sea una columna para Plotly
            df_plot = df_normalized.reset_index()

            # Crear gráfico de línea con Plotly Express
            fig_norm = px.line(df_plot,
                               x='Fecha',
                               y=selected_tickers,
                               title="Rendimiento Normalizado (Inicio Periodo = 100%)",
                               labels={'value': 'Rendimiento (%)', 'variable': 'Ticker', 'Fecha': 'Fecha'},
                               )
            fig_norm.update_layout(legend_title_text='Tickers')
            fig_norm.update_traces(hovertemplate='<b>%{fullData.name}</b><br>Fecha: %{x|%Y-%m-%d}<br>Rendimiento: %{y:.2f}%<extra></extra>')
            st.plotly_chart(fig_norm, use_container_width=True)

        except Exception as e:
            st.error(f"Error al generar el gráfico normalizado: {e}")

elif df_history.empty:
     st.warning("No hay datos históricos disponibles para graficar.")
else:
     st.info("Selecciona al menos un ticker en la barra lateral para ver el gráfico histórico.")

# --- Otros Gráficos (Ejemplos basados en Resumen) ---
st.subheader("🔍 Análisis Adicional (Datos de Resumen)")

viz_col1, viz_col2 = st.columns(2)

with viz_col1:
    st.markdown("##### **% Cambio Hoy**")
    if not df_summary.empty and 'Cambio Hoy (%)' in df_summary.columns:
        try:
            df_summary['Cambio Num'] = pd.to_numeric(df_summary['Cambio Hoy (%)'], errors='coerce')
            df_plot_change = df_summary.dropna(subset=['Cambio Num']).sort_values('Cambio Num', ascending=False)
            fig_change = px.bar(df_plot_change, x='Ticker', y='Cambio Num', title="Cambio Porcentual Hoy", labels={'Ticker': 'Ticker', 'Cambio Num': '% Cambio'}, color='Cambio Num', color_continuous_scale=px.colors.diverging.RdYlGn, text_auto='.2f')
            fig_change.update_layout(yaxis_title="% Cambio")
            st.plotly_chart(fig_change, use_container_width=True)
        except Exception as e: st.error(f"Error gráfico cambio diario: {e}")
    else: st.info("Columna '% Cambio Hoy (%)' no disponible.")

with viz_col2:
    st.markdown("##### **Posición en Rango 52 Semanas**")
    if not df_summary.empty and '% Rango 52 Sem' in df_summary.columns:
         try:
            df_summary['Rango Num'] = pd.to_numeric(df_summary['% Rango 52 Sem'], errors='coerce')
            df_plot_range = df_summary.dropna(subset=['Rango Num']).sort_values('Rango Num', ascending=False)
            fig_range = px.bar(df_plot_range, x='Ticker', y='Rango Num', title="% Rango Anual (0%=Mín, 100%=Máx)", labels={'Ticker': 'Ticker', 'Rango Num': '% Rango 52 Sem'}, color='Rango Num', color_continuous_scale=px.colors.sequential.Viridis, text_auto='.1f')
            fig_range.update_layout(yaxis_title="% Rango 52 Sem")
            st.plotly_chart(fig_range, use_container_width=True)
         except Exception as e: st.error(f"Error gráfico rango 52 sem: {e}")
    else: st.info("Columna '% Rango 52 Sem' no disponible.")

# --- Exploración de Datos Detallada (Expanders) ---
st.header("🔍 Exploración de Datos Detallada")

with st.expander("Ver Tabla Resumen General", expanded=False):
    if not df_summary.empty: st.dataframe(df_summary)
    else: st.info("No hay datos de resumen.")

with st.expander("Ver Tabla Historial (Últimos 10 días)", expanded=False):
     if not df_history.empty: st.dataframe(df_history.tail(10)) # Muestra el DF con su índice 'Fecha'
     else: st.info("No hay datos históricos.")

with st.expander("Ver Noticias Recientes", expanded=False):
     if not df_news.empty: st.dataframe(df_news[['Ticker', 'Título', 'Publicador', 'Fecha', 'Enlace']])
     else: st.info("No hay noticias recientes.")

# --- Información Adicional Sidebar ---
st.sidebar.header("Acerca de")
st.sidebar.info("Dashboard de datos financieros recopilados por `financial_updater.py`.")
st.sidebar.header("Fuente de Datos")
st.sidebar.markdown(f"Datos desde: `{EXCEL_FILE.name}`")
st.sidebar.markdown("---")
if st.sidebar.button("Recargar Datos"):
    st.cache_data.clear(); st.rerun()