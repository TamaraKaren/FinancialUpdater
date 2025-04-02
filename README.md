# FinancialUpdater: Sistema Automatizado de Datos y Dashboard Financiero

## Descripción General

Este proyecto implementa un sistema automatizado para recopilar datos financieros de múltiples activos (acciones, índices, criptomonedas) desde Yahoo Finance, procesarlos, calcular indicadores clave, almacenarlos en un archivo Excel detallado y visualizarlos a través de un dashboard web interactivo construido con Streamlit.

El objetivo principal es eliminar la recolección manual de datos y proporcionar una visión consolidada y actualizada del mercado y del rendimiento de los activos seleccionados.

## Características Principales

**Actualizador (`financial_updater.py`):**

*   **Recolección de Datos Multi-Activo:** Obtiene datos de una lista configurable de tickers (acciones, índices, cripto) usando `yfinance`.
*   **Actualizaciones Diferenciales Programadas:**
    *   **Tarea Rápida (cada 15 min):** Obtiene datos en vivo (precio, cambio, volumen) para una actualización frecuente.
    *   **Tarea Completa (cada 6 horas):** Descarga historial (1 año), calcula indicadores (SMA 50/200, RSI), obtiene datos fundamentales/perfil, estados financieros y noticias recientes.
*   **Almacenamiento Centralizado:** Guarda toda la información procesada en un archivo Excel (`Dynamic Financial Data.xlsx`) con múltiples hojas (Resumen, Live Data, Historial, Financieros por Ticker, Noticias).
*   **Formato Avanzado de Excel:** Aplica formato detallado (colores, números, anchos) y genera gráficos de rendimiento normalizado usando `openpyxl`.
*   **Logging Detallado:** Registra eventos, advertencias y errores en `Logs/financial_updater.log` y en consola.
*   **Ejecución en Segundo Plano:** Puede ejecutarse discretamente en Windows usando un archivo `.bat`.
*   **(Opcional):** Funcionalidad para enviar Excel por email, actualizar Google Sheets y guardar en SQLite.

**Dashboard (`dashboard.py`):**

*   **Interfaz Web Interactiva:** Dashboard construido con `Streamlit`.
*   **Fuente de Datos:** Lee directamente del archivo Excel generado por el actualizador, con caché (`st.cache_data`) para rendimiento.
*   **Visualizaciones Clave:**
    *   Tabla de Datos en Vivo.
    *   Gráfico Interactivo (Plotly) de Rendimiento Normalizado (1 año) con selección de tickers.
    *   Gráficos de Cambio Diario y Posición en Rango de 52 Semanas.
    *   Tablas de Datos Completas (Resumen, Historial, Noticias) en secciones desplegables.
*   **Fácil Lanzamiento:** Se inicia con un simple comando `streamlit run` o un archivo `.bat`.

## Cómo Funciona

1.  `financial_updater.py` se ejecuta (preferiblemente en segundo plano).
2.  Utiliza `apscheduler` para ejecutar tareas periódicas (rápidas y completas) que obtienen datos de `yfinance`.
3.  Los datos se procesan con `pandas` y se escriben/formatean en `Dynamic Financial Data.xlsx` usando `openpyxl`.
4.  `dashboard.py` (aplicación Streamlit) lee el archivo `.xlsx` actualizado.
5.  El usuario interactúa con el dashboard en el navegador, visualizando los datos y gráficos generados con `plotly`.

## Pila Tecnológica

*   **Lenguaje:** Python 3
*   **Datos:** `yfinance`, `pandas`, `numpy`
*   **Excel:** `openpyxl`
*   **Programación de Tareas:** `apscheduler`
*   **Dashboard Web:** `streamlit`
*   **Gráficos:** `plotly`, `plotly.express`
*   **Logging:** `logging`
*   **Gestión de Entorno:** `venv`
*   **Gestión de Dependencias:** `pip`, `requirements.txt`
*   **Control de Versiones:** Git, GitHub
*   **(Opcionales):** `smtplib`, `email`, `gspread`, `oauth2client` / `google-auth`, `sqlite3`, `cryptography`

## Instalación

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/TamaraKaren/FinancialUpdater.git
    cd FinancialUpdater
    ```
2.  **Crear y activar un entorno virtual:**
    ```bash
    python -m venv venv
    # En Windows:
    .\venv\Scripts\activate
    # En macOS/Linux:
    source venv/bin/activate
    ```
3.  **Instalar dependencias:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **(Opcional):** Configurar credenciales si se utilizan las funciones de email, Google Sheets o base de datos.

## Uso

1.  **Iniciar el Actualizador (Recomendado en segundo plano):**
    *   En Windows: Ejecutar `Iniciar_Updater_Background.bat`.
    *   Directamente (consola visible): `python financial_updater.py`
2.  **Lanzar el Dashboard:**
    *   En Windows: Ejecutar `Lanzar_Dashboard.bat`.
    *   Directamente: `streamlit run dashboard.py`
3.  Abrir la URL proporcionada por Streamlit (usualmente `http://localhost:8501`) en el navegador web.

## Capturas de Pantalla (Ejemplos)

![f19137a5-4867-40b7-b308-efa344dcc857](https://github.com/user-attachments/assets/8996fb7f-a11e-4093-9934-e30d0b13f09f)
![image](https://github.com/user-attachments/assets/7c9d1c65-7845-43dc-b90c-faf4a7fb1bc3)

## Mejoras Futuras Posibles

*   Externalizar configuraciones (tickers, rutas, credenciales) a archivos (`.env`, `.yaml`).
*   Manejo de errores más robusto y lógica de reintentos.
*   Añadir más indicadores técnicos y análisis.
*   Implementar formato condicional avanzado en Excel.
*   Sistema de alertas basado en umbrales.
*   Migrar de `oauth2client` a `google-auth`.
*   Refactorizar código en clases más especializadas.
*   Mejorar gestión de secretos.
