@echo off
echo Activando entorno virtual y lanzando Dashboard...

:: Cambia al directorio donde se encuentra este archivo .bat
cd /d "%~dp0"

:: Activa el entorno virtual (ajusta 'venv' si tu carpeta se llama diferente)
call venv\Scripts\activate.bat

:: Ejecuta el dashboard de Streamlit
echo Lanzando Streamlit...
streamlit run dashboard.py

:: Mantiene la ventana abierta al final para ver mensajes (opcional)
pause