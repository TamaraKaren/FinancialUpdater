@echo off
echo Iniciando Financial Updater en segundo plano...

:: Cambia al directorio donde se encuentra este archivo .bat
cd /d "%~dp0"

:: Activa el entorno virtual
call venv\Scripts\activate.bat

:: Inicia el script Python en segundo plano (minimized)
echo Ejecutando financial_updater.py en segundo plano...
start /B python financial_updater.py

echo Script 'financial_updater.py' iniciado en segundo plano.
echo Puedes cerrar esta ventana. Para detenerlo, necesitarás buscar el proceso en el Administrador de Tareas.

:: Opcional: No pausar para que la ventana se cierre automáticamente después de iniciar
:: pause