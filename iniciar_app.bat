@echo off
REM ============================================
REM  Academia Horizonte - Sistema de certificados
REM  Arranca la aplicacion web con doble clic.
REM  Para DETENERLA: cerrar esta ventana o
REM  presionar Ctrl+C en ella.
REM ============================================
cd /d "%~dp0"

REM El doble clic hereda el entorno del Explorador, que puede que
REM todavia no tenga a Python en el PATH (se instalo despues de
REM abrir la sesion). Por eso se agrega aqui explicitamente.
set "PYDIR=%LOCALAPPDATA%\Programs\Python\Python312"
set "PATH=%PYDIR%;%PYDIR%\Scripts;%PATH%"

REM Si de todos modos no se encuentra Python, se avisa y se sale
where python >nul 2>nul
if errorlevel 1 (
    echo No se encontro Python. Instalalo desde python.org y vuelve a intentarlo.
    pause
    exit /b 1
)

echo Instalando dependencias si faltan (primera vez)...
python -m pip install -r requirements.txt --quiet

echo Iniciando Academia Horizonte en http://127.0.0.1:5000 ...
echo Para detener: cierre esta ventana o presione Ctrl+C

REM Abre el navegador despues de 3 segundos (cuando el servidor ya arranco)
start "" /b cmd /c "timeout /t 3 >nul & start http://127.0.0.1:5000"

python app.py

pause
