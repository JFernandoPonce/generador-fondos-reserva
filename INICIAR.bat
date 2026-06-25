@echo off
title Generador de Informes - Fondos de Reserva 17D08
color 0A
echo.
echo  ================================================
echo   GENERADOR DE INFORMES - FONDOS DE RESERVA
echo   Direccion Distrital 17D08 - Educacion Ecuador
echo  ================================================
echo.
echo  Iniciando servidor local...
echo  Por favor espere...
echo.

:: Verificar que Python esté instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no esta instalado.
    echo  Descargue Python desde https://www.python.org/downloads/
    echo  y asegurese de marcar "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)

:: Ir al directorio de la aplicación
cd /d "%~dp0"

:: Instalar dependencias si faltan
echo  Verificando dependencias...
pip install flask openpyxl pandas python-docx xlrd -q --no-warn-script-location

echo.
echo  ================================================
echo   Servidor iniciado en: http://localhost:5000
echo.
echo   Abra su navegador (Chrome o Edge) y vaya a:
echo   http://localhost:5000
echo.
echo   Para cerrar: cierre esta ventana o Ctrl+C
echo  ================================================
echo.

:: Abrir navegador automáticamente
timeout /t 2 /nobreak >nul
start "" "http://localhost:5000"

:: Iniciar servidor
python app.py

pause
