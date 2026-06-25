#!/bin/bash
echo ""
echo "================================================"
echo "  GENERADOR DE INFORMES - FONDOS DE RESERVA"
echo "  Dirección Distrital 17D08 - Educación Ecuador"
echo "================================================"
echo ""

# Ir al directorio del script
cd "$(dirname "$0")"

# Verificar Python
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python3 no está instalado."
    echo "Instálelo con: brew install python (Mac) o apt install python3 (Ubuntu)"
    exit 1
fi

# Instalar dependencias
echo "Verificando dependencias..."
pip3 install flask openpyxl pandas python-docx xlrd -q

echo ""
echo "================================================"
echo "  Servidor iniciado en: http://localhost:5000"
echo ""
echo "  Abra su navegador y vaya a:"
echo "  http://localhost:5000"
echo ""
echo "  Para cerrar: Ctrl+C"
echo "================================================"
echo ""

# Abrir navegador
sleep 1
if command -v open &> /dev/null; then
    open "http://localhost:5000"   # Mac
elif command -v xdg-open &> /dev/null; then
    xdg-open "http://localhost:5000"   # Linux
fi

python3 app.py
