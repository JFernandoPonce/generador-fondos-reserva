# Generador de Informes — Fondos de Reserva

Aplicación web local que genera automáticamente el **Informe Técnico de Fondos de Reserva** en formato Word (`.docx`), idéntico al formato oficial, a partir de los archivos Excel del mes. Desarrollada para la **Dirección Distrital 17D08 – Educación** (Parroquias Rurales Conocoto a La Merced, Ecuador).

Convierte un proceso manual y repetitivo —cruzar la matriz del IESS con el distributivo del Ministerio de Educación y transcribir todo a Word— en una operación de cinco pasos que termina con el documento descargado.

> **Privacidad por diseño:** la aplicación corre **100 % en local** (servidor Flask en `localhost`), sin enviar nada a la nube. Los Excel de entrada contienen datos personales (cédulas, nombres de afiliados) y nunca salen de la máquina de quien la usa. Este repositorio **no incluye datos reales**, solo la plantilla y el código.

---

## Características

- **Cruce automático** de dos fuentes: la Matriz de Fondos de Reserva del IESS y el Distributivo de Trabajo del Ministerio de Educación.
- **Generación de Word oficial** a partir de `Plantilla.docx`, preservando logo, pie de página y márgenes exactos del formato institucional.
- **Manejo de no coincidencias:** si un funcionario de la matriz no aparece en el distributivo, la app solicita completar sus datos antes de generar.
- **Soporta `.xls` y `.xlsx`** (Excel viejo y nuevo).
- **Interfaz guiada de 5 pasos** servida en el navegador; sin instalación de programas más allá de Python.
- **Arranque de un clic** en Windows (`INICIAR.bat`) o por terminal en Mac/Linux (`iniciar.sh`).

## Stack tecnológico

| Componente | Tecnología |
|------------|------------|
| Servidor | Python + Flask (local, `127.0.0.1:5000`) |
| Lectura de Excel | pandas, openpyxl, xlrd |
| Generación de Word | python-docx + manipulación directa de OOXML (lxml) |
| Interfaz | HTML + JavaScript (una sola página) |
| Plantilla | `Plantilla.docx` (formato oficial) |

## Estructura del proyecto

```
.
├── INICIAR.bat              # Arranque en Windows (doble clic)
├── iniciar.sh               # Arranque en Mac/Linux
├── app.py                   # Servidor Flask + endpoints
├── generar_informe.py       # Motor de generación del Word
├── index.html               # Interfaz de la aplicación (5 pasos)
├── Plantilla.docx           # Plantilla oficial del informe
├── datos_generales_tbl.xml  # Estructura de tabla (encabezado)
├── resumen_tbl.xml          # Estructura de tabla (resumen)
├── requirements.txt
└── LEAME.txt                # Guía de uso para el operador
```

## Requisitos

- **Python 3.8+** ([descarga](https://www.python.org/downloads/) — marcar *"Add Python to PATH"* al instalar).
- Navegador moderno (Chrome o Edge recomendados).
- Conexión a internet solo la primera vez (para instalar dependencias).

## Instalación y uso

### Windows
1. Doble clic en **`INICIAR.bat`**. El script verifica Python, instala dependencias si faltan, levanta el servidor y abre el navegador.
2. Si el navegador no abre solo, ve a `http://localhost:5000`.
3. Sigue los 5 pasos. El informe Word se descarga al finalizar.
4. Para cerrar: cierra la ventana de comandos.

### Mac / Linux
```bash
bash iniciar.sh
```

### Manual (cualquier sistema)
```bash
pip install -r requirements.txt
python app.py
# abrir http://localhost:5000
```

## Flujo de la aplicación

1. **Archivos** — cargar los Excel del mes (matriz IESS + distributivo Mineduc).
2. **Configuración** — datos del encabezado (mes, número de informe, fechas, responsables).
3. **Novedades** — ingresar las del mes (encargos, CMYO, sectorización, jubilaciones, etc.).
4. **Sin coincidencia** — completar nombre y estructura de funcionarios que no crucen.
5. **Generar** — se produce y descarga el informe Word.

## Estado

Herramienta funcional, usada para la generación mensual real de informes en la Dirección Distrital 17D08. El repositorio se publica como portafolio y referencia técnica.

## Autoría

Desarrollado por **Juan Fernando Ponce** — automatización de procesos para el sector público (Educación, Ecuador).
