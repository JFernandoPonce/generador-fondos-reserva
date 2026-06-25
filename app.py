#!/usr/bin/env python3
"""
Servidor local - Generador de Informes Fondos de Reserva 17D08
Ejecutar: python3 app.py
Luego abrir: http://localhost:5000
"""
import os, sys, json, tempfile, traceback
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string

# Importar el generador
sys.path.insert(0, str(Path(__file__).parent))
from generar_informe import generate

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB máximo

# ─── LEER HTML DESDE ARCHIVO ──────────────────────────────────
HTML_PATH = Path(__file__).parent / 'index.html'
DATOS_PREV_PATH = Path(__file__).parent / 'datos_anteriores.json'


def _load_datos_anteriores():
    """Carga las novedades precargadas desde datos_anteriores.json (NO versionado).

    Si el archivo no existe o está corrupto, devuelve {} y la app funciona
    igual, solo sin precarga de novedades. Así nunca se rompe el arranque.
    """
    try:
        if DATOS_PREV_PATH.exists():
            return json.loads(DATOS_PREV_PATH.read_text(encoding='utf-8'))
    except Exception as e:
        print(f"[AVISO] No se pudo leer datos_anteriores.json: {e}")
    return {}


@app.route('/')
def index():
    html = HTML_PATH.read_text(encoding='utf-8')
    datos = _load_datos_anteriores()
    inject = ('<script>window.__PREV_DATA__ = '
              + json.dumps(datos, ensure_ascii=False) + ';</script>')
    return html.replace('<!--PREV_DATA_INJECT-->', inject)

# ─── ENDPOINT: GENERAR INFORME ────────────────────────────────
@app.route('/generar', methods=['POST'])
def generar():
    tmp_dir = None
    try:
        # Recibir archivos (ahora son 3)
        if 'planilla' not in request.files or 'matriz' not in request.files or 'distributivo' not in request.files:
            return jsonify({'error': 'Faltan archivos: planilla, matriz y/o distributivo'}), 400

        f_planilla = request.files['planilla']
        f_matriz   = request.files['matriz']
        f_dist     = request.files['distributivo']
        config_str = request.form.get('config', '{}')

        # Validar que sean Excel
        for f in [f_planilla, f_matriz, f_dist]:
            if not f.filename.endswith(('.xlsx', '.xls')):
                return jsonify({'error': f'Archivo {f.filename} no es Excel (.xlsx/.xls)'}), 400

        # Guardar temporalmente
        tmp_dir = tempfile.mkdtemp()
        path_planilla = os.path.join(tmp_dir, 'planilla' + Path(f_planilla.filename).suffix)
        path_matriz   = os.path.join(tmp_dir, 'matriz'   + Path(f_matriz.filename).suffix)
        path_dist     = os.path.join(tmp_dir, 'distributivo' + Path(f_dist.filename).suffix)
        path_output   = os.path.join(tmp_dir, 'informe.docx')

        f_planilla.save(path_planilla)
        f_matriz.save(path_matriz)
        f_dist.save(path_dist)

        # Archivo opcional: Excel de vinculaciones
        path_vin_excel = None
        if 'vinculaciones_excel' in request.files:
            f_vin = request.files['vinculaciones_excel']
            if f_vin and f_vin.filename:
                path_vin_excel = os.path.join(tmp_dir, 'vinculaciones' + Path(f_vin.filename).suffix)
                f_vin.save(path_vin_excel)

        # Completar rutas en config
        cfg = json.loads(config_str)
        cfg['path_planilla']     = path_planilla
        cfg['path_matriz']       = path_matriz
        cfg['path_distributivo'] = path_dist
        cfg['path_vin_excel']    = path_vin_excel
        cfg['output_path']       = path_output

        mes  = cfg.get('mes', 'MES').upper()
        anio = cfg.get('anio', '2026')

        # Generar
        plantilla_path = str(Path(__file__).parent / 'Plantilla.docx')
        total, sin_match, excel_path = generate(json.dumps(cfg), path_output, plantilla_path=plantilla_path)

        # Nombres descriptivos
        nombre_word  = f"INFORME_FONDOS_RESERVA_{mes}_{anio}.docx"
        nombre_excel = f"INFORME_FONDOS_RESERVA_{mes}_{anio}_TABLAS.xlsx"
        nombre_zip   = f"INFORME_FONDOS_RESERVA_{mes}_{anio}.zip"

        # Empaquetar Word + Excel en un ZIP
        import zipfile as zf
        zip_path = os.path.join(tmp_dir, 'resultado.zip')
        with zf.ZipFile(zip_path, 'w', zf.ZIP_DEFLATED) as zipf:
            zipf.write(path_output, nombre_word)
            if excel_path and os.path.exists(excel_path):
                zipf.write(excel_path, nombre_excel)

        return send_file(
            zip_path,
            as_attachment=True,
            download_name=nombre_zip,
            mimetype='application/zip'
        )

    except Exception as e:
        tb = traceback.format_exc()
        print(f"ERROR en /generar:\n{tb}")
        return jsonify({'error': str(e), 'detalle': tb}), 500

# ─── ENDPOINT: CRUCE DE DATOS (PREVIEW) ───────────────────────
@app.route('/cruce', methods=['POST'])
def cruce():
    tmp_dir = None
    try:
        if 'planilla' not in request.files or 'matriz' not in request.files or 'distributivo' not in request.files:
            return jsonify({'error': 'Faltan archivos: planilla, matriz y/o distributivo'}), 400

        f_planilla = request.files['planilla']
        f_matriz   = request.files['matriz']
        f_dist     = request.files['distributivo']

        tmp_dir = tempfile.mkdtemp()
        path_planilla = os.path.join(tmp_dir, 'planilla' + Path(f_planilla.filename).suffix)
        path_matriz   = os.path.join(tmp_dir, 'matriz'   + Path(f_matriz.filename).suffix)
        path_dist     = os.path.join(tmp_dir, 'distributivo' + Path(f_dist.filename).suffix)
        f_planilla.save(path_planilla)
        f_matriz.save(path_matriz)
        f_dist.save(path_dist)

        # Importar funciones del generador (ya maneja .xls y .xlsx)
        from generar_informe import norm_ced, decode_est, _excel_engine
        import pandas as pd

        # ── Distributivo (soporta .xls y .xlsx) ──────────────
        eng_d = _excel_engine(path_dist)
        df_dist = pd.read_excel(path_dist, header=0, engine=eng_d, dtype={'ESTRUCTURA PROGRAMÁTICA': str})
        df_dist.columns = [str(c).strip() for c in df_dist.columns]

        # Debug: verificar que tenemos las columnas necesarias
        print(f"[CRUCE] Distributivo columnas: {list(df_dist.columns)[:8]}...")
        print(f"[CRUCE] Distributivo filas: {len(df_dist)}")

        # Verificar que las columnas existan
        if 'ESTADO DEL PUESTO' not in df_dist.columns:
            return jsonify({'error': f'Columna "ESTADO DEL PUESTO" no encontrada en el Distributivo. '
                           f'Columnas: {list(df_dist.columns)[:8]}. '
                           f'Verifique que subió el Distributivo (Mineduc) como Archivo 2.'}), 400
        if 'NÚMERO IDENTIFICACIÓN' not in df_dist.columns:
            return jsonify({'error': f'Columna "NÚMERO IDENTIFICACIÓN" no encontrada. '
                           f'Columnas: {list(df_dist.columns)[:8]}'}), 400

        occ = {}
        for _, row in df_dist.iterrows():
            ep = str(row.get('ESTADO DEL PUESTO', '')).upper().strip()
            ced_val = row.get('NÚMERO IDENTIFICACIÓN', '')
            if ep == 'OCUPADO' and pd.notna(ced_val) and str(ced_val).strip():
                c   = norm_ced(ced_val)
                raw = str(row.get('ESTRUCTURA PROGRAMÁTICA', ''))
                raw = raw if raw != 'nan' else ''
                occ[c] = {
                    'nombre':     str(row.get('NOMBRES', '')).strip(),
                    'estructura': decode_est(raw) if raw else ''
                }

        print(f"[CRUCE] Ocupados encontrados: {len(occ)}")
        if len(occ) > 0:
            sample = list(occ.keys())[:3]
            print(f"[CRUCE] Muestra cédulas distributivo: {sample}")

        # ── Matriz IESS (soporta .xls y .xlsx) ───────────────
        eng_m = _excel_engine(path_matriz)
        xl = pd.ExcelFile(path_matriz, engine=eng_m)
        # Buscar hoja con cabecera Cedula en columnas reales (saltar hojas CSV de 1 col)
        sheet = None
        for sh in xl.sheet_names:
            try:
                sample = xl.parse(sh, header=None, nrows=8)
                if sample.shape[1] < 3:
                    continue
                for _, row in sample.iterrows():
                    if any('cedula' in str(v).lower() for v in row
                           if str(v).strip() and str(v).strip().lower() != 'nan'):
                        sheet = sh
                        break
                if sheet:
                    break
            except Exception:
                continue
        if sheet is None:
            sheet = xl.sheet_names[0]
        df_raw = xl.parse(sheet, header=None)

        # Buscar fila de cabecera (la que contiene 'Cedula')
        hrow = -1
        for i, row in df_raw.iterrows():
            if any('cedula' in str(v).lower() for v in row if str(v).strip() and str(v).strip().lower() != 'nan'):
                hrow = i; break
        if hrow < 0:
            return jsonify({'error': 'No se encontró la cabecera (Cedula) en la Matriz IESS. Verifique que subió la Matriz IESS como Archivo 2.'}), 400

        df_raw.columns = [str(c).strip() for c in df_raw.iloc[hrow]]
        df = df_raw.iloc[hrow+1:].reset_index(drop=True)
        col_ced = next((c for c in df.columns if 'cedula' in c.lower()), df.columns[0])
        df = df.dropna(subset=[col_ced])
        col_nom = next((c for c in df.columns if 'nombre' in c.lower()), col_ced)
        col_sol = next((c for c in df.columns if 'solicitud' in c.lower() and 'tiene' in c.lower()), None)
        col_der = next((c for c in df.columns if 'derecho' in c.lower() or c.strip().lower() in ('tiene','tiene ')), None)

        df['Cedula'] = df[col_ced].apply(norm_ced)

        print(f"[CRUCE] Matriz filas: {len(df)}")
        if len(df) > 0:
            sample_m = list(df['Cedula'].head(3))
            print(f"[CRUCE] Muestra cédulas matriz: {sample_m}")

        sin_match = []
        for _, r in df.iterrows():
            c = r['Cedula']
            if c not in occ:
                sin_match.append({
                    'cedula': c,
                    'nombre': str(r.get(col_nom, '')).strip(),
                    'tieneSolicitud': str(r.get(col_sol, '')).strip().upper() if col_sol else '',
                    'tieneDerecho':   str(r.get(col_der, '')).strip().upper() if col_der else '',
                })

        print(f"[CRUCE] Sin match: {len(sin_match)} de {len(df)}")

        return jsonify({
            'total_matriz': len(df),
            'total_ocupados': len(occ),
            'sin_match': sin_match
        })

    except Exception as e:
        return jsonify({'error': str(e), 'detalle': traceback.format_exc()}), 500


# ─── ENDPOINT: PARSEAR INFORME ANTERIOR ──────────────────────
@app.route('/parsear_informe', methods=['POST'])
def parsear_informe():
    """Lee un .docx del mes anterior y extrae las tablas de novedades."""
    tmp_dir = None
    try:
        if 'informe' not in request.files:
            return jsonify({'error': 'No se envió el archivo'}), 400
        f = request.files['informe']
        if not f.filename.endswith('.docx'):
            return jsonify({'error': 'El archivo debe ser .docx'}), 400

        tmp_dir = tempfile.mkdtemp()
        path = os.path.join(tmp_dir, 'informe.docx')
        f.save(path)

        # Extraer tablas de novedades usando python-docx
        from docx import Document
        import re
        doc = Document(path)

        novedades = []
        current_title = None
        current_headers = []
        current_rows = []

        # Patterns to detect novedad titles
        NOV_KEYWORDS = ['ENCARGO','SECTORIZ','DESVINCULAC','RENUNCIA','NOTIF','FALLECIM',
                        'JUBILAC','SUSPENSION','SUSPENS','CAMBIO','INCORPORAC','INGRESO DIRECTOR',
                        'VINCULAC','CMYO']

        def limpiar(t):
            return re.sub(r'\*+', '', t).strip()

        def es_titulo_novedad(t):
            tu = t.upper().strip()
            return (re.search(r'^\d+[\.-].*NOVEDAD', tu) or
                    any(kw in tu for kw in NOV_KEYWORDS)) and len(tu) > 5

        def flush_novedad():
            if current_title and current_rows:
                # Only data rows (skip if looks like header again)
                data_rows = [r for r in current_rows if r and not all(
                    str(v).startswith('**') for v in r if v)]
                if data_rows:
                    novedades.append({
                        'nombre': current_title,
                        'titulo': current_title,
                        'headers': current_headers,
                        'filas': data_rows
                    })

        # Iterate document elements in order
        for block in doc.element.body:
            # Paragraph
            from docx.oxml.ns import qn
            if block.tag.endswith('p'):
                text = ''.join(r.text or '' for r in block.iter(qn('w:t'))).strip()
                if es_titulo_novedad(text):
                    flush_novedad()
                    current_title = re.sub(r'^\d+[\.-\s]+NOVEDAD:\s*', '', text, flags=re.IGNORECASE).strip()
                    current_title = re.sub(r'^\d+[\.-\s]+', '', current_title).strip()
                    current_headers = []
                    current_rows = []
            # Table
            elif block.tag.endswith('tbl'):
                if current_title:
                    for row in block.iter(qn('w:tr')):
                        cells = []
                        for cell in row.iter(qn('w:tc')):
                            cell_text = ''.join(r.text or '' for r in cell.iter(qn('w:t'))).strip()
                            cells.append(cell_text)
                        if cells:
                            # First row with bold = header
                            first_cell_runs = list(block.iter(qn('w:b')))
                            if not current_headers:
                                current_headers = cells
                            else:
                                # Skip rows that look like repeated headers
                                if cells != current_headers:
                                    current_rows.append(cells)

        flush_novedad()

        return jsonify({'novedades': novedades, 'total': len(novedades)})

    except Exception as e:
        return jsonify({'error': str(e), 'detalle': traceback.format_exc()}), 500
    finally:
        if tmp_dir:
            import shutil
            try: shutil.rmtree(tmp_dir, ignore_errors=True)
            except: pass


if __name__ == '__main__':
    print("=" * 55)
    print("  GENERADOR DE INFORMES - FONDOS DE RESERVA 17D08")
    print("=" * 55)
    print("  Servidor iniciado en: http://localhost:5000")
    print("  Abra esa dirección en su navegador web.")
    print("  Para cerrar: presione Ctrl+C")
    print("=" * 55)
    app.run(host='127.0.0.1', port=5000, debug=False)
