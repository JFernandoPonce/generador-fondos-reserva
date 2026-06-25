"""
Tests del Generador de Informes — Fondos de Reserva.

Verifican que tras externalizar las novedades a datos_anteriores.json:
  1. La app importa y arranca sin error.
  2. Las rutas siguen registradas (/, /generar, /cruce, /parsear_informe).
  3. La ruta / sirve el HTML e inyecta correctamente las novedades del JSON.
  4. Si NO existe el JSON, la app sirve el HTML igual con __PREV_DATA__ = {} (no se rompe).
  5. El HTML servido NO contiene cédulas reales hardcodeadas.
  6. El JSON de novedades editado por el usuario se refleja en lo servido (atajo conectado).
"""
import json
import re
import shutil
from pathlib import Path

import app as appmod

HERE = Path(__file__).parent
DATOS = HERE / 'datos_anteriores.json'
CEDULA_RE = re.compile(r"'[0-9]{10}'|\"[0-9]{10}\"")

resultados = []
def check(nombre, condicion, detalle=""):
    estado = "OK  " if condicion else "FALLA"
    resultados.append((condicion, f"[{estado}] {nombre}" + (f" — {detalle}" if detalle and not condicion else "")))


client = appmod.app.test_client()

# --- Test 1: rutas registradas ---
rutas = {r.rule for r in appmod.app.url_map.iter_rules()}
for ruta in ['/', '/generar', '/cruce', '/parsear_informe']:
    check(f"ruta {ruta} registrada", ruta in rutas)

# --- Test 2: / responde 200 y es HTML ---
r = client.get('/')
check("GET / responde 200", r.status_code == 200, f"status={r.status_code}")
body = r.get_data(as_text=True)
check("respuesta es HTML", '<html' in body.lower() or '<!doctype' in body.lower())

# --- Test 3: el marcador fue reemplazado (no queda crudo) ---
check("marcador de inyección reemplazado", '<!--PREV_DATA_INJECT-->' not in body)
check("se inyectó window.__PREV_DATA__", 'window.__PREV_DATA__' in body)

# --- Test 4: las novedades del JSON real aparecen inyectadas ---
datos_reales = json.loads(DATOS.read_text(encoding='utf-8'))
m = re.search(r'window\.__PREV_DATA__ = (\{.*?\});</script>', body, flags=re.DOTALL)
check("bloque __PREV_DATA__ presente y parseable", bool(m))
if m:
    inyectado = json.loads(m.group(1))
    check("novedades inyectadas == datos_anteriores.json",
          inyectado == datos_reales,
          f"esperado {list(datos_reales)} got {list(inyectado)}")
    check("Encargos tiene 4 filas inyectadas",
          len(inyectado.get('Encargos', {}).get('filas', [])) == 4)

# --- Test 5: el index.html en disco NO tiene cédulas hardcodeadas ---
html_disk = (HERE / 'index.html').read_text(encoding='utf-8')
check("index.html sin cédulas hardcodeadas",
      len(CEDULA_RE.findall(html_disk)) == 0,
      f"encontradas {len(CEDULA_RE.findall(html_disk))}")

# --- Test 6: edición del usuario se refleja (atajo conectado) ---
backup = DATOS.read_text(encoding='utf-8')
try:
    DATOS.write_text(json.dumps({"PruebaNovedad": {"filas": [["9", "TEST"]]}},
                                ensure_ascii=False), encoding='utf-8')
    body2 = client.get('/').get_data(as_text=True)
    check("nueva novedad del usuario aparece servida", 'PruebaNovedad' in body2)
    check("novedad vieja ya no aparece (refleja el JSON actual)", 'GÓMEZ PRADO' not in body2)
finally:
    DATOS.write_text(backup, encoding='utf-8')

# --- Test 7: sin el JSON, la app no se rompe (degrada a {}) ---
tmp = HERE / 'datos_anteriores.json.bak'
shutil.move(str(DATOS), str(tmp))
try:
    r3 = client.get('/')
    check("GET / responde 200 SIN datos_anteriores.json", r3.status_code == 200)
    body3 = r3.get_data(as_text=True)
    check("degrada a __PREV_DATA__ = {} sin el JSON",
          'window.__PREV_DATA__ = {};' in body3)
finally:
    shutil.move(str(tmp), str(DATOS))

# --- Reporte ---
print("\n" + "=" * 60)
print("  RESULTADOS DE TESTS")
print("=" * 60)
for _, linea in resultados:
    print(" ", linea)
fallos = [r for ok, r in resultados if not ok]
print("=" * 60)
print(f"  {len(resultados) - len(fallos)}/{len(resultados)} tests OK")
print("=" * 60)
raise SystemExit(1 if fallos else 0)
