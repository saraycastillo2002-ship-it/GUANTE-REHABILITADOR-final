"""
Registro automático de datos del Guante Rehab -> Excel
------------------------------------------------------
Lee los datos del ESP32 por USB, guía la prueba dedo por dedo y guarda
un Excel con los datos crudos y un resumen listo para el informe.

Instalar (una sola vez):   pip install pyserial openpyxl
Usar con el guante:        python registrar_guante.py --puerto COM3
Probar sin el guante:      python registrar_guante.py --simular

IMPORTANTE antes de ejecutar:
  - Cierra el Monitor Serie de Arduino y la página (si estaba conectada por USB).
  - Al abrir el puerto el ESP32 se reinicia y toma el reposo: durante los
    primeros segundos mantén la mano ABIERTA y QUIETA.
"""
import re, sys, time, math, random, argparse
from datetime import datetime
from statistics import mean, stdev
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------- Ajustes (se pueden cambiar) ----------------
DEDOS = ["Pulgar", "Índice", "Medio", "Anular"]
REPS = 10             # cierres por dedo
SEG_DEDO = 25         # segundos de grabación por dedo
SEG_REPOSO = 10       # segundos de reposo
CV_MAX = 15.0         # % máximo de variación entre repeticiones
DIAF_MAX = 25.0       # % máximo de respuesta de los otros dedos
FP_MAX = 5.0          # % máximo de falsos positivos en reposo
MIN_REPS = 6          # repeticiones mínimas detectadas (de REPS)
# No existe una norma específica para guantes de datos con sensores flex.
# Estos límites son más flexibles y los fijó el equipo; ver nota en el Excel.

PATRON = re.compile(r"^\s*\[\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\]\s*$")


# ---------------- Fuentes de datos ----------------
class Puerto:
    def __init__(self, nombre):
        import serial
        self.s = serial.Serial(nombre, 115200, timeout=1)

    def _linea(self):
        m = PATRON.match(self.s.readline().decode(errors="ignore"))
        return [int(m.group(i)) for i in range(1, 5)] if m else None

    def esperar_inicio(self):
        print("Esperando al guante (mano abierta y quieta)...")
        t0 = time.time()
        while time.time() - t0 < 60:
            if self._linea() is not None:
                return True
        return False

    def leer(self, segundos, dedo):
        self.s.reset_input_buffer()
        datos, t0 = [], time.time()
        while time.time() - t0 < segundos:
            v = self._linea()
            if v: datos.append((round(time.time() - t0, 3), v))
        return datos


class Simulado:
    def esperar_inicio(self): return True

    def leer(self, segundos, dedo):
        datos, amp = [], 300
        for i in range(int(segundos * 25)):
            t = i / 25
            if i % 50 == 0: amp = random.gauss(300, 12)
            pulso = max(0, math.sin(2 * math.pi * t / 2.0)) if dedo is not None else 0
            v = []
            for c in range(4):
                x = amp * pulso if c == dedo else amp * 0.12 * pulso * random.random()
                x = int(x)
                v.append(x if x > 30 else 0)   # zona muerta del firmware
            datos.append((round(t, 3), v))
        return datos


# ---------------- Análisis ----------------
def segmentos(serie, umbral):
    segs, ini = [], None
    for i, v in enumerate(serie):
        if v > umbral and ini is None: ini = i
        elif v <= umbral and ini is not None: segs.append((ini, i)); ini = None
    if ini is not None: segs.append((ini, len(serie)))
    return [s for s in segs if s[1] - s[0] >= 3]


def analizar_dedo(k, datos):
    cols = list(zip(*[d[1] for d in datos]))
    a = cols[k]
    if max(a) < 60: return None
    segs = segmentos(a, max(a) * 0.5)
    picos = [max(a[i:j]) for i, j in segs]
    diaf = [max(max(cols[c][i:j]) for c in range(4) if c != k) / max(a[i:j]) * 100 for i, j in segs]
    r = {"reps": len(segs), "pico": mean(picos) if picos else None,
         "sd": stdev(picos) if len(picos) > 1 else None, "diaf": mean(diaf) if diaf else None}
    r["cv"] = r["sd"] / r["pico"] * 100 if r["sd"] is not None and r["pico"] else None
    return r


def analizar_reposo(datos):
    cols = list(zip(*[d[1] for d in datos]))
    return [sum(1 for x in c if x > 0) / len(c) * 100 for c in cols]


# ---------------- Excel ----------------
def guardar(nombre, voluntario, reposo, res, crudos):
    wb = Workbook(); ws = wb.active; ws.title = "Resumen"
    b = Font(name="Arial", bold=True); n = Font(name="Arial")
    h = PatternFill("solid", fgColor="1F4E78"); hw = Font(name="Arial", bold=True, color="FFFFFF")
    ws["A1"] = "Validación del Guante Rehab"; ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = f"Voluntario: {voluntario}   Fecha: {datetime.now():%d/%m/%Y %H:%M}"; ws["A2"].font = n

    def fila_enc(r, textos):
        for i, t in enumerate(textos):
            c = ws.cell(r, 1 + i, t); c.font = hw; c.fill = h; c.alignment = Alignment(horizontal="center", wrap_text=True)

    fila_enc(4, ["Dedo", "Repeticiones detectadas", "Pico medio (cuentas)", "Desv. estándar", "CV (%)",
                 "Diafonía (%)", "Falsos positivos en reposo (%)", "Resultado"])
    for i, d in enumerate(DEDOS):
        r = res[i]; fp = reposo[i]
        if r is None:
            vals = [d, 0, None, None, None, None, round(fp, 1), "SIN SEÑAL"]
        else:
            ok = (r["cv"] is not None and r["cv"] <= CV_MAX and r["diaf"] <= DIAF_MAX
                  and fp <= FP_MAX and r["reps"] >= MIN_REPS)
            vals = [d, r["reps"], round(r["pico"], 1), None if r["sd"] is None else round(r["sd"], 1),
                    None if r["cv"] is None else round(r["cv"], 1), round(r["diaf"], 1), round(fp, 1),
                    "CUMPLE" if ok else "NO CUMPLE"]
        for j, v in enumerate(vals):
            c = ws.cell(5 + i, 1 + j, v); c.font = n; c.alignment = Alignment(horizontal="center")
        res_c = ws.cell(5 + i, 8); res_c.font = Font(name="Arial", bold=True,
                color="006100" if res_c.value == "CUMPLE" else "9C0006")
    notas = [f"Criterios: CV ≤ {CV_MAX}%, diafonía ≤ {DIAF_MAX}%, falsos positivos ≤ {FP_MAX}%, "
             f"al menos {MIN_REPS} de {REPS} repeticiones detectadas.",
             "No existe una norma específica para guantes con sensores flex: los límites los fijó el equipo.",
             "Contexto de la literatura: Hosie et al. 2025 (Sensors 25(2):367) reportan error medio de 1.25° en sensores flex;",
             "otro estudio (Sensors 2025, 25, 7068) usa 5° como umbral de exactitud del goniómetro clínico.",
             "CV = desviación estándar / pico medio. Diafonía = respuesta de los otros dedos / respuesta del dedo activo.",
             "El firmware envía 0 si el cambio es menor a 30 cuentas, por eso respuestas pequeñas de otros dedos no se ven."]
    for i, t in enumerate(notas):
        ws.cell(10 + i, 1, t).font = Font(name="Arial", italic=True, size=9)
    for col, w in zip("ABCDEFGH", [14, 16, 16, 14, 10, 12, 20, 14]): ws.column_dimensions[col].width = w

    wd = wb.create_sheet("Datos")
    for i, t in enumerate(["Tiempo (s)", "Prueba"] + DEDOS):
        c = wd.cell(1, 1 + i, t); c.font = hw; c.fill = h
    fila = 2
    for prueba, datos in crudos:
        for t, v in datos:
            wd.cell(fila, 1, t); wd.cell(fila, 2, prueba)
            for j in range(4): wd.cell(fila, 3 + j, v[j])
            fila += 1
    for col, w in zip("ABCDEF", [12, 18, 10, 10, 10, 10]): wd.column_dimensions[col].width = w
    wd.freeze_panes = "A2"
    wb.save(nombre)


# ---------------- Programa ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--puerto", help="Ej.: COM3 (Windows) o /dev/ttyUSB0")
    ap.add_argument("--simular", action="store_true", help="Prueba sin guante con datos inventados")
    a = ap.parse_args()

    if a.simular: fuente = Simulado()
    else:
        puerto = a.puerto
        if not puerto:
            from serial.tools import list_ports
            for p in list_ports.comports(): print(" ", p.device, "-", p.description)
            puerto = input("Escribe el puerto del ESP32 (ej. COM3): ").strip()
        fuente = Puerto(puerto)
    voluntario = input("Nombre o código del voluntario: ").strip() or "sin nombre"

    if not fuente.esperar_inicio():
        print("No llegan datos del guante. Revisa el puerto y que el Monitor Serie esté cerrado."); sys.exit(1)

    crudos = []
    input(f"\n1) REPOSO: mano abierta y quieta {SEG_REPOSO} s. Pulsa ENTER para empezar...")
    rep = fuente.leer(SEG_REPOSO, None); crudos.append(("Reposo", rep))
    reposo = analizar_reposo(rep)

    res = []
    for k, d in enumerate(DEDOS):
        input(f"\n{k + 2}) {d.upper()}: cierra y abre SOLO el {d.lower()} {REPS} veces "
              f"(unos {SEG_DEDO} s). Pulsa ENTER y empieza...")
        datos = fuente.leer(SEG_DEDO, k); crudos.append((d, datos))
        r = analizar_dedo(k, datos); res.append(r)
        print("   ->", "sin señal" if r is None else f"{r['reps']} repeticiones, pico medio {r['pico']:.0f}")

    nombre = f"validacion_guante_{datetime.now():%Y%m%d_%H%M}.xlsx"
    guardar(nombre, voluntario, reposo, res, crudos)
    print(f"\nListo. Archivo guardado: {nombre}")


if __name__ == "__main__":
    main()
