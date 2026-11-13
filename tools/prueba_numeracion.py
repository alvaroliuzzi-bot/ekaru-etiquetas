# -*- coding: utf-8 -*-
"""Pruebas de la numeración v2 y del trabajo de impresión de E-Karú Etiquetas (sin pantalla ni impresora).

Uso:   python tools/prueba_numeracion.py
Usa un tkinter falso (tools/_tk_shim.py), así corre en cualquier PC y no abre ventanas. Si están instalados
Pillow, NumPy y OpenCV, además dibuja las etiquetas y vuelve a leer sus códigos de barras con un decodificador.
"""
import contextlib
import csv
import importlib
import json
import os
import random
import sys
import tempfile
import time
import traceback
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))

import _tk_shim  # noqa: E402

_tk_shim.install(force=True)
app = importlib.import_module("app")  # noqa: E402

TODAY = date(2026, 11, 12)
RESULTS = []


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def gs1_check(payload):
    """Dígito de control GS1 calculado de forma independiente (suma de pares/impares desde la derecha)."""
    digits = [int(c) for c in payload][::-1]
    total = sum(digits[0::2]) * 3 + sum(digits[1::2])
    return str((10 - total % 10) % 10)


# Decodificador EAN-13 propio, armado con la definición estándar (L; R = complemento de L; G = R al revés) y no con
# las tablas de app.py: así la prueba comprueba de verdad lo que se imprimiría.
_L = ("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
_R = tuple("".join("1" if c == "0" else "0" for c in code) for code in _L)
_G = tuple(code[::-1] for code in _R)
_PARITY = ("LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG", "LGGLLG", "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL")


def scanline_bits(elements, y):
    """Fila de puntos (1 = negro) que vería un lector a la altura «y» de la etiqueta de 160 puntos de ancho."""
    row = [0] * 160
    for e in app._text_to_bars(elements):
        if e[0] == "bar" and e[2] <= y < e[2] + e[4]:
            for x in range(max(0, e[1]), min(160, e[1] + e[3])):
                row[x] = 1
    return row


def decode_ean13(row):
    """Decodifica un EAN-13 de una fila de puntos (módulo = el trazo más fino). None si no es un EAN-13 limpio."""
    runs, last, count = [], None, 0
    for value in row:
        if value == last:
            count += 1
        else:
            if last is not None:
                runs.append((last, count))
            last, count = value, 1
    runs.append((last, count))
    while runs and runs[0][0] == 0:
        runs.pop(0)
    while runs and runs[-1][0] == 0:
        runs.pop()
    module = min(length for _, length in runs)
    bits = "".join(str(value) * round(length / module) for value, length in runs)
    if len(bits) != 95 or bits[:3] != "101" or bits[45:50] != "01010" or bits[-3:] != "101":
        return None
    left = [bits[3 + 7 * i:10 + 7 * i] for i in range(6)]
    right = [bits[50 + 7 * i:57 + 7 * i] for i in range(6)]
    digits, parity = [], ""
    for chunk in left:
        if chunk in _L:
            digits.append(str(_L.index(chunk)))
            parity += "L"
        elif chunk in _G:
            digits.append(str(_G.index(chunk)))
            parity += "G"
        else:
            return None
    if parity not in _PARITY:
        return None
    first = str(_PARITY.index(parity))
    for chunk in right:
        if chunk not in _R:
            return None
        digits.append(str(_R.index(chunk)))
    return first + "".join(digits)


def decode_label(elements):
    """Lee el código de barras de una etiqueta (la fila del medio de las barras más altas)."""
    bars = [e for e in elements if e[0] == "bar" and e[4] > 40]
    if not bars:
        return None
    return decode_ean13(scanline_bits(elements, bars[0][2] + bars[0][4] // 2))


# --------------------------------------------------------------------------- catálogo
def test_catalogo():
    numbers = sorted(app.products_by_number())
    check(numbers == list(range(1, 48)), f"los productos deberían ser 1..47: {numbers[:3]}…")
    padre = app.padre_by_code()
    check(len(padre) == 9, f"deberían ser 9 grupos PADRE distintos, hay {len(padre)}")
    for code in padre:
        check(app.is_valid_ean13(code), f"PADRE {code} no cierra el control")
    check(len(app.CODIGOS_ANTERIORES) == 47 and len(app.PADRE_ANTERIORES) == 9, "tablas de códigos anteriores incompletas")
    for old, new in app.PADRE_ANTERIORES.items():
        check(old[:12] == new[:12] or old[:11] == new[:11], f"el PADRE {old} → {new} cambia más que el último dígito")
        check(app.with_check(old[:12]) == new, f"{new} no es el control correcto de {old}")
        check(new in padre, f"el PADRE {new} no está en el catálogo")
    for number, old in app.CODIGOS_ANTERIORES.items():
        check(old[:7] == f"2{number:06d}", f"código anterior {old} no corresponde al Nº {number}")
    for name, data in app.PRODUCTS.items():
        if data["barcode"]:
            check(app.is_valid_ean13(data["barcode"]), f"{name}: el código del catálogo no cierra")
        else:
            check("Balde" in name, f"{name} sin código y no es un balde")


def test_ejemplos_del_informe():
    d = date(2026, 11, 10)
    check(app.codigo_producto(1, d) == "2200012611104", app.codigo_producto(1, d))
    check(app.spaced_code(app.codigo_producto(1, d)) == "22 0001 261110 4", "formato 22 0001 261110 4")
    check(app.codigo_pack(1) == "2300000000016", app.codigo_pack(1))
    check(app.spaced_code(app.codigo_pack(1)) == "23 0000000001 6", "formato 23 0000000001 6")
    check(app.codigo_id_balde(1) == "2400000000013", app.codigo_id_balde(1))
    check(app.PRODUCTS["Grupo PADRE 1 kg"]["barcode"] == "2900002800081", "Pote 1 kg: último dígito 8 → 1")
    check(app.spaced_code("2900002800081") == "2 900002 80008 1", "PADRE: 2 900002 80008 1")
    check(app.mod10_check("20000010110500041129") == "3", "balde con serie: control 3")
    check(app.codigo_insumo(1, d) == app.with_check("250001261110"), "insumo 25")
    check(app.aammdd(d) == "261110" and app.aammdd() == f"{date.today():%y%m%d}", "AAMMDD")


def test_digitos_de_control():
    rng = random.Random(7)
    for _ in range(3000):
        payload = "".join(rng.choice("0123456789") for _ in range(12))
        check(app.ean13_check_digit(payload) == gs1_check(payload), f"EAN-13 {payload}")
        check(app.mod10_check(payload) == gs1_check(payload), f"mod10 de 12 dígitos {payload}")
        long = "".join(rng.choice("0123456789") for _ in range(20))
        check(app.mod10_check(long) == gs1_check(long), f"mod10 de 20 dígitos {long}")
    for code in (app.codigo_producto(5, TODAY), app.codigo_pack(123456), app.codigo_id_balde(987654321),
                 app.codigo_insumo(9999, TODAY)):
        check(len(code) == 13 and app.is_valid_ean13(code), f"{code} tiene que ser EAN-13 válido")


def test_limites():
    for bad in (lambda: app.codigo_producto(0), lambda: app.codigo_producto(10000), lambda: app.codigo_insumo("x"),
                lambda: app.codigo_pack(-1), lambda: app.codigo_pack(10 ** 10), lambda: app.codigo_id_balde(None)):
        try:
            bad()
        except ValueError:
            continue
        raise AssertionError("debería rechazar un número fuera de rango")
    check(len(app.codigo_pack(app.MAX_NUMERO_CORRELATIVO)) == 13, "el número más alto cabe en 13 dígitos")


# --------------------------------------------------------------------------- print_code
def test_print_code_nueva():
    app.set_numbering(app.NUMERACION_NUEVA)
    app.set_elaboration_day(None)
    for number, name in app.products_by_number().items():
        code = app.print_code(name, TODAY)
        check(code == app.codigo_producto(number, TODAY), f"{name}: {code}")
        check(code.startswith("22") and app.is_valid_ean13(code), f"{name}: no es un 22 válido")
        check(code[2:6] == f"{number:04d}" and code[6:12] == "261112", f"{name}: Nº o fecha")
    for name, data in app.PRODUCTS.items():
        if data["barcode"][:2] == "29" or not data["barcode"]:
            check(app.print_code(name, TODAY) == data["barcode"], f"{name}: PADRE y baldes no cambian")
    app.set_elaboration_day(date(2026, 10, 1))
    check(app.print_code("PALITO CON COBERTURA")[6:12] == "261001", "fecha elegida a mano")
    check(app.print_code("PALITO CON COBERTURA", TODAY)[6:12] == "261112", "el parámetro «day» manda")
    app.set_elaboration_day(date.today())
    check(app.CODE_STATE["dia"] is None, "elegir hoy equivale a «automático»")
    app.set_elaboration_day(None)


def test_print_code_anterior():
    """La numeración anterior (versión 2.0) tiene que seguir saliendo igual."""
    app.set_numbering(app.NUMERACION_ACTUAL)
    try:
        for name, data in app.PRODUCTS.items():
            code = app.print_code(name, date(2026, 11, 3))
            if name.startswith(("PALITO", "ALFAJOR", "BROWNIE", "PASTAFROLA", "GALLETITA", "TARTA", "TORTA", "CONO FAN")):
                check(code == app.with_check(data["barcode"][:7] + "26307"), f"{name}: {code}")
            else:
                check(code == data["barcode"], f"{name}: sin fecha")
    finally:
        app.set_numbering(app.NUMERACION_NUEVA)


# --------------------------------------------------------------------------- lector
def read(code, today=TODAY, supplies=None):
    return app.interpretar_codigo(code, today, supplies)


def test_lector_ejemplos():
    r = read("2200012611104")
    check(r["estado"] == "ok" and r["producto"] == "PALITO CON COBERTURA", r)
    check(dict(r["campos"])["Fecha de elaboración"] == "10/11/2026", r["campos"])
    check(read("22 0001 261110 4")["estado"] == "ok", "el escáner puede traer espacios")
    r = read("200000101105000411293")
    check(r["estado"] == "ok" and r["tipo"] == "Balde con transacción", r)
    check(dict(r["campos"]) == {"Sabor": "000001", "Peso": "1,105 kg", "Transacción": "00041129"}, r["campos"])
    check(read("200000101105000411292")["estado"] == "rechazado", "21 dígitos con control malo")
    check(read("200000101105000411293"[:-1] + "3")["estado"] == "ok", "control bueno")
    check(read("2000001011058")["estado"] == "ok" and read("2000001011058")["tipo"] == "Balde", "balde de la balanza")
    check(read("2000001011059")["estado"] == "rechazado", "balde con control malo (regla 2)")
    check(read("2000001300009")["tipo"] != "Balde" or read("2000001300009")["estado"] == "rechazado", "peso fuera de rango")
    r = read("2300000000016")
    check(r["tipo"] == "Promo Pack" and dict(r["campos"])["Nº de paquete"] == "1", r)
    r = read("2400000000013")
    check(r["tipo"] == "ID de balde" and dict(r["campos"])["Nº de balde"] == "1", r)
    r = read("2500012611105", supplies={1: "Fruta a granel"})
    check(r["estado"] == "ok" and dict(r["campos"])["Insumo"] == "Fruta a granel", r)
    r = read("2900002800081")
    check(r["estado"] == "ok" and r["tipo"] == "Grupo PADRE" and "1 kg" in dict(r["campos"])["Grupo"], r)


def test_lector_rechazos():
    check(read("")["estado"] == "rechazado", "vacío")
    check(read("2200012611105")["tipo"] == "Dígito de control incorrecto", "control malo")
    check(read(app.with_check("220001261113"))["estado"] == "rechazado", "fecha futura (13/11 y hoy es 12/11)")
    check(read(app.with_check("220001261340"))["estado"] == "rechazado", "mes 13")
    check(read(app.with_check("229999261110"))["estado"] == "rechazado", "producto que no existe")
    check(read(app.with_check("220000261110"))["estado"] == "rechazado", "producto 0")
    for prefix in ("21", "26", "27", "28"):
        r = read(app.with_check(prefix + "0000000001"))
        check(r["estado"] == "rechazado" and r["regla"] == "Regla 6", f"{prefix} reservado: {r}")
    check(read(app.with_check("290000999999"))["estado"] == "rechazado", "PADRE desconocido")
    check(read(app.with_check("250001261340"))["estado"] == "rechazado", "insumo con fecha mala")
    check(read(app.with_check("250001261113"))["estado"] == "rechazado", "insumo con fecha futura")


def test_lector_externos():
    check(read("7840000000019")["estado"] in ("externo", "rechazado"), "EAN")
    ean = app.with_check("784000000001")
    r = read(ean)
    check(r["estado"] == "externo" and r["regla"].startswith("Regla 7"), r)
    check(read("ABC-123")["estado"] == "externo", "código con letras")
    check(read("12345")["estado"] == "externo", "largo raro")
    check(read(app.with_check("784000000002")[:-1] + "0")["estado"] in ("externo", "rechazado"), "EAN con control malo")
    bad = app.with_check("784000000001")
    bad = bad[:-1] + str((int(bad[-1]) + 1) % 10)
    check(read(bad)["estado"] == "rechazado", "EAN de proveedor con control malo se rechaza (regla 2)")


def test_transicion():
    table = app.tabla_transicion()
    for number, old in app.CODIGOS_ANTERIORES.items():
        r = read(old)
        check(r["estado"] == "viejo" and r["producto"] == app.products_by_number()[number], f"{old}: {r}")
    for number, name in app.products_by_number().items():
        r = read(app.PRODUCTS[name]["barcode"])
        check(r["estado"] == "viejo" and r["producto"] == name, f"{name}: código 2.0 sin fecha → {r['estado']}")
    for old, new in app.PADRE_ANTERIORES.items():
        r = read(old)
        check(r["estado"] == "viejo" and new in r["mensaje"], f"{old}: {r}")
    # numeración anterior con fecha (AADDD) de dos años seguidos: siempre «viejo», con la fecha correcta
    app.set_numbering(app.NUMERACION_ACTUAL)
    try:
        day = date(2024, 1, 1)
        while day <= date(2030, 12, 31):
            for name in ("PALITO CON COBERTURA", "TORTA OREO PREMIUM", "CONO FAN 3X25"):
                r = read(app.print_code(name, day), today=date(2031, 1, 1))
                check(r["estado"] == "viejo" and r["producto"] == name, f"{name} {day}: {r['estado']} {r['tipo']}")
                check(dict(r["campos"])["Elaborado"] == f"{day:%d/%m/%Y}", f"{name} {day}: {r['campos']}")
            day += timedelta(days=17)
    finally:
        app.set_numbering(app.NUMERACION_NUEVA)
    check(len(table) >= 47 + 9, "la tabla tiene los 47 productos y los 9 PADRE")


def test_sin_choques():
    """Ningún código nuevo se lee como balde ni como código viejo, y viceversa (todos los productos, 3 años)."""
    names = app.products_by_number()
    day = date(2026, 1, 1)
    while day <= date(2028, 12, 31):
        for number in names:
            code = app.codigo_producto(number, day)
            r = read(code, today=date(2030, 1, 1))
            check(r["estado"] == "ok" and r["tipo"] == "Producto terminado" and r["producto"] == names[number],
                  f"{code}: {r['estado']} {r['tipo']}")
            check(dict(r["campos"])["Fecha de elaboración"] == f"{day:%d/%m/%Y}", f"{code}: fecha")
        day += timedelta(days=1)
    # baldes de la balanza (sabor 1..99, 0,5–12 kg) → nunca «viejo» salvo el choque conocido del Informe
    old_static = set(app.tabla_transicion())
    clashes = 0
    for flavor in range(1, 100):
        for grams in list(range(500, 12001, 7)) + [10001]:  # 10001 = el campo de los palitos viejos (el choque)
            code = app.with_check(f"2{flavor:06d}{grams:05d}")
            r = read(code)
            if r["estado"] == "viejo":
                check(code in old_static, f"balde {code} leído como código viejo sin estar en la tabla")
                clashes += 1
            else:
                check(r["tipo"] == "Balde" and r["estado"] == "ok", f"{code}: {r['estado']} {r['tipo']}")
    print(f"      (choques balde ↔ código viejo hallados en la muestra: {clashes}; los resuelve la tabla de transición)")


def test_fechas_del_usuario():
    check(app.parse_user_date("12/11/2026") == date(2026, 11, 12), "dd/mm/aaaa")
    check(app.parse_user_date("12-11-26") == date(2026, 11, 12), "dd-mm-aa")
    check(app.parse_user_date("12.11.2026") == date(2026, 11, 12), "dd.mm.aaaa")
    check(app.parse_user_date("12112026") == date(2026, 11, 12), "ddmmaaaa")
    check(app.parse_user_date("031126") == date(2026, 11, 3), "ddmmaa")
    for bad in ("", "31/02/2026", "hoy", "1/2", "12/13/2026", "a/b/c", "12/11/26/1"):
        check(app.parse_user_date(bad) is None, f"«{bad}» no es una fecha")
    check(app.parse_aammdd("261110") == date(2026, 11, 10) and app.parse_aammdd("261310") is None, "AAMMDD")


# --------------------------------------------------------------------------- tabla de transición
def test_tabla_transicion_csv():
    rows = app.transition_rows(TODAY)
    check(len(rows) == 1 + 47 + 9, f"filas: {len(rows)}")
    with tempfile.TemporaryDirectory() as tmp:
        path = app.write_transition_csv(Path(tmp) / "t.csv", TODAY)
        raw = path.read_bytes()
        check(raw.startswith(b"\xef\xbb\xbf"), "UTF-8 con BOM para Excel")
        back = list(csv.reader(raw.decode("utf-8-sig").splitlines(), delimiter=";"))
        check(back == rows, "el CSV vuelve igual")
    for row in rows[1:48]:
        check(app.is_valid_ean13(row[5]) and row[5].startswith("22"), row)
        check(read(row[5], today=TODAY)["estado"] == "ok", row)
        check(read(row[3])["estado"] == "viejo", row)
    for row in rows[48:]:
        check(app.is_valid_ean13(row[5]) and row[5][:2] == "29", row)


# --------------------------------------------------------------------------- datos propios
def test_contadores_e_insumos():
    with tempfile.TemporaryDirectory() as tmp:
        old = (app.COUNTERS_FILE, app.SUPPLIES_FILE)
        app.COUNTERS_FILE, app.SUPPLIES_FILE = Path(tmp) / "c.json", Path(tmp) / "i.json"
        try:
            check(app.load_counters() == {"pack": 0, "balde": 0}, "sin archivo: todo en cero")
            app.save_counters({"pack": 12, "balde": 3})
            check(app.load_counters() == {"pack": 12, "balde": 3}, "ida y vuelta")
            app.COUNTERS_FILE.write_text("{rotura", encoding="utf-8")
            check(app.load_counters() == {"pack": 0, "balde": 0}, "archivo roto no tira la app")
            app.COUNTERS_FILE.write_text(json.dumps({"pack": -4, "balde": "x"}), encoding="utf-8")
            check(app.load_counters() == {"pack": 0, "balde": 0}, "valores inválidos se ignoran")
            names = [n for _, n in app.load_supplies()]
            check("Conos" in names and names == [n for _, n in app.INSUMOS_INICIALES], names)
            number, items = app.add_supply("  Frutilla   a granel ")
            check(number == 7 and (7, "Frutilla a granel") in items, items)
            check(app.load_supplies() == items, "se guardó")
            try:
                app.add_supply("frutilla a granel")
                raise AssertionError("debería rechazar un nombre repetido")
            except ValueError:
                pass
            try:
                app.add_supply("   ")
                raise AssertionError("debería rechazar un nombre vacío")
            except ValueError:
                pass
            app.SUPPLIES_FILE.write_text("no es json", encoding="utf-8")
            check(len(app.load_supplies()) == 6, "archivo roto: vuelve a la lista inicial")
        finally:
            app.COUNTERS_FILE, app.SUPPLIES_FILE = old


# --------------------------------------------------------------------------- impresión
CFG = dict(app.DEFAULT_CONFIG)


def test_tspl_igual_que_antes():
    """Para copias idénticas, el programa TSPL es el mismo de siempre; la serie agrega un PRINT por etiqueta."""
    elements = app.product_elements("PALITO CON COBERTURA")
    data = app.build_tspl(elements, 20, 30, 3, CFG)
    check(data.endswith(b"PRINT 1,3\r\n") and data.count(b"PRINT") == 1, "copias idénticas: un solo PRINT")
    check(data.startswith(b"SIZE 20 mm,30 mm\r\nGAP 2 mm,0 mm\r\nDIRECTION 1\r\nREFERENCE 0,0\r\nDENSITY 8\r\n"
                          b"SPEED 4\r\nCODEPAGE 1252\r\nCLS\r\n"), data[:120])
    labels = [app.serial_label_elements("PROMO PACK", "12 PALITOS", app.codigo_pack(n)) for n in (5, 6, 7)]
    series = app.build_tspl(labels[0], 20, 30, 3, CFG, labels)
    check(series.count(b"PRINT 1,1\r\n") == 3 and series.count(b"CLS\r\n") == 3, "serie: 3 etiquetas = 3 PRINT y 3 CLS")
    check(series.count(b"PRINT") == 3, "sin PRINT de más")
    parts = series.split(b"CLS\r\n")
    check(len(parts) == 4 and parts[1] != parts[2] != parts[3], "cada etiqueta es distinta")
    job = app.PrintJob(labels[0], 20, 30, 1, CFG, "grid", labels=labels)
    check(job.copies == 3 and job.labels and len(job.labels) == 3 and job.data == series, "PrintJob de serie")
    plain = app.PrintJob(elements, 20, 30, 5, CFG, "grid")
    check(plain.copies == 5 and plain.labels is None, "PrintJob común")
    many = app.PrintJob(elements, 20, 30, 1, CFG, "grid", labels=[elements] * 1500)
    check(many.copies == 999 and len(many.labels) == 999, "una serie nunca pasa de 999 etiquetas")


class FakeGdi:
    """GDI falso: anota cada rectángulo que se pinta, página por página."""

    def __init__(self):
        self.pages, self.fonts, self.text_calls = [], 0, 0

    def CreateDCW(self, *a):
        return 1

    def DeleteDC(self, h):
        pass

    def GetDeviceCaps(self, h, index):
        return {88: 203, 90: 203, 8: 812, 10: 264}[index]

    def StartDocW(self, h, doc):
        return 1

    def StartPage(self, h):
        self.pages.append([])
        return 1

    def EndPage(self, h):
        return 1

    def EndDoc(self, h):
        return 1

    def AbortDoc(self, h):
        raise AssertionError("no debería abortar")

    def PatBlt(self, h, x, y, w, hh, rop):
        self.pages[-1].append((x, y, w, hh))
        return 1

    def CreateFontW(self, *a):
        self.fonts += 1
        return self.fonts

    def SelectObject(self, h, o):
        return 0

    def DeleteObject(self, o):
        pass

    def SetBkMode(self, *a):
        pass

    def SetTextColor(self, *a):
        pass

    def TextOutW(self, *a):
        self.text_calls += 1
        return 1


def run_gdi(job):
    from ctypes import wintypes
    spooler = app.WinSpooler.__new__(app.WinSpooler)
    spooler.wt = wintypes
    spooler.gdi = FakeGdi()
    spooler._devmode = lambda *a, **k: None
    spooler._err = lambda: "error simulado"
    detail = spooler.send_gdi("HPRT", job, "prueba")
    return spooler.gdi, detail


def expected_rects(elements, dx, dy):
    """Rectángulos que debería pintar el driver para una etiqueta corrida (dx, dy) puntos (la negrita engrosa 1 punto)."""
    out = []

    def add(x, y, w, h):
        x0, x1 = round(x), round(x + w)
        y0, y1 = round(y), round(y + h)
        out.append((x0, y0, max(1, x1 - x0), max(1, y1 - y0)))

    for e in elements:
        if e[0] == "bar":
            add(e[1] + dx, e[2] + dy, e[3], e[4])
        elif e[0] == "text":
            _, x, y, font, text, bold = e
            fw = app.FONT_DOTS[font][0]
            for index, ch in enumerate(text):
                for rx, ry, rw, rh in (app.glyph_rects(font, ch) or []) if ch != " " else []:
                    add(x + dx + index * fw + rx, y + dy + ry, rw + (1 if bold else 0), rh)
    return out


def test_gdi_serie_y_paginas():
    cfg = dict(CFG, grid=True, cols=4, col_pitch_mm=23.0, margin_left_mm=7.7, margin_top_mm=0.0)
    labels = [app.serial_label_elements("PROMO PACK", "12 PALITOS", app.codigo_pack(n)) for n in range(1, 7)]
    job = app.PrintJob(labels[0], 20, 30, 1, cfg, "grid", labels=labels)
    gdi, detail = run_gdi(job)
    check(len(gdi.pages) == 2, f"6 etiquetas en 4 columnas = 2 páginas, hubo {len(gdi.pages)}")
    check("6 etiqueta(s) en 2 página(s)" in detail, detail)
    left, pitch = 7.7 * app.DPM, 23.0 * app.DPM
    page1 = sorted(gdi.pages[0])
    want1 = sorted(r for c in range(4) for r in expected_rects(labels[c], left + c * pitch, 0))
    check(page1 == want1, "página 1: cuatro etiquetas distintas, cada una en su columna")
    page2 = sorted(gdi.pages[1])
    want2 = sorted(r for c in range(2) for r in expected_rects(labels[4 + c], left + c * pitch, 0))
    check(page2 == want2, "página 2: las dos que quedan, desde la primera columna")
    # copias idénticas (como siempre)
    elements = app.product_elements("PALITO CON COBERTURA")
    job = app.PrintJob(elements, 20, 30, 6, cfg, "grid")
    gdi, detail = run_gdi(job)
    check(len(gdi.pages) == 2 and len(gdi.pages[0]) == 4 * len(gdi.pages[1]) // 2, "copias idénticas: 4 + 2")
    # sin columnas: una página por etiqueta, cada una con su contenido
    job = app.PrintJob(labels[0], 20, 30, 1, dict(cfg, grid=False), "label", labels=labels[:3])
    gdi, _ = run_gdi(job)
    check(len(gdi.pages) == 3 and sorted(gdi.pages[2]) == sorted(expected_rects(labels[2], 0, 0)), "una por página")
    check(gdi.text_calls == 0, "toda la tipografía sale como rectángulos (no como texto TrueType)")


def test_etiquetas_dentro_del_papel():
    labels = [("producto", app.small_label_elements("x", app.codigo_producto(47, TODAY), "")),
              ("insumo", app.insumo_label_elements("Isopor 1/2 kg con nombre largo largo largo", app.codigo_insumo(6, TODAY), TODAY)),
              ("pack", app.serial_label_elements("PROMO PACK", "12 PALITOS", app.codigo_pack(9999999999))),
              ("id", app.serial_label_elements("ID DE BALDE", "", app.codigo_id_balde(9999999999)))]
    for item in app.test_set_labels(TODAY):
        labels.append((item["titulo"], item["elementos"]))
    for title, elements in labels:
        check(not app.elements_overflow(elements, 160, 240), f"{title}: algo se sale de la etiqueta de 2 x 3 cm")
        check(any(e[0] == "bar" for e in elements) and any(e[0] == "text" for e in elements), f"{title}: vacío")
    tags = [e for e in app.test_set_labels(TODAY)[0]["elementos"] if e[0] == "text" and e[4] == "PRUEBA"]
    check(len(tags) == 1, "las etiquetas de prueba llevan la marca PRUEBA")


def test_juego_de_prueba():
    items = app.test_set_labels(TODAY)
    check([i["titulo"] for i in items] == ["Producto (22)", "Grupo PADRE (29)", "Insumo (25)", "Promo Pack (23)",
                                          "ID de balde (24)"], [i["titulo"] for i in items])
    for item in items:
        r = read(item["codigo"])
        check(r["estado"] == "ok", f"{item['titulo']}: {item['codigo']} → {r['estado']}")
        digits = "".join(e[4] for e in item["elementos"] if e[0] == "text" and e[4].isdigit())
        check(digits == item["codigo"], f"{item['titulo']}: el número impreso no coincide ({digits})")


# --------------------------------------------------------------------------- lectura de las etiquetas dibujadas
def test_lectura_de_codigos_de_barras():
    cases = [("producto 22", app.small_label_elements("x", app.codigo_producto(1, TODAY), ""), app.codigo_producto(1, TODAY)),
             ("producto 22 (47)", app.small_label_elements("x", app.codigo_producto(47, date(2027, 3, 9)), ""),
              app.codigo_producto(47, date(2027, 3, 9))),
             ("PADRE 29", app.small_label_elements("x", "2900002800081", ""), "2900002800081"),
             ("insumo 25", app.insumo_label_elements("Isopor 1 kg", app.codigo_insumo(4, TODAY), TODAY), app.codigo_insumo(4, TODAY)),
             ("Promo Pack 23", app.serial_label_elements("PROMO PACK", "12 PALITOS", app.codigo_pack(123)), app.codigo_pack(123)),
             ("ID de balde 24", app.serial_label_elements("ID DE BALDE", "", app.codigo_id_balde(77)), app.codigo_id_balde(77))]
    for item in app.test_set_labels(TODAY):
        cases.append(("juego de prueba · " + item["titulo"], item["elementos"], item["codigo"]))
    for title, elements, code in cases:
        found = decode_label(elements)
        check(found == code, f"{title}: se esperaba {code} y el lector vio {found}")
    # todos los productos, con la fecha de varios años: lo impreso se lee igual que lo generado
    for number, name in app.products_by_number().items():
        for day in (date(2026, 11, 12), date(2027, 2, 28), date(2028, 12, 31)):
            code = app.codigo_producto(number, day)
            found = decode_label(app.small_label_elements(name, code, ""))
            check(found == code, f"{name} {day}: {found} != {code}")
    check(decode_ean13([0] * 160 if False else scanline_bits([("bar", 10, 0, 4, 50)], 10)) is None, "basura no es un EAN")


# --------------------------------------------------------------------------- interfaz (con tkinter falso)
@contextlib.contextmanager
def sandbox():
    """Las pruebas de interfaz escriben en una carpeta temporal: nunca tocan ajustes, contadores ni cantidades reales."""
    names = ("CONFIG_FILE", "COUNTERS_FILE", "SUPPLIES_FILE", "QUANTITIES_FILE", "LOG_FILE", "TRANSITION_FILE")
    saved = {name: getattr(app, name) for name in names}
    state = dict(app.CODE_STATE)
    quantities = dict(app.PRODUCT_QUANTITIES)
    with tempfile.TemporaryDirectory() as tmp:
        for name in names:
            setattr(app, name, Path(tmp) / saved[name].name)
        _tk_shim.ANSWERS.clear()
        del _tk_shim.DIALOG_CALLS[:]
        try:
            yield Path(tmp)
        finally:
            for name, value in saved.items():
                setattr(app, name, value)
            app.CODE_STATE.update(state)
            app.PRODUCT_QUANTITIES.clear()
            app.PRODUCT_QUANTITIES.update(quantities)
            _tk_shim.ANSWERS.clear()


def wait_for(condition, timeout=3.0):
    """Ejecuta los eventos pendientes hasta que se cumpla la condición (los envíos corren en otro hilo)."""
    end = time.time() + timeout
    while time.time() < end:
        _tk_shim.pump(2)
        if condition():
            return True
        time.sleep(0.01)
    return bool(condition())


def build_app():
    backend = app.SimBackend()
    window = app.EkaruApp(backend)
    wait_for(lambda: window.printers)
    return window, backend


def find(window, cls_name, since=0, **match):
    """Busca widgets creados (desde el índice «since») por clase y atributos de configuración."""
    out = []
    for widget in _tk_shim.CREATED[since:]:
        if type(widget).__name__ == cls_name and widget._alive and all(widget.kw.get(k) == v for k, v in match.items()):
            out.append(widget)
    return out


def label_texts(since=0):
    return [str(w.kw.get("text", "")) for w in _tk_shim.CREATED[since:] if type(w).__name__ == "Label" and w._alive]


def set_text(entry, text):
    entry.delete(0, "end")
    entry.insert(0, text)


def send_job(window, backend, action):
    """Hace «action» (que imprime) y espera a que el trabajo llegue al backend. Devuelve el programa TSPL."""
    count = len(backend.jobs)
    action()
    check(wait_for(lambda: len(backend.jobs) > count), "no se envió ningún trabajo")
    wait_for(lambda: not window._busy)
    return backend.jobs[-1][1].decode("cp1252")


def test_interfaz_arranca_y_imprime():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        check(window.selected == "PALITO CON COBERTURA", window.selected)
        info = window.sel_info.kw.get("text", "")
        check("22 0001" in info and "EAN-13" in info, f"el panel muestra el código nuevo: {info!r}")
        check(window.num_label.kw.get("text") == "Códigos nuevos (v2)", window.num_label.kw.get("text"))
        _tk_shim.ANSWERS["askyesnocancel"] = True
        window.config_data["aviso_numeracion"] = False
        program = send_job(window, backend, lambda: window.print_product("PALITO CON COBERTURA", 3))
        check(window.config_data["aviso_numeracion"] is True, "después del aviso no se vuelve a preguntar")
        check(any(c[0] == "askyesnocancel" for c in _tk_shim.DIALOG_CALLS), "se mostró el aviso de la numeración nueva")
        check("PRINT 1,3" in program, program[-60:])
        check(app.codigo_producto(1) in program, "la etiqueta lleva el código nuevo de hoy")
        saved = json.loads(app.CONFIG_FILE.read_text(encoding="utf-8"))
        check(saved["aviso_numeracion"] is True and saved["numeracion"] == "nueva", saved)
        # una segunda impresión ya no pregunta
        calls = len(_tk_shim.DIALOG_CALLS)
        send_job(window, backend, lambda: window.print_product("PALITO CON COBERTURA", 1))
        check(len(_tk_shim.DIALOG_CALLS) == calls, "el aviso se muestra una sola vez")
        check(not _tk_shim.UNKNOWN_CALLS, f"métodos de tkinter desconocidos: {sorted(_tk_shim.UNKNOWN_CALLS)}")


def test_ui_aviso_de_numeracion_no_y_cancelar():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        window.config_data["aviso_numeracion"] = False
        # «Cancelar»: no se imprime nada y el aviso queda pendiente
        _tk_shim.ANSWERS["askyesnocancel"] = None
        window.print_product("PALITO CON COBERTURA", 2)
        _tk_shim.pump(3)
        check(not backend.jobs and not window.config_data["aviso_numeracion"], "cancelar no imprime ni confirma")
        check(app.CODE_STATE["numeracion"] == "nueva", "cancelar deja la numeración como estaba")
        # «No»: se sigue con la anterior (2.0) y se imprime con ese código
        _tk_shim.ANSWERS["askyesnocancel"] = False
        program = send_job(window, backend, lambda: window.print_product("PALITO CON COBERTURA", 2))
        old = app.print_code("PALITO CON COBERTURA", mode=app.NUMERACION_ACTUAL)
        check(app.CODE_STATE["numeracion"] == "actual" and window.config_data["numeracion"] == "actual", "pasó a la anterior")
        check(old[:7] == "2000001" and old in program, f"imprime el código anterior {old}")
        check(window.num_label.kw.get("text") == "Códigos anteriores (2.0)", window.num_label.kw.get("text"))
        check(window.config_data["aviso_numeracion"] is True, "queda confirmado")
        saved = json.loads(app.CONFIG_FILE.read_text(encoding="utf-8"))
        check(saved["numeracion"] == "actual", saved)


def test_ui_fecha_de_elaboracion():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        today = date.today()
        three_ago = today - timedelta(days=3)
        window.date_var.set(three_ago.strftime("%d/%m/%Y"))
        window._commit_date()
        check(app.CODE_STATE["dia"] == three_ago, app.CODE_STATE)
        check(app.spaced_code(app.codigo_producto(1, three_ago)) in window.sel_info.kw["text"], "el panel usa la fecha")
        check(window.date_entry.kw.get("text_color") == app.WARN_COLOR, "fecha fijada = en ámbar")
        # lo que se rechaza no cambia nada y explica por qué
        for text, word in (((today + timedelta(days=1)).strftime("%d/%m/%Y"), "futura"), ("hola", "no válida"),
                           ("01/01/2010", "más de"), ("31/02/2026", "no válida")):
            window.date_var.set(text)
            window._commit_date()
            check(app.CODE_STATE["dia"] == three_ago, f"{text!r} no debía cambiar la fecha")
            check(word in window.status_label.kw["text"], (text, window.status_label.kw["text"]))
            check(window.date_var.get() == app.fecha_texto(three_ago), "el campo vuelve a la fecha válida")
        # el código que se imprime lleva esa fecha
        window.config_data["aviso_numeracion"] = True
        program = send_job(window, backend, lambda: window.print_product("PALITO CON COBERTURA", 1))
        check(app.codigo_producto(1, three_ago) in program, "imprime con la fecha fijada")
        window._date_today()
        check(app.CODE_STATE["dia"] is None and window.date_var.get() == app.fecha_texto(today), "«Hoy» vuelve a hoy")
        check(window.date_entry.kw.get("text_color") == app.TEXT, "hoy = color normal")
        # escribir la fecha de hoy equivale a «hoy» (no queda fijada)
        window.date_var.set(three_ago.strftime("%d-%m-%y"))
        window._commit_date()
        check(app.CODE_STATE["dia"] == three_ago, "acepta 12-11-26")
        window.date_var.set(today.strftime("%d.%m.%Y"))
        window._commit_date()
        check(app.CODE_STATE["dia"] is None, "la fecha de hoy no queda fijada")
        # cambio de día a medianoche: si no está fijada, el campo y los códigos pasan al día nuevo
        window._painted_day = today - timedelta(days=1)
        window._day_tick()
        check(window._painted_day == today and "Cambió el día" in window.status_label.kw["text"], "avance de día")


def test_ui_busqueda_por_codigo():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        names = app.products_by_number()
        week_ago = date.today() - timedelta(days=7)
        for number in (3, 17, 47):
            window.search_var.set(app.spaced_code(app.codigo_producto(number, week_ago)))  # como lo escribe una persona
            window._refresh_list()
            check(window.visible == [names[number]] and window.selected == names[number], (number, window.visible))
            window.search_var.set(app.codigo_producto(number, week_ago))  # como lo tipea el escáner
            window._refresh_list()
            check(window.visible == [names[number]], (number, window.visible))
        # código viejo (catálogo original de e-Karú): también encuentra el producto
        window.search_var.set(app.CODIGOS_ANTERIORES[5])
        window._refresh_list()
        check(window.visible == [names[5]], window.visible)
        # búsqueda parcial por el prefijo de la numeración nueva y por texto normal
        window.search_var.set("220007")
        window._refresh_list()
        check(window.visible == [names[7]], window.visible)
        window.search_var.set("palito")
        window._refresh_list()
        check(len(window.visible) > 1 and all("PALITO" in n.upper() or "palito" in n.lower() for n in window.visible[:3]),
              window.visible[:3])
        window.search_var.set("")
        window._refresh_list()
        check(len(window.visible) == len(app.PRODUCTS), "sin búsqueda: todo el catálogo")


def test_ui_ajustes_numeracion():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        since = len(_tk_shim.CREATED)
        window.open_settings()
        menus = [m for m in find(window, "OptionMenu", since) if m.kw.get("values") == list(app.NUMERACION_TEXTO.values())]
        check(len(menus) == 1, "hay un menú de numeración en Ajustes")
        menus[0].choose(app.NUMERACION_TEXTO["actual"])
        find(window, "Button", since, text="Guardar")[-1].invoke()
        check(app.CODE_STATE["numeracion"] == "actual" and window.config_data["numeracion"] == "actual", "aplicó")
        check(window.config_data["aviso_numeracion"] is True, "elegir en Ajustes equivale a confirmar el aviso")
        check(window.num_label.kw.get("text") == "Códigos anteriores (2.0)", "barra lateral")
        saved = json.loads(app.CONFIG_FILE.read_text(encoding="utf-8"))
        check(saved["numeracion"] == "actual" and saved["aviso_numeracion"] is True, saved)
        check(app.print_code("PALITO CON COBERTURA")[:7] == "2000001", "la lista usa los códigos anteriores")
        check("2 000001" in window.sel_info.kw["text"], window.sel_info.kw["text"])
        # y de vuelta a la nueva
        since = len(_tk_shim.CREATED)
        window.open_settings()
        menus = [m for m in find(window, "OptionMenu", since) if m.kw.get("values") == list(app.NUMERACION_TEXTO.values())]
        check(menus[0].get() == app.NUMERACION_TEXTO["actual"], "Ajustes muestra la numeración activa")
        menus[0].choose(app.NUMERACION_TEXTO["nueva"])
        find(window, "Button", since, text="Guardar")[-1].invoke()
        check(app.CODE_STATE["numeracion"] == "nueva" and "22 0001" in window.sel_info.kw["text"], "volvió a la nueva")


def test_ui_lector():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        since = len(_tk_shim.CREATED)
        window.open_reader()
        box = find(window, "Entry", since, placeholder_text="Código de barras…")[0]
        today = date.today()
        cases = (
            (app.spaced_code(app.codigo_producto(1, today - timedelta(days=2))), "✔ CÓDIGO CORRECTO · Producto terminado",
             "Producto: PALITO CON COBERTURA"),
            (app.CODIGOS_ANTERIORES[1], "▲ CÓDIGO VIEJO", "Nº de producto: 0001"),
            ("2000001100010", "✖ RECHAZADO · Dígito de control incorrecto", "debería ser"),
            (app.codigo_producto(1, today + timedelta(days=5)), "✖ RECHAZADO · Producto terminado", "futura"),
            (app.codigo_pack(7), "✔ CÓDIGO CORRECTO · Promo Pack", "Nº de paquete: 7"),
            (app.codigo_insumo(2, today), "✔ CÓDIGO CORRECTO · Insumo interno", "Insumo: Conos"),
            ("7891000100103", "● CÓDIGO EXTERNO · EAN de proveedor", "equivalencias"),
            ("2 000001 01105 00041129 3", "✔ CÓDIGO CORRECTO · Balde con transacción", "Transacción: 00041129"),
        )
        for code, headline, detail in cases:
            box.kw["textvariable"].set(code)
            box._bindings["<Return>"](None)
            texts = label_texts(since)
            check(any(t.startswith(headline) for t in texts),
                  (code, headline, [t for t in texts if "CÓDIGO" in t or "RECHAZADO" in t]))
            check(any(detail in t for t in texts), (code, detail))
            check(box.kw["textvariable"].get() == "", "el campo queda libre para el próximo escaneo")
        hist = find(window, "Textbox", since)[0].get()
        check(hist.count("\n") == len(cases) and "Promo Pack" in hist, hist)
        # exportar la tabla desde el lector
        target = app.TRANSITION_FILE.parent / "salida.csv"
        _tk_shim.ANSWERS["asksaveasfilename"] = str(target)
        find(window, "Button", since, text="Tabla de transición (CSV)")[0].invoke()
        rows = list(csv.reader(target.read_bytes().decode("utf-8-sig").splitlines(), delimiter=";"))
        check(len(rows) == 1 + 47 + 9, len(rows))
        check(not _tk_shim.UNKNOWN_CALLS, f"métodos de tkinter desconocidos: {sorted(_tk_shim.UNKNOWN_CALLS)}")


def test_ui_etiquetas_internas():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        since = len(_tk_shim.CREATED)
        window.open_internal_labels()
        entries = find(window, "Entry", since)
        check(len(entries) == 7, f"campos: {len(entries)}")
        received, supply_qty, new_supply, pack_start, pack_qty, balde_start, balde_qty = entries
        segmented = find(window, "SegmentedButton", since)[0]
        print_button = find(window, "Button", since, text="Imprimir")[0]
        check(segmented.kw["values"] == ["Insumo (25)", "Promo Pack (23)", "ID de balde (24)"], segmented.kw["values"])
        # insumo: 2 etiquetas iguales con nombre, código 25 y fecha de recepción
        set_text(supply_qty, "2")
        program = send_job(window, backend, print_button.invoke)
        code = app.codigo_insumo(1, date.today())
        check("PRINT 1,2" in program and program.count("PRINT") == 1, program[-40:])
        check("FRUTA A GRANEL" in program and app.spaced_code(code) not in program and code in program, "etiqueta de insumo")
        # Promo Pack: serie numerada; el contador avanza solo si el envío salió bien
        segmented.choose("Promo Pack (23)")
        check(app.load_counters() == {"pack": 0, "balde": 0}, "contadores en cero")
        set_text(pack_qty, "5")
        program = send_job(window, backend, print_button.invoke)
        check(program.count("PRINT 1,1") == 5 and program.count("PRINT") == 5, "5 etiquetas, 5 PRINT")
        for n in range(1, 6):
            check(app.codigo_pack(n) in program, f"falta el Nº {n}")
        check(wait_for(lambda: app.load_counters()["pack"] == 5), app.load_counters())
        check(pack_start.get() == "6", f"la ventana propone el próximo número: {pack_start.get()}")
        # reimprimir una etiqueta vieja no hace retroceder el contador
        set_text(pack_start, "3")
        set_text(pack_qty, "1")
        program = send_job(window, backend, print_button.invoke)
        check(app.codigo_pack(3) in program, "reimpresión del Nº 3")
        wait_for(lambda: True)
        check(app.load_counters()["pack"] == 5, app.load_counters())
        # ID de balde (plan B)
        segmented.choose("ID de balde (24)")
        set_text(balde_qty, "2")
        program = send_job(window, backend, print_button.invoke)
        check(app.codigo_id_balde(1) in program and app.codigo_id_balde(2) in program, "ID de balde 1 y 2")
        check(wait_for(lambda: app.load_counters() == {"pack": 5, "balde": 2}), app.load_counters())
        # si el envío falla, el contador NO avanza
        backend.fail_with = "impresora sin papel"
        set_text(balde_start, "3")
        count = len(backend.jobs)
        print_button.invoke()
        wait_for(lambda: not window._busy and "sin papel" in window.status_label.kw["text"])
        check(len(backend.jobs) == count and app.load_counters()["balde"] == 2, "envío fallido: contador intacto")
        backend.fail_with = None
        wait_for(lambda: not window._busy)
        # entradas inválidas: se explica y no se envía nada
        count = len(backend.jobs)
        segmented.choose("Promo Pack (23)")
        set_text(pack_qty, "0")
        print_button.invoke()
        check(any("Cantidad: tiene que estar entre 1 y 999" in t for t in label_texts(since)), "mensaje de cantidad")
        set_text(pack_qty, "abc")
        print_button.invoke()
        check(any("Cantidad: escribí un número entero" in t for t in label_texts(since)), "mensaje de número")
        set_text(pack_qty, "1")
        set_text(pack_start, str(app.MAX_NUMERO_CORRELATIVO + 1))
        print_button.invoke()
        check(any("Nº inicial: tiene que estar entre" in t for t in label_texts(since)), "mensaje de rango")
        check(len(backend.jobs) == count, "nada se envió con datos inválidos")
        # insumo nuevo
        segmented.choose("Insumo (25)")
        set_text(new_supply, "Cobertura de chocolate")
        find(window, "Button", since, text="Agregar")[0].invoke()
        names = [n for _, n in app.load_supplies()]
        check(names[-1] == "Cobertura de chocolate" and len(names) == 7, names)
        menu = find(window, "OptionMenu", since)[0]
        check("0007 · Cobertura de chocolate" in menu.kw["values"] and menu.get() == "0007 · Cobertura de chocolate", menu.kw["values"])
        set_text(new_supply, "conos")
        find(window, "Button", since, text="Agregar")[0].invoke()
        check(any("Ya existe" in t for t in label_texts(since)), "no repite nombres")
        set_text(received, "99/99/2026")
        print_button.invoke()
        check(any("Fecha de recepción" in t for t in label_texts(since)), "fecha de recepción inválida")
        check(not _tk_shim.UNKNOWN_CALLS, f"métodos de tkinter desconocidos: {sorted(_tk_shim.UNKNOWN_CALLS)}")


def test_ui_juego_de_prueba_e_informe():
    with sandbox():
        app.set_numbering(app.NUMERACION_NUEVA)
        window, backend = build_app()
        program = send_job(window, backend, window.print_test_set)
        check(program.count("PRINT 1,1") == 5 and program.count("PRINT") == 5, "5 etiquetas")
        for item in app.test_set_labels(date.today()):
            check(item["codigo"] in program or app.spaced_code(item["codigo"]) in program, item["titulo"])
        wait_for(lambda: "Juego de prueba impreso" in window.status_label.kw["text"])
        check("Juego de prueba impreso" in window.status_label.kw["text"], window.status_label.kw["text"])
        report = window.build_report()
        for text in ("Numeración de los códigos: Nueva (v2)", "Fecha de elaboración en los códigos:", "Contadores: Promo Pack 0"):
            check(text in report, text)
        since = len(_tk_shim.CREATED)
        window.open_diagnostics()
        target = app.TRANSITION_FILE.parent / "diag.csv"
        _tk_shim.ANSWERS["asksaveasfilename"] = str(target)
        find(window, "Button", since, text="Tabla de transición (CSV)")[0].invoke()
        check(target.exists() and target.stat().st_size > 1000, "el diagnóstico exporta la tabla de transición")
        # los botones del diagnóstico siguen en una grilla 3 x 3
        buttons = [b for b in find(window, "Button", since) if b._gm == "grid"]
        check(len(buttons) == 9, len(buttons))


# --------------------------------------------------------------------------- ejecución
def main():
    tests = [(name, func) for name, func in globals().items() if name.startswith("test_") and callable(func)]
    failed = 0
    for name, func in tests:
        try:
            func()
            print(f"  OK   {name}")
        except Exception:  # noqa: BLE001
            failed += 1
            print(f"  FALLA {name}")
            traceback.print_exc()
    unknown = sorted(_tk_shim.UNKNOWN_CALLS)
    if unknown:
        print("  (métodos de tkinter que el sustituto no conoce y ignoró:", ", ".join(unknown), ")")
    print(f"\n{len(tests) - failed} de {len(tests)} pruebas correctas.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
