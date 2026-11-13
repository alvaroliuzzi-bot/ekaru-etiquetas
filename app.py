# -*- coding: utf-8 -*-
"""E-Karú Etiquetas — generador e impresor de etiquetas (HPRT HT400, TSPL/RAW).

Cambios v2.1: numeración de códigos v2 del Informe final (22 = producto + Nº de producto + fecha de
elaboración, 23 = Promo Pack, 24 = ID de balde, 25 = insumo interno, 29 = grupo PADRE sin cambios),
fecha de elaboración editable, etiquetas internas con numeración correlativa, lector de códigos con
las 7 reglas del Informe, juego de prueba, tabla de transición (CSV) y búsqueda por código escaneado.
La numeración anterior sigue disponible en Ajustes.

Cambios v2.0: impresión RAW directa por ctypes (sin pywin32), diagnóstico de
impresión, lista de catálogo virtualizada (sin parpadeo), vista previa idéntica
a lo impreso y códigos de barras dibujados desde los mismos bits verificados.
El catálogo (PRODUCTS) y los códigos de barras no se modificaron.
"""
import csv
import ctypes
import json
import os
import queue
import subprocess
import sys
import tempfile
import textwrap
import threading
import time
import traceback
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import font as tkfont


APP_NAME = "E-Karú Etiquetas"
VERSION = "v2.1.0"
BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DESIGNS_DIR = BASE_DIR / "diseños"
QUANTITIES_FILE = BASE_DIR / "cantidades_impresion.json"
CONFIG_FILE = BASE_DIR / "ekaru_config.json"
LOG_FILE = BASE_DIR / "ekaru_impresion.log"
ERROR_LOG = BASE_DIR / "ekaru_startup_error.log"
OUTPUT_DIR = BASE_DIR / "salida_impresion"
COUNTERS_FILE = BASE_DIR / "contadores.json"            # últimos números de Promo Pack (23) e ID de balde (24)
SUPPLIES_FILE = BASE_DIR / "insumos.json"               # lista de insumos internos (25): Nº + nombre
TRANSITION_FILE = BASE_DIR / "tabla_transicion_codigos.csv"
try:
    DESIGNS_DIR.mkdir(exist_ok=True)
except OSError:
    pass


# --------------------------------------------------------------------------- #
#  Catálogo (vinculante) — copiado sin cambios de la versión anterior
# --------------------------------------------------------------------------- #
FORMATS = {
    "Etiqueta pequeña (3 x 2 cm)": (30, 20),
    "Etiqueta grande (10 x 8 cm)": (100, 80),
}
PRODUCTS = {
    "PALITO CON COBERTURA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000001100011", "detail": "Palito"},
    "PALITO FRUTILLA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000002100010", "detail": "Palito"},
    "PALITO FRUTILLA AL AGUA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000003100019", "detail": "Palito"},
    "PALITO FRUTILLA C/ RELLENO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000004100018", "detail": "Palito"},
    "PALITO GRANIZADO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000005100017", "detail": "Palito"},
    "PALITO MBURUCUYA A LA CREMA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000006100016", "detail": "Palito"},
    "PALITO MENTA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000007100015", "detail": "Palito"},
    "PALITO OREO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000008100014", "detail": "Palito"},
    "PALITO PISTACHO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000009100013", "detail": "Palito"},
    "ALFAJOR CHICO 4 LECHES": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000010200023", "detail": "Alfajor"},
    "ALFAJOR CHICO DULCE DE LECHE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000011200022", "detail": "Alfajor"},
    "ALFAJOR DE MAIZENA*6": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000012200021", "detail": "Alfajor"},
    "ALFAJOR GRANDE ALMENDRAS": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000013200020", "detail": "Alfajor"},
    "ALFAJOR GRANDE DL C/MICROGALLETITAS": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000014200029", "detail": "Alfajor"},
    "ALFAJOR GRANDE PISTACHO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000015200028", "detail": "Alfajor"},
    "ALFAJOR MAICENA DELIFREE*4": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000016200027", "detail": "Alfajor"},
    "ALFAJOR CON CHOCOLATE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000017200026", "detail": "Alfajor"},
    "BROWNIE BANDEJA PACK": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000018300039", "detail": "Brownie"},
    "BROWNIE PORCION": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000019300038", "detail": "Brownie"},
    "BROWNIE DELIFREE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000020300034", "detail": "Brownie"},
    "BROWNIE CON HELADO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000021300033", "detail": "1 porción + 100 g helado"},
    "PASTAFROLA PORCION": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000022400046", "detail": "Postre"},
    "GALLETITA COOKIES RELLENO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000023400045", "detail": "Postre"},
    "BROWNIE GATITO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000024400044", "detail": "Postre"},
    "TARTA 3 QUESOS": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000025500057", "detail": "Tarta"},
    "TARTA ACELGA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000026500056", "detail": "Tarta"},
    "TARTA JAMON Y QUESO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000027500055", "detail": "Tarta"},
    "TARTA DE POLLO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000028500054", "detail": "Tarta"},
    "CONO FAN": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000029600067", "detail": "Cono"},
    "CONO FAN 2*1": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000030600063", "detail": "Cono"},
    "CONO FAN 3X25": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000031600062", "detail": "Cono"},
    "TORTA AMARENA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000032700075", "detail": "Torta"},
    "TORTA CON FRUTILLA PREMIUM": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000033700074", "detail": "Torta"},
    "TORTA BROWNIE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000034700073", "detail": "Torta"},
    "TORTA CHOCOLATE CON FRUTILLA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000035700072", "detail": "Torta"},
    "TORTA MATILDA CHOCOLATE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000036700071", "detail": "Torta"},
    "TORTA CREMA CON ALMENDRAS": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000037700070", "detail": "Torta"},
    "TORTA DDL CON CHOCOLATE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000038700079", "detail": "Torta"},
    "TORTA DE MBURUCUYA CON CHOCOLATE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000039700078", "detail": "Torta"},
    "TORTA HELADA RED VELVET": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000040700074", "detail": "Torta"},
    "TORTA INSIGNIA GLACE FRUTILLA": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000041700073", "detail": "Torta"},
    "TORTA MBURUCUYA CON FRUTILLA PREMIUM": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000042700072", "detail": "Torta"},
    "TORTA OREO PREMIUM": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000043700071", "detail": "Torta"},
    "TORTA PIE DE LIMON CON FRUTOS DEL BOSQUE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000044700070", "detail": "Torta"},
    "TORTA AMARENA DELIFREE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000045700079", "detail": "Torta DeliFREE"},
    "TORTA TRAMONTANA DELIFREE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000046700078", "detail": "Torta DeliFREE"},
    "TORTA CHOCOLATE FRUTILLA DELIFREE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2000047700077", "detail": "Torta DeliFREE"},
    "Cono 3": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900009800084", "detail": "Grupo PADRE · 3 bolas - 180 g"},
    "Cono 2": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900008800085", "detail": "Grupo PADRE · 2 bolas - 120 g"},
    "Vaso 1": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900005800088", "detail": "Grupo PADRE · 1 bola - 60 g"},
    "Vaso 2": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900006800087", "detail": "Grupo PADRE · 2 bolas - 120 g"},
    "Vaso 3": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900007800086", "detail": "Grupo PADRE · 3 bolas - 180 g"},
    "Milkshake CHICO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900001800082", "detail": "Grupo Vasos/Milkshakes/Smoothies · 2 bolas - 120 g"},
    "Milkshake GRANDE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900001800082", "detail": "Grupo Vasos/Milkshakes/Smoothies · 3 bolas - 180 g"},
    "Smoothie CHICO": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900001800082", "detail": "Grupo Vasos/Milkshakes/Smoothies · 2 bolas - 120 g"},
    "Smoothie GRANDE": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900001800082", "detail": "Grupo Vasos/Milkshakes/Smoothies · 3 bolas - 180 g"},
    "Grupo PADRE Vasos/Milkshakes/Smoothies": {
        "format": "Etiqueta pequeña (3 x 2 cm)",
        "barcode": "2900001800082",
        "detail": "Grupo PADRE según catálogo",
    },
    "Grupo PADRE 1 kg": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900002800081", "detail": "Pote 1 kg"},
    "Grupo PADRE 1/2 kg": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900003800080", "detail": "Pote 1/2 kg"},
    "Grupo PADRE 1/4 kg": {"format": "Etiqueta pequeña (3 x 2 cm)", "barcode": "2900004800089", "detail": "Pote 1/4 kg"},
    "Mascarpone Cheesecake frutos del bosque - Balde 6 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 6 kg",
    },
    "Menta Granizado - Balde 6 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 6 kg",
    },
    "Pistacho - Balde 6 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 6 kg",
    },
    "Mascarpone Cheesecake frutos del bosque - Balde 3 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 3 kg",
    },
    "Menta Granizado - Balde 3 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 3 kg",
    },
    "Pistacho - Balde 3 kg": {
        "format": "Etiqueta grande (10 x 8 cm)",
        "barcode": "",
        "detail": "Contenido neto: 3 kg",
    },
}


def is_valid_ean13(value):
    value = str(value).strip()
    if len(value) != 13 or not value.isdigit():
        return False
    total = sum(int(digit) * (1 if index % 2 == 0 else 3) for index, digit in enumerate(value[:12]))
    return (10 - total % 10) % 10 == int(value[-1])


def ean13_bits(value):
    if not is_valid_ean13(value):
        return None
    left = {
        "0": "0001101", "1": "0011001", "2": "0010011", "3": "0111101", "4": "0100011",
        "5": "0110001", "6": "0101111", "7": "0111011", "8": "0110111", "9": "0001011",
    }
    left_g = {
        "0": "0100111", "1": "0110011", "2": "0011011", "3": "0100001", "4": "0011101",
        "5": "0111001", "6": "0000101", "7": "0010001", "8": "0001001", "9": "0010111",
    }
    right = {
        "0": "1110010", "1": "1100110", "2": "1101100", "3": "1000010", "4": "1011100",
        "5": "1001110", "6": "1010000", "7": "1000100", "8": "1001000", "9": "1110100",
    }
    parity = {
        "0": "LLLLLL", "1": "LLGLGG", "2": "LLGGLG", "3": "LLGGGL", "4": "LGLLGG",
        "5": "LGGLLG", "6": "LGGGLL", "7": "LGLGLG", "8": "LGLGGL", "9": "LGGLGL",
    }[value[0]]
    bits = "101"
    for digit, code in zip(value[1:7], parity):
        bits += left[digit] if code == "L" else left_g[digit]
    bits += "01010"
    bits += "".join(right[digit] for digit in value[7:])
    return bits + "101"


def code128_bits(value):
    value = str(value)
    if not value:
        return ""
    if value.isdigit() and len(value) >= 2:
        codes = [105]
        pairs = len(value) - len(value) % 2
        codes.extend(int(value[index:index + 2]) for index in range(0, pairs, 2))
        if len(value) % 2:
            codes.extend((100, ord(value[-1]) - 32))
    else:
        codes = [104]
        codes.extend(ord(char) - 32 for char in value)
    checksum = (codes[0] + sum(index * code for index, code in enumerate(codes[1:], 1))) % 103
    codes.append(checksum)
    codes.append(106)
    patterns = (
        "11011001100", "11001101100", "11001100110", "10010011000", "10010001100",
        "10001001100", "10011001000", "10011000100", "10001100100", "11001001000",
        "11001000100", "11000100100", "10110011100", "10011011100", "10011001110",
        "10111001100", "10011101100", "10011100110", "11001110010", "11001011100",
        "11001001110", "11011100100", "11001110100", "11101101110", "11101001100",
        "11100101100", "11100100110", "11101100100", "11100110100", "11100110010",
        "11011011000", "11011000110", "11000110110", "10100011000", "10001011000",
        "10001000110", "10110001000", "10001101000", "10001100010", "11010001000",
        "11000101000", "11000100010", "10110111000", "10110001110", "10001101110",
        "10111011000", "10111000110", "10001110110", "11101110110", "11010001110",
        "11000101110", "11011101000", "11011100010", "11011101110", "11101011000",
        "11101000110", "11100010110", "11101101000", "11101100010", "11100011010",
        "11101111010", "11001000010", "11110001010", "10100110000", "10100001100",
        "10010110000", "10010000110", "10000101100", "10000100110", "10110010000",
        "10110000100", "10011010000", "10011000010", "10000110100", "10000110010",
        "11000010010", "11001010000", "11110111010", "11000010100", "10001111010",
        "10100111100", "10010111100", "10010011110", "10111100100", "10011110100",
        "10011110010", "11110100100", "11110010100", "11110010010", "11011011110",
        "11011110110", "11110110110", "10101111000", "10100011110", "10001011110",
        "10111101000", "10111100010", "11110101000", "11110100010", "10111011110",
        "10111101110", "11101011110", "11110101110", "11010000100", "11010010000",
        "11010011100",
    )
    return "".join(patterns[code] for code in codes[:-1]) + "11000111010" + "11"



def _load_quantities():
    try:
        data = json.loads(QUANTITIES_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {}
    if not isinstance(data, dict):
        return {}
    result = {}
    for name, value in data.items():
        try:
            number = int(value)
        except (TypeError, ValueError):
            continue
        if 1 <= number <= 999:
            result[str(name)] = number
    return result


PRODUCT_QUANTITIES = _load_quantities()


def spaced_code(value):
    """Dígitos agrupados como en el Informe: 22 0001 261110 4 · 23 0000000001 6 · 2 000002 26307 4 ·
    balde con transacción (21 dígitos) 2 000001 01105 00041129 3."""
    if len(value) == 21 and value.isdigit():
        return f"{value[0]} {value[1:7]} {value[7:12]} {value[12:20]} {value[20]}"
    if len(value) == 13 and value.isdigit():
        if value[:2] in ("22", "25"):
            return f"{value[:2]} {value[2:6]} {value[6:12]} {value[12]}"
        if value[:2] in ("23", "24"):
            return f"{value[:2]} {value[2:12]} {value[12]}"
        return f"{value[0]} {value[1:7]} {value[7:12]} {value[12]}"
    return value


def barcode_kind(value):
    if not value:
        return "Sin código de barras"
    return "EAN-13" if is_valid_ean13(value) else "Code 128"


def ean13_check_digit(first12):
    total = sum(int(d) * (1 if i % 2 == 0 else 3) for i, d in enumerate(first12))
    return str((10 - total % 10) % 10)


def date_field(day=None):
    """Fecha de elaboración como AADDD (año de 2 cifras + día del año, 3 cifras). 3-nov-2026 → 26307."""
    day = day or datetime.now()
    return f"{day.year % 100:02d}{day.timetuple().tm_yday:03d}"


# Productos con código hijo propio (cada uno su código; la categoría los agrupa).
# Vasos, Cono 2/3, Milkshakes, Smoothies y potes son solo grupos PADRE: código fijo, sin fecha.
DATED_CATEGORIES = {"Palitos", "Alfajores", "Brownies", "Postres", "Tartas", "Tortas"}


# --------------------------------------------------------------------------- #
#  Numeración v2 (Informe final, página «Códigos de barras»): un prefijo por familia dentro de 20–29
#    20  balde (lo imprime la balanza)  2 + sabor (6) + peso en g (5) + control. Sin cambios.
#    22  producto terminado             22 + Nº de producto (4) + fecha de elaboración AAMMDD (6) + control
#    23  Promo Pack                     23 + número de paquete (10) + control
#    24  ID de balde (plan B)           24 + número correlativo (10) + control
#    25  insumo interno                 25 + Nº de insumo (4) + fecha de recepción AAMMDD (6) + control
#    29  grupo PADRE                    se mantiene (2 9000NN 80008, con el último dígito bien calculado)
#    21 y 26–28 reservados · cualquier otro = EAN del proveedor
#  Todos los códigos internos tienen 13 dígitos con control EAN-13 válido (el balde con serie es Code 128 de 21).
# --------------------------------------------------------------------------- #
NUMERACION_NUEVA = "nueva"
NUMERACION_ACTUAL = "actual"
NUMERACION_TEXTO = {
    NUMERACION_NUEVA: "Nueva (v2): 22 + Nº de producto + fecha",
    NUMERACION_ACTUAL: "Anterior: 2 + producto + fecha AADDD",
}
CODE_STATE = {"numeracion": NUMERACION_NUEVA, "dia": None}  # dia=None → siempre la fecha de hoy
MAX_NUMERO_CORRELATIVO = 9_999_999_999
MAX_DIAS_ATRAS = 400   # la fecha de elaboración a mano no puede ser más vieja (evita errores de año al escribirla)

# Códigos de producto del catálogo de e-Karú ANTES de corregir el dígito de control (codigos_cambios.csv).
# Nº de producto → código anterior. El lector los acepta solo como «tabla de transición».
CODIGOS_ANTERIORES = {
    1: "2000001100014", 2: "2000002100017", 3: "2000003100012", 4: "2000004100019",
    5: "2000005100015", 6: "2000006100011", 7: "2000007100018", 8: "2000008100013",
    9: "2000009100016", 10: "2000010200021", 11: "2000011200028", 12: "2000012200025",
    13: "2000013200022", 14: "2000014200029", 15: "2000015200026", 16: "2000016200023",
    17: "2000017200020", 18: "2000018300037", 19: "2000019300034", 20: "2000020300031",
    21: "2000021300038", 22: "2000022400045", 23: "2000023400042", 24: "2000024400049",
    25: "2000025500056", 26: "2000026500053", 27: "2000027500050", 28: "2000028500057",
    29: "2000029600064", 30: "2000030600061", 31: "2000031600068", 32: "2000032700073",
    33: "2000033700076", 34: "2000034700079", 35: "2000035700072", 36: "2000036700075",
    37: "2000037700078", 38: "2000038700071", 39: "2000039700074", 40: "2000040700070",
    41: "2000041700073", 42: "2000042700076", 43: "2000043700079", 44: "2000044700072",
    45: "2000045700075", 46: "2000046700078", 47: "2000047700071",
}
# Grupos PADRE: código anterior (último dígito mal calculado) → código con control válido (el que imprime esta app).
PADRE_ANTERIORES = {
    "2900009800087": "2900009800084",
    "2900008800080": "2900008800085",
    "2900005800089": "2900005800088",
    "2900006800086": "2900006800087",
    "2900007800083": "2900007800086",
    "2900001800081": "2900001800082",
    "2900002800088": "2900002800081",
    "2900003800085": "2900003800080",
    "2900004800082": "2900004800089",
}
# Insumos internos de ejemplo (Nº definitivo = el que acuerden Glacé y e-Karú). Se pueden agregar más en la app.
INSUMOS_INICIALES = (
    (1, "Fruta a granel"), (2, "Conos"), (3, "Vasos"),
    (4, "Isopor 1 kg"), (5, "Isopor 1/2 kg"), (6, "Isopor 1/4 kg"),
)


def _as_date(day):
    """date a partir de date/datetime/None (None = hoy)."""
    if day is None:
        return date.today()
    return day.date() if isinstance(day, datetime) else day


def aammdd(day=None):
    """Fecha como AAMMDD (6 cifras). 10-nov-2026 → 261110."""
    day = _as_date(day)
    return f"{day.year % 100:02d}{day.month:02d}{day.day:02d}"


def parse_aammdd(text):
    """AAMMDD → date, o None si no es una fecha real (año 2000-2099)."""
    if len(text) != 6 or not text.isdigit():
        return None
    try:
        return date(2000 + int(text[:2]), int(text[2:4]), int(text[4:]))
    except ValueError:
        return None


def parse_user_date(text):
    """Fecha escrita por una persona: 12/11/2026, 12-11-26, 12.11.2026 o 12112026. None si no se entiende."""
    text = str(text).strip()
    for sep in ("-", "."):
        text = text.replace(sep, "/")
    if "/" in text:
        parts = text.split("/")
    elif text.isdigit() and len(text) in (6, 8):
        parts = [text[:2], text[2:4], text[4:]]
    else:
        return None
    if len(parts) != 3 or not all(p.strip().isdigit() for p in parts):
        return None
    day_, month, year = (int(p) for p in parts)
    if year < 100:
        year += 2000
    try:
        return date(year, month, day_)
    except ValueError:
        return None


def mod10_check(payload):
    """Dígito de control GS1 (módulo 10) de cualquier largo: pesos 3,1,3,1… contando desde la derecha."""
    total = sum(int(d) * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(payload)))
    return str((10 - total % 10) % 10)


def with_check(first12):
    return first12 + ean13_check_digit(first12)


def _number_in_range(value, low, high, what):
    try:
        number = int(value)
    except (TypeError, ValueError):
        raise ValueError(f"{what}: tiene que ser un número entero.") from None
    if not low <= number <= high:
        raise ValueError(f"{what}: tiene que estar entre {low} y {high}.")
    return number


def codigo_producto(numero, day=None):
    """22 + Nº de producto (4) + fecha de elaboración AAMMDD (6) + control."""
    return with_check(f"22{_number_in_range(numero, 1, 9999, 'Nº de producto'):04d}{aammdd(day)}")


def codigo_insumo(numero, day=None):
    """25 + Nº de insumo (4) + fecha de recepción AAMMDD (6) + control."""
    return with_check(f"25{_number_in_range(numero, 1, 9999, 'Nº de insumo'):04d}{aammdd(day)}")


def codigo_pack(numero):
    """23 + número de paquete (10) + control."""
    return with_check(f"23{_number_in_range(numero, 0, MAX_NUMERO_CORRELATIVO, 'Nº de paquete'):010d}")


def codigo_id_balde(numero):
    """24 + número correlativo (10) + control."""
    return with_check(f"24{_number_in_range(numero, 0, MAX_NUMERO_CORRELATIVO, 'Nº de balde'):010d}")


def product_number(name):
    """Nº de producto del catálogo (Palito con cobertura = 1 … 47). None: grupo PADRE o balde."""
    code = PRODUCTS[name]["barcode"]
    if len(code) == 13 and code[:3] == "200" and code.isdigit():
        return int(code[1:7])
    return None


def products_by_number():
    """{Nº: nombre} de los productos con código 22 (los 47 del catálogo)."""
    result = {}
    for name in PRODUCTS:
        number = product_number(name)
        if number is not None:
            result.setdefault(number, name)
    return result


def padre_by_code():
    """{código PADRE: [nombres]}."""
    result = {}
    for name, data in PRODUCTS.items():
        if data["barcode"][:2] == "29":
            result.setdefault(data["barcode"], []).append(name)
    return result


def padre_display(code):
    """Nombre legible de un grupo PADRE (prefiere «Grupo PADRE …»)."""
    names = padre_by_code().get(code, [])
    if not names:
        return None
    pick = next((n for n in names if n.upper().startswith("GRUPO PADRE")), names[0])
    detail = PRODUCTS[pick]["detail"]
    return pick if detail.startswith("Grupo PADRE") else f"{pick} ({detail})"


def print_code(name, day=None, mode=None):
    """Código que se imprime para un producto del catálogo.

    Numeración nueva (v2): productos = 22 + Nº de producto (4) + fecha de elaboración AAMMDD + control.
    Los grupos PADRE (29…) y los baldes se mantienen. Numeración anterior: 2 + producto (6) + fecha AADDD
    (5) + control, como en la versión 2.0. «day» = fecha de elaboración (por defecto la elegida en la app o hoy),
    «mode» = numeración (por defecto la activa).
    """
    code = PRODUCTS[name]["barcode"]
    if not code:
        return code
    day = day or CODE_STATE["dia"]
    if (mode or CODE_STATE["numeracion"]) == NUMERACION_NUEVA:
        number = product_number(name)
        return codigo_producto(number, day) if number is not None else code
    if not CODE_WITH_DATE or not is_valid_ean13(code):
        return code
    upper = name.upper()
    if product_category(name) in DATED_CATEGORIES or upper.startswith("CONO FAN"):
        first12 = code[:7] + date_field(day)
        return first12 + ean13_check_digit(first12)
    return code


def set_numbering(mode):
    CODE_STATE["numeracion"] = mode if mode in NUMERACION_TEXTO else NUMERACION_NUEVA


def set_elaboration_day(day):
    """Fecha de elaboración fija para los códigos; None = siempre la de hoy."""
    CODE_STATE["dia"] = None if day is None or day == date.today() else _as_date(day)


def effective_day():
    return CODE_STATE["dia"] or date.today()


# -- lector de códigos: las reglas del Informe, en el mismo orden ---------------------------------------------
_TRANSICION = {}


def tabla_transicion():
    """Código viejo exacto → (clase, dato, origen). Los 47 códigos de producto + los 9 PADRE anteriores."""
    if not _TRANSICION:
        for number, old in CODIGOS_ANTERIORES.items():
            _TRANSICION[old] = ("producto", number, "catálogo de e-Karú (antes de corregir el control)")
        for number, name in products_by_number().items():
            _TRANSICION.setdefault(PRODUCTS[name]["barcode"], ("producto", number, "e-Karú Etiquetas 2.0 (sin fecha)"))
        for old, new in PADRE_ANTERIORES.items():
            _TRANSICION[old] = ("padre", new, "grupo PADRE anterior (cambia solo el último dígito)")
    return _TRANSICION


def peso_texto(gramos):
    return f"{gramos / 1000:.3f}".replace(".", ",") + " kg"


def fecha_texto(day):
    return f"{day:%d/%m/%Y}"


def interpretar_codigo(texto, hoy=None, insumos=None):
    """Lee un código como lo hará e-Karú (reglas del Informe, en orden). No usa la interfaz.

    Devuelve un dict: estado («ok», «viejo», «rechazado» o «externo»), tipo, regla, campos [(nombre, valor)],
    mensaje y, si es un producto, su nombre en «producto». «insumos» = {Nº: nombre} (opcional).
    """
    hoy = hoy or date.today()
    code = "".join(str(texto).split()).replace("-", "")
    out = {"codigo": code, "estado": "rechazado", "tipo": "", "regla": "", "campos": [], "mensaje": "", "producto": None}

    def done(estado, tipo, regla, mensaje="", campos=None, producto=None):
        out.update(estado=estado, tipo=tipo, regla=regla, mensaje=mensaje, campos=list(campos or []),
                   producto=producto)
        return out

    if not code:
        return done("rechazado", "Sin código", "", "Escaneá una etiqueta o escribí el número del código.")
    if not code.isdigit():
        return done("externo", "Código no numérico", "Regla 7 · otro",
                    "No es un código interno de Glacé (los internos son solo números). e-Karú lo busca en las "
                    "equivalencias de proveedor; si no existe, pide vincularlo.")
    # Regla 1 · 21 dígitos que empiezan con 2 → balde con transacción
    if len(code) == 21 and code[0] == "2":
        grams = int(code[7:12])
        campos = [("Sabor", code[1:7]), ("Peso", peso_texto(grams)), ("Transacción", code[12:20])]
        if mod10_check(code[:20]) != code[20]:
            return done("rechazado", "Balde con transacción", "Regla 1",
                        f"El dígito de control no cierra (debería ser {mod10_check(code[:20])}). Escaneá de nuevo.", campos)
        if not 500 <= grams <= 12000:
            return done("rechazado", "Balde con transacción", "Regla 1",
                        "El peso está fuera del rango del envase (0,5 a 12 kg).", campos)
        return done("ok", "Balde con transacción", "Regla 1", "Si esa transacción ya está registrada, e-Karú avisa "
                    "«balde ya escaneado». El sabor tiene que existir en el maestro de e-Karú.", campos)
    # Tabla de transición: se busca el código exacto antes de las reglas 2 y 3
    if len(code) == 13 and code in tabla_transicion():
        kind, data, origin = tabla_transicion()[code]
        if kind == "producto":
            name = products_by_number().get(data, "(producto)")
            return done("viejo", "Código VIEJO de producto", "Tabla de transición",
                        f"Etiqueta con la numeración anterior ({origin}). Hay que reetiquetar con el código nuevo "
                        f"(22 + Nº + fecha).", [("Producto", name), ("Nº de producto", f"{data:04d}")], name)
        return done("viejo", "Código VIEJO de grupo PADRE", "Tabla de transición",
                    f"{origin}. El código correcto es {data}.", [("Grupo", padre_display(data) or data),
                                                                 ("Código correcto", data)])
    # Regla 2 · 13 dígitos → control EAN-13
    if len(code) == 13 and not is_valid_ean13(code):
        return done("rechazado", "Dígito de control incorrecto", "Regla 2",
                    f"El último dígito debería ser {ean13_check_digit(code[:12])}. Escaneá de nuevo.")
    prefix = code[:2] if (len(code) == 13 and code[0] == "2") else ""
    # Regla 3 · 20 → balde
    if prefix == "20":
        grams = int(code[7:12])
        number = int(code[1:7])
        # numeración anterior con fecha: 2 + producto (6) + AADDD (año + día del año). AADDD >= 24001 queda siempre
        # por encima del peso máximo de un balde (12 kg), así que no se confunde con un balde.
        old_day = None
        year, ordinal = 2000 + int(code[7:9]), int(code[9:12])
        if 1 <= number <= max(CODIGOS_ANTERIORES) and 2024 <= year <= 2060 and 1 <= ordinal <= 366:
            old_day = date(year, 1, 1) + timedelta(days=ordinal - 1)
            if old_day.year != year:  # día 366 en un año común
                old_day = None
        if old_day:
            name = products_by_number().get(number, "(producto)")
            return done("viejo", "Código VIEJO de producto con fecha", "Tabla de transición",
                        "Numeración anterior (2 + producto + fecha AADDD). Hay que reetiquetar con el código nuevo.",
                        [("Producto", name), ("Elaborado", fecha_texto(old_day))], name)
        campos = [("Sabor", code[1:7]), ("Peso", peso_texto(grams))]
        if not 500 <= grams <= 12000:
            return done("rechazado", "Balde", "Regla 3", "El peso está fuera del rango del envase (0,5 a 12 kg).", campos)
        return done("ok", "Balde", "Regla 3", "El sabor tiene que existir en el maestro de e-Karú.", campos)
    # Regla 4 · 22 → producto
    if prefix == "22":
        number, day = int(code[2:6]), parse_aammdd(code[6:12])
        name = products_by_number().get(number)
        campos = [("Producto", name or "(no existe)"), ("Nº de producto", f"{number:04d}"),
                  ("Fecha de elaboración", fecha_texto(day) if day else code[6:12] + " (no es una fecha)")]
        if name is None:
            return done("rechazado", "Producto terminado", "Regla 4", "Ese Nº de producto no existe en el catálogo.", campos)
        if day is None:
            return done("rechazado", "Producto terminado", "Regla 4", "La fecha de elaboración no es válida.", campos, name)
        if day > hoy:
            return done("rechazado", "Producto terminado", "Regla 4", "La fecha de elaboración es futura.", campos, name)
        campos.append(("Antigüedad", f"{(hoy - day).days} día(s)"))
        return done("ok", "Producto terminado", "Regla 4", "La fecha es el lote del producto.", campos, name)
    # Regla 5 · 23, 24, 25 y 29
    if prefix == "23":
        return done("ok", "Promo Pack", "Regla 5", "e-Karú lo asocia a los 12 palitos escaneados.",
                    [("Nº de paquete", str(int(code[2:12])))])
    if prefix == "24":
        return done("ok", "ID de balde", "Regla 5", "Etiqueta del plan B: se pega junto a la de la balanza.",
                    [("Nº de balde", str(int(code[2:12])))])
    if prefix == "25":
        number, day = int(code[2:6]), parse_aammdd(code[6:12])
        name = (insumos or {}).get(number)
        campos = [("Insumo", name or "(sin nombre en esta PC)"), ("Nº de insumo", f"{number:04d}"),
                  ("Fecha de recepción", fecha_texto(day) if day else code[6:12] + " (no es una fecha)")]
        if day is None:
            return done("rechazado", "Insumo interno", "Regla 5", "La fecha de recepción no es válida.", campos)
        if day > hoy:
            return done("rechazado", "Insumo interno", "Regla 5", "La fecha de recepción es futura.", campos)
        return done("ok", "Insumo interno", "Regla 5", "", campos)
    if prefix == "29":
        label = padre_display(code)
        if label is None:
            return done("rechazado", "Grupo PADRE", "Regla 5", "Ese grupo PADRE no existe en el catálogo.")
        return done("ok", "Grupo PADRE", "Regla 5", "Descuenta el grupo y los gramos del balde elegido.",
                    [("Grupo", label)])
    # Regla 6 · reservados
    if prefix in ("21", "26", "27", "28"):
        return done("rechazado", "Prefijo reservado", "Regla 6", f"El prefijo {prefix} queda libre para usos futuros.")
    # Regla 7 · cualquier otro → EAN de proveedor
    return done("externo", "EAN de proveedor", "Regla 7 · otro",
                "e-Karú lo busca en las equivalencias; si no existe, pide vincularlo una sola vez al insumo.")


def transition_rows(day=None):
    """Filas de la tabla de transición (47 productos + 9 grupos PADRE) para exportar."""
    day = day or date.today()
    rows = [["Nº", "Producto", "Categoría", "Código anterior (catálogo e-Karú)", "Código 2.0 (control corregido)",
             "Código nuevo v2 (ejemplo)", "Fecha del ejemplo", "Formato nuevo"]]
    for number, name in sorted(products_by_number().items()):
        rows.append([f"{number:04d}", name, product_category(name), CODIGOS_ANTERIORES.get(number, ""),
                     PRODUCTS[name]["barcode"], codigo_producto(number, day), fecha_texto(day),
                     "22 + Nº (4) + AAMMDD + control"])
    seen = set()
    for old, new in PADRE_ANTERIORES.items():
        if new in seen:
            continue
        seen.add(new)
        rows.append(["", padre_display(new) or new, "Grupo PADRE", old, new, new, "",
                     "29 se mantiene: solo cambia el último dígito"])
    return rows


def write_transition_csv(path=None, day=None):
    """Guarda la tabla de transición (CSV con «;», UTF-8 con BOM, abre bien en Excel). Devuelve la ruta."""
    path = Path(path) if path else TRANSITION_FILE
    tmp = path.with_name(path.name + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle, delimiter=";", lineterminator="\r\n").writerows(transition_rows(day))
    os.replace(tmp, path)
    return path


def duplicate_barcodes():
    seen = {}
    for product, data in PRODUCTS.items():
        if not data["barcode"]:
            continue
        seen.setdefault(data["barcode"], []).append(product)
    return {code: products for code, products in seen.items() if len(products) > 1}


def product_category(name):
    upper = name.upper()
    for prefix, category in PRODUCT_CATEGORIES.items():
        if upper.startswith(prefix):
            return category
    return "Otros"


PRODUCT_CATEGORIES = {
    "PALITO": "Palitos",
    "ALFAJOR": "Alfajores",
    "BROWNIE": "Brownies",
    "PASTAFROLA": "Postres",
    "GALLETITA": "Postres",
    "TARTA": "Tartas",
    "CONO": "Conos",
    "VASO": "Vasos",
    "TORTA": "Tortas",
    "MILKSHAKE": "Milkshakes",
    "SMOOTHIE": "Smoothies",
    "GRUPO PADRE": "Grupos PADRE",
    "MASCARPONE": "Baldes 10 x 8 cm",
    "MENTA": "Baldes 10 x 8 cm",
    "PISTACHO": "Baldes 10 x 8 cm",
}


# Etiquetas de baldes 10 x 8 cm: solo información (texto legal de los Word + tabla nutricional). Sin código de barras.
LARGE_LABEL_LINES = {
    'Mascarpone Cheesecake frutos del bosque - Balde 6 kg': {'title': 'HELADO DE CREMA SABOR CHEESECAKE CON FRUTOS DEL BOSQUE', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 190 kcal = 796 kJ    10% VD', 'Carbohidratos: 42 g (14% VD)    Azúcares: 42 g', 'Proteínas: 0,4 g (1% VD)    Grasas Totales: 2,3 g (4% VD)', 'Grasas Saturadas: 2 g (9% VD)    Grasas Trans: 0 g', 'Fibra Alimentaria: 0 g (0% VD)    Sodio: 63 mg (3% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, queso crema (glucosa de maíz, agua, sal, almidón, acidulante: ácido láctico (INS 270) , aromatizante sintético idéntico al natural , conservante: sorbato de potasio (INS 202)), leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), dextrosa, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), pasta frutos de bosque (azúcar, glucosa de maíz, frambuesa, mora, frutilla, agua, almidón, acidulante: ácido cítrico (INS 330), conservante: sorbato de potasio (INS 202), aromatizante artificial, espesantes: pectinas (INS 440), colorante idéntico al natural: caramelo IV (INS 150d) colorante artificial: amaranto (INS 123)), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar)', 'allerg': 'ALERGENOS: CONTIENE LECHE Y DERIVADOS. CONTIENE SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122459', 'neto': 'Contenido neto: 6 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'},
    'Mascarpone Cheesecake frutos del bosque - Balde 3 kg': {'title': 'HELADO DE CREMA SABOR CHEESECAKE CON FRUTOS DEL BOSQUE', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 190 kcal = 796 kJ    10% VD', 'Carbohidratos: 42 g (14% VD)    Azúcares: 42 g', 'Proteínas: 0,4 g (1% VD)    Grasas Totales: 2,3 g (4% VD)', 'Grasas Saturadas: 2 g (9% VD)    Grasas Trans: 0 g', 'Fibra Alimentaria: 0 g (0% VD)    Sodio: 63 mg (3% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, queso crema (glucosa de maíz, agua, sal, almidón, acidulante: ácido láctico (INS 270) , aromatizante sintético idéntico al natural , conservante: sorbato de potasio (INS 202)), leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), dextrosa, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), pasta frutos de bosque (azúcar, glucosa de maíz, frambuesa, mora, frutilla, agua, almidón, acidulante: ácido cítrico (INS 330), conservante: sorbato de potasio (INS 202), aromatizante artificial, espesantes: pectinas (INS 440), colorante idéntico al natural: caramelo IV (INS 150d) colorante artificial: amaranto (INS 123)), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar)', 'allerg': 'ALERGENOS: CONTIENE LECHE Y DERIVADOS. CONTIENE SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122459', 'neto': 'Contenido neto: 3 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'},
    'Menta Granizado - Balde 6 kg': {'title': 'HELADO DE CREMA SABOR MENTA GRANIZADO', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 118 kcal = 495 kJ    6% VD', 'Carbohidratos: 16 g (5% VD)    Azúcares: 14 g', 'Proteínas: 1,2 g (2% VD)    Grasas Totales: 5,5 g (10% VD)', 'Grasas Saturadas: 3,6 g (16% VD)    Grasas Trans: 0,1 g', 'Fibra Alimentaria: 0,1 g (0% VD)    Sodio: 28 mg (1% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), polvo de menta, dextrosa, menta, licor de menta, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar), chocolate para granizar (azúcar; aceite vegetal; cacao en polvo; emulsionantes: INS322, INS476; aromatizantes permitidos, aromatizado con etilvainillina.), esencia de menta.', 'allerg': 'CONTIENE LECHE Y DERIVADOS DE LECHE Y SOJA. PUEDE CONTENER MANÍ, AVELLANAS Y SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122417', 'neto': 'Contenido neto: 6 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'},
    'Menta Granizado - Balde 3 kg': {'title': 'HELADO DE CREMA SABOR MENTA GRANIZADO', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 118 kcal = 495 kJ    6% VD', 'Carbohidratos: 16 g (5% VD)    Azúcares: 14 g', 'Proteínas: 1,2 g (2% VD)    Grasas Totales: 5,5 g (10% VD)', 'Grasas Saturadas: 3,6 g (16% VD)    Grasas Trans: 0,1 g', 'Fibra Alimentaria: 0,1 g (0% VD)    Sodio: 28 mg (1% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), polvo de menta, dextrosa, menta, licor de menta, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar), chocolate para granizar (azúcar; aceite vegetal; cacao en polvo; emulsionantes: INS322, INS476; aromatizantes permitidos, aromatizado con etilvainillina.), esencia de menta', 'allerg': 'CONTIENE LECHE Y DERIVADOS DE LECHE Y SOJA. PUEDE CONTENER MANÍ, AVELLANAS Y SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122417', 'neto': 'Contenido neto: 3 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'},
    'Pistacho - Balde 6 kg': {'title': 'HELADO DE CREMA SABOR PISTACHO', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 193 kcal = 807 kJ    10% VD', 'Carbohidratos: 41 g (14% VD)    Azúcares: 40 g', 'Proteínas: 0,7 g (1% VD)    Grasas Totales: 2,9 g (5% VD)', 'Grasas Saturadas: 2,1 g (10% VD)    Grasas Trans: 0 g', 'Fibra Alimentaria: 0,1 g (0% VD)    Sodio: 71 mg (3% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), manteca de maní (maní tostado), dextrosa, pistacho tostado molido, esencia de pistacho, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar), sal.', 'allerg': 'ALERGENOS: CONTIENE LECHE, MANÍ, PISTACHO Y SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122416', 'neto': 'Contenido neto: 6 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'},
    'Pistacho - Balde 3 kg': {'title': 'HELADO DE CREMA SABOR PISTACHO', 'nutri': ['Porción 60 gramos (1 bola)', 'Valor Energético: 193 kcal = 807 kJ    10% VD', 'Carbohidratos: 41 g (14% VD)    Azúcares: 40 g', 'Proteínas: 0,7 g (1% VD)    Grasas Totales: 2,9 g (5% VD)', 'Grasas Saturadas: 2,1 g (10% VD)    Grasas Trans: 0 g', 'Fibra Alimentaria: 0,1 g (0% VD)    Sodio: 71 mg (3% VD)'], 'ingred': 'INGREDIENTES: Agua, azúcar, crema de leche, leche entera, leche descremada, agente de batido (maltodextrina, grasa vegetal hidrogenada, proteína láctea, emulsificante: mono y diglicéridos de ácidos grasos INS 471, emulsificante ésteres de propilenglicol de ácidos grasos INS 477, estabilizante pirofosfato ácido de sodio INS 450iii), manteca de maní (maní tostado), dextrosa, pistacho tostado molido, esencia de pistacho, fibra cítrica (maltodextrina, fibra cítrica y monoglicéridos de ácidos graso), emulsionantes ( Mono y diglicéridos de ácidos grasos INS 471, Carboximetilcelulosa (CMC) INS 466, Goma guar INS 412, Carragenina INS 407, Goma garrofín (goma de algarrobo) INS 410, Polisorbato 80 INS 433, azúcar), sal.', 'allerg': 'ALERGENOS: CONTIENE LECHE, MANÍ, PISTACHO Y SULFITOS.', 'maker': 'Elaborado por GLACE S.A. Dirección: Capitán Domingo Antonio Ortiz casi San Pedro 2571 - Asunción, Paraguay. Industria Paraguaya RE: 801002842/1 RSPA: 122416', 'neto': 'Contenido neto: 3 kg.', 'elab': 'Elaboración, Vencimiento y Lote: (Ver envase) Duración a -18° C (freezer): hasta vencimiento'}}


# --------------------------------------------------------------------------- #
#  Utilidades: registro, configuración, escritura atómica
# --------------------------------------------------------------------------- #
def log(message):
    """Escribe en ekaru_impresion.log; nunca lanza excepciones."""
    try:
        if LOG_FILE.exists() and LOG_FILE.stat().st_size > 512 * 1024:
            LOG_FILE.write_text("", encoding="utf-8")
        with LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {message}\n")
    except OSError:
        pass


def atomic_write_json(path, data):
    """Escritura atómica (segura con OneDrive): archivo temporal + replace."""
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


DEFAULT_CONFIG = {
    "printer": "",
    "printer_large": "",  # impresora para etiquetas 10 x 8 cm (FTX TTP010UE); vacío = detectar sola
    "gap_mm": 2.0,       # separación entre etiquetas; 0 = papel continuo/sin gap
    "density": 8,        # 0-15
    "speed": 4,          # 1-6
    "direction": 1,      # 0 o 1 (ver prueba si sale al revés)
    "offset_x": 0,       # puntos
    "offset_y": 0,
    "strip_accents": False,
    "mode": "gdi",       # "gdi" = como Ctrl+P (driver de Windows) · "raw" = TSPL directo
    # Rollo con varias etiquetas de 2 x 3 cm en el ancho: se llenan columnas de izquierda a derecha
    "grid": True,
    "cols": 4,
    "col_pitch_mm": 23.0,    # distancia entre el borde izquierdo de una etiqueta y el de la siguiente
    "margin_left_mm": 3.0,   # borde izquierdo de la 1.ª columna medido desde el borde de la página
    "margin_top_mm": 1.0,
    "row_pitch_mm": 33.0,    # alto de una fila de etiquetas (etiqueta + separación)
    "codigo_con_fecha": True,  # numeración anterior: códigos hijo con fecha de elaboración (ver print_code)
    "numeracion": "nueva",     # "nueva" = v2 del Informe (22 + Nº + AAMMDD…) · "actual" = la de la versión 2.0
    "aviso_numeracion": False,  # True cuando ya se mostró el aviso de la numeración nueva
}


def load_config():
    config = dict(DEFAULT_CONFIG)
    try:
        saved = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        if isinstance(saved, dict):
            for key, default in DEFAULT_CONFIG.items():
                value = saved.get(key)
                if isinstance(default, bool):
                    if isinstance(value, bool):
                        config[key] = value
                elif isinstance(default, (int, float)):
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        config[key] = type(default)(value)
                elif isinstance(value, str):
                    config[key] = value
    except (OSError, ValueError, TypeError):
        pass
    if config.get("numeracion") not in NUMERACION_TEXTO:
        config["numeracion"] = NUMERACION_NUEVA
    return config


_STARTUP_CONFIG = load_config()
CODE_WITH_DATE = bool(_STARTUP_CONFIG.get("codigo_con_fecha", True))
set_numbering(_STARTUP_CONFIG.get("numeracion"))


# -- datos propios de la numeración v2: contadores (23 y 24) e insumos (25) ---------------------------------
def load_counters():
    """Últimos números usados de Promo Pack («pack») e ID de balde («balde»). Nunca lanza excepciones."""
    counters = {"pack": 0, "balde": 0}
    try:
        data = json.loads(COUNTERS_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return counters
    if isinstance(data, dict):
        for key in counters:
            value = data.get(key)
            if isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= MAX_NUMERO_CORRELATIVO:
                counters[key] = value
    return counters


def save_counters(counters):
    atomic_write_json(COUNTERS_FILE, {"pack": int(counters.get("pack", 0)), "balde": int(counters.get("balde", 0))})


def load_supplies():
    """Lista de insumos internos [(Nº, nombre)], ordenada por Nº. Si no hay archivo, la lista inicial."""
    items = {}
    try:
        data = json.loads(SUPPLIES_FILE.read_text(encoding="utf-8"))
        for entry in data.get("insumos", []) if isinstance(data, dict) else []:
            number, name = entry.get("numero"), str(entry.get("nombre", "")).strip()
            if isinstance(number, int) and not isinstance(number, bool) and 1 <= number <= 9999 and name:
                items[number] = name
    except (OSError, ValueError, TypeError, AttributeError):
        items = {}
    if not items:
        items = dict(INSUMOS_INICIALES)
    return sorted(items.items())


def save_supplies(items):
    atomic_write_json(SUPPLIES_FILE, {"insumos": [{"numero": n, "nombre": name} for n, name in sorted(items)]})


def add_supply(name):
    """Agrega un insumo con el siguiente Nº (los números no se reutilizan). Devuelve (Nº, lista)."""
    name = " ".join(str(name).split())
    if not name:
        raise ValueError("El insumo necesita un nombre.")
    items = load_supplies()
    if any(existing.lower() == name.lower() for _, existing in items):
        raise ValueError("Ya existe un insumo con ese nombre.")
    number = max(n for n, _ in items) + 1
    if number > 9999:
        raise ValueError("No quedan números de insumo libres (máximo 9999).")
    items.append((number, name))
    save_supplies(items)
    return number, sorted(items)


def strip_accents(text):
    return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")


class PrintJob:
    """Un trabajo de impresión: elementos de la etiqueta + programa TSPL equivalente.

    Con «labels» (lista de listas de elementos) es una SERIE: cada etiqueta lleva su propio contenido
    (por ejemplo, números correlativos). Sin «labels» se imprimen «copies» copias idénticas.
    """

    def __init__(self, elements, w_mm, h_mm, copies, config, page="label", labels=None):
        self.page = page            # "label" | "grid" (varias columnas) | "strip" (fila de calibración)
        self.labels = [list(items) for items in labels][:999] if labels else None
        if self.labels:
            elements, copies = self.labels[0], len(self.labels)
        self.elements = list(elements)
        self.w_mm, self.h_mm = w_mm, h_mm
        self.copies = max(1, min(999, int(copies)))
        self.config = dict(config)
        self.on_success = None      # función opcional: se llama cuando el envío salió bien
        self.data = build_tspl(self.elements, w_mm, h_mm, self.copies, self.config, self.labels)


# --------------------------------------------------------------------------- #
#  Impresión: backends
# --------------------------------------------------------------------------- #
class PrintError(Exception):
    """Error de impresión con mensaje apto para el usuario."""


PRINTER_STATUS_FLAGS = [
    (0x00000001, "en pausa"), (0x00000002, "con error"), (0x00000004, "eliminando trabajo"),
    (0x00000008, "papel atascado"), (0x00000010, "sin papel/etiquetas"), (0x00000040, "problema de papel"),
    (0x00000080, "sin conexión"), (0x00001000, "no disponible"), (0x00100000, "requiere intervención"),
    (0x00400000, "tapa abierta"),
]
ATTR_WORK_OFFLINE = 0x00000400


class WinSpooler:
    """Envío RAW al spooler de Windows usando solo ctypes (sin pywin32)."""

    name = "Windows (spooler RAW)"
    is_real = True

    def __init__(self):
        from ctypes import wintypes
        self.wt = wintypes
        dll = ctypes.WinDLL("winspool.drv", use_last_error=True)
        h = ctypes.c_void_p

        class PRINTER_INFO_4(ctypes.Structure):
            _fields_ = [("pPrinterName", wintypes.LPWSTR), ("pServerName", wintypes.LPWSTR),
                        ("Attributes", wintypes.DWORD)]

        class PRINTER_INFO_6(ctypes.Structure):
            _fields_ = [("dwStatus", wintypes.DWORD)]

        class DOC_INFO_1(ctypes.Structure):
            _fields_ = [("pDocName", wintypes.LPWSTR), ("pOutputFile", wintypes.LPWSTR),
                        ("pDatatype", wintypes.LPWSTR)]

        self.PI4, self.PI6, self.DOC = PRINTER_INFO_4, PRINTER_INFO_6, DOC_INFO_1
        dll.EnumPrintersW.argtypes = [wintypes.DWORD, wintypes.LPWSTR, wintypes.DWORD, ctypes.c_void_p,
                                      wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(wintypes.DWORD)]
        dll.EnumPrintersW.restype = wintypes.BOOL
        dll.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(h), ctypes.c_void_p]
        dll.OpenPrinterW.restype = wintypes.BOOL
        dll.ClosePrinter.argtypes = [h]
        dll.ClosePrinter.restype = wintypes.BOOL
        dll.StartDocPrinterW.argtypes = [h, wintypes.DWORD, ctypes.c_void_p]
        dll.StartDocPrinterW.restype = wintypes.DWORD
        dll.EndDocPrinter.argtypes = [h]
        dll.EndDocPrinter.restype = wintypes.BOOL
        dll.StartPagePrinter.argtypes = [h]
        dll.StartPagePrinter.restype = wintypes.BOOL
        dll.EndPagePrinter.argtypes = [h]
        dll.EndPagePrinter.restype = wintypes.BOOL
        dll.WritePrinter.argtypes = [h, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        dll.WritePrinter.restype = wintypes.BOOL
        dll.GetPrinterW.argtypes = [h, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        dll.GetPrinterW.restype = wintypes.BOOL
        dll.GetDefaultPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        dll.GetDefaultPrinterW.restype = wintypes.BOOL
        self.dll = dll

    def _err(self):
        code = ctypes.get_last_error()
        return f"código de Windows {code}: {ctypes.FormatError(code).strip()}"

    def _enum(self):
        needed, returned = self.wt.DWORD(0), self.wt.DWORD(0)
        flags = 2 | 4  # LOCAL | CONNECTIONS
        self.dll.EnumPrintersW(flags, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(returned))
        if needed.value == 0:
            return []
        buf = ctypes.create_string_buffer(needed.value)
        if not self.dll.EnumPrintersW(flags, None, 4, buf, needed.value, ctypes.byref(needed), ctypes.byref(returned)):
            raise PrintError(f"No se pudo listar las impresoras ({self._err()}).")
        array = ctypes.cast(buf, ctypes.POINTER(self.PI4))
        return [(array[i].pPrinterName, array[i].Attributes) for i in range(returned.value)]

    def list_printers(self):
        return sorted(name for name, _ in self._enum() if name)

    def default_printer(self):
        size = self.wt.DWORD(0)
        self.dll.GetDefaultPrinterW(None, ctypes.byref(size))
        if size.value == 0:
            return ""
        buf = ctypes.create_unicode_buffer(size.value)
        return buf.value if self.dll.GetDefaultPrinterW(buf, ctypes.byref(size)) else ""

    def status(self, printer):
        """Devuelve (ok, texto). ok=False si Windows la marca sin conexión/pausada/error."""
        problems = []
        for name, attrs in self._enum():
            if name == printer and attrs & ATTR_WORK_OFFLINE:
                problems.append("modo «Usar impresora sin conexión» activado")
        handle = ctypes.c_void_p()
        if not self.dll.OpenPrinterW(printer, ctypes.byref(handle), None):
            return False, f"No se puede abrir «{printer}» ({self._err()})"
        try:
            needed = self.wt.DWORD(0)
            self.dll.GetPrinterW(handle, 6, None, 0, ctypes.byref(needed))
            if needed.value:
                buf = ctypes.create_string_buffer(needed.value)
                if self.dll.GetPrinterW(handle, 6, buf, needed.value, ctypes.byref(needed)):
                    state = ctypes.cast(buf, ctypes.POINTER(self.PI6))[0].dwStatus
                    problems += [text for bit, text in PRINTER_STATUS_FLAGS if state & bit]
        finally:
            self.dll.ClosePrinter(handle)
        if problems:
            return False, "Windows informa: " + ", ".join(problems)
        return True, "Lista"

    def send(self, printer, job, title="E-Karu etiqueta"):
        if not printer:
            raise PrintError("No hay impresora seleccionada.")
        if job.config.get("mode", "gdi") == "gdi":
            return self.send_gdi(printer, job, title)
        data = job.data
        handle = ctypes.c_void_p()
        if not self.dll.OpenPrinterW(printer, ctypes.byref(handle), None):
            raise PrintError(f"No se pudo abrir la impresora «{printer}» ({self._err()}).")
        job, finished = 0, False
        try:
            doc = self.DOC(title, None, "RAW")
            job = self.dll.StartDocPrinterW(handle, 1, ctypes.byref(doc))
            if not job:
                raise PrintError(f"Windows rechazó el trabajo RAW ({self._err()}). "
                                 "Revisá que el controlador de la impresora acepte datos RAW.")
            if not self.dll.StartPagePrinter(handle):
                raise PrintError(f"StartPage falló ({self._err()}).")
            written = self.wt.DWORD(0)
            buf = ctypes.create_string_buffer(data, len(data))
            if not self.dll.WritePrinter(handle, buf, len(data), ctypes.byref(written)):
                raise PrintError(f"No se pudieron enviar los datos ({self._err()}).")
            if written.value != len(data):
                raise PrintError(f"Envío incompleto: {written.value} de {len(data)} bytes.")
            self.dll.EndPagePrinter(handle)
            finished = True
        finally:
            if job:
                self.dll.EndDocPrinter(handle)
            self.dll.ClosePrinter(handle)
        return f"trabajo #{job} enviado a «{printer}» ({len(data)} bytes)"


    # -- impresión por el driver de Windows (igual que Ctrl+P) ------------------
    def _devmode(self, printer, w_mm, h_mm, keep_width=False):
        """DEVMODE con papel personalizado w_mm x h_mm (el driver suele venir con 4x6 pulgadas)."""
        try:
            wt = self.wt
            dp = self.dll.DocumentPropertiesW
            dp.argtypes = [wt.HWND, ctypes.c_void_p, wt.LPWSTR, ctypes.c_void_p, ctypes.c_void_p, wt.DWORD]
            dp.restype = ctypes.c_long
            handle = ctypes.c_void_p()
            if not self.dll.OpenPrinterW(printer, ctypes.byref(handle), None):
                return None
            try:
                size = dp(None, handle, printer, None, None, 0)
                if size <= 0:
                    return None
                buf = ctypes.create_string_buffer(size)
                if dp(None, handle, printer, buf, None, 2) < 0:  # DM_OUT_BUFFER
                    return None
                import struct
                fields = struct.unpack_from("<I", buf, 72)[0] | 0x2 | 0x4 | 0x8   # PAPERSIZE|LENGTH|WIDTH
                struct.pack_into("<I", buf, 72, fields)
                width10 = struct.unpack_from("<h", buf, 82)[0] if keep_width else int(w_mm * 10)
                if width10 <= 0:
                    width10 = 1016
                struct.pack_into("<hhh", buf, 78, 256, int(h_mm * 10), width10)  # DMPAPER_USER
                if dp(None, handle, printer, buf, buf, 10) < 0:  # DM_IN_BUFFER | DM_OUT_BUFFER
                    return None
                return buf
            finally:
                self.dll.ClosePrinter(handle)
        except Exception as error:  # noqa: BLE001 - si falla se usa el papel del driver
            log(f"DEVMODE no disponible: {error!r}")
            return None

    def send_gdi(self, printer, job, title):
        wt = self.wt
        gdi = getattr(self, "gdi", None)
        if gdi is None:
            gdi = ctypes.WinDLL("gdi32.dll", use_last_error=True)
            h = ctypes.c_void_p
            gdi.CreateDCW.argtypes = [wt.LPCWSTR, wt.LPCWSTR, wt.LPCWSTR, ctypes.c_void_p]
            gdi.CreateDCW.restype = h
            gdi.DeleteDC.argtypes = [h]
            gdi.GetDeviceCaps.argtypes = [h, ctypes.c_int]
            gdi.GetDeviceCaps.restype = ctypes.c_int
            gdi.StartDocW.argtypes = [h, ctypes.c_void_p]
            gdi.StartDocW.restype = ctypes.c_int
            gdi.StartPage.argtypes = [h]
            gdi.StartPage.restype = ctypes.c_int
            gdi.EndPage.argtypes = [h]
            gdi.EndPage.restype = ctypes.c_int
            gdi.EndDoc.argtypes = [h]
            gdi.EndDoc.restype = ctypes.c_int
            gdi.AbortDoc.argtypes = [h]
            gdi.PatBlt.argtypes = [h, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wt.DWORD]
            gdi.PatBlt.restype = wt.BOOL
            gdi.CreateFontW.argtypes = ([ctypes.c_int] * 5 + [wt.DWORD] * 8 + [wt.LPCWSTR])
            gdi.CreateFontW.restype = h
            gdi.SelectObject.argtypes = [h, h]
            gdi.SelectObject.restype = h
            gdi.DeleteObject.argtypes = [h]
            gdi.SetBkMode.argtypes = [h, ctypes.c_int]
            gdi.SetTextColor.argtypes = [h, wt.DWORD]
            gdi.TextOutW.argtypes = [h, ctypes.c_int, ctypes.c_int, wt.LPCWSTR, ctypes.c_int]
            gdi.TextOutW.restype = wt.BOOL
            self.gdi = gdi

        class DOCINFOW(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_int), ("lpszDocName", wt.LPCWSTR), ("lpszOutput", wt.LPCWSTR),
                        ("lpszDatatype", wt.LPCWSTR), ("fwType", wt.DWORD)]

        cfg = job.config
        if job.page == "label":
            devmode = self._devmode(printer, job.w_mm, job.h_mm)
        else:
            devmode = self._devmode(printer, 0, float(cfg.get("row_pitch_mm", 33.0)), keep_width=True)
        hdc = gdi.CreateDCW("WINSPOOL", printer, None, ctypes.cast(devmode, ctypes.c_void_p) if devmode else None)
        if not hdc:
            raise PrintError(f"No se pudo abrir el driver de «{printer}» ({self._err()}).")
        started = False
        try:
            dpi_x, dpi_y = gdi.GetDeviceCaps(hdc, 88), gdi.GetDeviceCaps(hdc, 90)
            page_w, page_h = gdi.GetDeviceCaps(hdc, 8), gdi.GetDeviceCaps(hdc, 10)
            if min(dpi_x, dpi_y, page_w, page_h) <= 0:
                raise PrintError("El driver no informó el tamaño de página. Revisá el tamaño de papel de la impresora.")
            label_w, label_h = job.w_mm * DPM, job.h_mm * DPM
            kx, ky = dpi_x / 203.0, dpi_y / 203.0
            fit = min(1.0, page_w / (label_w * kx), page_h / (label_h * ky)) if job.page == "label" else 1.0
            kx, ky = kx * fit, ky * fit
            ox = int(job.config.get("offset_x", 0) * kx)
            oy = int(job.config.get("offset_y", 0) * ky)
            doc = DOCINFOW(ctypes.sizeof(DOCINFOW), title, None, None, 0)
            if gdi.StartDocW(hdc, ctypes.byref(doc)) <= 0:
                raise PrintError(f"El driver rechazó el trabajo ({self._err()}).")
            started = True
            fonts = {}

            def font_for(fw, fh, bold):
                key = (fw, fh, bold)
                if key not in fonts:
                    fonts[key] = gdi.CreateFontW(-max(6, round(fh * ky * 1.18)), max(3, round(fw * kx)), 0, 0,
                                                 900 if bold else 700, 0, 0, 0, 1, 0, 0, 3, 49, "Courier New")
                return fonts[key]

            def box(x, y, w, h):
                x0, x1 = round(x * kx), round((x + w) * kx)
                y0, y1 = round(y * ky), round((y + h) * ky)
                gdi.PatBlt(hdc, ox + x0, oy + y0, max(1, x1 - x0), max(1, y1 - y0), 0x00000042)  # BLACKNESS

            def draw_all(dx, dy, items=None):
                for e in (job.elements if items is None else items):
                    if e[0] == "bar":
                        box(e[1] + dx, e[2] + dy, e[3], e[4])
                    elif e[0] == "box":
                        _, x, y, w, h, t = e
                        x, y = x + dx, y + dy
                        box(x, y, w, t)
                        box(x, y + h - t, w, t)
                        box(x, y, t, h)
                        box(x + w - t, y, t, h)
                    elif e[0] == "text":
                        _, x, y, font, text, bold = e
                        fw, fh = FONT_DOTS[font]
                        old = gdi.SelectObject(hdc, font_for(fw, fh, bold))
                        tx = ox + round((x + dx) * kx)
                        ty = oy + round((y + dy) * ky)
                        for index, ch in enumerate(text):
                            if ch == " ":
                                continue
                            rects = glyph_rects(font, ch)
                            if rects is None:
                                gdi.TextOutW(hdc, tx + round(index * fw * kx), ty, ch, 1)
                                continue
                            for rx, ry, rw, rh in rects:
                                # el título (negrita) se engrosa 1 punto: más calor, trazo más firme
                                box(x + dx + index * fw + rx, y + dy + ry, rw + (1 if bold else 0), rh)
                        gdi.SelectObject(hdc, old)

            cols = max(1, int(cfg.get("cols", 4)))
            if job.page == "grid":
                pages = -(-job.copies // cols)
            elif job.page == "strip":
                pages = 1
            else:
                pages = job.copies
            left = float(cfg.get("margin_left_mm", 0.0)) * DPM
            top = float(cfg.get("margin_top_mm", 0.0)) * DPM
            pitch = float(cfg.get("col_pitch_mm", 23.0)) * DPM
            remaining = job.copies
            series = getattr(job, "labels", None)  # serie: cada etiqueta con su propio contenido
            printed = 0
            for _ in range(pages):
                if gdi.StartPage(hdc) <= 0:
                    raise PrintError(f"El driver no pudo iniciar la página ({self._err()}).")
                gdi.SetBkMode(hdc, 1)
                gdi.SetTextColor(hdc, 0)
                if job.page == "grid":
                    here = min(cols, remaining)
                    remaining -= here
                    for column in range(here):
                        draw_all(left + column * pitch, top, series[printed + column] if series else None)
                    printed += here
                else:
                    draw_all(0, 0, series[printed] if series and printed < len(series) else None)
                    printed += 1
                if gdi.EndPage(hdc) <= 0:
                    raise PrintError(f"El driver no pudo terminar la página ({self._err()}).")
            gdi.EndDoc(hdc)
            started = False
            for handle in fonts.values():
                gdi.DeleteObject(handle)
        finally:
            if started:
                gdi.AbortDoc(hdc)
            gdi.DeleteDC(hdc)
        return (f"{job.copies} etiqueta(s) en {pages} página(s) por el driver de «{printer}» "
                f"({page_w}x{page_h} px, {dpi_x} dpi)")


class FileBackend:
    """Sin spooler (Linux/macOS o pruebas): guarda el trabajo como archivo .prn."""

    name = "Archivo .prn (sin impresora)"
    is_real = False
    PRINTER = "Archivo .prn (carpeta salida_impresion)"

    def __init__(self, directory=None):
        self.directory = Path(directory) if directory else OUTPUT_DIR

    def list_printers(self):
        return [self.PRINTER]

    def default_printer(self):
        return self.PRINTER

    def status(self, printer):
        return True, "Lista"

    def send(self, printer, job, title="E-Karu etiqueta"):
        data = job.data
        try:
            self.directory.mkdir(parents=True, exist_ok=True)
            path = self.directory / f"etiqueta_{datetime.now():%Y%m%d_%H%M%S_%f}.prn"
            path.write_bytes(data)
        except OSError as error:
            raise PrintError(f"No se pudo guardar el archivo: {error}") from error
        return f"guardado {path.name} ({len(data)} bytes)"


class SimBackend(FileBackend):
    """Backend de pruebas: guarda los trabajos en memoria."""

    name = "Simulado"
    PRINTER = "HPRT HT400 (simulada)"

    def __init__(self):
        self.jobs = []
        self.fail_with = None
        self.state = (True, "Lista")

    def list_printers(self):
        return [self.PRINTER, "Microsoft Print to PDF (simulada)"]

    def status(self, printer):
        return self.state

    def send(self, printer, job, title="E-Karu etiqueta"):
        if self.fail_with:
            raise PrintError(self.fail_with)
        self.jobs.append((printer, job.data))
        return f"trabajo simulado #{len(self.jobs)} ({len(job.data)} bytes)"


def make_backend():
    if os.environ.get("EKARU_BACKEND") == "sim":
        return SimBackend()
    if os.name == "nt":
        try:
            return WinSpooler()
        except Exception as error:  # noqa: BLE001 - se informa en el diagnóstico
            log(f"No se pudo inicializar winspool.drv: {error!r}")
    return FileBackend()


# --------------------------------------------------------------------------- #
#  Diseño de etiquetas (un único modelo para vista previa e impresión)
#  Elementos en PUNTOS de impresora (203 dpi ≈ 8 puntos/mm):
#    ("text", x, y, fuente "1".."5", texto, negrita)
#    ("bar", x, y, ancho, alto)           barra negra rellena
#    ("box", x, y, ancho, alto, grosor)   rectángulo
# --------------------------------------------------------------------------- #
DPM = 8
FONT_DOTS = {"1": (8, 12), "2": (12, 20), "3": (16, 24), "4": (24, 32), "5": (32, 48)}

# Fuente de mapa de bits (negrita, Liberation Mono) para el modo driver: se dibuja con
# rectángulos negros, igual que las barras, porque el texto TrueType sale fragmentado
# en papel térmico. Cada glifo = filas en hexadecimal (ancho = FONT_DOTS[font][0]).
BITMAP_FONT = {
    "1": {
        '!': "000018181818181800181800",
        '"': "003c24242400000000000000",
        '#': "000012367f24247e7e484800",
        '$': "00187e6e68783e0f4b7e3c08",
        '%': "0000f2b6b4f81e1f2d6d4700",
        '&': "00003c242c3872dececf7f00",
        "'": "001818181800000000000000",
        '(': "000c18183830303030303818",
        ')': "003018181c0c0c0c0c0c1c18",
        '*': "00183c3c3c24000000000000",
        '+': "00000018187e7e1818180000",
        ',': "000000000000000000383030",
        '-': "0000000000003c3c00000000",
        '.': "000000000000000000181800",
        '/': "0006060c0c18183030606000",
        '0': "00003c66667e7e66667e3c00",
        '1': "0000187858181818187e7f00",
        '2': "00003e66460e1c38707e7e00",
        '3': "00007e66061e1c06667e3c00",
        '4': "00000e1e3e3e6efeff0e0e00",
        '5': "00007e60607e7e06667e3c00",
        '6': "00003e76607c7e66667e3c00",
        '7': "00007e060c0c181838303000",
        '8': "00007e66667e3e66667e3c00",
        '9': "00007c6666667e3e067e3c00",
        ':': "000000001818000000181800",
        ';': "000000081818000000181830",
        '<': "000000063e7060781e060000",
        '=': "000000007e00007e7e000000",
        '>': "000000607c0e061e78600000",
        '?': "00007e66460e1c1800181800",
        '@': "003c624b5fa5a5e5effe4066",
        'A': "00001c3c3c3c667e7ec3c300",
        'B': "00007e66667e7e67637f7e00",
        'C': "00003e7662e0e0e0677e3c00",
        'D': "00007c6e66676767667e7c00",
        'E': "00007e60607e7e60607e7f00",
        'F': "00007e6060607e7e60606000",
        'G': "00003c7664e0eeee667e3c00",
        'H': "00006666667e7e6666666600",
        'I': "00007e1818181818187e7e00",
        'J': "00001e0e06060606667c3c00",
        'K': "0000666e6c787c7c6e666700",
        'L': "0000606060606060607e7e00",
        'M': "0000e7e7ffffdbdbdbc3c300",
        'N': "00006676767e7e7e6e6e6600",
        'O': "00003c6666e7e7e7667e3c00",
        'P': "00007e6e67667e7c60606000",
        'Q': "00003c7666e7e7e7667e3c18",
        'R': "00007e6f67667e7c6e666700",
        'S': "00007e6660783e07c37e7e00",
        'T': "0000ff181818181818181800",
        'U': "0000666666666666667e3c00",
        'V': "0000c3e766667e3c3c3c1800",
        'W': "0000c3c3dbdb7e7e7e666600",
        'X': "000066663c38183c7e66c300",
        'Y': "0000e7667e3c181818181800",
        'Z': "00007e0e0c1c1830707eff00",
        '[': "003c30303030303030303030",
        '\\': "006060303018180c0c060600",
        ']': "003c0c0c0c0c0c0c0c0c0c0c",
        '^': "00183c3c6466c20000000000",
        '_': "000000000000000000000000",
        '`': "003808000000000000000000",
        'a': "000000187c663e7e666e7f00",
        'b': "0060606c7e66666766767c00",
        'c': "000000187e666060667e3c00",
        'd': "000606367e6666e6667e3e00",
        'e': "000000187e667e7e60763c00",
        'f': "0e1e387e7e30303030303000",
        'g': "000000127e666666667e3e06",
        'h': "0060606c7e66666666666600",
        'i': "0018003878181818181c7f00",
        'j': "000c00387c0c0c0c0c0c0c0c",
        'k': "00606062666c787c6c666600",
        'l': "0078181818181818181e0f00",
        'm': "00000012fedbdbdbdbdbdb00",
        'n': "0000004c7e66666666666600",
        'o': "000000187e6667e7667e3c00",
        'p': "0000004c7e66666766767c60",
        'q': "000000327e6666e6666e3e06",
        'r': "000000267e78707070707000",
        's': "000000187e66783e06663c00",
        't': "0000307c7e30303030381e00",
        'u': "0000004266666666667e7e00",
        'v': "00000042e766667c3c3c1800",
        'w': "00000000c3dbdb5e7e666600",
        'x': "00000042663c3c183c66e600",
        'y': "00000042e766663c3c3c1818",
        'z': "0000003c7e0c1c38307e7e00",
        '{': "001e18181818307038181818",
        '|': "001818181818181818181818",
        '}': "0078181818180c0e1c181818",
        '~': "000000000070ff0400000000",
        'Á': "18101c3c3c3c667e7ec3c300",
        'É': "18107e60607e7e60607e7f00",
        'Í': "18107e1818181818187e7e00",
        'Ó': "18183c6666e7e7e7667e3c00",
        'Ú': "1800666666666666667e3c00",
        'Ü': "2c00666666666666667e3c00",
        'Ñ': "7c006676767e7e7e6e6e6600",
        'á': "000c18187c663e7e666e7f00",
        'é': "000c18187e667e7e60763c00",
        'í': "000c183878181818181c7f00",
        'ó': "000c18187e6667e7667e3c00",
        'ú': "000c184266666666667e7e00",
        'ü': "0024244266666666667e7e00",
        'ñ': "003e2c4c7e66666666666600",
        '°': "00183c243c18000000000000",
    },
    "2": {
        '!': "000000000060060060060060060060060060060000000060060060000000",
        '"': "00019839c39c198198198198198000000000000000000000000000000000",
        '#': "0000000000c40cc08c08c7fe7fe1881181187fe7fe310330230230000000",
        '$': "0000200601f83fc76e76e7607e03f81fc07e06e26676676e3fe1f8060060",
        '%': "000000000786d86ccccd8cd8db07b006005e0db1b31b333321b61e000000",
        '&': "0000000000f01f83983981b81f01e03c67eee6ce7ce3ce3c7ff3e7000000",
        "'": "000060060060060060060060060000000000000000000000000000000000",
        '(': "0000380300700600e00e00e00c01c01c01c01c01c01c01c00e00e00e00e0",
        ')': "0001c00c00e0060070070070030038038038038038038038070070070070",
        '*': "0000600600603fc3fc0f00f0198090000000000000000000000000000000",
        '+': "0000000000000000600600600607fe7fe7fe060060060060060000000000",
        ',': "0000000000000000000000000000000000000000000000e00e01c01c0180",
        '-': "0000000000000000000000000000000001f81f81f8000000000000000000",
        '.': "000000000000000000000000000000000000000000000060060060000000",
        '/': "00000600e00c01c0180380300700600e00c00c01c0180380300700000000",
        '0': "0000000001f83fc39c30c70e70e76e76e76e70e70e30c39c3fc1f8000000",
        '1': "0000000000700f03f07f06700700700700700700700700707fe7fe000000",
        '2': "0000000001f83fc39c70e00e01c03c0780f01e01c03807807fe7fe000000",
        '3': "0000000001f83fc79e70e00e01c0780f807c00e00e70e79e3fc1f8000000",
        '4': "00000000003c07c07c0fc0dc1dc39c31c71c71c7fe7fe01c01c01c000000",
        '5': "0000000003fc3fc3803003003f87fc7fe30e00e00e70e79e3fc1f8000000",
        '6': "0000000000f81fc39c38c7007307fc7fc78e70e70e30e39e1fc0f8000000",
        '7': "0000000007fe7fe3fe01c01c0380300700700e00e00e01c01c01c0000000",
        '8': "0000000001f83fc79c70e70e39c3f81f839c70e70e70e70e3fc1f8000000",
        '9': "0000000001f83fc79c71c70e70e70e3fe3fe1ee00e31c39c3f81f0000000",
        ':': "000000000000000000060060060060000000000000000060060060000000",
        ';': "0000000000000000000700700700600000000000000000700e00e00c00c0",
        '<': "00000000000000200e03e0fc7f07806007807e01f807e00e002000000000",
        '=': "0000000000000000007fe7fe7fe0000000007fe7fe7fe000000000000000",
        '>': "0000000000004007007c03f00fe01e00601e07e1f87e0700400000000000",
        '?': "0000000001f83fc79e70e70e00e01c0780700e00e00000e00e00e0000000",
        '@': "0000700f818c30620267e4db59b59bd1bd1bd1ad12d325be4cc60020430c",
        'A': "0000000000f00f00f01f81f819839c39c39c3fc7fe7fe70ee07e07000000",
        'B': "0000000007f87fc79e70e70e71c7fc7fc79e70e70770f70e7fe7fc000000",
        'C': "0000000001f83fc3de78e70c70070070070070070670e39e3fc1f8000000",
        'D': "0000000007f07f87fc71e70e70e70e70e70e70e70e71e73c7fc7f0000000",
        'E': "0000000007fe7fe7fc7007007007fc7fc7fc7007007007007fe7fe000000",
        'F': "0000000003fe3fe3fc3803803803803fe3fe3fc380380380380380000000",
        'G': "0000000001f83fc3dc70e70870070073e73e70e70e70e39e3fe1f8000000",
        'H': "00000000070e70e70e70e70e70e7fe7fe7fe70e70e70e70e70e70e000000",
        'I': "0000000007fe7fe3fc0700700700700700700700700700f07fe7fe000000",
        'J': "0000000000fc0fc0fc01c01c01c01c01c01c01c31c71c7bc3f81f0000000",
        'K': "00000000070e71e71c7387387707f07f07f07b873c71c71e70e70f000000",
        'L': "0000000003803803803803803803803803803803803803803fe3fe000000",
        'M': "00000000078e79e79e79e79e7fe6f66f66f6676666606606606606000000",
        'N': "00000000078e78e7ce7ce7ce7ce7ee76e76e77e73e73e73e71e71e000000",
        'O': "0000000001f83fc39c70e70e70e70e70e70e70e70e70e79e3fc1f8000000",
        'P': "0000000007f87fc7fe70e70e70e70e7fe7fc7f0700700700700700000000",
        'Q': "0000000001f83fc39c70e70e70e70e70e70e70e70e70e79e3fc1f80f0070",
        'R': "0000000007f87fe7fe70e70e70e79e7fc7f87b873c71c71e70e70f000000",
        'S': "0000000001f83fc79e70e7007803f01fc07e00e00ef0e70e7fe1fc000000",
        'T': "0000000007fe7fe7fe0f00f00f00f00f00f00f00f00f00f00f00f0000000",
        'U': "00000000070e70e70e70e70e70e70e70e70e70e70e70e39c3fc1f8000000",
        'V': "000000000e07f0770e70e70e38c39c39c19c1f81f81f80f00f00f0000000",
        'W': "000000000e07e07e07e076666666f66f67fe7fe7fe79e39c39c39c000000",
        'X': "00000000070e70e39c39c1f81f00f00f00f01f81f839c79c70ee07000000",
        'Y': "000000000e0770e70e39c39c1f81f80f00f00e00e00e00e00e00e0000000",
        'Z': "0000000003fe3fe3fe01c0380780700e00e01c03c03807807fe7fe000000",
        '[': "0000f80fc0f80c00c00c00c00c00c00c00c00c00c00c00c00c00c00c00c0",
        '\\': "0006007003003801801c00c00e006007003003003801801c00c00e000000",
        ']': "0001f03f01f0030030030030030030030030030030030030030030030030",
        '^': "0000000f00f01f01b819839830c70c60e606000000000000000000000000",
        '_': "000000000000000000000000000000000000000000000000000000000000",
        '`': "0000000e0070030000000000000000000000000000000000000000000000",
        'a': "0000000000000000001f83f839c11c01c3fc3fc71c71c73c7fe3ef000000",
        'b': "0003007007007007007787fc79e78e70e70e70e70e78e79e7fc778000000",
        'c': "0000000000000000001f83fc39c70e70070070070070e38e3fc1f8000000",
        'd': "00000c00e00e00e00e1ee3fe79e71e70e70e70e70e71e79e3fe1ee000000",
        'e': "0000000000000000001f83fc39c70e70e7fe7fe70070078e3fc1f8000000",
        'f': "00007e0fe0e00e00c07fe7fe1e00c00c00c00c00c00c00c00c00c0000000",
        'g': "0000000000000000001ee3fe39e70e70e70e70e70e70e39e3fe1ee00e00e",
        'h': "00030070070070070077c7fc7de78e70e70e70e70e70e70e70e70e000000",
        'i': "0000700700700000003f03f00700700700700700700700707fe7fe000000",
        'j': "0000300380300000003f83f8078038038038038038038038038038030070",
        'k': "00030038038038038038e39c3b83b83f03f03f03b83b839c39e38e000000",
        'l': "0003f03f03f007007007007007007007007007007007007007e03e000000",
        'm': "0000000000000000007ce7fe776666666666666666666666666666000000",
        'n': "00000000000000000077c77c7de78e70e70e70e70e70e70e70e70e000000",
        'o': "0000000000000000001f83fc39e70e70e70e70e70e70e79e3fc1f8000000",
        'p': "0000000000000000007787fc79e78e70e70e70e70e70e79e7fc778700700",
        'q': "0000000000000000001ee3fe79e71e70e70e70e70e70e79e3fe3ee00e00e",
        'r': "0000000000000000003be3be3fc3c03c0380380380380380380380000000",
        's': "0000000000000000001f83fc39c3083c03f81fc03e00e70e3fc1f8000000",
        't': "0000000000c00c01c03fc3fc1c01c01c01c01c01c01c01c01fc0fc000000",
        'u': "00000000000000000070e70e70e70e70e70e70e70e71e79e3fe3ee000000",
        'v': "000000000000000000e0770e70e70e39c39c39c1b81f81f80f00f0000000",
        'w': "000000000000000000c03e07e076666666f66f66f679e79e39c39c000000",
        'x': "00000000000000000070e39c39c1f80f00f00f01f81f839c31c70e000000",
        'y': "00000000000000000070670e70e30c39c39c1981d81f80f00f00f00600e0",
        'z': "0000000000000000003fc3fc03c0380700700e01c03c03807fe7fe000000",
        '{': "00003c07c0fc0e00e00e00e00e00e01c03c03801c00e00e00e00e00e00e0",
        '|': "000060060060060060060060060060060060060060060060060060060060",
        '}': "0003c03e03f007007007007007007003803c01c038070070070070070070",
        '~': "0000000000000000000000000007e27fe47e000000000000000000000000",
        'Á': "0700600000f00f00f01f81f819839c39c39c3fc7fe7fe70ee07e07000000",
        'É': "0700600007fe7fe7fc7007007007fc7fc7fc7007007007007fe7fe000000",
        'Í': "0300600007fe7fe3fc0700700700700700700700700700f07fe7fe000000",
        'Ó': "0700600001f83fc39c70e70e70e70e70e70e70e70e70e79e3fc1f8000000",
        'Ú': "07006000070e70e70e70e70e70e70e70e70e70e70e70e39c3fc1f8000000",
        'Ü': "19819800070e70e70e70e70e70e70e70e70e70e70e70e39c3fc1f8000000",
        'Ñ': "3fc33800078e78e7ce7ce7ce7ce7ee76e76e77e73e73e73e71e71e000000",
        'á': "0000000380700400001f83f839c11c01c3fc3fc71c71c73c7fe3ef000000",
        'é': "0000000380700400001f83fc39c70e70e7fe7fe70070078e3fc1f8000000",
        'í': "0000000380700400003f03f00700700700700700700700707fe7fe000000",
        'ó': "0000000380700400001f83fc39e70e70e70e70e70e70e79e3fc1f8000000",
        'ú': "00000003807004000070e70e70e70e70e70e70e70e71e79e3fe3ee000000",
        'ü': "00000019819800000070e70e70e70e70e70e70e70e71e79e3fe3ee000000",
        'ñ': "0000841ec37c01000077c77c7de78e70e70e70e70e70e70e70e70e000000",
        '°': "0000000f01f81981081981f80f0000000000000000000000000000000000",
    },
    "3": {
        '!': "00000000000003c003c003c003c003c003c003c003c003c003c003c003c003c00000000003c003c003c003c000000000",
        '"': "000000001e781e781e781e781e781e700c700c300c300000000000000000000000000000000000000000000000000000",
        '#': "0000000000000208070c0618061806183ffe7ffe0e380c300c300c301c707ffe7ffe1860186038e030c030c000000000",
        '$': "00000100018007e01ff83ffc3dbc799c79903d803f801ff00ff803fc01be019e719e799e7dbe3ffc1ff80ff001800180",
        '%': "00000000000038067e0eee0cc618c638c630e660eee07cc0118003be037e06670ee30ce318e33867307e603e00000000",
        '&': "00000000000003c00fe00ff01c701c701c700fe00fe01f803f0c7f1c739cf3dcf1f8f0f8f0f87fff7fff1f8f00000000",
        "'": "0000000003c003c003c003c003c003800380018001800000000000000000000000000000000000000000000000000000",
        '(': "0000000000e001e001c003c00380078007800780070007000f000f000f000f000f000f000700078007800780038003c0",
        ')': "000000000f000780078003c003c001e001e001e001e000e000e000f000f000f000e000e000e001e001e001e001c003c0",
        '*': "000000000180018019901ff81ff807c007e00ff00e700460000000000000000000000000000000000000000000000000",
        '+': "000000000000000000000000038003800380038003807ffe7ffe7ffe03c0038003800380038001800000000000000000",
        ',': "00000000000000000000000000000000000000000000000000000000000000000000000007800780070007000f000e00",
        '-': "00000000000000000000000000000000000000000000000000000ff00ff00ff000000000000000000000000000000000",
        '.': "00000000000000000000000000000000000000000000000000000000000000000000000003c003c003c003c000000000",
        '/': "00000000001c001c003800380070007000e000e001c001c00380038007000f000e001e001c003c003800780000000000",
        '0': "00000000000003c00ff01ff83e783c3c3c3c383c381c7b9c7b9c7b9c799c381c383c3c3c3c781ff80ff007e000000000",
        '1': "00000000000001c003e007e03fe03fe039e001e001e001e001e001e001e001e001e001e001e03ffe3ffe3ffe00000000",
        '2': "00000000000003c00ff01ff83e7c3c3c383c003c007c00f801f003e007c00f801f003e003c007ffc7ffc7ffc00000000",
        '3': "00000000000003c00ff03ff83e7c3c3c183c003c007803f003e003f8007c001c081e781e3c3c3ffc1ff80ff000000000",
        '4': "000000000000007000f801f801f803f8077807780e781c783c78387870787ffe7ffe7ffe007800780078007800000000",
        '5': "0000000000001ff83ffc3ffc3c003c003c003cc03ff03ff83ffc3c3c003e001e001e783c3c3c3ff81ff80fe000000000",
        '6': "00000000000003e00ff01ff81e7c3c3c3c00380039f03ff83ffc3c3c3c3c381e3c1e3c3c1e3c1ffc0ff807f000000000",
        '7': "0000000000003ffc3ffc3ffc3ffc003c0078007000f001e001e003c003c007800780078007800f000f000f0000000000",
        '8': "00000000000003c00ff01ff83c3c3c3c3c3c3c3c3e781ff00ff01ff83c3c781c781e781e3c3c3ffc1ff80ff000000000",
        '9': "00000000000003c00ff01ff83e783c3c783c783c783c3c3c3efc3ffc1fdc071c003c383c3c783ff81ff00fe000000000",
        ':': "000000000000000000000000000003c003c003c003c0000000000000000000000000000003c003c003c003c000000000",
        ';': "000000000000000000000000000003c003c003c003c0000000000000000000000000000003c003c00380078007000700",
        '<': "000000000000000000000000001e007e03fc1ff07f807e0070007c007f001fe007f800fe003e00060000000000000000",
        '=': "00000000000000000000000000003ffc7ffe7ffe3ffc00000000000000007ffe7ffe7ffe000000000000000000000000",
        '>': "00000000000000000000400078007e003fc00ff001fc007e001e003e00fe07f81fe07f007c0060000000000000000000",
        '?': "00000000000003c00ff03ff83e7c783c781c003c003c00f801f003e003c0038000000000078007800780078000000000",
        '@': "0000000007f00f781c1c180c300631de63fe66726e326c324c32cc72cc62dc664c666ce66dec67386210300030081c38",
        'A': "00000000000003c007e007e007e00fe00ef00e701e701e781c383c383ffc3ffc3ffc7ffe781e701ef00ff00f00000000",
        'B': "0000000000003fc03ff83ffc3ffc3c3c3c3c3c3c3c7c3ff83ff83ffc3c1e3c1e3c1e3c1e3c3e3ffe3ffc3ff000000000",
        'C': "00000000000003c00ff01ff83f7c3c3c7c1e78007800780078007800780078007c1e3c1e3e3c1ffc0ff807f000000000",
        'D': "0000000000003f003fe03ff83ff83c7c3c3c3c1e3c1e3c1e3c1e3c1e3c1e3c1e3c3c3c3c3cfc3ff83ff03fc000000000",
        'E': "0000000000003ffc3ffc3ffc3ffc3c003c003c003c003ffc3ffc3ffc3c003c003c003c003c003ffe3ffe3ffe00000000",
        'F': "0000000000001ffc3ffc3ffc3ffc3c003c003c003c003c003ffc3ffc3ffc3c003c003c003c003c003c003c0000000000",
        'G': "00000000000003c00ff01ff83ffc3c3c7c3c78007800780079fe79fe78fe781e781e3c1e3e3e1ffc0ff807f000000000",
        'H': "000000000000381c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3ffc3ffc3ffc3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c00000000",
        'I': "0000000000003ffc3ffc3ffc3ffc03c003c003c003c003c003c003c003c003c003c003c003c03ffc3ffc3ffc00000000",
        'J': "00000000000003f807f807f803f800780078007800780078007800780078007878787c783cf83ff01ff00fc000000000",
        'K': "000000000000381e3c1e3c3c3c783cf83cf03de03fc03fc03fc03fe03ff03cf03cf83c783c7c3c3e3c1e3c1f00000000",
        'L': "0000000000001c001e001e001e001e001e001e001e001e001e001e001e001e001e001e001e001ffe1ffe1ffe00000000",
        'M': "000000000000381c7c3e7c3e7e7e7e7e7e7e7e6e77ee77ee77ee73ce73ce73ce73ce700e700e700e700e700e00000000",
        'N': "0000000000003c1c3e1c3e1c3e1c3f1c3f1c3f1c3f9c3b9c3b9c39dc39dc39fc38fc38fc38fc387c387c387c00000000",
        'O': "00000000000003c00ff01ff83ffc3c3c783e781e781e781e781e781e781e781e781e3c3c3e7c1ff81ff807e000000000",
        'P': "0000000000003f803ff03ffc3ffc3c3e3c1e3c1e3c1e3c3e3ffc3ff83ff03c003c003c003c003c003c003c0000000000",
        'Q': "00000000000003c00ff01ff83ffc3c3c783e781e781e781e781e781e781e781e781e7c3c3c7c3ff81ff80ff003c001e0",
        'R': "0000000000003f803ff83ffc3ffe3c3e3c1e3c1e3c1e3cfc3ffc3ff03ff03cf03c783c783c3c3c3e3c1e3c1f00000000",
        'S': "00000000000007c01ff03ffc3c7c783e78187c003f003fe00ff803fc007e001e301e781e7c3e3ffc1ff80ff000000000",
        'T': "0000000000007ffe7ffe7ffe7ffe03c003c003c003c003c003c003c003c003c003c003c003c003c003c003c000000000",
        'U': "000000000000381c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3e7c1ff81ff00fe000000000",
        'V': "000000000000f00ff00f781e781e781e383c3c3c3c3c1c381e781e780e700ef00ff00fe007e007e003c003c000000000",
        'W': "000000000000e007e00ff00ff00ff00e700e73ce73ce73ce7bde7ffe3ffc3ffc3e7c3e7c3e7c3e7c3c381c3800000000",
        'X': "000000000000700e781e3c3c3c3c1e780ff00ff007e007c003c007e00fe00ff01e781e783c3c783e781ef00f00000000",
        'Y': "000000000000700ef81e781e3c3c3c3c1e781e780ff00fe007e003c003c003c003c003c003c003c003c003c000000000",
        'Z': "0000000000003ffc3ffc3ffc3ffc007800f800f001e003e003c007800f800f001e003c007c007ffe7ffe7ffe00000000",
        '[': "0000000007f807f807800700070007000700070007000700070007000700070007000700070007000700070007000700",
        '\\': "00000000380038001c001c000e000e00070007000380038001c001c000e000e0007000700038003c001c001c00000000",
        ']': "000000001fe01fe001e000e000e000e000e000e000e000e000e000e000e000e000e000e000e000e000e000e000e000e0",
        '^': "0000000003c007c007e00ee00ef00e701c701c383838381c701c701e700e000000000000000000000000000000000000",
        '_': "00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000ffff",
        '`': "000000000f00038001c000e0000000000000000000000000000000000000000000000000000000000000000000000000",
        'a': "000000000000000000000000000007e01ff03ff83c78003800380ff81ff83c787838787878783ffc3fbe1f1e00000000",
        'b': "000000003c003c003c003c003c003df03ff83ffc3e3c3c3c3c1e3c1e3c1e3c1e3c1e3c3c3e3c3ffc3ff839f000000000",
        'c': "000000000000000000000000000007e00ff81ff83c3c3c3c780078007800780078003c3c3c3c1ff81ff807e000000000",
        'd': "00000000003c003c003c003c003c0fbc1ffc3ffc3c7c783c783c783c783c783c783c783c3c7c3ffc1ffc0fbc00000000",
        'e': "000000000000000000000000000007e01ff01e783c3c381c781c7ffe7ffe780078003c003c1c1ffc1ff807f000000000",
        'f': "000000fc03fe07fe0780078007803ffc7ffc3ffc07800780078007800780078007800780078007800780078000000000",
        'g': "0000000000000000000000000000079c1fdc3ffc3c3c3c3c3c1c381c381c381c3c3c3c3c3e3c1ffc1fdc0f9c001c003c",
        'h': "000000003c003c003c003c003c003cf03ff83ffc3e3c3c3c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c00000000",
        'i': "0000000001c001c001c0000000001fc01fc01fc001c001c001c001c001c001c001c001c001c03ffe7ffe7ffe00000000",
        'j': "0000000001e001e000e0000000001fe01fe01fe001e001e001e001e001e001e001e001e001e001e001e001e001e001e0",
        'k': "000000003c003c003c003c003c003c1c3c3c3c783cf03de03fc03fc03fe03fe03cf03c783c783c3c3c3c3c1e00000000",
        'l': "000000001fc01fc003c001c001c001c001c001c001c001c001c001c001c001c001c001c001e001fe00fe007e00000000",
        'm': "0000000000000000000000000000773c7fbe7ffe73ce73ce73ce73ce738e738e738e738e738e738e738e738e00000000",
        'n': "000000000000000000000000000038f03bf83ffc3e3c3c3c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c00000000",
        'o': "000000000000000000000000000007e01ff83ffc3c3c783c781e781e781e781e781e783c3c3c3ff81ff807e000000000",
        'p': "000000000000000000000000000038f03ff83ffc3e3c3c3c3c1e3c1e3c1e3c1e3c1e3c3c3e3c3ffc3ff83df03c003c00",
        'q': "00000000000000000000000000000f1c1ffc3ffc3c7c783c783c783c783c783c783c783c3c7c3ffc1ffc0fbc003c003c",
        'r': "00000000000000000000000000001c7c1cfc1ffc1fc01f001f001e001e001e001e001e001e001e001e001e0000000000",
        's': "000000000000000000000000000007e01ff83ffc3c3c3c003f001fe01ff807fc003c001c381c3ffc1ff80ff000000000",
        't': "00000000000003000700070007003ff83ff83ff80f000f000f000f000f000f000f000f000f000ff807f803f800000000",
        'u': "0000000000000000000000000000383c383c383c383c383c383c383c383c383c383c3c3c3c7c3ffc1fbc0f3c00000000",
        'v': "0000000000000000000000000000700e781e781e783c3c3c3c3c1c781e781e700ef00ef00fe007e007e007c000000000",
        'w': "0000000000000000000000000000e007e007e00e700e738e73ce73ce73ce77ec367c3e7c3e7c3c7c3c3c3c3800000000",
        'x': "0000000000000000000000000000781e3c3c3c781e780ff007e007e003c007e00ff00ef01e783c3c783c781e00000000",
        'y': "0000000000000000000000000000701e781e781e3c3c3c3c1c381e781e780e700f700ff007e007e003c003c003c00780",
        'z': "00000000000000000000000000003ff83ff83ff8007800f001e003e003c007800f001e003c003ffc3ffc3ffc00000000",
        '{': "0000000001fc03fc03c007800780078007800780078007800f801f003e003f000f800780078007800780078007800780",
        '|': "0000000003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c003c0",
        '}': "000000003f803fc003c001e001e001e001e001e001e001e001e000f8007800f801f001e001e001e001e001e001e001e0",
        '~': "00000000000000000000000000000000000000000c007f867ffe67fe007c000000000000000000000000000000000000",
        'Á': "01e00380000003c007e007e007e00fe00ef00e701e701e781c383c383ffc3ffc3ffc7ffe781e701ef00ff00f00000000",
        'É': "01e0038000003ffc3ffc3ffc3ffc3c003c003c003c003ffc3ffc3ffc3c003c003c003c003c003ffe3ffe3ffe00000000",
        'Í': "01e003c000003ffc3ffc3ffc3ffc03c003c003c003c003c003c003c003c003c003c003c003c03ffc3ffc3ffc00000000",
        'Ó': "01e00380000003c00ff01ff83ffc3c3c783e781e781e781e781e781e781e781e781e3c3c3e7c1ff81ff807e000000000",
        'Ú': "01e003800000381c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3e7c1ff81ff00fe000000000",
        'Ü': "0e700e700000381c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3e7c1ff81ff00fe000000000",
        'Ñ': "1ff818f000003c1c3e1c3e1c3e1c3f1c3f1c3f1c3f9c3b9c3b9c39dc39dc39fc38fc38fc38fc387c387c387c00000000",
        'á': "0000000000f001e003c00300000007e01ff03ff83c78003800380ff81ff83c787838787878783ffc3fbe1f1e00000000",
        'é': "00000000007800e001c00380000007e01ff01e783c3c381c781c7ffe7ffe780078003c003c1c1ffc1ff807f000000000",
        'í': "00000000007800e001c0038000001fc01fc01fc001c001c001c001c001c001c001c001c001c03ffe7ffe7ffe00000000",
        'ó': "00000000007800e001c00300000007e01ff83ffc3c3c783c781e781e781e781e781e783c3c3c3ff81ff807e000000000",
        'ú': "0000000000f001e001c003000000383c383c383c383c383c383c383c383c383c383c3c3c3c7c3ffc1fbc0f3c00000000",
        'ü': "000000000e700e700e7000000000383c383c383c383c383c383c383c383c383c383c3c3c3c7c3ffc1fbc0f3c00000000",
        'ñ': "000000000f981ff819f81860000038f03bf83ffc3e3c3c3c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c3c1c00000000",
        '°': "00000000038007e00ff00c300c300c300e7007e003c00000000000000000000000000000000000000000000000000000",
    },
}


def glyph_rects(font, ch):
    """Rectángulos (x, y, ancho, alto) de un carácter; None si no está en la tabla."""
    table = BITMAP_FONT.get(font)
    if not table or ch not in table:
        return None
    fw, fh = FONT_DOTS[font]
    hx = (fw + 3) // 4
    code = table[ch]
    open_runs, rects = {}, []
    for row in range(fh + 1):
        runs = set()
        if row < fh:
            value = int(code[row * hx:(row + 1) * hx], 16)
            bits = format(value, "0%db" % (hx * 4))[hx * 4 - fw:]
            x = 0
            while x < fw:
                if bits[x] == "1":
                    end = x
                    while end < fw and bits[end] == "1":
                        end += 1
                    runs.add((x, end - x))
                    x = end
                else:
                    x += 1
        for key in list(open_runs):
            if key not in runs:
                start = open_runs.pop(key)
                rects.append((key[0], start, key[1], row - start))
        for key in runs:
            open_runs.setdefault(key, row)
    return rects


def barcode_modules(value):
    """Bits (1=barra) del código real: EAN-13 si es válido, si no Code 128."""
    value = str(value).strip()
    bits = ean13_bits(value)
    if bits:
        return bits, "EAN-13", 0
    return code128_bits(value), "Code 128", 10  # zona de silencio de 10 módulos por lado


def barcode_elements(value, x_center, y, height, max_width, max_narrow=4):
    """Barras como rectángulos BAR; ancho de módulo automático según el espacio."""
    bits, kind, quiet = barcode_modules(value)
    total = len(bits) + 2 * max(quiet, 7 if kind == "EAN-13" else 0)
    narrow = max(1, min(max_narrow, max_width // total))
    width = len(bits) * narrow
    x0 = x_center - width // 2
    elements, index = [], 0
    while index < len(bits):
        if bits[index] == "1":
            end = index
            while end < len(bits) and bits[end] == "1":
                end += 1
            elements.append(("bar", x0 + index * narrow, y, (end - index) * narrow, height))
            index = end
        else:
            index += 1
    return elements, narrow, width


def wrap_to(text, font, max_width):
    chars = max(4, max_width // FONT_DOTS[font][0])
    return textwrap.wrap(text, chars, break_long_words=True) or [""]


def _clip(lines, count):
    if len(lines) <= count:
        return lines
    kept = lines[:count]
    kept[-1] = kept[-1][:-1].rstrip() + "…" if len(kept[-1]) > 1 else kept[-1]
    return kept


def _text_to_bars(elements):
    """Convierte los textos en rectángulos con la fuente de mapa de bits."""
    out = []
    for e in elements:
        if e[0] != "text":
            out.append(e)
            continue
        _, x, y, font, text, bold = e
        fw = FONT_DOTS[font][0]
        for index, ch in enumerate(text):
            for rx, ry, rw, rh in (glyph_rects(font, ch) or []):
                out.append(("bar", x + index * fw + rx, y + ry, rw, rh))
    return out


def _rotate_cw(elements, land_h):
    """Gira 90° (sentido horario) un diseño apaisado a vertical: (x, y) -> (land_h - y - h, x)."""
    out = []
    for e in elements:
        if e[0] == "bar":
            _, x, y, w, h = e
            out.append(("bar", land_h - y - h, x, h, w))
        elif e[0] == "box":
            _, x, y, w, h, t = e
            out.append(("box", land_h - y - h, x, h, w, t))
    return out


def small_label_elements(name, barcode, detail, width=160, height=240):
    """Etiqueta vertical 2 x 3 cm (160 x 240 puntos): solo código de barras y número, centrados.
    Barras verticales: el código girado (barras en el sentido del avance del papel) se
    empastaba en el papel térmico y no se podía escanear."""
    margin = 6
    usable = width - 2 * margin
    digits_h, gap, bar_h = 20, 8, 130
    top = (height - (bar_h + gap + digits_h)) // 2
    bars, narrow, bar_w = barcode_elements(barcode, width // 2, top, bar_h, usable)
    elements = list(bars)
    text_w = len(barcode) * 12
    elements.append(("text", max(2, (width - text_w) // 2), top + bar_h + gap, "2", barcode, False))
    return elements


def large_label_elements(name, barcode, detail, width=800, height=640):
    """Etiqueta 10 x 8 cm: solo información (tabla nutricional, ingredientes, alérgenos, datos legales).
    Sin código de barras: los baldes llevan el código de su otro programa."""
    margin = 20
    usable = width - 2 * margin
    data = LARGE_LABEL_LINES.get(name)
    # Marco que ocupa toda la etiqueta: así el driver no recorta la hoja («Save Paper») y la siguiente
    # etiqueta arranca en su lugar; además delimita visualmente la ficha.
    elements = [("box", 3, 3, width - 6, height - 6, 2)]
    y = 14

    def block(text, font, bold=False, gap=4, max_lines=None, pitch=2, x=margin, room=usable):
        nonlocal y
        lines = wrap_to(text, font, room)
        if max_lines:
            lines = _clip(lines, max_lines)
        for line in lines:
            elements.append(("text", x, y, font, line, bold))
            y += FONT_DOTS[font][1] + pitch
        y += gap

    if not isinstance(data, dict):  # producto sin ficha: nombre y detalle
        block(name, "3", True, 10, 2)
        elements.append(("bar", margin, y, usable, 3))
        y += 12
        if detail:
            block(detail, "2")
        return elements
    # cabecera: marca y producto
    block("GLACE", "3", True, 2)
    block(data["title"], "3", True, 6, 3)
    elements.append(("bar", margin, y, usable, 3))
    y += 9
    # tabla nutricional (porción, energía y pares de nutrientes)
    rows = data["nutri"]
    block("INFORMACIÓN NUTRICIONAL - " + rows[0], "2", True, 3)
    pairs = [r.split("    ", 1) if "    " in r else [r, ""] for r in rows[1:]]
    left_w = max(len(a.strip()) for a, _ in pairs[1:]) * 12 + 8
    cell_h = 25
    top = y
    elements.append(("box", margin, top, usable, cell_h * len(pairs) + 2, 2))
    for i, (cell_a, cell_b) in enumerate(pairs):
        ty = top + 1 + i * cell_h
        if i:
            elements.append(("bar", margin, ty, usable, 1))
        first = cell_a.strip() if i else (cell_a.strip() + "   " + cell_b.strip()).strip()
        elements.append(("text", margin + 4, ty + 3, "2", first, False))
        if cell_b and i:
            elements.append(("bar", margin + left_w, ty, 1, cell_h))
            elements.append(("text", margin + left_w + 4, ty + 3, "2", cell_b.strip(), False))
    y = top + cell_h * len(pairs) + 2 + 8
    # ingredientes
    block(data["ingred"], "1", True, 8, None, 3)
    # alérgenos destacados en recuadro
    box_top = y
    y += 6
    block(data["allerg"], "2", True, 0, 2, 2, margin + 8, usable - 16)
    elements.append(("box", margin, box_top, usable, y - box_top + 6, 2))
    y += 6 + 10
    # datos legales y contenido neto
    block(data["maker"], "1", True, 8, None, 3)
    block(data["neto"], "3", True, 6)
    block(data["elab"], "1", True, 0, None, 3)
    return elements


def physical_mm(format_name):
    """Medida real de la etiqueta en el rollo (ancho x alto en mm).

    «Etiqueta pequeña (3 x 2 cm)» es el nombre del catálogo; el rollo real trae etiquetas de
    2 cm de ancho x 3 cm de alto (en sentido de avance del papel).
    """
    if format_name == "Etiqueta pequeña (3 x 2 cm)":
        return 20, 30
    return FORMATS[format_name]


def label_size_dots(format_name):
    w_mm, h_mm = physical_mm(format_name)
    return w_mm * DPM, h_mm * DPM


def product_elements(name):
    product = PRODUCTS[name]
    width, height = label_size_dots(product["format"])
    if product["format"] == "Etiqueta grande (10 x 8 cm)":
        return large_label_elements(name, product["barcode"], product["detail"], width, height)
    return small_label_elements(name, print_code(name), product["detail"], width, height)


def _center_x(text, font, width=160):
    return max(2, (width - len(text) * FONT_DOTS[font][0]) // 2)


def coded_label_elements(code, top, bar_h, width=160, margin=6):
    """Código de barras con su número debajo (igual que las etiquetas de producto); el bloque empieza en «top»."""
    bars, _, _ = barcode_elements(code, width // 2, top, bar_h, width - 2 * margin)
    return list(bars) + [("text", _center_x(code, "2", width), top + bar_h + 8, "2", code, False)]


def insumo_label_elements(nombre, code, day):
    """Etiqueta interna de insumo (2 x 3 cm): nombre, código 25 + Nº + fecha de recepción, y la fecha legible."""
    elements = []
    y = 10
    for line in wrap_to(strip_accents(nombre).upper(), "1", 148)[:2]:
        elements.append(("text", _center_x(line, "1"), y, "1", line, True))
        y += 14
    top = 46
    elements += coded_label_elements(code, top, 104)
    stamp = f"REC {_as_date(day):%d/%m/%y}"
    elements.append(("text", _center_x(stamp, "1"), top + 104 + 8 + 20 + 8, "1", stamp, False))
    return elements


def serial_label_elements(title, subtitle, code):
    """Etiqueta con numeración correlativa (Promo Pack, ID de balde): título, código y su número."""
    elements = [("text", _center_x(title, "2"), 14, "2", title, True)]
    if subtitle:
        elements.append(("text", _center_x(subtitle, "1"), 40, "1", subtitle, False))
    return elements + coded_label_elements(code, 62, 110)


def test_set_labels(day=None, supplies=None):
    """Una etiqueta de cada familia (22, 29, 25, 23 y 24) para probar los escáneres. Llevan la marca «PRUEBA»."""
    day = _as_date(day)
    number = min(products_by_number())
    product = products_by_number()[number]
    padre = next(name for name, data in PRODUCTS.items() if data["barcode"][:2] == "29")
    supply_number, supply_name = (supplies or load_supplies())[0]
    rows = [
        ("Producto (22)", codigo_producto(number, day), small_label_elements(product, codigo_producto(number, day), "")),
        ("Grupo PADRE (29)", PRODUCTS[padre]["barcode"], small_label_elements(padre, PRODUCTS[padre]["barcode"], "")),
        ("Insumo (25)", codigo_insumo(supply_number, day),
         insumo_label_elements(supply_name, codigo_insumo(supply_number, day), day)),
        ("Promo Pack (23)", codigo_pack(1), serial_label_elements("PROMO PACK", "12 PALITOS", codigo_pack(1))),
        ("ID de balde (24)", codigo_id_balde(1), serial_label_elements("ID DE BALDE", "", codigo_id_balde(1))),
    ]
    tag = ("text", _center_x("PRUEBA", "1"), 218, "1", "PRUEBA", True)
    return [{"titulo": title, "codigo": code, "elementos": elements + [tag]} for title, code, elements in rows]


def elements_overflow(elements, width, height):
    """Devuelve elementos que salen de la etiqueta (para pruebas y avisos)."""
    bad = []
    for e in elements:
        if e[0] == "text":
            fw, fh = FONT_DOTS[e[3]]
            x1, y1 = e[1] + fw * len(e[4]), e[2] + fh
        else:
            x1, y1 = e[1] + e[3], e[2] + e[4]
        if e[1] < 0 or e[2] < 0 or x1 > width or y1 > height:
            bad.append(e)
    return bad


def _tspl_text(value):
    return value.replace("\\", "\\\\").replace('"', '\\"').replace("\r", " ").replace("\n", " ")


def _tspl_elements(elements, ox, oy, config):
    """Líneas TSPL (bytes) de los elementos de una etiqueta."""
    chunks = []
    for e in elements:
        if e[0] == "text":
            _, x, y, font, text, bold = e
            text = strip_accents(text) if config.get("strip_accents") else text
            for dx in ((0, 1) if bold else (0,)):
                chunks.append(f'TEXT {x + ox + dx},{y + oy},"{font}",0,1,1,"'.encode("ascii")
                              + _tspl_text(text).encode("cp1252", "replace") + b'"\r\n')
        elif e[0] == "bar":
            _, x, y, w, h = e
            chunks.append(f"BAR {x + ox},{y + oy},{w},{h}\r\n".encode("ascii"))
        elif e[0] == "box":
            _, x, y, w, h, t = e
            chunks.append(f"BOX {x + ox},{y + oy},{x + ox + w},{y + oy + h},{t}\r\n".encode("ascii"))
    return chunks


def build_tspl(elements, w_mm, h_mm, copies, config, labels=None):
    """Programa TSPL completo (bytes cp1252): N copias idénticas, o una serie («labels») con una etiqueta distinta cada una."""
    copies = max(1, min(999, int(copies)))
    ox, oy = int(config.get("offset_x", 0)), int(config.get("offset_y", 0))
    gap = float(config.get("gap_mm", 2.0))
    head = [f"SIZE {w_mm} mm,{h_mm} mm"]
    head.append(f"GAP {gap:g} mm,0 mm" if gap > 0 else "GAP 0 mm,0 mm")
    head += [f"DIRECTION {int(config.get('direction', 1))}", "REFERENCE 0,0",
             f"DENSITY {int(config.get('density', 8))}", f"SPEED {int(config.get('speed', 4))}"]
    if not config.get("strip_accents"):
        head.append("CODEPAGE 1252")
    head.append("CLS")
    chunks = [("\r\n".join(head) + "\r\n").encode("ascii")]
    if labels:
        for index, items in enumerate(labels):
            if index:
                chunks.append(b"CLS\r\n")
            chunks.extend(_tspl_elements(items, ox, oy, config))
            chunks.append(b"PRINT 1,1\r\n")
        return b"".join(chunks)
    chunks.extend(_tspl_elements(elements, ox, oy, config))
    chunks.append(f"PRINT 1,{copies}\r\n".encode("ascii"))
    return b"".join(chunks)


def test_label_elements():
    """Etiqueta de calibración 2 x 3 cm: marco + regla en mm + código de barras."""
    width, height = 160, 240
    elements = [("box", 0, 0, width, height, 2)]
    for mm in range(1, 20):
        elements.append(("bar", mm * DPM, 2, 1, 14 if mm % 5 == 0 else 7))
        if mm % 5 == 0:
            elements.append(("text", mm * DPM - len(str(mm)) * 4, 18, "1", str(mm), False))
    for mm in range(1, 30):
        elements.append(("bar", 2, mm * DPM, 14 if mm % 5 == 0 else 7, 1))
        if mm % 5 == 0:
            elements.append(("text", 18, mm * DPM - 6, "1", str(mm), False))
    elements.append(("text", 30, 60, "2", "PRUEBA", True))
    elements.append(("text", 30, 84, "2", "E-KARU", True))
    bars, _, _ = barcode_elements("2000001100014", width // 2, 120, 70, width - 16)
    elements += bars
    elements.append(("text", 28, 196, "1", "2000001100014", False))
    elements.append(("text", 28, 214, "1", f"{datetime.now():%d/%m %H:%M}", False))
    return elements


def strip_calibration_elements(config):
    """Una fila completa: regla de 0 a 100 mm + recuadros donde caería cada etiqueta según los ajustes."""
    cols = max(1, int(config.get("cols", 4)))
    left = float(config.get("margin_left_mm", 3.0))
    pitch = float(config.get("col_pitch_mm", 23.0))
    top = float(config.get("margin_top_mm", 1.0))
    row_mm = float(config.get("row_pitch_mm", 33.0))
    elements = []
    for mm in range(0, 101):
        x = mm * DPM
        tall = 18 if mm % 10 == 0 else (11 if mm % 5 == 0 else 5)
        elements.append(("bar", x, 0, 1 if mm % 10 else 2, tall))
        if mm % 10 == 0:
            label = str(mm)
            elements.append(("text", max(0, x - len(label) * 4), 20, "1", label, False))
    for mm in range(0, int(row_mm) + 1):
        y = mm * DPM
        elements.append(("bar", 0, y, 14 if mm % 10 == 0 else (9 if mm % 5 == 0 else 4), 1))
        if mm % 10 == 0 and mm:
            elements.append(("text", 18, y - 6, "1", str(mm), False))
    for column in range(cols):
        x = int((left + column * pitch) * DPM)
        y = int(top * DPM)
        elements.append(("box", x, y, 160, 240, 1))
        elements.append(("text", x + 60, y + 100, "4", str(column + 1), True))
    elements.append(("text", 200, 4, "1", f"cols={cols} izq={left:g} paso={pitch:g} fila={row_mm:g}", False))
    return elements


# --------------------------------------------------------------------------- #
#  Icono en el escritorio
# --------------------------------------------------------------------------- #
SHORTCUT_FLAG = BASE_DIR / ".acceso_directo_creado"


def create_desktop_shortcut():
    """Crea «E-Karu.lnk» en el escritorio (con icono propio). Devuelve (ok, mensaje)."""
    if os.name != "nt":
        return False, "Solo disponible en Windows."
    icon_src = BASE_DIR / "ekaru.ico"
    local = Path(os.environ.get("LOCALAPPDATA", str(BASE_DIR))) / "EKaru"
    try:
        local.mkdir(parents=True, exist_ok=True)
        icon = local / "ekaru.ico"
        if icon_src.exists():
            icon.write_bytes(icon_src.read_bytes())
        exe = Path(sys.executable)
        pythonw = exe.with_name("pythonw.exe")
        target = pythonw if pythonw.exists() else exe
        script = (
            "$d=[Environment]::GetFolderPath('DesktopDirectory');"
            "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\\E-Karu.lnk');"
            f"$s.TargetPath='{target}';"
            f"$q=[char]34;$s.Arguments=$q+'{BASE_DIR / 'app.py'}'+$q;"
            f"$s.WorkingDirectory='{BASE_DIR}';"
            f"$s.IconLocation='{icon}';"
            "$s.Description='E-Karu Etiquetas';$s.Save();Write-Output $d"
        )
        result = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                                capture_output=True, text=True, timeout=40, creationflags=0x08000000)
        if result.returncode != 0:
            raise OSError(result.stderr.strip() or f"PowerShell código {result.returncode}")
        try:
            SHORTCUT_FLAG.write_text("ok", encoding="utf-8")
        except OSError:
            pass
        return True, f"Icono creado en el escritorio: {result.stdout.strip()}"
    except (OSError, subprocess.SubprocessError) as error:
        log(f"No se pudo crear el acceso directo: {error}")
        return False, f"No se pudo crear el icono: {error}"


# --------------------------------------------------------------------------- #
#  Interfaz
# --------------------------------------------------------------------------- #
BG = "#071F1C"
PANEL = "#0D2F2B"
CARD = "#12403A"
CARD_HOVER = "#17504A"
LINE = "#1E5C54"
ACCENT = "#1FB8A2"
ACCENT_DARK = "#158F7E"
TEXT = "#E8F6F3"
MUTED = "#8FBDB4"
OK_COLOR = "#46D69A"
WARN_COLOR = "#F5B13D"
ERR_COLOR = "#FF6B6B"
TICKET = "#FFFFFF"
INK = "#111827"
UI_FONT = "Segoe UI"
MONO_FONT = "Consolas"


_MONO = []


def mono_family():
    """Fuente monoespaciada disponible (Consolas en Windows)."""
    if not _MONO:
        available = set(tkfont.families())
        _MONO.append(next((f for f in ("Consolas", "Courier New", "DejaVu Sans Mono", "Liberation Mono")
                           if f in available), "Courier"))
    return _MONO[0]


def format_short(format_name):
    w, h = physical_mm(format_name)
    return f"{w // 10} × {h // 10} cm"


def normalize(text):
    return strip_accents(str(text)).lower()


def draw_label(canvas, elements, label_w, label_h, box_w, box_h, pad=14):
    """Dibuja la etiqueta en el canvas exactamente como la recibe la impresora."""
    canvas.delete("all")
    scale = min((box_w - 2 * pad) / label_w, (box_h - 2 * pad) / label_h)
    ox = (box_w - label_w * scale) / 2
    oy = (box_h - label_h * scale) / 2
    canvas.create_rectangle(ox + 3, oy + 4, ox + label_w * scale + 3, oy + label_h * scale + 4,
                            fill="#041412", outline="")
    canvas.create_rectangle(ox, oy, ox + label_w * scale, oy + label_h * scale, fill=TICKET, outline="#CBD5E1")
    for e in elements:
        if e[0] == "text":
            _, x, y, font, text, bold = e
            fw, fh = FONT_DOTS[font]
            size = max(5, int(min(fw * scale / 0.56, fh * scale / 1.18)))
            for index, ch in enumerate(text):
                rects = glyph_rects(font, ch) if ch != " " else []
                if rects is None:  # carácter sin mapa de bits: texto normal
                    canvas.create_text(ox + (x + index * fw) * scale, oy + y * scale, anchor="nw", text=ch,
                                       fill=INK, font=(mono_family(), -size, "bold" if bold else "normal"))
                    continue
                # igual que la impresión: cada carácter son rectángulos negros (misma forma y posición)
                for rx, ry, rw, rh in rects:
                    canvas.create_rectangle(ox + (x + index * fw + rx) * scale, oy + (y + ry) * scale,
                                            ox + (x + index * fw + rx + rw + (1 if bold else 0)) * scale,
                                            oy + (y + ry + rh) * scale, fill=INK, outline="", width=0)
        elif e[0] == "bar":
            _, x, y, w, h = e
            canvas.create_rectangle(ox + x * scale, oy + y * scale, ox + (x + w) * scale,
                                    oy + (y + h) * scale, fill=INK, outline=INK, width=0)
        elif e[0] == "box":
            _, x, y, w, h, t = e
            canvas.create_rectangle(ox + x * scale, oy + y * scale, ox + (x + w) * scale,
                                    oy + (y + h) * scale, outline=INK, width=max(1, t * scale))
    return scale, ox, oy


_RUN_CACHE = {}


def barcode_runs(value):
    if value not in _RUN_CACHE:
        bits, kind, quiet = barcode_modules(value)
        runs, i = [], 0
        while i < len(bits):
            if bits[i] == "1":
                j = i
                while j < len(bits) and bits[j] == "1":
                    j += 1
                runs.append((i, j))
                i = j
            else:
                i += 1
        _RUN_CACHE[value] = (runs, len(bits))
    return _RUN_CACHE[value]


class CatalogList(tk.Canvas):
    """Lista virtualizada: dibuja solo las filas visibles sobre un único Canvas.

    No crea widgets por fila, por eso no hay parpadeo ni lentitud al desplazar.
    """

    def __init__(self, master, app):
        super().__init__(master, bg=PANEL, highlightthickness=0, bd=0)
        self.app = app
        s = app.s
        self.s = s
        self.row_h = int(76 * s)
        self.items = []
        self.offset = 0
        self.hover = None
        self.hits = []
        self._pending = False
        self._entry = None
        self._entry_name = None
        self._bar_state = None
        self._bg = {}
        self._imgs = {}
        self._rows = {}
        self._row_hits = {}
        self._drawn_w = None
        self._drawn_offset = 0
        self._force_full = True
        px = lambda n: -max(8, int(n * s))
        self.f_name = tkfont.Font(family=UI_FONT, size=px(15), weight="bold")
        self.f_detail = tkfont.Font(family=UI_FONT, size=px(12))
        self.f_small = tkfont.Font(family=UI_FONT, size=px(11))
        self.f_btn = tkfont.Font(family=UI_FONT, size=px(13), weight="bold")
        self.f_code = tkfont.Font(family=MONO_FONT, size=px(10))
        self.vbar = ctk.CTkScrollbar(master, command=self._on_scrollbar, fg_color=PANEL,
                                     button_color=LINE, button_hover_color=ACCENT_DARK)
        self.bind("<Configure>", lambda e: self.schedule())
        self.bind("<MouseWheel>", self._on_wheel)
        self.bind("<Button-4>", lambda e: self._scroll_by(-self.row_h))
        self.bind("<Button-5>", lambda e: self._scroll_by(self.row_h))
        self.bind("<Button-1>", self._on_click)
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", self._on_leave)

    # -- datos / desplazamiento ------------------------------------------------
    def set_items(self, names, keep_offset=False):
        self.items = list(names)
        if not keep_offset:
            self.offset = 0
        self._clamp()
        self.schedule()

    def _max_offset(self):
        return max(0, len(self.items) * self.row_h - self.winfo_height())

    def _clamp(self):
        self.offset = max(0, min(self._max_offset(), self.offset))

    def _scroll_by(self, delta):
        old = self.offset
        self.offset += delta
        self._clamp()
        if self.offset != old:
            self.schedule(full=False)

    def _on_wheel(self, event):
        step = -1 if event.delta > 0 else 1
        self._scroll_by(step * int(self.row_h * 0.9))

    def _on_scrollbar(self, *args):
        total = max(1, len(self.items) * self.row_h)
        if args[0] == "moveto":
            self.offset = int(float(args[1]) * total)
        elif args[0] == "scroll":
            unit = self.row_h if args[2] == "units" else self.winfo_height()
            self.offset += int(args[1]) * unit
        self._clamp()
        self.schedule(full=False)

    def ensure_visible(self, name):
        if name in self.items:
            top = self.items.index(name) * self.row_h
            if top < self.offset:
                self.offset = top
            elif top + self.row_h > self.offset + self.winfo_height():
                self.offset = top + self.row_h - self.winfo_height()
            self._clamp()
            self.schedule(full=False)

    # -- dibujo ----------------------------------------------------------------
    def schedule(self, full=True):
        if full:
            self._force_full = True
        if not self._pending:
            self._pending = True
            self.after_idle(self.redraw)

    def _fit(self, text, font, width):
        if font.measure(text) <= width:
            return text
        while text and font.measure(text + "…") > width:
            text = text[:-1]
        return text.rstrip() + "…"

    def redraw(self):
        self._pending = False
        self._end_edit(commit=True)
        width, height = self.winfo_width(), self.winfo_height()
        if width < 50 or height < 50:
            return
        s, rh = self.s, self.row_h
        total = max(1, len(self.items) * rh)
        bar = (round(self.offset / total, 4), round(min(1.0, (self.offset + height) / total), 4))
        if bar != self._bar_state:
            self._bar_state = bar
            self.vbar.set(*bar)
        if self._force_full or width != self._drawn_w or not self.items:
            self.delete("all")
            self._rows = {}
            self._row_hits = {}
            self._bg = {}
            self._force_full = False
            self._drawn_w = width
            self._drawn_offset = self.offset
        if not self.items:
            self.hits = []
            self.create_text(width / 2, height / 2, text="Ningún producto coincide con la búsqueda.",
                             fill=MUTED, font=self.f_detail)
            return
        first = self.offset // rh
        last = min(len(self.items), (self.offset + height) // rh + 2)
        wanted = set(range(first, last))
        # filas que ya no se ven: se borran; las que siguen: se desplazan (sin redibujar)
        for index in [i for i in self._rows if i not in wanted]:
            self.delete(f"row{index}")
            del self._rows[index]
            self._row_hits.pop(index, None)
        dy = self._drawn_offset - self.offset
        if dy and self._rows:
            self.move("all", 0, dy)
            for index, hits in self._row_hits.items():
                self._row_hits[index] = [(a, b + dy, c, d + dy, e, f) for a, b, c, d, e, f in hits]
        self._drawn_offset = self.offset
        for index in sorted(wanted):
            if index in self._rows:
                continue
            before = self.find_all()
            last_id = before[-1] if before else 0
            self.hits = []
            self._draw_row(self.items[index], index * rh - self.offset, width)
            self._row_hits[index] = self.hits
            self._rows[index] = True
            for item in self.find_all():
                if item > last_id:
                    self.addtag_withtag(f"row{index}", item)
        self.hits = [h for i in sorted(self._row_hits) for h in self._row_hits[i]]

    def _barcode_image(self, code, w, h):
        key = (code, w, h)
        img = self._imgs.get(key)
        if img is None:
            runs, modules = barcode_runs(code)
            cols = [TICKET] * w
            for a, b in runs:
                for x in range(int(a * w / modules), max(int(a * w / modules) + 1, int(b * w / modules))):
                    if 0 <= x < w:
                        cols[x] = INK
            row = "{" + " ".join(cols) + "}"
            img = tk.PhotoImage(width=w, height=h)
            img.put(" ".join([row] * h))
            self._imgs[key] = img
        return img

    def _draw_row(self, name, y, width):
        s, rh = self.s, self.row_h
        product = PRODUCTS[name]
        selected = name == self.app.selected
        x0, x1, y0, y1 = 6 * s, width - 6 * s, y + 4 * s, y + rh - 4 * s
        fill = CARD_HOVER if (name == self.hover and not selected) else CARD
        self._bg[name] = self.create_rectangle(x0, y0, x1, y1, fill=fill, outline=ACCENT if selected else LINE,
                                               width=2 if selected else 1)
        self.hits.append((x0, y0, x1, y1, "row", name))
        # etiqueta blanca con código de barras real
        tw = int(150 * s)
        tx0, tx1, ty0, ty1 = x0 + 8 * s, x0 + 8 * s + tw, y0 + 7 * s, y1 - 7 * s
        self.create_rectangle(tx0, ty0, tx1, ty1, fill=TICKET, outline="")
        inner = int(tw - 20 * s)
        bar_bottom = ty1 - 15 * s
        bar_h = max(8, int(bar_bottom - (ty0 + 5 * s)))
        code = print_code(name)
        if code:
            self.create_image(tx0 + 10 * s, ty0 + 5 * s, anchor="nw", image=self._barcode_image(code, inner, bar_h))
            self.create_text((tx0 + tx1) / 2, ty1 - 2 * s, text=spaced_code(code), fill=INK, font=self.f_code, anchor="s")
        else:
            # miniatura de la ficha informativa (sin código de barras)
            gx0, gx1 = tx0 + 8 * s, tx1 - 8 * s
            gy = ty0 + 6 * s
            self.create_rectangle(gx0, gy, gx1, gy + 4 * s, fill=INK, outline="")
            self.create_rectangle(gx0, gy + 7 * s, gx0 + (gx1 - gx0) * 0.7, gy + 10 * s, fill=INK, outline="")
            self.create_rectangle(gx0, gy + 15 * s, gx1, gy + 31 * s, fill="", outline=INK)
            for k in range(1, 4):
                self.create_line(gx0, gy + (15 + 4 * k) * s, gx1, gy + (15 + 4 * k) * s, fill=INK)
            for k in range(4):
                self.create_rectangle(gx0, gy + (36 + 3.2 * k) * s, gx1 - (k % 2) * 14 * s,
                                      gy + (37.6 + 3.2 * k) * s, fill="#94A3B8", outline="")
        # columna derecha: cantidad + imprimir
        bw, bh = int(96 * s), int(36 * s)
        cy = (y0 + y1) / 2
        bx1 = x1 - 12 * s
        bx0 = bx1 - bw
        self.create_rectangle(bx0, cy - bh / 2, bx1, cy + bh / 2, fill=ACCENT, outline="", tags="btn")
        self.create_text((bx0 + bx1) / 2, cy, text="Imprimir", fill="#04201C", font=self.f_btn)
        self.hits.append((bx0, cy - bh / 2, bx1, cy + bh / 2, "print", name))
        sw = int(28 * s)
        nw = int(42 * s)
        plus1 = bx0 - 14 * s
        plus0 = plus1 - sw
        num1 = plus0 - 2 * s
        num0 = num1 - nw
        minus1 = num0 - 2 * s
        minus0 = minus1 - sw
        for a, b, label, act in ((minus0, minus1, "−", "minus"), (plus0, plus1, "+", "plus")):
            self.create_rectangle(a, cy - bh / 2, b, cy + bh / 2, fill=PANEL, outline=LINE)
            self.create_text((a + b) / 2, cy, text=label, fill=TEXT, font=self.f_btn)
            self.hits.append((a, cy - bh / 2, b, cy + bh / 2, act, name))
        self.create_rectangle(num0, cy - bh / 2, num1, cy + bh / 2, fill=BG, outline=LINE)
        self.create_text((num0 + num1) / 2, cy, text=str(self.app.get_qty(name)), fill=TEXT, font=self.f_btn)
        self.hits.append((num0, cy - bh / 2, num1, cy + bh / 2, "qty", name))
        # textos
        text_x = tx1 + 16 * s
        room = minus0 - 16 * s - text_x
        self.create_text(text_x, y0 + 10 * s, anchor="nw", text=self._fit(name, self.f_name, room),
                         fill=TEXT, font=self.f_name)
        self.create_text(text_x, y0 + 32 * s, anchor="nw", text=self._fit(product["detail"], self.f_detail, room),
                         fill=MUTED, font=self.f_detail)
        meta = f"{barcode_kind(print_code(name))}  ·  {format_short(product['format'])}"
        self.create_text(text_x, y0 + 51 * s, anchor="nw", text=self._fit(meta, self.f_small, room),
                         fill=ACCENT, font=self.f_small)

    # -- interacción -----------------------------------------------------------
    def _hit(self, x, y):
        for hit in self.hits:
            if hit[0] <= x <= hit[2] and hit[1] <= y <= hit[3] and hit[4] != "row":
                return hit
        for hit in self.hits:
            if hit[0] <= x <= hit[2] and hit[1] <= y <= hit[3]:
                return hit
        return None

    def _on_motion(self, event):
        hit = self._hit(event.x, event.y)
        name = hit[5] if hit else None
        self.configure(cursor="hand2" if hit and hit[4] != "row" else "")
        if name != self.hover:
            self._set_hover(name)

    def _set_hover(self, name):
        """Solo recolorea las dos filas afectadas (sin redibujar la lista: sin parpadeo)."""
        old, self.hover = self.hover, name
        for key, color in ((old, CARD), (name, CARD_HOVER)):
            if key is not None and key != self.app.selected and key in self._bg:
                self.itemconfigure(self._bg[key], fill=color)

    def _on_leave(self, _event):
        if self.hover is not None:
            self._set_hover(None)

    def _on_click(self, event):
        self.focus_set()
        hit = self._hit(event.x, event.y)
        if not hit:
            return
        action, name = hit[4], hit[5]
        if action == "row":
            self.app.select_product(name)
        elif action == "plus":
            self.app.select_product(name)
            self.app.set_qty(name, self.app.get_qty(name) + 1)
        elif action == "minus":
            self.app.select_product(name)
            self.app.set_qty(name, self.app.get_qty(name) - 1)
        elif action == "print":
            self.app.select_product(name)
            self.app.print_product(name)
        elif action == "qty":
            self.app.select_product(name)
            self._begin_edit(name, hit)

    def _begin_edit(self, name, hit):
        self._end_edit(commit=False)
        entry = tk.Entry(self, justify="center", bg=BG, fg=TEXT, insertbackground=TEXT, relief="flat",
                         highlightthickness=1, highlightcolor=ACCENT, highlightbackground=LINE,
                         font=self.f_btn)
        entry.insert(0, str(self.app.get_qty(name)))
        entry.select_range(0, "end")
        self.create_window((hit[0] + hit[2]) / 2, (hit[1] + hit[3]) / 2, window=entry,
                           width=hit[2] - hit[0], height=hit[3] - hit[1])
        entry.focus_set()
        entry.bind("<Return>", lambda e: self._end_edit(commit=True))
        entry.bind("<KP_Enter>", lambda e: self._end_edit(commit=True))
        entry.bind("<Escape>", lambda e: self._end_edit(commit=False))
        entry.bind("<FocusOut>", lambda e: self._end_edit(commit=True))
        self._entry, self._entry_name = entry, name

    def _end_edit(self, commit):
        entry, name = self._entry, self._entry_name
        if entry is None:
            return
        self._entry = self._entry_name = None
        value = entry.get()
        try:
            entry.destroy()
        except tk.TclError:
            pass
        if commit:
            try:
                self.app.set_qty(name, int(value.strip()))
            except ValueError:
                self.app.set_status("Cantidad inválida: usá un número entre 1 y 999.", "warn")
                self.schedule()


class EkaruApp(ctk.CTk):
    def __init__(self, backend=None):
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.title(f"{APP_NAME}  {VERSION}")
        self.configure(fg_color=BG)
        self.s = float(self._get_window_scaling())
        self.backend = backend or make_backend()
        self.config_data = load_config()
        set_numbering(self.config_data.get("numeracion"))
        self._painted_day = None
        self.printers = []
        self.selected = next(iter(PRODUCTS))
        self.visible = list(PRODUCTS)
        self._busy = False
        self._busy_since = 0.0
        self._results = queue.Queue()
        self._search_job = None
        self.last_job = None
        self.editor_items = []
        self.editor_sel = None
        # Tamaño adaptado a la pantalla (barra de estado siempre visible)
        screen_w = int(self.winfo_screenwidth() / self.s)
        screen_h = int(self.winfo_screenheight() / self.s)
        win_w, win_h = min(1320, screen_w - 40), min(800, screen_h - 110)
        self.geometry(f"{win_w}x{win_h}+{max(0, (screen_w - win_w) // 2)}+{max(0, (screen_h - win_h) // 2 - 20)}")
        self.minsize(min(1000, win_w), min(520, win_h))
        self.compact = screen_h < 860
        if self.compact:
            self.after(200, self._maximize)
        self.fonts = {
            "title": ctk.CTkFont(family=UI_FONT, size=22, weight="bold"),
            "h2": ctk.CTkFont(family=UI_FONT, size=16, weight="bold"),
            "body": ctk.CTkFont(family=UI_FONT, size=13),
            "small": ctk.CTkFont(family=UI_FONT, size=12),
            "btn": ctk.CTkFont(family=UI_FONT, size=14, weight="bold"),
        }
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self.main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(0, weight=1)
        self.pages = {}
        self._build_catalog_page()
        self._build_editor_page()
        self.status_bar = ctk.CTkFrame(self.main, fg_color=PANEL, corner_radius=0, height=34)
        self.status_bar.grid(row=1, column=0, sticky="ew")
        self.status_label = ctk.CTkLabel(self.status_bar, text="", anchor="w", font=self.fonts["small"],
                                         text_color=MUTED)
        self.status_label.pack(fill="x", padx=16, pady=6)
        self.show_page("catalogo")
        self._paint_date()
        self._refresh_list()
        self.select_product(self.selected)
        self.bind("<Control-p>", lambda e: self.print_product(self.selected))
        self.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(50, self.refresh_printers)
        self.after(300, self._set_icon)
        if os.name == "nt" and not SHORTCUT_FLAG.exists():
            self.after(2500, lambda: threading.Thread(target=self._auto_shortcut, daemon=True).start())
        self.after(100, self._poll_results)
        self.after(30_000, self._day_tick)

    def _maximize(self):
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def _auto_shortcut(self):
        ok, text = create_desktop_shortcut()
        log(("OK  " if ok else "ERR ") + text)

    def make_shortcut(self):
        ok, text = create_desktop_shortcut()
        self.set_status(("✔ " if ok else "✖ ") + text, "ok" if ok else "error")

    def _set_icon(self):
        icon = BASE_DIR / "ekaru.ico"
        if os.name == "nt" and icon.exists():
            try:
                self.iconbitmap(str(icon))
            except tk.TclError:
                pass

    # -- barra lateral ---------------------------------------------------------
    def _build_sidebar(self):
        side = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=0, width=240)
        side.grid(row=0, column=0, sticky="nsew")
        side.grid_propagate(False)
        small = getattr(self, "compact", False)
        logo = ctk.CTkLabel(side, text="E-Karú", font=ctk.CTkFont(family=UI_FONT, size=24 if small else 30,
                                                                   weight="bold"), text_color=ACCENT)
        logo.pack(anchor="w", padx=22, pady=(14 if small else 26, 0))
        ctk.CTkLabel(side, text=f"Etiquetas  ·  {VERSION}", font=self.fonts["small"],
                     text_color=MUTED).pack(anchor="w", padx=22, pady=(0, 2))
        self.num_label = ctk.CTkLabel(side, text=self._numbering_text(), font=self.fonts["small"], text_color=ACCENT)
        self.num_label.pack(anchor="w", padx=22, pady=(0, 10 if small else 18))
        self.nav = {}
        for key, label in (("catalogo", "Catálogo de productos"), ("editor", "Editor de etiquetas")):
            button = ctk.CTkButton(side, text=label, anchor="w", height=32 if small else 42, corner_radius=10,
                                   font=self.fonts["btn"], fg_color="transparent", hover_color=CARD,
                                   text_color=TEXT, command=lambda k=key: self.show_page(k))
            button.pack(fill="x", padx=14, pady=2 if small else 3)
            self.nav[key] = button
        ctk.CTkLabel(side, text="HERRAMIENTAS", font=self.fonts["small"], text_color=MUTED).pack(
            anchor="w", padx=22, pady=(4 if small else 10, 0))
        for label, command in (("Etiquetas internas", self.open_internal_labels),
                               ("Leer un código", self.open_reader),
                               ("Juego de prueba", self.print_test_set)):
            ctk.CTkButton(side, text=label, anchor="w", height=30 if small else 40, corner_radius=10,
                          font=self.fonts["btn"], fg_color="transparent", hover_color=CARD, text_color=TEXT,
                          command=command).pack(fill="x", padx=14, pady=1 if small else 2)
        box = ctk.CTkFrame(side, fg_color=CARD, corner_radius=12)
        box.pack(side="bottom", fill="x", padx=14, pady=(0, 10), before=logo)
        ctk.CTkLabel(box, text="IMPRESORA", font=self.fonts["small"], text_color=MUTED).pack(
            anchor="w", padx=14, pady=(8, 2))
        self.printer_var = tk.StringVar(value="Buscando…")
        self.printer_menu = ctk.CTkOptionMenu(box, variable=self.printer_var, values=["Buscando…"], height=34,
                                              fg_color=PANEL, button_color=ACCENT_DARK,
                                              button_hover_color=ACCENT, dropdown_fg_color=PANEL,
                                              font=self.fonts["small"], dynamic_resizing=False,
                                              command=self._printer_changed)
        self.printer_menu.pack(fill="x", padx=12)
        self.printer_state = ctk.CTkLabel(box, text="", font=self.fonts["small"], text_color=MUTED,
                                          wraplength=190, justify="left", anchor="w")
        self.printer_state.pack(fill="x", padx=14, pady=(2, 0))
        self.mode_var = tk.StringVar(value="Driver" if self.config_data.get("mode", "gdi") == "gdi" else "TSPL")
        ctk.CTkLabel(box, text="MODO DE IMPRESIÓN", font=self.fonts["small"], text_color=MUTED).pack(
            anchor="w", padx=14, pady=(2, 2))
        ctk.CTkSegmentedButton(box, values=["Driver", "TSPL"], variable=self.mode_var, height=30,
                               selected_color=ACCENT_DARK, selected_hover_color=ACCENT, unselected_color=PANEL,
                               font=self.fonts["small"], command=self._mode_changed).pack(fill="x", padx=12)
        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(fill="x", padx=12, pady=(6, 10))
        ctk.CTkButton(row, text="Actualizar", width=84, height=32, fg_color=PANEL, hover_color=LINE,
                      font=self.fonts["small"], command=self.refresh_printers).pack(side="left")
        ctk.CTkButton(row, text="Diagnóstico", width=96, height=32, fg_color=ACCENT_DARK, hover_color=ACCENT,
                      font=self.fonts["small"], command=self.open_diagnostics).pack(side="right")

    def show_page(self, key):
        for name, frame in self.pages.items():
            if name == key:
                frame.grid(row=0, column=0, sticky="nsew")
            else:
                frame.grid_forget()
        for name, button in self.nav.items():
            button.configure(fg_color=CARD if name == key else "transparent")

    # -- página catálogo -------------------------------------------------------
    def _build_catalog_page(self):
        page = ctk.CTkFrame(self.main, fg_color=BG, corner_radius=0)
        self.pages["catalogo"] = page
        page.grid_columnconfigure(0, weight=1)
        page.grid_rowconfigure(1, weight=1)
        top = ctk.CTkFrame(page, fg_color="transparent")
        top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(18, 10))
        ctk.CTkLabel(top, text="Catálogo de productos", font=self.fonts["title"], text_color=TEXT).pack(side="left")
        self.count_label = ctk.CTkLabel(top, text="", font=self.fonts["small"], text_color=MUTED)
        self.count_label.pack(side="right", padx=(12, 0))
        self.category_var = tk.StringVar(value="Todas las categorías")
        categories = ["Todas las categorías"] + list(dict.fromkeys(product_category(n) for n in PRODUCTS))
        ctk.CTkOptionMenu(top, variable=self.category_var, values=categories, width=190, height=36,
                          fg_color=CARD, button_color=ACCENT_DARK, button_hover_color=ACCENT,
                          dropdown_fg_color=PANEL, font=self.fonts["body"],
                          command=lambda v: self._refresh_list()).pack(side="right", padx=8)
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(top, textvariable=self.search_var, width=300, height=36,
                                         placeholder_text="Buscar producto o código de barras…", placeholder_text_color=MUTED,
                                         fg_color=CARD, border_color=LINE, font=self.fonts["body"])
        self.search_entry.pack(side="right", padx=8)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        listbox = ctk.CTkFrame(page, fg_color=PANEL, corner_radius=12)
        listbox.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=(0, 16))
        listbox.grid_columnconfigure(0, weight=1)
        listbox.grid_rowconfigure(0, weight=1)
        self.catalog = CatalogList(listbox, self)
        self.catalog.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=6)
        self.catalog.vbar.grid(row=0, column=1, sticky="ns", padx=(2, 4), pady=6)
        panel = ctk.CTkFrame(page, fg_color=PANEL, corner_radius=12, width=360)
        panel.grid(row=1, column=1, sticky="ns", padx=(0, 20), pady=(0, 16))
        panel.grid_propagate(False)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(panel, text="Vista previa de impresión", font=self.fonts["h2"], text_color=TEXT).grid(
            row=0, column=0, sticky="w", padx=18, pady=(14, 6))
        self.preview = tk.Canvas(panel, bg=CARD, highlightthickness=0, height=int(110 * self.s))
        self.preview.grid(row=1, column=0, sticky="nsew", padx=14)
        self.preview.bind("<Configure>", lambda e: self._draw_preview())
        self.sel_name = ctk.CTkLabel(panel, text="", font=self.fonts["h2"], text_color=TEXT, wraplength=320,
                                     justify="left", anchor="w")
        self.sel_name.grid(row=2, column=0, sticky="ew", padx=18, pady=(10, 0))
        self.sel_info = ctk.CTkLabel(panel, text="", font=self.fonts["small"], text_color=MUTED, wraplength=320,
                                     justify="left", anchor="w")
        self.sel_info.grid(row=3, column=0, sticky="ew", padx=18, pady=(2, 8))
        drow = ctk.CTkFrame(panel, fg_color="transparent")
        drow.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 8))
        self.date_title = ctk.CTkLabel(drow, text="Fecha de elaboración", font=self.fonts["body"], text_color=TEXT)
        self.date_title.pack(side="left")
        ctk.CTkButton(drow, text="Hoy", width=44, height=34, fg_color=CARD, hover_color=LINE,
                      font=self.fonts["small"], command=self._date_today).pack(side="right")
        self.date_var = tk.StringVar(value=fecha_texto(effective_day()))
        self.date_entry = ctk.CTkEntry(drow, textvariable=self.date_var, width=104, height=34, justify="center",
                                       fg_color=BG, border_color=LINE, font=self.fonts["body"])
        self.date_entry.pack(side="right", padx=4)
        self.date_entry.bind("<Return>", lambda e: self._commit_date())
        self.date_entry.bind("<FocusOut>", lambda e: self._commit_date())
        qrow = ctk.CTkFrame(panel, fg_color="transparent")
        qrow.grid(row=5, column=0, sticky="ew", padx=18)
        ctk.CTkLabel(qrow, text="Cantidad de etiquetas", font=self.fonts["body"], text_color=TEXT).pack(side="left")
        self.qty_var = tk.StringVar(value="1")
        ctk.CTkButton(qrow, text="+", width=34, height=34, fg_color=CARD, hover_color=LINE, font=self.fonts["btn"],
                      command=lambda: self.set_qty(self.selected, self.get_qty(self.selected) + 1)).pack(side="right")
        self.qty_entry = ctk.CTkEntry(qrow, textvariable=self.qty_var, width=56, height=34, justify="center",
                                      fg_color=BG, border_color=LINE, font=self.fonts["btn"])
        self.qty_entry.pack(side="right", padx=4)
        ctk.CTkButton(qrow, text="−", width=34, height=34, fg_color=CARD, hover_color=LINE, font=self.fonts["btn"],
                      command=lambda: self.set_qty(self.selected, self.get_qty(self.selected) - 1)).pack(side="right")
        self.qty_entry.bind("<Return>", lambda e: self._commit_panel_qty())
        self.qty_entry.bind("<FocusOut>", lambda e: self._commit_panel_qty())
        self.print_button = ctk.CTkButton(panel, text="Imprimir etiquetas", height=44, corner_radius=10,
                                          fg_color=ACCENT, hover_color="#27D3BB", text_color="#04201C",
                                          font=self.fonts["btn"], command=self._print_selected)
        self.print_button.grid(row=6, column=0, sticky="ew", padx=18, pady=(12, 6))
        row = ctk.CTkFrame(panel, fg_color="transparent")
        row.grid(row=7, column=0, sticky="ew", padx=18, pady=(0, 14))
        ctk.CTkButton(row, text="Etiqueta de prueba", height=36, fg_color=CARD, hover_color=LINE,
                      font=self.fonts["small"], command=self.print_test).pack(side="left", expand=True, fill="x",
                                                                              padx=(0, 4))
        ctk.CTkButton(row, text="Guardar .prn", height=36, fg_color=CARD, hover_color=LINE,
                      font=self.fonts["small"], command=self.save_selected_prn).pack(side="left", expand=True,
                                                                                     fill="x", padx=(4, 0))

    def _on_search_key(self, _event=None):
        if self._search_job:
            self.after_cancel(self._search_job)
        self._search_job = self.after(120, self._refresh_list)

    def _refresh_list(self):
        self._search_job = None
        query = normalize(self.search_var.get()).strip()
        compact = "".join(query.split())
        # Si se escaneó o se pegó un código completo (13 o 21 dígitos), se busca el producto que ese código nombra,
        # sea de la numeración nueva (22…, con cualquier fecha) o de las anteriores.
        scanned = None
        if compact.isdigit() and len(compact) in (13, 21):
            scanned = interpretar_codigo(compact).get("producto")
        category = self.category_var.get()
        names = []
        for name, product in PRODUCTS.items():
            if category != "Todas las categorías" and product_category(name) != category:
                continue
            if query and name != scanned:
                haystack = normalize(f"{name} {product['detail']} {product['barcode']} {print_code(name)}")
                if query not in haystack and not (compact.isdigit() and compact in haystack.replace(" ", "")):
                    continue
            names.append(name)
        self.visible = names
        self.count_label.configure(text=f"{len(names)} de {len(PRODUCTS)} productos")
        self.catalog.set_items(names)
        if scanned in names and scanned != self.selected:
            self.select_product(scanned)
            self.catalog.ensure_visible(scanned)

    def select_product(self, name):
        self.selected = name
        product = PRODUCTS[name]
        code = print_code(name)
        self.sel_name.configure(text=name)
        self.sel_info.configure(text=f"{product['detail']}\n{barcode_kind(code)}  ·  "
                                     f"{spaced_code(code) or 'sin código'}  ·  {format_short(product['format'])}")
        # ¿este código lleva fecha? (dos días seguidos dan códigos distintos)
        dated = bool(code) and print_code(name, date(2026, 1, 2)) != print_code(name, date(2026, 1, 3))
        self.date_title.configure(text="Fecha de elaboración" if dated else "Fecha (este código no la usa)",
                                  text_color=TEXT if dated else MUTED)
        self.qty_var.set(str(self.get_qty(name)))
        self._draw_preview()
        self.catalog.schedule()

    def _draw_preview(self):
        width, height = self.preview.winfo_width(), self.preview.winfo_height()
        if width < 40 or height < 40:
            return
        label_w, label_h = label_size_dots(PRODUCTS[self.selected]["format"])
        draw_label(self.preview, product_elements(self.selected), label_w, label_h, width, height)

    # -- cantidades ------------------------------------------------------------
    def get_qty(self, name):
        return PRODUCT_QUANTITIES.get(name, 1)

    def set_qty(self, name, value):
        value = max(1, min(999, int(value)))
        PRODUCT_QUANTITIES[name] = value
        try:
            atomic_write_json(QUANTITIES_FILE, PRODUCT_QUANTITIES)
        except OSError as error:
            self.set_status(f"No se pudo guardar la cantidad: {error}", "warn")
        if name == self.selected:
            self.qty_var.set(str(value))
        self.catalog.schedule()

    def _commit_panel_qty(self):
        try:
            self.set_qty(self.selected, int(self.qty_var.get().strip()))
        except ValueError:
            self.qty_var.set(str(self.get_qty(self.selected)))
            self.set_status("Cantidad inválida: usá un número entre 1 y 999.", "warn")

    def _print_selected(self):
        self._commit_panel_qty()
        self.print_product(self.selected)

    # -- numeración de los códigos y fecha de elaboración -----------------------
    def _numbering_text(self):
        return "Códigos nuevos (v2)" if CODE_STATE["numeracion"] == NUMERACION_NUEVA else "Códigos anteriores (2.0)"

    def _codes_changed(self):
        """Cambió la numeración o la fecha: se rehacen la lista, el panel y la vista previa."""
        self._refresh_list()
        self.select_product(self.selected)

    def _apply_numbering(self, mode):
        set_numbering(mode)
        self.config_data["numeracion"] = CODE_STATE["numeracion"]
        self.num_label.configure(text=self._numbering_text())
        self._codes_changed()

    def _paint_date(self):
        """Muestra en el campo la fecha que usan los códigos (en ámbar si está fijada a mano)."""
        self._painted_day = effective_day()
        self.date_var.set(fecha_texto(self._painted_day))
        self.date_entry.configure(text_color=WARN_COLOR if CODE_STATE["dia"] else TEXT)

    def _date_today(self):
        set_elaboration_day(None)
        self._paint_date()
        self._codes_changed()
        self.set_status("Fecha de elaboración: hoy.", "info")

    def _commit_date(self):
        typed = self.date_var.get().strip()
        if typed == fecha_texto(effective_day()):
            return  # sin cambios (también pasa al salir del campo)
        day, today = parse_user_date(typed), date.today()
        problem = None
        if day is None:
            problem = "Fecha no válida: escribí día/mes/año, por ejemplo 12/11/2026."
        elif day > today:
            problem = "La fecha de elaboración no puede ser futura: e-Karú rechazaría esos códigos (regla 4)."
        elif (today - day).days > MAX_DIAS_ATRAS:
            problem = f"Esa fecha es de hace más de {MAX_DIAS_ATRAS} días: revisá el año."
        if problem:
            self._paint_date()
            self.set_status(problem, "warn")
            return
        set_elaboration_day(day)
        self._paint_date()
        self._codes_changed()
        if CODE_STATE["dia"]:
            self.set_status(f"Fecha de elaboración fijada en {fecha_texto(day)}: los códigos con fecha usan ese día "
                            "hasta que toques «Hoy».", "info")
        else:
            self.set_status("Fecha de elaboración: hoy.", "info")

    def _day_tick(self):
        """Pasada la medianoche, si la fecha no está fijada, los códigos pasan solos al día nuevo."""
        try:
            if CODE_STATE["dia"] is None and self._painted_day != date.today():
                self._paint_date()
                self._codes_changed()
                self.set_status("Cambió el día: los códigos llevan la fecha de hoy.", "info")
        except Exception:  # noqa: BLE001 - el reloj nunca debe morir
            log("ERROR al cambiar de día: " + traceback.format_exc())
        self.after(30_000, self._day_tick)

    def _numbering_notice(self):
        """Primera impresión de productos con la numeración nueva: se explica el cambio y se pide confirmarlo.

        Devuelve True si se puede imprimir. «Sí» = numeración nueva · «No» = seguir con la anterior · «Cancelar» = no imprimir.
        """
        if self.config_data.get("aviso_numeracion") or CODE_STATE["numeracion"] != NUMERACION_NUEVA:
            return True
        example = products_by_number()[min(products_by_number())]
        new = spaced_code(print_code(example, mode=NUMERACION_NUEVA))
        old = spaced_code(print_code(example, mode=NUMERACION_ACTUAL))
        answer = messagebox.askyesnocancel(
            APP_NAME,
            "Esta versión imprime los códigos con la numeración nueva del Informe final.\n\n"
            f"{example}\n   nuevo:   {new}   (22 + Nº de producto + fecha + control)\n   antes:   {old}\n\n"
            "Las etiquetas viejas que ya están pegadas se reconocen en «Leer un código».\n"
            "Importante: la caja (e-Karú) tiene que tener cargados los códigos nuevos antes de vender con ellos.\n\n"
            "Sí = usar la numeración nueva\nNo = seguir con la anterior\nCancelar = no imprimir ahora\n\n"
            "Se puede cambiar cuando quieras en Diagnóstico → Ajustes.", parent=self)
        if answer is None:
            self.set_status("No se imprimió: elegí la numeración de los códigos cuando quieras.", "info")
            return False
        if answer is False:
            self._apply_numbering(NUMERACION_ACTUAL)
            self.set_status("Se sigue con la numeración anterior (2.0).", "info")
        self.config_data["aviso_numeracion"] = True
        self.save_config()
        log(f"AVISO de numeración confirmado: {CODE_STATE['numeracion']}")
        return True

    def export_transition(self, parent=None):
        """Guarda la tabla de transición (47 productos + 9 grupos PADRE) como CSV para Excel."""
        parent = parent or self
        path = filedialog.asksaveasfilename(parent=parent, defaultextension=".csv",
                                            initialfile="tabla_transicion_codigos.csv",
                                            filetypes=[("CSV para Excel", "*.csv"), ("Todos", "*.*")])
        if not path:
            return
        try:
            write_transition_csv(path, effective_day())
            self.set_status(f"✔ Tabla de transición guardada: {path}", "ok")
        except OSError as error:
            messagebox.showerror(APP_NAME, f"No se pudo guardar:\n{error}", parent=parent)

    # -- impresoras ------------------------------------------------------------
    def refresh_printers(self):
        self.printer_state.configure(text="Buscando impresoras…", text_color=MUTED)

        def work():
            try:
                names = self.backend.list_printers()
                default = self.backend.default_printer()
                self._results.put(("printers", names, default, None))
            except Exception as error:  # noqa: BLE001
                log(f"Error al listar impresoras: {error!r}")
                self._results.put(("printers", [], "", str(error)))

        threading.Thread(target=work, daemon=True).start()

    def _apply_printers(self, names, default, error):
        self.printers = names
        if error:
            self.printer_menu.configure(values=["(error)"])
            self.printer_var.set("(error)")
            self.printer_state.configure(text="No se pudo leer la lista de impresoras. Abrí «Diagnóstico».",
                                         text_color=ERR_COLOR)
            self.set_status(f"Error al listar impresoras: {error}", "error")
            return
        if not names:
            self.printer_menu.configure(values=["(ninguna)"])
            self.printer_var.set("(ninguna)")
            self.printer_state.configure(text="Windows no tiene impresoras instaladas.", text_color=ERR_COLOR)
            self.set_status("No hay impresoras instaladas en Windows.", "error")
            return
        saved = self.config_data.get("printer", "")
        pick = saved if saved in names else next(
            (n for n in names if "hprt" in n.lower() or "ht400" in n.lower()), default if default in names else names[0])
        self.printer_menu.configure(values=names)
        self.printer_var.set(pick)
        self._printer_changed(pick)

    def _printer_changed(self, name):
        self.config_data["printer"] = name
        self.save_config()
        self.printer_state.configure(text="Comprobando…", text_color=MUTED)

        def work():
            try:
                ok, text = self.backend.status(name)
            except Exception as error:  # noqa: BLE001
                ok, text = False, str(error)
            self._results.put(("pstate", name, ok, text))

        threading.Thread(target=work, daemon=True).start()

    def _mode_changed(self, value):
        self.config_data["mode"] = "gdi" if value == "Driver" else "raw"
        self.save_config()
        self.set_status("Modo: " + ("driver de Windows (como Ctrl+P)" if value == "Driver" else "TSPL directo (RAW)"), "info")

    def save_config(self):
        try:
            atomic_write_json(CONFIG_FILE, self.config_data)
        except OSError as error:
            log(f"No se pudo guardar la configuración: {error}")

    # -- estado / impresión ----------------------------------------------------
    def set_status(self, text, level="info"):
        color = {"info": MUTED, "ok": OK_COLOR, "warn": WARN_COLOR, "error": ERR_COLOR}.get(level, MUTED)
        self.status_label.configure(text=text, text_color=color)

    def print_product(self, name, quantity=None):
        quantity = quantity or self.get_qty(name)
        if not self._numbering_notice():
            return
        job = self.make_job(name, quantity)
        log(f"CLICK imprimir «{name}» × {quantity} · código {print_code(name) or '(sin código)'} "
            f"(numeración {CODE_STATE['numeracion']}, modo {self.config_data.get('mode')})")
        self.submit(job, f"{name} × {quantity}")

    def make_job(self, name, quantity):
        product = PRODUCTS[name]
        w_mm, h_mm = physical_mm(product["format"])
        small = product["format"] == "Etiqueta pequeña (3 x 2 cm)"
        page = "grid" if small and self.config_data.get("grid") else "label"
        return PrintJob(product_elements(name), w_mm, h_mm, quantity, self.config_data, page)

    def print_test(self):
        page = "grid" if self.config_data.get("grid") else "label"
        job = PrintJob(test_label_elements(), 20, 30, 1, self.config_data, page)
        self.submit(job, "etiqueta de prueba")

    def print_calibration(self):
        job = PrintJob(strip_calibration_elements(self.config_data), 100, self.config_data.get("row_pitch_mm", 33.0),
                       1, self.config_data, "strip")
        self.submit(job, "calibración de fila")

    def save_selected_prn(self):
        name = self.selected
        self.save_prn(self.make_job(name, self.get_qty(name)).data, name)

    def save_prn(self, data, hint="etiqueta"):
        if not data:
            messagebox.showinfo(APP_NAME, "Todavía no hay ningún trabajo para guardar.", parent=self)
            return
        safe = "".join(c if c.isalnum() else "_" for c in strip_accents(hint))[:40]
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".prn", initialfile=f"{safe}.prn",
                                            filetypes=[("Archivo de impresora", "*.prn"), ("Todos", "*.*")])
        if path:
            try:
                Path(path).write_bytes(data)
                self.set_status(f"Guardado: {path}", "ok")
            except OSError as error:
                messagebox.showerror(APP_NAME, f"No se pudo guardar:\n{error}", parent=self)

    def printer_for(self, job):
        """Etiquetas grandes (10 x 8 cm) van a la impresora FTX; las chicas, a la elegida en la barra lateral."""
        chosen = self.printer_var.get()
        if getattr(job, "w_mm", 0) < 90:
            return chosen
        saved = self.config_data.get("printer_large", "")
        if saved in self.printers:
            return saved
        for name in self.printers:
            if "ftx" in name.lower() or "ttp010" in name.lower():
                return name
        return chosen

    def submit(self, job, label, force=False):
        if self._busy and time.time() - self._busy_since < 20:
            log(f"IGNORADO {label}: hay un envío en curso")
            self.set_status("Ya hay un envío en curso, esperá un momento…", "warn")
            return
        printer = self.printer_for(job)
        if printer not in self.printers:
            log(f"IGNORADO {label}: impresora «{printer}» no está en la lista {self.printers}")
            messagebox.showwarning(APP_NAME, "Elegí una impresora en la barra lateral (o abrí «Diagnóstico»).",
                                   parent=self)
            return
        self._busy = True
        self._busy_since = time.time()
        self.last_job = job.data
        self.set_status(f"Enviando {label} a «{printer}»…", "info")

        def work():
            try:
                if not force:
                    ok, text = self.backend.status(printer)
                    if not ok:
                        log(f"BLOQUEADO {label}: {text}")
                        self._results.put(("blocked", label, job, text))
                        return
                detail = self.backend.send(printer, job, f"E-Karu {label}"[:60])
                log(f"OK  {label}: {detail}")
                self._results.put(("sent", label, detail, job.on_success))
            except PrintError as error:
                log(f"ERR {label}: {error}")
                self._results.put(("failed", label, str(error)))
            except Exception as error:  # noqa: BLE001
                log(f"ERR {label}: {traceback.format_exc()}")
                self._results.put(("failed", label, f"Error inesperado: {error!r}"))

        threading.Thread(target=work, daemon=True).start()

    def _poll_results(self):
        try:
            self._poll_once()
        except Exception:  # noqa: BLE001 - el bucle nunca debe morir
            log("ERROR en el bucle de resultados: " + traceback.format_exc())
            self._busy = False
        self.after(100, self._poll_results)

    def _poll_once(self):
        try:
            while True:
                kind, *rest = self._results.get_nowait()
                if kind == "printers":
                    self._apply_printers(*rest)
                elif kind == "pstate":
                    name, ok, text = rest
                    if name == self.printer_var.get():
                        self.printer_state.configure(text=("● " + text) if ok else ("▲ " + text),
                                                     text_color=OK_COLOR if ok else WARN_COLOR)
                elif kind == "sent":
                    self._busy = False
                    self.set_status(f"✔ {rest[0]}: {rest[1]}", "ok")
                    callback = rest[2] if len(rest) > 2 else None
                    if callback:  # por ejemplo, avanzar el contador de Promo Pack solo si el envío salió bien
                        try:
                            callback()
                        except Exception:  # noqa: BLE001
                            log("ERROR al registrar el envío: " + traceback.format_exc())
                elif kind == "failed":
                    self._busy = False
                    self.set_status(f"✖ {rest[0]}: {rest[1]}", "error")
                    messagebox.showerror("No se pudo imprimir",
                                         f"{rest[1]}\n\nAbrí «Diagnóstico» para ver el estado de la impresora "
                                         "y probar una etiqueta de prueba.", parent=self)
                elif kind == "blocked":
                    self._busy = False
                    label, job, text = rest
                    self.set_status(f"▲ La impresora no está lista: {text}", "warn")
                    if messagebox.askyesno("Impresora no lista",
                                           f"{text}\n\nWindows puede retener o descartar el trabajo.\n"
                                           "¿Enviar de todos modos?", parent=self):
                        self.submit(job, label, force=True)
        except queue.Empty:
            pass

    # -- diagnóstico y ajustes -------------------------------------------------
    def build_report(self):
        lines = [f"{APP_NAME} {VERSION}", f"Python {sys.version.split()[0]}  ({Path(sys.executable).name})",
                 f"Sistema: {sys.platform}", f"Modo de envío: {self.backend.name}",
                 f"Carpeta de la app: {BASE_DIR}", "", "Impresoras detectadas:"]
        if not self.printers:
            lines.append("  (ninguna)")
        for name in self.printers:
            try:
                ok, text = self.backend.status(name)
            except Exception as error:  # noqa: BLE001
                ok, text = False, str(error)
            mark = "✔" if ok else "▲"
            lines.append(f"  {mark} {name}  —  {text}")
        lines += ["", f"Seleccionada: {self.printer_var.get()}", "Ajustes: " + json.dumps(self.config_data, ensure_ascii=False)]
        counters = load_counters()
        lines += ["", f"Numeración de los códigos: {NUMERACION_TEXTO[CODE_STATE['numeracion']]}",
                  "Fecha de elaboración en los códigos: " + fecha_texto(effective_day())
                  + (" (fijada a mano)" if CODE_STATE["dia"] else " (hoy)"),
                  f"Contadores: Promo Pack {counters['pack']} · ID de balde {counters['balde']} · "
                  f"insumos cargados: {len(load_supplies())}"]
        lines += ["", "Últimos eventos:"]
        try:
            lines += ["  " + l for l in LOG_FILE.read_text(encoding="utf-8").splitlines()[-12:]] or ["  (sin eventos)"]
        except OSError:
            lines.append("  (sin eventos)")
        return "\n".join(lines)

    def open_diagnostics(self):
        win = ctk.CTkToplevel(self)
        win.title("Diagnóstico de impresión")
        win.geometry("760x600")
        win.configure(fg_color=BG)
        win.transient(self)
        text = ctk.CTkTextbox(win, fg_color=PANEL, text_color=TEXT, font=ctk.CTkFont(family=MONO_FONT, size=12),
                              wrap="word")
        text.pack(fill="both", expand=True, padx=16, pady=(16, 8))
        text.insert("1.0", self.build_report())
        text.configure(state="disabled")
        row = ctk.CTkFrame(win, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 16))

        def refresh():
            text.configure(state="normal")
            text.delete("1.0", "end")
            text.insert("1.0", self.build_report())
            text.configure(state="disabled")

        def copy():
            self.clipboard_clear()
            self.clipboard_append(self.build_report())
            self.set_status("Informe copiado al portapapeles.", "ok")

        def queue_window():
            if os.name == "nt":
                subprocess.Popen(["control", "printers"])
            else:
                messagebox.showinfo(APP_NAME, "Solo disponible en Windows.", parent=win)

        buttons = (("Etiqueta de prueba", self.print_test), ("Calibrar fila", self.print_calibration),
                   ("Guardar último trabajo", lambda: self.save_prn(self.last_job)),
                   ("Cola de Windows", queue_window), ("Ajustes", self.open_settings), ("Icono en el escritorio", self.make_shortcut),
                   ("Tabla de transición (CSV)", lambda: self.export_transition(win)),
                   ("Copiar informe", copy),
                   ("Actualizar", lambda: (self.refresh_printers(), win.after(800, refresh))))
        for i, (label, cmd) in enumerate(buttons):
            ctk.CTkButton(row, text=label, height=36, fg_color=ACCENT_DARK if i == 0 else CARD, hover_color=ACCENT,
                          font=self.fonts["small"], command=cmd).grid(row=i // 3, column=i % 3, padx=4, pady=4,
                                                                       sticky="ew")
            row.grid_columnconfigure(i % 3, weight=1)
        win.after(150, win.lift)

    def open_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Ajustes de impresión")
        win.geometry("480x680")
        win.configure(fg_color=BG)
        win.transient(self)
        holder = ctk.CTkFrame(win, fg_color="transparent")
        holder.pack(side="bottom", fill="x")
        body = ctk.CTkScrollableFrame(win, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True)
        fields = {}
        spec = (("cols", "Etiquetas por fila en el rollo (columnas)"),
                ("col_pitch_mm", "Distancia entre columnas (mm, de borde izquierdo a borde izquierdo)"),
                ("margin_left_mm", "Margen izquierdo de la 1.ª columna (mm)"),
                ("margin_top_mm", "Margen superior (mm)"),
                ("row_pitch_mm", "Alto de una fila: etiqueta + separación (mm)"),
                ("offset_x", "Desplazamiento horizontal fino (puntos)"),
                ("offset_y", "Desplazamiento vertical fino (puntos)"),
                ("gap_mm", "Separación entre etiquetas (mm, solo modo TSPL)"),
                ("density", "Densidad (0-15, solo modo TSPL)"), ("speed", "Velocidad (1-6, solo modo TSPL)"),
                ("direction", "Dirección (0 o 1, solo modo TSPL)"))
        for key, label in spec:
            ctk.CTkLabel(body, text=label, font=self.fonts["small"], text_color=MUTED, anchor="w",
                         wraplength=400, justify="left").pack(fill="x", padx=14, pady=(10, 0))
            entry = ctk.CTkEntry(body, height=34, fg_color=CARD, border_color=LINE)
            entry.insert(0, str(self.config_data[key]))
            entry.pack(fill="x", padx=14)
            fields[key] = entry
        ctk.CTkLabel(body, text="Modo de impresión", font=self.fonts["small"], text_color=MUTED, anchor="w").pack(
            fill="x", padx=14, pady=(12, 0))
        mode_var = tk.StringVar(value="Driver de Windows (como Ctrl+P)" if self.config_data.get("mode", "gdi") == "gdi"
                                else "TSPL directo (RAW)")
        ctk.CTkOptionMenu(body, variable=mode_var, values=["Driver de Windows (como Ctrl+P)", "TSPL directo (RAW)"],
                          fg_color=CARD, button_color=ACCENT_DARK, button_hover_color=ACCENT,
                          dropdown_fg_color=PANEL).pack(fill="x", padx=14)
        ctk.CTkLabel(body, text="Numeración de los códigos de barras de los productos", font=self.fonts["small"],
                     text_color=MUTED, anchor="w").pack(fill="x", padx=14, pady=(12, 0))
        num_var = tk.StringVar(value=NUMERACION_TEXTO[CODE_STATE["numeracion"]])
        ctk.CTkOptionMenu(body, variable=num_var, values=list(NUMERACION_TEXTO.values()), fg_color=CARD,
                          button_color=ACCENT_DARK, button_hover_color=ACCENT,
                          dropdown_fg_color=PANEL).pack(fill="x", padx=14)
        ctk.CTkLabel(body, text="«Nueva» sigue el Informe final (22 + Nº de producto + fecha). «Anterior» es la de la "
                                "versión 2.0. Los grupos PADRE y los baldes no cambian.", font=self.fonts["small"],
                     text_color=MUTED, anchor="w", wraplength=400, justify="left").pack(fill="x", padx=14, pady=(4, 0))
        grid_switch = ctk.CTkSwitch(body, text="Llenar las columnas del rollo (varias etiquetas por fila)",
                                    font=self.fonts["small"], progress_color=ACCENT)
        if self.config_data.get("grid"):
            grid_switch.select()
        grid_switch.pack(anchor="w", padx=14, pady=(14, 4))
        accents = ctk.CTkSwitch(body, text="Quitar tildes al imprimir (si salen símbolos raros)",
                                font=self.fonts["small"], progress_color=ACCENT)
        if self.config_data.get("strip_accents"):
            accents.select()
        accents.pack(anchor="w", padx=14, pady=(8, 16))

        def save():
            try:
                new = {
                    "cols": max(1, min(8, int(fields["cols"].get()))),
                    "col_pitch_mm": max(5.0, min(120.0, float(fields["col_pitch_mm"].get().replace(",", ".")))),
                    "margin_left_mm": max(0.0, min(60.0, float(fields["margin_left_mm"].get().replace(",", ".")))),
                    "margin_top_mm": max(0.0, min(30.0, float(fields["margin_top_mm"].get().replace(",", ".")))),
                    "row_pitch_mm": max(10.0, min(200.0, float(fields["row_pitch_mm"].get().replace(",", ".")))),
                    "gap_mm": max(0.0, min(10.0, float(fields["gap_mm"].get().replace(",", ".")))),
                    "density": max(0, min(15, int(fields["density"].get()))),
                    "speed": max(1, min(6, int(fields["speed"].get()))),
                    "direction": 1 if int(fields["direction"].get()) else 0,
                    "offset_x": max(-120, min(120, int(fields["offset_x"].get()))),
                    "offset_y": max(-120, min(120, int(fields["offset_y"].get()))),
                    "strip_accents": bool(accents.get()),
                    "grid": bool(grid_switch.get()),
                    "mode": "gdi" if mode_var.get().startswith("Driver") else "raw",
                    "numeracion": next((key for key, text in NUMERACION_TEXTO.items() if text == num_var.get()),
                                       NUMERACION_NUEVA),
                    "aviso_numeracion": True,  # elegirla acá equivale a confirmar el aviso
                }
            except ValueError:
                messagebox.showerror(APP_NAME, "Revisá los valores: deben ser números.", parent=win)
                return
            self.config_data.update(new)
            self._apply_numbering(new["numeracion"])
            self.save_config()
            self.set_status("Ajustes guardados.", "ok")
            win.destroy()

        ctk.CTkButton(holder, text="Guardar", height=40, fg_color=ACCENT, hover_color="#27D3BB", text_color="#04201C",
                      font=self.fonts["btn"], command=save).pack(fill="x", padx=24, pady=10)
        win.after(150, win.lift)

    # -- herramientas: juego de prueba, etiquetas internas y lector de códigos --------------
    def print_test_set(self):
        """Una etiqueta de cada familia (22, 29, 25, 23 y 24), con la marca PRUEBA, para probar los lectores."""
        items = test_set_labels(effective_day())
        labels = [item["elementos"] for item in items]
        page = "grid" if self.config_data.get("grid") else "label"
        job = PrintJob(labels[0], 20, 30, len(labels), self.config_data, page, labels=labels)
        job.on_success = lambda: self.set_status(
            f"✔ Juego de prueba impreso ({len(labels)} etiquetas). Escaneá cada una en «Leer un código»: todas tienen "
            "que dar «código correcto».", "ok")
        log("CLICK juego de prueba: " + ", ".join(item["codigo"] for item in items))
        self.submit(job, f"juego de prueba ({len(labels)} etiquetas)")

    def open_internal_labels(self):
        """Etiquetas que no son de producto: insumo interno (25), Promo Pack (23) e ID de balde (24)."""
        kinds = ("Insumo (25)", "Promo Pack (23)", "ID de balde (24)")
        win = ctk.CTkToplevel(self)
        win.title("Etiquetas internas")
        win.geometry("880x640")
        win.configure(fg_color=BG)
        win.transient(self)
        left = ctk.CTkFrame(win, fg_color=PANEL, corner_radius=12, width=450)
        left.pack(side="left", fill="y", padx=(16, 8), pady=16)
        left.pack_propagate(False)
        right = ctk.CTkFrame(win, fg_color=PANEL, corner_radius=12)
        right.pack(side="left", fill="both", expand=True, padx=(8, 16), pady=16)
        ctk.CTkLabel(right, text="Vista previa", font=self.fonts["h2"], text_color=TEXT).pack(
            anchor="w", padx=18, pady=(14, 6))
        caption = ctk.CTkLabel(right, text="", font=self.fonts["small"], text_color=MUTED, wraplength=330,
                               justify="left", anchor="w")
        caption.pack(side="bottom", fill="x", padx=18, pady=(4, 14))
        canvas = tk.Canvas(right, bg=CARD, highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=14)

        ctk.CTkLabel(left, text="Etiquetas internas", font=self.fonts["h2"], text_color=TEXT).pack(
            anchor="w", padx=18, pady=(14, 6))
        kind_var = tk.StringVar(value=kinds[0])
        ctk.CTkSegmentedButton(left, values=list(kinds), variable=kind_var, height=34, selected_color=ACCENT_DARK,
                               selected_hover_color=ACCENT, unselected_color=CARD, font=self.fonts["small"],
                               command=lambda value: switch()).pack(fill="x", padx=16, pady=(0, 4))
        ctk.CTkButton(left, text="Imprimir", height=44, corner_radius=10, fg_color=ACCENT, hover_color="#27D3BB",
                      text_color="#04201C", font=self.fonts["btn"], command=lambda: do_print()).pack(
            side="bottom", fill="x", padx=16, pady=(4, 14))
        note = ctk.CTkLabel(left, text="", font=self.fonts["small"], text_color=WARN_COLOR, wraplength=400,
                            justify="left", anchor="w")
        note.pack(side="bottom", fill="x", padx=18)
        body = ctk.CTkFrame(left, fg_color="transparent")
        body.pack(side="top", fill="both", expand=True, padx=18)

        def heading(parent, text):
            ctk.CTkLabel(parent, text=text, font=self.fonts["small"], text_color=MUTED, anchor="w",
                         justify="left", wraplength=400).pack(fill="x", pady=(10, 0))

        def field(parent, value=""):
            box = ctk.CTkEntry(parent, height=34, fg_color=CARD, border_color=LINE, font=self.fonts["body"])
            box.insert(0, value)
            box.pack(fill="x")
            box.bind("<KeyRelease>", lambda event: redraw())
            return box

        def read_int(box, what, low, high):
            try:
                value = int(box.get().strip())
            except ValueError:
                raise ValueError(f"{what}: escribí un número entero.") from None
            if not low <= value <= high:
                raise ValueError(f"{what}: tiene que estar entre {low} y {high}.")
            return value

        def alive():
            try:
                return bool(win.winfo_exists())
            except tk.TclError:
                return False

        # -- insumo interno (25)
        f_supply = ctk.CTkFrame(body, fg_color="transparent")
        supplies = {}
        supply_var = tk.StringVar()
        heading(f_supply, "Insumo (los números son provisorios hasta que Glacé y e-Karú los acuerden)")
        supply_menu = ctk.CTkOptionMenu(f_supply, variable=supply_var, values=["…"], height=34, fg_color=CARD,
                                        button_color=ACCENT_DARK, button_hover_color=ACCENT, dropdown_fg_color=PANEL,
                                        font=self.fonts["small"], dynamic_resizing=False,
                                        command=lambda value: redraw())
        supply_menu.pack(fill="x")
        heading(f_supply, "Fecha de recepción (día/mes/año)")
        received = field(f_supply, fecha_texto(date.today()))
        heading(f_supply, "Cantidad de etiquetas")
        supply_qty = field(f_supply, "1")
        heading(f_supply, "Agregar un insumo nuevo")
        add_row = ctk.CTkFrame(f_supply, fg_color="transparent")
        add_row.pack(fill="x")
        new_supply = ctk.CTkEntry(add_row, height=34, fg_color=CARD, border_color=LINE, font=self.fonts["body"],
                                  placeholder_text="Nombre, por ejemplo: Cobertura de chocolate",
                                  placeholder_text_color=MUTED)
        new_supply.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(add_row, text="Agregar", width=84, height=34, fg_color=ACCENT_DARK, hover_color=ACCENT,
                      font=self.fonts["small"], command=lambda: add_new_supply()).pack(side="left", padx=(8, 0))

        def load_supply_menu(select=None):
            supplies.clear()
            for number, name in load_supplies():
                supplies[f"{number:04d} · {name}"] = (number, name)
            values = list(supplies)
            supply_menu.configure(values=values)
            supply_var.set(next((v for v in values if supplies[v][0] == select), values[0]))

        def add_new_supply():
            try:
                number, _ = add_supply(new_supply.get())
            except (ValueError, OSError) as error:
                note.configure(text=str(error))
                return
            new_supply.delete(0, "end")
            load_supply_menu(select=number)
            redraw()

        # -- Promo Pack (23) e ID de balde (24): numeración correlativa
        serial = {}
        for key, text in (("pack", "Cada Promo Pack (12 palitos) lleva un número propio y único. Las etiquetas salen "
                                   "numeradas una tras otra."),
                          ("balde", "Plan B para el balde: una etiqueta con ID propio, que se pega junto a la de la "
                                    "balanza. Las etiquetas salen numeradas una tras otra.")):
            frame = ctk.CTkFrame(body, fg_color="transparent")
            heading(frame, text)
            info = ctk.CTkLabel(frame, text="", font=self.fonts["body"], text_color=TEXT, anchor="w", justify="left")
            info.pack(fill="x", pady=(10, 0))
            heading(frame, "Empezar en el Nº")
            start = field(frame, "1")
            heading(frame, "Cantidad de etiquetas")
            quantity = field(frame, "1")
            serial[key] = {"frame": frame, "info": info, "start": start, "qty": quantity}

        def show_counters(reset_start=False):
            counters = load_counters()
            for key, parts in serial.items():
                last = counters[key]
                parts["info"].configure(text=f"Último número impreso: {last}   ·   Próximo: {last + 1}")
                if reset_start:
                    parts["start"].delete(0, "end")
                    parts["start"].insert(0, str(last + 1))

        def bump(key, last):
            """Avanza el contador (solo si el envío salió bien y nunca hacia atrás)."""
            counters = load_counters()
            if last > counters[key]:
                counters[key] = last
                try:
                    save_counters(counters)
                except OSError as error:
                    log(f"No se pudo guardar el contador: {error}")
                    self.set_status(f"▲ Se imprimió, pero no se pudo guardar el contador: {error}", "warn")
            if alive():
                show_counters(reset_start=True)
                redraw()

        frames = {kinds[0]: f_supply, kinds[1]: serial["pack"]["frame"], kinds[2]: serial["balde"]["frame"]}

        def plan(full=True):
            """Lo que está en pantalla, listo para imprimir. Lanza ValueError con el motivo si falta algo."""
            kind = kind_var.get()
            if kind == kinds[0]:
                if not supplies:
                    raise ValueError("Todavía no hay insumos: agregá uno.")
                number, name = supplies[supply_var.get()]
                day, today = parse_user_date(received.get()), date.today()
                if day is None:
                    raise ValueError("Fecha de recepción: escribí día/mes/año, por ejemplo 12/11/2026.")
                if day > today:
                    raise ValueError("La fecha de recepción no puede ser futura.")
                if (today - day).days > MAX_DIAS_ATRAS:
                    raise ValueError(f"Esa fecha es de hace más de {MAX_DIAS_ATRAS} días: revisá el año.")
                copies = read_int(supply_qty, "Cantidad", 1, 999)
                code = codigo_insumo(number, day)
                return {"preview": insumo_label_elements(name, code, day), "labels": None, "copies": copies,
                        "code": code, "text": f"insumo «{name}» × {copies}", "after": None,
                        "caption": f"{spaced_code(code)}\n25 + Nº de insumo {number:04d} + recepción "
                                   f"{fecha_texto(day)} + control"}
            key = "pack" if kind == kinds[1] else "balde"
            title, subtitle, maker = (("PROMO PACK", "12 PALITOS", codigo_pack) if key == "pack"
                                      else ("ID DE BALDE", "", codigo_id_balde))
            first = read_int(serial[key]["start"], "Nº inicial", 0, MAX_NUMERO_CORRELATIVO)
            copies = read_int(serial[key]["qty"], "Cantidad", 1, 999)
            last = first + copies - 1
            if last > MAX_NUMERO_CORRELATIVO:
                raise ValueError("Se acaban los números disponibles: bajá la cantidad o el Nº inicial.")
            numbers = range(first, last + 1) if full else [first]
            labels = [serial_label_elements(title, subtitle, maker(n)) for n in numbers]
            what = "Promo Pack" if key == "pack" else "ID de balde"
            caption = spaced_code(maker(first)) + (f"\n… hasta {spaced_code(maker(last))}" if copies > 1 else "")
            return {"preview": labels[0], "labels": labels if full else None, "copies": copies, "code": maker(first),
                    "text": f"{what} Nº {first}" + (f" a {last}" if copies > 1 else ""),
                    "after": lambda: bump(key, last), "caption": caption}

        def redraw(*_):
            if not alive():
                return
            try:
                chosen = plan(full=False)
            except ValueError as error:
                note.configure(text=str(error))
                canvas.delete("all")
                caption.configure(text="")
                return
            note.configure(text="")
            width, height = canvas.winfo_width(), canvas.winfo_height()
            if width >= 40 and height >= 40:
                draw_label(canvas, chosen["preview"], 160, 240, width, height)
            caption.configure(text=chosen["caption"])

        def switch():
            for frame in frames.values():
                frame.pack_forget()
            frames[kind_var.get()].pack(fill="both", expand=True)
            redraw()

        def do_print():
            try:
                chosen = plan(full=True)
            except ValueError as error:
                note.configure(text=str(error))
                return
            note.configure(text="")
            page = "grid" if self.config_data.get("grid") else "label"
            if chosen["labels"]:
                job = PrintJob(chosen["labels"][0], 20, 30, len(chosen["labels"]), self.config_data, page,
                               labels=chosen["labels"])
            else:
                job = PrintJob(chosen["preview"], 20, 30, chosen["copies"], self.config_data, page)
            job.on_success = chosen["after"]
            log(f"CLICK etiquetas internas: {chosen['text']} · código {chosen['code']}")
            self.submit(job, chosen["text"])

        canvas.bind("<Configure>", redraw)
        load_supply_menu()
        show_counters(reset_start=True)
        switch()
        win.after(150, win.lift)
        win.after(250, redraw)

    def open_reader(self):
        """Lee un código como lo hará e-Karú: las reglas de lectura del Informe final, en el mismo orden."""
        win = ctk.CTkToplevel(self)
        win.title("Leer un código")
        win.geometry("740x660")
        win.configure(fg_color=BG)
        win.transient(self)
        ctk.CTkLabel(win, text="Leer un código", font=self.fonts["title"], text_color=TEXT).pack(
            anchor="w", padx=20, pady=(16, 2))
        ctk.CTkLabel(win, text="Escaneá una etiqueta (o escribí el número) y apretá Enter. Se aplican las reglas de "
                               "lectura del Informe final, en el mismo orden que usará e-Karú.",
                     font=self.fonts["small"], text_color=MUTED, wraplength=680, justify="left", anchor="w").pack(
            fill="x", padx=20, pady=(0, 10))
        row = ctk.CTkFrame(win, fg_color="transparent")
        row.pack(fill="x", padx=20)
        var = tk.StringVar()
        entry = ctk.CTkEntry(row, textvariable=var, height=44, fg_color=CARD, border_color=LINE,
                             font=ctk.CTkFont(family=MONO_FONT, size=18), placeholder_text="Código de barras…",
                             placeholder_text_color=MUTED)
        entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="Leer", width=90, height=44, fg_color=ACCENT, hover_color="#27D3BB",
                      text_color="#04201C", font=self.fonts["btn"], command=lambda: read()).pack(side="left", padx=(8, 0))
        verdict = ctk.CTkFrame(win, fg_color=CARD, corner_radius=12)
        verdict.pack(fill="x", padx=20, pady=(12, 8))
        headline = ctk.CTkLabel(verdict, text="Esperando un código…", font=self.fonts["h2"], text_color=MUTED,
                                anchor="w", justify="left", wraplength=660)
        headline.pack(fill="x", padx=16, pady=(12, 2))
        details = ctk.CTkLabel(verdict, text="", font=self.fonts["body"], text_color=TEXT, anchor="w",
                               justify="left", wraplength=660)
        details.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkLabel(win, text="Lecturas de esta sesión", font=self.fonts["small"], text_color=MUTED).pack(
            anchor="w", padx=20)
        bottom = ctk.CTkFrame(win, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=20, pady=(0, 16))
        history = ctk.CTkTextbox(win, fg_color=PANEL, text_color=TEXT,
                                 font=ctk.CTkFont(family=MONO_FONT, size=12), wrap="none")
        history.pack(side="top", fill="both", expand=True, padx=20, pady=(2, 8))
        history.configure(state="disabled")
        styles = {"ok": (OK_COLOR, "✔", "CÓDIGO CORRECTO"), "viejo": (WARN_COLOR, "▲", "CÓDIGO VIEJO"),
                  "rechazado": (ERR_COLOR, "✖", "RECHAZADO"), "externo": (ACCENT, "●", "CÓDIGO EXTERNO")}

        def write_history(text=None):
            history.configure(state="normal")
            if text is None:
                history.delete("1.0", "end")
            else:
                history.insert("1.0", text + "\n")
            history.configure(state="disabled")

        def read(_event=None):
            if not var.get().strip():
                return
            result = interpretar_codigo(var.get(), insumos=dict(load_supplies()))
            color, mark, title = styles[result["estado"]]
            headline.configure(text=f"{mark} {title} · {result['tipo']}", text_color=color)
            lines = [spaced_code(result["codigo"]) + (f"      [{result['regla']}]" if result["regla"] else "")]
            lines += [f"{name}: {value}" for name, value in result["campos"]]
            if result["mensaje"]:
                lines += ["", result["mensaje"]]
            details.configure(text="\n".join(lines))
            summary = "; ".join(f"{name}: {value}" for name, value in result["campos"][:2])
            write_history(f"{datetime.now():%H:%M:%S}  {mark} {spaced_code(result['codigo']):<30} "
                          f"{result['tipo']}" + (f" · {summary}" if summary else ""))
            log(f"LECTOR {result['codigo']}: {result['estado']} · {result['tipo']} ({result['regla']})")
            var.set("")
            entry.focus_set()

        ctk.CTkButton(bottom, text="Tabla de transición (CSV)", height=34, fg_color=CARD, hover_color=LINE,
                      font=self.fonts["small"], command=lambda: self.export_transition(win)).pack(side="left")
        ctk.CTkButton(bottom, text="Limpiar lecturas", height=34, fg_color=CARD, hover_color=LINE,
                      font=self.fonts["small"], command=lambda: write_history(None)).pack(side="right")
        entry.bind("<Return>", read)
        entry.bind("<KP_Enter>", read)
        win.after(150, win.lift)
        win.after(300, entry.focus_set)

    # -- página editor ---------------------------------------------------------
    def _build_editor_page(self):
        page = ctk.CTkFrame(self.main, fg_color=BG, corner_radius=0)
        self.pages["editor"] = page
        page.grid_columnconfigure(1, weight=1)
        page.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(page, text="Editor de etiquetas", font=self.fonts["title"], text_color=TEXT).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(18, 10))
        tools = ctk.CTkFrame(page, fg_color=PANEL, corner_radius=12, width=250)
        tools.grid(row=1, column=0, sticky="ns", padx=(20, 10), pady=(0, 16))
        tools.grid_propagate(False)
        self.ed_format = tk.StringVar(value=next(iter(FORMATS)))
        ctk.CTkOptionMenu(tools, variable=self.ed_format, values=list(FORMATS), height=34, fg_color=CARD,
                          button_color=ACCENT_DARK, button_hover_color=ACCENT, dropdown_fg_color=PANEL,
                          font=self.fonts["small"], command=lambda v: self._ed_redraw()).pack(fill="x", padx=14,
                                                                                               pady=(16, 8))
        for label, cmd in (("Agregar texto", self._ed_add_text), ("Agregar código de barras", self._ed_add_barcode),
                           ("Agregar rectángulo", self._ed_add_box), ("Agregar línea", self._ed_add_line),
                           ("Editar seleccionado", self._ed_edit), ("Eliminar seleccionado", self._ed_delete),
                           ("Limpiar todo", self._ed_clear)):
            ctk.CTkButton(tools, text=label, height=34, anchor="w", fg_color=CARD, hover_color=LINE,
                          font=self.fonts["small"], command=cmd).pack(fill="x", padx=14, pady=3)
        self.ed_product = tk.StringVar(value=next(iter(PRODUCTS)))
        ctk.CTkLabel(tools, text="Partir de un producto", font=self.fonts["small"], text_color=MUTED).pack(
            anchor="w", padx=16, pady=(14, 2))
        ctk.CTkOptionMenu(tools, variable=self.ed_product, values=list(PRODUCTS), height=32, fg_color=CARD,
                          button_color=ACCENT_DARK, button_hover_color=ACCENT, dropdown_fg_color=PANEL,
                          font=self.fonts["small"], dynamic_resizing=False,
                          command=self._ed_load_product).pack(fill="x", padx=14)
        for label, cmd in (("Guardar diseño", self._ed_save), ("Abrir diseño", self._ed_load)):
            ctk.CTkButton(tools, text=label, height=34, anchor="w", fg_color=CARD, hover_color=LINE,
                          font=self.fonts["small"], command=cmd).pack(fill="x", padx=14, pady=(10, 0))
        ctk.CTkButton(tools, text="Imprimir diseño (1)", height=42, fg_color=ACCENT, hover_color="#27D3BB",
                      text_color="#04201C", font=self.fonts["btn"], command=self._ed_print).pack(
            fill="x", padx=14, pady=(14, 8))
        stage = ctk.CTkFrame(page, fg_color=PANEL, corner_radius=12)
        stage.grid(row=1, column=1, sticky="nsew", padx=(0, 20), pady=(0, 16))
        stage.grid_columnconfigure(0, weight=1)
        stage.grid_rowconfigure(0, weight=1)
        self.ed_canvas = tk.Canvas(stage, bg=CARD, highlightthickness=0)
        self.ed_canvas.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.ed_canvas.bind("<Configure>", lambda e: self._ed_redraw())
        self.ed_canvas.bind("<Button-1>", self._ed_press)
        self.ed_canvas.bind("<B1-Motion>", self._ed_drag)
        self.ed_canvas.bind("<Double-Button-1>", lambda e: self._ed_edit())
        self.ed_canvas.bind("<Delete>", lambda e: self._ed_delete())
        self._ed_geom = (1, 0, 0)

    def report_callback_exception(self, exc, val, tb):
        text = "".join(traceback.format_exception(exc, val, tb))
        log("ERROR en la interfaz: " + text)
        self._busy = False
        self.set_status(f"✖ Error interno: {val}  (ver ekaru_impresion.log)", "error")

    def _ed_elements(self):
        elements = []
        for item in self.editor_items:
            t = item["type"]
            if t == "text":
                elements.append(("text", item["x"], item["y"], item["font"], item["text"], item.get("bold", False)))
            elif t == "barcode":
                width, _ = label_size_dots(self.ed_format.get())
                bars, _, _ = barcode_elements(item["value"], item["x"], item["y"], item["h"], width - 16)
                elements.extend(bars)
            elif t == "box":
                elements.append(("box", item["x"], item["y"], item["w"], item["h"], item.get("t", 2)))
            elif t == "line":
                elements.append(("bar", item["x"], item["y"], item["w"], item.get("t", 2)))
        return elements

    def _ed_bbox(self, item):
        t = item["type"]
        if t == "text":
            fw, fh = FONT_DOTS[item["font"]]
            return item["x"], item["y"], item["x"] + fw * len(item["text"]), item["y"] + fh
        if t == "barcode":
            width, _ = label_size_dots(self.ed_format.get())
            bars, _, w = barcode_elements(item["value"], item["x"], item["y"], item["h"], width - 16)
            return item["x"] - w // 2, item["y"], item["x"] + w // 2, item["y"] + item["h"]
        if t == "line":
            return item["x"], item["y"], item["x"] + item["w"], item["y"] + item.get("t", 2)
        return item["x"], item["y"], item["x"] + item["w"], item["y"] + item["h"]

    def _ed_redraw(self):
        width, height = self.ed_canvas.winfo_width(), self.ed_canvas.winfo_height()
        if width < 40 or height < 40:
            return
        label_w, label_h = label_size_dots(self.ed_format.get())
        self._ed_geom = draw_label(self.ed_canvas, self._ed_elements(), label_w, label_h, width, height, pad=30)
        if self.editor_sel is not None and self.editor_sel < len(self.editor_items):
            scale, ox, oy = self._ed_geom
            x0, y0, x1, y1 = self._ed_bbox(self.editor_items[self.editor_sel])
            self.ed_canvas.create_rectangle(ox + x0 * scale - 2, oy + y0 * scale - 2, ox + x1 * scale + 2,
                                            oy + y1 * scale + 2, outline=ACCENT, dash=(4, 3), width=2)

    def _ed_press(self, event):
        self.ed_canvas.focus_set()
        scale, ox, oy = self._ed_geom
        px, py = (event.x - ox) / scale, (event.y - oy) / scale
        self.editor_sel = None
        for index in range(len(self.editor_items) - 1, -1, -1):
            x0, y0, x1, y1 = self._ed_bbox(self.editor_items[index])
            if x0 - 3 <= px <= x1 + 3 and y0 - 3 <= py <= y1 + 3:
                self.editor_sel = index
                self._ed_grab = (px, py, dict(self.editor_items[index]))
                break
        self._ed_redraw()

    def _ed_drag(self, event):
        if self.editor_sel is None:
            return
        scale, ox, oy = self._ed_geom
        px, py = (event.x - ox) / scale, (event.y - oy) / scale
        gx, gy, original = self._ed_grab
        item = self.editor_items[self.editor_sel]
        item["x"] = round((original["x"] + px - gx) / 4) * 4
        item["y"] = round((original["y"] + py - gy) / 4) * 4
        self._ed_redraw()

    def _ed_add(self, item):
        self.editor_items.append(item)
        self.editor_sel = len(self.editor_items) - 1
        self._ed_redraw()

    def _ed_add_text(self):
        text = simpledialog.askstring(APP_NAME, "Texto:", parent=self)
        if text:
            self._ed_add({"type": "text", "x": 12, "y": 12, "font": "2", "text": text, "bold": False})

    def _ed_add_barcode(self):
        value = simpledialog.askstring(APP_NAME, "Valor del código de barras (solo números o letras):", parent=self)
        if value and value.strip():
            w, h = label_size_dots(self.ed_format.get())
            self._ed_add({"type": "barcode", "x": w // 2, "y": h // 2, "h": 48, "value": value.strip()})

    def _ed_add_box(self):
        w, h = label_size_dots(self.ed_format.get())
        self._ed_add({"type": "box", "x": 4, "y": 4, "w": w - 8, "h": h - 8, "t": 2})

    def _ed_add_line(self):
        w, h = label_size_dots(self.ed_format.get())
        self._ed_add({"type": "line", "x": 8, "y": h // 2, "w": w - 16, "t": 2})

    def _ed_edit(self):
        if self.editor_sel is None:
            self.set_status("Primero seleccioná un elemento en la etiqueta.", "warn")
            return
        item = self.editor_items[self.editor_sel]
        if item["type"] == "text":
            text = simpledialog.askstring(APP_NAME, "Texto:", initialvalue=item["text"], parent=self)
            size = simpledialog.askinteger(APP_NAME, "Tamaño de letra (1 = chica … 5 = muy grande):",
                                           initialvalue=int(item["font"]), minvalue=1, maxvalue=5, parent=self)
            if text:
                item["text"] = text
            if size:
                item["font"] = str(size)
        elif item["type"] == "barcode":
            value = simpledialog.askstring(APP_NAME, "Valor:", initialvalue=item["value"], parent=self)
            height = simpledialog.askinteger(APP_NAME, "Alto de las barras (puntos):", initialvalue=item["h"],
                                             minvalue=16, maxvalue=400, parent=self)
            if value and value.strip():
                item["value"] = value.strip()
            if height:
                item["h"] = height
        else:
            key = "w"
            number = simpledialog.askinteger(APP_NAME, "Ancho (puntos):", initialvalue=item.get(key, 40),
                                             minvalue=2, maxvalue=900, parent=self)
            if number:
                item["w"] = number
            if item["type"] == "box":
                number = simpledialog.askinteger(APP_NAME, "Alto (puntos):", initialvalue=item["h"], minvalue=2,
                                                 maxvalue=900, parent=self)
                if number:
                    item["h"] = number
        self._ed_redraw()

    def _ed_delete(self):
        if self.editor_sel is not None and self.editor_sel < len(self.editor_items):
            del self.editor_items[self.editor_sel]
            self.editor_sel = None
            self._ed_redraw()

    def _ed_clear(self):
        if not self.editor_items or messagebox.askyesno(APP_NAME, "¿Borrar todos los elementos?", parent=self):
            self.editor_items = []
            self.editor_sel = None
            self._ed_redraw()

    def _ed_load_product(self, name):
        product = PRODUCTS[name]
        self.ed_format.set(product["format"])
        items = []
        bars = []
        for e in product_elements(name):
            if e[0] == "text":
                items.append({"type": "text", "x": e[1], "y": e[2], "font": e[3], "text": e[4], "bold": e[5]})
            elif e[0] == "bar" and e[4] > 4:
                bars.append(e)
            elif e[0] == "bar":
                items.append({"type": "line", "x": e[1], "y": e[2], "w": e[3], "t": e[4]})
        if bars:
            left = min(b[1] for b in bars)
            right = max(b[1] + b[3] for b in bars)
            items.append({"type": "barcode", "x": (left + right) // 2, "y": bars[0][2], "h": bars[0][4],
                          "value": print_code(name)})
        self.editor_items, self.editor_sel = items, None
        self._ed_redraw()

    def _ed_save(self):
        path = filedialog.asksaveasfilename(parent=self, initialdir=str(DESIGNS_DIR), defaultextension=".json",
                                            filetypes=[("Diseño E-Karú", "*.json")])
        if path:
            try:
                atomic_write_json(Path(path), {"format": self.ed_format.get(), "items": self.editor_items})
                self.set_status(f"Diseño guardado: {path}", "ok")
            except OSError as error:
                messagebox.showerror(APP_NAME, f"No se pudo guardar:\n{error}", parent=self)

    def _ed_load(self):
        path = filedialog.askopenfilename(parent=self, initialdir=str(DESIGNS_DIR),
                                          filetypes=[("Diseño E-Karú", "*.json")])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            if data["format"] not in FORMATS or not isinstance(data["items"], list):
                raise ValueError("formato desconocido")
            for item in data["items"]:
                if item["type"] not in ("text", "barcode", "box", "line"):
                    raise ValueError("elemento desconocido")
        except (OSError, ValueError, KeyError, TypeError) as error:
            messagebox.showerror(APP_NAME, f"No se pudo abrir el diseño:\n{error}", parent=self)
            return
        self.ed_format.set(data["format"])
        self.editor_items, self.editor_sel = data["items"], None
        self._ed_redraw()

    def _ed_print(self):
        if not self.editor_items:
            self.set_status("El diseño está vacío.", "warn")
            return
        w_mm, h_mm = physical_mm(self.ed_format.get())
        small = self.ed_format.get() == "Etiqueta pequeña (3 x 2 cm)"
        page = "grid" if small and self.config_data.get("grid") else "label"
        self.submit(PrintJob(self._ed_elements(), w_mm, h_mm, 1, self.config_data, page), "diseño personalizado")


def main():
    try:
        app = EkaruApp()
        app.mainloop()
    except Exception:  # noqa: BLE001 - se registra para poder diagnosticar con pythonw
        text = traceback.format_exc()
        try:
            ERROR_LOG.write_text(f"{datetime.now():%Y-%m-%d %H:%M:%S}\n{text}", encoding="utf-8")
        except OSError:
            pass
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(APP_NAME, f"La aplicación no pudo iniciar.\n\n{text[-900:]}\n\nDetalle en:\n{ERROR_LOG}")
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    main()
