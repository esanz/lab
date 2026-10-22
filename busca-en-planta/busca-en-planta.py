#!/usr/bin/env python3
"""
busca-en-planta.py
==================

Busca una o varias direcciones IP en la hoja de cálculo de "Planta de Datos"
y muestra, para cada IP, el Punto de Servicio (PS) al que pertenece:

    <IP> --> <Código PS>, <Nombre del PS>, <Dirección>, <Municipio>,
             <Provincia>, <Código Postal>, <Coubicado>

Además del listado por pantalla genera un fichero CSV con el resultado.

Cómo se busca
-------------
1. Coincidencia exacta con cualquier columna cuyo nombre empiece por "IP"
   (IP Física, IP Virtual, IP Loopback, IP Estática, ... y sus "Respaldo").
2. Si no hay coincidencia exacta, se busca la IP dentro de los rangos de las
   columnas "Rango Direccionamiento ..." (pueden venir varios separados por
   ";"). Se queda con el rango más específico (prefijo más largo).
Se ignoran los valores 0.0.0.0 y 0.0.0.0/0, que la hoja usa como "vacío".

Uso:
    python3 busca-en-planta.py -p Planta_Datos.xlsx 10.13.2.2 10.238.80.77
    python3 busca-en-planta.py -p Planta_Datos.xlsx -f ips.txt
    python3 busca-en-planta.py -p Planta_Datos.xlsx ips.txt -o salida.csv
    cat ips.txt | python3 busca-en-planta.py -p Planta_Datos.xlsx -

Los argumentos posicionales pueden ser IPs o ficheros; de un fichero se
extraen todas las IPs que aparezcan (una por línea, separadas por comas,
con comentarios "#", etc.).

Requiere: openpyxl (pip install openpyxl)
"""

import argparse
import csv
import ipaddress
import os
import re
import sys
from datetime import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("Falta el módulo openpyxl. Instálalo con: pip install openpyxl")


HOJA_POR_DEFECTO = "Planta de datos"

# (clave interna, cabecera en la hoja, cabecera en el CSV de salida)
CAMPOS = [
    ("codigo", "Código de Punto de Servicio", "Código de Punto de Servicio"),
    ("nombre", "Nombre del PS", "Nombre del PS"),
    ("direccion", "Dirección", "Dirección"),
    ("municipio", "Municipio", "Municipio"),
    ("provincia", "Provincia", "Provincia"),
    ("cp", "Código Postal", "Código Postal"),
    ("coubicado", "Coubicado", "Coubicado"),
]

RE_IP = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
VACIOS = {ipaddress.ip_address("0.0.0.0")}


def texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


# ---------------------------------------------------------------------------
# Lectura de la planta
# ---------------------------------------------------------------------------

def cargar_planta(ruta, nombre_hoja=None):
    """Devuelve (registros, indice_ip, rangos).

    registros: lista de dict con los CAMPOS de cada fila.
    indice_ip: {ip_address: [(n_registro, columna), ...]}
    rangos:    [(ip_network, n_registro, columna), ...]
    """
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)

    if nombre_hoja:
        if nombre_hoja not in wb.sheetnames:
            sys.exit(f"No existe la hoja '{nombre_hoja}'. Hojas: {', '.join(wb.sheetnames)}")
        ws = wb[nombre_hoja]
    elif HOJA_POR_DEFECTO in wb.sheetnames:
        ws = wb[HOJA_POR_DEFECTO]
    else:
        ws = wb.worksheets[0]

    filas = ws.iter_rows(values_only=True)
    cabecera = [texto(c) for c in next(filas, [])]

    # Primera aparición de cada cabecera (hay alguna repetida en la hoja)
    pos = {}
    for i, c in enumerate(cabecera):
        pos.setdefault(c, i)

    faltan = [cab for _, cab, _ in CAMPOS if cab not in pos]
    if faltan:
        sys.exit(f"En la hoja '{ws.title}' faltan las columnas: {', '.join(faltan)}")

    cols_ip = [(i, c) for i, c in enumerate(cabecera) if c.upper().startswith("IP ")]
    cols_rango = [(i, c) for i, c in enumerate(cabecera)
                  if c.lower().startswith("rango direccionamiento")]

    registros, indice_ip, rangos = [], {}, []
    for fila in filas:
        if not fila or all(v is None for v in fila):
            continue
        fila = list(fila) + [None] * (len(cabecera) - len(fila))
        reg = {clave: texto(fila[pos[cab]]) for clave, cab, _ in CAMPOS}
        if not any(reg.values()):
            continue
        n = len(registros)
        registros.append(reg)

        for i, col in cols_ip:
            for s in RE_IP.findall(texto(fila[i])):
                try:
                    ip = ipaddress.ip_address(s)
                except ValueError:
                    continue
                if ip not in VACIOS:
                    indice_ip.setdefault(ip, []).append((n, col))

        for i, col in cols_rango:
            for s in re.split(r"[;,\s]+", texto(fila[i])):
                if not s:
                    continue
                try:
                    red = ipaddress.ip_network(s, strict=False)
                except ValueError:
                    continue
                if red.prefixlen > 0:
                    rangos.append((red, n, col))

    wb.close()
    return registros, indice_ip, rangos


# ---------------------------------------------------------------------------
# Búsqueda
# ---------------------------------------------------------------------------

def buscar(ip, registros, indice_ip, rangos):
    """Devuelve lista de (registro, tipo_coincidencia, detalle) sin duplicados."""
    encontrados = []
    if ip in indice_ip:
        tipo = "exacta"
        for n, col in indice_ip[ip]:
            encontrados.append((n, col))
    else:
        tipo = "rango"
        candidatos = [(red, n, col) for red, n, col in rangos if ip in red]
        if candidatos:
            mejor = max(red.prefixlen for red, _, _ in candidatos)
            encontrados = [(n, f"{col} {red}") for red, n, col in candidatos
                           if red.prefixlen == mejor]

    # Agrupa por PS (un mismo PS puede salir en varias columnas o filas)
    resultado, vistos = [], {}
    for n, detalle in encontrados:
        reg = registros[n]
        clave = tuple(reg[c] for c, _, _ in CAMPOS)
        if clave in vistos:
            if detalle not in vistos[clave][2]:
                vistos[clave][2].append(detalle)
            continue
        item = (reg, tipo, [detalle])
        vistos[clave] = item
        resultado.append(item)
    return resultado


# ---------------------------------------------------------------------------
# Entrada de IPs
# ---------------------------------------------------------------------------

def ips_de_texto(contenido):
    ips = []
    for linea in contenido.splitlines():
        linea = linea.split("#", 1)[0]
        ips.extend(RE_IP.findall(linea))
    return ips


def recoger_ips(args):
    brutas = []
    for a in args.ips:
        if a == "-":
            brutas.extend(ips_de_texto(sys.stdin.read()))
        elif os.path.isfile(a):
            with open(a, encoding="utf-8-sig", errors="replace") as f:
                brutas.extend(ips_de_texto(f.read()))
        else:
            brutas.append(a)
    for fichero in args.fichero or []:
        with open(fichero, encoding="utf-8-sig", errors="replace") as f:
            brutas.extend(ips_de_texto(f.read()))

    ips, vistas = [], set()
    for s in brutas:
        s = s.strip()
        try:
            ip = ipaddress.ip_address(s)
        except ValueError:
            print(f"Aviso: '{s}' no es una dirección IP válida; se ignora.", file=sys.stderr)
            continue
        if ip not in vistas:
            vistas.add(ip)
            ips.append(ip)
    return ips


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        prog="busca-en-planta",
        description="Busca direcciones IP en la Planta de Datos y muestra el Punto de Servicio.")
    ap.add_argument("-p", "--planta", required=True,
                    help="hoja de cálculo de la Planta de Datos (.xlsx)")
    ap.add_argument("ips", nargs="*",
                    help="direcciones IP, ficheros con IPs, o '-' para leer de la entrada estándar")
    ap.add_argument("-f", "--fichero", action="append",
                    help="fichero con direcciones IP (se puede repetir)")
    ap.add_argument("-H", "--hoja", help=f"nombre de la hoja (por defecto '{HOJA_POR_DEFECTO}')")
    ap.add_argument("-o", "--csv",
                    help="fichero CSV de salida (por defecto busca-en-planta_AAAAMMDD-HHMMSS.csv)")
    ap.add_argument("--sin-csv", action="store_true", help="no generar el fichero CSV")
    ap.add_argument("-v", "--detalle", action="store_true",
                    help="indica en pantalla en qué columna se encontró la IP")
    args = ap.parse_args()

    ips = recoger_ips(args)
    if not ips:
        ap.error("no se ha indicado ninguna dirección IP válida")

    registros, indice_ip, rangos = cargar_planta(args.planta, args.hoja)

    filas_csv = []
    for ip in ips:
        res = buscar(ip, registros, indice_ip, rangos)
        if not res:
            print(f"{ip} --> NO ENCONTRADA")
            filas_csv.append([str(ip)] + [""] * len(CAMPOS) + ["no encontrada", ""])
            continue
        for reg, tipo, detalles in res:
            valores = [reg[c] for c, _, _ in CAMPOS]
            linea = f"{ip} --> {', '.join(valores)}"
            if args.detalle:
                linea += f"  [{tipo}: {'; '.join(detalles)}]"
            print(linea)
            filas_csv.append([str(ip)] + valores + [tipo, "; ".join(detalles)])

    if not args.sin_csv:
        ruta = args.csv or datetime.now().strftime("busca-en-planta_%Y%m%d-%H%M%S.csv")
        # utf-8-sig y ';' para que Excel en español lo abra directamente
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Dirección IP"] + [c for _, _, c in CAMPOS]
                       + ["Coincidencia", "Columna"])
            w.writerows(filas_csv)
        print(f"\nCSV generado: {ruta}", file=sys.stderr)


if __name__ == "__main__":
    main()
