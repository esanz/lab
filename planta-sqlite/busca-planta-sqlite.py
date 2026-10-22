#!/usr/bin/env python3
"""
busca-planta-sqlite.py
======================

Igual que busca-en-planta.py, pero consultando la base de datos SQLite
generada por carga-planta-sqlite.py en lugar de la hoja de cálculo.

Para cada IP muestra el/los Punto(s) de Servicio (PS) cuya red de la columna
"Rango Direccionamiento Datos" (formato CIDR) contiene la IP:

    <IP> --> <Código PS>, <Nombre del PS>, <Dirección>, <Municipio>,
             <Provincia>, <Código Postal>, <Coubicado>

Si encajan varias filas se listan todas, una por línea. Además del listado
por pantalla genera un fichero CSV con el resultado.

Cómo se busca
-------------
Para cada longitud de prefijo presente en la tabla "rangos" se calcula la red
de la IP con ese prefijo y se busca por igualdad en el índice
(prefijo, inicio). Son como mucho 32 búsquedas por índice por IP, sea cual
sea el tamaño de la planta.

Uso:
    python3 busca-planta-sqlite.py 10.13.2.2 10.238.80.77
    python3 busca-planta-sqlite.py -d planta.sqlite -f ips.txt
    python3 busca-planta-sqlite.py ips.txt -o salida.csv
    cat ips.txt | python3 busca-planta-sqlite.py -

Los argumentos posicionales pueden ser IPs o ficheros; de un fichero se
extraen todas las IPs que aparezcan (una por línea, separadas por comas,
con comentarios "#", etc.).

No necesita módulos externos.
"""

import argparse
import csv
import ipaddress
import os
import re
import sqlite3
import sys
from datetime import datetime


BD_POR_DEFECTO = "planta.sqlite"
COLUMNA_RANGO = "Rango Direccionamiento Datos"

# Columnas de la tabla planta que se muestran (mismo nombre que en la hoja)
CAMPOS = [
    "Código de Punto de Servicio",
    "Nombre del PS",
    "Dirección",
    "Municipio",
    "Provincia",
    "Código Postal",
    "Coubicado",
]

RE_IP = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")


def q(nombre):
    """Identificador SQL entre comillas dobles."""
    return '"' + nombre.replace('"', '""') + '"'


# ---------------------------------------------------------------------------
# Búsqueda
# ---------------------------------------------------------------------------

class Planta:
    def __init__(self, ruta):
        if not os.path.isfile(ruta):
            sys.exit(f"No existe la base de datos '{ruta}'. "
                     "Genérala antes con carga-planta-sqlite.py")
        self.con = sqlite3.connect(f"file:{ruta}?mode=ro", uri=True)
        try:
            self.prefijos = [p for (p,) in self.con.execute(
                "SELECT DISTINCT prefijo FROM rangos ORDER BY prefijo DESC")]
            columnas = {r[1] for r in self.con.execute("PRAGMA table_info(planta)")}
        except sqlite3.DatabaseError as e:
            sys.exit(f"'{ruta}' no parece una base de datos de carga-planta-sqlite.py ({e})")
        faltan = [c for c in CAMPOS if c not in columnas]
        if faltan:
            sys.exit(f"En la tabla planta faltan las columnas: {', '.join(faltan)}")
        self.sql = (f"SELECT p.fila, r.cidr, {', '.join('p.' + q(c) for c in CAMPOS)} "
                    "FROM rangos r JOIN planta p ON p.fila = r.fila "
                    "WHERE r.prefijo = ? AND r.inicio = ?")

    def buscar(self, ip):
        """Devuelve [(fila, [cidr, ...], [valores CAMPOS]), ...] ordenado por fila."""
        n = int(ip)
        por_fila = {}
        for p in self.prefijos:
            inicio = n & (((1 << p) - 1) << (32 - p))
            for fila, cidr, *valores in self.con.execute(self.sql, (p, inicio)):
                por_fila.setdefault(fila, ([], valores))[0].append(cidr)
        return [(fila, cidrs, valores) for fila, (cidrs, valores) in sorted(por_fila.items())]

    def info_carga(self):
        try:
            return self.con.execute("SELECT fichero, hoja, fecha FROM carga").fetchone()
        except sqlite3.DatabaseError:
            return None


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
            ip = ipaddress.IPv4Address(s)
        except ValueError:
            print(f"Aviso: '{s}' no es una dirección IPv4 válida; se ignora.", file=sys.stderr)
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
        prog="busca-planta-sqlite",
        description="Busca direcciones IP en la Planta de Datos (SQLite) "
                    "y muestra el Punto de Servicio.")
    ap.add_argument("ips", nargs="*",
                    help="direcciones IP, ficheros con IPs, o '-' para leer de la entrada estándar")
    ap.add_argument("-d", "--bd", default=BD_POR_DEFECTO,
                    help=f"base de datos SQLite (por defecto {BD_POR_DEFECTO})")
    ap.add_argument("-f", "--fichero", action="append",
                    help="fichero con direcciones IP (se puede repetir)")
    ap.add_argument("-o", "--csv",
                    help="fichero CSV de salida (por defecto busca-en-planta_AAAAMMDD-HHMMSS.csv)")
    ap.add_argument("--sin-csv", action="store_true", help="no generar el fichero CSV")
    ap.add_argument("-v", "--detalle", action="store_true",
                    help="muestra en pantalla el rango CIDR y la fila de la hoja que encajan")
    args = ap.parse_args()

    ips = recoger_ips(args)
    if not ips:
        ap.error("no se ha indicado ninguna dirección IP válida")

    planta = Planta(args.bd)
    if args.detalle and (info := planta.info_carga()):
        print(f"[Planta: {info[0]}, hoja '{info[1]}', cargada el {info[2]}]", file=sys.stderr)

    filas_csv = []
    for ip in ips:
        res = planta.buscar(ip)
        if not res:
            print(f"{ip} --> NO ENCONTRADA")
            filas_csv.append([str(ip)] + [""] * len(CAMPOS) + ["", ""])
            continue
        for fila, cidrs, valores in res:
            valores = [v or "" for v in valores]
            rangos = "; ".join(cidrs)
            linea = f"{ip} --> {', '.join(valores)}"
            if args.detalle:
                linea += f"  [{rangos}, fila {fila}]"
            print(linea)
            filas_csv.append([str(ip)] + valores + [rangos, fila])

    if not args.sin_csv:
        ruta = args.csv or datetime.now().strftime("busca-en-planta_%Y%m%d-%H%M%S.csv")
        # utf-8-sig y ';' para que Excel en español lo abra directamente
        with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Dirección IP"] + CAMPOS + [COLUMNA_RANGO, "Fila en la hoja"])
            w.writerows(filas_csv)
        print(f"\nCSV generado: {ruta}", file=sys.stderr)


if __name__ == "__main__":
    main()
