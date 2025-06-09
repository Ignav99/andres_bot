#!/usr/bin/env python3
import sys
import os
import argparse
from collections import defaultdict
from src.reader import parse_calendar
from datetime import date

# Aseguramos que src/ esté en el PYTHONPATH
PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, PROYECTO_ROOT)

EXCEL = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
OUT_DIR = os.path.join(PROYECTO_ROOT, "data", "outputs")

def parse_args():
    p = argparse.ArgumentParser(
        description="Genera informes diarios de impuestos a partir del Excel"
    )
    p.add_argument(
        "--date", "-d",
        help="Fecha (YYYY-MM-DD) para la que generar el informe. Si no se indica, genera para todas.",
    )
    p.add_argument(
        "--company", "-c",
        help="Empresa para filtrar (ej. ENDESA). Si no se indica, incluye todas.",
    )
    return p.parse_args()

def main():
    args = parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1) Extraer todos los registros del Excel
    registros = parse_calendar(EXCEL)

    # 2) Filtrar si se pide una fecha concreta
    if args.date:
        try:
            filtro_fecha = date.fromisoformat(args.date)
            registros = [r for r in registros if r["fecha"] == filtro_fecha]
        except ValueError:
            print(f"Fecha inválida: {args.date}")
            sys.exit(1)

    # 3) Filtrar si se pide una empresa concreta
    if args.company:
        registros = [r for r in registros if r["empresa"].lower() == args.company.lower()]

    if not registros:
        print("No se encontraron registros con esos filtros.")
        return

    # 4) Agrupar por fecha
    por_fecha = defaultdict(list)
    for r in registros:
        por_fecha[r["fecha"]].append(r)

    # 5) Generar reportes
    for fecha, items in sorted(por_fecha.items()):
        fname = f"{fecha.isoformat()}"
        if args.company:
            fname += f"_{args.company}"
        filename = os.path.join(OUT_DIR, f"{fname}.txt")
        with open(filename, "w", encoding="utf-8") as f:
            header = f"Impuestos para {fecha.isoformat()}"
            if args.company:
                header += f" (Empresa: {args.company})"
            f.write(header + "\n\n")
            for it in items:
                line = f"- {it['empresa']} | {it['pais']} | {it['impuesto']} | {it['estado']}"
                f.write(line + "\n")
        print(f"Generado: {filename}")

if __name__ == "__main__":
    main()
