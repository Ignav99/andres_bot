#!/usr/bin/env python3
import sys, os
from collections import defaultdict
from src.reader import parse_calendar

PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
EXCEL = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
OUT_DIR = os.path.join(PROYECTO_ROOT, "data", "outputs")

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    regs = parse_calendar(EXCEL)
    # Agrupar por fecha
    por_fecha = defaultdict(list)
    for r in regs:
        por_fecha[r["fecha"]].append(r)

    # Para cada fecha, genera un .txt
    for fecha, items in por_fecha.items():
        filename = os.path.join(OUT_DIR, f"{fecha.isoformat()}.txt")
        with open(filename, "w", encoding="utf-8") as f:
            f.write(f"Impuestos para {fecha.isoformat()}\n\n")
            for it in items:
                f.write(f"- {it['empresa']} | {it['pais']} | {it['impuesto']} | {it['estado']}\n")
        print(f"  Generado: {filename}")

if __name__ == "__main__":
    main()
