#!/usr/bin/env python3
import sys, os

# Importar parse_calendar
PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, PROYECTO_ROOT)

from src.reader import parse_calendar

def main():
    path = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
    regs = parse_calendar(path)
    print(f"Total registros encontrados: {len(regs)}\n")
    for i, r in enumerate(regs[:10], start=1):
        print(f"{i:02d} | {r['fecha'].isoformat()} | {r['empresa']} | "
              f"{r['pais']} | {r['impuesto']} | {r['estado']}")

if __name__ == "__main__":
    main()
