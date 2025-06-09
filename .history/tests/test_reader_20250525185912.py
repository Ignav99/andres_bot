#!/usr/bin/env python3
from src.reader import parse_calendar

def main():
    path = "tax_calendar_25.xlsm"  # ajusta si está en subcarpeta data/, p.ej. "data/tax_calendar_25.xlsm"
    registros = parse_calendar(path)
    print(f"Total registros encontrados: {len(registros)}\n")

    # Imprime los primeros 10 registros para inspección
    for i, reg in enumerate(registros[:10], start=1):
        print(f"{i:02d} | {reg['fecha'].date()} | {reg['empresa']} | {reg['pais']} | "
              f"{reg['impuesto']} | {reg['estado']}")
    
if __name__ == "__main__":
    main()
