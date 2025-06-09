#!/usr/bin/env python3
import sys, os
import argparse
from collections import defaultdict
from datetime import date, datetime
from src.reader import parse_calendar

# Mapea códigos a descripciones
LEGEND = {
    "SI": "Send information",
    "RI": "Review information and doubts (EY Local)",
    "SD": "Send draft (EY Local)",
    "AD": "Approve draft",
    "OS": "Official Submission Deadline",
    "OP": "Official Payment Deadline",
    "SP": "Official Submission and Payment (same deadline)",
    "HS": "Public Holiday Spain",
    "HL": "Local Holiday – Non-working day"
}

PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
EXCEL = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
OUT_DIR = os.path.join(PROYECTO_ROOT, "data", "outputs")

def parse_args():
    p = argparse.ArgumentParser(
        description="Genera informes diarios de impuestos a partir del Excel"
    )
    p.add_argument(
        "--date", "-d",
        help="Fecha (YYYY-MM-DD) para generar el informe. Default: hoy o próximas fechas.",
    )
    p.add_argument(
        "--company", "-c",
        help="Empresa para filtrar (ej. ENDESA). Si no se indica, incluye todas.",
    )
    return p.parse_args()

def main():
    args = parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1) Extraer todos los registros
    regs = parse_calendar(EXCEL)

    # 2) Filtrar fechas a partir de hoy
    hoy = date.today()
    regs = [r for r in regs if r["fecha"] >= (date.fromisoformat(args.date) if args.date else hoy)]

    # 3) Filtrar empresa si se pide
    if args.company:
        regs = [r for r in regs if r["empresa"].lower() == args.company.lower()]

    if not regs:
        print("No hay registros para esos filtros.")
        return

    # 4) Agrupar por fecha → empresa → país → impuesto
    by_fecha = defaultdict(list)
    for r in regs:
        by_fecha[r["fecha"]].append(r)

    for fecha in sorted(by_fecha):
        fname = fecha.isoformat()
        if args.company:
            fname += f"_{args.company}"
        out_path = os.path.join(OUT_DIR, f"{fname}.txt")

        with open(out_path, "w", encoding="utf-8") as f:
            header = f"Informe de impuestos para {fecha.isoformat()}"
            if args.company:
                header += f" (Empresa: {args.company})"
            f.write(header + "\n\n")

            # Organizar en diccionario empresa → país → lista de (impuesto, estado)
            emp_map = defaultdict(lambda: defaultdict(list))
            for r in by_fecha[fecha]:
                emp_map[r["empresa"]][r["pais"]].append((r["impuesto"], r["estado"]))

            # Escribir por empresa
            for emp, paises in emp_map.items():
                f.write(f"Empresa: {emp}\n")
                for pais, items in paises.items():
                    f.write(f"  País: {pais}\n")
                    for impuesto, estado in items:
                        desc = LEGEND.get(estado, "")
                        f.write(f"    • {impuesto}: {estado}" + (f" — {desc}" if desc else "") + "\n")
                f.write("\n")
        print(f"Generado: {out_path}")

if __name__ == "__main__":
    main()
