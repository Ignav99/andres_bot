#!/usr/bin/env python3
import sys, os
import argparse
from collections import defaultdict
from datetime import date
from src.reader import parse_calendar

# Leyenda de códigos
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
    p = argparse.ArgumentParser(description="Genera informe de un solo día de impuestos")
    p.add_argument("-d", "--date", required=True,
                   help="Fecha (YYYY-MM-DD) para generar el informe (obligatorio).")
    p.add_argument("-c", "--company",
                   help="Empresa a filtrar (ej. X-ELIO). Si no se indica, todas.")
    return p.parse_args()

def main():
    args = parse_args()
    # Convertir argumento date a objeto date
    try:
        target_date = date.fromisoformat(args.date)
    except ValueError:
        print(f"Formato de fecha inválido: {args.date}")
        sys.exit(1)

    # Crear carpeta de salida
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1) Leer todos los registros
    registros = parse_calendar(EXCEL, target_date)

    # 2) Filtrar solo la fecha seleccionada
    registros = [r for r in registros if r["fecha"] == target_date]

    # 3) Filtrar empresa si se pidió
    if args.company:
        registros = [r for r in registros if r["empresa"].lower() == args.company.lower()]

    if not registros:
        print(f"No hay registros para {args.date}" + (f" y empresa {args.company}" if args.company else ""))
        return

    # 4) Ignorar cualquier registro cuyo país sea “Legend” o “Leyend”
    registros = [r for r in registros if r["pais"].lower() not in ("legend", "leyend")]

    # 5) Agrupar por empresa → país → lista de (impuesto, estado)
    emp_map = defaultdict(lambda: defaultdict(list))
    for r in registros:
        emp_map[r["empresa"]][r["pais"]].append((r["impuesto"], r["estado"]))

    # 6) Generar un único fichero .txt
    fname = target_date.isoformat()
    if args.company:
        fname += f"_{args.company}"
    out_path = os.path.join(OUT_DIR, f"{fname}.txt")

    with open(out_path, "w", encoding="utf-8") as f:
        header = f"Informe de impuestos para {target_date.isoformat()}"
        if args.company:
            header += f" (Empresa: {args.company})"
        f.write(header + "\n\n")

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
