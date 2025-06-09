#!/usr/bin/env python3
import sys, os
import argparse
from datetime import date
from collections import defaultdict
from src.reader import parse_calendar

PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
EXCEL = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
OUT_DIR = os.path.join(PROYECTO_ROOT, "data", "outputs")

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

def parse_args():
    p = argparse.ArgumentParser(
        description="Genera informes de impuestos por fecha y empresa"
    )
    grp = p.add_mutually_exclusive_group(required=True)
    grp.add_argument(
        "-d", "--date",
        help="Fecha (YYYY-MM-DD) para generar un único informe."
    )
    grp.add_argument(
        "--all-dates", action="store_true",
        help="Genera un informe por cada fecha marcada en el Excel."
    )
    p.add_argument(
        "-c", "--company",
        help="Empresa a filtrar (ej. X-ELIO). Si no se indica, se incluyen todas."
    )
    return p.parse_args()

def write_report_for(regs, target_date, company_filter=None):
    """
    regs: lista de registros ya filtrados por fecha (y empresa si toca)
    target_date: date
    """
    fname = target_date.isoformat()
    if company_filter:
        fname += f"_{company_filter}"
    path = os.path.join(OUT_DIR, f"{fname}.txt")

    # Agrupar por empresa → país
    emp_map = defaultdict(lambda: defaultdict(list))
    for r in regs:
        emp_map[r["empresa"]][r["pais"]].append((r["impuesto"], r["estado"]))

    with open(path, "w", encoding="utf-8") as f:
        header = f"Informe de impuestos para {target_date.isoformat()}"
        if company_filter:
            header += f" (Empresa: {company_filter})"
        f.write(header + "\n\n")

        for emp, paises in emp_map.items():
            f.write(f"Empresa: {emp}\n")
            for pais, items in paises.items():
                f.write(f"  País: {pais}\n")
                for impuesto, estado in items:
                    desc = LEGEND.get(estado, "")
                    line = f"    • {impuesto}: {estado}"
                    if desc:
                        line += f" — {desc}"
                    f.write(line + "\n")
            f.write("\n")

    print(f"Generado: {path}")

def main():
    args = parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    # Si --all-dates, primero extraemos todo y calculamos fechas únicas
    if args.all_dates:
        all_regs = parse_calendar(EXCEL)
        # Filtrar empresa si se pide
        if args.company:
            all_regs = [r for r in all_regs if r["empresa"].lower() == args.company.lower()]

        fechas = sorted({r["fecha"] for r in all_regs})
        if not fechas:
            print("No se encontraron registros para los filtros dados.")
            return

        for dt in fechas:
            regs_dt = [r for r in all_regs if r["fecha"] == dt]
            write_report_for(regs_dt, dt, args.company)

    else:
        # Fecha única
        try:
            dt = date.fromisoformat(args.date)
        except ValueError:
            print(f"Fecha inválida: {args.date}")
            sys.exit(1)

        regs = parse_calendar(EXCEL, target_date=dt)
        # Filtrar empresa si se pide
        if args.company:
            regs = [r for r in regs if r["empresa"].lower() == args.company.lower()]

        if not regs:
            print(f"No hay registros para {dt.isoformat()}" + (f" y empresa {args.company}" if args.company else ""))
            return

        write_report_for(regs, dt, args.company)

if __name__ == "__main__":
    main()
