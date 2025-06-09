#!/usr/bin/env python3
import sys, os
import argparse
from datetime import date, timedelta
from collections import defaultdict
from src.reader import parse_calendar

PROYECTO_ROOT = os.path.dirname(os.path.dirname(__file__))
EXCEL = os.path.join(PROYECTO_ROOT, "tax_calendar_25.xlsm")
OUT_DIR = os.path.join(PROYECTO_ROOT, "data", "outputs")

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
        description="Genera informes de impuestos por fecha, rango o todas las fechas"
    )
    grp = p.add_mutually_exclusive_group(required=True)
    grp.add_argument("-d", "--date",
                     help="Fecha única (YYYY-MM-DD).")
    grp.add_argument("--all-dates", action="store_true",
                     help="Informe para todas las fechas encontradas.")
    grp.add_argument("--range", nargs=2, metavar=("START","END"),
                     help="Rango de fechas (YYYY-MM-DD YYYY-MM-DD).")
    p.add_argument("-c", "--company",
                   help="Filtrar por empresa (ej. ALTADIA).")
    return p.parse_args()

def daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)

def write_report(regs, dt: date, company=None):
    fn = dt.isoformat()
    if company: fn += f"_{company}"
    path = os.path.join(OUT_DIR, f"{fn}.txt")

    grouped = defaultdict(lambda: defaultdict(list))
    for r in regs:
        grouped[r["empresa"]][r["pais"]].append((r["impuesto"], r["estado"]))

    with open(path, "w", encoding="utf-8") as f:
        header = f"Informe de impuestos para {dt.isoformat()}"
        if company: header += f" (Empresa: {company})"
        f.write(header+"\n\n")
        for emp, paises in grouped.items():
            f.write(f"Empresa: {emp}\n")
            for pais, items in paises.items():
                f.write(f"  País: {pais}\n")
                for imp, est in items:
                    desc = LEGEND.get(est,"")
                    line = f"    • {imp}: {est}"
                    if desc: line += f" — {desc}"
                    f.write(line+"\n")
            f.write("\n")
    print(f"Generado: {path}")

def main():
    args = parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    # Montamos la lista de fechas a procesar
    dates = []
    if args.date:
        try:
            dates = [date.fromisoformat(args.date)]
        except:
            print("Fecha inválida:", args.date); sys.exit(1)
    elif args.range:
        try:
            start = date.fromisoformat(args.range[0])
            end   = date.fromisoformat(args.range[1])
            if start > end:
                print("El inicio debe ser ≤ fin."); sys.exit(1)
            dates = list(daterange(start,end))
        except:
            print("Formato de rango inválido."); sys.exit(1)
    else:  # --all-dates
        all_regs = parse_calendar(EXCEL)
        if args.company:
            all_regs = [r for r in all_regs if r["empresa"].lower()==args.company.lower()]
        dates = sorted({r["fecha"] for r in all_regs})

    if not dates:
        print("No hay fechas para procesar con esos filtros.")
        return

    # Para cada fecha, extraemos e informamos
    for dt in dates:
        regs = parse_calendar(EXCEL, target_date=dt)
        if args.company:
            regs = [r for r in regs if r["empresa"].lower()==args.company.lower()]
        if regs:
            write_report(regs, dt, args.company)
        else:
            print(f"No hay registros para {dt.isoformat()}"+(f" y empresa {args.company}" if args.company else ""))

if __name__ == "__main__":
    main()
