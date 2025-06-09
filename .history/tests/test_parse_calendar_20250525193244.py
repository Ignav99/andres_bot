import os
import pytest
from src.reader import parse_calendar

# Usaremos un Excel de prueba más pequeño, ponlo en data/mini_prueba.xlsx
FIXTURE = os.path.join(os.path.dirname(__file__), "..", "data", "mini_prueba.xlsx")

def test_parse_calendar_devuelve_lista_no_vacía():
    regs = parse_calendar(FIXTURE)
    assert isinstance(regs, list)
    assert len(regs) > 0

def test_campos_registro():
    reg = parse_calendar(FIXTURE)[0]
    assert set(reg.keys()) == {"empresa", "pais", "impuesto", "fecha", "estado"}

def test_agrupar_por_fecha():
    regs = parse_calendar(FIXTURE)
    fechas = {r["fecha"] for r in regs}
    # Debe haber registros en al menos dos fechas distintas
    assert len(fechas) >= 2
