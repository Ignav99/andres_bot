import pytest
from src.reader import leer_excel

# Fixture: ruta al Excel de prueba
EXCEL_PRUEBA = 'data/ejemplo_poC.xlsx'


def test_leer_excel_devuelve_lista():
    registros = leer_excel(EXCEL_PRUEBA)
    assert isinstance(registros, list)


def test_registros_tienen_campos_minimos():
    registros = leer_excel(EXCEL_PRUEBA)
    assert all('hoja' in r and 'celda' in r and 'valor' in r and 'color' in r for r in registros)

# Podrías añadir tests concretos según colores esperados
