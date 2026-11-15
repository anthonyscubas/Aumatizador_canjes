import datetime as dt
from decimal import Decimal
import pytest
from canjes_app.rules import (fecha_de, validado_aceptado, valor_excel,
                              nombre_fechas, normalizar_boleta, obtener_fechas_nombre)

@pytest.mark.parametrize('valor,esperado', [
    ('2026-11-12', dt.date(2026,11,12)),
    ('12/11/2026', dt.date(2026,11,12)),
    ('20261112', dt.date(2026,11,12)),
    (dt.datetime(2026,11,12,9), dt.date(2026,11,12)),
    (None,None),
])
def test_fecha_de(valor, esperado):
    assert fecha_de(valor)==esperado

@pytest.mark.parametrize('valor,esperado', [(1,True),(0,False),(None,True),('verdadero',True),('FALSE',False)])
def test_validado(valor, esperado):
    assert validado_aceptado(valor) is esperado

@pytest.mark.parametrize('texto,esperado', [('001-otro','001'),('001,','001,'),(' 0001-XYZ ','0001'),('', '')])
def test_boleta(texto, esperado):
    assert normalizar_boleta(texto)==esperado

def test_no_se_repite_fecha_correo():
    assert obtener_fechas_nombre('base_canjes_lima_(del 12 y 12 nov).xlsx') == 'el día 12 de noviembre'

def test_periodos_y_conversiones():
    assert nombre_fechas(dt.date(2026,11,11),dt.date(2026,11,12))=='del 11 y 12 nov'
    assert valor_excel(Decimal('1.2'))==Decimal('1.2')
    assert valor_excel(None)==''


def test_fecha_erronea():
    with pytest.raises(ValueError): fecha_de('ayer')
