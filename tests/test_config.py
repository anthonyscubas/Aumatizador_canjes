import pytest
from canjes_app import config


def test_identificador_vista_valido(monkeypatch):
    monkeypatch.setattr(config, 'VIEW', 'ETL.VISTA_CANJES')
    assert config.vista_sql_segura() == '[ETL].[VISTA_CANJES]'

@pytest.mark.parametrize('valor', ['ETL.vista;DROP TABLE X', 'ETL.vista otra', 'ETL.vista.tabla', ''])
def test_identificador_vista_invalido(monkeypatch, valor):
    monkeypatch.setattr(config, 'VIEW', valor)
    with pytest.raises(ValueError): config.vista_sql_segura()
