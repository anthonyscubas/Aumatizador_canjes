"""Prueba integrada en memoria, sin Azure SQL, Outlook ni datos empresariales."""
import datetime as dt
from decimal import Decimal
from openpyxl import load_workbook
from canjes_app import exporter

HEADERS = [
    'CATEGORIA','FEC_CREACION','USUARIO','VALIDADO','OFICINA','MARCA',
    'ID_REP_CANJE','PREMIO_CANTIDAD','ID_PTO_VTA_COMPROBANTE','COD_LUCKY',
    'PTO_VTA_COMPROBANTE','CAMPANIA','ESCALA_ALIAS','NUM_COMPROBANTE','PREMIO_DESCRIPCION'
]

def row(**changes):
    base=dict(CATEGORIA='PANETONES',FEC_CREACION=dt.datetime(2026,11,12,10),
        USUARIO='operador_demo',VALIDADO=1,OFICINA='CUSCO',MARCA='SAYON',
        ID_REP_CANJE=1, PREMIO_CANTIDAD=Decimal('1.2'),ID_PTO_VTA_COMPROBANTE='10',
        COD_LUCKY='0001',PTO_VTA_COMPROBANTE='TIENDA DEMO',CAMPANIA='NORMAL',
        ESCALA_ALIAS='E2',NUM_COMPROBANTE='001-TEST',PREMIO_DESCRIPCION='CANJE DEMO')
    base.update(changes)
    return tuple(base[k] for k in HEADERS)

class Cursor:
    description=[(h,) for h in HEADERS]
    def __init__(self):
        self.rows=[row(),row(OFICINA='LIMA',MARCA='BLANCA FLOR'),row(CATEGORIA='OTROS'),
            row(USUARIO='upruebasps'),row(FEC_CREACION=dt.date(2026,11,10))]
        self.delivered=False
    def execute(self,sql): self.sql=sql
    def fetchmany(self,n):
        if self.delivered: return []
        self.delivered=True
        return self.rows
class Connection:
    def __init__(self): self.cur=Cursor()
    def cursor(self): return self.cur
    def close(self): pass

def test_end_to_end_export(monkeypatch,tmp_path):
    con=Connection()
    monkeypatch.setattr(exporter,'conexion',lambda password: con)
    monkeypatch.setattr(exporter,'columnas_como_exportador_bd_bi',lambda cursor: ', '.join(HEADERS))
    outputs,lima,prov=exporter.exportar(dt.date(2026,11,12),dt.date(2026,11,12),tmp_path,'demo',lambda msg: None)
    assert (lima,prov)==(1,1)
    assert all(p.exists() for p in outputs)
    wb=load_workbook(outputs[1],read_only=True,data_only=True)
    try:
        assert wb.sheetnames==['base_canjes','Validación']
        sh=wb['Validación']
        assert sh['A4'].value=='OFICINA'
        assert sh['A7'].value=='Total general'
        assert sh['I6'].value==1.2
    finally:
        wb.close()
    wb=load_workbook(outputs[0],read_only=True)
    try:
        assert wb['base_canjes'].max_row==2
    finally:
        wb.close()

def test_tabla_html_y_cuerpo_provincia(monkeypatch, tmp_path):
    from canjes_app.mailer import tabla_validacion_html, cuerpo_correo_provincia
    con=Connection()
    monkeypatch.setattr(exporter,'conexion',lambda password: con)
    monkeypatch.setattr(exporter,'columnas_como_exportador_bd_bi',lambda cursor: ', '.join(HEADERS))
    outputs, *_ = exporter.exportar(dt.date(2026,11,12),dt.date(2026,11,12),tmp_path,'demo',lambda msg: None)
    tabla=tabla_validacion_html(outputs[1])
    assert 'Total general' in tabla
    assert 'TIENDA DEMO' in tabla
    assert 'PREMIO POR FACTURA' not in tabla
    cuerpo=cuerpo_correo_provincia(outputs[1])
    assert 'el día 12 de noviembre' in cuerpo
    assert '<table' in cuerpo
