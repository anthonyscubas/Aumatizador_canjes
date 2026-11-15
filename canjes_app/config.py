"""Configuración externa. No almacenar secretos en el repositorio."""
import os
import re
import sys
from pathlib import Path

SERVER = os.getenv('CANJES_SQL_SERVER', '').strip()
DATABASE = os.getenv('CANJES_SQL_DATABASE', '').strip()
USER = os.getenv('CANJES_SQL_USER', '').strip()
VIEW = os.getenv('CANJES_SQL_VIEW', 'ETL.VW_REGISTRO_REPORTE_CANJE_SELL_OUT_BI').strip()
SERVICE = os.getenv('CANJES_KEYRING_SERVICE', 'Automatizador Canjes')
CORREO_REMITENTE = os.getenv('CANJES_OUTLOOK_SENDER', '').strip()
BATCH_SIZE = int(os.getenv('CANJES_BATCH_SIZE', '10000'))
CARPETA_CANJES = Path(os.getenv('CANJES_OUTPUT_DIR', str(Path.home() / 'Desktop' / 'CANJES'))).expanduser()
CARPETA_PROYECTO = (Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent.parent)
ARCHIVO_CORREOS = Path(os.getenv('CANJES_RECIPIENTS_FILE', str(CARPETA_PROYECTO / 'correos.xlsx'))).expanduser()
MESES = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
MESES_COMPLETOS = dict(zip(MESES, ('enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre')))


def vista_sql_segura():
    """Permite exclusivamente identificadores de vista schema.objeto."""
    if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*\.[A-Za-z_][A-Za-z0-9_]*', VIEW):
        raise ValueError('CANJES_SQL_VIEW debe tener formato esquema.vista (solo letras, números y _).')
    return '.'.join(f'[{part}]' for part in VIEW.split('.'))


def validar_configuracion_sql():
    faltan = [nombre for nombre, valor in [('CANJES_SQL_SERVER', SERVER), ('CANJES_SQL_DATABASE', DATABASE), ('CANJES_SQL_USER', USER)] if not valor]
    if faltan:
        raise ValueError('Configura las variables de entorno: ' + ', '.join(faltan))
    if not 1 <= BATCH_SIZE <= 100000:
        raise ValueError('CANJES_BATCH_SIZE debe estar entre 1 y 100000')
    vista_sql_segura()


def validar_configuracion_correo():
    if not CORREO_REMITENTE or '@' not in CORREO_REMITENTE:
        raise ValueError('Configura CANJES_OUTLOOK_SENDER con la dirección de la cuenta Outlook.')
