"""Reglas puras de fechas, limpieza, filtros y representación Excel."""
import datetime as dt
import re
from pathlib import Path
from .config import MESES, MESES_COMPLETOS


def fecha_de(valor):
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    if valor is None:
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    try:
        return dt.datetime.fromisoformat(texto.replace('Z', '+00:00')).date()
    except ValueError:
        pass
    for formato in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y%m%d'):
        try:
            return dt.datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    raise ValueError(f'Formato de FEC_CREACION desconocido: {texto[:40]}')


def validado_aceptado(valor):
    if valor is None:
        return True
    if isinstance(valor, (bool, int)):
        return valor == 1
    return str(valor).strip().upper() in ('', '1', 'TRUE', 'VERDADERO')


def valor_excel(valor):
    if valor is None:
        return ''
    if isinstance(valor, (bytes, bytearray, memoryview)):
        return bytes(valor).hex()
    if isinstance(valor, bool):
        return 'VERDADERO' if valor else 'FALSO'
    if isinstance(valor, dt.datetime):
        return valor.strftime('%d/%m/%Y %H:%M:%S')
    if isinstance(valor, dt.date):
        return valor.strftime('%d/%m/%Y')
    if isinstance(valor, dt.time):
        return valor.strftime('%H:%M:%S')
    return valor


def nombre_fechas(inicio, fin):
    if inicio.year == fin.year and inicio.month == fin.month:
        return f'del {inicio.day} y {fin.day} {MESES[inicio.month-1]}'
    return (f'del {inicio.day} {MESES[inicio.month-1]} {inicio.year} '
            f'y {fin.day} {MESES[fin.month-1]} {fin.year}')


def normalizar_boleta(valor):
    boleta = str(valor or '').strip()
    boleta = boleta.split('-', 1)[0].strip()
    boleta = boleta.strip()
    return boleta


def obtener_fechas_nombre(archivo):
    """
    Ejemplo:
    base_canjes_lima_(del 11 y 12 nov).xlsx

    Resultado:
    los días 11 y 12 de noviembre
    """

    nombre = Path(archivo).stem.lower()

    patron = (
        r'del\s+(\d{1,2})\s+y\s+'
        r'(\d{1,2})\s+'
        r'(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)'
        r'(?:\s+\d{4})?'
    )

    encontrado = re.search(patron, nombre)

    
    if encontrado:
        dia1 = int(encontrado.group(1))
        dia2 = int(encontrado.group(2))
        mes = MESES_COMPLETOS[encontrado.group(3)]

        if dia1 == dia2:
            return f'el día {dia1} de {mes}'

        return f'los días {dia1} y {dia2} de {mes}'


    patron_diferente = (
        r'del\s+(\d{1,2})\s+'
        r'(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)'
        r'\s+(\d{4})\s+y\s+'
        r'(\d{1,2})\s+'
        r'(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)'
        r'\s+(\d{4})'
    )

    encontrado = re.search(patron_diferente, nombre)

    if encontrado:
        d1, m1, a1, d2, m2, a2 = encontrado.groups()

        return (
            f'los días {d1} de {MESES_COMPLETOS[m1]} '
            f'de {a1} y {d2} de {MESES_COMPLETOS[m2]} '
            f'de {a2}'
        )

    raise ValueError(
        f'No se pudieron identificar las fechas: {archivo}'
    )


