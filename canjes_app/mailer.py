"""Lectura de destinatarios, HTML de validación y Outlook COM."""
import html
import re
from pathlib import Path
from decimal import Decimal
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from .config import CARPETA_CANJES, CARPETA_PROYECTO, ARCHIVO_CORREOS, CORREO_REMITENTE, validar_configuracion_correo
from .rules import obtener_fechas_nombre


def obtener_ultimos_excels():
    carpeta = CARPETA_CANJES

    if not carpeta.exists():
        raise FileNotFoundError(
            f'No existe la carpeta CANJES:\n{carpeta}'
        )

    archivos_lima = list(
        carpeta.glob('base_canjes_lima_(*).xlsx')
    )

    archivos_prov = list(
        carpeta.glob('base_canjes_provincia_(*).xlsx')
    )

    if not archivos_lima:
        raise FileNotFoundError(
            'No se encontró ningún Excel de Lima.'
        )

    if not archivos_prov:
        raise FileNotFoundError(
            'No se encontró ningún Excel de Provincia.'
        )

    def periodo(ruta, prefijo):
        nombre = ruta.name

        if not nombre.startswith(prefijo):
            return None

        return nombre[len(prefijo):]

    lima_por_periodo = {
        periodo(p, 'base_canjes_lima_'): p
        for p in archivos_lima
    }

    parejas = []

    for provincia in archivos_prov:
        sufijo = periodo(
            provincia,
            'base_canjes_provincia_'
        )

        lima = lima_por_periodo.get(sufijo)

        if lima is not None:
            parejas.append((
                min(
                    lima.stat().st_mtime,
                    provincia.stat().st_mtime
                ),
                lima,
                provincia
            ))

    if not parejas:
        raise ValueError(
            'No se encontraron archivos Lima y Provincia '
            'con el mismo periodo de fechas.'
        )

    _, archivo_lima, archivo_prov = max(
        parejas,
        key=lambda x: x[0]
    )

    return archivo_lima, archivo_prov


def leer_destinatarios():
    """
    correos.xlsx:

    A = Lima ENVIAR
    B = Lima COPIA
    C = Provincia ENVIAR
    D = Provincia COPIA
    """

    if not ARCHIVO_CORREOS.exists():
        raise FileNotFoundError(
            f'No existe correos.xlsx en:\n{CARPETA_PROYECTO}'
        )

    wb = load_workbook(
        ARCHIVO_CORREOS,
        read_only=True,
        data_only=True
    )

    try:
        ws = wb.active

        def obtener_columna(columna):
            correos = []

            for fila in range(3, ws.max_row + 1):
                valor = ws.cell(fila, columna).value

                if not valor:
                    continue

                partes = re.split(r'[;\n]+', str(valor))

                for correo in partes:
                    correo = correo.strip()

                    if correo:
                        if not re.fullmatch(
                            r'[^@\s]+@[^@\s]+\.[^@\s]+',
                            correo
                        ):
                            raise ValueError(
                                f'Correo no válido en '
                                f'{ws.cell(fila, columna).coordinate}: '
                                f'{correo}'
                            )

                        correos.append(correo)

            return list(dict.fromkeys(correos))

        return {
            'lima': {
                'para': obtener_columna(1),
                'cc': obtener_columna(2)
            },
            'provincia': {
                'para': obtener_columna(3),
                'cc': obtener_columna(4)
            }
        }

    finally:
        wb.close()


def tabla_validacion_html(archivo_provincia):
    """
    Extrae únicamente la primera tabla de Validación:
    - Inicio fijo: A4
    - Última columna: M
    - Última fila: Total general (dinámico)
    """

    wb = load_workbook(
        archivo_provincia,
        data_only=True
    )

    try:
        if 'Validación' not in wb.sheetnames:
            raise ValueError(
                'El Excel Provincia no tiene la hoja Validación.'
            )

        ws = wb['Validación']

        inicio_fila = 4
        inicio_col = 1
        fin_col = 13  

        fin_fila = None

        for fila in range(6, ws.max_row + 1):
            valor = ws.cell(fila, 1).value

            if (
                valor is not None
                and str(valor).strip().casefold() == 'total general'
            ):
                fin_fila = fila
                break

        if fin_fila is None:
            raise ValueError(
                'No se encontró Total general en la '
                'primera tabla de Validación.'
            )

        combinadas = {}

        for rango in ws.merged_cells.ranges:
            if (
                rango.min_row >= inicio_fila
                and rango.min_row <= fin_fila
                and rango.min_col >= inicio_col
                and rango.min_col <= fin_col
            ):
                combinadas[
                    (rango.min_row, rango.min_col)
                ] = rango

        tabla = [
            '<div style="width:100%;overflow-x:auto;">',
            '<table cellpadding="6" cellspacing="0" '
            'style="border-collapse:collapse;'
            'font-family:Arial,sans-serif;'
            'font-size:11px;'
            'border:1px solid #A9B7C6;">'
        ]

        for fila in range(inicio_fila, fin_fila + 1):

            es_cabecera = fila in (4, 5)
            es_total = fila == fin_fila

            tabla.append('<tr>')

            for columna in range(inicio_col, fin_col + 1):

                celda = ws.cell(fila, columna)

                if isinstance(celda, MergedCell):
                    continue

                rango = combinadas.get((fila, columna))

                rowspan = 1
                colspan = 1

                if rango is not None:
                    rowspan = min(
                        rango.max_row, fin_fila
                    ) - fila + 1

                    colspan = min(
                        rango.max_col, fin_col
                    ) - columna + 1

                valor = celda.value

                if valor is None:
                    valor = ''

                elif isinstance(valor, (int, float, Decimal)):
                    valor = f'{valor:,.2f}'

                valor = html.escape(str(valor))

                fondo = '#FFFFFF'
                color = '#111827'

                if es_cabecera:
                    fondo = '#DCE6F1'

                    if fila == 4 and columna in (10, 12):
                        fondo = '#002060'
                        color = '#FFFFFF'

                    elif fila == 4 and columna == 6:
                        fondo = '#FFF59D'

                    elif fila == 5 and columna in (10, 11):
                        fondo = '#FFF59D'

                elif es_total:
                    fondo = '#DCE6F1'

                alineacion = (
                    'left'
                    if columna in (1, 2, 3)
                    else 'center'
                )

                negrita = (
                    'bold'
                    if es_cabecera or es_total
                    else 'normal'
                )

                estilo = (
                    'border:1px solid #A9B7C6;'
                    f'background-color:{fondo};'
                    f'color:{color};'
                    f'font-weight:{negrita};'
                    f'text-align:{alineacion};'
                    'vertical-align:middle;'
                    'padding:7px 9px;'
                    'white-space:nowrap;'
                )

                etiqueta = 'th' if es_cabecera else 'td'

                tabla.append(
                    f'<{etiqueta} '
                    f'rowspan="{rowspan}" '
                    f'colspan="{colspan}" '
                    f'style="{estilo}">'
                    f'{valor}'
                    f'</{etiqueta}>'
                )

            tabla.append('</tr>')

        tabla.append('</table></div>')

        return ''.join(tabla)

    finally:
        wb.close()


def cuerpo_correo_lima(archivo):
    fechas = obtener_fechas_nombre(archivo)

    return f"""
    <html>
    <body style="font-family:Arial;font-size:12px;">
        <p>Buen día, Equipo.</p>

        <p>
        Les comparto los canjes registrados {fechas}.
        </p>

        <p>Saludos,</p>
    </body>
    </html>
    """


def cuerpo_correo_provincia(archivo):
    fechas = obtener_fechas_nombre(archivo)

    tabla = tabla_validacion_html(archivo)

    return f"""
    <html>
    <body style="font-family:Arial;font-size:12px;">
        <p>Buen día, Equipo.</p>

        <p>
        Les comparto los canjes registrados {fechas}.
        </p>

        <p>
        Comparto el detalle de canjes registrados
        por PDV a validar.
        </p>

        {tabla}

        <p>Saludos,</p>
    </body>
    </html>
    """


def enviar_correos_outlook(avisar):
    """
    Envía únicamente DOS correos:
    1. Lima-Ayacucho
    2. Provincia

    No envía correos individuales.
    """

    validar_configuracion_correo()
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()

    try:
        avisar('Buscando los últimos Excel...')

        archivo_lima, archivo_prov = obtener_ultimos_excels()

        avisar(f'Excel Lima: {archivo_lima.name}')
        avisar(f'Excel Provincia: {archivo_prov.name}')

        destinatarios = leer_destinatarios()

        if not destinatarios['lima']['para']:
            raise ValueError(
                'No hay destinatarios principales para Lima.'
            )

        if not destinatarios['provincia']['para']:
            raise ValueError(
                'No hay destinatarios principales para Provincia.'
            )

        avisar('Conectando a Outlook...')

        outlook = win32.Dispatch('Outlook.Application')
        sesion = outlook.Session

        cuenta = None

        for acc in sesion.Accounts:
            if (
                str(acc.SmtpAddress).strip().lower()
                == CORREO_REMITENTE.lower()
            ):
                cuenta = acc
                break

        if cuenta is None:
            raise RuntimeError(
                'No se encontró la cuenta '
                f'{CORREO_REMITENTE} en Outlook de escritorio.'
            )

        configuraciones = [
            {
                'nombre': 'LIMA - AYACUCHO',
                'archivo': archivo_lima,
                'destinatarios': destinatarios['lima'],
                'asunto': (
                    'VALIDACIÓN DE CANJES PANETON '
                    'LIMA - AYACUCHO'
                ),
                'cuerpo': cuerpo_correo_lima(archivo_lima)
            },
            {
                'nombre': 'PROVINCIA',
                'archivo': archivo_prov,
                'destinatarios': destinatarios['provincia'],
                'asunto': 'VALIDACIÓN DE CANJES PANETON',
                'cuerpo': cuerpo_correo_provincia(archivo_prov)
            }
        ]

        mensajes = []

        for config in configuraciones:

            correo = outlook.CreateItem(0)

            correo.SendUsingAccount = cuenta

            correo.To = '; '.join(
                config['destinatarios']['para']
            )

            correo.CC = '; '.join(
                config['destinatarios']['cc']
            )

            correo.Subject = config['asunto']
            correo.HTMLBody = config['cuerpo']

            correo.Attachments.Add(
                str(config['archivo'].resolve())
            )

            if not correo.Recipients.ResolveAll():
                raise RuntimeError(
                    'Outlook no pudo resolver todos los '
                    f'destinatarios de {config["nombre"]}. '
                    'No se envió ningún correo.'
                )

            mensajes.append((config['nombre'], correo))
        enviados = []

        for nombre, correo in mensajes:

            avisar(f'Enviando correo de {nombre}...')

            try:
                correo.Send()
            except Exception as exc:
                raise RuntimeError(
                    f'Outlook aceptó {len(enviados)} envío(s) '
                    f'antes del error al enviar {nombre}: {exc}. '
                    'Revisa Elementos enviados antes de reintentar.'
                ) from exc

            enviados.append(nombre)

            avisar(
                f'Outlook aceptó el envío de {nombre}.'
            )

        avisar(
            'Los 2 correos fueron entregados a Outlook '
            'para su envío.'
        )

        return len(enviados)

    finally:
        pythoncom.CoUninitialize()


