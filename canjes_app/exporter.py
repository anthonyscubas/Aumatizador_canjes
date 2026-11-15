"""Exportación en lotes y construcción de las tres tablas de validación."""
import datetime as dt
import os
import time
from pathlib import Path
from decimal import Decimal
from collections import defaultdict
import xlsxwriter
from .config import VIEW, BATCH_SIZE, vista_sql_segura
from .database import conexion, columnas_como_exportador_bd_bi
from .rules import fecha_de, validado_aceptado, valor_excel, nombre_fechas, normalizar_boleta


def exportar(inicio, fin, destino, password, avisar):
    if inicio > fin:
        raise ValueError('La fecha de inicio no puede superar la fecha final.')
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    sufijo = nombre_fechas(inicio, fin)
    finales = [destino / f'base_canjes_lima_({sufijo}).xlsx',
               destino / f'base_canjes_provincia_({sufijo}).xlsx']
    temporales = [destino / f'.canjes_lima_{os.getpid()}.partial.xlsx',
                  destino / f'.canjes_prov_{os.getpid()}.partial.xlsx']
    conexion_sql = None
    libros = []
    completados = False
    inicio_tiempo = time.perf_counter()
    recibidas = aceptadas = lima_n = prov_n = 0
    grupos = defaultdict(lambda: [0, 0])
    grupos_dia = defaultdict(lambda: defaultdict(lambda: Decimal('0')))
    lavagna_modificados = 0
    escalas_no_reconocidas = set()
    grupos_factura = defaultdict(
        lambda: {
            'BLANCA FLOR': Decimal('0'),
            'SAYON': Decimal('0')
        }
    )
    
    try:
        avisar('Conectando a Azure SQL...')
        conexion_sql = conexion(password)
        cur = conexion_sql.cursor()
        avisar('Leyendo estructura de la vista...')
        seleccion = columnas_como_exportador_bd_bi(cur)
        avisar('Ejecutando SELECT directo de la vista (sin filtros SQL)...')
        t0 = time.perf_counter()
        cur.execute(f'SELECT {seleccion} FROM {vista_sql_segura()}')
        avisar(f'SELECT iniciado en {time.perf_counter()-t0:.1f}s. Descargando filas...')
        encabezados = [d[0] for d in cur.description]
        pos = {str(c).strip().upper(): i for i, c in enumerate(encabezados)}
        necesarias = (
            'CATEGORIA', 'FEC_CREACION', 'USUARIO', 'VALIDADO',
            'OFICINA', 'MARCA', 'ID_REP_CANJE', 'PREMIO_CANTIDAD',
            'ID_PTO_VTA_COMPROBANTE', 'COD_LUCKY',
            'PTO_VTA_COMPROBANTE', 'CAMPANIA', 'ESCALA_ALIAS',
            'NUM_COMPROBANTE', 'PREMIO_DESCRIPCION'
        )
        faltan = [x for x in necesarias if x not in pos]
        if faltan:
            raise RuntimeError('Faltan columnas: ' + ', '.join(faltan))

        for i, ruta in enumerate(temporales):
            libros.append(xlsxwriter.Workbook(str(ruta), {
                'constant_memory': i == 0,
                'strings_to_urls': False,
                'strings_to_formulas': False,
                'use_zip64': True,
                'default_date_format': 'dd/mm/yyyy',
            }))
        wb_lima, wb_prov = libros
        sh_lima = wb_lima.add_worksheet('base_canjes')
        sh_prov = wb_prov.add_worksheet('base_canjes')
        formato_cab_l = wb_lima.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#203554'})
        formato_cab_p = wb_prov.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#203554'})
        for sh, formato in ((sh_lima, formato_cab_l), (sh_prov, formato_cab_p)):
            sh.write_row(0, 0, encabezados, formato)
            sh.freeze_panes(1, 0)
            sh.set_column(0, len(encabezados)-1, 20)

        fila_lima = fila_prov = 1
        n_lote = 0
        while True:
            n_lote += 1
            avisar(f'SQL: recibiendo lote {n_lote} (hasta {BATCH_SIZE:,} filas)...')
            t_lote = time.perf_counter()
            lote = cur.fetchmany(BATCH_SIZE)
            if not lote:
                break
            recibidas += len(lote)
            avisar(f'SQL: recibidas {recibidas:,} filas; lote en {time.perf_counter()-t_lote:.1f}s')

            for original in lote:
                if str(original[pos['CATEGORIA']] or '').strip().upper() != 'PANETONES':
                    continue
                fecha = fecha_de(original[pos['FEC_CREACION']])
                if fecha is None or not (inicio <= fecha <= fin):
                    continue
                usuario = str(original[pos['USUARIO']] or '').strip().lower()
                if usuario in ('upruebasps', 'ydiestraca'):
                    continue
                if not validado_aceptado(original[pos['VALIDADO']]):
                    continue
                aceptadas += 1
                fila = [valor_excel(v) for v in original]
                fila[pos['VALIDADO']] = 'VERDADERO'
                oficina = str(fila[pos['OFICINA']] or '').strip().upper()
                if oficina in ('LIMA', 'AYACUCHO'):
                    if fila_lima >= 1_048_576:
                        raise ValueError('Excel Lima supera el límite de filas.')
                    sh_lima.write_row(fila_lima, 0, fila)
                    fila_lima += 1
                    lima_n += 1
                else:
                    if oficina == 'JULIACA':
                        fila[pos['ID_PTO_VTA_COMPROBANTE']] = '1540885'
                        fila[pos['COD_LUCKY']] = '1000009918'
                        fila[pos['PTO_VTA_COMPROBANTE']] = 'CORPORATIVO DE COMERCIO E INVERSION'
                    campania = str(fila[pos['CAMPANIA']] or '').strip().upper()
                    escala = str(fila[pos['ESCALA_ALIAS']] or '').strip().upper()
                    escala = ''.join(escala.split())

                    premios_lavagna = {
                        'E1': Decimal('0.8'),
                        'E2': Decimal('1.2'),
                        'E3': Decimal('1.6')
                    }

                    if 'LAVAGNA' in campania:

                        if escala in premios_lavagna:
                            fila[pos['PREMIO_CANTIDAD']] = premios_lavagna[escala]
                            lavagna_modificados += 1
                        else:
                            escalas_no_reconocidas.add(escala)
                        
                    if fila_prov >= 1_048_576:
                        raise ValueError('Excel Provincia supera el límite de filas.')
                    sh_prov.write_row(fila_prov, 0, fila)
                    fila_prov += 1
                    prov_n += 1
                    marca = str(fila[pos['MARCA']] or '').strip().upper()

                    oficina_pdv = str(fila[pos['OFICINA']] or '').strip()
                    cod_lucky = str(fila[pos['COD_LUCKY']] or '').strip()
                    nombre_pdv = str(fila[pos['PTO_VTA_COMPROBANTE']] or '').strip()

                    clave = (oficina_pdv, cod_lucky, nombre_pdv, marca)

                    if fila[pos['ID_REP_CANJE']] not in ('', None):
                        grupos[clave][0] += 1

                    cantidad = fila[pos['PREMIO_CANTIDAD']]

                    if cantidad not in ('', None):

                        try:
                            premio = Decimal(str(cantidad))

                        except Exception:
                            raise ValueError(
                                f"PREMIO_CANTIDAD no numérico: {cantidad}"
                            )

                        grupos[clave][1] += premio

                        if marca in ('BLANCA FLOR', 'SAYON'):

                            clave_factura = (
                                str(fila[pos['OFICINA']] or '').strip(),
                                str(fila[pos['COD_LUCKY']] or '').strip(),
                                str(fila[pos['PTO_VTA_COMPROBANTE']] or '').strip(),
                                fecha.strftime('%d/%m/%Y'),
                                normalizar_boleta(fila[pos['NUM_COMPROBANTE']]),
                                str(fila[pos['CAMPANIA']] or '').strip(),
                                
                                str(fila[pos['PREMIO_DESCRIPCION']] or '').strip()
                            )

                            grupos_factura[clave_factura][marca] += premio
                            escala_dia = ''.join(
                                str(fila[pos['ESCALA_ALIAS']] or '').upper().split()
                            )

                            clave_dia = (
                                oficina_pdv,
                                cod_lucky,
                                nombre_pdv,
                                fecha.strftime('%d/%m/%Y'),
                                normalizar_boleta(fila[pos['NUM_COMPROBANTE']]),
                                str(fila[pos['CAMPANIA']] or '').strip(),
                                str(fila[pos['PREMIO_DESCRIPCION']] or '').strip()
                            )

                            grupos_dia[clave_dia][(marca, escala_dia)] += premio
            avisar(f'Filtradas {aceptadas:,} de {recibidas:,} | Lima {lima_n:,} | Provincia {prov_n:,}')
        avisar(f'LAVAGNA modificados: {lavagna_modificados}')

        if escalas_no_reconocidas:
            avisar(
                'ESCALA_ALIAS no reconocidos: '
                + repr(sorted(escalas_no_reconocidas))
            )

        avisar('Construyendo hoja Validación...')

        valid = wb_prov.add_worksheet('Validación')

        rojo = wb_prov.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#BD2530',
            'align': 'center',
            'valign': 'vcenter'
        })

        naranja = wb_prov.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#EF8D22',
            'align': 'center'
        })

        azul = wb_prov.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#123B71',
            'align': 'center'
        })

        cab = wb_prov.add_format({
            'bold': True,
            'bg_color': '#DCE6F1',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        cab_azul = wb_prov.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#002060',
            'border': 1,
            'align': 'center'
        })

        cab_amarillo = wb_prov.add_format({
            'bold': True,
            'bg_color': '#FFF59D',
            'border': 1,
            'align': 'center'
        })

        numero = wb_prov.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })

        texto = wb_prov.add_format({
            'border': 1
        })

        total_fmt = wb_prov.add_format({
            'bold': True,
            'bg_color': '#DCE6F1',
            'border': 1,
            'num_format': '#,##0.00'
        })


        valid.set_column('A:A', 17)
        valid.set_column('B:B', 19)
        valid.set_column('C:C', 43)
        valid.set_column('D:M', 15)
        valid.set_column('N:AO', 12)


        valid.set_row(0, 30)

        valid.merge_range(
            'A1:M1',
            'PREMIO POR PDV (Unidades)',
            rojo
        )

        valid.merge_range(
            'O1:X1',
            'PREMIO POR FACTURA',
            naranja
        )

        valid.merge_range(
            'AA1:AO1',
            'CANJES POR DIA DE RELEVO',
            azul
        )



        valid.merge_range('A4:A5', 'OFICINA', cab)
        valid.merge_range('B4:B5', 'COD_LUCKY', cab)
        valid.merge_range('C4:C5', 'PTO_VTA_COMPROBANTE', cab)

        valid.merge_range('D4:E4', 'BLANCA FLOR', cab)
        valid.merge_range('F4:G4', 'SAYON', cab_amarillo)

        valid.write('D5', 'Canjes', cab)
        valid.write('E5', 'Premio', cab)
        valid.write('F5', 'Canjes', cab)
        valid.write('G5', 'Premio', cab)

        valid.merge_range('H4:H5', 'Total Canjes', cab)
        valid.merge_range('I4:I5', 'Total Premio', cab)

        valid.merge_range('J4:K4', 'Supervisor', cab_azul)
        valid.write('J5', 'BF', cab_amarillo)
        valid.write('K5', 'Sayon', cab_amarillo)

        valid.merge_range('L4:M4', 'Diferencia', cab_azul)
        valid.write('L5', 'BF', cab)
        valid.write('M5', 'Sayon', cab)


        pdvs = {}

        for clave, valores in grupos.items():

            oficina, cod, nombre, marca = clave

            if marca not in ('BLANCA FLOR', 'SAYON'):
                raise ValueError(
                    f"Marca no contemplada en Validación: {marca}"
                )

            llave_pdv = (oficina, cod, nombre)

            if llave_pdv not in pdvs:
                pdvs[llave_pdv] = {
                    'BF_CANJES': 0,
                    'BF_PREMIO': Decimal('0'),
                    'SAYON_CANJES': 0,
                    'SAYON_PREMIO': Decimal('0')
                }

            if marca == 'BLANCA FLOR':
                pdvs[llave_pdv]['BF_CANJES'] += valores[0]
                pdvs[llave_pdv]['BF_PREMIO'] += valores[1]

            elif marca == 'SAYON':
                pdvs[llave_pdv]['SAYON_CANJES'] += valores[0]
                pdvs[llave_pdv]['SAYON_PREMIO'] += valores[1]


        fila_excel = 5

        totales = [Decimal('0')] * 6

        for (oficina, cod, nombre), d in sorted(pdvs.items()):

            bf_c = d['BF_CANJES']
            bf_p = d['BF_PREMIO']

            sy_c = d['SAYON_CANJES']
            sy_p = d['SAYON_PREMIO']

            total_c = bf_c + sy_c
            total_p = bf_p + sy_p

            valores = [
                oficina,
                cod,
                nombre,
                bf_c,
                bf_p,
                sy_c,
                sy_p,
                total_c,
                total_p
            ]

            for columna, valor in enumerate(valores):

                if columna < 3:
                    valid.write(fila_excel, columna, valor, texto)
                else:
                    valid.write_number(
                        fila_excel,
                        columna,
                        float(valor),
                        numero
                    )

            for i, valor in enumerate(valores[3:]):
                totales[i] += valor

            for columna in range(9, 13):
                valid.write_blank(
                    fila_excel,
                    columna,
                    None,
                    texto
                )

            fila_excel += 1


        valid.merge_range(
            fila_excel,
            0,
            fila_excel,
            2,
            'Total general',
            total_fmt
        )

        for i, valor in enumerate(totales):
            valid.write_number(
                fila_excel,
                i + 3,
                float(valor),
                total_fmt
            )


        valid.freeze_panes(5, 3)

        if fila_excel > 5:
            valid.autofilter(4, 0, fila_excel - 1, 12)

        avisar('Construyendo segunda tabla: PREMIO POR FACTURA...')


        factura_cab = wb_prov.add_format({
            'bold': True,
            'bg_color': '#DCE6F1',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        factura_sayon = wb_prov.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#F5A623',
            'border': 1,
            'align': 'center'
        })

        factura_num = wb_prov.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })

        factura_texto = wb_prov.add_format({
            'border': 1
        })

        factura_total = wb_prov.add_format({
            'bold': True,
            'bg_color': '#DCE6F1',
            'border': 1,
            'num_format': '#,##0.00'
        })


        valid.set_column('O:O', 17)
        valid.set_column('P:P', 18)
        valid.set_column('Q:Q', 42)
        valid.set_column('R:R', 16)
        valid.set_column('S:S', 18)
        valid.set_column('T:T', 26)
        valid.set_column('U:U', 35)
        valid.set_column('V:X', 17)


        valid.merge_range('O4:O5', 'OFICINA', factura_cab)
        valid.merge_range('P4:P5', 'COD_LUCKY', factura_cab)
        valid.merge_range('Q4:Q5', 'PTO_VTA_COMPROBANTE', factura_cab)
        valid.merge_range('R4:R5', 'FECHA', factura_cab)
        valid.merge_range('S4:S5', 'BOLETA', factura_cab)
        valid.merge_range('T4:T5', 'CAMPANIA', factura_cab)
        valid.merge_range('U4:U5', 'PREMIO_DESCRIPCION', factura_cab)

        valid.merge_range('V4:W4', 'MARCA', factura_cab)
        valid.write('V5', 'BLANCA FLOR', factura_cab)
        valid.write('W5', 'SAYON', factura_sayon)

        valid.merge_range('X4:X5', 'Total general', factura_cab)


        fila_factura = 5

        total_bf = Decimal('0')
        total_sayon = Decimal('0')

        for clave_factura, marcas in sorted(
            grupos_factura.items(),
            key=lambda x: tuple(str(v) for v in x[0])
        ):

            oficina, cod, nombre, fecha, boleta, campania, descripcion = clave_factura

            premio_bf = marcas['BLANCA FLOR']
            premio_sayon = marcas['SAYON']

            total_fila = premio_bf + premio_sayon

            datos_texto = [
                oficina,
                cod,
                nombre,
                fecha,
                boleta,
                campania,
                descripcion
            ]

            for i, valor in enumerate(datos_texto):

                valid.write(
                    fila_factura,
                    14 + i,
                    valor,
                    factura_texto
                )

            valid.write_number(
                fila_factura,
                21,
                float(premio_bf),
                factura_num
            )

            valid.write_number(
                fila_factura,
                22,
                float(premio_sayon),
                factura_num
            )

            valid.write_number(
                fila_factura,
                23,
                float(total_fila),
                factura_num
            )

            total_bf += premio_bf
            total_sayon += premio_sayon

            fila_factura += 1

        valid.merge_range(
            fila_factura,
            14,
            fila_factura,
            20,
            'Total general',
            factura_total
        )

        valid.write_number(
            fila_factura,
            21,
            float(total_bf),
            factura_total
        )

        valid.write_number(
            fila_factura,
            22,
            float(total_sayon),
            factura_total
        )

        valid.write_number(
            fila_factura,
            23,
            float(total_bf + total_sayon),
            factura_total
        )

        avisar(
            f"Segunda tabla completada: "
            f"{len(grupos_factura):,} combinaciones de factura"
        )
        avisar('Construyendo tercera tabla...')

        columnas_marca = [
            ('BLANCA FLOR', 'E1'),
            ('BLANCA FLOR', 'E3'),
            ('SAYON', 'E1'),
            ('SAYON', 'E2'),
            ('SAYON', 'E3')
        ]

        campos = [
            'OFICINA', 'COD_LUCKY', 'PTO_VTA_COMPROBANTE',
            'FECHA', 'BOLETA', 'CAMPANIA', 'PREMIO_DESCRIPCION'
        ]

        for i, campo in enumerate(campos):
            valid.merge_range(3, 26+i, 5, 26+i, campo, cab)

        valid.merge_range('AH4:AL4', 'MARCA', cab)

        valid.merge_range('AH5:AI5', 'BLANCA FLOR', cab)
        valid.merge_range('AJ5:AL5', 'SAYON', cab_amarillo)

        for i, (_, escala) in enumerate(columnas_marca):
            valid.write(5, 33+i, escala, cab)

        valid.merge_range('AM4:AM6', 'Total general', cab)

        valid.set_column('AA:AA', 17)
        valid.set_column('AB:AB', 19)
        valid.set_column('AC:AC', 43)
        valid.set_column('AD:AE', 16)
        valid.set_column('AF:AF', 26)
        valid.set_column('AG:AG', 35)
        valid.set_column('AH:AM', 16)

        fila_dia = 6
        totales_dia = [Decimal('0')] * 5

        for clave_dia, premios in sorted(grupos_dia.items()):

            for i, valor in enumerate(clave_dia):
                valid.write(fila_dia, 26+i, valor, texto)

            total_fila = Decimal('0')

            for i, combinacion in enumerate(columnas_marca):

                cantidad = premios.get(combinacion, Decimal('0'))

                valid.write_number(
                    fila_dia,
                    33+i,
                    float(cantidad),
                    numero
                )

                totales_dia[i] += cantidad
                total_fila += cantidad

            valid.write_number(
                fila_dia,
                38,
                float(total_fila),
                numero
            )

            fila_dia += 1

        valid.merge_range(
            fila_dia, 26,
            fila_dia, 32,
            'Total general',
            total_fmt
        )

        for i, cantidad in enumerate(totales_dia):
            valid.write_number(
                fila_dia,
                33+i,
                float(cantidad),
                total_fmt
            )

        valid.write_number(
            fila_dia,
            38,
            float(sum(totales_dia)),
            total_fmt
        )

        avisar(
            f'Tercera tabla completada: {len(grupos_dia):,} registros agrupados'
        )
        avisar(
            f"Validación completada: {len(pdvs):,} PDV"
        )
        for libro in libros:
            libro.close()
        libros.clear()
        for temporal, final in zip(temporales, finales):
            os.replace(temporal, final)
        completados = True
        avisar(f'COMPLETADO en {time.perf_counter()-inicio_tiempo:.1f}s | Lima {lima_n:,} | Provincia {prov_n:,}')
        return finales, lima_n, prov_n
    finally:
        if conexion_sql:
            try:
                conexion_sql.close()
            except Exception:
                pass
        for libro in libros:
            try:
                libro.close()
            except Exception:
                pass
        if not completados:
            for temporal in temporales:
                try:
                    temporal.unlink(missing_ok=True)
                except OSError:
                    pass


