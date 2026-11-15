from pathlib import Path
from openpyxl import Workbook

archivo = Path(__file__).resolve().parent / 'correos_demo.xlsx'
wb=Workbook()
ws=wb.active
ws.title='Destinatarios'
ws.append(['Lima ENVIAR','Lima COPIA','Provincia ENVIAR','Provincia COPIA'])
ws.append(['Ejemplos ficticios: reemplazar antes de usar','','',''])
ws.append(['lima@ejemplo.com','copia-lima@ejemplo.com','provincia@ejemplo.com','copia-prov@ejemplo.com'])
wb.save(archivo)
print(f'Creado: {archivo}')
