"""Conexión y proyección de columnas de SQL Server."""
import certifi
from .config import SERVER, DATABASE, USER, VIEW, validar_configuracion_sql, vista_sql_segura


def conexion(password):
    validar_configuracion_sql()
    import pytds
    return pytds.connect(
        server=SERVER, database=DATABASE, user=USER,
        password=password, port=1433, timeout=120, login_timeout=30,
        cafile=certifi.where(), validate_host=True, use_mars=False
    )


def columnas_como_exportador_bd_bi(cursor):
    """Replica el tratamiento de columnas de texto de tu exportador rápido."""
    cursor.execute('''
        SELECT c.name, t.name
        FROM sys.columns c
        JOIN sys.types t ON c.user_type_id = t.user_type_id
        WHERE c.object_id = OBJECT_ID(%s)
        ORDER BY c.column_id
    ''', (VIEW,))
    columnas = cursor.fetchall()
    if not columnas:
        raise RuntimeError('No se encontraron columnas de la vista.')
    expresiones = []
    for nombre, tipo in columnas:
        identificador = '[' + str(nombre).replace(']', ']]') + ']'
        if str(tipo).lower() in ('varchar', 'char', 'text'):
            expresiones.append(f'CONVERT(NVARCHAR(MAX), {identificador}) AS {identificador}')
        else:
            expresiones.append(identificador)
    return ', '.join(expresiones)


