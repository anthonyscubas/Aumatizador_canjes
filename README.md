# Automatizador de Canjes | Python, SQL Server y Outlook

**Sistema de automatización de reportes empresariales desarrollado en Python**, diseñado para optimizar la extracción, transformación, validación y distribución de información desde SQL Server hacia Microsoft Excel y Outlook.

El proyecto integra procesamiento de datos, generación de reportes dinámicos y una interfaz gráfica que permite ejecutar el flujo sin conocimientos de programación.

## Tecnologías utilizadas

- **Python:** desarrollo y automatización del proceso.
- **SQL Server / Azure SQL:** extracción de información empresarial.
- **Pandas:** transformación y procesamiento de datos.
- **XlsxWriter:** generación de reportes Excel.
- **Tkinter:** interfaz gráfica de escritorio.
- **PyWin32:** integración con Microsoft Outlook.
- **Keyring:** almacenamiento seguro de credenciales.
- **Pytest:** pruebas automatizadas.
- **GitHub Actions:** ejecución de pruebas de calidad.
- **PyInstaller:** distribución mediante ejecutable de Windows.

## Problema identificado

La elaboración de reportes de canjes requiere consultar información desde una base de datos, aplicar reglas de negocio, clasificar registros por zonas geográficas, generar archivos Excel y preparar correos para diferentes destinatarios.

Cuando estas actividades se realizan manualmente, aumentan el tiempo operativo y el riesgo de errores en la consolidación y distribución de la información.

## Solución desarrollada

Desarrollé una aplicación de escritorio que centraliza y automatiza el procesamiento de reportes de canjes mediante un flujo controlado.

**Funcionalidades principales:**

1. **Extracción de datos:** conexión a SQL Server y lectura de registros por lotes.
2. **Procesamiento:** filtrado por fechas, validación de registros y normalización de identificadores.
3. **Segmentación geográfica:** clasificación automática de información en reportes de Lima y Provincia.
4. **Reportes Excel:** creación de archivos estructurados con tablas de validación y agrupaciones.
5. **Reportes dinámicos:** identificación automática de rangos variables y totales generales.
6. **Automatización de correos:** generación de mensajes HTML con reportes adjuntos e integración con Outlook.
7. **Interfaz gráfica:** selección de fechas, seguimiento del procesamiento y confirmación del envío.
8. **Seguridad:** administración de credenciales y separación de configuraciones sensibles.

## Arquitectura del sistema

              INTERFAZ GRÁFICA
                  Tkinter
                     |
              CONTROL DEL FLUJO
                     |
              EXTRACCIÓN SQL
                     |
             PROCESAMIENTO PYTHON
                     |
              REGLAS DE NEGOCIO
                     |
             GENERACIÓN DE EXCEL
                     |
              +------+------+
              |             |
            LIMA         PROVINCIA
              |             |
              +------+------+
                     |
             VALIDACIÓN DE DATOS
                     |
             REVISIÓN DEL USUARIO
                     |
             ENVÍO CON OUTLOOK


La aplicación utiliza una arquitectura modular que separa la conexión a base de datos, las reglas de negocio, la exportación Excel, la gestión de correos y la interfaz gráfica.

## Aspectos técnicos destacados

### Procesamiento de información

- Extracción de registros mediante consultas SQL.
- Procesamiento por lotes con `fetchmany()`.
- Transformación y clasificación de registros.
- Normalización de identificadores de boletas.
- Aplicación de reglas empresariales para consolidar información.

### Automatización de Excel

- Generación independiente de reportes por región.
- Creación de tablas de validación.
- Agrupaciones por punto de venta, factura y fecha.
- Identificación dinámica de filas de totales.
- Protección frente a la interpretación de fórmulas en valores exportados.

### Integración con Outlook

- Generación automática de correos HTML.
- Incorporación de tablas de resumen en el mensaje.
- Adjuntos Excel según región.
- Configuración externa de destinatarios.
- Confirmación manual antes del envío.

### Calidad y mantenimiento

- Separación del código en módulos.
- Pruebas unitarias y de exportación con datos simulados.
- Integración continua mediante GitHub Actions.
- Configuración mediante variables de entorno.
- Manejo de excepciones y registros de ejecución.

## Flujo de funcionamiento

**1. Selección:** el usuario define el periodo de consulta.

**2. Extracción:** el sistema obtiene registros desde SQL Server.

**3. Transformación:** se aplican filtros y reglas de negocio.

**4. Consolidación:** se generan los reportes correspondientes a Lima y Provincia.

**5. Validación:** el usuario revisa los archivos Excel generados.

**6. Distribución:** previa confirmación, Outlook envía los reportes a los destinatarios configurados.

## Ejecución del proyecto

**Requisitos:**

- Windows 10/11.
- Python 3.11 o 3.13.
- Acceso autorizado a SQL Server.
- Microsoft Outlook clásico de escritorio.

Clonar el repositorio:

```bash
git clone https://github.com/TU-USUARIO/automatizador-canjes-python.git
cd automatizador-canjes-python
```

Instalar dependencias:

```bash
python -m pip install -r requirements.txt
```

Configurar las variables de entorno de conexión y destinatarios, siguiendo las indicaciones del archivo `.env.example`.

Ejecutar:

```bash
python main.py
```

Para generar un ejecutable de Windows:

```bat
build_windows.bat
```

## Pruebas

El proyecto incluye pruebas automatizadas de las reglas de procesamiento, validaciones, normalización y generación de archivos Excel.

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Las pruebas se ejecutan también mediante GitHub Actions.

**Estado:** pruebas locales con datos simulados completadas. La conexión real con Azure SQL y el envío de correos requieren validación en un entorno Windows autorizado.

## Impacto del proyecto

El sistema busca reducir las tareas repetitivas relacionadas con la generación y distribución de reportes empresariales.

**Beneficios principales:**

- Centralización del proceso de extracción y generación de reportes.
- Menor intervención manual en la transformación de información.
- Estandarización de archivos Excel.
- Automatización de la preparación de comunicaciones.
- Mayor trazabilidad del proceso mediante la interfaz y sus registros.

*Las métricas de tiempo ahorrado y mejora de productividad están pendientes de medición en operación real.*

## Consideraciones de seguridad

Este repositorio está diseñado como una demostración técnica de portafolio.

No incluye credenciales, datos de producción, destinatarios corporativos ni información confidencial. Su utilización en entornos empresariales requiere autorización, configuración y validación de las reglas de negocio correspondientes.

## Autor

**Desarrollo:** Proyecto de automatización y procesamiento de datos con Python.

**Áreas:** Data Analytics · Business Intelligence · Process Automation · Python Development.
