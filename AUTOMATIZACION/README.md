# Automatizacion de Reporte Mensual

Sistema local para generar reportes mensuales a partir de archivos XM, actualizar un comparativo en Excel y producir informes Word. Tambien incluye un flujo para consultar tickets de Freshdesk y generar el informe mensual de Helpdesk.


## Recordar

##  1. HELPDESK
-Utilizar el último documento HelpDesk actualizado del último mes, para así cuando acabe el mes actual, sacar solo de ese mes en este caso Julio dando rango de fechas de "01-07-2026 al 31-07-26", así tomará los tickets de ese mes y se evitará actualizar tickets ya hechos anteriormente.

## 2. ARCHIVOS XM
-Modificar el CSV de Summary o Instances en caso haya una modificación que hacer para no estar moviendo nada en el Comparativo.
-Tener el comparativo actualizado también para que solo se tenga que añadir el del último mes.
-Respetar el orden de las capturas para que no se distorcionen al modificar el tamaño.



## Estructura del proyecto

La estructura esperada es la carpeta principal `AUTOMATIZACION REPORTE MENSUAL` con la aplicacion dentro de `AUTOMATIZACION`.

```text
AUTOMATIZACION REPORTE MENSUAL/
|-- .env
|-- .venv/
|-- ARCHIVOS SUMMARY E INSTANCES/
|   |-- XM-summary-YYYY-MM.xlsx
|   |-- XM-Instances-YYYY-MM.xlsx
|   `-- Comparativo Instance ...
|-- PLANTILLAS/
|   |-- Informe de Consumos - Plantilla.docx
|   `-- XM Soporte_ Informe de Helpdesk- Inspira IT_Mes_Anio - plantilla.docx
`-- AUTOMATIZACION/
    |-- .env.example
    |-- requirements.txt
    |-- README.md
    `-- frontend/
        |-- app.py
        |-- automatizar_tablas_y_comparativo.py
        |-- static/
        |   |-- app.js
        |   `-- styles.css
        `-- templates/
            |-- index.html
            |-- xm.html
            `-- helpdesk.html
```

Carpetas como `__pycache__`, `_docx_images`, `_render_informe_estudiante`, `.frontend_uploads` y archivos `.log` son generados por Python o por la ejecucion del sistema. No son la fuente principal del proyecto y normalmente pueden ignorarse.

## Instalacion

Desde la carpeta principal `AUTOMATIZACION REPORTE MENSUAL`:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r '.\AUTOMATIZACION\requirements.txt'
```

## Configuracion

Para usar el modulo de Helpdesk, crea un archivo `.env` en la carpeta principal `AUTOMATIZACION REPORTE MENSUAL`. Puedes tomar como base `AUTOMATIZACION\.env.example`:

```env
FRESHDESK_API_KEY=tu_api_key
FRESHDESK_DOMAIN=inspirasupport.freshdesk.com
FRESHDESK_RESPONDER_ID=64035132350
```

Variables:

- `FRESHDESK_API_KEY`: API key del usuario Freshdesk.
- `FRESHDESK_DOMAIN`: dominio Freshdesk.
- `FRESHDESK_RESPONDER_ID`: ID del responsable usado para filtrar tickets.

## Ejecucion

Inicia el servidor local:

```powershell
cd '.\AUTOMATIZACION\'
python .\frontend\app.py
```

Luego abre en el navegador:

```text
http://127.0.0.1:8765
```

Rutas disponibles:

- `/`: dashboard principal.
- `/xm`: automatizacion del reporte XM.
- `/helpdesk`: generacion del informe Helpdesk desde Freshdesk.


## Dependencias principales

- `openpyxl`: lectura y escritura de Excel.
- `Pillow`: generacion y preparacion de imagenes.
- `python-docx`: edicion de documentos Word.
- `docxtpl` y `jinja2`: soporte para plantillas Word.
- `python-dotenv`: carga de variables de entorno.
- `requests`: integracion con Freshdesk.
