# Transito

Aplicacion en Python para sincronizar documentos de SharePoint con una base de datos PostgreSQL. El proyecto expone una API con FastAPI para buscar archivos Excel y CSV dentro de carpetas definidas en Microsoft Graph, registrar su informacion en BD y procesar el contenido de archivos de recaudos y carteras.

## Que Realiza

Este proyecto automatiza el registro de documentos relacionados con transito. En particular:

- Conecta con Microsoft Graph usando credenciales de Azure AD.
- Recorre un sitio de SharePoint y localiza archivos `.xlsx` y `.csv`.
- Registra metadatos del archivo en PostgreSQL.
- Procesa contenidos segun la ruta del documento:
	- Recaudos de multas.
	- Recaudos de derechos de transito.
	- Cartera de multas.
	- Cartera de derechos de transito.
- Expone una API HTTP para disparar la sincronizacion de forma manual o mediante webhook.
- Mantiene utilidades para crear y renovar suscripciones de Microsoft Graph.
- Puede publicar la API local con un tunel temporal de ngrok para recibir webhooks desde internet (Solo para pruebas o desarrollo).

## Diagrama

Este es el flujo general del proyecto:

![Diagrama del proyecto](./diagrama.png)

## Como funciona

El flujo principal esta dividido en estos componentes:

1. `main.py` levanta una aplicacion FastAPI con los endpoints `/registrar-documento`, `/webhook`, `/health` y `/`.
2. `sharepoint.py` obtiene un token de Microsoft, consulta el sitio de SharePoint, recorre carpetas y detecta archivos Excel o CSV.
3. Cuando encuentra un archivo, lo descarga, lo lee con `pandas` y lo transforma en objetos ORM de SQLAlchemy.
4. `models.py` define la conexion a PostgreSQL y los modelos `Archivo`, `Recaudo`, `Cartera` y `Auditoria`.
5. El guardado en base de datos se hace en lotes para mejorar el rendimiento.
6. `graph_subscriptions.py` administra suscripciones de Microsoft Graph para recibir eventos de cambio.
7. `run_tunnel.py` abre un tunel con ngrok y arranca la API para pruebas o webhooks externos.

## Trazabilidad y visualización del flujo (nueva)

Se agregó un sistema de trazabilidad para registrar y visualizar cada paso del procesamiento de archivos.

- Nuevas tablas en la BD: `file_processing_logs`, `processing_step_executions`.
- Nuevo módulo backend: `backend/processing_tracker.py` (clase `FileProcessingTracker`) para registrar pasos como `descargando`, `validando`, `procesando`, `guardando`, `finalizando`.
- Nuevos endpoints para consultar el flujo y detalles de procesamiento (ver "Endpoints importantes").
- Frontend: visualizador interactivo con React Flow en `frontend/app/welcome/welcome.tsx`.

La aplicacion usa estas carpetas o prefijos para clasificar los archivos en sharepoint:

- `/RECAUDO/MULTAS/`
- `/RECAUDO/DERECHOS DE TRANSITO/`
- `/CARTERA/MULTAS/`
- `/CARTERA/DERECHOS DE TRANSITO/`

## Requisitos

### Version de Python

Este proyecto esta preparado para **Python 3.14.4**.

### Dependencias principales

Se instalan desde `requirements.txt` e incluyen:

- `fastapi`
- `uvicorn[standard]`
- `requests`
- `msal`
- `python-dotenv`
- `SQLAlchemy`
- `psycopg2-binary`
- `pandas`
- `openpyxl`
- `pyngrok`

### Requisitos externos

- PostgreSQL accesible desde la maquina donde corre la aplicacion.
- Una aplicacion registrada en Azure AD con permisos para Microsoft Graph.
- Un sitio de SharePoint disponible en Microsoft 365.
- Ngrok solo si vas a usar `run_tunnel.py`.

## Variables de entorno

Crea un archivo `.env` en la raiz del proyecto con los valores necesarios.

```env
TENANT_ID=tu-tenant-id
CLIENT_ID=tu-client-id
CLIENT_SECRET=tu-client-secret
DOMINIO=tu-dominio.sharepoint.com
NOMBRE_SITIO=nombre-del-sitio
DATABASE_URL=postgresql://usuario:clave@localhost:5432/transito
HOST=0.0.0.0
NOTIFICATION_URL=https://tu-url-publica/webhook
GRAPH_RESOURCE=sites/tu-site-id/drive/root
GRAPH_CLIENT_STATE=SecretToken123
GRAPH_DB_FILE=subscription_info.json
NGROK_AUTHTOKEN=tu-token-de-ngrok
```

Notas importantes:

- `GRAPH_RESOURCE` es opcional si prefieres que el proyecto resuelva el site de SharePoint automaticamente.
- `NOTIFICATION_URL` debe ser una URL publica accesible por Microsoft Graph cuando uses webhooks.
- `DATABASE_URL` puede apuntar a cualquier PostgreSQL compatible.

## Instalacion en Windows

1. Instala Python 3.14.4 desde python.org y verifica que el launcher `py` esta disponible.
2. Abre una terminal en la carpeta del proyecto.
3. Crea y activa un entorno virtual:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

4. Instala las dependencias:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

5. Crea el archivo `.env` con tus credenciales y datos de conexion.
6. Verifica que PostgreSQL este ejecutandose y que la base de datos `transito` exista, o ajusta `DATABASE_URL`.

## Instalacion en Linux

1. Instala Python 3.14.4 en tu distribucion o con `pyenv`.
2. Abre una terminal en la carpeta del proyecto.
3. Crea y activa un entorno virtual:

```bash
python3.14 -m venv .venv
source .venv/bin/activate
```

4. Instala las dependencias:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

5. Crea el archivo `.env` con tus credenciales y datos de conexion.
6. Asegurate de tener PostgreSQL activo y accesible desde la aplicacion.

## Ejecucion

### Levantar la API principal

```bash
python main.py
```

La API queda disponible en `http://localhost:8000` por defecto.

### Levantar la API con tunel publico

```bash
python run_tunnel.py
```

Este comando abre ngrok en el puerto 8000 y publica una URL temporal para recibir webhooks.

### Gestionar suscripciones de Microsoft Graph

```bash
python graph_subscriptions.py
```

Este script crea, renueva o recrea la suscripcion y guarda su estado en `subscription_info.json`.

## Endpoints de la API

- `GET /` retorna informacion general del servicio.
- `GET /health` verifica que la API esta activa.
- `GET /registrar-documento` ejecuta una sincronizacion manual de SharePoint y persiste los resultados.
- `POST /webhook` recibe notificaciones de Microsoft Graph.

### Endpoints de trazabilidad (nuevos)

- `GET /api/archivos` — Lista archivos procesados. Soporta filtros opcionales:
	- `estado` (iniciado, procesando, completado, error)
	- `tipo` (MULTAS, DERECHOS, RECAUDO, CARTERA)
	- `organismo` (prefijo de ruta)
	- `skip`, `limit` (paginación)

- `GET /api/archivos/{archivo_id}/flujo` — Retorna detalle del flujo de procesamiento de un archivo:
	- `procesamiento`: estado general, duracion, filas procesadas/fallidas
	- `pasos`: lista ordenada de `ProcessingStepExecution` (nombre, orden, estado, duracion, registros procesados, mensajes de error)

- `GET /api/archivos/{archivo_id}/detalles` — Resumen con estadísticas por archivo (conteo de `recaudos`, `carteras`, duración y porcentaje de éxito).

El webhook también responde a la validación inicial de Microsoft Graph cuando la URL incluye `validationToken`.

## Que guarda en la base de datos

El proyecto crea y usa estas tablas:

- `archivos`: metadatos basicos de cada archivo detectado.
- `recaudos`: filas procesadas desde documentos de recaudo.
- `carteras`: filas procesadas desde documentos de cartera.
- `auditoria`: registro auxiliar por columna/archivo.

- `file_processing_logs`: registro principal por cada archivo (estado, duración, filas procesadas/fallidas).
- `processing_step_executions`: detalle de cada paso ejecutado por archivo (nombre, orden, estado, duración, registros procesados, errores).

## Detalles tecnicos

- La busqueda de archivos en SharePoint es recursiva y soporta paginacion.
- La lectura de archivos usa `pandas`, con soporte para Excel y CSV.
- El guardado se optimiza con inserciones por lote para evitar commits por fila.
- En carteras, el proyecto marca registros previos como inactivos antes de reimportar el mismo organismo y tipo.
- La aplicacion utiliza SQLAlchemy con PostgreSQL y sincroniza secuencias seriales al iniciar.

## Estructura del proyecto

- `main.py`: API FastAPI principal.
- `sharepoint.py`: acceso a Microsoft Graph, descarga y procesamiento de archivos.
- `models.py`: modelos ORM y conexion a PostgreSQL.
- `documents.py`: enum de rutas/categorias de documentos.
- `graph_subscriptions.py`: administracion de suscripciones de Microsoft Graph.
- `run_tunnel.py`: arranque con ngrok.
- `subscription_info.json`: estado local de la suscripcion activa.

## Flujo recomendado de uso

1. Configura las variables de entorno en `.env`.
2. Asegura la base de datos PostgreSQL.
3. Ejecuta `python main.py` para pruebas locales o `python run_tunnel.py` si necesitas webhook publico.
4. Si vas a usar notificaciones de Microsoft Graph, ejecuta `python graph_subscriptions.py` para crear la suscripcion.
5. Dispara `GET /registrar-documento` o espera a que llegue el webhook para sincronizar el contenido.

## Observaciones

- El proyecto asume que los nombres de columnas en los archivos de Excel o CSV coinciden con los esperados por `sharepoint.py`.
- Si cambian los encabezados de los archivos fuente, tambien debera ajustarse el mapeo en el codigo.
- Para produccion, conviene usar un servicio persistente, revisar el manejo de secretos y asegurar que la URL de notificacion permanezca accesible.
