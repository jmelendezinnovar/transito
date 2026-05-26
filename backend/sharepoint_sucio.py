import logging
import os
import tempfile
from typing import Dict, List, Optional

import msal
import pandas as pd
import requests
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from sqlalchemy.exc import SQLAlchemyError
from urllib3.util.retry import Retry
from datetime import datetime
from backend.models import Archivo, Session, session, Cartera, Ejecucion, Etapa, Recaudo
from backend.documents import WalletStatus, EjecucionEstado, EtapaNombre, EtapaEstado, Document, TransitoNombre
from backend.utils import extract_organismo, safe_text, safe_datetime, safe_int, safe_receipt
import time
from multiprocessing import Pool, cpu_count

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DOMINIO = os.getenv("DOMINIO")
NOMBRE_SITIO = os.getenv("NOMBRE_SITIO")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "100000"))
MAX_RECURSION_DEPTH = int(os.getenv("MAX_RECURSION_DEPTH", "20"))
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
CHUNK_SIZE = int(os.getenv("CHUNK_SIZE", "100000"))
PROGRESS_LOG_EVERY = int(os.getenv("PROGRESS_LOG_EVERY", "1000"))
NUM_WORKERS = min(6, max(2, cpu_count() - 1))
FILE_ID = (os.getenv("FILE_ID") or "").strip()

def build_sharepoint_view_url(web_url: Optional[str], archivo_id: str) -> str:
    base_url = (web_url or "").split("?")[0].rstrip("/")
    if not base_url:
        base_url = f"https://{DOMINIO}/:x:/s/{NOMBRE_SITIO}"
    return f"{base_url}/{archivo_id}"

def get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    token_result = app.acquire_token_for_client(SCOPE)
    return token_result["access_token"]

def get_graph_headers() -> Dict[str, str]:
    access_token = get_access_token()
    return {"Authorization": f"Bearer {access_token}"}

def get_sharepoint_site_and_drive(headers: Dict[str, str]) -> tuple[str, str]:
    site_url = f"https://graph.microsoft.com/v1.0/sites/{DOMINIO}:/sites/{NOMBRE_SITIO}:"
    resp = requests.get(site_url, headers=headers, timeout=30)
    if resp.status_code != 200:
        raise Exception(f"Error obteniendo site: {resp.status_code} - {resp.text}")

    site_id = resp.json()["id"]

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp_drives = requests.get(drives_url, headers=headers, timeout=30)
    if resp_drives.status_code != 200:
        raise Exception(f"Error obteniendo drives: {resp_drives.status_code} - {resp_drives.text}")

    drives_data = resp_drives.json().get("value", [])
    if not drives_data:
        raise Exception("No se encontraron drives en el sitio de SharePoint")

    return site_id, drives_data[0]["id"]

def download_with_retries(
    url: str,
    dest_path: str,
    headers: Dict[str, str] | None = None,
    timeout: tuple[int, int] = (10, 900),
    chunk_size: int = 8192,
    max_retries: int = 5,
):
    session = requests.Session()
    retries = Retry(
        total=max_retries,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"],
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    hdrs = headers or {}
    hdrs.setdefault("Accept-Encoding", "identity")
    hdrs.setdefault("Connection", "keep-alive")

    with session.get(url, headers=hdrs, stream=True, timeout=timeout) as response:
        response.raise_for_status()
        with open(dest_path, "wb") as file_handle:
            for chunk in response.iter_content(chunk_size=chunk_size):
                if chunk:
                    file_handle.write(chunk)

    return dest_path

def _list_children(drive_id: str, item_id: str, headers: Dict[str, str]) -> List[Dict]:
    if item_id == "root":
        children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
    else:
        children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"

    items: List[Dict] = []
    while children_url:
        response = requests.get(children_url, headers=headers, timeout=30)
        if response.status_code != 200:
            raise Exception(f"Error explorando SharePoint: {response.status_code} - {response.text}")

        payload = response.json()
        items.extend(payload.get("value", []))
        children_url = payload.get("@odata.nextLink")

    return items

def start_etapa(ejecucion: Ejecucion, nombre: EtapaNombre) -> Etapa:
    etapa = session.query(Etapa).filter_by(ejecucion_id=ejecucion.id, nombre=nombre.value).first()
    if etapa:
        etapa.estado = EtapaEstado.EN_PROCESO.value
        etapa.iniciado = datetime.now()
        etapa.finalizado = None
    else:
        etapa = Etapa(
            ejecucion_id=ejecucion.id,
            nombre=nombre.value,
            estado=EtapaEstado.EN_PROCESO.value,
            iniciado=datetime.now()
        )
        session.add(etapa)
    session.commit()
    return etapa

def complete_etapa(etapa: Etapa):
    etapa.estado = EtapaEstado.COMPLETADO.value
    etapa.finalizado = datetime.now()
    session.commit()

def fail_etapa(etapa: Etapa, error_state: str = EtapaEstado.FALLIDO.value):
    etapa.estado = error_state
    etapa.finalizado = datetime.now()
    session.commit()

def complete_ejecucion(ejecucion: Ejecucion):
    ejecucion.estado = EjecucionEstado.COMPLETADA.value
    ejecucion.finalizado = datetime.now()
    session.commit()

def fail_ejecucion(ejecucion: Ejecucion):
    ejecucion.estado = EjecucionEstado.FALLIDA.value
    ejecucion.finalizado = datetime.now()
    session.commit()

def obtener_tamaño_archivo(file_id: str, drive_id: str, headers: Dict) -> int:
    try:
        metadata_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}?$select=size"
        resp = requests.get(metadata_url, headers=headers, timeout=10)
        if resp.status_code == 200:
            return resp.json().get("size", 0)
        return 0
    except Exception as e:
        logger.warning(f"No se pudo obtener tamaño para {file_id}: {str(e)}")
        return 0

def _find_folder_by_name(
    drive_id: str,
    headers: Dict[str, str],
    folder_name: str,
    item_id: str = "root",
    path: str = "",
    depth: int = 0,
) -> Optional[Dict]:
    if depth > MAX_RECURSION_DEPTH:
        return None

    for item in _list_children(drive_id, item_id, headers):
        current_path = f"{path}/{item['name']}" if path else item["name"]
        if "folder" in item and item["name"].strip().upper() == folder_name.upper():
            return {
                "id": item["id"],
                "name": item["name"],
                "path": current_path,
                "webUrl": item.get("webUrl"),
            }

        if "folder" in item:
            nested = _find_folder_by_name(
                drive_id,
                headers,
                folder_name,
                item["id"],
                current_path,
                depth + 1,
            )
            if nested:
                return nested

    return None

def _collect_files_from_folder(
    drive_id: str,
    item_id: str,
    headers: Dict[str, str],
    path: str = "",
    depth: int = 0,
) -> List[Dict]:
    if depth > MAX_RECURSION_DEPTH:
        logger.warning(f"Límite de recursión alcanzado en: {path}")
        return []

    files: List[Dict] = []
    try:
        for item in _list_children(drive_id, item_id, headers):
            current_path = f"{path}/{item['name']}" if path else item["name"]
            if "file" in item and item["name"].lower().endswith((".xlsx", ".csv")):
                files.append(
                    {
                        "file_id": item["id"],
                        "file_name": item["name"],
                        "file_path": current_path,
                        "createdDateTime": item.get("createdDateTime"),
                        "webUrl": item.get("webUrl"),
                    }
                )
            elif "folder" in item:
                files.extend(
                    _collect_files_from_folder(
                        drive_id,
                        item["id"],
                        headers,
                        current_path,
                        depth + 1,
                    )
                )
    except Exception as error:
        logger.error(f"Error explorando {path}: {error}")

    return files

def _clean_numbered_prefix(value: str | None) -> str | None:
    if not value:
        return value

    text = str(value).strip()
    prefix, separator, rest = text.partition(".")
    if separator and prefix.strip().isdigit():
        cleaned = rest.strip()
    else:
        cleaned = text
    return cleaned or None


def _normalize_tabular_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.upper()
        .str.replace(" ", "_", regex=False)
    )
    return df


def _leer_excel_multihojas(
    tmp_path: str,
    file_path: str,
    header_rows: Optional[List[int]] = None,
) -> pd.DataFrame:
    workbook = pd.ExcelFile(tmp_path)
    sheet_names = workbook.sheet_names

    logger.info(
        "Hojas detectadas en %s: %s",
        file_path,
        ", ".join(sheet_names) if sheet_names else "sin hojas",
    )

    dataframes = []
    for sheet_index, sheet_name in enumerate(sheet_names):
        header_row = header_rows[sheet_index] if header_rows and sheet_index < len(header_rows) else 3
        sheet_df = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
        logger.info("Hoja %s: %s filas", sheet_name, len(sheet_df))
        dataframes.append(_normalize_tabular_columns(sheet_df))

    if not dataframes:
        return pd.DataFrame()

    return pd.concat(dataframes, ignore_index=True, sort=False)


def _leer_excel_villa_del_rosario_multas(tmp_path: str, file_path: str) -> pd.DataFrame:
    workbook = pd.ExcelFile(tmp_path)
    sheet_names = workbook.sheet_names

    logger.info(
        "Hojas detectadas en %s: %s",
        file_path,
        ", ".join(sheet_names) if sheet_names else "sin hojas",
    )

    parsed_rows: List[Dict] = []

    for sheet_name in sheet_names:
        raw_df = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
        logger.info("Hoja %s: %s filas crudas", sheet_name, len(raw_df))

        current_headers: Optional[List[str]] = None
        for row_index, row_values in enumerate(raw_df.itertuples(index=False, name=None), start=1):
            values = list(row_values)
            first_value = safe_text({"_0": values[0] if values else None}, "_0") or ""
            first_value = first_value.strip().upper()

            if not first_value and not any(values):
                continue

            if first_value.startswith("DETALLE RECAUDO") or first_value.startswith("DETALLE REC"):
                logger.info(
                    "Hoja %s fila %s omitida por titulo de seccion: %s",
                    sheet_name,
                    row_index,
                    first_value,
                )
                current_headers = None
                continue

            if first_value == "FUENTE":
                current_headers = []
                for column_index, cell_value in enumerate(values):
                    header_name = safe_text({"_0": cell_value}, "_0") or f"COL_{column_index}"
                    header_name = header_name.strip().upper().replace(" ", "_")
                    current_headers.append(header_name)

                logger.info(
                    "Encabezado detectado en hoja %s fila %s: %s",
                    sheet_name,
                    row_index,
                    current_headers,
                )
                continue

            if not current_headers:
                continue

            row_dict: Dict[str, object] = {}
            for column_index, header_name in enumerate(current_headers):
                cell_value = values[column_index] if column_index < len(values) else None
                row_dict[header_name] = cell_value

            parsed_rows.append(row_dict)

    if not parsed_rows:
        return pd.DataFrame()

    return pd.DataFrame(parsed_rows)


def _load_recaudos_multas_dataframe(file_path: str, tmp_path: str) -> pd.DataFrame:
    file_name = file_path.upper()

    if file_path.lower().endswith(".csv"):
        return _normalize_tabular_columns(pd.read_csv(tmp_path, low_memory=False))

    if not file_path.lower().endswith((".xlsx", ".xls")):
        raise ValueError(f"Formato de archivo no soportado: {file_path}")

    if TransitoNombre.VILLA_DEL_ROSARIO.value in file_name:
        return _leer_excel_villa_del_rosario_multas(tmp_path, file_path)

    if TransitoNombre.TURBACO.value in file_name:
        workbook = pd.ExcelFile(tmp_path)
        if not workbook.sheet_names:
            return pd.DataFrame()

        dataframes = []
        logger.info(
            "Hojas detectadas en %s: %s",
            file_path,
            ", ".join(workbook.sheet_names),
        )

        for sheet_index, sheet_name in enumerate(workbook.sheet_names):
            header_row = 3
            sheet_df = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
            logger.info("Hoja %s: %s filas", sheet_name, len(sheet_df))
            dataframes.append(_normalize_tabular_columns(sheet_df))

        return pd.concat(dataframes, ignore_index=True, sort=False)

    return _leer_excel_multihojas(tmp_path, file_path)

def contar_filas_archivo(file_path: str, df: pd.DataFrame) -> int:
    logger.info(f"Contando filas en {file_path}...")
    try:
        if df is not None and not df.empty:
            return len(df)
        return 0
    except Exception as e:
        logger.warning(f"Error contando filas de {file_path}: {str(e)}")
        return 0

def _extraer_anio_valido(value) -> Optional[int]:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, (int, float)):
        try:
            year = int(value)
            return year if 1900 <= year <= 2100 else None
        except Exception:
            return None

    text = str(value).strip()
    if not text:
        return None

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if not pd.isna(parsed):
        return int(parsed.year)

    digits = "".join(char for char in text if char.isdigit())
    if len(digits) >= 4:
        try:
            year = int(digits[-4:])
            return year if 1900 <= year <= 2100 else None
        except Exception:
            return None

    return None

def _leer_tabular_file(tmp_path: str, file_path: str) -> pd.DataFrame:
    archivo = file_path.lower()
    if archivo.endswith((".xlsx", ".xls")):
        return _leer_excel_multihojas(tmp_path, file_path)
    if archivo.endswith(".csv"):
        return pd.read_csv(tmp_path, low_memory=False)
    raise ValueError(f"Formato no soportado: {file_path}")
    
def get_datos_sucios_files() -> Dict:
    headers = get_graph_headers()
    site_id, drive_id = get_sharepoint_site_and_drive(headers)

    folder = {
        "id": "root",
        "name": "root",
        "path": "",
        "webUrl": None,
    }

    files = _collect_files_from_folder(drive_id, folder["id"], headers, folder["path"])
    if not files:
        logger.warning("No se encontraron archivos en SharePoint")
    else:
        logger.info("Se procesaran todos los archivos encontrados en SharePoint")

    if FILE_ID:
        total_files = len(files)
        files = [file_info for file_info in files if file_info.get("file_id") == FILE_ID]
        logger.info(
            "Filtro file_id aplicado (%s): %s de %s archivo(s) coinciden",
            FILE_ID,
            len(files),
            total_files,
        )
        if not files:
            logger.warning("No se encontraron archivos con file_id=%s", FILE_ID)

    files.sort(key=lambda file_info: file_info.get("createdDateTime", ""))

    for index, file_info in enumerate(files, 1):
        logger.info(f"{index}. {file_info['file_name']} (creado: {file_info.get('createdDateTime', 'N/A')})")

    return {
        "files": files,
        "headers": headers,
        "site_id": site_id,
        "drive_id": drive_id,
        "folder": folder,
    }

def get_or_create_ejecucion(archivo_id: str, estado: str) -> Ejecucion:
    ejecucion = session.query(Ejecucion).filter_by(archivo_id=archivo_id).first()
    if ejecucion:
        ejecucion.estado = estado
        if estado == EjecucionEstado.EN_PROCESO.value and ejecucion.finalizado is not None:
            ejecucion.finalizado = None
        session.commit()
        return ejecucion

    ejecucion = Ejecucion(archivo_id=archivo_id, estado=estado)
    session.add(ejecucion)
    session.commit()
    return ejecucion

def descargar_y_parsear_excel(file_id: str, file_path: str, headers: Dict, drive_id: str) -> Dict:
    try:
        # Generar URL de descarga con token fresco (no caduca como downloadUrl)
        # Esta URL usa Graph API que siempre funciona con token válido
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        start_time = time.time()

        # Descargar en streaming a archivo temporal con reintentos
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmp:
            tmp_path = tmp.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando {file_path}: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return None

        # Leer archivo desde disco (evita cargar contenido entero en memoria)
        try:
            df = _leer_tabular_file(tmp_path, file_path)
        except Exception as error:
            logger.error(str(error))
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return None
        
        # Normalizar columnas
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_", regex=False)
        )
        
        elapsed = time.time() - start_time
        logger.info(f"✓ Archivo parseado: {file_path} ({len(df):,} filas en {elapsed:.2f}s)")
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        
        return {"df": df, "file_path": file_path, "file_id": file_id}
        
    except Exception as e:
        logger.error(f"Error descargando {file_path}: {str(e)}")
        return None

def save_datos_sucios_to_database(excel_files, headers, site_id, drive_id, batch_size: int = BATCH_SIZE):
    if not session:
        raise Exception("Base de datos no disponible")
    
    results = {
        "guardados": [],
        "actualizados": [],
        "errores": [],
        "procesados": []
    }
    
    start_total = time.time()
    archivos_cartera = []
    archivos_recaudo = []
    
    for file_info in excel_files:
        ejecucion = None
        etapa_extraccion = None
        etapa_limpieza = None
        etapa_guardado = None
        try:
            existing = session.query(Archivo).filter_by(archivo_id=file_info["file_id"]).first()
            if existing:
                logger.info(f"Archivo ya registrado, se omite reprocesamiento: {file_info['file_name']} ({file_info['file_id']})")
                results["actualizados"].append(file_info["file_path"])
                continue

            file_download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_info['file_id']}/content"
            filas_count = 0
            file_view_url = build_sharepoint_view_url(file_info.get("webUrl"), file_info["file_id"])
            
            try:
                # Descargar en streaming a temp file y contar filas
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_info["file_path"])[1]) as tmpf:
                    tmp_path = tmpf.name

                try:
                    download_with_retries(file_download_url, tmp_path, headers=headers, timeout=(10, 900))
                    try:
                        if Document.RECAUDO_MULTAS.value in file_info["file_path"]:
                            df = _load_recaudos_multas_dataframe(file_info["file_path"], tmp_path)
                        else:
                            df = _leer_tabular_file(tmp_path, file_info["file_path"])
                    except Exception:
                        df = None
                    filas_count = contar_filas_archivo(file_info["file_path"], df)
                finally:
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
            except Exception as e:
                logger.warning(f"Error descargando {file_info['file_path']} para contar filas: {str(e)}")
                filas_count = 0

            file_record = Archivo(
                archivo_id=file_info["file_id"],
                nombre=file_info["file_name"],
                ruta=file_info["file_path"],
                url=file_view_url,
                filas=filas_count
            )
            session.add(file_record)
            session.commit()
            results["guardados"].append(file_info["file_path"])

            # PASO 2: Ahora sí crear Ejecucion (el Archivo ya existe)
            ejecucion = get_or_create_ejecucion(file_info["file_id"], EjecucionEstado.EN_PROCESO.value)

            etapa_extraccion = start_etapa(ejecucion, EtapaNombre.EXTRACCION)
            complete_etapa(etapa_extraccion)

            etapa_limpieza = start_etapa(ejecucion, EtapaNombre.LIMPIEZA)
            complete_etapa(etapa_limpieza)
            
            logger.info(f"Archivo registrado: {file_info['file_name']} ({filas_count} filas)")

            complete_etapa(etapa_limpieza)
            
            # Clasificar archivos por tipo
            if Document.CARTERA_MULTAS.value in file_info["file_path"] or Document.CARTERA_DERECHOS_DE_TRANSITO.value in file_info["file_path"]:
                etapa_guardado = start_etapa(ejecucion, EtapaNombre.GUARDADO)
                archivos_cartera.append({
                    "file_id": file_info["file_id"],
                    "file_path": file_info["file_path"],
                    "headers": headers,
                    "site_id": site_id,
                    "drive_id": drive_id,
                    "ejecucion_id": ejecucion.id,
                    "etapa_guardado_id": etapa_guardado.id
                })
            elif Document.RECAUDO_MULTAS.value in file_info["file_path"] or Document.RECAUDO_DERECHOS_DE_TRANSITO.value in file_info["file_path"]:
                etapa_guardado = start_etapa(ejecucion, EtapaNombre.GUARDADO)
                archivos_recaudo.append({
                    "file_id": file_info["file_id"],
                    "file_path": file_info["file_path"],
                    "headers": headers,
                    "site_id": site_id,
                    "drive_id": drive_id,
                    "ejecucion_id": ejecucion.id,
                    "etapa_guardado_id": etapa_guardado.id
                })
            else:
                complete_ejecucion(ejecucion)
                    
        except Exception as e:
            session.rollback()
            try:
                if etapa_extraccion:
                    fail_etapa(etapa_extraccion)
                if etapa_limpieza:
                    fail_etapa(etapa_limpieza)
                if etapa_guardado:
                    fail_etapa(etapa_guardado)
                if ejecucion:
                    fail_ejecucion(ejecucion)
            except Exception:
                pass
            results["errores"].append({
                "archivo": file_info["file_path"],
                "error": str(e)
            })
    
    # PASO 1.5: Ordenar archivos de cartera por tamaño (pequeños primero)
    logger.info(f"📦 Obteniendo tamaños de {len(archivos_cartera)} archivos de cartera...")
    archivos_con_tamaño = []
    for cartera_item in archivos_cartera:
        file_id = cartera_item["file_id"]
        file_path = cartera_item["file_path"]
        hdrs = cartera_item["headers"]
        s_id = cartera_item["site_id"]
        d_id = cartera_item["drive_id"]
        tamaño = obtener_tamaño_archivo(file_id, d_id, hdrs)
        archivos_con_tamaño.append((cartera_item, tamaño))
    
    # Ordenar por tamaño ascendente (pequeños primero para paralelismo óptimo)
    archivos_con_tamaño.sort(key=lambda x: x[1])
    archivos_cartera = [x[0] for x in archivos_con_tamaño]
    
    # Precarga de códigos existentes para cada organismo/tipo (evita N queries después)
    logger.info("🔍 Precargando códigos existentes en BD...")
    codigos_existentes_map = {}  # {(organismo, tipo_cartera): set(códigos)}
    
    # PASO 2: Procesar CARTERAS en paralelo (2-6 workers dinámicos)
    logger.info(f"🚀 Iniciando procesamiento paralelo de {len(archivos_cartera)} archivos de cartera...")
    if archivos_cartera:
        try:
            with Pool(processes=NUM_WORKERS) as pool:
                cartera_results = pool.map(
                    procesar_archivo_cartera,
                    [
                        (item["file_id"], item["file_path"], item["headers"], item["site_id"], item["drive_id"])
                        for item in archivos_cartera
                    ]
                )
                
            for cartera_item, resultado in zip(archivos_cartera, cartera_results):
                try:
                    etapa_guardado = session.query(Etapa).filter_by(id=cartera_item["etapa_guardado_id"]).first()
                    ejecucion = session.query(Ejecucion).filter_by(id=cartera_item["ejecucion_id"]).first()
                    if resultado:
                        if etapa_guardado:
                            complete_etapa(etapa_guardado)
                        if ejecucion:
                            complete_ejecucion(ejecucion)
                    else:
                        if etapa_guardado:
                            fail_etapa(etapa_guardado)
                        if ejecucion:
                            fail_ejecucion(ejecucion)
                except Exception as status_error:
                    logger.error(f"Error actualizando estado de cartera: {str(status_error)}")

                if resultado:
                    results["procesados"].append({
                        "archivo": resultado.get("archivo_path", "desconocido"),
                        "filas_procesadas": resultado["filas_procesadas"],
                        "guardadas": resultado["guardadas"],
                        "errores": resultado["errores"]
                    })
        except Exception as e:
            logger.error(f"Error en parallelización de carteras: {str(e)}")
    
    # PASO 3: Procesar RECAUDOS secuencialmente (más eficiente por su tamaño)
    logger.info(f"📊 Procesando {len(archivos_recaudo)} archivos de recaudos...")
    for file_info in archivos_recaudo:
        try:
            etapa_guardado = session.query(Etapa).filter_by(id=file_info["etapa_guardado_id"]).first()
            ejecucion = session.query(Ejecucion).filter_by(id=file_info["ejecucion_id"]).first()

            for enum_item in Document:
                if enum_item.value in file_info["file_path"]:
                    if enum_item == Document.RECAUDO_MULTAS:
                        resultado = get_recaudos_multas(
                            file_info["file_id"],
                            file_info["file_path"],
                            file_info["headers"],
                            file_info["site_id"],
                            file_info["drive_id"]
                        )
                    elif enum_item == Document.RECAUDO_DERECHOS_DE_TRANSITO:
                        resultado = get_recaudos_derechos(
                            file_info["file_id"],
                            file_info["file_path"],
                            file_info["headers"],
                            file_info["site_id"],
                            file_info["drive_id"]
                        )
                    else:
                        break
                    
                    if etapa_guardado:
                        complete_etapa(etapa_guardado)
                    if ejecucion:
                        complete_ejecucion(ejecucion)

                    results["procesados"].append({
                        "archivo": file_info["file_path"],
                        "tipo": enum_item.name,
                        "filas_procesadas": resultado["filas_procesadas"],
                        "guardadas": resultado["guardadas"],
                        "errores": resultado["errores"]
                    })
                    break
                    
        except Exception as e:
            session.rollback()
            try:
                if etapa_guardado:
                    fail_etapa(etapa_guardado)
                if ejecucion:
                    fail_ejecucion(ejecucion)
            except Exception:
                pass
            results["errores"].append({
                "archivo": file_info["file_path"],
                "error": str(e)
            })
    
    elapsed_total = time.time() - start_total
    logger.info(f"✅ Proceso completado en {elapsed_total:.2f}s (archivos: {len(excel_files)})")
    
    return results

def batch_insert_records(records: List, batch_size: int = BATCH_SIZE):
    """Inserta registros en lotes masivos (50k/commit) con optimización de BD"""
    if not records:
        return
    
    start_time = time.time()
    total_inserted = 0
    
    def model_to_mapping(obj):
        # Convierte una instancia ORM a mapping plano, excluyendo metadatos de SQLAlchemy
        d = {}
        for k, v in obj.__dict__.items():
            if k == '_sa_instance_state':
                continue
            d[k] = v
        return d

    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        db_session = Session()
        try:
            mappings = [model_to_mapping(r) for r in batch]
            cls = batch[0].__class__
            db_session.bulk_insert_mappings(cls, mappings)
            db_session.commit()
            total_inserted += len(batch)
            db_session.close()
            logger.debug(f"✓ Batch insertado: {len(batch):,} registros")
        except SQLAlchemyError as e:
            try:
                db_session.rollback()
            except Exception:
                pass
            logger.error(f"Error en batch insert: {str(e)}")
            try:
                db_session.close()
            except Exception:
                pass
    
    elapsed = time.time() - start_time
    logger.info(f"Total insertados: {total_inserted:,} registros en {elapsed:.2f}s ({total_inserted/elapsed:,.0f} filas/seg)")

def procesar_archivo_cartera(args_tuple):
    file_id, file_path, headers, site_id, drive_id = args_tuple
    
    # Descargar y parsear
    data = descargar_y_parsear_excel(file_id, file_path, headers, drive_id)
    if not data:
        return None
    
    df = data["df"]
    organismo = extract_organismo(file_path)
    
    # Determinar tipo y procesar
    resultado = None
    if Document.CARTERA_MULTAS.value in file_path:
        resultado = get_carteras_multas(file_id, file_path, headers, site_id, drive_id)
    elif Document.CARTERA_DERECHOS_DE_TRANSITO.value in file_path:
        resultado = get_carteras_derechos(file_id, file_path, headers, site_id, drive_id)
    
    if resultado:
        resultado["archivo_path"] = file_path
    
    return resultado

def get_carteras_multas(file_id, file_path, headers, site_id, drive_id):
    """Procesa archivo de carteras de MULTAS (CSV/XLSX) en chunks y lo guarda en la tabla `Cartera`."""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmpf:
            tmp_path = tmpf.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando archivo: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        # Leer archivo
        try:
            df = _load_recaudos_multas_dataframe(file_path, tmp_path)
        except Exception as error:
            logger.error(str(error))
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        if TransitoNombre.TURBACO.value in file_path.upper() and file_path.lower().endswith((".xlsx", ".xls")):
            total_interno = int((df["ORIGEN_RECAUDO"] == "RECAUDO INTERNO").sum()) if "ORIGEN_RECAUDO" in df.columns else 0
            total_externo = int((df["ORIGEN_RECAUDO"] == "RECAUDO EXTERNO").sum()) if "ORIGEN_RECAUDO" in df.columns else 0
            logger.info(
                "Excel TURBACO combinado: total=%s, interno=%s, externo=%s",
                len(df),
                total_interno,
                total_externo,
            )
        elif file_path.lower().endswith(".xlsx"):
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            logger.info(f"CSV leído: {len(df)} filas")

        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0

        logger.info(
            "Columnas detectadas en multas: %s",
            ", ".join(df.columns.tolist()[:60])
        )

        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            chunk_len = len(df_chunk)
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + chunk_len} ({chunk_len} filas)")

            db_session = Session()
            insert_mappings = []
            rows_in_chunk = 0
            descartes_apellido = 0
            descartes_estado = 0
            descartes_anio = 0
            muestra_descartes = []
            registros_descartados = []

            def registrar_descarte(motivo: str, row_dict: Dict, estado: str, anio: object) -> None:
                nonlocal registros_descartados
                detalle = {
                    "motivo": motivo,
                    "codigo": str(row_dict.get('CODIGO', '') or ''),
                    "nombre": safe_text(row_dict, 'NOMBRE_INFRACTOR'),
                    "apellido": safe_text(row_dict, 'APELLIDO'),
                    "estado": estado,
                    "anio": anio,
                    "comparendo": safe_text(row_dict, 'NUMERO_COMPARENDO') or safe_text(row_dict, 'COMPARENDO'),
                    "placa": safe_text(row_dict, 'PLACA'),
                    "fecha": safe_text(row_dict, 'FECHA_COMPARENDO') or safe_text(row_dict, 'FECHA'),
                }
                registros_descartados.append(detalle)
                logger.warning("Fila multa no registrada: %s", detalle)

            chunk_start_time = time.time()
            for idx, row in enumerate(df_chunk.itertuples(index=False, name='Record')):
                try:
                    rows_in_chunk += 1
                    row_dict = row._asdict()
                    codigo = str(row_dict.get('CODIGO', '') or '')
                    apellido = safe_text(row_dict, 'APELLIDO')
                    # Eliminar filas con apellido indeterminado
                    if apellido and apellido.strip().upper() == 'INDETERMINADO':
                        descartes_apellido += 1
                        if len(muestra_descartes) < 5:
                            muestra_descartes.append({
                                "motivo": "apellido",
                                "codigo": codigo,
                                "apellido": apellido,
                                "estado": safe_text(row_dict, 'ESTADO_COMPARENDO') or safe_text(row_dict, 'ESTADO_CARTERA') or safe_text(row_dict, 'ESTADO') or '',
                                "anio": safe_text(row_dict, 'AÑO_COMPARENDO') or safe_text(row_dict, 'FECHA_COMPARENDO') or safe_text(row_dict, 'FECHA') or ''
                            })
                        registrar_descarte(
                            "apellido",
                            row_dict,
                            safe_text(row_dict, 'ESTADO_COMPARENDO') or safe_text(row_dict, 'ESTADO_CARTERA') or safe_text(row_dict, 'ESTADO') or '',
                            safe_text(row_dict, 'AÑO_COMPARENDO') or safe_text(row_dict, 'FECHA_COMPARENDO') or safe_text(row_dict, 'FECHA') or ''
                        )
                        continue

                    # Filtrar por estado del comparendo (sancionado, vigente, acuerdo de pago)
                    estado_comp = (
                        safe_text(row_dict, 'ESTADO_COMPARENDO') or
                        safe_text(row_dict, 'ESTADO_CARTERA') or
                        safe_text(row_dict, 'ESTADO') or
                        ''
                    ).upper()
                    if not any(k in estado_comp for k in ('SANCION', 'VIGENTE', 'ACUERDO DE PAGO', 'ACUERDO')):
                        descartes_estado += 1
                        if len(muestra_descartes) < 5:
                            muestra_descartes.append({
                                "motivo": "estado",
                                "codigo": codigo,
                                "apellido": apellido,
                                "estado": estado_comp,
                                "anio": safe_text(row_dict, 'AÑO_COMPARENDO') or safe_text(row_dict, 'FECHA_COMPARENDO') or safe_text(row_dict, 'FECHA') or ''
                            })
                        registrar_descarte(
                            "estado",
                            row_dict,
                            estado_comp,
                            safe_text(row_dict, 'AÑO_COMPARENDO') or safe_text(row_dict, 'FECHA_COMPARENDO') or safe_text(row_dict, 'FECHA') or ''
                        )
                        continue

                    # Filtrar por año de comparendo (2015 en adelante)
                    year_source = (
                        safe_text(row_dict, "TO_CHAR(FECHA_COMPARENDO,'YYYY')") or
                        safe_text(row_dict, 'AÑO_COMPARENDO') or
                        safe_text(row_dict, 'FECHA_COMPARENDO') or
                        safe_text(row_dict, 'FECHA') or
                        ''
                    )
                    year_int = _extraer_anio_valido(year_source)
                    if year_int is None or year_int < 2015:
                        descartes_anio += 1
                        if len(muestra_descartes) < 5:
                            muestra_descartes.append({
                                "motivo": "anio",
                                "codigo": codigo,
                                "apellido": apellido,
                                "estado": estado_comp,
                                "anio": year_source
                            })
                        registrar_descarte(
                            "anio",
                            row_dict,
                            estado_comp,
                            year_source
                        )
                        continue

                    # Limpiar formatos de fecha (remover puntos)
                    fecha_raw = safe_text(row_dict, 'FECHA_COMPARENDO') or ''
                    fecha_clean = fecha_raw.replace('.', '') if fecha_raw else None
                    notif_raw = safe_text(row_dict, 'NOTIF_FECHA') or ''
                    notif_clean = notif_raw.replace('.', '') if notif_raw else None

                    insert_mappings.append({
                        'archivo_id': file_id,
                        'organismo': organismo,
                        'codigo': codigo,
                        'tipo_cartera': 'MULTAS',
                        'estado_cartera_final': WalletStatus.ACTIVE.value,
                        'fecha': fecha_clean,
                        'tipo_comparendo': safe_text(row_dict, 'TIPO_COMPARENDO'),
                        'clase': safe_text(row_dict, 'CLASE'),
                        'servicio': safe_text(row_dict, 'SERVICIO'),
                        'valor_inicial_cartera': safe_text(row_dict, 'CART_VALOR_INICIAL'),
                        'numero_referencia_cartera': safe_text(row_dict, 'CART_NRO_REFERENCIA'),
                        'estado_cartera': safe_text(row_dict, 'ESTADO_CARTERA'),
                        'fecha_inicio_cartera': safe_text(row_dict, 'CART_FECHA_INGRESO'),
                        'estado_gestion': safe_text(row_dict, 'ESTADO_GESTION'),
                        'capital': safe_text(row_dict, 'CAPITAL'),
                        'total': safe_text(row_dict, 'TOTAL'),
                        'resolucion_fecha': safe_text(row_dict, 'RESOLUCION_FECHA'),
                        'intereses': safe_text(row_dict, 'INTERESES'),
                        'placa': safe_text(row_dict, 'PLACA'),
                        'tipo_identificacion': safe_text(row_dict, 'TIPO_IDENTIFICACION'),
                        'numero_identificacion': safe_text(row_dict, 'NUMERO_IDENTIFICACION'),
                        'nombre_infractor': safe_text(row_dict, 'NOMBRE_INFRACTOR'),
                        'numero_comparendo': safe_text(row_dict, 'NUMERO_COMPARENDO'),
                        'estado_comparendo': estado_comp,
                        'infraccion': safe_text(row_dict, 'INFRACCION'),
                        'resolucion_sancion': safe_text(row_dict, 'RESOLUCION_SANCION'),
                        'mandamiento_de_pago': safe_text(row_dict, 'MANDAMIENTO_DE_PAGO'),
                        'fecha_mandamiento_de_pago': safe_text(row_dict, 'FECHA_MANDAMIENTO'),
                        'fecha_de_notificacion': notif_clean,
                        'clase_vehiculo': safe_text(row_dict, 'CLASE_VEHICULO'),
                        'año_comparendo': str(year_int),
                        'ciudad': safe_text(row_dict, 'NOMBRE_CIUDAD'),
                        'direccion': safe_text(row_dict, 'DIR_DIRECCION'),
                        'telefono': safe_text(row_dict, 'DIR_TELEFONO'),
                        'movil': safe_text(row_dict, 'MOVIL'),
                        'email': safe_text(row_dict, 'EMAIL')
                    })

                    total_guardadas += 1

                    if rows_in_chunk % PROGRESS_LOG_EVERY == 0:
                        elapsed = time.time() - chunk_start_time
                        logger.info(f"Progreso chunk: {rows_in_chunk}/{chunk_len} filas procesadas en este chunk, total guardadas: {total_guardadas} - {elapsed:.1f}s")

                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")

            try:
                if insert_mappings:
                    for j in range(0, len(insert_mappings), BATCH_SIZE):
                        sub = insert_mappings[j:j + BATCH_SIZE]
                        db_session.bulk_insert_mappings(Cartera, sub)
                db_session.commit()
            except SQLAlchemyError as e:
                logger.error(f"Error aplicando bulk updates/inserts: {e}")
                try:
                    db_session.rollback()
                except Exception:
                    pass
            finally:
                try:
                    db_session.close()
                except Exception:
                    pass

            logger.info(
                "Resumen chunk multas %s-%s: leidas=%s, guardadas=%s, descartadas_apellido=%s, descartadas_estado=%s, descartadas_anio=%s",
                chunk_start,
                chunk_start + chunk_len,
                rows_in_chunk,
                len(insert_mappings),
                descartes_apellido,
                descartes_estado,
                descartes_anio,
            )
            if muestra_descartes:
                logger.info("Muestra de descartes multas: %s", muestra_descartes)
            if registros_descartados:
                logger.info("Total de filas no registradas en este chunk: %s", len(registros_descartados))

        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")

        try:
            os.remove(tmp_path)
        except Exception:
            pass

        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

def get_carteras_derechos(file_id, file_path, headers, site_id, drive_id):
    """Procesa archivo de carteras de DERECHOS DE TRANSITO (CSV/XLSX) en chunks y lo guarda en la tabla `Cartera`."""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmpf:
            tmp_path = tmpf.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando archivo: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = _leer_excel_multihojas(tmp_path, file_path)
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path, low_memory=False)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0

        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            chunk_len = len(df_chunk)
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + chunk_len} ({chunk_len} filas)")

            db_session = Session()
            insert_mappings = []
            rows_in_chunk = 0
            chunk_start_time = time.time()
            for idx, row in enumerate(df_chunk.itertuples(index=False, name='Record')):
                try:
                    rows_in_chunk += 1
                    row_dict = row._asdict()
                    codigo = str(row_dict.get('CODIGO', '') or '')
                    nombre = safe_text(row_dict, 'NOMBRE')
                    apellido = safe_text(row_dict, 'APELLIDO')
                    # Descartar filas con apellido indeterminado o nombres que indiquen persona indeterminada
                    if apellido and (
                        apellido.strip().upper() == 'INDETERMINADO' or 
                        apellido.strip().upper() == 'DESCONOCIDO' or 
                        apellido.strip().upper() == 'NO IDENTIFICADO' or
                        apellido.strip().upper() == 'INDETERMINADA'
                    ):
                        continue

                    if nombre and (
                        nombre.strip().upper() == 'INDETERMINADO' or 
                        nombre.strip().upper() == 'DESCONOCIDO' or 
                        nombre.strip().upper() == 'NO IDENTIFICADO' or
                        nombre.strip().upper() == 'INDETERMINADA'
                    ):
                        continue

                    insert_mappings.append({
                            'archivo_id': file_id,
                            'organismo': organismo,
                            'codigo': codigo,
                            'tipo_cartera': 'DERECHOS DE TRANSITO',
                            'estado_cartera_final': WalletStatus.ACTIVE.value,
                            'fecha': safe_text(row_dict, 'FECHA_CARTERA') or safe_text(row_dict, 'FECHA CARTERA'),
                            'valor_inicial_cartera': safe_text(row_dict, 'CARTERA_VALOR_INICIAL'),
                            'numero_referencia_cartera': safe_text(row_dict, 'REFERENCIA'),
                            'estado_cartera': safe_text(row_dict, 'ESTADO_CARTERA') or safe_text(row_dict, 'ESTADO CARTERA'),
                            'estado_gestion': safe_text(row_dict, 'ESTADO_GESTION'),
                            'capital': safe_text(row_dict, 'CAPITAL'),
                            'total': safe_text(row_dict, 'TOTAL'),
                            'fecha_inicio_cartera': safe_text(row_dict, 'CARTERA_FECHA_DE_INGRESO'),
                            'intereses': safe_text(row_dict, 'INTERESES'),
                            'placa': safe_text(row_dict, 'PLACA'),
                            'clase': safe_text(row_dict, 'CLASE'),
                            'servicio': safe_text(row_dict, 'SERVICIO'),
                            'tipo_identificacion': safe_text(row_dict, 'TIPO_IDENTIFICACION'),
                            'numero_identificacion': safe_text(row_dict, 'NUMERO_IDENTIFICACION'),
                            'nombre_infractor': (
                                ' '.join(
                                    p for p in (
                                        nombre,
                                        apellido
                                    ) if p
                                ) or None
                            ),
                            'email': safe_text(row_dict, 'EMAIL'),
                            'modelo': safe_text(row_dict, 'MODELO'),
                            'telefono': safe_text(row_dict, 'TELEFONO_MOVIL'),
                            'fecha_propietario': safe_text(row_dict, 'FECHA_PROPIETARIO'),
                            'filtro_coactivo': safe_text(row_dict, 'FILTRO_COACTIVO'),
                            'clase_vehiculo': safe_text(row_dict, 'CLASE_VEHICULO'),
                            'direccion': safe_text(row_dict, 'DIRECCION'),
                            'mp_resolucion': safe_text(row_dict, 'MP_RESOLUCION'),
                            'fecha_mp': safe_text(row_dict, 'FECHA_MP')
                        })

                    total_guardadas += 1

                    if rows_in_chunk % PROGRESS_LOG_EVERY == 0:
                        elapsed = time.time() - chunk_start_time
                        logger.info(f"Progreso chunk: {rows_in_chunk}/{chunk_len} filas procesadas en este chunk, total guardadas: {total_guardadas} - {elapsed:.1f}s")

                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")

            try:
                if insert_mappings:
                    for j in range(0, len(insert_mappings), BATCH_SIZE):
                        sub = insert_mappings[j:j + BATCH_SIZE]
                        db_session.bulk_insert_mappings(Cartera, sub)
                db_session.commit()
            except SQLAlchemyError as e:
                logger.error(f"Error aplicando bulk updates/inserts (derechos): {e}")
                try:
                    db_session.rollback()
                except Exception:
                    pass
            finally:
                try:
                    db_session.close()
                except Exception:
                    pass

        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")

        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

def get_recaudos_multas(file_id, file_path, headers, site_id, drive_id):
    """
    Lee un archivo de recaudos de multas (Excel o CSV) en CHUNKS.
    Usa itertuples() en lugar de iterrows() (100x más rápido).
    Batch insert de 5000 filas (40 commits en lugar de 200k).
    """
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

        # Descargar en streaming a temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmpf:
            tmp_path = tmpf.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando archivo: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = _load_recaudos_multas_dataframe(file_path, tmp_path)
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path, low_memory=False)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0

        logger.info(
            "Encabezados detectados en multas (%s): %s",
            file_path,
            list(df.columns),
        )
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            chunk_len = len(df_chunk)
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + chunk_len} ({chunk_len} filas)")

            records = []
            rows_in_chunk = 0
            chunk_start_time = time.time()
            for row_dict in df_chunk.to_dict(orient="records"):
                try:
                    rows_in_chunk += 1
                    if chunk_start == 0 and rows_in_chunk <= 5:
                        logger.info("row_dict muestra #%s en multas: %s", rows_in_chunk, row_dict)

                    fuente = safe_text(row_dict, "FUENTE") or safe_text(row_dict, "ORIGEN_RECAUDO")
                    
                    recaudo = Recaudo(
                        archivo_id=file_id,
                        organismo=organismo,
                        tipo_recaudo="MULTAS",
                        origen_recaudo=safe_text(row_dict, "ORIGEN_RECAUDO"),
                        fuente=fuente,
                        fecha_pago=safe_datetime(row_dict, "FECHA_PAGO") or safe_datetime(row_dict, "_1"),
                        recibo=safe_receipt(row_dict, "RECIBO"),
                        valor_recibido=safe_text(row_dict, "VALOR_RECIBO"),
                        tipo_documento=safe_text(row_dict, "TIPO_DOCUMENTO"),
                        identificacion=safe_text(row_dict, "IDENTIFICACION"),
                        nombre=safe_text(row_dict, "NOMBRE"),
                        vehiculo_placa=safe_text(row_dict, "VEHI_PLACA"),
                        comparendo=safe_text(row_dict, "COMPARENDO") or safe_text(row_dict, "COMP_NUMERO"),
                        fecha_comparendo=safe_datetime(row_dict, "COMP_FECHA") or safe_datetime(row_dict, "FECHA_COMPARENDO"),
                        año_comparendo=safe_text(row_dict, "AÑO_COMPARENDO") or safe_text(row_dict, "ANIO_COMPARENDO"),
                        prescripcion=safe_text(row_dict, "PRESCRIPCION"),
                        tipo_comparendo=_clean_numbered_prefix(safe_text(row_dict, "TIPO_COMPARENDO")),
                        clase_vehiculo=safe_text(row_dict, "CLASE_VEHICULO"),
                        tipo=safe_text(row_dict, "TIPO"),
                        servicio_vehiculo=safe_text(row_dict, "SERVICIO_VEHICULO"),
                        valor_pagado=safe_text(row_dict, "VALOR_PAGADO"),
                        fecha_distribucion=safe_datetime(row_dict, "DISTRI_FECHA"),
                        resolucion_mp=safe_text(row_dict, "RESOLUCION_MP"),
                        valor_inicial_cargado=(
                            safe_text(row_dict, "VALOR_INICIAL_CAR") or
                            safe_text(row_dict, "VALOR_CARTERA") or
                            safe_text(row_dict, "VALOR_CAR")
                        ),
                        concepto=(
                            safe_text(row_dict, "CONCEPTO") or
                            safe_text(row_dict, "CONCEPTO_PRINCIPAL") or
                            safe_text(row_dict, "DETALLE") or
                            safe_text(row_dict, "DESCRIPCION")
                        ),
                        estado_cartera=safe_text(row_dict, "ESTADO_CARTERA"),
                        concepto_principal=safe_text(row_dict, "CONCEPTO_PRINCIPAL"),
                        gestion=safe_text(row_dict, "GESTION") or safe_text(row_dict, "GESTIÓN"),
                        descuento_cartera=safe_text(row_dict, "DESCUENTO_CARTERA"),
                        descuento_de_intereses=safe_text(row_dict, "DES_INTERESES"),
                        cantidad_de_descuento_cartera=safe_int(row_dict, "CANT_DESTO_CARTERA"),
                        cantidad_de_descuento_de_intereses=safe_int(row_dict, "CANT_DES_INTERESES"),
                        resolucion_sancion=safe_text(row_dict, "RESOLUCIÓN_SANCIÓN"),
                        fecha_resolucion_sancion=safe_datetime(row_dict, "FECHA_RESOLUCION_SANCIÓN"),
                        valor_pagado_de_intereses=safe_text(row_dict, "VALOR_PAGADO_INTERESES")
                    )
                    records.append(recaudo)
                    total_guardadas += 1

                    # Log de progreso periódico para detectar si el proceso se detuvo
                    if rows_in_chunk % PROGRESS_LOG_EVERY == 0:
                        elapsed = time.time() - chunk_start_time
                        logger.info(f"Progreso chunk: {rows_in_chunk}/{chunk_len} filas procesadas en este chunk, total guardadas: {total_guardadas} (chunk {chunk_start}-{chunk_start+chunk_len}) - {elapsed:.1f}s")
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote (batch_insert_records hace batch_size = 5000)
            if records:
                batch_insert_records(records, BATCH_SIZE)
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")

        try:
            os.remove(tmp_path)
        except Exception:
            pass

        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
        
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
    
def get_recaudos_derechos(file_id, file_path, headers, site_id, drive_id):
    """Versión optimizada: chunks + itertuples + batch insert"""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

        # Descargar en streaming a temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmpf:
            tmp_path = tmpf.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando archivo: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = _leer_excel_multihojas(tmp_path, file_path)
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path, low_memory=False)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            chunk_len = len(df_chunk)
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + chunk_len} ({chunk_len} filas)")

            records = []
            rows_in_chunk = 0
            chunk_start_time = time.time()
            for row in df_chunk.itertuples(index=False, name='Record'):
                try:
                    rows_in_chunk += 1
                    row_dict = row._asdict()
                    primer_valor = next(iter(row_dict.values()), None)
                    if (safe_text({"_0": primer_valor}, "_0") or "").strip().upper() == "FUENTE":
                        continue

                    fuente = safe_text(row_dict, "FUENTE") or safe_text(row_dict, "ORIGEN_RECAUDO")
                    
                    recaudo = Recaudo(
                        archivo_id=file_id,
                        organismo=organismo,
                        tipo_recaudo="DERECHOS DE TRANSITO",
                        fuente=fuente,
                        fecha_pago=safe_datetime(row_dict, "FECHA_PAGO") or safe_datetime(row_dict, "FECHA PAGO") or safe_datetime(row_dict, "_1"),
                        recibo=safe_receipt(row_dict, "RECIBO"),
                        recibo_pago=safe_text(row_dict, "RECIBO_PAGO"),
                        valor_recibido=safe_text(row_dict, "VALOR_RECIBO"),
                        tipo_documento=safe_text(row_dict, "TIPO_DOCUMENTO"),
                        identificacion=safe_text(row_dict, "IDENTIFICACION"),
                        nombre=safe_text(row_dict, "NOMBRE"),
                        vehiculo_placa=safe_text(row_dict, "VEHI_PLACA") or safe_text(row_dict, "PLACA"),
                        comparendo=safe_text(row_dict, "COMPARENDO") or safe_text(row_dict, "CODIGO"),
                        fecha_comparendo=safe_datetime(row_dict, "COMP_FECHA") or safe_datetime(row_dict, "FECHA_COMPARENDO"),
                        año_comparendo=safe_text(row_dict, "AÑO_COMPARENDO"),
                        prescripcion=safe_text(row_dict, "PRESCRIPCION"),
                        tipo_comparendo=safe_text(row_dict, "TIPO_COMPARENDO"),
                        clase_vehiculo=safe_text(row_dict, "CLASE_VEHICULO") or safe_text(row_dict, "CLASE"),
                        tipo=safe_text(row_dict, "TIPO"),
                        servicio_vehiculo=safe_text(row_dict, "SERVICIO_VEHICULO") or safe_text(row_dict, "SERVICIO"),
                        valor_pagado=safe_text(row_dict, "VALOR_PAGADO"),
                        fecha_distribucion=safe_datetime(row_dict, "DISTRI_FECHA"),
                        resolucion_mp=safe_text(row_dict, "RESOLUCION_MP"),
                        valor_inicial_cargado=safe_text(row_dict, "VALOR_INICIAL_CAR") or safe_text(row_dict, "CART_VALOR_INICIAL") or safe_text(row_dict, "CARTERA_VALOR_INICIAL"),
                        concepto=safe_text(row_dict, "CONCEPTO"),
                        fecha_cartera=safe_datetime(row_dict, "FECHA_CARTERA") or safe_datetime(row_dict, "FECHA CARTERA"),
                        estado_cartera=safe_text(row_dict, "ESTADO_CARTERA") or safe_text(row_dict, "ESTADO CARTERA"),
                        tipo_cartera=safe_text(row_dict, "TIPO_CARTERA") or "DERECHOS DE TRANSITO",
                        concepto_principal=safe_text(row_dict, "CONCEPTO_PRINCIPAL"),
                        gestion=safe_text(row_dict, "GESTION") or safe_text(row_dict, "GESTIÓN") or safe_text(row_dict, "ESTADO_GESTION"),
                        descuento_cartera=safe_text(row_dict, "DESCUENTO_CARTERA"),
                        descuento_de_intereses=safe_text(row_dict, "DES_INTERESES"),
                        cantidad_de_descuento_cartera=safe_int(row_dict, "CANT_DESTO_CARTERA") or safe_int(row_dict, "CANT_DESCUENTO_CARTERA"),
                        cantidad_de_descuento_de_intereses=safe_int(row_dict, "CANT_DES_INTERESES") or safe_int(row_dict, "CANT_DESCUENTO_INTERESES"),
                        resolucion_sancion=safe_text(row_dict, "RESOLUCIÓN_SANCIÓN") or safe_text(row_dict, "RESOLUCION_SANCION"),
                        fecha_resolucion_sancion=safe_datetime(row_dict, "FECHA_RESOLUCION_SANCIÓN") or safe_datetime(row_dict, "FECHA_RESOLUCION_SANCION"),
                        valor_pagado_de_intereses=safe_text(row_dict, "VALOR_PAGADO_INTERESES"),
                        acuerdos_de_pago=safe_text(row_dict, "ACUERDOS_DE_PAGO"),
                        referencia=safe_text(row_dict, "REFERENCIA"),
                        sistematizacion=safe_text(row_dict, "SISTEMATIZACION")
                    )
                    records.append(recaudo)
                    
                    total_guardadas += 1

                    if rows_in_chunk % PROGRESS_LOG_EVERY == 0:
                        elapsed = time.time() - chunk_start_time
                        logger.info(f"Progreso chunk: {rows_in_chunk}/{chunk_len} filas procesadas en este chunk, total guardadas: {total_guardadas} - {elapsed:.1f}s")
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote
            if records:
                batch_insert_records(records, BATCH_SIZE)
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")

        try:
            os.remove(tmp_path)
        except Exception:
            pass

        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
        
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
    
def get_carteras_multas(file_id, file_path, headers, site_id, drive_id):
    """Versión optimizada: chunks + itertuples + batch insert"""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"

        # Descargar en streaming a temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_path)[1]) as tmpf:
            tmp_path = tmpf.name

        try:
            download_with_retries(download_url, tmp_path, headers=headers, timeout=(10, 900))
        except Exception as e:
            logger.error(f"Error descargando archivo: {str(e)}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}

        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = _leer_excel_multihojas(tmp_path, file_path)
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path, low_memory=False)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            chunk_len = len(df_chunk)
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + chunk_len} ({chunk_len} filas)")

            # Crear sesión local para este chunk
            db_session = Session()

            insert_mappings = []
            rows_in_chunk = 0
            chunk_start_time = time.time()
            for idx, row in enumerate(df_chunk.itertuples(index=False, name='Record')):
                try:
                    rows_in_chunk += 1
                    row_dict = row._asdict()
                    codigo = str(row_dict.get('CODIGO', '') or '')

                    insert_mappings.append({
                            'archivo_id': file_id,
                            'organismo': organismo,
                            'codigo': codigo,
                            'tipo_cartera': 'MULTAS',
                            'estado_cartera_final': WalletStatus.ACTIVE.value,
                            'fecha': safe_text(row_dict, 'FECHA_COMPARENDO'),
                            'tipo_comparendo': safe_text(row_dict, 'TIPO_COMPARENDO'),
                            'clase': safe_text(row_dict, 'CLASE'),
                            'servicio': safe_text(row_dict, 'SERVICIO'),
                            'valor_inicial_cartera': safe_text(row_dict, 'CART_VALOR_INICIAL'),
                            'numero_referencia_cartera': safe_text(row_dict, 'CART_NRO_REFERENCIA'),
                            'estado_cartera': safe_text(row_dict, 'ESTADO_CARTERA'),
                            'fecha_inicio_cartera': safe_text(row_dict, 'CART_FECHA_INGRESO'),
                            'estado_gestion': safe_text(row_dict, 'ESTADO_GESTION'),
                            'capital': safe_text(row_dict, 'CAPITAL'),
                            'total': safe_text(row_dict, 'TOTAL'),
                            'resolucion_fecha': safe_text(row_dict, 'RESOLUCION_FECHA'),
                            'intereses': safe_text(row_dict, 'INTERESES'),
                            'placa': safe_text(row_dict, 'PLACA'),
                            'tipo_identificacion': safe_text(row_dict, 'TIPO_IDENTIFICACION'),
                            'numero_identificacion': safe_text(row_dict, 'NUMERO_IDENTIFICACION'),
                            'nombre_infractor': safe_text(row_dict, 'NOMBRE_INFRACTOR'),
                            'numero_comparendo': safe_text(row_dict, 'NUMERO_COMPARENDO'),
                            'estado_comparendo': safe_text(row_dict, 'ESTADO_COMPARENDO'),
                            'infraccion': safe_text(row_dict, 'INFRACCION'),
                            'resolucion_sancion': safe_text(row_dict, 'RESOLUCION_SANCION'),
                            'mandamiento_de_pago': safe_text(row_dict, 'MANDAMIENTO_DE_PAGO'),
                            'fecha_mandamiento_de_pago': safe_text(row_dict, 'FECHA_MANDAMIENTO'),
                            'fecha_de_notificacion': safe_text(row_dict, 'NOTIF_FECHA'),
                            'clase_vehiculo': safe_text(row_dict, 'CLASE_VEHICULO'),
                            'año_comparendo': safe_text(row_dict, "TO_CHAR(FECHA_COMPARENDO,'YYYY')"),
                            'ciudad': safe_text(row_dict, 'NOMBRE_CIUDAD'),
                            'direccion': safe_text(row_dict, 'DIR_DIRECCION'),
                            'telefono': safe_text(row_dict, 'DIR_TELEFONO'),
                            'movil': safe_text(row_dict, 'MOVIL'),
                            'email': safe_text(row_dict, 'EMAIL')
                        })
                    
                    total_guardadas += 1

                    if rows_in_chunk % PROGRESS_LOG_EVERY == 0:
                        elapsed = time.time() - chunk_start_time
                        logger.info(f"Progreso chunk: {rows_in_chunk}/{chunk_len} filas procesadas en este chunk, total guardadas: {total_guardadas} - {elapsed:.1f}s")

                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Aplicar actualizaciones e inserciones en bloque
            try:
                if insert_mappings:
                    # dividir insert_mappings en sub-batches si es grande
                    for j in range(0, len(insert_mappings), BATCH_SIZE):
                        sub = insert_mappings[j:j + BATCH_SIZE]
                        db_session.bulk_insert_mappings(Cartera, sub)
                db_session.commit()
            except SQLAlchemyError as e:
                logger.error(f"Error aplicando bulk updates/inserts: {e}")
                try:
                    db_session.rollback()
                except Exception:
                    pass
            finally:
                try:
                    db_session.close()
                except Exception:
                    pass
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")

        try:
            os.remove(tmp_path)
        except Exception:
            pass

        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}