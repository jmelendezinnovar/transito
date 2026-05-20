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
from backend.models import Archivo, Session, session, Cartera, Ejecucion, Etapa
from backend.documents import WalletStatus, EjecucionEstado, EtapaNombre, EtapaEstado, Document
from backend.utils import extract_organismo, safe_text
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

def marcar_carteras_inactivas(organismo: str, tipo_cartera: str):
    """Marca como inactivas todas las carteras del mismo organismo y tipo antes de reimportar."""
    db = Session()
    try:
        db.query(Cartera).filter_by(
            organismo=organismo,
            tipo_cartera=tipo_cartera
        ).update(
            {Cartera.estado_cartera_final: WalletStatus.INACTIVE.value},
            synchronize_session=False
        )
        db.commit()
    except Exception as e:
        try:
            db.rollback()
        except Exception:
            pass
        logger.error(f"Error marcando carteras inactivas para {organismo}/{tipo_cartera}: {str(e)}")
    finally:
        try:
            db.close()
        except Exception:
            pass

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

def contar_filas_archivo(file_path: str, df: pd.DataFrame) -> int:
    try:
        if df is not None and not df.empty:
            return len(df)
        return 0
    except Exception as e:
        logger.warning(f"Error contando filas de {file_path}: {str(e)}")
        return 0
    
def get_datos_sucios_files() -> Dict:
    headers = get_graph_headers()
    site_id, drive_id = get_sharepoint_site_and_drive(headers)

    folder = _find_folder_by_name(drive_id, headers, "DATOS SUCIOS")
    if not folder:
        raise Exception("No se encontro la carpeta DATOS SUCIOS en SharePoint")

    files = _collect_files_from_folder(drive_id, folder["id"], headers, folder["path"])
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
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(tmp_path)
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path)
        else:
            logger.error(f"Formato no soportado: {file_path}")
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
            file_download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_info['file_id']}/content"
            filas_count = 0
            file_view_url = build_sharepoint_view_url(file_info.get("webUrl"), file_info["file_id"])
            
            try:
                # Descargar en streaming a temp file y contar filas
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_info["file_path"])[1]) as tmpf:
                    tmp_path = tmpf.name

                try:
                    download_with_retries(file_download_url, tmp_path, headers=headers, timeout=(10, 900))
                    if file_info["file_path"].lower().endswith(".xlsx"):
                        df = pd.read_excel(tmp_path)
                    elif file_info["file_path"].lower().endswith(".csv"):
                        df = pd.read_csv(tmp_path)
                    else:
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

            existing = session.query(Archivo).filter_by(archivo_id=file_info["file_id"]).first()
            
            if existing:
                existing.nombre = file_info["file_name"]
                existing.ruta = file_info["file_path"]
                existing.url = file_view_url
                existing.filas = filas_count
                session.commit()
                results["actualizados"].append(file_info["file_path"])
            else:
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
    # logger.info(f"📊 Procesando {len(archivos_recaudo)} archivos de recaudos...")
    # for file_info in archivos_recaudo:
    #     try:
    #         etapa_guardado = session.query(Etapa).filter_by(id=file_info["etapa_guardado_id"]).first()
    #         ejecucion = session.query(Ejecucion).filter_by(id=file_info["ejecucion_id"]).first()

    #         for enum_item in Document:
    #             if enum_item.value in file_info["file_path"]:
    #                 if enum_item == Document.RECAUDO_MULTAS:
    #                     resultado = get_recaudos_multas(
    #                         file_info["file_id"],
    #                         file_info["file_path"],
    #                         file_info["headers"],
    #                         file_info["site_id"],
    #                         file_info["drive_id"]
    #                     )
    #                 elif enum_item == Document.RECAUDO_DERECHOS_DE_TRANSITO:
    #                     resultado = get_recaudos_derechos(
    #                         file_info["file_id"],
    #                         file_info["file_path"],
    #                         file_info["headers"],
    #                         file_info["site_id"],
    #                         file_info["drive_id"]
    #                     )
    #                 else:
    #                     break
                    
    #                 if etapa_guardado:
    #                     complete_etapa(etapa_guardado)
    #                 if ejecucion:
    #                     complete_ejecucion(ejecucion)

    #                 results["procesados"].append({
    #                     "archivo": file_info["file_path"],
    #                     "tipo": enum_item.name,
    #                     "filas_procesadas": resultado["filas_procesadas"],
    #                     "guardadas": resultado["guardadas"],
    #                     "errores": resultado["errores"]
    #                 })
    #                 break
                    
    #     except Exception as e:
    #         session.rollback()
    #         try:
    #             if etapa_guardado:
    #                 fail_etapa(etapa_guardado)
    #             if ejecucion:
    #                 fail_ejecucion(ejecucion)
    #         except Exception:
    #             pass
    #         results["errores"].append({
    #             "archivo": file_info["file_path"],
    #             "error": str(e)
    #         })
    
    elapsed_total = time.time() - start_total
    logger.info(f"✅ Proceso completado en {elapsed_total:.2f}s (archivos: {len(excel_files)})")
    
    return results

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
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(tmp_path)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path)
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
        marcar_carteras_inactivas(organismo, "MULTAS")
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
                    apellido = safe_text(row_dict, 'APELLIDO')
                    if apellido and apellido.strip().upper() == 'INDETERMINADO':
                        continue

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
            df = pd.read_excel(tmp_path)
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(tmp_path)
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
        marcar_carteras_inactivas(organismo, "DERECHOS DE TRANSITO")
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
                    if apellido and apellido.strip().upper() == 'INDETERMINADO':
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