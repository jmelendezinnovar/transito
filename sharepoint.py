import msal
import requests
import os
from dotenv import load_dotenv
from models import Archivo, Recaudo, Cartera, Auditoria, session, Session
from documents import Document, WalletStatus
import pandas as pd
import logging
from io import BytesIO
from typing import List, Dict
from multiprocessing import Pool
import time

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DOMINIO = os.getenv("DOMINIO")
NOMBRE_SITIO = os.getenv("NOMBRE_SITIO")

# CONFIGURACIÓN DE OPTIMIZACIÓN - AUMENTADO PARA MEJOR PERFORMANCE
BATCH_SIZE = 50000  # Insertar de a 50k filas por commit (10x más rápido)
CHUNK_SIZE = 50000  # Leer 50k filas a la vez de Excel
MAX_RECURSION_DEPTH = 20  # Límite de profundidad en búsqueda de carpetas
NUM_WORKERS = 2  # Procesos paralelos (3 cores activos, usar 2 para no saturar)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

def extract_organismo(file_path):
    """
        Extrae el organismo del filepath basado en el prefijo 
        definido en variables de entorno.
    """
    try:
        if "/" in file_path:
            organismo = file_path.split("/")[0].strip()
            return organismo if organismo else None
        
        return None
    except Exception as e:
        logger.error(f"Error extrayendo organismo de {file_path}: {str(e)}")
        return None

def safe_text(row_dict, key, default=None):
    value = row_dict.get(key, default)
    if value is None or pd.isna(value):
        return default

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return default

    return text

def safe_datetime(row_dict, key):
    value = row_dict.get(key)
    if value is None or pd.isna(value):
        return None

    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed

def safe_int(row_dict, key):
    value = row_dict.get(key)
    if value is None or pd.isna(value):
        return None

    if isinstance(value, str):
        value = value.strip()
        if not value or value.lower() == "nan":
            return None

    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None

def marcar_carteras_inactivas(organismo: str, tipo_cartera: str):
    """Marca como inactivas todas las carteras del mismo organismo y tipo antes de reimportar."""
    try:
        session.query(Cartera).filter_by(
            organismo=organismo,
            tipo_cartera=tipo_cartera
        ).update(
            {Cartera.estado_cartera_final: WalletStatus.INACTIVE.value},
            synchronize_session=False
        )
        session.commit()
    except Exception as e:
        session.rollback()
        logger.error(f"Error marcando carteras inactivas para {organismo}/{tipo_cartera}: {str(e)}")

def get_access_token():
    """Obtiene token de acceso de Microsoft"""
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    token_result = app.acquire_token_for_client(SCOPE)
    return token_result["access_token"]

def get_sharepoint_files():
    """Obtiene los archivos del SharePoint y retorna lista de archivos Excel encontrados"""
    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Obtener site_id
    url = f"https://graph.microsoft.com/v1.0/sites/{DOMINIO}:/sites/{NOMBRE_SITIO}:"
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Error obteniendo site: {resp.status_code} - {resp.text}")
    
    site_data = resp.json()
    site_id = site_data["id"]
    
    # Obtener drives
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp_drives = requests.get(drives_url, headers=headers)
    if resp_drives.status_code != 200:
        raise Exception(f"Error obteniendo drives: {resp_drives.status_code}")
    
    drives_data = resp_drives.json()
    main_drive = drives_data["value"][0]["id"]
    
    # Obtener archivos
    files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{main_drive}/root/children"
    resp_files = requests.get(files_url, headers=headers)
    if resp_files.status_code != 200:
        raise Exception(f"Error obteniendo archivos: {resp_files.status_code}")
    
    excel_files = search_excel_files(main_drive, "root", headers, site_id)
    
    # Retornar datos necesarios
    return {
        "files": excel_files,
        "headers": headers,
        "site_id": site_id,
        "drive_id": main_drive
    }

def search_excel_files(drive_id, item_id, headers, site_id, path="", depth=0):
    """Búsqueda de archivos Excel con paginación y límite de profundidad"""
    if depth > MAX_RECURSION_DEPTH:
        logger.warning(f"Límite de recursión alcanzado en: {path}")
        return []
    
    try:
        excel_files = []
        
        if item_id == "root":
            children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        else:
            children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
        
        # Paginar resultados (Graph API retorna max 200 items)
        while children_url:
            resp_children = requests.get(children_url, headers=headers, timeout=30)
            
            if resp_children.status_code != 200:
                logger.error(f"Error en {path}: {resp_children.status_code}")
                return excel_files
            
            response_data = resp_children.json()
            items = response_data.get("value", [])
            
            for item in items:
                current_path = f"{path}/{item['name']}" if path else item['name']
                item_id_val = item.get("id")
                
                if "file" in item and item["name"].lower().endswith((".xlsx", ".csv")):
                    excel_files.append({
                        "file_id": item_id_val,
                        "file_name": item["name"],
                        "file_path": current_path
                    })
                elif "folder" in item:
                    subfolder_files = search_excel_files(drive_id, item_id_val, headers, site_id, current_path, depth + 1)
                    excel_files.extend(subfolder_files)
            
            # Verificar paginación
            children_url = response_data.get("@odata.nextLink")
        
        return excel_files
        
    except Exception as e:
        logger.error(f"Error explorando {path}: {str(e)}")
        return []

def batch_insert_records(records: List, batch_size: int = BATCH_SIZE):
    """Inserta registros en lotes masivos (50k/commit) con optimización de BD"""
    if not records:
        return
    
    start_time = time.time()
    total_inserted = 0
    
    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        db_session = None
        try:
            db_session = Session()
            # Desactivar autoflush para mejor performance
            db_session.bulk_insert_mappings(batch[0].__class__, [record.__dict__ for record in batch])
            db_session.commit()
            db_session.close()
            total_inserted += len(batch)
            logger.info(f"✓ Batch insertado: {len(batch):,} registros")
        except Exception as e:
            try:
                if db_session is not None:
                    db_session.rollback()
                    db_session.close()
            except Exception:
                pass
            logger.error(f"Error en batch insert: {str(e)}")
    
    elapsed = time.time() - start_time
    logger.info(f"Total insertados: {total_inserted:,} registros en {elapsed:.2f}s ({total_inserted/elapsed:,.0f} filas/seg)")


def descargar_y_parsear_excel(file_id: str, file_path: str, headers: Dict, drive_id: str) -> Dict:
    """Descarga y parsea archivo Excel/CSV desde SharePoint"""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando {file_path}: {resp.status_code}")
            return None
        
        start_time = time.time()
        
        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
        else:
            logger.error(f"Formato no soportado: {file_path}")
            return None
        
        # Normalizar columnas
        df.columns = (
            df.columns.str.strip()
            .str.upper()
            .str.replace(" ", "_", regex=False)
        )
        
        elapsed = time.time() - start_time
        logger.info(f"✓ Archivo parseado: {file_path} ({len(df):,} filas en {elapsed:.2f}s)")
        
        return {"df": df, "file_path": file_path, "file_id": file_id}
        
    except Exception as e:
        logger.error(f"Error descargando {file_path}: {str(e)}")
        return None


def procesar_archivo_cartera(args_tuple):
    """Función auxiliar para procesar carteras en paralelo"""
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


def save_files_to_database(excel_files, headers, site_id, drive_id):
    """Guarda los archivos encontrados en la base de datos con PARALELIZACIÓN"""
    if not session:
        raise Exception("Base de datos no disponible")
    
    results = {
        "guardados": [],
        "actualizados": [],
        "errores": [],
        "procesados": []
    }
    
    start_total = time.time()
    
    # PASO 1: Registrar archivos en BD
    archivos_cartera = []
    archivos_recaudo = []
    
    for file_info in excel_files:
        try:
            existing = session.query(Archivo).filter_by(archivo_id=file_info["file_id"]).first()
            
            if existing:
                existing.nombre = file_info["file_name"]
                existing.ruta = file_info["file_path"]
                session.commit()
                results["actualizados"].append(file_info["file_path"])
            else:
                file_record = Archivo(
                    archivo_id=file_info["file_id"],
                    nombre=file_info["file_name"],
                    ruta=file_info["file_path"]
                )
                session.add(file_record)
                session.commit()
                results["guardados"].append(file_info["file_path"])
            
            # Clasificar archivos por tipo
            if Document.CARTERA_MULTAS.value in file_info["file_path"] or Document.CARTERA_DERECHOS_DE_TRANSITO.value in file_info["file_path"]:
                archivos_cartera.append((file_info["file_id"], file_info["file_path"], headers, site_id, drive_id))
            elif Document.RECAUDO_MULTAS.value in file_info["file_path"] or Document.RECAUDO_DERECHOS_DE_TRANSITO.value in file_info["file_path"]:
                archivos_recaudo.append(file_info)
                    
        except Exception as e:
            session.rollback()
            results["errores"].append({
                "archivo": file_info["file_path"],
                "error": str(e)
            })
    
    # PASO 2: Procesar CARTERAS en paralelo (2-3 workers)
    logger.info(f"🚀 Iniciando procesamiento paralelo de {len(archivos_cartera)} archivos de cartera...")
    if archivos_cartera:
        try:
            with Pool(processes=NUM_WORKERS) as pool:
                cartera_results = pool.map(procesar_archivo_cartera, archivos_cartera)
                
            for resultado in cartera_results:
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
            for enum_item in Document:
                if enum_item.value in file_info["file_path"]:
                    if enum_item == Document.RECAUDO_MULTAS:
                        resultado = get_recaudos_multas(
                            file_info["file_id"],
                            file_info["file_path"],
                            headers,
                            site_id,
                            drive_id
                        )
                    elif enum_item == Document.RECAUDO_DERECHOS_DE_TRANSITO:
                        resultado = get_recaudos_derechos(
                            file_info["file_id"],
                            file_info["file_path"],
                            headers,
                            site_id,
                            drive_id
                        )
                    else:
                        break
                    
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
            results["errores"].append({
                "archivo": file_info["file_path"],
                "error": str(e)
            })
    
    elapsed_total = time.time() - start_total
    logger.info(f"✅ Proceso completado en {elapsed_total:.2f}s (archivos: {len(excel_files)})")
    
    return results

def get_recaudos_multas(file_id, file_path, headers, site_id, drive_id):
    """
    Lee un archivo de recaudos de multas (Excel o CSV) en CHUNKS.
    Usa itertuples() en lugar de iterrows() (100x más rápido).
    Batch insert de 5000 filas (40 commits en lugar de 200k).
    """
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + len(df_chunk)} ({len(df_chunk)} filas)")
            
            records = []
            for row in df_chunk.itertuples(index=False, name='Record'):
                try:
                    row_dict = row._asdict()
                    
                    recaudo = Recaudo(
                        archivo_id=file_id,
                        organismo=organismo,
                        tipo_recaudo="MULTAS",
                        fuente=safe_text(row_dict, "FUENTE"),
                        fecha_pago=safe_datetime(row_dict, "FECHA_PAGO") or safe_datetime(row_dict, "_1"),
                        recibo=safe_text(row_dict, "RECIBO"),
                        valor_recibido=safe_text(row_dict, "VALOR_RECIBO"),
                        tipo_documento=safe_text(row_dict, "TIPO_DOCUMENTO"),
                        identificacion=safe_text(row_dict, "IDENTIFICACION"),
                        nombre=safe_text(row_dict, "NOMBRE"),
                        vehiculo_placa=safe_text(row_dict, "VEHI_PLACA"),
                        comparendo=safe_text(row_dict, "COMPARENDO"),
                        fecha_comparendo=safe_datetime(row_dict, "COMP_FECHA"),
                        año_comparendo=safe_text(row_dict, "AÑO_COMPARENDO"),
                        prescripcion=safe_text(row_dict, "PRESCRIPCION"),
                        tipo_comparendo=safe_text(row_dict, "TIPO_COMPARENDO"),
                        clase_vehiculo=safe_text(row_dict, "CLASE_VEHICULO"),
                        tipo=safe_text(row_dict, "TIPO"),
                        servicio_vehiculo=safe_text(row_dict, "SERVICIO_VEHICULO"),
                        valor_pagado=safe_text(row_dict, "VALOR_PAGADO"),
                        fecha_distribucion=safe_datetime(row_dict, "DISTRI_FECHA"),
                        resolucion_mp=safe_text(row_dict, "RESOLUCION_MP"),
                        valor_inicial_cargado=safe_text(row_dict, "VALOR_INICIAL_CAR"),
                        concepto=safe_text(row_dict, "CONCEPTO"),
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
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote (batch_insert_records hace batch_size = 5000)
            if records:
                batch_insert_records(records, BATCH_SIZE)
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")
        
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
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        total_guardadas = 0
        total_errores = 0
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + len(df_chunk)} ({len(df_chunk)} filas)")
            
            records = []
            for row in df_chunk.itertuples(index=False, name='Record'):
                try:
                    row_dict = row._asdict()

                    print(row_dict.keys())
                    
                    recaudo = Recaudo(
                        archivo_id=file_id,
                        organismo=organismo,
                        tipo_recaudo="DERECHOS DE TRANSITO",
                        fuente=safe_text(row_dict, "FUENTE"),
                        fecha_pago=safe_datetime(row_dict, "FECHA_PAGO") or safe_datetime(row_dict, "FECHA PAGO") or safe_datetime(row_dict, "_1"),
                        recibo=safe_text(row_dict, "RECIBO"),
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
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote
            if records:
                batch_insert_records(records, BATCH_SIZE)
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")
        
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
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        marcar_carteras_inactivas(organismo, "MULTAS")
        total_guardadas = 0
        total_errores = 0
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + len(df_chunk)} ({len(df_chunk)} filas)")
            
            records = []
            hay_actualizaciones = False
            for row in df_chunk.itertuples(index=False, name='Record'):
                try:
                    row_dict = row._asdict()
                    
                    codigo = str(row_dict.get("CODIGO", ""))
                    with session.no_autoflush:
                        cartera_existente = session.query(Cartera).filter_by(codigo=codigo, organismo=organismo).first()
                    
                    if cartera_existente:
                        cartera_existente.archivo_id = file_id
                        cartera_existente.organismo = organismo
                        cartera_existente.tipo_cartera = "MULTAS"
                        cartera_existente.estado_cartera_final = WalletStatus.ACTIVE.value
                        cartera_existente.fecha = safe_text(row_dict, "FECHA_COMPARENDO")
                        cartera_existente.tipo_comparendo = safe_text(row_dict, "TIPO_COMPARENDO")
                        cartera_existente.clase = safe_text(row_dict, "CLASE")
                        cartera_existente.servicio = safe_text(row_dict, "SERVICIO")
                        cartera_existente.valor_inicial_cartera = safe_text(row_dict, "CART_VALOR_INICIAL")
                        cartera_existente.numero_referencia_cartera = safe_text(row_dict, "CART_NRO_REFERENCIA")
                        cartera_existente.estado_cartera = safe_text(row_dict, "ESTADO_CARTERA")
                        cartera_existente.fecha_inicio_cartera = safe_text(row_dict, "CART_FECHA_INGRESO")
                        cartera_existente.estado_gestion = safe_text(row_dict, "ESTADO_GESTION")
                        cartera_existente.capital = safe_text(row_dict, "CAPITAL")
                        cartera_existente.total = safe_text(row_dict, "TOTAL")
                        cartera_existente.resolucion_fecha = safe_text(row_dict, "RESOLUCION_FECHA")
                        cartera_existente.intereses = safe_text(row_dict, "INTERESES")
                        cartera_existente.placa = safe_text(row_dict, "PLACA")
                        cartera_existente.tipo_identificacion = safe_text(row_dict, "TIPO_IDENTIFICACION")
                        cartera_existente.numero_identificacion = safe_text(row_dict, "NUMERO_IDENTIFICACION")
                        cartera_existente.nombre_infractor = safe_text(row_dict, "NOMBRE_INFRACTOR")
                        cartera_existente.numero_comparendo = safe_text(row_dict, "NUMERO_COMPARENDO")
                        cartera_existente.estado_comparendo = safe_text(row_dict, "ESTADO_COMPARENDO")
                        cartera_existente.infraccion = safe_text(row_dict, "INFRACCION")
                        cartera_existente.resolucion_sancion = safe_text(row_dict, "RESOLUCION_SANCION")
                        cartera_existente.mandamiento_de_pago = safe_text(row_dict, "MANDAMIENTO_DE_PAGO")
                        cartera_existente.fecha_mandamiento_de_pago = safe_text(row_dict, "FECHA_MANDAMIENTO")
                        cartera_existente.fecha_de_notificacion = safe_text(row_dict, "NOTIF_FECHA")
                        cartera_existente.clase_vehiculo = safe_text(row_dict, "CLASE_VEHICULO")
                        cartera_existente.año_comparendo = safe_text(row_dict, "TO_CHAR(FECHA_COMPARENDO,'YYYY')")
                        cartera_existente.ciudad = safe_text(row_dict, "NOMBRE_CIUDAD")
                        cartera_existente.direccion = safe_text(row_dict, "DIR_DIRECCION")
                        cartera_existente.telefono = safe_text(row_dict, "DIR_TELEFONO")
                        cartera_existente.movil = safe_text(row_dict, "MOVIL")
                        cartera_existente.email = safe_text(row_dict, "EMAIL")
                        hay_actualizaciones = True
                    else:
                        cartera = Cartera(
                            archivo_id=file_id,
                            organismo=organismo,
                            codigo=codigo,
                            tipo_cartera="MULTAS",
                            estado_cartera_final=WalletStatus.ACTIVE.value,
                            fecha=safe_text(row_dict, "FECHA_COMPARENDO"),
                            tipo_comparendo=safe_text(row_dict, "TIPO_COMPARENDO"),
                            clase=safe_text(row_dict, "CLASE"),
                            servicio=safe_text(row_dict, "SERVICIO"),
                            valor_inicial_cartera=safe_text(row_dict, "CART_VALOR_INICIAL"),
                            numero_referencia_cartera=safe_text(row_dict, "CART_NRO_REFERENCIA"),
                            estado_cartera=safe_text(row_dict, "ESTADO_CARTERA"),
                            fecha_inicio_cartera=safe_text(row_dict, "CART_FECHA_INGRESO"),
                            estado_gestion=safe_text(row_dict, "ESTADO_GESTION"),
                            capital=safe_text(row_dict, "CAPITAL"),
                            total=safe_text(row_dict, "TOTAL"),
                            resolucion_fecha=safe_text(row_dict, "RESOLUCION_FECHA"),
                            intereses = safe_text(row_dict, "INTERESES"),
                            placa = safe_text(row_dict, "PLACA"),
                            tipo_identificacion = safe_text(row_dict, "TIPO_IDENTIFICACION"),
                            numero_identificacion = safe_text(row_dict, "NUMERO_IDENTIFICACION"),
                            nombre_infractor = safe_text(row_dict, "NOMBRE_INFRACTOR"),
                            numero_comparendo = safe_text(row_dict, "NUMERO_COMPARENDO"),
                            estado_comparendo = safe_text(row_dict, "ESTADO_COMPARENDO"),
                            infraccion = safe_text(row_dict, "INFRACCION"),
                            resolucion_sancion = safe_text(row_dict, "RESOLUCION_SANCION"),
                            mandamiento_de_pago = safe_text(row_dict, "MANDAMIENTO_DE_PAGO"),
                            fecha_mandamiento_de_pago = safe_text(row_dict, "FECHA_MANDAMIENTO"),
                            fecha_de_notificacion = safe_text(row_dict, "NOTIF_FECHA"),
                            clase_vehiculo = safe_text(row_dict, "CLASE_VEHICULO"),
                            año_comparendo = safe_text(row_dict, "TO_CHAR(FECHA_COMPARENDO,'YYYY')"),
                            ciudad = safe_text(row_dict, "NOMBRE_CIUDAD"),
                            direccion = safe_text(row_dict, "DIR_DIRECCION"),
                            telefono = safe_text(row_dict, "DIR_TELEFONO"),
                            movil = safe_text(row_dict, "MOVIL"),
                            email = safe_text(row_dict, "EMAIL")
                        )
                        records.append(cartera)
                    
                    total_guardadas += 1
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote
            if records:
                batch_insert_records(records, BATCH_SIZE)

            if hay_actualizaciones:
                session.commit()
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")
        
        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
    
def get_carteras_derechos(file_id, file_path, headers, site_id, drive_id):
    """Versión optimizada: chunks + itertuples + batch insert"""
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        resp = requests.get(download_url, headers=headers, timeout=60)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
            df.columns = (
                df.columns.str.strip()
                .str.upper()
                .str.replace(" ", "_", regex=False)
            )
            logger.info(f"CSV leído: {len(df)} filas")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        organismo = extract_organismo(file_path)
        marcar_carteras_inactivas(organismo, "DERECHOS DE TRANSITO")
        total_guardadas = 0
        total_errores = 0
        
        # Procesar en chunks lógicos (iloc)
        for chunk_start in range(0, len(df), CHUNK_SIZE):
            df_chunk = df.iloc[chunk_start:chunk_start + CHUNK_SIZE]
            logger.info(f"Procesando filas {chunk_start} a {chunk_start + len(df_chunk)} ({len(df_chunk)} filas)")
            
            records = []
            hay_actualizaciones = False
            for row in df_chunk.itertuples(index=False, name='Record'):
                try:
                    row_dict = row._asdict()
                    
                    codigo = str(row_dict.get("CODIGO", ""))
                    with session.no_autoflush:
                        cartera_existente = session.query(Cartera).filter_by(codigo=codigo, organismo=organismo).first()
                    
                    if cartera_existente:
                        cartera_existente.archivo_id = file_id
                        cartera_existente.organismo = organismo
                        cartera_existente.tipo_cartera = "DERECHOS DE TRANSITO"
                        cartera_existente.estado_cartera_final = WalletStatus.ACTIVE.value
                        cartera_existente.fecha = safe_text(row_dict, "FECHA_CARTERA") or safe_text(row_dict, "FECHA CARTERA")
                        cartera_existente.valor_inicial_cartera = safe_text(row_dict, "CARTERA_VALOR_INICIAL")
                        cartera_existente.numero_referencia_cartera = safe_text(row_dict, "REFERENCIA")
                        cartera_existente.estado_cartera = safe_text(row_dict, "ESTADO_CARTERA") or safe_text(row_dict, "ESTADO CARTERA")
                        cartera_existente.estado_gestion = safe_text(row_dict, "ESTADO_GESTION")
                        cartera_existente.capital = safe_text(row_dict, "CAPITAL")
                        cartera_existente.total = safe_text(row_dict, "TOTAL")
                        cartera_existente.fecha_inicio_cartera = safe_text(row_dict, "CARTERA_FECHA_DE_INGRESO")
                        cartera_existente.intereses = safe_text(row_dict, "INTERESES")
                        cartera_existente.placa = safe_text(row_dict, "PLACA")
                        cartera_existente.clase = safe_text(row_dict, "CLASE")
                        cartera_existente.servicio = safe_text(row_dict, "SERVICIO")
                        cartera_existente.tipo_identificacion = safe_text(row_dict, "TIPO_IDENTIFICACION")
                        cartera_existente.numero_identificacion = safe_text(row_dict, "NUMERO_IDENTIFICACION")
                        cartera_existente.nombre_infractor = (
                            ' '.join(
                                p for p in (
                                    safe_text(row_dict, "NOMBRE"),
                                    safe_text(row_dict, "APELLIDO")
                                ) if p
                            ) or None
                        )
                        cartera_existente.email = safe_text(row_dict, "EMAIL")
                        cartera_existente.telefono = safe_text(row_dict, "TELEFONO_MOVIL")
                        cartera_existente.fecha_propietario = safe_text(row_dict, "FECHA_PROPIETARIO")
                        cartera_existente.filtro_coactivo = safe_text(row_dict, "FILTRO_COACTIVO")
                        cartera_existente.clase_vehiculo = safe_text(row_dict, "CLASE_VEHICULO")
                        cartera_existente.direccion = safe_text(row_dict, "DIRECCION")
                        cartera_existente.mp_resolucion = safe_text(row_dict, "MP_RESOLUCION")
                        cartera_existente.fecha_mp = safe_text(row_dict, "FECHA_MP")

                        hay_actualizaciones = True
                    else:
                        cartera = Cartera(
                            archivo_id=file_id,
                            organismo=organismo,
                            codigo=codigo,
                            tipo_cartera="DERECHOS DE TRANSITO",
                            estado_cartera_final=WalletStatus.ACTIVE.value,
                            fecha=safe_text(row_dict, "FECHA_CARTERA") or safe_text(row_dict, "FECHA CARTERA"),
                            valor_inicial_cartera=safe_text(row_dict, "CARTERA_VALOR_INICIAL"),
                            numero_referencia_cartera=safe_text(row_dict, "REFERENCIA"),
                            estado_cartera=safe_text(row_dict, "ESTADO_CARTERA") or safe_text(row_dict, "ESTADO CARTERA"),
                            estado_gestion=safe_text(row_dict, "ESTADO_GESTION"),
                            capital=safe_text(row_dict, "CAPITAL"),
                            total=safe_text(row_dict, "TOTAL"),
                            fecha_inicio_cartera=safe_text(row_dict, "CARTERA_FECHA_DE_INGRESO"),
                            intereses=safe_text(row_dict, "INTERESES"),
                            placa=safe_text(row_dict, "PLACA"),
                            clase=safe_text(row_dict, "CLASE"),
                            servicio=safe_text(row_dict, "SERVICIO"),
                            tipo_identificacion = safe_text(row_dict, "TIPO_IDENTIFICACION"),
                            numero_identificacion = safe_text(row_dict, "NUMERO_IDENTIFICACION"),
                            nombre_infractor = (
                                ' '.join(
                                    p for p in (
                                        safe_text(row_dict, "NOMBRE"),
                                        safe_text(row_dict, "APELLIDO")
                                    ) if p
                                ) or None
                            ),
                            email = safe_text(row_dict, "EMAIL"),
                            telefono = safe_text(row_dict, "TELEFONO_MOVIL"),
                            fecha_propietario = safe_text(row_dict, "FECHA_PROPIETARIO"),
                            filtro_coactivo = safe_text(row_dict, "FILTRO_COACTIVO"),
                            clase_vehiculo = safe_text(row_dict, "CLASE_VEHICULO"),
                            direccion = safe_text(row_dict, "DIRECCION"),
                            mp_resolucion = safe_text(row_dict, "MP_RESOLUCION"),
                            fecha_mp = safe_text(row_dict, "FECHA_MP")
                        )
                        records.append(cartera)
                    
                    total_guardadas += 1
                    
                except Exception as e:
                    total_errores += 1
                    logger.warning(f"Error en fila: {str(e)[:100]}")
            
            # Insertar lote
            if records:
                batch_insert_records(records, BATCH_SIZE)

            if hay_actualizaciones:
                session.commit()
        
        logger.info(f"Total guardadas: {total_guardadas}, Total errores: {total_errores}")
        
        return {
            "filas_procesadas": total_guardadas + total_errores,
            "guardadas": total_guardadas,
            "errores": total_errores
        }
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}