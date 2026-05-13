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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DOMINIO = os.getenv("DOMINIO")
NOMBRE_SITIO = os.getenv("NOMBRE_SITIO")

# CONFIGURACIÓN DE OPTIMIZACIÓN
BATCH_SIZE = 5000  # Insertar de a 5000 filas por commit
CHUNK_SIZE = 10000  # Leer 10k filas a la vez de Excel
MAX_RECURSION_DEPTH = 20  # Límite de profundidad en búsqueda de carpetas

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
    """Inserta registros en lotes (5000 filas/commit en lugar de 1 commit/fila)"""
    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        db_session = None
        try:
            db_session = Session()
            db_session.add_all(batch)
            db_session.commit()
            db_session.close()
            logger.info(f"Batch insertado: {len(batch)} registros")
        except Exception as e:
            try:
                if db_session is not None:
                    db_session.rollback()
                    db_session.close()
            except Exception:
                pass
            logger.error(f"Error en batch insert: {str(e)}")


def save_files_to_database(excel_files, headers, site_id, drive_id):
    """Guarda los archivos encontrados en la base de datos"""
    if not session:
        raise Exception("Base de datos no disponible")
    
    results = {
        "guardados": [],
        "actualizados": [],
        "errores": [],
        "procesados": []
    }
    
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
                            results["procesados"].append({
                                "archivo": file_info["file_path"],
                                "tipo": enum_item.name,
                                "filas_procesadas": resultado["filas_procesadas"],
                                "guardadas": resultado["guardadas"],
                                "errores": resultado["errores"]
                            })
                        elif enum_item == Document.RECAUDO_DERECHOS_DE_TRANSITO:
                            resultado = get_recaudos_derechos(
                                file_info["file_id"],
                                file_info["file_path"],
                                headers,
                                site_id,
                                drive_id
                            )
                            results["procesados"].append({
                                "archivo": file_info["file_path"],
                                "tipo": enum_item.name,
                                "filas_procesadas": resultado["filas_procesadas"],
                                "guardadas": resultado["guardadas"],
                                "errores": resultado["errores"]
                            })
                        elif enum_item == Document.CARTERA_MULTAS:
                            resultado = get_carteras_multas(
                                file_info["file_id"],
                                file_info["file_path"],
                                headers,
                                site_id,
                                drive_id
                            )
                            results["procesados"].append({
                                "archivo": file_info["file_path"],
                                "tipo": enum_item.name,
                                "filas_procesadas": resultado["filas_procesadas"],
                                "guardadas": resultado["guardadas"],
                                "errores": resultado["errores"]
                            })
                        elif enum_item == Document.CARTERA_DERECHOS_DE_TRANSITO:
                            resultado = get_carteras_derechos(
                                file_info["file_id"],
                                file_info["file_path"],
                                headers,
                                site_id,
                                drive_id
                            )
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
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
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
                        fuente=str(row_dict.get("FUENTE", "")),
                        fecha_pago=pd.to_datetime(row_dict.get("FECHA PAGO")) if row_dict.get("FECHA PAGO") else None,
                        recibo=str(row_dict.get("RECIBO", "")),
                        valor_recibido=str(row_dict.get("VALOR RECIBO", "")),
                        tipo_documento=str(row_dict.get("TIPO DOCUMENTO", "")) if row_dict.get("TIPO DOCUMENTO") else None,
                        identificacion=str(row_dict.get("IDENTIFICACION", "")),
                        nombre=str(row_dict.get("NOMBRE", "")) if row_dict.get("NOMBRE") else None,
                        vehiculo_placa=str(row_dict.get("VEHI PLACA", "")) if row_dict.get("VEHI PLACA") else None,
                        comparendo=str(row_dict.get("COMPARENDO", "")) if row_dict.get("COMPARENDO") else None,
                        fecha_comparendo=pd.to_datetime(row_dict.get("COMP FECHA")) if row_dict.get("COMP FECHA") else None,
                        año_comparendo=str(row_dict.get("AÑO COMPARENDO", "")) if row_dict.get("AÑO COMPARENDO") else None,
                        prescripcion=str(row_dict.get("PRESCRIPCION", "")),
                        tipo_comparendo=str(row_dict.get("TIPO COMPARENDO", "")),
                        clase_vehiculo=str(row_dict.get("CLASE VEHICULO", "")) if row_dict.get("CLASE VEHICULO") else None,
                        tipo=str(row_dict.get("TIPO", "")),
                        servicio_vehiculo=str(row_dict.get("SERVICIO VEHICULO", "")) if row_dict.get("SERVICIO VEHICULO") else None,
                        valor_pagado=str(row_dict.get("VALOR PAGADO", "")),
                        fecha_distribucion=pd.to_datetime(row_dict.get("DISTRI FECHA")) if row_dict.get("DISTRI FECHA") else None,
                        resolucion_mp=str(row_dict.get("RESOLUCION MP", "")) if row_dict.get("RESOLUCION MP") else None,
                        valor_inicial_cargado=str(row_dict.get("VALOR INICIAL CAR", "")) if row_dict.get("VALOR INICIAL CAR") else None,
                        concepto=str(row_dict.get("CONCEPTO", "")),
                        estado_cartera=str(row_dict.get("ESTADO CARTERA", "")) if row_dict.get("ESTADO CARTERA") else None,
                        concepto_principal=str(row_dict.get("CONCEPTO PRINCIPAL", "")) if row_dict.get("CONCEPTO PRINCIPAL") else None,
                        gestion=str(row_dict.get("GESTIÓN", "")),
                        descuento_cartera=str(row_dict.get("DESCUENTO CARTERA", "")) if row_dict.get("DESCUENTO CARTERA") else None,
                        descuento_de_intereses=str(row_dict.get("DES INTERESES", "")) if row_dict.get("DES INTERESES") else None,
                        cantidad_de_descuento_cartera=int(row_dict.get("CANT DESTO CARTERA")) if row_dict.get("CANT DESTO CARTERA") else None,
                        cantidad_de_descuento_de_intereses=int(row_dict.get("CANT DES INTERESES")) if row_dict.get("CANT DES INTERESES") else None,
                        resolucion_sancion=str(row_dict.get("RESOLUCIÓN SANCIÓN", "")) if row_dict.get("RESOLUCIÓN SANCIÓN") else None,
                        fecha_resolucion_sancion=pd.to_datetime(row_dict.get("FECHA RESOLUCION SANCIÓN")) if row_dict.get("FECHA RESOLUCION SANCIÓN") else None,
                        valor_pagado_de_intereses=str(row_dict.get("VALOR PAGADO INTERESES", "")) if row_dict.get("VALOR PAGADO INTERESES") else None
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
            df = pd.read_excel(BytesIO(resp.content), skiprows=3, skipfooter=1)
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content), skiprows=3, skipfooter=1)
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
                        tipo_recaudo="DERECHOS DE TRANSITO",
                        fuente=str(row_dict.get("FUENTE", "")),
                        fecha_pago=pd.to_datetime(row_dict.get("FECHA PAGO")) if row_dict.get("FECHA PAGO") and not isinstance(row_dict.get("FECHA PAGO"), str) else None,
                        recibo=str(row_dict.get("RECIBO", "")),
                        recibo_pago=str(row_dict.get("RECIBO PAGO", "")) if row_dict.get("RECIBO PAGO") else None,
                        valor_recibido=str(row_dict.get("VALOR RECIBO", "")),
                        tipo_documento=str(row_dict.get("TIPO DOCUMENTO", "")) if row_dict.get("TIPO DOCUMENTO") else None,
                        identificacion=str(row_dict.get("IDENTIFICACION", "")),
                        nombre=str(row_dict.get("NOMBRE", "")) if row_dict.get("NOMBRE") else None,
                        vehiculo_placa=str(row_dict.get("PLACA", "")) if row_dict.get("PLACA") else None,
                        comparendo=None,
                        fecha_comparendo=None,
                        año_comparendo=None,
                        prescripcion=str(row_dict.get("PRESCRIPCION", "")),
                        tipo_comparendo="",
                        clase_vehiculo=str(row_dict.get("CLASE VEHICULO", "")) if row_dict.get("CLASE VEHICULO") else None,
                        tipo=str(row_dict.get("TIPO", "")),
                        servicio_vehiculo=str(row_dict.get("SERVICIO VEHICULO", "")) if row_dict.get("SERVICIO VEHICULO") else None,
                        valor_pagado=str(row_dict.get("VALOR PAGADO", "")),
                        fecha_distribucion=pd.to_datetime(row_dict.get("FECHA CARTERA")) if row_dict.get("FECHA CARTERA") and not isinstance(row_dict.get("FECHA CARTERA"), str) else None,
                        resolucion_mp=str(row_dict.get("RESOLUCION MP", "")) if row_dict.get("RESOLUCION MP") else None,
                        valor_inicial_cargado=str(row_dict.get("CART VALOR INICIAL", "")) if row_dict.get("CART VALOR INICIAL") else None,
                        concepto=str(row_dict.get("CONCEPTO", "")),
                        fecha_cartera=pd.to_datetime(row_dict.get("FECHA CARTERA")) if row_dict.get("FECHA CARTERA") and not isinstance(row_dict.get("FECHA CARTERA"), str) else None,
                        estado_cartera=None,
                        tipo_cartera=str(row_dict.get("TIPO CARTERA", "")) if row_dict.get("TIPO CARTERA") else None,
                        concepto_principal=None,
                        gestion=str(row_dict.get("GESTIÓN", "")),
                        descuento_cartera=str(row_dict.get("DESCUENTO CARTERA", "")) if row_dict.get("DESCUENTO CARTERA") else None,
                        descuento_de_intereses=str(row_dict.get("DES INTERESES", "")) if row_dict.get("DES INTERESES") else None,
                        cantidad_de_descuento_cartera=int(row_dict.get("CANT_DESCUENTO_CARTERA")) if row_dict.get("CANT_DESCUENTO_CARTERA") and not isinstance(row_dict.get("CANT_DESCUENTO_CARTERA"), str) else None,
                        cantidad_de_descuento_de_intereses=int(row_dict.get("CANT_DESCUENTO_INTERESES")) if row_dict.get("CANT_DESCUENTO_INTERESES") and not isinstance(row_dict.get("CANT_DESCUENTO_INTERESES"), str) else None,
                        resolucion_sancion=str(row_dict.get("RESOLUCION LIQ", "")) if row_dict.get("RESOLUCION LIQ") else None,
                        fecha_resolucion_sancion=None,
                        valor_pagado_de_intereses=None,
                        acuerdos_de_pago=str(row_dict.get("ACUERDO DE PAGO", "")) if row_dict.get("ACUERDO DE PAGO") else None,
                        referencia=str(row_dict.get("REFERENCIA", "")) if row_dict.get("REFERENCIA") else None,
                        sistematizacion=str(row_dict.get("SISTEMATIZACION", "")) if row_dict.get("SISTEMATIZACION") else None
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
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
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
                        cartera_existente.fecha = str(row_dict.get("FECHA_COMPARENDO", "")) if row_dict.get("FECHA_COMPARENDO") else None
                        cartera_existente.tipo_comparendo = str(row_dict.get("TIPO_COMPARENDO", "")) if row_dict.get("TIPO_COMPARENDO") else None
                        cartera_existente.clase = str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None
                        cartera_existente.servicio = str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None
                        cartera_existente.valor_inicial_cartera = str(row_dict.get("CART_VALOR_INICIAL", "")) if row_dict.get("CART_VALOR_INICIAL") else None
                        cartera_existente.numero_referencia_cartera = str(row_dict.get("CART_NRO_REFERENCIA", "")) if row_dict.get("CART_NRO_REFERENCIA") else None
                        cartera_existente.estado_cartera = str(row_dict.get("ESTADO_CARTERA", "")) if row_dict.get("ESTADO_CARTERA") else None
                        cartera_existente.fecha_inicio_cartera = str(row_dict.get("CART_FECHA_INGRESO", "")) if row_dict.get("CART_FECHA_INGRESO") else None
                        cartera_existente.estado_gestion = str(row_dict.get("ESTADO_GESTION", "")) if row_dict.get("ESTADO_GESTION") else None
                        cartera_existente.capital = str(row_dict.get("CAPITAL", "")) if row_dict.get("CAPITAL") else None
                        cartera_existente.total = str(row_dict.get("TOTAL", "")) if row_dict.get("TOTAL") else None
                        hay_actualizaciones = True
                    else:
                        cartera = Cartera(
                            archivo_id=file_id,
                            organismo=organismo,
                            codigo=codigo,
                            tipo_cartera="MULTAS",
                            estado_cartera_final=WalletStatus.ACTIVE.value,
                            fecha=str(row_dict.get("FECHA_COMPARENDO", "")) if row_dict.get("FECHA_COMPARENDO") else None,
                            tipo_comparendo=str(row_dict.get("TIPO_COMPARENDO", "")) if row_dict.get("TIPO_COMPARENDO") else None,
                            clase=str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None,
                            servicio=str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None,
                            valor_inicial_cartera=str(row_dict.get("CART_VALOR_INICIAL", "")) if row_dict.get("CART_VALOR_INICIAL") else None,
                            numero_referencia_cartera=str(row_dict.get("CART_NRO_REFERENCIA", "")) if row_dict.get("CART_NRO_REFERENCIA") else None,
                            estado_cartera=str(row_dict.get("ESTADO_CARTERA", "")) if row_dict.get("ESTADO_CARTERA") else None,
                            fecha_inicio_cartera=str(row_dict.get("CART_FECHA_INGRESO", "")) if row_dict.get("CART_FECHA_INGRESO") else None,
                            estado_gestion=str(row_dict.get("ESTADO_GESTION", "")) if row_dict.get("ESTADO_GESTION") else None,
                            capital=str(row_dict.get("CAPITAL", "")) if row_dict.get("CAPITAL") else None,
                            total=str(row_dict.get("TOTAL", "")) if row_dict.get("TOTAL") else None
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
            logger.info(f"Excel leído: {len(df)} filas")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
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
                        cartera_existente.fecha = str(row_dict.get("FECHA_CARTERA", "")) if row_dict.get("FECHA_CARTERA") else None
                        cartera_existente.valor_inicial_cartera = str(row_dict.get("CARTERA_VALOR_INICIAL", "")) if row_dict.get("CARTERA_VALOR_INICIAL") else None
                        cartera_existente.numero_referencia_cartera = str(row_dict.get("REFERENCIA", "")) if row_dict.get("REFERENCIA") else None
                        cartera_existente.estado_cartera = str(row_dict.get("ESTADO_CARTERA", "")) if row_dict.get("ESTADO_CARTERA") else None
                        cartera_existente.estado_gestion = str(row_dict.get("ESTADO_GESTION", "")) if row_dict.get("ESTADO_GESTION") else None
                        cartera_existente.capital = str(row_dict.get("CAPITAL", "")) if row_dict.get("CAPITAL") else None
                        cartera_existente.total = str(row_dict.get("TOTAL", "")) if row_dict.get("TOTAL") else None
                        cartera_existente.fecha_inicio_cartera = str(row_dict.get("CARTERA_FECHA_DE_INGRESO", "")) if row_dict.get("CARTERA_FECHA_DE_INGRESO") else None
                        cartera_existente.intereses = str(row_dict.get("INTERESES", "")) if row_dict.get("INTERESES") else None
                        cartera_existente.placa = str(row_dict.get("PLACA", "")) if row_dict.get("PLACA") else None
                        cartera_existente.clase = str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None
                        cartera_existente.servicio = str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None
                        hay_actualizaciones = True
                    else:
                        cartera = Cartera(
                            archivo_id=file_id,
                            organismo=organismo,
                            codigo=codigo,
                            tipo_cartera="DERECHOS DE TRANSITO",
                            estado_cartera_final=WalletStatus.ACTIVE.value,
                            fecha=str(row_dict.get("FECHA_CARTERA", "")) if row_dict.get("FECHA_CARTERA") else None,
                            valor_inicial_cartera=str(row_dict.get("CARTERA_VALOR_INICIAL", "")) if row_dict.get("CARTERA_VALOR_INICIAL") else None,
                            numero_referencia_cartera=str(row_dict.get("REFERENCIA", "")) if row_dict.get("REFERENCIA") else None,
                            estado_cartera=str(row_dict.get("ESTADO_CARTERA", "")) if row_dict.get("ESTADO_CARTERA") else None,
                            estado_gestion=str(row_dict.get("ESTADO_GESTION", "")) if row_dict.get("ESTADO_GESTION") else None,
                            capital=str(row_dict.get("CAPITAL", "")) if row_dict.get("CAPITAL") else None,
                            total=str(row_dict.get("TOTAL", "")) if row_dict.get("TOTAL") else None,
                            fecha_inicio_cartera=str(row_dict.get("CARTERA_FECHA_DE_INGRESO", "")) if row_dict.get("CARTERA_FECHA_DE_INGRESO") else None,
                            intereses=str(row_dict.get("INTERESES", "")) if row_dict.get("INTERESES") else None,
                            placa=str(row_dict.get("PLACA", "")) if row_dict.get("PLACA") else None,
                            clase=str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None,
                            servicio=str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None
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