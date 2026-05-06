import msal
import requests
import os
from dotenv import load_dotenv
from models import Archivo, Recaudo, Auditoria, session
from documents import Document
import pandas as pd
import logging
from io import BytesIO

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DOMINIO = os.getenv("DOMINIO")
NOMBRE_SITIO = os.getenv("NOMBRE_SITIO")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

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
    indent = "  " * depth
    try:
        if item_id == "root":
            children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        else:
            children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
        
        resp_children = requests.get(children_url, headers=headers)
        
        if resp_children.status_code != 200:
            print(f"{indent}❌ Error: {resp_children.status_code}")
            return []
        
        response_data = resp_children.json()
        items = response_data.get("value", [])
        excel_files = []
        
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
        
        return excel_files
    
    except Exception as e:
        print(f"{indent}❌ EXCEPCIÓN explorando {path or 'raíz'}: {str(e)}")
        return []

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
    Lee un archivo de recaudos de multas (Excel o CSV) y procesa cada fila.
    Crea registros de Recaudo en BD y registra errores en Auditoria.
    """
    try:
        
        # Descargar archivo de SharePoint
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        resp = requests.get(download_url, headers=headers)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo según su extensión
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content))
            logger.info(f"Archivo Excel leído correctamente")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content))
            logger.info(f"Archivo CSV leído correctamente")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        total_filas = len(df)
        total_columnas = len(df.columns)
        
        logger.info(f"Total de filas: {total_filas}")
        logger.info(f"Total de columnas: {total_columnas}")
        logger.info(f"Columnas: {', '.join(df.columns.tolist())}")
        
        guardadas = 0
        errores = 0
        
        # Procesar cada fila
        for idx, row in df.iterrows():
            try:
                # Convertir fila a diccionario
                row_dict = row.to_dict()
                
                # Convertir NaN a None y valores vacíos
                row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
                
                recaudo = Recaudo(
                    archivo_id=file_id,
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
                
                session.add(recaudo)
                session.commit()
                guardadas += 1
                
            except Exception as e:
                session.rollback()
                errores += 1
                
                # Registrar error en tabla Auditoria
                try:
                    auditoria = Auditoria(
                        archivo_id=file_id,
                        columna=f"Fila {idx + 2} - Error: {str(e)}"
                    )
                    session.add(auditoria)
                    session.commit()
                except:
                    session.rollback()
                
                logger.warning(f"Fila {idx + 2} - Error: {str(e)[:100]}")
        
        logger.info(f"Filas guardadas: {guardadas}")
        logger.info(f"Filas con error: {errores}")
        
        return {
            "filas_procesadas": total_filas,
            "guardadas": guardadas,
            "errores": errores
        }
        
    except Exception as e:
        logger.error(f"Error procesando archivo {file_path}: {str(e)}")
        return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}