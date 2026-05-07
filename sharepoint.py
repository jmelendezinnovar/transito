import msal
import requests
import os
from dotenv import load_dotenv
from models import Archivo, Recaudo, Cartera, Auditoria, session
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
ORGANISMO_PREFIJO = os.getenv("ORGANISMO_PREFIJO", "SECRETARIA DE TRANSITO")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

def extract_organismo(file_path):
    """Extrae el organismo del filepath basado en el prefijo definido en variables de entorno.
    
    Ejemplo: 
    - Prefijo: "SECRETARIA DE TRANSITO"
    - Filepath: "SECRETARIA DE TRANSITO TURBACO/RECAUDO/DERECHOS DE TRANSITO/2026/05/archivo.xlsx"
    - Retorna: "TURBACO"
    """
    try:
        if ORGANISMO_PREFIJO not in file_path:
            logger.warning(f"Prefijo '{ORGANISMO_PREFIJO}' no encontrado en filepath: {file_path}")
            return None
        
        # Encontrar donde termina el prefijo
        prefix_end = file_path.find(ORGANISMO_PREFIJO) + len(ORGANISMO_PREFIJO)
        
        # Obtener el resto del path después del prefijo
        remaining_path = file_path[prefix_end:].lstrip(" ")
        
        # Extraer lo que viene hasta el siguiente "/"
        if "/" in remaining_path:
            organismo = remaining_path.split("/")[0].strip()
            return organismo if organismo else None
        
        return None
    except Exception as e:
        logger.error(f"Error extrayendo organismo de {file_path}: {str(e)}")
        return None

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
        
        # Extraer organismo del filepath
        organismo = extract_organismo(file_path)
        
        # Procesar solo la primera fila
        for idx, row in df.iterrows():
            try:
                # Convertir fila a diccionario
                row_dict = row.to_dict()
                row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
                
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
    
def get_recaudos_derechos(file_id, file_path, headers, site_id, drive_id):
    try:
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
        resp = requests.get(download_url, headers=headers)
        
        if resp.status_code != 200:
            logger.error(f"Error descargando archivo: {resp.status_code}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        # Leer archivo según su extensión
        if file_path.lower().endswith(".xlsx"):
            df = pd.read_excel(BytesIO(resp.content), skiprows=3, skipfooter=1)
            logger.info(f"Archivo Excel leído correctamente")
        elif file_path.lower().endswith(".csv"):
            df = pd.read_csv(BytesIO(resp.content), skiprows=3, skipfooter=1)
            logger.info(f"Archivo CSV leído correctamente")
        else:
            logger.error(f"Formato de archivo no soportado: {file_path}")
            return {"filas_procesadas": 0, "guardadas": 0, "errores": 0}
        
        total_filas = len(df)
        total_columnas = len(df.columns)
        
        logger.info(f"Total de filas: {total_filas}")
        logger.info(f"Total de columnas: {total_columnas}")
        
        guardadas = 0
        errores = 0
        
        # Extraer organismo del filepath
        organismo = extract_organismo(file_path)
        
        for idx, row in df.iterrows():
            try:
                row_dict = row.to_dict()
                row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
                
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
    
def get_carteras_multas(file_id, file_path, headers, site_id, drive_id):
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
        
        # Extraer organismo del filepath
        organismo = extract_organismo(file_path)
        
        # Procesar cada fila
        for idx, row in df.iterrows():
            try:
                # Convertir fila a diccionario
                row_dict = row.to_dict()
                
                # Convertir NaN a None y valores vacíos
                row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
                
                codigo = str(row_dict.get("CODIGO", ""))
                
                # Buscar si ya existe una cartera con ese código
                cartera_existente = session.query(Cartera).filter_by(codigo=codigo).first()
                
                if cartera_existente:
                    # Actualizar cartera existente
                    cartera_existente.archivo_id = file_id
                    cartera_existente.organismo = organismo
                    cartera_existente.tipo_cartera = "MULTAS"
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
                    cartera_existente.resolucion_fecha = str(row_dict.get("RESOLUCION_FECHA", "")) if row_dict.get("RESOLUCION_FECHA") else None
                    cartera_existente.intereses = str(row_dict.get("INTERESES", "")) if row_dict.get("INTERESES") else None
                    cartera_existente.placa = str(row_dict.get("PLACA", "")) if row_dict.get("PLACA") else None
                    cartera_existente.tipo_identificacion = str(row_dict.get("TIPO_IDENTIFICACION", "")) if row_dict.get("TIPO_IDENTIFICACION") else None
                    cartera_existente.numero_identificacion = str(row_dict.get("NUMERO_IDENTIFICACION", "")) if row_dict.get("NUMERO_IDENTIFICACION") else None
                    cartera_existente.nombre_infractor = str(row_dict.get("NOMBRE_INFRACTOR", "")) if row_dict.get("NOMBRE_INFRACTOR") else None
                    cartera_existente.numero_comparendo = str(row_dict.get("NUMERO_COMPARENDO", "")) if row_dict.get("NUMERO_COMPARENDO") else None
                    cartera_existente.estado_comparendo = str(row_dict.get("ESTADO_COMPARENDO", "")) if row_dict.get("ESTADO_COMPARENDO") else None
                    cartera_existente.infraccion = str(row_dict.get("INFRACCION", "")) if row_dict.get("INFRACCION") else None
                    cartera_existente.resolucion_sancion = str(row_dict.get("RESOLUCION_SANCION", "")) if row_dict.get("RESOLUCION_SANCION") else None
                    cartera_existente.mandamiento_de_pago = str(row_dict.get("MANDAMIENTO_DE_PAGO", "")) if row_dict.get("MANDAMIENTO_DE_PAGO") else None
                    cartera_existente.fecha_mandamiento_de_pago = str(row_dict.get("FECHA_MANDAMIENTO", "")) if row_dict.get("FECHA_MANDAMIENTO") else None
                    cartera_existente.fecha_de_notificacion = str(row_dict.get("NOTIF_FECHA", "")) if row_dict.get("NOTIF_FECHA") else None
                    cartera_existente.clase_vehiculo = str(row_dict.get("CLASE_VEHICULO", "")) if row_dict.get("CLASE_VEHICULO") else None
                    cartera_existente.año_comparendo = str(row_dict.get("TO_CHAR(FECHA_COMPARENDO,'YYYY')", "")) if row_dict.get("TO_CHAR(FECHA_COMPARENDO,'YYYY')") else None
                    cartera_existente.ciudad = str(row_dict.get("NOMBRE_CIUDAD", "")) if row_dict.get("NOMBRE_CIUDAD") else None
                    cartera_existente.direccion = str(row_dict.get("DIR_DIRECCION", "")) if row_dict.get("DIR_DIRECCION") else None
                    cartera_existente.telefono = str(row_dict.get("DIR_TELEFONO", "")) if row_dict.get("DIR_TELEFONO") else None
                    cartera_existente.movil = str(row_dict.get("MOVIL", "")) if row_dict.get("MOVIL") else None
                    cartera_existente.email = str(row_dict.get("EMAIL", "")) if row_dict.get("EMAIL") else None
                    
                    session.commit()
                    guardadas += 1
                    logger.info(f"Cartera {codigo} actualizada")
                else:
                    # Crear nueva cartera
                    cartera = Cartera(
                        archivo_id=file_id,
                        organismo=organismo,
                        codigo=codigo,
                        tipo_cartera="MULTAS",
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
                        total=str(row_dict.get("TOTAL", "")) if row_dict.get("TOTAL") else None,
                        resolucion_fecha=str(row_dict.get("RESOLUCION_FECHA", "")) if row_dict.get("RESOLUCION_FECHA") else None,
                        intereses=str(row_dict.get("INTERESES", "")) if row_dict.get("INTERESES") else None,
                        placa=str(row_dict.get("PLACA", "")) if row_dict.get("PLACA") else None,
                        tipo_identificacion=str(row_dict.get("TIPO_IDENTIFICACION", "")) if row_dict.get("TIPO_IDENTIFICACION") else None,
                        numero_identificacion=str(row_dict.get("NUMERO_IDENTIFICACION", "")) if row_dict.get("NUMERO_IDENTIFICACION") else None,
                        nombre_infractor=str(row_dict.get("NOMBRE_INFRACTOR", "")) if row_dict.get("NOMBRE_INFRACTOR") else None,
                        numero_comparendo=str(row_dict.get("NUMERO_COMPARENDO", "")) if row_dict.get("NUMERO_COMPARENDO") else None,
                        estado_comparendo=str(row_dict.get("ESTADO_COMPARENDO", "")) if row_dict.get("ESTADO_COMPARENDO") else None,
                        infraccion=str(row_dict.get("INFRACCION", "")) if row_dict.get("INFRACCION") else None,
                        resolucion_sancion=str(row_dict.get("RESOLUCION_SANCION", "")) if row_dict.get("RESOLUCION_SANCION") else None,
                        mandamiento_de_pago=str(row_dict.get("MANDAMIENTO_DE_PAGO", "")) if row_dict.get("MANDAMIENTO_DE_PAGO") else None,
                        fecha_mandamiento_de_pago=str(row_dict.get("FECHA_MANDAMIENTO", "")) if row_dict.get("FECHA_MANDAMIENTO") else None,
                        fecha_de_notificacion=str(row_dict.get("NOTIF_FECHA", "")) if row_dict.get("NOTIF_FECHA") else None,
                        clase_vehiculo=str(row_dict.get("CLASE_VEHICULO", "")) if row_dict.get("CLASE_VEHICULO") else None,
                        año_comparendo=str(row_dict.get("TO_CHAR(FECHA_COMPARENDO,'YYYY')", "")) if row_dict.get("TO_CHAR(FECHA_COMPARENDO,'YYYY')") else None,
                        ciudad=str(row_dict.get("NOMBRE_CIUDAD", "")) if row_dict.get("NOMBRE_CIUDAD") else None,
                        direccion=str(row_dict.get("DIR_DIRECCION", "")) if row_dict.get("DIR_DIRECCION") else None,
                        telefono=str(row_dict.get("DIR_TELEFONO", "")) if row_dict.get("DIR_TELEFONO") else None,
                        movil=str(row_dict.get("MOVIL", "")) if row_dict.get("MOVIL") else None,
                        email=str(row_dict.get("EMAIL", "")) if row_dict.get("EMAIL") else None
                    )
                    
                    session.add(cartera)
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
    
def get_carteras_derechos(file_id, file_path, headers, site_id, drive_id):
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
        
        # Extraer organismo del filepath
        organismo = extract_organismo(file_path)
        
        # Procesar cada fila
        for idx, row in df.iterrows():
            try:
                # Convertir fila a diccionario
                row_dict = row.to_dict()
                
                # Convertir NaN a None y valores vacíos
                row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
                
                codigo = str(row_dict.get("CODIGO", ""))
                
                # Buscar si ya existe una cartera con ese código
                cartera_existente = session.query(Cartera).filter_by(codigo=codigo).first()
                
                if cartera_existente:
                    # Actualizar cartera existente
                    cartera_existente.archivo_id = file_id
                    cartera_existente.organismo = organismo
                    cartera_existente.tipo_cartera = "DERECHOS DE TRANSITO"
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
                    cartera_existente.modelo = str(row_dict.get("MODELO", "")) if row_dict.get("MODELO") else None
                    cartera_existente.clase = str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None
                    cartera_existente.servicio = str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None
                    cartera_existente.tipo_identificacion = str(row_dict.get("TIPO_IDENTIFICACION", "")) if row_dict.get("TIPO_IDENTIFICACION") else None
                    cartera_existente.numero_identificacion = str(row_dict.get("NUMERO_IDENTIFICACION", "")) if row_dict.get("NUMERO_IDENTIFICACION") else None
                    cartera_existente.nombre_infractor = str(row_dict.get("NOMBRE", "")) + " " + str(row_dict.get("APELLIDO", "")) if row_dict.get("NOMBRE") or row_dict.get("APELLIDO") else None
                    cartera_existente.email = str(row_dict.get("EMAIL", "")) if row_dict.get("EMAIL") else None
                    cartera_existente.movil = str(row_dict.get("TELEFONO_MOVIL", "")) if row_dict.get("TELEFONO_MOVIL") else None
                    cartera_existente.fecha_propietario = str(row_dict.get("FECHA_PROPIETARIO", "")) if row_dict.get("FECHA_PROPIETARIO") else None
                    cartera_existente.filtro_coactivo = str(row_dict.get("FILTRO_COACTIVO", "")) if row_dict.get("FILTRO_COACTIVO") else None
                    cartera_existente.clase_vehiculo = str(row_dict.get("CLASE_VEHICULO", "")) if row_dict.get("CLASE_VEHICULO") else None
                    cartera_existente.direccion = str(row_dict.get("DIRECCION", "")) if row_dict.get("DIRECCION") else None
                    cartera_existente.mp_resolucion = str(row_dict.get("MP_RESOLUCION", "")) if row_dict.get("MP_RESOLUCION") else None
                    cartera_existente.fecha_mp = str(row_dict.get("FECHA_MP", "")) if row_dict.get("FECHA_MP") else None
                    
                    session.commit()
                    guardadas += 1
                    logger.info(f"Cartera {codigo} actualizada")
                else:
                    # Crear nueva cartera
                    cartera = Cartera(
                        archivo_id=file_id,
                        organismo=organismo,
                        codigo=codigo,
                        tipo_cartera="DERECHOS DE TRANSITO",
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
                        modelo=str(row_dict.get("MODELO", "")) if row_dict.get("MODELO") else None,
                        clase=str(row_dict.get("CLASE", "")) if row_dict.get("CLASE") else None,
                        servicio=str(row_dict.get("SERVICIO", "")) if row_dict.get("SERVICIO") else None,
                        tipo_identificacion=str(row_dict.get("TIPO_IDENTIFICACION", "")) if row_dict.get("TIPO_IDENTIFICACION") else None,
                        numero_identificacion=str(row_dict.get("NUMERO_IDENTIFICACION", "")) if row_dict.get("NUMERO_IDENTIFICACION") else None,
                        nombre_infractor=str(row_dict.get("NOMBRE", "")) + " " + str(row_dict.get("APELLIDO", "")) if row_dict.get("NOMBRE") or row_dict.get("APELLIDO") else None,
                        email=str(row_dict.get("EMAIL", "")) if row_dict.get("EMAIL") else None,
                        movil=str(row_dict.get("TELEFONO_MOVIL", "")) if row_dict.get("TELEFONO_MOVIL") else None,
                        fecha_propietario=str(row_dict.get("FECHA_PROPIETARIO", "")) if row_dict.get("FECHA_PROPIETARIO") else None,
                        filtro_coactivo=str(row_dict.get("FILTRO_COACTIVO", "")) if row_dict.get("FILTRO_COACTIVO") else None,
                        clase_vehiculo=str(row_dict.get("CLASE_VEHICULO", "")) if row_dict.get("CLASE_VEHICULO") else None,
                        direccion=str(row_dict.get("DIRECCION", "")) if row_dict.get("DIRECCION") else None,
                        mp_resolucion=str(row_dict.get("MP_RESOLUCION", "")) if row_dict.get("MP_RESOLUCION") else None,
                        fecha_mp=str(row_dict.get("FECHA_MP", "")) if row_dict.get("FECHA_MP") else None
                    )
                    
                    session.add(cartera)
                    session.commit()
                    guardadas += 1
                    logger.info(f"Cartera {codigo} creada")
                
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