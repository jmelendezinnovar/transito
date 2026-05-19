import logging
import os
import sys
from pathlib import Path
from urllib.parse import urlparse

from fastapi import FastAPI, HTTPException, Request, BackgroundTasks
from fastapi.responses import JSONResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from typing import Optional, List

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = CURRENT_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from backend.sharepoint import get_sharepoint_files, save_files_to_database
from backend.models import Session, Archivo, FileProcessingLog, ProcessingStepExecution

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

def _origin_from_url(url_value: str | None) -> str | None:
    if not url_value:
        return None
    parsed = urlparse(url_value)
    if parsed.scheme and parsed.netloc:
        return f"{parsed.scheme}://{parsed.netloc}"
    return None


def _build_allowed_origins() -> List[str]:
    origins = {
        "http://localhost:3000",
        "http://localhost:5173",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:5173",
        "https://graph.microsoft.com",
        "https://login.microsoftonline.com",
    }

    notification_origin = _origin_from_url(os.getenv("NOTIFICATION_URL"))
    if notification_origin:
        origins.add(notification_origin)

    dominio = (os.getenv("DOMINIO") or "").strip()
    if dominio:
        normalized_domain = dominio.replace("https://", "").replace("http://", "").strip("/")
        if normalized_domain:
            origins.add(f"https://{normalized_domain}")

    custom_origins = os.getenv("CORS_ALLOW_ORIGINS", "")
    if custom_origins:
        for origin in custom_origins.split(","):
            cleaned = origin.strip()
            if cleaned:
                origins.add(cleaned)

    return sorted(origins)

app = FastAPI(
    title="API de Registro de Documentos",
    description="API para buscar y registrar archivos Excel desde SharePoint",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=_build_allowed_origins(),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def procesar_sincronizacion_sharepoint():
    """Ejecuta la sincronización completa fuera del ciclo de respuesta del webhook."""
    try:
        data = get_sharepoint_files()
        excel_files = data["files"]
        headers = data["headers"]
        site_id = data["site_id"]
        drive_id = data["drive_id"]

        if not excel_files:
            logger.info("Webhook: no se encontraron archivos Excel para procesar")
            return

        results = save_files_to_database(excel_files, headers, site_id, drive_id)
        total_procesados = len(results["guardados"]) + len(results["actualizados"])
        logger.info(f"Webhook: proceso completado. {total_procesados} archivos procesados")
    except Exception as e:
        logger.error(f"Error procesando sincronización en background: {str(e)}")

@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):
    validation_token = request.query_params.get("validationToken")
    
    print("web hook actvado")

    if validation_token:
        return PlainTextResponse(content=validation_token, status_code=200)

    background_tasks.add_task(procesar_sincronizacion_sharepoint)
    return JSONResponse(status_code=202, content={"estado": "ok"})

@app.get("/registrar-documento")
async def registrar_documento():
    try:
        data = get_sharepoint_files()
        excel_files = data["files"]
        headers = data["headers"]
        site_id = data["site_id"]
        drive_id = data["drive_id"]
        
        if not excel_files:
            return JSONResponse(
                status_code=200,
                content={
                    "estado": "exitoso",
                    "mensaje": "No se encontraron archivos Excel",
                    "total_encontrados": 0,
                    "guardados": [],
                    "actualizados": [],
                    "procesados": [],
                    "errores": []
                }
            )
        
        results = save_files_to_database(excel_files, headers, site_id, drive_id)
        total_procesados = len(results["guardados"]) + len(results["actualizados"])
        
        return JSONResponse(
            status_code=200,
            content={
                "estado": "exitoso",
                "mensaje": f"Proceso completado. {total_procesados} archivos procesados",
                "total_encontrados": len(excel_files),
                "guardados": {
                    "cantidad": len(results["guardados"]),
                    "archivos": results["guardados"]
                },
                "actualizados": {
                    "cantidad": len(results["actualizados"]),
                    "archivos": results["actualizados"]
                },
                "procesados": {
                    "cantidad": len(results["procesados"]),
                    "detalles": results["procesados"]
                },
                "errores": {
                    "cantidad": len(results["errores"]),
                    "detalles": results["errores"]
                }
            }
        )
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail={
                "estado": "error",
                "mensaje": f"Error al procesar: {str(e)}"
            }
        )

@app.get("/health")
async def health_check():
    """
    Endpoint de salud para verificar que la API está funcionando
    """
    return {
        "estado": "ok",
        "servicio": "API de Registro de Documentos"
    }

@app.get("/")
async def root():
    return {
        "nombre": "API de Registro de Documentos",
        "version": "1.0.0",
        "endpoints": {
            "registrar": "/registrar-documento - POST para buscar y registrar archivos",
            "health": "/health - GET para verificar estado de la API",
            "docs": "/docs - Documentación interactiva (Swagger)"
        }
    }

# ==================== ENDPOINTS DE TRAZABILIDAD ====================

@app.get("/api/archivos")
async def listar_archivos(
    estado: Optional[str] = None,
    tipo: Optional[str] = None,
    organismo: Optional[str] = None,
    skip: int = 0,
    limit: int = 50
):
    try:
        db_session = Session()
        
        query = db_session.query(
            Archivo.id,
            Archivo.archivo_id,
            Archivo.nombre,
            Archivo.ruta,
            Archivo.url,
            Archivo.rows,
            Archivo.created_at,
            FileProcessingLog.status,
            FileProcessingLog.completed_steps,
            FileProcessingLog.total_steps,
            FileProcessingLog.error_message,
            FileProcessingLog.total_duration_ms,
            FileProcessingLog.rows_processed,
            FileProcessingLog.rows_failed
        ).outerjoin(
            FileProcessingLog,
            Archivo.archivo_id == FileProcessingLog.archivo_id
        )
        
        if estado:
            query = query.filter(FileProcessingLog.status == estado)
        
        if tipo:
            query = query.filter(Archivo.ruta.like(f"%{tipo}%"))
        
        if organismo:
            query = query.filter(Archivo.ruta.like(f"{organismo}%"))
        
        query = query.order_by(Archivo.created_at.desc())
        
        total = query.count()
        archivos = query.offset(skip).limit(limit).all()
        
        resultado = []
        for archivo in archivos:
            resultado.append({
                "id": archivo[0],
                "archivo_id": archivo[1],
                "nombre": archivo[2],
                "ruta": archivo[3],
                "url": archivo[4],
                "rows": archivo[5],
                "fecha_creacion": archivo[6].isoformat() if archivo[6] else None,
                "status": archivo[7] or "sin_procesar",
                "pasos_completados": archivo[8] or 0,
                "total_pasos": archivo[9] or 0,
                "error_message": archivo[10],
                "duracion_ms": archivo[11],
                "filas_procesadas": archivo[12] or 0,
                "filas_fallidas": archivo[13] or 0
            })
        
        db_session.close()
        
        return JSONResponse(
            status_code=200,
            content={
                "total": total,
                "skip": skip,
                "limit": limit,
                "archivos": resultado
            }
        )
    except Exception as e:
        logger.error(f"Error listando archivos: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/archivos/{archivo_id}/flujo")
async def obtener_flujo_archivo(archivo_id: str):
    """
    Obtiene el detalle del flujo de procesamiento de un archivo específico.
    Incluye todos los pasos ejecutados con sus detalles.
    """
    try:
        db_session = Session()
        
        # Obtener info del archivo
        archivo = db_session.query(Archivo).filter(Archivo.archivo_id == archivo_id).first()
        if not archivo:
            db_session.close()
            raise HTTPException(status_code=404, detail="Archivo no encontrado")
        
        # Obtener log de procesamiento
        log = db_session.query(FileProcessingLog).filter(
            FileProcessingLog.archivo_id == archivo_id
        ).first()
        
        pasos = []
        if log:
            # Obtener todos los pasos ejecutados
            step_executions = db_session.query(ProcessingStepExecution).filter(
                ProcessingStepExecution.file_processing_log_id == log.id
            ).order_by(ProcessingStepExecution.step_order).all()
            
            pasos = [
                {
                    "id": step.id,
                    "nombre": step.step_name,
                    "orden": step.step_order,
                    "status": step.status,
                    "duracion_ms": step.duration_ms,
                    "detalles": step.details,
                    "mensaje_error": step.error_message,
                    "registros_procesados": step.records_processed,
                    "inicio": step.started_at.isoformat() if step.started_at else None,
                    "fin": step.completed_at.isoformat() if step.completed_at else None
                }
                for step in step_executions
            ]
        
        resultado = {
            "archivo": {
                "id": archivo.id,
                "archivo_id": archivo.archivo_id,
                "nombre": archivo.nombre,
                "ruta": archivo.ruta,
                "url": archivo.url,
                "rows": archivo.rows,
                "fecha_creacion": archivo.created_at.isoformat() if archivo.created_at else None
            },
            "procesamiento": {
                "status": log.status if log else "sin_procesar",
                "total_pasos": log.total_steps if log else 0,
                "pasos_completados": log.completed_steps if log else 0,
                "mensaje_error": log.error_message if log else None,
                "duracion_total_ms": log.total_duration_ms if log else None,
                "filas_procesadas": log.rows_processed if log else 0,
                "filas_fallidas": log.rows_failed if log else 0,
                "inicio": log.started_at.isoformat() if log and log.started_at else None,
                "fin": log.completed_at.isoformat() if log and log.completed_at else None
            },
            "pasos": pasos
        }
        
        db_session.close()
        
        return JSONResponse(status_code=200, content=resultado)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error obteniendo flujo: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/archivos/{archivo_id}/detalles")
async def obtener_detalles_archivo(archivo_id: str):
    """
    Obtiene información completa de un archivo incluyendo su flujo y resumen.
    """
    try:
        db_session = Session()
        
        archivo = db_session.query(Archivo).filter(Archivo.archivo_id == archivo_id).first()
        if not archivo:
            db_session.close()
            raise HTTPException(status_code=404, detail="Archivo no encontrado")
        
        log = db_session.query(FileProcessingLog).filter(
            FileProcessingLog.archivo_id == archivo_id
        ).first()
        
        # Contar registros por tabla relacionada
        from backend.models import Recaudo, Cartera
        recaudos_count = db_session.query(Recaudo).filter(Recaudo.archivo_id == archivo_id).count()
        carteras_count = db_session.query(Cartera).filter(Cartera.archivo_id == archivo_id).count()
        
        db_session.close()
        
        return JSONResponse(
            status_code=200,
            content={
                "archivo": {
                    "id": archivo.id,
                    "archivo_id": archivo.archivo_id,
                    "nombre": archivo.nombre,
                    "ruta": archivo.ruta,
                    "url": archivo.url,
                    "rows": archivo.rows,
                    "fecha_creacion": archivo.created_at.isoformat() if archivo.created_at else None
                },
                "procesamiento": {
                    "status": log.status if log else "sin_procesar",
                    "duracion_total_ms": log.total_duration_ms if log else None,
                    "filas_procesadas": log.rows_processed if log else 0,
                    "filas_fallidas": log.rows_failed if log else 0,
                    "porcentaje_exito": round((log.rows_processed / (log.rows_processed + log.rows_failed) * 100), 2) if log and (log.rows_processed + log.rows_failed) > 0 else 0
                },
                "estadisticas": {
                    "recaudos_registrados": recaudos_count,
                    "carteras_registradas": carteras_count,
                    "total_registros_guardados": recaudos_count + carteras_count
                }
            }
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error obteniendo detalles: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    print("Iniciando servidor FastAPI...")
    load_dotenv()
    

    uvicorn.run(
        app,
        host=os.getenv("HOST", "0.0.0.0"),
        port=8000,
        log_level="info"
    )
