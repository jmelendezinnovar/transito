from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, PlainTextResponse
from sharepoint import get_sharepoint_files, save_files_to_database
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="API de Registro de Documentos",
    description="API para buscar y registrar archivos Excel desde SharePoint",
    version="1.0.0"
)


@app.post("/webhook")
async def webhook(request: Request):
    validation_token = request.query_params.get("validationToken")
    
    print("web hook actvado")

    if validation_token:
        return PlainTextResponse(content=validation_token, status_code=200)

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

if __name__ == "__main__":
    import uvicorn
    print("Iniciando servidor FastAPI...")
    print("Documentación disponible en: http://localhost:8000/docs")
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        log_level="info"
    )
