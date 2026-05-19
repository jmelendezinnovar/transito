"""
Módulo para rastrear el flujo de procesamiento de archivos.
Registra cada paso del procesamiento en la base de datos.
"""

import time
import logging
from datetime import datetime
from typing import Optional
from backend.models import Session, FileProcessingLog, ProcessingStepExecution

logger = logging.getLogger(__name__)


class FileProcessingTracker:
    """Gestor para rastrear el procesamiento de un archivo"""
    
    def __init__(self, archivo_id: str):
        self.archivo_id = archivo_id
        self.db_session = Session()
        self.processing_log: Optional[FileProcessingLog] = None
        self.current_step: Optional[ProcessingStepExecution] = None
        self.step_counter = 0
        self._init_processing_log()
    
    def _init_processing_log(self):
        """Inicializa el registro de procesamiento"""
        try:
            # Verificar si ya existe
            existing = self.db_session.query(FileProcessingLog).filter(
                FileProcessingLog.archivo_id == self.archivo_id
            ).first()
            
            if existing:
                self.processing_log = existing
            else:
                self.processing_log = FileProcessingLog(
                    archivo_id=self.archivo_id,
                    status="iniciado",
                    total_steps=6,  # Total de pasos esperados
                    completed_steps=0,
                    rows_processed=0,
                    rows_failed=0
                )
                self.db_session.add(self.processing_log)
                self.db_session.commit()
                logger.info(f"Iniciado procesamiento para archivo: {self.archivo_id}")
        except Exception as e:
            logger.error(f"Error inicializando log: {str(e)}")
            self.db_session.rollback()
    
    def start_step(self, step_name: str, details: str = "") -> 'FileProcessingTracker':
        """
        Inicia un nuevo paso de procesamiento.
        
        Args:
            step_name: Nombre del paso (ej: "descargando", "validando", "procesando")
            details: Detalles adicionales del paso
        
        Returns:
            self para encadenamiento
        """
        try:
            self.step_counter += 1
            self.current_step = ProcessingStepExecution(
                file_processing_log_id=self.processing_log.id,
                step_name=step_name,
                step_order=self.step_counter,
                status="ejecutando",
                details=details,
                records_processed=0,
                started_at=datetime.now()
            )
            self.db_session.add(self.current_step)
            self.db_session.commit()
            logger.info(f"Paso iniciado: {step_name}")
            return self
        except Exception as e:
            logger.error(f"Error iniciando paso: {str(e)}")
            self.db_session.rollback()
            return self
    
    def complete_step(self, records_processed: int = 0):
        """
        Completa el paso actual.
        
        Args:
            records_processed: Cantidad de registros procesados en este paso
        """
        try:
            if not self.current_step:
                return
            
            now = datetime.now()
            duration_ms = int((now - self.current_step.started_at).total_seconds() * 1000)
            
            self.current_step.status = "completado"
            self.current_step.duration_ms = duration_ms
            self.current_step.completed_at = now
            self.current_step.records_processed = records_processed
            
            self.db_session.commit()
            
            # Actualizar el log principal
            self.processing_log.completed_steps += 1
            self.processing_log.status = "procesando"
            self.db_session.commit()
            
            logger.info(f"Paso completado: {self.current_step.step_name} ({duration_ms}ms)")
        except Exception as e:
            logger.error(f"Error completando paso: {str(e)}")
            self.db_session.rollback()
    
    def step_error(self, error_message: str):
        """
        Marca el paso actual como fallido.
        
        Args:
            error_message: Mensaje del error
        """
        try:
            if not self.current_step:
                return
            
            now = datetime.now()
            duration_ms = int((now - self.current_step.started_at).total_seconds() * 1000) if self.current_step.started_at else 0
            
            self.current_step.status = "error"
            self.current_step.error_message = error_message
            self.current_step.duration_ms = duration_ms
            self.current_step.completed_at = now
            
            self.db_session.commit()
            
            # Actualizar el log principal como error
            self.processing_log.status = "error"
            self.processing_log.error_message = error_message
            self.db_session.commit()
            
            logger.error(f"Paso fallido: {self.current_step.step_name} - {error_message}")
        except Exception as e:
            logger.error(f"Error registrando fallo: {str(e)}")
            self.db_session.rollback()
    
    def update_progress(self, rows_processed: int = 0, rows_failed: int = 0):
        """
        Actualiza el progreso general del procesamiento.
        
        Args:
            rows_processed: Filas procesadas correctamente
            rows_failed: Filas que fallaron
        """
        try:
            self.processing_log.rows_processed = rows_processed
            self.processing_log.rows_failed = rows_failed
            self.db_session.commit()
        except Exception as e:
            logger.error(f"Error actualizando progreso: {str(e)}")
            self.db_session.rollback()
    
    def complete_processing(self):
        """
        Marca el procesamiento como completado.
        """
        try:
            now = datetime.now()
            start_time = self.processing_log.started_at
            duration_ms = int((now - start_time).total_seconds() * 1000)
            
            self.processing_log.status = "completado"
            self.processing_log.completed_at = now
            self.processing_log.total_duration_ms = duration_ms
            self.db_session.commit()
            
            logger.info(f"Procesamiento completado en {duration_ms}ms")
        except Exception as e:
            logger.error(f"Error completando procesamiento: {str(e)}")
            self.db_session.rollback()
    
    def fail_processing(self, error_message: str):
        """
        Marca el procesamiento como fallido.
        
        Args:
            error_message: Mensaje del error
        """
        try:
            now = datetime.now()
            start_time = self.processing_log.started_at
            duration_ms = int((now - start_time).total_seconds() * 1000)
            
            self.processing_log.status = "error"
            self.processing_log.error_message = error_message
            self.processing_log.completed_at = now
            self.processing_log.total_duration_ms = duration_ms
            self.db_session.commit()
            
            logger.error(f"Procesamiento fallido: {error_message}")
        except Exception as e:
            logger.error(f"Error registrando fallo de procesamiento: {str(e)}")
            self.db_session.rollback()
    
    def close(self):
        """Cierra la sesión de BD"""
        try:
            self.db_session.close()
        except Exception as e:
            logger.error(f"Error cerrando sesión: {str(e)}")
