import os
from sqlalchemy import ForeignKey, create_engine, Column, String, DateTime, Integer, func
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Usar pg8000 como driver PostgreSQL (puro Python, sin compilación necesaria)
DATABASE_URL = os.getenv("DATABASE_URL", "postgresql+pg8000://postgres:12345678@localhost:5432/transito")

engine = None
session = None
Session = None

try:
    engine = create_engine(DATABASE_URL, echo=False, pool_pre_ping=True)
    with engine.connect() as conn:
        logger.info("Conexión a PostgreSQL exitosa")
        
    Base = declarative_base()

    class Archivo(Base):
        __tablename__ = "archivos"
        
        id = Column(Integer, primary_key=True, autoincrement=True)
        archivo_id = Column(String, unique=True, nullable=False)
        nombre = Column(String, nullable=False)
        ruta = Column(String, nullable=False)
        created_at = Column(DateTime, default=func.now())

    class Recaudo(Base):
        __tablename__ = "recaudos"

        id = Column(Integer, primary_key=True, autoincrement=True)
        archivo_id = Column(String, ForeignKey("archivos.archivo_id"), nullable=False)
        fuente = Column(String, nullable=False)
        fecha_pago = Column(DateTime, nullable=False)
        recibo = Column(String, nullable=False)
        valor_recibido = Column(String, nullable=False)
        tipo_documento = Column(String, nullable=True)
        identificacion = Column(String, nullable=False)
        nombre = Column(String, nullable=True)
        vehiculo_placa = Column(String, nullable=True)
        comparendo = Column(String, nullable=True)
        fecha_comparendo = Column(DateTime, nullable=True)
        año_comparendo = Column(String, nullable=True)
        prescripcion = Column(String, nullable=False)
        tipo_comparendo = Column(String, nullable=False)
        clase_vehiculo = Column(String, nullable=True)
        tipo = Column(String, nullable=False)
        servicio_vehiculo = Column(String, nullable=True)
        valor_pagado = Column(String, nullable=False)
        fecha_distribucion = Column(DateTime, nullable=False)
        resolucion_mp = Column(String, nullable=True)
        valor_inicial_cargado = Column(String, nullable=True)
        concepto = Column(String, nullable=False)
        estado_cartera = Column(String, nullable=True)
        concepto_principal = Column(String, nullable=True)
        gestion = Column(String, nullable=False)
        descuento_cartera = Column(String, nullable=True)
        descuento_de_intereses = Column(String, nullable=True)
        cantidad_de_descuento_cartera = Column(Integer, nullable=True)
        cantidad_de_descuento_de_intereses = Column(Integer, nullable=True)
        resolucion_sancion = Column(String, nullable=True)
        fecha_resolucion_sancion = Column(DateTime, nullable=True)
        valor_pagado_de_intereses = Column(String, nullable=True)
        created_at = Column(DateTime, default=func.now())
    
    class Auditoria(Base):
        __tablename__ = "auditoria"

        id = Column(Integer, primary_key=True, autoincrement=True)
        archivo_id = Column(String, ForeignKey("archivos.archivo_id"), nullable=False)
        columna = Column(String, nullable=False)
        created_at = Column(DateTime, default=func.now())

    Base.metadata.create_all(engine)
    logger.info("Tablas creadas exitosamente")
    Session = sessionmaker(bind=engine)
    session = Session()
    
except Exception as e:
    logger.error(f"Error de conexión o creación de BD: {str(e)}")
    logger.info("Verifica:")
    logger.info("   - PostgreSQL está corriendo (por defecto puerto 5432)")
    logger.info("   - La BD 'transito' existe")
    logger.info("   - Usuario 'postgres' y contraseña son correctos")
    logger.info(f"   - URL: {DATABASE_URL}")
    engine = None
    session = None
    Base = declarative_base()
