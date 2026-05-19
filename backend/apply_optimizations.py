"""
Utilidad para aplicar optimizaciones de BD automáticamente
Ejecutar: python apply_optimizations.py
"""

import os
import logging
from backend.models import engine
from sqlalchemy import text

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_database_indexes():
    """Crea los índices de optimización en PostgreSQL"""
    indexes = [
        # CARTERAS
        """CREATE INDEX IF NOT EXISTS idx_cartera_codigo_org 
           ON carteras(codigo, organismo)
           WHERE estado_cartera_final IS NOT NULL""",
        
        """CREATE INDEX IF NOT EXISTS idx_cartera_tipo_org 
           ON carteras(tipo_cartera, organismo)""",
        
        """CREATE INDEX IF NOT EXISTS idx_cartera_organismo 
           ON carteras(organismo)""",
        
        # RECAUDOS
        """CREATE INDEX IF NOT EXISTS idx_recaudo_organismo_tipo 
           ON recaudos(organismo, tipo_recaudo)""",
        
        """CREATE INDEX IF NOT EXISTS idx_recaudo_archivo 
           ON recaudos(archivo_id)""",
        
        # ARCHIVOS
        """CREATE INDEX IF NOT EXISTS idx_archivo_id 
           ON archivos(archivo_id)""",
        
        # AUDITORIA
        """CREATE INDEX IF NOT EXISTS idx_auditoria_archivo 
           ON auditoria(archivo_id)""",
    ]
    
    if not engine:
        logger.error("❌ Engine de BD no disponible")
        return False
    
    try:
        with engine.begin() as conn:
            for idx_sql in indexes:
                conn.execute(text(idx_sql))
                logger.info(f"✓ Índice creado: {idx_sql.split('CREATE INDEX')[1][:50]}...")
            
            # Recolectar estadísticas
            conn.execute(text("ANALYZE carteras"))
            conn.execute(text("ANALYZE recaudos"))
            conn.execute(text("ANALYZE archivos"))
            
        logger.info("✅ Todos los índices creados exitosamente")
        return True
        
    except Exception as e:
        logger.error(f"❌ Error creando índices: {str(e)}")
        return False

def get_database_stats():
    """Obtiene estadísticas de tamaño de tablas"""
    if not engine:
        return {}
    
    try:
        with engine.connect() as conn:
            stats = {}
            for table in ["carteras", "recaudos", "archivos"]:
                result = conn.execute(
                    text(f"""
                        SELECT 
                            relname,
                            pg_size_pretty(pg_total_relation_size(relid)) as tamaño,
                            n_live_tup as registros
                        FROM pg_stat_user_tables
                        WHERE relname = '{table}'
                    """)
                ).fetchone()
                
                if result:
                    stats[table] = {
                        "tamaño": result[1],
                        "registros": result[2]
                    }
            
            return stats
    except Exception as e:
        logger.error(f"Error obteniendo estadísticas: {str(e)}")
        return {}

if __name__ == "__main__":
    logger.info("🚀 Aplicando optimizaciones de BD...")
    
    # Crear índices
    if create_database_indexes():
        logger.info("\n📊 Estadísticas de BD:")
        stats = get_database_stats()
        for tabla, datos in stats.items():
            logger.info(f"  {tabla}: {datos['registros']:,} registros ({datos['tamaño']})")
    
    logger.info("\n✅ Optimizaciones completadas")
