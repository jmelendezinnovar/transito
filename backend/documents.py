from enum import Enum

class Document(Enum):
    RECAUDO_MULTAS = "/RECAUDO/MULTAS/"
    RECAUDO_DERECHOS_DE_TRANSITO = "/RECAUDO/DERECHOS DE TRANSITO/"
    CARTERA_MULTAS = "/CARTERA/MULTAS/"
    CARTERA_DERECHOS_DE_TRANSITO = "/CARTERA/DERECHOS DE TRANSITO/"

class WalletStatus(Enum):
    ACTIVE = "ACTIVO"
    INACTIVE = "INACTIVO"

class EjecucionEstado(Enum):
    NO_INICIADA = "No iniciada"
    PENDIENTE = "Pendiente"
    EN_PROCESO = "Procesando"
    COMPLETADA = "Completado"
    FALLIDA = "Fallido"

class EtapaEstado(Enum):
    NO_INICIADO = "No iniciado"
    PENDIENTE = "Pendiente"
    EN_PROCESO = "Procesando"
    COMPLETADO = "Completado"
    FALLIDO = "Fallido"

class EtapaNombre(Enum):
    EXTRACCION = "Extrayendo"
    LIMPIEZA = "Limpiando"
    GUARDADO = "Guardando"

class TransitoNombre(Enum):
    BOLIVAR = "BOLIVAR"
    CLEMENCIA = "CLEMENCIA"
    SAN_JUAN = "SAN JUAN"
    TURBACO = "TURBACO"
    VILLA_DEL_ROSARIO = "VILLA DEL ROSARIO"