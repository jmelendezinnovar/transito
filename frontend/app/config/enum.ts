export enum EtapaNombre {
    EXTRACCION = "Extrayendo",
    LIMPIEZA = "Limpiando",
    GUARDADO = "Guardando"
}

export function getEtapaNombreDisplay(etapa: EtapaNombre): string {
    switch (etapa) {
        case EtapaNombre.EXTRACCION:
            return "Extracción de Datos";
        case EtapaNombre.LIMPIEZA:
            return "Limpieza de Datos";
        case EtapaNombre.GUARDADO:
            return "Guardado de Registros";
        default:
            return etapa;
    }
}