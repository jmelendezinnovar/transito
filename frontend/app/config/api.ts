type ApiEnv = ImportMetaEnv & {
  API_URL?: string;
  VITE_API_URL?: string;
};

const env = import.meta.env as ApiEnv;

export const API_BASE_URL = env.API_URL ?? env.VITE_API_URL ?? "http://localhost:8000";

export const API_ROUTES = {
  archivos: `${API_BASE_URL}/api/archivos`,
  archivoFlujo: (archivoId: string) => `${API_BASE_URL}/api/archivos/${archivoId}/flujo`,
  archivoDetalles: (archivoId: string) => `${API_BASE_URL}/api/archivos/${archivoId}/detalles`,
} as const;