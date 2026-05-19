"""
Microsoft Graph Subscriptions Manager
Automatiza la creación y renovación de suscripciones a cambios en recursos de SharePoint
usando Microsoft Graph API.
"""

import requests
import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv
import os

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

load_dotenv()

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
NOTIFICATION_URL = os.getenv('NOTIFICATION_URL')
DOMINIO = os.getenv('DOMINIO')
NOMBRE_SITIO = os.getenv('NOMBRE_SITIO')
GRAPH_RESOURCE = os.getenv('GRAPH_RESOURCE')
CLIENT_STATE = os.getenv('GRAPH_CLIENT_STATE', 'SecretToken123')
DB_FILE = Path(os.getenv('GRAPH_DB_FILE', 'subscription_info.json'))

REQUIRED_VARS = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'NOTIFICATION_URL', 'DOMINIO', 'NOMBRE_SITIO']
for var in REQUIRED_VARS:
    if not locals()[var]:
        logger.error(f"Variable de entorno requerida no configurada: {var}")
        raise ValueError(f"Falta configurar: {var}")

class GraphSubscriptionManager:
    """Gestor de suscripciones a Microsoft Graph"""
    
    BASE_URL = "https://graph.microsoft.com/v1.0"
    GRAPH_SCOPE = "https://graph.microsoft.com/.default"
    SUBSCRIPTION_EXPIRY_DAYS = 2
    TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    def __init__(self):
        self.token = None
        self.token_timestamp = None
        self.resource = GRAPH_RESOURCE if GRAPH_RESOURCE and "{site-id}" not in GRAPH_RESOURCE else None

    def resolve_resource(self, token: str) -> str:
        """
        Resuelve el recurso de Graph a partir del sitio de SharePoint.

        Si GRAPH_RESOURCE está definido, se usa como fallback. Si no, se consulta
        el sitio en Microsoft Graph para obtener el site_id y construir el resource.
        """
        if self.resource:
            logger.info("Usando GRAPH_RESOURCE desde variables de entorno")
            return self.resource

        logger.info("Resolviendo resource desde DOMINIO y NOMBRE_SITIO...")
        url = f"https://graph.microsoft.com/v1.0/sites/{DOMINIO}:/sites/{NOMBRE_SITIO}:"
        headers = {'Authorization': f'Bearer {token}'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        site_data = response.json()
        site_id = site_data.get('id')
        if not site_id:
            raise ValueError("No se pudo obtener site_id del sitio de SharePoint")

        self.resource = f"sites/{site_id}/drive/root"
        logger.info("Resource resuelto: %s", self.resource)
        return self.resource
    
    def get_token(self) -> Optional[str]:
        """
        Obtiene un token de acceso de Azure AD.
        
        Returns:
            str: Token de acceso válido
            
        Raises:
            requests.RequestException: Si hay error al obtener el token
        """
        try:
            logger.info("Obteniendo token de acceso...")
            
            data = {
                'grant_type': 'client_credentials',
                'client_id': CLIENT_ID,
                'client_secret': CLIENT_SECRET,
                'scope': self.GRAPH_SCOPE
            }
            
            response = requests.post(self.TOKEN_URL, data=data, timeout=10)
            response.raise_for_status()
            
            token = response.json().get('access_token')
            if not token:
                logger.error("No se obtuvo token en la respuesta")
                raise ValueError("Token vacío en respuesta")
            
            logger.info("Token obtenido exitosamente")
            self.token = token
            self.token_timestamp = datetime.now(timezone.utc)
            return token
            
        except requests.RequestException as e:
            logger.error(f"Error al obtener token: {e}")
            raise
    
    def save_sub_info(self, sub_id: str, sub_data: dict) -> None:
        """
        Guarda información de la suscripción localmente.
        
        Args:
            sub_id: ID de la suscripción
            sub_data: Datos completos de la suscripción
        """
        try:
            data = {
                "subscriptionId": sub_id,
                "createdAt": datetime.now(timezone.utc).isoformat(),
                "resource": self.resource,
                "notificationUrl": NOTIFICATION_URL,
                **sub_data
            }
            
            with open(DB_FILE, 'w') as f:
                json.dump(data, f, indent=2)
            
            logger.info(f"Información de suscripción guardada: {sub_id}")
            
        except IOError as e:
            logger.error(f"Error guardando información de suscripción: {e}")
            raise
    
    def load_sub_info(self) -> Optional[dict]:
        """
        Carga información de suscripción guardada localmente.
        
        Returns:
            dict: Información de suscripción o None si no existe
        """
        try:
            if DB_FILE.exists():
                with open(DB_FILE, 'r') as f:
                    data = json.load(f)
                logger.info(f"Información de suscripción cargada")
                return data
            else:
                logger.info(f"Archivo de suscripción no encontrado: {DB_FILE}")
                return None
                
        except (IOError, json.JSONDecodeError) as e:
            logger.error(f"Error cargando información de suscripción: {e}")
            return None
    
    def create_subscription(self, token: str) -> bool:
        """
        Crea una nueva suscripción en Microsoft Graph.
        
        Args:
            token: Token de acceso de Azure AD
            
        Returns:
            bool: True si la suscripción fue creada exitosamente
        """
        try:
            logger.info("Creando nueva suscripción...")
            resource = self.resolve_resource(token)
            
            # Calcular fecha de expiración
            expiry = (
                datetime.now(timezone.utc) + 
                timedelta(days=self.SUBSCRIPTION_EXPIRY_DAYS)
            ).isoformat().replace("+00:00", "Z")
            
            url = f"{self.BASE_URL}/subscriptions"
            
            payload = {
                "changeType": "updated",
                "notificationUrl": NOTIFICATION_URL,
                "resource": resource,
                "expirationDateTime": expiry,
                "clientState": CLIENT_STATE
            }
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            response = requests.post(url, headers=headers, json=payload, timeout=10)
            response.raise_for_status()
            
            if response.status_code == 201:
                sub_data = response.json()
                new_id = sub_data.get('id')
                self.save_sub_info(new_id, sub_data)
                logger.info(f"Suscripción creada exitosamente. ID: {new_id}")
                return True
            else:
                logger.error(f"Código de estado inesperado: {response.status_code}")
                return False
                
        except requests.RequestException as e:
            logger.error(f"Error al crear suscripción: {e}")
            return False
    
    def renew_subscription(self, token: str, sub_id: str) -> bool:
        """
        Renueva una suscripción existente.
        
        Args:
            token: Token de acceso de Azure AD
            sub_id: ID de la suscripción a renovar
            
        Returns:
            bool: True si la renovación fue exitosa
        """
        try:
            logger.info(f"Intentando renovar suscripción {sub_id}...")
            
            new_expiry = (
                datetime.now(timezone.utc) + 
                timedelta(days=self.SUBSCRIPTION_EXPIRY_DAYS)
            ).isoformat().replace("+00:00", "Z")
            
            url = f"{self.BASE_URL}/subscriptions/{sub_id}"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            payload = {"expirationDateTime": new_expiry}
            
            response = requests.patch(url, headers=headers, json=payload, timeout=10)
            
            if response.status_code == 200:
                logger.info("Suscripción renovada correctamente.")
                self.save_sub_info(sub_id, response.json())
                return True
                
            elif response.status_code == 404:
                logger.warning(
                    f"Suscripción {sub_id} no encontrada. "
                    "Será creada una nueva."
                )
                return False
                
            else:
                logger.error(
                    f"Error inesperado ({response.status_code}): {response.text}"
                )
                return False
                
        except requests.RequestException as e:
            logger.error(f"Error al renovar suscripción: {e}")
            return False
    
    def manage_subscription(self) -> bool:
        """
        Gestiona el ciclo de vida de las suscripciones.
        Intenta renovar la existente o crea una nueva si no existe.
        
        Returns:
            bool: True si la operación fue exitosa
        """
        try:
            # Obtener token
            token = self.get_token()
            if not token:
                logger.error("No se pudo obtener token de acceso")
                return False
            
            # Cargar información de suscripción existente
            sub_info = self.load_sub_info()

            if sub_info and sub_info.get('resource') and not self.resource:
                self.resource = sub_info.get('resource')
            
            if not sub_info:
                logger.info("No hay suscripción existente. Creando nueva...")
                return self.create_subscription(token)
            
            # Intentar renovar suscripción existente
            sub_id = sub_info.get('subscriptionId')
            if not sub_id:
                logger.error("ID de suscripción inválido en archivo local")
                return self.create_subscription(token)
            
            renewed = self.renew_subscription(token, sub_id)
            
            if not renewed:
                # Si falla la renovación, crear nueva
                return self.create_subscription(token)
            
            return True
            
        except Exception as e:
            logger.error(f"Error crítico en manage_subscription: {e}")
            return False
    
    def get_subscription_status(self, token: str, sub_id: str) -> Optional[dict]:
        """
        Obtiene el estado de una suscripción existente.
        
        Args:
            token: Token de acceso de Azure AD
            sub_id: ID de la suscripción
            
        Returns:
            dict: Información de la suscripción o None si hay error
        """
        try:
            url = f"{self.BASE_URL}/subscriptions/{sub_id}"
            headers = {'Authorization': f'Bearer {token}'}
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            return response.json()
            
        except requests.RequestException as e:
            logger.error(f"Error obteniendo estado de suscripción: {e}")
            return None
    
    def delete_subscription(self, token: str, sub_id: str) -> bool:
        """
        Elimina una suscripción.
        
        Args:
            token: Token de acceso de Azure AD
            sub_id: ID de la suscripción a eliminar
            
        Returns:
            bool: True si se eliminó exitosamente
        """
        try:
            logger.info(f"Eliminando suscripción {sub_id}...")
            
            url = f"{self.BASE_URL}/subscriptions/{sub_id}"
            headers = {'Authorization': f'Bearer {token}'}
            
            response = requests.delete(url, headers=headers, timeout=10)
            
            if response.status_code == 204:
                logger.info("Suscripción eliminada exitosamente")
                DB_FILE.unlink(missing_ok=True)
                return True
            else:
                logger.error(f"Error al eliminar: {response.status_code}")
                return False
                
        except requests.RequestException as e:
            logger.error(f"Error eliminando suscripción: {e}")
            return False


def main():
    """Función principal para ejecutar el gestor de suscripciones"""
    manager = GraphSubscriptionManager()
    success = manager.manage_subscription()
    
    if not success:
        logger.error("La operación no se completó exitosamente")
        exit(1)
    
    logger.info("Operación completada exitosamente")


if __name__ == "__main__":
    main()
