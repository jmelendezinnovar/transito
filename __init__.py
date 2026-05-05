import msal
import requests
import os
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DOMINIO = os.getenv("DOMINIO")
NOMBRE_SITIO = os.getenv("NOMBRE_SITIO")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

token_result = app.acquire_token_for_client(SCOPE)
access_token = token_result["access_token"]

print("Token obtenido correctamente")

url = f"https://graph.microsoft.com/v1.0/sites/{DOMINIO}:/sites/{NOMBRE_SITIO}"
headers = {"Authorization": f"Bearer {access_token}"}
resp = requests.get(url, headers=headers)
print("SITE ID")
print(resp.json())

url = f"https://graph.microsoft.com/v1.0/sites/{resp.json()}/drives"
resp = requests.get(url, headers=headers)
print(resp.json())