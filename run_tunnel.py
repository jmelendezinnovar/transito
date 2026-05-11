"""Levanta la API local y crea un túnel temporal público con ngrok."""

import logging
import os
import sys

import uvicorn
from dotenv import load_dotenv
from pyngrok import ngrok

from main import app


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()


def main() -> None:
    authtoken = os.getenv("NGROK_AUTHTOKEN")
    if not authtoken:
        logger.error(
            "Falta NGROK_AUTHTOKEN en .env. Crea una cuenta en ngrok y pega tu token ahí."
        )
        sys.exit(1)

    ngrok.set_auth_token(authtoken)
    public_url = ngrok.connect(8000, "http")
    logger.info("Túnel activo: %s", public_url)
    logger.info("Webhook público: %s/webhook", public_url)

    try:
        uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
    finally:
        ngrok.kill()


if __name__ == "__main__":
    main()