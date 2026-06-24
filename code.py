import pdfplumber
import pytesseract
import re
import os
import time
import warnings
import threading
import queue
from pdf2image import convert_from_path
from datetime import datetime
from PIL import Image, ImageOps
import sys
from msal import ConfidentialClientApplication
import requests
from dotenv import load_dotenv

load_dotenv()

warnings.filterwarnings("ignore")
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

pytesseract.pytesseract.tesseract_cmd = os.path.join(BASE_DIR, "Tesseract-OCR", "tesseract.exe")

TEMP_PATH = r"C:\Temp\Boletos"

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=authority,
    client_credential=CLIENT_SECRET
)

token = app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)
print(token)

ACCESS_TOKEN = token["access_token"]

HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}


def buscar_emails():
    url = (
        "https://graph.microsoft.com/v1.0/users/"
        "pedro.castro@apsa.com.br/messages"
        "?$top=20"
    )

    resp = requests.get(url, headers=HEADERS)

    dados = resp.json()

    print(dados)

    if "value" in dados:
        return dados["value"]

    return dados
    

print(buscar_emails())