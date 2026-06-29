from msal import ConfidentialClientApplication
import requests
import os
import base64
from dotenv import load_dotenv
import time
import re
from datetime import datetime
import pytesseract
import sys
import pdfplumber
from pdf2image import convert_from_path

load_dotenv()

TEMP_PATH = r"C:\Temp\Boletos"
os.makedirs(TEMP_PATH, exist_ok=True)

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


pytesseract.pytesseract.tesseract_cmd = os.path.join(BASE_DIR, "Tesseract-OCR", "tesseract.exe")


# PUXA IDs PARA AUTENTICAÇÃO
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
EMAIL = os.getenv("EMAIL")

MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "março": 3,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12
}


def tokens():

    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )

    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in token:
        raise Exception(token)

    return token["access_token"]

def headers():
    token = tokens()

    return{
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }



def listar_emails():

    url = (
        f"https://graph.microsoft.com/v1.0/users/{EMAIL}/"
        "mailFolders/Inbox/messages"
        "?$orderby=receivedDateTime desc"
        "&$top=5"
    )

    resposta = requests.get(url, headers=headers())

    print(resposta.json())
    return resposta.json()["value"]


def obter_email(message_id):

    url = (
        f"https://graph.microsoft.com/v1.0/users/{EMAIL}"
        f"/messages/{message_id}"
    )

    resposta = requests.get(
        url,
        headers=headers()
    )

    return resposta.json()




def baixar_pdfs(message_id):

    url = (
        f"https://graph.microsoft.com/v1.0/users/{EMAIL}"
        f"/messages/{message_id}/attachments"
    )

    resposta = requests.get(
        url,
        headers=headers()
    ).json()

    arquivos = []

    for anexo in resposta["value"]:

        if not anexo["name"].lower().endswith(".pdf"):
            continue

        caminho = os.path.join(TEMP_PATH, anexo["name"])

        with open(caminho, "wb") as f:
            f.write(base64.b64decode(anexo["contentBytes"]))

        arquivos.append(caminho)

    return arquivos


def extrair_qualquer_data(texto):
    """
    Extrai a primeira data encontrada nos formatos:
    dd/mm/yyyy, dd-mm-yyyy, dd mm yyyy,
    dd/mm/yy, dd-mm-yy, dd mm yy
    """
    texto = texto.replace("\n", " ")

    padrao = r"\b(\d{2}[\/\-\s]\d{2}[\/\-\s]\d{2,4})\b"

    for data_str in re.findall(padrao, texto):
        data_str = data_str.strip()

        # Normaliza separadores
        data_str = re.sub(r"[\/\-\s]", "/", data_str)

        for fmt in ("%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(data_str, fmt)
            except ValueError:
                pass

    return None

def extrair_validade(texto):
    texto = texto.lower().replace("\n", " ")

    datas = []

    palavras_chave = (
        r"vencimento|vence|pagamento\s*até|validade|"
        r"encimento|ence|vcto|data\s*vencimento|pagamento|pgto|venc"
    )

    padrao_numerico = re.compile(
        rf"({palavras_chave}).{{0,200}}?"
        r"(\d{2}[\/\-\s]\d{2}[\/\-\s]\d{2,4})",
        re.IGNORECASE
    )

    for m in padrao_numerico.finditer(texto):
        data_str = re.sub(r"[\/\-\s]", "/", m.group(2).strip())

        for fmt in ("%d/%m/%Y", "%d/%m/%y"):
            try:
                datas.append(datetime.strptime(data_str, fmt))
                break
            except ValueError:
                pass


    meses_regex = (
        r"janeiro|fevereiro|março|marco|abril|maio|junho|"
        r"julho|agosto|setembro|outubro|novembro|dezembro"
    )

    padrao_extenso = re.compile(
        rf"({palavras_chave}).{{0,200}}?"
        rf"(\d{{1,2}}(?:º)?\s*(?:de\s+)?"
        rf"({meses_regex})"
        rf"\s*(?:de\s+)?(\d{{2,4}}))",
        re.IGNORECASE
    )

    for m in padrao_extenso.finditer(texto):
        try:
            dia = int(re.search(r"\d{1,2}", m.group(2)).group())

            mes = MESES[m.group(3).lower()]

            ano = int(re.search(r"\d{2,4}$", m.group(2)).group())

            if ano < 100:
                ano += 2000

            datas.append(datetime(ano, mes, dia))

        except Exception:
            pass

    return min(datas) if datas else None

def extrair_texto_pdf(caminho_pdf):
    texto_pdf = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                conteudo = pagina.extract_text()
                if conteudo:
                    texto_pdf += conteudo + "\n"
    except Exception as erro:
        print(f"Erro leitura PDF: {erro}")
    return texto_pdf


def executar_ocr(caminho_pdf):

    imagens = convert_from_path(
        caminho_pdf,
        dpi=500
    )

    texto = ""

    for img in imagens:

        texto += pytesseract.image_to_string(
            img,
            lang="por"
        )

    return texto

def descobrir_data(email):

    assunto = email["subject"]

    data = extrair_qualquer_data(assunto)

    if data:
        return data

    corpo = email["body"]["content"]

    data = extrair_validade(corpo)

    if data:
        return data

    pdfs = baixar_pdfs(email["id"])

    datas = []

    for pdf in pdfs:

        texto = extrair_texto_pdf(pdf)

        validade = extrair_validade(texto)

        if not validade:

            texto = executar_ocr(pdf)

            validade = extrair_validade(texto)

        if validade:
            datas.append(validade)

        os.remove(pdf)

    if datas:
        return min(datas)

    return None

def processar_email(message_id):

    email = obter_email(message_id)

    assunto = email["subject"]

    if re.match(r"^\d{2}/\d{2}/\d{4}", assunto):
        return

    data = descobrir_data(email)

    if not data:
        return

    if data.date() <= datetime.now().date():

        novo_assunto = (
            f"{data:%d/%m/%Y} - Atenção - {assunto}"
        )

    else:

        novo_assunto = (
            f"{data:%d/%m/%Y} - {assunto}"
        )

    atualizar_assunto(
        message_id,
        novo_assunto
    )

def atualizar_assunto(message_id, novo_assunto):

    url = (
        f"https://graph.microsoft.com/v1.0/users/{EMAIL}"
        f"/messages/{message_id}"
    )

    requests.patch(
        url,
        headers=headers(),
        json={
            "subject": novo_assunto
        }
    )

def last_email():
    emails_processados = set()

    while True:
        emails = listar_emails()

        for email in emails:
            id_email= email["id"]

            if id_email in emails_processados:
                continue

            emails_processados.add(id_email)
            
            processar_email(id_email)
        
        time.sleep(3)

if __name__ == "__main__":
    last_email()