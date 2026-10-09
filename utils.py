import os
import re
import pandas as pd
from datetime import datetime

try:
    import keyring
    HAS_KEYRING = True
except ImportError:
    HAS_KEYRING = False

try:
    import fitz  
    import pytesseract
    from PIL import Image, ImageTk 
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import pyautogui
    import pyperclip
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    HAS_RPA = True
except ImportError:
    HAS_RPA = False

VERSAO_ATUAL = "1.0.0"
GITHUB_REPO = "JonatasManasses/HubAutomacao" 
APP_NAME = "Hub de Automação"
ROOT_DIR_NFTS = "C:/RoboNFTS"
AIRTABLE_BASE_ID = "appumhEucykyJuO7G"
AIRTABLE_TABLE_ID = "tblPhPCaNhd0ktJu4" 
AIRTABLE_CIDADES_TABLE_ID = "tblMpIjendFVYT97c"
AIRTABLE_VIEW_ID = "" 

COLUNAS_AIRTABLE = [
    "Razão Social", "CNPJ", "Apelido", "Código", 
    "Status da Empresa", "Tipo de Tributação", "CCM", "Cidade da Empresa", "Mensalidade",
    "Data de Abertura", "Data de Inativação"
]

def limpar_cnpj(texto_cnpj):
    num = re.sub(r'\D', '', str(texto_cnpj))
    if 11 < len(num) < 14: return num.zfill(14)
    elif 0 < len(num) < 11: return num.zfill(11)
    return num

def limpar_nome_arquivo(nome):
    if not nome: return ""
    return re.sub(r'[\\/*?:"<>|]', "", str(nome)).strip()

def seguro_float(v):
    if pd.isna(v): return 0.0
    s = str(v).replace('R$', '').strip()
    if ',' in s and '.' in s: s = s.replace('.', '').replace(',', '.')
    elif ',' in s: s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def formatar_data_br(d):
    if pd.isna(d) or not str(d).strip() or str(d).strip().lower() == 'nan': return ""
    if isinstance(d, (datetime, pd.Timestamp)): return d.strftime('%d/%m/%Y')
    d_str = str(d).strip().split(' ')[0]
    if re.match(r'^\d{2}-\d{2}-\d{4}$', d_str): d_str = d_str.replace('-', '/')
    if re.match(r'^\d{2}/\d{2}/\d{4}$', d_str): return d_str
    if re.match(r'^\d{4}-\d{2}-\d{2}$', d_str):
        ano, mes, dia = d_str.split('-')
        return f"{dia}/{mes}/{ano}"
    try: return pd.to_datetime(d_str, dayfirst=True).strftime('%d/%m/%Y')
    except: return d_str

def obter_token():
    if HAS_KEYRING:
        try: return keyring.get_password(APP_NAME, "AirtableToken")
        except: return None
    return None

def obter_certificado_padrao():
    if HAS_KEYRING:
        try: return keyring.get_password(APP_NAME, "CertificadoPadrão")
        except: return None
    return None

def validar_tributacao(tipo_guia, tipo_tributacao):
    tipo_tributacao = str(tipo_tributacao).strip().lower()
    if tipo_guia == "DAS" or tipo_guia == "Declaração PGDAS-D":
        validos_das = ["simples nacional fator r", "simples nacional anexo iii", "simples nacional anexo lll", "simples nacional anexo iv", "simples nacional anexo v"]
        return any(t in tipo_tributacao for t in validos_das)
    elif tipo_guia == "ISS": return "lucro presumido" in tipo_tributacao
    return True