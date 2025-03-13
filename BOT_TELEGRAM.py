import os
import json
import logging
import aiohttp
import requests
import datetime
import time
import telegram
import xml.etree.ElementTree as ET
from openpyxl import Workbook

from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes, MessageHandler, filters

import xml.etree.ElementTree as ET
from openpyxl import Workbook
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import zipfile
import os


__version__ = "0.1.4"
__author__ = "Alexandre B, J. Ayrton"
__credits__ = "Anderson, Josimar"

FileName = "WebHook.json"

# #BOT-ALEXO "7820581372:AAH8OLVNA6KDkFqIGTDJx2IEm_HDOH_bRcs" #ALEXO
# #DKT_TESTE "7829967937:AAG7ZajDn7lYbeIllyQJKYa1V5q8AbFDeMk" #DKT_TESTE

BOT_TOKEN = "7820581372:AAH8OLVNA6KDkFqIGTDJx2IEm_HDOH_bRcs"

TELEGRAM_GROUP_ID = "-1002292627707"


class IgnoreAttributeErrorFilter(logging.Filter):
    def filter(self, record):
        return "AttributeError" not in record.getMessage()

def send_log_to_telegram(message):
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendMessage'
    payload = {
        'chat_id': TELEGRAM_GROUP_ID,
        'text': message,
        'parse_mode': 'Markdown'
    }
    requests.post(url, json=payload)


logger = logging.getLogger()
logger.setLevel(logging.INFO)

file_handler = logging.FileHandler('LOGGING_EXEC.txt')
file_handler.setLevel(logging.INFO)
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(console_formatter)

file_handler.addFilter(IgnoreAttributeErrorFilter())
console_handler.addFilter(IgnoreAttributeErrorFilter())

logger.addHandler(file_handler)
logger.addHandler(console_handler)

ErroE101 = "❌ Atenção, excesso de argumentos. Verifique o comando informado e tente novamente!"
ErroP101 = "❌ Atenção, 'POP' não informado!"
ErroP102 = "❌ Atenção, 'POP' não existe na lista de templates. Verifique se foi informado corretamente ou notifique a equipe interna."
ErroF101 = "❌ Atenção, 'FSAN/SN' não informado para a consulta. Verifique o comando e tente novamente!"
ErroF102 = "❌ Atenção, o formato do campo 'FSAN/SN' está incorreto!"
ErroS101 = "❌ Atenção, 'SPLITTER' não informado. Verifique o comando e tente novamente!"
ErroN101 = "❌ Atenção, 'OLT/SLOT/PON' não informado. Verifique o comando e tente novamente!"
ErroN102 = "❌ Atenção, 'OLT/SLOT/PON' contém mais de duas '/'. Verifique o comando e tente novamente!"
ErroC101 = "❌ Atenção, verifique se a 'CTO' informada está correta e tente novamente."


class TelegramHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        send_log_to_telegram(log_entry)

telegram_handler = TelegramHandler()
telegram_handler.setLevel(logging.INFO)
telegram_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
telegram_handler.setFormatter(telegram_formatter)

logger.addHandler(telegram_handler)

logging.getLogger("aiohttp").setLevel(logging.WARNING)
logging.getLogger("telegram").setLevel(logging.WARNING)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)

SPLITTERS_VALIDOS = {"1/16", "1/8", "1/4"}

def kml_to_xlsx(kml_file, xlsx_file):
    tree = ET.parse(kml_file)
    root = tree.getroot()
    namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}

    wb = Workbook()
    ws = wb.active
    ws.title = "Placemarks"
    ws.append(["Nome do Placemark", "Longitude", "Latitude"])

    for placemark in root.findall(".//kml:Placemark", namespaces):
        name = placemark.find("kml:name", namespaces)
        coordinates = placemark.find(".//kml:coordinates", namespaces)
        
        if name is not None and coordinates is not None:
            coord_text = coordinates.text.strip()
            coord_parts = coord_text.split(",")
            
            if len(coord_parts) >= 2:
                lon, lat = coord_parts[:2]
                ws.append([name.text, lon.strip(), lat.strip()])
    
    wb.save(xlsx_file)

def ListaCidades():
    try:
        with open(FileName, 'r', encoding='utf-8') as file:
            dados = json.load(file)

        cidades = [f"{i + 1}. {entry['POP']} - {entry['CIDADE']}" for i, entry in enumerate(dados)]
        return "\n".join(cidades)
    
    except FileNotFoundError:
        return "Arquivo não encontrado."
    
    except json.JSONDecodeError:
        return "Erro ao ler o arquivo JSON."

def buscar_webhook_por_pop(pop):
    try:
        with open(FileName, 'r', encoding='utf-8') as arquivo_json:
            dados = json.load(arquivo_json)

        for entry in dados:
            if entry["POP"].upper() == pop.upper():
                return entry["WEBHOOK_LINK"]

        return None
    
    except FileNotFoundError:
        return "Arquivo não encontrado."
    
    except json.JSONDecodeError:
        return "Erro ao ler o arquivo JSON."

def kml_to_xlsx(kml_file, xlsx_file):
    tree = ET.parse(kml_file)
    root = tree.getroot()
    namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}

    wb = Workbook()
    ws = wb.active
    ws.title = "Placemarks"
    ws.append(["Id CTOs", "Latitude", "Longitude"])

    for placemark in root.findall(".//kml:Placemark", namespaces):
        name = placemark.find("kml:name", namespaces)
        coordinates = placemark.find(".//kml:coordinates", namespaces)
        
        if name is not None and coordinates is not None:
            coord_text = coordinates.text.strip()
            coord_parts = coord_text.split(",")
            
            if len(coord_parts) >= 2:
                lat, lon = coord_parts[:2]
                ws.append([name.text, lon.strip(), lat.strip()])
    
    wb.save(xlsx_file)


def extract_kml_from_kmz(kmz_file, extract_to):
    with zipfile.ZipFile(kmz_file, 'r') as kmz:
        for file in kmz.namelist():
            if file.endswith('.kml'):
                kmz.extract(file, extract_to)
                kml_file = os.path.join(extract_to, file)
                new_kml_file = os.path.join(extract_to, os.path.splitext(os.path.basename(kmz_file))[0] + '.kml')
                os.rename(kml_file, new_kml_file)
                return new_kml_file
    return


async def ajuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    comandos = [
        "| Ajuda - BOT-ALEXO",
        "\n\n- Atividades 🌟",
        "    /Atividades <POP>",
        "    Verifica se existem atividades pendentes no template.",
        "    EX: /Atividades TIE",

        "\n\n- Checar 🔍",
        "    /Checar <CTO> <FSAN>",
        "    Verifica OLT/SLOT/PON do cliente na CTO.",
        "    EX: /Checar TIE-001 FHTT0000000",

        "\n\n- Localizar 📍",
        "    /Localizar <CTO>",
        "    Retorna a localização de uma CTO.",
        "    EX: /Localizar TIE-001",

        "\n\n- Input 📝",
        "    /Input <CTO> <SPLITER>",
        "    Inputa as informações de data e splitter para o template.",
        "    EX: /Input TIE-001 1/16",

        "\n\n- Insert 📝",
        "    /Insert <CTO> <OLT/SLOT/PON>",
        "    Inputa as informações da CTO e splitter para o template na aba checar.",
        "    EX: /Insert TIE-001 1/1/1",

        "\n\n- Listar IDS 📋",
        "    /ListarIDs <POP> <OLT/SLOT/PON>",
        "    /TESTE",
        "    EX: /Input TIE-001 1/16",

        "\n\n- Nova CTO ➕",
        "    /NovaCTO <POP> <OLT/SLOT/PON> <SPLITER>",
        "    CTO QUE NÃO EXISTE NO KMZ.",
        "    EX: /NovaCTO TIE 1/1/1 1/16",

        "\n\n| Informações ℹ️:",
        f"    Versão: {__version__}",
        f"    Criadores: {__author__}",
        f"    Créditos:   {__credits__}"
    ]

    comandos_texto = "\n".join(comandos)

    logger.info(f"/Ajuda - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")

    await context.bot.send_message(chat_id=chat_id, text=comandos_texto)


async def fetch_data(webhook_link, payload):
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(webhook_link, json=payload) as response:
                if response.status == 200:
                    response_data = await response.json()
                    logger.info(f"Google App Script - Resposta: {response_data}")
                    return response_data
                else:
                    logger.error(f"Erro ao conectar ao Apps Script: {response.status} - {response.reason}")
                    return {"status": "error", "message": "Erro ao conectar ao servidor."}
    except Exception as e:
        logger.error(f"/fetch_data - Exceção ao acessar o Apps Script: {e}")
        return {"status": "error", "message": str(e)}


async def atividades(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return

    pop = context.args[0]
    pop = str(pop.upper())
    pop = pop.split('-')[0]

    webhook_link = buscar_webhook_por_pop(pop)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return

    payload = {"comando": "Atividades", "id": chat_id}

    logger.info(f"RECEBIDO: /Atividades - POP:{pop} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")

    try:
        data = await fetch_data(webhook_link, payload)
    except Exception as e:
        logger.error(f"Erro ao buscar dados: {e}")
        await update.message.reply_text(text="⚠️ Erro ao processar a solicitação.")
        return
    
    if data.get("status") == "sucesso":
        await context.bot.send_message(chat_id=chat_id, text=f"{data.get('mensagem')}")
        logger.info(f"Atividade: {data.get('mensagem')}")
    else:
        ErroWH104 = (
            "WH104."
            "\n\n| VERIFICAR SE A SIGLA DO POP FOI INFORMADO CORRETAMENTE!"
            "\n\nCaso persistir, informar o erro à equipe interna com urgência!"
            "\n\nCONTATOS:"
            "\n     - @J_Ayrton"
            "\n     - @AlexandreBarros_Desktop"
        )
        error_message = data.get("mensagem", ErroWH104)

        logger.error(f"ERRO WH104: COMANDO /Atividades - POP:{pop} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")

        await context.bot.send_message(chat_id=chat_id, text=f"⚠️ Erro: {error_message}")

    return webhook_link


async def checar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user


    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return

    if len(context.args) < 2:
        await update.message.reply_text(text=ErroF101)
        return

    cto, fsan = context.args[:2]

    cto = str(cto.upper())
    pop = cto.split('-')[0]

    VerificarIfen_CTO = cto.count('-')
    if VerificarIfen_CTO > 1 or VerificarIfen_CTO < 1:
        await update.message.reply_text(text=ErroC101)
        return

    if '/' in fsan or '-' in fsan:
        await update.message.reply_text(text=ErroF102)
        return

    webhook_link = buscar_webhook_por_pop(pop)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return

    payload = {"comando": "Checar", "cto": cto, "fsan": fsan}

    logger.info(f"/Checar recebido - CTO: {cto}, FSAN: {fsan} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    data = await fetch_data(webhook_link, payload)

    if data.get("status") == "sucesso":
        await update.message.reply_text(text=f"{data.get('confirmacao')}")
    else:
        await update.message.reply_text(text=f"⚠️ Erro: {data.get('mensagem')}")
    return webhook_link

async def localizar_cto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return

    cto = context.args[0].upper()

    VerificarIfen_CTO = cto.count('-')
    if VerificarIfen_CTO > 1 or VerificarIfen_CTO < 1:
        await update.message.reply_text(text=ErroC101)
        return

    pop = cto.split('-')[0]
    webhook_link = buscar_webhook_por_pop(pop)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return

    payload = {"comando": "Localizar", "cto": cto}

    logger.info(f"/Localizar recebido - POP: {pop}, CTO: {cto} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    
    data = await fetch_data(webhook_link, payload)

    if data.get("status") == "sucesso":
        await update.message.reply_text(text=f"{data.get('mensagem')}")
    else:
        await update.message.reply_text(text=f"⚠️ CTO NÃO ENCONTRADO!")


async def ExibirCidade(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_title = update.effective_chat.title
    user = update.effective_user

    cidade = ListaCidades()

    await update.message.reply_text(text=f"🌆 Cidades disponíveis:\n\n{cidade}")

    logger.info(f"/ExibirCidade recebido - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    

async def input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_title = update.effective_chat.title or "Chat Privado"
    user = update.effective_user


    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return

    if len(context.args) < 2:
        await update.message.reply_text(text=ErroS101)
        return

    cto, splitter = context.args[:2]

    VerificarIfen_CTO = cto.count('-')
    if VerificarIfen_CTO > 1 or VerificarIfen_CTO < 1:
        await update.message.reply_text(text=ErroC101)
        return

    VerificarBarra_SPL = splitter.count('/')
    if VerificarBarra_SPL > 1 or VerificarBarra_SPL < 1:
        await update.message.reply_text(text=ErroN102)
        return

    VerificarIfen_CTO = cto.count('-')
    if VerificarIfen_CTO > 1 or VerificarIfen_CTO < 1:
        await update.message.reply_text(text=ErroC101)
        return

    cto = cto.upper()
    pop = cto.split('-')[0]
    webhook_link = buscar_webhook_por_pop(pop)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return

    splitters = {"16", "8", "4"}
    splitter_final = splitter.split("/")[-1]
    if splitter_final not in splitters:
        await update.message.reply_text(
            text=f"❌ SPLITTER inválido! Use apenas 1/16, 1/8, 1/4."
        )
        return

    payload = {"comando": "Input", "cto": cto, "splitter": splitter_final}

    logger.info(f"/Input recebido - POP: {pop}, CTO: {cto} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    
    try:
        data = await fetch_data(webhook_link, payload)
    except Exception as e:
        logger.error(f"/Input recebido - POP: {pop}, CTO: {cto}, {e} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
        await update.message.reply_text(
            text="⚠️ Erro interno ao processar sua solicitação. Tente novamente mais tarde."
        )
        return

    if data.get("status") == "sucesso":
        await update.message.reply_text(text=f"{data.get('confirmacao')}")
    else:
        await update.message.reply_text(text=f" ⚠️ Erro: {data.get('mensagem')}")


async def AjudaAdm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title or "Chat Privado"
    
    comandos = (
        "| AjudaAdm:"

        "\n\n - EXIBIR O ID DO GRUPO:"
        "\n    /id"

        "\n\n- EXIBIR CIDADES SALVAS:"
        "\n    /ExibirCidade"

        "\n\n- EXCLUIR TEMPLATE EXISTENTE:"
        "\n    /ExcluirTemplate <cidade>"

        "\n\n- ADICIONAR NOVO TEMPLATE:"
        "\n    /AddTemplate <cidade> <POP> <WebHook>" 

        "\n\n- Grupo de logger:"
        "\n    https://t.me/+Ij5OdRrCgAVkNTIx"

        "\n\n- One Driver Backup:"
        "\n    https://1drv.ms/f/s!AltzaXN7TtjqkqR0OQJ0jYa9VSyhWg?e=bb1LEy"

        "\n\n- Compartilhar Webhook.json:"
        "\n    /CWH"

        "\n\n| *Quando o nome da cidade conter 'espaço' lembre-se \n de subistituir por hífen '-'."
    )

    await context.bot.send_message(chat_id=chat_id, text=comandos)

    logger.info(f"/ajudaadm - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    

async def CWH(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title or "Chat Privado"

    logger.info(f"/CWH - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    
    chat_id = update.effective_chat.id
    arquivo = open('WebHook.json', 'rb')
    await context.bot.send_document(chat_id=chat_id, document=arquivo)
    

async def AdcionarTemplate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    def ADC_WEBHOOK(CIDADE_ID, POP, WEBHOOK_LINK):
        novo_dado = {
            'CIDADE': CIDADE_ID,
            'POP': POP,
            'WEBHOOK_LINK': WEBHOOK_LINK
        }

        if os.path.exists(FileName):
            try:
                with open(FileName, 'r', encoding='utf-8') as arquivo_json:
                    dados_existentes = json.load(arquivo_json)
            except json.JSONDecodeError:
                dados_existentes = []
        else:
            dados_existentes = []

        dados_existentes.append(novo_dado)
        with open(FileName, 'w', encoding='utf-8') as arquivo_json:
            json.dump(dados_existentes, arquivo_json, ensure_ascii=False, indent=4)

    if not update.message:
        print("Erro: Não foi possivel capturar o update do Webhook.")
        return

    if len(context.args) < 3:
        await update.message.reply_text(
            text=(
                "❌ Formato inválido!\n\n"
                "Use: /AdcionarTemplate <CIDADE> <POP> <WEBHOOK>\n\n"
                "Exemplo:\n/AdcionarTemplate Rio_Claro RCA https://script.google.com/macros..."
            )
        )
        return

    CIDADE_ID, POP, WEBHOOK_LINK = context.args[:3]
    CIDADE_ID = CIDADE_ID.upper()
    POP = POP.upper()

    ADC_WEBHOOK(CIDADE_ID, POP, WEBHOOK_LINK)

    cidade = ListaCidades()

    await update.message.reply_text(text=f"✅ Novo template adicionado:\n\n- CIDADE: {CIDADE_ID}\n- POP: {POP}\n- WEBHOOK: {WEBHOOK_LINK}")
    await update.message.reply_text(text=f"Lista de cidades existentes:\n\n{cidade}")

    logger.info(f"/AdcionarTemplate - CIDADE:{CIDADE_ID}, POP:{POP} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")


async def ExcluirTemplate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user
    

    MensagemErro = "❌ Formato inválido!\n\n Use: /ExcluirTemplate <POP>\n    - Ex: /ExcluirTemplate TIE\n    * Importante que o nome da cidade seja exatamente igual ao registrado."

    if len(context.args) < 1:
        await update.message.reply_text(text=MensagemErro)
        return


    ExcluirCidade = '-'.join(context.args).upper()

    logger.info(f"/ExcluirTemplate - cidade:{ExcluirCidade} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")

    with open(FileName, 'r') as file:
        dados = json.load(file)

    dados_atualizados = [POP for POP in dados if POP['POP'] != ExcluirCidade]

    if len(dados) == len(dados_atualizados):
        await update.message.reply_text(text=f"⚠️ A cidade {ExcluirCidade} não foi encontrada.")
    else:
        with open(FileName, 'w') as file:
            json.dump(dados_atualizados, file, indent=4)

        await update.message.reply_text(text=f"✅ O POP '{ExcluirCidade}' foi excluído com sucesso!")
        cidade = ListaCidades()
        await update.message.reply_text(text=f"Lista de cidades existentes:\n\n{cidade}")


async def id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    logger.info(f"/id - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")
    
    await update.message.reply_text(f"A ID deste grupo é: {chat_id}, "
                                    f"Nome do grupo: {chat_title}")

async def Info(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user


    Inf = (
        "| Nome do BOT: Alexo"
        "\n\n - Alexo tem o intuito de ser um auxílio para os usuários técnicos, back-offices e internos, com a capacidade de gerar de editar plalhas inopputando informações direto do chat, assim reduzindo as margens se erros na inputação de diversos procedimentos por todas as equipes."
        f"\n\nVersão: {__version__}"
        f"\n\nCriador: {__author__}"
        f"\nCréditos: {__credits__}"
    )

    logger.info(f"/Info - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")

    await update.message.reply_text(Inf)

async def listarIDs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user


    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return
    
    if len(context.args) < 2:
        await update.message.reply_text(text=ErroS101)
        return
    
    pop, OLT_SLOT_PON = context.args[:2]
    pop = pop.upper()
    pop = pop.split('-')[0]

    VerificarBarra = OLT_SLOT_PON.count('/')
    if VerificarBarra > 2 or VerificarBarra < 2:
        await update.message.reply_text(text=ErroN102)
        return

    partes = OLT_SLOT_PON.split("/")

    olt = partes[0] 
    slot = partes[1]
    pon = partes[2] 


    payload = {"comando": "ListarIds", "olt": olt, "slot": slot, "pon": pon}

    logger.info(f"/ListarIDs - OLT:{olt}, SLOT:{slot}, PON:{pon} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")


    webhook_link = buscar_webhook_por_pop(pop)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return
    
    data = await fetch_data(webhook_link, payload)

    if data.get("status") == "sucesso":
        ctos = data.get('mensagem')
        ctos_com_contador = [f"{i+1}. {cto}" for i, cto in enumerate(ctos)]
        ctos_com_contador_str = '\n'.join(ctos_com_contador)
        await update.message.reply_text(text=f"IDs disponiveis:\n\n{ctos_com_contador_str}\n\n| Sempre use o Ids da CTO de número [1]")
    else:
        await update.message.reply_text(text=f"⚠️ Erro: {data.get('mensagem')}")

async def insert(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_title = update.effective_chat.title
    user = update.effective_user


    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return
    
    if len(context.args) < 2:
        await update.message.reply_text(text=ErroN101)
        return
    
    CTO, OLT_SLOT_PON = context.args[:2]
    
    CTO = str(CTO.upper())
    POP = CTO.split('-')[0]

    VerificarIfen_CTO = CTO.count('-')
    if VerificarIfen_CTO > 1 or VerificarIfen_CTO < 1:
        await update.message.reply_text(text=ErroC101)
        return

    VerificarBarra = OLT_SLOT_PON.count('/')
    if VerificarBarra > 2 or VerificarBarra < 2:
        await update.message.reply_text(text=ErroN102)
        return

    if "/" in OLT_SLOT_PON:
        partes = OLT_SLOT_PON.split("/")
        
        olt = partes[0]
        slot = partes[1]
        pon = partes[2]

    else:
        olt = OLT_SLOT_PON.upper()
        slot = ""
        pon = ""

    payload = {"comando": "Insert", "cto": CTO, "olt": olt, "slot": slot, "pon": pon}

    logger.info(f"/Insert - CTO:{CTO}, PON:{OLT_SLOT_PON} - Usuário:{user.first_name} {user.last_name}, Grupo:{chat_title}")


    webhook_link = buscar_webhook_por_pop(POP)

    if webhook_link is None:
        await update.message.reply_text(ErroP102)
        return
    
    data = await fetch_data(webhook_link, payload)

    if data.get("status") == "sucesso":
        ctos = data.get('mensagem')

        await update.message.reply_text(text=f"{ctos}")
    else:
        await update.message.reply_text(text=f"⚠️ Erro: {data.get('mensagem')}")


async def novaCTO(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    chat_title = update.effective_chat.title
    user = update.effective_user

    if len(context.args) < 1:
        await update.message.reply_text(text=ErroP101)
        return
    
    if len(context.args) < 2:
        await update.message.reply_text(text=ErroN101)
        return
    
    if len(context.args) < 3:
        await update.message.reply_text(text=ErroS101)
        return

    pop = context.args[0]
    pop = pop.split('-')[0]

    olt_slot_pon = context.args[1]
    VerificarBarra = olt_slot_pon.count('/')
    if VerificarBarra > 2 or VerificarBarra < 2:
        await update.message.reply_text(text=ErroN102)
        return


    splitter = context.args[2]

    VerificarBarra_SPL = splitter.count('/')
    if VerificarBarra_SPL > 1 or VerificarBarra_SPL < 1:
        await update.message.reply_text(text=ErroN102)
        return

    splitters = {"16", "8", "4"}
    splitter_final = splitter.split("/")[-1]

    if splitter_final not in splitters:
        await update.message.reply_text(
            text="❌ SPLITTER inválido! Use apenas 1/16, 1/8, 1/4."
        )
        return 
    await update.message.reply_text(
        text="📍 Por favor, envie a localização da CTO que deseja adicionar."
    )

    context.user_data['waiting_for_location'] = True
    context.user_data['pop'] = pop
    context.user_data['olt_slot_pon'] = olt_slot_pon
    context.user_data['splitter'] = splitter
    context.user_data['splitter_final'] = splitter_final

    logger.info(f"/NovaCTO - POP:{pop}, PON:{olt_slot_pon}, SPL:{splitter} - Usuário:{user}, Grupo:{chat_title}")


async def handle_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_title = update.effective_chat.title or "Chat Privado"
    chat_id = update.effective_chat.id
    user = update.effective_user.username or "Usuário Privado"
    user_id = update.effective_user.id


    def EquipeWH(chat_title, chat_id, user, user_id):
        FileNameEquipe = "EquipeWH.json"

        Novo_DadosEquipe = {
            'NomeEquipe': chat_title,
            'ID_Equipe': chat_id,
            'NomeTec': f"{user}",
            'ID_Tec': user_id,
        }

        if os.path.exists(FileNameEquipe):
            try:
                with open(FileNameEquipe, 'r', encoding='utf-8') as arquivo_json:
                    dados_existentes = json.load(arquivo_json)
            except json.JSONDecodeError:
                dados_existentes = []
        else:
            dados_existentes = []

        dados_existentes.append(Novo_DadosEquipe)
        with open(FileNameEquipe, 'w', encoding='utf-8') as arquivo_json:
            json.dump(dados_existentes, arquivo_json, ensure_ascii=False, indent=4)

    EquipeWH(chat_title, chat_id,user,user_id)

    

    if update.message and update.message.location:
        location = update.message.location
        latitude = location.latitude
        longitude = location.longitude
        accuracy = location.horizontal_accuracy if location.horizontal_accuracy else "Desconhecida"

        if context.user_data.get('waiting_for_location'):
            pop = context.user_data.get('pop')
            webhook_link = buscar_webhook_por_pop(pop)

            if webhook_link is None:
                await update.message.reply_text(ErroP102)
                return

            olt_slot_pon = context.user_data.get('olt_slot_pon')
            splitter = context.user_data.get('splitter')

            splitter_final = splitter.split("/")[-1]
            if splitter_final not in {"16", "8", "4"}:
                await update.message.reply_text(
                    text="❌ Formato do spliter inválido! [1]"
                )
                return

            filtros = olt_slot_pon.split("/")
            if len(filtros) != 3:
                await update.message.reply_text("❌ Formato da pon inválido! [2]")
                return

            olt, slot, pon = filtros

            payload = {
                "comando": "NovaCto",
                "olt": olt,
                "slot": slot,
                "pon": pon,
                "latitude": latitude,
                "longitude": longitude,
                "splitter": splitter_final,
                "id": chat_id
            }

            await update.message.reply_text(
                text=f"📍 Localização recebida: {latitude}, {longitude}\n"
                     f"Precisão: {accuracy} metros\n"
                     f"POP: {pop}\n"
                     f"OLT/SLOT/PON: {olt_slot_pon}\n"
                     f"Splitter: {splitter}\n"
                     "\nEnviando as informações para o template, aguarde..."
            )

            logger.info(f"payload: {payload} ////// Localização /NovaCTO recebida - Usuário:{user}, Grupo:{chat_title}")

            data = await fetch_data(webhook_link, payload)

            if data.get("status") == "sucesso":
                await update.message.reply_text(text=f"{data.get('mensagem')}")
            else:
                await update.message.reply_text(text=f"⚠️ Erro: {data.get('mensagem')}")

            context.user_data['waiting_for_location'] = False

        else:
            chat_title = update.effective_chat.title
            
            logger.info(f"Localização recebida - Usuário:{user}, Grupo:{chat_title}")

            Msg_Localizacao = f"📍 - Informações da localização\n\n| Coordenadas: {latitude}, {longitude}\n| Precisão: {accuracy} metros\n\n| Link-Maps: https://www.google.pt/maps?q={latitude},{longitude}"

            await update.message.reply_text(text=Msg_Localizacao)
    else:
        if update.message:
            await update.message.reply_text("❌ Não foi possível obter a localização. Por favor, envie uma localização válida.")
        else:
            pass


async def convert(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user = update.effective_user
    context.user_data['waiting_for_file'] = True
    await update.message.reply_text("Por favor, envie o arquivo KML/KMZ que você deseja converter.")


async def handle_arquivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    user = update.effective_user

    if context.user_data.get('waiting_for_file'):

        if update.message and update.message.document:
            document = update.message.document
            file_id = document.file_id
            file_name = document.file_name
            file_extension = file_name.split('.')[-1]

            file = await context.bot.get_file(file_id)
            file_path = f"{file_name}"

            await file.download_to_drive(file_path)

            logger.info(f"Arquivo Recebido - Arquivo:{file_name} - Usuário:{user.first_name} {user.last_name}")

            if file_extension == 'kml':
                Arq = "KML"
                xlsx_file = file_path.replace(file_extension, 'xlsx')
                kml_to_xlsx(file_path, xlsx_file)

                Mensagem_User = f"📄 - Conversor KML para XLSX: \n\nArquivo: {file_name}\nFomarto do arquivo: {Arq}\nConvertido para: XLSX\nNovo Arquivo: {xlsx_file}"


                await update.message.reply_text(text=Mensagem_User)
                await context.bot.send_document(chat_id=chat_id, document=open(xlsx_file, 'rb'))

            elif file_extension == 'kmz':
                extract_to = ""
                kml_file = extract_kml_from_kmz(file_path, extract_to)
                if kml_file:
                    xlsx_file = kml_file.replace('kml', 'xlsx')
                    Arq = "KMZ"
                    kml_to_xlsx(kml_file, xlsx_file)

                    Mensagem_User = f"📄 - Conversor KMZ para XLSX: \n\nArquivo: {file_name}\nFomarto do arquivo: {Arq}\nConvertido para: XLSX\nNovo Arquivo: {xlsx_file}"
                    await update.message.reply_text(text=Mensagem_User)
                    await context.bot.send_document(chat_id=chat_id, document=open(xlsx_file, 'rb'))
                else:
                    await update.message.reply_text("❌ Não foi possível extrair o arquivo KML do arquivo KMZ.")

            else:
                await update.message.reply_text(text=f"📄 Arquivo recebido: {file_name}\nTamanho: {document.file_size} bytes\nFile ID: {file_id}")
            
            context.user_data['waiting_for_file'] = False

        else:
            await update.message.reply_text("❌ Não foi possível identificar o arquivo. Por favor, envie um arquivo válido.")
    
    else:
        pass
    
        #await update.message.reply_text("❌ Por favor, use o comando /Convert antes de enviar o arquivo.")


if __name__ == "__main__":
    logger.info(f"Iniciando o BOT-ALEXO Versão:{__version__}")

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("Id", id))
    app.add_handler(CommandHandler("CWH", CWH))
    app.add_handler(CommandHandler("Info", Info))
    app.add_handler(CommandHandler("Input", input))
    app.add_handler(CommandHandler("Convert", convert))
    app.add_handler(CommandHandler("Insert", insert))
    app.add_handler(CommandHandler("Ajuda", ajuda))
    app.add_handler(CommandHandler("Checar", checar))
    app.add_handler(CommandHandler("NovaCTO", novaCTO))
    app.add_handler(CommandHandler("AjudaAdm", AjudaAdm))
    app.add_handler(CommandHandler("ListarIDs", listarIDs))
    app.add_handler(CommandHandler("Atividades", atividades))
    app.add_handler(CommandHandler("Localizar", localizar_cto))
    app.add_handler(CommandHandler("ExibirCidade", ExibirCidade))
    app.add_handler(CommandHandler("AddTemplate", AdcionarTemplate))
    app.add_handler(CommandHandler("ExcluirTemplate", ExcluirTemplate))

    app.add_handler(MessageHandler(filters.Document.ALL, handle_arquivo))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
    
    logger.info("Bot está rodando...")
    app.run_polling()

