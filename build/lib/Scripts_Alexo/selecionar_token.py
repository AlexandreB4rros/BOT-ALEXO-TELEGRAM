import os
import logging
from .version import __version__



def selecionar_token(bot_id):
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger()

    if bot_id == 1:
        bot_token = os.getenv("TOKEN_BOT_ALEXO")
        if not bot_token:
            logger.error("Token do Bot 1 não encontrado.")
            raise ValueError("Token do Bot 1 está ausente.")
        logger.info(f"BOT ALEXO - iniciando na Versão:{__version__}")
        return bot_token
    elif bot_id == 2:
        bot_token = os.getenv("TOKEN_DKT_TESTE")
        if not bot_token:
            logger.error("Token do Bot 2 não encontrado.")
            raise ValueError("Token do Bot 2 está ausente.")
        logger.info(f"BOT TESTE - iniciando na Versão:{__version__}")
        return bot_token
    else:
        logger.error("Nenhum bot válido foi selecionado, verifique o atributo.")
        raise ValueError("ID do bot inválido! Escolha 1 ou 2.")
