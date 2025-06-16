import os
import logging
from .version import __version__



def selecionar_token(bot_id):
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger()
    logger.info("Iniciando a seleção do token do bot...")
    if bot_id == 1:
        bot_token = os.getenv("TOKEN_BOT_ALEXO")
        db_database = os.getenv("DB_DATABASE")

        print(db_database)
        if not bot_token:
            logger.error("Token do Bot 1 não encontrado.")
            raise ValueError("Token do Bot 1 está ausente.")
        logger.info(f"BOT ALEXO - iniciando na Versão:{__version__}")
        logger.info(f"Banco de Dados: {db_database}")
        return bot_token, db_database
    elif bot_id == 2:
        bot_token = os.getenv("TOKEN_DKT_TESTE")
        db_database = os.getenv("DB_DATABASE_TESTE")
        if not bot_token:
            logger.error("Token do Bot 2 não encontrado.")
            raise ValueError("Token do Bot 2 está ausente.")
        logger.info(f"BOT TESTE - iniciando na Versão:{__version__}")
        return bot_token, db_database
    else:
        logger.error("Nenhum bot válido foi selecionado, verifique o atributo.")
        raise ValueError("ID do bot inválido! Escolha 1 ou 2.")

