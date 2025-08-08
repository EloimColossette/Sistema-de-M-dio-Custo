import os
import psycopg2
import logging

logger = logging.getLogger(__name__)

def conectar(ip=None, porta=None, usuario=None, senha=None, dbname=None):
    from login import TelaLogin
    config = TelaLogin.carregar_configuracoes_static()

    host     = ip     or config.get("DB_HOST",   "192.168.1.117")
    port     = porta  or config.get("DB_PORT",   "5432")
    user     = usuario or config.get("DB_USER",   os.environ.get("DB_USER", "Kametal"))
    password = senha   or config.get("DB_PASSWORD", os.environ.get("DB_PASSWORD","908513"))
    dbname   = dbname  or config.get("DB_NAME",   os.environ.get("DB_NAME", "Teste"))

    logger.info(f"Tentando conectar: {dbname}@{host}:{port} como {user}")
    dsn = (f"dbname={dbname} user={user} password={password} "
           f"host={host} port={port} client_encoding=UTF8")
    try:
        conn = psycopg2.connect(dsn)
        logger.info("Conexão estabelecida com sucesso!")
        return conn
    except Exception:
        logger.exception("Erro ao conectar")
        return None
