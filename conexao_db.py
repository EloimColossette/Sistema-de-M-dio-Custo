import os
import time
import random
import logging
import threading
import functools
from contextlib import contextmanager
from collections import deque
from typing import Optional
import psycopg2
import psycopg2.extras
from psycopg2 import pool, OperationalError, DatabaseError
import random
from typing import Optional

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


class CircuitoAbertoError(RuntimeError):
    """Lançado quando o circuit breaker está aberto."""
    pass


class PoolPostgres:
    """
    Pool Postgres com:
    - ThreadedConnectionPool do psycopg2
    - Limite de operações concorrentes (semaphore) -> backpressure
    - Circuit breaker simples
    - Retries configuráveis para criação do pool
    - Context managers para conexão e cursor (commit/rollback automático)
    """

    def __init__(
        self,
        ip: Optional[str] = None,
        porta: Optional[str] = None,
        usuario: Optional[str] = None,
        senha: Optional[str] = None,
        nome_bd: Optional[str] = None,
        minconn: int = 1,
        maxconn: int = 10,
        connect_timeout: int = 5,
        tentativas: int = 3,
        backoff_base: float = 0.5,
        application_name: str = "app",
        # parâmetros para alta concorrência / estabilidade
        max_operacoes_simultaneas: Optional[int] = None,  # limite de operações DB simultâneas no app
        tempo_espera_semaforo: float = 5.0,               # tempo (s) para esperar permissão do semáforo
        limite_falhas_para_circuito: int = 10,            # número de falhas para abrir circuito
        janela_falhas_segundos: int = 60,                 # janela para contar falhas
        cooldown_circuito_segundos: int = 30,             # tempo que o circuito fica aberto
        timeout_statement_ms_padrao: Optional[int] = 30_000,  # timeout por query (ms)
        timeout_idle_in_tx_ms: Optional[int] = 60_000,        # evita idle-in-transaction
    ):
        # tenta carregar configurações como no seu projeto (TelaLogin)
        try:
            from login import TelaLogin
            config = TelaLogin.carregar_configuracoes_static()
        except Exception:
            config = {}

        host = ip or config.get("DB_HOST", "192.168.1.117")
        port = porta or config.get("DB_PORT", "5432")
        user = usuario or config.get("DB_USER", os.environ.get("DB_USER", "Kametal"))
        password = senha or config.get("DB_PASSWORD", os.environ.get("DB_PASSWORD", "908513"))
        dbname = nome_bd or config.get("DB_NAME", os.environ.get("DB_NAME", "Teste"))

        self._con_params = dict(
            host=host,
            port=int(port),
            user=user,
            password=password,
            dbname=dbname,
            connect_timeout=int(connect_timeout),
            application_name=application_name,
            client_encoding="UTF8",
            # keepalives para detectar conexões mortas
            keepalives=1,
            keepalives_idle=30,
            keepalives_interval=10,
            keepalives_count=5,
        )

        self._minconn = minconn
        self._maxconn = maxconn
        self._tentativas = tentativas
        self._backoff_base = backoff_base
        self._pool: Optional[pool.ThreadedConnectionPool] = None

        # cria pool com retries
        self._criar_pool_com_retentativas()

        # backpressure / semáforo para limitar operações simultâneas
        self._max_operacoes = max_operacoes_simultaneas if max_operacoes_simultaneas is not None else maxconn
        self._semaforo = threading.BoundedSemaphore(self._max_operacoes)
        self._tempo_espera_semaforo = tempo_espera_semaforo

        # circuit breaker
        self._limite_falhas = limite_falhas_para_circuito
        self._janela_falhas = janela_falhas_segundos
        self._cooldown_circuito = cooldown_circuito_segundos
        self._registro_falhas = deque()
        self._circuito_aberto_ate = 0.0

        # timeouts
        self._timeout_statement_ms_padrao = timeout_statement_ms_padrao
        self._timeout_idle_in_tx_ms = timeout_idle_in_tx_ms

    def _criar_pool_com_retentativas(self):
        """Tenta criar o ThreadedConnectionPool com exponential backoff."""
        for tentativa in range(1, self._tentativas + 1):
            try:
                logger.info(
                    "Criando pool de conexões: %s -> %s (min=%d max=%d)",
                    self._con_params.get("user"),
                    self._con_params.get("host"),
                    self._minconn,
                    self._maxconn,
                )
                self._pool = pool.ThreadedConnectionPool(
                    self._minconn, self._maxconn, **self._con_params
                )
                logger.info("Pool criado com sucesso")
                return
            except Exception:
                logger.exception("Falha ao criar pool (tentativa %d/%d)", tentativa, self._tentativas)
                if tentativa == self._tentativas:
                    raise
                tempo_espera = self._backoff_base * (2 ** (tentativa - 1)) + random.random() * 0.1
                time.sleep(tempo_espera)

    # ----------------------------
    # Circuit breaker (simples)
    # ----------------------------
    def _registrar_falha(self):
        agora = time.time()
        self._registro_falhas.append(agora)
        corte = agora - self._janela_falhas
        while self._registro_falhas and self._registro_falhas[0] < corte:
            self._registro_falhas.popleft()
        if len(self._registro_falhas) >= self._limite_falhas:
            self._abrir_circuito()

    def _abrir_circuito(self):
        self._circuito_aberto_ate = time.time() + self._cooldown_circuito
        logger.warning("Circuito ABERTO até %s", time.ctime(self._circuito_aberto_ate))

    def _circuito_permite(self) -> bool:
        if time.time() < self._circuito_aberto_ate:
            return False
        # reseta contador se passou o cooldown
        self._registro_falhas.clear()
        self._circuito_aberto_ate = 0.0
        return True

    # ----------------------------
    # Obter / devolver conexões
    # ----------------------------
    def obter_conexao(self):
        """
        Adquire permissão (semaforo) e pega uma conexão do pool.
        Lança CircuitoAbertoError se o circuito estiver aberto.
        """
        if not self._circuito_permite():
            raise CircuitoAbertoError("Circuito aberto — muitas falhas recentes")

        adquirido = self._semaforo.acquire(timeout=self._tempo_espera_semaforo)
        if not adquirido:
            raise TimeoutError("Não foi possível adquirir permissão para usar o banco (carga alta)")

        try:
            if not self._pool:
                self._criar_pool_com_retentativas()
            conexao = self._pool.getconn()
            if getattr(conexao, "closed", 1):
                logger.warning("Conexão retirada do pool estava fechada — recriando diretamente")
                conexao = psycopg2.connect(**self._con_params)
            return conexao
        except Exception as exc:
            # liberar semáforo e registrar falha
            try:
                self._semaforo.release()
            except Exception:
                pass
            self._registrar_falha()
            raise

    def devolver_conexao(self, conexao, fechar: bool = False):
        """
        Devolve a conexão ao pool e libera o semáforo.
        Se fechar=True, fecha a conexão em vez de devolver.
        """
        if not conexao:
            try:
                self._semaforo.release()
            except Exception:
                pass
            return

        try:
            if fechar:
                try:
                    conexao.close()
                finally:
                    return
            self._pool.putconn(conexao)
        except Exception:
            logger.exception("Erro ao devolver conexão ao pool — fechando")
            try:
                conexao.close()
            except Exception:
                pass
        finally:
            try:
                self._semaforo.release()
            except Exception:
                pass

    def fechar_tudo(self):
        """Fecha todas as conexões do pool."""
        if self._pool:
            try:
                self._pool.closeall()
            except Exception:
                logger.exception("Erro ao fechar todas as conexões do pool")
            finally:
                self._pool = None

    # ----------------------------
    # Context managers (conexão / cursor)
    # ----------------------------
    @contextmanager
    def conexao(self, timeout_statement_ms: Optional[int] = None):
        """
        Context manager para obter uma conexão e aplicar timeouts de sessão local.
        Exemplo:
            with pg.conexao(timeout_statement_ms=5000) as conn:
                cur = conn.cursor(); cur.execute(...)
        """
        conn = None
        try:
            conn = self.obter_conexao()
            cur = conn.cursor()
            st_ms = timeout_statement_ms if timeout_statement_ms is not None else self._timeout_statement_ms_padrao
            if st_ms is not None:
                cur.execute("SET LOCAL statement_timeout = %s", (st_ms,))
            if self._timeout_idle_in_tx_ms is not None:
                cur.execute("SET LOCAL idle_in_transaction_session_timeout = %s", (self._timeout_idle_in_tx_ms,))
            cur.close()
            yield conn
        except Exception as exc:
            # registra falha para o circuit breaker se for erro operacional
            if isinstance(exc, (OperationalError, DatabaseError)):
                self._registrar_falha()
            raise
        finally:
            try:
                self.devolver_conexao(conn)
            except Exception:
                logger.exception("Erro ao devolver conexão no context manager")

    @contextmanager
    def cursor(self, dict_cursor: bool = False, commit: bool = False, timeout_statement_ms: Optional[int] = None):
        """
        Context manager para usar cursor com commit/rollback automático.
        Uso:
            with pg.cursor(dict_cursor=True, commit=True) as cur:
                cur.execute(...)
        """
        with self.conexao(timeout_statement_ms=timeout_statement_ms) as conn:
            factory = psycopg2.extras.RealDictCursor if dict_cursor else None
            cur = conn.cursor(cursor_factory=factory)
            try:
                yield cur
                if commit:
                    conn.commit()
            except Exception:
                try:
                    conn.rollback()
                except Exception:
                    logger.exception("Rollback falhou")
                raise
            finally:
                try:
                    cur.close()
                except Exception:
                    pass

    # ----------------------------
    # Decorator de retry para operações
    # ----------------------------
    def retry_em_erro_operacional(self, tentativas: int = 3, backoff_base: float = 0.1, excecoes_permitidas=(OperationalError,)):
        """
        Decorator para reexecutar função que faz queries em falhas transitórias.
        Uso:
            @pg.retry_em_erro_operacional(tentativas=4)
            def minha_func(...): ...
        """
        def deco(fn):
            @functools.wraps(fn)
            def wrapper(*args, **kwargs):
                tentativa = 0
                while True:
                    try:
                        return fn(*args, **kwargs)
                    except excecoes_permitidas as exc:
                        tentativa += 1
                        logger.warning("Operação falhou (%s). tentativa %d/%d", exc, tentativa, tentativas)
                        if tentativa >= tentativas:
                            self._registrar_falha()
                            raise
                        tempo_espera = backoff_base * (2 ** (tentativa - 1)) + random.random() * 0.05
                        time.sleep(tempo_espera)
            return wrapper
        return deco

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

# PoolPostgres (usando a classe definida acima no mesmo arquivo)

_pg = None

def conectar(ip: Optional[str] = None,
             porta: Optional[str] = None,
             usuario: Optional[str] = None,
             senha: Optional[str] = None,
             dbname: Optional[str] = None):
    """
    Versão autocontida de conectar() — cria o pool localmente se necessário e
    fornece fallback. Protege criação com lock armazenado como atributo da função,
    faz health-check e aplica SET LOCAL para timeouts. Usa a classe ConexaoPooled
    definida no módulo quando disponível; caso contrário, cria um wrapper mínimo.
    """
    global _pg

    # imports locais necessários para evitar dependências no topo do módulo
    import os
    import random
    import time

    # garante logger mínimo
    try:
        logger
    except NameError:
        import logging
        logger = logging.getLogger(__name__)
        logger.addHandler(logging.NullHandler())

    # cria lock como atributo da função (thread-safe, sem editar topo)
    if not hasattr(conectar, "_lock"):
        import threading
        conectar._lock = threading.Lock()

    # helper: criar wrapper local se ConexaoPooled não existir no módulo
    ConexaoWrapper = None
    if "ConexaoPooled" in globals() and callable(globals().get("ConexaoPooled")):
        ConexaoWrapper = globals().get("ConexaoPooled")
    else:
        # classe mínima compatível
        class _ConexaoPooledLocal:
            def __init__(self, real_conn, pool_obj):
                self._real = real_conn
                self._pool = pool_obj
                self._fechada = False

            def close(self):
                if self._fechada:
                    return
                try:
                    if self._pool is None:
                        try:
                            self._real.close()
                        except Exception:
                            pass
                    else:
                        # tenta devolver via API conhecida
                        try:
                            if hasattr(self._pool, "devolver_conexao"):
                                self._pool.devolver_conexao(self._real)
                            elif hasattr(self._pool, "putconn"):
                                self._pool.putconn(self._real)
                            else:
                                self._real.close()
                        except Exception:
                            try:
                                self._real.close()
                            except Exception:
                                pass
                finally:
                    self._fechada = True

            def fechar_de_verdade(self):
                try:
                    self._real.close()
                except Exception:
                    pass
                finally:
                    self._fechada = True

            def __getattr__(self, name):
                return getattr(self._real, name)

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                try:
                    if exc_type is not None:
                        try:
                            self._real.rollback()
                        except Exception:
                            pass
                finally:
                    try:
                        self.close()
                    except Exception:
                        pass

        ConexaoWrapper = _ConexaoPooledLocal

    # tenta criar / obter o pool global com proteção
    with conectar._lock:
        if _pg is None:
            # se houver PoolPostgres (definido no módulo db_pool.py), usar
            if 'PoolPostgres' in globals() and globals().get('PoolPostgres') is not None:
                try:
                    PCls = globals().get('PoolPostgres')
                    _pg = PCls(
                        minconn=1,
                        maxconn=30,
                        max_operacoes_simultaneas=20,
                        connect_timeout=3,
                        tentativas=3,
                        timeout_statement_ms_padrao=10_000,
                        limite_falhas_para_circuito=8,
                        janela_falhas_segundos=60,
                        cooldown_circuito_segundos=20,
                    )
                    logger.info("PoolPostgres criado localmente em conectar().")
                except Exception:
                    logger.exception("Falha ao criar PoolPostgres, tentaremos ThreadedConnectionPool fallback.")
                    _pg = None

            if _pg is None:
                # fallback para ThreadedConnectionPool do psycopg2
                try:
                    import psycopg2
                    from psycopg2 import pool as _pspool
                    # tenta ler credenciais do login.py ou env
                    try:
                        from login import TelaLogin
                        cfg = TelaLogin.carregar_configuracoes_static()
                    except Exception:
                        cfg = {}
                    host = ip or cfg.get("DB_HOST", os.environ.get("DB_HOST", "127.0.0.1"))
                    port = porta or cfg.get("DB_PORT", os.environ.get("DB_PORT", "5432"))
                    user = usuario or cfg.get("DB_USER", os.environ.get("DB_USER", os.environ.get("DB_USER", "Kametal")))
                    password = senha or cfg.get("DB_PASSWORD", os.environ.get("DB_PASSWORD", os.environ.get("DB_PASSWORD","908513")))
                    dbname_final = dbname or cfg.get("DB_NAME", os.environ.get("DB_NAME", os.environ.get("DB_NAME", "Teste")))
                    _pg = _pspool.ThreadedConnectionPool(1, 30,
                                                        host=host, port=port, user=user, password=password, dbname=dbname_final,
                                                        connect_timeout=3)
                    _pg._is_fallback = True
                    logger.info("ThreadedConnectionPool (fallback) criado em conectar().")
                except Exception:
                    logger.exception("Não foi possível criar pool fallback em conectar().")
                    _pg = None

    # agora tenta retirar conexão do pool (se houver)
    if _pg is not None:
        try:
            raw = None
            if hasattr(_pg, "obter_conexao"):
                raw = _pg.obter_conexao()
            elif hasattr(_pg, "getconn"):
                raw = _pg.getconn()
            else:
                raw = None

            if raw is not None:
                # health-check rápido
                try:
                    cur = raw.cursor()
                    try:
                        cur.execute("SELECT 1")
                        cur.fetchone()
                    finally:
                        cur.close()
                except Exception as exc:
                    logger.warning("Conexão retirada do pool inválida: %s", exc)
                    # tenta devolver/fechar e pegar outra (uma tentativa)
                    try:
                        if hasattr(_pg, "putconn"):
                            _pg.putconn(raw, close=True)
                        elif hasattr(_pg, "devolver_conexao"):
                            _pg.devolver_conexao(raw, fechar=True)
                        else:
                            try:
                                raw.close()
                            except Exception:
                                pass
                    except Exception:
                        pass
                    # tenta obter nova conexão (se disponível)
                    try:
                        if hasattr(_pg, "obter_conexao"):
                            raw = _pg.obter_conexao()
                        elif hasattr(_pg, "getconn"):
                            raw = _pg.getconn()
                        else:
                            raw = None
                    except Exception as exc2:
                        logger.warning("Falha ao tentar novo checkout do pool: %s", exc2)
                        raw = None

            if raw is not None:
                # aplica timeouts locais se o pool tiver atributos de timeout
                try:
                    st_ms = getattr(_pg, "_timeout_statement_ms_padrao", None) or getattr(_pg, "_default_statement_timeout_ms", None) or getattr(_pg, "_timeout_statement_ms", None)
                    idle_ms = getattr(_pg, "_timeout_idle_in_tx_ms", None) or getattr(_pg, "_timeout_idle_ms", None) or getattr(_pg, "_idle_in_transaction_session_timeout_ms", None)
                    if st_ms is not None or idle_ms is not None:
                        cur = raw.cursor()
                        try:
                            if st_ms is not None:
                                cur.execute("SET LOCAL statement_timeout = %s", (int(st_ms),))
                            if idle_ms is not None:
                                cur.execute("SET LOCAL idle_in_transaction_session_timeout = %s", (int(idle_ms),))
                        finally:
                            cur.close()
                except Exception:
                    logger.exception("Não foi possível aplicar SET LOCAL na conexão do pool (seguimos).")

                return ConexaoWrapper(raw, _pg)

        except Exception:
            logger.exception("Falha ao obter conexão do pool; tentando fallback direto.")

    # fallback direto com psycopg2.connect + retries para OperationalError
    try:
        import psycopg2
        from psycopg2 import OperationalError as PsyOpErr
    except Exception:
        psycopg2 = None
        PsyOpErr = Exception

    try:
        from login import TelaLogin
        cfg = TelaLogin.carregar_configuracoes_static()
    except Exception:
        cfg = {}

    host = ip or cfg.get("DB_HOST", os.environ.get("DB_HOST", "127.0.0.1"))
    port = porta or cfg.get("DB_PORT", os.environ.get("DB_PORT", "5432"))
    user = usuario or cfg.get("DB_USER", os.environ.get("DB_USER", os.environ.get("DB_USER", "Kametal")))
    password = senha or cfg.get("DB_PASSWORD", os.environ.get("DB_PASSWORD", os.environ.get("DB_PASSWORD","908513")))
    dbname_final = dbname or cfg.get("DB_NAME", os.environ.get("DB_NAME", os.environ.get("DB_NAME", "Teste")))

    if psycopg2 is None:
        logger.error("psycopg2 não disponível para fallback direto.")
        return None

    tentativas = 3
    backoff_base = 0.1
    tentativa = 0
    last_exc = None
    while tentativa < tentativas:
        try:
            dsn = (f"dbname={dbname_final} user={user} password={password} "
                   f"host={host} port={port} client_encoding=UTF8")
            conn = psycopg2.connect(dsn)
            # health-check
            try:
                cur = conn.cursor()
                try:
                    cur.execute("SELECT 1")
                    cur.fetchone()
                finally:
                    cur.close()
            except Exception:
                try:
                    conn.close()
                except Exception:
                    pass
                raise

            # aplicar timeouts locais se tivermos valores
            try:
                st_ms = getattr(_pg, "_timeout_statement_ms_padrao", None) if _pg is not None else None
                idle_ms = getattr(_pg, "_timeout_idle_in_tx_ms", None) if _pg is not None else None
                if st_ms is not None or idle_ms is not None:
                    cur = conn.cursor()
                    try:
                        if st_ms is not None:
                            cur.execute("SET LOCAL statement_timeout = %s", (int(st_ms),))
                        if idle_ms is not None:
                            cur.execute("SET LOCAL idle_in_transaction_session_timeout = %s", (int(idle_ms),))
                    finally:
                        cur.close()
            except Exception:
                logger.exception("Não foi possível aplicar SET LOCAL no fallback (não crítico).")

            return ConexaoWrapper(conn, None)

        except PsyOpErr as exc:
            last_exc = exc
            tentativa += 1
            if tentativa >= tentativas:
                logger.exception("Falha ao conectar (fallback) após %d tentativas: %s", tentativa, exc)
                break
            wait = backoff_base * (2 ** (tentativa - 1)) + random.random() * 0.05
            logger.warning("Falha operacional ao conectar (tentativa %d/%d). Esperando %.2fs. Erro: %s",
                           tentativa, tentativas, wait, exc)
            time.sleep(wait)
        except Exception as exc:
            logger.exception("Erro não-operacional ao tentar conectar: %s", exc)
            return None

    logger.error("Fallback direto não conseguiu obter conexão. Último erro: %s", last_exc)
    return None
