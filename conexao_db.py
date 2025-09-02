import os
import time
import threading
import logging
import weakref
from psycopg2 import connect
from psycopg2 import pool as pg_pool

logger = logging.getLogger(__name__)

# ---------------------------
# Configuráveis (defaults para produção)
# ---------------------------
_POOL_MIN = int(os.environ.get("DB_POOL_MIN", "1"))
_POOL_MAX = int(os.environ.get("DB_POOL_MAX", "50"))  # ajuste conforme limite do Postgres
_HEARTBEAT_ENABLED = True
_HEARTBEAT_INTERVAL = int(os.environ.get("DB_HEARTBEAT_INTERVAL", "60"))  # segundos entre pings
_CONN_KEEPALIVE_IDLE = int(os.environ.get("DB_KEEPALIVE_IDLE", "60"))
_CONN_KEEPALIVE_INTERVAL = int(os.environ.get("DB_KEEPALIVE_INTERVAL", "10"))
_CONN_KEEPALIVE_COUNT = int(os.environ.get("DB_KEEPALIVE_COUNT", "5"))
_CONNECT_TIMEOUT = int(os.environ.get("DB_CONNECT_TIMEOUT", "10"))  # seconds

# Prefill: False por padrão em produção para evitar consumir todas as conexões do servidor
_PREFILL_USE_MAXCONN = False

# timeout para adquirir o semáforo (evita bloqueio indefinido)
_SEM_TIMEOUT = int(os.environ.get("DB_SEM_TIMEOUT", "30"))  # segundos

# watchdog: tempo máximo que uma conexão pode ficar "emprestada" antes de ser forçosamente devolvida
_MAX_CHECKOUT_SECONDS = int(os.environ.get("DB_MAX_CHECKOUT_SECONDS", "120"))

# ---------------------------
# Estado do módulo
# ---------------------------
_pool = None
_pool_lock = threading.Lock()
_conn_params = None

_active_proxies = weakref.WeakSet()

_heartbeat_thread = None
_heartbeat_stop = threading.Event()

_checkout_semaphore = None

_heartbeat_conn = None

_total_checkouts = 0
_total_returns = 0
_total_forced_gc_returns = 0
_metrics_lock = threading.Lock()

# extras
_STATEMENT_TIMEOUT_MS = int(os.environ.get("DB_STATEMENT_TIMEOUT_MS", str(5 * 60 * 1000)))  # 5 minutos padrão
_APPLICATION_NAME = os.environ.get("DB_APPLICATION_NAME", "app_kametal")

# ---------------------------
# Internal helpers
# ---------------------------
def _init_pool_if_needed(minconn=_POOL_MIN, maxconn=_POOL_MAX, **connect_kwargs):
    """Inicializa pool ThreadedConnectionPool de forma idempotente."""
    global _pool, _conn_params, _heartbeat_thread, _checkout_semaphore, _heartbeat_conn

    with _pool_lock:
        if _pool is not None:
            return

        # Salva params para reconexão se necessário
        _conn_params = connect_kwargs.copy()

        # Valores de keepalive / timeout (ajustáveis)
        _conn_params.setdefault("connect_timeout", _CONNECT_TIMEOUT)
        _conn_params.setdefault("keepalives", 1)
        _conn_params.setdefault("keepalives_idle", _CONN_KEEPALIVE_IDLE)
        _conn_params.setdefault("keepalives_interval", _CONN_KEEPALIVE_INTERVAL)
        _conn_params.setdefault("keepalives_count", _CONN_KEEPALIVE_COUNT)
        _conn_params.setdefault("application_name", _APPLICATION_NAME)

        try:
            _pool = pg_pool.ThreadedConnectionPool(minconn, maxconn, **_conn_params)
            logger.info("Pool de conexões inicializado (min=%d max=%d)", minconn, maxconn)
        except Exception:
            logger.exception("Falha ao inicializar pool de conexões")
            _pool = None
            raise

        # cria semaphore alinhado ao maxconn do pool (impede "thread storm")
        try:
            _checkout_semaphore = threading.BoundedSemaphore(maxconn)
        except Exception:
            _checkout_semaphore = None

        # tenta criar conexão dedicada para heartbeat (não usar pool)
        try:
            _heartbeat_conn = connect(**_conn_params)
            try:
                _heartbeat_conn.set_session(autocommit=True)
            except Exception:
                pass
        except Exception:
            _heartbeat_conn = None
            logger.warning("Não foi possível criar conexão dedicada para heartbeat")

        # prefill (opcional — cria minconn conexões imediatamente)
        if _PREFILL_USE_MAXCONN:
            try:
                for _ in range(minconn):
                    c = _pool.getconn()
                    try:
                        _safe_ping_conn(c)
                    except Exception:
                        pass
                    _pool.putconn(c)
                logger.info("Pool prefill (minconn): %d conexões criadas.", minconn)
            except Exception:
                logger.debug("Prefill do pool falhou (não crítico).", exc_info=False)

        # Inicia heartbeat global que mantém conexões vivas (idle + em uso)
        if _HEARTBEAT_ENABLED:
            _heartbeat_stop.clear()
            _heartbeat_thread = threading.Thread(target=_heartbeat_loop, name="db-heartbeat", daemon=True)
            _heartbeat_thread.start()


def _safe_ping_conn(conn, timeout=5):
    """Executa SELECT 1 em conn; protege contra erros."""
    try:
        cur = conn.cursor()
        cur.execute("SELECT 1")
        _ = cur.fetchone()
        cur.close()
        return True
    except Exception:
        logger.debug("Ping falhou numa conexão (ignorado).", exc_info=False)
        return False


def _heartbeat_loop():
    """Loop que periodicamente pinga heartbeat dedicado e conexões ativas.
       Também força devolução de proxies que excederam _MAX_CHECKOUT_SECONDS.
    """
    global _pool, _heartbeat_conn
    logger.info("Heartbeat do pool iniciado (interval=%ss)", _HEARTBEAT_INTERVAL)
    while not _heartbeat_stop.wait(_HEARTBEAT_INTERVAL):
        # 1) ping na conexão dedicada (se existir)
        try:
            if _heartbeat_conn is None:
                try:
                    _heartbeat_conn = connect(**_conn_params)
                    try:
                        _heartbeat_conn.set_session(autocommit=True)
                    except Exception:
                        pass
                except Exception:
                    _heartbeat_conn = None
            else:
                _safe_ping_conn(_heartbeat_conn)
        except Exception:
            logger.exception("Erro ao pingar conexão dedicada do heartbeat")

        # 2) ping em proxies ativas (conexões que foram entregues e ainda não devolvidas)
        try:
            proxies = list(_active_proxies)
            for p in proxies:
                try:
                    real_conn = getattr(p, "_conn", None)
                    if real_conn is not None:
                        _safe_ping_conn(real_conn)
                except Exception:
                    logger.exception("Erro ao pingar conexão ativa")
        except Exception:
            logger.exception("Erro no heartbeat ao iterar proxies ativas")

        # 3) força devolução de proxies que estão emprestados por tempo excessivo
        try:
            now = time.monotonic()
            proxies = list(_active_proxies)
            for p in proxies:
                try:
                    ts = getattr(p, "_checkout_ts", None)
                    if ts is not None and (now - ts) > _MAX_CHECKOUT_SECONDS:
                        force = getattr(p, "_force_close", None)
                        if callable(force):
                            logger.warning(
                                "Heartbeat: forçando devolução de proxy com %ds de uso (limite=%ds)",
                                int(now - ts), _MAX_CHECKOUT_SECONDS
                            )
                            try:
                                force()
                            except Exception:
                                logger.exception("Erro ao forçar devolução via heartbeat")
                except Exception:
                    logger.exception("Erro ao verificar/forçar proxy antigo")
        except Exception:
            logger.exception("Erro no heartbeat ao forçar devoluções de proxies antigos")

    logger.info("Heartbeat do pool finalizado.")


# ---------------------------
# ConnectionProxy (com watchdog)
# ---------------------------
class ConnectionProxy:
    """
    Proxy que encapsula a conexão real e garante liberação (close).
    Possui watchdog que força devolução após _MAX_CHECKOUT_SECONDS.
    """
    __slots__ = ("_conn", "_closed", "_checkout_ts", "_watchdog")

    def __init__(self, real_conn):
        global _total_checkouts
        self._conn = real_conn
        self._closed = False
        self._checkout_ts = time.monotonic()
        self._watchdog = None

        try:
            _active_proxies.add(self)
        except Exception:
            logger.debug("Falha ao registrar proxy no conjunto de proxies ativas", exc_info=False)
        with _metrics_lock:
            try:
                _total_checkouts += 1
            except Exception:
                pass

        # inicia watchdog (timer) que força retorno caso não seja fechado a tempo
        try:
            if _MAX_CHECKOUT_SECONDS and _MAX_CHECKOUT_SECONDS > 0:
                t = threading.Timer(_MAX_CHECKOUT_SECONDS, self._watchdog_cb)
                t.daemon = True
                self._watchdog = t
                t.start()
        except Exception:
            self._watchdog = None

    def _watchdog_cb(self):
        """Callback executado pelo timer: força fechamento se proxy ainda aberto."""
        try:
            if not getattr(self, "_closed", True):
                logger.warning("Watchdog: checkout excedeu %ds — forçando devolução.", _MAX_CHECKOUT_SECONDS)
                try:
                    # força rollback e devolução ao pool
                    self._force_close()
                except Exception:
                    logger.exception("Erro ao forçar devolução pela watchdog")
        except Exception:
            pass

    def _force_close(self):
        """Força a devolução/fechamento sem depender do usuário."""
        global _total_forced_gc_returns, _checkout_semaphore, _pool
        if getattr(self, "_closed", True):
            return
        self._closed = True

        real = getattr(self, "_conn", None)
        self._conn = None

        try:
            if real is not None:
                try:
                    real.rollback()
                except Exception:
                    pass
                # devolve ao pool ou fecha
                if _pool:
                    try:
                        _pool.putconn(real)
                    except Exception:
                        try:
                            real.close()
                        except Exception:
                            pass
                else:
                    try:
                        real.close()
                    except Exception:
                        pass
        finally:
            # remove do conjunto de proxies ativas
            try:
                _active_proxies.discard(self)
            except Exception:
                pass
            with _metrics_lock:
                try:
                    _total_forced_gc_returns += 1
                except Exception:
                    pass
            # libera semáforo
            try:
                if _checkout_semaphore is not None:
                    _checkout_semaphore.release()
            except Exception:
                pass
            # cancela watchdog (se ainda rodando)
            w = getattr(self, "_watchdog", None)
            if w is not None:
                try:
                    w.cancel()
                except Exception:
                    pass
                self._watchdog = None

    def cursor(self, *args, **kwargs):
        if self._conn is None:
            raise RuntimeError("Conexão já foi fechada/devolvida")
        return self._conn.cursor(*args, **kwargs)

    def commit(self):
        if self._conn is None:
            raise RuntimeError("Conexão já foi fechada/devolvida")
        return self._conn.commit()

    def rollback(self):
        if self._conn is None:
            raise RuntimeError("Conexão já foi fechada/devolvida")
        return self._conn.rollback()

    def set_client_encoding(self, enc):
        if self._conn is None:
            raise RuntimeError("Conexão já foi fechada/devolvida")
        try:
            return self._conn.set_client_encoding(enc)
        except Exception:
            logger.debug("Falha ao set_client_encoding", exc_info=False)

    def close(self):
        """
        Devolve a conexão ao pool e marca proxy como fechado.
        Cancela watchdog e libera semáforo.
        """
        global _pool, _checkout_semaphore, _total_returns
        if self._closed:
            return
        self._closed = True

        real_conn = getattr(self, "_conn", None)
        self._conn = None

        try:
            if real_conn is not None:
                try:
                    real_conn.rollback()
                except Exception:
                    pass
                if _pool:
                    try:
                        _pool.putconn(real_conn)
                    except Exception:
                        logger.debug("Erro ao devolver conexão ao pool; tentando fechar", exc_info=False)
                        try:
                            real_conn.close()
                        except Exception:
                            pass
                else:
                    try:
                        real_conn.close()
                    except Exception:
                        pass
        finally:
            # remove do conjunto de proxies ativas (se presente)
            try:
                _active_proxies.discard(self)
            except Exception:
                pass
            with _metrics_lock:
                try:
                    _total_returns += 1
                except Exception:
                    pass
            # libera o semáforo
            try:
                if _checkout_semaphore is not None:
                    _checkout_semaphore.release()
            except Exception:
                pass
            # cancela watchdog (se existir)
            w = getattr(self, "_watchdog", None)
            if w is not None:
                try:
                    w.cancel()
                except Exception:
                    pass
                self._watchdog = None

    # suporte ao 'with'
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    # garantir recuperação final se objeto for coletado (backup)
    def __del__(self):
        try:
            if not getattr(self, "_closed", True) and getattr(self, "_conn", None) is not None:
                try:
                    self._force_close()
                except Exception:
                    pass
        except Exception:
            pass

    def __getattr__(self, item):
        conn = object.__getattribute__(self, "_conn")
        if conn is None:
            raise AttributeError(f"conexão fechada/devolvida: {item}")
        return getattr(conn, item)


# ---------------------------
# Encerramento / utilitários
# ---------------------------
def fechar_todas_conexoes():
    """Fecha todas conexões do pool e para o heartbeat. Chame ao encerrar a aplicação."""
    global _pool, _heartbeat_stop, _heartbeat_thread, _heartbeat_conn, _checkout_semaphore

    # para heartbeat
    try:
        _heartbeat_stop.set()
    except Exception:
        pass

    # libera proxies ativas (devolve ao pool)
    try:
        proxies = list(_active_proxies)
        for p in proxies:
            try:
                p.close()
            except Exception:
                pass
    except Exception:
        logger.exception("Erro ao liberar proxies ativas")

    # fecha pool completamente
    with _pool_lock:
        if _pool is not None:
            try:
                _pool.closeall()
                logger.info("Todas conexões do pool foram fechadas.")
            except Exception:
                logger.exception("Erro ao fechar pool")
            finally:
                _pool = None

    # fecha conexão dedicada do heartbeat
    try:
        if _heartbeat_conn is not None:
            try:
                _heartbeat_conn.close()
            except Exception:
                pass
    except Exception:
        pass

    # aguarda thread terminar (se existir)
    try:
        if _heartbeat_thread is not None and _heartbeat_thread.is_alive():
            _heartbeat_thread.join(timeout=2.0)
    except Exception:
        pass

    # remove semaphore (garante que nova inicialização recrie)
    try:
        _checkout_semaphore = None
    except Exception:
        pass


def imprimir_metricas_pool():
    """Imprime métricas básicas do pool e proxies ativas."""
    try:
        ap = len(list(_active_proxies))
    except Exception:
        ap = -1
    try:
        pool_max = getattr(_pool, "maxconn", None) or getattr(_pool, "maxsize", None)
    except Exception:
        pool_max = None
    with _metrics_lock:
        tc = _total_checkouts
        tr = _total_returns
        tf = _total_forced_gc_returns
    try:
        print(f"METRICAS_POOL: total_checkouts={tc} total_returns={tr} forced_gc_returns={tf} active_proxies={ap} pool_max={pool_max}")
    except Exception:
        logger.exception("Falha ao imprimir métricas do pool")


def obter_metricas_uso():
    """Retorna dicionário com métricas úteis (para logs/monitor)."""
    try:
        active = len(list(_active_proxies))
    except Exception:
        active = -1
    with _metrics_lock:
        tc = _total_checkouts
        tr = _total_returns
        tf = _total_forced_gc_returns
    est_in_use = tc - tr
    return {
        "total_checkouts": tc,
        "total_returns": tr,
        "forced_gc_returns": tf,
        "estimated_in_use": est_in_use,
        "active_proxies": active,
        "pool_max": getattr(_pool, "maxconn", None) if _pool else None,
    }


# ---------------------------
# Função pública: conectar(...)
# ---------------------------
def conectar(ip=None, porta=None, usuario=None, senha=None, dbname=None, minconn=None, maxconn=None):
    """
    Mantém compatibilidade com sua função original:
      conn = conectar(ip=None, porta=None, usuario=None, senha=None, dbname=None)
    """
    global _pool, _conn_params, _pool_lock, _checkout_semaphore

    # importa TelaLogin dentro da função para compatibilidade com seu projeto
    try:
        from login import TelaLogin
        config = TelaLogin.carregar_configuracoes_static()
    except Exception:
        logger.debug("Não foi possível carregar TelaLogin; usando variáveis de ambiente/defaults.", exc_info=False)
        config = {}

    host = ip or config.get("DB_HOST", os.environ.get("DB_HOST", "192.168.1.117"))
    port = porta or config.get("DB_PORT", os.environ.get("DB_PORT", "5432"))
    user = usuario or config.get("DB_USER", os.environ.get("DB_USER", "Kametal"))
    password = senha or config.get("DB_PASSWORD", os.environ.get("DB_PASSWORD", "908513"))
    database = dbname or config.get("DB_NAME", os.environ.get("DB_NAME", "Teste"))

    logger.info("Tentando conectar: %s@%s:%s", database, host, port)

    conn_kwargs = dict(
        host=host,
        port=int(port),
        user=user,
        password=password,
        database=database,
    )

    minc = minconn if minconn is not None else _POOL_MIN
    maxc = maxconn if maxconn is not None else _POOL_MAX

    # Inicializa pool quando for a primeira chamada
    try:
        _init_pool_if_needed(minconn=minc, maxconn=maxc, **conn_kwargs)
    except Exception:
        logger.exception("Não foi possível inicializar pool de conexões")
        return None

    # usa o semaphore criado no init do pool
    _semaphore = _checkout_semaphore
    if _semaphore is None:
        # fallback local (não sobrescreve global)
        _semaphore = threading.BoundedSemaphore(maxc)

    # tenta adquirir semáforo (timeout para não bloquear indefinidamente)
    try:
        acquired = _semaphore.acquire(timeout=_SEM_TIMEOUT)
    except Exception:
        acquired = False

    if not acquired:
        logger.error("Timeout ao aguardar recurso do pool (semaphore).")
        return None

    # tentativas para obter conexão (backoff exponencial); mantemos semáforo até devolver ou proxy.close()
    attempts = 3
    delay = 0.15
    last_exc = None

    for attempt in range(1, attempts + 1):
        real_conn = None
        try:
            real_conn = _pool.getconn()

            # se conexão veio 'closed' ou inválida, fecha e recria
            if real_conn is None or (hasattr(real_conn, "closed") and real_conn.closed):
                logger.warning("Conexão retirada do pool estava closed/nula — criando nova conexão direta.")
                try:
                    if real_conn is not None:
                        real_conn.close()
                except Exception:
                    pass
                real_conn = connect(**_conn_params)

            # valida via ping; se ping falhar, tenta recriar connection direta
            if not _safe_ping_conn(real_conn):
                logger.debug("Ping falhou na conexão retirada do pool — recriando conexão direta.")
                try:
                    real_conn.close()
                except Exception:
                    pass
                real_conn = connect(**_conn_params)

            # aplicar configurações de sessão (encoding, timeouts)
            try:
                try:
                    real_conn.set_client_encoding("UTF8")
                except Exception:
                    pass

                cur = real_conn.cursor()
                try:
                    cur.execute(f"SET statement_timeout = {int(_STATEMENT_TIMEOUT_MS)}")
                    cur.execute("SET client_min_messages = WARNING")
                finally:
                    cur.close()
            except Exception:
                logger.debug("Falha ao aplicar configurações de sessão (não crítico).", exc_info=False)

            # retorno: proxy que gerencia devolução + watchdog
            proxy = ConnectionProxy(real_conn)
            return proxy

        except Exception as e:
            last_exc = e
            logger.exception("Erro ao obter conexão do pool (tentativa %d/%d): %s", attempt, attempts, e)

            # se real_conn foi criado mas algo deu errado, tenta devolver/fechar para não vazar
            try:
                if 'real_conn' in locals() and real_conn is not None:
                    try:
                        if _pool:
                            _pool.putconn(real_conn)
                        else:
                            real_conn.close()
                    except Exception:
                        try:
                            real_conn.close()
                        except Exception:
                            pass
            except Exception:
                pass

            time.sleep(delay)
            delay *= 2
            continue

    # se chegou aqui, falhou todas as tentativas -> liberar semáforo e retornar None
    try:
        if _checkout_semaphore is not None:
            _checkout_semaphore.release()
        else:
            try:
                _semaphore.release()
            except Exception:
                pass
    except Exception:
        pass

    logger.error("Não foi possível obter conexão do pool após %d tentativas. Último erro: %s", attempts, repr(last_exc))
    return None
