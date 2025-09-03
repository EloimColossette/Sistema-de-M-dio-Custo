import os
import time
import threading
import logging
import weakref
import concurrent.futures
from psycopg2 import pool as pg_pool
from psycopg2 import connect as pg_connect

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

# extras
_STATEMENT_TIMEOUT_MS = int(os.environ.get("DB_STATEMENT_TIMEOUT_MS", str(5 * 60 * 1000)))  # 5 minutos padrão
_APPLICATION_NAME = os.environ.get("DB_APPLICATION_NAME", "app_kametal")

# Executor/background workers (configurável)
_DB_WORKER_THREADS = int(os.environ.get("DB_WORKER_THREADS", "8"))
_executor = None

# ---------------------------
# Estado do módulo
# ---------------------------
_pool = None
_pool_lock = threading.Lock()
_conn_params = None

_active_proxies = weakref.WeakSet()
_active_proxies_lock = threading.Lock()

_heartbeat_thread = None
_heartbeat_stop = threading.Event()

_checkout_semaphore = None

_heartbeat_conn = None

_total_checkouts = 0
_total_returns = 0
_total_forced_gc_returns = 0
_metrics_lock = threading.Lock()


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
            _heartbeat_conn = pg_connect(**_conn_params)
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
                    _heartbeat_conn = pg_connect(**_conn_params)
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
            with _active_proxies_lock:
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
            with _active_proxies_lock:
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
    Watchdog apenas loga (não força devolução) na versão pedida.
    """
    __slots__ = ("_conn", "_closed", "_checkout_ts", "_watchdog")

    def __init__(self, real_conn):
        global _total_checkouts
        self._conn = real_conn
        self._closed = False
        self._checkout_ts = time.monotonic()
        self._watchdog = None

        try:
            with _active_proxies_lock:
                _active_proxies.add(self)
        except Exception:
            logger.debug("Falha ao registrar proxy no conjunto de proxies ativas", exc_info=False)
        with _metrics_lock:
            try:
                _total_checkouts += 1
            except Exception:
                pass

        # inicia watchdog (timer) que avisa se não for fechado a tempo (não força fechamento)
        try:
            if _MAX_CHECKOUT_SECONDS and _MAX_CHECKOUT_SECONDS > 0:
                t = threading.Timer(_MAX_CHECKOUT_SECONDS, self._watchdog_cb)
                t.daemon = True
                self._watchdog = t
                t.start()
        except Exception:
            self._watchdog = None

    def _watchdog_cb(self):
        """Callback executado pelo timer: apenas avisa que o checkout excedeu o limite."""
        try:
            if not getattr(self, "_closed", True):
                logger.warning(
                    "Watchdog: checkout excedeu %ds — conexão ainda aberta (não será forçada).",
                    _MAX_CHECKOUT_SECONDS
                )
                # não chamamos _force_close() aqui conforme solicitado
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
                with _active_proxies_lock:
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
                with _active_proxies_lock:
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
# Executor (helpers para execução em background)
# ---------------------------
def _init_executor_if_needed(max_workers_hint=None):
    """Cria ThreadPoolExecutor se ainda não existir. Chamado dentro de conectar()."""
    global _executor
    if _executor is not None:
        return
    maxw = _DB_WORKER_THREADS
    if max_workers_hint:
        try:
            maxw = min(maxw, int(max_workers_hint))
        except Exception:
            pass
    try:
        _executor = concurrent.futures.ThreadPoolExecutor(max_workers=max(1, maxw))
        logger.info("DB executor iniciado (workers=%d)", maxw)
    except Exception:
        _executor = None
        logger.exception("Falha ao iniciar DB executor")


def _direct_execute(sql, params=None, commit=False, fetch=False, statement_timeout_ms=None, conn_params=None):
    """
    Abre uma conexão direta (psycopg2.connect) usando conn_params,
    executa e fecha. Retorna rows se fetch=True.
    Projetado para ser executado em executor (background).
    """
    if conn_params is None:
        raise RuntimeError("_direct_execute precisa de conn_params válidos")
    conn = None
    cur = None
    try:
        conn = pg_connect(**conn_params)
        try:
            cur = conn.cursor()
            if statement_timeout_ms is not None:
                cur.execute(f"SET statement_timeout = {int(statement_timeout_ms)}")
            cur.execute(sql, params or ())
            if commit:
                conn.commit()
            if fetch:
                rows = cur.fetchall()
                return rows
            return None
        finally:
            if cur is not None:
                try: cur.close()
                except Exception: pass
    finally:
        if conn is not None:
            try: conn.close()
            except Exception: pass


def _db_worker_execute(sql, params=None, commit=False, fetch=False, statement_timeout_ms=None):
    """
    Executa query no worker: obtém conexão com `conectar()`, executa, (commit opcional) e fecha.
    Retorna rows se fetch=True, caso contrário None.
    """
    conn = None
    try:
        conn = conectar()  # usar conectar() garante que o pool/executor estejam prontos
        if conn is None:
            raise RuntimeError("Não foi possível obter conexão para execução de query")
        # Se receber AsyncConnectionProxy (caso raro aqui), usamos sua API para submeter
        if hasattr(conn, 'execute_async') and not hasattr(conn, 'cursor'):
            fut = conn.submit(sql, params=params, commit=commit, fetch=fetch, statement_timeout_ms=statement_timeout_ms)
            return fut.result()

        # caso padrão: ConnectionProxy (síncrono)
        with conn:
            cur = conn.cursor()
            try:
                if statement_timeout_ms is not None:
                    cur.execute(f"SET statement_timeout = {int(statement_timeout_ms)}")
                cur.execute(sql, params or ())
                if commit:
                    try:
                        conn.commit()
                    except Exception:
                        logger.exception("Erro ao commitar")
                        raise
                if fetch:
                    return cur.fetchall()
                return None
            finally:
                try:
                    cur.close()
                except Exception:
                    pass
    finally:
        # nada a fazer aqui — ConnectionProxy.close() cuida da devolução
        pass


def execute_async(sql, params=None, commit=False, fetch=False,
                  statement_timeout_ms=None, callback=None, error_callback=None):
    """
    Envia a query para execução em background e retorna Future.
    - callback(result) será chamado (no thread do executor) quando a execução terminar com sucesso.
    - error_callback(exc) será chamado (no thread do executor) se der erro.
    """
    global _executor
    if _executor is None:
        # tenta inicializar com base no pool max (se disponível)
        try:
            maxhint = getattr(_pool, "maxconn", None) or getattr(_pool, "maxsize", None)
        except Exception:
            maxhint = None
        _init_executor_if_needed(maxhint)

    if _executor is None:
        raise RuntimeError("Executor não disponível para execução assíncrona")

    fut = _executor.submit(_db_worker_execute, sql, params, commit, fetch, statement_timeout_ms)

    if callback or error_callback:
        def _done_cb(f):
            try:
                res = f.result()
                if callback:
                    try:
                        callback(res)
                    except Exception:
                        logger.exception("Erro no callback de sucesso")
            except Exception as e:
                if error_callback:
                    try:
                        error_callback(e)
                    except Exception:
                        logger.exception("Erro no callback de erro")
                else:
                    logger.exception("Erro na execução assíncrona")
        fut.add_done_callback(_done_cb)

    return fut


def execute_sync_with_timeout(sql, params=None, commit=False, fetch=False,
                              statement_timeout_ms=None, timeout=None):
    """
    Executa a query em background, mas aguarda o resultado até `timeout` segundos.
    Retorna o resultado (se fetch=True) ou None. Lança TimeoutError se exceder tempo.
    Use isso com cautela na UI principal (melhor ainda usar execute_async + callback).
    """
    fut = execute_async(sql=sql, params=params, commit=commit, fetch=fetch,
                        statement_timeout_ms=statement_timeout_ms)
    try:
        return fut.result(timeout=timeout)
    except concurrent.futures.TimeoutError:
        try:
            fut.cancel()
        except Exception:
            pass
        raise


# ---------------------------
# AsyncConnectionProxy: retornado imediatamente na MainThread se pool não existir
# ---------------------------
class AsyncConnectionProxy:
    """
    Proxy retornado rapidamente quando conectar() é chamado da MainThread
    e o pool ainda nao foi inicializado. Fornece API assíncrona:
      - execute_async(...)
      - submit(...)
    Também fornece um fallback limitado: se o código tentar usar .cursor()
    ou outro atributo, o proxy aguarda curto período pelo pool e encaminha
    para um ConnectionProxy real (se o pool ficar pronto).
    """

    def __init__(self, conn_params):
        self._conn_params = conn_params
        _init_executor_if_needed()
        self._pending = []  # lista de futures (opcional para tracking)
        self._closed = False

    # ---------- helpers de espera ----------
    def _wait_for_pool_ready(self, timeout=2.0, poll=0.05):
        """Espera até `timeout` segundos o pool ficar pronto. Retorna True se pronto."""
        end = time.monotonic() + float(timeout)
        while _pool is None and time.monotonic() < end:
            time.sleep(poll)
        return _pool is not None

    def _get_sync_proxy_or_none(self, timeout=2.0):
        """
        Se o pool ficar pronto dentro do timeout, obtém e retorna um ConnectionProxy
        chamando conectar() (o fluxo normal). Caso contrário retorna None.
        """
        if not self._wait_for_pool_ready(timeout=timeout):
            return None
        # agora que o pool deve estar pronto, chamar conectar() deve retornar ConnectionProxy
        try:
            p = conectar()  # isto obtém ConnectionProxy síncrono (rápido se pool pronto)
            return p
        except Exception:
            return None

    # ---------- compatibilidade com cursor() (fallback) ----------
    def cursor(self, *args, timeout=2.0, **kwargs):
        """
        Tenta obter um ConnectionProxy real dentro de `timeout` segundos e
        retorna sua cursor() — se não for possível, levanta RuntimeError com instrução.
        Atenção: isso pode bloquear até `timeout` segundos.
        """
        p = self._get_sync_proxy_or_none(timeout=timeout)
        if p is None:
            raise RuntimeError(
                "Pool ainda não pronto (timeout). Use execute_async/submit para chamadas não-bloqueantes "
                "ou inicialize o pool em background no startup."
            )
        # p é um ConnectionProxy: delega
        return p.cursor(*args, **kwargs)

    # ---------- delegação genérica (tentativa) ----------
    def __getattr__(self, name):
        """
        Se algum código acessar outro método/atributo (por exemplo commit/rollback),
        tentamos obter um proxy síncrono rapidamente e delegar. Caso não seja possível,
        caímos para as APIs assíncronas ou erro instrutivo.
        """
        # métodos óbvios que suportamos assíncrono: execute_async, submit, close estão aqui
        if name in ("execute_async", "submit", "close"):
            return object.__getattribute__(self, name)

        # caso contrário, tentamos obter proxy síncrono com timeout pequeno
        p = self._get_sync_proxy_or_none(timeout=1.5)
        if p is None:
            raise AttributeError(
                f"'{self.__class__.__name__}' object has no attribute '{name}' (pool não pronto). "
                "Use execute_async/submit ou aguarde a inicialização do pool."
            )
        return getattr(p, name)

    # ---------- APIs assíncronas existentes ----------
    def submit(self, sql, params=None, commit=False, fetch=False, statement_timeout_ms=None):
        if self._closed:
            raise RuntimeError("Proxy já fechado")
        fut = _executor.submit(
            _direct_execute,
            sql, params, commit, fetch, statement_timeout_ms, self._conn_params
        )
        try:
            self._pending.append(fut)
        except Exception:
            pass

        def _done(_f):
            try:
                self._pending.remove(_f)
            except Exception:
                pass
        fut.add_done_callback(_done)
        return fut

    def execute_async(self, sql, params=None, commit=False, fetch=False,
                      statement_timeout_ms=None, callback=None, error_callback=None):
        """
        Conveniência: submete a execução e chama callback(result) no thread do executor.
        ATENÇÃO: callback é executado no thread do executor — se for atualizar UI, use root.after(...)
        """
        fut = self.submit(sql, params=params, commit=commit, fetch=fetch, statement_timeout_ms=statement_timeout_ms)

        if callback or error_callback:
            def _done_cb(f):
                try:
                    res = f.result()
                    if callback:
                        try:
                            callback(res)
                        except Exception:
                            logger.exception("Erro no callback de sucesso (AsyncConnectionProxy)")
                except Exception as e:
                    if error_callback:
                        try:
                            error_callback(e)
                        except Exception:
                            logger.exception("Erro no callback de erro (AsyncConnectionProxy)")
                    else:
                        logger.exception("Erro na execução assíncrona (AsyncConnectionProxy)")
            fut.add_done_callback(_done_cb)
        return fut

    def close(self):
        # não força nada — apenas marca fechado e deixa pendências terminarem
        self._closed = True

    def __enter__(self):
        raise RuntimeError("AsyncConnectionProxy não suporta 'with' (para evitar bloqueio). Use execute_async/submit.")

    def __exit__(self, exc_type, exc, tb):
        self.close()


# ---------------------------
# Encerramento / utilitários
# ---------------------------
def fechar_todas_conexoes():
    """Fecha todas conexões do pool e para o heartbeat. Chame ao encerrar a aplicação."""
    global _pool, _heartbeat_stop, _heartbeat_thread, _heartbeat_conn, _checkout_semaphore, _executor

    # para heartbeat
    try:
        _heartbeat_stop.set()
    except Exception:
        pass

    # libera proxies ativas (devolve ao pool)
    try:
        with _active_proxies_lock:
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

    # encerra executor
    try:
        if _executor is not None:
            try:
                _executor.shutdown(wait=False)
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
        with _active_proxies_lock:
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
        with _active_proxies_lock:
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

    Comportamento especial:
    - Se chamado da MainThread e o pool NÃO existir, inicia inicialização em background
      e retorna AsyncConnectionProxy imediatamente (não bloqueante).
    - Se não estiver na MainThread (worker thread) e o pool não existir, inicializa
      sincronamente (blocking) e retorna ConnectionProxy.
    """
    global _pool, _conn_params, _pool_lock, _checkout_semaphore, _executor

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

    # Se o pool já foi inicializado -> comportamento normal e rápido
    if _pool is not None:
        # segue fluxo original: adquirir semáforo e getconn()...
        _semaphore = _checkout_semaphore
        if _semaphore is None:
            _semaphore = threading.BoundedSemaphore(maxc)

        try:
            acquired = _semaphore.acquire(timeout=_SEM_TIMEOUT)
        except Exception:
            acquired = False

        if not acquired:
            logger.error("Timeout ao aguardar recurso do pool (semaphore).")
            return None

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
                    real_conn = pg_connect(**_conn_params)

                # valida via ping; se ping falhar, tenta recriar connection direta
                if not _safe_ping_conn(real_conn):
                    logger.debug("Ping falhou na conexão retirada do pool — recriando conexão direta.")
                    try:
                        real_conn.close()
                    except Exception:
                        pass
                    real_conn = pg_connect(**_conn_params)

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

    # --- Novo comportamento: se estiver na MainThread e pool ainda não criado,
    # inicializa o pool em background e retorna AsyncConnectionProxy IMEDIATAMENTE ---
    if threading.current_thread().name == "MainThread" and _pool is None:
        with _pool_lock:
            if _conn_params is None:
                _conn_params = conn_kwargs.copy()
                _conn_params.setdefault("connect_timeout", _CONNECT_TIMEOUT)
                _conn_params.setdefault("keepalives", 1)
                _conn_params.setdefault("keepalives_idle", _CONN_KEEPALIVE_IDLE)
                _conn_params.setdefault("keepalives_interval", _CONN_KEEPALIVE_INTERVAL)
                _conn_params.setdefault("keepalives_count", _CONN_KEEPALIVE_COUNT)
                _conn_params.setdefault("application_name", _APPLICATION_NAME)

        def _bg_init():
            try:
                _init_pool_if_needed(minconn=minc, maxconn=maxc, **_conn_params)
            except Exception:
                logger.exception("Inicialização background do pool falhou")

        t = threading.Thread(target=_bg_init, daemon=True, name="db-init-bg")
        t.start()

        # garante executor também
        _init_executor_if_needed(maxc)

        # retorna proxy assíncrono imediato (não bloqueante)
        return AsyncConnectionProxy(_conn_params)

    # --- Caso padrão (não main thread ou pool já existente) continua o fluxo síncrono normal ---
    try:
        _init_pool_if_needed(minconn=minc, maxconn=maxc, **conn_kwargs)
    except Exception:
        logger.exception("Não foi possível inicializar pool de conexões")
        return None

    # inicializa executor etc.
    _init_executor_if_needed(maxc)

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
                real_conn = pg_connect(**_conn_params)

            # valida via ping; se ping falhar, tenta recriar connection direta
            if not _safe_ping_conn(real_conn):
                logger.debug("Ping falhou na conexão retirada do pool — recriando conexão direta.")
                try:
                    real_conn.close()
                except Exception:
                    pass
                real_conn = pg_connect(**_conn_params)

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
