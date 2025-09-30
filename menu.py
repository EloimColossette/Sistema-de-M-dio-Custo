import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from logos import aplicar_icone
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
from conexao_db import conectar, logger
import threading
import select
import time
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
import numpy as np
import seaborn as sns
import mplcursors
from datetime import datetime
import pandas as pd
import pandas.api.types as ptypes
import importlib
import queue
import traceback
from datetime import datetime, date, timezone
import calendar
from versao import __version__

class Janela_Menu(tk.Tk):
    def __init__(self, user_id):
        super().__init__()
        self.user_id = user_id

        # --- Conexão com o banco de dados e Permissão ---
        self.conn = conectar()
        # Conexão separada para escutar notificações
        self.conn_listen = conectar()
        # garante que pg_notify / LISTEN funcione sem commit
        self.conn_listen.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        self._cur_listen = self.conn_listen.cursor()
        self._cur_listen.execute("LISTEN canal_atualizacao;")

        # --- atributos de controle de thread / estado (existem desde o início) ---
        self.encerrar = False
        self._stop_event = threading.Event()
        self._closing = False
        self.hora_job = None
        self._spinner_job = None
        self.atualizacao_pendente = False

        # Carrega permissões / nome (pode depender de conexões)
        self.permissoes = self.carregar_permissoes_usuario(user_id)
        self.user_name = self.carregar_nome_usuario(user_id)

        # --- Configurações gerais da janela ---
        self.title("Tela Inicial")
        self.geometry("1200x700")
        self.configure(bg="#ecf0f1")
        self.state("zoomed")
        self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        try:
            aplicar_icone(self, self.caminho_icone)
        except Exception:
            # não crítico — segue sem ícone se falhar
            pass

        # --- Cria os componentes da interface ---
        self._criar_sidebar()
        self.hora_job = None
        self._criar_cabecalho()      # atenção: este método deve criar self.right_frame
        self._criar_main_content()
        self._criar_menubar()
        self._criar_menu_lateral()
        self.configurar_estilos()

        # cria barra de abas customizada (setas + canvas)
        self._criar_barra_abas()

        # cria o notebook
        self.notebook = ttk.Notebook(self.main_content, style="Hidden.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        # --- Configurações do sistema de aviso de purga ---
        # thresholds (ajuste conforme desejar)
        self.purge_warning_days = 7
        self.purge_critical_days = 3

        # flag para evitar abrir o modal automaticamente várias vezes no mesmo ciclo
        self._auto_purge_modal_shown = False

        # estado do flash
        self._purge_flashing = False
        self._purge_flash_state = False
        self._last_purge_days = None

        # id do after para o piscar (para poder cancelar na destruição)
        self._purge_flash_job = None

        # cria abas mínimas (placeholder)
        self._criar_abas_minimal()
        self.update_idletasks()

        # --- somente inicializa o sistema de purga se o usuário tiver permissão ---
        if "Calculo_Produto" in self.permissoes:
            try:
                # primeira verificação da purga — garante badge correta antes da thread iniciar
                self.check_purge_status()
            except Exception as e:
                print("Erro em check_purge_status() na inicialização:", e)
        else:
            # garante que os atributos existam mesmo quando não houver permissão
            self.purge_frame = None
            self.purge_btn = None
            self.purge_badge = None
            self.purge_badge_amarelo = None

        # --- Agora sim: Inicia a thread de escuta APÓS os widgets existirem ---
        self._thread_escuta = threading.Thread(target=self._escutar_notificacoes, daemon=True)
        self._thread_escuta.start()

        # agenda carregamento pesado em segundo plano
        def _carregar_abas_pesadas():
            try:
                # chama a versão pesada depois que a janela já estiver visível
                self.after(1500, self._criar_abas)
            except Exception as e:
                print("Erro ao carregar abas:", e)

        threading.Thread(target=_carregar_abas_pesadas, daemon=True).start()

        # sincroniza barra de abas
        self._atualizar_barra_abas()

        # prewarm modules (opcional)
        self.prewarm_modules([
            "Base_produto", "Base_material", "Saida_NF", "Insercao_NF",
            "usuario", "Estoque", "media_custo", "relatorio_saida",
            "relatorio_cotacao", "registro_teste"
        ])

        # binds / fechamento
        self.bind("<FocusIn>", lambda e: self.configurar_estilos())
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def carregar_permissoes_usuario(self, user_id):
        """Carrega as permissões do usuário e retorna um conjunto com os nomes internos das permissões."""
        if not self.conn:
            return set()
        cursor = self.conn.cursor()
        cursor.execute("SELECT janela FROM permissoes WHERE usuario_id = %s", (user_id,))
        permissoes = {row[0] for row in cursor.fetchall()}
        cursor.close()
        return permissoes

    def _criar_sidebar(self):
        """Cria a sidebar de navegação com largura fixa (não encolhe mesmo se não houver botões)."""
        # Frame principal da sidebar com largura fixa
        sidebar_frame = tk.Frame(self, bg="#2c3e50", width=220)
        sidebar_frame.pack(side="left", fill="y")

        # Canvas para permitir rolagem (largura fixa para manter layout)
        self.sidebar_canvas = tk.Canvas(
            sidebar_frame,
            bg="#2c3e50",
            highlightthickness=0,
            width=220,
            bd=0
        )
        # importante: não usar expand=True aqui para evitar que o canvas cresça/encolha horizontalmente
        self.sidebar_canvas.pack(side="left", fill="y")

        # Scrollbar vertical
        scrollbar = tk.Scrollbar(
            sidebar_frame,
            orient="vertical",
            command=self.sidebar_canvas.yview,
            width=15,
            troughcolor="#1a252f",
            bg="#2980b9",
            activebackground="#3498db"
        )
        scrollbar.pack(side="right", pady=(60, 0), fill="y")
        self.sidebar_canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno dentro do canvas (onde os botões serão colocados)
        # definimos width=220 e pack_propagate(False) para que também não encolha
        self.sidebar_inner = tk.Frame(self.sidebar_canvas, bg="#2c3e50", width=220)

        # cria a window no canvas e guarda o id para poder ajustar largura quando o canvas mudar
        self.sidebar_window = self.sidebar_canvas.create_window((0, 0), window=self.sidebar_inner, anchor="nw")

        # quando o conteúdo interno mudar, atualiza a região de rolagem
        def _on_inner_config(e):
            # deixa um pouco de folga (200px) para reduzir o tamanho da thumb da scrollbar
            self.sidebar_canvas.configure(scrollregion=(0, 0, e.width, e.height + 200))

        self.sidebar_inner.bind("<Configure>", _on_inner_config)

        # quando o canvas mudar de tamanho (ex.: redimensionamento da janela), ajusta a largura da window interna
        def _on_canvas_config(e):
            try:
                # força a janela interna a ter a mesma largura do canvas (evita encolhimento)
                self.sidebar_canvas.itemconfig(self.sidebar_window, width=e.width)
            except Exception:
                pass

        self.sidebar_canvas.bind("<Configure>", _on_canvas_config)

        # rolagem com roda do mouse
        def _on_mousewheel(event):
            # padrão Windows: event.delta múltiplo de 120
            try:
                self.sidebar_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                # fallback mais simples
                self.sidebar_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        # bind global apenas para a sidebar (não deve competir com outros widgets)
        self.sidebar_canvas.bind_all("<MouseWheel>", _on_mousewheel)
      
    def criar_botao_redondo(self, master, comando=None, texto="❓", raio=14,
                        bg="#2c3e50", fg="white",
                        hover_enter_callback=None, hover_leave_callback=None):
        """
        Botão redondo com animações:
        - hover: cresce levemente e mostra sombra
        - click: ripple + callback
        - teclado: suporta Tab/Enter/Space e anel de foco
        Retorna: Canvas (widget)
        """
        # estado e parâmetros
        base_raio = raio
        max_raio = int(raio * 1.14)   # até quanto cresce no hover
        speed_ms = 15                 # intervalo de animação (menor = mais suave)
        step = 1                      # pixels por frame
        current_raio = base_raio
        anim_id = {"hover": None, "ripple": None}

        # canvas com mesmo bg do pai para integrar visual
        canvas = tk.Canvas(master, width=(max_raio*2)+8, height=(max_raio*2)+8,
                        highlightthickness=0, bg=master["bg"], bd=0, relief="flat", takefocus=1, cursor="hand2")

        cx = (max_raio + 4)
        cy = (max_raio + 4)

        # shadow (atrás, ligeiramente deslocada para baixo)
        shadow_offset = 2
        shadow_color = "#172029"
        shadow_id = canvas.create_oval(cx-base_raio+shadow_offset, cy-base_raio+shadow_offset,
                                    cx+base_raio+shadow_offset, cy+base_raio+shadow_offset,
                                    fill=shadow_color, outline=shadow_color)

        # círculo principal
        oval_id = canvas.create_oval(cx-base_raio, cy-base_raio, cx+base_raio, cy+base_raio,
                                    fill=bg, outline=bg)
        text_id = canvas.create_text(cx, cy, text=texto, fill=fg, font=("Helvetica", max(10, int(base_raio*0.9)), "bold"))

        # anel de foco (invisível até focus)
        focus_id = canvas.create_oval(cx-base_raio-6, cy-base_raio-6, cx+base_raio+6, cy+base_raio+6,
                                    outline="", width=2, dash=(3, 4))

        # helpers para redesenhar (muda coords e fonte conforme raio)
        def _desenhar(raio_atual):
            # atualiza shadow, oval, texto, focus ring
            canvas.coords(shadow_id,
                        cx-raio_atual+shadow_offset, cy-raio_atual+shadow_offset,
                        cx+raio_atual+shadow_offset, cy+raio_atual+shadow_offset)
            canvas.coords(oval_id,
                        cx-raio_atual, cy-raio_atual,
                        cx+raio_atual, cy+raio_atual)
            canvas.coords(text_id,
                        cx, cy)
            # fonte proporcional (inteiro)
            fsize = max(8, int(raio_atual * 0.9))
            canvas.itemconfig(text_id, font=("Helvetica", fsize, "bold"))
            canvas.coords(focus_id,
                        cx-raio_atual-6, cy-raio_atual-6,
                        cx+raio_atual+6, cy+raio_atual+6)

        # animação de hover: expandir/contrair suavemente
        def _animate_hover(target):
            # cancela animação anterior (se houver)
            if anim_id["hover"]:
                canvas.after_cancel(anim_id["hover"])
                anim_id["hover"] = None

            def step_fn():
                nonlocal current_raio
                if current_raio == target:
                    anim_id["hover"] = None
                    return
                # aproxima-se do target
                if current_raio < target:
                    current_raio = min(target, current_raio + step)
                else:
                    current_raio = max(target, current_raio - step)
                _desenhar(current_raio)
                anim_id["hover"] = canvas.after(speed_ms, step_fn)

            step_fn()

        # ripple ao clicar: cria um anel que cresce e some
        def _ripple():
            rip_id = canvas.create_oval(cx, cy, cx, cy, outline="#ffffff", width=2)
            max_r = int(base_raio * 2.4)
            frames = 10
            frame = {"i": 0}

            def ripple_step():
                frame["i"] += 1
                t = frame["i"] / frames
                r = int(base_raio + t * (max_r - base_raio))
                canvas.coords(rip_id, cx - r, cy - r, cx + r, cy + r)
                # reduz opacidade simulando trocando cor por tons (simples) -> aqui usa cinza claro progressivo
                alpha_val = int(220 * (1 - t))  # 220..0
                # transforma em hex (não alpha real) para simular fade
                hex_gray = f"#{alpha_val:02x}{alpha_val:02x}{alpha_val:02x}"
                canvas.itemconfig(rip_id, outline=hex_gray)
                if frame["i"] < frames:
                    anim_id["ripple"] = canvas.after(30, ripple_step)
                else:
                    canvas.delete(rip_id)
                    anim_id["ripple"] = None

            ripple_step()

        # eventos
        def _on_enter(evt=None):
            # cresce e ressalta
            _animate_hover(max_raio)
            # leve escurecimento do oval para indicar hover
            canvas.itemconfig(oval_id, fill="#374e60", outline="#374e60")
            # optional callback
            if callable(hover_enter_callback):
                try:
                    hover_enter_callback()
                except Exception:
                    pass

        def _on_leave(evt=None):
            _animate_hover(base_raio)
            canvas.itemconfig(oval_id, fill=bg, outline=bg)
            if callable(hover_leave_callback):
                try:
                    hover_leave_callback()
                except Exception:
                    pass

        def _on_click(evt=None):
            # ripple visual e comando
            _ripple()
            # press visual rápido (um contorno branco que some)
            canvas.itemconfig(oval_id, outline="#ffffff")
            canvas.after(90, lambda: canvas.itemconfig(oval_id, outline=bg))
            if callable(comando):
                try:
                    comando()
                except Exception:
                    # não deixar animação travar por exceção
                    pass

        # foco por teclado: mostra anel
        def _on_focus_in(evt=None):
            canvas.itemconfig(focus_id, outline="#ffffff")
        def _on_focus_out(evt=None):
            canvas.itemconfig(focus_id, outline="")

        # binds
        canvas.bind("<Enter>", lambda e: _on_enter(e))
        canvas.bind("<Leave>", lambda e: _on_leave(e))
        canvas.bind("<Button-1>", lambda e: _on_click(e))
        canvas.bind("<Return>", lambda e: _on_click(e))
        canvas.bind("<space>", lambda e: _on_click(e))
        canvas.bind("<FocusIn>", lambda e: _on_focus_in(e))
        canvas.bind("<FocusOut>", lambda e: _on_focus_out(e))

        # inicial desenha
        _desenhar(current_raio)

        return canvas

    def _criar_cabecalho(self):
        """Cria o cabeçalho com a data, hora, título e botões à direita."""
        self.cabecalho = tk.Frame(self, bg="#34495e", height=80)
        self.cabecalho.pack(side="top", fill="x")

        # grid 3 colunas...
        self.cabecalho.columnconfigure(0, weight=1)
        self.cabecalho.columnconfigure(1, weight=2)
        self.cabecalho.columnconfigure(2, weight=1)

        # Coluna 0: Data e Hora
        self.info_frame = tk.Frame(self.cabecalho, bg="#34495e")
        self.info_frame.grid(row=0, column=0, padx=15, sticky="w")

        self.data_label = tk.Label(self.info_frame, text="", fg="white", bg="#34495e",
                                   font=("Helvetica", 16, "bold"))
        self.data_label.pack(side="top", anchor="w")

        self.hora_label = tk.Label(self.info_frame, text="", fg="#f1c40f", bg="#34495e",
                                   font=("Helvetica", 16, "bold"))
        self.hora_label.pack(side="top", anchor="w")

        # Coluna 1: título centralizado
        titulo_label = tk.Label(self.cabecalho, text="Sistema Kametal", fg="white", bg="#34495e",
                                font=("Helvetica", 26, "bold"))
        titulo_label.grid(row=0, column=1, padx=10)

        # Coluna 2: lado direito
        self.right_frame = tk.Frame(self.cabecalho, bg="#34495e")
        self.right_frame.grid(row=0, column=2, padx=10, sticky="e")

        # Botão Atualizar (sempre visível)
        self.botao_atualizar = tk.Button(
            self.right_frame, text="Atualizar", fg="white", bg="#2980b9",
            font=("Helvetica", 12, "bold"), bd=0, relief="flat", command=self.atualizar_pagina
        )
        self.botao_atualizar.bind("<Enter>", lambda e: self.botao_atualizar.config(bg="#3498db"))
        self.botao_atualizar.bind("<Leave>", lambda e: self.botao_atualizar.config(bg="#2980b9"))
        self.botao_atualizar.pack(side="right", padx=(8, 0))

       # Botão Ajuda (profissional, com tooltip e atalho F1)
        self.help_frame = tk.Frame(self.right_frame, bg="#34495e")
        self.help_frame.pack(side="right", padx=(6, 4))

        self.botao_ajuda = self.criar_botao_redondo(
            self.help_frame,
            comando=self._abrir_ajuda_modal,
            texto="❓",
            raio=14,
            bg="#2c3e50",
            fg="white",
            hover_enter_callback=lambda: None,  # opcional: passe algo se quiser executar
            hover_leave_callback=lambda: None
        )
        self.botao_ajuda.pack(side="right", padx=(0, 6))

        try:
            # seu _create_tooltip deve aceitar o canvas; se não aceitar, adapte para usar widget=canvas
            self._create_tooltip(self.botao_ajuda, "Ajuda (F1)")
        except Exception:
            pass

        # Atalho global F1 para abrir Ajuda
        try:
            self.bind_all("<F1>", lambda e: self._abrir_ajuda_modal())
        except Exception:
            pass

        # --- só cria o sino/badges se o usuário tiver a permissão para Cálculo de NFs ---
        if "Calculo_Produto" in self.permissoes:
            self.purge_frame = tk.Frame(self.right_frame, bg="#34495e")
            self.purge_frame.pack(side="right", padx=(0, 4))

            # Botão sino
            self.purge_btn = tk.Button(
                self.purge_frame, text="🔔", fg="white", bg="#34495e",
                font=("Helvetica", 18), bd=0, relief="flat", 
                command=lambda: self.mostrar_aviso_purga(force=True)
            )
            self.purge_btn.pack(side="right")

            # Badge vermelho (notificação)
            self.purge_badge = tk.Label(
                self.purge_frame, text="", bg="#e74c3c", fg="white",
                font=("Helvetica", 8, "bold"), padx=4, pady=0
            )
            self.purge_badge.place(in_=self.purge_btn, relx=1.0, rely=0.0, anchor="ne", x=0, y=0)
            self.purge_badge.place_forget()  # começa escondido

            # Badge amarelo (opcional)
            self.purge_badge_amarelo = tk.Label(
                self.purge_frame, text="", bg="#f1c40f", fg="black",
                font=("Helvetica", 8, "bold"), padx=4, pady=0
            )
            self.purge_badge_amarelo.place(in_=self.purge_btn, relx=0.7, rely=0.0, anchor="ne", x=0, y=0)
            self.purge_badge_amarelo.place_forget()
        else:
            # define None para evitar AttributeError em outros trechos
            self.purge_frame = None
            self.purge_btn = None
            self.purge_badge = None
            self.purge_badge_amarelo = None

        # Atualiza hora inicial
        self.atualizar_hora()

    def atualizar_hora(self):
        now = datetime.now()
        data_formatada = now.strftime("%d/%m/%Y")
        self.data_label.config(text=data_formatada)

        hora_formatada = now.strftime("%H:%M:%S")
        self.hora_label.config(text=hora_formatada)

        # Armazena o ID do callback para que ele possa ser cancelado depois
        self.hora_job = self.after(1000, self.atualizar_hora)

    def criar_widgets_purge(self):
        # método só para manter estrutura — chamamos check direto no init
        self.check_purge_status()

    def check_purge_status(self, dias_aviso=None):
        """
        Atualiza a badge com base nos dias restantes até o fim do mês.
        - mostra nada se maior que purge_warning_days
        - mostra badge amarela se <= purge_warning_days
        - mostra badge vermelha + piscar se <= purge_critical_days
        - abre automaticamente o modal **uma vez** quando faltar 2 dias ou menos
        """
        # Proteção: não faz nada se o usuário não tiver a permissão (ou se os widgets não existirem)
        if "Calculo_Produto" not in self.permissoes or getattr(self, "purge_badge", None) is None:
            return
        
        try:
            if dias_aviso is None:
                dias_aviso = self.purge_warning_days

            hoje = date.today()
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            dias_restantes = ultimo_dia - hoje.day

            # se for maior que aviso -> esconde badge e para flash
            if dias_restantes > self.purge_warning_days:
                if self.purge_badge.winfo_ismapped():
                    self.purge_badge.pack_forget()
                # parar flash se estava piscando
                if self._purge_flashing:
                    self._purge_flashing = False
                    try:
                        self.purge_badge.config(bg="#f1c40f", fg="black")
                    except Exception:
                        pass
            else:
                # mostra badge com número
                self.purge_badge.config(text=str(dias_restantes))
                if not self.purge_badge.winfo_ismapped():
                    self.purge_badge.pack(side="right", padx=(0, 6))

                # crítico?
                if dias_restantes <= self.purge_critical_days:
                    # cor vermelha e inicia flash
                    self.purge_badge.config(bg="#e74c3c", fg="white")
                    if not self._purge_flashing:
                        self._purge_flashing = True
                        self._purge_flash_state = False
                        self._flash_purge_badge()
                else:
                    # nível aviso (amarelo)
                    self.purge_badge.config(bg="#f1c40f", fg="black")
                    if self._purge_flashing:
                        self._purge_flashing = False
                        self.purge_badge.config(bg="#f1c40f", fg="black")

            # --- ABRIR AUTOMATICAMENTE O MODAL quando faltar 2 dias ou menos ---
            try:
                # se faltar 2 dias ou menos e ainda não abrimos automaticamente, abre modal
                if dias_restantes <= 2 and not getattr(self, "_auto_purge_modal_shown", False):
                    # usa after para garantir que a chamada ocorra na thread principal de UI
                    try:
                        self.after(200, lambda: self.mostrar_aviso_purga(force=False))
                    except Exception:
                        # fallback direto (pelo menos tenta)
                        try:
                            self.mostrar_aviso_purga(force=False)
                        except Exception:
                            pass
                    # marca como já mostrado para não ficar abrindo repetidamente
                    self._auto_purge_modal_shown = True
                # se estamos acima de 2 dias, resetamos a flag (próximo mês / novo ciclo)
                elif dias_restantes > 2:
                    self._auto_purge_modal_shown = False
            except Exception:
                pass

            # --- Agendar próxima checagem (10 minutos) ---
            try:
                # cancela agendamento anterior se existir
                if hasattr(self, "_purge_job") and self._purge_job:
                    try:
                        self.after_cancel(self._purge_job)
                    except Exception:
                        pass
                # só agenda se a janela ainda existir
                if self.winfo_exists():
                    self._purge_job = self.after(10 * 60 * 1000, self.check_purge_status)
            except Exception:
                pass

        except Exception as e:
            print("Erro em check_purge_status:", e)

    def _flash_purge_badge(self):
        """
        Efeito de piscar da badge — alterna visibilidade/leves cores.
        Guarda o id do after em self._purge_flash_job para permitir cancelamento seguro.
        """
        try:
            # se a janela foi destruída, aborta
            try:
                if not self.winfo_exists():
                    return
            except Exception:
                return

            # se a flag estiver desligada, garante estado final e cancela job
            if not getattr(self, "_purge_flashing", False):
                try:
                    self.purge_badge.config(bg="#e74c3c", fg="white")
                except Exception:
                    pass
                # cancela qualquer after pendente
                try:
                    if getattr(self, "_purge_flash_job", None):
                        self.after_cancel(self._purge_flash_job)
                except Exception:
                    pass
                self._purge_flash_job = None
                return

            # alterna entre dois estados visuais
            try:
                if self._purge_flash_state:
                    self.purge_badge.config(bg="#e74c3c", fg="white")
                else:
                    self.purge_badge.config(bg="#ff9999", fg="white")
            except Exception:
                pass

            self._purge_flash_state = not self._purge_flash_state

            # cancela qualquer after antigo (prudência)
            try:
                if getattr(self, "_purge_flash_job", None):
                    self.after_cancel(self._purge_flash_job)
            except Exception:
                pass
            self._purge_flash_job = None

            # agenda próximo ciclo e guarda id
            try:
                self._purge_flash_job = self.after(700, self._flash_purge_badge)
            except Exception:
                # não foi possível agendar (janela provavelmente fechada)
                self._purge_flash_job = None
                self._purge_flashing = False

        except Exception as e:
            print("Erro no flash da badge:", e)
            try:
                if getattr(self, "_purge_flash_job", None):
                    self.after_cancel(self._purge_flash_job)
            except Exception:
                pass
            self._purge_flash_job = None
            self._purge_flashing = False

    def _cancel_all_afters(self):
        """
        Cancela todos os timers agendados no Tk (útil quando ids foram perdidos).
        Coloque este método na classe (próximo a _cancel_after_job).
        """
        try:
            info = self.tk.call('after', 'info')  # retorna '' ou ids separados por espaço
            if not info:
                return
            jobs = str(info).split()
            for j in jobs:
                try:
                    self.after_cancel(j)
                except Exception:
                    pass
        except Exception:
            pass

    def mostrar_aviso_purga(self, force: bool = False):
        """
        Modal profissional e legível para o usuário sobre a limpeza mensal de registros.
        force=True  -> usuário clicou no sino -> sempre abre
        force=False -> automático -> abre somente se crítico
        """
        if "Calculo_Produto" not in self.permissoes:
            return
        
        try:
            hoje = date.today()
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            dias_restantes = ultimo_dia - hoje.day

            # Se não for forçado e não estiver em nível de aviso, não abre modal
            if not force and dias_restantes > self.purge_warning_days:
                return

            # resumo (consulta segura ao banco)
            resumo = self._get_purge_summary_safe()

            # se automático e não crítico, mostra caixa simples
            if not force and dias_restantes > self.purge_critical_days:
                texto_simples = (
                    f"Faltam {dias_restantes} dia(s) para a limpeza mensal (remoção automática) dos registros do Histórico de Cálculo."
                    "\n\nSe quiser manter seus registros, use a opção Exportar antes da data indicada."
                )
                messagebox.showinfo("Aviso: Limpeza Mensal de Registros", texto_simples)
                return

            # --- janela modal ---
            modal = tk.Toplevel(self)
            modal.title("Limpeza Mensal de Registros")
            modal.transient(self)
            modal.grab_set()
            modal.resizable(False, False)

            # garante que toda a área de conteúdo da janela seja branca
            modal.configure(bg="white")

            w, h = 640, 380
            x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (self.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")

            try:
                aplicar_icone(modal, self.caminho_icone)
            except Exception:
                pass

            # --- estilos locais (com prefixo exclusivo para este modal) ---
            style = ttk.Style(modal)
            prefix = f"Purge{modal.winfo_id()}"  # id único da janela

            style.configure(f"{prefix}.Header.TLabel",
                            font=("Helvetica", 14, "bold"))
            style.configure(f"{prefix}.BodyWhite.TLabel",
                            font=("Helvetica", 11),
                            background="white", foreground="black")
            style.configure(f"{prefix}.DaysWhite.TLabel",
                            font=("Helvetica", 28, "bold"),
                            background="white", foreground="#c0392b")

            # cabeçalho (mantém escuro para destacar)
            header = tk.Frame(modal, bg="#2c3e50", height=72)
            header.pack(side="top", fill="x")
            tk.Label(header, text="🗑️", bg="#2c3e50", fg="white",
                    font=("Helvetica", 18)).pack(side="left", padx=12, pady=10)
            tk.Label(header, text="Limpeza Mensal de Registros", bg="#2c3e50", fg="white",
                    font=("Helvetica", 16, "bold")).pack(side="left", pady=10)

            # conteúdo (fundo branco)
            content = tk.Frame(modal, bg="white", padx=12, pady=12)
            content.pack(fill="both", expand=True)

            # mensagem e dias
            top = tk.Frame(content, bg="white")
            top.pack(fill="x", pady=(0, 10))
            msg = (
                f"Faltam {dias_restantes} dia(s) para a limpeza mensal (remoção automática) dos registros\n"
                "do Histórico de Cálculo.\n\n"
                "O sistema removerá automaticamente esses registros na virada do mês para liberar espaço.\n"
                "Se quiser manter os dados, exporte-os agora."
            )
            ttk.Label(top, text=msg, style=f"{prefix}.BodyWhite.TLabel",
                    wraplength=420, justify="left").pack(side="left", anchor="nw")
            ttk.Label(top, text=str(dias_restantes),
                    style=f"{prefix}.DaysWhite.TLabel").pack(side="right", anchor="ne", padx=6)

            # progress bar
            pb_frame = tk.Frame(content, bg="white")
            pb_frame.pack(fill="x", pady=(0, 8))
            ttk.Label(pb_frame, text="Progresso do mês:",
                    style=f"{prefix}.BodyWhite.TLabel").pack(anchor="w")
            progress = ttk.Progressbar(
                pb_frame, orient="horizontal", mode="determinate", length=560)
            progress.pack(pady=4)
            dias_passados = hoje.day
            percent = int((dias_passados / ultimo_dia) * 100)
            progress['value'] = percent

            # resumo dos dados afetados
            summary_frame = tk.Frame(content, bg="white")
            summary_frame.pack(fill="both", expand=True)
            ttk.Label(summary_frame, text="Resumo dos registros afetados:",
                    style=f"{prefix}.BodyWhite.TLabel").pack(anchor="w", pady=(0, 4))

            if isinstance(resumo, dict):
                for k, v in resumo.items():
                    ttk.Label(summary_frame, text=f"• {k}: {v} registro(s)",
                            style=f"{prefix}.BodyWhite.TLabel").pack(anchor="w")
            else:
                ttk.Label(summary_frame, text=str(resumo),
                        style=f"{prefix}.BodyWhite.TLabel").pack(anchor="w")

            # linha de botões (fundo branco também)
            btn_frame = tk.Frame(modal, bg="white")
            btn_frame.pack(side="bottom", fill="x", padx=12, pady=12)

            def _export_action():
                # chama o método de exportação que agora existe na classe
                try:
                    # mantemos modal aberto enquanto a exportação ocorre
                    self.exportar_historico(parent=modal)
                except Exception as e:
                    messagebox.showerror("Erro na exportação", f"Ocorreu um erro ao exportar:\n{e}")

            # Botões com cores personalizadas (export verde, fechar cinza)
            b_export = tk.Button(
                btn_frame,
                text="Exportar (Excel)",
                command=_export_action,
                bg="#27ae60", fg="white",
                activebackground="#2ecc71",
                bd=0, relief="raised",
                padx=12, pady=6,
            )
            b_export.pack(side="right", padx=6)

            b_close = tk.Button(
                btn_frame,
                text="Fechar",
                command=modal.destroy,
                bg="#95a5a6", fg="white",
                activebackground="#b2babb",
                bd=0, relief="raised",
                padx=12, pady=6,
            )
            b_close.pack(side="left", padx=6)

            modal.focus_set()
            modal.wait_window()
            self._last_purge_days = dias_restantes

        except Exception as e:
            print("Erro em mostrar_aviso_purga:", e)

    def exportar_historico(self, parent=None):
        """
        Exporta os registros da tabela 'calculo_historico' para Excel (xlsx).
        - Remove coluna 'id' (case-insensitive) antes de exportar.
        - Normaliza datetimes (incl. tz-aware) para strings compatíveis com Excel.
        """
        try:
            if not getattr(self, "conn", None):
                messagebox.showerror("Exportar", "Conexão com o banco indisponível.")
                return

            cur = self.conn.cursor()
            try:
                cur.execute("SELECT * FROM calculo_historico;")
                rows = cur.fetchall()
                cols = [desc[0] for desc in cur.description] if cur.description else []
            except Exception as e_query:
                cur.close()
                messagebox.showerror("Exportar", f"Erro ao consultar dados: {e_query}")
                return
            cur.close()

            if not rows:
                messagebox.showinfo("Exportar", "Não há registros para exportar.")
                return

            # monta DataFrame
            try:
                df = pd.DataFrame(rows, columns=cols)
            except Exception:
                df = pd.DataFrame(rows)

            # --- Remover coluna(s) ID (case-insensitive) ---
            id_cols = [c for c in df.columns if str(c).lower() == "id"]
            if id_cols:
                df.drop(columns=id_cols, inplace=True)

            # --- Função de conversão segura para células datetime ---
            def _cell_to_excel_safe(x):
                if isinstance(x, datetime):
                    # converte tz-aware para UTC e torna tz-naive
                    if x.tzinfo is not None:
                        x = x.astimezone(timezone.utc).replace(tzinfo=None)
                    return x.strftime("%Y-%m-%d %H:%M:%S")
                return x

            # Aplica conversão por coluna:
            for c in df.columns:
                try:
                    # se a coluna for datetime (tz-aware ou não), faz conversão vetorizada
                    if ptypes.is_datetime64_any_dtype(df[c]) or ptypes.is_datetime64tz_dtype(df[c]):
                        try:
                            # se tz-aware, converte para UTC e remove timezone; então formata
                            if ptypes.is_datetime64tz_dtype(df[c]):
                                df[c] = df[c].dt.tz_convert("UTC").dt.tz_localize(None).dt.strftime("%Y-%m-%d %H:%M:%S")
                            else:
                                df[c] = df[c].dt.strftime("%Y-%m-%d %H:%M:%S")
                        except Exception:
                            # fallback para map célula-a-célula
                            df[c] = df[c].map(_cell_to_excel_safe)
                    else:
                        # para outras colunas, aplica map que deixa intacto se não for datetime
                        df[c] = df[c].map(_cell_to_excel_safe)
                except Exception:
                    # caso algo dê errado, transforma a coluna em string segura
                    df[c] = df[c].astype(str)

            # pede local para salvar
            filetypes = [("Excel (*.xlsx)", "*.xlsx"), ("CSV (*.csv)", "*.csv"), ("Todos os arquivos", "*.*")]
            initialfile = f"historico_calculo_{date.today().strftime('%Y%m%d')}.xlsx"
            path = filedialog.asksaveasfilename(parent=parent, title="Salvar histórico", defaultextension=".xlsx",
                                                filetypes=filetypes, initialfile=initialfile)
            if not path:
                return  # usuário cancelou

            # se usuário escolheu csv, salva direto
            if path.lower().endswith(".csv"):
                try:
                    df.to_csv(path, index=False, encoding="utf-8-sig")
                    messagebox.showinfo("Exportar", f"Exportação concluída com sucesso (CSV):\n{path}")
                except Exception as e_csv_direct:
                    messagebox.showerror("Exportar", f"Falha ao salvar CSV:\n{e_csv_direct}")
                return

            # tenta salvar em Excel; se falhar, salva CSV como fallback
            try:
                df.to_excel(path, index=False)
                messagebox.showinfo("Exportar", f"Exportação concluída com sucesso:\n{path}")
            except Exception as e_xlsx:
                try:
                    csv_path = path if path.lower().endswith(".csv") else (path + ".csv")
                    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
                    messagebox.showwarning("Exportar",
                                        f"Falha ao gerar .xlsx ({e_xlsx}). Salvado como CSV:\n{csv_path}")
                except Exception as e_csv:
                    messagebox.showerror("Exportar", f"Falha ao salvar arquivo:\nErro .xlsx: {e_xlsx}\nErro .csv: {e_csv}")

        except Exception as e_outer:
            messagebox.showerror("Exportar", f"Erro inesperado ao exportar:\n{e_outer}")

    def _create_tooltip(self, widget, text, delay=450, max_width=None):
        """
        Tooltip melhorado: quebra automática de linhas e ajuste para não sair da tela.
        - widget: widget alvo
        - text: texto do tooltip
        - delay: ms até exibir
        - max_width: largura máxima do tooltip em pixels (opcional)
        """
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            try:
                screen_w = widget.winfo_screenwidth()
                screen_h = widget.winfo_screenheight()
            except Exception:
                screen_w, screen_h = 1024, 768

            # determina wraplength: respeita max_width se passado, senão calcula baseado na tela
            if max_width:
                wrap_len = max(120, min(max_width, screen_w - 80))
            else:
                wrap_len = min(360, max(200, screen_w - 160))

            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)

            label = tk.Label(
                win,
                text=text,
                bg="#333333",
                fg="white",
                font=("Segoe UI", 9),
                bd=0,
                padx=6,
                pady=4,
                wraplength=wrap_len
            )
            label.pack()

            # posição inicial: centrado horizontalmente sobre o widget, abaixo do widget
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajustar horizontalmente para não sair da tela
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # ajustar verticalmente: tenta abaixo; se não couber, tenta acima; senão limita dentro da tela
            if y + h > screen_h:
                y_above = widget.winfo_rooty() - h - 6
                if y_above > 10:
                    y = y_above
                else:
                    y = max(10, screen_h - h - 10)

            win.geometry(f"+{x}+{y}")
            tooltip["win"] = win

        def hide():
            if tooltip["after_id"]:
                try:
                    widget.after_cancel(tooltip["after_id"])
                except Exception:
                    pass
                tooltip["after_id"] = None
            if tooltip["win"]:
                try:
                    tooltip["win"].destroy()
                except Exception:
                    pass
                tooltip["win"] = None

        def schedule_show(e=None):
            tooltip["after_id"] = widget.after(delay, show)

        widget.bind("<Enter>", schedule_show)
        widget.bind("<Leave>", lambda e: hide())
        widget.bind("<ButtonPress>", lambda e: hide())

    def _abrir_ajuda_modal(self, contexto=None):
        """Modal de Ajuda com layout profissional (navegação à esquerda e conteúdo à direita).
        As seções exibidas no Listbox são montadas dinamicamente conforme as permissões do usuário."""
        try:
            modal = tk.Toplevel(self)
            modal.title("Ajuda — Sistema Kametal")
            modal.transient(self)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões e centralização
            w, h = 920, 700
            x = max(0, (self.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (self.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(780, 520)

            # Tentar aplicar ícone (se a função existir)
            try:
                aplicar_icone(modal, getattr(self, "caminho_icone", None))
            except Exception:
                pass

            # ======= Cabeçalho =======
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=18)
            tk.Label(header, text="Guia rápido e descrição das telas (F1)", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=8)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # ======= Corpo: Navigation (esq) + Conteúdo (dir) =======
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            # Left: navegação
            nav_frame = tk.Frame(body, width=240, bg="#f5f7fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)

            nav_title = tk.Label(nav_frame, text="Seções", bg="#f5f7fa",
                                font=("Segoe UI", 10, "bold"))
            nav_title.pack(anchor="nw", pady=(8, 6), padx=12)

            # --- Montagem dinâmica de seções com base em permissões ---
            # Mapeamento: seção -> permissão necessária (None = visível para todos)
            sections_with_perms = [
                ("Visão Geral", None),
                ("Saída de NFs", "SistemaNF"),
                ("Base de Produtos", "criar_interface_produto"),
                ("Base de Materiais", "criar_interface_materiais"),
                ("Entrada de NFs", "Janela_InsercaoNF"),
                ("Cálculo de NFs", "Calculo_Produto"),
                ("Média Custo", "criar_media_custo"),
                ("Relatório Item por Grupo", "RelatorioApp"),
                ("Relatório Cotação", "CadastroProdutosApp"),
                ("Registro de Teste", "RegistroTeste"),
                ("Usuários", "criar_tela_usuarios"),
                ("Abas e Visões Rápidas", None),
                ("Fluxo / Estoque (impactos)", None),
                ("Notificações (Sino)", "Calculo_Produto"),  # mesmo permissão do cálculo
                ("Exemplos Práticos", None),
                ("FAQ", None),
            ]

            def _build_user_permission_set():
                """Tenta extrair as permissões do usuário de várias formas comuns.
                Retorna:
                    - None se devemos usar self.has_permission(permission) (quando existir)
                    - set(...) de permissões normalizadas (strings em lowercase) se encontradas
                    - empty set() se encontrado explicitamente um objeto vazio
                    - None (fallback) se não for possível detectar — nesse caso exibimos todas as seções.
                """
                # Se a classe tiver um método has_permission, usaremos ele (sinalizado por None)
                if hasattr(self, "has_permission") and callable(getattr(self, "has_permission")):
                    return None

                candidates = []
                try:
                    if hasattr(self, "user_permissions"):
                        candidates.append(getattr(self, "user_permissions"))
                    if hasattr(self, "permissoes"):
                        candidates.append(getattr(self, "permissoes"))
                    if hasattr(self, "user") and getattr(self, "user") is not None:
                        u = getattr(self, "user")
                        if hasattr(u, "permissions"):
                            candidates.append(getattr(u, "permissions"))
                        if hasattr(u, "permissoes"):
                            candidates.append(getattr(u, "permissoes"))
                except Exception:
                    # qualquer erro aqui => sem dados de permissão detectados
                    return None

                perms = set()
                found_any = False
                for c in candidates:
                    if c is None:
                        continue
                    found_any = True
                    # lista/tupla/set de strings ou objetos convertíveis
                    if isinstance(c, (list, tuple, set)):
                        for p in c:
                            try:
                                perms.add(str(p).lower())
                            except Exception:
                                pass
                    # string única com separador (ex.: "perm1,perm2")
                    elif isinstance(c, str):
                        for p in c.split(","):
                            p = p.strip()
                            if p:
                                perms.add(p.lower())
                    # dicionário possivelmente com chaves sendo permissões
                    elif isinstance(c, dict):
                        for k in c.keys():
                            try:
                                perms.add(str(k).lower())
                            except Exception:
                                pass
                    else:
                        # tentativa de iterar
                        try:
                            for p in list(c):
                                perms.add(str(p).lower())
                            found_any = True
                        except Exception:
                            pass

                if not found_any:
                    return None
                return perms

            _user_perms = _build_user_permission_set()

            def has_perm(required):
                """Verifica se o usuário tem a permissão `required`.
                - Se required is None => True
                - Se self.has_permission exists => usa ele
                - Se coletamos um set de permissões => verifica membership
                - Se não conseguimos detectar nada => retorna True (exibir todas)"""
                if required is None:
                    return True
                # Prioriza método self.has_permission(...)
                if hasattr(self, "has_permission") and callable(getattr(self, "has_permission")):
                    try:
                        return bool(self.has_permission(required))
                    except Exception:
                        # se falhar, não bloqueia — tenta usar set abaixo
                        pass
                # Usa set coletado
                if isinstance(_user_perms, set):
                    return required.lower() in _user_perms
                # Fallback: se não sabemos as permissões, mostrar tudo (comodidade)
                return True

            # Filtra as seções visíveis
            visible_sections = [name for (name, perm) in sections_with_perms if has_perm(perm)]
            if not visible_sections:
                # fallback mínimo (garante que o modal não fique vazio)
                visible_sections = ["Visão Geral"]

            # Cria o listbox com as seções visíveis
            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#fefefe")
            for s in visible_sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            # Right: área de conteúdo (Text rolável)
            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)

            txt = tk.Text(content_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb = tk.Scrollbar(content_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            # Conteúdo por seção (substitua texto se quiser)
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "A Janela Menu é a tela inicial do Sistema Kametal. Reúne acesso rápido aos módulos principais e mostra resumos com as últimas movimentações (entradas/saídas, produtos e pesos).\n\n"
                "Áreas principais: barra lateral de navegação, cabeçalho superior, área central com abas e rodapé."
            )

            contents["Saída de NFs"] = (
                "Saída de NFs\n\n"
                "Esta janela registra e exibe as notas fiscais de saída (clientes). Contém: número da NF, cliente, itens, peso por item e total. Ao confirmar uma saída, o sistema subtrai (debita) o peso/quantidade dos produtos do estoque."
            )

            contents["Base de Produtos"] = (
                "Base de Produtos\n\n"
                "Catálogo com todos os produtos da empresa: nome, código, unidade e referência às matérias-primas.\n"
                "Use para consultar detalhes do produto, histórico e ligação com custos."
            )

            contents["Base de Materiais"] = (
                "Base de Materiais\n\n"
                "Armazena as matérias-primas (insumos) usadas na fabricação. Informações chave: estoque, unidade e composição (por exemplo, % de cobre, % de zinco). Esses valores são usados no cálculo do custo e na composição dos produtos."
            )

            contents["Entrada de NFs"] = (
                "Entrada de NFs\n\n"
                "Tela para registrar as notas fiscais recebidas de fornecedores. Ao salvar uma NF de entrada, os produtos/quantidades/pesos são adicionados (somados) ao estoque conforme informado."
            )

            contents["Cálculo de NFs"] = (
                "Cálculo de NFs\n\n"
                "Ferramenta que calcula o custo associado a cada NF de entrada. Aplica regras como perdas, rendimentos e rateio entre itens, usando informações de Base de Materiais e Base de Produtos para distribuir custos."
            )

            contents["Média Custo"] = (
                "Média Custo\n\n"
                "Exibe a média de custo por produto — útil para análises e comparações ao longo do tempo.\n"
                "A média é calculada a partir das entradas e ajustes aplicados no período."
            )

            contents["Relatório Item por Grupo"] = (
                "Relatório Item por Grupo\n\n"
                "Gera relatórios mensais ou por período que mostram o que entrou e saiu por grupos/famílias de itens.\n"
                "Também calcula a média de custo dos produtos no período selecionado."
            )

            contents["Relatório Cotação"] = (
                "Relatório Cotação\n\n"
                "Área para registrar cotações e valores de referência (preço do produto, câmbio). A partir desses dados o sistema monta gráficos que ajudam a visualizar tendências e impactos no custo."
            )

            contents["Registro de Teste"] = (
                "Registro de Teste\n\n"
                "Módulo para armazenar histórico de testes e inspeções feitos nos produtos que saem — por exemplo, controle de qualidade, amostras, observações técnicas. Útil para rastreabilidade."
            )

            contents["Usuários"] = (
                "Usuários\n\n"
                "Tela de gestão de usuários: criar, editar, remover e atribuir permissões. Controle de acesso importante para evitar ações não autorizadas (ex.: excluir registros, alterar preços)."
            )

            contents["Abas e Visões Rápidas"] = (
                "Abas e Visões Rápidas\n\n"
                "As abas centrais (ex.: 'Últimas NFs') mostram uma visão rápida do sistema: tabelas com as últimas movimentações, filtros e seleção para acessar detalhes. Sirvem para monitoramento diário."
            )

            contents["Fluxo / Estoque (impactos)"] = (
                "Fluxo / Estoque — Impacto das ações\n\n"
                " - Entrada de NF: soma ao estoque.\n"
                " - Saída de NF: subtrai do estoque.\n"
                " - Cálculo de NFs: em geral analítico; verifique se existe confirmação que aplica ajustes ao estoque.\n\n"
                "Boas práticas: validar pesos antes de confirmar, reconciliar inventário e manter backups."
            )

            contents["Notificações (Sino)"] = (
                "Notificações (Sino)\n\n"
                "O botão do sino (🔔) aparece no canto superior direito do cabeçalho **apenas** para usuários com permissão 'Calculo_Produto'. Ele serve para alertar sobre a limpeza mensal do histórico de cálculos e outros eventos enviados pelo servidor via NOTIFY.\n\n"
                "Comportamento e significado:\n"
                " - A badge (pequeno círculo com número) sobre o sino indica o número de dias restantes até a 'Limpeza Mensal de Registros'.\n"
                " - Cor AMARELA — aviso: faltando alguns dias para a limpeza (por exemplo, 5 dias).\n"
                " - Cor VERMELHA + EFEITO DE PISCAR — nível crítico: faltando pouquíssimos dias para a limpeza (por exemplo, 2 dias).\n"
                " - Clicar no sino força a abertura do modal de aviso (mesmo que não esteja crítico) — é o atalho manual para `mostrar_aviso_purga(force=True)`.\n\n"
                "Integração com PostgreSQL / NOTIFY:\n"
                " - A aplicação realiza `LISTEN canal_atualizacao` (conexão separada) e processa payloads recebidos do servidor.\n"
                "Mensagens recebidas do servidor e o que elas fazem:\n"
                " - 'purge:<n>'  — atualiza o número mostrado na badge para <n> dias restantes (ex.: 'purge:3' fará a badge mostrar '3').\n"
                " - 'purge_warning' ou 'purge' — pede que o sistema verifique novamente se a limpeza está próxima e atualize a informação exibida ao usuário.\n"
                " - 'menu_atualizado' — pede que o sistema recarregue os menus e o conteúdo da tela para mostrar informações mais recentes.\n"
                " - Outras mensagens desconhecidas — o sistema tentará atualizar a tela/menu para manter tudo sincronizado; se algo não atualizar, peça para a equipe de backend investigar.\n\n"
                " - O sistema pode abrir automaticamente o modal de aviso quando estiver em nível crítico (por ex. faltando 2 dias ou menos). Isso evita que o usuário perca o prazo para exportar/recuperar dados.\n\n"
                "Problemas comuns / Soluções:\n"
                " - Sino/badge não aparece: verifique se o usuário tem a permissão 'Calculo_Produto'.\n"
                " - Badge não atualiza: confirme que o backend está chamando `NOTIFY canal_atualizacao, 'purge:<n>'` ou 'purge_warning' e que a conexão `conn_listen` não foi fechada.\n"
                " - Modal não abre automaticamente: cheque os valores de `purge_warning_days` e `purge_critical_days` e se `check_purge_status` está sendo agendado corretamente.\n\n"
                "Observação: o sino/badge é uma camada de alerta — os dados reais e as ações (exportar, excluir, ajustar) estão disponíveis no modal que é aberto quando você clica no sino."
            )

            contents["Exemplos Práticos"] = (
                "Exemplos Práticos (resumido)\n\n"
                "1) Registrar entrada: Entrada de NFs → Nova NF → preencher produtos e salvar → estoque aumenta.\n"
                "2) Registrar saída: Saída de NFs → Nova NF de saída → preencher e confirmar → estoque diminui.\n"
                "3) Ver média: Média Custo / Relatório Item por Grupo → selecionar período → gerar relatório."
            )

            contents["FAQ"] = (
                "FAQ — Perguntas Frequentes\n\n"
                "Q: Onde vejo % de cobre/zinco de um material?\n"
                "R: Em 'Base de Materiais' (cada material armazena sua composição).\n\n"
                "Q: O que fazer se o estoque não bater?\n"
                "R: Verificar lançamentos recentes, executar inventário e revisar logs; se necessário, chamar suporte."
            )

            # Função para mostrar conteúdo na área de texto
            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            # Inicializa com a primeira seção visível selecionada
            try:
                listbox.selection_set(0)
                mostrar_secao(visible_sections[0])
            except Exception:
                # se algo falhar, tenta mostrar Visão Geral
                if "Visão Geral" in contents:
                    mostrar_secao("Visão Geral")

            # Bind para mudar seção ao clicar
            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    idx = sel[0]
                    try:
                        key = visible_sections[idx]
                    except Exception:
                        return
                    mostrar_secao(key)

            listbox.bind("<<ListboxSelect>>", on_select)

            # ======= Rodapé com botão Fechar =======
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right")

            # Atalho Esc para fechar modal
            modal.bind("<Escape>", lambda e: modal.destroy())

            modal.focus_set()
            modal.wait_window()

        except Exception as e:
            print("Erro ao abrir modal de ajuda (profissional):", e)

    def _get_purge_summary_safe(self):
        """
        Tenta consultar o banco e retornar um dicionário com contagens (ex: {'Cálculos': 123}).
        Se a tabela não existir ou ocorrer erro, retorna uma mensagem amigável.
        """
        try:
            cur = self.conn.cursor()
            # Exemplo: contar registros na sua tabela (ajuste o nome da tabela/coluna conforme seu schema)
            # Se sua tabela tiver outro nome, altere "calculo_historico" abaixo.
            cur.execute("SELECT COUNT(*) FROM calculo_historico;")
            total = cur.fetchone()[0]
            # você pode adicionar mais queries aqui (p.ex. por usuário, por NF, etc.)
            return {"Registros no Histórico": total}
        except Exception:
            # fallback: não quebra a UI, apenas retorna texto explicativo
            return "Não foi possível obter detalhes do banco (tabela ausente ou erro)."

    def _cancel_after_job(self, attr_name: str):
        """Cancela com segurança um job armazenado em self.<attr_name> e zera o atributo."""
        try:
            job = getattr(self, attr_name, None)
            if job:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                try:
                    setattr(self, attr_name, None)
                except Exception:
                    pass
        except Exception:
            pass

    def destroy(self):
        """Cancela callbacks e sinaliza threads antes de destruir a janela."""
        try:
            self.encerrar = True
        except Exception:
            pass

        # cancela agendamento de check_purge_status
        try:
            if hasattr(self, "_purge_job") and self._purge_job:
                try:
                    self.after_cancel(self._purge_job)
                except Exception:
                    pass
                self._purge_job = None
        except Exception:
            pass

        # cancela outros agendamentos
        try:
            if getattr(self, "hora_job", None):
                try:
                    self.after_cancel(self.hora_job)
                except Exception:
                    pass
                self.hora_job = None
        except Exception:
            pass

        try:
            if getattr(self, "_spinner_job", None):
                try:
                    self.after_cancel(self._spinner_job)
                except Exception:
                    pass
                self._spinner_job = None
        except Exception:
            pass

        # cancela job do flash da badge (se houver)
        try:
            try:
                self._purge_flashing = False
            except Exception:
                pass
            if getattr(self, "_purge_flash_job", None):
                try:
                    self.after_cancel(self._purge_flash_job)
                except Exception:
                    pass
                self._purge_flash_job = None
        except Exception:
            pass

        # cancela qualquer after pendente conhecido ou perdido (segurança extra)
        try:
            self._cancel_all_afters()
        except Exception:
            pass

        try:
            if getattr(self, "conn_listen", None):
                try:
                    self.conn_listen.close()
                except Exception:
                    pass
                self.conn_listen = None
        except Exception:
            pass

        # destrói janela base
        try:
            super(type(self), self).destroy()
        except tk.TclError:
            pass
        except Exception:
            pass

    def _criar_main_content(self):
        """Cria a área de conteúdo principal."""
        self.main_content = tk.Frame(self, bg="#ecf0f1")
        self.main_content.pack(side="right", expand=True, fill="both")

    def _criar_menubar(self):
        """Cria o menu superior com opções de Logout e Sair e, se permitido, o menu de Usuários."""
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Logout", command=self.logout)
        file_menu.add_command(label="Sair", command=self.destroy)
        menubar.add_cascade(label="Arquivo", menu=file_menu)

        if "criar_tela_usuarios" in self.permissoes:
            user_menu = tk.Menu(menubar, tearoff=0)
            user_menu.add_command(label="Gerenciar Usuários", command=self.abrir_usuarios)
            menubar.add_cascade(label="Usuários", menu=user_menu)

    def _criar_menu_lateral(self):
        """Cria os botões da sidebar conforme as permissões do usuário."""
        self.buttons = []

        # 1) Frame de usuário
        user_frame = tk.Frame(self.sidebar_inner, bg="#1a252f", bd=2, relief="ridge")
        user_frame.pack(pady=(10, 0), padx=10, fill="x")
        user_label = tk.Label(
            user_frame,
            text=f"Bem-vindo, {self.user_name}",
            fg="white",
            bg="#1a252f",
            font=("Helvetica", 10, "bold")
        )
        user_label.pack(expand=True, fill="both", padx=5, pady=5)

        # 2) Separador
        separator = tk.Frame(self.sidebar_inner, bg="#95a5a6", height=2)
        separator.pack(pady=(10, 10), padx=10, fill="x")

        # 3) Mapeamento de rótulos
        janelas = {
            "criar_interface_produto": "Base de Produtos",
            "criar_interface_materiais": "Base de Materiais",
            "SistemaNF": "Saída de NFs",
            "Janela_InsercaoNF": "Entrada de NFs",
            "Calculo_Produto": "Cálculo de NFs",
            "criar_media_custo": "Média Custo",
            "criar_tela_usuarios": "Gerenciar Usuários",
            "RelatorioApp": "Relatório Item por Grupo",
            "CadastroProdutosApp": "Relatório Cotação",
            "RegistroTeste": "Registro de Teste"
        }

        # 4) Dados de botões
        buttons_data = [
            ("criar_interface_produto", self.abrir_base_produtos),
            ("criar_interface_materiais", self.abrir_base_materiais),
            ("SistemaNF", self.abrir_saida_nf),
            ("Janela_InsercaoNF", self.abrir_insercao_nf),
            ("Calculo_Produto", self.Calculo_produto),
            ("criar_media_custo", self.abrir_media_custo),
            ("RelatorioApp", self.relatorio_item_grupo),
            ("CadastroProdutosApp", self.cotacao),
            ("RegistroTeste", self.registro_teste)
        ]

        # 5) Criação dos botões visíveis
        for internal_name, command in buttons_data:
            if internal_name in self.permissoes:
                # Botão principal
                button = tk.Button(
                    self.sidebar_inner,
                    text=janelas.get(internal_name, internal_name),
                    fg="white", bg="#2980b9",
                    font=("Helvetica", 12),
                    bd=0, relief="flat",
                    padx=10, pady=10,
                    command=command
                )
                button.bind("<Enter>", lambda e, b=button: b.config(bg="#3498db"))
                button.bind("<Leave>", lambda e, b=button: b.config(bg="#2980b9"))
                button.pack(pady=5, padx=10, fill="x")

                # Sombra
                shadow = tk.Label(self.sidebar_inner, bg="#2471a3")
                shadow.pack_forget()  # desativado no layout novo, opcional

                self.buttons.append((button, internal_name, shadow))

        # 6) --- Rodapé com versão (barra inferior da janela inteira) ---
        # Use self.root (ou self.master) para ligar ao topo da janela principal
        top = self.sidebar_inner.winfo_toplevel()

        versao_label = tk.Label(
            top,
            text=f"Versão {__version__}",
            bg="#1b2e3f",
            fg="lightgray",
            font=("Helvetica", 9)
        )
        # canto inferior esquerdo (10px de margem)
        versao_label.place(relx=0.0, rely=1.0, x=10, y=-10, anchor="sw")

    def _criar_abas(self):
        """Cria as abas do Notebook conforme self.permissoes e popula cada uma.
        Garante ao menos uma aba final (mensagem) se o usuário não tiver permissão para nada."""
        # --- cancela spinner/placeholder criado na inicialização para evitar que o "Carregando" fique sozinho ---
        try:
            if getattr(self, "_spinner_job", None):
                try:
                    self.after_cancel(self._spinner_job)
                except Exception:
                    pass
                self._spinner_job = None
        except Exception:
            pass

        # Se existir frame_placeholder (criado por _criar_abas_minimal), tente removê-lo
        try:
            if getattr(self, "frame_placeholder", None):
                try:
                    self.notebook.forget(self.frame_placeholder)
                except Exception:
                    pass
                # tentar destruir widgets internos se existirem
                try:
                    if getattr(self, "spinner_label", None):
                        try:
                            self.spinner_label.destroy()
                        except Exception:
                            pass
                        self.spinner_label = None
                except Exception:
                    pass
                self.frame_placeholder = None
        except Exception:
            pass

        # limpa abas existentes (segurança)
        for tab in self.notebook.tabs():
            try:
                self.notebook.forget(tab)
            except Exception:
                pass

        # --- cria abas condicionais como antes ---
        added = 0
        if "Janela_InsercaoNF" in self.permissoes or "SistemaNF" in self.permissoes:
            self.frame_nf = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_nf, text="Últimas NFs")
            try:
                self.criar_relatorio_nf()
            except Exception:
                pass
            added += 1

        if "criar_interface_materiais" in self.permissoes or "criar_interface_produto" in self.permissoes:
            self.frame_relatorios_produto_material = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_relatorios_produto_material, text="Relatórios Produto/Material")
            try:
                self.criar_relatorio_produto_material()
            except Exception:
                pass
            added += 1

        if "criar_media_custo" in self.permissoes:
            self.frame_relatorio = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_relatorio, text="Relatório de Estoque")
            try:
                self.criar_relatorio_estoque(self.conn)
            except Exception:
                pass
            self.frame_graficos = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_graficos, text="Gráficos de Custo")
            try:
                self.criar_grafico_mensal("Custo Médio de Cada Produto", [])
            except Exception:
                pass
            added += 1

        if "CadastroProdutosApp" in self.permissoes:
            self.frame_grafico_produtos = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_grafico_produtos, text="Gráfico Produtos")
            try:
                self.criar_grafico_produtos()
            except Exception:
                pass
            self.frame_grafico_dolar = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_grafico_dolar, text="Gráfico Dólar")
            try:
                self.criar_grafico_dolar()
            except Exception:
                pass
            added += 1

        if "RegistroTeste" in self.permissoes:
            self.frame_ultimos_registros_teste = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_ultimos_registros_teste, text="Últimos Registros de Teste")
            try:
                self.criar_aba_ultimos_registros_teste()
            except Exception:
                pass
            added += 1

        # Se nenhuma aba foi adicionada, inserimos uma aba com mensagem (mantém a área do notebook igual)
        if added == 0:
            frame_none = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(frame_none, text="Visão Rápida")
            tk.Label(
                frame_none,
                text="Você não tem permissão para visualizar os relatórios/abas desta área.",
                bg="#ecf0f1",
                font=("Helvetica", 12),
                wraplength=700,
                justify="center"
            ).pack(expand=True, fill="both", pady=40)

        # Atualiza a barra de abas (botões) para refletir o novo conteúdo
        try:
            self._atualizar_barra_abas()
        except Exception:
            pass

    def _criar_abas_minimal(self):
        self.frame_placeholder = tk.Frame(self.notebook, bg="#ecf0f1")
        self.notebook.add(self.frame_placeholder, text="Carregando...")

        self.spinner_label = tk.Label(
            self.frame_placeholder,
            text="⠋ Carregando...",
            font=("Segoe UI", 14),
            bg="#ecf0f1"
        )
        self.spinner_label.pack(pady=50)

        self._spinner_frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self._spinner_index = 0
        self._atualizar_spinner()

    def _atualizar_spinner(self):
        try:
            if not getattr(self, "spinner_label", None):
                return
            try:
                if not self.spinner_label.winfo_exists():
                    return
            except tk.TclError:
                return

            self.spinner_label.config(
                text=f"{self._spinner_frames[self._spinner_index]} Carregando..."
            )
            self._spinner_index = (self._spinner_index + 1) % len(self._spinner_frames)
            # agendar apenas se a janela ainda existir
            try:
                self._spinner_job = self.after(100, self._atualizar_spinner)
            except tk.TclError:
                # janela/destruição ocorreu no meio -> ignora
                self._spinner_job = None
                return
        except tk.TclError:
            return

    def configurar_estilos(self):
        """Configura os estilos do ttk para Notebook e Treeview."""
        style = ttk.Style(self)
        style.theme_use("alt")
        style.configure("Menu.TNotebook", background="#ecf0f1", borderwidth=0)
        style.configure("Menu.TNotebook.Tab",
                        font=("Arial", 15),
                        padding=[10, 5],
                        background="#bdc3c7",
                        foreground="black",
                        borderwidth=2,
                        relief="raised")
        style.map("Menu.TNotebook.Tab",
                  background=[("selected", "#ecf0f1")],
                  foreground=[("selected", "black")],
                  relief=[("selected", "ridge")])
        style.configure("Treeview",
                        background="white",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="white",
                        font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        style.map("Treeview",
                  background=[("selected", "#2980b9")],
                  foreground=[("selected", "white")])

        # === esconder as tabs nativas do Notebook ===
        # Alguns temas/versões requerem esconder tanto o layout do Notebook quanto o layout da Tab.
        try:
            style.layout("Hidden.TNotebook.Tab", [])  # remove layout das tabs
            style.layout("Hidden.TNotebook", [("Notebook.client", {"sticky": "nswe"})])  # mantém apenas a "client area"
        except Exception:
            # se falhar por causa do tema, relaxa — não é crítico, mas normalmente funciona
            pass

    def _criar_barra_abas(self):
        """Cria a barra de abas com setas e canvas rolável."""
        self.tabbar = tk.Frame(self.main_content, bg="#bdc3c7", height=36)
        self.tabbar.pack(side="top", fill="x")

        self.tab_left = tk.Button(self.tabbar, text="◀", bd=0, width=3,
                                  command=lambda: self._rolar_abas(-200))
        self.tab_left.pack(side="left", padx=2, pady=2)

        self.tab_right = tk.Button(self.tabbar, text="▶", bd=0, width=3,
                                   command=lambda: self._rolar_abas(200))
        self.tab_right.pack(side="right", padx=2, pady=2)

        self.tab_canvas = tk.Canvas(self.tabbar, bg="#bdc3c7", height=36, highlightthickness=0)
        self.tab_canvas.pack(side="left", fill="x", expand=True)

        self.tab_buttons_frame = tk.Frame(self.tab_canvas, bg="#bdc3c7")
        self.tab_canvas.create_window((0, 0), window=self.tab_buttons_frame, anchor="nw")

        # Atualiza região rolável quando os botões mudam
        def _on_cfg(e):
            self.tab_canvas.configure(scrollregion=self.tab_canvas.bbox("all"))
        self.tab_buttons_frame.bind("<Configure>", _on_cfg)

        # rolagem horizontal com Shift+Scroll (Windows). opcional.
        def _on_mousewheel_h(event):
            delta = int(-1 * (event.delta / 120))
            self.tab_canvas.xview_scroll(delta, "units")
        self.tab_canvas.bind_all("<Shift-MouseWheel>", _on_mousewheel_h)

    def _atualizar_barra_abas(self):
        """Reconstrói os botões das abas conforme self.notebook.tabs()."""
        # limpa
        for w in self.tab_buttons_frame.winfo_children():
            w.destroy()

        # recria botões para cada aba (mantemos a ordem do notebook)
        for tab_id in self.notebook.tabs():
            text = self.notebook.tab(tab_id, option="text")
            btn = tk.Button(
                self.tab_buttons_frame,
                text=text,
                font=("Arial", 12),
                bd=0,
                padx=10,
                pady=6,
                relief="flat",
                command=lambda t=tab_id: self._selecionar_e_destacar_aba(t)
            )
            btn.pack(side="left", padx=2, pady=2)

        self.update_idletasks()
        self._destacar_aba_selecionada()

    def _selecionar_e_destacar_aba(self, tab_id):
        """Seleciona a aba do notebook e garante que seu botão esteja visível/realçado."""
        self.notebook.select(tab_id)
        self._destacar_aba_selecionada()

        # garante visibilidade do botão correspondente
        btn = None
        selected_text = self.notebook.tab(tab_id, "text")
        for child in self.tab_buttons_frame.winfo_children():
            if getattr(child, "cget", lambda x: "")("text") == selected_text:
                btn = child
                break
        if not btn:
            return

        # coordenadas do botão
        x1 = btn.winfo_x()
        x2 = x1 + btn.winfo_width()
        canvas_w = self.tab_canvas.winfo_width()
        view_left = self.tab_canvas.canvasx(0)

        total_width = max(1, self.tab_buttons_frame.winfo_width())
        if x1 < view_left:
            self.tab_canvas.xview_moveto(x1 / total_width)
        elif x2 > view_left + canvas_w:
            self.tab_canvas.xview_moveto((x2 - canvas_w) / total_width)

    def _destacar_aba_selecionada(self):
        """Aplica visual diferente ao botão da aba selecionada."""
        selected = self.notebook.select()
        sel_text = self.notebook.tab(selected, "text") if selected else None
        for child in self.tab_buttons_frame.winfo_children():
            try:
                if child.cget("text") == sel_text:
                    child.config(bg="#ecf0f1", relief="raised")
                else:
                    child.config(bg="#bdc3c7", relief="flat")
            except Exception:
                pass

    def _rolar_abas(self, delta):
        """Rola horizontalmente o canvas das abas. delta em pixels aproximados."""
        units = int(delta / 10)
        self.tab_canvas.xview_scroll(units, "units")

    def carregar_nome_usuario(self, user_id):
        """Busca o nome do usuário no banco de dados usando o user_id."""
        if not self.conn:
            return ""
        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT nome FROM usuarios WHERE id = %s", (user_id,))
            resultado = cursor.fetchone()
            cursor.close()
            if resultado:
                return resultado[0]
            return ""
        except Exception as e:
            print("Erro ao buscar nome do usuário:", e)
            return ""

    def _escutar_notificacoes(self):
        """
        Thread que fica escutando NOTIFY do Postgres (self.conn_listen).
        - Processa notificações imediatamente dentro do loop, evitando UnboundLocalError.
        - Quando receber payloads relacionados à purga, força atualização do badge via self.after().
        - Agenda atualizar_pagina() na thread principal quando apropriado.
        """
        try:
            # loop principal da thread de escuta
            while not getattr(self, "encerrar", False) and not getattr(self, "_stop_event", threading.Event()).is_set():
                try:
                    # espera até 1 segundo por atividade (ajustável)
                    try:
                        ready = select.select([self.conn_listen], [], [], 1)
                    except Exception:
                        # se select falhar (por exemplo self.conn_listen inválido), pausa e continua
                        time.sleep(1)
                        continue

                    # faz poll da conexão para popular conn_listen.notifies
                    try:
                        self.conn_listen.poll()
                    except Exception:
                        # possível que a conexão tenha sido fechada — apenas ignore e continue
                        time.sleep(1)
                        continue

                    # processa todas as notificações pendentes
                    while getattr(self.conn_listen, "notifies", None):
                        notify = None
                        try:
                            notify = self.conn_listen.notifies.pop(0)
                        except Exception:
                            # nenhum notify real -> sai do while
                            break

                        # DEBUG: print para acompanhar no console
                        try:
                            print(f"[NOTIFY] canal={getattr(notify, 'channel', '')} payload={getattr(notify, 'payload', '')}")
                        except Exception:
                            pass

                        # processa payload do notify AQUI (onde 'notify' existe)
                        try:
                            payload = getattr(notify, "payload", "") or ""
                            # formato esperado: "purge:3" ou "purge_warning" ou "menu_atualizado"
                            if payload.startswith("purge:"):
                                # extrai número de dias (se possível) e força atualização imediata
                                try:
                                    dias = int(payload.split(":", 1)[1])
                                except Exception:
                                    dias = None
                                # atualiza badge/estado na thread principal
                                try:
                                    if dias is not None and getattr(self, "purge_badge", None):
                                        self.after(50, lambda d=dias: self.purge_badge.config(text=str(d)))
                                    # recalcula corretamente e ajusta visual
                                    self.after(100, self.check_purge_status)
                                except Exception:
                                    pass

                            elif payload in ("purge_warning", "purge"):
                                # apenas rechecagem da purga
                                try:
                                    self.after(100, self.check_purge_status)
                                except Exception:
                                    pass

                            elif payload == "menu_atualizado":
                                # força atualização completa da página (abas, sidebar, quick views)
                                try:
                                    self.after(100, self.atualizar_pagina)
                                except Exception:
                                    pass

                            else:
                                # outros payloads genéricos -> atualizar página
                                try:
                                    existe = False
                                    try:
                                        existe = self.winfo_exists()
                                    except Exception:
                                        existe = False

                                    if existe:
                                        self.after(150, self.atualizar_pagina)
                                except Exception:
                                    pass

                        except Exception as e_notify:
                            # falha ao processar um notify específico — loga e continua
                            print("Erro ao processar notify:", e_notify)

                except Exception as e_loop:
                    # erro transitório no loop de select/poll
                    if not getattr(self, "encerrar", False) and not getattr(self, "_stop_event", threading.Event()).is_set():
                        print(f"[Erro na thread de escuta]: {e_loop}")
                    # pequena pausa antes de tentar novamente
                    time.sleep(1)

        except Exception as e_outer:
            # se a thread morrer por erro não esperado, loga
            print(f"[Erro crítico na thread de escuta]: {e_outer}")

    def atualizar_pagina(self):
        try:
            # 1) Recarrega permissões
            self.permissoes = self.carregar_permissoes_usuario(self.user_id)

            # 2) Atualiza abas do Notebook sem destruir frames
            # => lista de (nome_interno, frame_objeto, título)
            abas = [
                ("Janela_InsercaoNF", self.frame_nf, "Últimas NFs"),
                ("SistemaNF",        self.frame_nf, "Últimas NFs"),           # mesma aba
                ("criar_interface_produto",   self.frame_relatorios_produto_material, "Relatórios Produto/Material"),
                ("criar_interface_materiais", self.frame_relatorios_produto_material, "Relatórios Produto/Material"),
                ("criar_media_custo",         self.frame_relatorio, "Relatório de Estoque"),
                ("criar_media_custo",         self.frame_graficos, "Gráficos de Custo"),
                ("CadastroProdutosApp",       self.frame_grafico_produtos, "Gráfico Produtos"),
                ("CadastroProdutosApp",       self.frame_grafico_dolar, "Gráfico Dólar"),
                ("RegistroTeste",             self.frame_ultimos_registros_teste, "Registro de Teste"),
            ]

            for nome, frame, titulo in abas:
                tem = nome in self.permissoes
                # verifica se o frame já está no notebook
                try:
                    idx = self.notebook.index(frame)
                    existe = True
                except tk.TclError:
                    existe = False

                if tem and not existe:
                    # adiciona pela primeira vez
                    self.notebook.add(frame, text=titulo)
                elif not tem and existe:
                    # oculta sem destruir
                    self.notebook.tab(frame, state="hidden")
                elif tem:
                    # garante que esteja visível
                    self.notebook.tab(frame, state="normal")

            # 3) Atualiza botões da sidebar sem destruir
            # => em self.buttons armazenamos tuplas (btn, nome_interno, shadow)
            for btn, nome_int, shadow in self.buttons:
                if nome_int in self.permissoes:
                    # Se o widget foi gerenciado por pack originalmente, reaplicamos pack()
                    try:
                        mgr = btn.winfo_manager()
                    except Exception:
                        mgr = None

                    if mgr == "pack":
                        if not btn.winfo_ismapped():
                            btn.pack(pady=5, padx=10, fill="x")
                        # sombra (se existir e também for pack)
                        if shadow:
                            try:
                                if shadow.winfo_manager() == "pack" and not shadow.winfo_ismapped():
                                    shadow.pack(padx=10)
                            except Exception:
                                pass
                    # Se foi gerenciado por place, garantimos que _default_y exista
                    elif mgr == "place":
                        # força cálculo de layout para ter coordenadas válidas
                        self.update_idletasks()
                        if not hasattr(btn, "_default_y"):
                            try:
                                btn._default_y = btn.winfo_y()
                            except Exception:
                                btn._default_y = 0
                        if shadow and not hasattr(shadow, "_default_y"):
                            try:
                                shadow._default_y = shadow.winfo_y()
                            except Exception:
                                shadow._default_y = 0
                        try:
                            btn.place(x=10, y=btn._default_y, width=200)
                        except Exception:
                            pass
                        if shadow:
                            try:
                                shadow.place(x=12, y=shadow._default_y, width=200, height=40)
                            except Exception:
                                pass
                    else:
                        # fallback: tenta pack como solução resiliente
                        try:
                            if not btn.winfo_ismapped():
                                btn.pack(pady=5, padx=10, fill="x")
                        except Exception:
                            pass
                else:
                    # ocultar independentemente do gerenciador usado
                    try:
                        mgr = btn.winfo_manager()
                    except Exception:
                        mgr = None

                    if mgr == "pack":
                        try:
                            if btn.winfo_ismapped():
                                btn.pack_forget()
                        except Exception:
                            pass
                    elif mgr == "place":
                        try:
                            btn.place_forget()
                        except Exception:
                            pass
                    else:
                        # tentativa genérica (pode falhar para alguns widgets, por isso o try)
                        try:
                            btn.forget()
                        except Exception:
                            pass

                    # esconder a sombra também
                    if shadow:
                        try:
                            mgr2 = shadow.winfo_manager()
                        except Exception:
                            mgr2 = None
                        if mgr2 == "pack":
                            try:
                                if shadow.winfo_ismapped():
                                    shadow.pack_forget()
                            except Exception:
                                pass
                        elif mgr2 == "place":
                            try:
                                shadow.place_forget()
                            except Exception:
                                pass
                        else:
                            try:
                                shadow.forget()
                            except Exception:
                                pass

            # 4) Atualiza conteúdo interno dos frames, caso necessário
    # --------------------------------------------------------

            # 4.1) Últimas NFs
            if "Janela_InsercaoNF" in self.permissoes or "SistemaNF" in self.permissoes:
                # limpa tudo dentro do frame_nf
                for w in self.frame_nf.winfo_children():
                    w.destroy()
                # repopula
                self.criar_relatorio_nf()

            # 4.2) Relatórios Produto/Material
            if "criar_interface_produto" in self.permissoes or "criar_interface_materiais" in self.permissoes:
                for w in self.frame_relatorios_produto_material.winfo_children():
                    w.destroy()
                self.criar_relatorio_produto_material()

            # 4.3) Relatório de Estoque
            if "criar_media_custo" in self.permissoes:
                for w in self.frame_relatorio.winfo_children():
                    w.destroy()
                self.criar_relatorio_estoque(self.conn)

            # 4.4) Gráficos de Custo
            if "criar_media_custo" in self.permissoes:
                for w in self.frame_graficos.winfo_children():
                    w.destroy()
                # NÂO chame _buscar_dados_mensal()
                self.criar_grafico_mensal("Custo Médio de Cada Produto")

            # 4.5) Gráfico de Produtos
            if "CadastroProdutosApp" in self.permissoes:
                for w in self.frame_grafico_produtos.winfo_children():
                    w.destroy()
                self.criar_grafico_produtos()

            # 4.6) Gráfico de Dólar
            if "CadastroProdutosApp" in self.permissoes:
                for w in self.frame_grafico_dolar.winfo_children():
                    w.destroy()
                self.criar_grafico_dolar()

            # 4.7) Últimos Registros de Teste
            if "RegistroTeste" in self.permissoes:
                # só tenta atualizar se o frame existir
                if hasattr(self, "frame_ultimos_registros_teste"):
                    for w in self.frame_ultimos_registros_teste.winfo_children():
                        w.destroy()
                    # repopula a aba
                    self.criar_aba_ultimos_registros_teste()

            # 5) Força o Tkinter a processar redraw
            self.update_idletasks()

        except Exception as e:
            print(f"[Erro ao atualizar a página]: {e}")
        finally:
            self.atualizacao_pendente = False

    def criar_relatorio_produto_material(self, tipo=None):
        """
        Cria relatórios de produtos e/ou materiais em uma única aba, conforme permissões.
        Se 'tipo' for especificado, cria apenas aquele. Caso contrário, usa permissões para decidir.
        """

        tipos = []
        if tipo:
            tipos.append(tipo)
        else:
            if "criar_interface_produto" in self.permissoes:
                tipos.append("produtos")
            if "criar_interface_materiais" in self.permissoes:
                tipos.append("materiais")

        # Container principal
        container = tk.Frame(self.frame_relatorios_produto_material, bg="#ecf0f1")
        container.pack(padx=10, pady=10, fill="both", expand=True)

        row = 0
        for tipo in tipos:
            if tipo == 'produtos':
                query = 'SELECT nome, "percentual_cobre", "percentual_zinco" FROM produtos'
                cols = [
                    (0, "Nome", 250, lambda v: v),
                    (1, "Percentual Cobre", 150, lambda v: f"{v}%"),
                    (2, "Percentual Zinco", 150, lambda v: f"{v}%"),
                ]
                sort_key = lambda row: row[0].lower()
                titulo = "Relatório de Produtos"
            elif tipo == 'materiais':
                query = "SELECT id, nome, fornecedor, valor, grupo FROM materiais"
                cols = [
                    (1, "Nome", 250, lambda v: v),
                    (2, "Fornecedor", 120, lambda v: v),
                    (3, "Valor", 80, lambda v: str(v).replace('.', ',')),
                    (4, "Grupo", 120, lambda v: v),
                ]
                sort_key = lambda row: row[1].lower()
                titulo = "Relatório de Materiais"
            else:
                continue

            conn = conectar()
            if conn is None:
                print("Erro: não foi possível conectar ao banco de dados.")
                return
            try:
                cursor = conn.cursor()
                cursor.execute(query)
                registros = cursor.fetchall()
            except Exception as e:
                print("Erro ao executar a consulta:", e)
                return
            finally:
                cursor.close()
                conn.close()

            registros.sort(key=sort_key)

            # --- LabelFrame externo para enquadrar tudo (título + tabela) ---
            quadro = tk.LabelFrame(container, text=titulo, font=("Arial", 12, "bold"),bg="#ecf0f1", fg="#000000", bd=2, relief="groove", labelanchor="nw")
            quadro.grid(row=row, column=0, padx=10, pady=10, sticky="nsew")
            container.grid_rowconfigure(row, weight=1)
            container.grid_columnconfigure(0, weight=1)

            # Frame da tabela (com fundo branco)
            frame_tabela = tk.Frame(quadro, bg="#ffffff", bd=1, relief="solid")
            frame_tabela.pack(padx=10, pady=(0, 10), fill="both", expand=True)

            # Treeview estilizado
            tree = ttk.Treeview(frame_tabela,
                                columns=[col[1] for col in cols],
                                show="headings",
                                height=12,
                                style="Treeview")

            for orig_idx, title, width, _ in cols:
                tree.column(title, anchor="center", width=width)
                tree.heading(title, text=title, anchor="center")

            for i, row_data in enumerate(registros):
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                valores = [fmt(row_data[orig_idx]) for orig_idx, _, _, fmt in cols]
                tree.insert("", "end", values=valores, tags=(tag,))

            tree.tag_configure("evenrow", background="#f2f2f2")
            tree.tag_configure("oddrow", background="white")

            scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            row += 1

    def criar_relatorio_nf(self):
        """
        Cria um relatório na aba 'Últimas NFs' mostrando:
        - Entradas de NF (caso o usuário tenha a permissão 'Janela_InsercaoNF')
        - Saídas de NF (caso o usuário tenha a permissão 'SistemaNF')
        """
        if not self.conn:
            print("Erro: Conexão com o banco de dados indisponível.")
            return

        # Função auxiliar para formatar datas
        def formatar_data(data_valor):
            try:
                if hasattr(data_valor, "strftime"):
                    return data_valor.strftime("%d/%m/%Y")
                else:
                    return datetime.strptime(str(data_valor), "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                return data_valor

        # Função auxiliar para formatar Peso (Peso Líquido ou Peso Integral)
        def formatar_valor(valor):
            try:
                # usa separador de milhar e vírgula decimal, sempre 3 casas
                s = f"{float(valor):,.3f}"  
                # converte para padrão BR: ponto de milhar, vírgula decimal
                s = s.replace(",", "X").replace(".", ",").replace("X", ".")
                return s + " Kg"
            except Exception:
                return valor

        # --- Se o usuário tem permissão para ver Entradas (Inserção NF) ---
        if "Janela_InsercaoNF" in self.permissoes:
            entradas_frame = tk.LabelFrame(
                self.frame_nf, text="Últimas Entradas de NF",
                bg="#ecf0f1", font=("Arial", 12, "bold")
            )
            entradas_frame.pack(fill="both", padx=10, pady=10, expand=True)
            
            # CONSULTA DAS ÚLTIMAS ENTRADAS (somar_produtos)
            try:
                cursor = self.conn.cursor()
                cursor.execute("""
                    SELECT data, nf, fornecedor, produto, peso_liquido 
                    FROM somar_produtos 
                    WHERE data = (SELECT MAX(data) FROM somar_produtos)
                    ORDER BY data DESC, nf::INTEGER DESC;
                """)
                entradas = cursor.fetchall()
                cursor.close()
            except Exception as e:
                print("Erro ao buscar entradas:", e)
                entradas = []

            # Cria um frame para o treeview das entradas e sua scrollbar
            frame_treeview_entradas = tk.Frame(entradas_frame)
            frame_treeview_entradas.pack(fill="both", padx=5, pady=5)

            colunas_entrada = ("Data", "NF", "Fornecedor", "Produto", "Peso Líquido")
            tree_entrada = ttk.Treeview(frame_treeview_entradas, columns=colunas_entrada, show="headings", height=7)
            tree_entrada.heading("Data", text="Data")
            tree_entrada.column("Data", width=80, anchor="center")
            tree_entrada.heading("NF", text="NF")
            tree_entrada.column("NF", width=80, anchor="center")
            tree_entrada.heading("Fornecedor", text="Fornecedor")
            tree_entrada.column("Fornecedor", width=150, anchor="center")
            tree_entrada.heading("Produto", text="Produto")
            tree_entrada.column("Produto", width=250, anchor="center")
            tree_entrada.heading("Peso Líquido", text="Peso Líquido")
            tree_entrada.column("Peso Líquido", width=150, anchor="center")

            for row in entradas:
                data_formatada = formatar_data(row[0])
                peso_formatado = formatar_valor(row[4]) if row[4] is not None else "N/A"
                tree_entrada.insert("", "end", values=(data_formatada, row[1], row[2], row[3], peso_formatado))

            # Cria e vincula a scrollbar para o treeview de entradas
            scrollbar_entradas = tk.Scrollbar(frame_treeview_entradas, orient="vertical", command=tree_entrada.yview)
            tree_entrada.configure(yscrollcommand=scrollbar_entradas.set)
            tree_entrada.pack(side="left", fill="both", expand=True)
            scrollbar_entradas.pack(side="right", fill="y")

        # --- Se o usuário tem permissão para ver Saídas (Saída NF) ---
        if "SistemaNF" in self.permissoes:
            saidas_frame = tk.LabelFrame(
                self.frame_nf, text="Últimas Saídas de NF",
                bg="#ecf0f1", font=("Arial", 12, "bold")
            )
            saidas_frame.pack(fill="both", padx=10, pady=10, expand=True)

            # CONSULTA DAS ÚLTIMAS SAÍDAS (nf)
            try:
                cursor = self.conn.cursor()
                cursor.execute("""
                    SELECT id, data, numero_nf, cliente 
                    FROM nf 
                    WHERE CAST(data AS date) = (
                        SELECT CAST(MAX(data) AS date) FROM nf
                    )
                    ORDER BY numero_nf::INTEGER DESC;
                """)
                saidas = cursor.fetchall()
                cursor.close()
            except Exception as e:
                print("Erro ao buscar saídas:", e)
                saidas = []

            # Cria um frame para o treeview das NFs (cabeçalho) e sua scrollbar
            frame_treeview_nf = tk.Frame(saidas_frame)
            frame_treeview_nf.pack(fill="both", padx=5, pady=5)
            
            colunas_saida = ("Data", "NF", "Cliente")
            tree_saida_header = ttk.Treeview(frame_treeview_nf, columns=colunas_saida, show="headings", height=5)
            tree_saida_header.heading("Data", text="Data")
            tree_saida_header.column("Data", width=80, anchor="center")
            tree_saida_header.heading("NF", text="NF")
            tree_saida_header.column("NF", width=80, anchor="center")
            tree_saida_header.heading("Cliente", text="Cliente")
            tree_saida_header.column("Cliente", width=300, anchor="center")
            
            # Insere os registros, usando o id da NF como iid
            for row in saidas:
                nf_id = row[0]
                data_formatada = formatar_data(row[1])
                tree_saida_header.insert("", "end", iid=nf_id, values=(data_formatada, row[2], row[3]))
            
            # Cria e vincula a scrollbar para o treeview de NFs
            scrollbar_nf = tk.Scrollbar(frame_treeview_nf, orient="vertical", command=tree_saida_header.yview)
            tree_saida_header.configure(yscrollcommand=scrollbar_nf.set)
            tree_saida_header.pack(side="left", fill="both", expand=True)
            scrollbar_nf.pack(side="right", fill="y")

            # Cria um frame e Treeview para os produtos da NF selecionada com sua própria scrollbar
            produtos_frame = tk.LabelFrame(
                saidas_frame, text="Produtos da NF Selecionada",
                bg="#ecf0f1", font=("Arial", 10, "bold")
            )
            produtos_frame.pack(fill="both", padx=5, pady=5, expand=True)
            
            # Adiciona a coluna "Base Produto" ao lado de "Produto" e "Peso"
            colunas_produtos = ("Produto", "Peso", "Base Produto")
            tree_produtos = ttk.Treeview(produtos_frame, columns=colunas_produtos, show="headings", height=5)
            tree_produtos.heading("Produto", text="Produto")
            tree_produtos.column("Produto", width=350, anchor="center")
            tree_produtos.heading("Peso", text="Peso")
            tree_produtos.column("Peso", width=50, anchor="center")
            tree_produtos.heading("Base Produto", text="Base Produto")
            tree_produtos.column("Base Produto", width=120, anchor="center")
            
            scrollbar_produtos = tk.Scrollbar(produtos_frame, orient="vertical", command=tree_produtos.yview)
            tree_produtos.configure(yscrollcommand=scrollbar_produtos.set)
            tree_produtos.pack(side="left", fill="both", expand=True)
            scrollbar_produtos.pack(side="right", fill="y")

            # Função para atualizar a lista de produtos quando uma NF for selecionada
            def atualizar_produtos(event):
                for item in tree_produtos.get_children():
                    tree_produtos.delete(item)
                selecionados = tree_saida_header.selection()
                if not selecionados:
                    return
                nf_id = selecionados[0]
                try:
                    cursor = self.conn.cursor()
                    cursor.execute("""
                        SELECT produto_nome, peso, base_produto
                        FROM produtos_nf 
                        WHERE nf_id = %s;
                    """, (nf_id,))
                    produtos = cursor.fetchall()
                    cursor.close()
                except Exception as e:
                    print("Erro ao buscar produtos para a NF selecionada:", e)
                    produtos = []
                for prod in produtos:
                    produto_nome = prod[0]
                    peso_formatada = formatar_valor(prod[1])
                    base_produto = prod[2] if prod[2] is not None else ""
                    tree_produtos.insert("", "end", values=(produto_nome, peso_formatada, base_produto))

            tree_saida_header.bind("<<TreeviewSelect>>", atualizar_produtos)

            # Seleciona o primeiro item se houver registros para mostrar os produtos inicialmente
            if saidas:
                primeiro_item = tree_saida_header.get_children()[0]
                tree_saida_header.selection_set(primeiro_item)
                atualizar_produtos(None)

    def criar_relatorio_estoque(self, conn=None):
        # tenta usar a versão em classe (MediaCusto). Se não existir, tenta as funções antigas.
        try:
            from media_custo import MediaCusto as _MediaCustoClass

            def buscar_produtos(conn):
                inst = object.__new__(_MediaCustoClass)
                inst.conn = conn
                return _MediaCustoClass.buscar_produtos(inst)

            def buscar_estoque(conn, nome_produto):
                inst = object.__new__(_MediaCustoClass)
                inst.conn = conn
                return _MediaCustoClass.buscar_estoque(inst, nome_produto)

        except Exception:
            from media_custo import buscar_produtos, buscar_estoque

        # Se não foi passada conexão, cria uma nova
        if conn is None:
            conn = conectar()
        if conn is None:
            print("Erro: Não foi possível conectar ao banco de dados.")
            return

        # Limpa widgets antigos da aba
        for widget in self.frame_relatorio.winfo_children():
            widget.destroy()

        # Busca produtos e quantidades
        produtos_bd = buscar_produtos(conn)
        produtos = []
        for produto in produtos_bd:
            nome_produto = produto[0]
            try:
                qtd_estoque = buscar_estoque(conn, nome_produto)
            except Exception as e:
                # segurança: se a função esperar assinatura diferente, tenta compatibilizar
                try:
                    qtd_estoque = buscar_estoque(conn, nome_produto)  # repetir por clareza; mantém contrato
                except Exception:
                    qtd_estoque = "0,000 kg"
                    print(f"Erro ao obter estoque para {nome_produto}: {e}")
            produtos.append((nome_produto, qtd_estoque))

        produtos.sort(key=lambda produto: produto[0].lower())

        # Moldura com fundo branco e título em negrito
        moldura = tk.LabelFrame(
            self.frame_relatorio,
            text="Relatório de Estoque",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            fg="#000000",
            bd=2,
            relief="groove",
            padx=10,
            pady=10
        )
        moldura.pack(padx=20, pady=20, fill="both", expand=True)

        frame_tabela = tk.Frame(moldura, bg="#FFFFFF")
        frame_tabela.pack(fill="both", expand=True)

        colunas = ("Produto", "Qtd Estoque")
        tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=15)

        tree.column("Produto", anchor="center", width=250)
        tree.column("Qtd Estoque", anchor="center", width=120)
        tree.heading("Produto", text="Produto", anchor="center")
        tree.heading("Qtd Estoque", text="Qtd Estoque", anchor="center")

        for i, produto in enumerate(produtos):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            tree.insert("", "end", values=produto, tags=(tag,))

        tree.tag_configure("evenrow", background="#f2f2f2")
        tree.tag_configure("oddrow", background="white")

        scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def criar_grafico_mensal(self, titulo, dados=None):
        """Cria um gráfico de barras com os dados de custo médio de cada produto somados com o custo empresa,
        e o adiciona à aba de gráficos."""
        sns.set(style="whitegrid")
        fig = Figure(figsize=(10, 6), dpi=100)
        ax = fig.add_subplot(111)

        try:
            conn = conectar()
            if conn is None:
                return
            with conn.cursor() as cursor:
                # Busca os nomes dos produtos
                cursor.execute("SELECT nome FROM produtos;")
                produtos = cursor.fetchall()
                categorias = [produto[0] for produto in produtos]
                medias_ponderadas = []
                for nome_produto in categorias:
                    cursor.execute(
                        """
                        SELECT SUM(eq.quantidade_estoque * eq.custo_total), 
                            SUM(eq.quantidade_estoque),
                            MAX(sp.custo_empresa)
                        FROM estoque_quantidade eq
                        JOIN somar_produtos sp ON eq.id_produto = sp.id
                        WHERE sp.produto = %s
                        """,
                        (nome_produto,)
                    )
                    resultado = cursor.fetchone()
                    if resultado is None:
                        medias_ponderadas.append(0)
                    else:
                        soma_custo_ponderado, soma_quantidade, custo_empresa = resultado
                        if soma_quantidade and soma_custo_ponderado:
                            media = soma_custo_ponderado / soma_quantidade
                            if custo_empresa is not None:
                                media += custo_empresa
                            medias_ponderadas.append(float(media))
                        else:
                            medias_ponderadas.append(0)
            conn.close()
        except Exception as e:
            print(f"Erro ao buscar dados do banco: {e}")
            return

        # Filtra os produtos com média maior que 0
        produtos_filtrados = [(c, m) for c, m in zip(categorias, medias_ponderadas) if m > 0]
        if not produtos_filtrados:
            print("Nenhum produto com custo superior a 0 para exibir no gráfico.")
            return

        # Ordena os produtos de forma decrescente: maior custo médio à esquerda e menor à direita
        produtos_filtrados.sort(key=lambda x: x[1], reverse=True)

        categorias_filtradas, medias_filtradas = zip(*produtos_filtrados)

        # Cria o gráfico de barras
        colors = sns.color_palette("Blues_r", len(medias_filtradas))
        bars = ax.bar(categorias_filtradas, medias_filtradas, color=colors, edgecolor='black', width=0.6)
        max_value = max(max(medias_filtradas) * 1.2, 100)
        ax.set_ylim(0, max_value)
        ax.set_title(titulo, fontsize=16, pad=20, fontweight='bold')
        ax.set_ylabel('Média de Custo Total', fontsize=12, fontweight='bold')
        ax.set_xticks(np.arange(len(categorias_filtradas)))
        ax.set_xticklabels(categorias_filtradas, rotation=30, ha='right', fontsize=10, fontweight='bold')
        fig.subplots_adjust(left=0.6, right=0.95, bottom=0.2, top=0.9)
        ax.yaxis.grid(True, linestyle='--', linewidth=0.7, alpha=0.7)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        # Adiciona as anotações com os valores em cada barra
        for bar, valor in zip(bars, medias_filtradas):
            valor_formatado = f'{valor:,.2f}'
            valor_formatado = "R$" + valor_formatado.replace(",", "v").replace(".", ",").replace("v", ".")
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + max_value * 0.02,
                valor_formatado,
                ha='center',
                va='bottom',
                fontsize=10,
                fontweight='bold',
                color='black'
            )

        # Configura o mplcursors para mostrar o nome do produto ao passar o mouse
        cursor_obj = mplcursors.cursor(bars, hover=True)
        cursor_obj.connect("add", lambda sel: sel.annotation.set_text(categorias_filtradas[sel.index]))
        fig.canvas.mpl_connect("button_press_event", lambda event: [cursor_obj.remove_selection(s) for s in cursor_obj.selections])
        fig.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.frame_graficos)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill='both', expand=True)
        canvas.draw()

        return canvas_widget

    def criar_grafico_produtos(self):
        """Cria um gráfico com as cotações de vários produtos ao longo do tempo e exibe no frame_grafico_produtos."""
        # Inicializa conexões e tooltip
        # Não precisamos inicializar tooltip aqui pois usamos mplcursors

        # Obter dados
        conn = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15
                FROM cotacao_produtos
                ORDER BY to_date(split_part(periodo, ' á ', 1), 'DD/MM/YY') DESC
            """
            )
            registros = cursor.fetchall()
            if not registros:
                messagebox.showwarning("Aviso", "Não há dados para exibir no gráfico.")
                return
            df = pd.DataFrame(
                registros,
                columns=["periodo","cobre","zinco","liga_62_38","liga_65_35","liga_70_30","liga_85_15"]
            )
            # Separar datas inicial e final
            df[['data_inicio', 'data_fim']] = df['periodo'].str.split(' á ', expand=True)
            df['data_inicio'] = pd.to_datetime(df['data_inicio'], format='%d/%m/%y')
            df['data_fim'] = pd.to_datetime(df['data_fim'], format='%d/%m/%y')

            # Agora criar uma lista para armazenar as novas linhas
            novas_linhas = []

            for _, row in df.iterrows():
                # Primeira linha - mês do início
                nova_linha_inicio = row.copy()
                nova_linha_inicio['periodo'] = row['data_inicio']
                novas_linhas.append(nova_linha_inicio)
                
                # Se o mês de início e fim forem diferentes, cria outra linha
                if row['data_inicio'].month != row['data_fim'].month:
                    nova_linha_fim = row.copy()
                    nova_linha_fim['periodo'] = row['data_fim']
                    novas_linhas.append(nova_linha_fim)

            # Criar novo DataFrame com todas as linhas
            df = pd.DataFrame(novas_linhas)
            df = df[df["periodo"].dt.year == datetime.now().year]
            df['mes'] = df["periodo"].dt.to_period("M")
            medias = df.groupby('mes').mean()
            datas = medias.index.to_timestamp()
        except Exception as e:
            messagebox.showerror("Erro BD", f"Falha ao buscar dados: {e}")
            return
        finally:
            cursor.close()
            conn.close()

        # Criar figura e eixo
        fig, ax = plt.subplots(figsize=(16, 4))
        fig.subplots_adjust(top=0.85, left=0.12, right=0.95, bottom=0.20)

        # Plotar linhas e armazenar referências
        self.linhas_produtos = []
        produtos = [
            ('Cobre', medias['cobre'], '#b87333'),
            ('Zinco', medias['zinco'], '#a8a8a8'),
            ('Liga 62/38', medias['liga_62_38'], '#0072B2'),
            ('Liga 65/35', medias['liga_65_35'], '#FF0000'),
            ('Liga 70/30', medias['liga_70_30'], '#009E73'),
            ('Liga 85/15', medias['liga_85_15'], '#CC79A7')
        ]
        for nome, serie, cor in produtos:
            linha, = ax.plot(datas, serie, label=nome, color=cor, marker='o')
            self.linhas_produtos.append((linha, nome))

        meses_pt = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(
            FuncFormatter(lambda x, pos: meses_pt[mdates.num2date(x).month - 1])
        )

        # Rótulos dos eixos
        ax.set_xlabel('Mês', labelpad=10, fontweight='bold')
        ax.set_ylabel('Preço', labelpad=10, fontweight='bold')

        # Legenda e título
        handles, labels = ax.get_legend_handles_labels()
        fig.legend(
            handles, labels,
            loc='upper center', bbox_to_anchor=(0.5, 0.98),
            ncol=len(labels), frameon=False
        )
        fig.suptitle(
            f"Gráfico de Produtos - {datetime.now().year}",
            y=0.92,
            fontweight='bold',
            fontsize=14  # Opcional: aumenta o tamanho da fonte
        )

        # Grade
        ax.grid(True, which='major', linestyle='--', linewidth=0.5)

        lines = [linha for linha, nome in self.linhas_produtos]
        names = [nome for linha, nome in self.linhas_produtos]

        cursor_obj = mplcursors.cursor(lines, hover=True)
        
        def _on_add(sel):
            name = names[lines.index(sel.artist)]
            x, y = sel.target
            date_obj = mdates.num2date(x)
            mes_num = date_obj.month
            year = date_obj.year
            # Mês abreviado em português
            meses_pt_abrev = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            mes_txt = meses_pt_abrev[mes_num - 1]
            valor = f"{y:.2f}".replace('.', ',')  # <-- aqui troca o ponto pela vírgula
            texto = (
                f"Nome: {name}\n"
                f"Mês: {mes_txt}\n"
                f"Ano: {year}\n"
                f"Valor: R$ {valor}"
            )

            sel.annotation.set_text(texto)
        
        cursor_obj.connect("add", _on_add)

        # Limpar seleção ao clicar fora
        # Criar canvas antes de conectar eventos
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico_produtos)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        canvas.mpl_connect(
            "button_press_event",
            lambda event: [cursor_obj.remove_selection(s) for s in list(cursor_obj.selections)]
        )

        # Salvar referência do canvas, se precisar
        self.canvas_produtos = canvas
        self.eixo_produtos = ax

    def criar_grafico_dolar(self):
        """Cria um gráfico da cotação do dólar ao longo do tempo no frame_grafico_dolar."""
        # Obter dados
        conn = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT data, dolar
                FROM cotacao_dolar
                ORDER BY data ASC
            """)
            registros = cursor.fetchall()
            if not registros:
                messagebox.showwarning("Aviso", "Não há dados para exibir no gráfico.")
                return
            
            df = pd.DataFrame(registros, columns=["data", "dolar"])
            df["data"] = pd.to_datetime(df["data"])
            df = df[df["data"].dt.year == datetime.now().year]
            df['mes'] = df["data"].dt.to_period("M")
            medias = df.groupby('mes').mean()
            datas = medias.index.to_timestamp()
        except Exception as e:
            messagebox.showerror("Erro BD", f"Falha ao buscar dados: {e}")
            return
        finally:
            cursor.close()
            conn.close()

        # Criar figura e eixo
        fig, ax = plt.subplots(figsize=(16, 4))
        fig.subplots_adjust(top=0.85, left=0.18, right=0.95, bottom=0.20)

        # Plotar
        linha, = ax.plot(datas, medias['dolar'], label="Dólar", color="#4682B4", marker='o')

        meses_pt = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: meses_pt[mdates.num2date(x).month-1]))

        # Trocar ponto por vírgula nas legendas do eixo Y
        ax.set_xlabel('Mês', labelpad=10, fontweight='bold')
        ax.set_ylabel('Cotação (R$)', labelpad=10, fontweight='bold')

        # Ajustar a legenda para subir mais
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.25), ncol=1, frameon=False)  # Aumentar o valor de 1.15 para 1.25

        fig.suptitle(
            f"Gráfico de Dólar - {datetime.now().year}", 
            y=0.92,
            fontweight='bold',
            fontsize=14  # Opcional: aumenta o tamanho da fonte
        )
        ax.grid(True, which='major', linestyle='--', linewidth=0.5)

        # >>>>>>>>>>>> CRIAÇÃO DO CANVAS <<<<<<<<<<<<
        self.canvas_dolar = FigureCanvasTkAgg(fig, master=self.frame_grafico_dolar)  # frame_grafico_dolar agora
        self.canvas_dolar.draw()
        self.canvas_dolar.get_tk_widget().pack(side="right", fill="both", expand=True)

        # Tooltip com valor com vírgula
        cursor_obj = mplcursors.cursor([linha], hover=True)

        def _on_add(sel):
            x, y = sel.target
            date_obj = mdates.num2date(x)
            valor = f"{y:.2f}".replace('.', ',')  # Agora com 2 casas decimais e troca ponto por vírgula
            texto = (
                f"Mês: {meses_pt[date_obj.month-1]}\n"
                f"Cotação: R$ {valor}\n"
                f"Ano: {date_obj.year}"
            )
            sel.annotation.set_text(texto)

        cursor_obj.connect("add", _on_add)

        self.canvas_dolar.mpl_connect(
            "button_press_event",
            lambda event: [cursor_obj.remove_selection(s) for s in list(cursor_obj.selections)]
        )

        # Alterar valores do eixo Y para vírgula
        def formatar_eixo_y(x, pos):
            return f"R$ {x:,.2f}".replace('.', ',').replace(',', 'X').replace('.', ',').replace('X', ',')

        ax.yaxis.set_major_formatter(FuncFormatter(formatar_eixo_y))
            
    def criar_aba_ultimos_registros_teste(self):
        """
        Cria uma aba/section 'Últimos Registros' exibindo todos os registros cuja
        data == MAX(data) (última data presente na tabela).
        """
        if not getattr(self, "conn", None):
            print("Erro: Conexão com o banco de dados indisponível.")
            return

        # formata data para dd/mm/YYYY
        def formatar_data(data_valor):
            try:
                if hasattr(data_valor, "strftime"):
                    return data_valor.strftime("%d/%m/%Y")
                else:
                    return datetime.datetime.strptime(str(data_valor), "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                return data_valor

        # formata area/alongamento com separador BR (opcional)
        def formatar_numero(valor, casas=4):
            try:
                s = f"{float(valor):,.{casas}f}"
                s = s.replace(",", "X").replace(".", ",").replace("X", ".")
                return s
            except Exception:
                return valor

        # limpa frame antigo (se existir)
        try:
            for w in self.frame_ultimos_registros_teste.winfo_children():
                w.destroy()
        except Exception:
            pass

        frame = self.frame_ultimos_registros_teste

        lblframe = tk.LabelFrame(frame, text="Últimos Registros", bg="#ecf0f1", font=("Arial", 12, "bold"))
        lblframe.pack(fill="both", padx=10, pady=10, expand=True)

        # frame do treeview
        frame_treeview = tk.Frame(lblframe)
        frame_treeview.pack(fill="both", padx=5, pady=5, expand=True)

        colunas = ("Data", "Código Barras", "OP", "Cliente", "Material", "Liga",
                "Dimensões", "Área", "LR Tração (N)", "LR Tração (MPa)",
                "Alongamento (%)", "Tempera", "Máquina", "Empresa")

        # guarda o treeview no self para poder atualizar de fora depois
        self.tree_registros_teste = ttk.Treeview(frame_treeview, columns=colunas, show="headings", height=10)

        # cabeçalhos / larguras (ajuste conforme desejar)
        larguras = {
            "Data": 90, "Código Barras": 150, "OP": 80, "Cliente": 200, "Material": 140,"Liga": 100, "Dimensões": 140, "Área": 90, "LR Tração (N)": 120,"LR Tração (MPa)": 150, "Alongamento (%)": 150, "Tempera": 100, "Máquina": 150, "Empresa": 150
        }

        # centraliza cabeçalho e conteúdo das colunas
        for c in colunas:
            self.tree_registros_teste.heading(c, text=c, anchor="center")
            self.tree_registros_teste.column(c, width=larguras.get(c, 100), anchor="center", stretch=True)

        # scrollbars (vertical + horizontal)
        vsb = ttk.Scrollbar(frame_treeview, orient="vertical", command=self.tree_registros_teste.yview)
        hsb = ttk.Scrollbar(frame_treeview, orient="horizontal", command=self.tree_registros_teste.xview)
        self.tree_registros_teste.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # layout com grid para que os scrollbars fiquem certinhos
        self.tree_registros_teste.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, columnspan=2, sticky="ew")

        # configura pesos para expansão correta
        frame_treeview.grid_rowconfigure(0, weight=1)
        frame_treeview.grid_columnconfigure(0, weight=1)

        # consulta e preenchimento
        try:
            cur = self.conn.cursor()
            cur.execute("""
                SELECT data, codigo_barras, op, cliente, material, liga,
                    dimensoes, area, lr_tracao_n, lr_tracao_mpa,
                    alongamento_percentual, tempera, maquina, empresa
                FROM registro_teste
                WHERE data = (SELECT MAX(data) FROM registro_teste)
                ORDER BY id;
            """)
            rows = cur.fetchall()
            cur.close()
        except Exception as e:
            print("Erro ao buscar últimos registros:", e)
            rows = []

        if not rows:
            # se não houver registros, mostra uma linha indicando isso
            self.tree_registros_teste.insert("", "end", values=("Nenhum registro encontrado",) + ("",) * (len(colunas)-1))
        else:
            for r in rows:
                data_fmt = formatar_data(r[0])
                area_fmt = formatar_numero(r[7], casas=4) if r[7] is not None else ""
                # ---- Alteração: mostrar apenas inteiro + '%' (ex: 45% ou 5%)
                if r[10] is not None and str(r[10]).strip() != "":
                    try:
                        # converte robustamente (troca vírgula por ponto se necessário)
                        raw = str(r[10]).replace(",", ".")
                        inteiro = int(float(raw) + 0.5)  # arredondamento half-up
                        along_fmt = f"{inteiro}%"
                    except Exception:
                        along_fmt = ""
                else:
                    along_fmt = ""
                # ---- fim alteração

                values = (
                    data_fmt,
                    r[1] or "",
                    r[2] or "",
                    r[3] or "",
                    r[4] or "",
                    r[5] or "",
                    r[6] or "",
                    area_fmt,
                    r[8] or "",
                    r[9] or "",
                    along_fmt,
                    r[11] or "",
                    r[12] or "",
                    r[13] or ""
                )
                self.tree_registros_teste.insert("", "end", values=values)

        # opcional: atualizar barra de abas customizada
        try:
            atualizar_barra = getattr(self, "_atualizar_barra_abas", None)
            if atualizar_barra:
                atualizar_barra()
        except Exception:
            pass

    def _open_async(self, module_name: str, attr_name: str = None,
                is_func: bool = False,
                init_args: tuple = None, init_kwargs: dict = None,
                hide_menu_after_open: bool = True):
        """
        Importa module_name (e pega attr_name) em thread background.
        Quando pronto, na thread principal instancia/chama o atributo.
        """
        init_args = tuple(init_args or ())
        init_kwargs = dict(init_kwargs or {})

        q = queue.Queue()

        def worker():
            try:
                mod = importlib.import_module(module_name)
                target = getattr(mod, attr_name) if attr_name else mod
                q.put(("ok", target))
            except Exception:
                q.put(("err", traceback.format_exc()))

        # opcional: mudar cursor para busy
        try:
            self.config(cursor="watch")
        except Exception:
            pass

        threading.Thread(target=worker, daemon=True).start()

        def check():
            try:
                status, payload = q.get_nowait()
            except queue.Empty:
                self.after(80, check)
                return

            # reset cursor
            try:
                self.config(cursor="")
            except Exception:
                pass

            if status == "err":
                messagebox.showerror(
                    "Erro ao abrir",
                    f"Falha ao carregar {module_name}:\n\n{payload}"
                )
                return

            target = payload
            try:
                if is_func:
                    target(*init_args, **init_kwargs)
                else:
                    try:
                        if len(init_args) == 0:
                            target(self)
                        else:
                            target(*init_args, **init_kwargs)
                    except TypeError:
                        target(self, *init_args, **init_kwargs)

                if hide_menu_after_open:
                    try:
                        self.withdraw()
                    except Exception:
                        pass
            except Exception:
                messagebox.showerror(
                    "Erro ao abrir",
                    f"Falha ao instanciar {attr_name or module_name}:\n\n{traceback.format_exc()}"
                )
                return

        self.after(80, check)

    def prewarm_modules(self, modules: list):
        """
        Importa em background uma lista de módulos para reduzir latência no primeiro clique.
        Passar lista de nomes de módulos, ex: ['Base_produto', 'Saida_NF', ...]
        """
        def w():
            for m in modules:
                try:
                    importlib.import_module(m)
                except Exception:
                    # falhas no prewarm não devem travar a UI
                    pass
        threading.Thread(target=w, daemon=True).start()

    def abrir_base_produtos(self):
        self._open_async("Base_produto", "InterfaceProduto")

    def abrir_base_materiais(self):
        self._open_async("Base_material", "InterfaceMateriais")

    def abrir_saida_nf(self):
        self._open_async("Saida_NF", "SistemaNF")

    def abrir_insercao_nf(self):
        # este módulo aguardava (self, self)
        self._open_async("Insercao_NF", "Janela_InsercaoNF", init_args=(self, self))

    def abrir_usuarios(self):
        self._open_async("usuario", "InterfaceUsuarios")

    def Calculo_produto(self):
         self._open_async("Estoque", "CalculoProduto", init_kwargs={"janela_menu": self})

    def abrir_media_custo(self):
        # função que recebe main_window=self
        self._open_async("media_custo", "criar_media_custo", is_func=True, init_kwargs={"main_window": self})

    def relatorio_item_grupo(self):
        self._open_async("relatorio_saida", "RelatorioApp")

    def cotacao(self):
        self._open_async("relatorio_cotacao", "CadastroProdutosApp")

    def registro_teste(self):
        # exemplo com kwargs específicos
        self._open_async("registro_teste", "RegistroTeste", init_args=(), init_kwargs={"janela_menu": self, "master": self})

    def logout(self):
        """Logout seguro: sinaliza encerrar, cancela afters, fecha conexões e destrói a janela."""
        from login import TelaLogin

        # evita reentrância
        if getattr(self, "_closing", False):
            return
        self._closing = True

        # remove temporariamente o handler WM_DELETE_WINDOW para evitar que o
        # gerenciador de janelas chame on_closing durante o teardown
        try:
            self.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass

        # garante que o piscar pare e cancela job do flash
        try:
            try:
                self._purge_flashing = False
            except Exception:
                pass
            if getattr(self, '_purge_flash_job', None):
                try:
                    self.after_cancel(self._purge_flash_job)
                except Exception:
                    pass
                self._purge_flash_job = None
        except Exception:
            pass

        # cleanup geral
        try:
            self._cleanup()
        except Exception:
            pass

        # cancela quaisquer afters restantes (segurança extra)
        try:
            self._cancel_all_afters()
        except Exception:
            pass

        # destruir janela com segurança
        try:
            try:
                self.update_idletasks()
            except Exception:
                pass
            super(type(self), self).destroy()
        except tk.TclError:
            pass
        except Exception:
            pass

        # abrir login; se falhar, mostra no console
        try:
            TelaLogin().run()
        except Exception as e:
            print("[Erro ao abrir TelaLogin]:", e)

    def carregar_janela(self, title, content):
        """Carrega conteúdo na área principal e restaura a janela do menu."""
        for widget in self.main_content.winfo_children():
            widget.destroy()
        title_label = tk.Label(self.main_content, text=title, bg="#ecf0f1", font=("Helvetica", 20, "bold"))
        title_label.pack(pady=30)
        content_label = tk.Label(self.main_content, text=content, bg="#ecf0f1", font=("Helvetica", 14))
        content_label.pack(pady=20)
        self.state("zoomed")
        self.deiconify()

    def _cleanup(self, wait_threads_timeout=1.0):
        """Limpa callbacks, sinaliza threads e fecha conexões — idempotente."""
        if getattr(self, "_cleanup_done", False):
            return
        self._cleanup_done = True

        # sinaliza para threads pararem
        try:
            self._stop_event.set()
        except Exception:
            pass
        try:
            self.encerrar = True
        except Exception:
            pass

        # cancelar after jobs conhecidos
        try:
            for job_attr in ("hora_job", "_spinner_job", "_purge_flash_job"):
                job = getattr(self, job_attr, None)
                if job:
                    try:
                        self.after_cancel(job)
                    except Exception:
                        pass
                    try:
                        setattr(self, job_attr, None)
                    except Exception:
                        pass
        except Exception:
            pass

        # cancela quaisquer afters restantes (segurança extra)
        try:
            self._cancel_all_afters()
        except Exception:
            pass

        # fechar conexão de listen (se existir)
        try:
            conn = getattr(self, "conn_listen", None)
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
                try:
                    self.conn_listen = None
                except Exception:
                    pass
        except Exception:
            pass

        # juntar thread de escuta (se existir)
        try:
            th = getattr(self, "_thread_escuta", None)
            if th and th.is_alive():
                try:
                    th.join(timeout=wait_threads_timeout)
                except Exception:
                    pass
        except Exception:
            pass

    def on_closing(self):
        # evita reentrância
        if getattr(self, "_closing", False):
            return
        self._closing = True

        # desliga o handler de WM_DELETE_WINDOW para evitar que callbacks
        # sejam chamados novamente durante o destroy
        try:
            self.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass

        if not messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            self._closing = False
            # restaura handler padrão
            try:
                self.protocol("WM_DELETE_WINDOW", self.on_closing)
            except Exception:
                pass
            return

        # garante que o piscar pare e cancela job do flash antes do cleanup
        try:
            try:
                self._purge_flashing = False
            except Exception:
                pass
            if getattr(self, "_purge_flash_job", None):
                try:
                    self.after_cancel(self._purge_flash_job)
                except Exception:
                    pass
                self._purge_flash_job = None
        except Exception:
            pass

        try:
            self._cleanup()
        except Exception:
            pass

        # cancela quaisquer afters restantes (segurança extra)
        try:
            self._cancel_all_afters()
        except Exception:
            pass

        # destruir a janela com segurança
        try:
            # garante redraw/flush antes de destruir
            try:
                self.update_idletasks()
            except Exception:
                pass
            super(type(self), self).destroy()
        except tk.TclError:
            pass
        except Exception:
            pass