import tkinter as tk
from tkinter import ttk, messagebox
from logos import aplicar_icone
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
from conexao_db import conectar, logger
import threading
import select
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT
import numpy as np
import seaborn as sns
import mplcursors
from datetime import datetime
import pandas as pd
import importlib
import queue
import traceback

class Janela_Menu(tk.Tk):
    def __init__(self, user_id):
        super().__init__()
        self.user_id = user_id

        # Conexão com o banco de dados e Permissão
        # … seu código atual …
        self.conn = conectar()
        # Conexão separada para escutar notificações
        self.conn_listen = conectar()  # nova conexão para o LISTEN
        self.conn_listen.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        self._cur_listen = self.conn_listen.cursor()
        self._cur_listen.execute("LISTEN canal_atualizacao;")

        # Inicia a thread de escuta
        self.encerrar = False  # <- define antes de iniciar escuta
        self._stop_event = threading.Event()   # sinal para threads pararem
        self._closing = False                 # evita reentrância em logout/on_closing
        self.hora_job = None                   # garante que exista o atributo desde o início
        self._spinner_job = None               # idem para spinner
        self.atualizacao_pendente = False
        self._thread_escuta = threading.Thread(target=self._escutar_notificacoes, daemon=True)
        self._thread_escuta.start()

        self.permissoes = self.carregar_permissoes_usuario(user_id)
        self.user_name = self.carregar_nome_usuario(user_id)

        # Configurações gerais da janela
        self.title("Tela Inicial")
        self.geometry("1200x700")
        self.configure(bg="#ecf0f1")
        self.state("zoomed")  # Janela maximizada
        self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self, self.caminho_icone)

        # Cria os componentes da interface
        self._criar_sidebar()
        self.hora_job = None  # Armazena o ID do callback do after
        self._criar_cabecalho()
        self._criar_main_content()
        self._criar_menubar()
        self._criar_menu_lateral()
        self.configurar_estilos()

        # cria barra de abas customizada (setas + canvas)
        self._criar_barra_abas()

        # cria o notebook
        self.notebook = ttk.Notebook(self.main_content, style="Hidden.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        # cria abas mínimas (placeholder)
        self._criar_abas_minimal()
        self.update_idletasks()  # garante que o placeholder apareça rápido

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

        # --- adiciona o prewarm aqui (opção imediata) ---
        self.prewarm_modules([
            "Base_produto", "Base_material", "Saida_NF", "Insercao_NF",
            "usuario", "Estoque", "media_custo", "relatorio_saida",
            "relatorio_cotacao", "registro_teste"
        ])

        # Atualiza estilos ao recuperar o foco e fecha o programa corretamente
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
        """Cria a sidebar de navegação com fundo em azul escuro e rolagem."""
        # Frame principal da sidebar
        sidebar_frame = tk.Frame(self, bg="#2c3e50", width=220)
        sidebar_frame.pack(side="left", fill="y")

        # Canvas para permitir rolagem
        self.sidebar_canvas = tk.Canvas(
            sidebar_frame,
            bg="#2c3e50",
            highlightthickness=0,
            width=220
        )
        self.sidebar_canvas.pack(side="left", fill="both", expand=True)

       # Scrollbar vertical (mais fina e com cor azul)
        scrollbar = tk.Scrollbar(
            sidebar_frame,
            orient="vertical",
            command=self.sidebar_canvas.yview,
            width=15,  # largura da barra de rolagem
            troughcolor="#1a252f",  # cor de fundo do trilho
            bg="#2980b9",           # cor da barra
            activebackground="#3498db"  # cor ao clicar
        )
        scrollbar.pack(side="right", pady=(60, 0), fill="y")

        self.sidebar_canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno que vai conter os widgets
        self.sidebar_inner = tk.Frame(self.sidebar_canvas, bg="#2c3e50")
        self.sidebar_canvas.create_window((0, 0), window=self.sidebar_inner, anchor="nw")

        # Atualiza região rolável automaticamente + “engana” para thumb menor
        self.sidebar_inner.bind(
            "<Configure>",
            lambda e: self.sidebar_canvas.configure(
                scrollregion=(0, 0, e.width, e.height + 200)  # +200 px extras para diminuir a thumb
            )
        )

        # Permite rolagem com roda do mouse
        def _on_mousewheel(event):
            self.sidebar_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.sidebar_canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _criar_cabecalho(self):
        """Cria o cabeçalho com a data (no formato DD/MM/YYYY), Hora, Título centralizado e Botão Atualizar à direita."""
        self.cabecalho = tk.Frame(self, bg="#34495e", height=80)
        self.cabecalho.pack(side="top", fill="x")

        # Configuração da grid para 3 colunas
        self.cabecalho.columnconfigure(0, weight=1)
        self.cabecalho.columnconfigure(1, weight=2)
        self.cabecalho.columnconfigure(2, weight=1)

        # Coluna 0: Data e Hora
        self.info_frame = tk.Frame(self.cabecalho, bg="#34495e")
        self.info_frame.grid(row=0, column=0, padx=15, sticky="w")

        # Rótulo para Data (apenas números, formato DD/MM/YYYY)
        self.data_label = tk.Label(
            self.info_frame, text="", fg="white", bg="#34495e",
            font=("Helvetica", 16, "bold")
        )
        self.data_label.pack(side="top", anchor="w")

        # Rótulo para Hora (fonte maior e destacada)
        self.hora_label = tk.Label(
            self.info_frame, text="", fg="#f1c40f", bg="#34495e",
            font=("Helvetica", 16, "bold")
        )
        self.hora_label.pack(side="top", anchor="w")

        # Coluna 1: Título centralizado
        titulo_label = tk.Label(
            self.cabecalho, text="Sistema Kametal", fg="white", bg="#34495e",
            font=("Helvetica", 26, "bold")
        )
        titulo_label.grid(row=0, column=1, padx=10)

        # Coluna 2: Botão Atualizar, alinhado à direita
        self.botao_atualizar = tk.Button(
            self.cabecalho, text="Atualizar", fg="white", bg="#2980b9",
            font=("Helvetica", 12, "bold"), bd=0, relief="flat", command=self.atualizar_pagina
        )
        self.botao_atualizar.bind("<Enter>", lambda e: self.botao_atualizar.config(bg="#3498db"))
        self.botao_atualizar.bind("<Leave>", lambda e: self.botao_atualizar.config(bg="#2980b9"))
        self.botao_atualizar.grid(row=0, column=2, padx=10, sticky="e")

        # Inicia a atualização do relógio
        self.atualizar_hora()

    def atualizar_hora(self):
        now = datetime.now()
        data_formatada = now.strftime("%d/%m/%Y")
        self.data_label.config(text=data_formatada)

        hora_formatada = now.strftime("%H:%M:%S")
        self.hora_label.config(text=hora_formatada)

        # Armazena o ID do callback para que ele possa ser cancelado depois
        self.hora_job = self.after(1000, self.atualizar_hora)

    def destroy(self):
        """Cancela callbacks e sinaliza threads antes de destruir a janela."""
        try:
            self.encerrar = True
        except Exception:
            pass

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

        try:
            if getattr(self, "conn_listen", None):
                try:
                    self.conn_listen.close()
                except Exception:
                    pass
                self.conn_listen = None
        except Exception:
            pass

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

    def _criar_abas(self):
        """Cria as abas do Notebook conforme self.permissoes e popula cada uma."""
        # limpa abas existentes
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        # Últimas NFs
        if "Janela_InsercaoNF" in self.permissoes or "SistemaNF" in self.permissoes:
            self.frame_nf = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_nf, text="Últimas NFs")
            self.criar_relatorio_nf()
        # Relatórios Produto/Material
        if "criar_interface_materiais" in self.permissoes or "criar_interface_produto" in self.permissoes:
            self.frame_relatorios_produto_material = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_relatorios_produto_material, text="Relatórios Produto/Material")
            self.criar_relatorio_produto_material()
        # Relatório de Estoque e Gráficos de Custo
        if "criar_media_custo" in self.permissoes:
            # relatório
            self.frame_relatorio = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_relatorio, text="Relatório de Estoque")
            self.criar_relatorio_estoque(self.conn)
            # gráfico
            self.frame_graficos = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_graficos, text="Gráficos de Custo")
            self.criar_grafico_mensal("Custo Médio de Cada Produto", [])
        # Gráfico Produtos
        if "CadastroProdutosApp" in self.permissoes:
            self.frame_grafico_produtos = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_grafico_produtos, text="Gráfico Produtos")
            self.criar_grafico_produtos()
        # Gráfico Dólar
        if "CadastroProdutosApp" in self.permissoes:
            self.frame_grafico_dolar = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_grafico_dolar, text="Gráfico Dólar")
            self.criar_grafico_dolar()
        # Últimos Registros de Teste (se permitido)
        if "RegistroTeste" in self.permissoes:
            self.frame_ultimos_registros_teste = tk.Frame(self.notebook, bg="#ecf0f1")
            self.notebook.add(self.frame_ultimos_registros_teste, text="Últimos Registros de Teste")
            self.criar_aba_ultimos_registros_teste()

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
        try:
            while not getattr(self, "encerrar", False) and not getattr(self, "_stop_event", threading.Event()).is_set():
                try:
                    # select bloqueia por 1s (como você já tinha)
                    select.select([self.conn_listen], [], [], 1)
                    self.conn_listen.poll()
                    while self.conn_listen.notifies:
                        notify = self.conn_listen.notifies.pop(0)
                        print(f"[NOTIFY] {notify.channel}: {notify.payload}")
                        try:
                            existe = self.winfo_exists()
                        except tk.TclError:
                            existe = False
                        if not getattr(self, "atualizacao_pendente", False) and existe:
                            self.atualizacao_pendente = True
                            try:
                                self.after(100, self.atualizar_pagina)
                            except tk.TclError:
                                # janela já fechada -> ignora
                                pass
                except Exception as e:
                    # se estiver encerrando, ignora erros transitórios
                    if not getattr(self, "encerrar", False) and not getattr(self, "_stop_event", threading.Event()).is_set():
                        print(f"[Erro na escuta]: {e}")
        except Exception as e_outer:
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
                    ORDER BY data DESC;
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
                    ORDER BY numero_nf ASC;
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

    def criar_relatorio_estoque(self, conn):
        from media_custo import buscar_estoque, buscar_produtos
        conn = conectar()
        if conn is None:
            print("Erro: Não foi possível conectar ao banco de dados.")
            return

        # Limpa widgets antigos da aba
        for widget in self.frame_relatorio.winfo_children():
            widget.destroy()

        produtos_bd = buscar_produtos(conn)
        produtos = []
        for produto in produtos_bd:
            nome_produto = produto[0]
            qtd_estoque = buscar_estoque(conn, nome_produto)
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

        try:
            self._cleanup()
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
            for job_attr in ("hora_job", "_spinner_job"):
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

        try:
            self._cleanup()
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