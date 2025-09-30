import tkinter as tk
from tkinter import ttk
from conexao_db import conectar  # Importa a função conectar
from tkinter import filedialog
import tkinter.messagebox as messagebox
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from relatorio_dolar import RelatorioDolar
from grafico_cotacao import AplicacaoGrafico
from datetime import datetime
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from matplotlib.backends.backend_pdf import PdfPages
import re
import threading

class CadastroProdutosApp:
    def __init__(self, root):
        self.root = root

        # Cria um Toplevel para o cadastro
        self.window = tk.Toplevel(self.root)
        self.window.title("Cotação")
        self.window.geometry("1000x600")
        self.window.state("zoomed")
        # self.window.option_add("*Font", "Helvetica 11")

        # Ícone e estilos
        aplicar_icone(self.window, r"C:\Sistema\logos\Kametal.ico")
        self.configurar_estilos()

        # Conexão com banco
        self.conn = conectar()
        if self.conn:
            self.cursor = self.conn.cursor()

        # Notebook e abas
        self.notebook = ttk.Notebook(self.window)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.tab_cadastro = ttk.Frame(self.notebook)
        self.tab_consulta = ttk.Frame(self.notebook)
        self.tab_relatorios = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_cadastro, text="Cotação de Produto")
        self.notebook.add(self.tab_consulta, text="Cotação de Dólar")
        self.notebook.add(self.tab_relatorios, text="Gráfico da Cotação")

        # Inicializa subcomponentes
        self.relatorio_dolar = None
        self.aba_grafico_produtos = AplicacaoGrafico(self.tab_relatorios)

        # Interface e dados
        self.criar_interface()
        self.carregar_dados()

        # Flag de encerramento e controle de loop
        self.encerrando    = False
        self._encerrar_id  = None

        # Rodapé com botões
        self.frame_voltar = ttk.Frame(self.window, padding=(10,5))
        self.frame_voltar.grid(row=1, column=0, sticky="ew")

        # frame específico para botões à direita (export / ajuda)
        self.frame_export = ttk.Frame(self.frame_voltar)
        self.frame_export.pack(side="right", padx=6, pady=5)

        # botões à esquerda (Voltar) permanecem no frame_voltar
        self.botao_voltar = ttk.Button(self.frame_voltar, text="Voltar", command=self.voltar, style="RelatorioCota.TButton")
        self.botao_voltar.pack(side="left", padx=(0,10), pady=5)

        # agora cria os botões de exportação dentro de frame_export
        self.botao_pdf = ttk.Button(self.frame_export, text="Exportar PDF", command=self.exportar_pdf, style="RelatorioCota.TButton")
        self.botao_excel = ttk.Button(self.frame_export, text="Exportar Excel", command=self.exportar_excel, style="RelatorioCota.TButton")
        self.botao_pdf.pack(side="right", padx=(10,0), pady=5)
        self.botao_excel.pack(side="right", padx=(10,0), pady=5)

        # --- Botão Ajuda (usa frame_export) ---
        self.botao_ajuda_cotacao = tk.Button(
            self.frame_export,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=self._abrir_ajuda_cotacao_modal
        )
        self.botao_ajuda_cotacao.pack(side="right", padx=6, pady=5)

        # hover, F1 bind no Toplevel e tooltip (mesmo que na opção A)
        self.botao_ajuda_cotacao.bind("<Enter>", lambda e: self.botao_ajuda_cotacao.config(bg="#3b5566"))
        self.botao_ajuda_cotacao.bind("<Leave>", lambda e: self.botao_ajuda_cotacao.config(bg="#2c3e50"))
        try:
            self.window.bind("<F1>", lambda e: self._abrir_ajuda_cotacao_modal())
        except Exception:
            pass
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_cotacao, "Ajuda — Relatórios de Cotação (F1)")
        except Exception:
            pass

        # Configuração de grid
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_rowconfigure(1, weight=0)
        self.window.grid_columnconfigure(0, weight=1)

        # 1) Configura o seu handler de bg_errors:
        self.window.report_callback_exception = self._handle_bg_error
        
        # Bind de abas e fechamento
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_tab_changed(self, event):
        aba_index = self.notebook.index(self.notebook.select())

        if aba_index == 1:  # Aba "Cotação de Dólar"
            if not self.relatorio_dolar:
                self.relatorio_dolar = RelatorioDolar(self.tab_consulta)

        elif aba_index == 2:  # Aba "Gráfico da Cotação"
            if not self.aba_grafico_produtos:
                self.aba_grafico_produtos = AplicacaoGrafico(self.tab_relatorios)

            # Atualiza o gráfico de Produtos e o Treeview de períodos
            self.aba_grafico_produtos.iniciar_atualizacao_produtos()
            self.aba_grafico_produtos.atualizar_treeview_periodos()
            
    def criar_interface(self):
        """Cria a interface dentro da aba de Cadastro"""
        
        # Frame superior para entrada do período e pesquisa
        self.frame_top = ttk.Frame(self.tab_cadastro, padding=(20, 10))
        self.frame_top.grid(row=0, column=0, sticky="ew")

        # Entrada de Período
        ttk.Label(self.frame_top, text="Período (DD/MM/AA a DD/MM/AA):").pack(side="left", padx=5)
        self.entrada_periodo = ttk.Entry(self.frame_top, width=25, font=("Helvetica", 12))
        self.entrada_periodo.pack(side="left", padx=5)
        self.entrada_periodo.bind("<KeyRelease>", self.formatar_periodo)

        # Barra de pesquisa (agora dentro do frame_top)
        ttk.Label(self.frame_top, text="Pesquisar:").pack(side="left", padx=20)
        self.entrada_pesquisa = ttk.Entry(self.frame_top, width=30, font=("Helvetica", 12))
        self.entrada_pesquisa.pack(side="left", padx=5)
        self.entrada_pesquisa.bind("<KeyRelease>", self.filtrar_dados)  # <-- essa linha ativa o filtro ao digitar

       
        # Frame do meio para entrada dos produtos
        self.frame_meio = ttk.Frame(self.tab_cadastro, padding=(20, 0))  # sem padding superior
        self.frame_meio.grid(row=1, column=0, sticky="ew")

        # Label de Produtos deslocado um pouco mais para baixo
        ttk.Label(self.frame_meio, text="Valores dos Produtos:").pack(side="left", padx=6, pady=(15, 0))  # ajustado de 10 para 15

        self.produtos = ["Cobre", "Zinco", "Liga 62/38", "Liga 65/35", "Liga 70/30", "Liga 85/15"]
        self.entradas = {}
        for produto in self.produtos:
            frame_produto = ttk.Frame(self.frame_meio)
            frame_produto.pack(side="left", padx=5)
            ttk.Label(frame_produto, text=produto).pack(side="top", padx=2)
            entrada = ttk.Entry(frame_produto, width=10, font=("Helvetica", 12))
            entrada.pack(side="top", padx=2)
            # bind do formatador automático:
            entrada.bind(
                "<KeyRelease>",
                lambda ev, ent=entrada: self._formatar_numero(ent)
            )
            self.entradas[produto] = entrada

        # Frame dos botões centralizados
        self.frame_btn = ttk.Frame(self.tab_cadastro, padding=(20, 10))
        self.frame_btn.grid(row=2, column=0, sticky="ew")
        self.frame_btn.grid_columnconfigure(0, weight=1)
        botoes_frame = ttk.Frame(self.frame_btn)
        botoes_frame.pack(anchor="center")
        self.botao_salvar = ttk.Button(botoes_frame, text="Salvar", command=self.salvar_dados, style="RelatorioCota.TButton")
        self.botao_salvar.pack(side="left", padx=5)
        self.botao_editar = ttk.Button(botoes_frame, text="Editar", command=self.editar_dados, style="RelatorioCota.TButton")
        self.botao_editar.pack(side="left", padx=5)
        self.botao_excluir = ttk.Button(botoes_frame, text="Excluir", command=self.excluir_dados, style="RelatorioCota.TButton")
        self.botao_excluir.pack(side="left", padx=5)
        self.botao_limpar = ttk.Button(botoes_frame, text="Limpar", command=self.limpar_entradas, style="RelatorioCota.TButton")
        self.botao_limpar.pack(side="left", padx=5)

        # Frame para o Treeview com Scrollbar
        self.frame_tree = ttk.Frame(self.tab_cadastro, padding=(20, 10))
        self.frame_tree.grid(row=3, column=0, sticky="nsew")
        self.tab_cadastro.rowconfigure(3, weight=1)
        self.tab_cadastro.columnconfigure(0, weight=1)

        # Definindo as colunas do Treeview, com a coluna "id" oculta
        colunas = ["id", "Período"] + self.produtos
        self.tree = ttk.Treeview(self.frame_tree, columns=colunas, show="headings", style="Cotacao.Treeview")
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.heading("id", text="ID")
        self.tree.column("id", width=0, minwidth=0, stretch=False)
        self.tree.heading("Período", text="Período")
        self.tree.column("Período", width=120, anchor="center")
        for coluna in self.produtos:
            self.tree.heading(coluna, text=coluna)
            self.tree.column(coluna, width=120, anchor="center")
        self.tree.config(style="Cotacao.Treeview", height=22)

        # Bind para que ao selecionar uma linha os dados apareçam nas entradas
        self.tree.bind('<<TreeviewSelect>>', self.selecionar_linha)

    def configurar_estilos(self):
        """
        Estilos LOCAIS para o relatório de cotação — NÃO altera 'Treeview' global
        nem toca nos estilos de botão.
        """
        estilo = ttk.Style(self.window)
        try:
            estilo.theme_use("alt")
        except Exception:
            pass

        # Estilo exclusivo para a Treeview deste módulo
        estilo.configure("Cotacao.Treeview",
                        font=("Arial", 10),
                        rowheight=27,
                        background="white",
                        foreground="black",
                        fieldbackground="white")
        estilo.configure("Cotacao.Treeview.Heading",
                        font=("Arial", 10, "bold"))
        estilo.map("Cotacao.Treeview",
                background=[("selected", "#0a64ad")],
                foreground=[("selected", "white")])

        # (opcional) estilo para linhas de total, se você usar algo assim:
        estilo.configure("Cotacao.Total.Treeview",
                        background="#FFD700",
                        foreground="black",
                        font=("Arial", 12, "bold"))
        
    def _create_tooltip(self, widget, text, delay=450):
        """Tooltip com quebra de linha automática e ajuste para não sair da tela."""
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            # cria janela do tooltip
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)

            # label com wrap para quebrar linhas
            label = tk.Label(
                win,
                text=text,
                bg="#333333",
                fg="white",
                font=("Segoe UI", 9),
                bd=0,
                padx=6,
                pady=4,
                wraplength=300  # máx. largura do tooltip (pixels)
            )
            label.pack()

            # calcula posição inicial
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            # força update para medir o tamanho real
            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # limites da tela
            screen_w = win.winfo_screenwidth()
            screen_h = win.winfo_screenheight()

            # ajusta posição horizontal se ultrapassar borda direita
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 0:
                x = 10

            # ajusta posição vertical se ultrapassar borda inferior
            if y + h > screen_h:
                y = widget.winfo_rooty() - h - 6  # mostra acima do widget

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

    def _abrir_ajuda_cotacao_modal(self, contexto=None):
        """
        Modal de ajuda para Relatórios / Cotação / Gráficos.
        Corrigido para usar `self.window` como parent do modal (evita AttributeError).
        """
        try:
            from tkinter import ttk

            # --- Janela modal ---
            modal = tk.Toplevel(self.window)              # <-- parent correto
            modal.title("Ajuda — Relatórios / Cotação")
            modal.transient(self.window)                  # <-- transient para o Toplevel pai
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões / centralização
            w, h = 900, 640
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # Ícone (silencioso se falhar)
            try:
                aplicar_icone(modal, r"C:\Sistema\logos\Kametal.ico")
            except Exception:
                pass

            # --- Cabeçalho ---
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Relatórios / Cotação", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # --- Corpo: navegação esquerda + conteúdo direita ---
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            # NAV FRAME (lado esquerdo) - sem scrollbar visível
            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0,12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10,6), padx=12)

            nav_list_container = tk.Frame(nav_frame, bg="#f6f8fa")
            nav_list_container.pack(fill="both", expand=True, padx=10, pady=(0,10))

            sections = [
                "Visão Geral",
                "Cotação — Aba Cadastro (Salvar / Editar / Excluir / Pesquisar)",
                "Cotação — Regras de Formatação (Período / Valores)",
                "Cotação Dólar — Inserir / Editar / Excluir / Média",
                "Editar Datas no Local — (Passo a Passo)",
                "Gráficos — Produtos",
                "Gráficos — Dólar",
                "Filtros e Pesquisa",
                "Exportação (Excel / PDF)",
                "FAQ"
            ]

            listbox = tk.Listbox(nav_list_container, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True)

            # --- Area de conteúdo (grid + painéis) ---
            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)
            content_frame.rowconfigure(0, weight=1)
            content_frame.columnconfigure(0, weight=1)

            # painel geral (padrão)
            general_frame = tk.Frame(content_frame, bg="white")
            general_frame.grid(row=0, column=0, sticky="nsew")
            general_frame.rowconfigure(0, weight=1)
            general_frame.columnconfigure(0, weight=1)

            txt_general = tk.Text(general_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_general = tk.Scrollbar(general_frame, command=txt_general.yview)  # scrollbar do painel de explicações
            txt_general.configure(yscrollcommand=sb_general.set)
            txt_general.grid(row=0, column=0, sticky="nsew")
            sb_general.grid(row=0, column=1, sticky="ns")

            # painel específico para edição/resumo (exemplo reaproveitável)
            editar_frame = tk.Frame(content_frame, bg="white")
            editar_frame.grid(row=0, column=0, sticky="nsew")
            editar_frame.rowconfigure(0, weight=1)
            editar_frame.columnconfigure(0, weight=1)

            txt_editar = tk.Text(editar_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_editar = tk.Scrollbar(editar_frame, command=txt_editar.yview)
            txt_editar.configure(yscrollcommand=sb_editar.set)
            txt_editar.grid(row=0, column=0, sticky="nsew")
            sb_editar.grid(row=0, column=1, sticky="ns")

            # --- Conteúdo das seções (textos detalhados solicitados) ---
            contents = {}

            contents["Visão Geral"] = (
            "Visão Geral\n\n"
            "Esta ajuda descreve as ações básicas em Cotação (Produtos) e Cotação (Dólar) e mostra, passo a passo, como editar datas *no próprio lugar* (sem precisar mexer no banco manualmente)."
            )

            contents["Cotação — Aba Cadastro (Salvar / Editar / Excluir / Pesquisar)"] = (
                "Cotação — Aba Cadastro (Produtos)\n\n"
                "Salvar:\n"
                "- Digite o Período no campo 'Período' usando apenas dígitos. Não insira '/' nem 'á'.\n"
                "- Preencha os valores dos produtos digitando só números (sem vírgula). O campo mostrará a vírgula automaticamente.\n"
                "- Clique em SALVAR.\n\n"
                "Editar:\n"
                "- Selecione a linha desejada na tabela (Treeview). Os campos de Período e valores aparecerão nas entradas acima.\n"
                "- Altere Período ou qualquer valor e clique em EDITAR para gravar a alteração.\n\n"
                "Excluir:\n"
                "- Selecione a linha e clique em EXCLUIR. Confirme para remover.\n\n"
                "Pesquisar:\n"
                "- Use o campo 'Pesquisar' para filtrar por período, data ou valor. Se digitar barras ('/'), a busca tenta interpretar como data; "
                "se digitar apenas números, a busca é por texto/substring."
            )

            contents["Cotação Dólar — Inserir / Editar / Excluir / Média"] = (
                "Cotação Dólar — Instruções essenciais\n\n"
                "Inserir:\n"
                "- Preencha 'Período' (opcional) com dígitos — o sistema formata as barras e o 'á'.\n"
                "- Em 'Data' digite apenas dígitos no formato DDMMYYYY (ex.: 01012025). O campo adicionará as barras automaticamente.\n"
                "- Em 'Dólar' digite apenas dígitos; o campo exibirá 4 casas decimais automaticamente.\n"
                "- Clique em INSERIR para gravar.\n\n"
                "Editar / Excluir:\n"
                "- Para editar uma data específica: selecione a linha *filha* (a linha que mostra Data e Dólar) dentro do Período.\n"
                "- O campo 'Data' será preenchido; altere para o novo valor (use apenas dígitos). Depois clique em EDITAR para atualizar.\n"
                "- Para remover, selecione e clique em EXCLUIR. Confirme para excluir.\n\n"
                "Média:\n"
                "- Ao carregar um Período, o sistema calcula automaticamente a média dos registros daquele Período e a exibe na interface."
            )

            contents["Cotação — Regras de Formatação (Período / Valores)"] = (
                "Regras de Formatação\n\n"
                "- Período: digite apenas dígitos; o sistema converte automaticamente para 'DD/MM/AA' e para 'DD/MM/AA á DD/MM/AA' quando houver intervalo.\n"
                "- Valores de produto: digite apenas dígitos (ex.: 12345 → o campo mostra 123,45). Não é preciso inserir vírgula manualmente.\n"
                "- Caso cole texto com pontos (separador de milhares), revise antes de salvar e remova pontos se necessário."
            )

            contents["Editar Datas no Local — (Passo a Passo)"] = (
                "Editar Datas no Local — Passo a passo (Produtos e Dólar)\n\n"
                "1) Selecionar o item correto:\n"
                "   - Cotação (Produtos): selecione a linha correspondente ao Período que quer alterar (a linha contém 'Período' e valores).\n"
                "   - Cotação (Dólar): expanda o Período e selecione a linha *de data* (filho) se quiser alterar apenas a Data ou o valor do Dólar.\n\n"
                "2) Os campos de entrada serão preenchidos automaticamente ao selecionar a linha.\n\n"
                "3) Alterar a Data no local:\n"
                "   - Em 'Data' (Dólar): digite DDMMYYYY (apenas números). O campo transforma para 'DD/MM/YYYY'.\n"
                "   - Em 'Período' (Produtos ou Dólar): digite os dígitos do período (ex.: 010125310125). O campo formata com barras e 'á'.\n\n"
                "4) Confirmar alteração:\n"
                "   - Clique em EDITAR para salvar a alteração do registro selecionado.\n"
                "   - Se quiser mudar a data para um Período diferente, edite o campo Período e então EDITAR (ou exclua e reinsira o registro no período desejado).\n\n"
                "Observações importantes:\n"
                "- Não há edição 'em massa' diretamente na árvore; edições em lote devem ser feitas via import/planilha ou através de excluir + inserir.\n"
                "- Para mover uma data de um período para outro sem perda de histórico, recomendamos exportar a linha, ajustar o Período e reimportar/inserir.\n"
            )

            contents["Gráficos — Produtos"] = (
                "Gráficos de Produtos — foco em filtros e interpretações\n\n"
                "- Controle de produtos: use as checkboxes para mostrar/ocultar séries (Cobre, Zinco, Ligas).\n"
                "- Datas/períodos: use a árvore de Períodos para escolher o intervalo exibido. A estatística lateral atualiza com Máximo, Mínimo e Média.\n"
                "- Estatísticas Gerais: o painel lateral exibe para cada produto:\n"
                "   • Máximo — o maior valor registrado no período selecionado;\n"
                "   • Mínimo — o menor valor registrado no período selecionado;\n"
                "   • Média — média aritmética dos pontos exibidos (soma / quantidade).\n"
                "  Use essas métricas para identificar amplitude (Máx − Mín) e tendência relativa entre produtos. As estatísticas são recalculadas automaticamente ao mudar filtros ou período.\n"
                "- Legenda: identifica a cor de cada produto; a legenda ajuda a relacionar a linha do gráfico com o produto correspondente."
            )

            contents["Gráficos — Dólar"] = (
                "Gráficos de Dólar — foco em datas e tendência\n\n"
                "- Selecione dias, meses ou anos na árvore lateral para filtrar o gráfico.\n"
                "- Estatísticas: o painel mostra Máximo, Mínimo e Média para a seleção atual e é atualizado automaticamente conforme você altera a seleção.\n"
                "- Tendência: o mini-gráfico (sparkline) indica direção geral (Alta / Queda / Estável) da série selecionada.\n"
                "- Tooltip: passe o mouse sobre os pontos da curva para ver um tooltip com a DATA e o VALOR exato do dólar; quando disponível, o tooltip também mostra a variação em relação ao ponto anterior."
            )

            contents["Filtros e Pesquisa"] = (
                "Filtros e Pesquisa\n\n"
                "- Use o campo 'Pesquisar' para localizar rapidamente períodos, datas ou valores.\n"
                "- Para busca por data exata, digite com '/' (ex.: 01/01/25) ou use apenas números para pesquisa por substring.\n"
            )

            contents["Exportação (Excel / PDF)"] = (
                "Exportação — lembretes úteis\n\n"
                "- Ao exportar para Excel ou PDF, confirme que todas as datas e valores foram corrigidos no local (Treeview) antes de gerar o arquivo.\n"
                "- Preferível exportar depois de ajustar médias/periodos para que os relatórios já reflitam as alterações."
            )

            contents["FAQ"] = (
                "FAQ — Perguntas rápidas\n\n"
                "Q: Posso editar várias datas ao mesmo tempo?\n"
                "A: Não diretamente na árvore. Edite linha a linha ou utilize exportação/importação para alteração em massa.\n\n"
                "Q: Preciso digitar barras no Período/Data?\n"
                "A: Não — digite apenas dígitos; a aplicação acrescenta '/' e 'á' automaticamente.\n\n"
                "Q: O que faço se transferir uma data para outro período?\n"
                "A: Edite o campo Período do registro e salve (ou exclua e reinsira no Período desejado). Para manter histórico, exporte antes de alterar."
            )

            # --- Função para exibir a seção correta (usa tkraise para painéis) ---
            def mostrar_secao(key):
                # Decide qual painel usar; exemplo: se for seção de edição usamos painel editar_frame
                usar_editar = (key == "Cotação — Regras de Formatação (Período / Valores)" or key == "Editar Estoque / Enviar para Resumo (Resumido)")
                if usar_editar:
                    txt_editar.configure(state="normal")
                    txt_editar.delete("1.0", "end")
                    txt_editar.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_editar.configure(state="disabled")
                    txt_editar.yview_moveto(0)
                    editar_frame.tkraise()
                else:
                    txt_general.configure(state="normal")
                    txt_general.delete("1.0", "end")
                    txt_general.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_general.configure(state="disabled")
                    txt_general.yview_moveto(0)
                    general_frame.tkraise()

            # Inicializa exibindo a primeira seção
            listbox.selection_set(0)
            mostrar_secao(sections[0])

            # Bind para seleção
            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])
            listbox.bind("<<ListboxSelect>>", on_select)

            # --- Bindings do mousewheel (melhora rolagem) ---
            def _on_mousewheel_text(event):
                if hasattr(event, "delta") and event.delta:
                    # Windows/macOS
                    txt_general.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    txt_editar.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    txt_general.yview_scroll(-1, "units")
                    txt_editar.yview_scroll(-1, "units")
                elif event.num == 5:
                    txt_general.yview_scroll(1, "units")
                    txt_editar.yview_scroll(1, "units")

            def _on_mousewheel_listbox(event):
                if hasattr(event, "delta") and event.delta:
                    listbox.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    listbox.yview_scroll(-1, "units")
                elif event.num == 5:
                    listbox.yview_scroll(1, "units")

            txt_general.bind("<MouseWheel>", _on_mousewheel_text)
            txt_general.bind("<Button-4>", _on_mousewheel_text)
            txt_general.bind("<Button-5>", _on_mousewheel_text)
            txt_editar.bind("<MouseWheel>", _on_mousewheel_text)
            txt_editar.bind("<Button-4>", _on_mousewheel_text)
            txt_editar.bind("<Button-5>", _on_mousewheel_text)
            listbox.bind("<MouseWheel>", _on_mousewheel_listbox)
            listbox.bind("<Button-4>", _on_mousewheel_listbox)
            listbox.bind("<Button-5>", _on_mousewheel_listbox)

            # --- Rodapé com botão Fechar ---
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right", padx=6)

            # bind ESC para fechar o modal
            modal.bind("<Escape>", lambda e: modal.destroy())

            # atalho F1 para o modal (opcional): vincula ao modal para que quando ele estiver aberto o F1 funcione
            try:
                modal.bind("<F1>", lambda e: None)
            except Exception:
                pass

            modal.focus_set()
            modal.wait_window()

        except Exception as e:
            # mostra erro mais informativo no console (para debugar)
            import traceback
            traceback.print_exc()
            print("Erro ao abrir modal de ajuda (Relatórios):", e)

   
    def obter_menor_id_disponivel(self):
        self.cursor.execute("SELECT id FROM cotacao_produtos ORDER BY id")
        ids = [row[0] for row in self.cursor.fetchall()]
        menor_id = 1
        for id_existente in ids:
            if id_existente == menor_id:
                menor_id += 1
            else:
                break
        return menor_id
    
    def _formatar_numero(self, entry_widget):
        """Formata o conteúdo do entry_widget para '1.234,56' enquanto digita."""
        texto = entry_widget.get()
        # 1) retira tudo que não for dígito
        dígitos = re.sub(r"\D", "", texto)
        if not dígitos:
            novo = "0,00"
        else:
            # 2) transforma em inteiro e divide por 100 para obter float
            valor = int(dígitos) / 100
            # 3) formata com vírgula e pontos de milhar
            #    primeiro gera '1,234.56' no estilo en_US e depois troca
            s = f"{valor:,.2f}"  
            # swap de separadores: '.'->'X', ','->'.', 'X'->','
            novo = s.replace(",", "X").replace(".", ",").replace("X", ".")

        # atualiza o entry sem disparar múltiplos binds
        entry_widget.unbind("<KeyRelease>")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, novo)
        entry_widget.bind(
            "<KeyRelease>",
            lambda ev, ent=entry_widget: self._formatar_numero(ent)
        )

    def formatar_valor(self, valor):
        """Formata o valor para o formato monetário 'R$'."""
        try:
            if isinstance(valor, str):  # Se o valor for uma string
                valor = valor.replace("R$", "").replace(".", "").replace(",", ".")
            
            valor_float = float(valor)  # Converte para float
            return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "R$ 0,00"  # Retorna "R$ 0,00" se ocorrer algum erro na conversão

    def converter_entrada(self, valor):
        """Converte valores de entrada para float, aceitando vírgula como separador decimal."""
        return float(valor.replace(",", ".")) if valor else None
    
    def formatar_periodo(self, event):
        # Obtém o texto atual e extrai apenas os dígitos
        raw_text = self.entrada_periodo.get()
        texto = ''.join(filter(str.isdigit, raw_text))
        
        # Limita a no máximo 12 dígitos (6 para cada data)
        if len(texto) > 12:
            texto = texto[:12]
        
        # Se ainda não temos 6 dígitos, não formata (permite digitar sem interferência)
        if len(texto) < 6:
            return

        # Formata a primeira data (primeiros 6 dígitos)
        primeira_data = f"{texto[0:2]}/{texto[2:4]}/{texto[4:6]}"
        
        # Se houver mais de 6 dígitos, formata o segundo período
        if len(texto) > 6:
            # Se houver pelo menos 2 dígitos para a segunda data, formata conforme disponíveis
            segunda_data = ""
            if len(texto) >= 8:
                segunda_data = f"{texto[6:8]}"
                if len(texto) >= 10:
                    segunda_data += f"/{texto[8:10]}"
                    if len(texto) >= 12:
                        segunda_data += f"/{texto[10:12]}"
                    else:
                        segunda_data += f"/{texto[10:]}"
                else:
                    segunda_data += f"/{texto[8:]}"
            else:
                segunda_data = texto[6:]
            texto_formatado = primeira_data + " á " + segunda_data
        else:
            texto_formatado = primeira_data

        # Atualiza a Entry com o texto formatado
        self.entrada_periodo.delete(0, tk.END)
        self.entrada_periodo.insert(0, texto_formatado)

    def carregar_dados(self):
        """
        Carrega os dados do banco e atualiza a Treeview.
        A ordenação é feita diretamente na query SQL usando a conversão da primeira data do campo "periodo".
        """
        query = """
            SELECT id, periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15 
            FROM cotacao_produtos
            ORDER BY to_date(split_part(periodo, ' á ', 1), 'DD/MM/YY') DESC
        """
        self.cursor.execute(query)
        registros = self.cursor.fetchall()

        # Limpa todos os itens atuais da Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insere os registros ordenados
        for row in registros:
            valores_formatados = [row[0], row[1]] + [self.formatar_valor(valor) for valor in row[2:]]
            self.tree.insert("", "end", values=valores_formatados)

        # Garante que nenhuma linha esteja selecionada
        self.tree.selection_remove(self.tree.selection())

    def filtrar_dados(self, event=None):
        """Filtra os dados no Treeview com base na pesquisa, considerando datas somente se o usuário digitar com '/'. Nenhuma formatação automática é feita."""
        termo_pesquisa = self.entrada_pesquisa.get().strip().lower()
        
        # Limpa os itens atuais do Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Recarrega os dados do banco de dados e aplica o filtro
        query = """
            SELECT id, periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15 
            FROM cotacao_produtos
            ORDER BY id DESC
        """
        self.cursor.execute(query)
        registros = self.cursor.fetchall()

        from datetime import datetime

        def parse_date(date_str):
            """Tenta converter uma string para datetime no formato dd/mm/yy."""
            try:
                return datetime.strptime(date_str, "%d/%m/%y")
            except Exception:
                return None

        for row in registros:
            periodo = row[1]
            match_data = False
            data_inicio = data_fim = None

            # Só tenta interpretar como data se o termo tiver barra (ex: 01/01/24)
            if "/" in termo_pesquisa:
                try:
                    # Extrai datas do período (ex: "01/01/24 á 07/01/24")
                    if "á" in periodo:
                        inicio_str, fim_str = [s.strip() for s in periodo.split("á")]
                    elif " a " in periodo:
                        inicio_str, fim_str = [s.strip() for s in periodo.split(" a ")]
                    else:
                        inicio_str = fim_str = periodo.strip()
                    data_inicio = parse_date(inicio_str)
                    data_fim = parse_date(fim_str)
                except Exception:
                    data_inicio = data_fim = None

                # Pesquisa como intervalo de datas se o termo tiver "á" ou " a "
                if "á" in termo_pesquisa or " a " in termo_pesquisa:
                    try:
                        if "á" in termo_pesquisa:
                            pesquisa_inicio_str, pesquisa_fim_str = [s.strip() for s in termo_pesquisa.split("á")]
                        else:
                            pesquisa_inicio_str, pesquisa_fim_str = [s.strip() for s in termo_pesquisa.split(" a ")]
                        pesquisa_inicio = parse_date(pesquisa_inicio_str)
                        pesquisa_fim = parse_date(pesquisa_fim_str)
                        if data_inicio and data_fim and pesquisa_inicio and pesquisa_fim:
                            if data_inicio <= pesquisa_fim and data_fim >= pesquisa_inicio:
                                match_data = True
                    except Exception:
                        pass
                else:
                    # Pesquisa por data única (dentro do intervalo)
                    try:
                        pesquisa_data = parse_date(termo_pesquisa)
                        if data_inicio and data_fim and pesquisa_data:
                            if data_inicio <= pesquisa_data <= data_fim:
                                match_data = True
                    except Exception:
                        pass

            # Sempre faz pesquisa por substring, como fallback ou pesquisa textual
            valores_formatados = [str(row[0]), periodo] + [self.formatar_valor(valor) for valor in row[2:]]
            if any(termo_pesquisa in str(valor).lower() for valor in valores_formatados):
                match_data = True

            if match_data:
                self.tree.insert("", "end", values=valores_formatados)

    def evento_composto(self, event):
        """Executa formatação e pesquisa simultaneamente."""
        self.formatar_data_pesquisa(event)
        self.filtrar_dados(event)

    def salvar_dados(self):
        """
        Salva os dados capturados das entradas no banco e atualiza a Treeview.
        Verifica antes se todos os campos foram preenchidos.
        """
        # 1) Verifica período
        periodo = self.entrada_periodo.get().strip()
        if not periodo:
            messagebox.showerror("Erro", "O campo Período é obrigatório.")
            return

        # 2) Verifica cada produto
        for produto, entry in self.entradas.items():
            if not entry.get().strip():
                messagebox.showerror("Erro", f"O valor de '{produto}' é obrigatório.")
                return

        # 3) Se passou na validação, converte e salva
        # Converte entradas formatadas tipo '1.234,56' para float
        def to_float(texto):
            t = re.sub(r"[^\d,]", "", texto)        # mantém dígitos e vírgula
            t = t.replace(".", "").replace(",", ".")# transforma em padrão Python
            try:
                return float(t)
            except ValueError:
                return None

        dados_float = [periodo] + [to_float(entry.get().strip()) for entry in self.entradas.values()]

        # 4) Obtém ID e executa o INSERT
        menor_id = self.obter_menor_id_disponivel()
        query = """
            INSERT INTO cotacao_produtos
            (id, periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        params = (menor_id, dados_float[0], *dados_float[1:])
        self.cursor.execute(query, params)
        self.conn.commit()

        # 5) Notifica e recarrega
        self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        self.conn.commit()
        self.carregar_dados()

        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
        self.limpar_entradas()

    def editar_dados(self):
        selected_item = self.tree.selection()  # Obtém o item selecionado
        if selected_item:
            valores = self.tree.item(selected_item, "values")
            novo_periodo = self.entrada_periodo.get()
            novos_produtos = [self.entradas[produto].get() for produto in self.produtos]

            # Formata os valores antes de exibir no Treeview (incluindo "R$")
            dados_formatados = [novo_periodo] + [self.formatar_valor(valor) for valor in novos_produtos]

            # Atualiza os valores no Treeview (aplicando "R$")
            self.tree.item(selected_item, values=[valores[0]] + dados_formatados)

            # Converte os valores para salvar no banco (removendo "R$" e ajustando separadores)
            dados_float = [novo_periodo]
            for valor in novos_produtos:
                try:
                    valor_sem_rs = valor.replace("R$", "").replace(".", "").replace(",", ".")
                    dados_float.append(float(valor_sem_rs) if valor else None)
                except ValueError:
                    dados_float.append(None)

            # Atualiza no banco de dados
            query = """
                UPDATE cotacao_produtos 
                SET periodo = %s, cobre = %s, zinco = %s, liga_62_38 = %s, liga_65_35 = %s, liga_70_30 = %s, liga_85_15 = %s
                WHERE id = %s
            """
            self.cursor.execute(query, (dados_float[0], dados_float[1], dados_float[2],
                                        dados_float[3], dados_float[4], dados_float[5],
                                        dados_float[6], valores[0]))
            self.conn.commit()

            self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            self.conn.commit()

            # Mensagem de sucesso
            messagebox.showinfo("Sucesso", "Dados editados com sucesso!")
            print("Dados atualizados no banco de dados.")

            # Limpa as entradas após editar
            self.limpar_entradas()
    
    def selecionar_linha(self, event=None):
        selected_items = self.tree.selection()
        if selected_items:
            # Usa o primeiro item selecionado
            item_id = selected_items[0]
            valores = self.tree.item(item_id, "values")
            
            # Preenche a entrada de período
            self.entrada_periodo.delete(0, tk.END)
            self.entrada_periodo.insert(0, valores[1])
            
            # Preenche as entradas dos produtos
            for i, produto in enumerate(self.produtos):
                self.entradas[produto].delete(0, tk.END)
                
                # Remove o "R$" e ajusta o formato para apenas o valor numérico
                valor_limpo = valores[i + 2]
                if isinstance(valor_limpo, str):
                    valor_limpo = valor_limpo.replace("R$", "").strip()  # Remove "R$"
                
                self.entradas[produto].insert(0, valor_limpo)
        
    def excluir_dados(self):
        selected_items = self.tree.selection()
        if selected_items:
            resposta = messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir os itens selecionados?")
            if resposta:
                for item in selected_items:
                    valores = self.tree.item(item, "values")
                    id_item = valores[0]
                    query = "DELETE FROM cotacao_produtos WHERE id = %s"
                    self.cursor.execute(query, (id_item,))
                    self.conn.commit()

                    self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    self.conn.commit()

                    self.tree.delete(item)
                    print(f"Dado com ID {id_item} excluído do banco de dados.")
                # Limpa as entradas após a exclusão
                self.limpar_entradas()
            else:
                print("Exclusão cancelada.")

    def exportar_excel(self):
        aba_index = self.notebook.index(self.notebook.select())

        if aba_index == 0:
            self.exportar_relatorio_excel_produto()

        elif aba_index == 1:
            if self.relatorio_dolar:
                self.relatorio_dolar.exportar_relatorio_excel_dolar()
            else:
                messagebox.showwarning("Exportação", "A aba Cotação de Dólar ainda não foi inicializada.")

        elif aba_index == 2:
            if self.aba_grafico_produtos:
                # Verifica qual subaba está ativa: Gráfico de Produtos ou Gráfico de Dólar
                subaba_index = self.aba_grafico_produtos.abas.index(self.aba_grafico_produtos.abas.select())
                if subaba_index == 0:
                    self.aba_grafico_produtos.exportar_grafico_excel()
                elif subaba_index == 1:
                    try:
                        self.aba_grafico_produtos.grafico_dolar.exportar_excel_grafico()
                    except Exception as e:
                        messagebox.showerror("Exportação", f"Erro ao exportar gráfico de Dólar:\n{str(e)}")
                else:
                    messagebox.showwarning("Exportação", "Subaba não reconhecida.")
            else:
                messagebox.showwarning("Exportação", "A aba 'Gráfico da Cotação' não foi inicializada.")

    def exportar_relatorio_excel_produto(self):
        """Abre uma janela com filtros de período e exporta os dados do Treeview para um arquivo Excel."""

        # Função para formatar o texto da data enquanto o usuário digita
        def formatar_data(event):
            entry = event.widget  # Obtém o campo de entrada que chamou a função
            data = entry.get().strip()
            # Remove caracteres não numéricos
            data = ''.join(filter(str.isdigit, data))
            # Mantém apenas os 8 primeiros caracteres
            data = data[:8]
            # Formata a data conforme a quantidade de caracteres
            if len(data) >= 2:
                data = data[:2] + '/' + data[2:]
            if len(data) >= 5:
                data = data[:5] + '/' + data[5:]
            # Atualiza o campo de entrada com a data formatada
            entry.delete(0, tk.END)
            entry.insert(0, data)

        def gerar_excel():
            data_inicio = entry_inicio.get().strip()
            data_fim    = entry_fim.get().strip()

            try:
                data_inicio_dt = datetime.strptime(data_inicio, "%d/%m/%y") if data_inicio else None
                data_fim_dt    = datetime.strptime(data_fim,    "%d/%m/%y") if data_fim    else None
            except ValueError:
                messagebox.showerror("Filtro de Data", "Formato de data inválido. Use DD/MM/YY.")
                return

            # obtém colunas e dados do Treeview…
            colunas_tree = self.tree["columns"]
            nomes_colunas = [col for col in colunas_tree if col != "id"]
            dados = [self.tree.item(item)["values"][1:] for item in self.tree.get_children()]

            df = pd.DataFrame(dados, columns=nomes_colunas)

            # Extrai Data_Inicio e Data_Fim do campo "Período"
            df["Data_Inicio"] = df["Período"].apply(
                lambda x: datetime.strptime(x.split(" á ")[0], "%d/%m/%y")
            )
            df["Data_Fim"] = df["Período"].apply(
                lambda x: datetime.strptime(x.split(" á ")[1], "%d/%m/%y")
            )

            # Filtra por overlap:
            #   queremos todo registro cujo período [Data_Inicio, Data_Fim]
            #   sobreponha [data_inicio_dt, data_fim_dt].
            if data_inicio_dt and data_fim_dt:
                df = df[
                    (df["Data_Inicio"] <= data_fim_dt) &
                    (df["Data_Fim"]    >= data_inicio_dt)
                ]
            elif data_inicio_dt:
                # tudo que passe por data_inicio_dt
                df = df[
                    (df["Data_Inicio"] <= data_inicio_dt) &
                    (df["Data_Fim"]    >= data_inicio_dt)
                ]
            elif data_fim_dt:
                # tudo que termine até data_fim_dt
                df = df[df["Data_Fim"] <= data_fim_dt]

            # reordena e limpa colunas auxiliares
            df.sort_values("Data_Inicio", ascending=False, inplace=True)
            df.drop(columns=["Data_Inicio", "Data_Fim"], inplace=True)

            # salva para Excel…
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Relatorio_Cotacao.xlsx",
                filetypes=[("Planilhas do Excel", "*.xlsx")],
                title="Salvar Relatório de Cotação de Produto"
            )
            if not caminho_arquivo:
                return

            try:
                with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Relatório Cotação", index=False)
                    ws = writer.sheets["Relatório Cotação"]
                    ws.auto_filter.ref = ws.dimensions

                messagebox.showinfo("Exportar Excel",
                                    f"Excel exportado com sucesso para:\n{caminho_arquivo}")
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Exportar Excel", f"Erro ao exportar Excel: {e}")

        # Cria janela popup
        popup = tk.Toplevel(self.root)
        popup.title("Exportar Excel - Filtro de Período")
        popup.geometry("300x180")
        popup.resizable(False, False)

        centralizar_janela(popup, 200, 200)
        aplicar_icone(popup, r"C:\Sistema\logos\Kametal.ico")
        popup.config(bg="#ecf0f1")

        # Estilo personalizado
        style = ttk.Style(popup)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        # Labels e Entradas
        ttk.Label(popup, text="Data Início (DD/MM/YY):", style="Custom.TLabel").pack(pady=5)
        entry_inicio = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_inicio.pack(pady=5)
        entry_inicio.bind("<KeyRelease>", formatar_data)

        ttk.Label(popup, text="Data Fim (DD/MM/YY):", style="Custom.TLabel").pack(pady=5)
        entry_fim = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_fim.pack(pady=5)
        entry_fim.bind("<KeyRelease>", formatar_data)

        # Botões
        botoes_frame = ttk.Frame(popup, style="Custom.TFrame")
        botoes_frame.pack(pady=10)

        ttk.Button(botoes_frame, text="Exportar Excel", command=gerar_excel).pack(side="left", padx=5)
        ttk.Button(botoes_frame, text="Cancelar", command=popup.destroy).pack(side="left", padx=5)

    def exportar_pdf(self):
        """Determina a aba ativa e delega para a exportação do PDF correspondente."""
        aba_index = self.notebook.index(self.notebook.select())

        if aba_index == 0:
            self.exportar_relatorio_pdf_produto()

        elif aba_index == 1:
            if self.relatorio_dolar:
                self.relatorio_dolar.exportar_relatorio_pdf_dolar()
            else:
                messagebox.showwarning("Exportação", "A aba Cotação de Dólar ainda não foi inicializada.")

        elif aba_index == 2:
            if self.aba_grafico_produtos:
                # Verifica qual subaba está ativa: Gráfico de Produtos ou Gráfico de Dólar
                subaba_index = self.aba_grafico_produtos.abas.index(self.aba_grafico_produtos.abas.select())
                if subaba_index == 0:
                    self.aba_grafico_produtos.exportar_grafico_pdf()
                elif subaba_index == 1:
                    try:
                        self.aba_grafico_produtos.grafico_dolar.exportar_grafico_pdf()
                    except Exception as e:
                        messagebox.showerror("Exportação", f"Erro ao exportar gráfico de Dólar:\n{str(e)}")
                else:
                    messagebox.showwarning("Exportação", "Subaba não reconhecida.")
            else:
                messagebox.showwarning("Exportação", "A aba 'Gráfico da Cotação' não foi inicializada.")

    def exportar_relatorio_pdf_produto(self):
        """Abre uma janela para filtro de período e exporta o PDF da cotação de produtos."""

        def formatar_data(event):
            entry = event.widget  # Obtém o campo de entrada que chamou a função
            data = entry.get().strip()
            # Remove caracteres não numéricos
            data = ''.join(filter(str.isdigit, data))
            # Mantém apenas os 8 primeiros caracteres
            data = data[:8]
            # Formata a data conforme a quantidade de caracteres
            if len(data) >= 2:
                data = data[:2] + '/' + data[2:]
            if len(data) >= 5:
                data = data[:5] + '/' + data[5:]
            # Atualiza o campo de entrada com a data formatada
            entry.delete(0, tk.END)
            entry.insert(0, data)

        def gerar_pdf():
            data_inicio = entry_inicio.get().strip()
            data_fim    = entry_fim.get().strip()

            # Converte strings para formato SQL (YYYY-MM-DD)
            data_inicio_sql = None
            data_fim_sql    = None
            try:
                if data_inicio:
                    data_inicio_sql = datetime.strptime(data_inicio, "%d/%m/%y").strftime("%Y-%m-%d")
                if data_fim:
                    data_fim_sql    = datetime.strptime(data_fim, "%d/%m/%y").strftime("%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Filtro de Data", "Formato inválido. Use DD/MM/YY.")
                return

            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialfile="relatorio_produto_cotacao.pdf",
                filetypes=[("Arquivo PDF", "*.pdf")],
                title="Salvar Relatório de Cotação de Produto"
            )
            if not caminho:
                return

            try:
                # === Busca os dados ===
                query = """
                    SELECT periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15
                    FROM cotacao_produtos
                    WHERE 1=1
                """
                params = []
                if data_inicio_sql and data_fim_sql:
                    # overlap: início ≤ data_fim  E  fim ≥ data_inicio
                    query += """
                        AND to_date(split_part(periodo,' á ',1),'DD/MM/YY') <= %s
                        AND to_date(split_part(periodo,' á ',2),'DD/MM/YY') >= %s
                    """
                    params.extend([data_fim_sql, data_inicio_sql])

                elif data_inicio_sql:
                    # pega tudo que contenha essa data
                    query += """
                        AND to_date(split_part(periodo,' á ',1),'DD/MM/YY') <= %s
                        AND to_date(split_part(periodo,' á ',2),'DD/MM/YY') >= %s
                    """
                    # aqui data_inicio_sql deve entrar nas duas posições
                    params.extend([data_inicio_sql, data_inicio_sql])

                elif data_fim_sql:
                    # tudo que termine até data_fim
                    query += " AND to_date(split_part(periodo,' á ',2),'DD/MM/YY') <= %s"
                    params.append(data_fim_sql)

                # ordenação (pela data de início, mais recente primeiro)
                query += """
                    ORDER BY to_date(split_part(periodo,' á ',1),'DD/MM/YY') DESC
                """

                # executa…
                self.cursor.execute(query, tuple(params))
                registros = self.cursor.fetchall()

                # === Geração do PDF via Platypus ===
                doc       = SimpleDocTemplate(caminho, pagesize=A4)
                elementos = []

                estilos = getSampleStyleSheet()
                # Título
                elementos.append(Paragraph("<b>Relatório de Cotação de Produto</b>", estilos["Title"]))
                elementos.append(Spacer(1, 20))

                # Cabeçalhos
                headers = ["Período", "Cobre", "Zinco",
                        "Liga 62/38", "Liga 65/35",
                        "Liga 70/30", "Liga 85/15"]
                dados = [headers]

                # Linhas de dados
                for row in registros:
                    linha = [row[0]] + [self.formatar_valor(v) for v in row[1:]]
                    dados.append(linha)

                # Monta a tabela com estilo uniforme
                tabela = Table(dados)
                tabela.setStyle(TableStyle([
                    ("BACKGROUND",    (0, 0), (-1, 0), colors.grey),
                    ("TEXTCOLOR",     (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                    ("GRID",          (0, 0), (-1, -1), 0.5, colors.black),
                ]))
                elementos.append(tabela)

                # Constrói o documento
                doc.build(elementos)

                messagebox.showinfo("Exportar PDF", f"PDF salvo em:\n{caminho}")
                popup.destroy()

            except Exception as e:
                messagebox.showerror("Exportar PDF", f"Erro ao exportar PDF: {e}")

        # --- Cria o popup de filtro (mantido igual) ---
        popup = tk.Toplevel(self.root)
        popup.title("Exportar PDF - Filtro de Período")
        popup.geometry("300x180")
        popup.resizable(False, False)
        centralizar_janela(popup, 300, 180)
        aplicar_icone(popup, r"C:\\Sistema\\logos\\Kametal.ico")
        popup.config(bg="#ecf0f1")

        style = ttk.Style(popup)
        style.theme_use("alt")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        ttk.Label(popup, text="Data Início (DD/MM/YY):", style="Custom.TLabel").pack(pady=5)
        entry_inicio = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_inicio.pack(pady=5)
        entry_inicio.bind("<KeyRelease>", formatar_data)

        ttk.Label(popup, text="Data Fim (DD/MM/YY):", style="Custom.TLabel").pack(pady=5)
        entry_fim = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_fim.pack(pady=5)
        entry_fim.bind("<KeyRelease>", formatar_data)

        botoes_frame = ttk.Frame(popup)
        botoes_frame.pack(pady=10)
        ttk.Button(botoes_frame, text="Exportar PDF", command=gerar_pdf).pack(side="left", padx=5)
        ttk.Button(botoes_frame, text="Cancelar", command=popup.destroy).pack(side="left", padx=5)

    def limpar_entradas(self):
        # Limpa a entrada do período
        self.entrada_periodo.delete(0, tk.END)
        # Limpa todas as entradas de produtos
        for produto in self.produtos:
            self.entradas[produto].delete(0, tk.END)

    def voltar(self):
        """Cancela callbacks, reexibe menu e destrói window em background."""
        # cancela callback agendado, se houver
        if getattr(self, "_encerrar_id", None) is not None:
            try:
                self.window.after_cancel(self._encerrar_id)
            except Exception:
                pass

        # mostra o menu (root) antes de qualquer destruição
        try:
            self.root.deiconify()
            self.root.state("zoomed")
            self.root.lift()
            self.root.update_idletasks()
            try:
                self.root.focus_force()
            except Exception:
                pass
        except Exception:
            pass

        def _cleanup_and_destroy():
            try:
                if hasattr(self, "cursor") and getattr(self, "cursor"):
                    try:
                        self.cursor.close()
                    except Exception:
                        pass
                if hasattr(self, "conn") and getattr(self, "conn"):
                    try:
                        self.conn.close()
                    except Exception:
                        pass
            finally:
                try:
                    if hasattr(self.window, "after"):
                        self.window.after(0, self.window.destroy)
                    else:
                        self.window.destroy()
                except Exception:
                    try:
                        self.window.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()
             
    def __del__(self):
        if hasattr(self, "conn"):
            self.conn.close()

    def _handle_bg_error(self, exc, val, tb):
        # Aqui você “intercepta” qualquer exceção disparada em callbacks agendados
        # Pode logar, mostrar um messagebox ou simplesmente ignorar:
        print("Erro em callback (ignorado):", val)

    def on_closing(self):
        # 1) confirmação
        if not messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            return

        # 2) sinaliza encerramento e cancela qualquer after próprio
        self.encerrando = True
        if getattr(self, "_encerrar_id", None):
            try:
                self.window.after_cancel(self._encerrar_id)
            except tk.TclError:
                pass

        # Se tiver outros after próprios, cancele aqui também:
        for job_attr in ("id_verificar", "job_atualizar_periodos", "job_atualizar_datas"):
            job = getattr(self, job_attr, None)
            if job:
                try:
                    self.window.after_cancel(job)
                except tk.TclError:
                    pass

        # 3) pede para os sub‐componentes fecharem
        if self.relatorio_dolar and hasattr(self.relatorio_dolar, "on_closing"):
            self.relatorio_dolar.on_closing()
        if self.aba_grafico_produtos and hasattr(self.aba_grafico_produtos, "on_closing"):
            self.aba_grafico_produtos.on_closing()

        # 4) inicia a checagem com ID próprio
        self._encerrar_id = self.window.after(200, self._verificar_encerramento)

    def _verificar_encerramento(self):
        # Em vez de confiar só na flag, cheque se cada janela ainda existe:
        grafico_produtos_rodando = (
            hasattr(self.aba_grafico_produtos, "winfo_exists")
            and self.aba_grafico_produtos.winfo_exists()
        )
        relatorio_dolar_rodando = (
            hasattr(self.relatorio_dolar, "winfo_exists")
            and self.relatorio_dolar.winfo_exists()
        )

        if not grafico_produtos_rodando and not relatorio_dolar_rodando:
            print("Todos encerrados, destruindo janela de cadastro")
            try:
                self.window.destroy()
            except tk.TclError:
                pass
        else:
            # Se ainda há componentes ativos, reagenda a checagem
            print("Aguardando encerramento dos subcomponentes…")
            self._encerrar_id = self.window.after(200, self._verificar_encerramento)