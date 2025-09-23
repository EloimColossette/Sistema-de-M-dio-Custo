import tkinter as tk
from tkinter import ttk
from conexao_db import conectar
import psycopg2
from tkinter import messagebox, filedialog
import pandas as pd
import sys
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from decimal import Decimal, InvalidOperation
from datetime import datetime, date, timedelta,timezone
import unicodedata
import threading
import traceback
import getpass
import psycopg2.extensions
import select
import time
import calendar

class CalculoProduto:
    def __init__(self, janela_menu=None, *args, **kwargs):
        # guarda referência do menu (se foi passada)
        self.janela_menu = janela_menu or kwargs.get("janela_menu", None)
        self.janela_menu = janela_menu
        self.root = tk.Toplevel()
        self.root.title("Calculo de Nfs")
        self.root.geometry("1200x700")
        self.root.state("zoomed")

         # determina usuário para registrar: prioridade
        # 1) menu.user_name  2) kwargs['usuario']  3) getpass.getuser()
        try:
            if self.janela_menu and getattr(self.janela_menu, "user_name", None):
                self.usuario = self.janela_menu.user_name
            else:
                self.usuario = kwargs.get("usuario") or getpass.getuser()
        except Exception:
            self.usuario = kwargs.get("usuario") or getpass.getuser()

        # Oculta a janela de menu
        self.janela_menu.withdraw()

        # Aplica o ícone (defina ou adapte a função aplicar_icone conforme necessário)
        aplicar_icone(self.root, "C:\\Sistema\\logos\\Kametal.ico")

        self.configurar_estilo()
        self.criar_widgets()
        self.carregar_dados_iniciais()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # reaplica estilo ao ser exibida ou ao ganhar foco — não cria função nova
        self.root.bind("<Map>", lambda e: self.configurar_estilo())
        self.root.bind("<FocusIn>", lambda e: self.configurar_estilo())

        # Chama a função que reorganiza os IDs ao inicializar a janela
        self.reiniciar_ids_estoque()

    def configurar_estilo(self):
        # usa o Style global (sem amarrar ao Toplevel) para garantir que theme_use afete tudo
        self.style = ttk.Style()
        try:
            # força o theme desejado (igual à primeira imagem)
            self.style.theme_use("alt")
        except Exception:
            pass

        # configura Treeview globalmente do jeito que você quer
        try:
            self.style.configure("Treeview",
                                background="white",
                                foreground="black",
                                rowheight=27,
                                fieldbackground="white",
                                font=("Arial", 10))
        except Exception:
            # se algum option não for suportado na plataforma, ignora
            try:
                self.style.configure("Treeview",
                                    background="white",
                                    foreground="black",
                                    rowheight=27,
                                    fieldbackground="white")
            except Exception:
                pass

        try:
            self.style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        except Exception:
            pass

        try:
            self.style.map("Treeview",
                        background=[("selected", "#0078D7")],
                        foreground=[("selected", "white")])
        except Exception:
            pass

        # garante que, se o treeview já existir, receba o estilo e altura desejada
        if hasattr(self, "treeview"):
            try:
                self.treeview.config(style="Treeview", height=15)  # volta pra 15, como antes
            except Exception:
                pass
    
    def criar_widgets(self):
        # Frame superior para entradas
        self.frame_top = ttk.Frame(self.root)
        self.frame_top.pack(padx=15, pady=(10, 0), fill="x")
    
        ttk.Label(self.frame_top, text="Produto:", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 1))
        self.entrada_produto = ttk.Combobox(self.frame_top, width=25, values=[], state="normal")
        self.entrada_produto.pack(side="left", padx=(0, 10))
    
        # Digitar Valor do Peso
        ttk.Label(self.frame_top, text="Digitar Valor do Peso:", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 1))
        self.entrada_valor = ttk.Entry(self.frame_top, width=10)
        self.entrada_valor.pack(side="left", padx=(0, 10))
        # aqui aplicamos o bind para formatação
        self.entrada_valor.bind("<KeyRelease>", lambda e: self.formatar_numero(e, 3))
    
        # Combobox para selecionar a operação
        ttk.Label(self.frame_top, text="Operação:", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 1))
        self.combo_operacao = ttk.Combobox(self.frame_top, values=["Adicionar", "Subtrair"], state="readonly", width=15)
        self.combo_operacao.set("Subtrair")
        self.combo_operacao.pack(side="left", padx=(1, 10))
    
        # Frame para os botões
        self.frame_botoes = ttk.Frame(self.root)
        self.frame_botoes.pack(padx=15, fill="x")

        # --- Botão de ajuda (correção: sem cget("background") no ttk.Frame) ---
        self.help_frame = tk.Frame(self.frame_top)   # sem bg = self.frame_top.cget("background")
        self.help_frame.pack(side="right", padx=(8, 12), pady=2)

        self.botao_ajuda_calculo = tk.Button(
            self.help_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=lambda: self._abrir_ajuda_calculo_modal()
        )
        self.botao_ajuda_calculo.bind("<Enter>", lambda e: self.botao_ajuda_calculo.config(bg="#3b5566"))
        self.botao_ajuda_calculo.bind("<Leave>", lambda e: self.botao_ajuda_calculo.config(bg="#2c3e50"))
        self.botao_ajuda_calculo.pack(side="right", padx=(4, 6), pady=4)

        # tooltip resumido (mostra pequeno texto no hover)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_calculo,
                    "Ajuda — Cálculo de NFs: F1",
                    max_width=380
                )
        except Exception:
            pass

        # também registra atalho F1 para abrir a ajuda (bind no Toplevel usado pela classe)
        try:
            self.root.bind_all("<F1>", lambda e: self._abrir_ajuda_calculo_modal())
        except Exception:
            try:
                self.root.bind("<F1>", lambda e: self._abrir_ajuda_calculo_modal())
            except Exception:
                pass
    
        # Frame para o Treeview e barras de rolagem
        self.frame_tree = ttk.Frame(self.root)
        self.frame_tree.pack(fill="both", expand=True, padx=10, pady=10)
    
        colunas = ("Data", "NF", "Produto", "Peso Líquido", "Qtd Estoque", "Qtd Cobre", "Qtd Zinco", "Qtd Sucata", "Valor Total NF", "Mão de Obra", "Matéria Prima", "Custo Manual", "Custo Total")
        self.treeview = ttk.Treeview(self.frame_tree, columns=colunas, show="headings", height=15)
        # Configuração de tags
        self.treeview.tag_configure("red_text", foreground="red")
        self.treeview.tag_configure("partial_red_text", foreground="Orange")
    
        for col in colunas:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, anchor="center", width=200)
    
        self.vsb = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.treeview.yview)
        self.hsb = ttk.Scrollbar(self.frame_tree, orient="horizontal", command=self.treeview.xview)
        self.treeview.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
    
        self.treeview.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        self.frame_tree.grid_rowconfigure(0, weight=1)
        self.frame_tree.grid_columnconfigure(0, weight=1)
    
        # Lista para armazenar itens ocultos do Treeview
        self.itens_ocultos = []
    
        # Botões
        # === Botões (criação) ===
        self.botao_atualizar = ttk.Button(
            self.frame_botoes,
            text="Calcular",
            command=self.atualizar_produto,
            style="Estoque.TButton"
        )
        self.botao_atualizar_custo_manual = ttk.Button(
            self.frame_botoes,
            text="Atualizar Custo Manual",
            command=self.atualizar_custo_manual,
            style="Estoque.TButton"
        )

        # cria (ou reutiliza) o botão de Histórico aqui
        self.botao_notificacao = ttk.Button(
            self.frame_botoes,
            text="Historico de Calculo",
            command=self.mostrar_historico,
            style="Estoque.TButton"
        )

        self.botao_voltar = ttk.Button(
            self.frame_botoes,
            text="Voltar",
            command=self.voltar_para_menu,
            style="Estoque.TButton"
        )

        # === Empacotar na ordem desejada ===
        self.botao_atualizar.pack(side="left", padx=5, pady=10)
        self.botao_atualizar_custo_manual.pack(side="left", padx=5, pady=10)
        self.botao_notificacao.pack(side="left", padx=5, pady=10)
        self.botao_voltar.pack(side="left", padx=5, pady=10)
    
        # Frame para a barra de pesquisa
        self.frame_pesquisa = ttk.Frame(self.root)
        self.frame_pesquisa.pack(padx=15, pady=(5,10), fill="x")
    
        ttk.Label(self.frame_pesquisa, text="Pesquisar:", font=("Arial", 10, "bold")).pack(side="left", padx=(10,1), pady=10)
        
        # Cria a StringVar para a pesquisa e adiciona o trace para formatação em tempo real
        self.search_var = tk.StringVar()
        self.trace_id = self.search_var.trace_add("write", self.formatar_data)
        
        # Cria o Entry usando a StringVar
        self.entrada_pesquisa = ttk.Entry(self.frame_pesquisa, width=25, textvariable=self.search_var)
        self.entrada_pesquisa.pack(side="left", padx=(0,10), pady=10)
    
        self.botao_pesquisar = ttk.Button(self.frame_pesquisa, text="Buscar", command=self.pesquisar, style="Estoque.TButton")
        self.botao_pesquisar.pack(side="left", padx=5, pady=10)
    
        self.botao_limpar_pesquisa = ttk.Button(self.frame_pesquisa, text="Limpar", command=self.limpar_pesquisa, style="Estoque.TButton")
        self.botao_limpar_pesquisa.pack(side="left", padx=5, pady=10)

        self.botao_exportar_excel = ttk.Button(self.frame_pesquisa, text="Exportar Excel", command=self.abrir_dialogo_exportacao, style="Estoque.TButton")
        self.botao_exportar_excel.pack(side="right", padx=5, pady=10)
    
        # Carrega os nomes dos produtos e associa ao combobox
        self.todos_produtos = self.carregar_nomes_produtos()
        self.entrada_produto['values'] = self.todos_produtos
        self.entrada_produto.bind('<KeyRelease>', self.atualizar_sugestoes_produto)
        self.entrada_pesquisa.bind("<Return>", lambda event: self.pesquisar())
        self.entrada_pesquisa.bind("<KeyRelease>", lambda event: self.limpar_pesquisa() if self.entrada_pesquisa.get() == "" else None)
    
        # Carrega os dados e atualiza o Treeview
        dados = self.carregar_dados()
        self.calcular_dados(dados)
        self.verificar_e_avisar_exportacao_inicio(dias_aviso=3)
        # garantir tabela de histórico criada
        self.garantir_tabela_historico()
        # garantir que a rotina de purga esteja rodando
        self.iniciar_rotina_purga_mensal()
        # purga automática silenciosa na abertura caso tenha virado o mês
        try:
            # chama o método que limpa registros de meses anteriores
            self.purga_automatica_na_abertura()
        except Exception:
            # não quer quebrar a inicialização se algo falhar aqui
            traceback.print_exc()

    def atualizar_treeview(self):
        """Função para atualizar o Treeview após a inserção de novos dados."""
        conn = conectar()
        cursor = conn.cursor()

        # Limpa o Treeview antes de recarregar os dados
        self.treeview.delete(*self.treeview.get_children())

        # Realiza a junção entre as tabelas 'somar_produtos' e 'estoque_quantidade'
        cursor.execute("""
            SELECT 
                sp.data,
                sp.nf,
                sp.produto,
                sp.peso_liquido,
                eq.quantidade_estoque,
                eq.qtd_cobre,
                eq.qtd_zinco,
                eq.qtd_sucata,
                eq.valor_total_nf,
                eq.mao_de_obra,
                eq.materia_prima,
                eq.custo_total_manual,
                eq.custo_total
            FROM 
                somar_produtos sp
            LEFT JOIN 
                estoque_quantidade eq
            ON 
                sp.id = eq.id_produto
            ORDER BY
                sp.data DESC, sp.nf::INTEGER DESC;
        """)
        resultados_atualizados = cursor.fetchall()

        for row in resultados_atualizados:
            # Formata a data para o formato dd/mm/yyyy, se aplicável
            if isinstance(row[0], datetime) or isinstance(row[0], date):
                data_formatada = row[0].strftime("%d/%m/%Y")
            else:
                data_formatada = row[0]

            # Substitui valores None por representações padrão
            nota_fiscal = row[1] if row[1] is not None else "N/A"
            produto_nome = row[2] if row[2] is not None else "N/A"
            peso_liquido = row[3] if row[3] is not None else 0
            quantidade_estoque = row[4] if row[4] is not None else peso_liquido  # Usa peso_liquido se None
            quantidade_cobre = row[5] if row[5] is not None else 0
            quantidade_zinco = row[6] if row[6] is not None else 0
            quantidade_sucata = row[7] if row[7] is not None else 0
            valor_total_nf = row[8] if row[8] is not None else 0
            mao_de_obra = row[9] if row[9] is not None else 0
            materia_prima = row[10] if row[10] is not None else 0
            custo_total_manual = row[11] if row[11] is not None else 0
            custo_total = row[12] if row[12] is not None else 0

            # Define o peso integral (pode ser ajustado conforme a lógica desejada)
            peso_integral = peso_liquido

            # Formata os valores para exibição (garante que 0 seja mostrado como "0,000", por exemplo)
            quantidade_cobre_formatada = self.formatar_valor(quantidade_cobre, 3)
            quantidade_zinco_formatada = self.formatar_valor(quantidade_zinco, 3)
            quantidade_sucata_formatada = self.formatar_valor(quantidade_sucata, 3)
            peso_liquido_com_virgula = self.formatar_valor(peso_liquido, 3)
            quantidade_estoque_com_virgula = self.formatar_valor(quantidade_estoque, 3)
            valor_total_formatado = self.formatar_valor(valor_total_nf)
            mao_de_obra_formatado = self.formatar_valor(mao_de_obra)
            materia_prima_formatado = self.formatar_valor(materia_prima)
            custo_total_manual_formatado = self.formatar_valor(custo_total_manual)
            custo_total_formatado = self.formatar_valor(custo_total)

            # Aplica a tag correta com base nas condições
            if quantidade_estoque == 0:
                tag = "red_text"
            elif quantidade_estoque != peso_liquido:
                tag = "partial_red_text"
            else:
                tag = ""

            self.treeview.insert(
                "",
                "end",
                values=(
                    data_formatada,           # Data formatada
                    nota_fiscal,              # Nota Fiscal
                    produto_nome,             # Produto
                    peso_liquido_com_virgula,  # Peso Líquido formatado
                    quantidade_estoque_com_virgula,  # Quantidade Estoque formatada
                    quantidade_cobre_formatada,      # Quantidade de Cobre formatada
                    quantidade_zinco_formatada,      # Quantidade de Zinco formatada
                    quantidade_sucata_formatada,     # Quantidade de Sucata formatada
                    valor_total_formatado,           # Valor Total NF formatado
                    mao_de_obra_formatado,           # Mão de Obra formatada
                    materia_prima_formatado,         # Matéria Prima formatada
                    custo_total_manual_formatado,    # Custo Manual formatado
                    custo_total_formatado            # Custo Total formatado
                ),
                tags=(tag,)
            )

        conn.commit()
        cursor.close()
        conn.close()

    def _create_tooltip(self, widget, text, delay=450, max_width=None):
        """
        Tooltip melhorado: quebra de linha e evita sair da tela.
        Use: self._create_tooltip(widget, 'texto longo...')
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

            # decide wraplength conforme a largura da tela
            if max_width:
                wrap_len = max(120, min(max_width, screen_w - 80))
            else:
                wrap_len = min(360, max(200, screen_w - 160))

            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)

            label = tk.Label(win, text=text, bg="#333333", fg="white",
                            font=("Segoe UI", 9), bd=0, padx=6, pady=4, wraplength=wrap_len)
            label.pack()

            # posição inicial: abaixo do widget, centrado
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajusta horizontalmente para não sair da tela
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # tenta posicionar abaixo; se não couber, posiciona acima
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

    def _abrir_ajuda_calculo_modal(self, contexto=None):
        """Modal de Ajuda — Cálculo de NFs: explica produto, peso, operação, calcular, atualizar manual, histórico, pesquisa e export."""
        try:
            modal = tk.Toplevel(self.root)
            modal.title("Ajuda — Cálculo de NFs")
            modal.transient(self.root)
            modal.grab_set()
            modal.configure(bg="white")

            # centraliza
            w, h = 920, 680
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # ícone (opcional)
            try:
                aplicar_icone(modal, "C:\\Sistema\\logos\\Kametal.ico")
            except Exception:
                pass

            # cabeçalho
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Cálculo de NFs", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # corpo (nav esquerda + conteúdo direita)
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10,6), padx=12)

            sections = [
                "Visão Geral",
                "Produto (lista suspensa)",
                "Peso (entrada)",
                "Operação (Adicionar/Subtrair)",
                "Botão Calcular (Comportamento)",
                "Atualizar Custo Manual",
                "Histórico de Cálculo",
                "Exportação do Histórico",
                "Pesquisa no Histórico",
                "Exportar Excel",
                "Barra de Pesquisa na Tela",
                "FAQ"
            ]

            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False, bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0,10))

            # area de conteúdo (text + scrollbar)
            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)

            txt = tk.Text(content_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb = tk.Scrollbar(content_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            # textos por seção
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "Esta tela permite ajustar manualmente o estoque com base nas NFs de entrada e acompanhar o histórico dessas alterações. "
                "Use os controles acima (produto, valor do peso e operação) para executar alterações rápidas; use 'Atualizar Custo Manual' "
                "para corrigir o custo de uma NF específica.\n"
            )

            contents["Produto (lista suspensa)"] = (
                "Produto — lista suspensa\n\n"
                " • A lista de produtos presente na combobox vem das NFs registradas (Entrada_NF). Ela lista todos os produtos já cadastrados nas entradas.\n"
                " • Você pode digitar parte do nome para filtrar; o sistema sugere itens existentes.\n"
                " • Se um produto novo for cadastrado na Entrada_NF/Base, reabra a tela ou atualize a lista (recarregar) para vê-lo na combobox.\n"
            )

            contents["Peso (entrada)"] = (
                "Peso — como informar\n\n"
                " • No campo de peso (entrada_valor) NÃO é necessário digitar a vírgula decimal — digite apenas dígitos e o sistema formatará automaticamente.\n"
                "   Ex.: digitar '2500' resultará em '25,00' (dependendo da formatação usada pela sua aplicação).\n"
                " • Use a caixa de seleção de operação (Adicionar/Subtrair) para informar se o valor será somado ao estoque ou subtraído dele.\n"
            )

            contents["Operação (Adicionar/Subtrair)"] = (
                "Operação — o que acontece\n\n"
                " • 'Subtrair': o sistema irá buscar as NFs mais antigas (FIFO) e subtrair a quantidade informada delas, respeitando o estoque disponível em cada NF.\n"
                "   Se o valor informado for maior que o disponível na NF atual, o sistema continuará buscando nas próximas NFs (ordem por data/número).\n\n"
                " • 'Adicionar': o sistema procura NFs mais recentes e tenta adicionar o valor informado nelas (até que o peso líquido da NF seja atingido). "
                "Se não houver espaço suficiente em NFs existentes, o sistema avisará sobre o restante que não pôde ser alocado.\n"
            )

            contents["Botão Calcular (Comportamento)"] = (
                "Botão Calcular — passo a passo\n\n"
                "1) Preencha Produto, Peso e selecione a Operação.\n"
                "2) Clique em 'Calcular'.\n"
                "   - Se for Subtrair: o sistema varre NFs (das mais antigas) subtraindo até consumir o valor informado.\n"
                "   - Se for Adicionar: o sistema varre NFs (das mais recentes) adicionando até preencher o peso líquido de cada NF.\n"
                "3) O sistema atualiza a coluna 'Qtd Estoque' das NFs alteradas e aplica as tags visuais:\n"
                "   • Amarelo: linha está em uso (estoque parcialmente utilizado)\n"
                "   • Vermelho: estoque da NF está zerado\n"
                "   • Preto: NF não foi alterada; estoque está igual ao peso líquido\n"
                "\nObservação técnica:\n"
                " • O processo tenta preservar consistência: se o valor informado ultrapassar o total disponível (no caso de subtração), o usuário recebe erro.\n"
                " • Para adicionar, o sistema tenta distribuir o valor até completar o peso_liquido das NFs (se houver espaço).\n"
            )

            contents["Atualizar Custo Manual"] = (
                "Atualizar Custo Manual — quando usar\n\n"
                " • Use este botão quando uma NF estiver com custo incorreto e você precisar corrigir manualmente o 'Custo Total Manual'.\n"
                " • Fluxo: selecione a linha (a NF específica) na Treeview, clique 'Atualizar Custo Manual', insira o valor correto e confirme. "
                "O valor será gravado na tabela e priorizado nos cálculos de custo.\n"
            )

            contents["Histórico de Cálculo"] = (
                "Histórico de Cálculo — o que é registrado\n\n"
                " • O histórico registra: usuário que executou a operação, NF de onde foi subtraído ou adicionou, produto, quantidade alterada, "
                "tipo de operação (adicionar/subtrair) e data/hora.\n"
                " • Cada ação (Calcular / Atualizar Custo Manual) que modifica estoque gera uma entrada no histórico.\n"
                " • Serve para auditoria e rastreabilidade das mudanças de estoque.\n"
            )

            contents["Exportação do Histórico"] = (
                "Exportação do Histórico — instruções\n\n"
                " • 'Exportar Visível': exporta apenas as linhas atualmente visíveis no histórico, com filtros e pesquisa aplicados.\n"
                " • 'Exportar Tudo': exporta todo o histórico do banco, independentemente dos filtros aplicados.\n"
                " • As exportações geram arquivos .xlsx com colunas: Usuário, NF, Produto, Quantidade, Operação, Data/Hora.\n"
                " • Recomenda-se exportar antes da purga mensal para manter registros de meses anteriores.\n"
            )

            contents["Pesquisa no Histórico"] = (
                "Pesquisa no Histórico — como usar\n\n"
                " • Use os filtros do histórico (produto, termo de busca) para localizar operações específicas.\n"
                " • É possível buscar por usuário, número da NF, produto, operação (adicionar/subtrair) ou parte do texto.\n"
                " • O resultado mostra as entradas ordenadas por data/hora (mais recentes primeiro)."
            )

            contents["Exportar Excel"] = (
                "Exportar Excel — como funciona\n\n"
                " • A exportação gera um arquivo .xlsx contendo as colunas visíveis (Data, NF, Produto, Quantidade, Operação, Usuário, etc.).\n"
                " • Antes de exportar, aplique os filtros desejados (data, produto) para reduzir o volume.\n"
            )

            contents["Barra de Pesquisa na Tela"] = (
                "Barra de Pesquisa (tela principal) — explicação\n\n"
                " • A barra de pesquisa filtra a Treeview principal por Data (formato dd/mm/aaaa), NF, Produto ou parte do nome.\n"
                " • Pressione Enter para executar a busca; 'Limpar' restaura a visão completa.\n"
            )

            contents["FAQ"] = (
                "FAQ — Perguntas rápidas\n\n"
                "Q: Posso digitar o peso com vírgula? \nA: Sim, mas não é necessário — o campo aceita apenas dígitos e faz a formatação.\n\n"
                "Q: O que acontece se eu adicionar mais do que o espaço disponível? \nA: O sistema tentará preencher NFs existentes; se não houver espaço, resta um valor não alocado e uma mensagem de aviso é exibida.\n\n"
                "Q: Quem aparece no histórico? \nA: O usuário obtido de `menu.user_name` (se disponível) ou o usuário do sistema (getpass.getuser())."
            )

            # função para mostrar seção
            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])

            listbox.bind("<<ListboxSelect>>", on_select)

            # rodapé com fechar
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right", padx=6)

            modal.bind("<Escape>", lambda e: modal.destroy())
            modal.focus_set()
            modal.wait_window()

        except Exception as e:
            print("Erro ao abrir modal de ajuda (Cálculo de NFs):", e)

    def formatar_valor(self, valor, casas_decimais=2):
        return f"{valor:,.{casas_decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    def formatar_numero(self, event, casas):
        entry = event.widget
        texto = entry.get()
        texto_limpo = ''.join(filter(str.isdigit, texto))
        if not texto_limpo:
            entry.delete(0, tk.END)
            return
        if len(texto_limpo) <= casas:
            texto_limpo = texto_limpo.rjust(casas+1, '0')
        inteiro = texto_limpo[:-casas]
        decimal = texto_limpo[-casas:]
        part_int = f"{int(inteiro):,}".replace(",", ".")
        valor_formatado = f"{part_int},{decimal}"
        entry.delete(0, tk.END)
        entry.insert(0, valor_formatado)

    def carregar_nomes_produtos(self):
        """Carrega os nomes dos produtos do banco de dados."""
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT produto FROM somar_produtos ORDER BY produto")
        produtos = [row[0] for row in cursor.fetchall()]
        cursor.close()
        conn.close()
        return produtos

    def atualizar_sugestoes_produto(self, event):
        """Atualiza as sugestões na Combobox com base no texto digitado."""
        texto_digitado = self.entrada_produto.get().lower()
        sugestoes = [produto for produto in self.todos_produtos if texto_digitado in produto.lower()]
        
        # Atualiza os valores da Combobox
        self.entrada_produto['values'] = sugestoes

        # Mostra o menu suspenso com sugestões
        self.entrada_produto.event_generate('<Down>')

    def remover_acentos(self, texto):
        return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

    def carregar_dados(self):
        conn = conectar()
        cursor = conn.cursor()
        query = """
            SELECT sp.data, sp.nf, sp.produto, sp.peso_liquido, sp.peso_integral, 
                eq.quantidade_estoque,
                sp.material_1, sp.fornecedor
            FROM somar_produtos sp
            LEFT JOIN estoque_quantidade eq ON sp.id = eq.id_produto
            ORDER BY sp.data ASC, sp.nf ASC;
        """
        cursor.execute(query)
        resultados = cursor.fetchall()
        cursor.close()
        conn.close()
        return resultados

    def carregar_dados_iniciais(self):
        self.todos_produtos = self.carregar_nomes_produtos()
        self.entrada_produto['values'] = self.todos_produtos
        self.atualizar_treeview()
        # Atualiza automaticamente o custo total sem intervenção do usuário
        self.atualizar_custo_total()

    def calcular_dados(self, resultados):
            """
            Processa os resultados, insere os dados gerais do estoque no Treeview e, em seguida, 
            chama os métodos complementares.
            """
            # Conecta ao banco e obtém um cursor
            conn = conectar()
            cursor = conn.cursor()

            # Limpa o Treeview
            for item in self.treeview.get_children():
                self.treeview.delete(item)

            # Processa cada linha dos resultados
            for row in resultados:
                # Inicializa as variáveis de metais e valor total da NF
                quantidade_cobre = None
                quantidade_zinco = None
                quantidade_sucata = None
                valor_total_nf = 0.0

                # Formata a data (índice 0)
                if isinstance(row[0], (datetime, date)):
                    data_formatada = row[0].strftime("%d/%m/%Y")
                else:
                    data_formatada = row[0]

                # Obtém os demais valores (ajuste os índices conforme sua consulta)
                nota_fiscal = row[1] if row[1] is not None else "N/A"
                produto_nome = row[2] if row[2] is not None else "N/A"
                peso_liquido = row[3] if row[3] is not None else 0
                peso_integral = row[4] if row[4] is not None else 0
                quantidade_estoque = row[5] if row[5] is not None else peso_liquido
                # Campos que originalmente já vieram do banco (cobre, zinco, etc.) serão substituídos
                # pelos valores calculados a seguir.
                fornecedor_nome = row[7]  # Supondo que o fornecedor esteja na coluna 7

                # Normaliza o nome do produto
                produto_nome_normalizado = produto_nome.strip().upper()
                produto_nome_normalizado = self.remover_acentos(produto_nome_normalizado)
                produto_nome_normalizado = produto_nome_normalizado.replace('~', '').replace('(', '').replace(')', '').replace(' ', '')
                produto_nome_normalizado = produto_nome_normalizado.replace('\r', '').replace('\n', '').replace('\t', '')

                # Verifica o produto na tabela somar_produtos com filtro pela NF
                cursor.execute("""
                    SELECT id, material_1, material_2, material_3, material_4, material_5
                    FROM somar_produtos
                    WHERE UPPER(REPLACE(REPLACE(REPLACE(REPLACE(produto, ' ', ''), '(', ''), ')', ''), '~', '')) = %s
                    AND nf = %s
                """, (produto_nome_normalizado, nota_fiscal))
                produto_result = cursor.fetchone()

                if produto_result:
                    id_produto = produto_result[0]
                    materiais_produto = produto_result[1:]

                    # Obtém os percentuais de cobre e zinco do produto
                    cursor.execute("""
                        SELECT percentual_cobre, percentual_zinco 
                        FROM produtos 
                        WHERE UPPER(REPLACE(REPLACE(REPLACE(REPLACE(nome, ' ', ''), '(', ''), ')', ''), '~', '')) = %s
                    """, (produto_nome_normalizado,))
                    produto_info = cursor.fetchone()
                    percentual_cobre = produto_info[0] / 100 if produto_info else None
                    percentual_zinco = produto_info[1] / 100 if produto_info else None

                    # Busca o valor total da NF na tabela estoque_quantidade
                    cursor.execute("""
                        SELECT eq.valor_total_nf
                        FROM estoque_quantidade eq
                        JOIN somar_produtos sp ON eq.id_produto = sp.id
                        WHERE sp.nf = %s
                    """, (nota_fiscal,))
                    valor_total_nf_result = cursor.fetchone()
                    if valor_total_nf_result:
                        valor_total_nf = valor_total_nf_result[0]
                    else:
                        valor_total_nf = 0.0

                    # Calcula os valores de metais separadamente
                    for material in materiais_produto:
                        if not material:
                            continue
                        material_normalizado = material.strip().upper()
                        cursor.execute("""
                            SELECT grupo, valor 
                            FROM materiais
                            WHERE UPPER(TRIM(nome)) = %s AND UPPER(TRIM(fornecedor)) = %s
                        """, (material_normalizado, fornecedor_nome.strip().upper()))
                        material_info = cursor.fetchone()
                        if material_info:
                            grupo = material_info[0].strip().lower()
                            valor_material = material_info[1]
                            if peso_liquido is not None and peso_integral is not None and valor_material:
                                if percentual_cobre is not None and grupo == "cobre":
                                    quantidade_cobre = ((peso_liquido - peso_integral) * percentual_cobre) / valor_material
                                elif percentual_zinco is not None and grupo == "zinco":
                                    quantidade_zinco = ((peso_liquido - peso_integral) * percentual_zinco) / valor_material
                                elif grupo == "sucata":
                                    try:
                                        quantidade_sucata = (peso_liquido - peso_integral) / valor_material
                                    except ZeroDivisionError:
                                        quantidade_sucata = None

                    # Atualiza a tabela estoque_quantidade para este produto
                    if id_produto is not None:
                        cursor.execute("""
                            INSERT INTO estoque_quantidade (id_produto, qtd_cobre, qtd_zinco, qtd_sucata, quantidade_estoque)
                            VALUES (%s, %s, %s, %s, %s)
                            ON CONFLICT (id_produto) DO UPDATE
                            SET qtd_cobre = EXCLUDED.qtd_cobre,
                                qtd_zinco = EXCLUDED.qtd_zinco,
                                qtd_sucata = EXCLUDED.qtd_sucata
                            WHERE estoque_quantidade.id_produto = EXCLUDED.id_produto
                        """, (id_produto, quantidade_cobre, quantidade_zinco, quantidade_sucata, peso_liquido))
                
            conn.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conn.commit()

            cursor.close()
            conn.close()

            # Após atualizar os dados do banco, chama os métodos complementares
            self.executar_calculo()
            self.mao_de_obra()
            self.custo_total()
            self.materia_prima()
            self.atualizar_treeview()

    def atualizar_produto(self):
        produto_nome = self.entrada_produto.get()
        valor_entrada = self.entrada_valor.get()
        operacao = self.combo_operacao.get()  # Captura a operação selecionada

        if not produto_nome or not valor_entrada or not operacao:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos!")
            return

        try:
            # Remove separadores de milhar (pontos) e substitui vírgulas por pontos
            valor_entrada = Decimal(valor_entrada.replace('.', '').replace(',', '.').strip())
        except InvalidOperation:
            messagebox.showerror("Erro", "O valor deve ser numérico e válido!")
            return

        conn = conectar()
        cursor = conn.cursor()

        # Seleciona os produtos a serem atualizados, ordenando por data e número da NF
        query = """
        SELECT sp.id, sp.nf, sp.produto, sp.peso_liquido, COALESCE(eq.quantidade_estoque, sp.peso_liquido)
        FROM somar_produtos sp
        LEFT JOIN estoque_quantidade eq ON sp.id = eq.id_produto
        WHERE sp.produto = %s
        ORDER BY sp.data ASC, sp.nf ASC
        """
        cursor.execute(query, (produto_nome,))
        resultados = cursor.fetchall()

        if not resultados:
            messagebox.showerror("Erro", "Produto não encontrado!")
            cursor.close()
            conn.close()
            return

        valor_restante = valor_entrada
        atualizacoes = []  # elementos: (id_produto, nf, nova_quantidade, quantidade_anterior)

        # Mais seguro: converte valores lidos
        try:
            resultados_conv = []
            for r in resultados:
                id_produto, nf, prod, peso_liquido, quantidade_estoque = r
                quantidade_estoque = Decimal(quantidade_estoque)
                resultados_conv.append((id_produto, nf, prod, peso_liquido, quantidade_estoque))
        except Exception:
            # se algo der errado, usa resultados originais (mas convertendo na hora)
            resultados_conv = resultados

        if operacao == "Subtrair":
            total_disponivel = sum(Decimal(r[4]) for r in resultados_conv)

            if valor_entrada > total_disponivel:
                messagebox.showerror(
                    "Erro",
                    f"O valor de entrada ({valor_entrada:.3f}) ultrapassa o total disponível nas notas fiscais ({total_disponivel:.3f})!"
                )
                cursor.close()
                conn.close()
                return

            for resultado in resultados_conv:
                id_produto, nf, prod, peso_liquido, quantidade_estoque = resultado
                quantidade_estoque = Decimal(quantidade_estoque)

                if valor_restante <= quantidade_estoque:
                    nova_quantidade_estoque = quantidade_estoque - valor_restante
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante = Decimal("0")
                    break
                else:
                    # zera esse registro e continua subtraindo do próximo
                    espaco_disponivel = quantidade_estoque
                    nova_quantidade_estoque = quantidade_estoque - espaco_disponivel  # = 0
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante -= espaco_disponivel

        elif operacao == "Adicionar":
            # percorre do mais novo para o mais antigo (LIFO)
            for resultado in reversed(resultados_conv):
                id_produto, nf, prod, peso_liquido, quantidade_estoque = resultado

                # conversões seguras (evita virar 0 por erro de conversão)
                try:
                    peso_liquido_dec = Decimal(str(peso_liquido)) if peso_liquido is not None else Decimal("0")
                except Exception:
                    peso_liquido_dec = Decimal("0")

                quantidade_estoque = Decimal(str(quantidade_estoque))

                espaco_disponivel = peso_liquido_dec - quantidade_estoque
                if espaco_disponivel <= 0:
                    # NF já está cheia, pula
                    continue

                if valor_restante <= espaco_disponivel:
                    nova_quantidade_estoque = quantidade_estoque + valor_restante
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante = Decimal("0")
                    break
                else:
                    # enche essa NF até o limite e segue para a anterior
                    nova_quantidade_estoque = peso_liquido_dec  # equivale a quantidade_estoque + espaco_disponivel
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante -= espaco_disponivel

            if valor_restante > 0:
                messagebox.showwarning("Aviso", f"Não foi possível adicionar todo o valor! Restante: {valor_restante:.3f}")

        # Atualiza a tabela no banco
        try:
            for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
                insert_query = """
                INSERT INTO estoque_quantidade (id_produto, quantidade_estoque)
                VALUES (%s, %s)
                ON CONFLICT (id_produto)
                DO UPDATE SET quantidade_estoque = EXCLUDED.quantidade_estoque
                """
                cursor.execute(insert_query, (id_produto, nova_quantidade))
            conn.commit()
        except Exception as e:
            # se der erro no update, desfaz e mostra mensagem
            conn.rollback()
            messagebox.showerror("Erro", f"Erro ao atualizar o banco: {e}")
            cursor.close()
            conn.close()
            return

        # Notifica, commit já feito acima
        try:
            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conn.commit()
        except Exception:
            # não crítico — prossegue
            pass

        # Atualiza o Treeview para refletir as alterações
        for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
            for item in self.treeview.get_children():
                valores = self.treeview.item(item, "values")
                # valores[1] = NF (coluna 1), valores[4] = Qtd Estoque (coluna 4)
                if str(valores[1]) == str(nf):
                    # Formata quantidade nova para exibição (vírgula)
                    try:
                        quantidade_formatada = f"{Decimal(nova_quantidade):.3f}".replace('.', ',')
                    except Exception:
                        quantidade_formatada = str(nova_quantidade)
                    # Mantemos as colunas anteriores, atualizamos a coluna Qtd Estoque (índice 4)
                    novos_valores = list(valores)
                    # garante que exista índice 4
                    if len(novos_valores) > 4:
                        novos_valores[4] = quantidade_formatada
                    else:
                        # fallback: amplia tupla até o índice 4
                        while len(novos_valores) <= 4:
                            novos_valores.append("")
                        novos_valores[4] = quantidade_formatada

                    # ajusta tag visual
                    try:
                        quantidade_decimal_tree = Decimal(str(quantidade_anterior))
                    except Exception:
                        quantidade_decimal_tree = quantidade_anterior

                    try:
                        nova_q_dec = Decimal(str(nova_quantidade))
                    except Exception:
                        nova_q_dec = nova_quantidade

                    if nova_q_dec == 0:
                        self.treeview.item(item, tags="red_text")
                    elif Decimal(str(nova_quantidade)) != Decimal(str(quantidade_decimal_tree)):
                        self.treeview.item(item, tags="partial_red_text")
                    else:
                        self.treeview.item(item, tags="")

                    self.treeview.item(item, values=tuple(novos_valores))

                    # ✅ Seleciona e dá foco na linha modificada
                    self.treeview.selection_set(item)
                    self.treeview.focus(item)
                    self.treeview.see(item)

                    break

        # <<< registra no histórico usando os valores corretos (quantidade_anterior) >>>
        usuario_para_registro = getattr(self, "usuario", None) or ""
        for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
            try:
                qt_anterior = Decimal(str(quantidade_anterior))
                qt_nova = Decimal(str(nova_quantidade))
            except Exception:
                try:
                    qt_anterior = Decimal(quantidade_anterior)
                    qt_nova = Decimal(nova_quantidade)
                except Exception:
                    qt_anterior = Decimal("0")
                    qt_nova = Decimal("0")

            if operacao == "Subtrair":
                quantidade_sub = qt_anterior - qt_nova
                if quantidade_sub > 0:
                    # tipo correto: "subtrair"
                    self.registrar_acao_historico(nf, produto_nome, quantidade_sub, "subtrair", usuario=usuario_para_registro)

            elif operacao == "Adicionar":
                quantidade_add = qt_nova - qt_anterior
                if quantidade_add > 0:
                    # tipo correto: "adicionar"
                    self.registrar_acao_historico(nf, produto_nome, quantidade_add, "adicionar", usuario=usuario_para_registro)

        # Mostra mensagem de sucesso só se teve alguma atualização
        if atualizacoes:
            messagebox.showinfo("Sucesso", "Cálculo realizado com sucesso!")

        # Fecha cursores/conexão
        try:
            cursor.close()
        except Exception:
            pass
        try:
            conn.close()
        except Exception:
            pass

    def calcular_valor_nf(self):
        """
        Calcula o valor total da nota fiscal e atualiza a tabela estoque_quantidade na coluna valor_total_nf.
        Fórmula:
        ((peso_integral x valor_integral)+ipi) +
        (qtd_cobre x valor_unitario_1) +
        (qtd_zinco x valor_unitario_2) +
        (qtd_sucata x valor_unitario_3) +
        ((peso_liquido - peso_integral) x valor_mao_obra_tm_metallica) +
        ((peso_liquido - peso_integral) x valor_unitario_energia)
        """
        try:
            conn = conectar()
            cursor = conn.cursor()

            # Recupera os dados necessários para o cálculo
            cursor.execute("""
                SELECT 
                    eq.id_produto,
                    sp.peso_integral,
                    sp.valor_integral,
                    sp.ipi,
                    eq.qtd_cobre,
                    sp.valor_unitario_1,
                    eq.qtd_zinco,
                    sp.valor_unitario_2,
                    eq.qtd_sucata,
                    sp.valor_unitario_3,
                    sp.peso_liquido,
                    sp.valor_mao_obra_tm_metallica,
                    sp.valor_unitario_energia
                FROM estoque_quantidade eq
                INNER JOIN somar_produtos sp ON eq.id_produto = sp.id
            """)
            resultados = cursor.fetchall()

            for row in resultados:
                (
                    id_produto,
                    peso_integral,
                    valor_integral,
                    ipi,
                    qtd_cobre,
                    valor_unitario_1,
                    qtd_zinco,
                    valor_unitario_2,
                    qtd_sucata,
                    valor_unitario_3,
                    peso_liquido,
                    valor_mao_obra_tm_metallica,
                    valor_unitario_energia
                ) = row

                # Função interna para tratar valores vazios ou None
                def tratar_valor(valor):
                    return float(valor) if valor and str(valor).strip() else 0

                peso_integral = tratar_valor(peso_integral)
                valor_integral = tratar_valor(valor_integral)
                ipi = tratar_valor(ipi)
                qtd_cobre = tratar_valor(qtd_cobre)
                valor_unitario_1 = tratar_valor(valor_unitario_1)
                qtd_zinco = tratar_valor(qtd_zinco)
                valor_unitario_2 = tratar_valor(valor_unitario_2)
                qtd_sucata = tratar_valor(qtd_sucata)
                valor_unitario_3 = tratar_valor(valor_unitario_3)
                peso_liquido = tratar_valor(peso_liquido)
                valor_mao_obra_tm_metallica = tratar_valor(valor_mao_obra_tm_metallica)
                valor_unitario_energia = tratar_valor(valor_unitario_energia)

                # Evita subtração negativa para casos onde peso_liquido < peso_integral
                diferenca_peso = max(0, peso_liquido - peso_integral)

                # Aplica o IPI como porcentagem
                valor_integral_com_ipi = valor_integral * (1 + ipi / 100)

                valor_total_nf = (
                    (peso_integral * valor_integral_com_ipi) +
                    (qtd_cobre * valor_unitario_1) +
                    (qtd_zinco * valor_unitario_2) +
                    (qtd_sucata * valor_unitario_3) +
                    (diferenca_peso * valor_mao_obra_tm_metallica) +
                    (diferenca_peso * valor_unitario_energia)
                )

                # Atualiza o valor total da nota fiscal na tabela estoque_quantidade
                cursor.execute("""
                    UPDATE estoque_quantidade
                    SET valor_total_nf = %s
                    WHERE id_produto = %s
                """, (valor_total_nf, id_produto))

            conn.commit()

        except psycopg2.Error as e:
            print(f"Erro ao calcular os valores: {e}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def executar_calculo(self):
        """Executa o cálculo do valor da nota fiscal e atualiza os valores."""
        try:
            self.calcular_valor_nf()
            print("Cálculos realizados e valores atualizados com sucesso!")
        except Exception as e:
            print(f"Erro ao calcular os valores: {e}")

    def mao_de_obra(self):
        """
        Calcula o valor de mão de obra e atualiza a tabela estoque_quantidade na coluna mao_de_obra.
        Se 'valor_integral' estiver preenchido (diferente de 0), utiliza:
            valor_mao_de_obra = valor_unitario_energia + valor_mao_obra_tm_metallica
        Caso contrário, realiza o cálculo alternativo:
            valor_mao_de_obra = (soma das duplicatas) / peso_liquido
        """
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT 
                sp.id, 
                sp.valor_integral,
                sp.valor_unitario_energia, 
                sp.valor_mao_obra_tm_metallica, 
                sp.peso_liquido,
                sp.duplicata_1, 
                sp.duplicata_2, 
                sp.duplicata_3, 
                sp.duplicata_4, 
                sp.duplicata_5, 
                sp.duplicata_6
            FROM 
                somar_produtos sp
        """)
        produtos = cursor.fetchall()

        for produto in produtos:
            id_produto = produto[0]
            valor_integral = produto[1] if produto[1] is not None else 0
            valor_unitario_energia = produto[2] if produto[2] is not None else 0
            valor_mao_obra_tm_metallica = produto[3] if produto[3] is not None else 0
            peso_liquido = produto[4] if produto[4] is not None else 0
            duplicatas = produto[5:11]

            if valor_integral != 0:
                valor_mao_de_obra = valor_unitario_energia + valor_mao_obra_tm_metallica
            else:
                if peso_liquido > 0:
                    duplicatas_corrigidas = [d if d is not None else 0 for d in duplicatas]
                    soma_duplicatas = sum(duplicatas_corrigidas)
                    valor_mao_de_obra = soma_duplicatas / peso_liquido
                else:
                    valor_mao_de_obra = 0

            cursor.execute("""
                UPDATE estoque_quantidade
                SET mao_de_obra = %s
                WHERE id_produto = %s
            """, (valor_mao_de_obra, id_produto))

        conn.commit()
        cursor.close()
        conn.close()

        self.atualizar_treeview()

    def materia_prima(self):
        """
        Calcula a matéria-prima com base na diferença entre custo_total e mao_de_obra,
        ou realiza um cálculo alternativo se 'valor_integral' estiver em branco ou for 0.
        Atualiza a coluna materia_prima na tabela estoque_quantidade.
        """
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                eq.id_produto,
                eq.custo_total,
                eq.mao_de_obra,
                sp.valor_integral,
                sp.valor_unitario_1,
                eq.qtd_cobre,
                sp.valor_unitario_2,
                eq.qtd_zinco,
                sp.valor_unitario_3,
                eq.qtd_sucata,
                sp.peso_liquido
            FROM estoque_quantidade eq
            LEFT JOIN somar_produtos sp ON eq.id_produto = sp.id
        """)
        resultados = cursor.fetchall()

        for row in resultados:
            (id_produto, custo_total, mao_de_obra, valor_integral,
             valor_unitario_1, qtd_cobre, valor_unitario_2, qtd_zinco,
             valor_unitario_3, qtd_sucata, peso_liquido) = row

            if valor_integral not in (None, 0):
                if custo_total is not None and mao_de_obra is not None:
                    materia_prima_valor = custo_total - mao_de_obra
                else:
                    materia_prima_valor = 0
            else:
                valor_unitario_1 = valor_unitario_1 or 0
                qtd_cobre = qtd_cobre or 0
                valor_unitario_2 = valor_unitario_2 or 0
                qtd_zinco = qtd_zinco or 0
                valor_unitario_3 = valor_unitario_3 or 0
                qtd_sucata = qtd_sucata or 0
                peso_liquido = peso_liquido or 1

                materia_prima_valor = (
                    (valor_unitario_1 * qtd_cobre) +
                    (valor_unitario_2 * qtd_zinco) +
                    (valor_unitario_3 * qtd_sucata)
                ) / peso_liquido

            cursor.execute("""
                UPDATE estoque_quantidade
                SET materia_prima = %s
                WHERE id_produto = %s
            """, (materia_prima_valor, id_produto))

            # Atualiza o Treeview para refletir a mudança (exemplo: adiciona o valor no final dos valores existentes)
            for item in self.treeview.get_children():
                values = self.treeview.item(item, "values")
                if str(values[1]) == str(id_produto):  # Ajuste conforme a posição do id_produto
                    novos_valores = values + (materia_prima_valor,)
                    self.treeview.item(item, values=novos_valores)
                    break

        conn.commit()
        cursor.close()
        conn.close()

        self.atualizar_treeview()

    def atualizar_custo_manual(self):
        # Verifica se o usuário selecionou alguma linha
        selected_item = self.treeview.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione uma linha para atualizar o custo total manual.")
            return

        # Obtém a janela principal a partir do treeview
        master_widget = self.treeview.winfo_toplevel()

        # Cria uma janela Toplevel customizada para entrada de valor
        dialog = tk.Toplevel(master_widget)
        dialog.title("Entrada de valor")
        dialog.resizable(False, False)
        dialog.transient(master_widget)  # Mantém a janela sempre acima da janela pai
        dialog.grab_set()                # Torna a janela modal

        # Aplica o ícone (certifique-se de que o caminho esteja correto)
        aplicar_icone(dialog, "C:\\Sistema\\logos\\Kametal.ico")

        # Centraliza a janela em relação ao pai
        master_widget.update_idletasks()
        width, height = 300, 200
        x = master_widget.winfo_rootx() + (master_widget.winfo_width() // 2) - (width // 2)
        y = master_widget.winfo_rooty() + (master_widget.winfo_height() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        # Configuração dos estilos (mesmo usados na função alterar_produto)
        style = ttk.Style(dialog)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TEntry", padding=5, relief="solid", font=("Arial", 10))
        style.configure("Alter.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Alter.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        # Cria um frame principal para organizar os widgets na janela de entrada
        frame = ttk.Frame(dialog, style="Custom.TFrame", padding=10)
        frame.pack(expand=True, fill="both")

        # Label com instrução
        label = ttk.Label(frame, text="Digite o valor para Custo Total Manual:", style="Custom.TLabel")
        label.pack(padx=20, pady=(20, 5))

        # Campo de entrada com estilo personalizado
        entry = ttk.Entry(frame, style="Custom.TEntry")
        entry.pack(padx=20, pady=5)
        entry.focus_set()

        # **1. Bind para formatação automática com 2 casas decimais**
        entry.bind("<KeyRelease>", lambda e: self.formatar_numero(e, 2))


        # Container para armazenar o valor de retorno
        result = {"value": None}

        def ok():
            result["value"] = entry.get()
            dialog.destroy()

        def cancelar():
            dialog.destroy()

        # Frame para os botões
        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.pack(padx=20, pady=(10, 20))
        ok_button = ttk.Button(button_frame, text="OK", command=ok, style="Alter.TButton")
        ok_button.pack(side="left", padx=5)
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=cancelar, style="Alter.TButton")
        cancel_button.pack(side="left", padx=5)

        # Eventos de teclado para facilitar o uso
        dialog.bind("<Return>", lambda event: ok())
        dialog.bind("<Escape>", lambda event: cancelar())

        # Aguarda o usuário fechar o diálogo
        master_widget.wait_window(dialog)
        valor_custo_manual = result["value"]

        # Se o usuário cancelar, retorna da função
        if valor_custo_manual is None:
            return

        # Substitui vírgula por ponto, se necessário, e converte para float
        valor_custo_manual = valor_custo_manual.replace(",", ".")
        try:
            valor_custo_manual = round(float(valor_custo_manual), 2)
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido. Por favor, insira um número válido.")
            return

        # Conecta-se ao banco de dados e atualiza os valores
        conn = conectar()
        cursor = conn.cursor()

        for item in selected_item:
            values = list(self.treeview.item(item)["values"])

            if len(values) > 12:
                values[11] = f"{valor_custo_manual:.2f}"
                self.treeview.item(item, values=values)

                produto_nome = values[2]
                nota_fiscal = values[1]

                cursor.execute("""
                    SELECT id_produto 
                    FROM estoque_quantidade eq
                    JOIN somar_produtos sp ON sp.id = eq.id_produto
                    WHERE CAST(sp.nf AS TEXT) = %s AND sp.produto = %s
                """, (str(nota_fiscal), produto_nome))
                id_produto = cursor.fetchone()

                if id_produto:
                    id_produto = id_produto[0]
                    cursor.execute("""
                        UPDATE estoque_quantidade
                        SET custo_total_manual = %s
                        WHERE id_produto = %s
                    """, (valor_custo_manual, id_produto))
                    conn.commit()
                else:
                    messagebox.showerror("Erro", "Produto não encontrado no banco de dados.")
            else:
                messagebox.showwarning("Aviso", "A linha selecionada não contém dados suficientes para atualizar o custo.")

        cursor.close()
        conn.close()

        self.atualizar_treeview()
        self.custo_total()

    def custo_total(self):
            """
            Calcula o custo total priorizando as regras:
            - Se houver um custo manual definido (maior que 0), usa-o.
            - Caso contrário, se 'valor_integral' estiver vazio ou 0, utiliza: mao_de_obra + materia_prima.
            - Caso 'valor_integral' possua valor, utiliza: valor_total_nf / peso_liquido.
            """
            conn = conectar()
            cursor = conn.cursor()

            cursor.execute("""
                SELECT 
                    eq.id_produto,
                    eq.mao_de_obra,
                    eq.materia_prima,
                    eq.valor_total_nf,
                    sp.peso_liquido,
                    sp.valor_integral,
                    eq.custo_total_manual
                FROM estoque_quantidade eq
                INNER JOIN somar_produtos sp ON eq.id_produto = sp.id
            """)
            resultados = cursor.fetchall()

            for row in resultados:
                id_produto, mao_de_obra, materia_prima, valor_total_nf, peso_liquido, valor_integral, custo_manual = row

                # Substitui valores None por 0
                mao_de_obra = mao_de_obra or 0
                materia_prima = materia_prima or 0

                # Se existir um custo manual (maior que 0), ele tem prioridade.
                if custo_manual and custo_manual > 0:
                    custo_total = custo_manual
                elif valor_integral is None or valor_integral == 0:
                    custo_total = mao_de_obra + materia_prima
                elif peso_liquido and peso_liquido != 0:
                    custo_total = valor_total_nf / peso_liquido
                else:
                    custo_total = 0

                cursor.execute("""
                    UPDATE estoque_quantidade
                    SET custo_total = %s
                    WHERE id_produto = %s
                """, (custo_total, id_produto))

            conn.commit()
            self.atualizar_treeview()
            cursor.close()
            conn.close()

    def atualizar_custo_total(self):
        """
        Atualiza o custo total chamando o método custo_total.
        Se exibir_mensagem for True, exibe uma mensagem de sucesso.
        """
        self.custo_total()

    def reordenar_treeview(self):
        """
        Reordena os itens do Treeview com base na data (coluna 0) e no número da NF (coluna 1).
        """
        # Obter todos os itens atualmente exibidos no treeview
        items = self.treeview.get_children('')
        itens_com_chave = []
        
        for item in items:
            valores = self.treeview.item(item, "values")
            # Supondo que a data esteja na coluna 0 no formato "dd/mm/yyyy"
            # e o número da NF esteja na coluna 1
            try:
                data_obj = datetime.strptime(valores[0], "%d/%m/%Y")
            except Exception:
                data_obj = datetime.min  # Se a data for inválida, posiciona no início
            try:
                nf_num = int(valores[1])
            except Exception:
                nf_num = 0  # Valor padrão caso não seja numérico
            itens_com_chave.append((item, data_obj, nf_num))
        
        # Ordena por data e, se as datas forem iguais, pelo número da NF
        itens_com_chave.sort(key=lambda x: (x[1], x[2]))
        
        # Reposiciona os itens no treeview na ordem correta
        for indice, (item, _, _) in enumerate(itens_com_chave):
            self.treeview.move(item, '', indice)

    def abrir_dialogo_exportacao(self):
        """
        Abre uma janela de diálogo para que o usuário informe os filtros de exportação.
        Os filtros serão: Data (intervalo), NF e Produto.
        """
        dialogo = tk.Toplevel(self.root)  # Vinculado à janela principal
        dialogo.title("Exportar NF - Filtros")
        dialogo.geometry("400x250")  # Mesmo tamanho do exemplo completo
        dialogo.resizable(False, False)
        centralizar_janela(dialogo, 400, 250)
        aplicar_icone(dialogo, "C:\\Sistema\\logos\\Kametal.ico")
        dialogo.config(bg="#ecf0f1")

        style = ttk.Style(dialogo)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        frame = ttk.Frame(dialogo, padding="15 15 15 15", style="Custom.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        dialogo.columnconfigure(0, weight=1)
        dialogo.rowconfigure(0, weight=1)

        # Função para formatar data diretamente nos campos de entrada
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

        # Linha 0: Data Inicial
        ttk.Label(frame, text="Data Inicial (dd/mm/yyyy):", style="Custom.TLabel").grid(row=0, column=0, sticky="e", padx=5, pady=(5,2))
        entry_data_inicial = ttk.Entry(frame, width=25)
        entry_data_inicial.grid(row=0, column=1, sticky="w", padx=5, pady=(5,2))
        entry_data_inicial.bind("<KeyRelease>", formatar_data)  # Aplica a formatação enquanto digita

        # Linha 1: Data Final
        ttk.Label(frame, text="Data Final (dd/mm/yyyy):", style="Custom.TLabel").grid(row=1, column=0, sticky="e", padx=5, pady=(5,2))
        entry_data_final = ttk.Entry(frame, width=25)
        entry_data_final.grid(row=1, column=1, sticky="w", padx=5, pady=(5,2))
        entry_data_final.bind("<KeyRelease>", formatar_data)  # Aplica a formatação enquanto digita

        # Linha 2: NF
        ttk.Label(frame, text="NF:", style="Custom.TLabel").grid(row=2, column=0, sticky="e", padx=5, pady=(5,2))
        entry_nf = ttk.Entry(frame, width=25)
        entry_nf.grid(row=2, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 3: Produto
        ttk.Label(frame, text="Produto:", style="Custom.TLabel").grid(row=3, column=0, sticky="e", padx=5, pady=(5,2))
        entry_produto = ttk.Entry(frame, width=25)
        entry_produto.grid(row=3, column=1, sticky="w", padx=5, pady=(5,2))

        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.grid(row=4, column=0, columnspan=2, pady=(15,5))

        def acao_exportar():
            filtro_data_inicial = entry_data_inicial.get().strip()
            filtro_data_final = entry_data_final.get().strip()
            filtro_nf = entry_nf.get().strip()
            filtro_produto = entry_produto.get().strip()

            # Processamento das datas:
            try:
                if filtro_data_inicial:
                    filtro_data_inicial = datetime.strptime(filtro_data_inicial, '%d/%m/%Y').strftime('%Y-%m-%d')
                if filtro_data_final:
                    filtro_data_final = datetime.strptime(filtro_data_final, '%d/%m/%Y').strftime('%Y-%m-%d')
            except ValueError:
                messagebox.showerror("Erro", "Formato de data inválido. Utilize dd/mm/yyyy.")
                return

            # Chama a função de exportação passando os filtros
            self.exportar_excel_filtrado_com_valores(
                filtro_data_inicial, filtro_data_final, filtro_nf, filtro_produto
            )
            dialogo.destroy()

        export_button = ttk.Button(button_frame, text="Exportar Excel", command=acao_exportar)
        export_button.grid(row=0, column=0, padx=5)
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=dialogo.destroy)
        cancel_button.grid(row=0, column=1, padx=5)

        frame.columnconfigure(1, weight=1)

    def exportar_excel_filtrado_com_valores(self, filtro_data_inicial, filtro_data_final, filtro_nf, filtro_produto):

        where_clauses = []
        parametros = []

        # Filtros de Data
        if filtro_data_inicial and filtro_data_final:
            where_clauses.append("s.data BETWEEN %s AND %s")
            parametros.extend([filtro_data_inicial, filtro_data_final])
        elif filtro_data_inicial:
            where_clauses.append("s.data >= %s")
            parametros.append(filtro_data_inicial)
        elif filtro_data_final:
            where_clauses.append("s.data <= %s")
            parametros.append(filtro_data_final)

        # Filtros NF e Produto
        if filtro_nf:
            where_clauses.append("s.nf ILIKE %s")
            parametros.append(f"%{filtro_nf}%")

        if filtro_produto:
            where_clauses.append("unaccent(s.produto) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_produto}%")

        where_sql = "WHERE " + " AND ".join(where_clauses) if where_clauses else ""

        query = f"""
            SELECT s.data, s.nf, s.produto, s.peso_liquido,
                eq.quantidade_estoque, eq.qtd_cobre, eq.qtd_zinco, eq.qtd_sucata,
                eq.valor_total_nf, eq.mao_de_obra, eq.materia_prima,
                eq.custo_total_manual, eq.custo_total
            FROM estoque_quantidade eq
            JOIN somar_produtos s ON eq.id_produto = s.id
            {where_sql}
            ORDER BY s.data ASC, s.nf ASC;
        """

        try:
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute(query, parametros)
            dados = cursor.fetchall()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na consulta: {e}")
            return
        finally:
            cursor.close()
            conn.close()

        colunas = ["Data", "NF", "Produto", "Peso Liquido",
                "Qtd Estoque", "Qtd Cobre", "Qtd Zinco", "Qtd Sucata",
                "Valor Total NF", "Mão de Obra", "Matéria Prima",
                "Custo Manual", "Custo Total"]
        import pandas as pd
        df = pd.DataFrame(dados, columns=colunas)

        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Relatorio_calculo_produto.xlsx"
        )

        if not caminho_arquivo:
            return

        try:
            df.to_excel(caminho_arquivo, index=False)
            messagebox.showinfo("Exportação", f"Exportação concluída com sucesso.\nArquivo salvo em:\n{caminho_arquivo}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

    def formatar_data(self, *args):
        """
        Formata o conteúdo do campo de pesquisa enquanto o usuário digita.
        Se o conteúdo tiver exatamente 8 dígitos (sem formatação), atualiza para dd/mm/aaaa.
        Se já estiver formatado (10 caracteres com barras nas posições corretas), não altera.
        """
        conteudo = self.search_var.get().strip()
    
        # Se já estiver formatado como data (ex: 25/04/2023), não faz nada.
        if len(conteudo) == 10 and conteudo[2] == '/' and conteudo[5] == '/':
            try:
                datetime.strptime(conteudo, "%d/%m/%Y")
                return
            except ValueError:
                pass
    
        # Extrai somente os dígitos
        digitos = ''.join(ch for ch in conteudo if ch.isdigit())
        if len(digitos) == 8:
            novo_conteudo = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
        else:
            novo_conteudo = conteudo  # Se não houver exatamente 8 dígitos, mantém o que foi digitado
    
        if novo_conteudo != conteudo:
            # Remove o trace para evitar recursividade
            self.search_var.trace_remove("write", self.trace_id)
            self.search_var.set(novo_conteudo)
            self.trace_id = self.search_var.trace_add("write", self.formatar_data)

    def remove_acento(self, texto):
        # Normaliza a string e remove os caracteres de marca (acento)
        return ''.join(
            c for c in unicodedata.normalize('NFD', texto)
            if unicodedata.category(c) != 'Mn'
        )

    def pesquisar(self):
        # Salva a ordem atual dos itens (apenas na primeira pesquisa)
        if not hasattr(self, "ordem_original") or not self.ordem_original:
            self.ordem_original = list(self.treeview.get_children())
        
        # Obtém o termo de busca sem acentos e em minúsculas
        termo = self.remove_acento(self.entrada_pesquisa.get().lower())
        
        # Reanexa todos os itens que estavam ocultos para que todos sejam avaliados.
        for item in self.itens_ocultos[:]:
            self.treeview.reattach(item, '', 'end')
            self.itens_ocultos.remove(item)
        
        # Itera por todos os itens atualmente exibidos
        for item in self.treeview.get_children():
            valores = self.treeview.item(item, "values")
            # Remove acentos de cada valor e faz a comparação
            if not any(termo in self.remove_acento(str(valor).lower()) for valor in valores):
                self.treeview.detach(item)
                if item not in self.itens_ocultos:
                    self.itens_ocultos.append(item)

    def limpar_pesquisa(self):
        self.entrada_pesquisa.delete(0, "end")

        # Reexibe todos os itens ocultos (reattach sem usar 'end' para preservar posição)
        for item in self.itens_ocultos:
            # Ao reattach sem especificar índice, ele volta a posição anterior; 
            # se precisar garantir posição, vamos reposicionar mais abaixo usando ordem_original
            self.treeview.reattach(item, "", "end")
        self.itens_ocultos.clear()

        # Se salvamos a ordem original antes da pesquisa, restauramos exatamente essa ordem
        if hasattr(self, "ordem_original") and self.ordem_original:
            for idx, item in enumerate(self.ordem_original):
                try:
                    self.treeview.move(item, '', idx)
                except Exception:
                    # se algum item não existir mais, apenas ignora
                    pass
            # limpamos o buffer da ordem original (próxima pesquisa nova ordem será salva novamente)
            self.ordem_original = []

    def reiniciar_ids_estoque(self):
        try:
            conn = conectar()
            cursor = conn.cursor()

            # Reorganiza os IDs, começando do 1
            cursor.execute("""
                WITH OrderedProducts AS (
                    SELECT id, ROW_NUMBER() OVER (ORDER BY id) AS new_id
                    FROM estoque_quantidade
                )
                UPDATE estoque_quantidade eq
                SET id = op.new_id
                FROM OrderedProducts op
                WHERE eq.id = op.id
            """)

            # Resetando a sequência do campo ID para garantir que o próximo seja 1
            cursor.execute("""
                SELECT setval(pg_get_serial_sequence('estoque_quantidade', 'id'), 1, false);
            """)

            conn.commit()

            # Fecha a conexão
            cursor.close()
            conn.close()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao reiniciar os IDs: {e}")

    def mostrar_historico(self):
        """
        Janela do Histórico com filtro por produto e barra de pesquisa.
        """
        try:
            conn = conectar()
            cur = conn.cursor()
            # busca produtos distintos para popular o filtro
            cur.execute("SELECT DISTINCT produto FROM calculo_historico WHERE produto IS NOT NULL ORDER BY produto;")
            produtos_db = [r[0] for r in cur.fetchall() if r[0] is not None]
            cur.close()
            conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar produtos para filtro: {e}")
            produtos_db = []

        largura, altura = 1000, 650
        dialog = tk.Toplevel(self.root)
        dialog.title("Histórico de Cálculo")
        dialog.transient(self.root)
        try:
            aplicar_icone(dialog, "C:\\Sistema\\logos\\Kametal.ico")
        except Exception:
            pass
        try:
            centralizar_janela(dialog, largura, altura)
        except Exception:
            pass

        # frame principal
        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill="both", expand=True)

        # --- filtros (top) ---
        filtros_frame = ttk.Frame(frame)
        filtros_frame.grid(row=0, column=0, sticky="ew", pady=(0,8))
        filtros_frame.columnconfigure(2, weight=1)

        ttk.Label(filtros_frame, text="Produto:", font=("Arial", 10)).grid(row=0, column=0, padx=(0,6))
        produto_vals = ["(Todos)"] + produtos_db
        produto_var = tk.StringVar(value="(Todos)")
        combo_produto = ttk.Combobox(filtros_frame, values=produto_vals, textvariable=produto_var, state="readonly", width=30)
        combo_produto.grid(row=0, column=1, padx=(0,8))

        ttk.Label(filtros_frame, text="Pesquisar:", font=("Arial", 10)).grid(row=0, column=2, padx=(0,4), sticky="e")
        termo_var = tk.StringVar()
        entrada_busca = ttk.Entry(filtros_frame, textvariable=termo_var)
        entrada_busca.grid(row=0, column=3, sticky="ew", padx=(0,8))

        btn_filtrar = ttk.Button(filtros_frame, text="Filtrar")
        btn_limpar = ttk.Button(filtros_frame, text="Limpar")
        btn_filtrar.grid(row=0, column=4, padx=(2,4))
        btn_limpar.grid(row=0, column=5)

        # --- Treeview + scrollbars ---
        vsb = ttk.Scrollbar(frame, orient="vertical")
        hsb = ttk.Scrollbar(frame, orient="horizontal")

        cols = ("Usuário", "NF", "Produto", "Quantidade", "Operação", "Data/Hora")
        tv = ttk.Treeview(frame, columns=cols, show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=20)

        for c in cols:
            tv.heading(c, text=c)
        tv.column("Usuário", width=140, anchor="center")
        tv.column("NF", width=100, anchor="center")
        tv.column("Produto", width=300, anchor="center")
        tv.column("Quantidade", width=100, anchor="center")
        tv.column("Operação", width=100, anchor="center")
        tv.column("Data/Hora", width=180, anchor="center")

        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        tv.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        # função para carregar do DB com filtros
        def carregar_rows(produto_filter="(Todos)", termo=""):
            try:
                conn = conectar()
                cur = conn.cursor()
                base = """
                    SELECT usuario, nf, produto, quantidade, tipo, timestamp
                    FROM calculo_historico
                """
                conditions = []
                params = []

                if produto_filter and produto_filter != "(Todos)":
                    conditions.append("produto = %s")
                    params.append(produto_filter)

                if termo and termo.strip():
                    # normaliza (remove acentos)
                    termo_proc = unicodedata.normalize("NFD", termo.strip().lower())
                    termo_proc = "".join(ch for ch in termo_proc if unicodedata.category(ch) != "Mn")

                    # mapeia termos da interface para os do banco
                    if termo_proc in ["adicionado"]:
                        termo_proc = "adicionar"
                    elif termo_proc in ["subtraido", "subtraído"]:  # com ou sem acento
                        termo_proc = "subtrair"

                    termo_like = f"%{termo_proc}%"
                    conditions.append("(usuario ILIKE %s OR CAST(nf AS TEXT) ILIKE %s OR produto ILIKE %s OR tipo ILIKE %s)")
                    params.extend([termo_like, termo_like, termo_like, termo_like])

                if conditions:
                    base += " WHERE " + " AND ".join(conditions)

                base += " ORDER BY timestamp DESC LIMIT 1000;"
                cur.execute(base, tuple(params))
                fetched = cur.fetchall()
                cur.close()
                conn.close()
                return fetched
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar histórico: {e}")
                return []

        # atualiza treeview com linhas (formata timestamp e operação)
        def atualizar_tv(rows):
            tv.delete(*tv.get_children())
            for usuario, nf, produto, quantidade, tipo, ts in rows:
                operacao = "Adicionado" if tipo == "adicionar" else "Subtraído"
                try:
                    datahora = ts.strftime("%d/%m/%Y %H:%M:%S") if ts else ""
                except Exception:
                    datahora = str(ts) if ts else ""
                usuario_para_mostrar = usuario or ""
                produto_para_mostrar = produto or ""
                quantidade_formatada = str(quantidade).replace(".", ",")
                tv.insert("", "end", values=(usuario_para_mostrar, nf, produto_para_mostrar, quantidade_formatada, operacao, datahora))

        # pegar as linhas visíveis (para exportar)
        def pegar_linhas_visiveis():
            linhas = []
            for iid in tv.get_children():
                vals = tv.item(iid, "values")
                if len(vals) >= 6:
                    linhas.append((vals[0], vals[1], vals[2], vals[3], vals[4], vals[5]))
                else:
                    padded = list(vals) + [""] * (6 - len(vals))
                    linhas.append(tuple(padded[:6]))
            return linhas

        # ações dos botões filtrar/limpar
        def aplicar_filtro(event=None):
            prod = produto_var.get()
            termo = termo_var.get()
            rows = carregar_rows(prod, termo)
            atualizar_tv(rows)

        def limpar_filtro():
            produto_var.set("(Todos)")
            termo_var.set("")
            rows = carregar_rows("(Todos)", "")
            atualizar_tv(rows)

        btn_filtrar.config(command=aplicar_filtro)
        btn_limpar.config(command=limpar_filtro)
        entrada_busca.bind("<Return>", aplicar_filtro)

        # botão exportar / fechar (abaixo)
        btn_frame_bottom = ttk.Frame(frame)
        btn_frame_bottom.grid(row=3, column=0, pady=10, sticky="ew")
        btn_frame_bottom.columnconfigure(0, weight=1)

        ttk.Button(btn_frame_bottom, text="Exportar Visível", command=lambda: self.exportar_linhas_para_excel(pegar_linhas_visiveis())).pack(side="right", padx=6)
        ttk.Button(btn_frame_bottom, text="Exportar Tudo", command=lambda: self.exportar_historico_para_excel()).pack(side="right", padx=6)
        ttk.Button(btn_frame_bottom, text="Fechar", command=dialog.destroy).pack(side="right", padx=6)

        # carrega dados iniciais (limit 500 -> usamos 1000 no carregar_rows; aqui mostramos os últimos 500 como antes)
        initial_rows = carregar_rows("(Todos)", "")
        if len(initial_rows) > 500:
            initial_rows = initial_rows[:500]
        atualizar_tv(initial_rows)

    def garantir_tabela_historico(self):
        """
        Cria a tabela calculo_historico no banco caso não exista.
        Chamar no __init__ da classe principal.
        """
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS calculo_historico (
                    id serial PRIMARY KEY,
                    usuario TEXT,
                    nf TEXT,
                    produto TEXT,
                    quantidade NUMERIC,
                    tipo TEXT,
                    timestamp timestamptz DEFAULT now()
                );
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_calculo_historico_produto ON calculo_historico(produto);")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_calculo_historico_nf ON calculo_historico(nf);")
            conn.commit()
            cur.close()
            conn.close()
        except Exception as e:
            print("Erro ao garantir tabela calculo_historico:", e)
            traceback.print_exc()

    def purga_automatica_na_abertura(self):
        """
        Limpa registros anteriores ao mês atual na tabela calculo_historico.
        Operação silenciosa feita na inicialização (preserva registros do mês corrente).

        Agora tenta renumerar os IDs dos registros remanescentes (começando em 1) **apenas**
        se não existirem foreign keys referenciando a tabela calculo_historico.
        """
        try:
            conn = conectar()
            cur = conn.cursor()
            # Apaga antigos
            cur.execute(
                "DELETE FROM calculo_historico WHERE timestamp < date_trunc('month', now()) RETURNING id;"
            )
            deleted = len(cur.fetchall())
            conn.commit()

            # tenta renumerar, se aplicável
            try:
                cur.execute("""
                    SELECT COUNT(*) FROM pg_constraint
                    WHERE confrelid = 'calculo_historico'::regclass AND contype = 'f';
                """)
                fk_count = cur.fetchone()[0] or 0

                if fk_count == 0:
                    cur.execute("SELECT COUNT(*) FROM calculo_historico;")
                    remaining = cur.fetchone()[0] or 0
                    if remaining > 0:
                        cur.execute("""
                            CREATE TABLE IF NOT EXISTS calculo_historico_tmp AS
                            SELECT usuario, nf, produto, quantidade, tipo, timestamp
                            FROM calculo_historico
                            ORDER BY timestamp, id;
                        """)
                        cur.execute("TRUNCATE calculo_historico RESTART IDENTITY;")
                        cur.execute("""
                            INSERT INTO calculo_historico (usuario, nf, produto, quantidade, tipo, timestamp)
                            SELECT usuario, nf, produto, quantidade, tipo, timestamp FROM calculo_historico_tmp ORDER BY timestamp, nf;
                        """)
                        cur.execute("DROP TABLE IF EXISTS calculo_historico_tmp;")
                        conn.commit()
                    else:
                        try:
                            cur.execute("SELECT setval(pg_get_serial_sequence('calculo_historico','id'), 1, false);")
                            conn.commit()
                        except Exception:
                            conn.rollback()
                else:
                    # Existem FKs; não renumerar
                    print("Renumeração automática ignorada: chaves estrangeiras apontam para calculo_historico.")
            except Exception as e_ren:
                conn.rollback()
                print("Erro ao tentar renumerar na purga automática:", e_ren)

            # notifica outras instâncias, se houver LISTEN
            try:
                cur.execute("NOTIFY historico_atualizado, 'purge';")
                conn.commit()
            except Exception:
                pass

            cur.close()
            conn.close()
            if deleted:
                print(f"Purga automática: removidas {deleted} linhas de calculo_historico")
        except Exception as e:
            print("Erro em purga_automatica_na_abertura:", e)
            traceback.print_exc()

    def registrar_acao_historico(self, nf, produto, quantidade, tipo, usuario=None):
        """
        Registra a ação no histórico (insere no BD e mantém cache local).
        tipo: 'subtrair' ou 'adicionar'
        """
        # garante que exista cache local
        if not hasattr(self, "history") or self.history is None:
            self.history = []

        try:
            qt = Decimal(str(quantidade))
        except Exception:
            qt = Decimal("0")

        agora = datetime.now()
        timestamp_iso = agora.isoformat()
        data_str = agora.strftime("%d/%m/%Y")
        hora_str = agora.strftime("%H:%M:%S")

        entrada = {
            "nf": str(nf),
            "produto": str(produto),
            "quantidade": f"{qt}",
            "tipo": tipo,
            "timestamp": timestamp_iso,
            "data": data_str,
            "hora": hora_str,
            "usuario": usuario or ""
        }

        # evita duplicatas imediatas locais
        if self.history:
            last = self.history[-1]
            if (last.get("nf") == entrada["nf"] and
                last.get("produto") == entrada["produto"] and
                last.get("quantidade") == entrada["quantidade"] and
                last.get("tipo") == entrada["tipo"]):
                return

        # grava no banco
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO calculo_historico (usuario, nf, produto, quantidade, tipo) VALUES (%s, %s, %s, %s, %s)",
                (usuario or "", str(nf), str(produto), Decimal(str(qt)), tipo)
            )
            conn.commit()
            # notifica outras instâncias
            cur.execute("NOTIFY historico_atualizado, 'novo';")
            conn.commit()
            cur.close()
            conn.close()
        except Exception as e:
            print("Erro gravando histórico no DB:", e)
            traceback.print_exc()

        # mantém cache local para compatibilidade
        self.history.append(entrada)

        # Limita a 100 registros por produto no cache local (mesma lógica do original)
        try:
            registros_produto = [e for e in self.history if e.get("produto") == produto]
            if len(registros_produto) > 100:
                # remove antigos e mantém últimos 100
                self.history = [e for e in self.history if e.get("produto") != produto]
                self.history.extend(registros_produto[-100:])
        except Exception:
            pass

        # persiste cache local e atualiza UI
        try:
            self.save_history()
        except Exception:
            pass
        try:
            self.atualizar_notificacao_badge()
        except Exception:
            pass

    def recarregar_historico_do_bd(self, limit=1000):
        """
        Carrega os últimos 'limit' registros do banco para self.history (cache local).
        Chamar após receber NOTIFY ou na inicialização.
        """
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("""
                SELECT usuario, nf, produto, quantidade, tipo, timestamp
                FROM calculo_historico
                ORDER BY timestamp DESC
                LIMIT %s
            """, (limit,))
            rows = cur.fetchall()
            cur.close()
            conn.close()
        except Exception as e:
            print("Erro carregando histórico do DB:", e)
            traceback.print_exc()
            return

        nova_history = []
        for r in rows:
            usuario_db, nf_db, produto_db, quantidade_db, tipo_db, ts_db = r
            try:
                ts_iso = ts_db.isoformat()
                data_str = ts_db.strftime("%d/%m/%Y")
                hora_str = ts_db.strftime("%H:%M:%S")
            except Exception:
                ts_iso = str(ts_db)
                data_str = ""
                hora_str = ""
            nova_history.append({
                "usuario": usuario_db,
                "nf": str(nf_db),
                "produto": str(produto_db),
                "quantidade": f"{quantidade_db}",
                "tipo": tipo_db,
                "timestamp": ts_iso,
                "data": data_str,
                "hora": hora_str
            })

        # ordena ascendente por timestamp para compatibilidade com UI
        self.history = list(reversed(nova_history))
        try:
            self.save_history()
        except Exception:
            pass
        try:
            self.root.after(0, self.atualizar_notificacao_badge)
        except Exception:
            pass

    def iniciar_escuta_historico(self):
        """
        Inicia uma thread que faz LISTEN historico_atualizado e chama recarregar_historico_do_bd
        quando receber notificação.
        """
        def listen_loop():
            while True:
                try:
                    conn = conectar()
                    conn.set_isolation_level(psycopg2.extensions.ISOLATION_LEVEL_AUTOCOMMIT)
                    cur = conn.cursor()
                    cur.execute("LISTEN historico_atualizado;")
                    # loop de escuta
                    while True:
                        if select.select([conn], [], [], 5) == ([], [], []):
                            continue
                        conn.poll()
                        while conn.notifies:
                            notify = conn.notifies.pop(0)
                            try:
                                self.root.after(0, lambda: self.recarregar_historico_do_bd())
                            except Exception:
                                pass
                except Exception as e:
                    print("Erro na thread de escuta:", e)
                    traceback.print_exc()
                    # aguarda antes de tentar reconectar
                    time.sleep(10)

        t = threading.Thread(target=listen_loop, daemon=True)
        t.start()

    def exportar_historico_para_excel(self):
        """
        Exporta todo o histórico do banco para um arquivo .xlsx (pede local de salvamento).
        Retorna o caminho do arquivo salvo, ou None se o usuário cancelar / ocorrer erro.
        Exporta apenas as colunas necessárias e com cabeçalhos:
        Usuário, NF, Produto, Quantidade, Operação, Data/Hora
        """
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("SELECT usuario, nf, produto, quantidade, tipo, timestamp FROM calculo_historico ORDER BY timestamp;")
            rows = cur.fetchall()
            cur.close()
            conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao recuperar histórico para exportar: {e}")
            return None

        if not rows:
            messagebox.showinfo("Exportar Histórico", "Não há registros para exportar.")
            return None

        # Process rows: format timestamp and map tipo to operação amigável
        processed = []
        for r in rows:
            usuario, nf, produto, quantidade, tipo, ts = r
            # format timestamp
            try:
                if isinstance(ts, datetime):
                    if ts.tzinfo is not None:
                        ts_naive = ts.astimezone(timezone.utc).replace(tzinfo=None)
                    else:
                        ts_naive = ts
                    ts_str = ts_naive.strftime("%d/%m/%Y %H:%M:%S")
                else:
                    ts_str = str(ts) if ts is not None else ""
            except Exception:
                ts_str = str(ts) if ts is not None else ""

            operacao = "Adicionado" if tipo == "adicionar" else "Subtraído" if tipo == "subtrair" else (tipo or "")
            # quantidade as string with comma decimal if numeric
            try:
                qtd_str = str(quantidade).replace(".", ",")
            except Exception:
                qtd_str = str(quantidade) if quantidade is not None else ""

            processed.append((usuario or "", nf or "", produto or "", qtd_str, operacao, ts_str))

        # Build DataFrame with desired headers and without id column
        import pandas as pd
        df = pd.DataFrame(processed, columns=["Usuário", "NF", "Produto", "Quantidade", "Operação", "Data/Hora"])

        # default filename
        default_name = f"historico_calculo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar histórico como..."
        )
        if not caminho:
            return None

        try:
            df.to_excel(caminho, index=False)
            messagebox.showinfo("Exportar Histórico", f"Histórico exportado com sucesso para:\n{caminho}")
            return caminho
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar o arquivo Excel: {e}")
            return None
        
    def exportar_linhas_para_excel(self, rows, sugestao_nome=None):
        """
        Exporta a lista de tuplas 'rows' para .xlsx.
        rows: lista de tuplas (usuario, nf, produto, quantidade, tipo, timestamp)
        """
        if not rows:
            messagebox.showinfo("Exportar Histórico", "Não há registros para exportar.")
            return

        # monta DataFrame
        try:
            df = pd.DataFrame(rows, columns=["Usuário", "NF", "Produto", "Quantidade", "Operação", "Data/Hora"])
        except Exception:
            # fallback simples
            df = pd.DataFrame([list(r) for r in rows], columns=["Usuário", "NF", "Produto", "Quantidade", "Operação", "Data/Hora"])

        # sugere nome
        if sugestao_nome:
            default_name = sugestao_nome
        else:
            default_name = f"historico_calculo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar histórico como..."
        )
        if not caminho:
            return

        try:
            df.to_excel(caminho, index=False)
            messagebox.showinfo("Exportar Histórico", f"Histórico exportado com sucesso para:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar o arquivo Excel: {e}")
   
    def iniciar_rotina_purga_mensal(self):
        """
        Inicia uma thread que agenda aviso no final do mês para apagar a tabela de histórico.
        """
        def segundos_ate_proxima_purga():
            hoje = date.today()
            if hoje.month == 12:
                primeiro_proximo = date(hoje.year + 1, 1, 1)
            else:
                primeiro_proximo = date(hoje.year, hoje.month + 1, 1)
            ultimo_dia = primeiro_proximo - timedelta(days=1)
            # agendamos para 23:55 do último dia do mês
            purge_dt = datetime(ultimo_dia.year, ultimo_dia.month, ultimo_dia.day, 23, 55)
            agora = datetime.now()
            delta = (purge_dt - agora).total_seconds()
            if delta < 0:
                # já passou, agenda para mês seguinte
                if primeiro_proximo.month == 12:
                    fn = date(primeiro_proximo.year + 1, 1, 1)
                else:
                    fn = date(primeiro_proximo.year, primeiro_proximo.month + 1, 1)
                ultimo_dia2 = fn - timedelta(days=1)
                purge_dt = datetime(ultimo_dia2.year, ultimo_dia2.month, ultimo_dia2.day, 23, 55)
                delta = (purge_dt - agora).total_seconds()
            return max(60, delta)

        def purge_loop():
            while True:
                try:
                    wait_sec = segundos_ate_proxima_purga()
                    time.sleep(wait_sec)
                    try:
                        # Corrigido: chama o método da instância
                        self.root.after(0, self.avisar_e_purgar)
                    except Exception:
                        pass
                    time.sleep(60)
                except Exception as e:
                    print("Erro no loop de purga:", e)
                    traceback.print_exc()
                    time.sleep(60)

        t = threading.Thread(target=purge_loop, daemon=True)
        t.start()

    def avisar_e_purgar(self):
        """
        Diálogo informativo: avisa que o histórico será limpo automaticamente na virada do mês.
        Oferece: Exportar (mantém dados) / Fechar. NÃO apaga nada aqui.
        """
        dialog = tk.Toplevel(self.root)
        dialog.title("Aviso — Histórico de Cálculo")
        dialog.transient(self.root)
        dialog.grab_set()
        try:
            aplicar_icone(dialog, "C:\\Sistema\\logos\\Kametal.ico")
        except Exception:
            pass
        try:
            centralizar_janela(dialog, 480, 200)
        except Exception:
            pass

        texto = (
            "Aviso: Está próximo do fim do mês. Este diálogo apenas avisa que o histórico "
            "de cálculo será limpo automaticamente quando o mês virar. Se desejar, "
            "exporte para Excel agora para manter uma cópia dos dados."
        )
        ttk.Label(dialog, text=texto, wraplength=440, justify="left").pack(padx=10, pady=(12,8))

        obs = (
            "Observação: a limpeza automática ocorrerá na virada do mês quando qualquer "
            "usuário abrir o sistema. Este diálogo NÃO executa a limpeza."
        )
        ttk.Label(dialog, text=obs, wraplength=440, justify="left").pack(padx=10)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=12)

        def on_export_only():
            dialog.destroy()
            try:
                caminho = self.exportar_historico_para_excel()
            except Exception as e:
                caminho = None
                print("Erro ao exportar histórico:", e)
                traceback.print_exc()
            if caminho:
                try:
                    messagebox.showinfo("Exportação", f"Histórico exportado para:\n{caminho}")
                except Exception:
                    pass
            else:
                try:
                    messagebox.showinfo("Exportação", "Exportação cancelada ou falhou. Registros permanecerão e serão limpos na virada do mês.")
                except Exception:
                    pass

        def on_close():
            dialog.destroy()

        ttk.Button(btn_frame, text="Exportar (manter dados)", command=on_export_only).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Fechar", command=on_close).pack(side="left", padx=6)

        try:
            dialog.focus_force()
        except Exception:
            pass

    def segundos_ate_proxima_purga(self):
        agora = datetime.now()
        ano = agora.year
        mes = agora.month

        # último dia do mês corrente
        ultimo_dia = calendar.monthrange(ano, mes)[1]
        purge_dt = datetime(ano, mes, ultimo_dia, 23, 55)

        delta = (purge_dt - agora).total_seconds()
        if delta < 0:
            # agendar para o próximo mês
            if mes == 12:
                ano2, mes2 = ano + 1, 1
            else:
                ano2, mes2 = ano, mes + 1
            ultimo_dia2 = calendar.monthrange(ano2, mes2)[1]
            purge_dt = datetime(ano2, mes2, ultimo_dia2, 23, 55)
            delta = (purge_dt - agora).total_seconds()

        # garante espera mínima (evita sleep(0))
        return max(60, int(delta))

    def executar_purga(self, total=False):
        """
        Executa a purga. Por padrão (total=False) apaga apenas registros anteriores ao mês atual.
        Se total=True apaga tudo — usar com extremo cuidado e apenas por ação explícita.

        Implementação atualizada:
        - se total=True: usa TRUNCATE ... RESTART IDENTITY (garante que próximo id comece em 1)
        - se parcial: deleta registros antigos e, SE NÃO existirem foreign keys referenciando
        calculo_historico, recria os registros restantes em ordem (renumerando os ids para 1..N).
        Se houver FKs, a renumeração é ignorada (para não quebrar referências).
        """
        try:
            conn = conectar()
            cur = conn.cursor()

            if total:
                # conta antes para informar quantas linhas foram removidas
                try:
                    cur.execute("SELECT COUNT(*) FROM calculo_historico;")
                    total_before = cur.fetchone()[0] or 0
                except Exception:
                    total_before = 0
                # Usa TRUNCATE para apagar tudo e reiniciar sequência
                try:
                    cur.execute("TRUNCATE calculo_historico RESTART IDENTITY;")
                    deleted = total_before
                    conn.commit()
                except Exception:
                    conn.rollback()
                    # Fallback: tentar DELETE + setval
                    cur.execute("DELETE FROM calculo_historico;")
                    deleted = cur.rowcount
                    cur.execute("SELECT setval(pg_get_serial_sequence('calculo_historico','id'), 1, false);")
                    conn.commit()

            else:
                # delete parcial: registros anteriores ao mês atual
                try:
                    cur.execute("DELETE FROM calculo_historico WHERE timestamp < date_trunc('month', now()) RETURNING id;")
                    removed_rows = cur.fetchall()
                    deleted = len(removed_rows)
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise

                # após remoção parcial, tenta renumerar os IDs restantes
                try:
                    # verifica se há FK referenciando calculo_historico
                    cur.execute("""
                        SELECT COUNT(*) FROM pg_constraint
                        WHERE confrelid = 'calculo_historico'::regclass AND contype = 'f';
                    """)
                    fk_count = cur.fetchone()[0] or 0

                    if fk_count == 0:
                        # apenas renumera se houver registros remanescentes
                        cur.execute("SELECT COUNT(*) FROM calculo_historico;")
                        remaining = cur.fetchone()[0] or 0
                        if remaining > 0:
                            # recria tabela temporária com os dados na ordem desejada e reinsera para renumerar IDs
                            cur.execute("""
                                CREATE TABLE calculo_historico_tmp AS
                                SELECT usuario, nf, produto, quantidade, tipo, timestamp
                                FROM calculo_historico
                                ORDER BY timestamp, id;
                            """)
                            cur.execute("TRUNCATE calculo_historico RESTART IDENTITY;")
                            cur.execute("""
                                INSERT INTO calculo_historico (usuario, nf, produto, quantidade, tipo, timestamp)
                                SELECT usuario, nf, produto, quantidade, tipo, timestamp FROM calculo_historico_tmp ORDER BY timestamp, nf;
                            """)
                            cur.execute("DROP TABLE IF EXISTS calculo_historico_tmp;")
                            conn.commit()
                        else:
                            # sem registros remanescentes, força sequência a 1
                            try:
                                cur.execute("SELECT setval(pg_get_serial_sequence('calculo_historico','id'), 1, false);")
                                conn.commit()
                            except Exception:
                                conn.rollback()
                    else:
                        # Há FKs — não renumerar; apenas logar para o usuário/admin
                        print("Renumeração de IDs ignorada: existem chaves estrangeiras apontando para calculo_historico.")
                except Exception as e_renum:
                    # Se algo falhar na renumeração, desfaz e continua (não queremos perder dados já válidos)
                    conn.rollback()
                    print("Falha ao renumerar IDs após purga parcial:", e_renum)

            # Notifica outras instâncias que escutem (opcional)
            try:
                cur.execute("NOTIFY historico_atualizado, 'purge';")
                conn.commit()
            except Exception:
                # NOTIFY não é crítico — ignora falhas aqui
                pass

            cur.close()
            conn.close()

            try:
                messagebox.showinfo("Purge Mensal", f"Purga concluída. Linhas removidas: {deleted}")
            except Exception:
                # Garantir que falha em mostrar diálogo não quebre a função
                print(f"Purga concluída. Linhas removidas: {deleted}")
        except Exception as e:
            try:
                messagebox.showerror("Erro", f"Erro ao apagar histórico: {e}")
            except Exception:
                print("Erro ao apagar histórico:", e)
            print("Erro no executar_purga:", e)
            traceback.print_exc()

    def verificar_e_avisar_exportacao_inicio(self, dias_aviso=3):
        try:
            hoje = date.today()
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            dias_restantes = ultimo_dia - hoje.day

            # mostra messagebox localmente (como você já faz)
            if dias_restantes <= dias_aviso:
                # ... seu código para mostrar messagebox local (pode manter)
                pass

            # Envia NOTIFY para o menu com payload 'purge:<dias>'
            payload = f"purge:{dias_restantes}"
            try:
                connn = conectar()
                cur = connn.cursor()
                cur.execute("SELECT pg_notify('canal_atualizacao', %s);", (payload,))
                connn.commit()
                cur.close()
                connn.close()
            except Exception:
                # não crítico; apenas logue se quiser
                pass

        except Exception as e:
            print("Erro em verificar_e_avisar_exportacao_inicio:", e)

    def voltar_para_menu(self):
        """Reexibe o menu imediatamente e faz a limpeza em background para não bloquear a UI."""
        try:
            self.janela_menu.deiconify()
            self.janela_menu.state("zoomed")
            self.janela_menu.lift()
            self.janela_menu.update_idletasks()
            self.janela_menu.focus_force()
        except Exception:
            pass

        def _cleanup_and_destroy():
            try:
                child = getattr(self, "root", None) or self
                if hasattr(child, "cursor") and getattr(child, "cursor"):
                    try:
                        child.cursor.close()
                    except Exception:
                        pass
                if hasattr(child, "conn") and getattr(child, "conn"):
                    try:
                        child.conn.close()
                    except Exception:
                        pass
            finally:
                try:
                    child.after(0, child.destroy)
                except Exception:
                    try:
                        child.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()
        
    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente"""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):

            # Fecha a conexão com o banco de dados, se estiver aberta
            if hasattr(self, "conn") and self.conn:
                self.conn.close()

            self.root.destroy()  # Fecha a janela
            sys.exit(0)  # Encerra o processo completamente