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
from datetime import datetime, date
import unicodedata
import threading
import json
import os

class CalculoProduto:
    def __init__(self, janela_menu):
        self.janela_menu = janela_menu
        self.root = tk.Toplevel()
        self.root.title("Calculo de Nfs")
        self.root.geometry("1200x700")
        self.root.state("zoomed")

        # Oculta a janela de menu
        self.janela_menu.withdraw()

        # Aplica o ícone (defina ou adapte a função aplicar_icone conforme necessário)
        aplicar_icone(self.root, "C:\\Sistema\\logos\\Kametal.ico")

        self.configurar_estilo()
        self.criar_widgets()
        self.carregar_dados_iniciais()

        # garante que a pasta 'config' exista
        config_dir = os.path.join(os.path.dirname(__file__), "config")
        os.makedirs(config_dir, exist_ok=True)

        # caminho do arquivo de histórico local dentro da pasta config
        self.history_path = os.path.join(config_dir, "History_Calculo.json")
        self.history = self.load_history()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Chama a função que reorganiza os IDs ao inicializar a janela
        self.reiniciar_ids_estoque()

    def configurar_estilo(self):
        self.style = ttk.Style(self.root)
        self.style.theme_use("alt")
        self.style.configure("Treeview", 
                            background="white", 
                            foreground="black", 
                            rowheight=27, 
                            fieldbackground="white")
        self.style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        self.style.map("Treeview", 
                    background=[("selected", "#0078D7")], 
                    foreground=[("selected", "white")])
        
        # Configura o Treeview com o mesmo estilo e 23 linhas visíveis
        if hasattr(self, "treeview"):
            self.treeview.config(style="Treeview", height=22)
    
    def criar_widgets(self):
        # Frame superior para entradas
        self.frame_top = ttk.Frame(self.root)
        self.frame_top.pack(padx=15, pady=(10, 0), fill="x")
    
        ttk.Label(self.frame_top, text="Produto:", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 1))
        self.entrada_produto = ttk.Combobox(self.frame_top, width=25, values=[], state="normal")
        self.entrada_produto.pack(side="left", padx=(0, 10))
        self.entrada_produto.bind("<<ComboboxSelected>>", self.on_select_produto)
    
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
            command=self.show_notifications,
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

        # atualiza badge/contador do botão agora que ele existe e está empacotado
        try:
            self.atualizar_notificacao_badge()
        except Exception:
            pass
    
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
                sp.data DESC, sp.nf ASC;
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

        if operacao == "Subtrair":
            total_disponivel = sum(Decimal(r[4]) for r in resultados)

            if valor_entrada > total_disponivel:
                messagebox.showerror(
                    "Erro",
                    f"O valor de entrada ({valor_entrada:.3f}) ultrapassa o total disponível nas notas fiscais ({total_disponivel:.3f})!"
                )
                cursor.close()
                conn.close()
                return

            for resultado in resultados:
                id_produto, nf, prod, peso_liquido, quantidade_estoque = resultado
                quantidade_estoque = Decimal(quantidade_estoque)

                if valor_restante <= quantidade_estoque:
                    nova_quantidade_estoque = quantidade_estoque - valor_restante
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante = Decimal("0")
                    break
                else:
                    nova_quantidade_estoque = Decimal("0")
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante -= quantidade_estoque

            if valor_restante > 0:
                messagebox.showwarning("Aviso", f"Não foi possível subtrair todo o valor! Restante: {valor_restante:.3f}")

        elif operacao == "Adicionar":
            # Pega as NFs mais recentes primeiro
            query = """
            SELECT sp.id, sp.nf, sp.produto, sp.peso_liquido, COALESCE(eq.quantidade_estoque, sp.peso_liquido)
            FROM somar_produtos sp
            LEFT JOIN estoque_quantidade eq ON sp.id = eq.id_produto
            WHERE sp.produto = %s
            ORDER BY sp.data DESC, sp.nf DESC
            """
            cursor.execute(query, (produto_nome,))
            resultados = cursor.fetchall()

            total_disponivel = sum(Decimal(r[3]) - Decimal(r[4]) for r in resultados)

            if valor_entrada > total_disponivel:
                messagebox.showerror(
                    "Erro",
                    f"O valor de entrada ({valor_entrada:.3f}) ultrapassa o espaço disponível nas notas fiscais ({total_disponivel:.3f})!"
                )
                cursor.close()
                conn.close()
                return

            for resultado in resultados:
                id_produto, nf, prod, peso_liquido, quantidade_estoque = resultado
                peso_liquido = Decimal(peso_liquido)
                quantidade_estoque = Decimal(quantidade_estoque)

                if valor_restante <= 0:
                    break

                espaco_disponivel = peso_liquido - quantidade_estoque

                if valor_restante <= espaco_disponivel:
                    nova_quantidade_estoque = quantidade_estoque + valor_restante
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante = Decimal("0")
                else:
                    nova_quantidade_estoque = peso_liquido
                    atualizacoes.append((id_produto, nf, nova_quantidade_estoque, quantidade_estoque))
                    valor_restante -= espaco_disponivel

            if valor_restante > 0:
                messagebox.showwarning("Aviso", f"Não foi possível adicionar todo o valor! Restante: {valor_restante:.3f}")

        # Atualiza a tabela no banco
        for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
            insert_query = """
            INSERT INTO estoque_quantidade (id_produto, quantidade_estoque)
            VALUES (%s, %s)
            ON CONFLICT (id_produto)
            DO UPDATE SET quantidade_estoque = EXCLUDED.quantidade_estoque
            """
            cursor.execute(insert_query, (id_produto, nova_quantidade))

        conn.commit()

        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        conn.commit()

        # Atualiza o Treeview para refletir as alterações
        for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
            for item in self.treeview.get_children():
                valores = self.treeview.item(item, "values")
                # valores[1] = NF (coluna 1), valores[4] = Qtd Estoque (coluna 4)
                if str(valores[1]) == str(nf):
                    # Formata quantidade nova para exibição (vírgula)
                    quantidade_formatada = f"{nova_quantidade:.3f}".replace('.', ',')
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

                    if nova_quantidade == 0:
                        self.treeview.item(item, tags="red_text")
                    elif Decimal(str(nova_quantidade)) != Decimal(str(quantidade_decimal_tree)):
                        self.treeview.item(item, tags="partial_red_text")
                    else:
                        self.treeview.item(item, tags="")

                    self.treeview.item(item, values=tuple(novos_valores))
                    break

        # <<< registra no histórico usando os valores corretos (quantidade_anterior) >>>
        for id_produto, nf, nova_quantidade, quantidade_anterior in atualizacoes:
            if operacao == "Subtrair":
                quantidade_sub = quantidade_anterior - nova_quantidade
                if quantidade_sub > 0:
                    self.record_action(nf, produto_nome, quantidade_sub, "subtrair")

            elif operacao == "Adicionar":
                quantidade_add = nova_quantidade - quantidade_anterior
                if quantidade_add > 0:
                    self.record_action(nf, produto_nome, quantidade_add, "adicionar")

        cursor.close()
        conn.close()
        messagebox.showinfo("Sucesso", "Operação realizada com sucesso!")

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

    # ------------------- Histórico local (JSON) -------------------
    def load_history(self):
        """Carrega o histórico local de ações (JSON)."""
        try:
            if os.path.exists(self.history_path):
                with open(self.history_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        # estrutura: lista de entradas
        return []

    def save_history(self):
        """Salva o histórico atual no arquivo JSON."""
        try:
            with open(self.history_path, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("Erro ao salvar histórico local:", e)

    def record_action(self, nf, produto, quantidade, tipo):
        """
        Registra uma ação local:
        tipo: 'subtrair' ou 'adicionar'
        quantidade: string ou número (trata como Decimal)
        Não grava entradas idênticas consecutivas (proteção contra duplicados).
        """
        # garante que history exista
        if not hasattr(self, "history") or self.history is None:
            self.history = []

        try:
            qt = Decimal(str(quantidade))
        except Exception:
            qt = Decimal("0")

        now = datetime.now()
        timestamp_iso = now.isoformat()
        data_str = now.strftime("%d/%m/%Y")
        hora_str = now.strftime("%H:%M:%S")

        entry = {
            "nf": str(nf),
            "produto": str(produto),
            "quantidade": f"{qt}",   # string para evitar problemas de json com Decimal
            "tipo": tipo,
            "timestamp": timestamp_iso,
            "data": data_str,
            "hora": hora_str
        }

        # Proteção contra duplicatas imediatas: se a última entrada for igual, ignora
        if self.history:
            last = self.history[-1]
            if (last.get("nf") == entry["nf"] and
                last.get("produto") == entry["produto"] and
                last.get("quantidade") == entry["quantidade"] and
                last.get("tipo") == entry["tipo"]):
                return  # já existe uma entrada idêntica imediatamente anterior

        self.history.append(entry)

        # Limita em 100 registros por produto (mantém só os últimos 100)
        registros_produto = [e for e in self.history if e.get("produto") == produto]
        if len(registros_produto) > 100:
            # Remove todas as entradas deste produto e adiciona apenas as últimas 100
            self.history = [e for e in self.history if e.get("produto") != produto]
            self.history.extend(registros_produto[-100:])

        self.save_history()
        self.atualizar_notificacao_badge()

    def compute_remaining_from_history(self, nf, produto, peso_liquido):
        """
        Calcula quanto resta na NF/produto aplicando o histórico local sobre o peso_liquido original.
        Retorna Decimal(remaining).
        """
        try:
            restante = Decimal(str(peso_liquido))
        except Exception:
            restante = Decimal("0")
        for e in self.history:
            if str(e.get("nf")) == str(nf) and str(e.get("produto")).strip() == str(produto).strip():
                try:
                    q = Decimal(str(e.get("quantidade", "0")))
                except Exception:
                    q = Decimal("0")
                if e.get("tipo") == "subtrair":
                    restante -= q
                elif e.get("tipo") == "adicionar":
                    restante += q
        # não deixa negativo se preferir
        return restante if restante >= 0 else Decimal("0")

    def atualizar_notificacao_badge(self):
        """Atualiza o texto do botão de notificações com a quantidade de NFs 'não processadas'."""
        try:
            dados = self.carregar_dados()  # usa sua função que traz (data, nf, produto, peso_liquido, ...)
        except Exception:
            dados = []
        count = 0
        seen = set()
        for row in dados:
            nf = row[1]; produto = row[2]; peso_liquido = row[3] or 0
            chave = f"{nf}|{produto}"
            if chave in seen:
                continue
            seen.add(chave)
            restante = self.compute_remaining_from_history(nf, produto, peso_liquido)
            # se restante == peso_liquido então ninguém mexeu (não processado)
            try:
                if Decimal(str(restante)) == Decimal(str(peso_liquido)):
                    count += 1
            except Exception:
                pass
        self.botao_notificacao.config(text=f"Historico de Calculo")

    def show_notifications(self):
        """Mostra janela de histórico filtrado por produto."""
        largura, altura = 700, 600
        dialog = tk.Toplevel(self.root)
        dialog.title("Notificações - Histórico")
        aplicar_icone(dialog, "C:\\Sistema\\logos\\Kametal.ico")

        # centraliza a janela
        centralizar_janela(dialog, largura, altura)

        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill="both", expand=True)

        # Lista suspensa de produtos (filtra valores None)
        produtos_unicos = sorted({e.get("produto") for e in (self.history or []) if e.get("produto")})
        ttk.Label(frame, text="Selecionar Produto:", font=("Arial", 10, "bold")).pack(anchor="w")

        produto_var = tk.StringVar()
        combo_produtos = ttk.Combobox(
            frame,
            textvariable=produto_var,
            values=produtos_unicos,
            state="readonly",
            width=40
        )
        combo_produtos.pack(fill="x", pady=5)
        combo_produtos.set('')  # garante que não apareça texto inicial

        # Container para treeview + scrollbars
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        # Treeview com 5 colunas (inclui Data)
        cols = ("Produto", "Quantidade", "Operação", "Data", "Hora")
        tv = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            height=15,
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )

        for c in cols:
            tv.heading(c, text=c)
        # ajustar larguras
        tv.column("Produto", anchor="w", width=240)
        tv.column("Quantidade", anchor="center", width=100)
        tv.column("Operação", anchor="center", width=100)
        tv.column("Data", anchor="center", width=90)
        tv.column("Hora", anchor="center", width=70)

        # associa scrollbars
        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        # layout com grid (para scrollbars funcionarem bem)
        tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # Função para preencher a tabela filtrada
        def atualizar_tabela(*args):
            tv.delete(*tv.get_children())  # limpa
            produto_sel = produto_var.get()
            if not produto_sel:
                return

            # filtra só o produto selecionado
            registros = [e for e in self.history if e.get("produto") == produto_sel]

            # ordena do mais recente para o mais antigo, usando timestamp
            registros.sort(key=lambda e: e.get("timestamp", ""), reverse=True)

            for e in registros:
                quantidade = e.get("quantidade", "")
                operacao = "Adicionado" if e.get("tipo") == "adicionar" else "Subtraído"

                # obtém data/hora
                if e.get("data") and e.get("hora"):
                    data_str = e.get("data")
                    hora_str = e.get("hora")
                else:
                    ts = e.get("timestamp")
                    try:
                        dt = datetime.fromisoformat(ts)
                        data_str = dt.strftime("%d/%m/%Y")
                        hora_str = dt.strftime("%H:%M:%S")
                    except Exception:
                        data_str, hora_str = "", ""

                tv.insert("", "end", values=(produto_sel, quantidade, operacao, data_str, hora_str))

        # Atualiza quando selecionar produto
        combo_produtos.bind("<<ComboboxSelected>>", atualizar_tabela)

        # Botão de fechar
        ttk.Button(frame, text="Fechar", command=dialog.destroy).pack(pady=5)

    def on_select_produto(self, event):
        produto = self.entrada_produto.get().strip()
        if not produto:
            return

        linhas = []
        for e in (self.history or []):
            if e.get("produto", "").strip().lower() == produto.lower():
                linhas.append(f"{e.get('timestamp')} | {e.get('tipo')} {e.get('quantidade')} kg")

        if not linhas:
            return   # não mostra mensagem nenhuma

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
