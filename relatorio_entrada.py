import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar
from conexao_db import conectar  # Importa a função para conectar ao banco
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela

class RelatorioEntradaApp:
    def __init__(self, root, parent, relatorio_resumo):
        self.root = root
        self.parent = parent
        self.relatorio_resumo = relatorio_resumo  # Instância da janela de resumo
        # Configura a janela principal
        self.root.title("Relatório de Entrada de Produtos")
        self.root.state("zoomed")
        aplicar_icone(self.root, r"C:\Sistema\logos\Kametal.ico")
        
        # Cria a interface na aba de entrada (parent)
        self.criar_interface(self.parent)
        self.configurar_estilos()
        self.carregar_produtos_base()

    def configurar_estilos(self):
        """Define estilos globais para o Treeview e os botões."""
        estilo = ttk.Style()
        estilo.theme_use("alt")

        # Estilo básico para linhas da tabela
        estilo.configure("Treeview", 
                        font=("Arial", 10), 
                        rowheight=27, 
                        background="white", 
                        foreground="black", 
                        fieldbackground="white")

        # Estilo do cabeçalho
        estilo.configure("Treeview.Heading", 
                        font=("Arial", 10, "bold"))

        # Estilo para a linha de total
        estilo.configure("Total.Treeview", 
                        background="#FFD700", 
                        foreground="black", 
                        font=("Arial", 12, "bold"))

        # Estilo para seleção
        estilo.map("Treeview", 
                background=[("selected", "blue")], 
                foreground=[("selected", "white")])

        # Estilo para os botões da classe RelatorioEntradaApp
        estilo.configure("RelatorioEntrada.TButton",
                        padding=(5, 2),
                        relief="raised",
                        background="#FF6347",  # Cor laranja para os botões
                        foreground="white",
                        font=("Arial", 10, "bold"),
                        borderwidth=2,
                        highlightbackground="#34495e",
                        highlightthickness=1)
        estilo.map("RelatorioEntrada.TButton", 
                background=[("active", "#FF4500")],
                foreground=[("active", "white")],
                relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Configura o Treeview com o mesmo estilo e 22 linhas visíveis
        if hasattr(self, "treeview"):
            self.treeview.config(style="Treeview", height=22)

    def criar_interface(self, parent):
        """Cria os filtros e a tabela para a aba de entrada."""
        frame_filtros = ttk.LabelFrame(parent, text="Filtros", padding=(10, 10))
        frame_filtros.pack(fill=tk.X, padx=20, pady=10)

        label_style = {"font": ("Arial", 10, "bold")}

        # Primeira linha: Campos de data e combobox
        tk.Label(frame_filtros, text="Data Inicial (MM/YYYY):", **label_style).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.entrada_data_inicial = tk.Entry(frame_filtros, width=12, font=("Arial", 10))
        self.entrada_data_inicial.grid(row=0, column=1, padx=5, pady=5)
        self.entrada_data_inicial.bind("<KeyRelease>", self.formatar_data)

        tk.Label(frame_filtros, text="Data Final (MM/YYYY):", **label_style).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.entrada_data_final = tk.Entry(frame_filtros, width=12, font=("Arial", 10))
        self.entrada_data_final.grid(row=0, column=3, padx=5, pady=5)
        self.entrada_data_final.bind("<KeyRelease>", self.formatar_data)

        tk.Label(frame_filtros, text="Base do Produto:", **label_style).grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.combobox_produto_entrada = ttk.Combobox(frame_filtros, width=30, font=("Arial", 10))
        self.combobox_produto_entrada.grid(row=0, column=5, padx=5, pady=5)

        # Botões
        frame_botoes = tk.Frame(frame_filtros)
        frame_botoes.grid(row=1, column=0, columnspan=6, pady=10)

        botao_relatorio = ttk.Button(
            frame_botoes, text="Gerar Relatório",
            command=lambda: self.gerar_relatorio_entrada(),
            style="RelatorioEntrada.TButton"  # Usando o estilo configurado
        )
        botao_relatorio.pack(side=tk.LEFT, padx=5)

        self.botao_calcular = ttk.Button(
            frame_botoes, text="Calcular Totais",
            command=self.atualizar_total,
            style="RelatorioEntrada.TButton",
            state=tk.DISABLED
        )
        self.botao_calcular.pack(side=tk.LEFT, padx=5)

        # **Novo botão "Editar Estoque"**
        self.botao_editar_estoque = ttk.Button(
            frame_botoes, text="Editar Estoque",
            command=self.abrir_janela_edicao,
            style="RelatorioEntrada.TButton",  # Usando o estilo configurado
            state=tk.DISABLED
        )
        self.botao_editar_estoque.pack(side=tk.LEFT, padx=5)

         # **Novo botão: Enviar para Resumo**
        self.botao_enviar_resumo = ttk.Button(
            frame_botoes, text="Enviar para Resumo",
            command=self.enviar_para_resumo,
            style="RelatorioEntrada.TButton",
            state=tk.DISABLED  # Inicialmente desabilitado; habilite conforme a lógica do seu app
        )
        self.botao_enviar_resumo.pack(side=tk.LEFT, padx=5)

        # Frame da Tabela
        frame_tabela = tk.Frame(parent)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        colunas = ("Data", "NF", "Produto", "Peso Líquido", "Qtd Estoque", "Custo Total")
        self.tabela_entrada = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

        for coluna, largura in zip(colunas, [100, 100, 300, 150, 150, 150]):
            self.tabela_entrada.heading(coluna, text=coluna)
            self.tabela_entrada.column(coluna, anchor=tk.CENTER, width=largura)

        scroll_y = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=self.tabela_entrada.yview)
        self.tabela_entrada.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tabela_entrada.pack(fill=tk.BOTH, expand=True)

        self.tabela_entrada.tag_configure("total", background="#FFD700", foreground="black", font=("Arial", 12, "bold"))

        # **Habilitar botão ao selecionar uma linha**
        self.tabela_entrada.bind("<<TreeviewSelect>>", lambda e: self.habilitar_botao_edicao())

        self.tabela_entrada.bind("<<TreeviewSelect>>", lambda e: self.habilitar_botoes())

    def habilitar_botoes(self):
        """Habilita os botões de edição e envio quando uma linha é selecionada."""
        self.botao_editar_estoque.config(state=tk.NORMAL)
        self.botao_enviar_resumo.config(state=tk.NORMAL)
        # Caso necessário, habilite também o botão de calcular totais

    def enviar_para_resumo(self):
        """Coleta os dados do produto selecionado e envia para a janela de resumo."""
        selected_item = self.tabela_entrada.focus()
        if not selected_item:
            tk.messagebox.showwarning("Atenção", "Selecione uma linha para enviar ao resumo.")
            return

        # Obtém os valores da linha selecionada
        valores = self.tabela_entrada.item(selected_item, "values")

        def parse_value(valor_str):
            """Converte uma string numérica para float, tratando separadores e unidades."""
            if not valor_str or valor_str.strip() == "":
                return 0
            
            valor_str = valor_str.strip()

            # Remove símbolos como "R$" e "Kg"
            for token in ["R$", "Kg", "kg"]:
                valor_str = valor_str.replace(token, "").strip()

            # Corrige separadores de milhar e decimal
            if "." in valor_str and "," in valor_str:
                valor_str = valor_str.replace(".", "").replace(",", ".")  # Exemplo: 7.721,50 → 7721.50
            elif "," in valor_str and "." not in valor_str:
                valor_str = valor_str.replace(",", ".")  # Exemplo: 77,70 → 77.70

            try:
                return float(valor_str)
            except ValueError:
                return 0  # Retorna 0 se a conversão falhar

        try:
            qtd_estoque = parse_value(valores[4]) if len(valores) > 4 else 0
            custo_medio = parse_value(valores[5]) if len(valores) > 5 else 0  # 🔴 Agora pega direto!

            print(f"DEBUG - Qtd Estoque: {qtd_estoque}, Custo Médio: {custo_medio}")

        except ValueError as e:
            tk.messagebox.showerror("Erro", f"Dados inválidos para cálculo: {e}")
            return

        # Nome do produto
        nome_produto = self.combobox_produto_entrada.get().strip()

        # Atualiza a Treeview do resumo corretamente
        self.relatorio_resumo.atualizar_dados(nome_produto, qtd_estoque, custo_medio)

        tk.messagebox.showinfo("Sucesso", 
            f"Dados enviados para o resumo:\nProduto: {nome_produto}\nQtd: {qtd_estoque}\nCusto Médio: {custo_medio:.2f}")

    def habilitar_botao_edicao(self):
        """Habilita o botão de edição quando há uma linha selecionada."""
        selecionado = self.tabela_entrada.selection()
        self.botao_editar_estoque.config(state=tk.NORMAL if selecionado else tk.DISABLED)

    def abrir_janela_edicao(self):
        item_selecionado = self.tabela_entrada.selection()
        if not item_selecionado:
            messagebox.showwarning("Aviso", "Selecione uma linha para editar o estoque.")
            return

        item = item_selecionado[0]
        # Se a linha tiver a tag "total", não permita a edição
        if "total" in self.tabela_entrada.item(item, "tags"):
            messagebox.showwarning("Aviso", "A linha de total não pode ser editada.")
            return

        item = item_selecionado[0]
        valores = self.tabela_entrada.item(item, "values")

        peso_liquido = float(valores[3].replace(" Kg", "").replace(".", "").replace(",", "."))
        estoque_atual = float(valores[4].replace(" Kg", "").replace(".", "").replace(",", "."))

        # Criar janela de edição
        largura, altura = 320, 180
        self.janela_edicao = tk.Toplevel(self.root)
        self.janela_edicao.title("Editar Estoque")
        self.janela_edicao.configure(bg="#ffffff")  # Fundo branco para um visual mais limpo
        self.janela_edicao.resizable(False, False)  # Impede redimensionamento

        # Aplica ícone
        aplicar_icone(self.janela_edicao, r"C:\Sistema\logos\Kametal.ico")

        # Centralizar janela na tela
        centralizar_janela(self.janela_edicao, largura, altura)

        # Configuração dos estilos (mesmo usados na função alterar_produto)
        style = ttk.Style(self.janela_edicao)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TEntry", padding=5, relief="solid", font=("Arial", 10))
        style.configure("Alter.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Alter.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        # Frame principal
        frame_principal = tk.Frame(self.janela_edicao, bg="#ffffff", padx=15, pady=15)
        frame_principal.pack(fill="both", expand=True)

        # Rótulo e entrada
        tk.Label(frame_principal, text="Novo Estoque:", font=("Arial", 11, "bold"), bg="#ffffff").pack(anchor="w", pady=(0, 5))

        self.entrada_novo_estoque = tk.Entry(frame_principal, font=("Arial", 11), justify="center")
        self.entrada_novo_estoque.pack(fill="x", pady=5)
        self.entrada_novo_estoque.insert(0, f"{estoque_atual:,.3f}".replace(".", ","))  # Formato BR

        def salvar_alteracao():
            try:
                # Aqui a função é corrigida, utilizando o `self` corretamente
                novo_estoque_str = self.entrada_novo_estoque.get()

                # Substitui a vírgula por ponto para permitir a conversão correta para float
                novo_estoque = float(novo_estoque_str.replace(",", "."))

                # Verifica se o novo estoque não é maior que o peso líquido
                if novo_estoque > peso_liquido:
                    messagebox.showerror("Erro", f"O estoque não pode ser maior que o peso líquido ({peso_liquido:.3f} Kg).")
                    return

                # Atualiza a tabela com o novo estoque formatado corretamente
                valores_atualizados = list(valores)
                valores_atualizados[4] = f"{novo_estoque:,.3f} Kg".replace(",", "X").replace(".", ",").replace("X", ".")

                # Atualiza a linha na tabela
                self.tabela_entrada.item(item, values=valores_atualizados)

                # ✅ Recalcula os totais automaticamente
                self.calcular_totais()

                # Fecha a janela de edição
                self.janela_edicao.destroy()

            except ValueError:
                messagebox.showerror("Erro", "Digite um valor numérico válido para o estoque.")

        # Frame para os botões
        frame_botoes = tk.Frame(frame_principal, bg="#ffffff", pady=10)
        frame_botoes.pack(fill="x")

        botao_salvar = ttk.Button(frame_botoes, text="Salvar", command=salvar_alteracao, style="Alter.TButton")
        botao_salvar.pack(side="left", expand=True, padx=5)

        botao_cancelar = ttk.Button(frame_botoes, text="Cancelar", command=self.janela_edicao.destroy, style="Alter.TButton")
        botao_cancelar.pack(side="right", expand=True, padx=5)

    def formatar_data(self, event):
        """Adiciona a barra automaticamente ao digitar o mês (formato MM/YYYY)."""
        entry = event.widget
        texto = entry.get()
        # Insere a barra automaticamente quando há 2 dígitos e ainda não existe a barra
        if len(texto) == 2 and '/' not in texto:
            entry.insert(tk.END, '/')

    def carregar_produtos_base(self):
        """Carrega os produtos na combobox de entrada da tabela 'produtos'."""
        try:
            conn = conectar()
            cur = conn.cursor()
            
            # Modifique a consulta para pegar os produtos da tabela 'produtos' (coluna nome)
            cur.execute("SELECT DISTINCT nome FROM produtos ORDER BY nome;")
            produtos = [row[0] for row in cur.fetchall()]

            # Preencher a Combobox com os produtos encontrados
            self.combobox_produto_entrada['values'] = produtos

            if produtos:
                self.combobox_produto_entrada.current(0)  # Seleciona o primeiro produto da lista

            cur.close()
            conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar produtos base: {e}")

    def gerar_relatorio_entrada(self):
        """Gera o relatório de entrada sem exibir a linha total."""
        data_inicial_str = self.entrada_data_inicial.get()
        data_final_str = self.entrada_data_final.get()
        produto_selecionado = self.combobox_produto_entrada.get()

        # Converte as datas
        try:
            mes_inicial, ano_inicial = map(int, data_inicial_str.split('/'))
            data_inicio = f"{ano_inicial}-{mes_inicial:02d}-01"
            mes_final, ano_final = map(int, data_final_str.split('/'))
            ultimo_dia = calendar.monthrange(ano_final, mes_final)[1]
            data_fim = f"{ano_final}-{mes_final:02d}-{ultimo_dia}"
        except:
            messagebox.showerror("Erro", "Formato de data inválido. Use 'MM/YYYY'.")
            return

        query = """
            SELECT 
                sp.data, sp.nf, sp.produto, sp.peso_liquido, eq.quantidade_estoque, eq.custo_total
            FROM 
                somar_produtos sp
            JOIN 
                produtos p ON sp.produto = p.nome
            LEFT JOIN 
                estoque_quantidade eq ON eq.id_produto = sp.id
            WHERE 
                sp.data::DATE BETWEEN %s AND %s
                AND LOWER(sp.produto) = LOWER(%s)
            ORDER BY 
                sp.data ASC, sp.nf ASC;
        """

        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute(query, (data_inicio, data_fim, produto_selecionado))
            resultados = cur.fetchall()

            for item in self.tabela_entrada.get_children():
                self.tabela_entrada.delete(item)

            if not resultados:
                messagebox.showinfo("Sem Registros", "Nenhum dado encontrado para o período selecionado.")
                self.botao_calcular.config(state=tk.DISABLED)
                return

            for reg in resultados:
                self.tabela_entrada.insert("", tk.END, values=(
                    reg[0].strftime("%d/%m/%Y"), reg[1], reg[2],
                    self.formatar_quantidade(reg[3]), self.formatar_quantidade(reg[4]),
                    f"R$ {reg[5]:.2f}".replace(".", ",")
                ))

            self.resultados = resultados  # Armazena os dados para calcular depois
            self.botao_calcular.config(state=tk.NORMAL)  # Habilita o botão Calcular Totais

            cur.close()
            conn.close()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar a consulta: {e}")

    def calcular_totais(self):
        """Calcula os totais com base nos dados exibidos no Treeview, evitando duplicações por NF."""
        dados_por_nf = {}

        for item_id in self.tabela_entrada.get_children():
            # Pula se o item tiver a tag "total"
            if "total" in self.tabela_entrada.item(item_id, "tags"):
                continue

            valores = self.tabela_entrada.item(item_id, "values")
            if not valores or valores[4] == "" or valores[5] == "":
                continue

            try:
                numero_nf = valores[1]  # Alterado: usar o índice 1, que representa a NF
                estoque = float(valores[4].replace(" Kg", "").replace(".", "").replace(",", "."))
                custo = float(valores[5].replace("R$", "").replace(".", "").replace(",", "."))
                # Substitui a entrada anterior da mesma NF (sempre pega a última)
                dados_por_nf[numero_nf] = (estoque, custo)
            except ValueError:
                continue

        # Agora soma apenas os valores únicos por NF
        total_estoque = 0
        total_custo = 0

        for estoque, custo in dados_por_nf.values():
            total_estoque += estoque
            total_custo += estoque * custo

        media_ponderada = total_custo / total_estoque if total_estoque > 0 else 0

        total_estoque_formatado = self.formatar_quantidade(total_estoque)
        custo_total_formatado = f"R$ {media_ponderada:.2f}".replace(".", ",")

        # Remove linha de total anterior (se existir)
        for item_id in self.tabela_entrada.get_children():
            if "total" in self.tabela_entrada.item(item_id, "tags"):
                self.tabela_entrada.delete(item_id)

        # Adiciona nova linha de total
        self.tabela_entrada.insert("", tk.END, values=(
            "", "", "TOTAL GERAL", "", total_estoque_formatado, custo_total_formatado
        ), tags=("total",))

    def atualizar_total(self):
        total_estoque = 0
        total_custo = 0

        for item in self.tabela_entrada.get_children():
            # Pula se for a linha de total
            if "total" in self.tabela_entrada.item(item, "tags"):
                continue

            valores = self.tabela_entrada.item(item, "values")
            if len(valores) < 6:
                continue

            try:
                estoque = float(valores[4].replace(" Kg", "").replace(".", "").replace(",", "."))
                custo_total = float(valores[5].replace("R$ ", "").replace(".", "").replace(",", "."))
                total_estoque += estoque
                total_custo += custo_total * estoque
            except ValueError as e:
                print(f"Erro ao converter valores: {valores} - {e}")

        media_ponderada = total_custo / total_estoque if total_estoque > 0 else 0

        total_estoque_formatado = self.formatar_quantidade(total_estoque)
        custo_total_formatado = f"R$ {media_ponderada:.2f}".replace(".", ",")

        atualizado = False
        for item in self.tabela_entrada.get_children():
            valores = self.tabela_entrada.item(item, "values")
            if "total" in self.tabela_entrada.item(item, "tags"):
                continue
            if len(valores) > 1 and valores[1] == "TOTAL":
                self.tabela_entrada.item(item, values=("", "TOTAL", "", "", total_estoque_formatado, custo_total_formatado))
                atualizado = True
                break

        if not atualizado:
            self.tabela_entrada.insert("", tk.END, values=(
                "", "", "TOTAL GERAL", "", total_estoque_formatado, custo_total_formatado
            ), tags=("total",))

    def formatar_quantidade(self, valor):
        """Formata o valor da quantidade para o padrão brasileiro (com vírgula como separador decimal e ponto como separador de milhar)."""
        # Formatação com 3 casas decimais, convertendo o número para string
        valor_formatado = f"{valor:,.3f}"

        # Substitui o ponto por vírgula como separador decimal e usa o ponto para milhar
        valor_formatado = valor_formatado.replace(",", "X").replace(".", ",").replace("X", ".")

        return valor_formatado + " Kg"