import tkinter as tk
from tkinter import ttk
from conexao_db import conectar  # Importa a função conectar da sua biblioteca

class RelatorioResumoApp(tk.Frame):  # HERDA DE tk.Frame
    def __init__(self, parent):
        super().__init__(parent)  # Inicializa como um Frame do Tkinter
        self.parent = parent
        self.tree = None
        
        # Configura o estilo da Treeview
        self.configurar_estilo_treeview()
        
        # Cria a interface com a Treeview com 3 colunas
        self.criar_interface()
        
        # Carrega os dados do banco de dados
        self.carregar_dados()

    def criar_interface(self):
        """Cria a interface da aba com uma Treeview com três cabeçalhos."""
        self.frame_tabela = tk.Frame(self)
        self.frame_tabela.pack(expand=True, fill="both", padx=10, pady=10)
        
        # Criação da Treeview com a coluna de árvore (#0) e duas colunas adicionais
        self.tree = ttk.Treeview(self.frame_tabela, 
                                 columns=("qtd_estoque", "custo_medio"),
                                 show="tree headings", 
                                 style="Custom.Treeview")
        # Configura os cabeçalhos das colunas
        self.tree.heading("#0", text="Nome do Produto")
        self.tree.heading("qtd_estoque", text="Qtd Estoque")
        self.tree.heading("custo_medio", text="Custo Médio")
        
        # Define a largura e alinhamento das colunas
        self.tree.column("#0", anchor="center", width=350)
        self.tree.column("qtd_estoque", anchor="center", width=100)
        self.tree.column("custo_medio", anchor="center", width=150)
        
        # Barra de rolagem vertical
        self.scrollbar_vertical = tk.Scrollbar(self.frame_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar_vertical.set)
        
        # Layout usando grid
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        
        self.frame_tabela.grid_rowconfigure(0, weight=1)
        self.frame_tabela.grid_columnconfigure(0, weight=1)

    def configurar_estilo_treeview(self):
        """Configura o estilo personalizado para a Treeview."""
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Teste outros temas como "default", "alt", "classic"
        
        self.style.configure("Custom.Treeview", 
                            font=("Arial", 10), 
                            rowheight=30, 
                            background="white", 
                            foreground="black", 
                            fieldbackground="white")
        
        self.style.map("Custom.Treeview", 
                    background=[("selected", "blue")], 
                    foreground=[("selected", "white")])

    def carregar_dados(self):
        """Consulta a tabela 'produtos' e insere os dados na Treeview, 
        preenchendo apenas a coluna 'Nome do Produto' e deixando as demais em branco."""
        try:
            conexao = conectar()  # Usa sua função de conexão
            cursor = conexao.cursor()
            
            # Consulta somente o nome dos produtos
            query = "SELECT nome FROM produtos ORDER BY nome ASC"
            cursor.execute(query)
            registros = cursor.fetchall()
            
            # Remove itens existentes no Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Insere cada nome na Treeview; as colunas "Qtd Estoque" e "Custo Médio" ficam em branco
            for registro in registros:
                nome = registro[0]
                self.tree.insert("", tk.END, text=nome, values=("", ""))
        except Exception as e:
            print("Erro ao carregar dados:", e)
        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'conexao' in locals():
                conexao.close()

    def atualizar_dados(self, nome_produto, total_estoque, media_ponderada):
        """
        Atualiza a Treeview com os novos dados para o produto especificado.
        Se o produto não for encontrado, ele é adicionado.
        """
        nome_procura = nome_produto.strip().lower()
        print(f"DEBUG - Atualizando {nome_produto}: Estoque={total_estoque}, Média Ponderada={media_ponderada}")

        # Formata a quantidade de estoque (formatação com separador de milhar e 3 casas decimais) e adiciona "Kg"
        total_estoque_formatado = f"{total_estoque:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " Kg"
        # Formata a média ponderada com 2 casas decimais e adiciona "R$"
        media_ponderada_formatada = "R$ " + f"{media_ponderada:.2f}".replace(".", ",")

        for item in self.tree.get_children():
            nome_item = self.tree.item(item)["text"].strip().lower()
            if nome_item == nome_procura:
                self.tree.item(item, values=(total_estoque_formatado, media_ponderada_formatada))
                print("DEBUG - Atualizado:", self.tree.item(item))
                return

        # Se o produto não for encontrado, insere-o na Treeview
        novo_item = self.tree.insert("", tk.END, text=nome_produto, values=(total_estoque_formatado, media_ponderada_formatada))
        print("DEBUG - Inserido novo:", self.tree.item(novo_item))