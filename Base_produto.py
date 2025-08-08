import tkinter as tk
from tkinter import messagebox, ttk
from conexao_db import conectar
from centralizacao_tela import centralizar_janela  # Importar a função
from exportacao import exportar_para_pdf, exportar_para_excel #, selecionar_diretorio
import sys
from logos import aplicar_icone
from tkinter import filedialog
import re

class InterfaceProduto:
    def __init__(self, janela_menu=None):
        self.janela_menu = janela_menu
        self.produto_ids = {}  # Dicionário para armazenar o ID real do produto (chave: item_id da Treeview)

        # Criação e configuração da janela principal da interface
        self.janela_base_produto = tk.Toplevel()
        self.janela_base_produto.title("Base de Produtos")

        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self.janela_base_produto, caminho_icone)

        self.janela_base_produto.state("zoomed")
        self.janela_base_produto.configure(bg="#ecf0f1")

        largura_tela = self.janela_base_produto.winfo_screenwidth()
        altura_tela = self.janela_base_produto.winfo_screenheight()
        self.janela_base_produto.geometry(f"{largura_tela}x{altura_tela}")
        self.janela_base_produto.attributes("-topmost", 0)

        # Chama o método para configurar os estilos da janela Produto
        self._configurar_estilo_produto()

        # Cabeçalho
        cabecalho = tk.Label(self.janela_base_produto, text="Base de Produtos", font=("Arial", 24, "bold"), bg="#34495e", fg="white", pady=15)
        cabecalho.pack(fill=tk.X)

        # Frame para Labels e Entradas
        frame_acoes = tk.Frame(self.janela_base_produto, bg="#ecf0f1")
        frame_acoes.pack(padx=20, pady=10, fill=tk.X)

        tk.Label(frame_acoes, text="Nome do Produto", bg="#ecf0f1", font=("Arial", 12))\
            .grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entrada_nome = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entrada_nome.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_acoes, text="Porcentagem de Cobre (%)", bg="#ecf0f1", font=("Arial", 12))\
            .grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entrada_cobre = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entrada_cobre.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_acoes, text="Porcentagem de Zinco (%)", bg="#ecf0f1", font=("Arial", 12))\
            .grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entrada_zinco = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entrada_zinco.grid(row=2, column=1, padx=5, pady=5)

        # Frame para os botões de ação
        frame_botoes_acao = tk.Frame(self.janela_base_produto, bg="#ecf0f1")
        frame_botoes_acao.pack(pady=10, fill=tk.X)

        botao_adicionar = ttk.Button(frame_botoes_acao, text="Adicionar", command=self.adicionar_produto)
        botao_adicionar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_alterar = ttk.Button(frame_botoes_acao, text="Alterar", command=self.alterar_produto)
        botao_alterar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_excluir = ttk.Button(frame_botoes_acao, text="Excluir", command=self.excluir_produto)
        botao_excluir.pack(side=tk.LEFT, padx=5, pady=5)

        botao_limpar = ttk.Button(frame_botoes_acao, text="Limpar", command=self.limpar_campos)
        botao_limpar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_voltar = ttk.Button(frame_botoes_acao, text="Voltar", command=self.voltar_para_menu)
        botao_voltar.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para os botões de exportação
        frame_botoes_exportacao = tk.Frame(self.janela_base_produto, bg="#ecf0f1")
        frame_botoes_exportacao.pack(pady=10, fill=tk.X)

        botao_exportar_excel = ttk.Button(frame_botoes_exportacao, text="Exportar Excel", command=self.exportar_excel_produtos)
        botao_exportar_excel.pack(side=tk.LEFT, padx=5, pady=5)

        botao_exportar_pdf = ttk.Button(frame_botoes_exportacao, text="Exportar PDF", command=self.exportar_pdf_produtos)
        botao_exportar_pdf.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para o Treeview
        frame_treeview = tk.Frame(self.janela_base_produto, bg="#ecf0f1")
        frame_treeview.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        scrollbar_vertical = tk.Scrollbar(frame_treeview, orient=tk.VERTICAL)
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        scrollbar_horizontal = tk.Scrollbar(frame_treeview, orient=tk.HORIZONTAL)
        scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

        self.lista_produtos = ttk.Treeview(
            frame_treeview, 
            columns=("Nome", "Cobre", "Zinco"),
            show="headings", 
            yscrollcommand=scrollbar_vertical.set, 
            xscrollcommand=scrollbar_horizontal.set
        )
        self.lista_produtos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.lista_produtos.heading("Nome", text="Nome", anchor="center")
        self.lista_produtos.heading("Cobre", text="Cobre (%)", anchor="center")
        self.lista_produtos.heading("Zinco", text="Zinco (%)", anchor="center")

        self.lista_produtos.column("Nome", anchor="center", width=200)
        self.lista_produtos.column("Cobre", anchor="center", width=150)
        self.lista_produtos.column("Zinco", anchor="center", width=150)

        self.atualizar_lista_produtos()

        scrollbar_vertical.config(command=self.lista_produtos.yview)
        scrollbar_horizontal.config(command=self.lista_produtos.xview)

        self.lista_produtos.bind("<ButtonRelease-1>", self.selecionar_produto)

        self.janela_base_produto.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.janela_base_produto.mainloop()

    def _configurar_estilo_produto(self):
        """Configura os estilos para a janela de produtos."""
        style = ttk.Style(self.janela_base_produto)
        style.theme_use("alt")
        # Estilo para frames (usado em containers, por exemplo)
        style.configure("Custom.TFrame", background="#ecf0f1")
        
        # Estilo para o Treeview (aumentamos o rowheight para 30)
        style.configure("Treeview", 
                        background="white", 
                        foreground="black", 
                        rowheight=20,      # Valor aumentado de 20 para 30
                        fieldbackground="white",
                        font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.map("Treeview", 
                background=[("selected", "#0078D7")], 
                foreground=[("selected", "white")])
        
        # Estilo para os botões de ação (Produto.TButton)
        style.configure("Produto.TButton",
                        padding=(5, 2),
                        relief="raised",
                        background="#2980b9",
                        foreground="white",
                        font=("Arial", 10),
                        borderwidth=2,
                        highlightbackground="#34495e",
                        highlightthickness=1)
        style.map("Produto.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")],
                relief=[("pressed", "sunken"), ("!pressed", "raised")])
        
        # Estilo para o botão de exportar para Excel
        style.configure("Excel.TButton",
                        padding=(5, 2),
                        relief="raised",
                        background="#27ae60",
                        foreground="white",
                        font=("Arial", 10),
                        borderwidth=2,
                        highlightbackground="#34495e",
                        highlightthickness=1)
        
        # Estilo para o botão de exportar para PDF
        style.configure("PDF.TButton",
                        padding=(5, 2),
                        background="#c0392b",
                        foreground="white",
                        font=("Arial", 10),
                        borderwidth=2,
                        highlightbackground="#34495e",
                        highlightthickness=1)
        return style

    def obter_menor_id_disponivel(self):
        """Obtém o menor ID disponível na tabela produtos."""
        conexao = conectar()
        if conexao is None:
            return None
        cursor = conexao.cursor()
        cursor.execute("SELECT id FROM produtos ORDER BY id")
        ids_existentes = cursor.fetchall()
        conexao.close()

        if not ids_existentes:
            return 1

        ids_existentes = [item[0] for item in ids_existentes]
        menor_id = 1
        while menor_id in ids_existentes:
            menor_id += 1
        return menor_id

    def atualizar_lista_produtos(self):
        """Atualiza a lista de produtos exibida na Treeview e o dicionário produto_ids."""
        for item in self.lista_produtos.get_children():
            self.lista_produtos.delete(item)
        self.produto_ids.clear()

        conexao = conectar()
        if conexao is None:
            return
        cursor = conexao.cursor()
        cursor.execute("SELECT id, nome, percentual_cobre, percentual_zinco FROM produtos ORDER BY nome ASC")
        produtos = cursor.fetchall()
        for produto in produtos:
            produto_id, nome, cobre, zinco = produto
            item_id = self.lista_produtos.insert("", "end", values=(nome, f"{int(cobre)}%", f"{int(zinco)}%"))
            self.produto_ids[item_id] = produto_id
        conexao.close()

    def remover_acentos(self, texto):
        """Remove acentos do texto."""
        texto = re.sub(r'[áàâãäåÁÀÂÃÄÅ]', 'a', texto)
        texto = re.sub(r'[éèêëÉÈÊË]', 'e', texto)
        texto = re.sub(r'[íìîïÍÌÎÏ]', 'i', texto)
        texto = re.sub(r'[óòôõöÓÒÔÕÖ]', 'o', texto)
        texto = re.sub(r'[úùûüÚÙÛÜ]', 'u', texto)
        texto = re.sub(r'[çÇ]', 'c', texto)
        return texto

    def adicionar_produto(self):
        """Adiciona um novo produto ao banco de dados."""
        nome = self.entrada_nome.get()
        percentual_cobre = self.entrada_cobre.get()
        percentual_zinco = self.entrada_zinco.get()

        if nome and percentual_cobre and percentual_zinco:
            nome_normalizado = self.remover_acentos(nome)
            nome_normalizado = ' '.join(nome_normalizado.split())
            novo_id = self.obter_menor_id_disponivel()
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            cursor.execute(
                "INSERT INTO produtos (id, nome, percentual_cobre, percentual_zinco) VALUES (%s, %s, %s, %s)",
                (novo_id, nome_normalizado, percentual_cobre, percentual_zinco)
            )
            conexao.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conexao.commit()

            conexao.close()

            item_id = self.lista_produtos.insert("", "end", values=(nome, percentual_cobre, percentual_zinco))
            self.produto_ids[item_id] = novo_id

            messagebox.showinfo("Sucesso", "Produto adicionado com sucesso!")
            self.atualizar_lista_produtos()

             # Atualizar o relatório de produtos automaticamente
            if self.janela_menu and hasattr(self.janela_menu, "frame_relatorios_produto_material"):
                for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                    widget.destroy()
                self.janela_menu.criar_relatorio_produto_material()  # Atualizando o relatório de produto


            self.limpar_campos()
        else:
            messagebox.showwarning("Campos obrigatórios", "Por favor, preencha todos os campos.")

    def excluir_produto(self):
        """Exclui os produtos selecionados."""
        selected_items = self.lista_produtos.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione ao menos um produto da lista!")
            return

        if messagebox.askyesno("Confirmação", f"Tem certeza que deseja excluir {len(selected_items)} produto(s)?"):
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            for item_id in selected_items:
                produto_id = self.produto_ids.get(item_id)
                if produto_id is None:
                    messagebox.showerror("Erro", "ID do produto não encontrado.")
                    continue
                try:
                    cursor.execute("DELETE FROM produtos WHERE id=%s", (produto_id,))
                    del self.produto_ids[item_id]
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao excluir o produto com ID {produto_id}: {e}")
            conexao.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conexao.commit()

            conexao.close()
            messagebox.showinfo("Sucesso", "Produtos excluídos com sucesso!")
            self.atualizar_lista_produtos()

            # Atualizar o relatório de produtos automaticamente
            if self.janela_menu and hasattr(self.janela_menu, "frame_relatorios_produto_material"):
                for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                    widget.destroy()
                self.janela_menu.criar_relatorio_produto_material()  # Atualizando o relatório de produto

            self.limpar_campos()

    def alterar_produto(self):
        """Altera o produto selecionado no banco de dados."""
        selected_items = self.lista_produtos.selection()
        if not selected_items:
            messagebox.showerror("Erro", "Nenhum produto selecionado.")
            return

        produto_id = self.produto_ids.get(selected_items[0])
        if produto_id is None:
            messagebox.showerror("Erro", "ID do produto não encontrado.")
            return

        novo_nome = self.entrada_nome.get()
        novo_cobre = self.entrada_cobre.get()
        novo_zinco = self.entrada_zinco.get()

        if not novo_nome or not novo_cobre or not novo_zinco:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
            return

        novo_nome_normalizado = self.remover_acentos(novo_nome)

        conexao = conectar()
        if conexao is None:
            return
        cursor = conexao.cursor()
        try:
            cursor.execute(
                """
                UPDATE produtos
                SET nome=%s, percentual_cobre=%s, percentual_zinco=%s
                WHERE id=%s
                """,
                (novo_nome_normalizado, novo_cobre, novo_zinco, produto_id)
            )
            conexao.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conexao.commit()

            messagebox.showinfo("Sucesso", "Produto alterado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao alterar o produto: {e}")
        finally:
            cursor.close()
            conexao.close()
        self.atualizar_lista_produtos()

        # Atualizar o relatório de produtos automaticamente
        if self.janela_menu and hasattr(self.janela_menu, "frame_relatorios_produto_material"):
            for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                widget.destroy()
            self.janela_menu.criar_relatorio_produto_material()  # Atualizando o relatório de produto

    def selecionar_produto(self, event):
        """Preenche os campos de entrada com os dados do produto selecionado."""
        selected = self.lista_produtos.selection()
        if selected:
            produto = self.lista_produtos.item(selected, "values")
            print("Produto selecionado:", produto)
            if len(produto) >= 3:
                self.entrada_nome.delete(0, tk.END)
                self.entrada_nome.insert(0, produto[0])
                self.entrada_cobre.delete(0, tk.END)
                self.entrada_cobre.insert(0, produto[1].replace('%', ''))
                self.entrada_zinco.delete(0, tk.END)
                self.entrada_zinco.insert(0, produto[2].replace('%', ''))
            else:
                print("O produto selecionado não contém dados suficientes.")
        else:
            print("Nenhum produto selecionado.")

    def limpar_campos(self):
        """Limpa os campos de entrada."""
        self.entrada_nome.delete(0, tk.END)
        self.entrada_cobre.delete(0, tk.END)
        self.entrada_zinco.delete(0, tk.END)

    def exportar_pdf_produtos(self):
        """Exporta os dados para um arquivo PDF."""
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="base_produtos.pdf"
        )
        if caminho_arquivo:
            exportar_para_pdf(caminho_arquivo, "produtos", ["Produto", "Cobre %", "Zinco %"], "Base de Produtos")

    def exportar_excel_produtos(self):
        """Exporta os dados para um arquivo Excel."""
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="base_produtos.xlsx"
        )
        if caminho_arquivo:
            exportar_para_excel(caminho_arquivo, "produtos", ["Produto", "Cobre %", "Zinco %"])

    def ordena_coluna(self, treeview, col, reverse):
        """Ordena a Treeview por uma coluna específica."""
        data = [(treeview.item(child)["values"], child) for child in treeview.get_children("")]
        data.sort(key=lambda x: x[0][col], reverse=reverse)
        for index, (item, tree_id) in enumerate(data):
            treeview.move(tree_id, '', index)

    def voltar_para_menu(self):
        """Fecha a janela de base do produto e reexibe o menu principal com atualização visual."""
        self.janela_base_produto.destroy()
        self.janela_menu.deiconify()
        self.janela_menu.state("zoomed")
        self.janela_menu.lift()  # Garante que fique no topo
        self.janela_menu.update()  # Força atualização visual
        
    def on_closing(self):
        """Função para lidar com o fechamento da janela."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar o programa?"):
            self.janela_base_produto.destroy()  # Fecha a janela atual
            sys.exit()  # Encerra completamente o programa
