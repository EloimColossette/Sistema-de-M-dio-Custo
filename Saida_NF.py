import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
from conexao_db import conectar
import sys
from centralizacao_tela import centralizar_janela
from logos import aplicar_icone
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import re
from decimal import Decimal

class SistemaNF(tk.Toplevel):
    def __init__(self, janela_menu=None, master=None):
        super().__init__(master=master)

        # Captura da resolução da tela
        screen_width = self.winfo_screenwidth()  # Largura da tela
        screen_height = self.winfo_screenheight()  # Altura da tela
        print(f"Resolução da tela: {screen_width}x{screen_height}")

        # Configuração do tamanho da janela como uma porcentagem da resolução da tela
        self.geometry(f"{int(screen_width * 0.8)}x{int(screen_height * 0.8)}")  # 80% da tela

        self.janela_menu = janela_menu 
        self.title("Saida de NF")
        self.config(bg="#f4f4f4")

        self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self, self.caminho_icone)

        # Maximizar a janela
        self.state("zoomed")

        # Conexão com o banco de dados
        try:
            self.conn = conectar()
            self.cursor = self.conn.cursor()
            print("Conexão com o banco de dados bem-sucedida.")
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            messagebox.showerror("Erro de Conexão", f"Não foi possível conectar ao banco de dados: {e}")
            return

        # Variáveis
        self.data_var = tk.StringVar()
        self.nf_var = tk.StringVar()
        self.cliente_var = tk.StringVar()
        self.doc_var = tk.StringVar()
        self.observacao_var = tk.StringVar()
        self.produto_entry = tk.StringVar()
        self.peso_entry = tk.StringVar()
        self.produtos = []
        self.base_produto_var = tk.StringVar()
        self.lista_nomes_produtos = self.obter_nomes_produtos()  # ou uma lista fixa, se preferir

        # Defina a variável delay_id para gerenciar o debounce
        self.delay_id = None

        # Criação dos campos de entrada e Treeview
        self.create_widgets()

        # Carregar os dados no Treeview
        self.atualizar_treeview()

        self.bind("<FocusIn>", lambda e: self.configurar_treeview())

        # Configurar fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self, bg="#f4f4f4")
        main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        # Configuração do estilo
        estilo = {'padx': 2, 'pady': 5}

        # Adicionando a barra de pesquisa
        search_frame = tk.Frame(main_frame, bg="#f4f4f4")
        search_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(search_frame, text="Pesquisar NF:", bg="#f4f4f4").pack(side=tk.LEFT, padx=5)

        # Variável de busca
        self.variavel_busca = tk.StringVar()
        self.entrada_busca = tk.Entry(search_frame, textvariable=self.variavel_busca, width=50)
        self.entrada_busca.pack(side=tk.LEFT, padx=5)
        # Aplica formatação apenas quando o usuário sai do campo
        self.entrada_busca.bind("<FocusOut>", self.ao_perder_foco)
        # Também formata ao pressionar Enter (antes de buscar)
        self.entrada_busca.bind("<Return>", lambda event: self.ao_perder_foco(event))
        self.entrada_busca.bind("<Return>", lambda event: self.buscar_nf())

        search_button = tk.Button(search_frame, text="Buscar", command=self.buscar_nf)
        search_button.pack(side=tk.LEFT, padx=5)

        # Novos botões para exportação
        tk.Button(search_frame, text="Excel", command=self.abrir_dialogo_exportacao, width=8).pack(side=tk.LEFT, padx=5, anchor="w")

        # Labels e Entradas em um Frame
        form_frame = tk.LabelFrame(main_frame, text="Dados da Nota Fiscal", bg="#f4f4f4")
        form_frame.pack(fill="x", padx=10, pady=10)

        # Campo Data com título
        tk.Label(form_frame, text="Data:", bg="#f4f4f4").grid(row=0, column=0, sticky="e")
        entry_data = tk.Entry(form_frame, textvariable=self.data_var, width=15)
        entry_data.grid(row=0, column=1, sticky="w")
        entry_data.bind("<FocusOut>", self.formatar_data)  # Formata a data ao perder o foco


        tk.Label(form_frame, text="NF:", bg="#f4f4f4").grid(row=0, column=2, **estilo, sticky="e")
        tk.Entry(form_frame, textvariable=self.nf_var, width=15).grid(row=0, column=3, **estilo, sticky="w")

        # CNPJ / CPF 1
        self.doc_var_1 = tk.StringVar()

        tk.Label(form_frame, text="CNPJ/CPF:", bg="#f4f4f4").grid(row=1, column=0, sticky="e")
        entry_doc = tk.Entry(form_frame, textvariable=self.doc_var, width=20)
        entry_doc.grid(row=1, column=1, sticky="w")
        entry_doc.bind("<FocusOut>", lambda e: (self.formatar_documento(e), self.verificar_documento(e)))

        # Campo Nome da Cliente (editável)
        tk.Label(form_frame, text="Cliente:", bg="#f4f4f4").grid(row=1, column=2, **estilo, sticky="e")
        tk.Entry(form_frame, textvariable=self.cliente_var, width=50).grid(row=1, column=3, columnspan=3, **estilo, sticky="w")

        tk.Label(form_frame, text="Observação:", bg="#f4f4f4").grid(row=2, column=0, **estilo, sticky="ne")
        self.observacao_text = tk.Text(form_frame, width=75, height=2, wrap="word")
        self.observacao_text.grid(row=2, column=1, columnspan=5, **estilo, sticky="w")

        # FRAME PARA PRODUTOS
        produto_frame = tk.LabelFrame(main_frame, text="Adicionar Produtos", bg="#f4f4f4")
        produto_frame.pack(fill="x", padx=10, pady=10)

        # Entrada de "Produto"
        tk.Label(produto_frame, text="Produto:", bg="#f4f4f4") \
            .grid(row=0, column=0, pady=5, sticky="e", padx=(0, 5))
        tk.Entry(produto_frame, textvariable=self.produto_entry, width=60) \
            .grid(row=0, column=1, pady=5, sticky="w", padx=(0, 10))

        # Entrada de "Peso"
        tk.Label(produto_frame, text="Peso:", bg="#f4f4f4") \
            .grid(row=0, column=2, pady=5, sticky="e", padx=(0, 5))
        entry_peso = tk.Entry(produto_frame, textvariable=self.peso_entry, width=10)
        entry_peso.grid(row=0, column=3, pady=5, sticky="w", padx=(0, 10))
        entry_peso.bind("<KeyRelease>", lambda e: self.formatar_numero(e, 3))

        # Entrada de "Base Produto" (combobox)
        tk.Label(produto_frame, text="Base Produto:", bg="#f4f4f4") \
            .grid(row=0, column=4, pady=5, sticky="e", padx=(0, 5))
        self.base_produto_combobox = ttk.Combobox(
            produto_frame,
            textvariable=self.base_produto_var,
            values=self.lista_nomes_produtos,
            width=30
        )
        self.base_produto_combobox.grid(row=0, column=5, pady=5, sticky="w", padx=(0, 5))

        # Botões de produtos
        botoes_produtos_frame = tk.Frame(produto_frame)
        botoes_produtos_frame.grid(row=3, column=0, columnspan=6, pady=5)
        tk.Button(botoes_produtos_frame, text="Adicionar Produto", command=self.adicionar_produto).grid(row=0, column=0, padx=2)
        tk.Button(botoes_produtos_frame, text="Alterar Produto", command=self.alterar_produto).grid(row=0, column=1, padx=2)
        tk.Button(botoes_produtos_frame, text="Excluir Produto", command=self.excluir_produto).grid(row=0, column=2, padx=2)

        # Listbox para exibição dos produtos adicionados
        lista_frame = tk.Frame(produto_frame, bg="#f4f4f4")
        lista_frame.grid(row=1, column=1, columnspan=4, sticky="nsew", pady=5)
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
        self.lista_produtos = tk.Listbox(lista_frame, width=80, height=4)
        self.lista_produtos.grid(row=0, column=0, sticky="nsew")
        scrollbar_lista = tk.Scrollbar(lista_frame, orient="vertical", command=self.lista_produtos.yview)
        scrollbar_lista.grid(row=0, column=1, sticky="ns")
        self.lista_produtos.config(yscrollcommand=scrollbar_lista.set)

        # FRAME PARA BOTÕES DA NF
        botoes_frame = tk.Frame(main_frame, bg="#f4f4f4")
        botoes_frame.pack(pady=10)

        tk.Button(botoes_frame, text="Salvar NF", command=self.salvar_nf, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_frame, text="Alterar NF", command=self.alterar_valores, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_frame, text="Excluir NF", command=self.excluir_valores, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_frame, text="Limpar", command=self.limpar_entradas, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(botoes_frame, text="Voltar", command=self.voltar_para_menu, width=15).pack(side=tk.LEFT, padx=5)

        # --- CONTAINER PARA O TREEVIEW E O LABEL DE SOMA ---
        container_frame = tk.Frame(main_frame, bg="#f4f4f4")
        container_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview ocupa a linha 0 do container
        self.tree = ttk.Treeview(container_frame, columns=("Data", "NF", "Produto", "Peso", "Cliente", "CNPJ/CPF", "Base Produto", "Observação"), show='headings')
        self.tree.grid(row=0, column=0, sticky="nsew")
        # Configurar scrollbars para o Treeview
        self.scrollbar_v = tk.Scrollbar(container_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar_v.grid(row=0, column=1, sticky="ns")
        self.scrollbar_h = tk.Scrollbar(container_frame, orient="horizontal", command=self.tree.xview)
        self.scrollbar_h.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=self.scrollbar_v.set, xscrollcommand=self.scrollbar_h.set)

        # Configurar colunas do Treeview
        colunas = ["Data", "NF", "Produto", "Peso", "Cliente", "CNPJ/CPF", "Base Produto", "Observação"]
        largura_coluna = [80, 50, 400, 100, 400, 130, 200, 120]  # ajuste as larguras conforme necessário
        for i, coluna in enumerate(colunas):
            self.tree.heading(coluna, text=coluna)
            self.tree.column(coluna, width=largura_coluna[i], anchor="center", stretch=True)
        self.tree.tag_configure("notification", foreground="red")
        self.tree.config(height=20)
        self.tree.bind("<Double-1>", self.mostrar_observacao)
        # Você pode vincular o evento de seleção conforme necessário
        self.tree.bind("<<TreeviewSelect>>", self.carregar_dados)

        # Configure o grid do container: a linha 0 (Treeview) expande e a linha 1 (label) não
        container_frame.rowconfigure(0, weight=1)
        container_frame.rowconfigure(1, weight=0)
        container_frame.columnconfigure(0, weight=1)

        #  FRAME PARA EXIBIR A SOMA
        soma_frame = tk.Frame(container_frame, bg="#f4f4f4", bd=0, relief="flat")
        soma_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(1, 0))

        # Label sem contorno: removendo bd e usando relief="flat"
        self.soma_label = tk.Label(soma_frame, text="Soma dos Pesos: 0,000 Kg", font=("Arial", 10, "bold"), bd=0, relief="flat")
        self.soma_label.pack(fill=tk.X)

        # Caso queira que o label de soma seja atualizado ao selecionar linhas,
        # vincule também um evento (por exemplo, com uma função que atualize a soma)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def obter_nomes_produtos(self):
        # Consulta ao banco de dados para retornar os nomes dos produtos
        conexao = conectar()
        cursor = conexao.cursor()
        cursor.execute("SELECT nome FROM produtos")
        resultados = cursor.fetchall()
        conexao.close()
        # Ordena alfabeticamente os nomes dos produtos
        return sorted([row[0] for row in resultados])

    def formatar_data(self, event=None):
        data = self.data_var.get().strip()

        # Remove qualquer caracter que não seja número
        data = ''.join(filter(str.isdigit, data))

        # Se a data tiver mais de 8 caracteres (depois de limpar), mantém apenas os 8 primeiros
        if len(data) > 8:
            data = data[:8]

        # Formata a data
        if len(data) == 8:
            data = data[:2] + '/' + data[2:4] + '/' + data[4:8]

        # Atualiza o campo de entrada com a data formatada
        self.data_var.set(data)

    def mostrar_observacao(self, event):
        """Exibe e permite a edição da observação completa em uma janela pop-up."""
        try:
            # Obter a lista de itens selecionados e usar o primeiro
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showwarning("Seleção Inválida", "Por favor, selecione um item na tabela.")
                return

            item = self.tree.item(selected_items[0])
            tags = item.get("tags", [])
            
            if tags:
                # A observação completa está armazenada na primeira tag
                obs_text = tags[0]

                # Criar a janela pop-up
                popup = tk.Toplevel(self)
                popup.title("Observação")
                popup.geometry("400x500")
                centralizar_janela(popup, 400, 500)  # Função auxiliar para centralizar a janela
                self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
                aplicar_icone(popup, self.caminho_icone)
                
                # Configurar o fundo da janela pop-up
                popup.config(bg="#ecf0f1")
                
                # Label para identificar a observação
                tk.Label(popup, text="Observação:", font=("Arial", 12, "bold"),
                        bg="#ecf0f1", fg="#34495e").pack(pady=10)
                
                # Widget de texto para exibir e editar a observação
                text_widget = tk.Text(popup, wrap="word", font=("Arial", 10))
                text_widget.insert("1.0", obs_text)
                text_widget.pack(expand=True, fill="both", padx=10, pady=10)
                
                # Frame para os botões com fundo neutro
                button_frame = tk.Frame(popup, bg="#ecf0f1")
                button_frame.pack(fill="x", pady=5)

                # Função para salvar a observação editada
                def salvar_observacao():
                    nova_observacao = text_widget.get("1.0", "end-1c").strip()
                    if nova_observacao:  # Certifica que não está vazia
                        try:
                            # Atualizar o banco de dados
                            nf_id = item["values"][1]  # Supondo que o número da NF seja a segunda coluna
                            self.cursor.execute('''
                                UPDATE nf
                                SET observacao = %s
                                WHERE numero_nf = %s
                            ''', (nova_observacao, str(nf_id)))
                            self.conn.commit()

                            # Atualizar a tag do item (a coluna exibirá a observação atualizada)
                            item["tags"] = (nova_observacao,)
                            self.atualizar_treeview()
                            self.limpar_entradas()

                            messagebox.showinfo("Sucesso", "Observação salva com sucesso!")
                            popup.destroy()
                        except Exception as e:
                            print(f"Erro ao salvar observação: {e}")
                            messagebox.showerror("Erro", f"Erro ao salvar observação: {e}")
                    else:
                        messagebox.showwarning("Aviso", "A observação não pode estar vazia.")

                # Função para remover a observação (definindo-a como vazia)
                def remover_observacao():
                    try:
                        nf_id = item["values"][1]
                        self.cursor.execute('''
                            UPDATE nf
                            SET observacao = %s
                            WHERE numero_nf = %s
                        ''', ("", str(nf_id)))
                        self.conn.commit()
                        
                        # Atualiza a tag para uma string vazia
                        item["tags"] = ("" ,)
                        self.atualizar_treeview()
                        self.limpar_entradas()

                        messagebox.showinfo("Sucesso", "Observação removida com sucesso!")
                        popup.destroy()
                    except Exception as e:
                        print(f"Erro ao remover observação: {e}")
                        messagebox.showerror("Erro", f"Erro ao remover observação: {e}")

                # Botão de salvar a observação editada (azul)
                tk.Button(button_frame, text="Salvar", command=salvar_observacao, font=("Arial", 10, "bold"),
                        relief="raised", padx=10, pady=5).pack(side="left", padx=10)
                
                # Botão para remover a observação (vermelho)
                tk.Button(button_frame, text="Remover", command=remover_observacao, font=("Arial", 10, "bold"),
                        relief="raised", padx=10, pady=5).pack(side="left", padx=10)
                
                # Botão de fechar a janela pop-up (tom escuro)
                tk.Button(button_frame, text="Fechar", command=lambda: (self.limpar_entradas(), popup.destroy()),font=("Arial", 10, "bold"),
                        relief="raised", padx=10, pady=5).pack(side="right", padx=10)
            else:
                messagebox.showinfo("Sem Observação", "Não há observação associada a este item.")
        except Exception as e:
            print(f"Erro ao exibir observação: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao tentar exibir a observação: {e}")

    def verificar_documento(self, event=None):
        # implementa sua lógica de busca, com raw = apenas dígitos
        doc_inserido = re.sub(r'\D', '', self.doc_var.get())
        if not doc_inserido:
            return
        for item in self.tree.get_children():
            valores = self.tree.item(item, "values")
            doc_na_linha = re.sub(r'\D', '', valores[5])
            if doc_inserido == doc_na_linha:
                self.cliente_var.set(valores[4])
                break
        else:
            self.cliente_var.set("")

    def formatar_documento(self, event):
        texto = self.doc_var.get()
        raw = re.sub(r'\D', '', texto)

        if len(raw) <= 11:
            raw = raw[:11]
            f = re.sub(r'^(\d{3})(\d)', r'\1.\2', raw)
            f = re.sub(r'^(\d{3}\.\d{3})(\d)', r'\1.\2', f)
            f = re.sub(r'^(\d{3}\.\d{3}\.\d{3})(\d)', r'\1-\2', f)
        else:
            raw = raw[:14]
            f = re.sub(r'^(\d{2})(\d)', r'\1.\2', raw)
            f = re.sub(r'^(\d{2}\.\d{3})(\d)', r'\1.\2', f)
            f = re.sub(r'^(\d{2}\.\d{3}\.\d{3})(\d)', r'\1/\2', f)
            f = re.sub(r'^(\d{2}\.\d{3}\.\d{3}\/\d{4})(\d)', r'\1-\2', f)

        self.doc_var.set(f)

        # -- Opcional: auto-foco no próximo widget --
        event.widget.tk_focusNext().focus_set()

    def formatar_numero(self, event, casas_decimais=3):
        entry = event.widget
        texto = entry.get()
        # mantém apenas dígitos
        apenas_digitos = ''.join(filter(str.isdigit, texto))
        if not apenas_digitos:
            entry.delete(0, tk.END)
            return
        # desloca para criar as casas decimais
        valor_int = int(apenas_digitos)
        valor = Decimal(valor_int) / (10 ** casas_decimais)
        # formata com separador de milhares e vírgula decimal
        s = f"{valor:,.{casas_decimais}f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        entry.delete(0, tk.END)
        entry.insert(0, s)

    def _pre_formatar(self, valor_str, casas_decimais=3):
        """Formata um valor_str já existente (com vírgula ou ponto) para pt‑BR."""
        text = valor_str.strip()
        try:
            # Se vier "1.234,567" (tanto "." quanto ","), tira os pontos de milhar
            if "." in text and "," in text:
                raw = text.replace(".", "").replace(",", ".")
            # Se vier apenas "50,000" (vírgula única), assume vírgula decimal
            elif "," in text:
                raw = text.replace(",", ".")
            # Se vier "1234.567" (ponto único), assume ponto decimal
            else:
                raw = text
            num = float(raw)
        except ValueError:
            return valor_str

        # Garante o formato com separador de milhar "." e vírgula decimal
        inteiros, decimais = f"{num:.{casas_decimais}f}".split(".")
        inteiros = f"{int(inteiros):,}".replace(",", ".")
        return f"{inteiros},{decimais}"

    def limpar_entradas(self):
        # Limpa os campos de entrada
        self.data_var.set("")
        self.nf_var.set("")
        self.cliente_var.set("")
        self.doc_var.set("")
        self.observacao_var.set("")
        self.produto_entry.set("")
        self.peso_entry.set("")
        self.observacao_text.delete(1.0, tk.END)  # Limpa o campo de texto
        self.lista_produtos.delete(0, tk.END)  # Limpa a lista de produtos
        self.variavel_busca.set("")
        self.base_produto_var.set("")

        # Restaura a exibição de todas as informações ocultas
        self.atualizar_treeview()
        print("Entradas limpas.")

    def atualizar_treeview(self):
        try:
            self.cursor.execute('''        
                SELECT nf.id, nf.data, nf.numero_nf, nf.cliente, nf.cnpj_cpf, nf.observacao,
                    string_agg(CONCAT(p.produto_nome, '; ', p.peso, '; ', COALESCE(p.base_produto, '')), ', ') as produtos
                FROM nf nf
                LEFT JOIN produtos_nf p ON nf.id = p.nf_id
                GROUP BY nf.id
                ORDER BY nf.data DESC, nf.numero_nf DESC;
            ''')
            rows = self.cursor.fetchall()

            # Limpa o Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            for row in rows:
                data_formatada = row[1].strftime("%d/%m/%Y") if row[1] else ""
                numero_nf = row[2]
                cliente = row[3].strip()  # Remover espaços extras
                cnpj_cpf = row[4]
                observacao = row[5] or ""
                produtos_str = row[6]  # String concatenada de produtos

                # Se houver observação, exibe apenas "Notificação" na coluna, mas guarda o texto completo na tag
                if observacao.strip():
                    observacao_exibida = "Notificação"
                    # guarda o texto real *e* marca de estilo
                    tags = (observacao, "notification")
                else:
                    observacao_exibida = ""
                    tags = ()

                # Cria a lista de produtos com os dados: nome, peso e base_produto
                produtos_lista = []
                for produto in produtos_str.split(', ') if produtos_str else []:
                    # Cada produto está no formato "produto_nome; peso; base_produto"
                    parts = produto.split('; ')
                    if len(parts) == 3:
                        nome, peso, base_produto = parts
                    else:
                        # Tratamento caso não haja todos os três valores
                        nome = parts[0]
                        peso = parts[1] if len(parts) > 1 else ""
                        base_produto = ""
                    try:
                        if peso:
                            peso_formatada = f"{float(peso.replace(',', '.')):.3f}".replace('.', ',')
                            peso_formatada += " Kg"
                        else:
                            peso_formatada = "0,000 Kg"
                    except ValueError:
                        peso_formatada = "Erro"
                    produtos_lista.append((nome, peso_formatada, base_produto))

                # Insere uma linha para cada produto, agora com a coluna "Base Produto" incluída
                for produto in produtos_lista:
                    self.tree.insert("", "end", values=(
                        data_formatada, numero_nf, produto[0], produto[1], cliente, cnpj_cpf, produto[2], observacao_exibida
                    ), tags=tags)
        except Exception as e:
            print(f"Erro ao atualizar o Treeview: {e}")
            messagebox.showerror("Erro", f"Erro ao atualizar a lista de notas fiscais: {e}")

    def configurar_treeview(self):
        """Configura o estilo do Treeview e define 10 linhas visíveis."""
        style = ttk.Style()
        style.theme_use("alt")  # ou outro tema que preferir
        style.configure("Treeview", 
                        background="white", 
                        foreground="black", 
                        rowheight=20,  # altura de cada linha (ajuste se necessário)
                        fieldbackground="white")
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.map("Treeview", 
                background=[("selected", "#0078D7")], 
                foreground=[("selected", "white")])
        
        # Define que o Treeview exibirá 10 linhas visíveis
        self.tree.config(height=10)

    def carregar_dados(self, event):
        # Obtém os itens selecionados
        selected_items = self.tree.selection()

        if selected_items:
            for selected_item in selected_items:
                item = self.tree.item(selected_item)
                values = item["values"]

                self.data_var.set(values[0])
                self.nf_var.set(values[1])
                self.produto_entry.set(values[2])
                # Remove o " Kg" da peso e mantém a vírgula
                peso_sem_kg = values[3].replace(" Kg", "")
                self.peso_entry.set(peso_sem_kg)
                self.cliente_var.set(values[4].strip())
                self.doc_var.set(values[5])
                # Define a variável da combobox "Base Produto"
                self.base_produto_var.set(values[6])
                tags = item.get("tags", [])
                # pega tudo que NÃO for a tag de estilo
                observacao_real = next((t for t in tags if t != "notification"), "")
                self.observacao_text.delete("1.0", "end")
                self.observacao_text.insert("1.0", observacao_real)

    def buscar_menor_id_disponivel(self, tabela):
        self.cursor.execute(f'''
            SELECT COALESCE(MIN(id) + 1, 1) FROM {tabela}
            WHERE id + 1 NOT IN (SELECT id FROM {tabela})
        ''')
        menor_id = self.cursor.fetchone()[0]
        return menor_id

    def ao_perder_foco(self, event):
        """
        Formata o conteúdo do campo de busca somente quando o usuário sai do campo
        ou pressiona Enter. Se o valor já estiver no formato dd/mm/yyyy, nada faz.
        Agora também formata CPF (11 dígitos), CNPJ (14 dígitos) e data (8 dígitos).
        """
        conteudo = self.variavel_busca.get().strip()

        # Se já estiver no formato "dd/mm/aaaa" válido, sai sem mudar
        if re.fullmatch(r"\d{2}/\d{2}/\d{4}", conteudo):
            try:
                datetime.strptime(conteudo, "%d/%m/%Y")
                return
            except ValueError:
                pass

        # Extrai só dígitos
        digitos = ''.join(ch for ch in conteudo if ch.isdigit())
        if not digitos:
            return

        # Data (8 dígitos) → dd/mm/aaaa
        if len(digitos) == 8:
            novo = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
        # CPF (11 dígitos) → 000.000.000-00
        elif len(digitos) == 11:
            novo = (
                digitos[:3] + '.'
                + digitos[3:6] + '.'
                + digitos[6:9] + '-'
                + digitos[9:]
            )
        # CNPJ (14 dígitos) → 00.000.000/0000-00
        elif len(digitos) == 14:
            novo = (
                digitos[:2] + '.'
                + digitos[2:5] + '.'
                + digitos[5:8] + '/'
                + digitos[8:12] + '-'
                + digitos[12:]
            )
        else:
            # Se não é data, CPF nem CNPJ, mantém o original
            novo = conteudo

        # Atualiza somente se realmente mudou
        if novo != conteudo:
            self.variavel_busca.set(novo)

    def formatar_termo_busca(self, termo):
        """
        Formata o termo de busca:
         - Se tiver 8 dígitos, formata como data (dd/mm/aaaa).
         - Se tiver 14 dígitos, formata como CNPJ (00.000.000/0000-00).
         - Caso contrário, retorna o termo original.
        """
        termo_digits = ''.join(ch for ch in termo if ch.isdigit())
        if len(termo_digits) == 8:
            return f"{termo_digits[:2]}/{termo_digits[2:4]}/{termo_digits[4:]}", "date"
        elif len(termo_digits) == 14:
            return f"{termo_digits[:2]}.{termo_digits[2:5]}.{termo_digits[5:8]}/{termo_digits[8:12]}-{termo_digits[12:]}", "cnpj"
        else:
            return termo, "text"

    def buscar_nf(self):
        """
        Realiza a busca das Notas Fiscais conforme o termo digitado,
        incluindo suporte para buscar por "Notificação". Ordena por data (desc)
        e, quando a data for igual, ordena pelo número da NF (desc).
        """
        # Garante que a formatação seja aplicada antes da busca
        self.ao_perder_foco(None)
        termo_original = self.variavel_busca.get().strip()
        if not termo_original:
            self.atualizar_treeview()
            return

        termo, termo_tipo = self.formatar_termo_busca(termo_original)
        print(f"Termo original: {termo_original} | Formatado: {termo} | Tipo: {termo_tipo}")

        try:
            # Caso especial: mostrar todas as NFs com observação (Notificação)
            if termo_original.lower() == "notificação":
                query = '''
                    SELECT nf.id, nf.data, nf.numero_nf, nf.cliente, nf.cnpj_cpf, nf.observacao,
                        string_agg(
                            CONCAT(p.produto_nome, '; ', p.peso, '; ', COALESCE(p.base_produto, '')),
                            ', '
                        ) AS produtos
                    FROM nf nf
                    LEFT JOIN produtos_nf p ON nf.id = p.nf_id
                    WHERE nf.observacao IS NOT NULL 
                    AND TRIM(nf.observacao) <> ''
                    GROUP BY nf.id
                    ORDER BY nf.data DESC, 
                            /* se numero_nf for texto, remova ::integer ou ajuste conforme necessário */
                            nf.numero_nf::integer DESC;
                '''
                params = []
            else:
                termo_data_sql = None
                data_filtro = False

                if termo_tipo == "date":
                    try:
                        termo_data = datetime.strptime(termo, "%d/%m/%Y")
                        termo_data_sql = termo_data.strftime("%Y-%m-%d")
                        data_filtro = True
                    except ValueError:
                        termo_tipo = "text"

                if data_filtro:
                    query = '''
                        SELECT nf.id, nf.data, nf.numero_nf, nf.cliente, nf.cnpj_cpf, nf.observacao,
                            string_agg(
                                CONCAT(p.produto_nome, '; ', p.peso, '; ', COALESCE(p.base_produto, '')),
                                ', '
                            ) AS produtos
                        FROM nf nf
                        LEFT JOIN produtos_nf p ON nf.id = p.nf_id
                        WHERE nf.data = %s
                        GROUP BY nf.id
                        ORDER BY nf.data DESC, 
                                nf.numero_nf::integer DESC;
                    '''
                    params = [termo_data_sql]
                else:
                    query = '''
                        SELECT nf.id, nf.data, nf.numero_nf, nf.cliente, nf.cnpj_cpf, nf.observacao,
                            string_agg(
                                CONCAT(p.produto_nome, '; ', p.peso, '; ', COALESCE(p.base_produto, '')),
                                ', '
                            ) AS produtos
                        FROM nf nf
                        LEFT JOIN produtos_nf p ON nf.id = p.nf_id
                        WHERE
                            unaccent(nf.numero_nf) ILIKE unaccent(%s) OR
                            unaccent(nf.cliente) ILIKE unaccent(%s) OR
                            unaccent(nf.cnpj_cpf) ILIKE unaccent(%s) OR
                            unaccent(p.produto_nome) ILIKE unaccent(%s) OR
                            unaccent(p.base_produto) ILIKE unaccent(%s)
                        GROUP BY nf.id
                        ORDER BY nf.data DESC, 
                                nf.numero_nf::integer DESC;
                    '''
                    params = [
                        f"%{termo}%", f"%{termo}%", f"%{termo}%", f"%{termo}%", f"%{termo}%"
                    ]

            print("Query SQL:", query)
            print("Parâmetros:", params)

            self.cursor.execute(query, tuple(params))
            rows = self.cursor.fetchall()

            # Limpa a treeview antes de inserir novos itens
            for item in self.tree.get_children():
                self.tree.delete(item)

            if rows:
                for row in rows:
                    data_formatada = row[1].strftime("%d/%m/%Y") if row[1] else ""
                    observacao = row[5] or ""
                    if observacao.strip():
                        observacao_exibida = "Notificação"
                        tags = (observacao, "notification")
                    else:
                        observacao_exibida = ""
                        tags = ()

                    produtos_str = row[6] or ""
                    produtos_lista = []
                    for produto in produtos_str.split(', '):
                        parts = produto.split('; ')
                        if len(parts) == 3:
                            nome, peso, base_produto = parts
                        else:
                            nome = parts[0]
                            peso = parts[1] if len(parts) > 1 else ""
                            base_produto = ""
                        peso = peso.strip()
                        try:
                            if peso.lower().endswith("kg"):
                                unit = " Kg"
                                numeric_part = peso[:-2].strip()
                                num = float(numeric_part.replace(',', '.'))
                                peso_formatada = "{:,.3f}".format(num)\
                                    .replace(",", "X")\
                                    .replace(".", ",")\
                                    .replace("X", ".") + unit
                            else:
                                num = float(peso.replace(',', '.'))
                                peso_formatada = "{:,.3f}".format(num)\
                                    .replace(",", "X")\
                                    .replace(".", ",")\
                                    .replace("X", ".") + " Kg"
                        except ValueError:
                            peso_formatada = "Erro"
                        produtos_lista.append((nome, peso_formatada, base_produto))

                    for produto in produtos_lista:
                        self.tree.insert(
                            "",
                            "end",
                            values=(
                                data_formatada,
                                row[2],        # número da NF
                                produto[0],    # nome do produto
                                produto[1],    # peso formatado
                                row[3],        # cliente
                                row[4],        # cnpj_cpf
                                produto[2],    # base_produto
                                observacao_exibida
                            ),
                            tags=tags
                        )
            else:
                messagebox.showinfo(
                    "Resultado da pesquisa",
                    "Nenhuma nota fiscal encontrada com o termo informado."
                )

        except Exception as e:
            print(f"Erro ao buscar NF: {e}")
            messagebox.showerror("Erro", f"Erro ao realizar a busca: {e}")

    def abrir_dialogo_exportacao(self):
        """
        Abre uma janela de diálogo para que o usuário informe os filtros de exportação,
        incluindo um intervalo de datas, NF, Cliente, CNPJ/CPF, produto e base produto.
        """
        # Cria a janela de diálogo (Toplevel)
        dialogo = tk.Toplevel(self)
        dialogo.title("Exportar NF - Filtros")
        dialogo.geometry("400x300")
        dialogo.resizable(False, False)
        # Centraliza a janela usando sua função importada
        centralizar_janela(dialogo, 400, 300)
        # Aplica o ícone (certifique-se de que a função 'aplicar_icone' esteja definida/importada)
        aplicar_icone(dialogo, self.caminho_icone)
        dialogo.config(bg="#ecf0f1")  # Fundo neutro

        # Configuração dos estilos personalizados utilizando ttk
        style = ttk.Style(dialogo)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])
        
        # Cria um frame principal com padding para organizar os widgets
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
            # Formata a data conforme a peso de caracteres
            if len(data) >= 2:
                data = data[:2] + '/' + data[2:]
            if len(data) >= 5:
                data = data[:5] + '/' + data[5:]
            # Atualiza o campo de entrada com a data formatada
            entry.delete(0, tk.END)
            entry.insert(0, data)

        # Função para formatar CNPJ enquanto digita
        def formatar_documento(event):
            entry = event.widget
            texto = entry.get()
            # Só dígitos
            raw = re.sub(r'\D', '', texto)

            # CPF: até 11 dígitos
            if len(raw) <= 11:
                raw = raw[:11]
                f = re.sub(r'^(\d{3})(\d)', r'\1.\2', raw)
                f = re.sub(r'^(\d{3}\.\d{3})(\d)', r'\1.\2', f)
                f = re.sub(r'^(\d{3}\.\d{3}\.\d{3})(\d)', r'\1-\2', f)
            # CNPJ: de 12 a 14 dígitos
            else:
                raw = raw[:14]
                f = re.sub(r'^(\d{2})(\d)', r'\1.\2', raw)
                f = re.sub(r'^(\d{2}\.\d{3})(\d)', r'\1.\2', f)
                f = re.sub(r'^(\d{2}\.\d{3}\.\d{3})(\d)', r'\1/\2', f)
                f = re.sub(r'^(\d{2}\.\d{3}\.\d{3}\/\d{4})(\d)', r'\1-\2', f)

            # Atualiza o campo só se mudou
            if f != texto:
                entry.delete(0, tk.END)
                entry.insert(0, f)

        # Linha 0: Data Inicial
        ttk.Label(frame, text="Data Inicial (dd/mm/yyyy):", style="Custom.TLabel")\
            .grid(row=0, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_data_inicial = ttk.Entry(frame, width=25)
        entry_data_inicial.grid(row=0, column=1, sticky="w", padx=5, pady=(5, 2))
        entry_data_inicial.bind("<KeyRelease>", formatar_data)

        # Linha 1: Data Final
        ttk.Label(frame, text="Data Final (dd/mm/yyyy):", style="Custom.TLabel")\
            .grid(row=1, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_data_final = ttk.Entry(frame, width=25)
        entry_data_final.grid(row=1, column=1, sticky="w", padx=5, pady=(5, 2))
        entry_data_final.bind("<KeyRelease>", formatar_data)

        # Linha 2: NF
        ttk.Label(frame, text="NF:", style="Custom.TLabel")\
            .grid(row=2, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_nf = ttk.Entry(frame, width=25)
        entry_nf.grid(row=2, column=1, sticky="w", padx=5, pady=(5, 2))
        
        # Linha 3: Cliente
        ttk.Label(frame, text="Cliente:", style="Custom.TLabel")\
            .grid(row=3, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_cliente = ttk.Entry(frame, width=35)
        entry_cliente.grid(row=3, column=1, sticky="w", padx=5, pady=(5, 2))
        
        # Linha 4: CNPJ
        self.doc_var_2 = tk.StringVar()
        ttk.Label(frame, text="CNPJ / CPF:", style="Custom.TLabel")\
            .grid(row=4, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_cnpj_cpf_2 = ttk.Entry(frame, textvariable=self.doc_var_2, width=25)
        entry_cnpj_cpf_2.grid(row=4, column=1, sticky="w", padx=5, pady=(5, 2))
        entry_cnpj_cpf_2.bind("<KeyRelease>", formatar_documento)
        
        # Linha 5: Produto
        ttk.Label(frame, text="Produto:", style="Custom.TLabel")\
            .grid(row=5, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_produto = ttk.Entry(frame, width=35)
        entry_produto.grid(row=5, column=1, sticky="w", padx=5, pady=(5, 2))
        
        # Linha 6: Base Produto
        ttk.Label(frame, text="Base Produto:", style="Custom.TLabel")\
            .grid(row=6, column=0, sticky="e", padx=5, pady=(5, 2))
        entry_base_produto = ttk.Entry(frame, width=35)
        entry_base_produto.grid(row=6, column=1, sticky="w", padx=5, pady=(5, 2))
        
        # Linha 7: Botões de Exportar e Cancelar
        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.grid(row=7, column=0, columnspan=2, pady=(15, 5))
        
        def acao_exportar():
            # Coleta os valores informados no diálogo
            filtro_data_inicial = entry_data_inicial.get().strip()
            filtro_data_final = entry_data_final.get().strip()
            filtro_nf = entry_nf.get().strip()
            filtro_cliente = entry_cliente.get().strip()
            filtro_cnpj_cpf = entry_cnpj_cpf_2.get().strip()
            filtro_produto = entry_produto.get().strip()
            filtro_base_produto = entry_base_produto.get().strip()
            
            # Lista para armazenar as cláusulas WHERE e os parâmetros
            where_clauses = []
            parametros = []
            
            # Processamento do intervalo de datas
            if filtro_data_inicial and filtro_data_final:
                try:
                    data_inicial = datetime.strptime(filtro_data_inicial, '%d/%m/%Y')
                    data_final = datetime.strptime(filtro_data_final, '%d/%m/%Y')
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido. Utilize dd/mm/yyyy.")
                    return
                filtro_data_inicial_fmt = data_inicial.strftime('%Y-%m-%d')
                filtro_data_final_fmt = data_final.strftime('%Y-%m-%d')
                where_clauses.append("nf.data BETWEEN %s AND %s")
                parametros.extend([filtro_data_inicial_fmt, filtro_data_final_fmt])
            elif filtro_data_inicial:
                try:
                    data_inicial = datetime.strptime(filtro_data_inicial, '%d/%m/%Y')
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido. Utilize dd/mm/yyyy.")
                    return
                filtro_data_inicial_fmt = data_inicial.strftime('%Y-%m-%d')
                where_clauses.append("nf.data >= %s")
                parametros.append(filtro_data_inicial_fmt)
            elif filtro_data_final:
                try:
                    data_final = datetime.strptime(filtro_data_final, '%d/%m/%Y')
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido. Utilize dd/mm/yyyy.")
                    return
                filtro_data_final_fmt = data_final.strftime('%Y-%m-%d')
                where_clauses.append("nf.data <= %s")
                parametros.append(filtro_data_final_fmt)
            
            # Filtragem para NF
            if filtro_nf:
                where_clauses.append("nf.numero_nf ILIKE %s")
                parametros.append(f"%{filtro_nf}%")
            
            # Filtragem para Cliente
            if filtro_cliente:
                where_clauses.append("unaccent(nf.cliente) ILIKE unaccent(%s)")
                parametros.append(f"%{filtro_cliente}%")
            
            # Filtragem para CNPJ
            if filtro_cnpj_cpf:
                where_clauses.append("nf.cnpj_cpf ILIKE %s")
                parametros.append(f"%{filtro_cnpj_cpf}%")
            
            # Filtragem para Produto
            if filtro_produto:
                where_clauses.append("unaccent(produtos_nf.produto_nome) ILIKE unaccent(%s)")
                parametros.append(f"%{filtro_produto}%")
            
            # Filtragem para Base Produto
            if filtro_base_produto:
                where_clauses.append("unaccent(produtos_nf.base_produto) ILIKE unaccent(%s)")
                parametros.append(f"%{filtro_base_produto}%")
            
            # Monta a cláusula WHERE se houver algum filtro
            where_sql = " AND ".join(where_clauses)
            if where_sql:
                where_sql = "WHERE " + where_sql
            
            # Monta a consulta SQL, incluindo a coluna base_produto
            sql = f"""
                SELECT nf.data, nf.numero_nf, produtos_nf.produto_nome, produtos_nf.peso,
                    nf.cliente, nf.cnpj_cpf, produtos_nf.base_produto, nf.observacao
                FROM nf
                JOIN produtos_nf ON nf.id = produtos_nf.nf_id
                {where_sql}
                ORDER BY nf.data ASC, nf.numero_nf ASC;
            """
            
            # Solicita ao usuário o caminho e nome do arquivo para salvar o Excel
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Relatorio_Saida_Nf.xlsx"
            )
            
            if caminho_arquivo:
                try:
                    self.cursor.execute(sql, parametros)
                    dados = self.cursor.fetchall()
                    
                    # Cria o DataFrame com os nomes das colunas conforme a consulta
                    df = pd.DataFrame(dados, columns=["data", "numero_nf", "produto_nome", "peso","cliente", "cnpj_cpf", "base_produto", "observacao"])
                    # Reordena as colunas para a ordem desejada:
                    df = df[["data", "numero_nf", "produto_nome", "peso","cliente", "cnpj_cpf", "base_produto", "observacao"]]
                    # Renomeia as colunas para exibição final no Excel
                    df.rename(columns={
                        "data": "Data",
                        "numero_nf": "NF",
                        "produto_nome": "Produto",
                        "peso": "Peso",
                        "cliente": "Cliente",
                        "cnpj_cpf": "CNPJ/CPF",
                        "base_produto": "Base Produto",
                        "observacao": "Observação"
                    }, inplace=True)
                    
                    # Exporta para Excel
                    df.to_excel(caminho_arquivo, index=False)
                    messagebox.showinfo("Exportação concluída", f"Arquivo Excel salvo em {caminho_arquivo}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")
                
                dialogo.destroy()
        
        # Botões de Exportar e Cancelar
        export_button = ttk.Button(button_frame, text="Exportar Excel", command=acao_exportar)
        export_button.grid(row=0, column=0, padx=5)
        
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=dialogo.destroy)
        cancel_button.grid(row=0, column=1, padx=5)
        
        # Configura o grid do frame para que a coluna dos campos de entrada expanda
        frame.columnconfigure(1, weight=1)

    def adicionar_produto(self):
        produto = self.produto_entry.get().strip()  # Remove espaços extras
        peso = self.peso_entry.get().strip()        # Remove espaços extras
        base_produto = self.base_produto_var.get().strip()  # Valor selecionado na combobox

        if produto and peso and base_produto:
            try:
                # 1. Remove separadores de milhar (pontos) e troca vírgula decimal por ponto
                peso_normalizado = peso.replace('.', '').replace(',', '.')
                # 2. Converte para float com um único argumento
                peso_valor = float(peso_normalizado)

                if peso_valor > 0:
                    # Adiciona à lista interna
                    self.produtos.append((produto, peso_valor, base_produto))

                    # Formata o peso no padrão brasileiro
                    peso_formatado = self._pre_formatar(str(peso_valor), 3)

                    # Atualiza a Listbox
                    self.lista_produtos.insert(tk.END, f"{produto} - {peso_formatado} - {base_produto}")

                    # Limpa os campos corretamente
                    self.produto_entry.set("")   # use a StringVar
                    self.peso_entry.set("")      # idem
                    self.base_produto_var.set("")
                else:
                    messagebox.showwarning("Erro", "O peso deve ser um número positivo.")
            except ValueError:
                messagebox.showwarning("Erro", f"Peso inválido: '{peso}'")
        else:
            messagebox.showwarning("Erro", "Por favor, preencha produto, peso e base produto.")

    def alterar_produto(self):
        # Verifica se algum item foi selecionado
        selected_item_index = self.lista_produtos.curselection()
        if not selected_item_index:
            tk.messagebox.showwarning("Aviso", "Selecione um produto para alterar!")
            return

        # Obtém o item selecionado
        selected_item = self.lista_produtos.get(selected_item_index)
        # Verifica se o item segue o formato esperado: "produto - peso - base_produto"
        parts = selected_item.split(" - ")
        if len(parts) == 3:
            produto_atual, peso_atual, base_produto_atual = parts
        else:
            tk.messagebox.showerror("Erro", "Formato de item desconhecido!")
            return

        # Cria uma janela Toplevel para a alteração dos dados
        top_alterar = tk.Toplevel(self)
        top_alterar.title("Alterar Produto")
        top_alterar.geometry("450x200")
        top_alterar.resizable(False, False)
        centralizar_janela(top_alterar, 450, 200)  # Certifique-se de que essa função esteja definida/importada
        aplicar_icone(top_alterar, self.caminho_icone)
        top_alterar.config(bg="#ecf0f1")  # Fundo neutro

        # Configuração dos estilos personalizados
        style = ttk.Style(top_alterar)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Alter.TButton", background="#2980b9", foreground="white",
                        font=("Arial", 10, "bold"), padding=5)
        style.map("Alter.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        # Cria um frame principal com padding para organizar os widgets
        frame = ttk.Frame(top_alterar, padding="15 15 15 15", style="Custom.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        top_alterar.columnconfigure(0, weight=1)
        top_alterar.rowconfigure(0, weight=1)

        # Linha 0: Rótulo e entrada para o novo nome do produto
        ttk.Label(frame, text="Novo Nome do Produto:", style="Custom.TLabel")\
            .grid(row=0, column=0, sticky="w", padx=5, pady=(5, 2))
        produto_entry = ttk.Entry(frame, width=25)
        produto_entry.insert(0, produto_atual)
        produto_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=(5, 2))
        frame.columnconfigure(1, weight=1)

        # Linha 1: Rótulo e entrada para a nova peso
        ttk.Label(frame, text="Novo Peso:", style="Custom.TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=(5, 2))
        peso_entry = ttk.Entry(frame, width=25)
        peso_entry.insert(0, self._pre_formatar(peso_atual, 3))
        peso_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=(5, 2))
        # bind no mesmo método de formatação
        peso_entry.bind("<KeyRelease>", lambda e: self.formatar_numero(e, 3))

        # Linha 2: Rótulo e combobox/entrada para a base do produto
        ttk.Label(frame, text="Base Produto:", style="Custom.TLabel")\
            .grid(row=2, column=0, sticky="w", padx=5, pady=(5, 2))
        # Aqui usamos uma Combobox para manter a consistência com a tela principal.
        base_produto_combobox = ttk.Combobox(frame, width=23, textvariable=tk.StringVar())
        # Se você tiver uma lista de nomes para a base do produto, atribua a ela.
        base_produto_combobox['values'] = self.lista_nomes_produtos if hasattr(self, 'lista_nomes_produtos') else []
        base_produto_combobox.set(base_produto_atual)
        base_produto_combobox.grid(row=2, column=1, sticky="ew", padx=5, pady=(5, 2))

        # Linha 3: Frame para os botões de Confirmar e Cancelar
        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.grid(row=3, column=0, columnspan=2, pady=(15, 5))

        def confirmar_alteracao():
            novo_produto = produto_entry.get().strip()
            novo_peso_str = peso_entry.get().strip()
            nova_base_produto = base_produto_combobox.get().strip()

            if not novo_produto:
                tk.messagebox.showwarning("Aviso", "Nome do produto não pode estar vazio!")
                return

            try:
                # Converte a peso para float, tratando vírgula como separador decimal
                novo_peso = float(novo_peso_str.replace(',', '.'))
                # e formata de volta pra armazenar no listbox
                peso_formatado = self._pre_formatar(str(novo_peso), 3)
            except ValueError:
                tk.messagebox.showerror("Erro", "Peso inválida!")
                return

            if not nova_base_produto:
                tk.messagebox.showwarning("Aviso", "Base Produto não pode estar vazio!")
                return

            # Atualiza o item na Listbox e na lista interna de produtos
            self.lista_produtos.delete(selected_item_index)
            novo_item = f"{novo_produto} - {peso_formatado} - {nova_base_produto}"
            self.lista_produtos.insert(selected_item_index, novo_item)
            # Atualiza o registro na lista interna; assume que self.produtos armazena tuplas (produto, peso, base_produto)
            self.produtos[selected_item_index[0]] = (novo_produto, novo_peso, nova_base_produto)

            tk.messagebox.showinfo("Alteração concluída", "Produto, peso e base alterados com sucesso!")
            top_alterar.destroy()

        def cancelar():
            top_alterar.destroy()

        # Botão Confirmar com estilo personalizado
        confirm_button = ttk.Button(button_frame, text="Confirmar", command=confirmar_alteracao, style="Alter.TButton")
        confirm_button.grid(row=0, column=0, padx=5)

        # Botão Cancelar com o mesmo estilo
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=cancelar, style="Alter.TButton")
        cancel_button.grid(row=0, column=1, padx=5)

    def excluir_produto(self):
        selected_item_index = self.lista_produtos.curselection()
        if not selected_item_index:
            messagebox.showwarning("Aviso", "Selecione um produto para excluir!")
            return

        idx = selected_item_index[0]
        item_text = self.lista_produtos.get(idx)
        parts = item_text.split(" - ")
        if len(parts) != 3:
            messagebox.showerror("Erro", "Formato do item inválido: use 'produto_nome - peso - base_produto'.")
            return

        produto_excluido = parts[0]
        quant_str = parts[1]
        base_produto_excluido = parts[2]
        # Normaliza formato brasileiro: remove pontos de milhares e troca vírgula por ponto
        quant_str_norm = quant_str.replace(".", "").replace(",", ".")
        try:
            peso_excluido = float(quant_str_norm)
        except ValueError:
            messagebox.showerror("Erro", f"Não foi possível converter peso '{quant_str}' para número.")
            return

        confirmacao = messagebox.askyesno("Confirmar Exclusão", f"Excluir {produto_excluido} ({quant_str})?")
        if not confirmacao:
            return
        removido_interno = False
        for i, (p, q, bp) in enumerate(self.produtos):
            if p == produto_excluido and bp == base_produto_excluido and abs(q - peso_excluido) < 1e-6:
                del self.produtos[i]
                removido_interno = True
                break

        if not removido_interno:
            try:
                self.cursor.execute(
                    "DELETE FROM produtos_nf WHERE produto_nome = ? AND peso = ? AND base_produto = ?",
                    (produto_excluido, peso_excluido, base_produto_excluido)
                )
                self.conn.commit()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir do banco: {e}")
                self.conn.rollback()
                return

        self.lista_produtos.delete(idx)

    def salvar_nf(self):
        # Verifica se os campos obrigatórios (exceto observação) foram preenchidos
        if (not self.data_var.get().strip() or
            not self.nf_var.get().strip() or
            not self.cliente_var.get().strip() or
            not self.doc_var.get().strip()):
            messagebox.showwarning("Campos Incompletos", "Por favor, preencha todos os campos obrigatórios (exceto observação).")
            return

        # Verifica se ao menos um produto foi adicionado
        if not self.produtos:
            messagebox.showwarning("Produtos", "Adicione pelo menos um produto à nota fiscal.")
            return

        # Verifica se todos os produtos possuem o campo "Base Produto" preenchido
        for produto, peso, base_produto in self.produtos:
            if not base_produto.strip():
                messagebox.showwarning("Campos Incompletos", "O campo 'Base Produto' é obrigatório para todos os produtos.")
                return

        try:
            # Obter o menor ID disponível para a tabela nf
            menor_id_nf = self.buscar_menor_id_disponivel("nf")

            # Coletar dados da interface
            data = self.data_var.get().strip()
            numero_nf = self.nf_var.get().strip()
            cliente = self.cliente_var.get().strip()
            doc = self.doc_var.get()
            observacao = self.observacao_text.get("1.0", "end-1c")  # Pode estar vazia

            # Inserir a nova nota fiscal com o menor ID disponível para nf
            self.cursor.execute('''
                INSERT INTO nf (id, data, numero_nf, cliente, cnpj_cpf, observacao)
                VALUES (%s, %s, %s, %s, %s, %s)
            ''', (menor_id_nf, data, numero_nf, cliente, doc, observacao))

            # Confirmar a inserção da nota fiscal
            self.conn.commit()

            # Inserir produtos associados à NF com o menor ID disponível para produtos_nf
            # Cada produto na lista deve ser uma tupla (produto, peso, base_produto)
            for produto, peso, base_produto in self.produtos:
                menor_id_produto = self.buscar_menor_id_disponivel("produtos_nf")
                self.cursor.execute('''
                    INSERT INTO produtos_nf (id, nf_id, produto_nome, peso, base_produto)
                    VALUES (%s, %s, %s, %s, %s)
                ''', (menor_id_produto, menor_id_nf, produto, peso, base_produto))

            # Confirmar a inserção dos produtos
            self.conn.commit()

            self.cursor.execute(
                "NOTIFY canal_atualizacao, 'menu_atualizado';"
            )
            # como já fizemos commit acima, basta outro commit para o NOTIFY
            self.conn.commit()

            # Atualizar Treeview
            self.atualizar_treeview()

            # Limpar lista de produtos, campos de entrada e lista de exibição
            self.produtos.clear()
            self.lista_produtos.delete(0, tk.END)

            # Limpar campos de entrada
            self.data_var.set("")
            self.nf_var.set("")
            self.cliente_var.set("")
            self.doc_var.set("")
            self.observacao_text.delete("1.0", tk.END)

            messagebox.showinfo("Sucesso", "Nota Fiscal salva com sucesso!")

        except Exception as e:
            print(f"Erro ao salvar NF: {e}")
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao salvar NF: {e}")

    def alterar_valores(self):
        selected_item = self.tree.selection()
        if selected_item:
            item = self.tree.item(selected_item)
            values = item["values"]

            # Obter os dados alterados
            nova_data = self.data_var.get().strip()
            novo_nf = self.nf_var.get().strip()
            novo_cliente = self.cliente_var.get().strip()
            novo_doc = self.doc_var.get().strip()
            nova_observacao = self.observacao_text.get("1.0", "end-1c")  # Captura o texto da Observação
            produto = self.produto_entry.get().strip()
            peso_str = self.peso_entry.get().strip()
            base_produto = self.base_produto_var.get().strip()  # Captura o valor da combobox "Base Produto"

            # Substituir a vírgula pelo ponto para aceitar decimais com vírgula
            peso_str = peso_str.replace(',', '.')
            try:
                peso = float(peso_str)
            except ValueError:
                messagebox.showerror("Erro", "Valor da peso inválido")
                return

            # Atualizar os valores no banco de dados
            try:
                # Atualiza a tabela nf
                self.cursor.execute(''' 
                    UPDATE nf 
                    SET data = %s, numero_nf = %s, cliente = %s, cnpj_cpf = %s, observacao = %s
                    WHERE numero_nf = %s
                ''', (nova_data, novo_nf, novo_cliente, novo_doc, nova_observacao, str(values[1])))

                # Atualiza a tabela produtos_nf incluindo a coluna base_produto
                self.cursor.execute(''' 
                    UPDATE produtos_nf 
                    SET produto_nome = %s, peso = %s, base_produto = %s
                    WHERE nf_id IN (SELECT id FROM nf WHERE numero_nf = %s) AND produto_nome = %s
                ''', (produto, peso, base_produto, novo_nf, values[2]))

                self.conn.commit()

                self.cursor.execute(
                    "NOTIFY canal_atualizacao, 'menu_atualizado';"
                )
                self.conn.commit()

                self.atualizar_treeview()  # Atualizar o Treeview após a alteração
                messagebox.showinfo("Sucesso", "Nota Fiscal alterada com sucesso!")

                if self.janela_menu and hasattr(self.janela_menu, 'frame_nf'):
                # limpa widgets antigos
                    for widget in self.janela_menu.frame_nf.winfo_children():
                        widget.destroy()
                # repopula com dados atualizados
                self.janela_menu.criar_relatorio_nf()

                # Limpar as entradas após a alteração bem-sucedida
                self.data_var.set("")
                self.nf_var.set("")
                self.cliente_var.set("")
                self.doc_var.set("")
                self.observacao_text.delete("1.0", tk.END)
                self.produto_entry.set("")
                self.peso_entry.set("")
                self.base_produto_var.set("")
                self.lista_produtos.delete(0, tk.END)  # Limpar a lista de produtos

            except Exception as e:
                print(f"Erro ao alterar valores: {e}")
                self.conn.rollback()
                messagebox.showerror("Erro", f"Erro ao alterar dados: {e}")
        else:
            messagebox.showwarning("Seleção Inválida", "Selecione uma linha para alterar.")

    def excluir_valores(self):
        selected_items = self.tree.selection()  # Obtém os itens selecionados
        if not selected_items:
            messagebox.showwarning("Seleção Inválida", "Selecione uma ou mais linhas para excluir.")
            return

        # Confirmar a exclusão com o usuário
        resposta = messagebox.askyesno("Confirmar Exclusão", "Tem certeza que deseja excluir as notas fiscais selecionadas?")
        if not resposta:
            return

        try:
            for selected_item in selected_items:
                item = self.tree.item(selected_item)
                values = item["values"]

                if not values:
                    continue  # Pula itens sem valores (caso ocorram)

                # Obtém o número da NF e remove espaços em branco no início e no final
                numero_nf = values[1]
                nf_param = str(numero_nf).strip()  # Remove espaços extras

                # Excluir os produtos associados à NF na tabela produtos_nf
                self.cursor.execute(''' 
                    DELETE FROM produtos_nf
                    WHERE nf_id IN (SELECT id FROM nf WHERE TRIM(numero_nf) = %s)
                ''', (nf_param,))

                # Excluir a NF da tabela nf
                self.cursor.execute(''' 
                    DELETE FROM nf WHERE TRIM(numero_nf) = %s
                ''', (nf_param,))

            # Confirmar alterações no banco
            self.conn.commit()

            self.cursor.execute(
                "NOTIFY canal_atualizacao, 'menu_atualizado';"
            )
            self.conn.commit()

            # Atualizar Treeview após exclusão
            self.atualizar_treeview()
            messagebox.showinfo("Sucesso", "Notas Fiscais excluídas com sucesso!")

            if self.janela_menu and hasattr(self.janela_menu, 'frame_nf'):
                # limpa widgets antigos
                for widget in self.janela_menu.frame_nf.winfo_children():
                    widget.destroy()
                # repopula com dados atualizados
                self.janela_menu.criar_relatorio_nf()

            # Limpar as entradas
            self.data_var.set("")
            self.nf_var.set("")
            self.cliente_var.set("")
            self.doc_var.set("")
            self.observacao_text.delete(1.0, tk.END)
            self.produto_entry.set("")
            self.peso_entry.set("")
            self.base_produto_var.set("")   # Limpa também a combobox de Base Produto
            self.lista_produtos.delete(0, tk.END)  # Limpar a lista de produtos

        except Exception as e:
            print(f"Erro ao excluir valores: {e}")
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao excluir dados: {e}")

    def atualizar_soma_selecionada(self, event=None):
        selected_items = self.tree.selection()
        print("Itens selecionados:", selected_items)  # Debug
        soma = 0.0
        for item in selected_items:
            valores = self.tree.item(item)['values']
            print("Valores do item:", valores)  # Debug
            try:
                # Obtém a string da peso (índice 3)
                quant_str = str(valores[3]).strip()
                # Remove a unidade "Kg" (maiúscula ou minúscula) e espaços
                quant_str = quant_str.replace("Kg", "").replace("kg", "").replace(" ", "")
                # Remove o separador de milhar e converte a vírgula decimal para ponto
                quant_str = quant_str.replace(".", "").replace(",", ".")
                # Converte para float e acumula
                soma += float(quant_str)
            except (ValueError, IndexError) as e:
                print(f"Erro convertendo peso: {e}")
                continue
        print("Soma calculada:", soma)  # Debug

        # Formata o número com separador de milhar e 3 casas decimais.
        # Exemplo padrão: "1,780.200"
        formatted_soma = f"{soma:,.3f}"
        # Troca os separadores: de "1,780.200" para "1.780,200"
        formatted_soma = formatted_soma.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
        
        self.soma_label.config(text=f"Soma dos Pesos: {formatted_soma} Kg")

    def on_tree_select(self, event):
        self.carregar_dados(event)
        self.atualizar_soma_selecionada(event)

    def voltar_para_menu(self):
        """Fecha a janela atual e reexibe o menu principal com atualização visual."""
        self.destroy()                           # Fecha a janela atual
        self.janela_menu.deiconify()             # Reexibe a janela do menu
        self.janela_menu.state("zoomed")         # Garante que fique maximizada
        self.janela_menu.lift()                  # Garante que fique no topo
        self.janela_menu.update()                # Força atualização visual
        
    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente"""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Fecha a conexão com o banco de dados, se estiver aberta
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
                print("Conexão com o banco de dados fechada.")

            self.destroy()  # Fecha a janela
            sys.exit(0)  # Encerra o processo completamente
