import tkinter as tk
from tkinter import messagebox, ttk
from conexao_db import conectar
from centralizacao_tela import centralizar_janela  # Importar a função
from exportacao import exportar_para_pdf, exportar_para_excel #, selecionar_diretorio
import sys
from logos import aplicar_icone
from tkinter import filedialog
import re
import threading

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

        self.header_container = tk.Frame(self.janela_base_produto, bg="#34495e")
        self.header_container.pack(fill=tk.X)

        # Label principal (título) — armazenado como atributo caso precise alterar depois
        self.cabecalho = tk.Label(self.header_container,
                                text="Base de Produtos",
                                font=("Arial", 24, "bold"),
                                bg="#34495e", fg="white", pady=15)
        self.cabecalho.pack(side="left", fill=tk.X, expand=True)

        # Frame para o botão de ajuda (mantém estilo do cabeçalho)
        self.help_frame = tk.Frame(self.header_container, bg="#34495e")
        self.help_frame.pack(side="right", padx=(8, 12), pady=6)

        self.botao_ajuda_produtos = tk.Button(
            self.help_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=lambda: self._abrir_ajuda_produtos_modal()
        )
        self.botao_ajuda_produtos.bind("<Enter>", lambda e: self.botao_ajuda_produtos.config(bg="#3b5566"))
        self.botao_ajuda_produtos.bind("<Leave>", lambda e: self.botao_ajuda_produtos.config(bg="#2c3e50"))
        self.botao_ajuda_produtos.pack(side="right", padx=(4, 6), pady=4)

        # Tooltip (seu método será adicionado abaixo)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_produtos, "Ajuda — Produtos (F1)")
        except Exception:
            pass

        # Atalho F1 (vinculado ao Toplevel desta interface)
        try:
            # bind_all no toplevel para garantir funcionamento quando a janela estiver ativa
            self.janela_base_produto.bind_all("<F1>", lambda e: self._abrir_ajuda_produtos_modal())
        except Exception:
            try:
                self.janela_base_produto.bind("<F1>", lambda e: self._abrir_ajuda_produtos_modal())
            except Exception:
                pass

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

        # Oculta o menu ao abrir
        if self.janela_menu is not None:
            self.janela_menu.withdraw()

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

    def _abrir_ajuda_produtos_modal(self, contexto=None):
        """Modal profissional com explicações: adicionar, editar, excluir, limpar, exportar Excel/PDF."""
        try:
            modal = tk.Toplevel(self.janela_base_produto)
            modal.title("Ajuda — Base de Produtos")
            modal.transient(self.janela_base_produto)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões / centralização
            w, h = 920, 680
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # aplicar ícone (mesma função que você usa)
            try:
                caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
                aplicar_icone(modal, caminho_icone)
            except Exception:
                pass

            # Cabeçalho
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Base de Produtos", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # Corpo: nav esquerda + conteúdo direita
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10, 6), padx=12)

            sections = [
                "Visão Geral",
                "Adicionar Produto",
                "Editar Produto",
                "Excluir Produto",
                "Botão Limpar (Campos)",
                "Exportar Excel / PDF",
                "Validações e Boas Práticas",
                "Exemplos Rápidos",
                "FAQ"
            ]
            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)

            txt = tk.Text(content_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb = tk.Scrollbar(content_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            # Conteúdos detalhados
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "A janela 'Base de Produtos' é o catálogo central dos produtos da empresa. "
                "Aqui você cadastra o nome do produto e as porcentagens de cobre e zinco que o compõem. "
                "A Treeview lista os produtos; selecione um item para carregar seus dados nos campos acima."
            )

            contents["Adicionar Produto"] = (
                "Adicionar Produto — Passo a passo\n\n"
                "1) Preencha 'Nome do Produto', 'Porcentagem de Cobre (%)' e 'Porcentagem de Zinco (%)'.\n"
                "2) Verifique se os valores de porcentagem são numéricos e somam ≤ 100% (ou segundo sua regra).\n"
                "3) Clique em 'Adicionar'. O sistema calcula um novo ID (menor disponível) e realiza INSERT no banco.\n"
                "4) Após sucesso, a Treeview é atualizada automaticamente e os campos são limpos.\n\n"
                "Validações típicas:\n"
                " - Nome não pode ser vazio.\n"
                " - Cobre e Zinco devem ser números inteiros ou decimais (ex.: 70 ou 70.5).\n"
                " - A soma das porcentagens pode ter limites definidos pelo seu processo (valide conforme regra)."
            )

            contents["Editar Produto"] = (
                "Editar Produto — Passo a passo\n\n"
                "1) Selecione o produto na lista (clique na linha).\n"
                "2) Os campos 'Nome', 'Cobre' e 'Zinco' são preenchidos automaticamente.\n"
                "3) Alterar os valores desejados e clique em 'Alterar'. O sistema faz UPDATE no banco.\n"
                "4) Em caso de erro, será exibida mensagem com o motivo; caso contrário a lista é atualizada.\n\n"
                "Dicas:\n"
                " - Normalizar nomes (remover acentos) é feito automaticamente pelo método 'remover_acentos'.\n"
                " - Verifique se outro usuário não alterou o mesmo registro simultaneamente."
            )

            contents["Excluir Produto"] = (
                "Excluir Produto — Passo a passo\n\n"
                "1) Selecione um ou mais produtos na Treeview.\n"
                "2) Clique em 'Excluir' e confirme a ação na caixa de diálogo.\n"
                "3) O sistema executa DELETE por ID e notifica o canal de atualização (NOTIFY) para recarregar\n"
                "   outros relatórios/painéis quando aplicável.\n\n"
                "Recomendações:\n"
                " - Faça backup antes de operações em lote.\n"
                " - Considere desativar em vez de excluir se precisar manter histórico."
            )

            contents["Botão Limpar (Campos)"] = (
                "Botão Limpar — comportamento\n\n"
                " - Ao clicar em 'Limpar', todos os campos de entrada acima são esvaziados.\n"
                " - Use o botão quando quiser iniciar cadastro de novo produto ou cancelar edição.\n"
                " - O botão não altera o banco, apenas limpa a interface. Após limpar, o item selecionado na\n"
                "   Treeview permanece selecionado até que você selecione outro."
            )

            contents["Exportar Excel / PDF"] = (
                "Exportar Excel / PDF — como funciona\n\n"
                "Exportar Excel:\n"
                "  • Clique em 'Exportar Excel' e escolha o local e nome do arquivo (.xlsx).\n"
                "  • A função exportar_para_excel(caminho, 'produtos', [colunas]) é chamada e cria a planilha.\n\n"
                "Exportar PDF:\n"
                "  • Clique em 'Exportar PDF' e escolha local/nome (.pdf).\n"
                "  • A função exportar_para_pdf(caminho, 'produtos', cabeçalho, titulo) é chamada.\n\n"
                "Boas práticas:\n"
                " - Verifique permissões de escrita no diretório selecionado.\n - Para grandes volumes, exporte em background para não travar a UI (pode-se usar threading)."
            )

            contents["Validações e Boas Práticas"] = (
                "Validações e Boas Práticas\n\n"
                " - Valide entradas numéricas usando try/float ou regex antes de gravar.\n"
                " - Normalize nomes removendo acentos (já incluso em remover_acentos).\n"
                " - Use transações e commit apenas após todas as operações concluídas.\n"
                " - Notifique outros módulos via NOTIFY para manter relatórios atualizados."
            )

            contents["Exemplos Rápidos"] = (
                "Exemplos Rápidos\n\n"
                "Exemplo 1 — Adicionar 'Fio Latão 1mm':\n"
                "  • Nome: Fio Latão 1mm\n  • Cobre: 70\n  • Zinco: 30\n  • Clique em Adicionar -> lista atualizada.\n\n"
                "Exemplo 2 — Exportar:\n"
                "  • Clique em Exportar Excel -> escolha pasta -> abra arquivo gerado para revisar."
            )

            contents["FAQ"] = (
                "FAQ — Perguntas rápidas\n\n"
                "Q: Porque a Treeview mostra '70%' quando o banco tem 70?\n"
                "R: O seu código formata para exibir com '%' para leitura. Internamente o banco mantém números.\n\n"
                "Q: O export falha ao salvar?\n"
                "R: Verifique permissões e caminho selecionado; tente salvar em 'Documentos' como teste."
            )

            # Função para mostrar seção
            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            # Inicializa com a primeira seção selecionada
            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])

            listbox.bind("<<ListboxSelect>>", on_select)

            # Rodapé: Fechar
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right", padx=6)

            # Atalhos
            modal.bind("<Escape>", lambda e: modal.destroy())

            modal.focus_set()
            modal.wait_window()

        except Exception as e:
            print("Erro ao abrir modal de ajuda (Produtos):", e)

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
                    self.after(0, self.janela_base_produto.destroy)
                except Exception:
                    try:
                        self.janela_base_produto.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()
        
    def on_closing(self):
        """Função para lidar com o fechamento da janela."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar o programa?"):
            self.janela_base_produto.destroy()  # Fecha a janela atual
            sys.exit()  # Encerra completamente o programa
