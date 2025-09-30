import tkinter as tk
from tkinter import messagebox, ttk
from conexao_db import conectar
from exportacao import exportar_para_pdf, exportar_para_excel
from logos import aplicar_icone
from tkinter import filedialog
import sys
import json
import os
import psycopg2
import re
import threading

# Função para conectar ao banco de dados PostgreSQL
class InterfaceMateriais:
    def __init__(self, janela_menu=None):
        self.janela_menu = janela_menu
        self.janela_materiais = tk.Toplevel()
        self.janela_materiais.title("Gerenciamento de Materiais")

        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self.janela_materiais, caminho_icone)

        # Maximiza a janela e define o fundo com um tom neutro e moderno
        self.janela_materiais.state("zoomed")
        self.janela_materiais.configure(bg="#ecf0f1")
        
        # Configura os estilos padrões
        self.configurar_estilos()
        
        # ---------- Cabeçalho com container e botão de Ajuda ----------
        self.header_container = tk.Frame(self.janela_materiais, bg="#34495e")
        self.header_container.pack(fill=tk.X)

        self.cabecalho = tk.Label(
            self.header_container,
            text="Gerenciamento de Materiais",
            font=("Arial", 24, "bold"),
            bg="#34495e", fg="white", pady=15
        )
        self.cabecalho.pack(side="left", fill=tk.X, expand=True)

        # Frame para o botão de ajuda (mesma linha do cabeçalho)
        self.help_frame = tk.Frame(self.header_container, bg="#34495e")
        self.help_frame.pack(side="right", padx=(8, 12), pady=6)

        self.botao_ajuda_materiais = tk.Button(
            self.help_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=lambda: self._abrir_ajuda_materiais_modal()
        )
        self.botao_ajuda_materiais.bind("<Enter>", lambda e: self.botao_ajuda_materiais.config(bg="#3b5566"))
        self.botao_ajuda_materiais.bind("<Leave>", lambda e: self.botao_ajuda_materiais.config(bg="#2c3e50"))
        self.botao_ajuda_materiais.pack(side="right", padx=(4, 6), pady=4)

        # Tooltip (usa o método abaixo)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_materiais, "Ajuda — Materiais (F1)")
        except Exception:
            pass

        # Atalho F1 (vinculado ao Toplevel desta interface)
        try:
            self.janela_materiais.bind_all("<F1>", lambda e: self._abrir_ajuda_materiais_modal())
        except Exception:
            try:
                self.janela_materiais.bind("<F1>", lambda e: self._abrir_ajuda_materiais_modal())
            except Exception:
                pass

        # Frame para Labels e Entradas
        frame_acoes = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_acoes.pack(padx=20, pady=10, fill=tk.X)

        # Label e entrada para o nome do material
        tk.Label(frame_acoes, text="Nome do Material", bg="#ecf0f1", font=("Arial", 12)) .grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_nome = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entry_nome.grid(row=0, column=1, padx=5, pady=5)

        # Combobox para fornecedor
        tk.Label(frame_acoes, text="Fornecedor", bg="#ecf0f1", font=("Arial", 12)) .grid(row=1, column=0, padx=5, pady=5, sticky="e")
        fornecedores_carregados = self.load_fornecedores()
        if not fornecedores_carregados:
            fornecedores_carregados = ["Termomecanica", "Metallica"]
        self.combobox_fornecedor = ttk.Combobox(frame_acoes, width=23, font=("Arial", 12), state="readonly")
        self.combobox_fornecedor['values'] = fornecedores_carregados
        self.combobox_fornecedor.grid(row=1, column=1, padx=5, pady=5)

        self.entry_novo_fornecedor = tk.Entry(frame_acoes, width=20, font=("Arial", 12))
        self.entry_novo_fornecedor.grid(row=1, column=2, padx=5, pady=5)

        btn_adicionar_fornecedor = tk.Button(frame_acoes, text="Adicionar", font=("Arial", 10), command=self.adicionar_fornecedor)
        btn_adicionar_fornecedor.grid(row=1, column=3, padx=5, pady=5)
        btn_excluir_fornecedor = tk.Button(frame_acoes, text="Excluir", font=("Arial", 10), command=self.excluir_fornecedor)
        btn_excluir_fornecedor.grid(row=1, column=4, padx=5, pady=5)

        # Outras entradas: Valor e Grupo
        tk.Label(frame_acoes, text="Valor", bg="#ecf0f1", font=("Arial", 12)) .grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_valor = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entry_valor.grid(row=2, column=1, padx=5, pady=5)
        # bind para formatação automática com 2 casas decimais:
        self.entry_valor.bind(
            "<KeyRelease>",
            lambda ev, ent=self.entry_valor: self._formatar_numero(ent, casas=4)
        )

        # oculta menu ao abrir
        if self.janela_menu is not None:
            self.janela_menu.withdraw()

        tk.Label(frame_acoes, text="Grupo", bg="#ecf0f1", font=("Arial", 12)) .grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.combobox_grupo = ttk.Combobox(frame_acoes, values=["Cobre", "Zinco", "Sucata"], state="readonly", font=("Arial", 12), width=23)
        self.combobox_grupo.grid(row=3, column=1, padx=5, pady=5)

        # Frame para os botões de ação dos materiais
        frame_botoes_acao = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_botoes_acao.pack(pady=10, fill=tk.X)

        botao_adicionar = ttk.Button(frame_botoes_acao, text="Adicionar", command=self.adicionar_material)
        botao_adicionar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_alterar = ttk.Button(frame_botoes_acao, text="Alterar", command=self.alterar_material)
        botao_alterar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_excluir = ttk.Button(frame_botoes_acao, text="Excluir", command=self.excluir_materiais)
        botao_excluir.pack(side=tk.LEFT, padx=5, pady=5)

        botao_limpar = ttk.Button(frame_botoes_acao, text="Limpar", command=self.limpar_campos)
        botao_limpar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_voltar = ttk.Button(frame_botoes_acao, text="Voltar", command=lambda: self.voltar_para_menu())
        botao_voltar.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para os botões de exportação
        frame_botoes_exportacao = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_botoes_exportacao.pack(pady=10, fill=tk.X)

        botao_exportar_excel = ttk.Button(frame_botoes_exportacao, text="Exportar Excel", command=self.exportar_excel_materiais)
        botao_exportar_excel.pack(side=tk.LEFT, padx=5, pady=5)

        botao_exportar_pdf = ttk.Button(frame_botoes_exportacao, text="Exportar PDF", command=self.exportar_pdf_materiais)
        botao_exportar_pdf.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para o Treeview
        frame_treeview = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_treeview.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # Barras de rolagem
        scrollbar_vertical = tk.Scrollbar(frame_treeview, orient=tk.VERTICAL)
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        scrollbar_horizontal = tk.Scrollbar(frame_treeview, orient=tk.HORIZONTAL)
        scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

        self.lista_materiais = ttk.Treeview(
            frame_treeview,
            columns=("ID", "Nome", "Fornecedor", "Valor", "Grupo"),
            show="headings",
            yscrollcommand=scrollbar_vertical.set,
            xscrollcommand=scrollbar_horizontal.set
        )
        self.lista_materiais.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configuração das colunas do Treeview
        self.lista_materiais.heading("ID", text="ID", anchor="center")
        self.lista_materiais.heading("Nome", text="Nome", anchor="center")
        self.lista_materiais.heading("Fornecedor", text="Fornecedor", anchor="center")
        self.lista_materiais.heading("Valor", text="Valor", anchor="center")
        self.lista_materiais.heading("Grupo", text="Grupo", anchor="center")
        
        self.lista_materiais.column("ID", width=50, anchor="center", stretch=False)
        self.lista_materiais.column("Nome", width=200, anchor="center")
        self.lista_materiais.column("Fornecedor", width=150, anchor="center")
        self.lista_materiais.column("Valor", width=100, anchor="center")
        self.lista_materiais.column("Grupo", width=100, anchor="center")

        # Exibe apenas as colunas desejadas (ocultando a coluna "ID")
        self.lista_materiais["displaycolumns"] = ("Nome", "Fornecedor", "Valor", "Grupo")

        # Vincula as barras de rolagem ao Treeview
        scrollbar_vertical.config(command=self.lista_materiais.yview)
        scrollbar_horizontal.config(command=self.lista_materiais.xview)

        # Evento de clique no Treeview para selecionar um item
        self.lista_materiais.bind("<ButtonRelease-1>", self.selecionar_material)

        # Atualiza a lista de materiais
        self.atualizar_lista_materiais()

        self.janela_materiais.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.janela_materiais.mainloop()

    def configurar_estilos(self):
        """Configura os estilos utilizados na interface."""
        estilo = ttk.Style()
        estilo.theme_use("alt")
        estilo.configure("Custom.TFrame", background="#ecf0f1")
        
        # Estilo para o Treeview
        estilo.configure("Treeview", 
                         background="white", 
                         foreground="black", 
                         rowheight=20, 
                         fieldbackground="white",
                         font=("Arial", 10))
        estilo.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        estilo.map("Treeview", 
                   background=[("selected", "#0078D7")], 
                   foreground=[("selected", "white")])
        
        # Estilos para os botões
        estilo.configure("Material.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#2980b9",      # Azul profissional
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("Material.TButton",
                   background=[("active", "#3498db")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])
        
        estilo.configure("Excel.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#27ae60",      # Verde profissional
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        
        estilo.configure("PDF.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#c0392b",      # Vermelho sofisticado
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)

    def _create_tooltip(self, widget, text, delay=450):
        """Tooltip melhorado: quebra de linha automática e ajuste para não sair da tela."""
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            # calcula largura máxima adequada para o tooltip (não maior que a tela)
            try:
                screen_w = widget.winfo_screenwidth()
                screen_h = widget.winfo_screenheight()
            except Exception:
                screen_w, screen_h = 1024, 768

            wrap_len = min(360, max(200, screen_w - 80))  # largura do texto em pixels, com limites

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

            # posição inicial centrada horizontalmente sobre o widget, abaixo do widget
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            # força o cálculo do tamanho real do tooltip
            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajustar horizontalmente para não sair da tela
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # ajustar verticalmente: se não couber embaixo, tenta acima do widget
            if y + h > screen_h:
                y_above = widget.winfo_rooty() - h - 6
                if y_above > 10:
                    y = y_above
                else:
                    # caso não caiba nem acima nem abaixo, limita para caber na tela
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

    def _abrir_ajuda_materiais_modal(self, contexto=None):
        """Modal profissional com instruções para Materiais e Fornecedores (Adicionar/Edit/Excluir/Limpar/Exportar)."""
        try:
            modal = tk.Toplevel(self.janela_materiais)
            modal.title("Ajuda — Materiais")
            modal.transient(self.janela_materiais)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões / centralização
            w, h = 920, 680
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # Ícone (se aplicável)
            try:
                caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
                aplicar_icone(modal, caminho_icone)
            except Exception:
                pass

            # Cabeçalho do modal
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Gerenciamento de Materiais", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # Corpo: navegação à esquerda + conteúdo à direita
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10, 6), padx=12)

            sections = [
                "Visão Geral",
                "Adicionar Material",
                "Editar Material",
                "Excluir Material",
                "Adicionar / Excluir Fornecedor",
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

            # Conteúdo por seção
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "A tela de Materiais centraliza o cadastro de matérias-primas, valores, fornecedores e agrupamentos.\n"
                "Use os campos para inserir Nome, Fornecedor, Valor e Grupo. A Treeview lista os registros; selecione um item para carregar seus dados nos campos."
            )

            contents["Adicionar Material"] = (
                "Adicionar Material — Passo a passo\n\n"
                "1) Preencha 'Nome do Material', selecione/insira o Fornecedor, informe o Valor e escolha o Grupo.\n"
                "2) Verifique o formato do valor (ex.: 1234,56) e clique em 'Adicionar'.\n"
                "3) O sistema gera um ID (menor disponível) e executa INSERT no banco; em seguida atualiza a lista.\n\n"
                "Validações:\n"
                " - Nome e fornecedor não podem ficar em branco.\n"
                " - Valor deve ser numérico; use o campo para ver formatação automática.\n"
            )

            contents["Editar Material"] = (
                "Editar Material — Passo a passo\n\n"
                "1) Selecione o material na Treeview.\n"
                "2) Os campos serão preenchidos automaticamente.\n"
                "3) Altere Nome / Fornecedor / Valor / Grupo e clique em 'Alterar'. O sistema faz UPDATE no banco.\n"
                "4) Se houver conflito ou erro de banco, será exibida mensagem com detalhes.\n"
            )

            contents["Excluir Material"] = (
                "Excluir Material — Passo a passo\n\n"
                "1) Selecione um ou mais materiais na lista.\n"
                "2) Clique em 'Excluir' e confirme na caixa de diálogo.\n"
                "3) O sistema executa DELETE por ID e envia NOTIFY para atualizar relatórios/painéis.\n\n"
                "Recomendações:\n"
                " - Faça backup antes de excluir muitos registros; considere sinalizar/desativar em vez de excluir."
            )

            contents["Adicionar / Excluir Fornecedor"] = (
                "Adicionar / Excluir Fornecedor — Passo a passo\n\n"
                "Adicionar:\n"
                "  • No campo 'Novo Fornecedor' digite o nome e clique em 'Adicionar'.\n"
                "  • Se não existir, será incluído na combobox e salvo em config/fornecedores.json.\n\n"
                "Excluir:\n"
                "  • Digite o nome do fornecedor no campo 'Novo Fornecedor' e clique em 'Excluir'.\n"
                "  • O item será removido da combobox e do arquivo JSON.\n\n"
                "Observações:\n"
                " - Fornecedores utilizados em materiais existentes não são removidos do banco automaticamente; ao excluir um fornecedor, verifique se registros dependentes precisam ser ajustados."
            )

            contents["Botão Limpar (Campos)"] = (
                "Botão Limpar — comportamento\n\n"
                " - Ao clicar em 'Limpar', os campos Nome, Fornecedor, Valor e Grupo são esvaziados.\n"
                " - O botão não altera o banco de dados; serve apenas para limpar a interface antes de novo cadastro."
            )

            contents["Exportar Excel / PDF"] = (
                "Exportar Excel / PDF — como funciona\n\n"
                "Exportar Excel:\n"
                "  • Clique em 'Exportar Excel' e escolha local/nome do arquivo (.xlsx).\n"
                "  • A função exportar_para_excel(caminho, 'materiais', [colunas]) será chamada.\n\n"
                "Exportar PDF:\n"
                "  • Clique em 'Exportar PDF' e escolha local/nome (.pdf).\n"
                "  • A função exportar_para_pdf(caminho, 'materiais', cabeçalhos, 'Base de Materiais') será chamada.\n\n"
                "Dicas:\n"
                " - Verifique permissões de gravação no diretório escolhido.\n - Para grandes volumes, execute export em background (threading) para não travar a UI."
            )

            contents["Validações e Boas Práticas"] = (
                "Validações e Boas Práticas\n\n"
                " - Use converter_valor para validar e converter valores.\n"
                " - Normalize nomes e use transações no banco.\n"
                " - Dispare NOTIFY para manter painéis/relatórios sincronizados."
            )

            contents["Exemplos Rápidos"] = (
                "Exemplos Rápidos\n\n"
                "Exemplo 1 — Adicionar material:\n"
                "  • Nome: 'Sucata A'\n  • Fornecedor: 'Termomecanica'\n  • Valor: 1234,56\n  • Grupo: 'Sucata'\n  • Clique em Adicionar → lista atualizada.\n\n"
                "Exemplo 2 — Exportar:\n"
                "  • Clique em Exportar PDF → selecione local → abra o arquivo para conferência."
            )

            contents["FAQ"] = (
                "FAQ — Materiais / Fornecedores\n\n"
                "Q: Posso remover um fornecedor usado por materiais existentes?\n"
                "A: Sim, mas verifique dependências. O código remove do JSON de fornecedores, não altera registros no banco.\n\n"
                "Q: Por que o valor aparece com vírgula?\n"
                "A: A interface formata números para o padrão brasileiro (vírgula como separador decimal)."
            )

            # Função para mostrar seção
            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            # Inicializa
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
            print("Erro ao abrir modal de ajuda (Materiais):", e)   

    def _formatar_numero(self, entry_widget, casas=2):
        """
        Formata o conteúdo do entry_widget com 'casas' casas decimais
        e separadores de milhar no padrão brasileiro.
        """
        texto = entry_widget.get()
        dígitos = re.sub(r"\D", "", texto)  # só mantém dígitos
        if not dígitos:
            novo = "0," + "0"*casas
        else:
            fator = 10**casas
            valor = int(dígitos) / fator
            # formata no estilo en_US (ex: "1,234.56") e depois converte:
            s = f"{valor:,.{casas}f}"
            novo = s.replace(",", "X").replace(".", ",").replace("X", ".")

        # atualiza o entry sem reentrar no bind
        entry_widget.unbind("<KeyRelease>")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, novo)
        entry_widget.bind(
            "<KeyRelease>",
            lambda ev, ent=entry_widget: self._formatar_numero(ent, casas)
        )

    def converter_valor(self, valor):
        try:
            return float(valor.replace(',', '.'))
        except ValueError:
            messagebox.showerror("Erro", "O valor deve ser um número!")
            return None

    def obter_menor_id_disponivel(self):
        """Obtém o menor ID disponível na tabela materiais"""
        try:
            conexao = conectar()
            if conexao is None:
                return None
            cursor = conexao.cursor()
            cursor.execute("SELECT id FROM materiais ORDER BY id")
            ids = cursor.fetchall()
            conexao.close()
            
            if not ids:
                return 1
            ids_usados = [item[0] for item in ids]
            menor_id_disponivel = 1
            for id_usado in ids_usados:
                if id_usado != menor_id_disponivel:
                    return menor_id_disponivel
                menor_id_disponivel += 1
            return menor_id_disponivel
        except psycopg2.Error as e:
            print(f"Erro ao acessar o banco de dados: {e}")
            return None

    def adicionar_material(self):
        nome = self.entry_nome.get()
        fornecedor = self.combobox_fornecedor.get().strip()
        valor = self.entry_valor.get()
        grupo = self.combobox_grupo.get().strip()

        if not nome or not fornecedor or not valor:
            messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos!")
            return

        valor_convertido = self.converter_valor(valor)
        if valor_convertido is None:
            return

        conexao = conectar()
        if conexao is None:
            return

        cursor = conexao.cursor()
        menor_id = self.obter_menor_id_disponivel()
        if menor_id is None:
            messagebox.showerror("Erro", "Não foi possível determinar o menor ID disponível.")
            conexao.close()
            return

        cursor.execute(
            "INSERT INTO materiais (id, nome, fornecedor, valor, grupo) VALUES (%s, %s, %s, %s, %s)",
            (menor_id, nome, fornecedor, valor_convertido, grupo)
        )
        conexao.commit()

        # dispara notificação
        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        conexao.commit()

        conexao.close()
        messagebox.showinfo("Sucesso", "Material adicionado com sucesso!")
        
        self.atualizar_lista_materiais()

        # Atualizar o relatório de produto e material automaticamente
        if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorios_produto_material'):
            for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                widget.destroy()
            self.janela_menu.criar_relatorio_produto_material()  # Alteração aqui para chamar a função correta

    def atualizar_lista_materiais(self):
        """Atualiza a lista de materiais exibida na Treeview"""
        for item in self.lista_materiais.get_children():
            self.lista_materiais.delete(item)
        try:
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            cursor.execute("SELECT id, nome, fornecedor, valor, grupo FROM materiais ORDER BY nome ASC")
            materiais = cursor.fetchall()
            for material in materiais:
                id_material, nome, fornecedor, valor, grupo = material
                # Formata o valor convertendo ponto para vírgula
                valor_formatado = str(valor).replace('.', ',')
                self.lista_materiais.insert("", "end", values=(id_material, nome, fornecedor, valor_formatado, grupo))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar lista de materiais: {e}")
        finally:
            if 'conexao' in locals() and conexao:
                conexao.close()

    def selecionar_material(self, event):
        try:
            selected_item = self.lista_materiais.selection()[0]
            item = self.lista_materiais.item(selected_item)
            values = item['values']
            print("Valores selecionados:", values)
            if len(values) >= 4:
                self.entry_nome.delete(0, tk.END)
                self.entry_nome.insert(0, values[1])
                self.combobox_fornecedor.set(values[2])
                self.entry_valor.delete(0, tk.END)
                self.entry_valor.insert(0, str(values[3]).replace('.', ','))
                self.combobox_grupo.set(values[4])
            else:
                print("Os valores selecionados não contêm dados suficientes.")
        except IndexError:
            print("Nenhum material selecionado.")
        except Exception as e:
            print("Erro inesperado:", e)

    def alterar_material(self):
        try:
            selected_item = self.lista_materiais.selection()[0]
        except IndexError:
            messagebox.showwarning("Aviso", "Selecione um material da lista!")
            return

        nome = self.entry_nome.get()
        fornecedor = self.combobox_fornecedor.get().strip()
        valor = self.entry_valor.get()
        grupo = self.combobox_grupo.get().strip()

        if not nome or not fornecedor or not valor:
            messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos!")
            return

        valor_convertido = self.converter_valor(valor)
        if valor_convertido is None:
            return

        conexao = conectar()
        if conexao is None:
            return
        cursor = conexao.cursor()
        item_id = self.lista_materiais.item(selected_item, 'values')[0]
        cursor.execute("UPDATE materiais SET nome=%s, fornecedor=%s, valor=%s, grupo=%s WHERE id=%s", 
                       (nome, fornecedor, valor_convertido, grupo, item_id))
        conexao.commit()

        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        conexao.commit()

        conexao.close()
        messagebox.showinfo("Sucesso", "Material alterado com sucesso!")
        self.atualizar_lista_materiais()

        # Atualizar o relatório de matéria-prima automaticamente
        if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorio_materia_prima'):
            for widget in self.janela_menu.frame_relatorio_materia_prima.winfo_children():
                widget.destroy()
            self.janela_menu.criar_relatorio_materia_prima()

    def limpar_campos(self):
        self.entry_nome.delete(0, tk.END)
        self.combobox_fornecedor.set('')
        self.entry_valor.delete(0, tk.END)
        self.combobox_grupo.set('')

    def excluir_materiais(self):
        selected_items = self.lista_materiais.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione ao menos um material da lista!")
            return

        if messagebox.askyesno("Confirmação", f"Tem certeza que deseja excluir {len(selected_items)} materiais?"):
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            for item in selected_items:
                material_id = self.lista_materiais.item(item, 'values')[0]
                cursor.execute("DELETE FROM materiais WHERE id=%s", (material_id,))
            conexao.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conexao.commit()

            conexao.close()
            messagebox.showinfo("Sucesso", "Materiais excluídos com sucesso!")
            self.atualizar_lista_materiais()

            # Atualizar o relatório de produto e material automaticamente
            if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorios_produto_material'):
                for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                    widget.destroy()
                self.janela_menu.criar_relatorio_produto_material()  # Alteração aqui para chamar a função correta

            self.limpar_campos()

    def exportar_excel_materiais(self):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="base_materiais.xlsx"
        )
        if caminho_arquivo:
            exportar_para_excel(caminho_arquivo, "materiais", ["Produto", "Fornecedor", "Valor", "Grupo"])

    def exportar_pdf_materiais(self):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="base_materiais.pdf"
        )
        if caminho_arquivo:
            exportar_para_pdf(
                caminho_arquivo,
                "materiais",
                ["Produto", "Fornecedor", "Valor", "Grupo"],
                "Base de Materiais"
            )

    def ordena_coluna(self, treeview, coluna, reverse):
        """Ordena a Treeview por uma coluna específica"""
        data = [(treeview.item(child)["values"], child) for child in treeview.get_children("")]
        data.sort(key=lambda x: x[0][coluna], reverse=reverse)
        for index, (item, tree_id) in enumerate(data):
            treeview.move(tree_id, '', index)
        treeview.heading(coluna, command=lambda: self.ordena_coluna(treeview, coluna, not reverse))

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
                    self.after(0, self.janela_materiais.destroy)
                except Exception:
                    try:
                        self.janela_materiais.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()
        
    def adicionar_fornecedor(self):
        novo_fornecedor = self.entry_novo_fornecedor.get().strip()
        if novo_fornecedor:
            valores_existentes = list(self.combobox_fornecedor['values'])
            if novo_fornecedor not in valores_existentes:
                valores_existentes.append(novo_fornecedor)
                self.combobox_fornecedor['values'] = valores_existentes
                self.salvar_fornecedores(valores_existentes)
                messagebox.showinfo("Sucesso", f"Fornecedor '{novo_fornecedor}' adicionado!")
                self.entry_novo_fornecedor.delete(0, tk.END)
            else:
                messagebox.showwarning("Aviso", f"O fornecedor '{novo_fornecedor}' já existe!")
        else:
            messagebox.showwarning("Aviso", "Digite o nome do fornecedor!")

    def excluir_fornecedor(self):
        fornecedor_para_excluir = self.entry_novo_fornecedor.get().strip()
        if fornecedor_para_excluir:
            valores_existentes = list(self.combobox_fornecedor['values'])
            if fornecedor_para_excluir in valores_existentes:
                valores_existentes.remove(fornecedor_para_excluir)
                self.combobox_fornecedor['values'] = valores_existentes
                self.salvar_fornecedores(valores_existentes)
                messagebox.showinfo("Sucesso", f"Fornecedor '{fornecedor_para_excluir}' removido!")
                self.entry_novo_fornecedor.delete(0, tk.END)
            else:
                messagebox.showwarning("Aviso", f"O fornecedor '{fornecedor_para_excluir}' não existe!")
        else:
            messagebox.showwarning("Aviso", "Digite o nome do fornecedor para remover!")

    def salvar_fornecedores(self, fornecedores):
        """Salva os fornecedores em um arquivo JSON dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)  # Diretório do script
        caminho_pasta = os.path.join(caminho_base, "config")  # Criar pasta 'config'
        os.makedirs(caminho_pasta, exist_ok=True)  # Cria a pasta se não existir

        caminho_arquivo = os.path.join(caminho_pasta, "fornecedores.json")  # Caminho completo

        try:
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                json.dump(fornecedores, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print("Erro ao salvar fornecedores:", e)

    def load_fornecedores(self):
        """Carrega os fornecedores do arquivo JSON dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)
        caminho_arquivo = os.path.join(caminho_base, "config", "fornecedores.json")

        if os.path.exists(caminho_arquivo):
            try:
                with open(caminho_arquivo, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            except Exception as e:
                print("Erro ao carregar fornecedores:", e)

        return []
    
    def on_closing(self):
        """Função para lidar com o fechamento da janela."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar o programa?"):
            self.janela_materiais.destroy()  # Fecha a janela atual
            sys.exit()  # Encerra completamente o programa