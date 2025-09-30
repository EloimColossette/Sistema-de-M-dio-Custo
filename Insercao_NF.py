import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from decimal import Decimal
import sys
from conexao_db import conectar
import pandas as pd
import json
import os
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
import threading

class Janela_InsercaoNF(tk.Toplevel):
    def __init__(self, parent, janela_menu):
        super().__init__()
        # Configuração inicial da janela
        self.resizable(True, True)
        self.geometry("900x500")
        self.state("normal")

        self.janela_menu = janela_menu
        self.parent = parent
        self.title("Entrada de NFs")
        self.state("zoomed")

        # caminho do ícone (mantive seu valor)
        self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        try:
            aplicar_icone(self, self.caminho_icone)
        except Exception:
            pass

        # Estilos
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("alt")
        except Exception:
            pass
        estilo.configure(".", font=("Arial", 10))
        estilo.configure("Treeview", rowheight=25)
        estilo.configure("Treeview.Heading", font=("Courier", 10, "bold"))

        # DB (tente conectar; se falhar, mantenha None para não travar)
        try:
            self.conn = conectar()
            self.cursor = self.conn.cursor()
        except Exception:
            self.conn = None
            self.cursor = None

        # Mapeamento de colunas e ordem (mantive seu mapeamento)
        self.colunas_fixas = {
            "data": "Data",
            "nf": "NF",
            "fornecedor": "Fornecedor",
            "material_1": "Material 1",
            "material_2": "Material 2",
            "material_3": "Material 3",
            "material_4": "Material 4",
            "material_5": "Material 5",
            "produto": "Produto",
            "custo_empresa": "Custo Empresa",
            "ipi": "IPI",
            "valor_integral": "Valor Integral",
            "valor_unitario_1": "Valor Unitário 1",
            "valor_unitario_2": "Valor Unitário 2",
            "valor_unitario_3": "Valor Unitário 3",
            "valor_unitario_4": "Valor Unitário 4",
            "valor_unitario_5": "Valor Unitário 5",
            "duplicata_1": "Duplicata 1",
            "duplicata_2": "Duplicata 2",
            "duplicata_3": "Duplicata 3",
            "duplicata_4": "Duplicata 4",
            "duplicata_5": "Duplicata 5",
            "duplicata_6": "Duplicata 6",
            "valor_unitario_energia": "Valor Unitário Energia",
            "valor_mao_obra_tm_metallica": "Valor Mão de Obra TM/Metallica",
            "peso_liquido": "Peso Líquido",
            "peso_integral": "Peso Integral"
        }

        self.ordem_colunas = [
            "data", "nf", "fornecedor",
            "material_1", "material_2", "material_3", "material_4", "material_5",
            "produto", "custo_empresa", "ipi", "valor_integral",
            "valor_unitario_1", "valor_unitario_2", "valor_unitario_3",
            "valor_unitario_4", "valor_unitario_5",
            "duplicata_1", "duplicata_2", "duplicata_3", "duplicata_4",
            "duplicata_5", "duplicata_6",
            "valor_unitario_energia", "valor_mao_obra_tm_metallica",
            "peso_liquido", "peso_integral"
        ]

        self.colunas_ocultaveis = [
            "material_1", "material_2", "material_3", "material_4", "material_5",
            "valor_unitario_1", "valor_unitario_2", "valor_unitario_3", "valor_unitario_4", "valor_unitario_5",
            "duplicata_1", "duplicata_2", "duplicata_3", "duplicata_4", "duplicata_5", "duplicata_6"
        ]

        # visibilidade e larguras
        self.colunas_visiveis = list(self.colunas_fixas.keys())
        self.colunas_ocultas = []
        self.larguras_colunas = {col: 120 for col in self.colunas_fixas.keys()}

        # Mapeamento reverso (rótulo -> chave)
        self.display_to_key = {v: k for k, v in self.colunas_fixas.items()}

        # Treeview
        self.tree = ttk.Treeview(self, columns=self.colunas_visiveis, show="headings")
        self.configurar_treeview()

        # Scrollbars
        scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar_y.set)
        scrollbar_y.grid(row=0, column=4, sticky="ns")

        scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscroll=scrollbar_x.set)
        scrollbar_x.grid(row=1, column=0, columnspan=4, sticky="ew")

        # Frame para operações com linhas
        self.frame_operacoes_linhas = tk.LabelFrame(self, text="Operações com Linhas")
        self.frame_operacoes_linhas.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=5)

        self.btn_insert = tk.Button(self.frame_operacoes_linhas, text="Inserir Linha", command=self.inserir_linha)
        self.btn_insert.grid(row=0, column=0, padx=5, pady=5)

        self.btn_edit = tk.Button(self.frame_operacoes_linhas, text="Editar Linha", command=self.editar_linha)
        self.btn_edit.grid(row=0, column=1, padx=5, pady=5)

        self.btn_remove_items = tk.Button(self.frame_operacoes_linhas, text="Excluir Itens", command=self.remove_selected_items)
        self.btn_remove_items.grid(row=0, column=2, padx=5, pady=5)

        self.btn_media_ponderada = tk.Button(self.frame_operacoes_linhas, text="Calculadora de Média", command=self.chamar_calculadora_media_ponderada)
        self.btn_media_ponderada.grid(row=0, column=3, padx=5, pady=5)

        # Campo de pesquisa com trace
        self.search_var = tk.StringVar()
        self.trace_id = self.search_var.trace_add("write", self.formatar_data_em_tempo_real)

        self.lbl_pesquisa = tk.Label(self.frame_operacoes_linhas, text="Buscar:")
        self.lbl_pesquisa.grid(row=0, column=4, padx=5, pady=5, sticky="e")

        self.entry_pesquisa = tk.Entry(self.frame_operacoes_linhas, width=50, textvariable=self.search_var)
        self.entry_pesquisa.grid(row=0, column=5, padx=5, pady=5, sticky="ew")
        self.entry_pesquisa.bind("<Return>", lambda event: self.pesquisar())

        self.btn_pesquisar = tk.Button(self.frame_operacoes_linhas, text="Pesquisar", command=self.pesquisar)
        self.btn_pesquisar.grid(row=0, column=6, padx=5, pady=5)

        self.frame_operacoes_linhas.columnconfigure(4, weight=1)

        # Layout principal
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

        # Frame para operações com colunas (ocultar/mostrar)
        self.frame_operacoes_colunas = tk.LabelFrame(self, text="Exibir/Ocultar Colunas")
        self.frame_operacoes_colunas.grid(row=4, column=0, columnspan=4, sticky="ew", padx=10, pady=5)

        self.combobox_colunas = ttk.Combobox(
            self.frame_operacoes_colunas,
            values=[self.colunas_fixas[c] for c in self.colunas_ocultaveis],
            state="readonly"
        )
        self.combobox_colunas.grid(row=0, column=0, padx=10, pady=5)
        self.combobox_colunas.set("Selecione uma coluna")

        self.btn_ocultar = tk.Button(self.frame_operacoes_colunas, text="Ocultar Coluna", command=self.ocultar_coluna_selecionada)
        self.btn_ocultar.grid(row=0, column=1, padx=5, pady=5)

        self.btn_mostrar = tk.Button(self.frame_operacoes_colunas, text="Mostrar Coluna", command=self.mostrar_coluna_selecionada)
        self.btn_mostrar.grid(row=0, column=2, padx=5, pady=5)

        self.btn_voltar = tk.Button(self.frame_operacoes_colunas, text="Voltar", command=self.voltar_para_menu)
        self.btn_voltar.grid(row=0, column=3, padx=5, pady=5)

        self.frame_operacoes_colunas.grid_columnconfigure(4, weight=1)

        self.btn_exportar = tk.Button(self.frame_operacoes_colunas, text="Exportação Excel", command=self.abrir_dialogo_exportacao)
        self.btn_exportar.grid(row=0, column=5, padx=5, pady=5, sticky="e")

        # --- Botão Ajuda (colocado na frame_operacoes_colunas existente) ---
        self.help_frame = tk.Frame(self.frame_operacoes_colunas, bg="#f4f4f4")
        self.help_frame.grid(row=0, column=6, padx=(8, 12), pady=2, sticky="e")

        self.botao_ajuda_estoque = tk.Button(
            self.help_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=lambda: self._abrir_ajuda_estoque_modal()
        )
        self.botao_ajuda_estoque.pack(side="right", padx=(4, 6), pady=4)
        self.botao_ajuda_estoque.bind("<Enter>", lambda e: self.botao_ajuda_estoque.config(bg="#3b5566"))
        self.botao_ajuda_estoque.bind("<Leave>", lambda e: self.botao_ajuda_estoque.config(bg="#2c3e50"))

        # Tooltip seguro (se definir _create_tooltip)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_estoque,
                                     "Ajuda — Estoque: F1 ")
        except Exception:
            pass

        # Atalho F1 -> abre modal (bind no próprio Toplevel)
        try:
            self.bind_all("<F1>", lambda e: self._abrir_ajuda_estoque_modal())
        except Exception:
            pass

        self.dados_colunas_ocultas = {}

        # Carregamentos iniciais
        self.carregar_dados()
        self.carregar_estado_colunas()

        # Fechamento
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

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

            # determina wraplength
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

            # posição inicial: centrado horizontalmente sobre o widget, abaixo dele
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajuste horizontal
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # ajuste vertical: tenta abaixo; se não couber, mostra acima; senão limita
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

    def _abrir_ajuda_estoque_modal(self, contexto=None):
        """Modal de Ajuda para Entrada de NF / Estoque — instruções estendidas.
        Inclui explicação de campos (formatação automática, onde cadastrar novos itens, etc.)
        concentrados na seção 'Adicionar NF'."""
        try:
            modal = tk.Toplevel(self)
            modal.title("Ajuda — Entrada de NF / Estoque")
            modal.transient(self)
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
                aplicar_icone(modal, self.caminho_icone)
            except Exception:
                pass

            # Cabeçalho
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Entrada de Notas Fiscais (Estoque)", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # Corpo
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=300, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10, 6), padx=12)

            sections = [
                "Visão Geral",
                "Adicionar NF",
                "Editar NF",
                "Excluir NF",
                "Calculadora: Média Ponderada",
                "Pesquisa",
                "Colunas Dinâmicas (Ocultar / Mostrar)",
                "Exportar Excel",
                "FAQ"
            ]
            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            # Área de conteúdo
            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)
            content_frame.rowconfigure(0, weight=1)
            content_frame.columnconfigure(0, weight=1)

            # Painel geral
            general_frame = tk.Frame(content_frame, bg="white")
            general_frame.grid(row=0, column=0, sticky="nsew")
            general_frame.rowconfigure(0, weight=1)
            general_frame.columnconfigure(0, weight=1)

            txt_general = tk.Text(general_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_general = tk.Scrollbar(general_frame, command=txt_general.yview)
            txt_general.configure(yscrollcommand=sb_general.set)
            txt_general.grid(row=0, column=0, sticky="nsew")
            sb_general.grid(row=0, column=1, sticky="ns")

            # Painel específico para "Adicionar NF"
            adicionar_frame = tk.Frame(content_frame, bg="white")
            adicionar_frame.grid(row=0, column=0, sticky="nsew")
            adicionar_frame.rowconfigure(0, weight=1)
            adicionar_frame.columnconfigure(0, weight=1)

            txt_adicionar = tk.Text(adicionar_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_adicionar = tk.Scrollbar(adicionar_frame, command=txt_adicionar.yview)
            txt_adicionar.configure(yscrollcommand=sb_adicionar.set)
            txt_adicionar.grid(row=0, column=0, sticky="nsew")
            sb_adicionar.grid(row=0, column=1, sticky="ns")

            # Conteúdos
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "Esta janela é destinada à **Entrada de produto vindo do fornecedor** — ou seja, registrar o que entrou em estoque a partir de uma Nota Fiscal (NF). Esta tela NÃO realiza cálculos de custo aqui: as informações principais já vêm na própria NF (produto, peso, base, fornecedor, etc.).\n\n"
                "Fluxo típico: abrir Nova Entrada → informar dados da NF (data, número, fornecedor) → adicionar os produtos conforme a NF → salvar.\n\n"
                "AVISO IMPORTANTE — ORDEM CORRETA PARA ATUALIZAÇÃO DE CUSTO/ESTOQUE:\n"
                " - Depois de salvar a NF na tela Entrada, abra a janela 'Cálculo NFs' e confirme que o custo total e o estoque da(s) linha(s) estão corretos. Somente após confirmar/registrar a NF em 'Cálculo NFs' proceda para a janela 'Média Custo'.\n"
                " - Se você pular esse passo e for direto para 'Média Custo', a nota recém-entrada **não** será considerada no cálculo de custo nem atualizará corretamente o saldo de estoque.\n\n"
                "Observação: toda regra de custo/cálculo (se existir) é tratada em outros módulos; aqui gravamos a entrada conforme a NF."
            )

            contents["Adicionar NF"] = (
                "Adicionar NF — instruções completas\n\n"
                "1) Campos principais e formatação automática\n"
                "   - Data: não é necessário digitar as barras. Ex.: '01012025' será automaticamente formatado como '01/01/2025'.\n"
                "   - Campos numéricos: não é necessário digitar a vírgula. Basta digitar os números que o sistema formata sozinho. "
                "Ex.: '2500' vira '25,00'.\n"
                "Isso vale para os seguintes campos:\n"
                "       • custo da empresa\n"
                "       • IPI\n"
                "       • valor integral\n"
                "       • valor unitário 1 a 5\n"
                "       • duplicata 1 a 6\n"
                "       • valor unitário energia\n"
                "       • valor mão de obra TM/Metallica\n"
                "       • peso líquido\n"
                "       • peso integral\n"
                "   - Na calculadora (Peso / Valor) também não é necessário digitar a vírgula: o sistema faz a formatação automática.\n\n"
                "2) Materiais / Produtos / Fornecedor — onde cadastrar novos itens\n"
                "   - Caso queira que as entradas de Materiais, Produto ou Fornecedor tenham novos registros, abra a janela 'Materiais'. "
                "Lá você pode cadastrar fornecedores e materiais novos que aparecerão automaticamente nas listas suspensas desta tela.\n"
                "   - Se quiser cadastrar produtos novos, utilize a janela 'Base' (cadastro de produtos/base). Após salvar, o item aparecerá na lista suspensa de produtos.\n\n"
                "3) Adicionar vários produtos na mesma NF\n"
                "   - Na seção de 'Adicionar Produto', selecione o produto/base e informe o Peso.\n"
                "   - Clique em 'Adicionar Produto' para inserir na lista temporária da NF.\n"
                "   - Repita quantas vezes forem necessárias: cada produto se tornará uma linha na Treeview, todos vinculados ao mesmo número de NF.\n\n"
                "4) Salvar NF\n"
                "   - Depois de adicionar todos os produtos da nota, clique em 'Salvar NF' para gravar a nota e os itens no banco de dados.\n\n"
                "AVISO IMPORTANTE — FLUXO RECOMENDADO (Entrada → Cálculo NFs → Média Custo):\n"
                " - Após salvar a NF aqui, **não** vá diretamente para 'Média Custo'. Primeiro abra 'Cálculo NFs' e:\n"
                "    1) Localize a NF / as linhas correspondentes.\n"
                "    2) Verifique o custo total e confirme/subtraia/ajuste o estoque conforme necessário.\n"
                "    3) Confirme para que o sistema registre corretamente a entrada no módulo de cálculo.\n"
                " - Só após esses passos abra 'Média Custo' para que a janela de média inclua a NF recém-entrada no cálculo de custo e atualize o estoque corretamente.\n\n"
                "Dica: Se um item cadastrado recentemente (fornecedor, material ou produto) não aparecer na lista, feche e reabra a tela de Entrada para recarregar as opções."
            )

            contents["Editar NF"] = (
                "Editar NF — instruções\n\n"
                " • Selecione a linha (registro) correspondente à NF/produto que deseja editar. Os campos serão carregados para edição.\n"
                " • Altere somente os campos necessários (por exemplo: quantidade, base do produto, descrições) e confirme para atualizar os registros no banco.\n"
                " • Esta janela trata da entrada física conforme a NF; recursos específicos de observação/alerta (como linhas em vermelho) pertencem a outras telas (ex.: Saída NF).\n"
            )

            contents["Excluir NF"] = (
                "Excluir NF — instruções\n\n"
                " • Selecione uma ou mais linhas da NF que deseja remover e clique em 'Excluir'. Você será solicitado a confirmar a operação.\n"
                " • A exclusão remove o registro da entrada (ou marca como cancelada, dependendo da implementação que você escolher) — tome cuidado para não perder dados importantes."
            )

            contents["Calculadora: Média Ponderada"] = (
                "Calculadora — Média Ponderada (quando usar)\n\n"
                " • Use a média ponderada somente quando **a mesma NF contém mais de uma ocorrência do mesmo material** com preços diferentes.\n"
                " • Procedimento: selecione as linhas que representam o mesmo material dentro da NF e aplique a média ponderada — o sistema calculará o custo médio considerando peso/quantidade como fator de ponderação.\n"
                " • Em entradas simples (um único registro por material) não é necessário utilizar essa calculadora.\n"
                " • Nas entradas da calculadora (Peso e Valor) digite apenas dígitos: o campo formata a vírgula automaticamente."
            )

            contents["Pesquisa"] = (
                "Pesquisa — dicas de uso\n\n"
                " • Pesquise por Data (dd/mm/aaaa), NF, Fornecedor ou Produto.\n"
                " • Para conjuntos muito grandes, combine filtros (por exemplo data + fornecedor) para reduzir os resultados.\n"
                " • Use o botão 'Limpar' para restabelecer a visão completa."
            )

            contents["Colunas Dinâmicas (Ocultar / Mostrar)"] = (
                "Colunas Dinâmicas — ocultar e mostrar\n\n"
                " • Quando uma NF possui muitos materiais diferentes (ou se houver muitas colunas geradas por diferentes materiais), a visualização pode ficar poluída. Utilize a funcionalidade de ocultar/mostrar colunas para exibir apenas o que interessa.\n"
                " • Recomenda-se ocultar colunas de materiais pouco relevantes durante análises rápidas e restaurá-las quando precisar de detalhes."
            )

            contents["Exportar Excel"] = (
                "Exportar Excel — como usar\n\n"
                " • Clique em 'Exportação Excel', filtre o conjunto desejado (por data, NF, fornecedor, produto) e escolha onde salvar o arquivo (.xlsx).\n"
                " • O sistema exportará apenas as colunas visíveis e os registros filtrados — aplique filtros e ajuste colunas antes de exportar para um relatório mais enxuto."
            )

            contents["FAQ"] = (
                "FAQ — Rápido\n\n"
                "Q: Cadastrei um material e ele não apareceu na lista suspensa?\n"
                "A: Confirme se salvou o material na janela 'Materiais' e, se necessário, reabra a tela ou recarregue os dados nesta janela para atualizar as comboboxes.\n\n"
                "Q: Esta tela faz cálculos de custo automático?\n"
                "A: Não — esta tela registra a entrada conforme a NF. Qualquer cálculo de custo é feito em módulos específicos, quando aplicável."
            )

            # função que mostra a seção
            def mostrar_secao(key):
                if key == "Adicionar NF":
                    txt_adicionar.configure(state="normal")
                    txt_adicionar.delete("1.0", "end")
                    txt_adicionar.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_adicionar.configure(state="disabled")
                    txt_adicionar.yview_moveto(0)
                    adicionar_frame.tkraise()
                else:
                    txt_general.configure(state="normal")
                    txt_general.delete("1.0", "end")
                    txt_general.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_general.configure(state="disabled")
                    txt_general.yview_moveto(0)
                    general_frame.tkraise()

            # inicializa
            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])

            listbox.bind("<<ListboxSelect>>", on_select)

            # rodapé
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
            print("Erro ao abrir modal de ajuda (Estoque):", e)

    def abrir_dialogo_exportacao(self):
        """
        Abre uma janela de diálogo para que o usuário informe os filtros de exportação.
        Os filtros serão: Data (intervalo), NF, Fornecedor, Produto, Material 1, Material 2,
        Material 3, Material 4 e Material 5.
        """
        dialogo = tk.Toplevel(self)
        dialogo.title("Exportar NF - Filtros")
        dialogo.geometry("400x500")  # Ajustado para caber todos os campos
        dialogo.resizable(False, False)
        centralizar_janela(dialogo, 400, 500)
        aplicar_icone(dialogo, self.caminho_icone)
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

        # Linha 3: Fornecedor
        ttk.Label(frame, text="Fornecedor:", style="Custom.TLabel").grid(row=3, column=0, sticky="e", padx=5, pady=(5,2))
        entry_fornecedor = ttk.Entry(frame, width=35)
        entry_fornecedor.grid(row=3, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 4: Produto
        ttk.Label(frame, text="Produto:", style="Custom.TLabel").grid(row=4, column=0, sticky="e", padx=5, pady=(5,2))
        entry_produto = ttk.Entry(frame, width=35)
        entry_produto.grid(row=4, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 5: Material 1
        ttk.Label(frame, text="Material 1:", style="Custom.TLabel").grid(row=5, column=0, sticky="e", padx=5, pady=(5,2))
        entry_material1 = ttk.Entry(frame, width=35)
        entry_material1.grid(row=5, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 6: Material 2
        ttk.Label(frame, text="Material 2:", style="Custom.TLabel").grid(row=6, column=0, sticky="e", padx=5, pady=(5,2))
        entry_material2 = ttk.Entry(frame, width=35)
        entry_material2.grid(row=6, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 7: Material 3
        ttk.Label(frame, text="Material 3:", style="Custom.TLabel").grid(row=7, column=0, sticky="e", padx=5, pady=(5,2))
        entry_material3 = ttk.Entry(frame, width=35)
        entry_material3.grid(row=7, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 8: Material 4
        ttk.Label(frame, text="Material 4:", style="Custom.TLabel").grid(row=8, column=0, sticky="e", padx=5, pady=(5,2))
        entry_material4 = ttk.Entry(frame, width=35)
        entry_material4.grid(row=8, column=1, sticky="w", padx=5, pady=(5,2))

        # Linha 9: Material 5
        ttk.Label(frame, text="Material 5:", style="Custom.TLabel").grid(row=9, column=0, sticky="e", padx=5, pady=(5,2))
        entry_material5 = ttk.Entry(frame, width=35)
        entry_material5.grid(row=9, column=1, sticky="w", padx=5, pady=(5,2))

        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.grid(row=10, column=0, columnspan=2, pady=(15,5))

        def acao_exportar():
            filtro_data_inicial = entry_data_inicial.get().strip()
            filtro_data_final   = entry_data_final.get().strip()
            filtro_nf         = entry_nf.get().strip()
            filtro_fornecedor = entry_fornecedor.get().strip()
            filtro_produto    = entry_produto.get().strip()
            filtro_material1  = entry_material1.get().strip()
            filtro_material2  = entry_material2.get().strip()
            filtro_material3  = entry_material3.get().strip()
            filtro_material4  = entry_material4.get().strip()
            filtro_material5  = entry_material5.get().strip()

            # Processamento das datas:
            # Se ambos os campos estiverem preenchidos, utiliza intervalo;
            # Se apenas um for preenchido, filtra por >= ou <= conforme o campo.
            if filtro_data_inicial and filtro_data_final:
                try:
                    if "/" in filtro_data_inicial:
                        data_inicial = datetime.strptime(filtro_data_inicial, '%d/%m/%Y')
                    elif "-" in filtro_data_inicial:
                        data_inicial = datetime.strptime(filtro_data_inicial, '%Y-%m-%d')
                    else:
                        raise ValueError("Formato inválido para a data inicial.")

                    if "/" in filtro_data_final:
                        data_final = datetime.strptime(filtro_data_final, '%d/%m/%Y')
                    elif "-" in filtro_data_final:
                        data_final = datetime.strptime(filtro_data_final, '%Y-%m-%d')
                    else:
                        raise ValueError("Formato inválido para a data final.")
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido. Utilize dd/mm/yyyy ou yyyy-mm-dd.")
                    return

                filtro_data_inicial_fmt = data_inicial.strftime('%Y-%m-%d')
                filtro_data_final_fmt   = data_final.strftime('%Y-%m-%d')
            elif filtro_data_inicial:  # Apenas data inicial
                try:
                    if "/" in filtro_data_inicial:
                        data_inicial = datetime.strptime(filtro_data_inicial, '%d/%m/%Y')
                    elif "-" in filtro_data_inicial:
                        data_inicial = datetime.strptime(filtro_data_inicial, '%Y-%m-%d')
                    else:
                        raise ValueError("Formato inválido para a data inicial.")
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido para a data inicial. Utilize dd/mm/yyyy ou yyyy-mm-dd.")
                    return
                filtro_data_inicial_fmt = data_inicial.strftime('%Y-%m-%d')
                filtro_data_final_fmt = None
            elif filtro_data_final:  # Apenas data final
                try:
                    if "/" in filtro_data_final:
                        data_final = datetime.strptime(filtro_data_final, '%d/%m/%Y')
                    elif "-" in filtro_data_final:
                        data_final = datetime.strptime(filtro_data_final, '%Y-%m-%d')
                    else:
                        raise ValueError("Formato inválido para a data final.")
                except ValueError:
                    messagebox.showerror("Erro", "Formato de data inválido para a data final. Utilize dd/mm/yyyy ou yyyy-mm-dd.")
                    return
                filtro_data_inicial_fmt = None
                filtro_data_final_fmt = data_final.strftime('%Y-%m-%d')
            else:
                filtro_data_inicial_fmt = None
                filtro_data_final_fmt = None

            # Chama a função de exportação passando os filtros
            self.exportar_excel_filtrado_com_valores(
                filtro_data_inicial_fmt,
                filtro_data_final_fmt,
                filtro_nf,
                filtro_fornecedor,
                filtro_produto,
                filtro_material1,
                filtro_material2,
                filtro_material3,
                filtro_material4,
                filtro_material5
            )
            dialogo.destroy()

        export_button = ttk.Button(button_frame, text="Exportar Excel", command=acao_exportar)
        export_button.grid(row=0, column=0, padx=5)
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=dialogo.destroy)
        cancel_button.grid(row=0, column=1, padx=5)

        frame.columnconfigure(1, weight=1)

    def exportar_excel_filtrado_com_valores(self, filtro_data_inicial, filtro_data_final, filtro_nf, filtro_fornecedor, filtro_produto, filtro_material1, filtro_material2, filtro_material3, filtro_material4,filtro_material5):
        where_clauses = []
        parametros = []

        # Filtro de data (intervalo ou único)
        if filtro_data_inicial and filtro_data_final:
            where_clauses.append("data BETWEEN %s AND %s")
            parametros.extend([filtro_data_inicial, filtro_data_final])
        elif filtro_data_inicial:
            where_clauses.append("data >= %s")
            parametros.append(filtro_data_inicial)
        elif filtro_data_final:
            where_clauses.append("data <= %s")
            parametros.append(filtro_data_final)

        if filtro_nf:
            where_clauses.append("unaccent(nf) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_nf}%") 

        if filtro_fornecedor:
            where_clauses.append("unaccent(fornecedor) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_fornecedor}%")

        if filtro_produto:
            where_clauses.append("unaccent(produto) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_produto}%")

        if filtro_material1:
            where_clauses.append("unaccent(material_1) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_material1}%")

        if filtro_material2:
            where_clauses.append("unaccent(material_2) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_material2}%")

        if filtro_material3:
            where_clauses.append("unaccent(material_3) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_material3}%")

        if filtro_material4:
            where_clauses.append("unaccent(material_4) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_material4}%")

        if filtro_material5:
            where_clauses.append("unaccent(material_5) ILIKE unaccent(%s)")
            parametros.append(f"%{filtro_material5}%")

        where_sql = ""
        if where_clauses:
            where_sql = "WHERE " + " AND ".join(where_clauses)

        query = f"""
            SELECT data, nf, fornecedor, produto, material_1, material_2, material_3, material_4, material_5, custo_empresa, ipi, valor_integral, valor_unitario_1, valor_unitario_2, valor_unitario_3, valor_unitario_4, valor_unitario_5, duplicata_1, duplicata_2, duplicata_3, duplicata_4, duplicata_5, duplicata_6, valor_unitario_energia, valor_mao_obra_tm_metallica, peso_liquido, peso_integral
            FROM somar_produtos
            {where_sql}
            ORDER BY somar_produtos.data ASC, somar_produtos.nf ASC;
        """

        try:
            self.cursor.execute(query, parametros)
            dados = self.cursor.fetchall()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na consulta: {e}")
            return

        colunas = ["Data", "NF", "Fornecedor", "Produto", "Material 1", "Material 2", "Material 3", "Material 4", "Material 5", "Custo Empresa", "IPI", "Valor Integral", "Valor Unitario 1", "Valor Unitario 2", "Valor Unitario 3", "Valor Unitario 4", "Valor Unitario 5", "Duplicata 1", "Duplicata 2", "Duplicata 3", "Duplicata 4", "Duplicata 5", "Duplicata 6", "Valor Unitário Energia", "Valor Mão de Obra TM/Metallica", "Peso Líquido", "Peso Integral"]
        df = pd.DataFrame(dados, columns=colunas)

        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Relatorio_Entrada_Nf.xlsx"
        )

        if not caminho_arquivo:
            return

        try:
            df.to_excel(caminho_arquivo, index=False)
            messagebox.showinfo("Exportação", f"Exportação concluída com sucesso.\nArquivo salvo em:\n{caminho_arquivo}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

    def chamar_calculadora_media_ponderada(self):
            import Calculadora_Sistema as calc  # Importa o arquivo da calculadora
            calc.CalculadoraMediaPonderada()  # Chama a função que cria a janela da calculadora

    def configurar_treeview(self):
        style = ttk.Style(self)
        style.theme_use("alt")
        style.configure("Treeview", 
                        background="white", 
                        foreground="black", 
                        rowheight=25, 
                        fieldbackground="white")
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.map("Treeview", 
                background=[("selected", "#0078D7")], 
                foreground=[("selected", "white")])
        
        self.tree.config(columns=self.colunas_visiveis, style="Treeview", height=22)
        for coluna in self.colunas_visiveis:
            self.tree.heading(coluna, text=self.colunas_fixas[coluna], anchor="center")
            self.tree.column(coluna, width=self.larguras_colunas.get(coluna, 250), anchor="center", minwidth=250)

    def atualizar_treeview(self):
        # Limpa o Treeview antes de atualizar as colunas
        self.tree.delete(*self.tree.get_children())

        # Atualiza as colunas do Treeview
        self.tree["columns"] = self.colunas_fixas

        for coluna in self.colunas_fixas:
            self.tree.heading(coluna, text=coluna.replace('"', ''), anchor="center")
            self.tree.column(coluna, width=100, stretch=tk.NO, anchor="center")

        # Coleta dados do banco de dados
        self.cursor.execute(f"SELECT {', '.join(self.colunas_fixas)} FROM somar_produtos")
        rows = self.cursor.fetchall()

        for row in rows:
            valores_a_inserir = [str(val) if val is not None else '' for val in row]
            self.tree.insert("", "end", values=valores_a_inserir)

    def ocultar_coluna(self, nome_coluna):
        """Oculta a coluna no Treeview junto com os dados, sem afetar as outras colunas."""
        if nome_coluna in self.colunas_ocultas:
            messagebox.showwarning("Aviso", f"A coluna '{self.colunas_fixas[nome_coluna]}' já está oculta.")
        elif nome_coluna in self.colunas_visiveis:
            # Armazena os dados da coluna antes de ocultá-la
            dados_coluna = [self.tree.set(item, nome_coluna) for item in self.tree.get_children()]
            self.dados_colunas_ocultas[nome_coluna] = dados_coluna

            # Remove a coluna da lista de visíveis e a adiciona à lista de ocultas
            self.colunas_visiveis.remove(nome_coluna)
            self.colunas_ocultas.append(nome_coluna)

            # Reconfigura o Treeview para incluir a nova coluna visível
            self.configurar_treeview()

            # Salva o estado das colunas e recarrega os dados para atualizar a exibição
            self.salvar_estado_colunas()
            self.carregar_dados()  # Recarrega os dados para mostrar as informações da nova coluna

    def mostrar_coluna(self, nome_coluna):
        """Mostra a coluna no Treeview na posição correta e com dados atualizados."""
        if nome_coluna in self.colunas_visiveis:
            # Exibe mensagem de aviso se a coluna já estiver visível
            messagebox.showwarning("Aviso", f"A coluna '{self.colunas_fixas[nome_coluna]}' já está visível.")
        elif nome_coluna in self.colunas_ocultas:
            # Remove a coluna da lista de ocultas e a adiciona à lista de visíveis
            self.colunas_ocultas.remove(nome_coluna)
            self.colunas_visiveis.append(nome_coluna)
            
            # Ordena as colunas visíveis conforme a ordem predefinida
            self.colunas_visiveis = sorted(self.colunas_visiveis, key=lambda x: self.ordem_colunas.index(x))

            # Reconfigura o Treeview para incluir a nova coluna visível
            self.configurar_treeview()

            # Salva o estado das colunas e recarrega os dados para atualizar a exibição
            self.salvar_estado_colunas()
            self.carregar_dados()  # Recarrega os dados para mostrar as informações da nova coluna

    def ocultar_coluna_selecionada(self):
        """Oculta a coluna selecionada na combobox."""
        nome_amigavel = self.combobox_colunas.get()
        if nome_amigavel != "Selecione uma coluna":
            # Encontra a chave da coluna com o nome amigável selecionado
            nome_coluna = [chave for chave, valor in self.colunas_fixas.items() if valor == nome_amigavel][0]
            if nome_coluna in self.colunas_ocultaveis:  # Verifica se a coluna pode ser oculta
                self.ocultar_coluna(nome_coluna)

    def mostrar_coluna_selecionada(self):
        """Mostra a coluna selecionada na combobox."""
        nome_amigavel = self.combobox_colunas.get()
        if nome_amigavel != "Selecione uma coluna":
            # Encontra a chave da coluna com o nome amigável selecionado
            nome_coluna = [chave for chave, valor in self.colunas_fixas.items() if valor == nome_amigavel][0]
            if nome_coluna in self.colunas_ocultaveis:  # Verifica se a coluna pode ser mostrada
                self.mostrar_coluna(nome_coluna)

    def carregar_estado_colunas(self):
        """Carrega o estado das colunas (ocultas) do arquivo dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)
        caminho_arquivo = os.path.join(caminho_base, "config", "estado_colunas.json")

        if os.path.exists(caminho_arquivo):
            with open(caminho_arquivo, "r") as f:
                estado = json.load(f)
                self.colunas_ocultas = estado.get("colunas_ocultas", [])
                self.colunas_visiveis = [col for col in self.colunas_visiveis if col not in self.colunas_ocultas]
                self.configurar_treeview()

    def salvar_estado_colunas(self):
        """Salva o estado das colunas (ocultas) em um arquivo dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)  # Diretório do script
        caminho_pasta = os.path.join(caminho_base, "config")  # Criar pasta config
        os.makedirs(caminho_pasta, exist_ok=True)  # Criar a pasta se não existir

        caminho_arquivo = os.path.join(caminho_pasta, "estado_colunas.json")  # Caminho completo

        with open(caminho_arquivo, "w") as f:
            estado = {"colunas_ocultas": self.colunas_ocultas}
            json.dump(estado, f)
            
    def formatar_data_em_tempo_real(self, *args):
        """
        Se o conteúdo for composto apenas por 8 dígitos,
        formata-o como data (dd/mm/aaaa) em tempo real.
        """
        conteudo = self.search_var.get()
        # Extrai apenas os dígitos
        digitos = ''.join(ch for ch in conteudo if ch.isdigit())
        
        # Se houver exatamente 8 dígitos, formata a data
        if len(digitos) == 8:
            novo_conteudo = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
        else:
            novo_conteudo = conteudo

        # Se o novo conteúdo for diferente, atualiza a variável sem causar recursão
        if novo_conteudo != conteudo:
            self.search_var.trace_remove("write", self.trace_id)
            self.search_var.set(novo_conteudo)
            self.trace_id = self.search_var.trace_add("write", self.formatar_data_em_tempo_real)

    def formatar_numero(self, event, casas_decimais=2):
        entry = event.widget
        texto = entry.get()
        apenas_digitos = ''.join(filter(str.isdigit, texto))
        if not apenas_digitos:
            entry.delete(0, tk.END)
            return
        valor_int = int(apenas_digitos)
        valor = Decimal(valor_int) / (10 ** casas_decimais)
        s = f"{valor:,.{casas_decimais}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        entry.delete(0, tk.END)
        entry.insert(0, s)

    def pesquisar(self):
        """Filtra os dados no Treeview com base no termo de pesquisa."""
        # Obtém o termo de pesquisa (já formatado em tempo real se for data)
        termo_pesquisa = self.entry_pesquisa.get().strip().lower()
        print(f"Termo de pesquisa: {termo_pesquisa}")  # Debug

        # Limpa o Treeview
        self.tree.delete(*self.tree.get_children())

        # Define as colunas visíveis para a consulta
        colunas_visiveis = [
            "data", "nf", "fornecedor", "material_1", "material_2", "material_3", "material_4", "material_5",
            "produto", "custo_empresa", "ipi", "valor_integral", "valor_unitario_1", "valor_unitario_2", 
            "valor_unitario_3", "valor_unitario_4", "valor_unitario_5", "duplicata_1", "duplicata_2", "duplicata_3", 
            "duplicata_4", "duplicata_5", "duplicata_6", "valor_unitario_energia", "valor_mao_obra_tm_metallica", 
            "peso_liquido", "peso_integral"
        ]
        # Cria a string com os nomes das colunas para a consulta SQL
        colunas_query = ', '.join(f'"{col}"' for col in colunas_visiveis)

        termo_data_sql = None
        data_filtro = False

        # Tenta converter o termo para uma data no formato dd/mm/yyyy
        try:
            termo_data = datetime.strptime(termo_pesquisa, "%d/%m/%Y")
            termo_data_sql = termo_data.strftime("%Y-%m-%d")  # Formata para o padrão SQL
            data_filtro = True
        except ValueError:
            pass  # Se não for uma data, segue a busca por texto

        if data_filtro:
            # Consulta para data
            query = f"""
                SELECT {colunas_query} FROM somar_produtos
                WHERE "data" = %s
            """
            termo = termo_data_sql
            print(f"Query de pesquisa por data: {query}")  # Debug
            self.cursor.execute(query, (termo,))
        else:
            # Consulta para busca por texto (fornecedor, produto, nf)
            query = f"""
                SELECT {colunas_query} FROM somar_produtos
                WHERE unaccent("fornecedor") ILIKE unaccent(%s)
                    OR unaccent("produto") ILIKE unaccent(%s)
                    OR unaccent("nf") ILIKE unaccent(%s)
            """
            termo = f"%{termo_pesquisa}%"
            print(f"Query de pesquisa: {query}")  # Debug
            self.cursor.execute(query, (termo, termo, termo))

        rows = self.cursor.fetchall()

        # Se não houver resultados, exibe uma mensagem
        if not rows:
            messagebox.showinfo("Resultados", "Nenhum resultado encontrado.")
            return

        print(f"Linhas encontradas: {len(rows)}")  # Debug
        print(f"Dados retornados: {rows}")  # Debug

        # Processa os dados para exibir no Treeview
        for row in rows:
            valores_formatados = []
            for coluna in colunas_visiveis:
                # Obtém o índice da coluna na lista
                index = colunas_visiveis.index(coluna)
                valor = row[index]  # Acessa o valor da linha correspondente à coluna

                # Pula as colunas que não devem ser exibidas
                if coluna in self.colunas_ocultas:
                    continue

                print(f"Coluna: {coluna}, Valor: {valor}")  # Debug

                if coluna == "data" and valor:
                    # Formata a data para dd/mm/yyyy
                    valor_formatado = valor.strftime("%d/%m/%Y")
                elif ("valor" in coluna or "custo" in coluna or "duplicata" in coluna or 
                      coluna == "peso_liquido" or coluna == "ipi"):
                    # Formata valores numéricos
                    if valor is not None:
                        valor_formatado = "{:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
                    else:
                        valor_formatado = ""
                else:
                    valor_formatado = str(valor) if valor is not None else ""

                valores_formatados.append(valor_formatado)

            print(f"Valores a inserir: {valores_formatados}")  # Debug
            self.tree.insert("", "end", values=valores_formatados)

    def carregar_dados(self):
        """Carrega os dados do banco de dados no Treeview."""
        # Carrega o estado das colunas ocultas para definir as colunas visíveis
        self.carregar_estado_colunas()
        
        # Limpa o Treeview antes de carregar novos dados
        self.tree.delete(*self.tree.get_children())

        # Obtém apenas as colunas visíveis
        colunas_visiveis = [coluna for coluna in self.colunas_fixas.keys() if coluna in self.colunas_visiveis]

        # Adiciona ordenação por data na consulta
        try:
            query = f"""
                SELECT {', '.join(colunas_visiveis)} 
                FROM somar_produtos 
                ORDER BY data DESC, nf::INTEGER DESC;
            """
            self.cursor.execute(query)
            rows = self.cursor.fetchall()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
            return

        for row in rows:
            valores_a_inserir = []

            for i, coluna in enumerate(colunas_visiveis):
                valor = row[i]

                # Se o valor for None, insere uma string vazia
                if valor is None:
                    valores_a_inserir.append('')
                    continue

                # Converte a data para o formato desejado, caso seja a primeira coluna
                if i == 0:  # Supondo que a primeira coluna seja a data
                    valor = valor.strftime("%d/%m/%Y")
                else:
                    # Formata valores numéricos
                    if coluna == "peso_liquido" or coluna == "peso_integral":
                        # Formata a coluna peso_liquido
                        valor_decimal = Decimal(valor)  # Converte para Decimal
                        valor_formatado = "{:.3f}".format(valor_decimal).replace(".", ",")  # Exibe com 3 casas decimais
                        valores_a_inserir.append(valor_formatado)
                        continue
                    # Caso especial: IPI
                    elif coluna == "ipi":
                        # valor já é float/Decimal; formata com 2 casas e vírgula
                        ipi_dec = Decimal(valor)
                        ipi_fmt = "{:.2f}".format(ipi_dec).replace('.', ',')
                        valores_a_inserir.append(ipi_fmt)
                        continue
                    elif "valor" in coluna or "custo" in coluna or "duplicata" in coluna:
                        # Formatação padrão para as demais colunas
                        valor_formatado = "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")
                        valores_a_inserir.append(valor_formatado)
                        continue

                valores_a_inserir.append(valor)

            # Insere os valores processados no Treeview
            self.tree.insert("", "end", values=valores_a_inserir)

    def carregar_produtos(self):
        try:
            self.cursor.execute("SELECT nome FROM produtos")
            produtos = [row[0] for row in self.cursor.fetchall()]
            return produtos
        except Exception as e:
            print("Erro ao carregar produtos do banco de dados:", e)
            return []
        
    def carregar_fornecedores_materiais(self, tipo):
        """Carrega fornecedores ou materiais do banco de dados."""
        if tipo == "fornecedor":
            query = "SELECT DISTINCT fornecedor FROM materiais"
        elif tipo == "material":
            query = "SELECT DISTINCT nome FROM materiais"
        
        # Conecte-se ao banco de dados e execute a query
        cursor = self.conn.cursor()
        cursor.execute(query)
        resultados = cursor.fetchall()
        cursor.close()
        
        # Retorna uma lista com os valores extraídos
        return [linha[0] for linha in resultados]
        
    def inserir_linha(self):
        """Método para inserir uma nova linha no banco de dados."""
        janela_inserir = tk.Toplevel(self)
        janela_inserir.title("Inserir Nova Linha")

        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(janela_inserir, caminho_icone)

        # Função auxiliar para centralizar a janela
        def centralizar_janela3(janela):
            janela.update_idletasks()
            largura_janela = janela.winfo_width()
            altura_janela = janela.winfo_height()
            largura_tela = janela.winfo_screenwidth()
            altura_tela = janela.winfo_screenheight()
            x = (largura_tela // 2) - (largura_janela // 2)
            y = (altura_tela // 2) - (altura_janela // 2)
            janela.geometry(f'{largura_janela}x{altura_janela}+{x}+{y}')

        # Configurar tamanho e centralizar a janela
        janela_inserir.geometry("1000x600")
        janela_inserir.update_idletasks()
        centralizar_janela3(janela_inserir)

        # Definir fundo da janela
        janela_inserir.configure(bg="#ecf0f1")

        # Fonte dos títulos das entradas
        LABEL_FONT = ("Arial", 9, "bold")

        # Variável para o campo de data
        data_var = tk.StringVar()

        # Função para formatar a data
        def formatar_data(event=None):
            data = data_var.get().strip()
            data = ''.join(filter(str.isdigit, data))
            if len(data) > 8:
                data = data[:8]
            if len(data) == 8:
                data = data[:2] + '/' + data[2:4] + '/' + data[4:8]
            data_var.set(data)

        # Estilos exclusivos e locais para esta janela
        style = ttk.Style(janela_inserir)
        # NÃO mudar o tema global, apenas configure estilos locais
        style.configure("Inserir.TCombobox", padding=5, relief="flat",
                        background="#ecf0f1", font=("Arial", 10))
        style.configure("Inserir.TEntry", padding=5, relief="solid",
                        background="#ecf0f1", font=("Arial", 10))

        # Título da janela
        titulo = tk.Label(janela_inserir, text="Inserir Nova Linha",
                        bg="#34495e", fg="white",
                        font=("Arial", 16, "bold"))
        titulo.pack(fill=tk.X, pady=(0, 10))

        # Frame principal
        frame_campos = tk.Frame(janela_inserir, bg="#ecf0f1")
        frame_campos.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)

        # Carregar dados
        fornecedores = self.carregar_fornecedores_materiais("fornecedor")
        materiais = self.carregar_fornecedores_materiais("material")

        # Lista de títulos amigáveis
        titulos_amigaveis_visiveis = [self.colunas_fixas[col] for col in self.colunas_visiveis]

        # Organiza em três colunas
        num_colunas = 3
        self.entries = {}
        for i, titulo_text in enumerate(titulos_amigaveis_visiveis):
            linha = i // num_colunas
            coluna = i % num_colunas

            # Label com fonte reduzida
            label = tk.Label(frame_campos, text=titulo_text, bg="#ecf0f1",
                            font=LABEL_FONT)
            label.grid(row=linha, column=coluna * 2, padx=10, pady=8, sticky="e")

            # Escolha do widget
            if titulo_text == "Fornecedor":
                fornecedores = sorted(self.carregar_fornecedores_materiais("fornecedor"))
                entry = ttk.Combobox(frame_campos, values=[""] + fornecedores,
                                    style="Inserir.TCombobox", width=22, state="readonly")
            elif titulo_text.startswith("Material"):
                materiais = sorted(self.carregar_fornecedores_materiais("material"))
                entry = ttk.Combobox(frame_campos, values=[""] + materiais,
                                    style="Inserir.TCombobox", width=22, state="readonly")
            elif titulo_text == "Produto":
                produtos = sorted(self.carregar_produtos())
                entry = ttk.Combobox(frame_campos, values=[""] + produtos,
                                    style="Inserir.TCombobox", width=22, state="readonly")
            elif titulo_text == "Data":
                entry = ttk.Entry(frame_campos, textvariable=data_var,
                                style="Inserir.TEntry", width=24)
                entry.bind("<FocusOut>", formatar_data)
            else:
                entry = ttk.Entry(frame_campos, style="Inserir.TEntry", width=24)
                if any(k in titulo_text for k in ("Valor", "Duplicata", "Peso", "Custo Empresa", "IPI")):
                    casas = 3 if "Peso" in titulo_text else 2
                    entry.bind("<KeyRelease>", lambda e, c=casas: self.formatar_numero(e, c))

            entry.grid(row=linha, column=coluna * 2 + 1, padx=10, pady=8)
            self.entries[titulo_text] = entry

        # Frame para botões
        frame_botoes = tk.Frame(janela_inserir, bg="#ecf0f1")
        frame_botoes.pack(pady=20)

        btn_confirmar = tk.Button(frame_botoes, text="Confirmar",
                                command=lambda: self.confirmar_inserir(janela_inserir),
                                bg="#27ae60", fg="white",
                                font=("Arial", 12, "bold"), width=15)
        btn_confirmar.grid(row=0, column=0, padx=10)

        btn_cancelar = tk.Button(frame_botoes, text="Cancelar",
                                command=janela_inserir.destroy,
                                bg="#c0392b", fg="white",
                                font=("Arial", 12, "bold"), width=15)
        btn_cancelar.grid(row=0, column=1, padx=10)

        # Ajuste de colunas
        for coluna in range(num_colunas * 2):
            frame_campos.grid_columnconfigure(coluna, weight=1)

    def confirmar_inserir(self, janela_inserir):
        """Confirma a inserção de uma nova linha."""
        # Coleta os dados dos campos de entrada, removendo espaços e quebras de linha
        dados_inseridos = {
            label: entry.get().strip().replace('\n', ' ').replace('\r', '')
            for label, entry in self.entries.items()
        }

        # Campos obrigatórios
        required_fields = ["Data", "NF", "Fornecedor", "Produto", "Peso Líquido"]
        for campo in required_fields:
            if not dados_inseridos.get(campo):
                messagebox.showerror("Erro", f"O campo '{campo}' é obrigatório.")
                return

        nf = dados_inseridos.get("NF")

        # Verifica se a NF já existe no banco de dados
        try:
            query_verificar_nf = "SELECT COUNT(*) FROM somar_produtos WHERE nf = %s"
            self.cursor.execute(query_verificar_nf, (nf,))
            nf_existente = self.cursor.fetchone()[0]

            if nf_existente > 0:
                messagebox.showerror("Erro", f"A nota fiscal '{nf}' já está cadastrada!")
                return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar nota fiscal: {e}")
            return

        valores = []
        # Lista de campos numéricos que, se vazios, receberão 0
        numeric_fields_default_zero = [
            "Valor Integral", "Valor Unitário 1", "Valor Unitário 2", "Valor Unitário 3",
            "Valor Unitário 4", "Valor Unitário 5", "Duplicata 1", "Duplicata 2", "Duplicata 3",
            "Duplicata 4", "Duplicata 5", "Duplicata 6", "Valor Unitário Energia",
            "Valor Mão de Obra TM/Metallica", "Peso Integral"
        ]

        # Percorre as colunas conforme o mapeamento (usando self.colunas_fixas)
        for coluna in self.colunas_fixas.keys():
            coluna_amigavel = self.colunas_fixas[coluna]

            if coluna_amigavel in dados_inseridos:
                valor = dados_inseridos[coluna_amigavel]

                # Converter data
                if coluna == "data":
                    try:
                        valor = datetime.strptime(valor, "%d/%m/%Y").date()
                    except ValueError:
                        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/YYYY.")
                        return

                # Tratamento especial para o campo IPI: aceita vírgula e remove o símbolo de porcentagem
                elif coluna == "ipi":
                    if valor:
                        valor = valor.replace("%", "").replace(",", ".")
                        try:
                            valor_float = float(valor)
                            valores.append(valor_float)
                        except ValueError:
                            messagebox.showerror("Erro", f"Valor inválido para o campo '{coluna_amigavel}'")
                            return
                        continue

                # Processar campos numéricos (exceto IPI, já tratado)
                elif "valor" in coluna or "custo" in coluna or "duplicata" in coluna or coluna in ["peso_liquido", "peso_integral"]:
                    if valor:
                        valor = valor.replace(".", "").replace(",", ".")
                        try:
                            valor_float = float(valor)
                            valores.append(valor_float)
                            # Exibe debug formatado (opcional)
                            valor_formatado = "{:,.2f}".format(valor_float).replace(",", "X").replace(".", ",").replace("X", ".")
                            print(f"Valor formatado para exibição: {valor_formatado}")
                        except ValueError:
                            messagebox.showerror("Erro", f"Valor inválido para o campo '{coluna_amigavel}'")
                            return
                        continue

                # Se o campo estiver vazio e for numérico, atribui 0; caso contrário, None ou o próprio valor
                if valor == "":
                    if coluna_amigavel in numeric_fields_default_zero:
                        valores.append(0)
                    else:
                        valores.append(None)
                else:
                    valores.append(valor)
            else:
                valores.append(None)

        try:
            self.cursor.execute("""
                SELECT COALESCE(MIN(t1.id + 1), 1) AS menor_id_disponivel
                FROM somar_produtos t1
                LEFT JOIN somar_produtos t2 ON t1.id + 1 = t2.id
                WHERE t2.id IS NULL
            """)
            menor_id_disponivel = self.cursor.fetchone()[0]

            print("Menor ID disponível encontrado:", menor_id_disponivel)

            colunas_fixas_list = list(self.colunas_fixas.keys())
            if menor_id_disponivel is not None:
                colunas_fixas_list.insert(0, 'id')
                valores.insert(0, menor_id_disponivel)

            query = f"""
                INSERT INTO somar_produtos ({', '.join(colunas_fixas_list)}) 
                VALUES ({', '.join(['%s'] * len(valores))})
            """
            print("Query de inserção:", query)
            print("Valores de inserção:", valores)

            self.cursor.execute(query, valores)
            self.conn.commit()

            self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            self.conn.commit()

            messagebox.showinfo("Sucesso", "Linha inserida com sucesso!")
            self.carregar_dados()  # Atualiza o Treeview

            if self.janela_menu and hasattr(self.janela_menu, 'frame_nf'):
            # limpa widgets antigos
                for widget in self.janela_menu.frame_nf.winfo_children():
                    widget.destroy()
            # repopula com dados atualizados
            self.janela_menu.criar_relatorio_nf()

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao inserir dados: {e}")

        janela_inserir.destroy()

    def editar_linha(self):
        """Método para editar uma linha selecionada no Treeview."""
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            messagebox.showwarning("Seleção", "Selecione uma linha para editar.")
            return

        item_id = item_selecionado[0]
        valores_atualizados = self.tree.item(item_id, 'values')
        nf_original = valores_atualizados[1]  # NF é o segundo valor

        # Cria a janela
        janela_editar = tk.Toplevel(self)
        janela_editar.title("Editar Linha")
        aplicar_icone(janela_editar, self.caminho_icone)

        # Centralizar janela
        def centralizar_janela(janela):
            janela.update_idletasks()
            largura_janela = janela.winfo_width()
            altura_janela = janela.winfo_height()
            largura_tela = janela.winfo_screenwidth()
            altura_tela = janela.winfo_screenheight()
            x = (largura_tela // 2) - (largura_janela // 2)
            y = (altura_tela // 2) - (altura_janela // 2)
            janela.geometry(f'{largura_janela}x{altura_janela}+{x}+{y}')

        janela_editar.geometry("1000x600")
        janela_editar.update_idletasks()
        centralizar_janela(janela_editar)

        janela_editar.configure(bg="#ecf0f1")

        # Fonte dos títulos
        LABEL_FONT = ("Arial", 9, "bold")

        # Estilos exclusivos e locais
        style = ttk.Style(janela_editar)
        # NÃO alterar o tema global
        style.configure("Editar.TCombobox", padding=5, relief="flat",
                        background="#ecf0f1", font=("Arial", 10))
        style.configure("Editar.TEntry", padding=5, relief="solid",
                        background="#ecf0f1", font=("Arial", 10))

        # Função para formatar a data
        def formatar_data(event=None):
            data = event.widget.get().strip()
            data = ''.join(filter(str.isdigit, data))
            if len(data) > 8:
                data = data[:8]
            if len(data) == 8:
                data = data[:2] + '/' + data[2:4] + '/' + data[4:8]
            event.widget.delete(0, tk.END)
            event.widget.insert(0, data)

        # Título
        titulo = tk.Label(janela_editar, text="Editar Linha", bg="#34495e", fg="white",
                        font=("Arial", 16, "bold"))
        titulo.pack(fill=tk.X, pady=(0, 10))

        # Frame principal
        frame_campos = tk.Frame(janela_editar, bg="#ecf0f1")
        frame_campos.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)

        # Colunas visíveis
        titulos_amigaveis_visiveis = [self.colunas_fixas[col] for col in self.colunas_visiveis]

        # Carregar listas para combobox
        fornecedores = self.carregar_fornecedores_materiais("fornecedor")
        materiais = self.carregar_fornecedores_materiais("material")
        produtos = self.carregar_produtos()

        num_colunas = 3
        self.entries_edit = {}

        for i, titulo_text in enumerate(titulos_amigaveis_visiveis):
            linha = i // num_colunas
            coluna = i % num_colunas

            # Label com fonte reduzida
            label = tk.Label(frame_campos, text=titulo_text, bg="#ecf0f1", font=LABEL_FONT)
            label.grid(row=linha, column=coluna * 2, padx=10, pady=8, sticky="e")

            # Escolhe widget
            if titulo_text == "Fornecedor":
                fornecedores = sorted(self.carregar_fornecedores_materiais("fornecedor"))
                entry = ttk.Combobox(frame_campos, values=[""] + fornecedores,
                                    style="Editar.TCombobox", width=22, state="readonly")
                entry.set(valores_atualizados[i])
            elif titulo_text == "IPI":
                entry = ttk.Entry(frame_campos, style="Editar.TEntry", width=24)
                valor_ipi = valores_atualizados[i]
                entry.insert(0, valor_ipi.replace('.', ','))
                entry.bind("<KeyRelease>", lambda e, c=2: self.formatar_numero(e, c))
            elif titulo_text in ("Peso Líquido", "Peso Integral"):
                entry = ttk.Entry(frame_campos, style="Editar.TEntry", width=24)
                raw = valores_atualizados[i]
                try:
                    dec = Decimal(raw.replace('.', '').replace(',', '.'))
                    fmt = f"{dec:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except Exception:
                    fmt = raw
                entry.insert(0, fmt)
                entry.bind("<KeyRelease>", lambda e: self.formatar_numero(e, 3))
            elif titulo_text.startswith("Material"):
                materiais = sorted(self.carregar_fornecedores_materiais("material"))
                entry = ttk.Combobox(frame_campos, values=[""] + materiais,
                                    style="Editar.TCombobox", width=22, state="readonly")
                entry.set(valores_atualizados[i])
            elif titulo_text == "Produto":
                produtos = sorted(self.carregar_produtos())
                entry = ttk.Combobox(frame_campos, values=[""] + produtos,
                                    style="Editar.TCombobox", width=22, state="readonly")
                entry.set(valores_atualizados[i])
            elif titulo_text == "Data":
                entry = ttk.Entry(frame_campos, width=24)
                entry.insert(0, valores_atualizados[i])
                entry.bind("<KeyRelease>", formatar_data)
            else:
                entry = ttk.Entry(frame_campos, style="Editar.TEntry", width=24)
                entry.insert(0, valores_atualizados[i])
                if any(k in titulo_text for k in ("Valor", "Duplicata", "Peso", "Custo Empresa", "IPI")):
                    casas = 3 if "Peso" in titulo_text else 2
                    entry.bind("<KeyRelease>", lambda e, c=casas: self.formatar_numero(e, c))

            entry.grid(row=linha, column=coluna * 2 + 1, padx=10, pady=8)
            self.entries_edit[titulo_text] = entry

        # Botões
        frame_botoes = tk.Frame(janela_editar, bg="#ecf0f1")
        frame_botoes.pack(pady=20)

        btn_confirmar = tk.Button(frame_botoes, text="Confirmar",
                                command=lambda: self.confirmar_editar(item_id, nf_original, janela_editar),
                                bg="#27ae60", fg="white", font=("Arial", 12, "bold"), width=15)
        btn_confirmar.grid(row=0, column=0, padx=10)

        btn_cancelar = tk.Button(frame_botoes, text="Cancelar",
                                command=janela_editar.destroy,
                                bg="#c0392b", fg="white", font=("Arial", 12, "bold"), width=15)
        btn_cancelar.grid(row=0, column=1, padx=10)

        for coluna in range(num_colunas * 2):
            frame_campos.grid_columnconfigure(coluna, weight=1)

    def confirmar_editar(self, item_id, nf_original, janela_edit):
        """Confirma a edição de uma linha."""
        # Coleta os dados atualizados dos campos de entrada
        dados_atualizados = {
            label: entry.get().strip().replace('\n', ' ').replace('\r', '')
            for label, entry in self.entries_edit.items()
        }

        # Verifica os campos obrigatórios
        required_fields = ["Data", "NF", "Fornecedor", "Produto", "Peso Líquido"]
        for campo in required_fields:
            if not dados_atualizados.get(campo):
                messagebox.showerror("Erro", f"O campo '{campo}' é obrigatório.")
                return

        # Mapeamento dos nomes amigáveis para os nomes reais do banco de dados
        mapeamento_colunas = {
            "Data": "data",
            "NF": "nf",
            "Fornecedor": "fornecedor",
            "Material 1": "material_1",
            "Material 2": "material_2",
            "Material 3": "material_3",
            "Material 4": "material_4",
            "Material 5": "material_5",
            "Produto": "produto",
            "Custo Empresa": "custo_empresa",
            "Valor Integral": "valor_integral",
            "Valor Unitário 1": "valor_unitario_1",
            "Valor Unitário 2": "valor_unitario_2",
            "Valor Unitário 3": "valor_unitario_3",
            "Valor Unitário 4": "valor_unitario_4",
            "Valor Unitário 5": "valor_unitario_5",
            "IPI": "ipi",
            "Duplicata 1": "duplicata_1",
            "Duplicata 2": "duplicata_2",
            "Duplicata 3": "duplicata_3",
            "Duplicata 4": "duplicata_4",
            "Duplicata 5": "duplicata_5",
            "Duplicata 6": "duplicata_6",
            "Valor Unitário Energia": "valor_unitario_energia",
            "Valor Mão de Obra TM/Metallica": "valor_mao_obra_tm_metallica",
            "Peso Líquido": "peso_liquido",
            "Peso Integral": "peso_integral"
        }

        numeric_fields_default_zero = [
            "Valor Integral", "Valor Unitário 1", "Valor Unitário 2", "Valor Unitário 3",
            "Valor Unitário 4", "Valor Unitário 5", "Valor Unitário Energia",
            "Valor Mão de Obra TM/Metallica", "Peso Integral", "Custo Empresa",
            "Duplicata 1", "Duplicata 2", "Duplicata 3", "Duplicata 4", "Duplicata 5",
            "Duplicata 6", "Quantidade Usada"
        ]

        set_clause = []
        valores = []

        for coluna_amigavel, coluna_db in mapeamento_colunas.items():
            if coluna_amigavel in dados_atualizados:
                valor = dados_atualizados[coluna_amigavel]
                if valor == "" and coluna_amigavel in numeric_fields_default_zero:
                    valor = "0"
                if valor == "" and coluna_amigavel not in numeric_fields_default_zero:
                    valor = None

                if coluna_db == "data":
                    try:
                        valor = datetime.strptime(valor, "%d/%m/%Y").date()
                    except ValueError:
                        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/YYYY.")
                        return
                elif "valor" in coluna_db or "custo" in coluna_db or "duplicata" in coluna_db or coluna_db in ["peso_liquido", "peso_integral", "quantidade_usada"]:
                    if valor is not None:
                        try:
                            valor = float(valor.replace(".", "").replace(",", "."))
                        except ValueError:
                            messagebox.showerror("Erro", f"Valor inválido para o campo '{coluna_amigavel}'")
                            return
                elif coluna_db == "ipi":
                    if valor:
                        # Remove o símbolo de porcentagem e substitui a vírgula pelo ponto
                        valor = valor.replace("%", "").replace(",", ".")
                        try:
                            valor = float(valor)
                        except ValueError:
                            messagebox.showerror("Erro", f"Valor inválido para o campo '{coluna_amigavel}'")
                            return

                set_clause.append(f'"{coluna_db}" = %s')
                valores.append(valor)

        if not set_clause:
            messagebox.showerror("Erro", "Nenhuma coluna foi alterada.")
            return

        valores.append(nf_original)

        try:
            query = f"UPDATE somar_produtos SET {', '.join(set_clause)} WHERE \"nf\" = %s"
            self.cursor.execute(query, valores)
            self.conn.commit()

            self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            self.conn.commit()

            # Atualiza o Treeview com os novos valores (conforme sua lógica)
            self.tree.item(item_id, values=list(dados_atualizados.values()))

            messagebox.showinfo("Sucesso", "Linha editada com sucesso!")
            self.carregar_dados()

            if self.janela_menu and hasattr(self.janela_menu, 'frame_nf'):
            # limpa widgets antigos
                for widget in self.janela_menu.frame_nf.winfo_children():
                    widget.destroy()
            # repopula com dados atualizados
            self.janela_menu.criar_relatorio_nf()

            janela_edit.destroy()

        except Exception as e:
            self.conn.rollback()
            messagebox.showerror("Erro", f"Erro ao editar dados: {e}")

    def remove_selected_items(self):
        """Remove as linhas selecionadas do Treeview e das tabelas no banco de dados."""
        item_selecionado = self.tree.selection()
        if not item_selecionado:
            messagebox.showwarning("Seleção", "Selecione uma linha para excluir.")
            return

        # Obter os valores das linhas selecionadas para mostrar na mensagem
        nfs_para_excluir = []
        for item_id in item_selecionado:
            valores = self.tree.item(item_id, 'values')
            nf_item = valores[1].strip()  # Acessar a coluna NF (segundo elemento)
            nfs_para_excluir.append(nf_item)

        # Mensagem de confirmação
        confirmacao_msg = f"Tem certeza que deseja excluir as linhas com NF: {', '.join(nfs_para_excluir)}?"
        if not messagebox.askyesno("Confirmação", confirmacao_msg):
            return

        try:
            # Processar a exclusão de cada item selecionado
            for item_id in item_selecionado:
                valores = self.tree.item(item_id, 'values')
                nf_item = valores[1].strip()  # Acessar a coluna NF (segundo elemento)
                print(f"Valor de NF para exclusão: {nf_item}")  # Para depuração

                # Buscar o ID na tabela somar_produtos
                self.cursor.execute("SELECT id FROM somar_produtos WHERE LOWER(TRIM(nf)) = LOWER(%s) LIMIT 1", (nf_item,))
                id_item = self.cursor.fetchone()  # Obtém apenas um ID correspondente

                if id_item is not None:
                    # Excluir da tabela estoque_quantidade usando o ID de somar_produtos
                    self.cursor.execute("DELETE FROM estoque_quantidade WHERE id_produto = %s", (id_item[0],))
                    print(f"Registro relacionado na tabela estoque_quantidade com id_produto {id_item[0]} excluído com sucesso.")

                    # Excluir da tabela somar_produtos com base no ID
                    self.cursor.execute("DELETE FROM somar_produtos WHERE id = %s", (id_item[0],))
                    self.conn.commit()  # Comita as alterações após excluir
                    self.tree.delete(item_id)  # Remove do Treeview

                    # Adicionando um print para verificar se a exclusão foi bem-sucedida
                    print(f"Registro com NF {nf_item} e ID {id_item[0]} excluído com sucesso.")
                else:
                    messagebox.showerror("Erro", f"Registro com NF {nf_item} não encontrado.")

            # Reiniciar as sequências de IDs das tabelas (PostgreSQL)
            self.cursor.execute("SELECT SETVAL(pg_get_serial_sequence('somar_produtos', 'id'), COALESCE(MAX(id), 1)) FROM somar_produtos")
            self.cursor.execute("SELECT SETVAL(pg_get_serial_sequence('estoque_quantidade', 'id'), COALESCE(MAX(id), 1)) FROM estoque_quantidade")
            self.conn.commit()  # Comitar após reiniciar as sequências

            self.cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            self.conn.commit()

            print("IDs das tabelas reiniciados com sucesso.")

            # Atualizar o relatório do menu, se estiver presente
            if self.janela_menu and hasattr(self.janela_menu, 'frame_nf'):
                for widget in self.janela_menu.frame_nf.winfo_children():
                    widget.destroy()
                self.janela_menu.criar_relatorio_nf()

        except Exception as e:
            self.conn.rollback()  # Reverter as alterações em caso de erro
            messagebox.showerror("Erro", f"Erro ao excluir a linha: {e}")

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
                    self.after(0, self.destroy)
                except Exception:
                    try:
                        self.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()
        
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
