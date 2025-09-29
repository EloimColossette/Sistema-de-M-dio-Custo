# registro_teste.py
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont, filedialog
from conexao_db import conectar
from logos import aplicar_icone
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import re
from centralizacao_tela import centralizar_janela
import pandas as pd
import threading

class RegistroTeste(tk.Toplevel):
    def __init__(self, janela_menu=None, master=None):
        super().__init__(master=master)
        self.janela_menu = janela_menu
        self.title("Registro de Teste")
        self.config(bg="#f4f4f4")

        # Fonte local
        fixed_font = tkfont.Font(family="Arial", size=10)

        aplicar_icone(self, r"C:\Sistema\logos\Kametal.ico")

        # Tamanho
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{int(sw*0.8)}x{int(sh*0.8)}")
        self.state("zoomed")

        # Estilo local por instância (isolado) – removemos estilos de botão
        self._prefix = f"Reg{abs(id(self))}"
        self.style = ttk.Style(self)
        self.style.configure(f"{self._prefix}.TLabel", font=("Arial", 10))
        self.style.configure(f"{self._prefix}.TFrame", background=self.cget("bg"))
        self.style.configure(f"{self._prefix}.Treeview", rowheight=18)
        self.style.configure(f"{self._prefix}.Treeview.Heading", font=("Arial", 10, "bold"))

        # Conexão DB
        try:
            self.conn   = conectar()
            self.cursor = self.conn.cursor()
        except Exception as e:
            messagebox.showerror("Erro de Conexão", str(e))
            self.destroy()
            return
        
        # controle de atualização automática da área:
        self.area_manual = False      # True quando usuário editou manualmente a área
        self.last_auto_area = ""      # último valor auto-calculado (string com vírgula)

        # controle de atualização automática do MPa:
        self.mpa_manual = False
        self.last_auto_mpa = ""

        # Montagem da UI
        self.create_widgets(fixed_font=fixed_font)
        self.atualizar_treeview()
        self.configure_treeview()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self, fixed_font=None):
        if fixed_font is None:
            fixed_font = ("Arial", 10)

        # largura padrão para todos os campos
        ENTRY_WIDTH = 25  

        main = tk.Frame(self, bg="#f4f4f4")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # Barra de pesquisa
        search_frame = tk.Frame(main, bg="#f4f4f4")
        search_frame.pack(fill="x", pady=(0,10), padx=10)
        tk.Label(search_frame, text="Pesquisar:", bg="#f4f4f4", font=fixed_font).pack(side="left")

        self.search_entry = tk.Entry(search_frame, width=ENTRY_WIDTH, font=fixed_font)
        self.search_entry.pack(side="left", fill="x", expand=False, padx=(5,0))
        self.search_entry.bind("<KeyRelease>", lambda ev: self._filter_rows(ev.widget.get()))

        # Botão Exportar Excel
        tk.Button(search_frame, text="Exportar Excel", command=self.abrir_dialogo_exportacao, width=15)\
            .pack(side="right")
        
         # --- Botão de Ajuda discreto (❓) ---
        self.botao_ajuda = tk.Button(
            search_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=self._abrir_ajuda_registro_modal
        )
        self.botao_ajuda.pack(side=tk.RIGHT, padx=6, pady=10)

        # efeito hover
        self.botao_ajuda.bind("<Enter>", lambda e: self.botao_ajuda.config(bg="#3b5566"))
        self.botao_ajuda.bind("<Leave>", lambda e: self.botao_ajuda.config(bg="#2c3e50"))

        # Atalho F1 para abrir ajuda (quando esta janela estiver ativa)
        try:
            self.bind("<F1>", lambda e: self._abrir_ajuda_registro_modal())
        except Exception:
            pass

        # Tooltip (se desejar hover)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda, "Ajuda — Registro Teste (F1)")
        except Exception:
            pass

        # Formulário
        form = tk.LabelFrame(main, text="Dados do Registro", bg="#f4f4f4", padx=10, pady=10)
        form.pack(fill="x", padx=10, pady=10)

        labels = [
            "Data", "Código de Barras", "O.P.", "Cliente", "Material", "Liga", "Dimensões","Área", "L.R. Tração (N)", "L.R. Tração (MPa)", "Alongamento (%)", "Tempera","Máquina", "Empresa"
        ]

        self.entries = {}
        for i, lbl in enumerate(labels):
            row, col = divmod(i, 3)
            tk.Label(form, text=lbl + ":", bg="#f4f4f4", font=fixed_font)\
            .grid(row=row, column=col*2, sticky="e", padx=5, pady=5)

            e = tk.Entry(form, width=ENTRY_WIDTH, font=fixed_font)

            if lbl == "Data":
                self.date_entry = e
                e.bind("<KeyRelease>", self._on_date_key)
            elif lbl == "Alongamento (%)":
                self.along_entry = e
                e.bind("<KeyRelease>", self._on_along_key)
            elif lbl == "Dimensões":
                self.dim_entry = e
                # Só Dimensões precisa dessa formatação especial
                e.bind("<KeyRelease>", lambda ev, en=e: self._on_decimal_key(ev, en))
                e.bind("<KeyRelease>", self._on_dim_key, add="+")
                e.bind("<FocusOut>", self._on_dim_key, add="+")
            elif lbl == "Área":
                self.area_entry = e
                # NÃO aplicar _on_decimal_key aqui
                e.bind("<KeyRelease>", self._on_area_key, add="+")
                e.bind("<FocusOut>", self._on_area_key, add="+")      # garante recálculo/limpeza ao perder foco
                e.bind("<Double-Button-1>", self._reset_area_auto, add="+")
                # -> cor de fundo para destacar que é campo com cálculo automático
                try:
                    e.config(bg="#fff7cc")  # amarelo suave
                except Exception:
                    pass
                # chama o método da classe
                self.add_tooltip(e, "Área: calculada automaticamente — não precisa digitar.", delay=400)
            elif lbl == "L.R. Tração (N)":
                self.lr_n_entry = e
                # apenas cálculo automático do MPa
                e.bind("<KeyRelease>", self._on_lr_n_key)
                e.bind("<FocusOut>", self._on_lr_n_key)
            elif lbl == "L.R. Tração (MPa)":
                self.lr_mpa_entry = e
                # sem formatação automática, só controle manual/auto
                e.bind("<KeyRelease>", self._on_mpa_key, add="+")
                e.bind("<Double-Button-1>", self._reset_mpa_auto, add="+")
                # cor de fundo para destacar
                try:
                    e.config(bg="#fff7cc")
                except Exception:
                    pass
                self.add_tooltip(e, "L.R. Tração (MPa): calculado automaticamente (N ÷ Área). Só digite se precisar.", delay=400)
            elif lbl == "Tempera":
                self.temper_entry = e
                e.bind("<KeyRelease>", self._on_tempera_key)

            e.grid(row=row, column=col*2+1, sticky="w", padx=5, pady=5)
            self.entries[lbl] = e

        # Botões
        btn_frame = tk.Frame(main, bg="#f4f4f4")
        btn_frame.pack(pady=5)
        for text, cmd in [
            ("Salvar", self.salvar), ("Alterar", self.alterar),
            ("Excluir", self.excluir), ("Limpar", self.limpar),
            ("Voltar", self.voltar_para_menu)
        ]:
            tk.Button(btn_frame, text=text, command=cmd, width=15)\
            .pack(side="left", padx=3)

        # Treeview
        tree_frame = tk.Frame(main, bg="#f4f4f4")
        tree_frame.pack(fill="both", expand=True, pady=(10,0))
        vsb = tk.Scrollbar(tree_frame, orient="vertical")
        hsb = tk.Scrollbar(tree_frame, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        cols = labels[:]
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                                style=f"{self._prefix}.Treeview")
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.tree.pack(side="left", fill="both", expand=True)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, anchor="center")
        self.tree.column("Data", width=80)
        self.tree.column("Área", width=100)
        self.tree.column("Alongamento (%)", width=150)

        self.tree.bind("<<TreeviewSelect>>", self.carregar_dados)

    def atualizar_treeview(self):
        """Carrega todos os registros sem filtro e exibe tudo."""
        try:
            self.cursor.execute(
                """
                SELECT *
                FROM registro_teste
                ORDER BY
                data DESC,
                LEFT(
                    regexp_replace(unaccent(lower(cliente)), '[^a-z]', '', 'g'),
                    1
                ) ASC,
                codigo_barras ASC;
                """
            )
            self.all_rows = self.cursor.fetchall()
            self._filter_rows("")  # exibe tudo novamente
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar: {e}")

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

    def _abrir_ajuda_registro_modal(self, contexto=None):
        """Modal de Ajuda para Registro de Teste — versão estilizada (layout com navegação + painéis)."""
        try:
            modal = tk.Toplevel(self)
            modal.title("Ajuda — Registro de Teste")
            modal.transient(self)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões / centralização
            w, h = 920, 680
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # Ícone (silencioso se falhar)
            try:
                # tenta usar caminho de ícone da instância, se existir; senão ignora
                ic = getattr(self, "caminho_icone", r"C:\Sistema\logos\Kametal.ico")
                aplicar_icone(modal, ic)
            except Exception:
                pass

            # Cabeçalho do modal
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Registro de Teste", bg="#2b3e50", fg="white",
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
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold"))\
                .pack(anchor="nw", pady=(10, 6), padx=12)

            sections = [
                "Visão Geral",
                "Salvar / Editar / Excluir / Limpar",
                "Entradas e Formatação",
                "Área & L.R. Tração (Cálculos automáticos)",
                "Alongamento (%)",
                "Tempera (normalização)",
                "Pesquisa",
                "Exportação (Excel)",
                "FAQ"
            ]

            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            # --- Content area (usa grid + tkraise para estabilidade) ---
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

            # Painel dedicado para Entradas (com scrollbar própria)
            entradas_frame = tk.Frame(content_frame, bg="white")
            entradas_frame.grid(row=0, column=0, sticky="nsew")
            entradas_frame.rowconfigure(0, weight=1)
            entradas_frame.columnconfigure(0, weight=1)

            txt_entradas = tk.Text(entradas_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_entradas = tk.Scrollbar(entradas_frame, command=txt_entradas.yview)
            txt_entradas.configure(yscrollcommand=sb_entradas.set)
            txt_entradas.grid(row=0, column=0, sticky="nsew")
            sb_entradas.grid(row=0, column=1, sticky="ns")

            # Conteúdos
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "Tela Registro de Teste — registre ensaios com Data, Dimensões, Área, L.R. Tração (N/MPa), Alongamento, Tempera, Máquina e Empresa.\n"
                "Use os botões para Salvar, Alterar, Excluir e Limpar. A lista (Treeview) mostra os registros; selecione uma linha para carregar no formulário."
            )

            contents["Salvar / Editar / Excluir / Limpar"] = (
                "Salvar / Editar / Excluir / Limpar\n\n"
                "- Salvar: preencha os campos (Data, Dimensões, L.R. Tração (N), etc.) e clique em 'Salvar'.\n"
                "- Editar (Alterar): selecione um registro na tabela; os campos serão preenchidos. Altere o que for necessário e clique em 'Alterar' para persistir.\n"
                "- Excluir: selecione um ou mais registros e clique em 'Excluir'. Confirme para remover do banco.\n"
                "- Limpar: limpa todos os campos do formulário (não apaga registros salvos no banco).\n\n"
                "Observação: operações críticas (Excluir/Alterar) devem pedir confirmação — mantenha sempre backup antes de operações em lote."
            )

            contents["Entradas e Formatação"] = (
                "Entradas e Formatação — regras e exemplos\n\n"
                "- Data: não é necessário digitar barras. Digite apenas dígitos: ex.: '01012025' → será formatado como '01/01/2025'.\n\n"
                "- Dimensões: digite apenas dígitos (sem vírgula). A formatação interna converte para valor com casas decimais quando necessário.\n"
                "  Exemplo: se o campo espera 2 casas decimais, digitar '12345' pode resultar em '123,45'.\n\n"
                "- Valores numéricos (Área, L.R. Tração, etc.): NÃO digite vírgula manualmente — digite só dígitos; o sistema aplica a vírgula/decimais automaticamente.\n"
                "  Você também pode colar valores já formatados; verifique o campo antes de salvar.\n\n"
                "- Campos em AMARELO (Área e L.R. Tração (MPa)) indicam campos calculados automaticamente (veja seção específica)."
            )

            contents["Área & L.R. Tração (Cálculos automáticos)"] = (
                "Área e L.R. Tração — cálculo automático e comportamento\n\n"
                "- Cálculo de Área:\n"
                "  • A área é calculada automaticamente a partir de Dimensões quando possível. Fórmula utilizada (exemplo para seção circular):\n"
                "    area = dimensao * dimensao * 0.7854  (π/4 ≈ 0.7854)\n"
                "  • O campo 'Área' é exibido em AMARELO para indicar que é um campo derivado.\n\n"
                "- Cálculo de L.R. Tração (MPa):\n"
                "  • L.R. Tração (MPa) = L.R. Tração (N) ÷ Área\n"
                "  • O campo 'L.R. Tração (MPa)' também aparece em AMARELO e é recalculado sempre que Dimensões ou L.R. Tração (N) mudam.\n\n"
                "- Comportamento e edição manual:\n"
                "  • Você pode apagar o valor automático, digitar um número manualmente ou deixar em branco — o sistema não irá quebrar.\n"
                "  • Depois de editar manualmente, dupla-clique no campo (Área ou L.R. Tração (MPa)) para restaurar o cálculo automático.\n"
                "  • Por isso os campos aparecem em amarelo: alertar o usuário que são calculados, mas editáveis."
            )

            contents["Alongamento (%)"] = (
                "Alongamento (%)\n\n"
                "- Não digite o símbolo '%' na entrada de Alongamento. Digite apenas o número (ex.: '12'). O sistema apresentará '12%' automaticamente.\n"
                "- O campo aceita valores vazios e entradas manuais; formatará o valor ao perder o foco."
            )

            contents["Tempera (normalização)"] = (
                "Tempera — normalização de texto\n\n"
                "- Ao digitar no campo 'Tempera', o sistema aplica normalizações automáticas para palavras conhecidas.\n"
                "- Algumas palavras/abreviações são expandidas e a palavra 'Duro' é adicionada automaticamente quando apropriado (ex.: ao digitar certas siglas ou palavras-chave o sistema pode inserir ' Duro' no final).\n"
                "- Se quiser evitar a adição automática, digite o texto completo desejado explicitamente (o sistema tentará normalizar, mas o texto manual prevalece)."
            )

            contents["Pesquisa"] = (
                "Pesquisa — como usar\n\n"
                "- O campo 'Pesquisar' filtra a lista (Treeview) ao digitar: você pode pesquisar por Data, Código de Barras, O.P., Cliente, Material, Máquina, Empresa, etc.\n"
                "- Para busca por data exata, digite com '/' (ex.: '01/01/2025') ou digite apenas números para busca por substring ('01012025').\n"
                "- A pesquisa é incremental; apague o texto para voltar à listagem completa."
            )

            contents["Exportação (Excel)"] = (
                "Exportação (Excel)\n\n"
                "- Clique em 'Exportar Excel' para abrir o diálogo de exportação.\n"
                "- Você pode escolher filtros: Data inicial / Data final e filtros textuais (O.P., Cliente, Material, Máquina, Empresa).\n"
                "- O arquivo .xlsx gerado conterá colunas com Data, Dimensões, Área, L.R. Tração (N), L.R. Tração (MPa), Alongamento (%), Tempera, Máquina e Empresa.\n"
                "- Antes de exportar, confirme que os campos calculados estão como deseja (área/l.r. tração), pois os valores exportados são os mostrados na interface."
            )

            contents["FAQ"] = (
                "FAQ — Perguntas rápidas\n\n"
                "Q: Posso deixar campos calculados em branco?\n"
                "A: Sim — o sistema aceita; se os dados de entrada estiverem disponíveis (ex.: Dimensões) o cálculo automático é refeito.\n\n"
                "Q: Posso editar manualmente valores calculados?\n"
                "A: Sim — campos em amarelo são editáveis. Duplo-clique para restaurar o cálculo automático.\n\n"
                "Q: Tempera sempre terá 'Duro' adicionado?\n"
                "A: O sistema adiciona 'Duro' em casos onde a normalização do texto determinar que faz sentido. Você pode substituir manualmente o texto."
            )

            # Função que mostra a seção correta (usa tkraise para alternar painéis)
            def mostrar_secao(key):
                if key == "Entradas e Formatação":
                    txt_entradas.configure(state="normal")
                    txt_entradas.delete("1.0", "end")
                    txt_entradas.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_entradas.configure(state="disabled")
                    txt_entradas.yview_moveto(0)
                    entradas_frame.tkraise()
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

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])

            listbox.bind("<<ListboxSelect>>", on_select)

            # Rodapé e Fechar (botão estiloso)
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
            import traceback
            traceback.print_exc()
            print("Erro ao abrir modal de ajuda (Registro de Teste):", e)


    def abrir_dialogo_exportacao(self):
        dialogo = tk.Toplevel(self)
        dialogo.title("Exportar Registro - Filtros")
        dialogo.geometry("400x400")  # altura aumentada
        dialogo.resizable(False, False)
        centralizar_janela(dialogo, 400, 400)
        aplicar_icone(dialogo, r"C:\Sistema\logos\Kametal.ico")
        dialogo.config(bg="#ecf0f1")

        # Configuração dos estilos personalizados utilizando ttk
        style = ttk.Style(dialogo)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white", font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        frame = ttk.Frame(dialogo, padding="15", style="Custom.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        dialogo.columnconfigure(0, weight=1)
        dialogo.rowconfigure(0, weight=1)

        # Função de formatação de data
        def formatar_data(event):
            entry = event.widget
            digits = ''.join(ch for ch in entry.get() if ch.isdigit())[:8]

            parts = []
            if len(digits) >= 2:
                parts.append(digits[:2])
                if len(digits) >= 4:
                    parts.append(digits[2:4])
                    if len(digits) > 4:
                        parts.append(digits[4:])
                else:
                    parts.append(digits[2:])
            else:
                parts.append(digits)

            novo = '/'.join(parts)
            entry.delete(0, tk.END)
            entry.insert(0, novo)

        # Campos de filtro correspondentes a registro_teste
        filtros = [
            "Data Inicial", 
            "Data Final", 
            "Código de Barras", 
            "O.P.",
            "Cliente", 
            "Material", 
            "Máquina", 
            "Empresa"
        ]
        entries = {}
        binds = {"Data Inicial": formatar_data, "Data Final": formatar_data}
        for i, lbl in enumerate(filtros):
            ttk.Label(frame, text=f"{lbl}:", style="Custom.TLabel").grid(
                row=i, column=0, sticky="e", padx=5, pady=4)
            e = ttk.Entry(frame, width=20)
            e.grid(row=i, column=1, sticky="w", padx=5, pady=4)
            if lbl in binds:
                e.bind("<KeyRelease>", binds[lbl])
            entries[lbl] = e

        def acao_exportar():
            vals = {lbl: entries[lbl].get().strip() for lbl in filtros}
            where, params = [], []
            # Data
            if vals["Data Inicial"]:
                try:
                    di = datetime.strptime(vals["Data Inicial"], '%d/%m/%Y')
                    where.append("data >= %s"); params.append(di.strftime('%Y-%m-%d'))
                except: pass
            if vals["Data Final"]:
                try:
                    dfm = datetime.strptime(vals["Data Final"], '%d/%m/%Y')
                    where.append("data <= %s"); params.append(dfm.strftime('%Y-%m-%d'))
                except: pass
            # Outros filtros textuais
            for col, field in [
                ("codigo_barras","Código de Barras"),
                ("op","O.P."),
                ("cliente","Cliente"),
                ("material","Material"),
                ("maquina","Máquina"),
                ("empresa","Empresa")
            ]:
                if vals[field]:
                    where.append(f"{col} ILIKE %s"); params.append(f"%{vals[field]}%")
            sql = (
                "SELECT data, codigo_barras, op, cliente, material, liga, "
                "dimensoes, area, lr_tracao_n, lr_tracao_mpa, alongamento_percentual, tempera, maquina, empresa FROM registro_teste"
            )
            if where:
                sql += " WHERE " + " AND ".join(where)
            sql += " ORDER BY data ASC"

            arquivo = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel','*.xlsx')],
                initialfile='RegistroTeste.xlsx'
            )
            if arquivo:
                try:
                    self.cursor.execute(sql, params)
                    dados = self.cursor.fetchall()
                    cols = [
                        "Data", "Código de Barras", "O.P.", "Cliente", "Material",
                        "Liga", "Dimensões", "Área", "L.R. Tração (N)",
                        "L.R. Tração (MPa)", "Alongamento (%)", "Tempera", "Máquina", "Empresa"
                    ]
                    df = pd.DataFrame(dados, columns=cols)
                    df['Data'] = df['Data'].apply(lambda d: d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d))
                    df.to_excel(arquivo, index=False)
                    messagebox.showinfo("Exportação", f"Arquivo salvo em {arquivo}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível exportar: {e}")
            dialogo.destroy()

        # Botões
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.grid(row=len(filtros), column=0, columnspan=2, pady=10)
        ttk.Button(
            btn_frame, text="Exportar Excel",
            command=acao_exportar, style="Custom.TButton"
        ).grid(row=0, column=0, padx=5)
        ttk.Button(
            btn_frame, text="Cancelar",
            command=dialogo.destroy, style="Custom.TButton"
        ).grid(row=0, column=1, padx=5)
        frame.columnconfigure(1, weight=1)

    def _filter_rows(self, term):
        """
        term: string com o que foi digitado no Entry.
        Se term == "", exibe tudo.
        """
        term = term.lower().strip()
        self.tree.delete(*self.tree.get_children())

        for row in getattr(self, 'all_rows', []):
            rec_id, raw = row[0], row[1]
            data = raw.strftime("%d/%m/%Y") if hasattr(raw, "strftime") else str(raw or "")

            # Lê os ranges de tração e substitui vazios por "0,0 - 0,0"
            lr_n_val = row[9]
            lr_mpa_val = row[10]
            lr_n = lr_n_val if (lr_n_val is not None and str(lr_n_val).strip() != "") else "0,0 - 0,0"
            lr_mpa = lr_mpa_val if (lr_mpa_val is not None and str(lr_mpa_val).strip() != "") else "0,0 - 0,0"

            vals = [
                data,
                row[2], row[3], row[4], row[5],
                row[6], row[7],
                f"{row[8]:.4f}".replace(".", ",") if row[8] is not None else "",
                lr_n, lr_mpa,
                f"{row[11]:.0f}%" if row[11] is not None else "",
                row[12], row[13], row[14]
            ]
            if not term or any(term in str(v).lower() for v in vals):
                self.tree.insert("", "end", iid=str(rec_id), values=vals)

    def configure_treeview(self):
        style = ttk.Style(self)
        style.theme_use("alt")   # mantém seu tema
        style.configure(f"{self._prefix}.Treeview", rowheight=20)
        style.configure(f"{self._prefix}.Treeview.Heading", font=("Arial", 10, "bold"))
        sel_bg = '#2980B9'   # azul da sua imagem
        sel_fg = 'white'     # texto branco como na imagem

        style.map(
            f"{self._prefix}.Treeview",
            background=[('selected', sel_bg)],
            foreground=[('selected', sel_fg)]
        )
        self.tree.config(height=10)

    def carregar_dados(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]

        for lbl, v in zip(self.entries.keys(), vals):
            e = self.entries[lbl]
            e.delete(0, tk.END)
            if lbl == "Data":
                # já vem formatado "DD/MM/YYYY"
                e.insert(0, v)
            elif lbl == "Alongamento (%)":
                # já vem formatado "60,00%"
                e.insert(0, v)
            else:
                e.insert(0, v)

    def _on_date_key(self, event, entry=None):
        e = entry or self.date_entry
        digits = re.sub(r'\D', '', e.get())[:8]

        parts = []
        if len(digits) >= 2:
            parts.append(digits[:2])
            if len(digits) >= 4:
                parts.append(digits[2:4])
                if len(digits) > 4:
                    parts.append(digits[4:])
            else:
                parts.append(digits[2:])
        else:
            parts.append(digits)

        novo = '/'.join(parts)
        e.delete(0, tk.END)
        e.insert(0, novo)

    def _on_along_key(self, event):
        e = self.along_entry
        t = e.get()
        core = ''.join(c for c in t if c.isdigit() or c in ",.") 
        novo = (core + "%") if core else ""
        pos = e.index(tk.INSERT)
        e.delete(0, tk.END)
        e.insert(0, novo)
        e.icursor(min(pos, len(novo)))

    def _on_decimal_key(self, event, entry):
        """
        Formata automaticamente para que antes das cinco últimas 
        casas haja uma vírgula. Ex.: "123456789" -> "1234,56789".
        """
        t = entry.get()
        # pega só dígitos
        digits = re.sub(r'\D', '', t)

        if not digits:
            novo = ""
        elif len(digits) <= 3:
            # até cinco dígitos, mostra-os puros (sem vírgula)
            novo = digits
        else:
            # insere vírgula antes dos 5 últimos dígitos
            inteiro = digits[:-3]
            dec = digits[-3:]
            novo = f"{inteiro},{dec}"

        # posiciona o cursor próximo da posição original
        pos = entry.index(tk.INSERT)
        entry.delete(0, tk.END)
        entry.insert(0, novo)
        entry.icursor(min(pos, len(novo)))

    def _on_tempera_key(self, event):
        # cancela qualquer formatação pendente
        if hasattr(self, "_after_id"):
            self.after_cancel(self._after_id)
        # agenda a formatação para daqui a 600ms
        self._after_id = self.after(600, self._format_tempera)

    def _format_tempera(self):
        e = self.temper_entry
        text = e.get().strip()

        if not text:
            novo = ""
        else:
            low = text.lower()

            # dicionário de abreviações para nomes completos
            abrevs = {
                "rec": "Recozido",
                "recoz": "Recozido",
                "recozido": "Recozido",
                "mol": "Mola",
                "mola": "Mola",
                "du": "Duro",
                "duro": "Duro"
            }

            # verifica se a entrada corresponde a algum valor no dicionário
            completo = abrevs.get(low)

            if completo:
                # se for um caso conhecido, mantém apenas o nome completo
                novo = completo
            else:
                # remove qualquer "duro" isolado do texto e normaliza espaços
                core = re.sub(r'\bduro\b', '', text, flags=re.IGNORECASE).strip()
                core = re.sub(r'\s+', ' ', core)
                # adiciona "Duro" ao final se houver texto, senão só "Duro"
                novo = f"{core} Duro" if core else "Duro"

        # atualiza o Entry preservando a posição do cursor
        try:
            pos = e.index(tk.INSERT)
        except Exception:
            pos = None
        e.delete(0, tk.END)
        e.insert(0, novo)
        if pos is not None:
            e.icursor(min(pos, len(novo)))

    def _on_dim_key(self, event=None):
        """
        Recalcula a área a partir do valor das dimensões:
        area = dimensoes * dimensoes * 0.7854
        Insere no campo Área somente se o usuário não tiver
        feito uma edição manual (self.area_manual == False)
        ou caso o campo Área contenha exatamente o último valor auto.
        Se Dimensões for apagado, limpa também Área (e MPa).
        """
        if event and getattr(event, "keysym", None) == "Tab":
            return
        txt = (self.dim_entry.get().strip() if hasattr(self, "dim_entry") else "")

        # Se Dimensões foi limpada -> limpar Área e MPa
        if txt == "":
            try:
                self.area_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_area = ""
            self.area_manual = False

            # Limpa MPa também (mesmo que estivesse em modo manual)
            try:
                self.lr_mpa_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_mpa = ""
            self.mpa_manual = False
            return

        norm = self.normalize_number(txt)  # retorna string com ponto decimal ou None
        if not norm:
            return

        try:
            val = Decimal(norm)
        except Exception:
            return

        try:
            area_dec = (val * val * Decimal("0.7854")).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
        except Exception:
            return

        # formata com vírgula como separador decimal (4 casas)
        area_str = format(area_dec, 'f').replace('.', ',')

        current_area = (self.area_entry.get().strip() if hasattr(self, "area_entry") else "")
        # atualiza se usuário não editou manualmente, ou se o campo ainda contém o último valor automático
        if (not self.area_manual) or (current_area == self.last_auto_area) or (current_area == ""):
            self.area_entry.delete(0, tk.END)
            self.area_entry.insert(0, area_str)
            self.last_auto_area = area_str
            self.area_manual = False

        # solicita recalcular MPa depois do loop de eventos, garantindo que o Entry já
        # contenha o novo valor (protege contra leituras prematuras).
        try:
            self.after(1, lambda: self._on_lr_n_key(None))
        except Exception:
            pass

    def _on_area_key(self, event=None):
        """
        Chamado quando o usuário digita no campo 'Área' ou perde o foco — marca que o valor
        foi editado manualmente para evitar sobrescrita automática e solicita recálculo do MPa.
        Se o campo Área for limpo, também limpamos o MPa (mesmo que MPa estivesse em modo manual).
        """
        self.area_manual = True

        # Agendamos a ação para depois do loop de eventos: se área vazia -> limpar MPa,
        # caso contrário -> recalcular MPa (se possível).
        try:
            self.after(1, self._handle_area_change)
        except Exception:
            try:
                self._handle_area_change()
            except Exception:
                pass

    def _handle_area_change(self):
        """Helper chamado logo após alteração na área (garante conteúdo atualizado)."""
        area_txt = (self.area_entry.get().strip() if hasattr(self, "area_entry") else "")
        if area_txt == "":
            # limpa MPa sempre que área for apagada
            try:
                self.lr_mpa_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_mpa = ""
            self.mpa_manual = False
        else:
            # tenta recalcular o MPa (se N estiver presente)
            try:
                self._on_lr_n_key(None)
            except Exception:
                pass

    def _reset_area_auto(self, event=None):
        """
        Ao dar duplo-clique no campo Área, remove a marca de edição manual
        e recalcula imediatamente (se houver valor em Dimensões).
        """
        self.area_manual = False
        # chama o recalculo (event pode ser None)
        self._on_dim_key(event)

    def _on_lr_n_key(self, event=None):
        """
        Recalcula o MPa sempre que o usuário digitar em L.R. Tração (N) ou quando a
        área for atualizada (manual/automática).
        Fórmula: MPa = N / Área

        Se N ou Área estiverem vazios/invalidos, limpa o campo MPa (como solicitado).
        """
        if event and getattr(event, "keysym", None) == "Tab":
            return
        txt_n = (self.lr_n_entry.get().strip() if hasattr(self, "lr_n_entry") else "")
        txt_area = (self.area_entry.get().strip() if hasattr(self, "area_entry") else "")

        norm_n = self.normalize_number(txt_n)
        norm_area = self.normalize_number(txt_area)

        # Se algum dos dois estiver vazio/invalid -> limpar MPa
        if not norm_n or not norm_area:
            try:
                self.lr_mpa_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_mpa = ""
            self.mpa_manual = False
            return

        try:
            val_n = Decimal(norm_n)
            val_area = Decimal(norm_area)
            if val_area == 0:
                # evita divisão por zero: limpamos MPa
                try:
                    self.lr_mpa_entry.delete(0, tk.END)
                except Exception:
                    pass
                self.last_auto_mpa = ""
                self.mpa_manual = False
                return
        except Exception:
            try:
                self.lr_mpa_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_mpa = ""
            self.mpa_manual = False
            return

        try:
            mpa_dec = (val_n / val_area).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except Exception:
            try:
                self.lr_mpa_entry.delete(0, tk.END)
            except Exception:
                pass
            self.last_auto_mpa = ""
            self.mpa_manual = False
            return

        # formata com vírgula
        mpa_str = format(mpa_dec, 'f').replace('.', ',')

        current_mpa = (self.lr_mpa_entry.get().strip() if hasattr(self, "lr_mpa_entry") else "")

        if (not self.mpa_manual) or (current_mpa == self.last_auto_mpa) or (current_mpa == ""):
            self.lr_mpa_entry.delete(0, tk.END)
            self.lr_mpa_entry.insert(0, mpa_str)
            self.last_auto_mpa = mpa_str
            self.mpa_manual = False

    def _on_mpa_key(self, event=None):
        """Marca que o usuário digitou manualmente o MPa"""
        self.mpa_manual = True

    def _reset_mpa_auto(self, event=None):
        """Duplo clique → volta para modo automático"""
        self.mpa_manual = False
        self._on_lr_n_key(event)

    def normalize_number(self, s):
        s = s.strip()
        if not s:
            return None
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return s

    def add_tooltip(self, widget, text, delay=400):
        """Método de tooltip dentro da classe. Use: self.add_tooltip(widget, texto)."""
        def _unschedule():
            tt_id = getattr(widget, "_tt_id", None)
            if tt_id:
                try:
                    widget.after_cancel(tt_id)
                except Exception:
                    pass
                widget._tt_id = None

        def _show():
            if getattr(widget, "_tt_win", None) or not text:
                return
            try:
                x = widget.winfo_rootx() + 20
                y = widget.winfo_rooty() + widget.winfo_height() + 10
            except Exception:
                x, y = 100, 100
            tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            lbl = tk.Label(tw, text=text, justify="left",
                           background="#ffffe0", relief="solid", borderwidth=1,
                           font=("Arial", 8))
            lbl.pack(ipadx=4, ipady=2)
            widget._tt_win = tw

        def _schedule(event=None):
            _unschedule()
            try:
                widget._tt_id = widget.after(delay, _show)
            except Exception:
                widget._tt_id = None

        def _hide(event=None):
            _unschedule()
            tw = getattr(widget, "_tt_win", None)
            widget._tt_win = None
            if tw:
                try:
                    tw.destroy()
                except Exception:
                    pass

        widget.bind("<Enter>", _schedule)
        widget.bind("<Leave>", _hide)
        widget.bind("<ButtonPress>", _hide)

        def remove():
            try:
                widget.unbind("<Enter>")
                widget.unbind("<Leave>")
                widget.unbind("<ButtonPress>")
            except Exception:
                pass
            _hide()
            for attr in ("_tt_id", "_tt_win"):
                if hasattr(widget, attr):
                    try:
                        delattr(widget, attr)
                    except Exception:
                        pass
        return remove

    def salvar(self):
        # helper para pegar valor de cada Entry
        get = lambda lbl: self.entries[lbl].get().strip()

        # pega a string da data, mantendo as barras
        data_str    = get("Data")                # ex.: "23/07/2025"
        cod         = get("Código de Barras")
        op          = get("O.P.")
        cliente     = get("Cliente")
        material    = get("Material")
        liga        = get("Liga")
        dims        = get("Dimensões")
        area_s      = get("Área").replace(".", "").replace(",", ".")
        lr_n_s      = get("L.R. Tração (N)")
        lr_mpa_s    = get("L.R. Tração (MPa)")
        along_str   = get("Alongamento (%)").replace("%", "")
        tempera     = get("Tempera")
        maquina     = get("Máquina")
        empresa     = get("Empresa")

        # Valida data
        if not data_str:
            return messagebox.showerror("Erro", "O campo Data é obrigatório.")
        try:
            data_date = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return messagebox.showerror("Erro", "Data inválida! Use DD/MM/AAAA.")

        # Converte apenas área e alongamento (numéricos)
        def to_float(s, nome):
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                messagebox.showerror("Erro", f"{nome} deve ser número.")
                raise

        try:
            area  = to_float(area_s,  "Área")
            along = to_float(along_str, "Alongamento (%)")
        except:
            return
        
         # pega menor ID disponível
        try:
            self.cursor.execute(
                """
                SELECT COALESCE(
                    (SELECT MIN(t1.id) + 1
                     FROM registro_teste t1
                     LEFT JOIN registro_teste t2 ON t2.id = t1.id + 1
                     WHERE t2.id IS NULL), 1
                );
                """
            )
            next_id = self.cursor.fetchone()[0]
        except Exception as e:
            return messagebox.showerror("Erro ao gerar ID", str(e))

        # Monta parâmetros **mantendo lr_n_s e lr_mpa_s como texto**
        params = (
            next_id,
            cod, data_date, op, cliente, material, liga, dims,
            area, lr_n_s, lr_mpa_s, along, tempera, maquina, empresa
        )
        sql = """
            INSERT INTO registro_teste
            (id, codigo_barras, data, op, cliente, material, liga, dimensoes,
            area, lr_tracao_n, lr_tracao_mpa, alongamento_percentual,
            tempera, maquina, empresa)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
        """

        try:
            self.cursor.execute(sql, params)
            self.conn.commit()
            self.cursor.execute(
                "NOTIFY canal_atualizacao, 'menu_atualizado';"
            )
            # como já fizemos commit acima, basta outro commit para o NOTIFY
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso!")
            self.limpar() 
            self.atualizar_treeview()
        except Exception as e:
            messagebox.showerror("Falha ao salvar", str(e))

    def alterar(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("Atenção", "Selecione um registro para alterar.")
        rec_id = int(sel[0])

        get = lambda lbl: self.entries[lbl].get().strip()
        data_str = get("Data")
        cod       = get("Código de Barras")
        op        = get("O.P.")
        cliente   = get("Cliente")
        material  = get("Material")
        liga      = get("Liga")
        dims      = get("Dimensões")

        # área e alongamento como string “1000,12345”
        area_s = self.normalize_number(get("Área"))

        # retira o '%' antes de normalizar
        raw_along = get("Alongamento (%)").replace("%", "")
        along_s   = self.normalize_number(raw_along)

        # ranges de tração como texto puro
        lr_n_s   = get("L.R. Tração (N)")
        lr_mpa_s = get("L.R. Tração (MPa)")

        tempera = get("Tempera")
        maquina = get("Máquina")
        empresa = get("Empresa")

        # valida data
        if not data_str:
            return messagebox.showerror("Erro", "O campo Data é obrigatório.")
        try:
            data_date = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return messagebox.showerror("Erro", "Data inválida! Use DD/MM/AAAA.")

        # converte e quantiza area e alongamento
        try:
            area_dec = Decimal(area_s).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
            if abs(area_dec) >= Decimal("1E8"):
                raise InvalidOperation()
        except InvalidOperation:
            return messagebox.showerror(
                "Erro",
                "Área inválida: máximo ±10⁸ e até 4 casas decimais."
            )

        try:
            along_dec = Decimal(along_s).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            if abs(along_dec) >= Decimal("1E4"):
                raise InvalidOperation()
        except InvalidOperation:
            return messagebox.showerror(
                "Erro",
                "Alongamento inválido: máximo ±10⁴ e até 2 casas decimais."
            )

        # agora monta o UPDATE com o nome correto da coluna de data
        sql = """
            UPDATE registro_teste
            SET codigo_barras          = %s,
                data                    = %s,
                op                     = %s,
                cliente                = %s,
                material               = %s,
                liga                   = %s,
                dimensoes              = %s,
                area                   = %s,
                lr_tracao_n            = %s,
                lr_tracao_mpa          = %s,
                alongamento_percentual = %s,
                tempera                = %s,
                maquina                = %s,
                empresa                = %s
            WHERE id = %s;
        """
        params = (
            cod, data_date, op, cliente, material, liga, dims,
            area_dec, lr_n_s, lr_mpa_s, along_dec,
            tempera, maquina, empresa,
            rec_id
        )

        try:
            self.cursor.execute(sql, params)
            self.conn.commit()
            self.cursor.execute(
                "NOTIFY canal_atualizacao, 'menu_atualizado';"
            )
            # como já fizemos commit acima, basta outro commit para o NOTIFY
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Registro alterado com sucesso!")
            self.atualizar_treeview()
            self.limpar()
        except Exception as e:
            messagebox.showerror("Erro ao alterar", str(e))

    def excluir(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("Atenção", "Selecione ao menos um registro para excluir.")

        ids = [int(item) for item in sel]
        count = len(ids)
        pergunta = f"Deseja realmente excluir {count} registro(s)?"
        if not messagebox.askyesno("Confirmação", pergunta):
            return

        try:
            # exclui múltiplos de uma vez
            placeholders = ",".join(["%s"] * count)
            sql = f"DELETE FROM registro_teste WHERE id IN ({placeholders});"
            self.cursor.execute(sql, tuple(ids))
            self.conn.commit()
            self.cursor.execute(
                "NOTIFY canal_atualizacao, 'menu_atualizado';"
            )
            # como já fizemos commit acima, basta outro commit para o NOTIFY
            self.conn.commit()

            self.atualizar_treeview()
            self.limpar()
            messagebox.showinfo("Sucesso", f"{count} registro(s) excluído(s) com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao excluir", str(e))

    def limpar(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def voltar_para_menu(self):
        """Reexibe o menu imediatamente e faz a limpeza em background para não bloquear a UI."""
        try:
            if getattr(self, "janela_menu", None):
                self.janela_menu.deiconify()
                self.janela_menu.state("zoomed")
                self.janela_menu.lift()
                self.janela_menu.update_idletasks()
                try:
                    self.janela_menu.focus_force()
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
                    self.after(0, self.destroy)
                except Exception:
                    try:
                        self.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()

    def on_closing(self):
        self.conn.close()
        self.destroy()
