import tkinter as tk
from tkinter import ttk, messagebox,filedialog, TclError
import queue
import threading
from datetime import datetime, date
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from conexao_db import conectar
from matplotlib.ticker import MultipleLocator
import matplotlib.dates as mdates
from matplotlib.dates import num2date
import calendar
import re
import locale
from grafico_dolar import AplicacaoGraficoDolar
import tempfile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import matplotlib.patches as mpatches
import re

# Define a localidade para nomes de meses em português
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")  # Linux/macOS
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, "pt_BR")  # Windows
    except locale.Error:
        pass  # Caso não funcione, mantém os nomes em inglês

# Configurações gerais do matplotlib para um visual clean e moderno
matplotlib.rcParams.update({
    "font.family": "Segoe UI",
    "font.size": 10,
    "axes.titlesize": 12,
    "axes.labelsize": 10,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "figure.facecolor": "white",
    "axes.edgecolor": "#333333",
    "grid.color": "#cccccc",
    "grid.linestyle": "--",
})

class AplicacaoGrafico(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(background="#f2f2f2")
        self.pack(fill=tk.BOTH, expand=True)

        self.EstiloGraficoCotacao()

        # Flag para indicar se a atualização dos gráficos deve ocorrer
        self.update_active = True

        self.tooltip = None
        self.tooltip_fixed = False
        self.fila_produtos = queue.Queue()
        self.encerrando = False
        self.iniciar_verificacao()
        # inicia o loop de verificação
        self._agendar_verificacao()

        self.abas = ttk.Notebook(self)
        self.abas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.aba_produtos = ttk.Frame(self.abas)
        self.aba_dolar = ttk.Frame(self.abas)

        self.abas.add(self.aba_produtos, text="Gráfico de Produtos")
        self.abas.add(self.aba_dolar, text="Gráfico de Dólar")

        self.configurar_aba_produtos()
        self.grafico_dolar = AplicacaoGraficoDolar(self.aba_dolar)  # Instancia o gráfico do Dólar
        self.linhas_produtos = [] 
        self.verificar_fila()

        # conecta apenas _mostrar_tooltip para motion e click
        self.canvas_produtos.mpl_connect("motion_notify_event", self._mostrar_tooltip)
        self.canvas_produtos.mpl_connect("button_press_event", self._mostrar_tooltip)
        self.canvas_produtos.mpl_connect("figure_leave_event", self._esconder_tooltip)

        # Vincula o evento de mudança de aba para pausar/retomar atualizações.
        self.abas.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Garante que o on_closing seja chamado ao fechar a janela
        self.winfo_toplevel().protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_tab_changed(self, event):
        try:
            tab_atual = self.abas.select()
            texto_aba = self.abas.tab(tab_atual, "text")
            print(f"Aba atual: {texto_aba}")

            # Aba de Produtos
            if texto_aba == "Gráfico de Produtos":
                self.update_active = True
                print("Ativando atualizações do gráfico de Produtos")
                self.iniciar_atualizacao_produtos()
                self.atualizar_treeview_periodos()      # <- popula seu Treeview de período
                self.verificar_fila()

            # Aba de Dólar
            elif texto_aba == "Gráfico de Dólar":
                # Pare qualquer verificação pendente de produtos
                self.update_active = False
                if hasattr(self, 'id_verificar'):
                    try: self.after_cancel(self.id_verificar)
                    except: pass

                print("Ativando atualizações do gráfico de Dólar")
                # Se for interessante, escopo separado de update_active para dólar
                # dispara o refresh do gráfico e do Treeview de datas
                self.grafico_dolar.iniciar_atualizacao_dolar()
                self.grafico_dolar.atualizar_treeview_datas()

            # Qualquer outra aba...
            else:
                self.update_active = False
                print("Pausando atualizações")
                if hasattr(self, 'id_verificar'):
                    try: self.after_cancel(self.id_verificar)
                    except: pass

        except tk.TclError as e:
            print("Erro no on_tab_changed:", e)

    def EstiloGraficoCotacao(self):
        style = ttk.Style(self)
        style.theme_use("alt")

        cor_fundo = "#f2f2f2"
        cor_treeview = "#ffffff"

        style.configure("TFrame", background=cor_fundo)
        style.configure("TLabel", background=cor_fundo, foreground="#333333")
        style.configure("TButton", background="#e6e6e6", foreground="#333333")

        style.configure("Estat.TLabelframe",
                        background=cor_fundo,
                        borderwidth=1,
                        relief="groove")
        style.configure("Estat.TLabelframe.Label", font=("Segoe UI", 10, "bold"),
                        foreground="#1f4e79", background=cor_fundo)
        style.configure("Estat.TLabel", background=cor_fundo, foreground="#444",
                        font=("Segoe UI", 9))
        style.configure("EstatValor.TLabel", background=cor_fundo, foreground="#0b5394",
                        font=("Segoe UI", 10, "bold"))

        style.configure("Custom.Treeview",
                        background=cor_treeview,
                        fieldbackground=cor_treeview,
                        foreground="#333333",
                        borderwidth=0,
                        rowheight=22,
                        font=("Segoe UI", 9))

        style.configure("Custom.Treeview.Heading",
                        foreground="#000000",
                        font=("Segoe UI", 9, "bold"))

        style.map("Custom.Treeview",
                  background=style.map("Treeview", "background"),
                  foreground=style.map("Treeview", "foreground"))

    def configurar_aba_produtos(self):
        frame_principal = ttk.Frame(self.aba_produtos)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        frame_filtros = ttk.Frame(frame_principal)
        frame_filtros.pack(side=tk.TOP, fill=tk.X, pady=5)

        # Cria o frame da árvore dos períodos
        frame_tree_btn = ttk.LabelFrame(frame_filtros, text="📅 Períodos", style="Estat.TLabelframe")
        frame_tree_btn.grid(row=0, column=0, rowspan=3, padx=5, pady=5, sticky=tk.NW)

        btn_atualizar = ttk.Button(frame_tree_btn, text="Atualizar Período", command=self.atualizar_treeview_periodos)
        btn_atualizar.grid(row=0, column=0, padx=5, pady=(5, 2), sticky=tk.W)

        self.treeview_periodos = ttk.Treeview(frame_tree_btn, selectmode="extended", height=7, style="Custom.Treeview")
        self.treeview_periodos.grid(row=1, column=0, padx=(5, 0), pady=(2, 5), sticky=tk.W + tk.E)
        self.treeview_periodos["show"] = "tree"

        scrollbar = ttk.Scrollbar(frame_tree_btn, orient="vertical", command=self.treeview_periodos.yview)
        scrollbar.grid(row=1, column=1, padx=5, pady=(2, 5), sticky=tk.NS)
        self.treeview_periodos.configure(yscrollcommand=scrollbar.set)

        self.treeview_periodos.insert("", "end", text="📅 Períodos Agrupados", open=True)
        self.preencher_treeview_periodos()
        # Use um método dedicado para o binding com tratamento de erros
        self.treeview_periodos.bind("<<TreeviewSelect>>", self.treeview_select_handler)

        lbl_produtos = ttk.Label(frame_filtros, text="Produtos:", font=("Segoe UI", 10, "bold"))
        lbl_produtos.grid(row=0, column=1, padx=5, pady=(8, 0), sticky=tk.NW)

        self.produtos_selecionados = {}
        lista_produtos = [
            ("Cobre", "cobre"),
            ("Zinco", "zinco"),
            ("Liga 62/38", "liga_62_38"),
            ("Liga 65/35", "liga_65_35"),
            ("Liga 70/30", "liga_70_30"),
            ("Liga 85/15", "liga_85_15")
        ]

        frame_checkboxes = ttk.Frame(frame_filtros)
        frame_checkboxes.grid(row=0, column=2, columnspan=6, padx=5, pady=5, sticky=tk.NW)
        for col, (texto, chave) in enumerate(lista_produtos):
            var = tk.BooleanVar(self, value=True)
            cb = ttk.Checkbutton(frame_checkboxes, text=texto, variable=var, command=self.iniciar_atualizacao_produtos)
            cb.grid(row=0, column=col, padx=4, pady=2)
            self.produtos_selecionados[chave] = var

        frame_estatisticas = ttk.LabelFrame(
            frame_filtros,
            text="📊 Estatísticas dos Produtos",
            padding=(10, 8),
            labelanchor="n",
            style="Estat.TLabelframe",
            relief=tk.GROOVE,
            borderwidth=1
        )
        frame_estatisticas.grid(row=1, column=1, columnspan=6, padx=5, pady=(0, 5), sticky=tk.NW)

        frame_tree_scroll = ttk.Frame(frame_estatisticas)
        frame_tree_scroll.pack(fill=tk.BOTH, expand=True)

        scrollbar_vertical = ttk.Scrollbar(frame_tree_scroll, orient=tk.VERTICAL)
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_estatisticas = ttk.Treeview(
            frame_tree_scroll,
            columns=("produto", "max", "min", "media"),
            show="headings",
            height=6,
            yscrollcommand=scrollbar_vertical.set,
            style="Custom.Treeview"
        )
        scrollbar_vertical.config(command=self.tree_estatisticas.yview)

        self.tree_estatisticas.heading("produto", text="Produto")
        self.tree_estatisticas.heading("max", text="Máximo")
        self.tree_estatisticas.heading("min", text="Mínimo")
        self.tree_estatisticas.heading("media", text="Média")

        self.tree_estatisticas.column("produto", width=120, anchor=tk.CENTER)
        self.tree_estatisticas.column("max", width=80, anchor=tk.CENTER)
        self.tree_estatisticas.column("min", width=80, anchor=tk.CENTER)
        self.tree_estatisticas.column("media", width=80, anchor=tk.CENTER)

        self.tree_estatisticas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.frame_legenda = ttk.Frame(frame_filtros)
        self.frame_legenda.grid(row=0, column=8, rowspan=3, padx=10, pady=5, sticky=tk.N)
        self.lbl_legenda = ttk.Label(self.frame_legenda, text="Legenda:", font=("Segoe UI", 10, "bold"))
        self.lbl_legenda.pack(anchor="w")
        self.legenda_labels = {}

        frame_grafico = ttk.Frame(frame_principal)
        frame_grafico.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.fig_produtos = Figure(figsize=(8, 6), dpi=100)
        self.eixo_produtos = self.fig_produtos.add_subplot(111)
        self.canvas_produtos = FigureCanvasTkAgg(self.fig_produtos, master=frame_grafico)
        self.canvas_produtos.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas_produtos, frame_grafico)
        toolbar.update()
        self.canvas_produtos._tkcanvas.pack(fill=tk.BOTH, expand=True)

        self.iniciar_atualizacao_produtos()
        self.id_verificar = self.after(1000, self.verificar_fila)

    def atualizar_treeview_periodos(self):
        # sai se já estamos encerrando ou se o widget foi destruído
        if self.encerrando or not getattr(self, "treeview_periodos", None) or \
        not self.treeview_periodos.winfo_exists():
            return
        
        self.preencher_treeview_periodos()
        # Opcional: forçar o redesenho imediato do widget
        self.treeview_periodos.update_idletasks()

    def formatar_produto(self, produto):
        """Formata o nome do produto com substituições específicas para ligas."""
        nome = str(produto).lower()
        if nome.startswith("liga_"):
            partes = nome.split("_")
            if len(partes) == 3:
                nome = f"Liga {partes[1]}/{partes[2]}"
                return nome
        return nome.capitalize()

    def treeview_select_handler(self, event):
        if self.encerrando:
            return
        """Callback de seleção da treeview com debounce para evitar atualizações excessivas."""
        try:
            if hasattr(self, "treeview_debounce_id") and self.treeview_debounce_id:
                self.after_cancel(self.treeview_debounce_id)
            # Debounce com delay de 300ms
            self.treeview_debounce_id = self.after(300, self._executar_atualizacao_produtos_segura)
        except tk.TclError:
            pass

    def _executar_atualizacao_produtos_segura(self):
        """Executa a atualização de forma segura, ignorando erros de interface (como ao trocar de aba)."""
        try:
            self.iniciar_atualizacao_produtos()
        except tk.TclError:
            pass
        finally:
            self.treeview_debounce_id = None  # Limpa o ID após execução

    def atualizar_legenda(self):
        """Atualiza a legenda exibida no canto superior direito."""
        try:
            # Remove todos os labels antigos, exceto o lbl_legenda
            for widget in self.frame_legenda.winfo_children():
                if isinstance(widget, ttk.Label) and widget != self.lbl_legenda:
                    widget.destroy()
            # Cria novos labels para cada produto
            for linha, produto, _ in self.linhas_produtos:
                cor = linha.get_color()
                produto_formatado = self.formatar_produto(produto)
                lbl_legenda = ttk.Label(self.frame_legenda,
                                        text=f"⬤ {produto_formatado}",
                                        foreground=cor,
                                        font=("Segoe UI", 9))
                lbl_legenda.pack(anchor="w", pady=1)
                self.legenda_labels[produto] = lbl_legenda
        except tk.TclError:
            pass

    def preencher_treeview_periodos(self):
        # Conexão e consulta ao banco de dados (continua igual)
        conn = conectar()
        if conn is None:
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados!")
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT periodo FROM cotacao_produtos ORDER BY periodo ASC")
            periodos = [row[0] for row in cursor.fetchall()]
            cursor.close()
            conn.close()
        
            # Limpa todos os itens atuais do treeview
            for item in self.treeview_periodos.get_children():
                self.treeview_periodos.delete(item)
        
            # Inserindo o nó raiz
            raiz_id = self.treeview_periodos.insert("", tk.END, text="📅 Períodos Agrupados", open=True)
        
            # Agrupamento dos períodos conforme sua lógica
            grupos = {}
            for p in periodos:
                p_str = p.strip()
                datas = re.split(r'\s*[aá]\s*', p_str)
                datas = [d.strip() for d in datas if d.strip()]
                if not datas:
                    continue
                try:
                    dia1, mes1, ano1 = datas[0].split("/")[:3]
                    ano_formatado = f"20{ano1}" if len(ano1) == 2 and ano1.isdigit() else "20??"
                except Exception:
                    continue
                if ano_formatado not in grupos:
                    grupos[ano_formatado] = {}
                if mes1 not in grupos[ano_formatado]:
                    grupos[ano_formatado][mes1] = set()
                grupos[ano_formatado][mes1].add(p_str)
                if len(datas) > 1:
                    try:
                        dia1, mes1, ano1 = datas[0].split("/")[:3]
                        dia2, mes2, ano2 = datas[1].split("/")[:3]
                        
                        data_inicio = datetime.strptime(f"{dia1}/{mes1}/{ano1}", "%d/%m/%y")
                        data_fim = datetime.strptime(f"{dia2}/{mes2}/{ano2}", "%d/%m/%y")
                        
                        data_atual = data_inicio
                        while data_atual <= data_fim:
                            ano = str(data_atual.year)
                            mes = str(data_atual.month).zfill(2)

                            if ano not in grupos:
                                grupos[ano] = {}
                            if mes not in grupos[ano]:
                                grupos[ano][mes] = set()
                            grupos[ano][mes].add(p_str)

                            # Pula para o primeiro dia do próximo mês
                            if data_atual.month == 12:
                                data_atual = datetime(data_atual.year + 1, 1, 1)
                            else:
                                data_atual = datetime(data_atual.year, data_atual.month + 1, 1)
                    except Exception:
                        pass
                        
            # Insere os grupos na treeview
            for ano in sorted(grupos.keys(), key=lambda x: int(x) if x.isdigit() else 0):
                ano_id = self.treeview_periodos.insert(raiz_id, tk.END, text=ano, open=False)
                for mes in sorted(grupos[ano].keys(), key=lambda m: int(m)):
                    MESES_PT = {
                        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
                        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
                        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
                    }
                    nome_mes = MESES_PT.get(int(mes), f"Mês {mes}")
                    mes_id = self.treeview_periodos.insert(ano_id, tk.END, text=nome_mes, open=False)
                    def get_start_date(periodo_str):
                        try:
                            return datetime.strptime(re.split(r'\s*[aá]\s*', periodo_str)[0].strip(), "%d/%m/%y")
                        except Exception:
                            return datetime.max
                    for periodo in sorted(list(grupos[ano][mes]), key=get_start_date):
                        self.treeview_periodos.insert(mes_id, tk.END, text=periodo)
        
            # Força o redesenho do treeview
            self.treeview_periodos.update_idletasks()
        
        except tk.TclError:
            pass
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar períodos: {e}")

    def obter_periodos_selecionados(self):
        """Retorna uma lista de períodos selecionados na treeview."""
        selecionados = set()

        def coletar_periodos(item):
            filhos = self.treeview_periodos.get_children(item)
            if not filhos:
                texto = self.treeview_periodos.item(item, 'text')
                if texto and texto != "📅 Períodos Agrupados":
                    selecionados.add(texto)
            else:
                for filho in filhos:
                    coletar_periodos(filho)

        try:
            for item in self.treeview_periodos.selection():
                coletar_periodos(item)
        except tk.TclError:
            pass

        return list(selecionados)

    def buscar_dados_produtos(self):
        # Sai cedo se já estivermos encerrando ou se o widget não existir
        if self.encerrando or not self.treeview_periodos.winfo_exists():
            return []
        
        """Busca os dados dos produtos de acordo com os períodos selecionados."""
        conn = conectar()
        if conn is None:
            return []
        try:
            cursor = conn.cursor()
            periodos_selecionados = self.obter_periodos_selecionados()
            if periodos_selecionados:
                periodos_str = ", ".join(f"'{p}'" for p in periodos_selecionados)
                consulta = f"""
                    SELECT periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15 
                    FROM cotacao_produtos
                    WHERE periodo IN ({periodos_str})
                    ORDER BY periodo
                """
            else:
                consulta = """
                    SELECT periodo, cobre, zinco, liga_62_38, liga_65_35, liga_70_30, liga_85_15 
                    FROM cotacao_produtos
                    ORDER BY periodo
                """
            cursor.execute(consulta)
            dados = cursor.fetchall()
            cursor.close()
            conn.close()

            # Verifica como os períodos foram selecionados para determinar o agrupamento
            selecionados_tree = self.treeview_periodos.selection()
            agrupados_por_mes = agrupados_por_ano = False
            for item in selecionados_tree:
                texto_item = self.treeview_periodos.item(item, 'text')
                filhos = self.treeview_periodos.get_children(item)
                if texto_item == "📅 Períodos Agrupados":
                    agrupados_por_ano = True
                elif filhos and self.treeview_periodos.get_children(filhos[0]):
                    agrupados_por_mes = True

            if agrupados_por_ano or agrupados_por_mes:
                dados_por_chave = {}
                for linha in dados:
                    periodo = linha[0]
                    try:
                        datas = re.split(r'\s*[aá]\s*', periodo)
                        data_inicio = datetime.strptime(datas[0].strip(), "%d/%m/%y")
                        data_fim = datetime.strptime(datas[1].strip(), "%d/%m/%y") if len(datas) > 1 else data_inicio

                        # Agora percorre todos os meses entre data_inicio e data_fim
                        atual = data_inicio
                        while atual <= data_fim:
                            chave = (atual.year, atual.month) if agrupados_por_mes else atual.year
                            if chave not in dados_por_chave:
                                dados_por_chave[chave] = []
                            dados_por_chave[chave].append(linha[1:])

                            # Avança para o primeiro dia do próximo mês
                            if atual.month == 12:
                                atual = datetime(atual.year + 1, 1, 1)
                            else:
                                atual = datetime(atual.year, atual.month + 1, 1)
                    except Exception:
                        continue

                # Calcula médias
                dados_agrupados = []
                for chave, valores in sorted(dados_por_chave.items()):
                    media = [sum(col) / len(col) for col in zip(*valores)]
                    if agrupados_por_mes:
                        ano, mes = chave
                        periodo_label = f"{calendar.month_name[mes].capitalize()} {ano}"
                    else:
                        periodo_label = str(chave)
                    dados_agrupados.append((periodo_label, *media))

                return dados_agrupados

            # Se não estiver agrupado, apenas retorna os dados crus
            return dados

        except tk.TclError:
            return []
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar dados: {e}")
            return []

    def iniciar_atualizacao_produtos(self):
        """Inicia a atualização dos dados dos produtos em uma thread separada."""
        thread = threading.Thread(target=self.thread_buscar_produtos, daemon=True)
        thread.start()

    def thread_buscar_produtos(self):
        dados = self.buscar_dados_produtos()
        if dados:
            self.fila_produtos.put(dados)

    def iniciar_verificacao(self):
        self.id_verificar = self.after(1000, self.verificar_fila)

    def _agendar_verificacao(self):
        if not self.encerrando:
            self.id_verificar = self.after(1000, self.verificar_fila)

    def verificar_fila(self):
        # Atualiza só se não estivermos encerrando
        if self.encerrando or not self.winfo_exists():
            return

        # Processa fila
        try:
            dados = self.fila_produtos.get_nowait()
            self.atualizar_grafico_produtos(dados)
        except queue.Empty:
            pass

        # Agenda próxima verificação só se não estiver encerrando
        if not self.encerrando:
            self.id_verificar = self.after(200, self.verificar_fila)
        else:
            print("Finalizando AplicacaoGrafico")
            self._esperar_encerramento()

    def atualizar_grafico_produtos(self, dados):
        if not dados or not self.winfo_ismapped():
            return

        try:
            self.eixo_produtos.clear()
            self.eixo_produtos.set_facecolor("white")
            self.eixo_produtos.grid(True, color="#e6e6e6", linestyle="--", linewidth=0.8)
        except:
            return

        import locale
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # ou 'Portuguese_Brazil.1252' no Windows

        try:
            periodos_dt, periodos_str, dados_sorted = [], [], []

            for linha in dados:
                p = linha[0]
                try:
                    if " á " in p:
                        dt_obj = datetime.strptime(p.split(" á ")[0].strip(), "%d/%m/%y").date()
                    else:
                        partes = p.strip().split()
                        if len(partes) == 2:
                            mes_nome, ano = partes
                            data = datetime.strptime(f"{mes_nome} {ano}", "%B %Y")
                            dt_obj = data.date()
                        else:
                            raise ValueError("Formato inválido de período")
                except Exception:
                    dt_obj = date.max  # Coloca no final da ordenação se erro
                periodos_dt.append(dt_obj)
                periodos_str.append(p)
                dados_sorted.append(linha)

            # --- ordena por data e constrói as tuplas (periodo_dt, periodo_str, dados) ---
            combinados = list(zip(periodos_dt, periodos_str, dados_sorted))
            combinados.sort()
            periodos_dt, periodos_str_sorted, dados_sorted = zip(*combinados)

            # <-- AQUI você armazena a lista de períodos exatamente como aparecerá no eixo X -->
            self._periodos_str_plotados = list(periodos_str_sorted)

            x_positions = list(range(len(periodos_str_sorted)))
            indices = {
                "cobre": 1, "zinco": 2,
                "liga_62_38": 3, "liga_65_35": 4,
                "liga_70_30": 5, "liga_85_15": 6
            }

            cores = {
                "cobre": "#b87333", "zinco": "#a8a8a8",
                "liga_62_38": "#0072B2", "liga_65_35": "#FF0000",
                "liga_70_30": "#009E73", "liga_85_15": "#CC79A7",
            }

            self.linhas_produtos = []
            for produto, idx in indices.items():
                if self.produtos_selecionados[produto].get():
                    valores = [linha[idx] for linha in dados_sorted]
                    linha_plot, = self.eixo_produtos.plot(
                        x_positions, valores, marker='o', label=produto.replace("_", " "), linewidth=2,
                        color=cores.get(produto, "#000000")
                    )
                    self.linhas_produtos.append((linha_plot, produto, valores))

            self.eixo_produtos.set_title("Cotação dos Produtos", fontsize=14)
            self.eixo_produtos.set_xlabel("Período")
            self.eixo_produtos.set_ylabel("Valor")
            self.eixo_produtos.set_xticks(x_positions)

            # Abreviação automática dos meses, preservando outros formatos
            rotulos_abreviados = []
            for dt, texto_original in zip(periodos_dt, periodos_str_sorted):
                partes = texto_original.strip().split()
                if len(partes) == 2:
                    rotulos_abreviados.append(dt.strftime("%b %Y").capitalize())
                else:
                    rotulos_abreviados.append(texto_original)

            self.eixo_produtos.set_xticklabels(
                rotulos_abreviados, rotation=45, ha="right", fontsize=9
            )
            self.eixo_produtos.yaxis.set_major_locator(MultipleLocator(10))
            self.fig_produtos.tight_layout()
            self.canvas_produtos.draw()

        except Exception:
            return

        self.atualizar_estatisticas()
        self.atualizar_legenda()


    def formatar_produto(self, nome):
        nome = nome.lower().replace(" ", "_")
        nome_formatado = nome.replace("_", " ").title()
        nome_formatado = nome_formatado.replace("62 38", "62/38")
        nome_formatado = nome_formatado.replace("65 35", "65/35")
        nome_formatado = nome_formatado.replace("70 30", "70/30")
        nome_formatado = nome_formatado.replace("85 15", "85/15")
        return nome_formatado

    def atualizar_estatisticas(self):
        # → Sai se estivermos encerrando ou se o widget não existir
        if (self.encerrando
            or not getattr(self, "tree_estatisticas", None)
            or not self.tree_estatisticas.winfo_exists()):
            return
        try:
            self.tree_estatisticas.delete(*self.tree_estatisticas.get_children())
            for linha_plot, chave, valores in self.linhas_produtos:
                if not valores:
                    continue
                max_val = max(valores)
                min_val = min(valores)
                media_val = round(sum(valores) / len(valores), 2)
                print("Label original:", linha_plot.get_label())
                nome_formatado = self.formatar_produto(linha_plot.get_label())
                print("Nome formatado:", nome_formatado)
                def formatar(v): return f"R$ {str(round(v, 2)).replace('.', ',')}"
                self.tree_estatisticas.insert("", "end", values=(
                    nome_formatado, formatar(max_val), formatar(min_val), formatar(media_val)
                ))
        except:
            pass

    def _mostrar_tooltip(self, event):
        """
        Exibe uma tooltip anotada no ponto mais próximo do mouse,
        mostrando:
        - Nome do produto
        - Período (no mesmo formato em que foi plotado no eixo X:
            ex: "01/01/2025 a 31/01/2025", ou "Jan 2025", ou "2025")
        - Valor (R$) formatado

        Pressionar o botão esquerdo (click) fixa/desfixa a tooltip.
        """
        try:
            canvas = self.canvas_produtos
            eixo   = self.eixo_produtos

            # Se o widget foi destruído, aborta
            if not canvas.get_tk_widget().winfo_exists():
                return

            # Se não temos dados de coordenada, aborta
            if event.xdata is None or event.ydata is None:
                return

            # Clique esquerdo: fixa/desfixa
            if event.button == 1:
                if self.tooltip and not self.tooltip_fixed:
                    self.tooltip_fixed = True
                    return
                elif self.tooltip and self.tooltip_fixed:
                    self.tooltip.set_visible(False)
                    canvas.draw_idle()
                    self.tooltip = None
                    self.tooltip_fixed = False
                    return

            # Se já está fixado e é só movimento, não atualiza
            if self.tooltip_fixed and event.name == 'motion_notify_event':
                return

            # --- 1) Encontra o ponto MAIS PRÓXIMO do mouse no plot ---
            pt_prox = None
            dmin    = float('inf')
            linha_origem = None

            for linha in eixo.lines:
                # cada linha contém um array de (x, y)
                for x_val, y_val in linha.get_xydata():
                    d = ((event.xdata - x_val)**2 + (event.ydata - y_val)**2) ** 0.5
                    # tolerância de 2 unidades (ajuste se quiser maior/menor)
                    if d < dmin and d < 2:
                        dmin = d
                        pt_prox = (x_val, y_val)
                        linha_origem = linha

            # Se não encontrou ponto próximo, esconde a tooltip e sai
            if pt_prox is None:
                if self.tooltip and not self.tooltip_fixed:
                    self.tooltip.set_visible(False)
                    self.tooltip = None
                    canvas.draw_idle()
                return

            x_val, y_val = pt_prox  # ponto mais próximo

            # --- 2) Formatação do NOME DO PRODUTO ---
            import re
            label_raw = linha_origem.get_label() or "Produto"
            if label_raw.lower().startswith("liga"):
                m = re.match(r"liga[_ ]?(\d+)[_ ]?(\d+)", label_raw.lower())
                if m:
                    label = f"Liga {m.group(1)}/{m.group(2)}"
                else:
                    label = label_raw.replace("_", " ").capitalize()
            else:
                label = label_raw.replace("_", " ").capitalize()

            # --- 3) Formatação do VALOR ---
            valor_str = f"{y_val:.2f}".replace(".", ",")

            # --- 4) Traduzir x_val (índice) para o PERÍODO em texto ---
            # Aqui assumimos que, em atualizar_grafico_produtos, 
            # você fez:
            #   self._periodos_str_plotados = list(periodos_str_sorted)
            # e que cada ponto plotado em eixo.lines usa índice inteiro (0,1,2,...).
            try:
                # Converte x_val para inteiro mais próximo
                idx = int(round(x_val))
                if (
                    hasattr(self, "_periodos_str_plotados")
                    and 0 <= idx < len(self._periodos_str_plotados)
                ):
                    texto_original = self._periodos_str_plotados[idx]

                    # Inicializa data_formatada com texto_original para garantir definição
                    data_formatada = texto_original

                    # Se o texto original for "Mês Ano", tenta converter para abreviação
                    partes = texto_original.strip().split()
                    if len(partes) == 2:
                        import locale
                        from datetime import datetime
                        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # ou 'Portuguese_Brazil.1252' no Windows
                        try:
                            data_dt = datetime.strptime(f"{partes[0]} {partes[1]}", "%B %Y")
                            data_formatada = data_dt.strftime("%b %Y").capitalize()
                        except Exception:
                            # Mantém texto_original se o parse falhar
                            pass
                    # Caso não sejam duas palavras, data_formatada permanece texto_original

                else:
                    raise IndexError

            except Exception:
                # Se der qualquer erro (índice fora de faixa, etc.), define como string do índice
                data_formatada = str(x_val)

            # --- 5) Monta o texto final da tooltip ---
            texto = (
                f"Produto: {label}\n"
                f"Período: {data_formatada}\n"
                f"Valor: R$ {valor_str}"
            )

            # --- 6) Posiciona a annotation SEMPRE ABAIXO DO PONTO e seta preta ---
            inv = eixo.transData.transform
            x_canvas, _ = inv((x_val, y_val))
            largura, _ = canvas.get_width_height()

            # deslocamento vertical negativo força a tooltip ficar abaixo
            dy = -30
            va = 'top'   # alinha a parte superior da box ao ponto, então aparece abaixo

            # deslocamento horizontal padrão
            dx = 10
            ha = 'left'

            # Se estiver próximo da borda direita, inverte horizontalmente
            if x_canvas > largura - 100:
                dx, ha = -100, 'right'

            desloc = (dx, dy)

            # --- 7) Cria ou atualiza a annotation (tooltip) com seta preta ---
            if self.tooltip is None:
                self.tooltip = eixo.annotate(
                    texto,
                    xy=(x_val, y_val),
                    xytext=desloc,
                    textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.9),
                    arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0', color='black'),
                    fontsize=10,
                    ha=ha,
                    va=va,
                    zorder=100
                )
                self.tooltip.set_clip_on(False)
            else:
                self.tooltip.xy = (x_val, y_val)
                self.tooltip.set_text(texto)
                self.tooltip.set_position(desloc)
                self.tooltip.set_ha(ha)
                self.tooltip.set_va(va)
                self.tooltip.get_bbox_patch().set_color('Yellow')
                self.tooltip.set_visible(True)

            canvas.draw_idle()

        except TclError:
            return

    def _esconder_tooltip(self, event):
        try:
            if self.tooltip and not self.tooltip_fixed:
                self.tooltip.set_visible(False)
                self.tooltip = None
                self.canvas_produtos.draw_idle()
        except TclError:
            pass
    
    def exportar_grafico_excel(self):
        try:
            # 1. Verificar qual aba está ativa
            aba_index = self.abas.index(self.abas.select())
            if aba_index == 0:
                # Aba "Gráfico de Produtos"
                fig = self.fig_produtos
                sheet_name = "Gráfico de Produtos"
            elif aba_index == 1:
                # Aba "Gráfico de Dólar"
                try:
                    fig = self.grafico_dolar.fig
                except AttributeError:
                    messagebox.showerror("Exportação", "O gráfico de Dólar não está disponível para exportação.")
                    return
                sheet_name = "Gráfico de Dólar"
            else:
                messagebox.showerror("Exportação", "A aba selecionada não possui um gráfico para exportar.")
                return

            # 2. Obter o eixo principal da figura original
            ax = fig.axes[0] if fig.axes else None
            if ax is None:
                messagebox.showerror("Exportação", "Nenhum eixo encontrado no gráfico.")
                return

            # Salvar informações de título e dos rótulos dos eixos para exportação
            titulo = ax.get_title()
            xlabel = ax.get_xlabel()
            ylabel = ax.get_ylabel()

            # Salvar os handles e labels para a legenda (de cores dos produtos)
            handles, labels = ax.get_legend_handles_labels()

            # 3. Criar uma nova figura para exportação sem interferir na interface
            fig_export = Figure(figsize=(6.5, 4.5), dpi=150)
            ax_export = fig_export.add_subplot(111)

            # 4. Copiar os dados (linhas) do gráfico original para a nova figura
            for line in ax.get_lines():
                ax_export.plot(
                    line.get_xdata(),
                    line.get_ydata(),
                    label=line.get_label(),
                    color=line.get_color(),
                    linestyle=line.get_linestyle(),
                    marker=line.get_marker()
                )

            # 5. Configurar título, rótulos dos eixos e grid na figura exportada
            ax_export.set_title(titulo)
            ax_export.set_xlabel(xlabel)
            ax_export.set_ylabel(ylabel)
            ax_export.grid(True)

            # 6. Copiar os ticks e os respectivos rótulos (por exemplo, períodos) do eixo x
            xticks = ax.get_xticks()
            xticklabels = [item.get_text() for item in ax.get_xticklabels()]
            ax_export.set_xticks(xticks)
            ax_export.set_xticklabels(xticklabels, rotation=45, ha='right')  # ajuste a rotação se necessário

            # 7. Adicionar a legenda de cores dos produtos na figura exportada
            # Se houver handles e labels, recrie a legenda
            if handles and labels:
                ax_export.legend(handles, labels, loc="upper right", fontsize=8)

            fig_export.tight_layout()

            # 8. Salvar a figura exportada num arquivo temporário (PNG)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                temp_file = tmp.name

            # Redefinir o tamanho da imagem (largura fixa e altura proporcional)
            largura_desejada = 6.5  # em polegadas
            altura_desejada = 3.5   # em polegadas
            largura_original = fig_export.get_size_inches()[0]
            proporcao = largura_desejada / largura_original
            altura_ajustada = altura_desejada * proporcao
            fig_export.set_size_inches(largura_desejada, altura_ajustada)
            fig_export.savefig(temp_file, bbox_inches="tight", dpi=150)

            # 9. Solicitar ao usuário o local e o nome do arquivo Excel para salvar a imagem
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Planilhas do Excel", "*.xlsx")],
                title="Salvar Gráfico em Excel",
                initialfile="Grafico_de_Cotação_Produto.xlsx"
            )
            if not caminho_arquivo:
                return  # Usuário cancelou

            # 10. Inserir a imagem exportada em um arquivo Excel novo
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name

            img = XLImage(temp_file)
            ws.add_image(img, 'A1')

            wb.save(caminho_arquivo)
            messagebox.showinfo("Exportação", f"Gráfico exportado com sucesso!\n{caminho_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro na Exportação", f"Erro ao exportar o gráfico:\n{str(e)}")

    def exportar_grafico_pdf(self):
        """Exporta o gráfico de produtos em um arquivo PDF, adicionando temporariamente uma legenda
        que indica qual cor corresponde a cada produto e, depois, removendo essa legenda da interface."""
        # Abre o diálogo para escolher onde salvar o PDF
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="grafico_produtos.pdf",
            filetypes=[("Arquivo PDF", "*.pdf")],
            title="Salvar Gráfico em PDF"
        )

        if not caminho_arquivo:
            return

        try:
            # Obtém o eixo principal da figura
            ax = self.fig_produtos.get_axes()[0]
            
            # Salva a legenda original (se existir) para restaurá-la depois
            legenda_original = ax.get_legend()

            # Cria a legenda temporária com base em self.linhas_produtos
            if hasattr(self, "linhas_produtos") and self.linhas_produtos:
                legend_handles = [
                    mpatches.Patch(color=linha.get_color(), label=self.formatar_produto(produto))
                    for linha, produto, _ in self.linhas_produtos
                ]
                legenda_temp = ax.legend(handles=legend_handles, title="Produtos",
                                        loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3)
            else:
                legenda_temp = None

            # Salva o gráfico com a legenda no PDF
            self.fig_produtos.savefig(caminho_arquivo, format="pdf", bbox_inches="tight")
            
            # Remove a legenda temporária, de modo que a interface fique inalterada
            if legenda_temp is not None:
                legenda_temp.remove()
                # Redesenha a figura para atualizar a interface (caso você esteja usando um canvas)
                self.canvas_produtos.draw()

            # Se havia uma legenda original, restaura-a
            if legenda_original is not None:
                ax.legend(handles=legenda_original.legendHandles, labels=[t.get_text() for t in legenda_original.get_texts()])

            messagebox.showinfo("Exportar PDF", f"Gráfico exportado com sucesso para:\n{caminho_arquivo}")
        except Exception as e:
            messagebox.showerror("Exportar PDF", f"Erro ao exportar o gráfico: {e}")

    def on_closing(self):
        print("Fechando AplicacaoGrafico")
        # 1) marque que estamos encerrando
        self.encerrando = True

        # 2) cancele TODOS os agendamentos de after que esta classe pode ter
        for job_attr in (
            "id_verificar",
            "job_atualizar_periodos",
            "job_atualizar_datas",
            "sparkline_debounce_id",
            "treeview_debounce_id",
        ):
            job = getattr(self, job_attr, None)
            if job:
                try:
                    self.after_cancel(job)
                except tk.TclError:
                    pass

        # 3) peça para o Toplevel de dólar fechar também
        if hasattr(self, 'grafico_dolar'):
            self.grafico_dolar.on_closing()

        # 4) aguarde um instante e finalize
        self.after(100, self._esperar_encerramento)

    def _esperar_encerramento(self):
        # Se o gráfico de dólar já está encerrando, fecha a janela principal
        if hasattr(self, 'grafico_dolar') and self.grafico_dolar.encerrando:
            print("Gráficos encerrados, fechando janela.")
            self.winfo_toplevel().destroy()
        else:
            # Senão, aguarda mais um pouco
            print("Aguardando encerramento dos gráficos…")
            self.after(100, self._esperar_encerramento)