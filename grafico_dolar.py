import tkinter as tk
from tkinter import ttk, messagebox, filedialog,TclError
import queue
import threading
from datetime import datetime
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import MultipleLocator
from conexao_db import conectar
import locale
from matplotlib.ticker import FuncFormatter
import pandas as pd
import tempfile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')  # Para nomes de meses em português

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

class AplicacaoGraficoDolar(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.configure(background="#f2f2f2")
        self.pack(fill=tk.BOTH, expand=True)

        self.EstiloGrafico()  # Aplicar estilos

        # Variáveis e fila
        self.tooltip = None
        self.tooltip_fixed = False
        self.fila_dolar = queue.Queue()

        self.encerrando = False
        self.id_verificar = None
        self.id_atualizacao = None

        # O conteúdo será desenhado diretamente neste frame
        self.aba_dolar = self
        self.configurar_aba_dolar()

        self.linhas_dolar = []
        self.executando = True  
        self.verificar_fila()

        # … protocolo, estilos, etc.
        self._agendar_verificacao()

        # Defina o temporizador para atualizar o valor do dólar
        self.id_atualizacao_dolar = self.after(1000, self.verificar_fila)  # Atualiza a cada 10 segundos

        self.canvas_destroyed = False
        self.cid = self.canvas_dolar.mpl_connect('motion_notify_event', self.mostrar_tooltip)
        # dispare o evento de cleanup
        tk_widget = self.canvas_dolar.get_tk_widget()
        tk_widget.bind('<Destroy>', self._on_destroy, add='+')

    def EstiloGrafico(self):
        style = ttk.Style(self)
        style.theme_use("alt")  # "alt" respeita bem a aparência nativa

        cor_fundo = "#f2f2f2"
        cor_treeview = "#ffffff"

        # Estilo geral dos frames e botões
        style.configure("TFrame", background=cor_fundo)
        style.configure("TLabel", background=cor_fundo, foreground="#333333")
        style.configure("TButton", background="#e6e6e6", foreground="#333333")

        # Estilos para os painéis internos
        style.configure("Estat.TLabelframe", background=cor_fundo, borderwidth=1, relief="groove")
        style.configure("Estat.TLabelframe.Label", font=("Segoe UI", 10, "bold"),
                        foreground="#1f4e79", background=cor_fundo)
        style.configure("Estat.TLabel", background=cor_fundo, foreground="#444",
                        font=("Segoe UI", 9))
        style.configure("EstatValor.TLabel", background=cor_fundo, foreground="#0b5394",
                        font=("Segoe UI", 10, "bold"))

        # Estilo para Treeview
        style.configure("Custom.Treeview",
                        background=cor_treeview,
                        fieldbackground=cor_treeview,
                        foreground="#333333",
                        borderwidth=0,
                        rowheight=22,
                        font=("Segoe UI", 9))
        
        # Cabeçalho do Treeview
        style.configure("Custom.Treeview.Heading",
                        foreground="#000000",
                        font=("Segoe UI", 9, "bold"))
        
        style.map("Custom.Treeview",
                background=style.map("Treeview", "background"),
                foreground=style.map("Treeview", "foreground"))

    def configurar_aba_dolar(self):
        frame_principal = ttk.Frame(self.aba_dolar)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        frame_filtros = ttk.Frame(frame_principal)
        frame_filtros.pack(side=tk.TOP, fill=tk.X, pady=5)

        # --- Painel de Datas ---
        # Colocamos o Treeview das datas dentro de um LabelFrame com título "📅 Datas"
        frame_datas = ttk.LabelFrame(frame_filtros, text="📅 Datas", style="Estat.TLabelframe")
        frame_datas.grid(row=0, column=0, padx=5, pady=5, sticky=tk.NW)

        # Usamos um container com grid para que o Treeview e sua barra fiquem lado a lado
        frame_datas_container = ttk.Frame(frame_datas, style="TFrame")
        frame_datas_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        frame_datas_container.grid_rowconfigure(0, weight=1)
        frame_datas_container.grid_columnconfigure(0, weight=1)

        self.treeview_datas = ttk.Treeview(frame_datas_container,
                                           selectmode="extended",
                                           style="Custom.Treeview")
        self.treeview_datas.grid(row=0, column=0, sticky="nsew")
        self.treeview_datas["show"] = "tree"

        scrollbar_datas = ttk.Scrollbar(frame_datas_container,
                                        orient="vertical",
                                        command=self.treeview_datas.yview)
        scrollbar_datas.grid(row=0, column=1, sticky="ns")
        self.treeview_datas.configure(yscrollcommand=scrollbar_datas.set)

        self.treeview_datas.insert("", "end", text="📅 Datas Agrupadas", open=True)
        self.preencher_treeview_datas()
        self.treeview_datas.bind("<<TreeviewSelect>>", lambda event: self.iniciar_atualizacao_dolar())

         # --- Botão para Atualizar a Treeview de Datas ---
        btn_atualizar_datas = ttk.Button(frame_datas,text="Atualizar Datas",command=self.atualizar_treeview_datas)
        btn_atualizar_datas.pack(anchor=tk.W, padx=5, pady=5)

        # --- Painel de Estatísticas Gerais em Treeview ---
        frame_estatisticas = ttk.LabelFrame(frame_filtros,
                                             text="📊 Estatísticas Gerais",
                                             padding=(10, 8),
                                             labelanchor="n",
                                             style="Estat.TLabelframe",
                                             relief=tk.GROOVE,
                                             borderwidth=1)
        frame_estatisticas.grid(row=0, column=1, padx=(20, 0), pady=10, sticky=tk.N)

        self.treeview_estatisticas = ttk.Treeview(frame_estatisticas,
                                                  columns=("estatistica", "valor"),
                                                  show="headings",
                                                  height=3,
                                                  style="Custom.Treeview")
        self.treeview_estatisticas.heading("estatistica", text="Estatística")
        self.treeview_estatisticas.heading("valor", text="Valor")
        for col in ("estatistica", "valor"):
            self.treeview_estatisticas.column(col, anchor="center", width=100)
        self.treeview_estatisticas.pack(padx=5, pady=5)

        # --- Painel Lateral com Sparkline e Tendência ---
        frame_lateral = ttk.LabelFrame(frame_filtros,
                                        text="📈 Mini Gráfico e Tendência",
                                        padding=(10, 8),
                                        labelanchor="n",
                                        style="Estat.TLabelframe",
                                        relief=tk.GROOVE,
                                        borderwidth=1)
        frame_lateral.grid(row=0, column=2, padx=(20, 5), pady=10, sticky=tk.N)

        self.fig_sparkline = Figure(figsize=(2.2, 1), dpi=100)
        self.eixo_sparkline = self.fig_sparkline.add_subplot(111)
        self.eixo_sparkline.axis("off")
        self.canvas_sparkline = FigureCanvasTkAgg(self.fig_sparkline, master=frame_lateral)
        self.canvas_sparkline.get_tk_widget().pack(padx=5, pady=(0, 5))

        self.lbl_tendencia = ttk.Label(frame_lateral, text="Tendência: ➡️", font=("Segoe UI", 10, "bold"))
        self.lbl_tendencia.pack(anchor="center")

        # --- Gráfico Principal ---
        frame_grafico = ttk.Frame(frame_principal)
        frame_grafico.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.fig_dolar = Figure(figsize=(7, 5), dpi=100)
        self.eixo_dolar = self.fig_dolar.add_subplot(111)
        self.canvas_dolar = FigureCanvasTkAgg(self.fig_dolar, master=frame_grafico)
        self.canvas_dolar.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas_dolar, frame_grafico)
        toolbar.update()
        self.canvas_dolar._tkcanvas.pack(fill=tk.BOTH, expand=True)

        self.iniciar_atualizacao_dolar()
        self.id_verificar = self.after(1000, self.verificar_fila)

    def atualizar_treeview_datas(self):
        if self.encerrando or not getattr(self, "treeview_datas", None) \
        or not self.treeview_datas.winfo_exists():
            return
        # Chama o método que preenche a Treeview de datas
        self.preencher_treeview_datas()
        # Força a atualização imediata do widget
        self.treeview_datas.update_idletasks()

    def debounced_atualizar_sparkline_e_tendencia(self, valores):
        if self.encerrando or not self.winfo_exists():
            return
        """
        Agenda a atualização do sparkline e tendência com debounce para evitar atualizações
        excessivas e travamentos.
        """
        try:
            # Cancela uma atualização pendente, se houver
            if hasattr(self, "sparkline_debounce_id") and self.sparkline_debounce_id is not None:
                self.after_cancel(self.sparkline_debounce_id)
            # Agenda a atualização para ocorrer após 300ms (ajuste se necessário)
            self.sparkline_debounce_id = self.after(300, lambda: self.atualizar_sparkline_e_tendencia(valores))
        except tk.TclError:
            pass

    def atualizar_sparkline_e_tendencia(self, valores):
        if not valores or not self.winfo_exists() or not self.executando:
            return

        try:
            self.eixo_sparkline.clear()
            self.eixo_sparkline.plot(valores, color="#4a90e2", linewidth=1.5)
            self.eixo_sparkline.axis("off")
            self.canvas_sparkline.draw()

            if valores[-1] > valores[0]:
                tendencia = "⬆️ Alta"
                cor = "green"
            elif valores[-1] < valores[0]:
                tendencia = "⬇️ Queda"
                cor = "red"
            else:
                tendencia = "➡️ Estável"
                cor = "gray"

            self.lbl_tendencia.config(text=f"Tendência: {tendencia}", foreground=cor)
        except tk.TclError:
            return
        
    def preencher_treeview_datas(self):
        conn = conectar()
        if conn is None:
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados!")
            return
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT data FROM cotacao_dolar ORDER BY data ASC")
            datas = [row[0].strftime("%d/%m/%Y") for row in cursor.fetchall()]
            cursor.close()
            conn.close()

            try:
                for item in self.treeview_datas.get_children():
                    self.treeview_datas.delete(item)

                raiz_id = self.treeview_datas.insert("", "end", text="📅 Datas Agrupadas", open=True)

                grupos = {}
                for d in datas:
                    try:
                        dt = datetime.strptime(d, "%d/%m/%Y")
                    except Exception:
                        continue
                    ano = dt.year
                    mes = dt.month
                    dia = dt.day
                    grupos.setdefault(ano, {}).setdefault(mes, set()).add(d)

                nomes_meses = {
                    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
                    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
                    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
                }

                for ano in sorted(grupos.keys()):
                    ano_id = self.treeview_datas.insert(raiz_id, "end", text=str(ano), open=False)
                    for mes in sorted(grupos[ano].keys()):
                        mes_nome = nomes_meses.get(mes, f"Mês {mes}")
                        mes_id = self.treeview_datas.insert(ano_id, "end", text=mes_nome, open=False)
                        for d in sorted(list(grupos[ano][mes]), key=lambda x: datetime.strptime(x, "%d/%m/%Y")):
                            self.treeview_datas.insert(mes_id, "end", text=d)
            except tk.TclError:
                return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar datas: {e}")

    def obter_selecao_agrupada(self):
        """
        Retorna uma tupla: (lista de datas selecionadas, modo_agrupamento)
        modo_agrupamento: "dia", "mes", "ano"
        """
        selecionados = []
        modo = "dia"
        selecoes = self.treeview_datas.selection()
        if selecoes:
            # Verifica o nível do primeiro item selecionado para definir o modo
            primeiro = selecoes[0]
            nivel = self.obter_nivel_item(primeiro)
            if nivel == "ano":
                modo = "ano"
            elif nivel == "mes":
                modo = "mes"
            else:
                modo = "dia"
        
        for item in selecoes:
            nivel = self.obter_nivel_item(item)
            if modo == "ano" and nivel == "ano":
                filhos = self.treeview_datas.get_children(item)
                for mes in filhos:
                    dias = self.treeview_datas.get_children(mes)
                    for dia in dias:
                        selecionados.append(self.treeview_datas.item(dia, 'text'))
            elif modo == "mes" and nivel == "mes":
                dias = self.treeview_datas.get_children(item)
                for dia in dias:
                    selecionados.append(self.treeview_datas.item(dia, 'text'))
            elif modo == "dia" and nivel == "dia":
                selecionados.append(self.treeview_datas.item(item, 'text'))
        return selecionados, modo
    
    def obter_nivel_item(self, item_id):
        parent = self.treeview_datas.parent(item_id)
        if parent == "":
            return "raiz"
        elif self.treeview_datas.parent(parent) == "":
            return "ano"
        elif self.treeview_datas.parent(self.treeview_datas.parent(parent)) == "":
            return "mes"
        else:
            return "dia"

    def obter_selecao_agrupada(self):
        """
        Retorna uma tupla: (lista de datas selecionadas, modo_agrupamento)
        onde modo_agrupamento pode ser:
        - "raiz": quando o nó "📅 Datas Agrupadas" for selecionado (exibe média por ano),
        - "ano": quando um nó de ano (ex: "2024") for selecionado (exibe média por mês),
        - "dia": quando um nó de dia for selecionado (exibe dados diários).
        """
        selecionados = []
        selecoes = self.treeview_datas.selection()
        if not selecoes:
            return selecionados, "dia"
        
        # Verifica o nível do primeiro item selecionado
        primeiro = selecoes[0]
        nivel = self.obter_nivel_item(primeiro)
        if nivel == "raiz":
            modo = "raiz"
            # Para o nó raiz, percorre todos os anos, meses e dias
            for ano in self.treeview_datas.get_children(primeiro):
                for mes in self.treeview_datas.get_children(ano):
                    for dia in self.treeview_datas.get_children(mes):
                        selecionados.append(self.treeview_datas.item(dia, 'text'))
        elif nivel == "ano":
            modo = "ano"
            for item in selecoes:
                # Se o item for um ano, adicione todos os dias dos seus filhos (meses)
                for mes in self.treeview_datas.get_children(item):
                    for dia in self.treeview_datas.get_children(mes):
                        selecionados.append(self.treeview_datas.item(dia, 'text'))
        else:
            # Se for "mes" ou "dia", assume-se dados diários
            modo = "dia"
            for item in selecoes:
                # Se o nó selecionado for de nível "mes", percorre seus dias
                if self.obter_nivel_item(item) == "mes":
                    for dia in self.treeview_datas.get_children(item):
                        selecionados.append(self.treeview_datas.item(dia, 'text'))
                else:
                    selecionados.append(self.treeview_datas.item(item, 'text'))
        return selecionados, modo
    
    def buscar_dados_dolar(self):
        # Sai cedo se já estivermos encerrando ou se o widget não existir
        if (self.encerrando
            or not getattr(self, "treeview_datas", None)
            or not self.treeview_datas.winfo_exists()):
            return [], "dia"

        conn = conectar()
        if conn is None:
            return [], "dia"

        try:
            cursor = conn.cursor()
            datas, modo = self.obter_selecao_agrupada()
            if datas:
                datas_formatadas = [datetime.strptime(d, "%d/%m/%Y").date() for d in datas]
                datas_str = ", ".join(f"'{d}'" for d in datas_formatadas)
                consulta = f"""
                    SELECT data, dolar
                    FROM cotacao_dolar
                    WHERE data IN ({datas_str})
                    ORDER BY data
                """
            else:
                modo = "dia"
                consulta = """
                    SELECT data, dolar
                    FROM cotacao_dolar
                    ORDER BY data
                """

            cursor.execute(consulta)
            dados = [(d.strftime("%d/%m/%Y"), v) for d, v in cursor.fetchall()]
            cursor.close()
            conn.close()
            return dados, modo

        except tk.TclError:
            # Widget destruído durante a busca: sai silenciosamente
            return [], "dia"
        except Exception as e:
            # Outros erros: notifica o usuário
            messagebox.showerror("Erro ao buscar dados", str(e))
            return [], "dia"

    def iniciar_atualizacao_dolar(self):
        if self.encerrando:
            return

        thread = threading.Thread(target=self.thread_buscar_dolar, daemon=True)
        thread.start()

        # dispara a primeira verificação só via after
        if not self.encerrando:
            self.id_atualizacao_dolar = self.after(1000, self.verificar_fila)

    def thread_buscar_dolar(self):
        # Busca os dados e coloca na fila para a thread principal atualizar a interface
        dados = self.buscar_dados_dolar()
        if dados:
            self.fila_dolar.put(dados)

    def _agendar_verificacao(self):
        if not self.encerrando:
            self.id_verificar = self.after(1000, self.verificar_fila)

    def verificar_fila(self):
        if self.encerrando or not self.winfo_exists():
            return

        try:
            dados_com_modo = self.fila_dolar.get_nowait()
            self.atualizar_grafico_dolar(dados_com_modo)
        except queue.Empty:
            pass

        if not self.encerrando:
            self._agendar_verificacao()

        else:
            print("Finalizando AplicacaoGraficoDolar")
            self._finalizar()

    def atualizar_grafico_dolar(self, dados_com_modo):
         # → checa flag de encerramento e existência da janela
        if (self.encerrando or not self.winfo_exists()
            # → checa se self.canvas_dolar existe
            or not getattr(self, "canvas_dolar", None)
            # → pega o widget Tk real e checa se ele existe
            or not self.canvas_dolar.get_tk_widget().winfo_exists()):
            return
        # Desempacota os dados e o modo
        dados, modo = dados_com_modo
        self.modo_atual = modo

        self.eixo_dolar.clear()
        self.eixo_dolar.set_facecolor("white")
        self.fig_dolar.patch.set_facecolor("white")
        self.eixo_dolar.grid(True, color="#e6e6e6", linestyle="--", linewidth=0.8)

        datas_dt, valores = [], []
        for linha in dados:
            try:
                # Converte a data para datetime e o valor para float
                dt = datetime.strptime(linha[0].strip(), "%d/%m/%Y")
                valor = float(linha[1])
                datas_dt.append(dt)
                valores.append(valor)
            except Exception:
                continue

        if not datas_dt or not valores:
            self.eixo_dolar.text(0.5, 0.5, "Nenhum dado encontrado",
                                horizontalalignment="center",
                                verticalalignment="center",
                                fontsize=12, fontweight="bold", color="#666666")
            self.canvas_dolar.draw()
            self.id_atualizacao_dolar = self.aba_dolar.after(5000, self.iniciar_atualizacao_dolar)
            return

        # Processa os dados de acordo com o modo de agrupamento:
        if modo == "raiz":
            agrupado = {}
            for dt, val in zip(datas_dt, valores):
                chave = dt.strftime("%Y")
                agrupado.setdefault(chave, []).append(val)
            chaves_ordenadas = sorted(agrupado.keys())
            valores_estat = [sum(agrupado[k]) / len(agrupado[k]) for k in chaves_ordenadas]
            x = list(range(len(chaves_ordenadas)))
            labels = chaves_ordenadas
            self.datas_str_sorted = chaves_ordenadas
        elif modo == "ano":
            agrupado = {}
            for dt, val in zip(datas_dt, valores):
                chave = dt.strftime("%Y-%m")
                agrupado.setdefault(chave, []).append(val)
            chaves_ordenadas = sorted(agrupado.keys())
            valores_estat = [sum(agrupado[k]) / len(agrupado[k]) for k in chaves_ordenadas]
            x = list(range(len(chaves_ordenadas)))
            labels = [datetime.strptime(k, "%Y-%m").strftime("%b %Y").capitalize() for k in chaves_ordenadas]
            self.datas_str_sorted = labels
        else:  # "dia"
            combinados = sorted(zip(datas_dt, valores), key=lambda item: item[0])
            datas_dt_sorted, valores_sorted = zip(*combinados)
            labels = [dt.strftime("%d/%m/%Y") for dt in datas_dt_sorted]
            x = list(range(len(labels)))
            valores_estat = valores_sorted
            self.datas_str_sorted = labels

        # Plotagem do gráfico
        self.eixo_dolar.plot(x, valores_estat, marker="o", linestyle="-", color="#0072B2")
        max_labels = 15
        step = max(1, len(labels) // max_labels)
        indices_labels = list(range(0, len(labels), step))
        if indices_labels[-1] != len(labels) - 1:
            indices_labels.append(len(labels) - 1)
        self.eixo_dolar.set_xticks([x[i] for i in indices_labels])
        self.eixo_dolar.set_xticklabels([labels[i] for i in indices_labels], rotation=45, ha="right", fontsize=9)

        # Título e rótulos dos eixos
        self.eixo_dolar.set_title("Cotação do Dólar", fontsize=14, fontweight="bold")
        self.eixo_dolar.set_xlabel("Data", fontsize=12)
        self.eixo_dolar.set_ylabel("Dólar", fontsize=12)
        self.eixo_dolar.yaxis.set_major_locator(MultipleLocator(0.5))
        
        # Função formatadora para o eixo y – exibe R$ com ponto para milhares e vírgula para decimais
        def pt_br_formatter(x, pos):
            s = f"{x:,.2f}"   # Exemplo: "1,234.56"
            s = s.replace(",", "X").replace(".", ",").replace("X", ".")
            return f"R$ {s}"
        
        self.eixo_dolar.yaxis.set_major_formatter(FuncFormatter(pt_br_formatter))

        self.fig_dolar.tight_layout()
        self.canvas_dolar.draw()

        self.tooltip = None
        self.tooltip_fixed = False

        self.debounced_atualizar_sparkline_e_tendencia(valores_estat)
        self.atualizar_estatisticas_dolar_treeview(valores_estat)
        self.aba_dolar.after(5000, self.iniciar_atualizacao_dolar)

    def atualizar_estatisticas_dolar_treeview(self, valores):
        if (self.encerrando
            or not valores
            or not self.winfo_exists()
            or not self.executando
            or not getattr(self, "treeview_estatisticas", None)
            or not self.treeview_estatisticas.winfo_exists()):
            return
        try:
            if self.treeview_estatisticas and self.treeview_estatisticas.winfo_exists():
                for i in self.treeview_estatisticas.get_children():
                    self.treeview_estatisticas.delete(i)

                media = sum(valores) / len(valores)
                maximo = max(valores)
                minimo = min(valores)

                estatisticas = [
                    ("Máx", f"R$ {maximo:.2f}".replace(".", ",")),
                    ("Mín", f"R$ {minimo:.2f}".replace(".", ",")),
                    ("Média", f"R$ {media:.2f}".replace(".", ","))
                ]

                for nome, valor in estatisticas:
                    self.treeview_estatisticas.insert("", "end", values=(nome, valor))
        except tk.TclError:
            pass

    def _on_destroy(self, event):
        # marca flag e desconecta callback Matplotlib
        self.canvas_destroyed = True
        if hasattr(self, 'cid'):
            self.canvas_dolar.mpl_disconnect(self.cid)
            del self.cid

    def mostrar_tooltip(self, event):
        try:
            # se o canvas já tiver sido destruído, desconecta e sai
            tk_widget = self.canvas_dolar.get_tk_widget()
            if self.canvas_destroyed or not tk_widget.winfo_exists():
                if hasattr(self, 'cid'):
                    self.canvas_dolar.mpl_disconnect(self.cid)
                    del self.cid
                return

            # coordenadas inválidas?
            if event.xdata is None or event.ydata is None:
                return

            # clique esquerdo fixa/desfixa
            if event.button == 1:
                if self.tooltip and not self.tooltip_fixed:
                    self.tooltip_fixed = True
                    return
                elif self.tooltip and self.tooltip_fixed:
                    self.tooltip.set_visible(False)
                    self.canvas_dolar.draw_idle()
                    self.tooltip = None
                    self.tooltip_fixed = False
                    return

            if self.tooltip_fixed:
                return  # não atualiza enquanto fixada

            # pega linha única
            linha = self.eixo_dolar.lines[0] if self.eixo_dolar.lines else None
            if not linha:
                return

            # encontra ponto próximo
            ponto_mais_proximo = None
            menor_dist = float('inf')
            for x_val, y_val in linha.get_xydata():
                d = ((event.xdata - x_val)**2 + (event.ydata - y_val)**2)**0.5
                if d < menor_dist and d < 2:
                    menor_dist = d
                    ponto_mais_proximo = (x_val, y_val)

            if ponto_mais_proximo is None:
                if self.tooltip:
                    self.tooltip.set_visible(False)
                    self.tooltip = None
                    self.canvas_dolar.draw_idle()
                return

            x_val, y_val = ponto_mais_proximo

            try:
                data_str = self.datas_str_sorted[int(x_val)]
            except (IndexError, TypeError):
                data_str = "Desconhecida"

            valor_str = f"{y_val:.2f}".replace('.', ',')
            texto = f"Data: {data_str}\nDólar: R$ {valor_str}"

            # calculo de deslocamento
            inv = self.eixo_dolar.transData.transform
            x_canvas, _ = inv((x_val, y_val))
            largura, _ = self.canvas_dolar.get_width_height()
            if x_canvas > largura - 100:
                desloc = (-100, 10); ha = 'right'
            else:
                desloc = (10, 10);  ha = 'left'

            # cria ou atualiza annotation
            if self.tooltip is None:
                self.tooltip = self.eixo_dolar.annotate(
                    texto,
                    xy=(x_val, y_val),
                    xytext=desloc,
                    textcoords='offset points',
                    bbox=dict(boxstyle='round,pad=0.3', fc='yellow', alpha=0.9),
                    arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'),
                    fontsize=10,
                    ha=ha,
                    zorder=100  
                )
                self.tooltip.set_clip_on(False)  # opcional
            else:
                self.tooltip.xy = (x_val, y_val)
                self.tooltip.set_text(texto)
                self.tooltip.set_position(desloc)
                self.tooltip.set_ha(ha)
                self.tooltip.set_visible(True)

            # só aqui chamamos draw_idle uma vez
            self.canvas_dolar.draw_idle()

        except TclError:
            # widget já não existe: ignora todo o resto
            return

    def esconder_tooltip(self, event):
        try:
            if self.tooltip and not self.tooltip_fixed:
                self.tooltip.set_visible(False)
                self.canvas_dolar.draw_idle()
                self.tooltip = None
        except TclError:
            # ignora se o canvas tiver morrido
            pass

    def exportar_excel_grafico(self):
        try:
            if not hasattr(self, "fig_dolar") or self.fig_dolar is None:
                raise ValueError("O gráfico de dólar não foi carregado corretamente.")

            ax_original = self.fig_dolar.axes[0]

            # Criar uma nova figura para exportação com altura reduzida (6.5 x 3.5 polegadas)
            fig_export = Figure(figsize=(6.5, 3.5), dpi=150)
            ax_export = fig_export.add_subplot(111)

            # Copiar as linhas do gráfico original
            for line in ax_original.get_lines():
                ax_export.plot(
                    line.get_xdata(),
                    line.get_ydata(),
                    label=line.get_label(),
                    color=line.get_color(),
                    linestyle=line.get_linestyle(),
                    marker=line.get_marker()
                )

            # Copiar título, rótulos dos eixos e grid
            ax_export.set_title(ax_original.get_title())
            ax_export.set_xlabel(ax_original.get_xlabel())
            ax_export.set_ylabel(ax_original.get_ylabel())
            ax_export.grid(True)

            # Copiar os xticks e os rótulos (as datas) do eixo x
            xticks = ax_original.get_xticks()
            xticklabels = [label.get_text() for label in ax_original.get_xticklabels()]
            ax_export.set_xticks(xticks)
            ax_export.set_xticklabels(xticklabels, rotation=45, ha='right')

            # Incluir a legenda, se existir
            if ax_original.get_legend():
                handles, labels = ax_original.get_legend_handles_labels()
                ax_export.legend(handles, labels, loc="upper right", fontsize=8)

            # Configurar o eixo y com o mesmo formato (ponto para milhares e vírgula para decimais)
            ax_export.yaxis.set_major_locator(MultipleLocator(0.5))
            def pt_br_formatter(x, pos):
                s = f"{x:,.2f}"  # Ex: "1,234.56"
                s = s.replace(",", "X").replace(".", ",").replace("X", ".")
                return f"R$ {s}"
            ax_export.yaxis.set_major_formatter(FuncFormatter(pt_br_formatter))

            fig_export.tight_layout()

            # Gerar a figura (usando um backend sem GUI) e salvá-la em um arquivo temporário
            canvas_export = FigureCanvasAgg(fig_export)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                temp_file = tmp.name

            fig_export.savefig(temp_file, bbox_inches="tight", dpi=150)

            # Solicitar o local e nome do arquivo Excel para salvar a imagem
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Planilhas do Excel", "*.xlsx")],
                title="Salvar Gráfico de Dólar em Excel",
                initialfile="Grafico_Cotacao_Dolar.xlsx"
            )
            if not caminho_arquivo:
                return

            # Criar um arquivo Excel e inserir a imagem exportada
            wb = Workbook()
            ws = wb.active
            ws.title = "Gráfico de Dólar"
            img = XLImage(temp_file)
            ws.add_image(img, 'A1')

            wb.save(caminho_arquivo)
            messagebox.showinfo("Exportação", f"Gráfico de dólar exportado com sucesso para:\n{caminho_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro na Exportação", f"Erro ao exportar o gráfico de dólar:\n{str(e)}")

    def exportar_grafico_pdf(self):
        """Exporta o gráfico de dólar em um arquivo PDF, incluindo uma legenda temporária."""
        # Abre diálogo para o usuário escolher onde salvar o PDF
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile="grafico_dolar.pdf",
            filetypes=[("Arquivo PDF", "*.pdf")],
            title="Salvar Gráfico de Dólar em PDF"
        )

        if not caminho_arquivo:
            return

        try:
            # Obtém o eixo do gráfico (primeiro eixo da figura)
            ax = self.fig_dolar.get_axes()[0]

            # Salva a legenda original (se houver) para restaurá-la depois
            legenda_original = ax.get_legend()

            # Cria uma legenda temporária caso haja linhas ploteadas
            lines = ax.get_lines()
            if lines:
                legenda_temp = ax.legend(
                    lines,
                    ["Cotação do Dólar"],
                    title="Legenda",
                    loc="upper center",
                    bbox_to_anchor=(0.5, -0.1),
                    ncol=1
                )
            else:
                legenda_temp = None

            # Salva a figura em PDF
            self.fig_dolar.savefig(caminho_arquivo, format="pdf", bbox_inches="tight")

            # Remove a legenda temporária para que a interface não seja afetada
            if legenda_temp is not None:
                legenda_temp.remove()
                self.canvas_dolar.draw()

            # Se havia uma legenda original, restaura-a
            if legenda_original is not None:
                ax.legend(
                    handles=legenda_original.legendHandles,
                    labels=[t.get_text() for t in legenda_original.get_texts()]
                )

            messagebox.showinfo("Exportar PDF", f"Gráfico exportado com sucesso para:\n{caminho_arquivo}")

        except Exception as e:
            messagebox.showerror("Exportar PDF", f"Erro ao exportar o gráfico: {e}")

    def on_closing(self):
        print("Fechando gráfico de dólar")
        self.encerrando = True

        # lista de todos os IDs de after que você usa
        for job_attr in (
            "id_verificar",
            "id_atualizacao_dolar",
            "job_atualizar_periodos",
            "job_atualizar_datas",
            "sparkline_debounce_id",
            "treeview_debounce_id"
        ):
            job = getattr(self, job_attr, None)
            if job:
                try:
                    self.after_cancel(job)
                except tk.TclError:
                    pass

        # Dê um tempinho para a fila ou thread finalizar, e destrói de fato
        self.after(100, self._finalizar)

    def _finalizar(self):
        try:
            if self.encerrando:
                self.destroy()
            else:
                self.after(100, self._finalizar)
        except tk.TclError:
            pass
