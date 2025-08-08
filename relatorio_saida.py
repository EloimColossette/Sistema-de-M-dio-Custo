import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar
from conexao_db import conectar  # Importa a função para conectar ao banco
from relatorio_entrada import RelatorioEntradaApp
from relatorio_resumo import RelatorioResumoApp
from logos import aplicar_icone
import pandas as pd
from tkinter import filedialog
import xlsxwriter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

class RelatorioApp(tk.Toplevel):  # Herda de Toplevel
    def __init__(self, root):
        super().__init__(root)  # Associa o Toplevel à janela principal (menu)
        self.title("Relatórios de Produtos")
        self.state("zoomed")

        aplicar_icone(self, r"C:\Sistema\logos\Kametal.ico")

        # Configuração do estilo do ttk com uma paleta mais profissional
        self.configurar_estilos()

        # Criar um frame para os botões (Voltar e Exportar)
        self.frame = tk.Frame(self)
        self.frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Criar o botão Voltar
        self.botao_voltar = ttk.Button(self.frame, text="Voltar", command=self.voltar, style="Voltar.TButton")
        self.botao_voltar.pack(side=tk.LEFT, padx=10, pady=10)

        # Botão Exportar para PDF
        self.botao_exportar_pdf = ttk.Button(
            self.frame,
            text="Exportar para PDF",
            command=self.exportar_para_pdf,
            style="ExportarPDF.TButton"
        )
        self.botao_exportar_pdf.pack(side=tk.RIGHT, padx=10, pady=10)

        # Botão Exportar para Excel
        self.botao_exportar = ttk.Button(
            self.frame, text="Exportar para Excel", 
            command=self.exportar_para_excel, style="ExportarExcel.TButton"
        )
        self.botao_exportar.pack(side=tk.RIGHT, padx=10, pady=10)

       # Criar abas (Notebook)
        self.abas = ttk.Notebook(self)
        self.abas.pack(fill=tk.BOTH, expand=True)

        # Aba Relatório de Saída
        self.aba_saida = ttk.Frame(self.abas)
        self.abas.add(self.aba_saida, text="Relatório de Saída")

        # Aba Relatório de Entrada
        self.aba_entrada = ttk.Frame(self.abas)
        self.abas.add(self.aba_entrada, text="Relatório de Entrada")

        # Aba Relatório Resumo
        self.aba_resumo = ttk.Frame(self.abas)
        self.abas.add(self.aba_resumo, text="Relatório Resumo")

        # Instancia a interface do relatório resumo na aba_resumo
        self.relatorioResumo = RelatorioResumoApp(self.aba_resumo)
        self.relatorioResumo.pack(expand=True, fill="both")

        # Agora, instancie a interface do relatório de entrada, passando o relatório resumo
        self.relatorioEntrada = RelatorioEntradaApp(self, self.aba_entrada, self.relatorioResumo)

        # Cria a interface para a aba de saída
        self.criar_interface(self.aba_saida, tipo="saida")

        self.carregar_produtos_base()

        # Configurar o evento de fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def configurar_estilos(self):
        """Define estilos globais para o Treeview e os botões."""
        estilo = ttk.Style()
        estilo.theme_use("alt")

        # Estilo básico para linhas da tabela
        estilo.configure("Treeview", 
                         font=("Arial", 10), 
                         rowheight=35,  # Aumenta a altura das linhas
                         background="white", 
                         foreground="black", 
                         fieldbackground="white")

        # Estilo do cabeçalho
        estilo.configure("Treeview.Heading", 
                         font=("Arial", 10, "bold"))

        # Estilo dos botões
        estilo.configure("Voltar.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#FF6347",  # Cor laranja para Voltar
                         foreground="white",
                         font=("Arial", 10, "bold"),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("Voltar.TButton", 
                   background=[("active", "#FF4500")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])

        estilo.configure("ExportarExcel.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#1E90FF",  # Cor azul para Exportar
                         foreground="white",
                         font=("Arial", 10, "bold"),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("ExportarExcel.TButton", 
                   background=[("active", "#1C86EE")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Estilo para linha de total
        estilo.configure("total_arame.Treeview", 
                         background="#FF6347",  # Cor laranja para arame
                         foreground="white", 
                         font=("Arial", 12, "bold"))
        
        estilo.configure("total_fio.Treeview", 
                         background="#1E90FF",  # Cor azul para fio
                         foreground="white", 
                         font=("Arial", 12, "bold"))
        
        estilo.configure("total_geral.Treeview", 
                         background="#FFD700",  # Cor dourada para o total geral
                         foreground="black", 
                         font=("Arial", 12, "bold"))

        # Verificar se o estilo foi aplicado corretamente
        print("Estilo configurado!")

    def criar_interface(self, parent, tipo):
        """Cria os filtros e a tabela para cada aba."""
        frame_filtros = ttk.LabelFrame(parent, text="Filtros", padding=(10, 10))
        frame_filtros.pack(fill=tk.X, padx=20, pady=10)

        label_style = {"font": ("Arial", 10, "bold")}

        # Linha 0: Entrada de Data e Combobox
        tk.Label(frame_filtros, text="Mês (MM/YYYY):", **label_style).grid(
            row=0, column=0, padx=(5,2), pady=5, sticky=tk.W
        )
        entrada_data = tk.Entry(frame_filtros, width=12, font=("Arial", 10))
        entrada_data.grid(row=0, column=1, padx=(2,15), pady=5, sticky=tk.W)

        tk.Label(frame_filtros, text="Base do Produto:", **label_style).grid(
            row=0, column=2, padx=(5,2), pady=5, sticky=tk.W
        )
        combobox_produto = ttk.Combobox(frame_filtros, width=30, font=("Arial", 10))
        combobox_produto.grid(row=0, column=3, padx=(2,15), pady=5, sticky=tk.W)

        botao_relatorio = ttk.Button(
            frame_filtros, text="Gerar Relatório",
            command=lambda: self.gerar_relatorio(tipo, entrada_data, combobox_produto),
            style="ExportarExcel.TButton"
        )
        botao_relatorio.grid(row=0, column=4, padx=(10,5), pady=5, sticky=tk.W)

        # Criar um Frame para manter os elementos alinhados
        frame_pesquisa = tk.Frame(frame_filtros)
        frame_pesquisa.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)

        # Rótulo de pesquisa
        tk.Label(frame_pesquisa, text="Pesquisar:", **label_style).pack(side=tk.LEFT, padx=(0, 5))

        # Campo de pesquisa
        pesquisa_entry = tk.Entry(frame_pesquisa, width=20, font=("Arial", 10))
        pesquisa_entry.pack(side=tk.LEFT)

        # Botão de pesquisa
        botao_pesquisar = ttk.Button(
            frame_pesquisa, text="Pesquisar",
            command=lambda: self.pesquisar(pesquisa_entry.get(), tabela),
            style="Voltar.TButton"
        )
        botao_pesquisar.pack(side=tk.LEFT, padx=5)

        # Tabela de Resultados
        frame_tabela = tk.Frame(parent)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        colunas = ("Data", "NF", "Produto", "Peso")
        tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

        for coluna, largura in zip(colunas, [100, 100, 300, 150]):
            tabela.heading(coluna, text=coluna)
            tabela.column(coluna, anchor=tk.CENTER, width=largura)

        scroll_y = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=tabela.yview)
        tabela.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tabela.pack(fill=tk.BOTH, expand=True)

        # Configura as tags de estilo para os totais
        tabela.tag_configure("total_arame", background="#FF6347", foreground="white", font=("Arial", 12, "bold"))
        tabela.tag_configure("total_fio", background="#1E90FF", foreground="white", font=("Arial", 12, "bold"))
        tabela.tag_configure("total_geral", background="#FFD700", foreground="black", font=("Arial", 12, "bold"))

        if tipo == "saida":
            self.entrada_data_saida = entrada_data
            self.combobox_produto_saida = combobox_produto
            self.tabela_saida = tabela
        else:
            self.entrada_data_entrada = entrada_data
            self.combobox_produto_entrada = combobox_produto
            self.tabela_entrada = tabela

        def formatar_data(event):
            """Adiciona a barra automaticamente ao digitar o mês"""
            texto = entrada_data.get()
            if len(texto) == 2 and '/' not in texto:
                entrada_data.insert(tk.END, '/')

        entrada_data.bind("<KeyRelease>", formatar_data)

    def pesquisar(self, texto_pesquisa, tabela):
        """Filtra a tabela conforme o texto inserido na barra de pesquisa e reinsere os totais conforme o termo pesquisado.
        
        - Se a pesquisa estiver em branco, exibe todos os totais (arame, fio e total geral, desde que tenham valor > 0).
        - Se o termo for 'arame' (e não 'fio'), exibe apenas o total de arame e o total geral.
        - Se o termo for 'fio' (e não 'arame'), exibe apenas o total de fio e o total geral.
        - Caso contrário, exibe somente o total geral.
        """
        texto_pesquisa = texto_pesquisa.lower().strip()
        
        if not hasattr(self, 'resultados_saida'):
            messagebox.showinfo("Informação", "Gere o relatório antes de pesquisar.")
            return
        
        # Limpa a Treeview
        for item in tabela.get_children():
            tabela.delete(item)
        
        # Filtra os registros de detalhes de acordo com a pesquisa
        if texto_pesquisa:
            resultados_filtrados = [
                reg for reg in self.resultados_saida if texto_pesquisa in " ".join(map(str, reg)).lower()
            ]
        else:
            resultados_filtrados = self.resultados_saida
        
        # Insere os registros filtrados (detalhes)
        for reg in resultados_filtrados:
            tabela.insert("", tk.END, values=(
                reg[0].strftime("%d/%m/%Y"), reg[1], reg[2], self.formatar_peso(reg[3])
            ))
        
        # Seleciona os totais a serem exibidos
        totais_para_exibir = []
        if texto_pesquisa == "":
            # Campo de pesquisa vazio: exibe todos os totais armazenados (eles já devem ter sido filtrados para não incluir arame/fio com 0)
            totais_para_exibir = self.totais_saida[:]
        else:
            # Se houver termo na pesquisa:
            # Se pesquisar "arame" e não "fio"
            if "arame" in texto_pesquisa and "fio" not in texto_pesquisa:
                for total, tag in self.totais_saida:
                    if tag in ("total_arame", "total_geral"):
                        totais_para_exibir.append((total, tag))
            # Se pesquisar "fio" e não "arame"
            elif "fio" in texto_pesquisa and "arame" not in texto_pesquisa:
                for total, tag in self.totais_saida:
                    if tag in ("total_fio", "total_geral"):
                        totais_para_exibir.append((total, tag))
            # Se pesquisar ambos ou outro termo: exibe somente o total geral
            else:
                for total, tag in self.totais_saida:
                    if tag == "total_geral":
                        totais_para_exibir.append((total, tag))
        
        # Insere os totais, se houver
        if totais_para_exibir:
            tabela.insert("", tk.END, values=("", "", "", ""))  # Linha em branco para separar os detalhes dos totais
            for total, tag in totais_para_exibir:
                tabela.insert("", tk.END, values=total, tags=(tag,))
                tabela.insert("", tk.END, values=("", "", "", ""))  # Linha em branco entre os totais (opcional)

    def voltar(self):
        """Fecha a janela atual e reexibe o menu principal com atualização visual."""
        # Cancela callback agendado, se houver
        if getattr(self, "_encerrar_id", None) is not None:
            try:
                self.after_cancel(self._encerrar_id)
            except Exception:
                pass

        # Fecha a janela atual antes de mostrar o menu
        self.destroy()

        # Reexibe o menu principal
        self.master.deiconify()
        self.master.state("zoomed")
        self.master.lift()
        self.master.update()

    def carregar_produtos_base(self):
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT base_produto FROM produtos_nf ORDER BY base_produto;")
            produtos = [row[0] for row in cur.fetchall()]
            # Atualiza a combobox da aba de saída
            self.combobox_produto_saida['values'] = produtos
            if produtos:
                self.combobox_produto_saida.current(0)
            cur.close()
            conn.close()
            
            # Atualiza a combobox da aba de entrada chamando o método da instância de RelatorioEntradaApp
            if hasattr(self, 'relatorioEntrada'):
                self.relatorioEntrada.carregar_produtos_base()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar produtos base: {e}")

    def gerar_relatorio(self, tipo, entrada_data, combobox_produto):
        """Gera o relatório de entrada ou saída, baseado no tipo informado."""
        data_inserida = entrada_data.get()
        produto_selecionado = combobox_produto.get()

        try:
            mes, ano = data_inserida.split('/')
            ano, mes = int(ano), int(mes)
            data_inicio = f"{ano}-{mes:02d}-01"
            ultimo_dia = calendar.monthrange(ano, mes)[1]
            data_fim = f"{ano}-{mes:02d}-{ultimo_dia}"
        except Exception as e:
            messagebox.showerror("Erro", f"Formato de data inválido. Use 'MM/YYYY': {e}")
            return

        # Consulta SQL dinâmica baseada no tipo
        query = """
            SELECT data, numero_nf, produto_nome, peso
            FROM {}
            WHERE data BETWEEN %s AND %s AND base_produto = %s
            ORDER BY data ASC, numero_nf ASC;
        """.format("nf JOIN produtos_nf ON nf.id = produtos_nf.nf_id" if tipo == "saida"
                else "entrada JOIN produtos_entrada ON entrada.id = produtos_entrada.entrada_id")

        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute(query, (data_inicio, data_fim, produto_selecionado))
            resultados = cur.fetchall()

            self.resultados_saida = resultados  # Armazena os dados completos


            # Limpa a tabela antes de inserir novos dados
            tabela = self.tabela_saida if tipo == "saida" else self.tabela_entrada
            for item in tabela.get_children():
                tabela.delete(item)

            # Se não houver registros, exibe mensagem e sai da função
            if not resultados:
                messagebox.showinfo("Sem Registros", f"Não há registro de {tipo} para '{produto_selecionado}' no mês {data_inserida}.")
                return

            # Após calcular os totais
            total_arame = 0
            total_fio = 0
            total_geral = 0

            # Insere os registros na tabela e acumula os totais
            for reg in resultados:
                produto_nome = reg[2].lower()
                peso = reg[3]

                if "arame" in produto_nome:
                    total_arame += peso
                elif "fio" in produto_nome:
                    total_fio += peso
                total_geral += peso

                tabela.insert("", tk.END, values=(
                    reg[0].strftime("%d/%m/%Y"), reg[1], reg[2], self.formatar_peso(peso)
                ))

            # Monta a lista de totais apenas com os que possuem valor maior que 0
            totais = []
            if total_arame > 0:
                totais.append((("", "TOTAL ARAME", self.formatar_peso(total_arame), ""), "total_arame"))
            if total_fio > 0:
                totais.append((("", "TOTAL FIO", self.formatar_peso(total_fio), ""), "total_fio"))
            if total_geral > 0:
                totais.append((("", "TOTAL GERAL", self.formatar_peso(total_geral), ""), "total_geral"))

            # Armazena os totais para uso na pesquisa
            self.totais_saida = totais

            # Insere uma linha em branco após os registros, se houver totais para exibir
            if totais:
                tabela.insert("", tk.END, values=("", "", "", ""))

            # Insere os totais na tabela, separando cada um por uma linha em branco
            for total, tag in totais:
                tabela.insert("", tk.END, values=total, tags=(tag,))
                tabela.insert("", tk.END, values=("", "", "", ""))

            cur.close()
            conn.close()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar a consulta: {e}")

    def formatar_peso(self, valor):
        return f"{valor:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " Kg"

    def exportar_para_excel(self):
        """Exporta os relatórios de saída, entrada e resumo para um arquivo Excel com abas separadas."""
        try:
            nome_padrao = "Relatorio_item_grupo.xlsx"
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=nome_padrao,
                filetypes=[("Arquivo Excel", "*.xlsx")],
                title="Salvar Relatório"
            )

            if not caminho_arquivo:
                return  # Usuário cancelou a ação

            workbook = xlsxwriter.Workbook(caminho_arquivo)

            # Aba "Relatório de Saída" (sem coluna Produto)
            if self.tabela_saida.get_children():
                aba_saida = workbook.add_worksheet("Relatório de Saída")
                self.exportar_tabela_para_aba(self.tabela_saida, aba_saida, incluir_produto=False)

            # Aba "Relatório de Entrada" (sem coluna Produto)
            if hasattr(self.relatorioEntrada, "tabela_entrada") and self.relatorioEntrada.tabela_entrada.get_children():
                aba_entrada = workbook.add_worksheet("Relatório de Entrada")
                self.exportar_tabela_para_aba(self.relatorioEntrada.tabela_entrada, aba_entrada, incluir_produto=False)

            # Aba "Resumo" (com coluna Produto)
            if self.relatorioResumo.tree.get_children():
                aba_resumo = workbook.add_worksheet("Resumo")
                self.exportar_tabela_para_aba(self.relatorioResumo.tree, aba_resumo, incluir_produto=True)

            workbook.close()
            messagebox.showinfo("Sucesso", f"Relatório salvo como:\n{caminho_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

    def exportar_tabela_para_aba(self, tabela, aba, incluir_produto=False):
        """
        Exporta os dados de uma Treeview para uma aba do Excel.
        
        :param tabela: Treeview a ser exportada.
        :param aba: Aba do Excel onde os dados serão escritos.
        :param incluir_produto: Se True, adiciona a coluna do nome do produto (#0), senão remove.
        """
        colunas = tabela["columns"]

        if incluir_produto:
            cabecalhos = ["Produto"] + [tabela.heading(col)["text"] for col in colunas]
        else:
            cabecalhos = [tabela.heading(col)["text"] for col in colunas]

        # Escreve os cabeçalhos no Excel
        for col_idx, cabecalho in enumerate(cabecalhos):
            aba.write(0, col_idx, cabecalho)

        # Escreve os dados das tabelas
        for row_idx, item in enumerate(tabela.get_children(), start=1):
            valores = tabela.item(item, "values")
            nome_produto = tabela.item(item, "text")  # Nome do produto na coluna #0

            if incluir_produto:
                valores = (nome_produto,) + valores  # Adiciona o nome do produto
            # Caso contrário, mantém apenas os valores das colunas sem o nome do produto

            for col_idx, valor in enumerate(valores):
                aba.write(row_idx, col_idx, valor)

    def exportar_para_pdf(self):
        aba_atual = self.abas.tab(self.abas.select(), "text")
        # Seleciona a Treeview e configurações conforme aba
        if aba_atual == "Relatório de Saída":
            tree = getattr(self, 'tabela_saida', None)
            nome_padrao = "Relatorio_saida.pdf"
            titulo_pdf = "Relatório de Saída"
        elif aba_atual == "Relatório de Entrada":
            tree = self.relatorioEntrada.tabela_entrada
            nome_padrao = "Relatorio_entrada.pdf"
            titulo_pdf = "Relatório de Entrada"
        elif aba_atual == "Relatório Resumo":
            tree = self.relatorioResumo.tree
            nome_padrao = "Relatorio_resumo.pdf"
            titulo_pdf = "Relatório Resumo"
        else:
            messagebox.showerror("Erro", "Esta aba não possui dados para exportar.")
            return

        # Verifica se há dados
        if not tree or not tree.get_children():
            messagebox.showerror("Erro", "Não há dados para exportar nesta aba.")
            return

        # Diálogo de salvamento com nome padrão
        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF","*.pdf")],
            initialfile=nome_padrao,
            title="Salvar como"
        )
        if not caminho:
            return

        try:
            estilos = getSampleStyleSheet()
            doc = SimpleDocTemplate(caminho, pagesize=A4)
            elementos = []
            # Título
            elementos.append(Paragraph(f"<b>{titulo_pdf}</b>", estilos["Title"]))
            elementos.append(Spacer(1, 20))

            # Monta dados para o PDF
            if aba_atual == "Relatório Resumo":
                # Inclui a coluna de árvore (#0)
                cabecalhos = [tree.heading("#0")["text"]] + [tree.heading(col)["text"] for col in tree["columns"]]
                dados = [cabecalhos]
                for iid in tree.get_children():
                    texto = tree.item(iid)["text"]
                    valores = tree.item(iid)["values"]
                    dados.append([texto] + [str(v) for v in valores])
            else:
                # Abas Saída e Entrada
                cabecalhos = [tree.heading(col)["text"] for col in tree["columns"]]
                dados = [cabecalhos]
                for iid in tree.get_children():
                    dados.append([str(v) for v in tree.item(iid)["values"]])

            tabela_pdf = Table(dados)
            tabela_pdf.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.grey),
                ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("BOTTOMPADDING",(0,0),(-1,0),12),
                ("GRID",(0,0),(-1,-1),0.5,colors.black),
            ]))
            elementos.append(tabela_pdf)
            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"Relatório exportado para: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para PDF: {e}")

    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Fecha a conexão com o banco de dados, se estiver aberta
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
                print("Conexão com o banco de dados fechada.")

            self.destroy()  # Fecha a janela
            