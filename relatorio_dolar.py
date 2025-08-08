import tkinter as tk
from tkinter import ttk
from datetime import datetime, date
from conexao_db import conectar
from tkinter import messagebox, filedialog
import pandas as pd
from openpyxl.styles import Font
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from calendar import monthrange
import re

class RelatorioDolar:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame  # Agora, vamos usar o parent_frame fornecido diretamente

        # Conectar ao banco no início
        self.conexao = conectar()
        if self.conexao:
            print("Conexão bem-sucedida com o banco de dados.")
        else:
            print("Falha na conexão com o banco de dados.")

        self.ultimo_periodo = ""  # Variável para armazenar o último período inserido

        self.configurar_estilos()
        self.criar_interface()
        self.carregar_dados()

        self.tree.bind("<<TreeviewSelect>>", self.on_item_selecionado)

    def criar_interface(self):
        # Frame superior para entradas de Período, Data e Dólar
        self.frame_inputs = ttk.Frame(self.parent_frame, padding=(20, 10))
        self.frame_inputs.grid(row=0, column=0, sticky="ew")
        
        ttk.Label(self.frame_inputs, text="Período (DD/MM/AA a DD/MM/AA):").pack(side="left", padx=5)
        self.entrada_periodo = ttk.Entry(self.frame_inputs, width=25, font=("Helvetica", 12))
        self.entrada_periodo.pack(side="left", padx=5)
        self.entrada_periodo.bind("<KeyRelease>", self.formatar_periodo)
        
        ttk.Label(self.frame_inputs, text="Data (DD/MM/AAAA):").pack(side="left", padx=5)
        self.entrada_data = ttk.Entry(self.frame_inputs, width=15, font=("Helvetica", 12))
        self.entrada_data.pack(side="left", padx=5)
        self.entrada_data.bind("<KeyRelease>", self.formatar_data)
        
        ttk.Label(self.frame_inputs, text="Dólar:").pack(side="left", padx=5)
        self.entrada_dolar = ttk.Entry(self.frame_inputs, width=15, font=("Helvetica", 12))
        self.entrada_dolar.pack(side="left", padx=5)
        # bind usando _formatar_numero com 4 casas
        self.entrada_dolar.bind(
            "<KeyRelease>",
            lambda ev, ent=self.entrada_dolar: self._formatar_numero(ent, casas=4)
        )
        
        # --- Novo: Frame de pesquisa ---
        self.frame_search = ttk.Frame(self.parent_frame, padding=(0, 5))  # Padding à esquerda zerado
        self.frame_search.grid(row=1, column=0, sticky="w")  # Alinha o frame à esquerda
        # Não é necessário configurar a coluna para expandir, já que queremos a barra à esquerda
        ttk.Label(self.frame_search, text="Pesquisar:").grid(row=0, column=0, padx=(0,3), sticky="w")
        self.entrada_pesquisa = ttk.Entry(self.frame_search, width=50, font=("Helvetica", 12))
        self.entrada_pesquisa.grid(row=0, column=1, padx=(0,3), sticky="w")
        self.entrada_pesquisa.bind("<KeyRelease>", self.pesquisar_registros)
        
        # --- Frame para os botões ---
        self.frame_btn = ttk.Frame(self.parent_frame, padding=(20, 10))
        self.frame_btn.grid(row=2, column=0, sticky="ew")
        self.frame_btn.grid_columnconfigure(0, weight=1)
        botoes_frame = ttk.Frame(self.frame_btn)
        botoes_frame.pack(anchor="center")
        self.botao_inserir = ttk.Button(botoes_frame, text="Inserir", command=self.salvar_registro, style="RelatorioCota.TButton")
        self.botao_inserir.pack(side="left", padx=5)
        self.botao_editar = ttk.Button(botoes_frame, text="Editar", command=self.editar_registro, style="RelatorioCota.TButton")
        self.botao_editar.pack(side="left", padx=5)
        self.botao_excluir = ttk.Button(botoes_frame, text="Excluir", command=self.excluir_registro, style="RelatorioCota.TButton")
        self.botao_excluir.pack(side="left", padx=5)
        self.botao_limpar = ttk.Button(botoes_frame, text="Limpar", command=self.limpar_entradas, style="RelatorioCota.TButton")
        self.botao_limpar.pack(side="left", padx=5)
        
        # --- Frame para o Treeview ---
        self.frame_tree = ttk.Frame(self.parent_frame, padding=(20, 10))
        self.frame_tree.grid(row=3, column=0, sticky="nsew")
        self.parent_frame.rowconfigure(3, weight=1)
        self.parent_frame.columnconfigure(0, weight=1)
        
        colunas = ("Data", "Dólar")
        self.tree = ttk.Treeview(self.frame_tree, columns=colunas, show="tree headings")
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        self.tree.heading("#0", text="Período")
        self.tree.column("#0", width=200, anchor="center")
        self.tree.heading("Data", text="Data")
        self.tree.column("Data", width=150, anchor="center")
        self.tree.heading("Dólar", text="Dólar")
        self.tree.column("Dólar", width=150, anchor="center")
        self.tree.config(style="Treeview", height=22)
        
        # Bind para que ao selecionar uma linha os dados apareçam nas entradas
        self.tree.bind("<<TreeviewSelect>>", self.on_item_selecionado)

    def configurar_estilos(self):
        estilo = ttk.Style()
        estilo.theme_use("alt")
        estilo.configure("Treeview", 
                         font=("Arial", 10), 
                         rowheight=27, 
                         background="white", 
                         foreground="black", 
                         fieldbackground="white")
        estilo.configure("Treeview.Heading", 
                         font=("Arial", 10, "bold"))
        estilo.configure("RelatorioCota.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#A9A9A9",
                         foreground="white",
                         font=("Arial", 10, "bold"),
                         borderwidth=2,
                         highlightbackground="#696969",
                         highlightthickness=1)
        estilo.map("RelatorioCota.TButton", 
                   background=[("active", "#808080")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])

    @staticmethod
    def parse_periodo(periodo):
        try:
            # Corrige acentos comuns e espaços extras
            periodo = periodo.replace("á", "a").replace("Á", "A").replace("à", "a").replace("À", "A").strip()

            # Divide o período usando o separador correto
            if " a " in periodo:
                segunda_data_str = periodo.split(" a ")[-1].strip()
            else:
                # Caso esteja colado, tipo "08/07/24a12/07/24"
                segunda_data_str = periodo.split("a")[-1].strip()

            # Converte ano de 2 dígitos para 4 (ex: 24 -> 2024)
            partes_data = segunda_data_str.split("/")
            if len(partes_data[-1]) == 2:
                partes_data[-1] = "20" + partes_data[-1]
                segunda_data_str = "/".join(partes_data)

            # Converte a string para datetime
            dt = datetime.strptime(segunda_data_str, "%d/%m/%Y")
            return dt

        except Exception as e:
            print(f"Erro ao converter período '{periodo}': {e}")
            return datetime.min  # Retorna mínima possível para evitar quebra na ordenação

    def carregar_dados(self):
        """
        Carrega os dados do banco de dados e os exibe no Treeview,
        ordenando os períodos do mais recente para o mais antigo e,
        dentro de cada período, as datas também de forma decrescente.
        """
        try:
            cursor = self.conexao.cursor()

            # Consultar os dados da tabela no banco
            cursor.execute("SELECT periodo, data, dolar FROM cotacao_dolar")
            registros = cursor.fetchall()

            if registros:
                # Organiza os dados por período
                periodos = {}
                for periodo, data, dolar in registros:
                    data_formatada = data.strftime("%d/%m/%Y")
                    dolar_formatado = f"R$ {dolar:.4f}".replace(".", ",")

                    if periodo not in periodos:
                        periodos[periodo] = []
                    periodos[periodo].append((data_formatada, dolar_formatado))

                # Usa a função parse_periodo corrigida para ordenar do mais recente ao mais antigo
                def parse_periodo(periodo):
                    try:
                        periodo = periodo.replace("á", "a").replace("à", "a").strip()
                        if " a " in periodo:
                            segunda_data_str = periodo.split(" a ")[-1].strip()
                        else:
                            segunda_data_str = periodo.split("a")[-1].strip()

                        partes = segunda_data_str.split("/")
                        if len(partes[2]) == 2:
                            partes[2] = "20" + partes[2]
                            segunda_data_str = "/".join(partes)

                        return datetime.strptime(segunda_data_str, "%d/%m/%Y")
                    except Exception as e:
                        print(f"Erro ao converter período '{periodo}': {e}")
                        return datetime.min

                # Ordena os períodos usando a data final (mais recente primeiro)
                periodos_ordenados = sorted(periodos.items(), key=lambda x: parse_periodo(x[0]), reverse=True)

                # Limpa o Treeview
                for item in self.tree.get_children():
                    self.tree.delete(item)

                # Preenche o Treeview
                for periodo, registros_periodo in periodos_ordenados:
                    registros_ordenados = sorted(registros_periodo, key=lambda x: datetime.strptime(x[0], "%d/%m/%Y"), reverse=True)
                    id_periodo = self.tree.insert("", "end", text=periodo, open=False)
                    for data, dolar in registros_ordenados:
                        self.tree.insert(id_periodo, "end", values=(data, dolar))
                    self.calcular_media(id_periodo)

                # Atualiza o último período carregado (o mais recente)
                self.ultimo_periodo = periodos_ordenados[0][0]

            cursor.close()

        except Exception as e:
            print(f"Erro ao carregar dados: {e}")

    def pesquisar_registros(self, event=None):
        texto_pesquisa = self.entrada_pesquisa.get().strip()
        
        # Limpa o Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            cursor = self.conexao.cursor()
            if texto_pesquisa == "":
                cursor.execute("""
                    SELECT periodo, data, dolar 
                    FROM cotacao_dolar 
                    ORDER BY data DESC
                """)
            else:
                busca = f"%{texto_pesquisa}%"
                cursor.execute("""
                    SELECT periodo, data, dolar
                    FROM cotacao_dolar
                    WHERE periodo ILIKE %s OR to_char(data, 'DD/MM/YYYY') ILIKE %s OR dolar::text ILIKE %s
                    ORDER BY data DESC
                """, (busca, busca, busca))
            
            registros = cursor.fetchall()
            if registros:
                periodos = {}
                for periodo, data, dolar in registros:
                    data_formatada = data.strftime("%d/%m/%Y")
                    dolar_formatado = f"R$ {dolar:.4f}".replace(".", ",")
                    if periodo not in periodos:
                        periodos[periodo] = []
                    periodos[periodo].append((data_formatada, dolar_formatado))
                
                # Reinsere os dados filtrados no Treeview
                for periodo, registros_periodo in periodos.items():
                    aberto = False if texto_pesquisa == "" else True
                    id_periodo = self.tree.insert("", "end", text=periodo, open=aberto)
                    for item in registros_periodo:
                        if isinstance(item, tuple) and len(item) == 2:
                            data, dolar = item
                            self.tree.insert(id_periodo, "end", values=(data, dolar))
                        else:
                            print(f"[AVISO] Registro ignorado por formato inválido: {item}")
                    try:
                        self.calcular_media(id_periodo)
                    except Exception as e:
                        print(f"[ERRO] Falha ao calcular média para {periodo}: {e}")
            cursor.close()
            
        except Exception as e:
            print(f"Erro na pesquisa: {e}")

    def formatar_periodo(self, event):
        # Obtém o texto atual e extrai apenas os dígitos
        raw_text = self.entrada_periodo.get()
        texto = ''.join(filter(str.isdigit, raw_text))
        
        # Limita a no máximo 12 dígitos (6 para cada data)
        if len(texto) > 12:
            texto = texto[:12]
        
        # Se ainda não temos 6 dígitos, não formata (permite digitar sem interferência)
        if len(texto) < 6:
            return

        # Formata a primeira data (primeiros 6 dígitos)
        primeira_data = f"{texto[0:2]}/{texto[2:4]}/{texto[4:6]}"
        
        # Se houver mais de 6 dígitos, formata o segundo período
        if len(texto) > 6:
            # Se houver pelo menos 2 dígitos para a segunda data, formata conforme disponíveis
            segunda_data = ""
            if len(texto) >= 8:
                segunda_data = f"{texto[6:8]}"
                if len(texto) >= 10:
                    segunda_data += f"/{texto[8:10]}"
                    if len(texto) >= 12:
                        segunda_data += f"/{texto[10:12]}"
                    else:
                        segunda_data += f"/{texto[10:]}"
                else:
                    segunda_data += f"/{texto[8:]}"
            else:
                segunda_data = texto[6:]
            texto_formatado = primeira_data + " á " + segunda_data
        else:
            texto_formatado = primeira_data

        # Atualiza a Entry com o texto formatado
        self.entrada_periodo.delete(0, tk.END)
        self.entrada_periodo.insert(0, texto_formatado)

    def formatar_data(self,event):
        texto = self.entrada_data.get().replace("/", "")[:8]  # Remove barras existentes e limita a 8 caracteres
        novo_texto = ""

        for i, char in enumerate(texto):
            if i in [2, 4]:  # Adiciona '/' após o dia e o mês
                novo_texto += "/"
            novo_texto += char

        self.entrada_data.delete(0, tk.END)
        self.entrada_data.insert(0, novo_texto)

    def _formatar_numero(self, entry_widget, casas=2):
        """
        Formata o conteúdo do entry_widget com 'casas' casas decimais
        e separadores de milhar no padrão brasileiro.
        """
        texto = entry_widget.get()
        dígitos = re.sub(r"\D", "", texto)  # só números
        if not dígitos:
            # se não tiver nada, preenche com zeros
            novo = "0," + "0"*casas
        else:
            # divide pelo fator 10**casas para float
            fator = 10**casas
            valor = int(dígitos) / fator
            # formata no estilo en_US com vírgulas e pontos
            style = f"{{valor:,.{casas}f}}"
            s = style.format(valor=valor)  # ex: "1,234.5678" para casas=4
            # troca separadores para PT-BR
            novo = s.replace(",", "X").replace(".", ",").replace("X", ".")

        # atualiza sem recursão
        entry_widget.unbind("<KeyRelease>")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, novo)
        entry_widget.bind(
            "<KeyRelease>",
            lambda ev, ent=entry_widget: self._formatar_numero(ent, casas)
        )

    def salvar_registro(self):
        periodo = self.entrada_periodo.get().strip()
        data   = self.entrada_data.get().strip()
        dolar  = self.entrada_dolar.get().strip().replace(",", ".")  # já numericamente “.0000”

        # Se não digitou período, usamos o último registrado
        if not periodo and self.ultimo_periodo:
            periodo = self.ultimo_periodo

        # Agora, apenas Data e Dólar são obrigatórios
        if not data or not dolar:
            messagebox.showwarning("Aviso", "Preencha Data e Dólar antes de salvar!")
            return

        # Valida data
        try:
            data_validada = datetime.strptime(data, "%d/%m/%Y").date()
        except ValueError:
            messagebox.showerror("Erro", "Data inválida! Use o formato DD/MM/AAAA.")
            return

        try:
            cursor = self.conexao.cursor()
            # Gera o menor ID livre
            cursor.execute("""
                WITH missing_ids AS (
                    SELECT generate_series(1, COALESCE((SELECT MAX(id) FROM cotacao_dolar), 0) + 1) AS candidate
                )
                SELECT candidate
                FROM missing_ids
                EXCEPT
                SELECT id FROM cotacao_dolar
                ORDER BY candidate
                LIMIT 1;
            """)
            menor_id = cursor.fetchone()[0]

            # Insere no banco
            cursor.execute("""
                INSERT INTO cotacao_dolar (id, periodo, data, dolar)
                VALUES (%s, %s, %s, %s)
            """, (menor_id, periodo, data_validada, dolar))
            self.conexao.commit()

            # Notifica e atualiza Treeview
            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            self.conexao.commit()

            # Mensagem de sucesso
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso!")

            # Insere/atualiza na Treeview
            item_pai = ""
            for child in self.tree.get_children():
                if self.tree.item(child, "text") == periodo:
                    item_pai = child
                    break
            if not item_pai:
                item_pai = self.tree.insert("", 0, text=periodo, open=True)

            dolar_formatado = f"R$ {float(dolar):.4f}".replace(".", ",")
            self.tree.insert(item_pai, "end", values=(data, dolar_formatado))
            self.calcular_media(item_pai)

            self.ultimo_periodo = periodo
            cursor.close()

        except Exception as e:
            self.conexao.rollback()
            messagebox.showerror("Erro", f"Erro ao salvar registro:\n{e}")
            return

        # Limpa apenas Data e Dólar (e mantém período para próxima edição/inserção, se quiser)
        # Se preferir também limpar período, basta descomentar a linha abaixo:
        # self.entrada_periodo.delete(0, tk.END)
        self.entrada_data.delete(0, tk.END)
        self.entrada_dolar.delete(0, tk.END)
    
    def calcular_media(self, item_pai):
        periodo = self.tree.item(item_pai, "text")
        
        try:
            cursor = self.conexao.cursor()
            cursor.execute("""
                SELECT data, dolar
                FROM cotacao_dolar
                WHERE periodo = %s
            """, (periodo,))
            
            registros = cursor.fetchall()
            
            # Converte os valores para float e calcula a média
            valores = [float(dolar) for _, dolar in registros]
            
            if valores:
                media = sum(valores) / len(valores)
                filhos = self.tree.get_children(item_pai)
                
                # Verifica se já existe uma linha de média para o período e a remove
                for filho in filhos:
                    if self.tree.item(filho, 'values')[0] == "Média:":
                        self.tree.delete(filho)
                        break
                
                # Formata a média para exibir com R$ e vírgula (ex.: R$ 1.234,56)
                media_formatada = f"R$ {media:.4f}".replace(".", ",")
                self.tree.insert(item_pai, "end", values=("Média:", media_formatada), tags=("Medium",))
                self.tree.tag_configure("Medium", font=("Helvetica", 10, "bold"))
            
            cursor.close()
            
        except Exception as e:
            print(f"Erro ao calcular a média: {e}")
    
    def editar_registro(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item = selected_item[0]

        if not self.tree.parent(item):  # É um período
            periodo_atual = self.tree.item(item, "text")
            periodo_novo = self.entrada_periodo.get().strip()

            if not periodo_novo:
                periodo_novo = periodo_atual

            try:
                cursor = self.conexao.cursor()

                if periodo_novo != periodo_atual:
                    cursor.execute("""
                        UPDATE cotacao_dolar
                        SET periodo = %s
                        WHERE periodo = %s
                    """, (periodo_novo, periodo_atual))
                    self.conexao.commit()

                    cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    self.conexao.commit()

                    self.tree.item(item, text=periodo_novo)
                    messagebox.showinfo("Sucesso", "Período atualizado com sucesso!")

                cursor.close()

            except Exception as e:
                self.conexao.rollback()
                messagebox.showerror("Erro", f"Erro ao editar o período:\n{e}")

        else:  # É um registro de dado
            periodo_atual = self.tree.item(self.tree.parent(item), "text")
            data_nova = self.entrada_data.get().strip()
            dolar_novo = self.entrada_dolar.get().strip().replace(",", ".")

            try:
                cursor = self.conexao.cursor()

                cursor.execute("""
                    SELECT data, dolar FROM cotacao_dolar
                    WHERE periodo = %s AND data = %s
                """, (periodo_atual, self.tree.item(item, "values")[0]))
                registro_atual = cursor.fetchone()

                if registro_atual:
                    data_atual, dolar_atual = registro_atual

                    if not data_nova:
                        data_nova = data_atual.strftime("%Y-%m-%d")
                    if not dolar_novo:
                        dolar_novo = str(dolar_atual)

                    # Valida nova data
                    try:
                        data_convertida = datetime.strptime(data_nova, "%d/%m/%Y").date()
                    except ValueError:
                        data_convertida = datetime.strptime(data_nova, "%Y-%m-%d").date()

                    if data_convertida != data_atual or float(dolar_novo) != float(dolar_atual):
                        cursor.execute("""
                            UPDATE cotacao_dolar
                            SET data = %s, dolar = %s
                            WHERE periodo = %s AND data = %s
                        """, (data_convertida, dolar_novo, periodo_atual, data_atual))
                        self.conexao.commit()

                        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                        self.conexao.commit()

                        data_formatada = data_convertida.strftime("%d/%m/%Y")
                        dolar_formatado = f"R$ {float(dolar_novo):.4f}".replace(".", ",")

                        self.tree.item(item, values=(data_formatada, dolar_formatado))
                        self.calcular_media(self.tree.parent(item))
                        messagebox.showinfo("Sucesso", "Registro editado com sucesso!")

                cursor.close()

            except Exception as e:
                self.conexao.rollback()
                messagebox.showerror("Erro", f"Erro ao editar o registro:\n{e}")

        # Limpa campos
        self.entrada_periodo.delete(0, tk.END)
        self.entrada_data.delete(0, tk.END)
        self.entrada_dolar.delete(0, tk.END)

    def carregar_dados_selecionados(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item = selected_item[0]

        # Se o item não tem um "pai", então é um "período"
        if not self.tree.parent(item):
            periodo = self.tree.item(item, "text")
            print(f"Período selecionado: {periodo}")  # Depuração: Verificando o valor do período
            
            # Limpar e carregar o valor do "período"
            self.entrada_data.delete(0, 'end')
            self.entrada_dolar.delete(0, 'end')
            self.entrada_periodo.delete(0, 'end')
            self.entrada_periodo.insert(0, periodo)

        else:
            # Caso contrário, é um "registro de dado", carregar os valores de "data" e "dolar"
            periodo = self.tree.item(self.tree.parent(item), "text")
            data, dolar = self.tree.item(item, "values")
            
            print(f"Registro selecionado: periodo={periodo}, data={data}, dolar={dolar}")  # Depuração

            # Remover o "R$ " do valor do dólar antes de inserir no campo
            dolar_sem_rs = dolar.replace("R$ ", "").strip()

            # Carregar os valores nos campos de entrada
            self.entrada_data.delete(0, 'end')
            self.entrada_data.insert(0, data)

            self.entrada_dolar.delete(0, 'end')
            self.entrada_dolar.insert(0, dolar_sem_rs)

    def exportar_relatorio_excel_dolar(self):
        def formatar_data(event):
            entry = event.widget
            data = entry.get().strip()
            data = ''.join(filter(str.isdigit, data))
            data = data[:8]
            if len(data) >= 2:
                data = data[:2] + '/' + data[2:]
            if len(data) >= 5:
                data = data[:5] + '/' + data[5:]
            entry.delete(0, tk.END)
            entry.insert(0, data)

        def gerar_excel():
            data_inicio = entry_inicio.get().strip()
            data_fim = entry_fim.get().strip()

            data_inicio_dt = None
            data_fim_dt = None

            try:
                if data_inicio:
                    data_inicio_dt = datetime.strptime(data_inicio, "%d/%m/%Y").date()
                if data_fim:
                    data_fim_dt = datetime.strptime(data_fim, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showerror("Filtro de Data", "Formato de data inválido. Use DD/MM/YYYY.")
                return

            try:
                dados = []
                for item_periodo in self.tree.get_children():
                    periodo = self.tree.item(item_periodo)['text']
                    registros_periodo = self.tree.get_children(item_periodo)
                    grupo_registros = []
                    incluir_grupo = False

                    for item_registro in registros_periodo:
                        valores = self.tree.item(item_registro)['values']
                        if valores and len(valores) >= 2:
                            data_str, dolar_str = valores[0], valores[1]
                            try:
                                data_dt = datetime.strptime(data_str, "%d/%m/%Y").date()
                            except Exception:
                                continue

                            if data_inicio_dt or data_fim_dt:
                                if data_inicio_dt and data_dt < data_inicio_dt:
                                    pass
                                elif data_fim_dt and data_dt > data_fim_dt:
                                    pass
                                else:
                                    incluir_grupo = True
                            else:
                                incluir_grupo = True

                            grupo_registros.append((periodo, data_str, dolar_str, data_dt))
                    if incluir_grupo and grupo_registros:
                        dados.extend(grupo_registros)

                if not dados:
                    messagebox.showinfo("Exportação", "Nenhum dado encontrado para o período selecionado.")
                    return

                df_relatorio = pd.DataFrame(dados, columns=["Período", "Data", "Dólar", "Data_dt"])

                def parse_dolar(valor):
                    v = valor.replace("R$", "").strip().replace(".", "").replace(",", ".")
                    try:
                        return float(v)
                    except:
                        return None

                df_relatorio["Dólar_Num"] = df_relatorio["Dólar"].apply(parse_dolar)

                periodos_ordenados = sorted(
                    df_relatorio.groupby("Período"),
                    key=lambda x: x[1]["Data_dt"].max(),
                    reverse=True
                )

                df_final = pd.DataFrame(columns=["Período", "Data", "Dólar"])
                linhas_em_negrito = []

                for periodo, grupo in periodos_ordenados:
                    grupo = grupo.copy()
                    grupo = grupo.sort_values("Data_dt", ascending=False)
                    media = grupo["Dólar_Num"].mean()

                    df_final = pd.concat([df_final, grupo[["Período", "Data", "Dólar"]]], ignore_index=True)

                    media_formatada = f"R$ {media:,.4f}" if media is not None else ""
                    linha_media = pd.DataFrame([[periodo, "Média", media_formatada]], columns=["Período", "Data", "Dólar"])
                    df_final = pd.concat([df_final, linha_media], ignore_index=True)

                    linha_media_excel = len(df_final)
                    linhas_em_negrito.append(linha_media_excel)

                    df_final = pd.concat([df_final, pd.DataFrame([["", "", ""]], columns=["Período", "Data", "Dólar"])],
                                        ignore_index=True)

                caminho_arquivo = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Planilhas do Excel", "*.xlsx")],
                    title="Salvar Relatório Dólar como",
                    initialfile="Relatorio_Dolar.xlsx"
                )
                if not caminho_arquivo:
                    return

                with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
                    df_final.to_excel(writer, sheet_name="Relatório Dólar", index=False)
                    wb = writer.book
                    ws = writer.sheets["Relatório Dólar"]

                    ws.auto_filter.ref = ws.dimensions

                    negrito = Font(bold=True)
                    for cell in ws[1]:
                        cell.font = negrito

                    for linha_idx in linhas_em_negrito:
                        for cell in ws[linha_idx + 1]:
                            cell.font = negrito

                messagebox.showinfo("Exportação", f"Arquivo exportado com sucesso!\n{caminho_arquivo}")
                popup.destroy()

            except Exception as e:
                messagebox.showerror("Erro na Exportação", f"Erro ao exportar Relatório Dólar:\n{str(e)}")

        popup = tk.Toplevel(self.parent_frame)
        popup.title("Exportar Excel - Filtro de Data")
        popup.geometry("300x180")
        popup.resizable(False, False)

        centralizar_janela(popup, 200, 200)
        aplicar_icone(popup, r"C:\\Sistema\\logos\\Kametal.ico")
        popup.config(bg="#ecf0f1")

        style = ttk.Style(popup)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white", font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton", background=[("active", "#3498db")], foreground=[("active", "white")])

        ttk.Label(popup, text="Data Início (DD/MM/YYYY):", style="Custom.TLabel").pack(pady=5)
        entry_inicio = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_inicio.pack(pady=5)
        entry_inicio.bind("<KeyRelease>", formatar_data)

        ttk.Label(popup, text="Data Fim (DD/MM/YYYY):", style="Custom.TLabel").pack(pady=5)
        entry_fim = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_fim.pack(pady=5)
        entry_fim.bind("<KeyRelease>", formatar_data)

        botoes_frame = ttk.Frame(popup, style="Custom.TFrame")
        botoes_frame.pack(pady=10)

        ttk.Button(botoes_frame, text="Exportar Excel", command=gerar_excel).pack(side="left", padx=5)
        ttk.Button(botoes_frame, text="Cancelar", command=popup.destroy).pack(side="left", padx=5)

    def exportar_relatorio_pdf_dolar(self):
        """
        Abre uma janela para filtro de período e exporta o PDF da cotação de Dólar,
        agrupando os registros em faixas semanais.
        """
        def formatar_data(event):
            entry = event.widget
            data = entry.get().strip()
            data = ''.join(filter(str.isdigit, data))
            data = data[:8]
            if len(data) >= 2:
                data = data[:2] + '/' + data[2:]
            if len(data) >= 5:
                data = data[:5] + '/' + data[5:]
            entry.delete(0, tk.END)
            entry.insert(0, data)

        def calcular_periodo(data_obj):
            dia, mes, ano = data_obj.day, data_obj.month, data_obj.year
            if dia <= 7:
                inicio, fim = data_obj.replace(day=1), data_obj.replace(day=7)
            elif dia <= 14:
                inicio, fim = data_obj.replace(day=8), data_obj.replace(day=14)
            elif dia <= 21:
                inicio, fim = data_obj.replace(day=15), data_obj.replace(day=21)
            elif dia <= 28:
                inicio, fim = data_obj.replace(day=22), data_obj.replace(day=28)
            else:
                inicio = data_obj.replace(day=29)
                ultimo = monthrange(ano, mes)[1]
                fim = data_obj.replace(day=ultimo)
            return f"{inicio.strftime('%d/%m/%y')} á {fim.strftime('%d/%m/%y')}"

        def gerar_pdf():
            # lê e converte filtros
            inicio_txt = entry_inicio.get().strip()
            fim_txt    = entry_fim.get().strip()
            try:
                dt_inicio = datetime.strptime(inicio_txt, "%d/%m/%Y").date() if inicio_txt else None
                dt_fim    = datetime.strptime(fim_txt,    "%d/%m/%Y").date() if fim_txt else None
            except ValueError:
                messagebox.showerror("Filtro de Data", "Use o formato DD/MM/YYYY.")
                return

            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialfile="relatorio_dolar.pdf",
                filetypes=[("PDF","*.pdf")],
                title="Salvar Relatório de Cotação de Dólar"
            )
            if not caminho:
                return

            try:
                # busca todos os registros
                cur = self.conexao.cursor()
                cur.execute("SELECT data, dolar FROM cotacao_dolar ORDER BY data DESC")
                raw = cur.fetchall()
                cur.close()

                if not raw:
                    messagebox.showinfo("Exportar PDF", "Nenhum registro encontrado.")
                    return

                # processa e agrupa por período
                grupos = {}
                for data_raw, valor in raw:
                    if isinstance(data_raw, str):
                        data_obj = datetime.strptime(data_raw, "%Y-%m-%d").date()
                    elif isinstance(data_raw, datetime):
                        data_obj = data_raw.date()
                    else:
                        data_obj = data_raw
                    periodo = calcular_periodo(data_obj)
                    grupos.setdefault(periodo, []).append((data_obj, valor))

                # filtra grupos pelo intervalo
                selecionados = {}
                for per, regs in grupos.items():
                    if dt_inicio or dt_fim:
                        if any((not dt_inicio or d>=dt_inicio) and (not dt_fim or d<=dt_fim) for d,_ in regs):
                            selecionados[per] = regs
                    else:
                        selecionados[per] = regs

                if not selecionados:
                    messagebox.showinfo("Exportar PDF", "Nenhum dado para o período.")
                    return

                # ordena períodos e registros internos
                def chave(p): return datetime.strptime(p.split("á")[0].strip(), "%d/%m/%y")
                ordenado = sorted(selecionados.items(), key=lambda x: chave(x[0]), reverse=True)

                # monta linhas finais
                linhas = []
                for per, regs in ordenado:
                    regs_ordenados = sorted(regs, key=lambda x: x[0], reverse=True)
                    for d, v in regs_ordenados:
                        data_fmt = d.strftime("%d/%m/%Y")
                        dolar_fmt = f"R$ {float(v):.4f}".replace(".", ",")
                        linhas.append((per, data_fmt, dolar_fmt))

                # === começa Platypus ===
                doc       = SimpleDocTemplate(caminho, pagesize=A4)
                elementos = []
                estilos   = getSampleStyleSheet()

                # título
                elementos.append(Paragraph("<b>Relatório de Cotação de Dólar</b>", estilos["Title"]))
                elementos.append(Spacer(1, 20))

                # cabeçalhos e dados
                headers = ["Período", "Data", "Dólar"]
                dados   = [headers] + linhas

                tabela = Table(dados, hAlign="CENTER")
                tabela.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,0), colors.grey),
                    ("TEXTCOLOR",     (0,0), (-1,0), colors.whitesmoke),
                    ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                    ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0,0), (-1,0), 12),
                    ("GRID",          (0,0), (-1,-1), 0.5, colors.black),
                ]))

                elementos.append(tabela)
                doc.build(elementos)

                messagebox.showinfo("Exportar PDF", f"PDF salvo em:\n{caminho}")
                popup.destroy()

            except Exception as e:
                messagebox.showerror("Exportar PDF", f"Erro ao exportar PDF: {e}")

        # cria popup de filtro (igual ao anterior)
        popup = tk.Toplevel(self.parent_frame)
        popup.title("Exportar PDF - Filtro de Data")
        popup.geometry("300x200")
        popup.resizable(False, False)
        centralizar_janela(popup, 300, 200)
        aplicar_icone(popup, r"C:\Sistema\logos\Kametal.ico")
        popup.config(bg="#ecf0f1")

        style = ttk.Style(popup)
        style.theme_use("alt")
        style.configure("Custom.TLabel",  background="#ecf0f1", foreground="#34495e", font=("Arial",10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white",
                        font=("Arial",10,"bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active","#3498db")], foreground=[("active","white")])

        ttk.Label(popup, text="Data Início (DD/MM/YYYY):", style="Custom.TLabel").pack(pady=5)
        entry_inicio = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_inicio.pack(pady=5)
        entry_inicio.bind("<KeyRelease>", formatar_data)

        ttk.Label(popup, text="Data Fim (DD/MM/YYYY):", style="Custom.TLabel").pack(pady=5)
        entry_fim = tk.Entry(popup, font=("Arial", 10), relief="solid", borderwidth=1)
        entry_fim.pack(pady=5)
        entry_fim.bind("<KeyRelease>", formatar_data)

        botoes = ttk.Frame(popup)
        botoes.pack(pady=10)
        ttk.Button(botoes, text="Exportar PDF", command=gerar_pdf).pack(side="left", padx=5)
        ttk.Button(botoes, text="Cancelar", command=popup.destroy).pack(side="left", padx=5)
    
    def on_item_selecionado(self, event):
        self.carregar_dados_selecionados()
    
    def excluir_registro(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        # Pergunta se o usuário realmente quer excluir
        confirmacao = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir o(s) registro(s)?")
        if not confirmacao:
            return  # Se o usuário escolher "Não", interrompe a função

        for item in selected_item:
            parent = self.tree.parent(item)
            
            if not parent:  # Se não tem pai, é um período
                periodo = self.tree.item(item, "text")
                try:
                    cursor = self.conexao.cursor()
                    cursor.execute("""
                        DELETE FROM cotacao_dolar WHERE periodo = %s
                    """, (periodo,))
                    self.conexao.commit()

                    cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    self.conexao.commit()

                    # Remove do Treeview
                    for child in self.tree.get_children(item):
                        self.tree.delete(child)
                    self.tree.delete(item)

                    cursor.close()

                except Exception as e:
                    print(f"Erro ao excluir os registros do período: {e}")
            
            else:  # Se tem um pai, é um registro de data e dólar
                data = self.tree.item(item, "values")[0]
                dolar = self.tree.item(item, "values")[1]
                periodo = self.tree.item(parent, "text")

                # Limpa a formatação do valor do dólar
                dolar_limpo = dolar.replace("R$ ", "").replace(".", "").replace(",", ".")

                try:
                    cursor = self.conexao.cursor()
                    cursor.execute("""
                        DELETE FROM cotacao_dolar WHERE periodo = %s AND data = %s AND dolar = %s
                    """, (periodo, data, dolar_limpo))
                    self.conexao.commit()

                    cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    self.conexao.commit()

                    # Remove do Treeview
                    self.tree.delete(item)
                    self.calcular_media(parent)

                    cursor.close()

                except Exception as e:
                    print(f"Erro ao excluir o registro: {e}")

        # Mensagem de sucesso
        messagebox.showinfo("Sucesso", "Registro(s) excluído(s) com sucesso!")

        # Limpar campos após excluir
        self.entrada_periodo.delete(0, tk.END)
        self.entrada_data.delete(0, tk.END)
        self.entrada_dolar.delete(0, tk.END)
    
    def limpar_entradas(self):
        self.entrada_periodo.delete(0, tk.END)
        self.entrada_data.delete(0, tk.END)
        self.entrada_dolar.delete(0, tk.END)

    def on_closing(self):
        pass  # Só pra teste
        # Aqui você pode desconectar eventos, encerrar conexões, etc.
