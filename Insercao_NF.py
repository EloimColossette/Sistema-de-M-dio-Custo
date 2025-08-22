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
    # Arrumar a largura do treeview
    def __init__(self, parent, janela_menu):
        super().__init__()
        # Configuração inicial da janela
        self.resizable(True, True)  # Permite redimensionamento
        self.geometry("900x500")  # Define o tamanho inicial
        self.state("normal")  # Garante o estado inicial como "normal"

        self.janela_menu = janela_menu  # Armazena a referência da janela do menu
        self.parent = parent
        self.title("Entrada de Nfs")
        self.state("zoomed")  # Se precisar que abra maximizada, pode manter após configurar tudo

        self.caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self, self.caminho_icone)

        # Resetar estilos para evitar herança de outras janelas
        style = ttk.Style(self)
        style.theme_use("alt")  # Define o tema padrão
        style.configure(".", font=("Arial", 10))  # Fonte padrão para todos os widgets
        style.configure("Treeview", rowheight=25)  # Altura das linhas no Treeview
        style.configure("Treeview.Heading", font=("Courier", 10, "bold"))  # Cabeçalhos do Treeview

        # Configuração do banco de dados usando a biblioteca conectar
        self.conn = conectar()
        self.cursor = self.conn.cursor()

        # Mapeamento de colunas do banco de dados para nomes amigáveis
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
            "produto", "custo_empresa", "ipi", "valor_integral",  # Adicionando IPI após Custo empresa
            "valor_unitario_1", "valor_unitario_2", "valor_unitario_3",
            "valor_unitario_4", "valor_unitario_5",
            "duplicata_1", "duplicata_2", "duplicata_3", "duplicata_4",
            "duplicata_5", "duplicata_6",
            "valor_unitario_energia", "valor_mao_obra_tm_metallica",  # Adicionando Energia e Mão de obra
            "peso_liquido", "peso_integral"  # Adicionando Quantidade Usada e Peso Integral
        ]

        # Definir as colunas que podem ser ocultadas
        self.colunas_ocultaveis = ["material_1", "material_2", "material_3", "material_4", "material_5", "valor_unitario_1", "valor_unitario_2", "valor_unitario_3", "valor_unitario_4", "valor_unitario_5", "duplicata_1", "duplicata_2", "duplicata_3", "duplicata_4", "duplicata_5", "duplicata_6"]

        # Inicializar colunas visíveis com todas as colunas
        self.colunas_visiveis = list(self.colunas_fixas.keys())
        self.colunas_ocultas = []  # Lista para armazenar colunas ocultas
        self.larguras_colunas = {coluna: 120 for coluna in self.colunas_fixas.keys()}  # Defina uma largura padrão para todas as colunas

        # Criar o Treeview com colunas fixas
        self.tree = ttk.Treeview(self, columns=self.colunas_visiveis, show="headings")
        self.configurar_treeview()

        # Adicionando barra de rolagem vertical
        scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar_y.set)
        scrollbar_y.grid(row=0, column=4, sticky="ns")

        # Adicionando barra de rolagem horizontal
        scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscroll=scrollbar_x.set)
        scrollbar_x.grid(row=1, column=0, columnspan=4, sticky="ew")

        # Frame para operações com linhas
        self.frame_operacoes_linhas = tk.LabelFrame(self, text="Operações com Linhas")
        self.frame_operacoes_linhas.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=5)

        # Botões para operações com linhas
        self.btn_insert = tk.Button(self.frame_operacoes_linhas, text="Inserir Linha", command=self.inserir_linha)
        self.btn_insert.grid(row=0, column=0, padx=5, pady=5)

        self.btn_edit = tk.Button(self.frame_operacoes_linhas, text="Editar Linha", command=self.editar_linha)
        self.btn_edit.grid(row=0, column=1, padx=5, pady=5)

        self.btn_remove_items = tk.Button(self.frame_operacoes_linhas, text="Excluir Itens", command=self.remove_selected_items)
        self.btn_remove_items.grid(row=0, column=2, padx=5, pady=5)

        # Botão para calcular a média ponderada
        self.btn_media_ponderada = tk.Button(self.frame_operacoes_linhas, text="Calculadora de Média", command=self.chamar_calculadora_media_ponderada)
        self.btn_media_ponderada.grid(row=0, column=3, padx=5, pady=5)

        # Crie uma StringVar e vincule-a ao Entry
        self.search_var = tk.StringVar()
        # Adicione o trace para formatação em tempo real
        self.trace_id = self.search_var.trace_add("write", self.formatar_data_em_tempo_real)

        self.lbl_pesquisa = tk.Label(self.frame_operacoes_linhas, text="Buscar:")
        self.lbl_pesquisa.grid(row=0, column=4, padx=5, pady=5, sticky="e")

        # Use a StringVar no textvariable do Entry
        self.entry_pesquisa = tk.Entry(self.frame_operacoes_linhas, width=50, textvariable=self.search_var)
        self.entry_pesquisa.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        # Vincula o pressionamento da tecla Enter à função de pesquisa
        self.entry_pesquisa.bind("<Return>", lambda event: self.pesquisar())

        self.btn_pesquisar = tk.Button(self.frame_operacoes_linhas, text="Pesquisar", command=self.pesquisar)
        self.btn_pesquisar.grid(row=0, column=6, padx=5, pady=5)

        # Ajustar as colunas para que a última coluna se expanda corretamente
        self.frame_operacoes_linhas.columnconfigure(4, weight=1)

        # Configurar o layout da janela
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # Adicionando o Treeview à janela
        self.tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

        # Frame para operações com colunas (ocultar e mostrar)
        self.frame_operacoes_colunas = tk.LabelFrame(self, text="Exibir/Ocultar Colunas")
        self.frame_operacoes_colunas.grid(row=4, column=0, columnspan=4, sticky="ew", padx=10, pady=5)

       # Combobox para selecionar colunas
        self.combobox_colunas = ttk.Combobox(
            self.frame_operacoes_colunas, 
            values=[self.colunas_fixas[coluna] for coluna in self.colunas_ocultaveis], 
            state="readonly"
        )
        self.combobox_colunas.grid(row=0, column=0, padx=10, pady=5)
        self.combobox_colunas.set("Selecione uma coluna")

        # Botões para ocultar e mostrar colunas
        self.btn_ocultar = tk.Button(self.frame_operacoes_colunas, text="Ocultar Coluna", command=self.ocultar_coluna_selecionada)
        self.btn_ocultar.grid(row=0, column=1, padx=5, pady=5)

        self.btn_mostrar = tk.Button(self.frame_operacoes_colunas, text="Mostrar Coluna", command=self.mostrar_coluna_selecionada)
        self.btn_mostrar.grid(row=0, column=2, padx=5, pady=5)

        # Botão de Voltar
        self.btn_voltar = tk.Button(self.frame_operacoes_colunas, text="Voltar", command=self.voltar_para_menu)
        self.btn_voltar.grid(row=0, column=3, padx=5, pady=5)

        # Coluna "espaçadora" para empurrar o botão Excel para a direita
        self.frame_operacoes_colunas.grid_columnconfigure(4, weight=1)

        # Botão de Excel na coluna 5, fixado à direita
        self.btn_exportar = tk.Button(self.frame_operacoes_colunas, text="Exportação Excel", command=self.abrir_dialogo_exportacao)
        self.btn_exportar.grid(row=0, column=5, padx=5, pady=5, sticky="e")

        self.dados_colunas_ocultas = {}

        # Carregar dados do banco de dados
        self.carregar_dados()

        # Carregar estado das colunas
        self.carregar_estado_colunas()

        # Configurar o comportamento ao fechar a janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

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
                ORDER BY data DESC, nf ASC;
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

        # Definir fundo e estilos padronizados
        janela_inserir.configure(bg="#ecf0f1")

        # Fonte dos títulos das entradas (reduzida)
        LABEL_FONT = ("Arial", 9, "bold")

        # Criar uma variável para o campo de data
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

        # Estilos exclusivos para essa janela
        style = ttk.Style(janela_inserir)
        style.theme_use("alt")
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

        # Estilos
        style = ttk.Style(janela_editar)
        style.theme_use("alt")
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
