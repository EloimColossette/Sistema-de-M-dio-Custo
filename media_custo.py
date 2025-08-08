import tkinter as tk
from tkinter import ttk, messagebox
import sys
from conexao_db import conectar
from logos import aplicar_icone
from menu import Janela_Menu
from aba_sudeste_media_custo import criar_media_custo_sudeste
from aba_centro_oeste_media_custo import criar_media_aba_centro_oeste  
from decimal import Decimal
from exportacao import exportar_notebook_para_excel

# Função para conectar ao banco de dados
def conectar_banco():
    try:
        conn = conectar()
        return conn
    except Exception as e:
        print("Erro ao conectar ao banco:", e)
        return None

# Função para buscar os nomes dos produtos no banco de dados
def buscar_produtos(conn):
    """
    Busca todos os nomes dos produtos na tabela 'produtos' do banco de dados.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: Lista de tuplas contendo os nomes dos produtos ou lista vazia em caso de erro.
    """
    try:
        cursor = conn.cursor()  # Cria um cursor para executar comandos SQL
        cursor.execute("SELECT nome FROM produtos ORDER BY nome;")  # Executa a consulta SQL
        produtos = cursor.fetchall()  # Recupera todos os resultados da consulta
        cursor.close()  # Fecha o cursor para liberar recursos
        return produtos
    except Exception as e:
        print("Erro ao buscar nome_produtos:", e)
        return []  # Retorna uma lista vazia em caso de erro

# Função para buscar e somar as quantidades de estoque de um produto específico
def buscar_estoque(conn, nome_produto):
    """
    Busca e soma a quantidade de estoque de um determinado produto.
    
    A função faz um JOIN entre as tabelas 'estoque_quantidade' e 'somar_produtos'
    para localizar o produto desejado e somar suas quantidades.
    
    :param conn: Conexão ativa com o banco de dados.
    :param nome_produto: Nome do produto a ser consultado.
    :return: String formatada representando a quantidade de estoque com 3 casas decimais e " kg".
    """
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT SUM(eq.quantidade_estoque)
            FROM estoque_quantidade eq
            JOIN somar_produtos sp ON eq.id_produto = sp.id
            WHERE sp.produto = %s
        """, (nome_produto,))
        
        quantidade = cursor.fetchone()[0]  # Recupera a soma do estoque
        cursor.close()

        if quantidade:
            # Formata o número para 3 casas decimais, ajustando separadores e adicionando " kg"
            return f"{quantidade:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"
        else:
            return "0,000 kg"  
    except Exception as e:
        print(f"Erro ao buscar estoque para o produto {nome_produto}: {e}")
        return "0,000 kg"
    
# Função para calcular a média ponderada do custo do estoque de um produto
def calcular_media_ponderada(conn, nome_produto):
    """
    Calcula a média ponderada do custo de um produto com base na quantidade em estoque.
    
    Realiza um JOIN entre 'estoque_quantidade' e 'somar_produtos', multiplica cada 
    quantidade pelo seu custo total e depois divide pela soma das quantidades.
    
    :param conn: Conexão ativa com o banco de dados.
    :param nome_produto: Nome do produto para o qual será calculada a média.
    :return: String formatada representando o custo médio ponderado como moeda.
    """
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT SUM(eq.quantidade_estoque * eq.custo_total), SUM(eq.quantidade_estoque)
            FROM estoque_quantidade eq
            JOIN somar_produtos sp ON eq.id_produto = sp.id
            WHERE sp.produto = %s
        """, (nome_produto,))

        soma_custo_ponderado, soma_quantidade = cursor.fetchone()  # Obtém os somatórios necessários
        cursor.close()

        if soma_quantidade and soma_custo_ponderado:
            media_ponderada = soma_custo_ponderado / soma_quantidade  # Calcula a média ponderada
            # Formata o valor para 2 casas decimais e adiciona "R$ " à frente
            return "R$ " + f"{media_ponderada:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return "R$ 0,00"
    except Exception as e:
        print(f"Erro ao calcular média ponderada para o produto {nome_produto}: {e}")
        return "R$ 0,00"
    
# Função para calcular o custo do produto incluindo o custo da mão de obra da empresa (empresa)
def calcular_custo_empresa(conn, nome_produto, media_custo_estoque): 
    """
    Soma o custo da mão de obra (empresa) ao custo médio do estoque.
    
    Se a média de custo do estoque for 0, retorna 0 imediatamente,
    ignorando o custo da empresa (valor cadastrado na tabela).
    
    :param conn: Conexão ativa com o banco de dados.
    :param nome_produto: Nome do produto.
    :param media_custo_estoque: Média ponderada do custo do estoque (float).
    :return: String formatada representando o custo total (estoque + empresa) como moeda.
    """
    # Se o custo médio for 0, ignora o valor da empresa
    if media_custo_estoque == 0:
        return "R$ 0,00"
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT custo_empresa FROM somar_produtos WHERE produto = %s
        """, (nome_produto,))
        custo_empresa = cursor.fetchone()
        cursor.close()

        if custo_empresa and custo_empresa[0] is not None:
            custo_empresa = float(custo_empresa[0])  # Converte o valor para float
            resultado = media_custo_estoque + custo_empresa  # Soma o custo da mão de obra
            return "R$ " + f"{resultado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return "R$ 0,00"  # Caso não haja valor registrado, retorna zero
    except Exception as e:
        print(f"Erro ao calcular custo empresa para o produto {nome_produto}: {e}")
        return "R$ 0,00"

# Função para calcular o custo com acréscimo de um percentual (exemplo: 5%)
def calcular_custo_5(custo_base, percentual):
    """
    Calcula o custo base com acréscimo de um percentual específico.
    
    Aplica a fórmula para aumentar o custo base de forma que ele corresponda a 100%
    após o acréscimo. Por exemplo, para 5% de acréscimo, divide-se por 0,95.
    
    :param custo_base: Valor base (pode ser float ou string formatada como moeda).
    :param percentual: Percentual de acréscimo (ex: 5 para 5%).
    :return: Valor final formatado como string no formato monetário.
    """
    try:
        # Se o custo base vier como string, converte para float removendo símbolos e separadores
        if isinstance(custo_base, str):
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.95  # Aplica o acréscimo (dividindo por 0.95 para 5%)
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de {percentual}%: {e}")
        return "R$ 0,00"
    
# Função para calcular o custo com acréscimo de 10%
def calcular_custo_10(custo_base):
    """
    Calcula o custo base com acréscimo de 10%.
    
    Para 10%, divide-se o custo base por 0,9.
    
    :param custo_base: Valor base (float ou string formatada).
    :return: Valor final formatado como moeda.
    """
    try:
        if isinstance(custo_base, str):
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.9  # Aplica o acréscimo de 10%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 10%: {e}")
        return "R$ 0,00"

# Função para calcular o custo com acréscimo de 15%
def calcular_custo_15(custo_base):
    """
    Calcula o custo base com acréscimo de 15%.
    
    Para 15%, divide-se o custo base por 0,85.
    
    :param custo_base: Valor base (float ou string formatada).
    :return: Valor final formatado como moeda.
    """
    try:
        if isinstance(custo_base, str):
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.85  # Aplica o acréscimo de 15%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 15%: {e}")
        return "R$ 0,00"
    
# Função para calcular o custo com acréscimo de 20%
def calcular_custo_20(custo_base):
    """
    Calcula o custo base com acréscimo de 20%.
    
    Para 20%, divide-se o custo base por 0,8.
    
    :param custo_base: Valor base (float ou string formatada).
    :return: Valor final formatado como moeda.
    """
    try:
        if isinstance(custo_base, str):
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.8  # Aplica o acréscimo de 20%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 20%: {e}")
        return "R$ 0,00"
    
# Função para calcular o total do estoque somando as quantidades de todos os produtos
def calcular_total_estoque(conn):
    """
    Percorre todos os produtos e acumula a quantidade de estoque de cada um.
    
    Converte a quantidade de cada produto para float e soma os valores.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: String representando o total do estoque formatado com 3 casas decimais e " kg".
    """
    total_estoque = 0.0
    produtos = buscar_produtos(conn)  # Obtém a lista de produtos

    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)  # Obtém o estoque do produto

        # Converte o estoque (string) para float, ajustando formatação
        estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))
        total_estoque += estoque_float  # Acumula o estoque total

    # Formata o total acumulado para 3 casas decimais e adiciona " kg"
    return f"{total_estoque:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"

# Função para calcular a média ponderada total do custo e do custo empresa do estoque
def calcular_media_ponderada_total_com_empresa(conn):
    """
    Calcula a média ponderada do custo total do estoque e também do custo empresa.
    
    Para cada produto, multiplica o custo (ou custo empresa) pela quantidade em estoque
    e acumula esses valores. Em seguida, divide pela soma total do estoque para obter a média.
    
    Se a média de custo do produto for 0, o custo empresa não é buscado (ou seja, considerado 0).
    
    :param conn: Conexão ativa com o banco de dados.
    :return: Dupla de strings formatadas representando a média ponderada do custo e do empresa.
    """
    produtos = buscar_produtos(conn)

    soma_ponderada_custo = 0.0
    soma_ponderada_empresa = 0.0
    soma_estoque = 0.0

    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)
        # Converte o estoque para float (removendo " kg", pontos e ajustando vírgula)
        estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))
        
        if estoque_float > 0:
            # Obtém o custo médio ponderado do produto e converte para float
            media_custo_str = calcular_media_ponderada(conn, nome_produto)
            media_custo = float(media_custo_str.replace("R$ ", "").replace(".", "").replace(",", "."))
            
            # Se o custo médio for 0, ignora o custo empresa
            if media_custo == 0:
                custo_empresa = 0.0
            else:
                custo_empresa_str = calcular_custo_empresa(conn, nome_produto, media_custo)
                custo_empresa = float(custo_empresa_str.replace("R$ ", "").replace(".", "").replace(",", "."))
            
            # Acumula os valores ponderados (valor * quantidade)
            soma_ponderada_custo += estoque_float * media_custo
            soma_ponderada_empresa += estoque_float * custo_empresa
            soma_estoque += estoque_float

    # Evita divisão por zero se o estoque total for zero
    if soma_estoque == 0:
        return "R$ 0,000", "R$ 0,000"

    # Calcula as médias ponderadas dividindo os somatórios pelo total do estoque
    media_ponderada_total = soma_ponderada_custo / soma_estoque
    media_ponderada_empresa = soma_ponderada_empresa / soma_estoque

    # Se a média ponderada total for 0, força a média do empresa para 0
    if media_ponderada_total == 0:
        media_ponderada_empresa = 0

    # Formata os valores para 2 casas decimais e adiciona o prefixo "R$ "
    media_ponderada_total_str = f"R$ {media_ponderada_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    media_ponderada_empresa_str = f"R$ {media_ponderada_empresa:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_ponderada_total_str, media_ponderada_empresa_str

# Função para calcular o custo empresa total com acréscimo de 5%
def calcular_media_total_5(conn):
    """
    Calcula o custo empresa total (média ponderada) com acréscimo de 5%.
    
    A operação consiste em dividir o custo empresa ponderado por 0,95 e formatar o resultado.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: String formatada representando o custo com acréscimo de 5%.
    """
    # Obtém a média ponderada do custo empresa
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
    
    media_5 = media_ponderada_empresa_float / 0.95  # Aplica acréscimo de 5%
    media_5_str = f"R$ {media_5:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_5_str

# Função para calcular o custo empresa total com acréscimo de 10%
def calcular_media_total_10(conn):
    """
    Calcula o custo empresa total (média ponderada) com acréscimo de 10%.
    
    Divide o custo empresa ponderado por 0,9.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: String formatada representando o custo com acréscimo de 10%.
    """
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
    
    media_10 = media_ponderada_empresa_float / 0.9  # Aplica acréscimo de 10%
    media_10_str = f"R$ {media_10:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_10_str

# Função para calcular o custo empresa total com acréscimo de 15%
def calcular_media_total_15(conn):
    """
    Calcula o custo empresa total (média ponderada) com acréscimo de 15%.
    
    Divide o custo empresa ponderado por 0,85.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: String formatada representando o custo com acréscimo de 15%.
    """
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
    
    media_15 = media_ponderada_empresa_float / 0.85  # Aplica acréscimo de 15%
    media_15_str = f"R$ {media_15:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_15_str

# Função para calcular o custo empresa total com acréscimo de 20%
def calcular_media_total_20(conn):
    """
    Calcula o custo empresa total (média ponderada) com acréscimo de 20%.
    
    Divide o custo empresa ponderado por 0,8.
    
    :param conn: Conexão ativa com o banco de dados.
    :return: String formatada representando o custo com acréscimo de 20%.
    """
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
    
    media_20 = media_ponderada_empresa_float / 0.8  # Aplica acréscimo de 20%
    media_20_str = f"R$ {media_20:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_20_str

# Função para voltar ao menu
def voltar_para_menu(janela_media_custo, main_window):
    """Fecha a janela de média de custo e reexibe o menu principal com atualização visual."""
    janela_media_custo.destroy()         # Fecha a janela atual
    main_window.deiconify()              # Reexibe a janela principal
    main_window.state("zoomed")          # Maximiza a janela principal
    main_window.lift()                   # Garante que fique no topo
    main_window.update()                 # Força atualização visual
   
# Função para criar a janela de média de custo
def criar_media_custo(font_size=12, main_window=None):
    user_id = main_window.user_id  # recupera o user_id do main_window
    conn = conectar_banco()  # Conecta ao banco de dados aqui
    if not conn:
        print("Erro ao conectar ao banco de dados.")
        return

    # Criando a janela Toplevel (não usamos mais o Tk())
    janela_media_custo = tk.Toplevel()
    janela_media_custo.title("Tabela de Produtos")
    janela_media_custo.geometry("800x600")
    janela_media_custo.state("zoomed")

    caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
    aplicar_icone(janela_media_custo, caminho_icone)

    # Criando o notebook (componentes de abas)
    notebook = ttk.Notebook(janela_media_custo)
    notebook.pack(expand=True, fill="both")

    aba_icm = tk.Frame(notebook)
    aba_sudeste = tk.Frame(notebook)
    aba_centro_oeste_nordeste = tk.Frame(notebook)

    notebook.add(aba_icm, text="ICM 18%")
    notebook.add(aba_sudeste, text="Região Sudeste ICM 12%")
    notebook.add(aba_centro_oeste_nordeste, text="Centro-Oeste e Nordeste 7%")

    # Função para carregar os dados ao selecionar a aba
    def carregar_dados(event):
        aba_atual = notebook.index(notebook.select())  # Obtém o índice da aba ativa
        conn = conectar_banco()  # Conecta ao banco
        
        if conn:
            if aba_atual == 1:  # Sudeste 12%
                for widget in aba_sudeste.winfo_children():
                    widget.destroy()
                criar_media_custo_sudeste(aba_sudeste, conn)

            elif aba_atual == 2:  # Centro-Oeste/Nordeste 7%
                for widget in aba_centro_oeste_nordeste.winfo_children():
                    widget.destroy()
                criar_media_aba_centro_oeste(aba_centro_oeste_nordeste, conn)

        conn.close()  # Fecha a conexão

    # Vincular o evento ao Notebook
    notebook.bind("<<NotebookTabChanged>>", carregar_dados)

    frame_tabela = tk.Frame(aba_icm)
    frame_tabela.pack(expand=True, fill="both", padx=10, pady=10)

    tree = ttk.Treeview(frame_tabela, columns=("Produto", "Estoque", "Custo Estoque", "Custo Empresa", "Custo 5%", "Custo 10%", "Custo 15%", "Custo 20%"), show="headings", style="Custom.Treeview" )
    tree.heading("Produto", text="Nome do Produto")
    tree.heading("Estoque", text="Estoque")
    tree.heading("Custo Estoque", text="Média de Custo do Estoque")
    tree.heading("Custo Empresa", text="Média Custo Empresa")
    tree.heading("Custo 5%", text="Custo com 5%")
    tree.heading("Custo 10%", text="Custo com 10%")
    tree.heading("Custo 15%", text="Custo com 15%")
    tree.heading("Custo 20%", text="Custo com 20%")
    
    tree.column("Produto", anchor="center")
    tree.column("Estoque", anchor="center")
    tree.column("Custo Estoque", anchor="center")
    tree.column("Custo Empresa", anchor="center")
    tree.column("Custo 5%", anchor="center")
    tree.column("Custo 10%", anchor="center")
    tree.column("Custo 15%", anchor="center")
    tree.column("Custo 20%", anchor="center")
    
   # Estilo personalizado
    style = ttk.Style()
    style.theme_use("alt")

    # Configurações para a Treeview
    style.configure("Treeview", 
                    background="white", 
                    foreground="black", 
                    rowheight=25,  # Ajuste de altura da linha
                    fieldbackground="white")

    style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
    style.configure("Treeview", rowheight=30)
    style.configure("Custom.Treeview", font=("Arial", 10), rowheight=30)

    # Mapeando a cor de fundo e de texto quando uma linha é selecionada
    style.map("Treeview", 
            background=[("selected", "#0078D7")], 
            foreground=[("selected", "white")])

    # Configurações para o botão com o estilo "Estoque.TButton"
    style.configure("Estoque.TButton",
                    padding=(5, 2),
                    relief="raised",  # Estilo da borda
                    background="#D3D3D3",  # Cor de fundo cinza claro
                    foreground="black",  # Cor do texto preta
                    font=("Arial", 10),
                    borderwidth=2,  # Largura da borda
                    highlightbackground="black",  # Cor da borda em foco
                    highlightthickness=1)  # Espessura do contorno de foco

    # Mapeando o estilo para os estados ativo e desabilitado do botão
    style.map("Estoque.TButton",
            background=[("active", "#C0C0C0")],  # Cor de fundo cinza médio quando ativo
            foreground=[("active", "black")],  # Cor do texto preta no estado ativo
            relief=[("pressed", "sunken"), ("!pressed", "raised")])  # Mudança na borda ao pressionar

    scrollbar_vertical = tk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
    scrollbar_horizontal = tk.Scrollbar(frame_tabela, orient="horizontal", command=tree.xview)
    tree.config(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
    
    tree.grid(row=0, column=0, sticky="nsew")
    scrollbar_vertical.grid(row=0, column=1, sticky="ns")
    scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

    frame_tabela.grid_rowconfigure(0, weight=1)
    frame_tabela.grid_columnconfigure(0, weight=1)

    # Inserindo os dados de produtos na tabela
    produtos = buscar_produtos(conn)
    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)
        
        # Calcula a Média de Custo do Estoque
        media_custo_estoque_str = calcular_media_ponderada(conn, nome_produto)
        
        # Converte a string de preço para float para realizar cálculos
        media_custo_estoque = float(media_custo_estoque_str.replace("R$ ", "").replace(".", "").replace(",", "."))

        # Calcula o Custo empresa
        custo_empresa = calcular_custo_empresa(conn, nome_produto, media_custo_estoque)

        # Aplicando 5%
        custo_5 = calcular_custo_5(custo_empresa, 5)

        # Aplicando 10%
        custo_10 = calcular_custo_10(custo_empresa)

        # Aplicando 15%
        custo_15 = calcular_custo_15(custo_empresa)

        # Aplicando 20%
        custo_20 = calcular_custo_20(custo_empresa) 

        # Define a tag como vazia por padrão
        tag_cor = ""

        # Se o estoque for "0,000 kg" ou vazio, definir a tag para "estoque_vazio"
        if estoque.strip() == "0,000 kg" or not estoque.strip():
            tag_cor = "estoque_vazio"

        # Inserindo os valores na tabela e aplicando a tag correspondente
        item_id = tree.insert("", "end", values=(nome_produto, estoque, media_custo_estoque_str, custo_empresa, custo_5, custo_10, custo_15, custo_20), tags=(tag_cor,))

    total_estoque = calcular_total_estoque(conn)
    media_ponderada_total, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)
    media_5 = calcular_media_total_5(conn)
    media_10 = calcular_media_total_10(conn)
    media_15 = calcular_media_total_15(conn)
    media_20 = calcular_media_total_20(conn)

    # Configuração da aparência do total
    tree.tag_configure("total", background="yellow", font=("Arial", 12, "bold"), foreground="red")


    # Inserindo a linha "Total" sempre no final
    tree.insert("", "end", values=("TOTAL", total_estoque, media_ponderada_total, media_ponderada_empresa, media_5, media_10, media_15, media_20,  "", "", "", "", "", ""), tags=("total",))

    # Aqui você insere o label com a frase desejada em vermelho
    observacao_label = tk.Label(janela_media_custo, text="Observação: Acrescentar 3,25%", fg="red", font=("Arial", 12, "bold"))
    observacao_label.pack(pady=10)

    # Configuração da tag para cor vermelha
    tree.tag_configure("estoque_vazio", foreground="red")

    janela_media_custo.option_add("*TButton*background", "#D3D3D3")
    janela_media_custo.option_add("*TButton*foreground", "black")
    janela_media_custo.option_add("*TButton*font", ("Arial", 10))

    # Cria um frame para os botões (sem expandir para toda a largura)
    frame_botoes = tk.Frame(janela_media_custo, bg="#ecf0f1")
    frame_botoes.pack(pady=10)

    # Define uma largura fixa (em número de caracteres) para os botões
    btn_voltar = ttk.Button(frame_botoes, text="Voltar", style="Estoque.TButton",
                            command=lambda: voltar_para_menu(janela_media_custo, main_window), width=20)
    btn_exportar = ttk.Button(frame_botoes, text="Exportar para Excel", style="Estoque.TButton",
                            command=lambda: exportar_notebook_para_excel(notebook), width=20)

    # Posiciona os botões lado a lado sem expandir
    btn_voltar.grid(row=0, column=0, padx=10, pady=5)
    btn_exportar.grid(row=0, column=1, padx=10, pady=5)

    def on_closing():
        """Fecha a janela e encerra o programa corretamente"""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Fecha a conexão com o banco de dados, se estiver aberta
            if conn:
                conn.close()
                print("Conexão com o banco de dados fechada.")

            janela_media_custo.destroy()  # Fecha apenas esta janela
            if main_window:  # Se houver uma janela principal, pode fechá-la também
                main_window.destroy()
            
            sys.exit(0)  # Encerra completamente o programa (opcional)
    
    janela_media_custo.protocol("WM_DELETE_WINDOW", on_closing)

    janela_media_custo.mainloop()
