import tkinter as tk
from tkinter import ttk
from conexao_db import conectar
from decimal import Decimal, ROUND_HALF_UP

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
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT nome FROM produtos ORDER BY nome;")
        produtos = cursor.fetchall()
        cursor.close()
        return produtos
    except Exception as e:
        print("Erro ao buscar nome_produtos:", e)
        return []

# Função para buscar e somar as quantidades de estoque
def buscar_estoque(conn, nome_produto):
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT SUM(eq.quantidade_estoque)
            FROM estoque_quantidade eq
            JOIN somar_produtos sp ON eq.id_produto = sp.id
            WHERE sp.produto = %s
        """, (nome_produto,))
        
        quantidade = cursor.fetchone()[0]
        cursor.close()

        if quantidade:
            return f"{quantidade:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"
        else:
            return "0,000 kg"  
    except Exception as e:
        print(f"Erro ao buscar estoque para o produto {nome_produto}: {e}")
        return "0,000 kg"
    
# Função para calcular a média ponderada do custo do estoque
def calcular_media_ponderada(conn, nome_produto):
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT SUM(eq.quantidade_estoque * eq.custo_total), SUM(eq.quantidade_estoque)
            FROM estoque_quantidade eq
            JOIN somar_produtos sp ON eq.id_produto = sp.id
            WHERE sp.produto = %s
        """, (nome_produto,))

        resultado = cursor.fetchone()
        cursor.close()

        if resultado:
            # Desembrulha os valores da tupla diretamente
            soma_custo_ponderado, soma_quantidade = resultado

            # Converte os valores Decimal para float antes de realizar os cálculos
            soma_custo_ponderado = float(soma_custo_ponderado)
            soma_quantidade = float(soma_quantidade)

            # Verifica se soma_quantidade é maior que 0 para evitar divisão por zero
            if soma_quantidade > 0:
                media_ponderada = soma_custo_ponderado / soma_quantidade
                
                # Aplicando o cálculo adicional
                custo_calculado = (media_ponderada * 0.7275) / 0.8375

                # Retornando o valor formatado com "R$"
                return "R$ " + f"{custo_calculado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                return "R$ 0,00"  # Evita divisão por zero
        else:
            print(f"Erro: A consulta retornou None para {nome_produto}")
            return "R$ 0,00"  # Retorna 0 se a consulta não retornar nada
    except Exception as e:
        return "R$ 0,00"
    
# Função para calcular a média ponderada do custo do estoque mais a mão de obra da empresa
def calcular_custo_empresa(conn, nome_produto):
    try:
        cursor = conn.cursor()

        # Obtendo os dados de quantidade_estoque e custo_total
        cursor.execute("""
            SELECT eq.quantidade_estoque, eq.custo_total
            FROM estoque_quantidade eq
            JOIN somar_produtos sp ON eq.id_produto = sp.id
            WHERE sp.produto = %s
        """, (nome_produto,))

        produtos = cursor.fetchall()
        
        # Somando para calcular a média ponderada
        soma_custo_ponderado = 0
        soma_quantidade = 0

        for produto in produtos:
            peso = produto[0]  # quantidade_estoque
            custo = produto[1]  # custo_total
            
            # Verifica se os valores não são None
            if peso is not None and custo is not None:
                soma_custo_ponderado += peso * custo
                soma_quantidade += peso

        # Calculando a média ponderada
        if soma_quantidade > 0:
            media_ponderada = soma_custo_ponderado / soma_quantidade
        else:
            media_ponderada = 0  # Caso não haja produtos válidos

        # Se a média ponderada for 0, ignora o custo Empresa e retorna 0
        if media_ponderada == 0:
            cursor.close()
            return "R$ 0,00"
        
        # Obtendo o custo Empresa cadastrado na empresa
        cursor.execute("""
            SELECT custo_empresa FROM somar_produtos WHERE produto = %s
        """, (nome_produto,))
        custo_empresa = cursor.fetchone()
        cursor.close()

        if custo_empresa and custo_empresa[0] is not None:
            custo_empresa = float(custo_empresa[0])
            media_ponderada = float(media_ponderada)
            
            # Soma o custo Empresa à média ponderada e aplica o cálculo adicional
            resultado = media_ponderada + custo_empresa
            resultado_final = (resultado * 0.7275) / 0.8375

            return "R$ " + f"{resultado_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return "R$ 0,00"
    except Exception as e:
        return "R$ 0,00"

# Função para calcular a media ponderada + mão de obra da empresa e + 5%
def calcular_custo_5(custo_base, percentual):
    """
    Calcula o custo base com acréscimo de um percentual específico.

    :param custo_base: Valor base (float) ao qual será aplicado o acréscimo.
    :param percentual: Percentual de acréscimo (ex: 5 para 5%).
    :return: Valor final formatado como string no formato monetário.
    """
    try:
        if isinstance(custo_base, str):  # Caso receba string, converte para float
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.95  # Aplicando o acréscimo
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de {percentual}%: {e}")
        return "R$ 0,00"
    
# Função para calcular a media ponderada + mão de obra da empresa e + 10%
def calcular_custo_10(custo_base):
    """
    Calcula o custo base com acréscimo de 10%.

    :param custo_base: Valor base (float) ao qual será aplicado o acréscimo.
    :return: Valor final formatado como string no formato monetário.
    """
    try:
        if isinstance(custo_base, str):  # Caso receba string, converte para float
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.9  # Aplicando o acréscimo de 10%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 10%: {e}")
        return "R$ 0,00"

# Função para calcular a media ponderada + mão de obra da empresa e + 15%
def calcular_custo_15(custo_base):
    """
    Calcula o custo base com acréscimo de 15%.

    :param custo_base: Valor base (float) ao qual será aplicado o acréscimo.
    :return: Valor final formatado como string no formato monetário.
    """
    try:
        if isinstance(custo_base, str):  # Caso receba string, converte para float
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.85  # Aplicando o acréscimo de 15%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 15%: {e}")
        return "R$ 0,00"
    
# Função para calcular a media ponderada + mão de obra da empresa e + 20%
def calcular_custo_20(custo_base):
    """
    Calcula o custo base com acréscimo de 20%.

    :param custo_base: Valor base (float) ao qual será aplicado o acréscimo.
    :return: Valor final formatado como string no formato monetário.
    """
    try:
        if isinstance(custo_base, str):  # Caso receba string, converte para float
            custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())

        custo_final = custo_base / 0.8  # Aplicando o acréscimo de 20%
        return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    except Exception as e:
        print(f"Erro ao calcular acréscimo de 20%: {e}")
        return "R$ 0,00"

def calcular_total_estoque(conn):
    """Calcula o total do estoque somando todos os produtos."""
    total_estoque = 0.0
    produtos = buscar_produtos(conn)

    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)

        # Converte o estoque para float, tratando separadores e unidade " kg"
        estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))

        # Acumula o total
        total_estoque += estoque_float

    # Retorna o total formatado com três casas decimais e " kg"
    return f"{total_estoque:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"

def calcular_media_ponderada_total_com_empresa(conn):
    """Calcula a média ponderada de custo do estoque total e o custo empresa ponderado."""
    produtos = buscar_produtos(conn)

    soma_ponderada_custo = 0.0
    soma_ponderada_empresa = 0.0
    soma_estoque = 0.0

    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)

        # Converte o estoque para float
        estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))

        if estoque_float > 0:
            # Obtém a média de custo do produto e converte para float
            media_custo_str = calcular_media_ponderada(conn, nome_produto)
            media_custo = float(media_custo_str.replace("R$ ", "").replace(".", "").replace(",", "."))
            
            # Se a média de custo for 0, ignora o custo empresa para esse produto
            if media_custo == 0:
                custo_empresa = 0.0
            else:
                custo_empresa_str = calcular_custo_empresa(conn, nome_produto)
                custo_empresa = float(custo_empresa_str.replace("R$ ", "").replace(".", "").replace(",", "."))
            
            # Acumula os valores ponderados (valor * quantidade)
            soma_ponderada_custo += estoque_float * media_custo
            soma_ponderada_empresa += estoque_float * custo_empresa
            soma_estoque += estoque_float

    # Evita divisão por zero
    if soma_estoque == 0:
        return "R$ 0,000", "R$ 0,000"

    # Calcula as médias ponderadas
    media_ponderada_total = soma_ponderada_custo / soma_estoque
    media_ponderada_empresa = soma_ponderada_empresa / soma_estoque

    # Retorna os valores formatados
    media_ponderada_total_str = f"R$ {media_ponderada_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    media_ponderada_empresa_str = f"R$ {media_ponderada_empresa:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
    return media_ponderada_total_str, media_ponderada_empresa_str
    
def calcular_media_total_5(conn):
    """Calcula o custo empresa dividido por 0,95 e retorna o valor formatado como moeda."""
    # Obtém o custo empresa ponderado
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)

    # Converte o custo empresa para float removendo o "R$ " e substituindo as vírgulas por pontos
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))

    # Realiza o cálculo dividindo o custo empresa por 0,95
    media_5 = media_ponderada_empresa_float / 0.95

    # Formatação para moeda
    media_5_str = f"R$ {media_5:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return media_5_str  # Retorna o valor já formatado como moeda

def calcular_media_total_10(conn):
    """Calcula o custo empresa dividido por 0,9 e retorna o valor formatado como moeda."""
    # Obtém o custo empresa ponderado
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)

    # Converte o custo empresa para float
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))

    # Realiza o cálculo dividindo o custo empresa por 0,9
    media_10 = media_ponderada_empresa_float / 0.9

    # Formatação para moeda
    media_10_str = f"R$ {media_10:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return media_10_str  # Retorna o valor já formatado como moeda

def calcular_media_total_15(conn):
    """Calcula o custo empresa dividido por 0,85 e retorna o valor formatado como moeda."""
    # Obtém o custo empresa ponderado
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)

    # Converte o custo empresa para float
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))

    # Realiza o cálculo dividindo o custo empresa por 0,85
    media_15 = media_ponderada_empresa_float / 0.85

    # Formatação para moeda
    media_15_str = f"R$ {media_15:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return media_15_str  # Retorna o valor já formatado como moeda

def calcular_media_total_20(conn):
    """Calcula o custo empresa dividido por 0,8 e retorna o valor formatado como moeda."""
    # Obtém o custo empresa ponderado
    _, media_ponderada_empresa = calcular_media_ponderada_total_com_empresa(conn)

    # Converte o custo empresa para float
    media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))

    # Realiza o cálculo dividindo o custo empresa por 0,8
    media_20 = media_ponderada_empresa_float / 0.8

    # Formatação para moeda
    media_20_str = f"R$ {media_20:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return media_20_str  # Retorna o valor já formatado como moeda

def criar_media_aba_centro_oeste(aba_centro_oeste, conn):
    # Criação da frame onde a tabela vai ficar
    frame_tabela = tk.Frame(aba_centro_oeste)
    frame_tabela.pack(expand=True, fill="both", padx=10, pady=10)

    # Criação da treeview para a tabela
    tree = ttk.Treeview(frame_tabela, columns=("Produto", "Estoque", "Custo Estoque", "Custo empresa", "Custo 5%", "Custo 10%", "Custo 15%", "Custo 20%"), show="headings", style="Custom.Treeview")
    tree.heading("Produto", text="Nome do Produto")
    tree.heading("Estoque", text="Estoque")
    tree.heading("Custo Estoque", text="Média de Custo do Estoque")
    tree.heading("Custo empresa", text="Média Custo Empresa")
    tree.heading("Custo 5%", text="Custo com 5%")
    tree.heading("Custo 10%", text="Custo com 10%")
    tree.heading("Custo 15%", text="Custo com 15%")
    tree.heading("Custo 20%", text="Custo com 20%")

    tree.column("Produto", anchor="center")
    tree.column("Estoque", anchor="center")
    tree.column("Custo Estoque", anchor="center")
    tree.column("Custo empresa", anchor="center")
    tree.column("Custo 5%", anchor="center")
    tree.column("Custo 10%", anchor="center")
    tree.column("Custo 15%", anchor="center")
    tree.column("Custo 20%", anchor="center")

    # Estilo da tabela
    style = ttk.Style()
    style.theme_use("alt")
    style.configure("Treeview", 
                    background="white", 
                    foreground="black", 
                    rowheight=25,  
                    fieldbackground="white")
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
    style.configure("Treeview", rowheight=30)
    style.configure("Custom.Treeview", font=("Arial", 10), rowheight=30)
    style.map("Treeview", 
              background=[("selected", "#0078D7")], 
              foreground=[("selected", "white")])

    scrollbar_vertical = tk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
    scrollbar_horizontal = tk.Scrollbar(frame_tabela, orient="horizontal", command=tree.xview)
    tree.config(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
    
    tree.grid(row=0, column=0, sticky="nsew")
    scrollbar_vertical.grid(row=0, column=1, sticky="ns")
    scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

    frame_tabela.grid_rowconfigure(0, weight=1)
    frame_tabela.grid_columnconfigure(0, weight=1)

    # # Inserindo os dados de produtos na tabela
    produtos = buscar_produtos(conn)
    for produto in produtos:
        nome_produto = produto[0]
        estoque = buscar_estoque(conn, nome_produto)
        
        # Calcula a Média de Custo do Estoque
        media_custo_estoque_str = calcular_media_ponderada(conn, nome_produto)
        
        media_custo_estoque = float(media_custo_estoque_str.replace("R$ ", "").replace(".", "").replace(",", "."))

        # Calculando os custos
        custo_empresa = calcular_custo_empresa(conn, nome_produto)  # Passando apenas os dois argumentos necessários
        custo_5 = calcular_custo_5(custo_empresa, 5)
        custo_10 = calcular_custo_10(custo_empresa)
        custo_15 = calcular_custo_15(custo_empresa)
        custo_20 = calcular_custo_20(custo_empresa)

        tag_cor = ""
        if estoque.strip() == "0,000 kg" or not estoque.strip():
            tag_cor = "estoque_vazio"

        tree.insert("", "end", values=(nome_produto, estoque, media_custo_estoque_str, custo_empresa, custo_5, custo_10, custo_15, custo_20), tags=(tag_cor,))

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

    tree.tag_configure("estoque_vazio", foreground="red")

    aba_centro_oeste.option_add("*TButton*background", "#D3D3D3")
    aba_centro_oeste.option_add("*TButton*foreground", "black")
    aba_centro_oeste.option_add("*TButton*font", ("Arial", 10))
