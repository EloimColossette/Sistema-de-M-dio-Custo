from conexao_db import conectar
import unicodedata
import pandas as pd
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from tkinter import messagebox, filedialog
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import re
from openpyxl import load_workbook

# Nome padrão para a janela
NOME_JANELA = "Sistema Kametal"

# Registrar a fonte Arial
pdfmetrics.registerFont(TTFont('Arial', 'C:/Windows/Fonts/arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold', 'C:/Windows/Fonts/arialbd.ttf'))

# Função para deixar o texto no formato utf-8 para pdf
def normalizar_texto(texto):
    """Normaliza o texto mantendo os caracteres acentuados."""
    return unicodedata.normalize('NFKC', texto)

# Função para limpar espaços e outras anomalias no pdf
def limpar_texto(texto):
    """Limpa o texto removendo quebras de linha, tabs e espaços extras."""
    texto_limpo = re.sub(r'[\n\r\t]+', ' ', texto)
    return texto_limpo.strip()

# Função para exportar em PDF
def exportar_para_pdf(caminho_arquivo, tabela, colunas, titulo):
    """Exporta os dados da base para um arquivo PDF"""
    # Conecta usando a função 'conectar' da biblioteca 'conexao_db'
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(f"SELECT * FROM {tabela}")
    dados = cursor.fetchall()
    conexao.close()

    # Remove a coluna ID da lista de dados (assumindo que seja a primeira coluna)
    dados = [linha[1:] for linha in dados]
    # Ordena os dados pela primeira coluna
    dados = sorted(dados, key=lambda x: x[0])

    # Exemplo usando Platypus em A4 — ajuste conforme desejar:
    doc       = SimpleDocTemplate(caminho_arquivo, pagesize=A4)
    elementos = []
    estilos   = getSampleStyleSheet()

    # Título
    elementos.append(Paragraph(f"<b>{titulo}</b>", estilos["Title"]))
    elementos.append(Spacer(1, 20))

    # Cabeçalhos
    dados_tabela = [colunas] + [[str(v) for v in row] for row in dados]
    tabela_pdf   = Table(dados_tabela, hAlign="CENTER")
    tabela_pdf.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elementos.append(tabela_pdf)

    # Gera o PDF
    doc.build(elementos)

    # Exibe mensagem de sucesso
    messagebox.showinfo("Exportação concluída", f"PDF exportado com sucesso para:\n{caminho_arquivo}")

# Função para exportar em Excel
def exportar_para_excel(caminho_arquivo, tabela, colunas):
    """Exporta os dados da base para um arquivo Excel"""
    # Conecta usando a função 'conectar' da biblioteca 'conexao_db'
    conexao = conectar()
    cursor = conexao.cursor()
    cursor.execute(f"SELECT * FROM {tabela}")  # Consulta baseada na tabela fornecida
    dados = cursor.fetchall()
    conexao.close()

    # Remove a coluna ID da lista de dados
    dados = [linha[1:] for linha in dados]

    # Verifica se a quantidade de colunas nos dados corresponde à quantidade esperada
    num_colunas_dados = len(dados[0]) if dados else 0
    num_colunas_spec = len(colunas)

    if num_colunas_dados != num_colunas_spec:
        raise ValueError(f"Number of columns in data ({num_colunas_dados}) does not match the number of columns specified ({num_colunas_spec})")

    # Criando o DataFrame e salvando o arquivo Excel
    df = pd.DataFrame(dados, columns=colunas)
    df.to_excel(caminho_arquivo, index=False)
    messagebox.showinfo("Exportação concluída", f"Arquivo Excel salvo em {caminho_arquivo}")

# Função para buscar um widget do tipo Treeview recursivamente
def find_treeview(widget):
    if widget.winfo_class() == "Treeview":
        return widget
    for child in widget.winfo_children():
        result = find_treeview(child)
        if result is not None:
            return result
    return None

# Função para exportar dados do Treeview para Excel
def exportar_notebook_para_excel(notebook):
    file_path = filedialog.asksaveasfilename(
         defaultextension=".xlsx",
         filetypes=[("Excel files", "*.xlsx")],
         title="Salvar como",
         initialfile="relatorio_medio_custo.xlsx"
    )
    if not file_path:   
         return

    active_tab = notebook.select()
    aba = notebook.nametowidget(active_tab)
    tree = find_treeview(aba)
    if tree is None:
         messagebox.showerror("Exportação", "Nenhum Treeview encontrado na aba ativa.")
         return

    col_ids = list(tree["columns"])
    col_names = [tree.heading(col)["text"].strip() if tree.heading(col)["text"].strip() != "" else col 
                 for col in col_ids]

    dados = []
    for item in tree.get_children():
         row = list(tree.item(item)["values"])
         if len(row) < len(col_ids):
              row.extend([""] * (len(col_ids) - len(row)))
         elif len(row) > len(col_ids):
              row = row[:len(col_ids)]
         dados.append(row)

    try:
         df = pd.DataFrame(dados, columns=col_names)
         df.to_excel(file_path, index=False)
         messagebox.showinfo("Exportação", f"Dados exportados com sucesso para Excel em:\n{file_path}")
    except Exception as e:
         messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")