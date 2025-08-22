# utilitarios.py
from PIL import Image, ImageTk
import tkinter as tk

# Função para colocar logotipo
def carregar_logotipo(janela, caminho_imagem, largura=300, altura=100, cor_fundo="#FFA500"):
    """
    Função para carregar e exibir o logotipo em uma janela.

    :param janela: A janela ou frame onde o logotipo será exibido.
    :param caminho_imagem: Caminho para a imagem do logotipo.
    :param largura: Largura desejada da imagem.
    :param altura: Altura desejada da imagem.
    :param cor_fundo: Cor de fundo onde o logotipo será exibido.
    """
    try:
        imagem_logo = Image.open(caminho_imagem)
        imagem_logo = imagem_logo.resize((largura, altura))  # Ajustar o tamanho conforme necessário
        imagem_tk = ImageTk.PhotoImage(imagem_logo)
        logo_label = tk.Label(janela, image=imagem_tk, bg=cor_fundo)
        logo_label.image = imagem_tk  # Necessário para manter a referência da imagem
        logo_label.pack(side="top", pady=10)  # Posicionar o logotipo no topo da janela
        print("Logotipo carregado com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar o logotipo: {e}")

# Função para clocar imagens 
def definir_icone(janela, caminho_icone):
    """
    Função para alterar o ícone padrão da janela do Tkinter (substituir a "pena").

    :param janela: A janela Tkinter onde o ícone será alterado.
    :param caminho_icone: Caminho para a imagem do ícone (formato .ico é recomendado).
    """
    try:
        icone = tk.PhotoImage(file=caminho_icone)
        janela.iconphoto(False, icone)
        print("Ícone alterado com sucesso.")
    except Exception as e:
        print(f"Erro ao definir o ícone: {e}")

# Função para aplicação do logotipo
def aplicar_icone(janela, caminho_icone):
    """ Aplica o ícone à janela fornecida. """
    try:
        janela.iconbitmap(caminho_icone)
    except Exception as e:
        print(f"Erro ao definir o ícone: {e}")