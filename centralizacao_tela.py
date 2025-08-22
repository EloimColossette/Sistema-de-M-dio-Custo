# Função para centralizar janelas pequena
def centralizar_janela(janela, largura, altura):
    # Obtém a largura e altura da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcula a posição x e y para centralizar a janela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Define a geometria da janela
    janela.geometry(f'{largura}x{altura}+{x}+{y}')
# Função para criar a janela em tela cheia
def centralizar_janela_tela_cheia(janela, largura, altura):
    # Obtém a largura e altura da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcula a posição x e y para centralizar a janela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Define a geometria da janela
    janela.geometry(f'{largura}x{altura}+{x}+{y}')

def centralizar_janela2(janela):
      # Atualiza a janela para calcular suas dimensões
    janela.update_idletasks()
    
    # Obtém a largura e altura da janela
    largura = janela.winfo_width()
    altura = janela.winfo_height()

    # Obtém a largura e altura da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcula a posição x e y para centralizar a janela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Define a geometria da janela
    janela.geometry(f'+{x}+{y}')