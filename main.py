import time
from login import TelaLogin
import tkinter as tk
import sys
from centralizacao_tela import centralizar_janela

def main():
    inicio = time.time()
    print("Iniciando a aplicação...")

    print("Criando a splash screen...")
    splash = tk.Tk()
    splash.overrideredirect(True)

    splash_width = 400
    splash_height = 300
    centralizar_janela(splash, splash_width, splash_height)
    splash.configure(background='#2C3E50')

    frame = tk.Frame(splash, background='#2C3E50')
    frame.pack(expand=True, fill='both')

    label = tk.Label(frame, text="Carregando Sistema Kametal...", font=("Helvetica", 18, "bold"), fg='white', background='#2C3E50')
    label.pack(expand=True)
    print("Splash screen exibida.")

    # Reduzido para 2000ms (2 segundos)
    splash.after(2000, splash.destroy)
    splash.mainloop()
    print("Splash finalizada.")

    print("Iniciando tela de login...")
    app = TelaLogin()
    # Registra a impressão do tempo de inicialização após a criação da janela
    app.janela_login.after_idle(lambda: print("Tempo de inicialização: {:.2f} segundos".format(time.time() - inicio)))
    app.run()

    sys.exit(0)

if __name__ == "__main__":
    main()