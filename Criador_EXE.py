import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import threading
import queue
import sys
import shutil

def resource_path(relative_path):
    """Obtém o caminho absoluto para um recurso, funcionando tanto em desenvolvimento quanto no bundle."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Configure as variáveis de ambiente para que o Tkinter encontre os arquivos do Tcl/Tk
if hasattr(sys, '_MEIPASS'):
    # Se o aplicativo estiver empacotado, os dados foram extraídos para sys._MEIPASS.
    print("sys._MEIPASS:", sys._MEIPASS)
    tcl_path = os.path.join(sys._MEIPASS, 'tcl', 'tcl8.6')
    tk_path = os.path.join(sys._MEIPASS, 'tk', 'tk8.6')
    print("TCL_LIBRARY:", tcl_path)
    print("TK_LIBRARY:", tk_path)
    os.environ['TCL_LIBRARY'] = tcl_path
    os.environ['TK_LIBRARY'] = tk_path
else:
    # Em modo de desenvolvimento, utiliza os caminhos da instalação local do Python.
    os.environ['TCL_LIBRARY'] = r'C:\Users\Kmt-02\AppData\Local\Programs\Python\Python313\tcl\tcl8.6'
    os.environ['TK_LIBRARY'] = r'C:\Users\Kmt-02\AppData\Local\Programs\Python\Python313\tcl\tk8.6'

def centralizar_janela(janela, largura, altura):
    """Centraliza a janela na tela."""
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

class ConstrutorDeExe(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ExeMaster")
        
        # Define tamanho e centraliza a janela
        largura, altura = 800, 500
        centralizar_janela(self, largura, altura)
        self.configure(bg="#2E3F4F")
        self.resizable(False, False)

        # Define o ícone da janela
        try:
            self.iconbitmap("C:\\Sistema\\logos\\engrenagens.ico")
        except Exception as e:
            print("Não foi possível definir o ícone:", e)

        # Fila para mensagens das threads
        self.queue = queue.Queue()

        self.criar_componentes()
        # Inicia o processamento da fila
        self.processar_queue()

    def criar_componentes(self):
        """Cria e organiza os componentes da interface."""
        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure("TFrame", background="#2E3F4F")
        estilo.configure("TLabel", background="#2E3F4F", foreground="white", font=("Helvetica", 12))
        estilo.configure("TButton", background="#1ABC9C", foreground="black", font=("Helvetica", 10, "bold"))
        estilo.configure("TRadiobutton", background="#2E3F4F", foreground="white", font=("Helvetica", 10))

        # Título
        rotulo_titulo = ttk.Label(self, text="ExeMaster", font=("Helvetica", 18, "bold"))
        rotulo_titulo.pack(pady=10)

        # Frame principal
        quadro_principal = ttk.Frame(self)
        quadro_principal.pack(pady=10, padx=20, fill=tk.X)

        # Nome do Sistema
        ttk.Label(quadro_principal, text="Nome do Sistema:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_nome = ttk.Entry(quadro_principal, width=50)
        self.entrada_nome.grid(row=0, column=1, padx=5, pady=5)

        # Caminho do Ícone
        ttk.Label(quadro_principal, text="Ícone (.ico):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_icone = ttk.Entry(quadro_principal, width=50)
        self.entrada_icone.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(quadro_principal, text="Selecionar", command=self.selecionar_icone).grid(row=1, column=2, padx=5, pady=5)

        # Caminho do Arquivo Principal
        ttk.Label(quadro_principal, text="Arquivo Principal (.py):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_principal = ttk.Entry(quadro_principal, width=50)
        self.entrada_principal.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(quadro_principal, text="Selecionar", command=self.selecionar_principal).grid(row=2, column=2, padx=5, pady=5)

        # Caminho do Diretório de Saída
        ttk.Label(quadro_principal, text="Diretório de Saída:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_saida = ttk.Entry(quadro_principal, width=50)
        self.entrada_saida.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(quadro_principal, text="Selecionar", command=self.selecionar_saida).grid(row=3, column=2, padx=5, pady=5)

        # Tipo de Build (OneDir ou OneFile) - ajustando o grid para nova linha
        ttk.Label(quadro_principal, text="Tipo de Build:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.tipo_build = tk.StringVar(value="onedir")
        quadro_opcoes = ttk.Frame(quadro_principal)
        quadro_opcoes.grid(row=4, column=1, sticky=tk.W)
        ttk.Radiobutton(quadro_opcoes, text="OneDir", variable=self.tipo_build, value="onedir", style="TRadiobutton").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(quadro_opcoes, text="OneFile", variable=self.tipo_build, value="onefile", style="TRadiobutton").pack(side=tk.LEFT, padx=10)

        # Botão Criar EXE
        botao_criar = ttk.Button(self, text="Criar EXE", command=self.iniciar_criacao_exe)
        botao_criar.pack(pady=15)

        # Label de Status (para mostrar a última mensagem)
        self.status_label = ttk.Label(self, text="Status: Aguardando início...", font=("Helvetica", 10))
        self.status_label.pack(pady=5)

        # Área de saída
        quadro_saida = ttk.Frame(self)
        quadro_saida.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.texto_saida = tk.Text(quadro_saida, height=8, wrap=tk.WORD, state=tk.DISABLED, bg="#1C2833", fg="white")
        self.texto_saida.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Barra de rolagem (opcional)
        barra_rolagem = ttk.Scrollbar(quadro_saida, orient=tk.VERTICAL, command=self.texto_saida.yview)
        barra_rolagem.pack(side=tk.RIGHT, fill=tk.Y)
        self.texto_saida.config(yscrollcommand=barra_rolagem.set)

    def selecionar_icone(self):
        """Abre um diálogo para seleção do ícone."""
        caminho = filedialog.askopenfilename(title="Selecione o ícone", filetypes=[("Arquivos ICO", "*.ico")])
        if caminho:
            self.entrada_icone.delete(0, tk.END)
            self.entrada_icone.insert(0, caminho)

    def selecionar_principal(self):
        """Abre um diálogo para seleção do arquivo principal."""
        caminho = filedialog.askopenfilename(title="Selecione o arquivo Python", filetypes=[("Arquivos Python", "*.py")])
        if caminho:
            self.entrada_principal.delete(0, tk.END)
            self.entrada_principal.insert(0, caminho)

    def selecionar_saida(self):
        """Abre um diálogo para seleção do diretório de saída."""
        diretorio = filedialog.askdirectory(title="Selecione o diretório de saída")
        if diretorio:
            self.entrada_saida.delete(0, tk.END)
            self.entrada_saida.insert(0, diretorio)

    def iniciar_criacao_exe(self):
        """Valida os dados e inicia a criação do EXE em uma thread separada."""
        nome_sistema = self.entrada_nome.get().strip()
        caminho_icone = self.entrada_icone.get().strip()
        arquivo_principal = self.entrada_principal.get().strip()
        diretorio_saida = self.entrada_saida.get().strip()

        if not nome_sistema:
            messagebox.showerror("Erro", "Informe o nome do sistema.")
            return
        if not os.path.exists(caminho_icone):
            messagebox.showerror("Erro", "Ícone não encontrado.")
            return
        if not os.path.exists(arquivo_principal):
            messagebox.showerror("Erro", "Arquivo principal não encontrado.")
            return
        if not os.path.isdir(diretorio_saida):
            messagebox.showerror("Erro", "Diretório de saída não encontrado.")
            return

        self.status_label.config(text="Status: Iniciando build...")
        self.adicionar_saida(f"Iniciando build: {nome_sistema}")

        thread_build = threading.Thread(
            target=self.criar_exe,
            args=(nome_sistema, caminho_icone, arquivo_principal, diretorio_saida),
            daemon=True
        )
        thread_build.start()

    def criar_exe(self, nome_sistema, caminho_icone, arquivo_principal, diretorio_saida):
        """Executa o comando do PyInstaller e envia as mensagens para a fila."""
        opcao_build = "--onedir" if self.tipo_build.get() == "onedir" else "--onefile"
        comando = [
            "pyinstaller", opcao_build, "--windowed",
            "--icon", caminho_icone, "--name", nome_sistema,
            "--distpath", diretorio_saida,
            "--workpath", os.path.join(diretorio_saida, "build"),
            "--specpath", os.path.join(diretorio_saida, "spec"),
            "--add-data=C:\\Users\\Kmt-02\\AppData\\Local\\Programs\\Python\\Python313\\tcl;tcl",
            "--clean", "-y", "--log-level=DEBUG", arquivo_principal
        ]

        try:
            processo = subprocess.Popen(
                comando,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,  # Força o buffer de linha
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            def ler_stdout():
                for linha in processo.stdout:
                    if linha:
                        linha = linha.strip()
                        self.queue.put(("stdout", linha))
                processo.stdout.close()

            def ler_stderr():
                for linha in processo.stderr:
                    if linha:
                        linha = linha.strip()
                        if "DEBUG:" in linha:
                            continue
                        self.queue.put(("stderr", linha))
                processo.stderr.close()

            thread_stdout = threading.Thread(target=ler_stdout, daemon=True)
            thread_stderr = threading.Thread(target=ler_stderr, daemon=True)
            thread_stdout.start()
            thread_stderr.start()

            processo.wait()
            thread_stdout.join()
            thread_stderr.join()

            # Se o build foi bem sucedido, copie a pasta externa para o diretório de saída
            if processo.returncode == 0:
                # Caminho da pasta externa que deseja copiar (exemplo: "C:\Sistema\logos")
                origem = r"C:\Sistema\logos"
                # Destino: dentro do diretório de saída
                destino = os.path.join(diretorio_saida, os.path.basename(origem))

                try:
                    if os.path.exists(destino):
                        shutil.rmtree(destino)  # Remove se já existir, para evitar conflito
                    shutil.copytree(origem, destino)
                    self.queue.put(("stdout", f"Pasta '{origem}' copiada para '{destino}'"))
                except Exception as e:
                    self.queue.put(("stderr", f"Erro ao copiar a pasta: {e}"))
                
                self.queue.put(("status", "Build finalizado com sucesso!"))
                self.queue.put(("final", "sucesso"))
            else:
                self.queue.put(("status", "Build finalizado com erros!"))
                self.queue.put(("final", "erro"))
        except Exception as e:
            self.queue.put(("status", f"Erro na execução: {e}"))
            self.queue.put(("final", "erro"))

    def processar_queue(self):
        """Processa as mensagens da fila e atualiza a interface na thread principal."""
        count = 0
        try:
            while count < 10:
                tipo, mensagem = self.queue.get_nowait()
                if tipo == "stdout":
                    self.adicionar_saida(mensagem)
                    self.status_label.config(text=f"Status: {mensagem}")
                elif tipo == "stderr":
                    self.adicionar_saida(mensagem)
                elif tipo == "status":
                    self.status_label.config(text=f"Status: {mensagem}")
                elif tipo == "final":
                    if mensagem == "sucesso":
                        # Exibe uma mensagem modal e aguarda o OK do usuário
                        if messagebox.showinfo("Sucesso", "EXE criado com sucesso!") == "ok":
                            self.limpar_texto_saida()
                    else:
                        messagebox.showerror("Erro", "Erro ao criar EXE.")
                self.queue.task_done()
                count += 1
        except queue.Empty:
            pass
        self.after(100, self.processar_queue)

    def adicionar_saida(self, texto):
        """Adiciona mensagens na área de saída."""
        self.texto_saida.config(state=tk.NORMAL)
        self.texto_saida.insert(tk.END, texto + "\n")
        self.texto_saida.see(tk.END)
        self.texto_saida.config(state=tk.DISABLED)

    def limpar_texto_saida(self):
        """Limpa a área de saída."""
        self.texto_saida.config(state=tk.NORMAL)
        self.texto_saida.delete("1.0", tk.END)
        self.texto_saida.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = ConstrutorDeExe()
    app.mainloop()