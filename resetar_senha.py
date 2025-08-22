import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
import sys
from centralizacao_tela import centralizar_janela
from conexao_db import conectar

class TelaResetarSenha:
    def __init__(self, janela_principal):
        """
        Monta a interface de reset de senha.
        A janela_principal é a janela que será ocultada enquanto a tela de reset estiver ativa.
        """
        self.janela_principal = janela_principal
        # Oculta a janela principal
        self.janela_principal.withdraw()
        
        # Cria a janela de reset
        self.janela_resetar = tk.Toplevel(self.janela_principal)
        self.janela_resetar.title("Resetar Senha")
        largura_tela = 450
        altura_tela = 450
        centralizar_janela(self.janela_resetar, largura_tela, altura_tela)
        self.janela_resetar.geometry(f"{largura_tela}x{altura_tela}")
        
        # Define o ícone (utilizando iconbitmap do Tkinter)
        try:
            self.janela_resetar.iconbitmap("C:\\Sistema\\logos\\Kametal.ico")
        except Exception as e:
            print(f"Erro ao definir o ícone com iconbitmap: {e}")
        
        # Configura o fundo da janela
        self.janela_resetar.configure(bg="#ecf0f1")
        
        # Cria o frame principal com CustomTkinter
        self.frame_principal = ctk.CTkFrame(self.janela_resetar, corner_radius=15, fg_color="#ecf0f1")
        self.frame_principal.pack(padx=20, pady=20, fill="both", expand=True)
        
        # Título e instruções
        self.label_titulo = ctk.CTkLabel(
            self.frame_principal,
            text="Resetar Senha",
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color="#34495e"
        )
        self.label_titulo.pack(pady=(20, 10))
        
        self.label_instrucoes = ctk.CTkLabel(
            self.frame_principal,
            text="Por favor, insira seu usuário e a nova senha para resetar.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#2c3e50"
        )
        self.label_instrucoes.pack(pady=(0, 20))
        
        # Frame para os campos de entrada
        self.frame_campos = ctk.CTkFrame(self.frame_principal, corner_radius=10, fg_color="#ffffff")
        self.frame_campos.pack(pady=10, padx=20, fill="x", expand=False)
        
        self.label_usuario = ctk.CTkLabel(
            self.frame_campos,
            text="Usuário:",
            anchor="w",
            font=ctk.CTkFont(size=12),
            text_color="#34495e"
        )
        self.label_usuario.grid(row=0, column=0, padx=10, pady=15, sticky="w")
        self.entrada_usuario = ctk.CTkEntry(self.frame_campos, placeholder_text="Digite seu usuário", width=200)
        self.entrada_usuario.grid(row=0, column=1, padx=10, pady=15, sticky="ew")
        
        self.label_nova_senha = ctk.CTkLabel(
            self.frame_campos,
            text="Nova Senha:",
            anchor="w",
            font=ctk.CTkFont(size=12),
            text_color="#34495e"
        )
        self.label_nova_senha.grid(row=1, column=0, padx=10, pady=15, sticky="w")
        self.entrada_nova_senha = ctk.CTkEntry(
            self.frame_campos, placeholder_text="Digite a nova senha", show="*", width=200
        )
        self.entrada_nova_senha.grid(row=1, column=1, padx=10, pady=15, sticky="ew")
        
        # Checkbox para mostrar/ocultar a senha
        self.mostrar_senha_var = ctk.BooleanVar(value=False)
        def toggle_senha():
            if self.mostrar_senha_var.get():
                self.entrada_nova_senha.configure(show="")
            else:
                self.entrada_nova_senha.configure(show="*")
        self.checkbox_mostrar = ctk.CTkCheckBox(
            self.frame_campos,
            text="Mostrar Senha",
            variable=self.mostrar_senha_var,
            command=toggle_senha,
            text_color="#34495e",
            font=ctk.CTkFont(size=12)
        )
        self.checkbox_mostrar.grid(row=2, column=1, padx=10, pady=(0, 15), sticky="w")
        
        self.frame_campos.columnconfigure(1, weight=1)
        
        # Frame para os botões
        self.frame_botoes = ctk.CTkFrame(self.frame_principal, corner_radius=10, fg_color="#ecf0f1")
        self.frame_botoes.pack(pady=20)
        
        self.botao_resetar = ctk.CTkButton(
            self.frame_botoes,
            text="Resetar Senha",
            width=150,
            fg_color="#2980b9",
            text_color="white",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=self.resetar_senha
        )
        self.botao_resetar.grid(row=0, column=0, padx=10, pady=10)
        
        self.botao_voltar = ctk.CTkButton(
            self.frame_botoes,
            text="Voltar",
            width=150,
            fg_color="#c0392b",
            text_color="white",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=self.voltar
        )
        self.botao_voltar.grid(row=0, column=1, padx=10, pady=10)
        
        # Define o fechamento da janela resetar
        self.janela_resetar.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def resetar_senha(self):
        """Reseta a senha de um usuário no banco de dados."""

        usuario = self.entrada_usuario.get().strip()
        nova_senha = self.entrada_nova_senha.get().strip()

        # Validação da nova senha
        if not (nova_senha.isdigit() and 1 <= len(nova_senha) <= 6):
            messagebox.showerror("Erro", "A senha deve conter entre 1 e 6 dígitos numéricos.")
            return

        # Conectar ao banco
        conexao = conectar()  # Substitui self.conectar_banco_postgres()
        if not conexao:
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados.")
            return

        try:
            cursor = conexao.cursor()
            cursor.execute("UPDATE usuarios SET senha = %s WHERE usuario = %s", (nova_senha, usuario))

            if cursor.rowcount == 0:
                messagebox.showerror("Erro", "Usuário não encontrado.")
            else:
                conexao.commit()
                messagebox.showinfo("Sucesso", "Senha alterada com sucesso.")
                
                # Fecha a janela de redefinição de senha e retorna à tela principal
                self.janela_resetar.destroy()
                self.janela_principal.deiconify()

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao alterar a senha: {e}")
        
        finally:
            cursor.close()
            conexao.close()

    def voltar(self):
        self.janela_resetar.destroy()
        self.janela_principal.deiconify()

    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente."""

        if not messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            return  # Se o usuário cancelar, simplesmente retorna

        print("Fechando o programa corretamente...")

        # Fecha a conexão com o banco de dados, se estiver aberta
        try:
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
                print("Conexão com o banco de dados fechada.")
        except Exception as e:
            print(f"Erro ao fechar a conexão: {e}")

        # Fecha a janela principal
        self.janela_resetar.destroy()

        # Encerra o programa
        sys.exit(0)