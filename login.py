import tkinter as tk
from tkinter import messagebox
import sys
from menu import Janela_Menu # Importa a classe Janela_Menu
from centralizacao_tela import centralizar_janela  # Função para centralizar a janela
from resetar_senha import TelaResetarSenha  # Tela para resetar senha
from logos import aplicar_icone        # Função para aplicar o ícone na janela
import customtkinter as ctk
from conexao_db import conectar
import json
import os

class TelaLogin:
    def __init__(self):
        # Configura o CustomTkinter
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Cria a janela principal de login
        self.janela_login = ctk.CTk()
        self.janela_login.title("Login")
        largura_janela = 450
        altura_janela = 450
        centralizar_janela(self.janela_login, largura_janela, altura_janela)
        self.janela_login.geometry(f"{largura_janela}x{altura_janela}")

        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        try:
            aplicar_icone(self.janela_login, caminho_icone)
        except Exception as e:
            print("Ícone não encontrado:", e)
        
        # Menu superior
        self.menu = tk.Menu(self.janela_login)
        self.janela_login.config(menu=self.menu)
        
        # Submenu de configurações
        self.menu_configuracoes = tk.Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Configurações", menu=self.menu_configuracoes)
        self.menu_configuracoes.add_command(label="Conectar ao Banco de Dados", command=self.abrir_configuracoes_db)
        self.menu_configuracoes.add_separator()
        self.menu_configuracoes.add_command(label="Fechar", command=self.janela_login.quit)
        
        # Frame principal
        self.frame_principal = ctk.CTkFrame(self.janela_login, corner_radius=15, fg_color="#ecf0f1")
        self.frame_principal.pack(padx=20, pady=20, fill="both", expand=True)
        
        # Título da tela de login
        self.label_titulo = ctk.CTkLabel(
            self.frame_principal,
            text="Login",
            font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"),
            text_color="#2c3e50"
        )
        self.label_titulo.pack(pady=(10, 20))
        
        # Frame para os campos de entrada (usuário e senha)
        self.frame_campos = ctk.CTkFrame(self.frame_principal, corner_radius=10, fg_color="#ffffff")
        self.frame_campos.pack(pady=10, padx=20, fill="x", expand=False)
        
        # Campo Usuário
        self.label_usuario = ctk.CTkLabel(
            self.frame_campos, text="Usuário:", anchor="w", 
            font=ctk.CTkFont(size=12), text_color="#2c3e50"
        )
        self.label_usuario.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entrada_usuario = ctk.CTkEntry(
            self.frame_campos, placeholder_text="Digite seu usuário", width=200
        )
        self.entrada_usuario.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Campo Senha
        self.label_senha = ctk.CTkLabel(
            self.frame_campos, text="Senha:", anchor="w", 
            font=ctk.CTkFont(size=12), text_color="#2c3e50"
        )
        self.label_senha.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.entrada_senha = ctk.CTkEntry(
            self.frame_campos, placeholder_text="Digite sua senha", show="*", width=200
        )
        self.entrada_senha.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        
        self.frame_campos.columnconfigure(1, weight=1)
        
        # Checkbox para mostrar/ocultar a senha
        self.mostrar_senha_var = ctk.BooleanVar(value=False)
        def toggle_senha():
            if self.mostrar_senha_var.get():
                self.entrada_senha.configure(show="")
            else:
                self.entrada_senha.configure(show="*")
        self.checkbox_mostrar = ctk.CTkCheckBox(
            self.frame_campos, text="Mostrar Senha", variable=self.mostrar_senha_var,
            command=toggle_senha, text_color="#2c3e50", font=ctk.CTkFont(size=12)
        )
        self.checkbox_mostrar.grid(row=2, column=1, padx=10, pady=(0, 10), sticky="w")
        
        # Frame para os botões
        self.frame_botoes = ctk.CTkFrame(self.frame_principal, corner_radius=10, fg_color="#ecf0f1")
        self.frame_botoes.pack(pady=20)
        
        self.botao_login = ctk.CTkButton(
            self.frame_botoes, text="Login", width=120,
            command=lambda: self.verificar_login(), fg_color="#2980b9"
        )
        self.botao_login.grid(row=0, column=0, padx=15, pady=10)
        
        self.botao_fechar = ctk.CTkButton(
            self.frame_botoes, text="Fechar", width=120,
            fg_color="#c0392b", command=self.janela_login.destroy
        )
        self.botao_fechar.grid(row=0, column=1, padx=15, pady=10)
        
        self.botao_esqueci_senha = ctk.CTkButton(
            self.frame_principal,
            text="Esqueci minha senha",
            width=200,
            fg_color="#27ae60",
            command=lambda: TelaResetarSenha(self.janela_login)
        )
        self.botao_esqueci_senha.pack(pady=(10, 15))
        
        self.janela_login.bind("<Return>", lambda event: self.verificar_login())

        # Garante que o "X" da janela feche corretamente
        self.janela_login.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def solicitar_nova_senha(self, user_id, usuario):
        # Cria a janela para alteração de senha usando CTkToplevel
        tela_alteracao = ctk.CTkToplevel(self.janela_login)
        tela_alteracao.title("Alterar Senha Inicial")
        tela_alteracao.geometry("400x400")
        tela_alteracao.resizable(False, False)
        centralizar_janela(tela_alteracao, 400, 400)
        tela_alteracao.transient(self.janela_login)
        tela_alteracao.grab_set()
        
        frame_principal = ctk.CTkFrame(tela_alteracao, corner_radius=15, fg_color="#ecf0f1")
        frame_principal.pack(padx=20, pady=20, fill="both", expand=True)
        
        caminho_icone = r"C:\Sistema\logos\Kametal.ico"
        def aplicar_icon():
            try:
                aplicar_icone(tela_alteracao, caminho_icone)
            except Exception as e:
                print("Ícone não encontrado:", e)
        aplicar_icon()
        tela_alteracao.after(200, aplicar_icon)
        
        label_titulo = ctk.CTkLabel(
            frame_principal,
            text="Alterar Senha Inicial",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#2c3e50"
        )
        label_titulo.pack(pady=(0, 15))
        
        # Cria a entrada para a nova senha
        nova_senha_entry = ctk.CTkEntry(
            frame_principal,
            placeholder_text="Digite a nova senha",
            show="*",
            width=250
        )
        nova_senha_entry.pack(pady=(5, 10))
        
        # Cria a entrada para confirmar a nova senha
        confirma_senha_entry = ctk.CTkEntry(
            frame_principal,
            placeholder_text="Confirme a nova senha",
            show="*",
            width=250
        )
        confirma_senha_entry.pack(pady=(5, 10))
        
        # Cria a variável e o CheckBox para mostrar/ocultar as senhas
        var_mostrar_senha = tk.BooleanVar(value=False)
        def toggle_password():
            if var_mostrar_senha.get():
                nova_senha_entry.configure(show="")  # Exibe a senha
                confirma_senha_entry.configure(show="")  # Exibe a senha
            else:
                nova_senha_entry.configure(show="*")  # Oculta a senha
                confirma_senha_entry.configure(show="*")  # Oculta a senha

        check_mostrar = ctk.CTkCheckBox(
            frame_principal,
            text="Mostrar Senha",
            variable=var_mostrar_senha,
            command=toggle_password
        )
        check_mostrar.pack(pady=(5, 10))
        
        def salvar_nova_senha():
            nova_senha = nova_senha_entry.get().strip()
            confirma_senha = confirma_senha_entry.get().strip()
            if not nova_senha or not confirma_senha:
                messagebox.showerror("Erro", "Todos os campos são obrigatórios.", parent=tela_alteracao)
                return
            if nova_senha != confirma_senha:
                messagebox.showerror("Erro", "As senhas não conferem.", parent=tela_alteracao)
                return
            if not (nova_senha.isdigit() and 1 <= len(nova_senha) <= 6):
                messagebox.showerror("Erro", "A nova senha deve conter entre 1 e 6 dígitos numéricos.", parent=tela_alteracao)
                return
            
            conexao = conectar()  # Usa a conexão da biblioteca
            if not conexao:
                messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados.", parent=tela_alteracao)
                return

            try:
                cursor = conexao.cursor()
                cursor.execute(
                    "UPDATE usuarios SET senha = %s, primeiro_login = FALSE WHERE id = %s",
                    (nova_senha, user_id)
                )
                conexao.commit()
                messagebox.showinfo("Sucesso", "Senha atualizada com sucesso!", parent=tela_alteracao)
                # Fecha a tela de alteração e abre a tela do menu
                tela_alteracao.destroy()
                self.janela_login.withdraw()
                app = Janela_Menu(user_id)  # Certifique-se de que Janela_Menu está importada corretamente
                app.mainloop()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao atualizar a senha: {e}", parent=tela_alteracao)
            finally:
                cursor.close()
                conexao.close()
        
        botao_salvar = ctk.CTkButton(
            frame_principal,
            text="Salvar",
            command=salvar_nova_senha,
            width=200,
            height=40,
            fg_color="#2980b9",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        botao_salvar.pack(pady=20)
        
        # Faz com que pressionar Enter acione a função salvar_nova_senha
        tela_alteracao.bind("<Return>", lambda event: salvar_nova_senha())
    
    def verificar_login(self):
        usuario = self.entrada_usuario.get().strip()
        senha = self.entrada_senha.get().strip()

        if not usuario or not senha:
            return messagebox.showwarning("Aviso", "Preencha todos os campos.")

        if not (senha.isdigit() and 1 <= len(senha) <= 6):
            return messagebox.showerror("Erro", "Senha inválida.")

        # carrega IP/porta
        cfg  = self.carregar_configuracoes()
        host = cfg.get("DB_HOST")
        port = cfg.get("DB_PORT")

        try:
            conexao = conectar(ip=host, porta=port)
            if not conexao:
                return messagebox.showerror("Erro", "Falha ao conectar ao banco de dados.")

            cursor = conexao.cursor()
            cursor.execute(
                "SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name='usuarios')"
            )
            if not cursor.fetchone()[0]:
                return messagebox.showerror("Erro", "Tabela 'usuarios' não encontrada.")

            cursor.execute(
                "SELECT id, senha, primeiro_login FROM usuarios WHERE usuario = %s;",
                (usuario,)
            )
            resultado = cursor.fetchone()
            if not resultado:
                return messagebox.showerror("Login", "Usuário não encontrado.")

            user_id, senha_bd, primeiro_login = resultado
            if senha_bd != senha:
                return messagebox.showerror("Login", "Usuário ou senha incorretos.")

            if primeiro_login:
                return self.solicitar_nova_senha(user_id, usuario)

            # somente aqui, se passou por todas as validações:
            messagebox.showinfo("Login", "Login realizado com sucesso!")
            self.janela_login.withdraw()
            app = Janela_Menu(user_id)
            app.mainloop()

        except Exception as e:
            messagebox.showerror("Erro", f"{e}")
            return  # <— evita cair no sucesso

        finally:
            try:
                cursor.close()
                conexao.close()
            except:
                pass
    
    def TelaResetarSenha(self):
        messagebox.showinfo("Resetar Senha", "Funcionalidade de resetar senha não implementada.")

    @staticmethod
    def carregar_configuracoes_static():
        caminho_arquivo = os.path.join("config", "config.json")
        if not os.path.exists(caminho_arquivo):
            return {}  # Retorna dicionário vazio se o arquivo não existe

        with open(caminho_arquivo, "r") as f:
            return json.load(f)

    def carregar_configuracoes(self):
        return self.carregar_configuracoes_static()

    def salvar_configuracoes(self, ip, porta, dbname, usuario_db, senha_db):
        cfg = {
            "DB_HOST": ip,
            "DB_PORT": porta,
            "DB_NAME": dbname,
            "DB_USER": usuario_db,
            "DB_PASSWORD": senha_db
        }

        # Cria a pasta "config" se não existir
        pasta_config = "config"
        os.makedirs(pasta_config, exist_ok=True)

        # Caminho completo para o arquivo
        caminho_arquivo = os.path.join(pasta_config, "config.json")

        # Salva o JSON
        with open(caminho_arquivo, "w") as f:
            json.dump(cfg, f, indent=4)

    def abrir_configuracoes_db(self):
        cfg = self.carregar_configuracoes()
        ip_atual = cfg.get("DB_HOST", "")
        porta_atual = cfg.get("DB_PORT", "")
        nome_atual = cfg.get("DB_NAME", "")
        usuario_atual = cfg.get("DB_USER", "")
        senha_atual = cfg.get("DB_PASSWORD", "")

        janela_config = tk.Toplevel(self.janela_login)
        janela_config.title("Configurações de Conexão com Banco de Dados")
        largura_janela, altura_janela = 400, 350
        centralizar_janela(janela_config, largura_janela, altura_janela)
        janela_config.geometry(f"{largura_janela}x{altura_janela}")
        janela_config.resizable(True, True)
        janela_config.attributes("-topmost", True)

        try:
            janela_config.wm_iconbitmap(r"C:\Sistema\logos\Kametal.ico")
        except Exception as e:
            print("Não foi possível aplicar o ícone:", e)

        # Validação de IP
        def validar_ip(valor):
            if valor == "": return True
            if valor.isdigit() and len(valor) <= 3:
                return 0 <= int(valor) <= 255
            return False
        validar_cmd = janela_config.register(validar_ip)

        # Frame para IP compacto
        frame_ip = ctk.CTkFrame(janela_config, fg_color="transparent")
        frame_ip.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        ctk.CTkLabel(janela_config, text="IP:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        largura_ip = 40
        entradas = []
        for i in range(4):
            e = ctk.CTkEntry(frame_ip, width=largura_ip, validate="key", validatecommand=(validar_cmd, "%P"))
            entradas.append(e)
        # layout com pontos
        for idx, widget in enumerate([entradas[0], '.', entradas[1], '.', entradas[2], '.', entradas[3]]):
            if isinstance(widget, str):
                ctk.CTkLabel(frame_ip, text=widget).grid(row=0, column=2*idx-1)
            else:
                widget.grid(row=0, column=2*idx)
        for entry, part in zip(entradas, ip_atual.split('.')):
            entry.insert(0, part)

        # Porta
        ctk.CTkLabel(janela_config, text="Porta:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        entrada_porta = ctk.CTkEntry(janela_config, width=100)
        entrada_porta.grid(row=1, column=1, sticky="w", padx=5)
        entrada_porta.insert(0, porta_atual)

        # Campo DB Name
        ctk.CTkLabel(janela_config, text="DB Name:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        entrada_dbname = ctk.CTkEntry(janela_config, width=150)
        entrada_dbname.grid(row=2, column=1, sticky="w", padx=5)
        entrada_dbname.insert(0, nome_atual)

        # Usuário DB
        ctk.CTkLabel(janela_config, text="Usuário:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        entrada_user = ctk.CTkEntry(janela_config, width=150)
        entrada_user.grid(row=3, column=1, sticky="w", padx=5)
        entrada_user.insert(0, usuario_atual)

        # Senha DB e checkbox
        ctk.CTkLabel(janela_config, text="Senha:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
        frame_pass = ctk.CTkFrame(janela_config, fg_color="transparent")
        frame_pass.grid(row=4, column=1, sticky="w", padx=5)
        entrada_pass = ctk.CTkEntry(frame_pass, width=150, show="*")
        entrada_pass.pack(side="left")
        entrada_pass.insert(0, senha_atual)
        # Checkbox para mostrar/ocultar
        mostrar_senha_var = tk.BooleanVar(value=False)
        def toggle_senha_config():
            entrada_pass.configure(show="" if mostrar_senha_var.get() else "*")
        chk_show = ctk.CTkCheckBox(frame_pass, text="Mostrar Senha", variable=mostrar_senha_var, command=toggle_senha_config)
        chk_show.pack(side="left", padx=(10,0))

        # Função salvar
        def salvar_config():
            ip = '.'.join(e.get().strip() for e in entradas)
            porta = entrada_porta.get().strip()
            dbnm  = entrada_dbname.get().strip()
            user = entrada_user.get().strip()
            pwd  = entrada_pass.get().strip()
            if all([*ip.split('.'), porta, dbnm, user, pwd]):
                self.salvar_configuracoes(ip, porta, dbnm, user, pwd)
                messagebox.showinfo("Configurações", "Salvas com sucesso!", parent=janela_config)
                janela_config.destroy()
            else:
                messagebox.showerror("Erro", "Preencha todos os campos.", parent=janela_config)

        # Botões Salvar e Fechar
        frame_botoes = ctk.CTkFrame(janela_config, fg_color="transparent")
        frame_botoes.grid(row=5, column=0, columnspan=2, pady=15)
        ctk.CTkButton(frame_botoes, text="Salvar", width=80, command=salvar_config).pack(side="left", padx=(0,10))
        ctk.CTkButton(frame_botoes, text="Fechar", width=80, fg_color="#c0392b",command=janela_config.destroy).pack(side="left")

        janela_config.grid_columnconfigure(1, weight=1)

    def on_closing(self):
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Verifica se a janela ainda existe antes de tentar fechá-la
            if self.janela_login:
                try:
                    self.janela_login.destroy()
                except Exception as e:
                    print("Erro ao destruir a janela:", e)

            sys.exit(0)  # Encerra o programa

    def run(self):
        self.janela_login.mainloop()