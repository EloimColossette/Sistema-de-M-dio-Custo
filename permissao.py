import tkinter as tk
from tkinter import messagebox, ttk
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from conexao_db import conectar
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
# Importar outras funções necessárias, como conectar(), centralizar_janela(), aplicar_icone(), etc.

class InterfacePermissoes:
    def __init__(self, user_id=None, main_window=None, update_callback=None):
        self.update_callback = update_callback
        self.main_window = main_window

        self.janela_permissoes = tk.Toplevel()
        self.janela_permissoes.title("Gerenciar Permissões")
        self.janela_permissoes.geometry("400x500")
        aplicar_icone(self.janela_permissoes, "C:\\Sistema\\logos\\Kametal.ico")
        centralizar_janela(self.janela_permissoes, 400, 500)
        self.janela_permissoes.config(bg="#ecf0f1")

        self._aplicar_estilos_ttk()

        frame = ttk.Frame(self.janela_permissoes, padding=10, style="Custom.TFrame")
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Selecione um usuário:", style="Custom.TLabel").pack(pady=(0,5))

        self.combo_usuarios = ttk.Combobox(frame, state="readonly", font=("Arial", 10))
        self.combo_usuarios.pack(pady=(0,10))

        self.janelas = {
            "criar_interface_produto": "Base de Produtos",
            "criar_interface_materiais": "Base de Materiais",
            "Janela_InsercaoNF": "Inserção de NF",
            "Calculo_Produto": "Calculo de NF",
            "SistemaNF": "Saída NF",
            "criar_media_custo": "Média Custo",
            "criar_tela_usuarios": "Gerenciamento de Usuário",
            "RelatorioApp": "Relatorio Item por Grupo",
            "CadastroProdutosApp": "Cotação",
            "RegistroTeste": "Registro de Teste"
        }
        self.permissoes_vars = {nome: tk.BooleanVar() for nome in self.janelas.keys()}

        permissoes_container = ttk.Frame(frame)
        permissoes_container.pack(fill="both", expand=True, pady=10)

        canvas = tk.Canvas(permissoes_container, bg="#ecf0f1", highlightthickness=0)
        scrollbar = ttk.Scrollbar(permissoes_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.frame_permissoes = ttk.LabelFrame(canvas, text="Permissões", style="Custom.TLabelframe", padding=10)
        frame_window = canvas.create_window((0, 0), window=self.frame_permissoes, anchor="nw")

        def ajustar_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(frame_window, width=canvas.winfo_width(), height=canvas.winfo_height())
        self.frame_permissoes.bind("<Configure>", ajustar_scroll)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        for nome, var in self.permissoes_vars.items():
            ttk.Checkbutton(self.frame_permissoes, text=self.janelas[nome],
                            variable=var, style="Custom.TCheckbutton").pack(anchor="w", pady=2)

        def carregar_permissoes_event(event=None):
            user_id_selecionado = self.combo_usuarios.get().split(" - ")[0]
            permissoes_atribuidas = self.carregar_permissoes_usuario(user_id_selecionado)
            for nome, var in self.permissoes_vars.items():
                var.set(nome in permissoes_atribuidas)

        self.combo_usuarios.bind("<<ComboboxSelected>>", carregar_permissoes_event)

        def salvar():
            user_id_selecionado = self.combo_usuarios.get().split(" - ")[0]
            permissoes_selecionadas = [nome for nome, var in self.permissoes_vars.items() if var.get()]
            self.salvar_permissoes(user_id_selecionado, permissoes_selecionadas)
            if self.update_callback:
                self.update_callback()

             # 2) Atualiza a janela de menu principal
            if self.main_window and hasattr(self.main_window, "atualizar_pagina"):
                self.main_window.atualizar_pagina()

            self.janela_permissoes.destroy()

        ttk.Button(frame, text="Salvar", command=salvar).pack(pady=10)

        usuarios = self.obter_usuarios()
        self.combo_usuarios["values"] = [f"{id} - {nome}" for id, nome in usuarios]

        if user_id:
            for valor in self.combo_usuarios["values"]:
                if valor.startswith(str(user_id) + " -"):
                    self.combo_usuarios.set(valor)
                    carregar_permissoes_event()
                    break

    def _aplicar_estilos_ttk(self):
        style = ttk.Style(self.janela_permissoes)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TLabelframe", background="#ecf0f1", foreground="#34495e", font=("Arial", 10, "bold"))
        style.configure("Custom.TCheckbutton", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Perm.TButton", background="#2980b9", foreground="white", font=("Arial", 10, "bold"), padding=5)
        style.map("Perm.TButton",
                  background=[("active", "#3498db")],
                  foreground=[("active", "white")])
        
    def obter_usuarios(self):
        """Obtém a lista de usuários do banco de dados."""
        conexao = conectar()  # Usando a conexão centralizada
        if not conexao:
            return []

        try:
            cursor = conexao.cursor()
            cursor.execute("SELECT id, nome FROM usuarios")
            usuarios = cursor.fetchall()
            return usuarios
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao obter usuários: {e}")
            return []
        finally:
            cursor.close()
            conexao.close()

    def carregar_permissoes_usuario(self, user_id):
        """Carrega as permissões do usuário a partir do banco de dados."""
        conn = conectar()  # Substituí self.conectar_banco() por conectar()
        if not conn:
            return set()  # Retorna um conjunto vazio se não houver conexão

        try:
            cursor = conn.cursor()
            cursor.execute("SELECT janela FROM permissoes WHERE usuario_id = %s", (int(user_id),))
            permissoes_existentes = {row[0] for row in cursor.fetchall()}
            return permissoes_existentes
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar permissões do usuário: {e}")
            return set()
        finally:
            cursor.close()
            conn.close()

    def salvar_permissoes(self, user_id, permissoes_selecionadas):
        """Salva as permissões do usuário no banco de dados."""
        if not user_id or not str(user_id).isdigit():
            messagebox.showerror("Erro", "ID do usuário inválido!")
            return

        conn = conectar()  # Usando a conexão centralizada
        if not conn:
            return

        try:
            cursor = conn.cursor()
            # Obtém as permissões atuais do usuário
            cursor.execute("SELECT janela FROM permissoes WHERE usuario_id = %s", (int(user_id),))
            permissoes_atuais = {row[0] for row in cursor.fetchall()}

            novas_permissoes = set(permissoes_selecionadas)
            permissoes_a_adicionar = novas_permissoes - permissoes_atuais
            permissoes_a_remover = permissoes_atuais - novas_permissoes

            if permissoes_a_remover:
                cursor.executemany(
                    "DELETE FROM permissoes WHERE usuario_id = %s AND janela = %s",
                    [(int(user_id), janela) for janela in permissoes_a_remover]
                )
            if permissoes_a_adicionar:
                cursor.executemany(
                    "INSERT INTO permissoes (usuario_id, janela, permitido) VALUES (%s, %s, %s)",
                    [(int(user_id), janela, True) for janela in permissoes_a_adicionar]
                )
            conn.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conn.commit()

            messagebox.showinfo("Sucesso", "Permissões atualizadas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar permissões: {e}")
        finally:
            cursor.close()
            conn.close()
