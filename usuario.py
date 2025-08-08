import tkinter as tk
from tkinter import messagebox, ttk
import sys
from permissao import InterfacePermissoes
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from conexao_db import conectar
import psycopg2

class InterfaceUsuarios:
    # Dicionário para mapear nomes internos para rótulos amigáveis (opcional)
    janelas = {
        "criar_interface_produto": "Base de Produtos",
        "criar_interface_materiais": "Base de Materiais",
        "Janela_InsercaoNF": "Inserção de NF",
        "Calculo_Produto": "Cálculo de NF",
        "SistemaNF": "Saída NF",
        "criar_media_custo": "Média Custo",
        "criar_tela_usuarios": "Gerenciamento de Usuário"
    }
    
    def __init__(self, janela_menu):
        if not isinstance(janela_menu, tk.Tk):
            messagebox.showerror("Erro", "A janela de menu não foi passada corretamente.")
            return

        self.janela_menu = janela_menu
        self.janela_menu.withdraw()  # Oculta o menu ao abrir a tela de usuários

        # Cria a janela de usuários como Toplevel
        self.janela_usuarios = tk.Toplevel(self.janela_menu)
        self.janela_usuarios.title("Gerenciamento de Usuários")
        self.janela_usuarios.state("zoomed")
        self.janela_usuarios.configure(bg="#ecf0f1")
        
        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self.janela_usuarios, caminho_icone)
        
        # Configura os estilos chamando o método interno
        self._configurar_estilo()
        
        # Cabeçalho (igual à janela Material)
        cabecalho = tk.Label(self.janela_usuarios, text="Gerenciamento de Usuários", font=("Arial", 24, "bold"), bg="#34495e", fg="white", pady=15)
        cabecalho.pack(fill=tk.X)
        
        # Frame para as entradas de cadastro
        frame_cadastros = ttk.Frame(self.janela_usuarios, style="Custom.TFrame")
        frame_cadastros.pack(padx=20, pady=10, fill=tk.X)
        
        tk.Label(frame_cadastros, text="Nome:", font=("Arial", 12)).grid(row=0, column=0, sticky="e", padx=10, pady=5)
        self.entrada_nome = tk.Entry(frame_cadastros, width=30, font=("Arial", 12), borderwidth=2, relief="groove")
        self.entrada_nome.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Label(frame_cadastros, text="Usuário:", font=("Arial", 12)).grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.entrada_usuario = tk.Entry(frame_cadastros, width=30, font=("Arial", 12), borderwidth=2, relief="groove")
        self.entrada_usuario.grid(row=1, column=1, padx=10, pady=5)
        
        # Label e Entry para a senha
        tk.Label(frame_cadastros, text="Senha:", font=("Arial", 12))\
            .grid(row=2, column=0, sticky="e", padx=10, pady=5)
        self.entrada_senha = tk.Entry(frame_cadastros, width=30, font=("Arial", 12),
                                    borderwidth=2, relief="groove", show="*")
        self.entrada_senha.grid(row=2, column=1, padx=10, pady=5)

        # Variável para controlar a exibição da senha – agora como atributo da instância
        self.var_mostrar_senha = tk.BooleanVar(self.janela_usuarios, value=False)

        def toggle_senha():
            print("toggle_senha chamado; var =", self.var_mostrar_senha.get())
            if self.var_mostrar_senha.get():
                self.entrada_senha.config(show="")   # Exibe a senha
            else:
                self.entrada_senha.config(show="*")  # Oculta a senha

        # Checkbutton para mostrar/ocultar a senha
        check_senha = tk.Checkbutton(frame_cadastros, text="Mostrar Senha",
                                    variable=self.var_mostrar_senha,
                                    command=toggle_senha, font=("Arial", 10),
                                    onvalue=True, offvalue=False)
        check_senha.grid(row=2, column=2, padx=10, pady=5)

        # Frame para os botões de ação
        frame_botoes = ttk.Frame(self.janela_usuarios, style="Custom.TFrame")
        frame_botoes.pack(padx=20, pady=10, fill=tk.X)
        
        ttk.Button(frame_botoes, text="Adicionar", command=self.adicionar_usuario).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(frame_botoes, text="Excluir", command=self.excluir_usuario).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(frame_botoes, text="Alterar Usuário", command=self.alterar_usuario).pack(side=tk.LEFT, padx=10, pady=5)
        ttk.Button(frame_botoes, text="Voltar", command=self.voltar_ao_menu).pack(side=tk.LEFT, padx=10, pady=5)
        
        # Frame para o Treeview
        frame_tabela = ttk.Frame(self.janela_usuarios, style="Custom.TFrame")
        frame_tabela.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        container = ttk.Frame(frame_tabela, style="Custom.TFrame")
        container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(container, 
                                 columns=("ID", "Nome", "Usuário", "Senha", "Permissões"), 
                                 show="headings")
        self.tree.heading("ID", text="ID", anchor="center")
        self.tree.heading("Nome", text="Nome", anchor="center")
        self.tree.heading("Usuário", text="Usuário", anchor="center")
        self.tree.heading("Senha", text="Senha", anchor="center")
        self.tree.heading("Permissões", text="Permissões", anchor="center")
        
        self.tree.column("ID", anchor="center", width=50)
        self.tree.column("Nome", anchor="center", width=200)
        self.tree.column("Usuário", anchor="center", width=150)
        self.tree.column("Senha", anchor="center", width=150)
        self.tree.column("Permissões", anchor="center", width=250)
        
        # Exibe apenas as colunas desejadas (ocultando a coluna "ID")
        self.tree["displaycolumns"] = ("Nome", "Usuário", "Senha", "Permissões")
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree.yview)
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vertical_scrollbar.set)
        horizontal_scrollbar = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=self.tree.xview)
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=horizontal_scrollbar.set)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        
        # Botão "Abrir Permissões" posicionado abaixo do Treeview
        ttk.Button(self.janela_usuarios, text="Abrir Permissões", command=self.abrir_permissoes)\
            .pack(pady=(5, 15))
        
        # Exibe os usuários na Treeview
        self.exibir_usuarios()
        
        # Dicionário para mapear IDs (para exclusão e alteração)
        self.usuario_ids = {}
        
        self.janela_usuarios.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.janela_usuarios.mainloop()
    
    def _configurar_estilo(self):
        """Configura os estilos para frames, Treeview e botões."""
        style = ttk.Style(self.janela_usuarios)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        
        # Estilo para o Treeview
        style.configure("Treeview", 
                        background="white", 
                        foreground="black", 
                        rowheight=20, 
                        fieldbackground="white",
                        font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.map("Treeview", 
                  background=[("selected", "#0078D7")], 
                  foreground=[("selected", "white")])
        
        # Estilo para os botões (com o nome "Usuarios.TButton")
        style.configure("Usuarios.TButton",
                        padding=(5, 2),
                        relief="raised",
                        background="#2980b9",      # Azul profissional
                        foreground="white",
                        font=("Arial", 10),
                        borderwidth=2,
                        highlightbackground="#34495e",
                        highlightthickness=1)
        style.map("Usuarios.TButton",
                  background=[("active", "#3498db")],
                  foreground=[("active", "white")],
                  relief=[("pressed", "sunken"), ("!pressed", "raised")])
        return style

    def proximo_id_disponivel(self):
        """Obtém o próximo ID disponível na tabela de usuários."""
        conexao = conectar()  # Usando a função do arquivo db_connection.py
        if conexao:
            cursor = conexao.cursor()
            cursor.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM usuarios")
            proximo_id = cursor.fetchone()[0]
            cursor.close()
            conexao.close()
            return proximo_id
        return None

    def adicionar_usuario(self):
        """Adiciona um novo usuário ao banco de dados."""
        nome = self.entrada_nome.get()
        usuario = self.entrada_usuario.get()
        senha = self.entrada_senha.get()

        if nome and usuario and senha:
            if senha.isdigit() and 1 <= len(senha) <= 6:
                id_disponivel = self.proximo_id_disponivel()
                if id_disponivel is None:
                    return
                conexao = conectar()  # Usando a conexão do arquivo externo
                if conexao:
                    cursor = conexao.cursor()
                    try:
                        cursor.execute(
                            "INSERT INTO usuarios (id, nome, usuario, senha) VALUES (%s, %s, %s, %s)",
                            (id_disponivel, nome, usuario, senha)
                        )
                        conexao.commit()

                        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                        conexao.commit()

                        messagebox.showinfo("Sucesso", "Usuário adicionado com sucesso.")
                        self.exibir_usuarios()
                    except psycopg2.IntegrityError:
                        messagebox.showerror("Erro", "Usuário já existente.")
                    except Exception as e:
                        messagebox.showerror("Erro", f"Ocorreu um erro ao adicionar o usuário: {e}")
                    finally:
                        cursor.close()
                        conexao.close()
            else:
                messagebox.showerror("Erro", "A senha deve conter entre 1 e 6 dígitos numéricos.")
        else:
            messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos.")

    def excluir_usuario(self):
        """Exclui o(s) usuário(s) selecionado(s)."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione pelo menos um usuário para excluir.")
            return

        confirm = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza de que deseja excluir {len(selected_items)} usuário(s)?")
        if confirm:
            conexao = conectar()  # Usando a conexão do arquivo externo
            if conexao:
                try:
                    cursor = conexao.cursor()
                    user_ids = [self.tree.item(item, "values")[0] for item in selected_items]

                    # Convertendo os IDs para inteiros (caso venham como strings)
                    user_ids = tuple(map(int, user_ids))

                    cursor.execute("DELETE FROM usuarios WHERE id IN %s", (user_ids,))
                    conexao.commit()

                    cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    conexao.commit()

                    # Removendo os usuários da interface gráfica
                    for item in selected_items:
                        self.tree.delete(item)

                    messagebox.showinfo("Sucesso", "Usuário(s) excluído(s) com sucesso!")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao excluir usuário(s): {e}")
                finally:
                    cursor.close()
                    conexao.close()
                    
    def abrir_permissoes(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um usuário para gerenciar permissões.")
            return
        user_id = self.tree.item(selected_item, "values")[0]
        InterfacePermissoes(user_id, update_callback=self.exibir_usuarios)

    def alterar_usuario(self):
        """Altera os dados do usuário selecionado."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Aviso", "Selecione um usuário para alterar os dados.")
            return

        # Obtém os dados do usuário selecionado (id, Nome, Usuário, Senha, Permissões)
        item_values = self.tree.item(selected_item, "values")
        user_id = item_values[0]

        edit_window = tk.Toplevel(self.janela_usuarios)
        edit_window.title("Alterar Dados do Usuário")
        edit_window.geometry("400x400")
        edit_window.resizable(False, False)
        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(edit_window, caminho_icone)
        centralizar_janela(edit_window, 500, 400)
        edit_window.configure(bg="#ecf0f1")

        header_frame = tk.Frame(edit_window, bg="#34495e", pady=5)
        header_frame.pack(fill=tk.X)
        header_label = tk.Label(header_frame, text="Alterar Dados do Usuário",
                                bg="#34495e", fg="white", font=("Arial", 18, "bold"))
        header_label.pack(padx=10, pady=5)

        field_frame = tk.Frame(edit_window, bg="#ecf0f1", padx=20, pady=20)
        field_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(field_frame, text="Novo Nome:", bg="#ecf0f1", font=("Arial", 12)).grid(row=0, column=0, sticky="e", padx=10, pady=10)
        nome_entry = ttk.Entry(field_frame, width=30)
        nome_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(field_frame, text="Novo Usuário:", bg="#ecf0f1", font=("Arial", 12)).grid(row=1, column=0, sticky="e", padx=10, pady=10)
        usuario_entry = ttk.Entry(field_frame, width=30)
        usuario_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(field_frame, text="Nova Senha:", bg="#ecf0f1", font=("Arial", 12)).grid(row=2, column=0, sticky="e", padx=10, pady=10)
        senha_entry = ttk.Entry(field_frame, width=30, show="*")
        senha_entry.grid(row=2, column=1, padx=10, pady=10)

        # Preenche os campos com os valores atuais (item_values = (id, Nome, Usuário, Senha, Permissões))
        if len(item_values) >= 4:
            nome_entry.insert(0, item_values[1])
            usuario_entry.insert(0, item_values[2])
            senha_entry.insert(0, item_values[3])

        # Cria uma variável para controlar a exibição da senha, associada à janela de edição
        var_mostrar_senha = tk.BooleanVar(edit_window, value=False)

        def toggle_senha():
            print("toggle_senha chamado; valor =", var_mostrar_senha.get())
            if var_mostrar_senha.get():
                senha_entry.config(show="")   # Exibe a senha
            else:
                senha_entry.config(show="*")  # Oculta a senha

        # Adiciona o Checkbutton ao field_frame
        check_senha = tk.Checkbutton(field_frame, text="Mostrar Senha", variable=var_mostrar_senha,
                                     command=toggle_senha, bg="#ecf0f1", onvalue=True, offvalue=False)
        check_senha.grid(row=2, column=2, padx=10, pady=10)

        button_frame = tk.Frame(edit_window, bg="#ecf0f1", pady=10)
        button_frame.pack(fill=tk.X)
        style = ttk.Style(edit_window)
        style.configure("Usuario.TButton",
                        padding=(5, 2),
                        relief="raised",
                        background="#2980b9",
                        foreground="white",
                        font=("Arial", 10, "bold"),
                        borderwidth=2)
        style.map("Usuario.TButton",
                  background=[("active", "#3498db")],
                  foreground=[("active", "white")],
                  relief=[("pressed", "sunken"), ("!pressed", "raised")])
        
        def salvar_alteracoes():
            novo_nome = nome_entry.get().strip()
            novo_usuario = usuario_entry.get().strip()
            nova_senha = senha_entry.get().strip()

            if not (novo_nome and novo_usuario and nova_senha):
                messagebox.showerror("Erro", "Todos os campos são obrigatórios.", parent=edit_window)
                return

            if not (nova_senha.isdigit() and 1 <= len(nova_senha) <= 6):
                messagebox.showerror("Erro", "A nova senha deve conter entre 1 e 6 dígitos numéricos.", parent=edit_window)
                return

            conexao = conectar()  # Usa a função conectar() para obter a conexão
            if conexao:
                try:
                    cursor = conexao.cursor()
                    cursor.execute(
                        "UPDATE usuarios SET nome = %s, usuario = %s, senha = %s WHERE id = %s",
                        (novo_nome, novo_usuario, nova_senha, user_id)
                    )
                    conexao.commit()

                    cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
                    conexao.commit()
                    
                    self.exibir_usuarios()  # Atualiza a exibição dos usuários
                    edit_window.destroy()   # Fecha a janela de edição
                except Exception as e:
                    messagebox.showerror("Erro", f"Ocorreu um erro ao atualizar o usuário: {e}")
                finally:
                    cursor.close()
                    conexao.close()

        save_button = ttk.Button(button_frame, text="Salvar", command=salvar_alteracoes, style="Usuario.TButton")
        cancel_button = ttk.Button(button_frame, text="Cancelar", command=edit_window.destroy, style="Usuario.TButton")
        save_button.pack(side=tk.LEFT, padx=10, expand=True)
        cancel_button.pack(side=tk.LEFT, padx=10, expand=True)

    def exibir_usuarios(self):
        """Consulta e exibe os usuários no Treeview."""
        conn = conectar()  # Substituí self.conectar_banco() por conectar()
        if not conn:
            return

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    u.id, 
                    u.nome, 
                    u.usuario, 
                    u.senha, 
                    COALESCE(STRING_AGG(p.janela, ', ' ORDER BY p.janela), 'Sem permissões') AS permissoes
                FROM usuarios u
                LEFT JOIN permissoes p ON u.id = p.usuario_id
                GROUP BY u.id, u.nome, u.usuario, u.senha
                ORDER BY u.id;
            """)
            usuarios = cursor.fetchall()

            # Limpa a Treeview antes de adicionar novos dados
            self.tree.delete(*self.tree.get_children())

            for usuario in usuarios:
                user_id, nome, usuario_nome, senha, permissoes = usuario
                # Converte os nomes internos para rótulos amigáveis, se necessário
                permissoes_legiveis = ', '.join(self.janelas.get(janela, janela) for janela in permissoes.split(', '))
                self.tree.insert("", "end", values=(user_id, nome, usuario_nome, senha, permissoes_legiveis))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exibir usuários: {e}")
        finally:
            cursor.close()
            conn.close()

    def voltar_ao_menu(self):
        """Fecha a janela de usuários e reexibe o menu principal com atualização visual."""
        self.janela_menu.deiconify()             # Reexibe a janela do menu
        self.janela_menu.state("zoomed")         # Garante que fique maximizada
        self.janela_menu.lift()                  # Garante que fique no topo
        self.janela_menu.update()                # Força atualização visual
        self.janela_usuarios.destroy()           # Fecha a janela de usuários

    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Fecha a conexão com o banco de dados, se houver uma conexão ativa
            try:
                if hasattr(self, "conn") and self.conn:
                    self.conn.close()
                    print("Conexão com o banco de dados fechada.")
            except Exception as e:
                print(f"Erro ao fechar a conexão com o banco: {e}")

            # Fecha as janelas corretamente
            if hasattr(self, "janela_usuarios"):
                self.janela_usuarios.destroy()

            if hasattr(self, "janela_menu"):
                self.janela_menu.destroy()

            sys.exit(0)  # Encerra o programa completamente