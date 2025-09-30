import tkinter as tk
from tkinter import messagebox, ttk
import sys
from permissao import InterfacePermissoes
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela
from conexao_db import conectar
import psycopg2
import threading

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
        self.header_container = tk.Frame(self.janela_usuarios, bg="#34495e")
        self.header_container.pack(fill=tk.X)

        # Label principal (título) — armazenado como atributo para uso posterior se necessário
        self.cabecalho = tk.Label(self.header_container,
                                text="Gerenciamento de Usuários",
                                font=("Arial", 24, "bold"),
                                bg="#34495e", fg="white", pady=15)
        self.cabecalho.pack(side="left", fill=tk.X, expand=True)

        # ---------- Botão de Ajuda seguro (colocado no mesmo header_container) ----------
        self.help_frame = tk.Frame(self.header_container, bg="#34495e")
        self.help_frame.pack(side="right", padx=(8, 12), pady=6)

        self.botao_ajuda_usuarios = tk.Button(
            self.help_frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=lambda: self._abrir_ajuda_usuarios_modal()
        )
        self.botao_ajuda_usuarios.bind("<Enter>", lambda e: self.botao_ajuda_usuarios.config(bg="#3b5566"))
        self.botao_ajuda_usuarios.bind("<Leave>", lambda e: self.botao_ajuda_usuarios.config(bg="#2c3e50"))
        self.botao_ajuda_usuarios.pack(side="right", padx=(4, 6), pady=4)

        # Tooltip (pequeno, não bloqueante) — usa o método já definido na classe
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_usuarios, "Ajuda de Usuários (F1)")
        except Exception:
            pass

        # Atalho F1: vincular ao Toplevel dessa janela (mais seguro que bind_all global)
        try:
            # vincula o evento F1 ao toplevel da janela de usuários (apenas quando janela estiver aberta)
            self.janela_usuarios.bind_all("<F1>", lambda e: self._abrir_ajuda_usuarios_modal())
            # opcionalmente, também permitir fechar modal com Esc (feito dentro do modal)
        except Exception:
            try:
                # fallback: bind simples na janela de usuários
                self.janela_usuarios.bind("<F1>", lambda e: self._abrir_ajuda_usuarios_modal())
            except Exception:
                pass
        
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

    def _create_tooltip(self, widget, text, delay=450, max_width=None):
        """
        Tooltip melhorado: quebra automática de linhas e ajuste para não sair da tela.
        - widget: widget alvo
        - text: texto do tooltip
        - delay: ms até exibir
        - max_width: largura máxima do tooltip em pixels (opcional)
        """
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            try:
                screen_w = widget.winfo_screenwidth()
                screen_h = widget.winfo_screenheight()
            except Exception:
                screen_w, screen_h = 1024, 768

            # determina wraplength: respeita max_width se passado, senão calcula baseado na tela
            if max_width:
                wrap_len = max(120, min(max_width, screen_w - 80))
            else:
                wrap_len = min(360, max(200, screen_w - 160))

            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)

            label = tk.Label(
                win,
                text=text,
                bg="#333333",
                fg="white",
                font=("Segoe UI", 9),
                bd=0,
                padx=6,
                pady=4,
                wraplength=wrap_len
            )
            label.pack()

            # posição inicial: centrado horizontalmente sobre o widget, abaixo do widget
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajustar horizontalmente para não sair da tela
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # ajustar verticalmente: tenta abaixo; se não couber, tenta acima; senão limita dentro da tela
            if y + h > screen_h:
                y_above = widget.winfo_rooty() - h - 6
                if y_above > 10:
                    y = y_above
                else:
                    y = max(10, screen_h - h - 10)

            win.geometry(f"+{x}+{y}")
            tooltip["win"] = win

        def hide():
            if tooltip["after_id"]:
                try:
                    widget.after_cancel(tooltip["after_id"])
                except Exception:
                    pass
                tooltip["after_id"] = None
            if tooltip["win"]:
                try:
                    tooltip["win"].destroy()
                except Exception:
                    pass
                tooltip["win"] = None

        def schedule_show(e=None):
            tooltip["after_id"] = widget.after(delay, show)

        widget.bind("<Enter>", schedule_show)
        widget.bind("<Leave>", lambda e: hide())
        widget.bind("<ButtonPress>", lambda e: hide())

    def _abrir_ajuda_usuarios_modal(self, contexto=None):
        """Abre modal profissional explicando como adicionar, editar, excluir usuários e gerenciar permissões."""
        try:
            modal = tk.Toplevel(self.janela_usuarios)
            modal.title("Ajuda — Usuários")
            modal.transient(self.janela_usuarios)  # deixa modal "transient" em relação à janela de usuários
            modal.grab_set()

            # Dimensões e centralização
            w, h = 900, 650
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
            aplicar_icone(modal, caminho_icone)

            # Cabeçalho
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Gerenciamento de Usuários", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # Corpo: nav esquerda + conteúdo direita
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=240, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0, 12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10, 6), padx=12)

            sections = [
                "Visão Geral",
                "Como Adicionar Usuário",
                "Como Editar Usuário",
                "Como Excluir Usuário",
                "Como Gerenciar Permissões",
                "Validações e Erros Comuns",
                "Boas Práticas / Segurança",
                "FAQ"
            ]
            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)

            txt = tk.Text(content_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb = tk.Scrollbar(content_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            # Conteúdo das seções (personalize conforme desejar)
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "A tela de Usuários permite criar, editar, remover usuários e atribuir permissões a funcionalidades (janelas) do sistema. A interface apresenta campos (Nome, Usuário, Senha), uma lista (Treeview) com os usuários existentes e botões para ações."
            )

            contents["Como Adicionar Usuário"] = (
                "Como Adicionar Usuário — Passo a passo\n\n"
                "1) Preencha os campos: Nome, Usuário e Senha.\n"
                "   • Observação: no código atual a senha é validada como 1 a 6 dígitos numéricos.\n"
                "2) (Opcional) Marque 'Mostrar Senha' para conferir a senha.\n"
                "3) Clique em 'Adicionar'.\n"
                "4) O sistema executa INSERT no banco; se sucesso, a Treeview é atualizada automaticamente.\n\n"
                "Dicas:\n"
                " - Use nomes claros e padrão de login previsível (ex.: primeiro.nome).\n"
                " - Após criar o usuário, abra Permissões para configurar acessos."
            )

            contents["Como Editar Usuário"] = (
                "Como Editar Usuário — Passo a passo\n\n"
                "1) Selecione o usuário na lista (Treeview).\n"
                "2) Clique em 'Alterar Usuário' (ou botão equivalente).\n"
                "3) Na janela de edição altere Nome / Usuário / Senha conforme necessário.\n"
                "4) Clique em 'Salvar' para aplicar as mudanças (UPDATE no banco).\n\n"
                "Observações:\n"
                " - Validações são aplicadas (campos obrigatórios e formato de senha).\n"
                " - O código dispara notificações para sincronizar outras janelas após alteração."
            )

            contents["Como Excluir Usuário"] = (
                "Como Excluir Usuário — Passo a passo\n\n"
                "1) Selecione um ou mais usuários na Treeview.\n"
                "2) Clique em 'Excluir'.\n"
                "3) Confirme a exclusão na caixa de diálogo que aparece.\n"
                "4) O sistema executa DELETE por ID e atualiza a lista.\n\n"
                "Recomendações:\n"
                " - Prefira desativar o usuário em vez de excluir, se precisar manter histórico.\n"
                " - Faça backup antes de operações em lote."
            )

            contents["Como Gerenciar Permissões"] = (
                "Como Gerenciar Permissões — Passo a passo\n\n"
                "1) Selecione o usuário e clique em 'Abrir Permissões'. Isso abre a janela de permissões (InterfacePermissoes).\n"
                "2) No combobox, escolha o usuário (id - nome).\n"
                "3) Marque as checkboxes correspondentes às janelas/modulos aos quais quer dar acesso.\n"
                "   • Desmarque para remover acesso.\n"
                "4) Clique em 'Salvar'. O procedimento atual sincroniza a tabela de permissões:\n"
                "   geralmente faz DELETE das permissões atuais do usuário e INSERT das marcadas (ou calcula diffs).\n\n"
                "Notas técnicas:\n"
                " - Garanta que as chaves de 'janela' usadas nas permissões correspondam aos nomes que o código\n"
                "   verifica antes de liberar acesso a cada funcionalidade.\n"
                " - Após salvar, as mudanças devem refletir imediatamente (NOTIFY) ou ao atualizar a janela."
            )

            contents["Validações e Erros Comuns"] = (
                "Validações e Erros Comuns\n\n"
                " - Senha inválida: verifique o formato (no layout atual, somente dígitos 1-6).\n"
                " - Usuário duplicado: IntegrityError do banco — verifique mensagens retornadas e mostre alerta.\n"
                " - Falha de conexão com banco: operações de CRUD vão falhar — checar logs e conexão.\n"
            )

            contents["Boas Práticas / Segurança"] = (
                "Boas Práticas e Segurança\n\n"
                " - Use senhas fortes e, se possível, migre para armazenamento com hash (bcrypt/sha).\n"
                " - Aplique princípio do menor privilégio: conceda apenas as permissões necessárias.\n"
                " - Mantenha logs de auditoria e backups regulares do banco de dados."
            )

            contents["FAQ"] = (
                "FAQ — Usuários\n\n"
                "Q: Posso criar um usuário com o mesmo nome de outro?\n"
                "A: Não é recomendado; o banco pode rejeitar por constraint. Use identificadores únicos.\n\n"
                "Q: Como ver quais permissões um usuário tem?\n"
                "A: Verifique a coluna de permissões na Treeview ou abra a janela de Permissões para visualizar.\n"
            )

            # Função para mostrar conteúdo
            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            # Inicializa com a primeira seção
            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])

            listbox.bind("<<ListboxSelect>>", on_select)

            # Rodapé com botão Fechar
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right", padx=6)

            # Atalhos
            modal.bind("<Escape>", lambda e: modal.destroy())

            modal.focus_set()
            modal.wait_window()

        except Exception as e:
            print("Erro ao abrir modal de ajuda (Usuários):", e)

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
        """Reexibe o menu imediatamente e faz a limpeza em background para não bloquear o mainloop."""
        try:
            self.janela_menu.deiconify()
            try:
                self.janela_menu.state("zoomed")
            except Exception:
                pass
            self.janela_menu.lift()
            self.janela_menu.update_idletasks()
            try:
                self.janela_menu.focus_force()
            except Exception:
                pass
        except Exception:
            pass

        def _cleanup_and_destroy():
            try:
                # tenta fechar cursor/conn no self.janela_usuarios (preferencial) ou em self (fallback)
                child = getattr(self, "janela_usuarios", None) or self
                if hasattr(child, "cursor") and getattr(child, "cursor"):
                    try:
                        child.cursor.close()
                    except Exception:
                        pass
                if hasattr(child, "conn") and getattr(child, "conn"):
                    try:
                        child.conn.close()
                    except Exception:
                        pass
                # também fechar se estiver em self diretamente
                if child is not self:
                    if hasattr(self, "cursor") and getattr(self, "cursor"):
                        try:
                            self.cursor.close()
                        except Exception:
                            pass
                    if hasattr(self, "conn") and getattr(self, "conn"):
                        try:
                            self.conn.close()
                        except Exception:
                            pass
            finally:
                # destrói a janela filha na thread principal
                try:
                    if child is not None and hasattr(child, "after"):
                        child.after(0, child.destroy)
                    else:
                        # fallback
                        if child is not None:
                            child.destroy()
                except Exception:
                    try:
                        if child is not None:
                            child.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()

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