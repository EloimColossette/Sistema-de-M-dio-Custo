import tkinter as tk
from tkinter import messagebox, ttk
from conexao_db import conectar
from exportacao import exportar_para_pdf, exportar_para_excel
from logos import aplicar_icone
from tkinter import filedialog
import sys
import json
import os
import psycopg2
import re

# Função para conectar ao banco de dados PostgreSQL
class InterfaceMateriais:
    def __init__(self, janela_menu=None):
        self.janela_menu = janela_menu
        self.janela_materiais = tk.Toplevel()
        self.janela_materiais.title("Gerenciamento de Materiais")

        caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
        aplicar_icone(self.janela_materiais, caminho_icone)

        # Maximiza a janela e define o fundo com um tom neutro e moderno
        self.janela_materiais.state("zoomed")
        self.janela_materiais.configure(bg="#ecf0f1")
        
        # Configura os estilos padrões
        self.configurar_estilos()
        
        # Cabeçalho com cor de fundo escura e texto branco
        cabecalho = tk.Label(
            self.janela_materiais, text="Gerenciamento de Materiais",
            font=("Arial", 24, "bold"),
            bg="#34495e", fg="white", pady=15
        )
        cabecalho.pack(fill=tk.X)

        # Frame para Labels e Entradas
        frame_acoes = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_acoes.pack(padx=20, pady=10, fill=tk.X)

        # Label e entrada para o nome do material
        tk.Label(frame_acoes, text="Nome do Material", bg="#ecf0f1", font=("Arial", 12)) .grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_nome = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entry_nome.grid(row=0, column=1, padx=5, pady=5)

        # Combobox para fornecedor
        tk.Label(frame_acoes, text="Fornecedor", bg="#ecf0f1", font=("Arial", 12)) .grid(row=1, column=0, padx=5, pady=5, sticky="e")
        fornecedores_carregados = self.load_fornecedores()
        if not fornecedores_carregados:
            fornecedores_carregados = ["Termomecanica", "Metallica"]
        self.combobox_fornecedor = ttk.Combobox(frame_acoes, width=23, font=("Arial", 12), state="readonly")
        self.combobox_fornecedor['values'] = fornecedores_carregados
        self.combobox_fornecedor.grid(row=1, column=1, padx=5, pady=5)

        self.entry_novo_fornecedor = tk.Entry(frame_acoes, width=20, font=("Arial", 12))
        self.entry_novo_fornecedor.grid(row=1, column=2, padx=5, pady=5)

        btn_adicionar_fornecedor = tk.Button(frame_acoes, text="Adicionar", font=("Arial", 10), command=self.adicionar_fornecedor)
        btn_adicionar_fornecedor.grid(row=1, column=3, padx=5, pady=5)
        btn_excluir_fornecedor = tk.Button(frame_acoes, text="Excluir", font=("Arial", 10), command=self.excluir_fornecedor)
        btn_excluir_fornecedor.grid(row=1, column=4, padx=5, pady=5)

        # Outras entradas: Valor e Grupo
        tk.Label(frame_acoes, text="Valor", bg="#ecf0f1", font=("Arial", 12)) .grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.entry_valor = tk.Entry(frame_acoes, width=25, font=("Arial", 12))
        self.entry_valor.grid(row=2, column=1, padx=5, pady=5)
        # bind para formatação automática com 2 casas decimais:
        self.entry_valor.bind(
            "<KeyRelease>",
            lambda ev, ent=self.entry_valor: self._formatar_numero(ent, casas=4)
        )

        tk.Label(frame_acoes, text="Grupo", bg="#ecf0f1", font=("Arial", 12)) .grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.combobox_grupo = ttk.Combobox(frame_acoes, values=["Cobre", "Zinco", "Sucata"], state="readonly", font=("Arial", 12), width=23)
        self.combobox_grupo.grid(row=3, column=1, padx=5, pady=5)

        # Frame para os botões de ação dos materiais
        frame_botoes_acao = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_botoes_acao.pack(pady=10, fill=tk.X)

        botao_adicionar = ttk.Button(frame_botoes_acao, text="Adicionar", command=self.adicionar_material)
        botao_adicionar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_alterar = ttk.Button(frame_botoes_acao, text="Alterar", command=self.alterar_material)
        botao_alterar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_excluir = ttk.Button(frame_botoes_acao, text="Excluir", command=self.excluir_materiais)
        botao_excluir.pack(side=tk.LEFT, padx=5, pady=5)

        botao_limpar = ttk.Button(frame_botoes_acao, text="Limpar", command=self.limpar_campos)
        botao_limpar.pack(side=tk.LEFT, padx=5, pady=5)

        botao_voltar = ttk.Button(frame_botoes_acao, text="Voltar", command=lambda: self.voltar_para_menu())
        botao_voltar.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para os botões de exportação
        frame_botoes_exportacao = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_botoes_exportacao.pack(pady=10, fill=tk.X)

        botao_exportar_excel = ttk.Button(frame_botoes_exportacao, text="Exportar Excel", command=self.exportar_excel_materiais)
        botao_exportar_excel.pack(side=tk.LEFT, padx=5, pady=5)

        botao_exportar_pdf = ttk.Button(frame_botoes_exportacao, text="Exportar PDF", command=self.exportar_pdf_materiais)
        botao_exportar_pdf.pack(side=tk.LEFT, padx=5, pady=5)

        # Frame para o Treeview
        frame_treeview = ttk.Frame(self.janela_materiais, style="Custom.TFrame")
        frame_treeview.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # Barras de rolagem
        scrollbar_vertical = tk.Scrollbar(frame_treeview, orient=tk.VERTICAL)
        scrollbar_vertical.pack(side=tk.RIGHT, fill=tk.Y)

        scrollbar_horizontal = tk.Scrollbar(frame_treeview, orient=tk.HORIZONTAL)
        scrollbar_horizontal.pack(side=tk.BOTTOM, fill=tk.X)

        self.lista_materiais = ttk.Treeview(
            frame_treeview,
            columns=("ID", "Nome", "Fornecedor", "Valor", "Grupo"),
            show="headings",
            yscrollcommand=scrollbar_vertical.set,
            xscrollcommand=scrollbar_horizontal.set
        )
        self.lista_materiais.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configuração das colunas do Treeview
        self.lista_materiais.heading("ID", text="ID", anchor="center")
        self.lista_materiais.heading("Nome", text="Nome", anchor="center")
        self.lista_materiais.heading("Fornecedor", text="Fornecedor", anchor="center")
        self.lista_materiais.heading("Valor", text="Valor", anchor="center")
        self.lista_materiais.heading("Grupo", text="Grupo", anchor="center")
        
        self.lista_materiais.column("ID", width=50, anchor="center", stretch=False)
        self.lista_materiais.column("Nome", width=200, anchor="center")
        self.lista_materiais.column("Fornecedor", width=150, anchor="center")
        self.lista_materiais.column("Valor", width=100, anchor="center")
        self.lista_materiais.column("Grupo", width=100, anchor="center")

        # Exibe apenas as colunas desejadas (ocultando a coluna "ID")
        self.lista_materiais["displaycolumns"] = ("Nome", "Fornecedor", "Valor", "Grupo")

        # Vincula as barras de rolagem ao Treeview
        scrollbar_vertical.config(command=self.lista_materiais.yview)
        scrollbar_horizontal.config(command=self.lista_materiais.xview)

        # Evento de clique no Treeview para selecionar um item
        self.lista_materiais.bind("<ButtonRelease-1>", self.selecionar_material)

        # Atualiza a lista de materiais
        self.atualizar_lista_materiais()

        self.janela_materiais.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.janela_materiais.mainloop()

    def configurar_estilos(self):
        """Configura os estilos utilizados na interface."""
        estilo = ttk.Style()
        estilo.theme_use("alt")
        estilo.configure("Custom.TFrame", background="#ecf0f1")
        
        # Estilo para o Treeview
        estilo.configure("Treeview", 
                         background="white", 
                         foreground="black", 
                         rowheight=20, 
                         fieldbackground="white",
                         font=("Arial", 10))
        estilo.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        estilo.map("Treeview", 
                   background=[("selected", "#0078D7")], 
                   foreground=[("selected", "white")])
        
        # Estilos para os botões
        estilo.configure("Material.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#2980b9",      # Azul profissional
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("Material.TButton",
                   background=[("active", "#3498db")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])
        
        estilo.configure("Excel.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#27ae60",      # Verde profissional
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        
        estilo.configure("PDF.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#c0392b",      # Vermelho sofisticado
                         foreground="white",
                         font=("Arial", 10),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        
    def _formatar_numero(self, entry_widget, casas=2):
        """
        Formata o conteúdo do entry_widget com 'casas' casas decimais
        e separadores de milhar no padrão brasileiro.
        """
        texto = entry_widget.get()
        dígitos = re.sub(r"\D", "", texto)  # só mantém dígitos
        if not dígitos:
            novo = "0," + "0"*casas
        else:
            fator = 10**casas
            valor = int(dígitos) / fator
            # formata no estilo en_US (ex: "1,234.56") e depois converte:
            s = f"{valor:,.{casas}f}"
            novo = s.replace(",", "X").replace(".", ",").replace("X", ".")

        # atualiza o entry sem reentrar no bind
        entry_widget.unbind("<KeyRelease>")
        entry_widget.delete(0, "end")
        entry_widget.insert(0, novo)
        entry_widget.bind(
            "<KeyRelease>",
            lambda ev, ent=entry_widget: self._formatar_numero(ent, casas)
        )

    def converter_valor(self, valor):
        try:
            return float(valor.replace(',', '.'))
        except ValueError:
            messagebox.showerror("Erro", "O valor deve ser um número!")
            return None

    def obter_menor_id_disponivel(self):
        """Obtém o menor ID disponível na tabela materiais"""
        try:
            conexao = conectar()
            if conexao is None:
                return None
            cursor = conexao.cursor()
            cursor.execute("SELECT id FROM materiais ORDER BY id")
            ids = cursor.fetchall()
            conexao.close()
            
            if not ids:
                return 1
            ids_usados = [item[0] for item in ids]
            menor_id_disponivel = 1
            for id_usado in ids_usados:
                if id_usado != menor_id_disponivel:
                    return menor_id_disponivel
                menor_id_disponivel += 1
            return menor_id_disponivel
        except psycopg2.Error as e:
            print(f"Erro ao acessar o banco de dados: {e}")
            return None

    def adicionar_material(self):
        nome = self.entry_nome.get()
        fornecedor = self.combobox_fornecedor.get().strip()
        valor = self.entry_valor.get()
        grupo = self.combobox_grupo.get().strip()

        if not nome or not fornecedor or not valor:
            messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos!")
            return

        valor_convertido = self.converter_valor(valor)
        if valor_convertido is None:
            return

        conexao = conectar()
        if conexao is None:
            return

        cursor = conexao.cursor()
        menor_id = self.obter_menor_id_disponivel()
        if menor_id is None:
            messagebox.showerror("Erro", "Não foi possível determinar o menor ID disponível.")
            conexao.close()
            return

        cursor.execute(
            "INSERT INTO materiais (id, nome, fornecedor, valor, grupo) VALUES (%s, %s, %s, %s, %s)",
            (menor_id, nome, fornecedor, valor_convertido, grupo)
        )
        conexao.commit()

        # dispara notificação
        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        conexao.commit()

        conexao.close()
        messagebox.showinfo("Sucesso", "Material adicionado com sucesso!")
        
        self.atualizar_lista_materiais()

        # Atualizar o relatório de produto e material automaticamente
        if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorios_produto_material'):
            for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                widget.destroy()
            self.janela_menu.criar_relatorio_produto_material()  # Alteração aqui para chamar a função correta

    def atualizar_lista_materiais(self):
        """Atualiza a lista de materiais exibida na Treeview"""
        for item in self.lista_materiais.get_children():
            self.lista_materiais.delete(item)
        try:
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            cursor.execute("SELECT id, nome, fornecedor, valor, grupo FROM materiais ORDER BY nome ASC")
            materiais = cursor.fetchall()
            for material in materiais:
                id_material, nome, fornecedor, valor, grupo = material
                # Formata o valor convertendo ponto para vírgula
                valor_formatado = str(valor).replace('.', ',')
                self.lista_materiais.insert("", "end", values=(id_material, nome, fornecedor, valor_formatado, grupo))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar lista de materiais: {e}")
        finally:
            if 'conexao' in locals() and conexao:
                conexao.close()

    def selecionar_material(self, event):
        try:
            selected_item = self.lista_materiais.selection()[0]
            item = self.lista_materiais.item(selected_item)
            values = item['values']
            print("Valores selecionados:", values)
            if len(values) >= 4:
                self.entry_nome.delete(0, tk.END)
                self.entry_nome.insert(0, values[1])
                self.combobox_fornecedor.set(values[2])
                self.entry_valor.delete(0, tk.END)
                self.entry_valor.insert(0, str(values[3]).replace('.', ','))
                self.combobox_grupo.set(values[4])
            else:
                print("Os valores selecionados não contêm dados suficientes.")
        except IndexError:
            print("Nenhum material selecionado.")
        except Exception as e:
            print("Erro inesperado:", e)

    def alterar_material(self):
        try:
            selected_item = self.lista_materiais.selection()[0]
        except IndexError:
            messagebox.showwarning("Aviso", "Selecione um material da lista!")
            return

        nome = self.entry_nome.get()
        fornecedor = self.combobox_fornecedor.get().strip()
        valor = self.entry_valor.get()
        grupo = self.combobox_grupo.get().strip()

        if not nome or not fornecedor or not valor:
            messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos!")
            return

        valor_convertido = self.converter_valor(valor)
        if valor_convertido is None:
            return

        conexao = conectar()
        if conexao is None:
            return
        cursor = conexao.cursor()
        item_id = self.lista_materiais.item(selected_item, 'values')[0]
        cursor.execute("UPDATE materiais SET nome=%s, fornecedor=%s, valor=%s, grupo=%s WHERE id=%s", 
                       (nome, fornecedor, valor_convertido, grupo, item_id))
        conexao.commit()

        cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
        conexao.commit()

        conexao.close()
        messagebox.showinfo("Sucesso", "Material alterado com sucesso!")
        self.atualizar_lista_materiais()

        # Atualizar o relatório de matéria-prima automaticamente
        if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorio_materia_prima'):
            for widget in self.janela_menu.frame_relatorio_materia_prima.winfo_children():
                widget.destroy()
            self.janela_menu.criar_relatorio_materia_prima()

    def limpar_campos(self):
        self.entry_nome.delete(0, tk.END)
        self.combobox_fornecedor.set('')
        self.entry_valor.delete(0, tk.END)
        self.combobox_grupo.set('')

    def excluir_materiais(self):
        selected_items = self.lista_materiais.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione ao menos um material da lista!")
            return

        if messagebox.askyesno("Confirmação", f"Tem certeza que deseja excluir {len(selected_items)} materiais?"):
            conexao = conectar()
            if conexao is None:
                return
            cursor = conexao.cursor()
            for item in selected_items:
                material_id = self.lista_materiais.item(item, 'values')[0]
                cursor.execute("DELETE FROM materiais WHERE id=%s", (material_id,))
            conexao.commit()

            cursor.execute("NOTIFY canal_atualizacao, 'menu_atualizado';")
            conexao.commit()

            conexao.close()
            messagebox.showinfo("Sucesso", "Materiais excluídos com sucesso!")
            self.atualizar_lista_materiais()

            # Atualizar o relatório de produto e material automaticamente
            if self.janela_menu and hasattr(self.janela_menu, 'frame_relatorios_produto_material'):
                for widget in self.janela_menu.frame_relatorios_produto_material.winfo_children():
                    widget.destroy()
                self.janela_menu.criar_relatorio_produto_material()  # Alteração aqui para chamar a função correta

            self.limpar_campos()

    def exportar_excel_materiais(self):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="base_materiais.xlsx"
        )
        if caminho_arquivo:
            exportar_para_excel(caminho_arquivo, "materiais", ["Produto", "Fornecedor", "Valor", "Grupo"])

    def exportar_pdf_materiais(self):
        caminho_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="base_materiais.pdf"
        )
        if caminho_arquivo:
            exportar_para_pdf(
                caminho_arquivo,
                "materiais",
                ["Produto", "Fornecedor", "Valor", "Grupo"],
                "Base de Materiais"
            )

    def ordena_coluna(self, treeview, coluna, reverse):
        """Ordena a Treeview por uma coluna específica"""
        data = [(treeview.item(child)["values"], child) for child in treeview.get_children("")]
        data.sort(key=lambda x: x[0][coluna], reverse=reverse)
        for index, (item, tree_id) in enumerate(data):
            treeview.move(tree_id, '', index)
        treeview.heading(coluna, command=lambda: self.ordena_coluna(treeview, coluna, not reverse))

    def voltar_para_menu(self):
        """Fecha a janela de materiais e reexibe o menu principal com atualização visual."""
        self.janela_materiais.destroy()
        self.janela_menu.deiconify()
        self.janela_menu.state("zoomed")
        self.janela_menu.lift()  # Garante que fique no topo
        self.janela_menu.update()  # Força atualização visual
        
    def adicionar_fornecedor(self):
        novo_fornecedor = self.entry_novo_fornecedor.get().strip()
        if novo_fornecedor:
            valores_existentes = list(self.combobox_fornecedor['values'])
            if novo_fornecedor not in valores_existentes:
                valores_existentes.append(novo_fornecedor)
                self.combobox_fornecedor['values'] = valores_existentes
                self.salvar_fornecedores(valores_existentes)
                messagebox.showinfo("Sucesso", f"Fornecedor '{novo_fornecedor}' adicionado!")
                self.entry_novo_fornecedor.delete(0, tk.END)
            else:
                messagebox.showwarning("Aviso", f"O fornecedor '{novo_fornecedor}' já existe!")
        else:
            messagebox.showwarning("Aviso", "Digite o nome do fornecedor!")

    def excluir_fornecedor(self):
        fornecedor_para_excluir = self.entry_novo_fornecedor.get().strip()
        if fornecedor_para_excluir:
            valores_existentes = list(self.combobox_fornecedor['values'])
            if fornecedor_para_excluir in valores_existentes:
                valores_existentes.remove(fornecedor_para_excluir)
                self.combobox_fornecedor['values'] = valores_existentes
                self.salvar_fornecedores(valores_existentes)
                messagebox.showinfo("Sucesso", f"Fornecedor '{fornecedor_para_excluir}' removido!")
                self.entry_novo_fornecedor.delete(0, tk.END)
            else:
                messagebox.showwarning("Aviso", f"O fornecedor '{fornecedor_para_excluir}' não existe!")
        else:
            messagebox.showwarning("Aviso", "Digite o nome do fornecedor para remover!")

    def salvar_fornecedores(self, fornecedores):
        """Salva os fornecedores em um arquivo JSON dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)  # Diretório do script
        caminho_pasta = os.path.join(caminho_base, "config")  # Criar pasta 'config'
        os.makedirs(caminho_pasta, exist_ok=True)  # Cria a pasta se não existir

        caminho_arquivo = os.path.join(caminho_pasta, "fornecedores.json")  # Caminho completo

        try:
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                json.dump(fornecedores, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print("Erro ao salvar fornecedores:", e)

    def load_fornecedores(self):
        """Carrega os fornecedores do arquivo JSON dentro da pasta 'config'."""
        caminho_base = os.path.dirname(__file__)
        caminho_arquivo = os.path.join(caminho_base, "config", "fornecedores.json")

        if os.path.exists(caminho_arquivo):
            try:
                with open(caminho_arquivo, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            except Exception as e:
                print("Erro ao carregar fornecedores:", e)

        return []
    
    def on_closing(self):
        """Função para lidar com o fechamento da janela."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar o programa?"):
            self.janela_materiais.destroy()  # Fecha a janela atual
            sys.exit()  # Encerra completamente o programa