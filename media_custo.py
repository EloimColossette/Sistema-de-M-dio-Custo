import tkinter as tk
from tkinter import ttk, messagebox
import sys
import threading
from conexao_db import conectar
from logos import aplicar_icone
from aba_sudeste_media_custo import criar_media_custo_sudeste
from aba_centro_oeste_media_custo import criar_media_aba_centro_oeste
from decimal import Decimal
from exportacao import exportar_notebook_para_excel


class MediaCusto:
    def __init__(self, main_window=None, font_size=12):
        self.main_window = main_window
        self.font_size = font_size
        self.conn = None

        # Cria janela
        self.janela_media_custo = tk.Toplevel()
        self.janela_media_custo.title("Tabela de Produtos")
        self.janela_media_custo.geometry("800x600")
        try:
            self.janela_media_custo.state("zoomed")
        except Exception:
            pass

        # Aplica ícone
        try:
            caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
            aplicar_icone(self.janela_media_custo, caminho_icone)
        except Exception:
            pass

        if self.main_window is not None:
            try:
                self.main_window.withdraw()
            except Exception:
                pass

        # Monta UI
        self._criar_notebook_e_tabelas()

        # Preenche ICM
        try:
            self.conn = self.conectar_banco()
            if self.conn:
                self._popular_aba_icm()
        except Exception:
            pass

        # Observação e botões
        self.observacao_label = tk.Label(self.janela_media_custo, text="Observação: Acrescentar 3,25%", fg="red", font=("Arial", 12, "bold"))
        self.observacao_label.pack(pady=10)

        frame_botoes = tk.Frame(self.janela_media_custo, bg="#ecf0f1")
        frame_botoes.pack(pady=10)

        btn_voltar = ttk.Button(frame_botoes, text="Voltar", width=20,
                                command=lambda: self.voltar_para_menu(self.janela_media_custo, self.main_window))
        btn_exportar = ttk.Button(frame_botoes, text="Exportar para Excel", width=20,
                                  command=lambda: exportar_notebook_para_excel(self.notebook))

        btn_voltar.grid(row=0, column=0, padx=10, pady=5)
        btn_exportar.grid(row=0, column=1, padx=10, pady=5)

        # ---- Botão de Ajuda discreto (aparece ao lado dos botões Voltar / Exportar) ----
        self.botao_ajuda_media_pequeno = tk.Button(
            frame_botoes,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=self._abrir_ajuda_media_modal
        )
        # coloca na mesma linha, coluna 2 (ajuste se quiser mais deslocamento)
        self.botao_ajuda_media_pequeno.grid(row=0, column=2, padx=10, pady=5)

        # efeito hover
        self.botao_ajuda_media_pequeno.bind("<Enter>", lambda e: self.botao_ajuda_media_pequeno.config(bg="#3b5566"))
        self.botao_ajuda_media_pequeno.bind("<Leave>", lambda e: self.botao_ajuda_media_pequeno.config(bg="#2c3e50"))

        # Tooltip (utiliza seu utilitário _create_tooltip se existir)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda_media_pequeno, "Ajuda — Média de Custo (F1)")
        except Exception:
            pass

        # Atalho F1 (vinculado à janela de média de custo) - já existia, mas garantir localmente
        try:
            self.janela_media_custo.bind("<F1>", lambda e: self._abrir_ajuda_media_modal())
        except Exception:
            pass

        self.janela_media_custo.protocol("WM_DELETE_WINDOW", self._on_closing)

        try:
            self.janela_media_custo.focus_force()
        except Exception:
            pass

    def _criar_notebook_e_tabelas(self):
        self.notebook = ttk.Notebook(self.janela_media_custo)
        self.notebook.pack(expand=True, fill="both")

        self.aba_icm = tk.Frame(self.notebook)
        self.aba_sudeste = tk.Frame(self.notebook)
        self.aba_centro_oeste_nordeste = tk.Frame(self.notebook)

        self.notebook.add(self.aba_icm, text="ICM 18%")
        self.notebook.add(self.aba_sudeste, text="Região Sudeste ICM 12%")
        self.notebook.add(self.aba_centro_oeste_nordeste, text="Centro-Oeste e Nordeste 7%")

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        frame_tabela = tk.Frame(self.aba_icm)
        frame_tabela.pack(expand=True, fill="both", padx=10, pady=10)

        self.tree_icm = ttk.Treeview(
            frame_tabela,
            columns=("Produto", "Estoque", "Custo Estoque", "Custo Empresa", "Custo 5%", "Custo 10%", "Custo 15%", "Custo 20%"),
            show="headings",
            style="Custom.Treeview"
        )
        self.tree_icm.heading("Produto", text="Nome do Produto")
        self.tree_icm.heading("Estoque", text="Estoque")
        self.tree_icm.heading("Custo Estoque", text="Média de Custo do Estoque")
        self.tree_icm.heading("Custo Empresa", text="Média Custo Empresa")
        self.tree_icm.heading("Custo 5%", text="Custo com 5%")
        self.tree_icm.heading("Custo 10%", text="Custo com 10%")
        self.tree_icm.heading("Custo 15%", text="Custo com 15%")
        self.tree_icm.heading("Custo 20%", text="Custo com 20%")

        for col in ("Produto", "Estoque", "Custo Estoque", "Custo Empresa", "Custo 5%", "Custo 10%", "Custo 15%", "Custo 20%"):
            self.tree_icm.column(col, anchor="center")

        style = ttk.Style()
        try:
            style.theme_use("alt")
        except Exception:
            pass
        style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        style.configure("Treeview", rowheight=30)
        style.configure("Custom.Treeview", font=("Arial", 10), rowheight=30)
        style.map("Treeview", background=[("selected", "#0078D7")], foreground=[("selected", "white")])

        scrollbar_vertical = tk.Scrollbar(frame_tabela, orient="vertical", command=self.tree_icm.yview)
        scrollbar_horizontal = tk.Scrollbar(frame_tabela, orient="horizontal", command=self.tree_icm.xview)
        self.tree_icm.config(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

        self.tree_icm.grid(row=0, column=0, sticky="nsew")
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

        frame_tabela.grid_rowconfigure(0, weight=1)
        frame_tabela.grid_columnconfigure(0, weight=1)

    def _create_tooltip(self, widget, text, delay=450):
        """Tooltip melhorado: quebra de linha automática e ajuste para não sair da tela."""
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            # calcula largura máxima adequada para o tooltip (não maior que a tela)
            try:
                screen_w = widget.winfo_screenwidth()
                screen_h = widget.winfo_screenheight()
            except Exception:
                screen_w, screen_h = 1024, 768

            wrap_len = min(360, max(200, screen_w - 80))  # largura do texto em pixels, com limites

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

            # posição inicial centrada horizontalmente sobre o widget, abaixo do widget
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            # força o cálculo do tamanho real do tooltip
            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # ajustar horizontalmente para não sair da tela
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 10:
                x = 10

            # ajustar verticalmente: se não couber embaixo, tenta acima do widget
            if y + h > screen_h:
                y_above = widget.winfo_rooty() - h - 6
                if y_above > 10:
                    y = y_above
                else:
                    # caso não caiba nem acima nem abaixo, limita para caber na tela
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

    def _abrir_ajuda_media_modal(self, contexto=None):
        """Abre modal de ajuda explicando a janela, as 3 abas e a exportação."""
        try:
            modal = tk.Toplevel(self.janela_media_custo)
            modal.title("Ajuda — Média de Custo")
            modal.transient(self.janela_media_custo)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões e centralização
            w, h = 900, 640
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # Ícone (tenta aplicar, falha silenciosa se não disponível)
            try:
                caminho_icone = "C:\\Sistema\\logos\\Kametal.ico"
                aplicar_icone(modal, caminho_icone)
            except Exception:
                pass

            # Cabeçalho do modal
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Média de Custo", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # Corpo: nav esquerda + conteúdo direita
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0,12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10,6), padx=12)

            sections = [
                "Visão Geral",
                "Aba ICM (ICM 18%)",
                "Aba Sudeste (ICM 12%)",
                "Aba Centro-Oeste / Nordeste (7%)",
                "Exportação",
                "Boas Práticas",
                "FAQ"
            ]
            listbox = tk.Listbox(nav_frame, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            for s in sections:
                listbox.insert("end", s)
            listbox.pack(fill="both", expand=True, padx=10, pady=(0,10))

            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)

            txt = tk.Text(content_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb = tk.Scrollbar(content_frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")

            # Conteúdos (ajustei para refletir o que seu código faz; vejo isso em media_custo.py). :contentReference[oaicite:1]{index=1}
            contents = {}
            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "Esta janela apresenta a Tabela de Produtos com cálculos de média ponderada do custo\n"
                "e simulações com acréscimos. Use as abas para ver as versões por regime/região fiscal."
            )

            contents["Aba ICM (ICM 18%)"] = (
                "ICM 18% — explicação\n\n"
                "- Aba padrão com cálculo usando fatores aplicados no projeto (métodos em MediaCusto).\n"
                "- Colunas: Produto, Estoque, Média de Custo do Estoque, Média Custo Empresa, Custo 5/10/15/20%.\n"
                "- A linha TOTAL agrega estoque e as médias ponderadas."
            )

            contents["Aba Sudeste (ICM 12%)"] = (
                "Região Sudeste (ICM 12%) — detalhes operacionais\n\n"
                "Esta aba aplica a fórmula ajustada para o regime/regra fiscal do Sudeste, usando alíquota efetiva de 12%.\n\n"
                "O que muda em relação à aba ICM 18%:\n"
                " - O fator de conversão aplicado às médias ponderadas é diferente (reduz o impacto tributário).\n"
                " - São usados os mesmos dados de entrada (estoque e custos do banco), porém o resultado final\n"
                "   (Média de Custo do Estoque / Custo Empresa) é recalculado com coeficientes próprios desta região.\n\n"
                "Observações práticas:\n"
                " - Use esta aba para simular preços de venda/markups quando o cliente ou operação estiver no Sudeste.\n"
                " - Se os valores estiverem muito distantes da aba ICM 18%, verifique as fórmulas em\n"
                "   'aba_sudeste_media_custo.py' (funções que ajustam alíquotas e coeficientes).\n"
                " - Recomendação: antes de fechar a negociação, exporte para Excel e valide o cálculo com a contabilidade."
            )

            contents["Aba Centro-Oeste / Nordeste (7%)"] = (
                "Centro-Oeste / Nordeste (7%) — detalhes operacionais\n\n"
                "Nesta aba os cálculos usam o coeficiente definido para regiões com alíquota efetiva aproximada de 7%.\n\n"
                "O que considerar:\n"
                " - Além de alterar apenas a alíquota, o módulo pode aplicar ajustes regionais de custo logístico\n"
                "   ou incentivos/fatores locais (ver 'aba_centro_oeste_media_custo.py').\n"
                " - Use esta aba para avaliar margens em centros de distribuição localizados nessas regiões,\n"
                "   pois custos e tributações podem reduzir o preço final comparado ao Sudeste/ICM 18%.\n\n"
                "Boas práticas:\n"
                " - Compare as três abas antes de definir o preço de venda para um cliente: ICM 18% (base),\n"
                "   Sudeste (12%) e Centro-Oeste/Nordeste (7%). Isso ajuda a mapear risco de margem por região.\n"
                " - Ao detectar discrepâncias grandes, verifique as tabelas auxiliares (frete, armazenagem) e\n"
                "   valide se algum ajuste extra está sendo aplicado no módulo regional."
            )

            contents["Exportação"] = (
                "Exportação\n\n"
                "- O botão 'Exportar para Excel' (na parte inferior da janela) chama a função\n"
                "  exportar_notebook_para_excel(self.notebook) e salva cada aba em planilhas separadas.\n"
                "- Verifique permissões de escrita e selecione um diretório com espaço suficiente.\n"
                "- Para exportar em PDF, podemos acrescentar botão/rotina similar se desejar."
            )

            contents["Boas Práticas"] = (
                "Boas Práticas\n\n"
                "- Recarregue os dados antes de exportar para garantir consistência.\n"
            )

            contents["FAQ"] = (
                "FAQ\n\n"
                "Q: Posso exportar apenas uma aba?\n"
                "A: A função atual exporta todo o notebook; posso adaptar para exportar só a aba ativa."
            )

            def mostrar_secao(key):
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                txt.configure(state="disabled")
                txt.yview_moveto(0)

            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])
            listbox.bind("<<ListboxSelect>>", on_select)

            # Rodapé com Fechar
            ttk.Separator(modal, orient="horizontal").pack(fill="x")
            rodape = tk.Frame(modal, bg="white")
            rodape.pack(side="bottom", fill="x", padx=12, pady=10)
            btn_close = tk.Button(rodape, text="Fechar", bg="#34495e", fg="white",
                                bd=0, padx=12, pady=8, command=modal.destroy)
            btn_close.pack(side="right", padx=6)

            modal.bind("<Escape>", lambda e: modal.destroy())
            modal.focus_set()
            modal.wait_window()
        except Exception as e:
            print("Erro ao abrir modal de ajuda (Média de Custo):", e)

    def _on_tab_changed(self, event):
        aba_atual = self.notebook.index(self.notebook.select())
        conn = self.conectar_banco()
        if not conn:
            return

        try:
            if aba_atual == 1:  # Sudeste
                for widget in self.aba_sudeste.winfo_children():
                    widget.destroy()
                criar_media_custo_sudeste(self.aba_sudeste, conn)
            elif aba_atual == 2:  # Centro-Oeste/Nordeste
                for widget in self.aba_centro_oeste_nordeste.winfo_children():
                    widget.destroy()
                criar_media_aba_centro_oeste(self.aba_centro_oeste_nordeste, conn)
        finally:
            try:
                conn.close()
            except Exception:
                pass

    def _popular_aba_icm(self):
        # usa self.conn se disponível, caso contrário tenta abrir uma conexão temporária
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None

        if not conn:
            return

        try:
            # limpa
            for w in self.tree_icm.get_children():
                self.tree_icm.delete(w)

            produtos = self.buscar_produtos()
            for produto in produtos:
                nome_produto = produto[0]
                estoque = self.buscar_estoque(nome_produto)

                media_custo_estoque_str = self.calcular_media_ponderada(nome_produto)
                try:
                    media_custo_estoque = float(media_custo_estoque_str.replace("R$ ", "").replace(".", "").replace(",", "."))
                except Exception:
                    media_custo_estoque = 0.0

                custo_empresa = self.calcular_custo_empresa(nome_produto, media_custo_estoque)
                custo_5 = self.calcular_custo_5(custo_empresa, 5)
                custo_10 = self.calcular_custo_10(custo_empresa)
                custo_15 = self.calcular_custo_15(custo_empresa)
                custo_20 = self.calcular_custo_20(custo_empresa)

                tag_cor = ""
                if estoque.strip() == "0,000 kg" or not estoque.strip():
                    tag_cor = "estoque_vazio"

                self.tree_icm.insert("", "end", values=(nome_produto, estoque, media_custo_estoque_str, custo_empresa, custo_5, custo_10, custo_15, custo_20), tags=(tag_cor,))

            total_estoque = self.calcular_total_estoque()
            media_ponderada_total, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
            media_5 = self.calcular_media_total_5()
            media_10 = self.calcular_media_total_10()
            media_15 = self.calcular_media_total_15()
            media_20 = self.calcular_media_total_20()

            self.tree_icm.tag_configure("total", background="yellow", font=("Arial", 12, "bold"), foreground="red")
            self.tree_icm.insert("", "end", values=("TOTAL", total_estoque, media_ponderada_total, media_ponderada_empresa, media_5, media_10, media_15, media_20), tags=("total",))
            self.tree_icm.tag_configure("estoque_vazio", foreground="red")
        finally:
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def conectar_banco(self):
        try:
            conn = conectar()
            return conn
        except Exception as e:
            print("Erro ao conectar ao banco:", e)
            return None

    def buscar_produtos(self):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None

        if not conn:
            return []

        try:
            cursor = conn.cursor()
            cursor.execute("SELECT nome FROM produtos ORDER BY nome;")
            produtos = cursor.fetchall()
            cursor.close()
            return produtos
        except Exception as e:
            print("Erro ao buscar nome_produtos:", e)
            return []
        finally:
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def buscar_estoque(self, nome_produto):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None

        if not conn:
            return "0,000 kg"

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT SUM(eq.quantidade_estoque)
                FROM estoque_quantidade eq
                JOIN somar_produtos sp ON eq.id_produto = sp.id
                WHERE sp.produto = %s
            """, (nome_produto,))
            quantidade = cursor.fetchone()[0]
            cursor.close()

            if quantidade:
                return f"{quantidade:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"
            else:
                return "0,000 kg"
        except Exception as e:
            print(f"Erro ao buscar estoque para o produto {nome_produto}: {e}")
            return "0,000 kg"
        finally:
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def calcular_media_ponderada(self, nome_produto):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None

        if not conn:
            return "R$ 0,00"

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT SUM(eq.quantidade_estoque * eq.custo_total), SUM(eq.quantidade_estoque)
                FROM estoque_quantidade eq
                JOIN somar_produtos sp ON eq.id_produto = sp.id
                WHERE sp.produto = %s
            """, (nome_produto,))

            resultado = cursor.fetchone()
            cursor.close()

            if resultado:
                soma_custo_ponderado, soma_quantidade = resultado

                soma_custo_ponderado = float(soma_custo_ponderado) if soma_custo_ponderado is not None else 0.0
                soma_quantidade = float(soma_quantidade) if soma_quantidade is not None else 0.0

                if soma_quantidade > 0:
                    media_ponderada = soma_custo_ponderado / soma_quantidade
                    custo_calculado = (media_ponderada * 0.7275) / 0.7875
                    return "R$ " + f"{custo_calculado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                else:
                    return "R$ 0,00"
            else:
                return "R$ 0,00"
        except Exception as e:
            print(f"Erro ao calcular média ponderada para o produto {nome_produto}: {e}")
            return "R$ 0,00"
        finally:
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def calcular_custo_empresa(self, nome_produto, media_custo_estoque=None):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None
        cursor = None

        if not conn:
            return "R$ 0,00"

        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT eq.quantidade_estoque, eq.custo_total
                FROM estoque_quantidade eq
                JOIN somar_produtos sp ON eq.id_produto = sp.id
                WHERE sp.produto = %s
            """, (nome_produto,))
            produtos = cursor.fetchall()

            soma_custo_ponderado = 0.0
            soma_quantidade = 0.0

            for produto in produtos:
                peso = produto[0]
                custo = produto[1]
                if peso is not None and custo is not None:
                    soma_custo_ponderado += float(peso) * float(custo)
                    soma_quantidade += float(peso)

            if soma_quantidade > 0:
                media_ponderada = soma_custo_ponderado / soma_quantidade
            else:
                media_ponderada = 0.0

            if media_ponderada == 0:
                return "R$ 0,00"

            cursor.execute("""
                SELECT custo_empresa FROM somar_produtos WHERE produto = %s
            """, (nome_produto,))
            custo_empresa = cursor.fetchone()

            if custo_empresa and custo_empresa[0] is not None:
                custo_empresa_val = float(custo_empresa[0])
                resultado = float(media_ponderada) + custo_empresa_val
                resultado_final = (resultado * 0.7275) / 0.7875
                return "R$ " + f"{resultado_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                return "R$ 0,00"
        except Exception as e:
            print(f"Erro ao calcular custo_empresa para o produto {nome_produto}: {e}")
            return "R$ 0,00"
        finally:
            try:
                if cursor:
                    cursor.close()
            except Exception:
                pass
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def calcular_custo_5(self, custo_base, percentual=5):
        try:
            if isinstance(custo_base, str):
                custo_base = float(custo_base.replace("R$", "").replace(".", "").replace(",", ".").strip())
            custo_final = custo_base / (1 - percentual / 100.0)
            return "R$ " + f"{custo_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception as e:
            print(f"Erro ao calcular acréscimo de {percentual}%: {e}")
            return "R$ 0,00"

    def calcular_custo_10(self, custo_base):
        return self.calcular_custo_5(custo_base, 10)

    def calcular_custo_15(self, custo_base):
        return self.calcular_custo_5(custo_base, 15)

    def calcular_custo_20(self, custo_base):
        return self.calcular_custo_5(custo_base, 20)

    def calcular_total_estoque(self):
        total_estoque = 0.0
        produtos = self.buscar_produtos()

        for produto in produtos:
            nome_produto = produto[0]
            estoque = self.buscar_estoque(nome_produto)
            estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))
            total_estoque += estoque_float

        return f"{total_estoque:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " kg"

    def calcular_media_ponderada_total_com_empresa(self):
        produtos = self.buscar_produtos()

        soma_ponderada_custo = 0.0
        soma_ponderada_empresa = 0.0
        soma_estoque = 0.0

        for produto in produtos:
            nome_produto = produto[0]
            estoque = self.buscar_estoque(nome_produto)
            estoque_float = float(estoque.replace(" kg", "").replace(".", "").replace(",", "."))

            if estoque_float > 0:
                media_custo_str = self.calcular_media_ponderada(nome_produto)
                media_custo = float(media_custo_str.replace("R$ ", "").replace(".", "").replace(",", "."))

                if media_custo == 0:
                    custo_empresa = 0.0
                else:
                    custo_empresa_str = self.calcular_custo_empresa(nome_produto)
                    custo_empresa = float(custo_empresa_str.replace("R$ ", "").replace(".", "").replace(",", "."))

                soma_ponderada_custo += estoque_float * media_custo
                soma_ponderada_empresa += estoque_float * custo_empresa
                soma_estoque += estoque_float

        if soma_estoque == 0:
            return "R$ 0,000", "R$ 0,000"

        media_ponderada_total = soma_ponderada_custo / soma_estoque
        media_ponderada_empresa = soma_ponderada_empresa / soma_estoque

        media_ponderada_total_str = f"R$ {media_ponderada_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        media_ponderada_empresa_str = f"R$ {media_ponderada_empresa:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        return media_ponderada_total_str, media_ponderada_empresa_str

    def calcular_media_total_5(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
        media_5 = media_ponderada_empresa_float / 0.95
        media_5_str = f"R$ {media_5:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return media_5_str

    def calcular_media_total_10(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
        media_10 = media_ponderada_empresa_float / 0.9
        media_10_str = f"R$ {media_10:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return media_10_str

    def calcular_media_total_15(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
        media_15 = media_ponderada_empresa_float / 0.85
        media_15_str = f"R$ {media_15:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return media_15_str

    def calcular_media_total_20(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace("R$ ", "").replace(".", "").replace(",", "."))
        media_20 = media_ponderada_empresa_float / 0.8
        media_20_str = f"R$ {media_20:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return media_20_str

    def _on_closing(self):
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            if self.conn:
                try:
                    self.conn.close()
                except Exception:
                    pass
            try:
                self.janela_media_custo.destroy()
            except Exception:
                pass
            if self.main_window:
                try:
                    self.main_window.destroy()
                except Exception:
                    pass
            try:
                sys.exit(0)
            except Exception:
                pass

    def voltar_para_menu(self, janela_media_custo, main_window):
        try:
            main_window.deiconify()
            main_window.state("zoomed")
            main_window.lift()
            main_window.update_idletasks()
            try:
                main_window.focus_force()
            except Exception:
                pass
        except Exception:
            pass

        def _cleanup_and_destroy():
            try:
                if hasattr(janela_media_custo, "cursor") and getattr(janela_media_custo, "cursor"):
                    try:
                        janela_media_custo.cursor.close()
                    except Exception:
                        pass
                if hasattr(janela_media_custo, "conn") and getattr(janela_media_custo, "conn"):
                    try:
                        janela_media_custo.conn.close()
                    except Exception:
                        pass
            finally:
                try:
                    if hasattr(janela_media_custo, "after"):
                        janela_media_custo.after(0, janela_media_custo.destroy)
                    else:
                        janela_media_custo.destroy()
                except Exception:
                    try:
                        janela_media_custo.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()

# Compatibilidade: função de módulo que instancia a classe
def criar_media_custo(font_size=12, main_window=None):
    return MediaCusto(main_window=main_window, font_size=font_size)
