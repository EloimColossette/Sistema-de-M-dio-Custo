import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar
from conexao_db import conectar  # Importa a função para conectar ao banco
from relatorio_entrada import RelatorioEntradaApp
from relatorio_resumo import RelatorioResumoApp
from logos import aplicar_icone
import pandas as pd
from tkinter import filedialog
import xlsxwriter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import threading

class RelatorioApp(tk.Toplevel):  # Herda de Toplevel
    def __init__(self, root):
        super().__init__(root)  # Associa o Toplevel à janela principal (menu)
        self.title("Relatórios de Produtos")
        self.state("zoomed")

        aplicar_icone(self, r"C:\Sistema\logos\Kametal.ico")

        # Configuração do estilo do ttk com uma paleta mais profissional
        self.configurar_estilos()

        # Criar um frame para os botões (Voltar e Exportar)
        self.frame = tk.Frame(self)
        self.frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Criar o botão Voltar
        self.botao_voltar = ttk.Button(self.frame, text="Voltar", command=self.voltar, style="Voltar.TButton")
        self.botao_voltar.pack(side=tk.LEFT, padx=10, pady=10)

        # Botão Exportar para PDF
        self.botao_exportar_pdf = ttk.Button(
            self.frame,
            text="Exportar para PDF",
            command=self.exportar_para_pdf,
            style="ExportarPDF.TButton"
        )
        self.botao_exportar_pdf.pack(side=tk.RIGHT, padx=10, pady=10)

        # Botão Exportar para Excel
        self.botao_exportar = ttk.Button(
            self.frame, text="Exportar para Excel", 
            command=self.exportar_para_excel, style="ExportarExcel.TButton"
        )
        self.botao_exportar.pack(side=tk.RIGHT, padx=10, pady=10)

        # --- Botão de Ajuda discreto (❓) ---
        self.botao_ajuda = tk.Button(
            self.frame,
            text="❓",
            fg="white",
            bg="#2c3e50",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            relief="flat",
            width=3,
            command=self._abrir_ajuda_relatorios_modal
        )
        self.botao_ajuda.pack(side=tk.RIGHT, padx=6, pady=10)

        # efeito hover
        self.botao_ajuda.bind("<Enter>", lambda e: self.botao_ajuda.config(bg="#3b5566"))
        self.botao_ajuda.bind("<Leave>", lambda e: self.botao_ajuda.config(bg="#2c3e50"))

        # Atalho F1 para abrir ajuda (quando esta janela estiver ativa)
        try:
            self.bind("<F1>", lambda e: self._abrir_ajuda_relatorios_modal())
        except Exception:
            pass

        # Tooltip (se desejar hover)
        try:
            if hasattr(self, "_create_tooltip"):
                self._create_tooltip(self.botao_ajuda, "Ajuda — Relatorios (F1)")
        except Exception:
            pass

       # Criar abas (Notebook)
        self.abas = ttk.Notebook(self)
        self.abas.pack(fill=tk.BOTH, expand=True)

        # Aba Relatório de Saída
        self.aba_saida = ttk.Frame(self.abas)
        self.abas.add(self.aba_saida, text="Relatório de Saída")

        # Aba Relatório de Entrada
        self.aba_entrada = ttk.Frame(self.abas)
        self.abas.add(self.aba_entrada, text="Relatório de Entrada")

        # Aba Relatório Resumo
        self.aba_resumo = ttk.Frame(self.abas)
        self.abas.add(self.aba_resumo, text="Relatório Resumo")

        # Instancia a interface do relatório resumo na aba_resumo
        self.relatorioResumo = RelatorioResumoApp(self.aba_resumo)
        self.relatorioResumo.pack(expand=True, fill="both")

        # Agora, instancie a interface do relatório de entrada, passando o relatório resumo
        self.relatorioEntrada = RelatorioEntradaApp(self, self.aba_entrada, self.relatorioResumo)

        # Cria a interface para a aba de saída
        self.criar_interface(self.aba_saida, tipo="saida")

        self.carregar_produtos_base()

        # Configurar o evento de fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def configurar_estilos(self):
        """Define estilos globais para o Treeview e os botões."""
        estilo = ttk.Style()
        estilo.theme_use("alt")

        # Estilo básico para linhas da tabela
        estilo.configure("Saida.Treeview",
                 font=("Arial", 10),
                 rowheight=35,
                 background="white",
                 foreground="black",
                 fieldbackground="white")

        estilo.configure("Saida.Treeview.Heading",
                        font=("Arial", 10, "bold"))

        # Estilo dos botões
        estilo.configure("Voltar.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#FF6347",  # Cor laranja para Voltar
                         foreground="white",
                         font=("Arial", 10, "bold"),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("Voltar.TButton", 
                   background=[("active", "#FF4500")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])

        estilo.configure("ExportarExcel.TButton",
                         padding=(5, 2),
                         relief="raised",
                         background="#1E90FF",  # Cor azul para Exportar
                         foreground="white",
                         font=("Arial", 10, "bold"),
                         borderwidth=2,
                         highlightbackground="#34495e",
                         highlightthickness=1)
        estilo.map("ExportarExcel.TButton", 
                   background=[("active", "#1C86EE")],
                   foreground=[("active", "white")],
                   relief=[("pressed", "sunken"), ("!pressed", "raised")])

        # Estilo para linha de total
        estilo.configure("total_arame.Treeview", 
                         background="#FF6347",  # Cor laranja para arame
                         foreground="white", 
                         font=("Arial", 12, "bold"))
        
        estilo.configure("total_fio.Treeview", 
                         background="#1E90FF",  # Cor azul para fio
                         foreground="white", 
                         font=("Arial", 12, "bold"))
        
        estilo.configure("total_geral.Treeview", 
                         background="#FFD700",  # Cor dourada para o total geral
                         foreground="black", 
                         font=("Arial", 12, "bold"))

        # Verificar se o estilo foi aplicado corretamente
        print("Estilo configurado!")

    def _create_tooltip(self, widget, text, delay=450):
        """Tooltip com quebra de linha automática e ajuste para não sair da tela."""
        tooltip = {"win": None, "after_id": None}

        def show():
            if tooltip["win"] or not widget.winfo_exists():
                return

            # cria janela do tooltip
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.attributes("-topmost", True)

            # label com wrap para quebrar linhas
            label = tk.Label(
                win,
                text=text,
                bg="#333333",
                fg="white",
                font=("Segoe UI", 9),
                bd=0,
                padx=6,
                pady=4,
                wraplength=300  # máx. largura do tooltip (pixels)
            )
            label.pack()

            # calcula posição inicial
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 6

            # força update para medir o tamanho real
            win.update_idletasks()
            w, h = win.winfo_width(), win.winfo_height()

            # limites da tela
            screen_w = win.winfo_screenwidth()
            screen_h = win.winfo_screenheight()

            # ajusta posição horizontal se ultrapassar borda direita
            if x + w > screen_w:
                x = screen_w - w - 10
            if x < 0:
                x = 10

            # ajusta posição vertical se ultrapassar borda inferior
            if y + h > screen_h:
                y = widget.winfo_rooty() - h - 6  # mostra acima do widget

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

    def _abrir_ajuda_relatorios_modal(self, contexto=None):
        """
        Modal de ajuda para Relatórios (Saída / Entrada / Resumo).
        Implementação baseada no padrão grid+tkraise (estável) e com scrollbar
        dedicada apenas para o painel de explicações (lado direito).
        """
        try:
            from tkinter import ttk

            # --- Janela modal ---
            modal = tk.Toplevel(self)
            modal.title("Ajuda — Relatórios")
            modal.transient(self)
            modal.grab_set()
            modal.configure(bg="white")

            # Dimensões / centralização
            w, h = 900, 640
            x = max(0, (modal.winfo_screenwidth() // 2) - (w // 2))
            y = max(0, (modal.winfo_screenheight() // 2) - (h // 2))
            modal.geometry(f"{w}x{h}+{x}+{y}")
            modal.minsize(760, 480)

            # Ícone (silencioso se falhar)
            try:
                aplicar_icone(modal, r"C:\Sistema\logos\Kametal.ico")
            except Exception:
                pass

            # --- Cabeçalho ---
            header = tk.Frame(modal, bg="#2b3e50", height=64)
            header.pack(side="top", fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="Ajuda — Relatórios", bg="#2b3e50", fg="white",
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=16)
            tk.Label(header, text="F1 abre esta ajuda — Esc fecha", bg="#2b3e50",
                    fg="#cbd7e6", font=("Segoe UI", 10)).pack(side="left", padx=8, pady=10)

            ttk.Separator(modal, orient="horizontal").pack(fill="x")

            # --- Corpo: navegação esquerda + conteúdo à direita ---
            body = tk.Frame(modal, bg="white")
            body.pack(fill="both", expand=True, padx=14, pady=12)

            # NAV FRAME (lado esquerdo) - sem scrollbar visual
            nav_frame = tk.Frame(body, width=260, bg="#f6f8fa")
            nav_frame.pack(side="left", fill="y", padx=(0,12), pady=2)
            nav_frame.pack_propagate(False)
            tk.Label(nav_frame, text="Seções", bg="#f6f8fa", font=("Segoe UI", 10, "bold")).pack(anchor="nw", pady=(10,6), padx=12)

            # container do listbox (sem scrollbar)
            nav_list_container = tk.Frame(nav_frame, bg="#f6f8fa")
            nav_list_container.pack(fill="both", expand=True, padx=10, pady=(0,10))

            sections = [
                "Visão Geral",
                "Relatório de Saída",
                "Relatório de Entrada",
                "Relatório Resumo",
                "Filtros e Pesquisa",
                "Editar Estoque / Enviar para Resumo",
                "Exportação (Excel / PDF)",
                "Boas Práticas / Segurança",
                "FAQ"
            ]

            listbox = tk.Listbox(nav_list_container, bd=0, highlightthickness=0, activestyle="none",
                                font=("Segoe UI", 10), selectmode="browse", exportselection=False,
                                bg="#ffffff")
            listbox.pack(fill="both", expand=True)
            for s in sections:
                listbox.insert("end", s)

            # --- Area de conteúdo (usa grid + tkraise para painéis distintos) ---
            content_frame = tk.Frame(body, bg="white")
            content_frame.pack(side="right", fill="both", expand=True)
            content_frame.rowconfigure(0, weight=1)
            content_frame.columnconfigure(0, weight=1)

            # Painel "geral" (padrão para a maioria das seções)
            general_frame = tk.Frame(content_frame, bg="white")
            general_frame.grid(row=0, column=0, sticky="nsew")
            general_frame.rowconfigure(0, weight=1)
            general_frame.columnconfigure(0, weight=1)

            txt_general = tk.Text(general_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_general = tk.Scrollbar(general_frame, command=txt_general.yview)  # scrollbar ligada ao Text
            txt_general.configure(yscrollcommand=sb_general.set)
            txt_general.grid(row=0, column=0, sticky="nsew")
            sb_general.grid(row=0, column=1, sticky="ns")

            # Se quiser painéis específicos por seção, crie frames adicionais como abaixo.
            # Aqui criamos um painel específico para 'Editar Estoque' (exemplo), para demonstrar como alternar painéis.
            editar_frame = tk.Frame(content_frame, bg="white")
            editar_frame.grid(row=0, column=0, sticky="nsew")
            editar_frame.rowconfigure(0, weight=1)
            editar_frame.columnconfigure(0, weight=1)

            txt_editar = tk.Text(editar_frame, wrap="word", font=("Segoe UI", 11), bd=0, padx=12, pady=12)
            sb_editar = tk.Scrollbar(editar_frame, command=txt_editar.yview)
            txt_editar.configure(yscrollcommand=sb_editar.set)
            txt_editar.grid(row=0, column=0, sticky="nsew")
            sb_editar.grid(row=0, column=1, sticky="ns")

            # --- Conteúdo das seções (strings preparadas) ---
            # Aqui estão os textos de cada seção. Substitua/edite conforme necessário.
            # OBS importante: incluí explicações claras sobre formatação automática da data e tratamento de vírgula no editar.
            contents = {}

            contents["Visão Geral"] = (
                "Visão Geral\n\n"
                "Esta janela reúne três relatórios principais: Relatório de Saída, Relatório de Entrada e Relatório Resumo.\n"
                "Objetivo: permitir conferência, consolidação e exportação das movimentações de estoque por período.\n\n"
                "O fluxo típico:\n"
                "1. Selecione o mês (MM/YYYY) e a Base do Produto.\n"
                "2. Clique em 'Gerar Relatório' para carregar os dados na Treeview.\n"
                "3. Use 'Pesquisar' para filtrar os resultados já carregados (filtro local).\n"
                "4. Use 'Calcular Totais' para visualizar somas/medias; 'Editar Estoque' para ajustes locais; 'Enviar para Resumo' para consolidar.\n\n"
                "Observação: a janela de Resumo serve como painel consolidado — pense nela como a visão\n"
                "final antes de exportar para contabilidade ou enviar resumos por e-mail/PDF."
            )

            # Relatório de Saída: inclui a nota sobre digitar data sem barra (o sistema formata automaticamente)
            contents["Relatório de Saída"] = (
                "Relatório de Saída — passo a passo\n\n"
                "O que mostra: notas fiscais / saídas por período (Data, Nº NF, Produto, Peso).\n\n"
                "Como usar:\n"
                "1) Escolha o mês na caixa de data. Você pode digitar 'MMYYYY' (ex.: 092025) sem a barra — o sistema irá inserir a barra automaticamente e apresentar '09/2025'.\n"
                "2) Escolha a Base do Produto e clique em 'Gerar Relatório'.\n"
                "3) Verifique as colunas principais: Data / NF / Produto / Peso. Linhas de TOTAL aparecem ao final.\n\n"
                "Totais específicos:\n"
                "- O relatório soma automaticamente pesos por categorias (por exemplo, subtotais 'arame' e 'fio') e um total geral.\n"
                "- Use a pesquisa para ver subtotais específicos (ex.: pesquisar 'arame' retornará apenas linhas e o subtotal de arame).\n\n"
                "Conferência de divergências:\n"
                "- Se encontrar diferença entre documento fiscal e sistema, verifique o número da NF e o lote/entrada correspondente.\n"
                "- Use a coluna 'Observações' para registrar uma justificativa temporária.\n\n"
                "Dica: sempre gere o relatório antes de exportar para garantir que os dados exibidos sejam os últimos."
            )

            # Relatório de Entrada: também aceita data sem barra
            contents["Relatório de Entrada"] = (
                "Relatório de Entrada — passo a passo e cálculo de custo\n\n"
                "O que mostra: entradas de mercadorias (Data, Nº NF, Produto, Peso Líquido, Qtd Estoque, Custo Total).\n\n"
                "Como usar:\n"
                "1) Selecione o mês na caixa de data. Basta digitar 'MMYYYY' (ex.: 092025) — o sistema completa para 'MM/YYYY' automaticamente.\n"
                "2) Clique em 'Gerar Relatório' para carregar as entradas do período.\n"
                "3) Para calcular médias ponderadas por produto (quando há várias NFs do mesmo produto), utilize 'Calcular Totais' na tela de Entrada.\n\n"
                "Cálculo de média ponderada (conceito):\n"
                "- Para um mesmo produto com várias entradas: Custo Médio = (Σ (qtd_i * custo_unit_i)) / Σ(qtd_i).\n"
                "- Use arredondamento consistente (ex.: Decimal quantize com 4 casas para custo unitário, 2/3 casas para valores dependendo do padrão da empresa).\n\n"
                "Correções e auditoria:\n"
                "- Se precisar modificar alguma NF, selecione a linha e use 'Editar Estoque' para ajuste local.\n"
                "- Para persistir alterações no banco, implemente UPDATE no método de salvar da edição (veja seção 'Editar Estoque')."
            )

            contents["Relatório Resumo"] = (
                "Relatório Resumo — o que é e como usar\n\n"
                "O Resumo é um painel condensado dos produtos: mostra Nome do Produto, Qtd Estoque e Custo Médio.\n\n"
                "Fluxo de trabalho comum:\n"
                "1) Gere Relatório de Entrada/ Saída e identifique os produtos que precisam ser consolidados.\n"
                "2) Selecione linha(s) e clique 'Enviar para Resumo'. Isso preenche/atualiza a linha do produto no Resumo.\n"
                "3) No Resumo você pode revisar, ajustar manualmente Qtd ou Custo Médio e então exportar a visão consolidada.\n\n"
                "Regras e prioridades:\n"
                "- Valores enviados para o Resumo sobrescrevem apenas os campos indicados (p.ex. Qtd Estoque e Custo Médio) — mantenha log de alterações.\n"
                "- O Resumo é ideal para relatórios gerenciais/fechamento mensal antes de gerar os arquivos para contabilidade.\n\n"
                "Exportação: gere o Resumo atualizado e use 'Exportar Excel' para criar uma planilha com a aba 'Resumo' separada."
            )

            contents["Filtros e Pesquisa"] = (
                "Filtros e Pesquisa — como extrair rapidamente o que precisa\n\n"
                "Pesquisa local (campo 'Pesquisar'):\n"
                "- Filtra os registros já carregados na Treeview — não faz nova query ao banco.\n"
                "- Suporta busca por partes do texto (substring). Exemplos: '670312', 'Solda', 'arame'.\n\n"
                "Formato da data (atenção):\n"
                "- Você não precisa digitar a barra '/'. Digite 'MMYYYY' (ex.: 092025) ou 'MM/YYYY' — o campo aceitará ambos e irá formatar automaticamente para 'MM/YYYY'.\n\n"
                "Filtros recomendados antes de gerar:\n"
                "- Mês — obrigatório para relatórios mensais.\n"
                "- Base do Produto — reduz volume e acelera consulta.\n\n"
                "Boas práticas:\n"
                "- Pesquise depois de carregar os dados para evitar consultas desnecessárias.\n"
                "- Para buscas complexas (ex.: intervalo de notas, fornecedor específico), adicione parâmetros na query que gera o relatório."
            )

            # Editar Estoque: indica que o campo aceita vírgula e ponto e que o sistema normaliza
            contents["Editar Estoque / Enviar para Resumo"] = (
                "Editar Estoque / Enviar para Resumo — instruções detalhadas\n\n"
                "Editar Estoque (passo a passo):\n"
                "1) Selecione exatamente a linha do item que deseja alterar (não selecione a linha TOTAL).\n"
                "2) Clique em 'Editar Estoque' — abrirá caixa de diálogo para alterar quantidade (qtd) e, opcionalmente, custo.\n"
                "3) Valide as entradas: quantidade não pode ser negativa; custo deve ser número positivo.\n"
                "   Observação importante: não é necessário digitar vírgula manualmente — o campo aceita números com vírgula ou ponto e o sistema converte/normaliza automaticamente para o formato interno.\n"
                "   Exemplos aceitos: '1000', '1234.56' ou '1234,56'. Evite separadores de milhares (ex.: '1.234,56'); prefira '1234.56' ou '1234,56'.\n"
                "4) Ao confirmar, a interface atualiza a Treeview e recalcula os totais.\n\n"
                "Persistindo a alteração no banco (recomendação):\n"
                "- Para alterar permanentemente: execute um SQL UPDATE no registro identificado por chave (ex.: produto_id, nf, lote).\n"
                "  Exemplo SQL (conceitual):\n"
                "    UPDATE estoque SET quantidade = :nova_qtd, custo_medio = :novo_custo WHERE produto_id = :id AND lote = :lote;\n"
                "- Após UPDATE, re-execute a query do relatório para recarregar os dados e evitar inconsistências.\n\n"
                "Enviar para Resumo:\n"
                "- Selecionando uma linha e clicando 'Enviar para Resumo' os campos relevantes (Qtd Estoque, Custo Médio) são copiados para a aba Resumo. Se o produto já existe, atualize a linha existente (merge), caso contrário inclua nova linha.\n\n"
                "Aviso: a implementação atual pode apenas alterar a visualização local; verifique se há persistência no banco se quiser que a alteração seja definitiva."
            )

            contents["Exportação (Excel / PDF)"] = (
                "Exportação — melhores práticas e opções\n\n"
                "Exportar para Excel (.xlsx):\n"
                "- Gera um arquivo com abas: 'Saída', 'Entrada', 'Resumo'. Cada aba contém todos os registros atualmente carregados.\n"
                "- Nome sugerido: relatorio_<tipo>_YYYYMMDD_HHMMSS.xlsx (ex.: relatorio_saida_20250929_101523.xlsx).\n"
                "- Inclua uma aba adicional 'META' ou 'TOTAIS' com os agregados principais (somas/medias) para facilitar conferência.\n"
                "- Recomendação técnica: use pandas.DataFrame + df.to_excel(...) ou openpyxl para controlar estilos e colunas.\n\n"
                "Exportar para PDF:\n"
                "- As duas abordagens comuns:\n"
                "  1) Construir PDF direto (reportlab) — controle preciso de layout, bom para relatórios prontos para impressão.\n"
                "  2) Gerar HTML formatado e imprimir/converter em PDF (mais simples para tabelas largas).\n"
                "- Configure largura de colunas e quebras de página (cada N linhas) para evitar cortes estranhos.\n\n"
                "Permissões e locais:\n"
                "- Verifique permissões de escrita antes de salvar.\n"
                "- Sempre gere o relatório primeiro e só então exporte (garante que os dados exportados reflitam a view atual)."
            )

            contents["Cálculo Totais"] = (
                "Botão 'Calcular Totais' — o que deve fazer\n\n"
                "Objetivo: agregar e exibir rapidamente os principais números do relatório ativo (quantidade, peso, valor).\n\n"
                "Itens a calcular normalmente:\n"
                "- Soma das quantidades (Σ qtd).\n"
                "- Soma dos pesos (Σ peso).\n"
                "- Soma do valor total (Σ qtd * custo_unitário) ou (Σ valor_nota).\n"
                "- Totais por categoria (ex.: subtotal 'arame', subtotal 'fio') e total geral.\n\n"
                "Apresentação:\n"
                "- Exibir os totais em labels no rodapé da janela OU inserir uma linha 'TOTAL' fixa ao final da Treeview.\n"
                "- Use Decimal com quantização adequada para evitar erros de ponto flutuante ao somar valores monetários.\n\n"
                "Observação: se houver filtros ativos (pesquisa), o cálculo deve considerar apenas as linhas visíveis (comportamento esperado)."
            )

            contents["Boas Práticas / Segurança"] = (
                "Boas práticas e segurança operacional\n\n"
                "- Faça backup do banco antes de operações em lote (importações ou atualizações massivas).\n"
                "- Valide formatos de data (MM/YYYY) antes de executar queries para evitar consultas vazias.\n"
            )

            contents["FAQ"] = (
                "FAQ — Perguntas Frequentes\n\n"
                "Q: Posso exportar apenas uma aba?\n"
                "A: Sim. A funcionalidade de PDF exporta a aba ativa; o Excel pode ser configurado para incluir apenas as abas desejadas.\n\n"
                "Q: Como recuperar dados antes de uma edição equivocada?\n"
                "A: Restaure o backup do banco ou implemente log de alterações para reverter (UPDATE inverso).\n\n"
                "Q: Por que alguns totais não batem com o ERP/contabilidade?\n"
                "A: Possíveis causas: difusão de critérios (peso bruto vs peso líquido), notas canceladas não filtradas, ou divergências de base de produto.\n\n"
                "Q: O que significa 'Enviar para Resumo'?\n"
                "A: Copia/atualiza os campos selecionados (Qtd, Custo) na aba Resumo para consolidação final antes da exportação.\n\n"
                "Se precisar, posso adicionar respostas específicas da sua instalação (ex.: validações extras, regras de negócio da contabilidade)."
            )

            # --- Função para exibir a seção correta ---
            def mostrar_secao(key):
                """
                Exibe o conteúdo da seção:
                - Para a maioria das seções usamos general_frame (txt_general + sb_general).
                - Para seções que carecem de painel específico (ex.: 'Editar Estoque') podemos usar editar_frame.
                - Chamamos tkraise() para trazer o painel visível.
                """
                if key == "Editar Estoque / Enviar para Resumo":
                    # exemplo: usar painel dedicado para "Editar Estoque"
                    txt_editar.configure(state="normal")
                    txt_editar.delete("1.0", "end")
                    txt_editar.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_editar.configure(state="disabled")
                    txt_editar.yview_moveto(0)
                    editar_frame.tkraise()
                else:
                    txt_general.configure(state="normal")
                    txt_general.delete("1.0", "end")
                    txt_general.insert("1.0", contents.get(key, "Conteúdo não disponível."))
                    txt_general.configure(state="disabled")
                    txt_general.yview_moveto(0)
                    general_frame.tkraise()

            # Inicializa exibindo a primeira seção
            listbox.selection_set(0)
            mostrar_secao(sections[0])

            def on_select(evt):
                sel = listbox.curselection()
                if sel:
                    mostrar_secao(sections[sel[0]])
            listbox.bind("<<ListboxSelect>>", on_select)

            # --- Bindings do mousewheel para ajudar rolagem ---
            def _on_mousewheel_text(event):
                if hasattr(event, "delta") and event.delta:
                    # Windows / Mac
                    txt_general.yview_scroll(int(-1 * (event.delta / 120)), "units")
                    txt_editar.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    txt_general.yview_scroll(-1, "units")
                    txt_editar.yview_scroll(-1, "units")
                elif event.num == 5:
                    txt_general.yview_scroll(1, "units")
                    txt_editar.yview_scroll(1, "units")

            def _on_mousewheel_listbox(event):
                if hasattr(event, "delta") and event.delta:
                    listbox.yview_scroll(int(-1 * (event.delta / 120)), "units")
                elif event.num == 4:
                    listbox.yview_scroll(-1, "units")
                elif event.num == 5:
                    listbox.yview_scroll(1, "units")

            # bind nos widgets corretos
            txt_general.bind("<MouseWheel>", _on_mousewheel_text)
            txt_general.bind("<Button-4>", _on_mousewheel_text)
            txt_general.bind("<Button-5>", _on_mousewheel_text)
            txt_editar.bind("<MouseWheel>", _on_mousewheel_text)
            txt_editar.bind("<Button-4>", _on_mousewheel_text)
            txt_editar.bind("<Button-5>", _on_mousewheel_text)

            listbox.bind("<MouseWheel>", _on_mousewheel_listbox)
            listbox.bind("<Button-4>", _on_mousewheel_listbox)
            listbox.bind("<Button-5>", _on_mousewheel_listbox)

            # --- Rodapé com botão Fechar ---
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
            print("Erro ao abrir modal de ajuda (Relatórios):", e)

    def criar_interface(self, parent, tipo):
        """Cria os filtros e a tabela para cada aba."""
        frame_filtros = ttk.LabelFrame(parent, text="Filtros", padding=(10, 10))
        frame_filtros.pack(fill=tk.X, padx=20, pady=10)

        label_style = {"font": ("Arial", 10, "bold")}

        # Linha 0: Entrada de Data e Combobox
        tk.Label(frame_filtros, text="Mês (MM/YYYY):", **label_style).grid(
            row=0, column=0, padx=(5,2), pady=5, sticky=tk.W
        )
        entrada_data = tk.Entry(frame_filtros, width=12, font=("Arial", 10))
        entrada_data.grid(row=0, column=1, padx=(2,15), pady=5, sticky=tk.W)

        tk.Label(frame_filtros, text="Base do Produto:", **label_style).grid(
            row=0, column=2, padx=(5,2), pady=5, sticky=tk.W
        )
        combobox_produto = ttk.Combobox(frame_filtros, width=30, font=("Arial", 10))
        combobox_produto.grid(row=0, column=3, padx=(2,15), pady=5, sticky=tk.W)

        botao_relatorio = ttk.Button(
            frame_filtros, text="Gerar Relatório",
            command=lambda: self.gerar_relatorio(tipo, entrada_data, combobox_produto),
            style="ExportarExcel.TButton"
        )
        botao_relatorio.grid(row=0, column=4, padx=(10,5), pady=5, sticky=tk.W)

        # Criar um Frame para manter os elementos alinhados
        frame_pesquisa = tk.Frame(frame_filtros)
        frame_pesquisa.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)

        # Rótulo de pesquisa
        tk.Label(frame_pesquisa, text="Pesquisar:", **label_style).pack(side=tk.LEFT, padx=(0, 5))

        # Campo de pesquisa
        pesquisa_entry = tk.Entry(frame_pesquisa, width=20, font=("Arial", 10))
        pesquisa_entry.pack(side=tk.LEFT)

        # Botão de pesquisa
        botao_pesquisar = ttk.Button(
            frame_pesquisa, text="Pesquisar",
            command=lambda: self.pesquisar(pesquisa_entry.get(), tabela),
            style="Voltar.TButton"
        )
        botao_pesquisar.pack(side=tk.LEFT, padx=5)

        # Tabela de Resultados
        frame_tabela = tk.Frame(parent)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        colunas = ("Data", "NF", "Produto", "Peso")
        tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", style="Saida.Treeview")

        for coluna, largura in zip(colunas, [100, 100, 300, 150]):
            tabela.heading(coluna, text=coluna)
            tabela.column(coluna, anchor=tk.CENTER, width=largura)

        scroll_y = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=tabela.yview)
        tabela.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tabela.pack(fill=tk.BOTH, expand=True)

        # Configura as tags de estilo para os totais
        tabela.tag_configure("total_arame", background="#FF6347", foreground="white", font=("Arial", 12, "bold"))
        tabela.tag_configure("total_fio", background="#1E90FF", foreground="white", font=("Arial", 12, "bold"))
        tabela.tag_configure("total_geral", background="#FFD700", foreground="black", font=("Arial", 12, "bold"))

        if tipo == "saida":
            self.entrada_data_saida = entrada_data
            self.combobox_produto_saida = combobox_produto
            self.tabela_saida = tabela
        else:
            self.entrada_data_entrada = entrada_data
            self.combobox_produto_entrada = combobox_produto
            self.tabela_entrada = tabela

        def formatar_data(event):
            """Adiciona a barra automaticamente ao digitar o mês"""
            texto = entrada_data.get()
            if len(texto) == 2 and '/' not in texto:
                entrada_data.insert(tk.END, '/')

        entrada_data.bind("<KeyRelease>", formatar_data)

    def pesquisar(self, texto_pesquisa, tabela):
        """Filtra a tabela conforme o texto inserido na barra de pesquisa e reinsere os totais conforme o termo pesquisado.
        
        - Se a pesquisa estiver em branco, exibe todos os totais (arame, fio e total geral, desde que tenham valor > 0).
        - Se o termo for 'arame' (e não 'fio'), exibe apenas o total de arame e o total geral.
        - Se o termo for 'fio' (e não 'arame'), exibe apenas o total de fio e o total geral.
        - Caso contrário, exibe somente o total geral.
        """
        texto_pesquisa = texto_pesquisa.lower().strip()
        
        if not hasattr(self, 'resultados_saida'):
            messagebox.showinfo("Informação", "Gere o relatório antes de pesquisar.")
            return
        
        # Limpa a Treeview
        for item in tabela.get_children():
            tabela.delete(item)
        
        # Filtra os registros de detalhes de acordo com a pesquisa
        if texto_pesquisa:
            resultados_filtrados = [
                reg for reg in self.resultados_saida if texto_pesquisa in " ".join(map(str, reg)).lower()
            ]
        else:
            resultados_filtrados = self.resultados_saida
        
        # Insere os registros filtrados (detalhes)
        for reg in resultados_filtrados:
            tabela.insert("", tk.END, values=(
                reg[0].strftime("%d/%m/%Y"), reg[1], reg[2], self.formatar_peso(reg[3])
            ))
        
        # Seleciona os totais a serem exibidos
        totais_para_exibir = []
        if texto_pesquisa == "":
            # Campo de pesquisa vazio: exibe todos os totais armazenados (eles já devem ter sido filtrados para não incluir arame/fio com 0)
            totais_para_exibir = self.totais_saida[:]
        else:
            # Se houver termo na pesquisa:
            # Se pesquisar "arame" e não "fio"
            if "arame" in texto_pesquisa and "fio" not in texto_pesquisa:
                for total, tag in self.totais_saida:
                    if tag in ("total_arame", "total_geral"):
                        totais_para_exibir.append((total, tag))
            # Se pesquisar "fio" e não "arame"
            elif "fio" in texto_pesquisa and "arame" not in texto_pesquisa:
                for total, tag in self.totais_saida:
                    if tag in ("total_fio", "total_geral"):
                        totais_para_exibir.append((total, tag))
            # Se pesquisar ambos ou outro termo: exibe somente o total geral
            else:
                for total, tag in self.totais_saida:
                    if tag == "total_geral":
                        totais_para_exibir.append((total, tag))
        
        # Insere os totais, se houver
        if totais_para_exibir:
            tabela.insert("", tk.END, values=("", "", "", ""))  # Linha em branco para separar os detalhes dos totais
            for total, tag in totais_para_exibir:
                tabela.insert("", tk.END, values=total, tags=(tag,))
                tabela.insert("", tk.END, values=("", "", "", ""))  # Linha em branco entre os totais (opcional)

    def voltar(self):
        """Cancela callbacks, reexibe menu e destrói a janela em background."""
        # cancela callback agendado, se houver
        if getattr(self, "_encerrar_id", None) is not None:
            try:
                self.after_cancel(self._encerrar_id)
            except Exception:
                pass

        # Reexibe o menu principal antes de fechar a janela atual
        try:
            self.master.deiconify()
            self.master.state("zoomed")
            self.master.lift()
            self.master.update_idletasks()
            try:
                self.master.focus_force()
            except Exception:
                pass
        except Exception:
            pass

        def _cleanup_and_destroy():
            try:
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
                try:
                    self.after(0, self.destroy)
                except Exception:
                    try:
                        self.destroy()
                    except Exception:
                        pass

        threading.Thread(target=_cleanup_and_destroy, daemon=True).start()

    def carregar_produtos_base(self):
        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT base_produto FROM produtos_nf ORDER BY base_produto;")
            produtos = [row[0] for row in cur.fetchall()]
            # Atualiza a combobox da aba de saída
            self.combobox_produto_saida['values'] = produtos
            if produtos:
                self.combobox_produto_saida.current(0)
            cur.close()
            conn.close()
            
            # Atualiza a combobox da aba de entrada chamando o método da instância de RelatorioEntradaApp
            if hasattr(self, 'relatorioEntrada'):
                self.relatorioEntrada.carregar_produtos_base()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar produtos base: {e}")

    def gerar_relatorio(self, tipo, entrada_data, combobox_produto):
        """Gera o relatório de entrada ou saída, baseado no tipo informado."""
        data_inserida = entrada_data.get()
        produto_selecionado = combobox_produto.get()

        try:
            mes, ano = data_inserida.split('/')
            ano, mes = int(ano), int(mes)
            data_inicio = f"{ano}-{mes:02d}-01"
            ultimo_dia = calendar.monthrange(ano, mes)[1]
            data_fim = f"{ano}-{mes:02d}-{ultimo_dia}"
        except Exception as e:
            messagebox.showerror("Erro", f"Formato de data inválido. Use 'MM/YYYY': {e}")
            return

        # Consulta SQL dinâmica baseada no tipo
        query = """
            SELECT data, numero_nf, produto_nome, peso
            FROM {}
            WHERE data BETWEEN %s AND %s AND base_produto = %s
            ORDER BY data ASC, numero_nf ASC;
        """.format("nf JOIN produtos_nf ON nf.id = produtos_nf.nf_id" if tipo == "saida"
                else "entrada JOIN produtos_entrada ON entrada.id = produtos_entrada.entrada_id")

        try:
            conn = conectar()
            cur = conn.cursor()
            cur.execute(query, (data_inicio, data_fim, produto_selecionado))
            resultados = cur.fetchall()

            self.resultados_saida = resultados  # Armazena os dados completos


            # Limpa a tabela antes de inserir novos dados
            tabela = self.tabela_saida if tipo == "saida" else self.tabela_entrada
            for item in tabela.get_children():
                tabela.delete(item)

            # Se não houver registros, exibe mensagem e sai da função
            if not resultados:
                messagebox.showinfo("Sem Registros", f"Não há registro de {tipo} para '{produto_selecionado}' no mês {data_inserida}.")
                return

            # Após calcular os totais
            total_arame = 0
            total_fio = 0
            total_geral = 0

            # Insere os registros na tabela e acumula os totais
            for reg in resultados:
                produto_nome = reg[2].lower()
                peso = reg[3]

                if "arame" in produto_nome:
                    total_arame += peso
                elif "fio" in produto_nome:
                    total_fio += peso
                total_geral += peso

                tabela.insert("", tk.END, values=(
                    reg[0].strftime("%d/%m/%Y"), reg[1], reg[2], self.formatar_peso(peso)
                ))

            # Monta a lista de totais apenas com os que possuem valor maior que 0
            totais = []
            if total_arame > 0:
                totais.append((("", "TOTAL ARAME", self.formatar_peso(total_arame), ""), "total_arame"))
            if total_fio > 0:
                totais.append((("", "TOTAL FIO", self.formatar_peso(total_fio), ""), "total_fio"))
            if total_geral > 0:
                totais.append((("", "TOTAL GERAL", self.formatar_peso(total_geral), ""), "total_geral"))

            # Armazena os totais para uso na pesquisa
            self.totais_saida = totais

            # Insere uma linha em branco após os registros, se houver totais para exibir
            if totais:
                tabela.insert("", tk.END, values=("", "", "", ""))

            # Insere os totais na tabela, separando cada um por uma linha em branco
            for total, tag in totais:
                tabela.insert("", tk.END, values=total, tags=(tag,))
                tabela.insert("", tk.END, values=("", "", "", ""))

            cur.close()
            conn.close()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao executar a consulta: {e}")

    def formatar_peso(self, valor):
        return f"{valor:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") + " Kg"

    def exportar_para_excel(self):
        """Exporta os relatórios de saída, entrada e resumo para um arquivo Excel com abas separadas."""
        try:
            nome_padrao = "Relatorio_item_grupo.xlsx"
            caminho_arquivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=nome_padrao,
                filetypes=[("Arquivo Excel", "*.xlsx")],
                title="Salvar Relatório"
            )

            if not caminho_arquivo:
                return  # Usuário cancelou a ação

            workbook = xlsxwriter.Workbook(caminho_arquivo)

            # Aba "Relatório de Saída" (sem coluna Produto)
            if self.tabela_saida.get_children():
                aba_saida = workbook.add_worksheet("Relatório de Saída")
                self.exportar_tabela_para_aba(self.tabela_saida, aba_saida, incluir_produto=False)

            # Aba "Relatório de Entrada" (sem coluna Produto)
            if hasattr(self.relatorioEntrada, "tabela_entrada") and self.relatorioEntrada.tabela_entrada.get_children():
                aba_entrada = workbook.add_worksheet("Relatório de Entrada")
                self.exportar_tabela_para_aba(self.relatorioEntrada.tabela_entrada, aba_entrada, incluir_produto=False)

            # Aba "Resumo" (com coluna Produto)
            if self.relatorioResumo.tree.get_children():
                aba_resumo = workbook.add_worksheet("Resumo")
                self.exportar_tabela_para_aba(self.relatorioResumo.tree, aba_resumo, incluir_produto=True)

            workbook.close()
            messagebox.showinfo("Sucesso", f"Relatório salvo como:\n{caminho_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

    def exportar_tabela_para_aba(self, tabela, aba, incluir_produto=False):
        """
        Exporta os dados de uma Treeview para uma aba do Excel.
        
        :param tabela: Treeview a ser exportada.
        :param aba: Aba do Excel onde os dados serão escritos.
        :param incluir_produto: Se True, adiciona a coluna do nome do produto (#0), senão remove.
        """
        colunas = tabela["columns"]

        if incluir_produto:
            cabecalhos = ["Produto"] + [tabela.heading(col)["text"] for col in colunas]
        else:
            cabecalhos = [tabela.heading(col)["text"] for col in colunas]

        # Escreve os cabeçalhos no Excel
        for col_idx, cabecalho in enumerate(cabecalhos):
            aba.write(0, col_idx, cabecalho)

        # Escreve os dados das tabelas
        for row_idx, item in enumerate(tabela.get_children(), start=1):
            valores = tabela.item(item, "values")
            nome_produto = tabela.item(item, "text")  # Nome do produto na coluna #0

            if incluir_produto:
                valores = (nome_produto,) + valores  # Adiciona o nome do produto
            # Caso contrário, mantém apenas os valores das colunas sem o nome do produto

            for col_idx, valor in enumerate(valores):
                aba.write(row_idx, col_idx, valor)

    def exportar_para_pdf(self):
        aba_atual = self.abas.tab(self.abas.select(), "text")
        # Seleciona a Treeview e configurações conforme aba
        if aba_atual == "Relatório de Saída":
            tree = getattr(self, 'tabela_saida', None)
            nome_padrao = "Relatorio_saida.pdf"
            titulo_pdf = "Relatório de Saída"
        elif aba_atual == "Relatório de Entrada":
            tree = self.relatorioEntrada.tabela_entrada
            nome_padrao = "Relatorio_entrada.pdf"
            titulo_pdf = "Relatório de Entrada"
        elif aba_atual == "Relatório Resumo":
            tree = self.relatorioResumo.tree
            nome_padrao = "Relatorio_resumo.pdf"
            titulo_pdf = "Relatório Resumo"
        else:
            messagebox.showerror("Erro", "Esta aba não possui dados para exportar.")
            return

        # Verifica se há dados
        if not tree or not tree.get_children():
            messagebox.showerror("Erro", "Não há dados para exportar nesta aba.")
            return

        # Diálogo de salvamento com nome padrão
        caminho = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF","*.pdf")],
            initialfile=nome_padrao,
            title="Salvar como"
        )
        if not caminho:
            return

        try:
            estilos = getSampleStyleSheet()
            doc = SimpleDocTemplate(caminho, pagesize=A4)
            elementos = []
            # Título
            elementos.append(Paragraph(f"<b>{titulo_pdf}</b>", estilos["Title"]))
            elementos.append(Spacer(1, 20))

            # Monta dados para o PDF
            if aba_atual == "Relatório Resumo":
                # Inclui a coluna de árvore (#0)
                cabecalhos = [tree.heading("#0")["text"]] + [tree.heading(col)["text"] for col in tree["columns"]]
                dados = [cabecalhos]
                for iid in tree.get_children():
                    texto = tree.item(iid)["text"]
                    valores = tree.item(iid)["values"]
                    dados.append([texto] + [str(v) for v in valores])
            else:
                # Abas Saída e Entrada
                cabecalhos = [tree.heading(col)["text"] for col in tree["columns"]]
                dados = [cabecalhos]
                for iid in tree.get_children():
                    dados.append([str(v) for v in tree.item(iid)["values"]])

            tabela_pdf = Table(dados)
            tabela_pdf.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.grey),
                ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("BOTTOMPADDING",(0,0),(-1,0),12),
                ("GRID",(0,0),(-1,-1),0.5,colors.black),
            ]))
            elementos.append(tabela_pdf)
            doc.build(elementos)
            messagebox.showinfo("Sucesso", f"Relatório exportado para: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para PDF: {e}")

    def on_closing(self):
        """Fecha a janela e encerra o programa corretamente."""
        if messagebox.askyesno("Fechar", "Tem certeza que deseja fechar esta janela?"):
            print("Fechando o programa corretamente...")

            # Fecha a conexão com o banco de dados, se estiver aberta
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
                print("Conexão com o banco de dados fechada.")

            self.destroy()  # Fecha a janela
            