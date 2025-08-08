import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from logos import aplicar_icone
from centralizacao_tela import centralizar_janela

# ----- Constantes de Estilo -----
PRIMARY_BG = '#F2F3F5'
HEADER_BG = '#2C3E50'
ACCENT_ACTIVE = '#1F618D'
ERROR_COLOR = '#C0392B'
FONT_MAIN = ('Segoe UI', 11)
FONT_HEADER = ('Segoe UI', 18, 'bold')
FONT_RESULT = ('Segoe UI', 14, 'bold')
FONT_HINT = ('Segoe UI', 9, 'italic')


class DicaDeFerramenta(ttk.Frame):
    """Exibe uma dica ao passar o mouse sobre o widget alvo."""

    def __init__(self, widget, texto: str):
        super().__init__(widget.master)
        self.widget = widget
        self.texto = texto
        self.janela_dica = None
        widget.bind('<Enter>', self._mostrar)
        widget.bind('<Leave>', self._esconder)

    def _mostrar(self, evento=None):
        if self.janela_dica or not self.texto:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.janela_dica = janela = tk.Toplevel(self.widget)
        janela.wm_overrideredirect(True)
        janela.wm_geometry(f'+{x}+{y}')
        rotulo = ttk.Label(
            janela,
            text=self.texto,
            background='#FFFFE0',
            relief='solid',
            borderwidth=1,
            font=FONT_HINT,
            padding=(4, 2)
        )
        rotulo.pack()

    def _esconder(self, evento=None):
        if self.janela_dica:
            self.janela_dica.destroy()
            self.janela_dica = None

class EntradaComPlaceholder(ttk.Entry):
    """Entrada com texto de espaço reservado (placeholder)."""

    def __init__(self, pai, texto_espaco_reservado: str, **kwargs):
        super().__init__(pai, **kwargs)
        self.texto_espaco_reservado = texto_espaco_reservado
        self.cor_padrao = '#777'
        self.bind('<FocusIn>', self._limpar)
        self.bind('<FocusOut>', self._adicionar)
        self._adicionar()

    def _limpar(self, evento=None):
        if self.get() == self.texto_espaco_reservado:
            self.delete(0, tk.END)
            self.config(foreground='black')

    def _adicionar(self, evento=None):
        if not self.get():
            self.insert(0, self.texto_espaco_reservado)
            self.config(foreground=self.cor_padrao)

class CalculadoraMediaPonderada(tk.Toplevel):
    """Janela principal da Calculadora de Média Ponderada."""

    def __init__(self, pai: tk.Tk = None):
        super().__init__(pai)
        self.title('Calculadora de Média Ponderada')
        self.configure(bg=PRIMARY_BG)
        largura, altura = 450, 650
        centralizar_janela(self, largura, altura)
        self.minsize(450, 650)
        aplicar_icone(self, Path('C:/Sistema/logos/Kametal.ico'))

        self._configurar_estilo()
        self._criar_cabecalho()
        self._criar_corpo()
        self._criar_barra_status()

        self.entradas: list[tuple[EntradaComPlaceholder, EntradaComPlaceholder]] = []
        self._adicionar_linha()
        self._vincular_atalhos()

    def _configurar_estilo(self):
        estilo = ttk.Style(self)
        estilo.theme_use('alt')
        estilo.configure('TFrame', background=PRIMARY_BG)
        estilo.configure('Header.TLabel', background=HEADER_BG, foreground='white', font=FONT_HEADER)
        estilo.configure('Hint.TLabel', background=PRIMARY_BG, foreground='#555', font=FONT_HINT)
        # Novo estilo para labels com fundo igual ao da janela
        estilo.configure('Padrao.TLabel', background=PRIMARY_BG, font=FONT_MAIN)

    def _criar_cabecalho(self):
        quadro_cabecalho = ttk.Frame(self)
        quadro_cabecalho.pack(fill='x', pady=(10, 0))

        canvas = tk.Canvas(quadro_cabecalho, height=50, bg=HEADER_BG, highlightthickness=0)
        canvas.pack(fill='x')

        etiqueta = ttk.Label(quadro_cabecalho, text='Calculadora Média Ponderada',style='Header.TLabel', background=HEADER_BG)
        etiqueta.place(relx=0.5, rely=0.5, anchor='center')

        self._pos_barra = 0
        self._dir_barra = 1  # 1 = direita, -1 = esquerda

    def _criar_corpo(self):
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=20, pady=10)

        # Canvas com Scrollbar
        canvas = tk.Canvas(container, bg=PRIMARY_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Frame dentro do canvas
        self.quadro_entradas = ttk.Frame(canvas)
        self.quadro_entradas.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        self._frame_id = canvas.create_window((0, 0), window=self.quadro_entradas, anchor='nw')

        # Atualiza o tamanho do frame ao redimensionar
        def ajustar_largura(event):
            canvas.itemconfig(self._frame_id, width=event.width)

        canvas.bind('<Configure>', ajustar_largura)

        # Cabeçalho das colunas
        ttk.Label(self.quadro_entradas, text='Peso', style='Padrao.TLabel').grid(row=0, column=0, padx=10, pady=4)
        ttk.Label(self.quadro_entradas, text='Valor', style='Padrao.TLabel').grid(row=0, column=1, padx=10, pady=4)
        self.quadro_entradas.columnconfigure((0, 1), weight=1)

        # Botões e resultado (fora da área de rolagem)
        controle = ttk.Frame(self)
        controle.pack(pady=5)

        acoes = [
            ('Adicionar: Ctrl+N', self._adicionar_linha, 'Adiciona linha: Ctrl+N'),
            ('Remover: Ctrl+D', self._remover_linha, 'Remove linha: Ctrl+D'),
            ('Calcular: Ctrl+C', self._calcular, 'Calcula média: Ctrl+C'),
            ('Limpar: Ctrl+L', self._limpar_tudo, 'Limpa: Ctrl+L')
        ]
        for idx, (texto, comando, dica) in enumerate(acoes):
            botao = ttk.Button(controle, text=texto, command=comando, style='Accent.TButton')
            botao.grid(row=0, column=idx, padx=5)
            DicaDeFerramenta(botao, dica)

        quadro_resultado = ttk.Frame(self)
        quadro_resultado.pack(fill='x', padx=20, pady=(5, 0))
        ttk.Label(quadro_resultado, text='Resultado:', style='Padrao.TLabel').pack(side='left', padx=10, pady=10)
        self.entrada_resultado = ttk.Entry(quadro_resultado, font=FONT_RESULT, justify='center', state='readonly')
        self.entrada_resultado.pack(side='left', fill='x', expand=True, padx=(10, 0), pady=10)
    
    def _criar_barra_status(self):
        self.variavel_status = tk.StringVar(value='Pronto')
        status = ttk.Label(self, textvariable=self.variavel_status, anchor='w',font=('Segoe UI', 10), relief='sunken')
        status.pack(fill='x', side='bottom')

    def _vincular_atalhos(self):
        self.bind_all('<Control-n>', lambda e: self._adicionar_linha())
        self.bind_all('<Control-d>', lambda e: self._remover_linha())
        self.bind_all('<Control-c>', lambda e: self._calcular())
        self.bind_all('<Control-l>', lambda e: self._limpar_tudo())
        self.bind_all('<Control-q>', lambda e: self.destroy())

    def _adicionar_linha(self):
        linha = len(self.entradas) + 1
        entrada_peso = EntradaComPlaceholder(self.quadro_entradas, 'Digite o peso', font=FONT_MAIN)
        entrada_valor = EntradaComPlaceholder(self.quadro_entradas, 'Digite o valor', font=FONT_MAIN)

        entrada_peso.grid(row=linha, column=0, padx=10, pady=4, sticky='ew')
        entrada_valor.grid(row=linha, column=1, padx=10, pady=4, sticky='ew')

        entrada_peso.bind('<KeyRelease>', lambda e: self._formatar_entrada(entrada_peso, 3))
        entrada_valor.bind('<KeyRelease>', lambda e: self._formatar_entrada(entrada_valor, 2))

        self.entradas.append((entrada_valor, entrada_peso))  # mantém ordem lógica
        self.variavel_status.set(f'Linha {linha} adicionada')

    def _remover_linha(self):
        if self.entradas:
            entrada_valor, entrada_peso = self.entradas.pop()
            entrada_valor.destroy()
            entrada_peso.destroy()
            self.variavel_status.set('Linha removida')

    def _formatar_entrada(self, entrada: EntradaComPlaceholder, casas_decimais: int):
        texto = entrada.get().replace(',', '').replace('.', '')
        if texto.isdigit():
            while len(texto) < casas_decimais + 1:
                texto = '0' + texto
            inteiro = texto[:-casas_decimais]
            decimal = texto[-casas_decimais:]
            novo_valor = f'{int(inteiro)},{decimal}'
            entrada.delete(0, tk.END)
            entrada.insert(0, novo_valor)
        elif not texto:
            pass
        else:
            entrada.delete(0, tk.END)

    def _calcular(self):
        try:
            valores, pesos = [], []
            for entrada_valor, entrada_peso in self.entradas:
                val = entrada_valor.get().replace(',', '.')
                pes = entrada_peso.get().replace(',', '.')
                if val == entrada_valor.texto_espaco_reservado or pes == entrada_peso.texto_espaco_reservado:
                    raise ValueError('Entradas incompletas')
                valores.append(float(val))
                pesos.append(float(pes))

            total_peso = sum(pesos)
            if total_peso == 0:
                raise ZeroDivisionError('Peso total zero')

            media = sum(v * p for v, p in zip(valores, pesos)) / total_peso
            self.entrada_resultado.config(state='normal')
            self.entrada_resultado.delete(0, tk.END)
            self.entrada_resultado.insert(0, f'R$ {media:.2f}'.replace('.', ','))
            self.entrada_resultado.config(state='readonly')
            self.variavel_status.set('Cálculo concluído')
        except ZeroDivisionError:
            messagebox.showerror('Erro', 'Peso total não pode ser zero.', parent=self)
            self.variavel_status.set('Erro: peso zero')
        except Exception:
            messagebox.showerror('Erro', 'Verifique os valores informados.', parent=self)
            self.variavel_status.set('Erro de entrada')

    def _limpar_tudo(self):
        for entrada_valor, entrada_peso in self.entradas:
            entrada_valor.destroy()
            entrada_peso.destroy()
        self.entradas.clear()
        self.entrada_resultado.config(state='normal')
        self.entrada_resultado.delete(0, tk.END)
        self.entrada_resultado.config(state='readonly')
        self._adicionar_linha()
        self.variavel_status.set('Tudo limpo')

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    CalculadoraMediaPonderada(root).mainloop()
