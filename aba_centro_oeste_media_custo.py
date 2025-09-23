import tkinter as tk
from tkinter import ttk
from conexao_db import conectar
from decimal import Decimal, ROUND_HALF_UP

class AbaCentroOesteMediaCusto:

    def __init__(self, aba_centro_oeste, conn):
        self.aba = aba_centro_oeste
        self.conn = conn

        # Criação da frame onde a tabela vai ficar
        self.frame_tabela = tk.Frame(self.aba)
        self.frame_tabela.pack(expand=True, fill='both', padx=10, pady=10)

        # Criação da treeview para a tabela
        self.tree = ttk.Treeview(self.frame_tabela, columns=(
            'Produto', 'Estoque', 'Custo Estoque', 'Custo empresa',
            'Custo 5%', 'Custo 10%', 'Custo 15%', 'Custo 20%'
        ), show='headings', style='Custom.Treeview')

        self.tree.heading('Produto', text='Nome do Produto')
        self.tree.heading('Estoque', text='Estoque')
        self.tree.heading('Custo Estoque', text='Média de Custo do Estoque')
        self.tree.heading('Custo empresa', text='Média Custo Empresa')
        self.tree.heading('Custo 5%', text='Custo com 5%')
        self.tree.heading('Custo 10%', text='Custo com 10%')
        self.tree.heading('Custo 15%', text='Custo com 15%')
        self.tree.heading('Custo 20%', text='Custo com 20%')

        for col in ('Produto', 'Estoque', 'Custo Estoque', 'Custo empresa', 'Custo 5%', 'Custo 10%', 'Custo 15%', 'Custo 20%'):
            self.tree.column(col, anchor='center')

        # Estilo da tabela
        style = ttk.Style()
        try:
            style.theme_use('alt')
        except Exception:
            pass
        style.configure('Treeview', background='white', foreground='black', rowheight=25, fieldbackground='white')
        style.configure('Treeview.Heading', font=('Arial', 10, 'bold'))
        style.configure('Treeview', rowheight=30)
        style.configure('Custom.Treeview', font=('Arial', 10), rowheight=30)
        style.map('Treeview', background=[('selected', '#0078D7')], foreground=[('selected', 'white')])

        self.scrollbar_vertical = tk.Scrollbar(self.frame_tabela, orient='vertical', command=self.tree.yview)
        self.scrollbar_horizontal = tk.Scrollbar(self.frame_tabela, orient='horizontal', command=self.tree.xview)
        self.tree.config(yscrollcommand=self.scrollbar_vertical.set, xscrollcommand=self.scrollbar_horizontal.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        self.scrollbar_vertical.grid(row=0, column=1, sticky='ns')
        self.scrollbar_horizontal.grid(row=1, column=0, sticky='ew')

        self.frame_tabela.grid_rowconfigure(0, weight=1)
        self.frame_tabela.grid_columnconfigure(0, weight=1)

        # Popula os dados
        try:
            self._popular()
        except Exception as e:
            print('Erro ao popular aba Centro-Oeste/Nordeste:', e)

        # Ajustes visuais da aba
        try:
            self.aba.option_add('*TButton*background', '#D3D3D3')
            self.aba.option_add('*TButton*foreground', 'black')
            self.aba.option_add('*TButton*font', ('Arial', 10))
        except Exception:
            pass

    def _popular(self):
        # limpa
        for w in self.tree.get_children():
            self.tree.delete(w)

        produtos = self.buscar_produtos()
        for produto in produtos:
            nome_produto = produto[0]
            estoque = self.buscar_estoque(nome_produto)

            media_custo_estoque_str = self.calcular_media_ponderada(nome_produto)
            try:
                media_custo_estoque = float(media_custo_estoque_str.replace('R$ ', '').replace('.', '').replace(',', '.'))
            except Exception:
                media_custo_estoque = 0.0

            custo_empresa = self.calcular_custo_empresa(nome_produto)
            custo_5 = self.calcular_custo_5(custo_empresa, 5)
            custo_10 = self.calcular_custo_10(custo_empresa)
            custo_15 = self.calcular_custo_15(custo_empresa)
            custo_20 = self.calcular_custo_20(custo_empresa)

            tag_cor = ''
            if estoque.strip() == '0,000 kg' or not estoque.strip():
                tag_cor = 'estoque_vazio'

            self.tree.insert('', 'end', values=(nome_produto, estoque, media_custo_estoque_str, custo_empresa, custo_5, custo_10, custo_15, custo_20), tags=(tag_cor,))

        total_estoque = self.calcular_total_estoque()
        media_ponderada_total, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_5 = self.calcular_media_total_5()
        media_10 = self.calcular_media_total_10()
        media_15 = self.calcular_media_total_15()
        media_20 = self.calcular_media_total_20()

        self.tree.tag_configure('total', background='yellow', font=('Arial', 12, 'bold'), foreground='red')
        self.tree.insert('', 'end', values=('TOTAL', total_estoque, media_ponderada_total, media_ponderada_empresa, media_5, media_10, media_15, media_20), tags=('total',))
        self.tree.tag_configure('estoque_vazio', foreground='red')

    def conectar_banco(self):
        try:
            conn = conectar()
            return conn
        except Exception as e:
            print('Erro ao conectar ao banco:', e)
            return None

    def buscar_produtos(self):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None

        if not conn:
            return []

        try:
            cursor = conn.cursor()
            cursor.execute('SELECT nome FROM produtos ORDER BY nome;')
            produtos = cursor.fetchall()
            cursor.close()
            return produtos
        except Exception as e:
            print('Erro ao buscar nome_produtos:', e)
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
            return '0,000 kg'

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
                return f"{quantidade:,.3f}".replace(',', 'X').replace('.', ',').replace('X', '.') + ' kg'
            else:
                return '0,000 kg'
        except Exception as e:
            print(f'Erro ao buscar estoque para o produto {nome_produto}: {e}')
            return '0,000 kg'
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
            return 'R$ 0,00'

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
                    # fórmula específica desta aba (note 0.8375)
                    custo_calculado = (media_ponderada * 0.7275) / 0.8375
                    return 'R$ ' + f"{custo_calculado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                else:
                    return 'R$ 0,00'
            else:
                print(f'Erro: A consulta retornou None para {nome_produto}')
                return 'R$ 0,00'
        except Exception as e:
            print(f'Erro ao calcular média ponderada para {nome_produto}: {e}')
            return 'R$ 0,00'
        finally:
            if conexao_temporaria:
                try:
                    conn.close()
                except Exception:
                    pass

    def calcular_custo_empresa(self, nome_produto):
        conn = self.conn or self.conectar_banco()
        conexao_temporaria = self.conn is None
        cursor = None

        if not conn:
            return 'R$ 0,00'

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
                peso = produto[0]  # quantidade_estoque
                custo = produto[1]  # custo_total

                if peso is not None and custo is not None:
                    soma_custo_ponderado += float(peso) * float(custo)
                    soma_quantidade += float(peso)

            if soma_quantidade > 0:
                media_ponderada = soma_custo_ponderado / soma_quantidade
            else:
                media_ponderada = 0.0  # Caso não haja produtos válidos

            if media_ponderada == 0:
                return 'R$ 0,00'

            cursor.execute("""
                SELECT custo_empresa FROM somar_produtos WHERE produto = %s
            """, (nome_produto,))
            custo_empresa = cursor.fetchone()

            if custo_empresa and custo_empresa[0] is not None:
                custo_empresa_val = float(custo_empresa[0])
                media_ponderada = float(media_ponderada)

                resultado = media_ponderada + custo_empresa_val
                resultado_final = (resultado * 0.7275) / 0.8375

                return 'R$ ' + f"{resultado_final:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            else:
                return 'R$ 0,00'
        except Exception as e:
            print(f'Erro ao calcular custo_empresa para {nome_produto}: {e}')
            return 'R$ 0,00'
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
                custo_base = float(custo_base.replace('R$', '').replace('.', '').replace(',', '.').strip())
            custo_final = custo_base / (1 - percentual / 100.0)
            return 'R$ ' + f"{custo_final:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        except Exception as e:
            print(f'Erro ao calcular acréscimo de {percentual}%: {e}')
            return 'R$ 0,00'

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

            # Converte o estoque para float, tratando separadores e unidade ' kg'
            estoque_float = float(estoque.replace(' kg', '').replace('.', '').replace(',', '.'))

            total_estoque += estoque_float

        return f"{total_estoque:,.3f}".replace(',', 'X').replace('.', ',').replace('X', '.') + ' kg'

    def calcular_media_ponderada_total_com_empresa(self):
        produtos = self.buscar_produtos()

        soma_ponderada_custo = 0.0
        soma_ponderada_empresa = 0.0
        soma_estoque = 0.0

        for produto in produtos:
            nome_produto = produto[0]
            estoque = self.buscar_estoque(nome_produto)

            # Converte o estoque para float
            estoque_float = float(estoque.replace(' kg', '').replace('.', '').replace(',', '.'))

            if estoque_float > 0:
                media_custo_str = self.calcular_media_ponderada(nome_produto)
                media_custo = float(media_custo_str.replace('R$ ', '').replace('.', '').replace(',', '.'))

                if media_custo == 0:
                    custo_empresa = 0.0
                else:
                    custo_empresa_str = self.calcular_custo_empresa(nome_produto)
                    custo_empresa = float(custo_empresa_str.replace('R$ ', '').replace('.', '').replace(',', '.'))

                soma_ponderada_custo += estoque_float * media_custo
                soma_ponderada_empresa += estoque_float * custo_empresa
                soma_estoque += estoque_float

        if soma_estoque == 0:
            return 'R$ 0,000', 'R$ 0,000'

        media_ponderada_total = soma_ponderada_custo / soma_estoque
        media_ponderada_empresa = soma_ponderada_empresa / soma_estoque

        media_ponderada_total_str = f"R$ {media_ponderada_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        media_ponderada_empresa_str = f"R$ {media_ponderada_empresa:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

        return media_ponderada_total_str, media_ponderada_empresa_str

    def calcular_media_total_5(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace('R$ ', '').replace('.', '').replace(',', '.'))
        media_5 = media_ponderada_empresa_float / 0.95
        media_5_str = f"R$ {media_5:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        return media_5_str

    def calcular_media_total_10(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace('R$ ', '').replace('.', '').replace(',', '.'))
        media_10 = media_ponderada_empresa_float / 0.9
        media_10_str = f"R$ {media_10:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        return media_10_str

    def calcular_media_total_15(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace('R$ ', '').replace('.', '').replace(',', '.'))
        media_15 = media_ponderada_empresa_float / 0.85
        media_15_str = f"R$ {media_15:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        return media_15_str

    def calcular_media_total_20(self):
        _, media_ponderada_empresa = self.calcular_media_ponderada_total_com_empresa()
        media_ponderada_empresa_float = float(media_ponderada_empresa.replace('R$ ', '').replace('.', '').replace(',', '.'))
        media_20 = media_ponderada_empresa_float / 0.8
        media_20_str = f"R$ {media_20:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        return media_20_str

# Compatibilidade: função no módulo com a mesma assinatura do original
def criar_media_aba_centro_oeste(aba_centro_oeste, conn):
    return AbaCentroOesteMediaCusto(aba_centro_oeste, conn)
