import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont, filedialog
from conexao_db import conectar
from logos import aplicar_icone
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import re
from centralizacao_tela import centralizar_janela
import pandas as pd

class RegistroTeste(tk.Toplevel):
    def __init__(self, janela_menu=None, master=None):
        super().__init__(master=master)
        self.janela_menu = janela_menu
        self.title("Registro de Teste")
        self.config(bg="#f4f4f4")

        # Fonte e Ícone
        fixed_font = tkfont.Font(family="Arial", size=10)
        self.option_add("*Font", fixed_font)
        aplicar_icone(self, r"C:\Sistema\logos\Kametal.ico")

        # Tamanho
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{int(sw*0.8)}x{int(sh*0.8)}")
        self.state("zoomed")

        # Conexão DB
        try:
            self.conn   = conectar()
            self.cursor = self.conn.cursor()
        except Exception as e:
            messagebox.showerror("Erro de Conexão", str(e))
            self.destroy()
            return

        # Montagem da UI
        self.create_widgets()
        self.atualizar_treeview()
        self.configure_treeview()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        main = tk.Frame(self, bg="#f4f4f4")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        # 1) Barra de pesquisa
        search_frame = tk.Frame(main, bg="#f4f4f4")
        search_frame.pack(fill="x", pady=(0,10), padx=10)
        tk.Label(search_frame, text="Pesquisar:", bg="#f4f4f4").pack(side="left")

        # entry sem StringVar, vamos usá‑lo diretamente
        self.search_entry = tk.Entry(search_frame, width=25)
        self.search_entry.pack(side="left", fill="x", expand=False, padx=(5,0))
        # a cada tecla, passa o texto para o filtro
        self.search_entry.bind("<KeyRelease>", lambda ev: self._filter_rows(ev.widget.get()))

        # Botão Exportar Excel acima da Treeview, alinhado à direita
        export_btn = ttk.Button(search_frame, text="Exportar Excel", command=self.abrir_dialogo_exportacao, width=15, style="Mini.TButton")
        export_btn.pack(side="right")

        form = tk.LabelFrame(main, text="Dados do Registro", bg="#f4f4f4", padx=10, pady=10)
        form.pack(fill="x", padx=10, pady=10)

        labels = [
            "Data", 
            "Código de Barras", 
            "O.P.", 
            "Cliente", 
            "Material",
            "Liga", 
            "Dimensões", 
            "Área", 
            "L.R. Tração (N)", 
            "L.R. Tração (MPa)", 
            "Alongamento (%)", 
            "Tempera", 
            "Máquina", 
            "Empresa"
        ]

        self.entries = {}
        for i, lbl in enumerate(labels):
            row, col = divmod(i, 3)
            tk.Label(form, text=lbl + ":", bg="#f4f4f4")\
              .grid(row=row, column=col*2, sticky="e", padx=5, pady=5)

            # 1) Cria o Entry em 'e' SEMPRE
            e = tk.Entry(form, width=15)

            # 2) Se for Data, guarda e bind
            if lbl == "Data":
                self.date_entry = e
                e.bind("<KeyRelease>", self._on_date_key)

            # 3) Se for Alongamento (%), guarda e bind
            elif lbl == "Alongamento (%)":
                self.along_entry = e
                e.bind("<KeyRelease>", self._on_along_key)

             # Dimensões (decimal automático)
            elif lbl == "Dimensões":
                self.dim_entry = e
                e.bind("<KeyRelease>", lambda ev, en=e: self._on_decimal_key(ev, en))

            # Área (decimal automático)
            elif lbl == "Área":
                self.area_entry = e
                e.bind("<KeyRelease>", lambda ev, en=e: self._on_decimal_key(ev, en))

            elif lbl == "Tempera":
                self.temper_entry = e
                e.bind("<KeyRelease>", self._on_tempera_key)

            # 4) Grida e armazena em entries
            e.grid(row=row, column=col*2+1, sticky="w", padx=5, pady=5)
            self.entries[lbl] = e

        # Botões
        style = ttk.Style(self)
        style.configure("Mini.TButton", font=("Arial",10), padding=(2,1))
        btn_frame = tk.Frame(main, bg="#f4f4f4")
        btn_frame.pack(pady=5)
        for text, cmd in [
            ("Salvar", self.salvar), ("Alterar", self.alterar),
            ("Excluir", self.excluir), ("Limpar", self.limpar),
            ("Voltar", self.voltar_para_menu)
        ]:
            ttk.Button(btn_frame, text=text, command=cmd, width=15, style="Mini.TButton")\
               .pack(side="left", padx=3)

        # Treeview
        tree_frame = tk.Frame(main)
        tree_frame.pack(fill="both", expand=True, pady=(10,0))
        vsb = tk.Scrollbar(tree_frame, orient="vertical")
        hsb = tk.Scrollbar(tree_frame, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        cols = labels[:]
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.tree.pack(side="left", fill="both", expand=True)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, anchor="center")
        self.tree.column("Data", width=80)
        self.tree.column("Área", width=100)
        self.tree.column("Alongamento (%)", width=100)

        self.tree.bind("<<TreeviewSelect>>", self.carregar_dados)

    def atualizar_treeview(self):
        """Carrega todos os registros sem filtro e exibe tudo."""
        try:
            self.cursor.execute(
                """
                SELECT *
                FROM registro_teste
                ORDER BY
                data DESC,
                LEFT(
                    regexp_replace(unaccent(lower(cliente)), '[^a-z]', '', 'g'),
                    1
                ) ASC,
                codigo_barras ASC;
                """
            )
            self.all_rows = self.cursor.fetchall()
            self._filter_rows("")  # exibe tudo novamente
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar: {e}")

    def abrir_dialogo_exportacao(self):
        dialogo = tk.Toplevel(self)
        dialogo.title("Exportar Registro - Filtros")
        dialogo.geometry("400x400")  # altura aumentada
        dialogo.resizable(False, False)
        centralizar_janela(dialogo, 400, 400)
        aplicar_icone(dialogo, r"C:\Sistema\logos\Kametal.ico")
        dialogo.config(bg="#ecf0f1")

        # Configuração dos estilos personalizados utilizando ttk
        style = ttk.Style(dialogo)
        style.theme_use("alt")
        style.configure("Custom.TFrame", background="#ecf0f1")
        style.configure("Custom.TLabel", background="#ecf0f1", foreground="#34495e", font=("Arial", 10))
        style.configure("Custom.TButton", background="#2980b9", foreground="white", font=("Arial", 10, "bold"), padding=5)
        style.map("Custom.TButton",
                background=[("active", "#3498db")],
                foreground=[("active", "white")])

        frame = ttk.Frame(dialogo, padding="15", style="Custom.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        dialogo.columnconfigure(0, weight=1)
        dialogo.rowconfigure(0, weight=1)

        # Função de formatação de data
        def formatar_data(event):
            entry = event.widget
            data = ''.join(filter(str.isdigit, entry.get()))[:8]
            if len(data)>=2: data=data[:2]+'/'+data[2:]
            if len(data)>=5: data=data[:5]+'/'+data[5:]
            entry.delete(0, tk.END); entry.insert(0,data)

        # Campos de filtro correspondentes a registro_teste
        filtros = [
            "Data Inicial", 
            "Data Final", 
            "Código de Barras", 
            "O.P.",
            "Cliente", 
            "Material", 
            "Máquina", 
            "Empresa"
        ]
        entries = {}
        binds = {"Data Inicial": formatar_data, "Data Final": formatar_data}
        for i, lbl in enumerate(filtros):
            ttk.Label(frame, text=f"{lbl}:", style="Custom.TLabel").grid(
                row=i, column=0, sticky="e", padx=5, pady=4)
            e = ttk.Entry(frame, width=20)
            e.grid(row=i, column=1, sticky="w", padx=5, pady=4)
            if lbl in binds:
                e.bind("<KeyRelease>", binds[lbl])
            entries[lbl] = e

        def acao_exportar():
            vals = {lbl: entries[lbl].get().strip() for lbl in filtros}
            where, params = [], []
            # Data
            if vals["Data Inicial"]:
                try:
                    di = datetime.strptime(vals["Data Inicial"], '%d/%m/%Y')
                    where.append("data >= %s"); params.append(di.strftime('%Y-%m-%d'))
                except: pass
            if vals["Data Final"]:
                try:
                    dfm = datetime.strptime(vals["Data Final"], '%d/%m/%Y')
                    where.append("data <= %s"); params.append(dfm.strftime('%Y-%m-%d'))
                except: pass
            # Outros filtros textuais
            for col, field in [
                ("codigo_barras","Código de Barras"),
                ("op","O.P."),
                ("cliente","Cliente"),
                ("material","Material"),
                ("maquina","Máquina"),
                ("empresa","Empresa")
            ]:
                if vals[field]:
                    where.append(f"{col} ILIKE %s"); params.append(f"%{vals[field]}%")
            sql = (
                "SELECT data, codigo_barras, op, cliente, material, liga, "
                "dimensoes, area, lr_tracao_n, lr_tracao_mpa, alongamento_percentual, tempera, maquina, empresa FROM registro_teste"
            )
            if where:
                sql += " WHERE " + " AND ".join(where)
            sql += " ORDER BY data ASC"

            arquivo = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel','*.xlsx')],
                initialfile='RegistroTeste.xlsx'
            )
            if arquivo:
                try:
                    self.cursor.execute(sql, params)
                    dados = self.cursor.fetchall()
                    cols = [
                        "Data", "Código de Barras", "O.P.", "Cliente", "Material",
                        "Liga", "Dimensões", "Área", "L.R. Tração (N)",
                        "L.R. Tração (MPa)", "Alongamento (%)", "Tempera", "Máquina", "Empresa"
                    ]
                    df = pd.DataFrame(dados, columns=cols)
                    df['Data'] = df['Data'].apply(lambda d: d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d))
                    df.to_excel(arquivo, index=False)
                    messagebox.showinfo("Exportação", f"Arquivo salvo em {arquivo}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Não foi possível exportar: {e}")
            dialogo.destroy()

        # Botões
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.grid(row=len(filtros), column=0, columnspan=2, pady=10)
        ttk.Button(
            btn_frame, text="Exportar Excel",
            command=acao_exportar, style="Custom.TButton"
        ).grid(row=0, column=0, padx=5)
        ttk.Button(
            btn_frame, text="Cancelar",
            command=dialogo.destroy, style="Custom.TButton"
        ).grid(row=0, column=1, padx=5)
        frame.columnconfigure(1, weight=1)

    def _filter_rows(self, term):
        """
        term: string com o que foi digitado no Entry.
        Se term == \"\", exibe tudo.
        """
        term = term.lower().strip()
        print(f"[DEBUG] _filter_rows() chamado com termo: '{term}'")
        self.tree.delete(*self.tree.get_children())

        for row in getattr(self, 'all_rows', []):
            rec_id, raw = row[0], row[1]
            data = raw.strftime("%d/%m/%Y") if hasattr(raw, "strftime") else str(raw or "")
            vals = [
                data,
                row[2], row[3], row[4], row[5],
                row[6], row[7],
                f"{row[8]:.5f}".replace(".", ",") if row[8] is not None else "",
                row[9] or "", row[10] or "",
                f"{row[11]:.2f}".replace(".", ",") + "%" if row[11] is not None else "",
                row[12], row[13], row[14]
            ]
            if not term or any(term in str(v).lower() for v in vals):
                self.tree.insert("", "end", iid=str(rec_id), values=vals)

        shown = len(self.tree.get_children())
        print(f"[DEBUG] linhas exibidas após filtro: {shown}")

    def configure_treeview(self):
        style = ttk.Style()
        style.theme_use("alt")
        style.configure("Treeview", rowheight=20)
        self.tree.config(height=10)

    def carregar_dados(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]

        for lbl, v in zip(self.entries.keys(), vals):
            e = self.entries[lbl]
            e.delete(0, tk.END)
            if lbl == "Data":
                # já vem formatado "DD/MM/YYYY"
                e.insert(0, v)
            elif lbl == "Alongamento (%)":
                # já vem formatado "60,00%"
                e.insert(0, v)
            else:
                e.insert(0, v)

    def _on_date_key(self, event, entry=None):
        e = entry or self.date_entry
        digits = re.sub(r'\D', '', e.get())[:8]

        parts = []
        if len(digits) >= 2:
            parts.append(digits[:2])
            if len(digits) >= 4:
                parts.append(digits[2:4])
                if len(digits) > 4:
                    parts.append(digits[4:])
            else:
                parts.append(digits[2:])
        else:
            parts.append(digits)

        novo = '/'.join(parts)
        e.delete(0, tk.END)
        e.insert(0, novo)

    def _on_along_key(self, event):
        e = self.along_entry
        t = e.get()
        core = ''.join(c for c in t if c.isdigit() or c in ",.") 
        novo = (core + "%") if core else ""
        pos = e.index(tk.INSERT)
        e.delete(0, tk.END)
        e.insert(0, novo)
        e.icursor(min(pos, len(novo)))

    def _on_decimal_key(self, event, entry):
        """
        Formata automaticamente para que antes das cinco últimas 
        casas haja uma vírgula. Ex.: "123456789" -> "1234,56789".
        """
        t = entry.get()
        # pega só dígitos
        digits = re.sub(r'\D', '', t)

        if not digits:
            novo = ""
        elif len(digits) <= 5:
            # até cinco dígitos, mostra-os puros (sem vírgula)
            novo = digits
        else:
            # insere vírgula antes dos 5 últimos dígitos
            inteiro = digits[:-5]
            dec = digits[-5:]
            novo = f"{inteiro},{dec}"

        # posiciona o cursor próximo da posição original
        pos = entry.index(tk.INSERT)
        entry.delete(0, tk.END)
        entry.insert(0, novo)
        entry.icursor(min(pos, len(novo)))

    def _on_tempera_key(self, event):
        """
        Garante que o campo termine com ' Duro'.
        Ex.: 'Extra' -> 'Extra Duro'
        """
        e = self.temper_entry
        text = e.get().strip()
        # remove sufixo já presente
        core = re.sub(r'\s*[dD]uro$', '', text)
        novo = f"{core} Duro" if core else ""
        pos = e.index(tk.INSERT)
        e.delete(0, tk.END)
        e.insert(0, novo)
        # ajusta cursor antes de ' Duro'
        new_pos = min(pos, len(core))
        e.icursor(new_pos)

    def normalize_number(self, s):
        """
        Recebe uma string como “14.570,0000” ou “14570.0000” e retorna
        algo como “14570.0000” pronto p/ Decimal().
        """
        s = s.strip()
        if not s:
            return None

        # se tem ambos, . = milhar, , = decimal
        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        # se só tem vírgula, ela é decimal
        elif "," in s:
            s = s.replace(",", ".")
        # se só tem ponto, assume-se que é decimal e deixamos
        return s

    def salvar(self):
        # helper para pegar valor de cada Entry
        get = lambda lbl: self.entries[lbl].get().strip()

        # pega a string da data, mantendo as barras
        data_str    = get("Data")                # ex.: "23/07/2025"
        cod         = get("Código de Barras")
        op          = get("O.P.")
        cliente     = get("Cliente")
        material    = get("Material")
        liga        = get("Liga")
        dims        = get("Dimensões")
        area_s      = get("Área").replace(".", "").replace(",", ".")
        lr_n_s      = get("L.R. Tração (N)")
        lr_mpa_s    = get("L.R. Tração (MPa)")
        along_str   = get("Alongamento (%)").replace("%", "")
        tempera     = get("Tempera")
        maquina     = get("Máquina")
        empresa     = get("Empresa")

        # Valida data
        if not data_str:
            return messagebox.showerror("Erro", "O campo Data é obrigatório.")
        try:
            data_date = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return messagebox.showerror("Erro", "Data inválida! Use DD/MM/AAAA.")

        # Converte apenas área e alongamento (numéricos)
        def to_float(s, nome):
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                messagebox.showerror("Erro", f"{nome} deve ser número.")
                raise

        try:
            area  = to_float(area_s,  "Área")
            along = to_float(along_str, "Alongamento (%)")
        except:
            return
        
         # pega menor ID disponível
        try:
            self.cursor.execute(
                """
                SELECT COALESCE(
                    (SELECT MIN(t1.id) + 1
                     FROM registro_teste t1
                     LEFT JOIN registro_teste t2 ON t2.id = t1.id + 1
                     WHERE t2.id IS NULL), 1
                );
                """
            )
            next_id = self.cursor.fetchone()[0]
        except Exception as e:
            return messagebox.showerror("Erro ao gerar ID", str(e))

        # Monta parâmetros **mantendo lr_n_s e lr_mpa_s como texto**
        params = (
            next_id,
            cod, data_date, op, cliente, material, liga, dims,
            area, lr_n_s, lr_mpa_s, along, tempera, maquina, empresa
        )
        sql = """
            INSERT INTO registro_teste
            (id, codigo_barras, data, op, cliente, material, liga, dimensoes,
            area, lr_tracao_n, lr_tracao_mpa, alongamento_percentual,
            tempera, maquina, empresa)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
        """

        try:
            self.cursor.execute(sql, params)
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Registro salvo com sucesso!")
            self.limpar() 
            self.atualizar_treeview()
        except Exception as e:
            messagebox.showerror("Falha ao salvar", str(e))

    def alterar(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("Atenção", "Selecione um registro para alterar.")
        rec_id = int(sel[0])

        get = lambda lbl: self.entries[lbl].get().strip()
        data_str = get("Data")
        cod       = get("Código de Barras")
        op        = get("O.P.")
        cliente   = get("Cliente")
        material  = get("Material")
        liga      = get("Liga")
        dims      = get("Dimensões")

        # área e alongamento como string “1000,12345”
        area_s = self.normalize_number(get("Área"))

        # retira o '%' antes de normalizar
        raw_along = get("Alongamento (%)").replace("%", "")
        along_s   = self.normalize_number(raw_along)

        # ranges de tração como texto puro
        lr_n_s   = get("L.R. Tração (N)")
        lr_mpa_s = get("L.R. Tração (MPa)")

        tempera = get("Tempera")
        maquina = get("Máquina")
        empresa = get("Empresa")

        # valida data
        if not data_str:
            return messagebox.showerror("Erro", "O campo Data é obrigatório.")
        try:
            data_date = datetime.strptime(data_str, "%d/%m/%Y").date()
        except ValueError:
            return messagebox.showerror("Erro", "Data inválida! Use DD/MM/AAAA.")

        # converte e quantiza area e alongamento
        try:
            area_dec = Decimal(area_s).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
            if abs(area_dec) >= Decimal("1E8"):
                raise InvalidOperation()
        except InvalidOperation:
            return messagebox.showerror(
                "Erro",
                "Área inválida: máximo ±10⁸ e até 4 casas decimais."
            )

        try:
            along_dec = Decimal(along_s).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP)
            if abs(along_dec) >= Decimal("1E4"):
                raise InvalidOperation()
        except InvalidOperation:
            return messagebox.showerror(
                "Erro",
                "Alongamento inválido: máximo ±10⁴ e até 2 casas decimais."
            )

        # agora monta o UPDATE com o nome correto da coluna de data
        sql = """
            UPDATE registro_teste
            SET codigo_barras          = %s,
                data                    = %s,
                op                     = %s,
                cliente                = %s,
                material               = %s,
                liga                   = %s,
                dimensoes              = %s,
                area                   = %s,
                lr_tracao_n            = %s,
                lr_tracao_mpa          = %s,
                alongamento_percentual = %s,
                tempera                = %s,
                maquina                = %s,
                empresa                = %s
            WHERE id = %s;
        """
        params = (
            cod, data_date, op, cliente, material, liga, dims,
            area_dec, lr_n_s, lr_mpa_s, along_dec,
            tempera, maquina, empresa,
            rec_id
        )

        try:
            self.cursor.execute(sql, params)
            self.conn.commit()
            messagebox.showinfo("Sucesso", "Registro alterado com sucesso!")
            self.atualizar_treeview()
            self.limpar()
        except Exception as e:
            messagebox.showerror("Erro ao alterar", str(e))

    def excluir(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("Atenção", "Selecione ao menos um registro para excluir.")

        ids = [int(item) for item in sel]
        count = len(ids)
        pergunta = f"Deseja realmente excluir {count} registro(s)?"
        if not messagebox.askyesno("Confirmação", pergunta):
            return

        try:
            # exclui múltiplos de uma vez
            placeholders = ",".join(["%s"] * count)
            sql = f"DELETE FROM registro_teste WHERE id IN ({placeholders});"
            self.cursor.execute(sql, tuple(ids))
            self.conn.commit()

            self.atualizar_treeview()
            self.limpar()
            messagebox.showinfo("Sucesso", f"{count} registro(s) excluído(s) com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao excluir", str(e))

    def limpar(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def voltar_para_menu(self):
        self.destroy()
        if self.janela_menu:
            try:
                self.janela_menu.deiconify()
                self.janela_menu.state("zoomed")
                self.janela_menu.lift()
                self.janela_menu.update()
            except Exception:
                pass

    def on_closing(self):
        self.conn.close()
        self.destroy()
