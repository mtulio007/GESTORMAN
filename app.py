import os
import csv
import sqlite3
import unicodedata
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime


# Algumas instalações do Python para Windows não conseguem descobrir Tcl/Tk
# quando o perfil do Windows possui caracteres acentuados. O caminho é obtido
# a partir do módulo tkinter instalado, antes de criar a primeira janela.
_python_root = Path(tk.__file__).resolve().parent.parent.parent
_tcl_root = _python_root / "tcl"
if (_tcl_root / "tcl8.6" / "init.tcl").is_file():
    os.environ.setdefault("TCL_LIBRARY", str(_tcl_root / "tcl8.6"))
if (_tcl_root / "tk8.6" / "tk.tcl").is_file():
    os.environ.setdefault("TK_LIBRARY", str(_tcl_root / "tk8.6"))


ALMOX_COLUMNS = (
    "origem", "codigo", "descricao", "descricao_ingles", "entradas",
    "saida", "saldo", "data_inventario", "responsavel", "lead_time",
    "media_consumo", "estoque_minimo", "estoque_maximo", "est_seguranca",
    "ponto_pedido", "ressuprimento", "status", "custo_medio", "custo_ttl",
)
RECEBIMENTO_COLUMNS = (
    "data_recebimento", "codigo", "descricao", "und", "qtd",
    "fornecedor", "num_nota_fiscal", "data_protocolo",
)
SAIDA_COLUMNS = (
    "data", "codigo", "descricao", "un", "qtd",
    "turno", "aplicacao", "requisitante",
)
OS_COLUMNS = (
    "setor", "numero_equipamento", "solicitante", "hora_parada", "tipo_servico",
    "prioridade", "especialidade", "descricao", "tecnico", "turno", "hora_inicio",
    "hora_final", "tempo_parada", "tempo_servico", "tempo_resposta", "situacao",
)

TECNICO_TURNO_MAP = {
    "AMAURY SILVA": "B",
    "ANDRÉ RICARDO": "A",
    "CLEBER ROGER": "A",
    "ERIC MORAIS": "A",
    "GABRIEL": "A",
    "GABRIEL GOMES": "A",
    "JOÃO V. CARDOSO": "C",
    "JOVERLEY BATALHA": "B",
    "MARLISSON ALVES": "A",
    "OZAMIR": "C",
    "PAULO A. CORREA": "C",
    "RONÉLIO MARINHO": "B",
    "WILLIAN BRAZ": "C",
}


def migrate_almoxarifado(connection):
    """Cria ou atualiza a tabela de almoxarifado sem perder os registros existentes."""
    table_exists = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'almoxarifado'"
    ).fetchone()
    if not table_exists:
        connection.execute("""
            CREATE TABLE almoxarifado (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                origem TEXT NOT NULL,
                codigo TEXT NOT NULL,
                descricao TEXT NOT NULL,
                descricao_ingles TEXT,
                entradas TEXT,
                saida TEXT,
                saldo TEXT,
                data_inventario TEXT,
                responsavel TEXT,
                lead_time TEXT,
                media_consumo TEXT,
                estoque_minimo TEXT,
                estoque_maximo TEXT,
                est_seguranca TEXT,
                ponto_pedido TEXT,
                ressuprimento TEXT,
                status TEXT,
                custo_medio TEXT,
                custo_ttl TEXT,
                UNIQUE(origem, codigo)
            )
        """)
        return

    current_columns = {
        row[1] for row in connection.execute("PRAGMA table_info(almoxarifado)")
    }
    if set(ALMOX_COLUMNS).issubset(current_columns):
        return

    def value_for(*names):
        for name in names:
            if name in current_columns:
                return name
        return "''"

    legacy_values = (
        value_for("origem", "base"), value_for("codigo"), value_for("descricao"),
        value_for("descricao_ingles"), value_for("entradas"), value_for("saida"),
        value_for("saldo"), value_for("data_inventario"), value_for("responsavel"),
        value_for("lead_time", "dias_entrega"), value_for("media_consumo"),
        value_for("estoque_minimo"), value_for("estoque_maximo"), value_for("est_seguranca"),
        value_for("ponto_pedido"), value_for("ressuprimento"), value_for("status"),
        value_for("custo_medio", "custo"), value_for("custo_ttl"),
    )
    connection.execute("ALTER TABLE almoxarifado RENAME TO almoxarifado_anterior")
    connection.execute("""
        CREATE TABLE almoxarifado (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            origem TEXT NOT NULL,
            codigo TEXT NOT NULL,
            descricao TEXT NOT NULL,
            descricao_ingles TEXT,
            entradas TEXT,
            saida TEXT,
            saldo TEXT,
            data_inventario TEXT,
            responsavel TEXT,
            lead_time TEXT,
            media_consumo TEXT,
            estoque_minimo TEXT,
            estoque_maximo TEXT,
            est_seguranca TEXT,
            ponto_pedido TEXT,
            ressuprimento TEXT,
            status TEXT,
            custo_medio TEXT,
            custo_ttl TEXT,
            UNIQUE(origem, codigo)
        )
    """)
    connection.execute(
        f"INSERT INTO almoxarifado (id, {', '.join(ALMOX_COLUMNS)}) "
        f"SELECT id, {', '.join(legacy_values)} FROM almoxarifado_anterior"
    )
    connection.execute("DROP TABLE almoxarifado_anterior")


class GestorMan(tk.Tk):
    """Tela principal do GestorMan — cadastro de Ordens de Serviço."""

    COLUMNS = OS_COLUMNS
    ALMOX_COLUMNS = ALMOX_COLUMNS
    RECEBIMENTO_COLUMNS = RECEBIMENTO_COLUMNS
    SAIDA_COLUMNS = SAIDA_COLUMNS

    def __init__(self):
        super().__init__()
        self.database_path = Path(__file__).with_name("gestorman.db")
        self._init_database()
        self.title("GestorMan | Gestão de Manutenção")
        self.geometry("1280x760")
        self.minsize(1020, 620)
        self.state("zoomed")
        self.configure(bg="#edf2f7")
        self._build_style()
        self._build_layout()

    def _init_database(self):
        """Cria ou migra a tabela local de itens de almoxarifado."""
        with sqlite3.connect(self.database_path) as connection:
            migrate_almoxarifado(connection)
            connection.execute("""
                CREATE TABLE IF NOT EXISTS ordens_servico (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    setor TEXT NOT NULL,
                    numero_equipamento TEXT NOT NULL,
                    solicitante TEXT NOT NULL,
                    hora_parada TEXT,
                    tipo_servico TEXT,
                    prioridade TEXT,
                    especialidade TEXT,
                    descricao TEXT NOT NULL,
                    tecnico TEXT NOT NULL,
                    turno TEXT,
                    hora_inicio TEXT,
                    hora_final TEXT,
                    tempo_parada TEXT,
                    tempo_servico TEXT,
                    tempo_resposta TEXT,
                    situacao TEXT
                )
            """)
            connection.execute("""
                CREATE TABLE IF NOT EXISTS recebimento_almoxarifado (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_recebimento TEXT,
                    codigo TEXT NOT NULL,
                    descricao TEXT NOT NULL,
                    und TEXT,
                    qtd TEXT,
                    fornecedor TEXT,
                    num_nota_fiscal TEXT,
                    data_protocolo TEXT
                )
            """)
            connection.execute("""
                CREATE TABLE IF NOT EXISTS saida_almoxarifado (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data TEXT,
                    codigo TEXT NOT NULL,
                    descricao TEXT NOT NULL,
                    un TEXT,
                    qtd TEXT,
                    turno TEXT,
                    aplicacao TEXT,
                    requisitante TEXT
                )
            """)

    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 8), background="white",
                        fieldbackground="white", foreground="#26384a")
        style.configure("Treeview.Heading", font=("Segoe UI", 8, "bold"),
                        background="#dce7f2", foreground="#304860", relief="flat")
        style.map("Treeview", background=[("selected", "#1678bf")], foreground=[("selected", "white")])
        style.configure("TCombobox", padding=4)

    def _build_layout(self):
        sidebar = tk.Frame(self, bg="#1f3448", width=245)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="GESTORMAN", bg="#193044", fg="white", anchor="w",
                 font=("Segoe UI", 14, "bold"), padx=22, pady=16).pack(fill="x")
        tk.Label(sidebar, text="GESTÃO DE MANUTENÇÃO", bg="#193044", fg="#9eb5c8",
                 anchor="w", font=("Segoe UI", 7), padx=23).pack(fill="x", pady=(0, 14))

        self.menu_area = tk.Frame(sidebar, bg="#1f3448")
        self.menu_area.pack(fill="both", expand=True, pady=4)
        self._menu_header("Cadastros", True, selected=True)
        self._menu_subitem("Ordens de Serviços", self.show_os)
        self._menu_subitem("Recebimento de Itens", self.show_almox_recebimento)
        self._menu_subitem("Saída de Itens", self.show_almox_saida)
        self._menu_subitem("Resumo", self.show_almox_resumo)

        self.main = tk.Frame(self, bg="#edf2f7")
        self.main.pack(side="left", fill="both", expand=True)
        self.show_os()

    def _menu_item(self, icon, text, command):
        tk.Button(self.menu_area, text=f" {icon}   {text}", command=command, anchor="w",
                  bd=0, relief="flat", bg="#1f3448", activebackground="#2e4b64",
                  activeforeground="white", fg="#d5e3ef", font=("Segoe UI", 9),
                  padx=17, pady=8).pack(fill="x")

    def _menu_header(self, text, expanded, selected=False):
        mark = "⌄" if expanded else "›"
        bg = "#1f3448"
        tk.Label(self.menu_area, text=f" {mark}   {text}", anchor="w", bg=bg, fg="white",
                 font=("Segoe UI", 10, "bold"), padx=18, pady=8).pack(fill="x")

    def _menu_subitem(self, text, command):
        tk.Button(self.menu_area, text=text, command=command, anchor="w", bd=0,
                  relief="flat", bg="#1f3448", activebackground="#2e4b64",
                  activeforeground="white", fg="#b9c9d7", font=("Segoe UI", 8),
                  padx=32, pady=7, cursor="hand2").pack(fill="x")

    def _clear_main(self):
        for widget in self.main.winfo_children():
            widget.destroy()

    def _page_header(self, title, breadcrumb):
        topbar = tk.Frame(self.main, bg="white", height=74)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        tk.Label(topbar, text=title, bg="white", fg="#243b53",
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=30, pady=18)
        tk.Label(topbar, text=breadcrumb, bg="white", fg="#718096",
                 font=("Segoe UI", 8)).pack(side="right", padx=30)
        body = tk.Frame(self.main, bg="#edf2f7")
        body.pack(fill="both", expand=True, padx=25, pady=20)
        return body

    def show_os(self):
        self._clear_main()
        body = self._page_header("Cadastro de Ordem de Serviço", "Manutenção  /  Ordem de Serviços")
        self._build_form(body)
        self._build_actions(body)
        self._build_table(body)
        self._load_os_items()
        self.new_os()

    def show_almoxarifado(self):
        self._clear_main()
        body = self._page_header("Cadastro de Almoxarifado", "Cadastros  /  Almoxarifado")
        self._build_almox_form(body)
        self._build_almox_actions(body)
        self._build_almox_table(body)
        self._load_almox_items()

    def _build_almox_form(self, parent):
        box = tk.LabelFrame(parent, text="  Informações do Item de Almoxarifado  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        box.grid_columnconfigure(0, weight=1)
        box.grid_columnconfigure(1, weight=1)
        box.grid_columnconfigure(2, weight=2)
        box.grid_columnconfigure(3, weight=2)
        box.grid_columnconfigure(4, weight=0, minsize=70)

        fields = [
            ("Origem", "origem"), ("Código", "codigo"), ("Descrição", "descricao"),
            ("Descrição Inglês", "descricao_ingles"), ("Entradas", "entradas"),
            ("Saída", "saida"), ("Saldo", "saldo"), ("Data Inventário", "data_inventario"),
            ("Responsável", "responsavel"), ("Lead Time (dias)", "lead_time"),
            ("Média Consumo", "media_consumo"), ("Estoque Mínimo", "estoque_minimo"),
            ("Estoque Máximo", "estoque_maximo"), ("Est. Segurança", "est_seguranca"),
            ("Ponto Pedido", "ponto_pedido"), ("Ressuprimento", "ressuprimento"),
            ("Status", "status"), ("Custo Médio", "custo_medio"), ("Custo TTL", "custo_ttl"),
        ]
        # Campos que são calculados automaticamente (read-only)
        read_only_fields = ("entradas", "saida", "saldo")
        
        self.almox_vars = {key: tk.StringVar() for _, key in fields}
        for idx, (label, key) in enumerate(fields):
            row, col = divmod(idx, 5)
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
            # Campos de entradas, saída e saldo são somente leitura (calculados automaticamente)
            if key in read_only_fields:
                tk.Entry(cell, textvariable=self.almox_vars[key], font=("Segoe UI", 8),
                         relief="solid", bd=1, state="readonly", readonlybackground="#e8eef5").pack(fill="x", ipady=2)
            else:
                tk.Entry(cell, textvariable=self.almox_vars[key], font=("Segoe UI", 8),
                         relief="solid", bd=1).pack(fill="x", ipady=2)

    def _build_almox_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "＋  Novo Item", self.new_almox_item, "#ffffff", "#315a7c")
        self._action(actions, "▣  Inserir", self.insert_almox_item, "#1678bf", "white")
        self._action(actions, "✓  Salvar edição", self.update_almox_item, "#ffffff", "#315a7c")
        self._action(actions, "⇧  Carga TXT", self.import_almox_txt, "#ffffff", "#315a7c")
        self._action(actions, "✕  Excluir", self.delete_almox_item, "#ffffff", "#b23b3b")

    def _build_almox_table(self, parent):
        search = tk.Frame(parent, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        search_content = tk.Frame(search, bg="#edf2f7")
        search_content.pack()
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.almox_search = tk.StringVar()
        self.almox_search.trace_add("write", self.filter_almox_items)
        tk.Entry(search_content, textvariable=self.almox_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        headings = ("ORIGEM", "CÓDIGO", "DESCRIÇÃO", "DESCRIÇÃO INGLÊS", "ENTRADAS",
                "SAÍDA", "SALDO", "DATA INVENTÁRIO", "RESPONSÁVEL", "LEAD TIME (DIAS)",
                "MÉDIA CONSUMO", "ESTOQUE MÍNIMO", "ESTOQUE MÁXIMO", "EST. SEGURANÇA",
                "PONTO PEDIDO", "RESSUPRIMENTO", "STATUS", "CUSTO MÉDIO", "CUSTO TTL")
        widths = (120, 100, 230, 180, 90, 80, 80, 110, 130, 95, 100, 100, 100, 105, 95, 110, 90, 110, 100)
        self.almox_grid = ttk.Treeview(frame, columns=self.ALMOX_COLUMNS, show="headings", selectmode="browse")
        # Configura tag para valores negativos (saldo negativo em vermelho)
        self.almox_grid.tag_configure("negative", foreground="#d32f2f")
        for col, heading, width in zip(self.ALMOX_COLUMNS, headings, widths):
            self.almox_grid.heading(col, text=heading)
            # Centraliza colunas de valores numéricos
            anchor = "center" if col in ("entradas", "saida", "saldo") else "w"
            self.almox_grid.column(col, width=width, minwidth=65, anchor=anchor, stretch=(col in ("descricao", "descricao_ingles")))
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.almox_grid.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.almox_grid.xview)
        self.almox_grid.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.almox_grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.almox_grid.bind("<<TreeviewSelect>>", self._select_almox_row)

    def _load_almox_items(self):
        """Consulta no SQLite os itens com cálculos de entradas, saída e saldo."""
        for item in self.almox_grid.get_children():
            self.almox_grid.delete(item)
        term = self.almox_search.get().strip() if hasattr(self, "almox_search") else ""
        
        # Query com LEFT JOINs para calcular entradas, saída e saldo a partir de outras tabelas
        # Fórmula do Saldo: Entrada - Saída
        query = """
            SELECT 
                a.id,
                a.origem,
                a.codigo,
                a.descricao,
                a.descricao_ingles,
                COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(r.qtd), '.', ''), ',', '.') AS REAL)), 0) as entradas,
                COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(s.qtd), '.', ''), ',', '.') AS REAL)), 0) as saida,
                (COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(r.qtd), '.', ''), ',', '.') AS REAL)), 0) - COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(s.qtd), '.', ''), ',', '.') AS REAL)), 0)) as saldo,
                a.data_inventario,
                a.responsavel,
                a.lead_time,
                a.media_consumo,
                a.estoque_minimo,
                a.estoque_maximo,
                a.est_seguranca,
                a.ponto_pedido,
                a.ressuprimento,
                a.status,
                a.custo_medio,
                a.custo_ttl
            FROM almoxarifado a
            LEFT JOIN recebimento_almoxarifado r ON a.codigo = r.codigo
            LEFT JOIN saida_almoxarifado s ON a.codigo = s.codigo
        """
        
        params = ()
        if term:
            condition = " OR ".join(f"a.{column} LIKE ?" for column in self.ALMOX_COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.ALMOX_COLUMNS)
        
        query += " GROUP BY a.id, a.codigo ORDER BY a.origem, a.codigo"
        
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        
        for row in rows:
            row_id = row[0]
            values = row[1:]
            # Formata os valores numéricos como inteiro
            formatted_values = []
            saldo_value = None
            for idx, val in enumerate(values):
                if idx in (5, 6, 7):  # entradas, saida, saldo
                    if isinstance(val, (int, float)):
                        int_val = int(val)
                        if idx == 7:  # Saldo com sinal
                            saldo_value = int_val
                            formatted_values.append(f"{int_val:+d}")
                        else:
                            formatted_values.append(str(int_val))
                    else:
                        if idx == 7:
                            formatted_values.append("+0")
                        else:
                            formatted_values.append("0")
                else:
                    formatted_values.append(str(val) if val else "")
            # Aplica tag "negative" se o saldo for negativo
            tags = ("negative",) if saldo_value is not None and saldo_value < 0 else ()
            self.almox_grid.insert("", "end", iid=str(row_id), values=formatted_values, tags=tags)

    def filter_almox_items(self, *_args):
        if hasattr(self, "almox_grid"):
            self._load_almox_items()

    def new_almox_item(self):
        for var in self.almox_vars.values():
            var.set("")
        self.almox_grid.selection_remove(self.almox_grid.selection())

    def insert_almox_item(self):
        required = ("origem", "codigo", "descricao")
        if any(not self.almox_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha origem, código e descrição.")
            return
        # Exclui campos calculados automaticamente
        editable_fields = tuple(col for col in self.ALMOX_COLUMNS if col not in ("entradas", "saida", "saldo"))
        values = tuple(self.almox_vars[key].get().strip() for key in editable_fields)
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(editable_fields)
                placeholders = ", ".join("?" for _ in editable_fields)
                connection.execute(f"INSERT INTO almoxarifado ({columns}) VALUES ({placeholders})", values)
        except sqlite3.IntegrityError:
            messagebox.showwarning("Código duplicado", "Já existe um item com esta origem e código.")
            return
        self._load_almox_items()
        self.new_almox_item()

    def update_almox_item(self):
        selected = self.almox_grid.selection()
        if not selected:
            messagebox.showinfo("Salvar edição", "Selecione um item para editar.")
            return
        required = ("origem", "codigo", "descricao")
        if any(not self.almox_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha origem, código e descrição.")
            return
        # Exclui campos calculados automaticamente
        editable_fields = tuple(col for col in self.ALMOX_COLUMNS if col not in ("entradas", "saida", "saldo"))
        values = tuple(self.almox_vars[key].get().strip() for key in editable_fields)
        assignments = ", ".join(f"{column} = ?" for column in editable_fields)
        try:
            with sqlite3.connect(self.database_path) as connection:
                connection.execute(f"UPDATE almoxarifado SET {assignments} WHERE id = ?",
                                   (*values, int(selected[0])))
        except sqlite3.IntegrityError:
            messagebox.showwarning("Código duplicado", "Já existe um item com esta origem e código.")
            return
        self._load_almox_items()
        self.new_almox_item()

    def delete_almox_item(self):
        selected = self.almox_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir item", "Selecione um item na lista.")
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM almoxarifado WHERE id = ?", (int(selected[0]),))
        self._load_almox_items()
        self.new_almox_item()

    @staticmethod
    def _normalize_header(value):
        """Normaliza títulos para reconhecer cabeçalhos com ou sem acentos."""
        text = unicodedata.normalize("NFKD", value).encode("ASCII", "ignore").decode()
        return "".join(character for character in text.upper() if character.isalnum())

    def _read_txt_rows(self, file_path):
        """Lê TXT separado por ;, tab, | ou vírgula, com cabeçalho opcional."""
        content = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                content = Path(file_path).read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            raise ValueError("Não foi possível identificar a codificação do arquivo TXT.")

        lines = [line for line in content.splitlines() if line.strip()]
        if not lines:
            raise ValueError("O arquivo TXT está vazio.")
        try:
            delimiter = csv.Sniffer().sniff("\n".join(lines[:10]), delimiters=";\t|,").delimiter
        except csv.Error:
            delimiter = ";"
        rows = list(csv.reader(lines, delimiter=delimiter))
        headers = [self._normalize_header(value) for value in rows[0]]
        aliases = {
            "ORIGEM": "origem", "BASE": "origem", "CODIGO": "codigo", "DESCRICAO": "descricao",
            "DESCRICAOINGLES": "descricao_ingles", "DESCRIPTIONENGLISH": "descricao_ingles",
            "ENTRADAS": "entradas", "SAIDA": "saida", "SAÍDA": "saida", "SALDO": "saldo",
            "ESTOQUEMINIMO": "estoque_minimo", "ESTOQUEMAXIMO": "estoque_maximo",
            "ESTSEGURANCA": "est_seguranca", "EST. SEGURANCA": "est_seguranca",
            "MEDIACONSUMO": "media_consumo", "MEDIA CONSUMO": "media_consumo",
            "QTDATUAL": "entradas", "QTD": "entradas", "QUANTIDADE": "entradas",
            "DATAINVENTARIO": "data_inventario", "RESPONSAVEL": "responsavel",
            "LEADTIME": "lead_time", "DIASENTREGA": "lead_time",
            "PONTOPEDIDO": "ponto_pedido", "PONTO_PEDIDO": "ponto_pedido",
            "RESSUPRIMENTO": "ressuprimento", "STATUS": "status",
            "CUSTOMEDIO": "custo_medio", "CUSTO_MEDIO": "custo_medio",
            "CUSTOUNITARIO": "custo_medio", "CUSTO": "custo_medio",
            "CUSTOTTL": "custo_ttl", "CUSTOTOTAL": "custo_ttl",
        }
        indexes = {aliases[header]: position for position, header in enumerate(headers) if header in aliases}
        has_header = all(field in indexes for field in ("origem", "codigo", "descricao"))
        data_rows = rows[1:] if has_header else rows
        loaded_rows, skipped = [], 0
        for row in data_rows:
            if has_header:
                values = tuple(row[indexes[field]].strip() if indexes.get(field, -1) < len(row) else ""
                               for field in self.ALMOX_COLUMNS)
            else:
                values = tuple(value.strip() for value in row[:len(self.ALMOX_COLUMNS)])
                values += ("",) * (len(self.ALMOX_COLUMNS) - len(values))
            if not any(values):
                continue
            if not all(values[index].strip() for index in (0, 1, 2)):
                skipped += 1
                continue
            loaded_rows.append(values)
        return loaded_rows, skipped

    def import_almox_txt(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT de almoxarifado",
            filetypes=(("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")),
        )
        if not file_path:
            return
        try:
            rows, skipped = self._read_txt_rows(file_path)
            if not rows:
                messagebox.showwarning("Carga TXT", "Nenhum item válido foi encontrado no arquivo.")
                return
            columns = ", ".join(self.ALMOX_COLUMNS)
            placeholders = ", ".join("?" for _ in self.ALMOX_COLUMNS)
            updates = ", ".join(f"{column} = excluded.{column}" for column in self.ALMOX_COLUMNS[2:])
            query = (f"INSERT INTO almoxarifado ({columns}) VALUES ({placeholders}) "
                     f"ON CONFLICT(origem, codigo) DO UPDATE SET {updates}")
            with sqlite3.connect(self.database_path) as connection:
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_almox_items()
        message = f"{len(rows)} item(ns) carregado(s) ou atualizado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter origem, código e descrição."
        messagebox.showinfo("Carga TXT concluída", message)

    def _select_almox_row(self, _event):
        selected = self.almox_grid.selection()
        if not selected:
            return
        values = self.almox_grid.item(selected[0], "values")
        for key, value in zip(self.ALMOX_COLUMNS, values):
            self.almox_vars[key].set(value)

    def show_almox_recebimento(self):
        self._clear_main()
        body = self._page_header("Recebimento de Itens de Almoxarifado", "Cadastros  /  Almoxarifado  /  Recebimento")
        self._build_recebimento_form(body)
        self._build_recebimento_actions(body)
        self._build_recebimento_table(body)
        self._load_recebimento_items()
        self.new_recebimento_item()

    def _build_recebimento_form(self, parent):
        box = tk.LabelFrame(parent, text="  Dados de Recebimento  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        box.grid_columnconfigure(0, weight=1)
        box.grid_columnconfigure(1, weight=1)
        box.grid_columnconfigure(2, weight=1)
        box.grid_columnconfigure(3, weight=1)

        fields = [
            ("Data Recebimento", "data_recebimento"), ("Código", "codigo"), 
            ("Descrição", "descricao"), ("Unidade", "und"),
            ("Quantidade", "qtd"), ("Fornecedor", "fornecedor"), 
            ("Nº Nota Fiscal", "num_nota_fiscal"), ("Data / Protocolo", "data_protocolo"),
        ]
        self.recebimento_vars = {key: tk.StringVar() for _, key in fields}
        for idx, (label, key) in enumerate(fields):
            row, col = divmod(idx, 4)
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
            tk.Entry(cell, textvariable=self.recebimento_vars[key], font=("Segoe UI", 8),
                     relief="solid", bd=1).pack(fill="x", ipady=2)

    def _build_recebimento_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "＋  Novo Recebimento", self.new_recebimento_item, "#ffffff", "#315a7c")
        self._action(actions, "▣  Inserir", self.insert_recebimento_item, "#1678bf", "white")
        self._action(actions, "✓  Salvar edição", self.update_recebimento_item, "#ffffff", "#315a7c")
        self._action(actions, "⇧  Carga TXT", self.import_recebimento_txt, "#ffffff", "#315a7c")
        self._action(actions, "✕  Excluir", self.delete_recebimento_item, "#ffffff", "#b23b3b")

    def _build_recebimento_table(self, parent):
        search = tk.Frame(parent, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        search_content = tk.Frame(search, bg="#edf2f7")
        search_content.pack()
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.recebimento_search = tk.StringVar()
        self.recebimento_search.trace_add("write", self.filter_recebimento_items)
        tk.Entry(search_content, textvariable=self.recebimento_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        headings = ("DATA RECEBIMENTO", "CÓDIGO", "DESCRIÇÃO", "UND", "QTD", 
                    "FORNECEDOR", "Nº NOTA FISCAL", "DATA / PROTOCOLO")
        widths = (140, 100, 250, 70, 80, 180, 130, 140)
        self.recebimento_grid = ttk.Treeview(frame, columns=self.RECEBIMENTO_COLUMNS, show="headings", selectmode="browse")
        for col, heading, width in zip(self.RECEBIMENTO_COLUMNS, headings, widths):
            self.recebimento_grid.heading(col, text=heading)
            self.recebimento_grid.column(col, width=width, minwidth=65, stretch=(col == "descricao"))
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.recebimento_grid.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.recebimento_grid.xview)
        self.recebimento_grid.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.recebimento_grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.recebimento_grid.bind("<<TreeviewSelect>>", self._select_recebimento_row)

    def _load_recebimento_items(self):
        """Consulta no SQLite os itens de recebimento que atendem ao texto pesquisado."""
        for item in self.recebimento_grid.get_children():
            self.recebimento_grid.delete(item)
        term = self.recebimento_search.get().strip() if hasattr(self, "recebimento_search") else ""
        columns = ", ".join(self.RECEBIMENTO_COLUMNS)
        query = f"SELECT id, {columns} FROM recebimento_almoxarifado"
        params = ()
        if term:
            condition = " OR ".join(f"{column} LIKE ?" for column in self.RECEBIMENTO_COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.RECEBIMENTO_COLUMNS)
        query += " ORDER BY data_recebimento DESC"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            self.recebimento_grid.insert("", "end", iid=str(row_id), values=values)

    def filter_recebimento_items(self, *_args):
        if hasattr(self, "recebimento_grid"):
            self._load_recebimento_items()

    def new_recebimento_item(self):
        for var in self.recebimento_vars.values():
            var.set("")
        self.recebimento_grid.selection_remove(self.recebimento_grid.selection())

    def insert_recebimento_item(self):
        required = ("data_recebimento", "codigo", "descricao")
        if any(not self.recebimento_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha data de recebimento, código e descrição.")
            return
        values = tuple(self.recebimento_vars[key].get().strip() for key in self.RECEBIMENTO_COLUMNS)
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(self.RECEBIMENTO_COLUMNS)
                placeholders = ", ".join("?" for _ in self.RECEBIMENTO_COLUMNS)
                connection.execute(f"INSERT INTO recebimento_almoxarifado ({columns}) VALUES ({placeholders})", values)
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao inserir", f"Não foi possível inserir o item.\n\n{error}")
            return
        self._load_recebimento_items()
        self.new_recebimento_item()

    def update_recebimento_item(self):
        selected = self.recebimento_grid.selection()
        if not selected:
            messagebox.showinfo("Salvar edição", "Selecione um item para editar.")
            return
        required = ("data_recebimento", "codigo", "descricao")
        if any(not self.recebimento_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha data de recebimento, código e descrição.")
            return
        values = tuple(self.recebimento_vars[key].get().strip() for key in self.RECEBIMENTO_COLUMNS)
        assignments = ", ".join(f"{column} = ?" for column in self.RECEBIMENTO_COLUMNS)
        try:
            with sqlite3.connect(self.database_path) as connection:
                connection.execute(f"UPDATE recebimento_almoxarifado SET {assignments} WHERE id = ?",
                                   (*values, int(selected[0])))
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao atualizar", f"Não foi possível atualizar o item.\n\n{error}")
            return
        self._load_recebimento_items()
        self.new_recebimento_item()

    def delete_recebimento_item(self):
        selected = self.recebimento_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir item", "Selecione um item na lista.")
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM recebimento_almoxarifado WHERE id = ?", (int(selected[0]),))
        self._load_recebimento_items()
        self.new_recebimento_item()

    def show_almox_resumo(self):
        """Exibe o saldo consolidado de entradas e saídas por item."""
        self._clear_main()
        body = self._page_header("Resumo de Movimentação de Itens", "Cadastros  /  Almoxarifado  /  Resumo")

        search = tk.Frame(body, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        search_content = tk.Frame(search, bg="#edf2f7")
        search_content.pack()
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.resumo_search = tk.StringVar()
        self.resumo_search.trace_add("write", self._load_resumo_items)
        tk.Entry(search_content, textvariable=self.resumo_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(body, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        columns = ("codigo", "descricao", "entradas", "saidas", "saldo")
        headings = ("CÓDIGO", "DESCRIÇÃO", "ENTRADAS", "SAÍDAS", "SALDO")
        widths = (130, 430, 110, 110, 110)
        self.resumo_grid = ttk.Treeview(frame, columns=columns, show="headings")
        for col, heading, width in zip(columns, headings, widths):
            self.resumo_grid.heading(col, text=heading)
            self.resumo_grid.column(col, width=width, minwidth=80,
                                    anchor="center" if col in ("entradas", "saidas", "saldo") else "w",
                                    stretch=(col == "descricao"))
        self.resumo_grid.tag_configure("negative", foreground="#b23b3b")
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.resumo_grid.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.resumo_grid.xview)
        self.resumo_grid.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.resumo_grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self._load_resumo_items()

    def _load_resumo_items(self, *_args):
        """Carrega entradas, saídas e saldo por código sem multiplicar movimentações."""
        if not hasattr(self, "resumo_grid"):
            return
        for item in self.resumo_grid.get_children():
            self.resumo_grid.delete(item)
        term = self.resumo_search.get().strip() if hasattr(self, "resumo_search") else ""
        query = """
            WITH recebimentos AS (
                SELECT codigo, MAX(descricao) AS descricao,
                       COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(qtd), '.', ''), ',', '.') AS REAL)), 0) AS entradas
                FROM recebimento_almoxarifado GROUP BY codigo
            ), saidas AS (
                SELECT codigo, MAX(descricao) AS descricao,
                       COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(qtd), '.', ''), ',', '.') AS REAL)), 0) AS saidas
                FROM saida_almoxarifado GROUP BY codigo
            ), codigos AS (
                SELECT codigo FROM recebimentos UNION SELECT codigo FROM saidas
            )
            SELECT c.codigo, COALESCE(r.descricao, s.descricao, ''),
                   COALESCE(r.entradas, 0), COALESCE(s.saidas, 0),
                   COALESCE(r.entradas, 0) - COALESCE(s.saidas, 0)
            FROM codigos c
            LEFT JOIN recebimentos r ON r.codigo = c.codigo
            LEFT JOIN saidas s ON s.codigo = c.codigo
        """
        params = ()
        if term:
            query += " WHERE c.codigo LIKE ? OR COALESCE(r.descricao, s.descricao, '') LIKE ?"
            params = (f"%{term}%", f"%{term}%")
        query += " ORDER BY c.codigo"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for codigo, descricao, entradas, saidas, saldo in rows:
            def format_qtd(value, signed=False):
                value = float(value)
                formatted = str(int(value)) if value.is_integer() else f"{value:.2f}".replace(".", ",")
                return f"{value:+g}" if signed and value.is_integer() else (f"+{formatted}" if signed and value >= 0 else formatted)
            self.resumo_grid.insert("", "end", values=(
                codigo, descricao, format_qtd(entradas), format_qtd(saidas), format_qtd(saldo, signed=True)
            ), tags=("negative",) if saldo < 0 else ())

    def _select_recebimento_row(self, _event):
        selected = self.recebimento_grid.selection()
        if not selected:
            return
        values = self.recebimento_grid.item(selected[0], "values")
        for key, value in zip(self.RECEBIMENTO_COLUMNS, values):
            self.recebimento_vars[key].set(value)

    def _read_recebimento_txt_rows(self, file_path):
        """Lê TXT de recebimento separado por ;, tab, | ou vírgula, com cabeçalho opcional."""
        content = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                content = Path(file_path).read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            raise ValueError("Não foi possível identificar a codificação do arquivo TXT.")

        lines = [line for line in content.splitlines() if line.strip()]
        if not lines:
            raise ValueError("O arquivo TXT está vazio.")
        try:
            delimiter = csv.Sniffer().sniff("\n".join(lines[:10]), delimiters=";\t|,").delimiter
        except csv.Error:
            delimiter = ";"
        rows = list(csv.reader(lines, delimiter=delimiter))
        headers = [self._normalize_header(value) for value in rows[0]]
        aliases = {
            "DATARECEBIMENTO": "data_recebimento", "DATA_RECEBIMENTO": "data_recebimento",
            "DATAREC": "data_recebimento", "CODIGO": "codigo", "DESCRICAO": "descricao",
            "UND": "und", "UNIDADE": "und", "QTD": "qtd", "QUANTIDADE": "qtd",
            "FORNECEDOR": "fornecedor", "NUMNOTAFISCAL": "num_nota_fiscal",
            "NUMERONOTAFISCAL": "num_nota_fiscal", "NOTAFISCAL": "num_nota_fiscal",
            "NNOTAFISCAL": "num_nota_fiscal", "DATAPROTOCOLO": "data_protocolo",
            "DATA_PROTOCOLO": "data_protocolo", "PROTOCOLO": "data_protocolo",
        }
        indexes = {aliases[header]: position for position, header in enumerate(headers) if header in aliases}
        has_header = all(field in indexes for field in ("data_recebimento", "codigo", "descricao"))
        data_rows = rows[1:] if has_header else rows
        loaded_rows, skipped = [], 0
        for row in data_rows:
            if has_header:
                values = tuple(row[indexes[field]].strip() if indexes.get(field, -1) < len(row) else ""
                               for field in self.RECEBIMENTO_COLUMNS)
            else:
                values = tuple(value.strip() for value in row[:len(self.RECEBIMENTO_COLUMNS)])
                values += ("",) * (len(self.RECEBIMENTO_COLUMNS) - len(values))
            if not any(values):
                continue
            if not all(values[index].strip() for index in (0, 1, 2)):
                skipped += 1
                continue
            loaded_rows.append(values)
        return loaded_rows, skipped

    def import_recebimento_txt(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT de recebimento de materiais",
            filetypes=(("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")),
        )
        if not file_path:
            return
        try:
            rows, skipped = self._read_recebimento_txt_rows(file_path)
            if not rows:
                messagebox.showwarning("Carga TXT", "Nenhum item válido foi encontrado no arquivo.")
                return
            columns = ", ".join(self.RECEBIMENTO_COLUMNS)
            placeholders = ", ".join("?" for _ in self.RECEBIMENTO_COLUMNS)
            query = f"INSERT INTO recebimento_almoxarifado ({columns}) VALUES ({placeholders})"
            with sqlite3.connect(self.database_path) as connection:
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_recebimento_items()
        message = f"{len(rows)} item(ns) carregado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter data, código e descrição."
        messagebox.showinfo("Carga TXT concluída", message)

    def show_almox_saida(self):
        self._clear_main()
        body = self._page_header("Saída de Itens de Almoxarifado", "Cadastros  /  Almoxarifado  /  Saída")
        self._build_saida_form(body)
        self._build_saida_actions(body)
        self._build_saida_table(body)
        self._load_saida_items()
        self.new_saida_item()

    def _build_saida_form(self, parent):
        box = tk.LabelFrame(parent, text="  Dados de Saída  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        box.grid_columnconfigure(0, weight=1)
        box.grid_columnconfigure(1, weight=1)
        box.grid_columnconfigure(2, weight=1)
        box.grid_columnconfigure(3, weight=1)

        fields = [
            ("Data", "data"), ("Código", "codigo"), 
            ("Descrição", "descricao"), ("Unidade", "un"),
            ("Quantidade", "qtd"), ("Turno", "turno"), 
            ("Aplicação", "aplicacao"), ("Requisitante", "requisitante"),
        ]
        self.saida_vars = {key: tk.StringVar() for _, key in fields}
        for idx, (label, key) in enumerate(fields):
            row, col = divmod(idx, 4)
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
            tk.Entry(cell, textvariable=self.saida_vars[key], font=("Segoe UI", 8),
                     relief="solid", bd=1).pack(fill="x", ipady=2)

    def _build_saida_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "＋  Nova Saída", self.new_saida_item, "#ffffff", "#315a7c")
        self._action(actions, "▣  Inserir", self.insert_saida_item, "#1678bf", "white")
        self._action(actions, "✓  Salvar edição", self.update_saida_item, "#ffffff", "#315a7c")
        self._action(actions, "⇧  Carga TXT", self.import_saida_txt, "#ffffff", "#315a7c")
        self._action(actions, "✕  Excluir", self.delete_saida_item, "#ffffff", "#b23b3b")

    def _build_saida_table(self, parent):
        search = tk.Frame(parent, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        search_content = tk.Frame(search, bg="#edf2f7")
        search_content.pack()
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.saida_search = tk.StringVar()
        self.saida_search.trace_add("write", self.filter_saida_items)
        tk.Entry(search_content, textvariable=self.saida_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        headings = ("DATA", "CÓDIGO", "DESCRIÇÃO", "UN", "QTD", 
                    "TURNO", "APLICAÇÃO", "REQUISITANTE")
        widths = (120, 100, 250, 70, 80, 80, 180, 150)
        self.saida_grid = ttk.Treeview(frame, columns=self.SAIDA_COLUMNS, show="headings", selectmode="browse")
        for col, heading, width in zip(self.SAIDA_COLUMNS, headings, widths):
            self.saida_grid.heading(col, text=heading)
            self.saida_grid.column(col, width=width, minwidth=65, stretch=(col == "descricao"))
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.saida_grid.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.saida_grid.xview)
        self.saida_grid.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.saida_grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.saida_grid.bind("<<TreeviewSelect>>", self._select_saida_row)

    def _load_saida_items(self):
        """Consulta no SQLite os itens de saída que atendem ao texto pesquisado."""
        for item in self.saida_grid.get_children():
            self.saida_grid.delete(item)
        term = self.saida_search.get().strip() if hasattr(self, "saida_search") else ""
        columns = ", ".join(self.SAIDA_COLUMNS)
        query = f"SELECT id, {columns} FROM saida_almoxarifado"
        params = ()
        if term:
            condition = " OR ".join(f"{column} LIKE ?" for column in self.SAIDA_COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.SAIDA_COLUMNS)
        query += " ORDER BY data DESC"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            self.saida_grid.insert("", "end", iid=str(row_id), values=values)

    def filter_saida_items(self, *_args):
        if hasattr(self, "saida_grid"):
            self._load_saida_items()

    def new_saida_item(self):
        for var in self.saida_vars.values():
            var.set("")
        self.saida_grid.selection_remove(self.saida_grid.selection())

    def insert_saida_item(self):
        required = ("data", "codigo", "descricao")
        if any(not self.saida_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha data, código e descrição.")
            return
        values = tuple(self.saida_vars[key].get().strip() for key in self.SAIDA_COLUMNS)
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(self.SAIDA_COLUMNS)
                placeholders = ", ".join("?" for _ in self.SAIDA_COLUMNS)
                connection.execute(f"INSERT INTO saida_almoxarifado ({columns}) VALUES ({placeholders})", values)
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao inserir", f"Não foi possível inserir o item.\n\n{error}")
            return
        self._load_saida_items()
        self.new_saida_item()

    def update_saida_item(self):
        selected = self.saida_grid.selection()
        if not selected:
            messagebox.showinfo("Salvar edição", "Selecione um item para editar.")
            return
        required = ("data", "codigo", "descricao")
        if any(not self.saida_vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha data, código e descrição.")
            return
        values = tuple(self.saida_vars[key].get().strip() for key in self.SAIDA_COLUMNS)
        assignments = ", ".join(f"{column} = ?" for column in self.SAIDA_COLUMNS)
        try:
            with sqlite3.connect(self.database_path) as connection:
                connection.execute(f"UPDATE saida_almoxarifado SET {assignments} WHERE id = ?",
                                   (*values, int(selected[0])))
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao atualizar", f"Não foi possível atualizar o item.\n\n{error}")
            return
        self._load_saida_items()
        self.new_saida_item()

    def delete_saida_item(self):
        selected = self.saida_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir item", "Selecione um item na lista.")
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM saida_almoxarifado WHERE id = ?", (int(selected[0]),))
        self._load_saida_items()
        self.new_saida_item()

    def _select_saida_row(self, _event):
        selected = self.saida_grid.selection()
        if not selected:
            return
        values = self.saida_grid.item(selected[0], "values")
        for key, value in zip(self.SAIDA_COLUMNS, values):
            self.saida_vars[key].set(value)

    def _read_saida_txt_rows(self, file_path):
        """Lê TXT de saída separado por ;, tab, | ou vírgula, com cabeçalho opcional."""
        content = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                content = Path(file_path).read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            raise ValueError("Não foi possível identificar a codificação do arquivo TXT.")

        lines = [line for line in content.splitlines() if line.strip()]
        if not lines:
            raise ValueError("O arquivo TXT está vazio.")
        try:
            delimiter = csv.Sniffer().sniff("\n".join(lines[:10]), delimiters=";\t|,").delimiter
        except csv.Error:
            delimiter = ";"
        rows = list(csv.reader(lines, delimiter=delimiter))
        headers = [self._normalize_header(value) for value in rows[0]]
        aliases = {
            "DATA": "data", "CODIGO": "codigo", "DESCRICAO": "descricao",
            "UN": "un", "UNIDADE": "un", "QTD": "qtd", "QUANTIDADE": "qtd",
            "TURNO": "turno", "APLICACAO": "aplicacao", "REQUISITANTE": "requisitante",
        }
        indexes = {aliases[header]: position for position, header in enumerate(headers) if header in aliases}
        has_header = all(field in indexes for field in ("data", "codigo", "descricao"))
        data_rows = rows[1:] if has_header else rows
        loaded_rows, skipped = [], 0
        for row in data_rows:
            if has_header:
                values = tuple(row[indexes[field]].strip() if indexes.get(field, -1) < len(row) else ""
                               for field in self.SAIDA_COLUMNS)
            else:
                values = tuple(value.strip() for value in row[:len(self.SAIDA_COLUMNS)])
                values += ("",) * (len(self.SAIDA_COLUMNS) - len(values))
            if not any(values):
                continue
            if not all(values[index].strip() for index in (0, 1, 2)):
                skipped += 1
                continue
            loaded_rows.append(values)
        return loaded_rows, skipped

    def import_saida_txt(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT de saída de materiais",
            filetypes=(("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")),
        )
        if not file_path:
            return
        try:
            rows, skipped = self._read_saida_txt_rows(file_path)
            if not rows:
                messagebox.showwarning("Carga TXT", "Nenhum item válido foi encontrado no arquivo.")
                return
            columns = ", ".join(self.SAIDA_COLUMNS)
            placeholders = ", ".join("?" for _ in self.SAIDA_COLUMNS)
            query = f"INSERT INTO saida_almoxarifado ({columns}) VALUES ({placeholders})"
            with sqlite3.connect(self.database_path) as connection:
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_saida_items()
        message = f"{len(rows)} item(ns) carregado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter data, código e descrição."
        messagebox.showinfo("Carga TXT concluída", message)

    def _build_form(self, parent):
        box = tk.LabelFrame(parent, text="  Informações Gerais da Ordem de Serviço  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        box.grid_columnconfigure(0, weight=1)
        box.grid_columnconfigure(1, weight=2)
        box.grid_columnconfigure(2, weight=1)
        box.grid_columnconfigure(3, weight=1)
        box.grid_columnconfigure(4, weight=1)
        box.grid_columnconfigure(5, weight=0, minsize=90)

        fields = [
            ("Setor", "setor", "combo"), ("Nº Equipamento", "numero_equipamento", "combo"),
            ("Solicitante", "solicitante", "combo"), ("Hora Parada", "hora_parada", "entry"),
            ("Tipo Serviço", "tipo_servico", "combo"), ("Prioridade", "prioridade", "combo"),
            ("Especialidade", "especialidade", "combo"), ("Descrição", "descricao", "entry"),
            ("Técnico", "tecnico", "combo"), ("Turno", "turno", "combo"),
            ("Hora Início", "hora_inicio", "entry"), ("Hora Final", "hora_final", "entry"),
            ("Tempo Parada", "tempo_parada", "entry"), ("Tempo Serviço", "tempo_servico", "entry"),
            ("Tempo de Resposta", "tempo_resposta", "entry"), ("Situação", "situacao", "combo"),
        ]
        positions = {
            "setor": (0, 0, 1),
            "numero_equipamento": (0, 1, 1),
            "solicitante": (0, 2, 1),
            "hora_parada": (0, 3, 1),
            "tipo_servico": (0, 4, 1),
            "prioridade": (0, 5, 1),
            "especialidade": (1, 0, 1),
            "descricao": (1, 1, 3),
            "tecnico": (1, 4, 1),
            "turno": (1, 5, 1),
            "hora_inicio": (2, 0, 1),
            "hora_final": (2, 1, 1),
            "tempo_parada": (2, 2, 1),
            "tempo_servico": (2, 3, 1),
            "tempo_resposta": (2, 4, 1),
            "situacao": (2, 5, 1),
        }
        self.vars = {key: tk.StringVar() for _, key, _ in fields}
        defaults = {"tipo_servico": "CORRETIVA",
                    "prioridade": "URGENTE", "especialidade": "MECÂNICA",
                    "situacao": "ABERTA"}
        for key, value in defaults.items():
            self.vars[key].set(value)

        options = {
            "setor": (
                "ADM",
                "C.COSTURA",
                "C.LONA",
                "CABLEADORA",
                "CHILLER",
                "COMPRESSOR 100HP",
                "EXT.COATING",
                "EXT.COMPOSTO",
                "EXT.MULTIFILAMENTO",
                "EXT.RÁFIA",
                "GALPÃO 2",
                "M.COSTURA PANO",
                "PRENSA",
                "PRENSA DE TUBETES E RESÍDUOS",
                "QUALIDADE",
                "REDE DE AR COMPRIMIDO",
                "REDE ELÉTRICA",
                "SECADOR DE AR 100HP",
                "TEC.ALÇA",
                "TEC.CADARÇO",
                "TEC.LEVE",
                "TEC.PESADA",
            ),
            "numero_equipamento": (
                "CA0101", "CB0101", "CC0201", "CC0202", "CC0303", "CC0401",
                "CL001", "CL002", "CL003", "CL004", "EC001", "EL001", "EL002",
                "EL003", "EM001", "ER001", "ER002", "ER003", "ER004", "ER005",
                "ER006", "ER007", "ER008", "ER009", "ER010", "ER011", "ER012",
                "ER013", "ER014", "ER015", "ER016", "ER017", "ER018", "ER019",
                "ER020", "ER021", "ER022", "ER023", "ER024", "ER025", "ER026",
                "ER027", "ER028", "ER029", "ER030", "ER031", "ER032", "ER033",
                "ER034", "ER035", "ER036", "ER037", "ER038", "ER039", "ER040",
                "ER041", "ER042", "ER043", "ER044", "ER045", "ER046", "ER047",
                "ER048", "ER049", "ER050", "ER051", "ER052", "ER053", "ER054",
                "ER055", "ER056", "ER057", "ER058", "ER059", "ER060", "ER061",
                "ER062", "ER063", "ER064", "ER065", "ER066", "ER067", "ER068",
                "ER069", "ER070", "ER071", "ER072", "ER073", "ER074", "ER075",
                "ER076", "ER077", "ER078", "ER079", "ER080", "ER081", "ER082",
                "ER083", "ER084", "ER085", "ER086", "ER087", "ER088", "ER089",
                "ER090", "ER091", "ER092", "ER093", "ER094", "ER095", "ER096",
                "ER097", "ER098", "ER099", "ER100",
                "MC0101", "MC0202", "MC0303", "MC0404", "PR001", "PR002", "PR003",
                "PR004", "PR005", "PR006", "PR007", "PR008", "PR009", "PR010",
                "SA001", "TA001", "TA002", "TA003", "TA004", "TA005", "TC001",
                "TC002", "TL001", "TL002", "TL003", "TL004", "TL005", "TL006",
                "TL007", "TL008", "TL009", "TL010", "TL011", "TL012", "TL013",
                "TL014", "TL015", "TL016", "TL017", "TL018", "TL019", "TL020",
                "TL021", "TL022", "TP001", "TP002", "TP003", "TP004", "TP005",
                "TP006", "TP007", "TP008", "TP009", "TP010",
            ),
            "solicitante": (
                "ALENA",
                "ALEN AR",
                "ALESSANDRA",
                "ALESSANDRO",
                "ALISSON",
                "ALDO",
                "AMIA",
                "ANABEL",
                "ANACLE",
                "ANASIA",
                "ANTONIA",
                "ANTONIO",
                "ANTÔNIO",
                "ARA",
                "CLARICE",
                "CLICB ROGER",
                "CRISTIANE",
                "DANIEL",
                "DIANA",
                "DIOGO",
                "DINA",
                "EDMAR",
                "EDVAN DO",
                "EFFERSON",
                "ELANIR",
                "ELIENEIR",
                "ELANA",
                "ELANE",
                "ELISSON",
                "ELTON",
                "ELLINE",
                "EMERSON",
                "ERIC MORAIS",
                "ESTEFANY",
                "EUCILENE",
                "EVALDO",
                "FABIANA",
                "FABIANA CARVALHO",
                "FABIANE",
                "FRANCISCA",
                "FRANCISCO",
                "FRANCISCO ARIAS",
                "FRANCISLDO",
                "IRGN",
                "IRINA",
                "ISABEL",
                "JACOB",
                "JAIRO",
                "JAKSON",
                "JAKSON N",
                "JANE",
                "JHON",
                "JOAINE",
                "JOANE",
                "JOÃO",
                "JOÃO V. CARDOSO",
                "JOCIANE",
                "JOHN",
                "JOHNY",
                "JONY",
                "JOSIANE",
                "JOYCE",
                "JUAREZ",
                "JUCILAINA",
                "JUCILAINA A",
                "JUCYLANA",
                "JUCYLAURA",
                "KARINA",
                "KARINA DIAS",
                "KETH",
                "KETTI",
                "KETILANE",
                "KELLEY",
                "LAURA",
                "LEANDRO",
                "LEIDYANNE",
                "LEIADRO",
                "LENDRO",
                "LENE",
                "LIZIANE",
                "LU",
                "LUCAS",
                "LUCAS DE SOUSA",
                "LUCIANE",
                "LUCIANA",
                "LUCIANE",
                "LUCIENE",
                "MAER",
                "MAIK",
                "MAIR",
                "MARCELA",
                "MARCÉLO",
                "MARCO A",
                "MARLENE",
                "MARCOS",
                "MARCOS VINÍUS",
                "MARIAN",
                "MARIA RODRIGUES",
                "MARILENE",
                "MARLINA",
                "MARINES",
                "MARIA",
                "MARLSON",
                "MARTA",
                "MARTHA",
                "MAYK",
                "MEL",
                "MIRIAM",
                "MONIQUE",
                "NICE",
                "NILCENÉIA",
                "NILSON",
                "NOIRMA",
                "OZAMIR",
                "OZANIR",
                "RAFAEL",
                "RAIMUNDO",
                "RAIMUNDO C",
                "RAIMUNDO ASTRO",
                "RAUNDA",
                "RICK",
                "RICKNER",
                "ROLDANNA",
                "RONALDO",
                "RONÉLIO MARINHO",
                "ROSE",
                "ROSEANE",
                "ROSELEIA",
                "ROSELY",
                "ROSILEIA",
                "ROZIANE",
                "ROTH",
                "SANDRA",
                "SHEILA",
                "SONY",
                "TATIHSSA",
                "TALISSA",
                "TALISSA",
                "TANIA",
                "TATIANE",
                "THIAGO",
                "TIAGO DANTAS",
                "WANDER",
                "WEMERSON",
                "WESLEY",
                "WILSON",
            ), "tipo_servico": ("CORRETIVA", "PREVENTIVA", "PREDITIVA"),
            "prioridade": ("URGENTE", "ALTA", "NORMAL", "BAIXA"), "especialidade": ("MECÂNICA", "ELÉTRICA", "INSTRUMENTAÇÃO"),
            "tecnico": (
                "AMAURY SILVA",
                "ANDRÉ RICARDO",
                "CLEBER ROGER",
                "ERIC MORAIS",
                "GABRIEL",
                "GABRIEL GOMES",
                "JOÃO V. CARDOSO",
                "JOVERLEY BATALHA",
                "MARLISSON ALVES",
                "OZAMIR",
                "PAULO A. CORREA",
                "RONÉLIO MARINHO",
                "WILLIAN BRAZ",
            ), "turno": ("A", "B", "C"),
            "situacao": ("ABERTA", "EM ATENDIMENTO", "CONCLUÍDA", "ENCERRADA"),
        }
        for label, key, kind in fields:
            row, col, span = positions[key]
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, columnspan=span, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
            if kind == "combo":
                widget = ttk.Combobox(cell, textvariable=self.vars[key], values=options[key], state="readonly",
                                      font=("Segoe UI", 8))
                if key == "tecnico":
                    widget.bind("<<ComboboxSelected>>", self._apply_turno_by_tecnico)
            else:
                widget = tk.Entry(cell, textvariable=self.vars[key], font=("Segoe UI", 8),
                                  relief="solid", bd=1)
            widget.pack(fill="x", ipady=2)

    def _apply_turno_by_tecnico(self, _event=None):
        tecnico = self.vars["tecnico"].get().strip().upper()
        turno = TECNICO_TURNO_MAP.get(tecnico)
        if turno:
            self.vars["turno"].set(turno)

    def _build_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "＋  Nova OS", self.new_os, "#ffffff", "#315a7c")
        self._action(actions, "▣  Inserir", self.insert_os, "#1678bf", "white")
        self._action(actions, "✓  Salvar edição", self.update_os, "#ffffff", "#315a7c")
        self._action(actions, "⇧  Carga TXT", self.import_os_txt, "#ffffff", "#315a7c")
        self._action(actions, "✕  Excluir", self.delete_os, "#ffffff", "#b23b3b")

    def _action(self, parent, text, command, bg, fg):
        tk.Button(parent, text=text, command=command, bg=bg, fg=fg, font=("Segoe UI", 8, "bold"),
                  relief="solid", bd=1, cursor="hand2", padx=16, pady=7, activebackground=bg).pack(side="left", padx=5)

    def _build_table(self, parent):
        search = tk.Frame(parent, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        search_content = tk.Frame(search, bg="#edf2f7")
        search_content.pack()
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.os_search = tk.StringVar()
        self.os_search.trace_add("write", self.filter_os_items)
        tk.Entry(search_content, textvariable=self.os_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        headings = ("SETOR", "Nº EQUIPAMENTO", "SOLICITANTE", "HORA PARADA", "TIPO SERVIÇO", "PRIORIDADE",
                    "ESPECIALIDADE", "DESCRIÇÃO", "TÉCNICO", "TURNO", "HORA INÍCIO", "HORA FINAL",
                    "TEMPO PARADA", "TEMPO SERVIÇO", "TEMPO DE RESPOSTA", "SITUAÇÃO")
        widths = (95, 115, 110, 125, 110, 95, 115, 250, 130, 60, 120, 120, 110, 110, 130, 120)
        self.grid = ttk.Treeview(frame, columns=self.COLUMNS, show="headings", selectmode="browse")
        for col, heading, width in zip(self.COLUMNS, headings, widths):
            self.grid.heading(col, text=heading)
            self.grid.column(col, width=width, minwidth=55, stretch=(col == "descricao"))
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.grid.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=self.grid.xview)
        self.grid.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.grid.bind("<<TreeviewSelect>>", self._select_row)

    def _load_os_items(self):
        """Carrega as Ordens de Serviço gravadas no banco SQLite."""
        for item in self.grid.get_children():
            self.grid.delete(item)
        columns = ", ".join(self.COLUMNS)
        term = self.os_search.get().strip() if hasattr(self, "os_search") else ""
        query = f"SELECT id, {columns} FROM ordens_servico"
        params = ()
        if term:
            condition = " OR ".join(f"{column} LIKE ?" for column in self.COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.COLUMNS)
        query += " ORDER BY id DESC"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            self.grid.insert("", "end", iid=str(row_id), values=values)

    def filter_os_items(self, *_args):
        if hasattr(self, "grid"):
            self._load_os_items()

    def new_os(self):
        for var in self.vars.values():
            var.set("")
        self.vars["hora_parada"].set(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.vars["tipo_servico"].set("CORRETIVA")
        self.vars["prioridade"].set("URGENTE")
        self.vars["especialidade"].set("MECÂNICA")
        self.vars["situacao"].set("ABERTA")

    def insert_os(self):
        required = ("setor", "numero_equipamento", "solicitante", "descricao", "tecnico")
        if any(not self.vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha setor, nº equipamento, solicitante, descrição e técnico.")
            return
        values = tuple(self.vars[key].get().strip() for key in self.COLUMNS)
        columns = ", ".join(self.COLUMNS)
        placeholders = ", ".join("?" for _ in self.COLUMNS)
        with sqlite3.connect(self.database_path) as connection:
            connection.execute(f"INSERT INTO ordens_servico ({columns}) VALUES ({placeholders})", values)
        self._load_os_items()
        self.new_os()

    def _read_os_txt_rows(self, file_path):
        """Lê Ordens de Serviço de um TXT separado por ;, tab, | ou vírgula."""
        content = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                content = Path(file_path).read_text(encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            raise ValueError("Não foi possível identificar a codificação do arquivo TXT.")

        lines = [line for line in content.splitlines() if line.strip()]
        if not lines:
            raise ValueError("O arquivo TXT está vazio.")
        try:
            delimiter = csv.Sniffer().sniff("\n".join(lines[:10]), delimiters=";\t|,").delimiter
        except csv.Error:
            delimiter = ";"
        rows = list(csv.reader(lines, delimiter=delimiter))
        headers = [self._normalize_header(value) for value in rows[0]]
        aliases = {
            "SETOR": "setor", "NEQUIPAMENTO": "numero_equipamento", "NOEQUIPAMENTO": "numero_equipamento",
            "NUMEROEQUIPAMENTO": "numero_equipamento", "NUMERODOEQUIPAMENTO": "numero_equipamento",
            "SOLICITANTE": "solicitante", "HORAPARADA": "hora_parada",
            "TIPOSERVICO": "tipo_servico", "PRIORIDADE": "prioridade",
            "ESPECIALIDADE": "especialidade", "DESCRICAO": "descricao",
            "TECNICO": "tecnico", "TURNO": "turno", "HORAINICIO": "hora_inicio",
            "HORAFINAL": "hora_final", "TEMPOPARADA": "tempo_parada",
            "TEMPOSERVICO": "tempo_servico", "TEMPODERESPOSTA": "tempo_resposta",
            "TEMPORESPOSTA": "tempo_resposta", "SITUACAO": "situacao",
        }
        indexes = {aliases[header]: position for position, header in enumerate(headers) if header in aliases}
        required = ("setor", "numero_equipamento", "solicitante", "descricao", "tecnico")
        has_header = all(field in indexes for field in required)
        data_rows = rows[1:] if has_header else rows
        loaded_rows, skipped = [], 0
        for row in data_rows:
            if has_header:
                values = tuple(row[indexes[field]].strip() if indexes.get(field, -1) < len(row) else ""
                               for field in self.COLUMNS)
            else:
                values = tuple(value.strip() for value in row[:len(self.COLUMNS)])
                values += ("",) * (len(self.COLUMNS) - len(values))
            if not any(values):
                continue
            if any(not values[self.COLUMNS.index(field)] for field in required):
                skipped += 1
                continue
            loaded_rows.append(values)
        return loaded_rows, skipped

    def import_os_txt(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT de Ordens de Serviço",
            filetypes=(("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")),
        )
        if not file_path:
            return
        try:
            rows, skipped = self._read_os_txt_rows(file_path)
            if not rows:
                messagebox.showwarning("Carga TXT", "Nenhuma Ordem de Serviço válida foi encontrada no arquivo.")
                return
            columns = ", ".join(self.COLUMNS)
            placeholders = ", ".join("?" for _ in self.COLUMNS)
            with sqlite3.connect(self.database_path) as connection:
                connection.executemany(
                    f"INSERT INTO ordens_servico ({columns}) VALUES ({placeholders})", rows
                )
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_os_items()
        self.new_os()
        message = f"{len(rows)} Ordem(ns) de Serviço carregada(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter os campos obrigatórios."
        messagebox.showinfo("Carga TXT concluída", message)

    def delete_os(self):
        selected = self.grid.selection()
        if not selected:
            messagebox.showinfo("Excluir OS", "Selecione uma Ordem de Serviço na lista.")
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM ordens_servico WHERE id = ?", (int(selected[0]),))
        self._load_os_items()
        self.new_os()

    def update_os(self):
        selected = self.grid.selection()
        if not selected:
            messagebox.showinfo("Salvar edição", "Selecione uma Ordem de Serviço para editar.")
            return
        required = ("setor", "numero_equipamento", "solicitante", "descricao", "tecnico")
        if any(not self.vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha setor, nº equipamento, solicitante, descrição e técnico.")
            return
        values = tuple(self.vars[key].get().strip() for key in self.COLUMNS)
        assignments = ", ".join(f"{column} = ?" for column in self.COLUMNS)
        with sqlite3.connect(self.database_path) as connection:
            connection.execute(f"UPDATE ordens_servico SET {assignments} WHERE id = ?",
                               (*values, int(selected[0])))
        self._load_os_items()
        self.new_os()

    def _select_row(self, _event):
        selected = self.grid.selection()
        if not selected:
            return
        values = self.grid.item(selected[0], "values")
        for key, value in zip(self.COLUMNS, values):
            self.vars[key].set(value)

    def _show_placeholder(self):
        messagebox.showinfo("GestorMan", "Módulo inicial em desenvolvimento.")


if __name__ == "__main__":
    GestorMan().mainloop()
