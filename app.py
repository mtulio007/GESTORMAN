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
    "origem", "codigo", "descricao", "descricao_ingles", "estoque_minimo",
    "qtd_atual", "data_inventario", "responsavel", "data_pedido",
    "data_entrega", "lead_time", "custo_unitario", "custo_ttl",
)
OS_COLUMNS = (
    "setor", "numero_equipamento", "solicitante", "hora_parada", "tipo_servico",
    "prioridade", "especialidade", "descricao", "tecnico", "turno", "hora_inicio",
    "hora_final", "tempo_parada", "tempo_servico", "tempo_resposta", "situacao",
)


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
                estoque_minimo TEXT,
                qtd_atual TEXT,
                data_inventario TEXT,
                responsavel TEXT,
                data_pedido TEXT,
                data_entrega TEXT,
                lead_time TEXT,
                custo_unitario TEXT,
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
        value_for("descricao_ingles"), value_for("estoque_minimo"),
        value_for("qtd_atual", "qtd"), value_for("data_inventario"),
        value_for("responsavel"), value_for("data_pedido"), value_for("data_entrega"),
        value_for("lead_time", "dias_entrega"), value_for("custo_unitario", "custo"),
        value_for("custo_ttl"),
    )
    connection.execute("ALTER TABLE almoxarifado RENAME TO almoxarifado_anterior")
    connection.execute("""
        CREATE TABLE almoxarifado (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            origem TEXT NOT NULL,
            codigo TEXT NOT NULL,
            descricao TEXT NOT NULL,
            descricao_ingles TEXT,
            estoque_minimo TEXT,
            qtd_atual TEXT,
            data_inventario TEXT,
            responsavel TEXT,
            data_pedido TEXT,
            data_entrega TEXT,
            lead_time TEXT,
            custo_unitario TEXT,
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
        self._menu_subitem("Consulta de OS", self.show_os)
        self._menu_subitem("Almoxarifado", self.show_almoxarifado)
        self._menu_header("Pedidos", False)
        self._menu_header("Hora Extras", False)
        self._menu_header("Pedidos", False)

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
                 font=("Segoe UI", 8, "bold"), padx=18, pady=8).pack(fill="x")

    def _menu_subitem(self, text, command):
        tk.Button(self.menu_area, text=text, command=command, anchor="w", bd=0,
                  relief="flat", bg="#1f3448", activebackground="#2e4b64",
                  activeforeground="white", fg="#b9c9d7", font=("Segoe UI", 7),
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
            ("Descrição Inglês", "descricao_ingles"), ("Estoque Mínimo", "estoque_minimo"),
            ("Qtd Atual", "qtd_atual"), ("Data Inventário", "data_inventario"),
            ("Responsável", "responsavel"), ("Data Pedido", "data_pedido"),
            ("Data Entrega", "data_entrega"), ("Lead Time", "lead_time"),
            ("Custo Unitário", "custo_unitario"), ("Custo TTL", "custo_ttl"),
        ]
        self.almox_vars = {key: tk.StringVar() for _, key in fields}
        for idx, (label, key) in enumerate(fields):
            row, col = divmod(idx, 5)
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
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
        headings = ("ORIGEM", "CÓDIGO", "DESCRIÇÃO", "DESCRIÇÃO INGLÊS", "ESTOQUE MÍNIMO",
                    "QTD ATUAL", "DATA INVENTÁRIO", "RESPONSÁVEL", "DATA PEDIDO",
                    "DATA ENTREGA", "LEAD TIME", "CUSTO UNITÁRIO", "CUSTO TTL")
        widths = (120, 100, 230, 210, 105, 80, 110, 130, 105, 105, 85, 110, 100)
        self.almox_grid = ttk.Treeview(frame, columns=self.ALMOX_COLUMNS, show="headings", selectmode="browse")
        for col, heading, width in zip(self.ALMOX_COLUMNS, headings, widths):
            self.almox_grid.heading(col, text=heading)
            self.almox_grid.column(col, width=width, minwidth=65, stretch=(col in ("descricao", "descricao_ingles")))
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
        """Consulta no SQLite os itens que atendem ao texto pesquisado."""
        for item in self.almox_grid.get_children():
            self.almox_grid.delete(item)
        term = self.almox_search.get().strip() if hasattr(self, "almox_search") else ""
        columns = ", ".join(self.ALMOX_COLUMNS)
        query = f"SELECT id, {columns} FROM almoxarifado"
        params = ()
        if term:
            condition = " OR ".join(f"{column} LIKE ?" for column in self.ALMOX_COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.ALMOX_COLUMNS)
        query += " ORDER BY origem, codigo"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            self.almox_grid.insert("", "end", iid=str(row_id), values=values)

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
        values = tuple(self.almox_vars[key].get().strip() for key in self.ALMOX_COLUMNS)
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(self.ALMOX_COLUMNS)
                placeholders = ", ".join("?" for _ in self.ALMOX_COLUMNS)
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
        values = tuple(self.almox_vars[key].get().strip() for key in self.ALMOX_COLUMNS)
        assignments = ", ".join(f"{column} = ?" for column in self.ALMOX_COLUMNS)
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
            "ESTOQUEMINIMO": "estoque_minimo", "QTDATUAL": "qtd_atual",
            "QTD": "qtd_atual", "QUANTIDADE": "qtd_atual", "DATAINVENTARIO": "data_inventario",
            "RESPONSAVEL": "responsavel", "DATAPEDIDO": "data_pedido",
            "DATAENTREGA": "data_entrega", "LEADTIME": "lead_time",
            "DIASENTREGA": "lead_time", "CUSTOUNITARIO": "custo_unitario",
            "CUSTO": "custo_unitario", "CUSTOTTL": "custo_ttl", "CUSTOTOTAL": "custo_ttl",
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

    def _build_form(self, parent):
        box = tk.LabelFrame(parent, text="  Informações Gerais da Ordem de Serviço  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        for col in range(6):
            box.grid_columnconfigure(col, weight=1)

        fields = [
            ("Setor", "setor", "combo"), ("Nº Equipamento", "numero_equipamento", "entry"),
            ("Solicitante", "solicitante", "combo"), ("Hora Parada", "hora_parada", "entry"),
            ("Tipo Serviço", "tipo_servico", "combo"), ("Prioridade", "prioridade", "combo"),
            ("Especialidade", "especialidade", "combo"), ("Descrição", "descricao", "entry"),
            ("Técnico", "tecnico", "combo"), ("Turno", "turno", "combo"),
            ("Hora Início", "hora_inicio", "entry"), ("Hora Final", "hora_final", "entry"),
            ("Tempo Parada", "tempo_parada", "entry"), ("Tempo Serviço", "tempo_servico", "entry"),
            ("Tempo de Resposta", "tempo_resposta", "entry"), ("Situação", "situacao", "combo"),
        ]
        self.vars = {key: tk.StringVar() for _, key, _ in fields}
        defaults = {"tipo_servico": "CORRETIVA",
                    "prioridade": "URGENTE", "especialidade": "MECÂNICA",
                    "situacao": "ABERTA"}
        for key, value in defaults.items():
            self.vars[key].set(value)

        options = {
            "setor": ("TEC.LEVE", "TEC.PESADA", "PRODUÇÃO"),
            "solicitante": ("WILLIAN", "CARLOS", "MARIA"), "tipo_servico": ("CORRETIVA", "PREVENTIVA", "PREDITIVA"),
            "prioridade": ("URGENTE", "ALTA", "NORMAL", "BAIXA"), "especialidade": ("MECÂNICA", "ELÉTRICA", "INSTRUMENTAÇÃO"),
            "tecnico": ("GABRIEL GOMES", "JOÃO SILVA", "ANA LIMA"), "turno": ("A", "B", "C"),
            "situacao": ("ABERTA", "EM ATENDIMENTO", "CONCLUÍDA", "ENCERRADA"),
        }
        for idx, (label, key, kind) in enumerate(fields):
            row, col = divmod(idx, 6)
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 7)).pack(fill="x")
            if kind == "combo":
                widget = ttk.Combobox(cell, textvariable=self.vars[key], values=options[key], state="readonly",
                                      font=("Segoe UI", 8))
            else:
                widget = tk.Entry(cell, textvariable=self.vars[key], font=("Segoe UI", 8),
                                  relief="solid", bd=1)
            widget.pack(fill="x", ipady=2)

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
