import os
import csv
import hashlib
import sqlite3
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from datetime import datetime, timedelta


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
CODIGOS_COLUMNS = ("codigo", "descricao", "descricao_ingles")
DATETIME_FORMAT = "%d/%m/%Y %H:%M:%S"
DATETIME_FIELDS = ("hora_inicio", "hora_final")

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
    CODIGOS_COLUMNS = CODIGOS_COLUMNS

    def __init__(self):
        super().__init__()
        self.database_path = Path(__file__).with_name("gestorman.db")
        self._init_database()
        self.title("GestorMan | Gestão de Manutenção v.001")
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
            connection.execute("""
                CREATE TABLE IF NOT EXISTS cadastro_codigos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT NOT NULL UNIQUE,
                    descricao TEXT NOT NULL,
                    descricao_ingles TEXT
                )
            """)
            connection.execute("""
                CREATE TABLE IF NOT EXISTS acompanhamento_pedidos (
                    codigo TEXT PRIMARY KEY,
                    lead_time REAL,
                    media_consumo REAL,
                    ponto_pedido REAL,
                    ressuprimento TEXT,
                    status TEXT NOT NULL DEFAULT 'NORMAL'
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
        tk.Label(sidebar, text="GESTÃO DE ESTOQUE", bg="#193044", fg="#9eb5c8",
                 anchor="w", font=("Segoe UI", 7), padx=23).pack(fill="x", pady=(0, 14))

        self.menu_area = tk.Frame(sidebar, bg="#1f3448")
        self.menu_area.pack(fill="both", expand=True, pady=4)
        tk.Label(sidebar, text="M@rcoSoft Alrights Reserved.", bg="#1f3448", fg="white",
                 anchor="center", font=("Segoe UI", 7)).pack(side="bottom", fill="x", pady=(8, 14))
        self.menu_sections = {}
        almoxarifado_menu = self._menu_header("Almoxarifado", "almoxarifado", expanded=False)
        self._menu_subitem(almoxarifado_menu, "Cadastro de Códigos", self.show_codigos)
        self._menu_subitem(almoxarifado_menu, "Recebimento de Itens", self.show_almox_recebimento)
        self._menu_subitem(almoxarifado_menu, "Saída de Itens", self.show_almox_saida)
        self._menu_subitem(almoxarifado_menu, "Resumo", self.show_almox_resumo)
        self._menu_subitem(almoxarifado_menu, "Acompanhamento de Pedidos", self.show_acompanhamento_pedidos)
        manutencao_menu = self._menu_header("Manutenção", "manutencao", expanded=False)
        self._menu_subitem(manutencao_menu, "Ordens de Serviços", self.show_os)

        self.main = tk.Frame(self, bg="#edf2f7")
        self.main.pack(side="left", fill="both", expand=True)
        self.show_codigos()

    def _menu_item(self, icon, text, command):
        tk.Button(self.menu_area, text=f" {icon}   {text}", command=command, anchor="w",
                  bd=0, relief="flat", bg="#1f3448", activebackground="#2e4b64",
                  activeforeground="white", fg="#d5e3ef", font=("Segoe UI", 9),
                  padx=17, pady=8).pack(fill="x")

    def _menu_header(self, text, key, expanded=True):
        """Cria uma seção retrátil da barra lateral e retorna o contêiner de seus itens."""
        section = tk.Frame(self.menu_area, bg="#1f3448")
        section.pack(fill="x")
        header = tk.Button(section, anchor="w", bd=0, relief="flat", bg="#1f3448",
                           activebackground="#2e4b64", activeforeground="white", fg="white",
                           font=("Segoe UI", 10, "bold"), padx=18, pady=8, cursor="hand2",
                           command=lambda: self._toggle_menu_section(key))
        header.pack(fill="x")
        submenu = tk.Frame(section, bg="#1f3448")
        if expanded:
            submenu.pack(fill="x")
        self.menu_sections[key] = {"text": text, "header": header, "submenu": submenu, "expanded": expanded}
        self._update_menu_header(key)
        return submenu

    def _toggle_menu_section(self, key):
        section = self.menu_sections[key]
        section["expanded"] = not section["expanded"]
        if section["expanded"]:
            section["submenu"].pack(fill="x")
        else:
            section["submenu"].pack_forget()
        self._update_menu_header(key)

    def _update_menu_header(self, key):
        section = self.menu_sections[key]
        mark = "⌄" if section["expanded"] else "›"
        section["header"].configure(text=f" {mark}   {section['text']}")

    def _menu_subitem(self, parent, text, command):
        tk.Button(parent, text=text, command=command, anchor="w", bd=0,
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

    def _build_record_counter(self, parent):
        """Mostra, no rodapé da tela, o total atualmente exibido na grade."""
        self.record_counter = tk.StringVar(value="Registros cadastrados: 0")
        footer = tk.Frame(parent, bg="#edf2f7")
        footer.pack(side="bottom", fill="x", pady=(8, 0))
        tk.Label(footer, textvariable=self.record_counter, bg="#edf2f7", fg="#64748b",
                 font=("Segoe UI", 8)).pack(side="right")

    def _set_record_counter(self, total):
        if hasattr(self, "record_counter"):
            self.record_counter.set(f"Registros cadastrados: {total}")

    @staticmethod
    def _auto_fit_tree_columns(grid, columns):
        """Ajusta as colunas ao maior conteúdo visível, incluindo o cabeçalho."""
        measure = tkfont.Font(font=("Segoe UI", 8))
        limits = {"descricao": 420, "aplicacao": 280, "requisitante": 220, "fornecedor": 220}
        for column in columns:
            heading = str(grid.heading(column, "text")).replace("\n", " ")
            width = measure.measure(heading) + 28
            for item_id in grid.get_children():
                value = str(grid.set(item_id, column))
                width = max(width, measure.measure(value) + 24)
            grid.column(column, width=max(65, min(width, limits.get(column, 180))))

    def _confirm_txt_import(self, table, screen_name):
        """Pergunta como tratar dados já cadastrados antes de uma nova carga."""
        with sqlite3.connect(self.database_path) as connection:
            total = connection.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        if not total:
            return "append"
        choice = messagebox.askyesnocancel(
            "Dados já cadastrados",
            f"A tela de {screen_name} já possui {total} registro(s).\n\n"
            "Sim: limpar os dados existentes e carregar o TXT.\n"
            "Não: manter os dados e carregar os novos registros após os existentes.\n"
            "Cancelar: não realizar a carga.",
            icon="warning",
        )
        if choice is None:
            return None
        return "replace" if choice else "append"

    def show_os(self):
        self._clear_main()
        body = self._page_header("Cadastro de Ordem de Serviço", "Manutenção  /  Ordem de Serviços")
        self._build_form(body)
        self._build_actions(body)
        self._build_table(body)
        self._build_record_counter(body)
        self._load_os_items()
        self.new_os()

    def show_codigos(self):
        """Exibe o cadastro mestre de códigos de materiais."""
        self._clear_main()
        body = self._page_header("Cadastro de Códigos", "Cadastros  /  Códigos")
        self._build_codigos_form(body)
        self._build_codigos_actions(body)
        self._build_codigos_table(body)
        self._build_record_counter(body)
        self._load_codigos()
        self.new_codigo()

    def _build_codigos_form(self, parent):
        box = tk.LabelFrame(parent, text="  Informações do Código  ",
                            bg="#eef4fa", fg="#28435e", font=("Segoe UI", 9, "bold"),
                            padx=14, pady=12, bd=1, relief="groove")
        box.pack(fill="x")
        box.grid_columnconfigure(1, weight=1)
        box.grid_columnconfigure(2, weight=1)
        self.codigo_vars = {key: tk.StringVar() for key in self.CODIGOS_COLUMNS}
        fields = (("Código", "codigo"), ("Descrição", "descricao"),
                  ("Descrição Inglês", "descricao_ingles"))
        for column, (label, key) in enumerate(fields):
            cell = tk.Frame(box, bg="#eef4fa")
            cell.grid(row=0, column=column, sticky="ew", padx=4, pady=4)
            tk.Label(cell, text=label, bg="#eef4fa", fg="#38546e", anchor="w",
                     font=("Segoe UI", 8)).pack(fill="x")
            tk.Entry(cell, textvariable=self.codigo_vars[key], font=("Segoe UI", 9),
                     relief="solid", bd=1).pack(fill="x", ipady=3)

    def _build_codigos_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "＋  Novo Código", self.new_codigo, "#ffffff", "#315a7c")
        self._action(actions, "▣  Inserir", self.insert_codigo, "#1678bf", "white")
        self._action(actions, "✓  Salvar edição", self.update_codigo, "#ffffff", "#315a7c")
        self._action(actions, "⇧  Carga TXT", self.import_codigos_txt, "#ffffff", "#315a7c")
        self._action(actions, "✕  Excluir", self.delete_codigo, "#ffffff", "#b23b3b")

    def _build_codigos_table(self, parent):
        search = tk.Frame(parent, bg="#edf2f7")
        search.pack(fill="x", pady=(0, 8))
        content = tk.Frame(search, bg="#edf2f7")
        content.pack()
        tk.Label(content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.codigo_search = tk.StringVar()
        self.codigo_search.trace_add("write", self.filter_codigos)
        tk.Entry(content, textvariable=self.codigo_search, width=42, font=("Segoe UI", 9),
                 relief="solid", bd=1).pack(side="left", ipady=3)

        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        self.codigos_grid = ttk.Treeview(frame, columns=self.CODIGOS_COLUMNS,
                                         show="headings", selectmode="browse")
        for key, title, width in (("codigo", "CÓDIGO", 180), ("descricao", "DESCRIÇÃO", 450),
                                  ("descricao_ingles", "DESCRIÇÃO INGLÊS", 450)):
            self.codigos_grid.heading(key, text=title)
            self.codigos_grid.column(key, width=width, minwidth=120, anchor="w", stretch=True)
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.codigos_grid.yview)
        self.codigos_grid.configure(yscrollcommand=scroll_y.set)
        self.codigos_grid.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.codigos_grid.bind("<<TreeviewSelect>>", self._select_codigo)

    def _load_codigos(self):
        for item in self.codigos_grid.get_children():
            self.codigos_grid.delete(item)
        term = self.codigo_search.get().strip()
        query = "SELECT id, codigo, descricao, descricao_ingles FROM cadastro_codigos"
        params = ()
        if term:
            query += " WHERE codigo LIKE ? OR descricao LIKE ? OR descricao_ingles LIKE ?"
            params = (f"%{term}%",) * 3
        query += " ORDER BY codigo"
        with sqlite3.connect(self.database_path) as connection:
            for row_id, *values in connection.execute(query, params):
                self.codigos_grid.insert("", "end", iid=str(row_id), values=values)
        self._set_record_counter(len(self.codigos_grid.get_children()))

    def filter_codigos(self, *_args):
        if hasattr(self, "codigos_grid"):
            self._load_codigos()

    def new_codigo(self):
        for var in self.codigo_vars.values():
            var.set("")
        if hasattr(self, "codigos_grid"):
            self.codigos_grid.selection_remove(self.codigos_grid.selection())

    def _codigo_values(self):
        values = tuple(self.codigo_vars[key].get().strip() for key in self.CODIGOS_COLUMNS)
        if not values[0] or not values[1]:
            messagebox.showwarning("Campos obrigatórios", "Preencha código e descrição.")
            return None
        return values

    def insert_codigo(self):
        values = self._codigo_values()
        if not values:
            return
        try:
            with sqlite3.connect(self.database_path) as connection:
                connection.execute("INSERT INTO cadastro_codigos (codigo, descricao, descricao_ingles) VALUES (?, ?, ?)", values)
        except sqlite3.IntegrityError:
            messagebox.showwarning("Código duplicado", "Já existe um cadastro com este código.")
            return
        self._load_codigos()
        self.new_codigo()

    def update_codigo(self):
        selected = self.codigos_grid.selection()
        if not selected:
            messagebox.showinfo("Salvar edição", "Selecione um código para editar.")
            return
        values = self._codigo_values()
        if not values:
            return
        try:
            with sqlite3.connect(self.database_path) as connection:
                connection.execute("UPDATE cadastro_codigos SET codigo = ?, descricao = ?, descricao_ingles = ? WHERE id = ?",
                                   (*values, int(selected[0])))
        except sqlite3.IntegrityError:
            messagebox.showwarning("Código duplicado", "Já existe um cadastro com este código.")
            return
        self._load_codigos()
        self.new_codigo()

    def delete_codigo(self):
        selected = self.codigos_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir código", "Selecione um código na lista.")
            return
        if not messagebox.askyesno("Excluir código", "Deseja mesmo excluir o registro?"):
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM cadastro_codigos WHERE id = ?", (int(selected[0]),))
        self._load_codigos()
        self.new_codigo()

    def _select_codigo(self, _event):
        selected = self.codigos_grid.selection()
        if selected:
            for key, value in zip(self.CODIGOS_COLUMNS, self.codigos_grid.item(selected[0], "values")):
                self.codigo_vars[key].set(value)

    def import_codigos_txt(self):
        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT de códigos",
            filetypes=(("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")),
        )
        if not file_path:
            return
        try:
            rows, skipped = self._read_codigos_txt_rows(file_path)
            if not rows:
                messagebox.showwarning("Carga TXT", "Nenhum código válido foi encontrado no arquivo.")
                return
            rows = tuple((codigo or self._codigo_provisorio(descricao), descricao, descricao_ingles)
                         for codigo, descricao, descricao_ingles in rows)
            mode = self._confirm_txt_import("cadastro_codigos", "Cadastro de Códigos")
            if mode is None:
                return
            with sqlite3.connect(self.database_path) as connection:
                if mode == "replace":
                    connection.execute("DELETE FROM cadastro_codigos")
                connection.executemany(
                    "INSERT INTO cadastro_codigos (codigo, descricao, descricao_ingles) VALUES (?, ?, ?) "
                    "ON CONFLICT(codigo) DO UPDATE SET descricao = excluded.descricao, descricao_ingles = excluded.descricao_ingles",
                    rows,
                )
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_codigos()
        message = f"{len(rows)} código(s) carregado(s) ou atualizado(s)."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter código e descrição."
        messagebox.showinfo("Carga TXT concluída", message)

    def _read_codigos_txt_rows(self, file_path):
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
            delimiter = "\t"
        rows = list(csv.reader(lines, delimiter=delimiter))
        headers = [self._normalize_header(value) for value in rows[0]]
        aliases = {"CODIGO": "codigo", "DESCRICAO": "descricao", "DESCRICAOINGLES": "descricao_ingles",
                   "DESCRIPTIONENGLISH": "descricao_ingles"}
        indexes = {aliases[header]: position for position, header in enumerate(headers) if header in aliases}
        has_header = "descricao" in indexes
        loaded, skipped = [], 0
        for row in (rows[1:] if has_header else rows):
            values = tuple(row[indexes[key]].strip() if has_header and indexes.get(key, -1) < len(row)
                           else (row[position].strip() if position < len(row) else "")
                           for position, key in enumerate(self.CODIGOS_COLUMNS))
            if not any(values):
                continue
            if not values[1]:
                skipped += 1
                continue
            loaded.append(values)
        return loaded, skipped

    def show_almoxarifado(self):
        self._clear_main()
        body = self._page_header("Cadastro de Almoxarifado", "Cadastros  /  Almoxarifado")
        self._build_almox_form(body)
        self._build_almox_actions(body)
        self._build_almox_table(body)
        self._build_record_counter(body)
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
            ("Média Consumo", "media_consumo"), ("Est. Inicial", "estoque_minimo"),
            ("Est. Final", "estoque_maximo"), ("Est. Segurança", "est_seguranca"),
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
                "MÉDIA CONSUMO", "EST. INICIAL", "EST. FINAL", "EST. SEGURANÇA",
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
                COALESCE((
                    SELECT SUM(CAST(REPLACE(REPLACE(TRIM(r.qtd), '.', ''), ',', '.') AS REAL))
                    FROM recebimento_almoxarifado r WHERE r.codigo = a.codigo
                ), 0) as entradas,
                COALESCE((
                    SELECT SUM(CAST(REPLACE(REPLACE(TRIM(s.qtd), '.', ''), ',', '.') AS REAL))
                    FROM saida_almoxarifado s WHERE s.codigo = a.codigo
                ), 0) as saida,
                (COALESCE((
                    SELECT SUM(CAST(REPLACE(REPLACE(TRIM(r.qtd), '.', ''), ',', '.') AS REAL))
                    FROM recebimento_almoxarifado r WHERE r.codigo = a.codigo
                ), 0) - COALESCE((
                    SELECT SUM(CAST(REPLACE(REPLACE(TRIM(s.qtd), '.', ''), ',', '.') AS REAL))
                    FROM saida_almoxarifado s WHERE s.codigo = a.codigo
                ), 0)) as saldo,
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
        """
        
        params = ()
        if term:
            condition = " OR ".join(f"a.{column} LIKE ?" for column in self.ALMOX_COLUMNS)
            query += f" WHERE {condition}"
            params = tuple(f"%{term}%" for _ in self.ALMOX_COLUMNS)
        
        query += " ORDER BY a.origem, a.codigo"
        
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
                elif idx in (11, 12, 13):  # est. inicial, final e segurança
                    formatted_values.append(str(int(self._pedido_numero(val))) if str(val or "").strip() else "")
                else:
                    formatted_values.append(str(val) if val else "")
            # Aplica tag "negative" se o saldo for negativo
            tags = ("negative",) if saldo_value is not None and saldo_value < 0 else ()
            self.almox_grid.insert("", "end", iid=str(row_id), values=formatted_values, tags=tags)
        self._set_record_counter(len(self.almox_grid.get_children()))

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
        if not messagebox.askyesno("Excluir item", "Deseja mesmo excluir o registro?"):
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

    @classmethod
    def _codigo_provisorio(cls, descricao):
        """Gera um código estável para que movimentações sem código possam ser conciliadas."""
        descricao_normalizada = cls._normalize_header(descricao)
        digest = hashlib.sha1(descricao_normalizada.encode("utf-8")).hexdigest()[:10].upper()
        return f"SEM-CODIGO-{digest}"

    def _preparar_codigos_importados(self, connection, rows, columns):
        """Preenche códigos ausentes a partir da descrição antes de gravar movimentações.

        A mesma descrição gera sempre o mesmo código técnico, permitindo que entradas e
        saídas importadas em arquivos diferentes sejam consideradas no mesmo saldo.
        """
        codigo_index = columns.index("codigo")
        descricao_index = columns.index("descricao")
        prepared, pendentes = [], 0
        for row in rows:
            values = list(row)
            codigo, descricao = values[codigo_index].strip(), values[descricao_index].strip()
            if not codigo:
                if descricao:
                    codigo = self._codigo_provisorio(descricao)
                    connection.execute(
                        "INSERT INTO cadastro_codigos (codigo, descricao) VALUES (?, ?) "
                        "ON CONFLICT(codigo) DO NOTHING", (codigo, descricao),
                    )
                else:
                    # Sem código e sem descrição não existe chave para conciliação.
                    # O identificador único conserva a linha para correção posterior.
                    digest = hashlib.sha1("|".join(values).encode("utf-8")).hexdigest()[:10].upper()
                    codigo = f"PENDENTE-{digest}"
                    pendentes += 1
                values[codigo_index] = codigo
            prepared.append(tuple(values))
        return prepared, pendentes

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
        has_header = all(field in indexes for field in ("origem", "descricao"))
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
            if not values[0].strip():
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
            mode = self._confirm_txt_import("almoxarifado", "Cadastro de Almoxarifado")
            if mode is None:
                return
            columns = ", ".join(self.ALMOX_COLUMNS)
            placeholders = ", ".join("?" for _ in self.ALMOX_COLUMNS)
            updates = ", ".join(f"{column} = excluded.{column}" for column in self.ALMOX_COLUMNS[2:])
            query = (f"INSERT INTO almoxarifado ({columns}) VALUES ({placeholders}) "
                     f"ON CONFLICT(origem, codigo) DO UPDATE SET {updates}")
            with sqlite3.connect(self.database_path) as connection:
                if mode == "replace":
                    connection.execute("DELETE FROM almoxarifado")
                rows, pendentes = self._preparar_codigos_importados(connection, rows, self.ALMOX_COLUMNS)
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_almox_items()
        message = f"{len(rows)} item(ns) carregado(s) ou atualizado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter origem."
        if pendentes:
            message += f"\n{pendentes} linha(s) sem código e descrição foi(ram) marcada(s) como pendente(s)."
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
        self._build_recebimento_actions(body)
        self._build_recebimento_table(body)
        self._build_record_counter(body)
        self._load_recebimento_items()

    def _build_recebimento_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "Salvar novos/edições", self.save_recebimento_items, "#ffffff", "#27835c")
        self._action(actions, "＋  Novo Recebimento", self.new_recebimento_item, "#ffffff", "#315a7c")
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
        self._new_recebimento_rows = {}
        self._edited_recebimento_rows = {}
        self._new_recebimento_sequence = 0
        self.recebimento_grid.bind("<Double-1>", self._edit_recebimento_cell)
        self.recebimento_grid.bind("<F2>", self._edit_selected_recebimento_cell)

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
        # Registros novos ficam sempre na primeira linha, logo abaixo do cabeçalho.
        query += " ORDER BY id DESC"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            values = self._edited_recebimento_rows.get(str(row_id), values)
            self.recebimento_grid.insert("", "end", iid=str(row_id), values=values)
        for item_id, values in getattr(self, "_new_recebimento_rows", {}).items():
            self.recebimento_grid.insert("", 0, iid=item_id, values=values, tags=("novo_recebimento",))
        self._auto_fit_tree_columns(self.recebimento_grid, self.RECEBIMENTO_COLUMNS)
        self._set_record_counter(len(self.recebimento_grid.get_children()))

    def filter_recebimento_items(self, *_args):
        if hasattr(self, "recebimento_grid"):
            self._load_recebimento_items()

    def new_recebimento_item(self):
        """Inclui uma linha editável no topo, preenchida com data e hora atuais."""
        self._new_recebimento_sequence += 1
        item_id = f"novo_recebimento_{self._new_recebimento_sequence}"
        values = [""] * len(self.RECEBIMENTO_COLUMNS)
        values[0] = datetime.now().strftime("%d/%m/%Y")
        self._new_recebimento_rows[item_id] = values
        self.recebimento_grid.insert("", 0, iid=item_id, values=values, tags=("novo_recebimento",))
        self.recebimento_grid.selection_set(item_id)
        self.recebimento_grid.focus(item_id)
        self.recebimento_grid.see(item_id)
        self._auto_fit_tree_columns(self.recebimento_grid, self.RECEBIMENTO_COLUMNS)
        self.after_idle(lambda: self._open_recebimento_cell_editor(item_id, "#2"))

    def save_recebimento_items(self):
        """Grava de uma vez as novas linhas e as edicoes pendentes."""
        new_rows = list(self._new_recebimento_rows.items())
        incomplete = [item_id for item_id, values in new_rows if not values[1] or not values[2]]
        if incomplete:
            messagebox.showwarning(
                "Salvar registros",
                "Preencha codigo e descricao em todos os novos recebimentos antes de salvar.",
            )
            self.recebimento_grid.selection_set(incomplete[0])
            self.recebimento_grid.focus(incomplete[0])
            self.recebimento_grid.see(incomplete[0])
            return
        if not new_rows and not self._edited_recebimento_rows:
            messagebox.showinfo("Salvar registros", "Nao ha registros novos ou alteracoes para salvar.")
            return
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(self.RECEBIMENTO_COLUMNS)
                placeholders = ", ".join("?" for _ in self.RECEBIMENTO_COLUMNS)
                connection.executemany(
                    f"INSERT INTO recebimento_almoxarifado ({columns}) VALUES ({placeholders})",
                    [values for _item_id, values in new_rows],
                )
                assignments = ", ".join(f"{column} = ?" for column in self.RECEBIMENTO_COLUMNS)
                connection.executemany(
                    f"UPDATE recebimento_almoxarifado SET {assignments} WHERE id = ?",
                    [tuple(values) + (int(item_id),) for item_id, values in self._edited_recebimento_rows.items()],
                )
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao salvar", f"Nao foi possivel salvar os registros.\n\n{error}")
            return
        saved_total = len(new_rows) + len(self._edited_recebimento_rows)
        self._new_recebimento_rows.clear()
        self._edited_recebimento_rows.clear()
        self._load_recebimento_items()
        messagebox.showinfo("Salvar registros", f"{saved_total} registro(s) salvo(s) com sucesso.")

    def delete_recebimento_item(self):
        selected = self.recebimento_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir item", "Selecione um item na lista.")
            return
        item_id = selected[0]
        if item_id in self._new_recebimento_rows:
            self._new_recebimento_rows.pop(item_id, None)
            self.recebimento_grid.delete(item_id)
            self._set_record_counter(len(self.recebimento_grid.get_children()))
            return
        if not messagebox.askyesno("Excluir item", "Deseja mesmo excluir o registro?"):
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM recebimento_almoxarifado WHERE id = ?", (int(item_id),))
        self._edited_recebimento_rows.pop(item_id, None)
        self._load_recebimento_items()

    def _edit_recebimento_cell(self, event):
        if self.recebimento_grid.identify_region(event.x, event.y) != "cell":
            return
        item_id = self.recebimento_grid.identify_row(event.y)
        column = self.recebimento_grid.identify_column(event.x)
        if item_id and column:
            self._open_recebimento_cell_editor(item_id, column)

    def _edit_selected_recebimento_cell(self, _event):
        selected = self.recebimento_grid.selection()
        if selected:
            self._open_recebimento_cell_editor(selected[0], "#1")
        return "break"

    def _open_recebimento_cell_editor(self, item_id, column):
        bbox = self.recebimento_grid.bbox(item_id, column)
        if not bbox:
            return
        index = int(column[1:]) - 1
        x, y, width, height = bbox
        editor = tk.Entry(self.recebimento_grid, font=("Segoe UI", 8))
        editor.insert(0, self.recebimento_grid.item(item_id, "values")[index])
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        editor.select_range(0, "end")
        editor.bind("<Return>", lambda _event: self._save_recebimento_cell_and_advance(item_id, index, editor))
        editor.bind("<FocusOut>", lambda _event: self._save_recebimento_cell(item_id, index, editor))
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def _save_recebimento_cell_and_advance(self, item_id, index, editor):
        """Salva a célula atual e leva o cursor à próxima coluna com Enter."""
        saved_id = self._save_recebimento_cell(item_id, index, editor)
        next_index = index + 1
        if saved_id and next_index < len(self.RECEBIMENTO_COLUMNS):
            self.after_idle(
                lambda: self._open_recebimento_cell_editor(saved_id, f"#{next_index + 1}")
            )
        return "break"

    def _save_recebimento_cell(self, item_id, index, editor):
        if not editor.winfo_exists():
            return
        value = editor.get().strip()
        editor.destroy()
        values = list(self.recebimento_grid.item(item_id, "values"))
        values[index] = value

        if item_id in self._new_recebimento_rows:
            self._new_recebimento_rows[item_id] = values
            self.recebimento_grid.item(item_id, values=values)
            self._auto_fit_tree_columns(self.recebimento_grid, self.RECEBIMENTO_COLUMNS)
            return item_id

        self._edited_recebimento_rows[item_id] = values
        self.recebimento_grid.item(item_id, values=values)
        self._auto_fit_tree_columns(self.recebimento_grid, self.RECEBIMENTO_COLUMNS)
        return item_id

    def show_almox_resumo(self):
        """Exibe o saldo consolidado de entradas e saídas por item."""
        self._clear_main()
        body = self._page_header("Resumo de Movimentação de Itens", "Cadastros  /  Almoxarifado  /  Resumo")

        summary_bar = tk.Frame(body, bg="#edf2f7")
        summary_bar.pack(fill="x", pady=(0, 10))
        search_content = tk.Frame(summary_bar, bg="#edf2f7")
        search_content.pack(side="left", anchor="w")
        tk.Label(search_content, text="Pesquisar", bg="#edf2f7", fg="#38546e",
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(2, 8))
        self.resumo_search = tk.StringVar()
        self.resumo_search.trace_add("write", self._load_resumo_items)
        tk.Entry(search_content, textvariable=self.resumo_search, width=34, font=("Segoe UI", 8),
                 relief="solid", bd=1).pack(side="left", ipady=3)
        tk.Button(search_content, text="Baixar Excel", command=self.export_resumo_excel,
                  bg="#1678bf", activebackground="#0f5f99", activeforeground="white",
                  fg="white", bd=0, font=("Segoe UI", 8, "bold"), cursor="hand2",
                  padx=12, pady=5).pack(side="left", padx=(10, 0))

        totals_content = tk.Frame(summary_bar, bg="#e4edf6", bd=1, relief="solid", padx=18, pady=7)
        totals_content.pack(side="right", anchor="e")
        self.resumo_total_vars = {
            "entradas": tk.StringVar(value="+0"),
            "saidas": tk.StringVar(value="+0"),
            "saldo": tk.StringVar(value="+0"),
        }
        for label, key, color in (("ENTRADAS", "entradas", "#166534"),
                                  ("SAÍDAS", "saidas", "#9a3412"),
                                  ("SALDO", "saldo", "#0f4c81")):
            item = tk.Frame(totals_content, bg="#e4edf6")
            item.pack(side="left", padx=14)
            tk.Label(item, text=label, bg="#e4edf6", fg="#64748b",
                     font=("Segoe UI", 7, "bold")).pack(anchor="e")
            tk.Label(item, textvariable=self.resumo_total_vars[key], bg="#e4edf6", fg=color,
                     font=("Segoe UI", 11, "bold")).pack(anchor="e")

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
        self._build_record_counter(body)
        self._load_resumo_items()

    def export_resumo_excel(self):
        """Exporta para Excel exatamente os dados atualmente exibidos no resumo."""
        if not hasattr(self, "resumo_grid"):
            return
        rows = [self.resumo_grid.item(item, "values") for item in self.resumo_grid.get_children()]
        if not rows:
            messagebox.showinfo("Baixar Excel", "N\u00e3o h\u00e1 dados para exportar.")
            return
        filename = f"resumo_movimentacao_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Salvar resumo em Excel", initialfile=filename,
            defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")],
        )
        if not file_path:
            return

        headings = ("C\u00d3DIGO", "DESCRI\u00c7\u00c3O", "ENTRADAS", "SA\u00cdDAS", "SALDO")
        total_values = (
            "TOTAIS", "",
            self.resumo_total_vars["entradas"].get(),
            self.resumo_total_vars["saidas"].get(),
            self.resumo_total_vars["saldo"].get(),
        )

        def cell(value, style=0):
            return f'<c t="inlineStr" s="{style}"><is><t>{escape(str(value))}</t></is></c>'

        sheet_rows = [
            f'<row r="1">{"".join(cell(value, 1) for value in headings)}</row>'
        ]
        for index, row in enumerate(rows, start=2):
            sheet_rows.append(f'<row r="{index}">{"".join(cell(value) for value in row)}</row>')
        sheet_rows.append(
            f'<row r="{len(rows) + 2}">{"".join(cell(value, 1) for value in total_values)}</row>'
        )
        worksheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="55" customWidth="1"/><col min="3" max="5" width="16" customWidth="1"/></cols>
<sheetData>""" + "".join(sheet_rows) + "</sheetData></worksheet>"
        styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"""
        try:
            with zipfile.ZipFile(file_path, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>""")
                archive.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""")
                archive.writestr("xl/workbook.xml", """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Resumo" sheetId="1" r:id="rId1"/></sheets></workbook>""")
                archive.writestr("xl/_rels/workbook.xml.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""")
                archive.writestr("xl/worksheets/sheet1.xml", worksheet)
                archive.writestr("xl/styles.xml", styles)
        except (OSError, zipfile.BadZipFile) as error:
            messagebox.showerror("Baixar Excel", f"N\u00e3o foi poss\u00edvel criar o arquivo.\n\n{error}")
            return
        messagebox.showinfo("Baixar Excel", "Arquivo Excel criado com sucesso.")

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
        query += " ORDER BY COALESCE(r.entradas, 0) - COALESCE(s.saidas, 0) DESC, c.codigo"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        total_entradas = total_saidas = total_saldo = 0
        for codigo, descricao, entradas, saidas, saldo in rows:
            def format_qtd(value, signed=False):
                value = float(value)
                formatted = str(int(value)) if value.is_integer() else f"{value:.2f}".replace(".", ",")
                return f"{value:+g}" if signed and value.is_integer() else (f"+{formatted}" if signed and value >= 0 else formatted)
            self.resumo_grid.insert("", "end", values=(
                codigo, descricao, format_qtd(entradas), format_qtd(saidas), format_qtd(saldo, signed=True)
            ), tags=("negative",) if saldo < 0 else ())
            total_entradas += float(entradas)
            total_saidas += float(saidas)
            total_saldo += float(saldo)
        if hasattr(self, "resumo_total_vars"):
            def format_total(value):
                formatted = f"{abs(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{'+' if value >= 0 else '-'}{formatted}"
            self.resumo_total_vars["entradas"].set(format_total(total_entradas))
            self.resumo_total_vars["saidas"].set(format_total(total_saidas))
            self.resumo_total_vars["saldo"].set(format_total(total_saldo))
        self._set_record_counter(len(self.resumo_grid.get_children()))

    @staticmethod
    def _pedido_numero(value):
        """Converte números digitados nos formatos 12,5 e 1.250,00."""
        text = str(value or "").strip().replace(" ", "")
        if not text:
            return 0.0
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return 0.0

    @staticmethod
    def _pedido_formatar(value, casas=2):
        value = float(value or 0)
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        return f"{value:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def show_acompanhamento_pedidos(self):
        """Acompanha níveis de estoque e necessidade de reposição por item."""
        self._clear_main()
        body = self._page_header("Acompanhamento de Pedidos", "Cadastros  /  Almoxarifado  /  Pedidos")

        top = tk.Frame(body, bg="#edf2f7")
        top.pack(fill="x", pady=(0, 10))
        tk.Label(top, text="A data-limite é calculada pela cobertura do saldo (saldo ÷ média de consumo). O lead time padrão é 20 dias e pode ser alterado.",
                 bg="#edf2f7", fg="#526b82", font=("Segoe UI", 8)).pack(side="left")
        self.pedidos_search = tk.StringVar()
        self.pedidos_search.trace_add("write", self._load_acompanhamento_pedidos)
        tk.Entry(top, textvariable=self.pedidos_search, width=28, font=("Segoe UI", 8), relief="solid", bd=1).pack(side="right", ipady=3)
        tk.Label(top, text="Pesquisar", bg="#edf2f7", fg="#38546e", font=("Segoe UI", 8, "bold")).pack(side="right", padx=(0, 8))

        content = tk.Frame(body, bg="#edf2f7")
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)
        grid_frame = tk.Frame(content, bg="white", bd=1, relief="solid")
        grid_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        columns = ("codigo", "descricao", "saldo", "lead_time", "media", "minimo", "maximo", "seguranca", "ponto", "ressuprimento", "status")
        headings = ("CÓDIGO", "DESCRIÇÃO", "SALDO", "LEAD\n(DIAS)", "MÉDIA\nCONSUMO", "ESTOQUE\nMÍNIMO", "ESTOQUE\nMÁXIMO", "EST.\nSEGURANÇA", "RESSUPRIMENTO", "RESSUPRIMENTO\n(LIMITE)", "STATUS")
        self.pedidos_grid = ttk.Treeview(grid_frame, columns=columns, show="headings", selectmode="browse")
        self.pedidos_grid.configure(displaycolumns=tuple(col for col in columns if col != "ressuprimento"))
        self.pedidos_columns = columns
        self.pedidos_headings = dict(zip(columns, headings))
        for col, heading in zip(columns, headings):
            self.pedidos_grid.heading(col, text=heading)
            self.pedidos_grid.column(col, width=90, minwidth=60, stretch=False,
                                    anchor="w" if col == "descricao" else "center")
        self.pedidos_grid.tag_configure("NORMAL", background="#e8f2ff", foreground="#145c9e")
        self.pedidos_grid.tag_configure("SOLICITAR", background="#fff5c2", foreground="#866400")
        self.pedidos_grid.tag_configure("URGENTE", background="#fde4e4", foreground="#b42318")
        self.pedidos_grid.bind("<Double-1>", self._editar_parametro_pedido)
        self.pedidos_grid.bind("<Button-1>", self._block_pedidos_column_resize, add="+")
        sy = ttk.Scrollbar(grid_frame, orient="vertical", command=self.pedidos_grid.yview)
        sx = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.pedidos_grid.xview)
        self.pedidos_grid.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.pedidos_grid.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)

        self._build_record_counter(body)
        self._load_acompanhamento_pedidos()

    def _load_acompanhamento_pedidos(self, *_args):
        if not hasattr(self, "pedidos_grid"):
            return
        for item in self.pedidos_grid.get_children():
            self.pedidos_grid.delete(item)
        term = self.pedidos_search.get().strip() if hasattr(self, "pedidos_search") else ""
        query = """
            WITH recebimentos AS (SELECT codigo, MAX(descricao) descricao, COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(qtd), '.', ''), ',', '.') AS REAL)), 0) entradas FROM recebimento_almoxarifado GROUP BY codigo),
            saidas AS (SELECT codigo, MAX(descricao) descricao, COALESCE(SUM(CAST(REPLACE(REPLACE(TRIM(qtd), '.', ''), ',', '.') AS REAL)), 0) saidas FROM saida_almoxarifado GROUP BY codigo),
            codigos AS (SELECT codigo FROM recebimentos UNION SELECT codigo FROM saidas)
            SELECT c.codigo, COALESCE(r.descricao, s.descricao, ''), COALESCE(r.entradas, 0)-COALESCE(s.saidas, 0),
                   p.lead_time, p.media_consumo, p.ponto_pedido, p.ressuprimento, COALESCE(p.status, 'NORMAL')
            FROM codigos c LEFT JOIN recebimentos r ON r.codigo=c.codigo LEFT JOIN saidas s ON s.codigo=c.codigo
            LEFT JOIN acompanhamento_pedidos p ON p.codigo=c.codigo
        """
        params = ()
        if term:
            query += " WHERE c.codigo LIKE ? OR COALESCE(r.descricao, s.descricao, '') LIKE ?"
            params = (f"%{term}%", f"%{term}%")
        query += " ORDER BY c.codigo"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for codigo, descricao, saldo, lead, media, ponto, ressuprimento, status in rows:
            saldo = float(saldo or 0)
            lead = 20 if lead is None else float(lead)
            minimo, maximo, seguranca = saldo * .30, saldo * 1.10, saldo * .50
            ponto, ressuprimento = self._calcular_datas_pedido(saldo, media, lead)
            values = (codigo, descricao, self._pedido_formatar(saldo), self._pedido_formatar(lead), self._pedido_formatar(media), self._pedido_formatar(minimo), self._pedido_formatar(maximo), self._pedido_formatar(seguranca), ponto or "-", ressuprimento or "-", status)
            self.pedidos_grid.insert("", "end", iid=codigo, values=values, tags=(status,))
        self._ajustar_colunas_pedidos()
        self._set_record_counter(len(rows))

    def _ajustar_colunas_pedidos(self):
        """Dimensiona a grade pelo cabeçalho e pelos valores, sem deixar lacunas excessivas."""
        medida = tkfont.Font(font=("Segoe UI", 8))
        limites = {"codigo": (75, 130), "descricao": (150, 360), "ressuprimento": (105, 140), "status": (90, 100)}
        for coluna in self.pedidos_columns:
            cabecalho = max(self.pedidos_headings[coluna].split("\n"), key=len)
            largura = medida.measure(cabecalho) + 28
            for item in self.pedidos_grid.get_children():
                largura = max(largura, medida.measure(str(self.pedidos_grid.set(item, coluna))) + 24)
            minimo, maximo = limites.get(coluna, (70, 115))
            self.pedidos_grid.column(coluna, width=max(minimo, min(largura, maximo)), minwidth=minimo, stretch=False)

    @staticmethod
    def _calcular_datas_pedido(saldo, media_consumo, lead_time):
        """Retorna data-limite de suprimento e ponto de pedido a partir da cobertura do estoque."""
        media = float(media_consumo or 0)
        if media <= 0:
            return "", ""
        dias_cobertura = max(0, float(saldo or 0) / media)
        data_ressuprimento = datetime.now() + timedelta(days=dias_cobertura)
        ponto_pedido = data_ressuprimento - timedelta(days=max(0, float(lead_time or 20)))
        return ponto_pedido.strftime("%d/%m/%Y"), data_ressuprimento.strftime("%d/%m/%Y")

    def _block_pedidos_column_resize(self, event):
        """Impede alteração manual das larguras pelos separadores do cabeçalho."""
        if self.pedidos_grid.identify_region(event.x, event.y) == "separator":
            return "break"

    def _editar_parametro_pedido(self, event):
        """Edita lead time e média de consumo diretamente na grade."""
        coluna = self.pedidos_grid.identify_column(event.x)
        item = self.pedidos_grid.identify_row(event.y)
        if not item or coluna not in ("#4", "#5"):
            return
        indice = int(coluna[1:]) - 1
        campo = self.pedidos_columns[indice]
        x, y, largura, altura = self.pedidos_grid.bbox(item, coluna)
        editor = tk.Entry(self.pedidos_grid, font=("Segoe UI", 8), justify="center")
        editor.insert(0, self.pedidos_grid.item(item, "values")[indice])
        editor.place(x=x, y=y, width=largura, height=altura)
        editor.focus_set()
        editor.select_range(0, "end")
        editor.bind("<Return>", lambda _event: self._salvar_celula_pedido(item, campo, editor))
        editor.bind("<FocusOut>", lambda _event: self._salvar_celula_pedido(item, campo, editor))
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def _salvar_celula_pedido(self, codigo, campo, editor):
        if not editor.winfo_exists():
            return
        valor = self._pedido_numero(editor.get())
        editor.destroy()
        with sqlite3.connect(self.database_path) as connection:
            row = connection.execute("SELECT lead_time, media_consumo FROM acompanhamento_pedidos WHERE codigo = ?", (codigo,)).fetchone()
            lead, media = row or (20, 0)
            lead = 20 if lead is None else float(lead)
            media = 0 if media is None else float(media)
            if campo == "lead_time":
                lead = valor or 20
            else:
                media = valor
            saldo = self._pedido_numero(self.pedidos_grid.item(codigo, "values")[2])
            ponto, ressuprimento = self._calcular_datas_pedido(saldo, media, lead)
            connection.execute("INSERT INTO acompanhamento_pedidos (codigo, lead_time, media_consumo, ponto_pedido, ressuprimento) VALUES (?, ?, ?, ?, ?) ON CONFLICT(codigo) DO UPDATE SET lead_time=excluded.lead_time, media_consumo=excluded.media_consumo, ponto_pedido=excluded.ponto_pedido, ressuprimento=excluded.ressuprimento", (codigo, lead, media, ponto, ressuprimento))
        self._load_acompanhamento_pedidos()
        self.pedidos_grid.selection_set(codigo)

    def _select_acompanhamento_item(self, _event):
        selected = self.pedidos_grid.selection()
        if not selected:
            return
        codigo = selected[0]
        values = self.pedidos_grid.item(codigo, "values")
        self.pedido_item_var.set(f"{values[0]} — {values[1]}")
        with sqlite3.connect(self.database_path) as connection:
            row = connection.execute("SELECT lead_time, media_consumo FROM acompanhamento_pedidos WHERE codigo = ?", (codigo,)).fetchone()
        lead, media = row or (20, "")
        lead = 20 if lead is None else lead
        saldo = self._pedido_numero(values[2])
        ponto, ressuprimento = self._calcular_datas_pedido(saldo, media, lead)
        self.pedido_vars["lead_time"].set(str(lead).replace(".", ","))
        self.pedido_vars["media_consumo"].set("" if media is None else str(media).replace(".", ","))
        self.pedido_vars["ponto_pedido"].set(ponto)
        self.pedido_vars["ressuprimento"].set(ressuprimento)

    def _save_acompanhamento_item(self):
        selected = self.pedidos_grid.selection()
        if not selected:
            messagebox.showinfo("Pedidos", "Selecione um item para informar os parâmetros.")
            return
        codigo = selected[0]
        lead = self._pedido_numero(self.pedido_vars["lead_time"].get()) or 20
        media = self._pedido_numero(self.pedido_vars["media_consumo"].get())
        saldo = self._pedido_numero(self.pedidos_grid.item(codigo, "values")[2])
        ponto, ressuprimento = self._calcular_datas_pedido(saldo, media, lead)
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("INSERT INTO acompanhamento_pedidos (codigo, lead_time, media_consumo, ponto_pedido, ressuprimento) VALUES (?, ?, ?, ?, ?) ON CONFLICT(codigo) DO UPDATE SET lead_time=excluded.lead_time, media_consumo=excluded.media_consumo, ponto_pedido=excluded.ponto_pedido, ressuprimento=excluded.ressuprimento", (codigo, lead, media, ponto, ressuprimento))
        self._load_acompanhamento_pedidos()
        self.pedidos_grid.selection_set(codigo)

    def _set_acompanhamento_status(self, status):
        selected = self.pedidos_grid.selection()
        if not selected:
            messagebox.showinfo("Status", "Selecione um item para definir o status.")
            return
        codigo = selected[0]
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("INSERT INTO acompanhamento_pedidos (codigo, status) VALUES (?, ?) ON CONFLICT(codigo) DO UPDATE SET status=excluded.status", (codigo, status))
        self._load_acompanhamento_pedidos()
        self.pedidos_grid.selection_set(codigo)

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
        has_header = all(field in indexes for field in ("data_recebimento", "descricao"))
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
            if not values[0].strip():
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
            mode = self._confirm_txt_import("recebimento_almoxarifado", "Recebimento de Itens")
            if mode is None:
                return
            columns = ", ".join(self.RECEBIMENTO_COLUMNS)
            placeholders = ", ".join("?" for _ in self.RECEBIMENTO_COLUMNS)
            query = f"INSERT INTO recebimento_almoxarifado ({columns}) VALUES ({placeholders})"
            with sqlite3.connect(self.database_path) as connection:
                if mode == "replace":
                    connection.execute("DELETE FROM recebimento_almoxarifado")
                rows, pendentes = self._preparar_codigos_importados(connection, rows, self.RECEBIMENTO_COLUMNS)
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_recebimento_items()
        message = f"{len(rows)} item(ns) carregado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter data."
        if pendentes:
            message += f"\n{pendentes} linha(s) sem código e descrição foi(ram) marcada(s) como pendente(s)."
        messagebox.showinfo("Carga TXT concluída", message)

    def show_almox_saida(self):
        self._clear_main()
        body = self._page_header("Saída de Itens de Almoxarifado", "Cadastros  /  Almoxarifado  /  Saída")
        self._build_saida_actions(body)
        self._build_saida_table(body)
        self._build_record_counter(body)
        self._load_saida_items()

    def _build_saida_actions(self, parent):
        actions = tk.Frame(parent, bg="#edf2f7")
        actions.pack(pady=16)
        self._action(actions, "Salvar novos/edições", self.save_saida_items, "#ffffff", "#27835c")
        self._action(actions, "＋  Nova Saída", self.new_saida_item, "#ffffff", "#315a7c")
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
        self._new_saida_rows = {}
        self._edited_saida_rows = {}
        self._new_saida_sequence = 0
        self.saida_grid.bind("<Double-1>", self._edit_saida_cell)
        self.saida_grid.bind("<F2>", self._edit_selected_saida_cell)

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
        # Registros novos ficam sempre na primeira linha, logo abaixo do cabeçalho.
        query += " ORDER BY id DESC"
        with sqlite3.connect(self.database_path) as connection:
            rows = connection.execute(query, params).fetchall()
        for row_id, *values in rows:
            values = self._edited_saida_rows.get(str(row_id), values)
            self.saida_grid.insert("", "end", iid=str(row_id), values=values)
        # Mantém visíveis as linhas que ainda estão sendo preenchidas na tabela.
        for item_id, values in getattr(self, "_new_saida_rows", {}).items():
            self.saida_grid.insert("", 0, iid=item_id, values=values, tags=("nova_saida",))
        self._auto_fit_tree_columns(self.saida_grid, self.SAIDA_COLUMNS)
        self._set_record_counter(len(self.saida_grid.get_children()))

    def filter_saida_items(self, *_args):
        if hasattr(self, "saida_grid"):
            self._load_saida_items()

    def new_saida_item(self):
        """Inclui uma linha editável no topo da tabela, com a data atual."""
        self._new_saida_sequence += 1
        item_id = f"nova_saida_{self._new_saida_sequence}"
        values = [""] * len(self.SAIDA_COLUMNS)
        values[0] = datetime.now().strftime("%d/%m/%Y")
        self._new_saida_rows[item_id] = values
        self.saida_grid.insert("", 0, iid=item_id, values=values, tags=("nova_saida",))
        self.saida_grid.selection_set(item_id)
        self.saida_grid.focus(item_id)
        self.saida_grid.see(item_id)
        self._auto_fit_tree_columns(self.saida_grid, self.SAIDA_COLUMNS)
        self.after_idle(lambda: self._open_saida_cell_editor(item_id, "#2"))

    def save_saida_items(self):
        """Grava de uma vez as novas linhas e as edicoes pendentes."""
        new_rows = list(self._new_saida_rows.items())
        incomplete = [item_id for item_id, values in new_rows if not values[1] or not values[2]]
        if incomplete:
            messagebox.showwarning(
                "Salvar registros",
                "Preencha codigo e descricao em todas as novas saidas antes de salvar.",
            )
            self.saida_grid.selection_set(incomplete[0])
            self.saida_grid.focus(incomplete[0])
            self.saida_grid.see(incomplete[0])
            return
        if not new_rows and not self._edited_saida_rows:
            messagebox.showinfo("Salvar registros", "Nao ha registros novos ou alteracoes para salvar.")
            return
        try:
            with sqlite3.connect(self.database_path) as connection:
                columns = ", ".join(self.SAIDA_COLUMNS)
                placeholders = ", ".join("?" for _ in self.SAIDA_COLUMNS)
                connection.executemany(
                    f"INSERT INTO saida_almoxarifado ({columns}) VALUES ({placeholders})",
                    [values for _item_id, values in new_rows],
                )
                assignments = ", ".join(f"{column} = ?" for column in self.SAIDA_COLUMNS)
                connection.executemany(
                    f"UPDATE saida_almoxarifado SET {assignments} WHERE id = ?",
                    [tuple(values) + (int(item_id),) for item_id, values in self._edited_saida_rows.items()],
                )
        except sqlite3.Error as error:
            messagebox.showerror("Erro ao salvar", f"Nao foi possivel salvar os registros.\n\n{error}")
            return
        saved_total = len(new_rows) + len(self._edited_saida_rows)
        self._new_saida_rows.clear()
        self._edited_saida_rows.clear()
        self._load_saida_items()
        messagebox.showinfo("Salvar registros", f"{saved_total} registro(s) salvo(s) com sucesso.")

    def delete_saida_item(self):
        selected = self.saida_grid.selection()
        if not selected:
            messagebox.showinfo("Excluir item", "Selecione um item na lista.")
            return
        item_id = selected[0]
        if item_id in self._new_saida_rows:
            self._new_saida_rows.pop(item_id, None)
            self.saida_grid.delete(item_id)
            self._set_record_counter(len(self.saida_grid.get_children()))
            return
        if not messagebox.askyesno("Excluir item", "Deseja mesmo excluir o registro?"):
            return
        with sqlite3.connect(self.database_path) as connection:
            connection.execute("DELETE FROM saida_almoxarifado WHERE id = ?", (int(item_id),))
        self._edited_saida_rows.pop(item_id, None)
        self._load_saida_items()

    def _edit_saida_cell(self, event):
        if self.saida_grid.identify_region(event.x, event.y) != "cell":
            return
        item_id = self.saida_grid.identify_row(event.y)
        column = self.saida_grid.identify_column(event.x)
        if item_id and column:
            self._open_saida_cell_editor(item_id, column)

    def _edit_selected_saida_cell(self, _event):
        selected = self.saida_grid.selection()
        if selected:
            self._open_saida_cell_editor(selected[0], "#1")
        return "break"

    def _open_saida_cell_editor(self, item_id, column):
        bbox = self.saida_grid.bbox(item_id, column)
        if not bbox:
            return
        index = int(column[1:]) - 1
        x, y, width, height = bbox
        editor = tk.Entry(self.saida_grid, font=("Segoe UI", 8))
        editor.insert(0, self.saida_grid.item(item_id, "values")[index])
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        editor.select_range(0, "end")
        editor.bind("<Return>", lambda _event: self._save_saida_cell_and_advance(item_id, index, editor))
        editor.bind("<FocusOut>", lambda _event: self._save_saida_cell(item_id, index, editor))
        editor.bind("<Escape>", lambda _event: editor.destroy())

    def _save_saida_cell_and_advance(self, item_id, index, editor):
        """Salva a célula atual e leva o cursor à próxima coluna com Enter."""
        saved_id = self._save_saida_cell(item_id, index, editor)
        next_index = index + 1
        if saved_id and next_index < len(self.SAIDA_COLUMNS):
            self.after_idle(lambda: self._open_saida_cell_editor(saved_id, f"#{next_index + 1}"))
        return "break"

    def _save_saida_cell(self, item_id, index, editor):
        if not editor.winfo_exists():
            return
        value = editor.get().strip()
        editor.destroy()
        values = list(self.saida_grid.item(item_id, "values"))
        values[index] = value

        if item_id in self._new_saida_rows:
            self._new_saida_rows[item_id] = values
            self.saida_grid.item(item_id, values=values)
            self._auto_fit_tree_columns(self.saida_grid, self.SAIDA_COLUMNS)
            return item_id

        self._edited_saida_rows[item_id] = values
        self.saida_grid.item(item_id, values=values)
        self._auto_fit_tree_columns(self.saida_grid, self.SAIDA_COLUMNS)
        return item_id

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
        has_header = all(field in indexes for field in ("data", "descricao"))
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
            if not values[0].strip():
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
            mode = self._confirm_txt_import("saida_almoxarifado", "Saída de Itens")
            if mode is None:
                return
            columns = ", ".join(self.SAIDA_COLUMNS)
            placeholders = ", ".join("?" for _ in self.SAIDA_COLUMNS)
            query = f"INSERT INTO saida_almoxarifado ({columns}) VALUES ({placeholders})"
            with sqlite3.connect(self.database_path) as connection:
                if mode == "replace":
                    connection.execute("DELETE FROM saida_almoxarifado")
                rows, pendentes = self._preparar_codigos_importados(connection, rows, self.SAIDA_COLUMNS)
                connection.executemany(query, rows)
        except (OSError, ValueError, csv.Error) as error:
            messagebox.showerror("Carga TXT", f"Não foi possível carregar o arquivo.\n\n{error}")
            return
        self._load_saida_items()
        message = f"{len(rows)} item(ns) carregado(s) no banco."
        if skipped:
            message += f"\n{skipped} linha(s) ignorada(s) por não conter data."
        if pendentes:
            message += f"\n{pendentes} linha(s) sem código e descrição foi(ram) marcada(s) como pendente(s)."
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
                if key in DATETIME_FIELDS:
                    widget.bind("<KeyRelease>", lambda _event, field=key: self._apply_datetime_mask(field))
                    widget.bind("<FocusOut>", lambda _event, field=key: self._apply_datetime_mask(field))
            widget.pack(fill="x", ipady=2)

    def _apply_datetime_mask(self, field):
        """Converte números digitados/colados em dd/MM/aaaa HH:mm:ss."""
        value = self.vars[field].get()
        digits = "".join(character for character in value if character.isdigit())[:14]
        separators = ((2, "/"), (4, "/"), (8, " "), (10, ":"), (12, ":"))
        masked = digits
        for position, separator in separators:
            if len(digits) > position:
                masked = masked[:position] + separator + masked[position:]
        if masked != value:
            self.vars[field].set(masked)

    def _validate_os_datetimes(self):
        """Garante que os horários preenchidos tenham data e hora válidas."""
        labels = {"hora_inicio": "Hora Início", "hora_final": "Hora Final"}
        for field in DATETIME_FIELDS:
            value = self.vars[field].get().strip()
            if not value:
                continue
            try:
                datetime.strptime(value, DATETIME_FORMAT)
            except ValueError:
                messagebox.showwarning(
                    "Data e hora inválidas",
                    f"{labels[field]} deve estar no formato dd/MM/aaaa HH:mm:ss.\n"
                    "Exemplo: 13/08/2026 17:03:26",
                )
                return False
        return True

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
        self._set_record_counter(len(self.grid.get_children()))

    def filter_os_items(self, *_args):
        if hasattr(self, "grid"):
            self._load_os_items()

    def new_os(self):
        for var in self.vars.values():
            var.set("")
        current_datetime = datetime.now().strftime(DATETIME_FORMAT)
        self.vars["hora_parada"].set(current_datetime)
        self.vars["hora_inicio"].set(current_datetime)
        self.vars["hora_final"].set(current_datetime)
        self.vars["tipo_servico"].set("CORRETIVA")
        self.vars["prioridade"].set("URGENTE")
        self.vars["especialidade"].set("MECÂNICA")
        self.vars["situacao"].set("ABERTA")

    def insert_os(self):
        required = ("setor", "numero_equipamento", "solicitante", "descricao", "tecnico")
        if any(not self.vars[key].get().strip() for key in required):
            messagebox.showwarning("Campos obrigatórios", "Preencha setor, nº equipamento, solicitante, descrição e técnico.")
            return
        if not self._validate_os_datetimes():
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
            mode = self._confirm_txt_import("ordens_servico", "Ordens de Serviço")
            if mode is None:
                return
            columns = ", ".join(self.COLUMNS)
            placeholders = ", ".join("?" for _ in self.COLUMNS)
            with sqlite3.connect(self.database_path) as connection:
                if mode == "replace":
                    connection.execute("DELETE FROM ordens_servico")
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
        if not messagebox.askyesno("Excluir OS", "Deseja mesmo excluir o registro?"):
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
        if not self._validate_os_datetimes():
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
