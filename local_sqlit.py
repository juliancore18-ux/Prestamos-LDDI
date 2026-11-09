import os
import json
import uuid
import sqlite3
import datetime
import tkinter as tk
from contextlib import closing
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================
# SISTEMA DE PRÉSTAMOS — LDDI  (versión SQLite)
# =============================================================


# =============================================================
# TEMA / ESTILOS CENTRALIZADOS
# =============================================================

class Theme:
    # Fondos
    BG          = "#F4F6F8"
    CARD        = "#FFFFFF"
    STATUS_BG   = "#E8F0E9"
    HEADER_BG   = "#1B5E20"

    # Acentos
    GREEN       = "#2E7D32"
    GREEN_HOVER = "#1B5E20"
    GREEN_LIGHT = "#4CAF50"
    GREEN_DARK  = "#1B5E20"
    ACCENT      = "#1565C0"

    # Texto
    FG          = "#212121"
    FG_MUTED    = "#616161"
    FG_HINT     = "#9E9E9E"
    FG_WHITE    = "#FFFFFF"

    # Estados Treeview
    ROW_OK      = "#E8F5E9"
    ROW_WARN    = "#FFF8E1"
    ROW_LATE    = "#FFEBEE"
    BORDER      = "#E0E0E0"

    # Fuentes (Segoe UI en Windows, fallback a Arial)
    FONT_TITLE  = ("Segoe UI", 18, "bold")
    FONT_HEAD   = ("Segoe UI", 12, "bold")
    FONT_BODY   = ("Segoe UI", 10)
    FONT_SMALL  = ("Segoe UI", 9)
    FONT_HINT   = ("Segoe UI", 8, "italic")
    FONT_BTN    = ("Segoe UI", 11, "bold")


# =============================================================
# CONFIG
# =============================================================

class Config:
    EQUIPOS = ["Laptop 17", "Laptop 19", "Laptop 20", "Laptop 23",
               "Laptop 24", "Laptop 25", "Laptop 26", "Laptop 29",
               "Laptop 30", "Laptop 31", "Laptop 32", "Laptop 34",
               "Laptop 54", "Laptop 58", "Laptop 59", "Laptop 60",
               "Laptop 61", "Laptop 62", "Laptop 63", "Laptop 66", "Laptop 67"]

    PROYECTORES = ["Proyector 12", "Proyector 4", "Proyector 9", "Proyector 65",
                   "Proyector 64", "Proyector 63", "Proyector 50", "Proyector 51",
                   "Proyector 52", "Proyector 53", "Proyector 54", "Proyector 55",
                   "Proyector 56", "Proyector 57", "Proyector 58", "Proyector 59",
                   "Proyector 60", "Proyector 61", "Proyector 62"]

    EXTRAS = ["EXTENSION", "HDMI", "BOCINAS"]

    MATRICULA_MAX  = 10
    MATRICULA_MIN  = 5
    DURACION_HORAS = 2
    ALERTA_MINUTOS = 15

    BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
    DB_PATH    = os.path.join(BASE_DIR, "prestamos.db")
    JSON_PATH  = os.path.join(BASE_DIR, "prestamos.json")   # solo para migración
    EXCEL_PATH = os.path.join(BASE_DIR, "PRESTAMOS.xlsx")
    HEADERS    = [
        "ID", "Matrícula", "Aula",
        "Laptop", "Proyector", "Extras",
        "Salida", "Límite", "Devolución",
        "Estado",
    ]

    SIN_LAPTOP    = "(ninguna)"
    SIN_PROYECTOR = "(ninguno)"


# =============================================================
# DATA (SQLite)
# =============================================================

class PrestamosDB:
    """
    Capa de acceso a datos con SQLite.
    Se abre una conexión corta por operación, así varias PCs/instancias
    pueden compartir el mismo archivo .db sin quedar bloqueado.
    """

    _SCHEMA = """
    CREATE TABLE IF NOT EXISTS prestamos (
        id               TEXT PRIMARY KEY,
        matricula        TEXT NOT NULL,
        salon            TEXT NOT NULL,
        equipo           TEXT NOT NULL DEFAULT '',
        proyector        TEXT NOT NULL DEFAULT '',
        extras           TEXT NOT NULL DEFAULT '',
        hora_salida      TEXT NOT NULL,
        hora_limite      TEXT NOT NULL,
        hora_devolucion  TEXT NOT NULL DEFAULT '',
        estado           TEXT NOT NULL DEFAULT 'PRESTADO'
                         CHECK (estado IN ('PRESTADO','DEVUELTO','DEVUELTO TARDE'))
    );

    CREATE INDEX IF NOT EXISTS idx_prestamos_estado ON prestamos(estado);

    -- Garantías a nivel de base de datos (solo aplican a préstamos activos)
    CREATE UNIQUE INDEX IF NOT EXISTS ux_activo_matricula
        ON prestamos(matricula) WHERE estado = 'PRESTADO';
    CREATE UNIQUE INDEX IF NOT EXISTS ux_activo_equipo
        ON prestamos(equipo) WHERE estado = 'PRESTADO' AND equipo <> '';
    CREATE UNIQUE INDEX IF NOT EXISTS ux_activo_proyector
        ON prestamos(proyector) WHERE estado = 'PRESTADO' AND proyector <> '';
    """

    _COLS = ("id", "matricula", "salon", "equipo", "proyector", "extras",
             "hora_salida", "hora_limite", "hora_devolucion", "estado")

    def __init__(self, path: str, json_legacy: str | None = None):
        self.path = path
        self._init_db()
        if json_legacy:
            self._migrar_json(json_legacy)

    # ---------- conexión ----------

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        try:
            with closing(self._connect()) as conn:
                conn.execute("PRAGMA journal_mode=WAL")
                conn.executescript(self._SCHEMA)
                conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror(
                "Error de base de datos",
                f"No se pudo inicializar la base de datos:\n{e}",
            )
            raise

    # ---------- migración desde JSON ----------

    def _migrar_json(self, json_path: str):
        """Importa prestamos.json una sola vez y lo renombra a .migrado."""
        if not os.path.exists(json_path):
            return
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            prestamos = data.get("prestamos", [])

            with closing(self._connect()) as conn:
                with conn:  # transacción
                    for p in prestamos:
                        conn.execute(
                            """INSERT OR IGNORE INTO prestamos
                               (id, matricula, salon, equipo, proyector, extras,
                                hora_salida, hora_limite, hora_devolucion, estado)
                               VALUES (?,?,?,?,?,?,?,?,?,?)""",
                            (
                                p.get("id") or str(uuid.uuid4()),
                                p.get("matricula", ""),
                                p.get("salon", ""),
                                p.get("equipo", "") or "",
                                p.get("proyector", "") or "",
                                p.get("extras", "") or "",
                                p.get("hora_salida", ""),
                                p.get("hora_limite", ""),
                                p.get("hora_devolucion", "") or "",
                                p.get("estado", "PRESTADO"),
                            ),
                        )
            os.replace(json_path, json_path + ".migrado")
            messagebox.showinfo(
                "Migración completada",
                f"Se importaron {len(prestamos)} registros desde prestamos.json.\n"
                "El archivo original fue renombrado a prestamos.json.migrado.",
            )
        except (json.JSONDecodeError, OSError, sqlite3.Error) as e:
            messagebox.showwarning(
                "Migración fallida",
                f"No se pudo migrar prestamos.json:\n{e}\n"
                "El archivo no fue modificado.",
            )

    # ---------- consultas ----------

    def activos(self) -> list[dict]:
        with closing(self._connect()) as conn:
            rows = conn.execute(
                "SELECT * FROM prestamos WHERE estado = 'PRESTADO' ORDER BY hora_salida"
            ).fetchall()
        return [dict(r) for r in rows]

    def todos(self) -> list[dict]:
        with closing(self._connect()) as conn:
            rows = conn.execute(
                "SELECT * FROM prestamos ORDER BY hora_salida"
            ).fetchall()
        return [dict(r) for r in rows]

    def obtener(self, prestamo_id: str) -> dict | None:
        with closing(self._connect()) as conn:
            row = conn.execute(
                "SELECT * FROM prestamos WHERE id = ?", (prestamo_id,)
            ).fetchone()
        return dict(row) if row else None

    def alumno_tiene_prestamo_activo(self, matricula: str) -> bool:
        with closing(self._connect()) as conn:
            row = conn.execute(
                "SELECT 1 FROM prestamos WHERE matricula = ? AND estado = 'PRESTADO' LIMIT 1",
                (matricula,),
            ).fetchone()
        return row is not None

    # ---------- escritura ----------

    def agregar(self, prestamo: dict):
        """
        Inserta un préstamo. Lanza ValueError si el alumno o el equipo
        ya están en uso (lo detectan los índices únicos de la BD).
        """
        prestamo["id"] = str(uuid.uuid4())
        try:
            with closing(self._connect()) as conn:
                with conn:
                    conn.execute(
                        """INSERT INTO prestamos
                           (id, matricula, salon, equipo, proyector, extras,
                            hora_salida, hora_limite, hora_devolucion, estado)
                           VALUES (:id, :matricula, :salon, :equipo, :proyector,
                                   :extras, :hora_salida, :hora_limite,
                                   :hora_devolucion, :estado)""",
                        {k: prestamo[k] for k in self._COLS},
                    )
        except sqlite3.IntegrityError as e:
            msg = str(e)
            if "matricula" in msg:
                raise ValueError("Esa matrícula ya tiene un préstamo activo.") from e
            if "equipo" in msg:
                raise ValueError("Esa laptop ya fue prestada a otra persona.") from e
            if "proyector" in msg:
                raise ValueError("Ese proyector ya fue prestado a otra persona.") from e
            raise ValueError(f"No se pudo registrar el préstamo: {e}") from e

    def devolver(self, prestamo_id: str, estado_final: str, hora_devolucion: str) -> dict | None:
        with closing(self._connect()) as conn:
            with conn:
                cur = conn.execute(
                    """UPDATE prestamos
                       SET estado = ?, hora_devolucion = ?
                       WHERE id = ? AND estado = 'PRESTADO'""",
                    (estado_final, hora_devolucion, prestamo_id),
                )
            if cur.rowcount == 0:
                return None
            row = conn.execute(
                "SELECT * FROM prestamos WHERE id = ?", (prestamo_id,)
            ).fetchone()
        return dict(row) if row else None


# =============================================================
# SERVICIO (lógica de negocio)
# =============================================================

class PrestamoService:
    def __init__(self, db: PrestamosDB):
        self.db = db

    def crear_prestamo(
        self,
        matricula: str,
        salon: str,
        equipo: str,
        proyector: str,
        extras: list,
    ) -> dict:
        ahora  = datetime.datetime.now()
        limite = ahora + datetime.timedelta(hours=Config.DURACION_HORAS)
        prestamo = {
            "id":               "",
            "matricula":        matricula,
            "salon":            salon,
            "equipo":           equipo,
            "proyector":        proyector,
            "extras":           ", ".join(extras),
            "hora_salida":      ahora.isoformat(),
            "hora_limite":      limite.isoformat(),
            "hora_devolucion":  "",
            "estado":           "PRESTADO",
        }
        self.db.agregar(prestamo)
        return prestamo

    def registrar_devolucion(self, prestamo_id: str) -> dict | None:
        """Calcula el estado ANTES de persistir."""
        prestamo = self.db.obtener(prestamo_id)
        if not prestamo or prestamo["estado"] != "PRESTADO":
            return None

        ahora        = datetime.datetime.now()
        limite       = datetime.datetime.fromisoformat(prestamo["hora_limite"])
        estado_final = "DEVUELTO TARDE" if ahora > limite else "DEVUELTO"
        return self.db.devolver(prestamo_id, estado_final, ahora.isoformat())

    def validar(
        self,
        matricula: str,
        salon: str,
        equipo: str,
        proyector: str,
    ) -> str | None:
        """Devuelve mensaje de error o None si todo es válido."""
        if not matricula:
            return "La matrícula es requerida."
        if not matricula.isdigit():
            return "La matrícula debe contener solo dígitos."
        if len(matricula) < Config.MATRICULA_MIN:
            return f"La matrícula debe tener al menos {Config.MATRICULA_MIN} dígitos."
        if len(matricula) > Config.MATRICULA_MAX:
            return f"La matrícula no puede tener más de {Config.MATRICULA_MAX} dígitos."
        if not salon:
            return "El aula es requerida."
        if self.db.alumno_tiene_prestamo_activo(matricula):
            return f"La matrícula {matricula} ya tiene un préstamo activo."
        if not equipo and not proyector:
            return "Debes seleccionar al menos una laptop o un proyector."
        return None


# =============================================================
# EXCEL
# =============================================================

class Excel:
    _ESTADO_COLORES = {
        "PRESTADO":       "FFF9C4",
        "DEVUELTO":       "C8E6C9",
        "DEVUELTO TARDE": "FFCDD2",
    }
    _COL_WIDTHS = [36, 14, 10, 14, 14, 22, 19, 19, 19, 16]

    def __init__(self, path: str):
        self.path = path

    def _crear(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Registros"
        ws.append(Config.HEADERS)

        header_fill   = PatternFill("solid", fgColor="1565C0")
        header_font   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        header_border = Border(bottom=Side(style="medium", color="FFFFFF"))

        for col in range(1, len(Config.HEADERS) + 1):
            cell            = ws.cell(row=1, column=col)
            cell.fill       = header_fill
            cell.font       = header_font
            cell.alignment  = Alignment(horizontal="center", vertical="center")
            cell.border     = header_border

        for i, ancho in enumerate(self._COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"
        wb.save(self.path)

    def _fmt_hora(self, iso: str) -> str:
        """Formatea un string ISO a 'YYYY-MM-DD HH:MM', o '' si vacío."""
        return iso[:16].replace("T", " ") if iso else ""

    def actualizar_o_agregar(self, prestamo: dict):
        if not os.path.exists(self.path):
            self._crear()

        try:
            wb = load_workbook(self.path)
        except PermissionError:
            messagebox.showerror(
                "Archivo bloqueado",
                f"No se pudo acceder a '{self.path}'.\n"
                "Cierra el archivo en Excel e intenta de nuevo.",
            )
            return

        ws = wb["Registros"]
        pid            = prestamo["id"]
        fila_existente = None

        for row in ws.iter_rows(min_row=2):
            if row[0].value == pid:
                fila_existente = row[0].row
                break

        datos = [
            prestamo["id"],
            prestamo["matricula"],
            prestamo["salon"],
            prestamo["equipo"],
            prestamo["proyector"],
            prestamo.get("extras", ""),
            self._fmt_hora(prestamo["hora_salida"]),
            self._fmt_hora(prestamo["hora_limite"]),
            self._fmt_hora(prestamo["hora_devolucion"]),
            prestamo["estado"],
        ]

        if fila_existente:
            for col, valor in enumerate(datos, 1):
                ws.cell(row=fila_existente, column=col, value=valor)
            target_row = fila_existente
        else:
            ws.append(datos)
            target_row = ws.max_row

        color = self._ESTADO_COLORES.get(prestamo["estado"], "FFFFFF")
        fill  = PatternFill("solid", fgColor=color)
        font  = Font(name="Arial", size=10)
        for col in range(1, len(Config.HEADERS) + 1):
            cell           = ws.cell(row=target_row, column=col)
            cell.fill      = fill
            cell.font      = font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[target_row].height = 18

        try:
            wb.save(self.path)
        except PermissionError:
            messagebox.showerror(
                "Archivo bloqueado",
                f"No se pudo guardar '{self.path}'.\n"
                "Cierra el archivo en Excel e intenta de nuevo.",
            )


# =============================================================
# APP
# =============================================================

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Sistema de Préstamos de Equipos — LDDI")
        self.root.geometry("720x620")
        self.root.minsize(620, 540)
        self.root.resizable(True, True)
        self.root.configure(bg=Theme.BG)

        # Centrar ventana
        self.root.update_idletasks()
        w, h = 820, 720
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self._setup_styles()

        self.db      = PrestamosDB(Config.DB_PATH, json_legacy=Config.JSON_PATH)
        self.service = PrestamoService(self.db)
        self.excel   = Excel(Config.EXCEL_PATH)

        self.devoluciones_map: dict[str, str] = {}
        self._notificados: set[str]           = set()

        self._ui()
        self._monitorear()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Notebook
        style.configure("TNotebook", background=Theme.BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                        font=Theme.FONT_BODY,
                        padding=[16, 8],
                        background="#E0E0E0",
                        foreground=Theme.FG)
        style.map("TNotebook.Tab",
                  background=[("selected", Theme.GREEN), ("active", "#C8E6C9")],
                  foreground=[("selected", "white"), ("active", Theme.FG)])

        # Treeview
        style.configure("Treeview",
                        background=Theme.CARD,
                        foreground=Theme.FG,
                        fieldbackground=Theme.CARD,
                        font=Theme.FONT_BODY,
                        rowheight=28)
        style.configure("Treeview.Heading",
                        font=Theme.FONT_HEAD,
                        background=Theme.GREEN,
                        foreground="white",
                        relief="flat")
        style.map("Treeview.Heading",
                  background=[("active", Theme.GREEN_DARK)])

        # Combobox
        style.configure("TCombobox", font=Theme.FONT_BODY)

    def _hover(self, widget, normal_bg, hover_bg):
        widget.bind("<Enter>", lambda e: widget.config(bg=hover_bg))
        widget.bind("<Leave>", lambda e: widget.config(bg=normal_bg))

    # ----------------------------------------------------------
    # CONSTRUCCIÓN DE UI
    # ----------------------------------------------------------

    def _ui(self):
        # Header
        header = tk.Frame(self.root, bg=Theme.HEADER_BG, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="Sistema de Préstamos de Equipos",
            font=Theme.FONT_TITLE,
            bg=Theme.HEADER_BG,
            fg=Theme.FG_WHITE,
        ).pack(side="left", padx=20, pady=10)

        tk.Label(
            header,
            text="LDDI",
            font=("Segoe UI", 10, "bold"),
            bg=Theme.HEADER_BG,
            fg="#A5D6A7",
        ).pack(side="right", padx=20)

        # Notebook
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        self.tab_prestamo   = tk.Frame(nb, bg=Theme.BG)
        self.tab_devolucion = tk.Frame(nb, bg=Theme.BG)
        self.tab_historial  = tk.Frame(nb, bg=Theme.BG)

        nb.add(self.tab_prestamo,   text="  📤  Préstamo  ")
        nb.add(self.tab_devolucion, text="  📥  Devolución  ")
        nb.add(self.tab_historial,  text="  📋  Activos  ")

        self._ui_prestamo()
        self._ui_devolucion()
        self._ui_historial()

        # Status bar
        status_frame = tk.Frame(self.root, bg=Theme.STATUS_BG, height=32)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)

        self.lbl_estado = tk.Label(
            status_frame,
            text="● Listo",
            font=Theme.FONT_BODY,
            bg=Theme.STATUS_BG,
            fg=Theme.GREEN_DARK,
            anchor="w",
            padx=12,
        )
        self.lbl_estado.pack(fill="x", expand=True)

        self._refresh()

    def _ui_prestamo(self):
        f = self.tab_prestamo

        # Título de sección
        tk.Label(
            f,
            text="Nuevo préstamo",
            font=Theme.FONT_HEAD,
            bg=Theme.BG,
            fg=Theme.GREEN_DARK,
        ).pack(pady=(16, 8), anchor="w", padx=24)

        # Tarjeta del formulario
        card = tk.Frame(f, bg=Theme.CARD, highlightbackground=Theme.BORDER,
                        highlightthickness=1)
        card.pack(padx=24, pady=4, fill="x")

        form = tk.Frame(card, bg=Theme.CARD)
        form.pack(padx=20, pady=16, fill="x")

        self.matricula = tk.StringVar()
        self.salon     = tk.StringVar()

        # Matrícula
        tk.Label(form, text="Matrícula", font=Theme.FONT_BODY, bg=Theme.CARD,
                 fg=Theme.FG).grid(row=0, column=0, sticky="w", pady=(0, 2))
        entry_mat = tk.Entry(form, textvariable=self.matricula, font=Theme.FONT_BODY,
                             relief="solid", bd=1, highlightthickness=1,
                             highlightcolor=Theme.GREEN, highlightbackground=Theme.BORDER)
        entry_mat.grid(row=1, column=0, sticky="ew", pady=(0, 12), ipady=4)
        tk.Label(form, text=f"{Config.MATRICULA_MIN}–{Config.MATRICULA_MAX} dígitos",
                 font=Theme.FONT_HINT, bg=Theme.CARD, fg=Theme.FG_HINT).grid(
            row=1, column=1, sticky="w", padx=8)

        # Aula
        tk.Label(form, text="Aula", font=Theme.FONT_BODY, bg=Theme.CARD,
                 fg=Theme.FG).grid(row=2, column=0, sticky="w", pady=(0, 2))
        entry_aula = tk.Entry(form, textvariable=self.salon, font=Theme.FONT_BODY,
                              relief="solid", bd=1, highlightthickness=1,
                              highlightcolor=Theme.GREEN, highlightbackground=Theme.BORDER)
        entry_aula.grid(row=3, column=0, sticky="ew", pady=(0, 12), ipady=4)

        # Laptop
        tk.Label(form, text="Laptop", font=Theme.FONT_BODY, bg=Theme.CARD,
                 fg=Theme.FG).grid(row=4, column=0, sticky="w", pady=(0, 2))
        self.cb_equipo = ttk.Combobox(form, state="readonly", font=Theme.FONT_BODY, width=32)
        self.cb_equipo.grid(row=5, column=0, sticky="ew", pady=(0, 12))

        # Proyector
        tk.Label(form, text="Proyector", font=Theme.FONT_BODY, bg=Theme.CARD,
                 fg=Theme.FG).grid(row=6, column=0, sticky="w", pady=(0, 2))
        self.cb_proy = ttk.Combobox(form, state="readonly", font=Theme.FONT_BODY, width=32)
        self.cb_proy.grid(row=7, column=0, sticky="ew", pady=(0, 12))

        # Extras
        tk.Label(form, text="Extras", font=Theme.FONT_BODY, bg=Theme.CARD,
                 fg=Theme.FG).grid(row=8, column=0, sticky="w", pady=(0, 4))

        extras_frame = tk.Frame(form, bg=Theme.CARD)
        extras_frame.grid(row=9, column=0, sticky="w", pady=(0, 8))

        self.extras_vars: dict[str, tk.BooleanVar] = {}
        for i, extra in enumerate(Config.EXTRAS):
            var = tk.BooleanVar()
            self.extras_vars[extra] = var
            cb = tk.Checkbutton(
                extras_frame, text=extra, variable=var,
                bg=Theme.CARD, font=Theme.FONT_BODY,
                activebackground=Theme.CARD, selectcolor=Theme.CARD,
            )
            cb.pack(side="left", padx=(0, 16))

        form.columnconfigure(0, weight=1)

        # Notas
        notes = tk.Frame(f, bg=Theme.BG)
        notes.pack(padx=24, pady=(8, 0), fill="x")
        tk.Label(
            notes,
            text="⚠  Se debe seleccionar al menos una laptop o un proyector.",
            font=Theme.FONT_SMALL, bg=Theme.BG, fg=Theme.FG_MUTED, anchor="w",
        ).pack(anchor="w")
        tk.Label(
            notes,
            text="Un alumno no puede tener dos préstamos activos simultáneos.",
            font=Theme.FONT_SMALL, bg=Theme.BG, fg=Theme.FG_MUTED, anchor="w",
        ).pack(anchor="w")

        # Botón
        self.btn_prestar = tk.Button(
            f,
            text="Registrar préstamo",
            font=Theme.FONT_BTN,
            bg=Theme.GREEN,
            fg="white",
            activebackground=Theme.GREEN_HOVER,
            activeforeground="white",
            relief="flat",
            padx=24,
            pady=10,
            cursor="hand2",
            command=self.prestar,
        )
        self.btn_prestar.pack(pady=20)
        self._hover(self.btn_prestar, Theme.GREEN, Theme.GREEN_HOVER)

        tk.Label(
            f, text="Creado en LDDI",
            font=Theme.FONT_HINT, bg=Theme.BG, fg=Theme.FG_HINT,
        ).pack(pady=(0, 12))

    def _ui_devolucion(self):
        f = self.tab_devolucion

        tk.Label(
            f, text="Registrar devolución",
            font=Theme.FONT_HEAD, bg=Theme.BG, fg=Theme.GREEN_DARK,
        ).pack(pady=(16, 8), anchor="w", padx=24)

        # Tarjeta
        card = tk.Frame(f, bg=Theme.CARD, highlightbackground=Theme.BORDER,
                        highlightthickness=1)
        card.pack(padx=24, pady=4, fill="x")

        inner = tk.Frame(card, bg=Theme.CARD)
        inner.pack(padx=20, pady=16, fill="x")

        tk.Label(
            inner, text="Selecciona el préstamo activo:",
            bg=Theme.CARD, font=Theme.FONT_BODY, fg=Theme.FG,
        ).pack(anchor="w")

        self.cb_devolucion = ttk.Combobox(inner, width=55, state="readonly", font=Theme.FONT_BODY)
        self.cb_devolucion.pack(fill="x", pady=8)

        self.btn_devolver = tk.Button(
            inner,
            text="Registrar devolución",
            font=Theme.FONT_BTN,
            bg=Theme.GREEN_DARK,
            fg="white",
            activebackground="#0D3B12",
            activeforeground="white",
            relief="flat",
            padx=24,
            pady=10,
            cursor="hand2",
            command=self.devolver,
        )
        self.btn_devolver.pack(pady=16)
        self._hover(self.btn_devolver, Theme.GREEN_DARK, "#0D3B12")

        # Info
        self.info_frame = tk.LabelFrame(
            f, text="  Información del préstamo  ",
            bg=Theme.CARD, font=Theme.FONT_BODY, fg=Theme.FG,
            highlightbackground=Theme.BORDER, highlightthickness=1,
            relief="flat",
        )
        self.info_frame.pack(padx=24, fill="x", pady=(8, 16))

        self.info_label = tk.Label(
            self.info_frame,
            text="Selecciona un préstamo para ver detalles.",
            bg=Theme.CARD, font=Theme.FONT_BODY,
            justify="left", anchor="w", fg=Theme.FG,
            wraplength=700,
        )
        self.info_label.pack(padx=12, pady=12, anchor="w", fill="x")

        self.cb_devolucion.bind("<<ComboboxSelected>>", self._mostrar_info_prestamo)

    def _ui_historial(self):
        f = self.tab_historial

        tk.Label(
            f, text="Préstamos activos",
            font=Theme.FONT_HEAD, bg=Theme.BG, fg=Theme.GREEN_DARK,
        ).pack(pady=(16, 8), anchor="w", padx=24)

        # Barra de búsqueda
        busq_frame = tk.Frame(f, bg=Theme.BG)
        busq_frame.pack(padx=24, fill="x", pady=(0, 8))

        tk.Label(
            busq_frame, text="Buscar matrícula:",
            bg=Theme.BG, font=Theme.FONT_BODY, fg=Theme.FG,
        ).pack(side="left")

        self.busq_var = tk.StringVar()
        self.busq_var.trace_add("write", lambda *_: self._refresh())
        entry_busq = tk.Entry(
            busq_frame, textvariable=self.busq_var,
            width=22, font=Theme.FONT_BODY,
            relief="solid", bd=1, highlightthickness=1,
            highlightcolor=Theme.GREEN, highlightbackground=Theme.BORDER,
        )
        entry_busq.pack(side="left", padx=8, ipady=3)

        # Leyenda de colores
        legend = tk.Frame(busq_frame, bg=Theme.BG)
        legend.pack(side="right")
        for color, text in [
            (Theme.ROW_OK, "En tiempo"),
            (Theme.ROW_WARN, "Por vencer"),
            (Theme.ROW_LATE, "Vencido"),
        ]:
            tk.Label(legend, text="  ", bg=color, width=2).pack(side="left", padx=(8, 2))
            tk.Label(legend, text=text, font=Theme.FONT_SMALL, bg=Theme.BG, fg=Theme.FG_MUTED).pack(side="left")

        # Contenedor del Treeview
        tree_container = tk.Frame(f, bg=Theme.BG)
        tree_container.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        cols = (
            "Matrícula", "Aula", "Laptop", "Proyector",
            "Extras", "Salida", "Límite", "Estado",
        )
        self.tree = ttk.Treeview(tree_container, columns=cols, show="headings", height=12)

        anchos = [100, 80, 110, 120, 140, 90, 90, 100]
        for col, ancho in zip(cols, anchos):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=ancho, anchor="center", minwidth=70)

        scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.tag_configure("vencido",    background=Theme.ROW_LATE)
        self.tree.tag_configure("por_vencer", background=Theme.ROW_WARN)
        self.tree.tag_configure("ok",         background=Theme.ROW_OK)

    # ----------------------------------------------------------
    # LÓGICA INTERNA
    # ----------------------------------------------------------

    def _set_busy(self, ocupado: bool):
        """Deshabilita botones durante operaciones para evitar doble clic."""
        estado = "disabled" if ocupado else "normal"
        self.btn_prestar.config(state=estado)
        self.btn_devolver.config(state=estado)
        if ocupado:
            self.lbl_estado.config(text="⏳  Guardando...", fg="#555555")
        self.root.update_idletasks()

    def _refresh(self):
        activos = self.db.activos()
        ahora   = datetime.datetime.now()
        filtro  = self.busq_var.get().strip() if hasattr(self, "busq_var") else ""

        equipos_ocupados = {p["equipo"]    for p in activos if p["equipo"]}
        proy_ocupados    = {p["proyector"] for p in activos if p["proyector"]}

        self.cb_equipo["values"] = [Config.SIN_LAPTOP]    + [e for e in Config.EQUIPOS     if e not in equipos_ocupados]
        self.cb_proy["values"]   = [Config.SIN_PROYECTOR] + [p for p in Config.PROYECTORES if p not in proy_ocupados]

        if self.cb_equipo.get() not in self.cb_equipo["values"]:
            self.cb_equipo.set(Config.SIN_LAPTOP)
        if self.cb_proy.get() not in self.cb_proy["values"]:
            self.cb_proy.set(Config.SIN_PROYECTOR)

        # Opciones de devolución
        self.devoluciones_map.clear()
        opciones = []
        for p in activos:
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = int((limite - ahora).total_seconds() / 60)
            aviso = (
                f"⚠ {restante} min" if 0 < restante <= Config.ALERTA_MINUTOS
                else ("🔴 VENCIDO"  if restante <= 0 else "")
            )
            txt = f'{p["matricula"]} → {p["equipo"] or "-"} / {p["proyector"] or "-"}  {aviso}'
            opciones.append(txt)
            self.devoluciones_map[txt] = p["id"]

        self.cb_devolucion["values"] = opciones
        if self.cb_devolucion.get() not in opciones:
            self.cb_devolucion.set("")
            self.info_label.config(text="Selecciona un préstamo para ver detalles.")

        if not opciones:
            self.info_label.config(text="Sin préstamos activos.")

        # Árbol con filtro por matrícula
        for item in self.tree.get_children():
            self.tree.delete(item)

        for p in activos:
            if filtro and filtro not in p["matricula"]:
                continue
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = (limite - ahora).total_seconds() / 60
            tag = (
                "vencido"    if restante <= 0
                else "por_vencer" if restante <= Config.ALERTA_MINUTOS
                else "ok"
            )
            self.tree.insert("", "end", values=(
                p["matricula"],
                p["salon"],
                p["equipo"]    or "-",
                p["proyector"] or "-",
                p.get("extras", "") or "-",
                p["hora_salida"][11:16],
                p["hora_limite"][11:16],
                "VENCIDO" if restante <= 0 else "PRESTADO",
            ), tags=(tag,))

    def _monitorear(self):
        """Revisa vencimientos cada minuto; notifica una sola vez por préstamo."""
        activos     = self.db.activos()
        ahora       = datetime.datetime.now()
        ids_activos = {p["id"] for p in activos}

        self._notificados &= ids_activos

        for p in activos:
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = (limite - ahora).total_seconds() / 60
            if restante <= 0 and p["id"] not in self._notificados:
                self._notificados.add(p["id"])
                messagebox.showwarning(
                    "Préstamo vencido",
                    f"El préstamo de la matrícula {p['matricula']}\n"
                    f"({p['equipo'] or '-'} / {p['proyector'] or '-'}) ha vencido.",
                )

        self._refresh()
        self.root.after(60_000, self._monitorear)

    def _mostrar_info_prestamo(self, _event=None):
        sel = self.cb_devolucion.get()
        pid = self.devoluciones_map.get(sel)
        if not pid:
            return
        p = self.db.obtener(pid)
        if not p or p["estado"] != "PRESTADO":
            return

        limite   = datetime.datetime.fromisoformat(p["hora_limite"])
        ahora    = datetime.datetime.now()
        restante = int((limite - ahora).total_seconds() / 60)
        estado_txt = f"⏱ {restante} min restantes" if restante > 0 else "🔴 TIEMPO VENCIDO"

        self.info_label.config(
            text=(
                f"Matrícula: {p['matricula']}     Aula: {p['salon']}\n"
                f"Laptop: {p['equipo'] or 'N/A'}     Proyector: {p['proyector'] or 'N/A'}\n"
                f"Extras: {p.get('extras', '') or 'N/A'}\n"
                f"Salida: {p['hora_salida'][11:16]}     "
                f"Límite: {p['hora_limite'][11:16]}     {estado_txt}"
            )
        )

    # ----------------------------------------------------------
    # ACCIONES
    # ----------------------------------------------------------

    def prestar(self):
        matricula = self.matricula.get().strip()
        salon     = self.salon.get().strip()
        equipo    = self.cb_equipo.get()
        proy      = self.cb_proy.get()

        equipo = "" if equipo == Config.SIN_LAPTOP    else equipo
        proy   = "" if proy   == Config.SIN_PROYECTOR else proy
        extras = [k for k, v in self.extras_vars.items() if v.get()]

        error = self.service.validar(matricula, salon, equipo, proy)
        if error:
            messagebox.showerror("Error", error)
            return

        resumen = (
            f"Matrícula: {matricula}\n"
            f"Aula: {salon}\n"
            f"Laptop: {equipo or 'N/A'}\n"
            f"Proyector: {proy or 'N/A'}\n"
            f"Extras: {', '.join(extras) or 'N/A'}\n\n"
            "¿Confirmar préstamo?"
        )
        if not messagebox.askokcancel("Confirmar préstamo", resumen):
            return

        self._set_busy(True)
        try:
            prestamo = self.service.crear_prestamo(matricula, salon, equipo, proy, extras)
            self.excel.actualizar_o_agregar(prestamo)
        except ValueError as e:
            # Conflicto detectado por la BD (otro usuario se adelantó)
            self._set_busy(False)
            messagebox.showerror("Error", str(e))
            self._refresh()
            return
        finally:
            self._set_busy(False)

        self.matricula.set("")
        self.salon.set("")
        for v in self.extras_vars.values():
            v.set(False)

        limite = datetime.datetime.fromisoformat(prestamo["hora_limite"])
        self.lbl_estado.config(
            text=f"✔  Préstamo registrado. Límite: {limite.strftime('%H:%M')}",
            fg=Theme.GREEN_DARK,
        )
        self._refresh()

    def devolver(self):
        seleccion = self.cb_devolucion.get()
        if not seleccion:
            messagebox.showerror("Error", "Selecciona un préstamo activo.")
            return

        prestamo_id = self.devoluciones_map.get(seleccion)
        if not prestamo_id:
            messagebox.showerror("Error", "Préstamo no encontrado.")
            return

        if not messagebox.askokcancel(
            "Confirmar devolución", "¿Registrar la devolución de este equipo?"
        ):
            return

        self._set_busy(True)
        try:
            p = self.service.registrar_devolucion(prestamo_id)
        finally:
            self._set_busy(False)

        if not p:
            messagebox.showerror("Error", "No se encontró el préstamo activo en la base de datos.")
            self._refresh()
            return

        self.excel.actualizar_o_agregar(p)

        if p["estado"] == "DEVUELTO TARDE":
            messagebox.showwarning(
                "Entrega tardía", "El equipo fue devuelto fuera del tiempo límite."
            )

        self.cb_devolucion.set("")
        self.info_label.config(text="Selecciona un préstamo para ver detalles.")
        self.lbl_estado.config(
            text=f"✔  Devolución registrada ({p['estado']})",
            fg=Theme.GREEN,
        )
        self._refresh()


# =============================================================
# MAIN
# =============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app  = App(root)
    root.mainloop()