import os
import json
import uuid
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageTk

# =========================
# PALETA DE COLORES
# =========================

class Colors:
    # Fondo y superficies
    BG          = "#F0F4F8"       # fondo general
    SURFACE     = "#FFFFFF"       # tarjetas / paneles
    SURFACE_ALT = "#F7FAFC"       # superficies secundarias

    # Acento principal — verde institucional con más peso visual
    PRIMARY     = "#2D7D46"
    PRIMARY_LT  = "#EAF6EE"
    PRIMARY_BTN = "#38944F"       # botón principal
    PRIMARY_HOV = "#2D7D46"       # hover

    # Acento secundario — azul para historial
    SECONDARY   = "#57A519"
    SECONDARY_LT= "#EBF2FB"

    # Semáforo de estados
    OK          = "#2D7D46"
    WARN        = "#C17D00"
    DANGER      = "#C0392B"
    WARN_BG     = "#FFF8E1"
    DANGER_BG   = "#FEECEB"
    OK_BG       = "#EAF6EE"

    # Texto
    TEXT        = "#1A202C"
    TEXT_MUTED  = "#718096"
    TEXT_INV    = "#FFFFFF"

    # Bordes
    BORDER      = "#DDE3EA"
    BORDER_FOCUS= "#38944F"

    # Sidebar
    SIDEBAR     = "#1E2D40"
    SIDEBAR_ACT = "#2D7D46"
    SIDEBAR_TXT = "#CBD5E0"
    SIDEBAR_HOV = "#57A519"


# =========================
# CONFIG
# =========================

# PARTE ELABORADA POR EZEQUIEL

# EZEQUIEL COLOCO 30 EQUIPOS
class Config:
    EQUIPOS         = ["Laptop 17","Laptop 19","Laptop 20","Laptop 23",
                       "Laptop 24","Laptop 25","Laptop 26","Laptop 29",
                       "Laptop 30","Laptop 31","Laptop 32","Laptop 34",
                       "Laptop 54","Laptop 58","Laptop 59","Laptop 60",
                       "Laptop 61","Laptop 62","Laptop 63","Laptop 66","Laptop 67"]
    PROYECTORES     = ["Proyector 12", "Proyector 4", "Proyector 9","Proyector 65",
                       "Proyector 64","Proyector 63","Proyector 50", "Proyector 51",
                       "Proyector 52","Proyector 53","Proyector 54","Proyector 55",
                       "Proyector 56","Proyector 57","Proyector 58", "Proyector 59",
                       "Proyector 60","Proyector 61","Proyector 62"]
    EXTRAS          = ["EXTENSION", "HDMI", "BOCINAS"]
    MATRICULA_MAX   = 10
    MATRICULA_MIN   = 5
    DURACION_HORAS  = 2
    ALERTA_MINUTOS  = 15

    JSON_PATH       = "prestamos.json"
    EXCEL_PATH      = "PRESTAMOS.xlsx"

    HEADERS = [
        "ID", "Matrícula", "Aula",
        "Laptop", "Proyector", "Extras",
        "Salida", "Límite", "Devolución",
        "Estado"
    ]


# =========================
# DATA
# =========================

class PrestamosDB:
    def __init__(self, path):
        self.path = path
        self.data = {"prestamos": []}
        self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
        except json.JSONDecodeError as e:
            messagebox.showerror("Error de datos",
                f"El archivo JSON está corrupto:\n{e}\nSe iniciará con datos vacíos.")
            self.data = {"prestamos": []}
        except IOError as e:
            messagebox.showerror("Error de archivo", f"No se pudo leer el archivo:\n{e}")
            self.data = {"prestamos": []}

    def save(self):
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
            os.replace(tmp, self.path)
        except IOError as e:
            messagebox.showerror("Error al guardar", f"No se pudo guardar:\n{e}")
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    def _reload(self):
        self._load()

    def activos(self):
        self._reload()
        return [p for p in self.data["prestamos"] if p["estado"] == "PRESTADO"]

    def todos(self):
        self._reload()
        return self.data["prestamos"]

    def agregar(self, prestamo):
        self._reload()
        ids_existentes = {p["id"] for p in self.data["prestamos"]}
        nuevo_id = str(uuid.uuid4())
        while nuevo_id in ids_existentes:
            nuevo_id = str(uuid.uuid4())
        prestamo["id"] = nuevo_id
        self.data["prestamos"].append(prestamo)
        self.save()

    def devolver(self, prestamo_id, estado_final: str, hora_devolucion: str):
        self._reload()
        for p in self.data["prestamos"]:
            if p["id"] == prestamo_id:
                p["estado"] = estado_final
                p["hora_devolucion"] = hora_devolucion
                self.save()
                return p
        return None

    def alumno_tiene_prestamo_activo(self, matricula):
        return any(p["matricula"] == matricula for p in self.activos())


# =========================
# SERVICIO
# =========================

class PrestamoService:
    def __init__(self, db: PrestamosDB):
        self.db = db

    def crear_prestamo(self, matricula, salon, equipo, proyector, extras) -> dict:
        ahora  = datetime.datetime.now()
        limite = ahora + datetime.timedelta(hours=Config.DURACION_HORAS)
        prestamo = {
            "id": "",
            "matricula": matricula,
            "salon": salon,
            "equipo": equipo,
            "proyector": proyector,
            "extras": ", ".join(extras),
            "hora_salida": ahora.isoformat(),
            "hora_limite": limite.isoformat(),
            "hora_devolucion": "",
            "estado": "PRESTADO",
        }
        self.db.agregar(prestamo)
        return prestamo

    def registrar_devolucion(self, prestamo_id: str) -> dict | None:
        prestamo = next((p for p in self.db.activos() if p["id"] == prestamo_id), None)
        if not prestamo:
            return None
        ahora  = datetime.datetime.now()
        limite = datetime.datetime.fromisoformat(prestamo["hora_limite"])
        estado = "DEVUELTO TARDE" if ahora > limite else "DEVUELTO"
        return self.db.devolver(prestamo_id, estado, ahora.isoformat())

    def validar(self, matricula, salon, equipo, proyector) -> str | None:
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
            return "Selecciona al menos una laptop o un proyector."
        return None


# =========================
# EXCEL
# =========================

class Excel:
    def __init__(self, path):
        self.path = path
        self._ESTADO_COLORES = {
            "PRESTADO":       "FFF9C4",
            "DEVUELTO":       "C8E6C9",
            "DEVUELTO TARDE": "FFCDD2",
        }

    def _crear(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Registros"
        ws.append(Config.HEADERS)
        header_fill   = PatternFill("solid", fgColor="1565C0")
        header_font   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        header_border = Border(bottom=Side(style="medium", color="FFFFFF"))
        for col in range(1, len(Config.HEADERS) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = header_border
        anchos = [36, 14, 10, 14, 14, 22, 19, 19, 19, 16]
        for i, ancho in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"
        wb.save(self.path)

    def actualizar_o_agregar(self, prestamo: dict):
        if not os.path.exists(self.path):
            self._crear()
        try:
            wb = load_workbook(self.path)
        except PermissionError:
            messagebox.showerror("Archivo bloqueado",
                f"Cierra '{self.path}' en Excel e intenta de nuevo.")
            return
        ws = wb["Registros"]
        pid = prestamo["id"]
        fila_existente = None
        for row in ws.iter_rows(min_row=2):
            if row[0].value == pid:
                fila_existente = row[0].row
                break
        datos = [
            prestamo["id"], prestamo["matricula"], prestamo["salon"],
            prestamo["equipo"], prestamo["proyector"],
            prestamo.get("extras", ""),
            prestamo["hora_salida"][:16].replace("T", " "),
            prestamo["hora_limite"][:16].replace("T", " "),
            prestamo["hora_devolucion"][:16].replace("T", " ") if prestamo["hora_devolucion"] else "",
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
            cell = ws.cell(row=target_row, column=col)
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[target_row].height = 18
        try:
            wb.save(self.path)
        except PermissionError:
            messagebox.showerror("Archivo bloqueado",
                f"Cierra '{self.path}' en Excel e intenta de nuevo.")


# =========================
# WIDGETS PERSONALIZADOS
# =========================

class RoundedButton(tk.Canvas):
    """Botón con esquinas redondeadas y efecto hover."""
    def __init__(self, parent, text, command=None, bg=Colors.PRIMARY_BTN,
                 fg=Colors.TEXT_INV, hover=Colors.PRIMARY_HOV,
                 width=200, height=40, radius=10, font_size=10, **kwargs):
        super().__init__(parent, width=width, height=height,
                         bg=parent.cget("bg"), highlightthickness=0, **kwargs)
        self._bg      = bg
        self._hover   = hover
        self._fg      = fg
        self._text    = text
        self._command = command
        self._radius  = radius
        self._font    = tkfont.Font(family="Segoe UI", size=font_size, weight="bold")
        self._enabled = True
        self._draw(bg)
        self.bind("<Enter>",          lambda e: self._draw(hover) if self._enabled else None)
        self.bind("<Leave>",          lambda e: self._draw(bg)    if self._enabled else None)
        self.bind("<ButtonRelease-1>",lambda e: self._click())

    def _round_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [
            x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
            x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
            x1, y2, x1, y2-r, x1, y1+r, x1, y1,
        ]
        return self.create_polygon(pts, smooth=True, **kw)

    def _draw(self, color):
        self.delete("all")
        w = int(self.cget("width"))
        h = int(self.cget("height"))
        c = "#A0ADB8" if not self._enabled else color
        self._round_rect(2, 2, w-2, h-2, self._radius, fill=c, outline="")
        self.create_text(w//2, h//2, text=self._text, fill=self._fg,
                         font=self._font, anchor="center")

    def _click(self):
        if self._enabled and self._command:
            self._command()

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        self._draw(self._bg if enabled else self._bg)


class Card(tk.Frame):
    """Panel con sombra simulada, borde suave y padding."""
    def __init__(self, parent, title="", accent=Colors.PRIMARY, **kwargs):
        super().__init__(parent, bg=Colors.SURFACE, relief="flat",
                         highlightbackground=Colors.BORDER, highlightthickness=1, **kwargs)
        if title:
            hdr = tk.Frame(self, bg=accent, height=4)
            hdr.pack(fill="x", side="top")
            tk.Label(self, text=title, bg=Colors.SURFACE,
                     fg=Colors.TEXT, font=("Segoe UI", 11, "bold"),
                     anchor="w", padx=16, pady=10).pack(fill="x")
            ttk.Separator(self, orient="horizontal").pack(fill="x", padx=0)


class StyledEntry(tk.Frame):
    """Campo de texto con borde de color y placeholder."""
    def __init__(self, parent, placeholder="", textvariable=None, **kwargs):
        super().__init__(parent, bg=Colors.SURFACE,
                         highlightbackground=Colors.BORDER, highlightthickness=1)
        self._ph   = placeholder
        self._tvar = textvariable or tk.StringVar()
        self._entry = tk.Entry(
            self, textvariable=self._tvar,
            relief="flat", bg=Colors.SURFACE, fg=Colors.TEXT,
            font=("Segoe UI", 10), insertbackground=Colors.PRIMARY,
            bd=0, **kwargs
        )
        self._entry.pack(fill="x", padx=10, pady=7)
        self._entry.bind("<FocusIn>",  self._on_focus)
        self._entry.bind("<FocusOut>", self._off_focus)

    def _on_focus(self, _=None):
        self.config(highlightbackground=Colors.BORDER_FOCUS, highlightthickness=2)

    def _off_focus(self, _=None):
        self.config(highlightbackground=Colors.BORDER, highlightthickness=1)

    def get(self):       return self._tvar.get()
    def set(self, v):    self._tvar.set(v)
    @property
    def textvariable(self): return self._tvar


class StatusBar(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=Colors.SURFACE,
                         highlightbackground=Colors.BORDER, highlightthickness=1, **kwargs)
        self._dot = tk.Label(self, text="●", bg=Colors.SURFACE,
                             fg=Colors.OK, font=("Segoe UI", 12))
        self._dot.pack(side="left", padx=(12, 4), pady=6)
        self._lbl = tk.Label(self, text="Creado en los Laboratorios Divisionales de Docencia e Informatica", bg=Colors.SURFACE,
                             fg=Colors.TEXT_MUTED, font=("Segoe UI", 9))
        self._lbl.pack(side="left", pady=6)

    def set(self, text, kind="ok"):
        colors = {"ok": Colors.OK, "warn": Colors.WARN, "error": Colors.DANGER,
                  "busy": Colors.TEXT_MUTED}
        self._dot.config(fg=colors.get(kind, Colors.TEXT_MUTED))
        self._lbl.config(text=text)


# =========================
# APP
# =========================

class App:
    VIEWS = ["prestamo", "devolucion", "historial"]
    NAV_LABELS = [
        ("📤", "Nuevo préstamo",    "prestamo"),
        ("📥", "Registrar entrega", "devolucion"),
        ("📋", "Activos",           "historial"),
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("Préstamo de Equipos Local")
        self.root.geometry("900x640")
        self.root.minsize(820, 580)
        self.root.configure(bg=Colors.BG)

        # Intentar icono de ventana (silencioso si no existe)
        try:
            self.root.iconbitmap("icon.ico")
        except Exception:
            pass

        self.db      = PrestamosDB(Config.JSON_PATH)
        self.service = PrestamoService(self.db)
        self.excel   = Excel(Config.EXCEL_PATH)
        self.devoluciones_map: dict[str, str] = {}
        self._notificados: set[str] = set()
        self._current_view = tk.StringVar(value="prestamo")

        self._build_layout()
        self._build_nav()
        self._build_views()
        self._build_statusbar()
        self._switch_view("prestamo")
        self._monitorear()

    # ─────────────────────────────
    # LAYOUT GENERAL (sidebar + main)
    # ─────────────────────────────
    def _build_layout(self):
        self.sidebar  = tk.Frame(self.root, bg=Colors.SIDEBAR, width=200)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.main_area = tk.Frame(self.root, bg=Colors.BG)
        self.main_area.pack(side="left", fill="both", expand=True)

        # Header del sidebar
        hdr = tk.Frame(self.sidebar, bg=Colors.SIDEBAR, pady=20)
        hdr.pack(fill="x")

        # Logo (opcional). Si no existe el archivo o falla la carga,
        # se usa el emoji como respaldo y la app sigue funcionando.
        try:
            img = Image.open("Lddi.png")
            img = img.resize((150, 100))
            self._logo_img = ImageTk.PhotoImage(img)  # referencia en self: evita que el GC la borre
            tk.Label(hdr, image=self._logo_img, bg=Colors.SIDEBAR).pack()
        except Exception:
            tk.Label(hdr, text="💻", bg=Colors.SIDEBAR, fg=Colors.TEXT_INV,
                     font=("Segoe UI", 28)).pack()

        tk.Label(hdr, text="Préstamos", bg=Colors.SIDEBAR, fg=Colors.TEXT_INV,
                 font=("Segoe UI", 12, "bold")).pack()
        tk.Label(hdr, text="LDDI", bg=Colors.SIDEBAR, fg=Colors.SIDEBAR_TXT,
                 font=("Segoe UI", 9)).pack()

        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", pady=(0, 8))

    def _build_nav(self):
        self._nav_btns = {}
        for icon, label, key in self.NAV_LABELS:
            btn = tk.Frame(self.sidebar, bg=Colors.SIDEBAR, cursor="hand2")
            btn.pack(fill="x", padx=0)
            inner = tk.Frame(btn, bg=Colors.SIDEBAR, padx=16, pady=12)
            inner.pack(fill="x")
            ic = tk.Label(inner, text=icon, bg=Colors.SIDEBAR, fg=Colors.TEXT_INV,
                          font=("Segoe UI", 13))
            ic.pack(side="left")
            lb = tk.Label(inner, text=label, bg=Colors.SIDEBAR, fg=Colors.SIDEBAR_TXT,
                          font=("Segoe UI", 9), anchor="w")
            lb.pack(side="left", padx=10)
            for w in (btn, inner, ic, lb):
                w.bind("<Enter>",          lambda e, b=btn, i=inner, _ic=ic, _lb=lb:
                       self._nav_hover(b, i, _ic, _lb, True))
                w.bind("<Leave>",          lambda e, b=btn, i=inner, _ic=ic, _lb=lb:
                       self._nav_hover(b, i, _ic, _lb, False))
                w.bind("<ButtonRelease-1>",lambda e, k=key: self._switch_view(k))
            self._nav_btns[key] = (btn, inner, ic, lb)

        # Espacio flexible + versión
        tk.Frame(self.sidebar, bg=Colors.SIDEBAR).pack(fill="both", expand=True)
        tk.Label(self.sidebar, text="v2.0", bg=Colors.SIDEBAR, fg=Colors.SIDEBAR_TXT,
                 font=("Segoe UI", 8), pady=12).pack()

    def _nav_hover(self, btn, inner, ic, lb, entering):
        if entering:
            for w in (btn, inner, ic, lb):
                try: w.config(bg=Colors.SIDEBAR_HOV)
                except Exception: pass
        else:
            # Restaurar; respetar activo
            key = next(k for k, v in self._nav_btns.items() if v[0] is btn)
            active = self._current_view.get() == key
            c = Colors.SIDEBAR_ACT if active else Colors.SIDEBAR
            for w in (btn, inner, ic, lb):
                try: w.config(bg=c)
                except Exception: pass

    def _switch_view(self, key):
        self._current_view.set(key)
        # Colorear nav
        for k, (btn, inner, ic, lb) in self._nav_btns.items():
            c = Colors.SIDEBAR_ACT if k == key else Colors.SIDEBAR
            for w in (btn, inner, ic, lb):
                try: w.config(bg=c)
                except Exception: pass
            lb.config(fg=Colors.TEXT_INV if k == key else Colors.SIDEBAR_TXT,
                      font=("Segoe UI", 9, "bold" if k == key else "normal"))
        # Mostrar panel
        for k, panel in self._panels.items():
            if k == key:
                panel.pack(fill="both", expand=True, padx=20, pady=16)
            else:
                panel.pack_forget()
        if key == "historial":
            self._refresh()

    # ─────────────────────────────
    # VISTAS
    # ─────────────────────────────
    def _build_views(self):
        self._panels = {}
        for key in self.VIEWS:
            self._panels[key] = tk.Frame(self.main_area, bg=Colors.BG)

        self._build_prestamo(self._panels["prestamo"])
        self._build_devolucion(self._panels["devolucion"])
        self._build_historial(self._panels["historial"])

    # ── PRÉSTAMO ──────────────────
    def _build_prestamo(self, parent):
        tk.Label(parent, text="Nuevo préstamo", bg=Colors.BG,
                 fg=Colors.TEXT, font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(4, 14))

        card = Card(parent, title="Datos del alumno", accent=Colors.PRIMARY)
        card.pack(fill="x", pady=(0, 12))

        body = tk.Frame(card, bg=Colors.SURFACE, padx=16, pady=12)
        body.pack(fill="x")

        # Matrícula
        tk.Label(body, text="Matrícula", bg=Colors.SURFACE, fg=Colors.TEXT_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.matricula_var = tk.StringVar()
        self.ent_matricula = StyledEntry(body, textvariable=self.matricula_var, width=30)
        self.ent_matricula.grid(row=1, column=0, sticky="ew", padx=(0, 16))

        # Aula
        tk.Label(body, text="Aula", bg=Colors.SURFACE, fg=Colors.TEXT_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w", pady=(0, 2))
        self.salon_var = tk.StringVar()
        self.ent_salon = StyledEntry(body, textvariable=self.salon_var, width=20)
        self.ent_salon.grid(row=1, column=1, sticky="ew")
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=1)

        # Equipos
        card2 = Card(parent, title="Equipos a prestar", accent=Colors.PRIMARY)
        card2.pack(fill="x", pady=(0, 12))

        body2 = tk.Frame(card2, bg=Colors.SURFACE, padx=16, pady=12)
        body2.pack(fill="x")

        # Laptop
        tk.Label(body2, text="Laptop", bg=Colors.SURFACE, fg=Colors.TEXT_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=(0,2))
        self.cb_equipo = self._styled_combo(body2, width=22)
        self.cb_equipo.grid(row=1, column=0, sticky="ew", padx=(0,16))

        # Proyector
        tk.Label(body2, text="Proyector", bg=Colors.SURFACE, fg=Colors.TEXT_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=1, sticky="w", pady=(0,2))
        self.cb_proy = self._styled_combo(body2, width=22)
        self.cb_proy.grid(row=1, column=1, sticky="ew", padx=(0,16))

        # Extras
        tk.Label(body2, text="Extras", bg=Colors.SURFACE, fg=Colors.TEXT_MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w", pady=(0,2))
        extras_f = tk.Frame(body2, bg=Colors.SURFACE)
        extras_f.grid(row=1, column=2, sticky="w")
        self.extras_vars: dict[str, tk.BooleanVar] = {}
        for extra in Config.EXTRAS:
            var = tk.BooleanVar()
            self.extras_vars[extra] = var
            cb = tk.Checkbutton(extras_f, text=extra, variable=var,
                                bg=Colors.SURFACE, fg=Colors.TEXT,
                                activebackground=Colors.SURFACE,
                                selectcolor=Colors.PRIMARY_LT,
                                font=("Segoe UI", 9), cursor="hand2",
                                relief="flat")
            cb.pack(side="left", padx=(0, 8))

        body2.columnconfigure(0, weight=1)
        body2.columnconfigure(1, weight=1)
        body2.columnconfigure(2, weight=1)

        # Nota
        nota = tk.Frame(parent, bg=Colors.WARN_BG,
                        highlightbackground="#F6C342", highlightthickness=1)
        nota.pack(fill="x", pady=(0, 16))
        tk.Label(nota, text="⚠  Se requiere al menos laptop o proyector. "
                 "Un alumno no puede tener dos préstamos activos a la vez.",
                 bg=Colors.WARN_BG, fg=Colors.WARN,
                 font=("Segoe UI", 9), wraplength=560, justify="left",
                 padx=12, pady=8).pack(anchor="w")

        self.btn_prestar = RoundedButton(
            parent, text="✔  Registrar préstamo",
            command=self.prestar, width=240, height=44, font_size=11
        )
        self.btn_prestar.pack(anchor="w")

    # ── DEVOLUCIÓN ────────────────
    def _build_devolucion(self, parent):
        tk.Label(parent, text="Registrar entrega", bg=Colors.BG,
                 fg=Colors.TEXT, font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(4, 14))

        card = Card(parent, title="Préstamo activo", accent=Colors.SECONDARY)
        card.pack(fill="x", pady=(0, 12))

        body = tk.Frame(card, bg=Colors.SURFACE, padx=16, pady=14)
        body.pack(fill="x")

        tk.Label(body, text="Selecciona el préstamo a devolver:", bg=Colors.SURFACE,
                 fg=Colors.TEXT_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 4))
        self.cb_devolucion = self._styled_combo(body, width=55)
        self.cb_devolucion.pack(fill="x")
        self.cb_devolucion.bind("<<ComboboxSelected>>", self._mostrar_info_prestamo)

        # Detalle
        card2 = Card(parent, title="Detalle del préstamo", accent=Colors.SECONDARY)
        card2.pack(fill="x", pady=(0, 16))

        self.info_frame_inner = tk.Frame(card2, bg=Colors.SURFACE, padx=16, pady=14)
        self.info_frame_inner.pack(fill="x")

        self._info_fields = {}
        rows = [
            ("Matrícula", "Aula"),
            ("Laptop",    "Proyector"),
            ("Extras",    "Tiempo restante"),
        ]
        for r, pair in enumerate(rows):
            for c, field in enumerate(pair):
                tk.Label(self.info_frame_inner, text=field, bg=Colors.SURFACE,
                         fg=Colors.TEXT_MUTED, font=("Segoe UI", 9)
                         ).grid(row=r*2, column=c, sticky="w", padx=(0, 30))
                lbl = tk.Label(self.info_frame_inner, text="—", bg=Colors.SURFACE,
                               fg=Colors.TEXT, font=("Segoe UI", 10, "bold"), anchor="w")
                lbl.grid(row=r*2+1, column=c, sticky="w", padx=(0, 30), pady=(0, 8))
                self._info_fields[field] = lbl

        self.btn_devolver = RoundedButton(
            parent, text="↩  Registrar devolución",
            command=self.devolver, width=240, height=44, font_size=11,
            bg="#57A519", hover="#57A519"
        )
        self.btn_devolver.pack(anchor="w")

    # ── HISTORIAL ─────────────────
    def _build_historial(self, parent):
        hdr = tk.Frame(parent, bg=Colors.BG)
        hdr.pack(fill="x", pady=(4, 14))
        tk.Label(hdr, text="Préstamos activos", bg=Colors.BG,
                 fg=Colors.TEXT, font=("Segoe UI", 16, "bold")).pack(side="left")

        # Contador
        self.lbl_count = tk.Label(hdr, text="0 activos", bg=Colors.PRIMARY_LT,
                                  fg=Colors.PRIMARY, font=("Segoe UI", 9, "bold"),
                                  padx=10, pady=4,
                                  highlightbackground=Colors.PRIMARY,
                                  highlightthickness=1)
        self.lbl_count.pack(side="left", padx=14, pady=2)

        # Búsqueda
        busq_f = tk.Frame(parent, bg=Colors.BG)
        busq_f.pack(fill="x", pady=(0, 10))
        tk.Label(busq_f, text="🔍", bg=Colors.BG,
                 font=("Segoe UI", 11)).pack(side="left")
        self.busq_var = tk.StringVar()
        self.busq_var.trace_add("write", lambda *_: self._refresh())
        busq_ent = tk.Entry(busq_f, textvariable=self.busq_var, relief="flat",
                            bg=Colors.SURFACE, fg=Colors.TEXT,
                            font=("Segoe UI", 10), width=20,
                            highlightbackground=Colors.BORDER, highlightthickness=1,
                            insertbackground=Colors.PRIMARY)
        busq_ent.pack(side="left", padx=8, ipady=5, ipadx=6)

        # Treeview
        style = ttk.Style()
        style.configure("Custom.Treeview",
                        background=Colors.SURFACE,
                        fieldbackground=Colors.SURFACE,
                        foreground=Colors.TEXT,
                        rowheight=34,
                        font=("Segoe UI", 9))
        style.configure("Custom.Treeview.Heading",
                        background=Colors.BG,
                        foreground=Colors.TEXT_MUTED,
                        font=("Segoe UI", 9, "bold"),
                        relief="flat")
        style.map("Custom.Treeview", background=[("selected", Colors.PRIMARY_LT)])

        cols = ("Matrícula", "Aula", "Laptop", "Proyector", "Extras", "Salida", "Límite", "Estado")
        self.tree = ttk.Treeview(parent, columns=cols, show="headings",
                                 style="Custom.Treeview", height=11)

        anchos = [90, 70, 110, 120, 130, 100, 100, 110]
        for col, ancho in zip(cols, anchos):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=ancho, anchor="center")

        scroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tree.tag_configure("vencido",    background="#FEECEB", foreground=Colors.DANGER)
        self.tree.tag_configure("por_vencer", background="#FFF8E1", foreground=Colors.WARN)
        self.tree.tag_configure("ok",         background="#EAF6EE", foreground=Colors.OK)

    # ─────────────────────────────
    # HELPERS UI
    # ─────────────────────────────
    def _styled_combo(self, parent, width=30) -> ttk.Combobox:
        style = ttk.Style()
        style.configure("Flat.TCombobox",
                        fieldbackground=Colors.SURFACE,
                        background=Colors.SURFACE,
                        foreground=Colors.TEXT,
                        selectbackground=Colors.PRIMARY_LT,
                        relief="flat")
        cb = ttk.Combobox(parent, state="readonly", width=width,
                          font=("Segoe UI", 10), style="Flat.TCombobox")
        return cb

    def _build_statusbar(self):
        self.status_bar = StatusBar(self.main_area)
        self.status_bar.pack(fill="x", side="bottom", padx=20, pady=(0, 12))

    # ─────────────────────────────
    # LÓGICA / ACCIONES
    # ─────────────────────────────
    def _set_busy(self, ocupado: bool):
        self.btn_prestar.set_enabled(not ocupado)
        self.btn_devolver.set_enabled(not ocupado)
        if ocupado:
            self.status_bar.set("Guardando…", "busy")
        self.root.update_idletasks()

    def _refresh(self):
        activos = self.db.activos()
        ahora   = datetime.datetime.now()
        filtro  = self.busq_var.get().strip() if hasattr(self, "busq_var") else ""

        equipos_ocupados = {p["equipo"]    for p in activos}
        proy_ocupados    = {p["proyector"] for p in activos}

        self.cb_equipo["values"] = ["(ninguna)"] + [e for e in Config.EQUIPOS
                                                     if e not in equipos_ocupados]
        self.cb_proy["values"]   = ["(ninguno)"] + [p for p in Config.PROYECTORES
                                                     if p not in proy_ocupados]
        if self.cb_equipo.get() not in self.cb_equipo["values"]:
            self.cb_equipo.set("(ninguna)")
        if self.cb_proy.get() not in self.cb_proy["values"]:
            self.cb_proy.set("(ninguno)")

        # Combo devolución
        self.devoluciones_map.clear()
        opciones = []
        for p in activos:
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = int((limite - ahora).total_seconds() / 60)
            aviso    = (f"⚠ {restante} min" if 0 < restante <= Config.ALERTA_MINUTOS
                        else ("🔴 VENCIDO" if restante <= 0 else ""))
            txt = f'{p["matricula"]} → {p["equipo"] or "-"} / {p["proyector"] or "-"}  {aviso}'
            opciones.append(txt)
            self.devoluciones_map[txt] = p["id"]

        self.cb_devolucion["values"] = opciones
        if self.cb_devolucion.get() not in opciones:
            self.cb_devolucion.set("")
            self._clear_info()

        # Árbol
        for item in self.tree.get_children():
            self.tree.delete(item)
        total = 0
        for p in activos:
            if filtro and filtro not in p["matricula"]:
                continue
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = (limite - ahora).total_seconds() / 60
            tag      = ("vencido"    if restante <= 0
                        else "por_vencer" if restante <= Config.ALERTA_MINUTOS
                        else "ok")
            self.tree.insert("", "end", values=(
                p["matricula"], p["salon"],
                p["equipo"] or "—", p["proyector"] or "—",
                p.get("extras", "") or "—",
                p["hora_salida"][11:16],
                p["hora_limite"][11:16],
                "⏰ VENCIDO" if restante <= 0 else "● ACTIVO",
            ), tags=(tag,))
            total += 1

        if hasattr(self, "lbl_count"):
            self.lbl_count.config(text=f"{total} activo{'s' if total != 1 else ''}")

    def _clear_info(self):
        defaults = {
            "Matrícula": "—", "Aula": "—", "Laptop": "—",
            "Proyector": "—", "Extras": "—", "Tiempo restante": "—"
        }
        for k, lbl in self._info_fields.items():
            lbl.config(text=defaults.get(k, "—"), fg=Colors.TEXT)

    def _mostrar_info_prestamo(self, _=None):
        sel = self.cb_devolucion.get()
        pid = self.devoluciones_map.get(sel)
        if not pid:
            return
        p = next((x for x in self.db.activos() if x["id"] == pid), None)
        if not p:
            return

        limite   = datetime.datetime.fromisoformat(p["hora_limite"])
        ahora    = datetime.datetime.now()
        restante = int((limite - ahora).total_seconds() / 60)
        if restante > 0:
            tiempo_txt = f"{restante} min restantes"
            tiempo_col = Colors.OK if restante > Config.ALERTA_MINUTOS else Colors.WARN
        else:
            tiempo_txt = "⏰ Tiempo vencido"
            tiempo_col = Colors.DANGER

        self._info_fields["Matrícula"].config(text=p["matricula"])
        self._info_fields["Aula"].config(text=p["salon"])
        self._info_fields["Laptop"].config(text=p["equipo"] or "—")
        self._info_fields["Proyector"].config(text=p["proyector"] or "—")
        self._info_fields["Extras"].config(text=p.get("extras", "") or "—")
        self._info_fields["Tiempo restante"].config(text=tiempo_txt, fg=tiempo_col)

    def _monitorear(self):
        activos    = self.db.activos()
        ahora      = datetime.datetime.now()
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
                    f"({p['equipo'] or '—'} / {p['proyector'] or '—'}) ha vencido."
                )

        self._refresh()
        self.root.after(60_000, self._monitorear)

    # ─────────────────────────────
    # ACCIONES PRINCIPALES
    # ─────────────────────────────
    def prestar(self):
        matricula = self.matricula_var.get().strip()
        salon     = self.salon_var.get().strip()
        equipo    = self.cb_equipo.get()
        proy      = self.cb_proy.get()
        equipo    = "" if equipo == "(ninguna)" else equipo
        proy      = "" if proy   == "(ninguno)" else proy
        extras    = [k for k, v in self.extras_vars.items() if v.get()]

        error = self.service.validar(matricula, salon, equipo, proy)
        if error:
            messagebox.showerror("Error de validación", error)
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
        finally:
            self._set_busy(False)

        # Reset
        self.matricula_var.set("")
        self.salon_var.set("")
        for v in self.extras_vars.values():
            v.set(False)

        limite = datetime.datetime.fromisoformat(prestamo["hora_limite"])
        self.status_bar.set(
            f"Préstamo registrado — devolver antes de las {limite.strftime('%H:%M')}", "ok")
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
        if not messagebox.askokcancel("Confirmar devolución", "¿Registrar la devolución?"):
            return

        self._set_busy(True)
        try:
            p = self.service.registrar_devolucion(prestamo_id)
        finally:
            self._set_busy(False)

        if not p:
            messagebox.showerror("Error", "No se encontró el préstamo.")
            return

        self.excel.actualizar_o_agregar(p)

        if p["estado"] == "DEVUELTO TARDE":
            messagebox.showwarning("Entrega tardía",
                "El equipo fue devuelto fuera del tiempo límite.")

        self.cb_devolucion.set("")
        self._clear_info()
        self.status_bar.set(f"Devolución registrada — {p['estado']}", "ok")
        self._refresh()


# =========================
# MAIN
# =========================

if __name__ == "__main__":
    root = tk.Tk()

    # Fuente global Segoe UI (si está disponible)
    try:
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        root.option_add("*Font", default_font)
    except Exception:
        pass

    app = App(root)
    root.mainloop()