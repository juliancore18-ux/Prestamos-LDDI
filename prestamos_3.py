import os
import sys
import json
import uuid
import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QCheckBox, QPushButton, QTabWidget,
    QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QAbstractItemView, QFrame, QGraphicsDropShadowEffect,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QColor, QBrush

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================
# SISTEMA DE PRÉSTAMOS — LDDI (versión PyQt6)
# Solo se modificó la capa visual (Theme + construcción de UI).
# La lógica de negocio (DB, servicio, Excel) permanece intacta.
# =============================================================


# =============================================================
# TEMA / ESTILOS CENTRALIZADOS
# =============================================================

class Theme:
    # Paleta
    BG            = "#F4F6F8"
    SURFACE       = "#FFFFFF"
    BORDER        = "#E1E4E8"

    PRIMARY       = "#2E8B47"
    PRIMARY_HOVER = "#279140"
    PRIMARY_PRESS = "#1F7A36"
    PRIMARY_DARK  = "#1B5E20"

    ACCENT        = "#1565C0"

    DANGER        = "#D32F2F"
    WARN          = "#F9A825"
    WARN_BG       = "#FFF8E1"
    OK_BG         = "#E8F5E9"
    LATE_BG       = "#FFEBEE"

    FG_TITLE      = "#FFFFFF"
    FG_MUTED      = "#6B7280"
    FG_HINT       = "#9AA1A9"
    FG_BODY       = "#2C3338"
    FG_HEAD       = "#1A1D21"

    STATUS_BG     = "#EEF1F4"

    FONT_FAMILY   = "Segoe UI"

    FONT_TITLE  = (FONT_FAMILY, 15, "bold")
    FONT_HEAD   = (FONT_FAMILY, 12, "bold")
    FONT_BODY   = (FONT_FAMILY, 10)
    FONT_SMALL  = (FONT_FAMILY, 9)
    FONT_HINT   = (FONT_FAMILY, 8, "italic")

    @staticmethod
    def qss(font_family: str) -> str:
        return f"""
        * {{
            font-family: '{font_family}';
        }}

        QMainWindow, QWidget#central {{
            background-color: {Theme.BG};
        }}

        /* ---------- Tabs ---------- */
        QTabWidget::pane {{
            border: none;
            background-color: {Theme.BG};
        }}
        QTabBar::tab {{
            background-color: transparent;
            color: {Theme.FG_MUTED};
            padding: 10px 18px;
            margin-right: 4px;
            font-size: 10.5pt;
            font-weight: 600;
            border-bottom: 3px solid transparent;
        }}
        QTabBar::tab:selected {{
            color: {Theme.PRIMARY_DARK};
            border-bottom: 3px solid {Theme.PRIMARY};
        }}
        QTabBar::tab:hover:!selected {{
            color: {Theme.FG_BODY};
        }}

        /* ---------- Labels ---------- */
        QLabel {{
            color: {Theme.FG_BODY};
            font-size: 10pt;
        }}

        /* ---------- Inputs ---------- */
        QLineEdit, QComboBox {{
            background-color: {Theme.SURFACE};
            border: 1.5px solid {Theme.BORDER};
            border-radius: 6px;
            padding: 6px 10px;
            font-size: 10pt;
            color: {Theme.FG_BODY};
            selection-background-color: {Theme.PRIMARY};
        }}
        QLineEdit:focus, QComboBox:focus {{
            border: 1.5px solid {Theme.PRIMARY};
        }}
        QLineEdit:hover, QComboBox:hover {{
            border: 1.5px solid #B9C0C7;
        }}
        QComboBox::drop-down {{
            border: none;
            width: 24px;
        }}
        QComboBox QAbstractItemView {{
            border: 1px solid {Theme.BORDER};
            selection-background-color: {Theme.PRIMARY};
            selection-color: white;
            background-color: {Theme.SURFACE};
            outline: none;
        }}

        /* ---------- Checkboxes ---------- */
        QCheckBox {{
            font-size: 10pt;
            color: {Theme.FG_BODY};
            spacing: 6px;
        }}
        QCheckBox::indicator {{
            width: 16px;
            height: 16px;
            border: 1.5px solid {Theme.BORDER};
            border-radius: 4px;
            background-color: {Theme.SURFACE};
        }}
        QCheckBox::indicator:checked {{
            background-color: {Theme.PRIMARY};
            border: 1.5px solid {Theme.PRIMARY};
        }}

        /* ---------- Buttons ---------- */
        QPushButton#btnPrimary {{
            background-color: {Theme.PRIMARY};
            color: white;
            font-weight: 600;
            font-size: 10pt;
            padding: 10px 26px;
            border: none;
            border-radius: 6px;
        }}
        QPushButton#btnPrimary:hover  {{ background-color: {Theme.PRIMARY_HOVER}; }}
        QPushButton#btnPrimary:pressed {{ background-color: {Theme.PRIMARY_PRESS}; }}

        QPushButton#btnSecondary {{
            background-color: {Theme.ACCENT};
            color: white;
            font-weight: 600;
            font-size: 10pt;
            padding: 10px 26px;
            border: none;
            border-radius: 6px;
        }}
        QPushButton#btnSecondary:hover  {{ background-color: #0D47A1; }}
        QPushButton#btnSecondary:pressed {{ background-color: #09367A; }}

        /* ---------- GroupBox ---------- */
        QGroupBox {{
            background-color: {Theme.SURFACE};
            border: 1px solid {Theme.BORDER};
            border-radius: 8px;
            margin-top: 14px;
            padding: 14px;
            font-weight: 600;
            color: {Theme.FG_HEAD};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: {Theme.PRIMARY_DARK};
        }}

        /* ---------- Table ---------- */
        QTableWidget {{
            background-color: {Theme.SURFACE};
            border: 1px solid {Theme.BORDER};
            border-radius: 8px;
            gridline-color: {Theme.BORDER};
            font-size: 9.5pt;
        }}
        QHeaderView::section {{
            background-color: {Theme.FG_HEAD};
            color: white;
            padding: 8px;
            border: none;
            font-weight: 600;
            font-size: 9.5pt;
        }}
        QTableWidget::item {{
            padding: 4px;
        }}
        QTableWidget::item:selected {{
            background-color: #CFE8D6;
            color: {Theme.FG_BODY};
        }}

        /* ---------- Scrollbars ---------- */
        QScrollBar:vertical {{
            background: {Theme.BG};
            width: 10px;
            margin: 0px;
        }}
        QScrollBar::handle:vertical {{
            background: #C7CDD3;
            border-radius: 5px;
            min-height: 24px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: #AEB6BD;
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0px;
        }}
        """


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
    JSON_PATH  = os.path.join(BASE_DIR, "prestamos.json")
    EXCEL_PATH = os.path.join(BASE_DIR, "PRESTAMOS.xlsx")
    HEADERS    = [
        "ID", "Matrícula", "Aula",
        "Laptop", "Proyector", "Extras",
        "Salida", "Límite", "Devolución",
        "Estado",
    ]

    # Opciones vacías para los comboboxes
    SIN_LAPTOP    = "(ninguna)"
    SIN_PROYECTOR = "(ninguno)"


# =============================================================
# DATA
# =============================================================

class PrestamosDB:
    def __init__(self, path: str):
        self.path = path
        self.data: dict = {"prestamos": []}
        self._load()

    def _load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
        except json.JSONDecodeError as e:
            QMessageBox.critical(
                None,
                "Error de datos",
                f"El archivo JSON está corrupto:\n{e}\n"
                "Se iniciará con datos vacíos.",
            )
            self.data = {"prestamos": []}
        except IOError as e:
            QMessageBox.critical(
                None,
                "Error de archivo",
                f"No se pudo leer el archivo de datos:\n{e}",
            )
            self.data = {"prestamos": []}

    def save(self):
        """Escritura atómica para evitar corrupción."""
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=4, ensure_ascii=False)
            os.replace(tmp, self.path)
        except IOError as e:
            QMessageBox.critical(None, "Error al guardar", f"No se pudo guardar el archivo:\n{e}")
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    def _reload(self):
        """Recarga desde disco antes de operaciones críticas."""
        self._load()

    def activos(self) -> list:
        self._reload()
        return [p for p in self.data["prestamos"] if p["estado"] == "PRESTADO"]

    def todos(self) -> list:
        self._reload()
        return self.data["prestamos"]

    def agregar(self, prestamo: dict):
        self._reload()
        ids_existentes = {p["id"] for p in self.data["prestamos"]}
        nuevo_id = str(uuid.uuid4())
        while nuevo_id in ids_existentes:
            nuevo_id = str(uuid.uuid4())
        prestamo["id"] = nuevo_id
        self.data["prestamos"].append(prestamo)
        self.save()

    def devolver(self, prestamo_id: str, estado_final: str, hora_devolucion: str) -> dict | None:
        self._reload()
        for p in self.data["prestamos"]:
            if p["id"] == prestamo_id:
                p["estado"]          = estado_final
                p["hora_devolucion"] = hora_devolucion
                self.save()
                return p
        return None

    def alumno_tiene_prestamo_activo(self, matricula: str) -> bool:
        return any(p["matricula"] == matricula for p in self.activos())


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
            "id":              "",
            "matricula":       matricula,
            "salon":           salon,
            "equipo":          equipo,
            "proyector":       proyector,
            "extras":          ", ".join(extras),
            "hora_salida":     ahora.isoformat(),
            "hora_limite":     limite.isoformat(),
            "hora_devolucion": "",
            "estado":          "PRESTADO",
        }
        self.db.agregar(prestamo)
        return prestamo

    def registrar_devolucion(self, prestamo_id: str) -> dict | None:
        """Calcula el estado ANTES de persistir."""
        prestamo = next(
            (p for p in self.db.activos() if p["id"] == prestamo_id), None
        )
        if not prestamo:
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
            cell           = ws.cell(row=1, column=col)
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = header_border

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
            QMessageBox.critical(
                None,
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
            QMessageBox.critical(
                None,
                "Archivo bloqueado",
                f"No se pudo guardar '{self.path}'.\n"
                "Cierra el archivo en Excel e intenta de nuevo.",
            )


# =============================================================
# APP
# =============================================================

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Préstamos de Equipos")
        self.setFixedSize(720, 740)
        self.setStyleSheet(Theme.qss(Theme.FONT_FAMILY))

        self.db      = PrestamosDB(Config.JSON_PATH)
        self.service = PrestamoService(self.db)
        self.excel   = Excel(Config.EXCEL_PATH)

        self.devoluciones_map: dict[str, str] = {}
        self._notificados: set[str]           = set()

        self._ui()
        self._monitorear()

    # ----------------------------------------------------------
    # CONSTRUCCIÓN DE UI
    # ----------------------------------------------------------

    def _ui(self):
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ---- Cabecera con degradado sutil ----
        header = QFrame()
        header.setStyleSheet(
            f"background-color: {Theme.PRIMARY_DARK};"
        )
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(24, 16, 24, 16)
        header_layout.setSpacing(2)

        titulo = QLabel("Sistema de Préstamos de Equipos")
        titulo.setFont(QFont(Theme.FONT_FAMILY, 15, QFont.Weight.Bold))
        titulo.setStyleSheet("color: white; background: transparent;")
        header_layout.addWidget(titulo)

        subtitulo = QLabel("Laptops · Proyectores · Accesorios")
        subtitulo.setFont(QFont(Theme.FONT_FAMILY, 9))
        subtitulo.setStyleSheet("color: rgba(255,255,255,0.75); background: transparent;")
        header_layout.addWidget(subtitulo)

        main_layout.addWidget(header)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        main_layout.addWidget(self.tabs, 1)

        self.tab_prestamo   = QWidget()
        self.tab_devolucion = QWidget()
        self.tab_historial  = QWidget()

        self.tabs.addTab(self.tab_prestamo,   "📤  Préstamo")
        self.tabs.addTab(self.tab_devolucion, "📥  Devolución")
        self.tabs.addTab(self.tab_historial,  "📋  Activos")

        self._ui_prestamo()
        self._ui_devolucion()
        self._ui_historial()

        self.lbl_estado = QLabel("Listo")
        self.lbl_estado.setFont(QFont(Theme.FONT_FAMILY, 9))
        self.lbl_estado.setStyleSheet(
            f"background-color: {Theme.STATUS_BG}; color: {Theme.FG_BODY}; "
            f"padding: 8px 16px; border-top: 1px solid {Theme.BORDER};"
        )
        main_layout.addWidget(self.lbl_estado)

        self._refresh()

    def _card(self) -> tuple[QFrame, QVBoxLayout]:
        """Crea un contenedor tipo 'tarjeta' con sombra suave."""
        card = QFrame()
        card.setStyleSheet(
            f"QFrame {{ background-color: {Theme.SURFACE}; border: 1px solid {Theme.BORDER}; "
            "border-radius: 10px; }}"
        )
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 30))
        card.setGraphicsEffect(shadow)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(20, 18, 20, 18)
        card_layout.setSpacing(12)
        return card, card_layout

    def _seccion_titulo(self, texto: str) -> QLabel:
        lbl = QLabel(texto)
        lbl.setFont(QFont(Theme.FONT_FAMILY, 12, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color: {Theme.PRIMARY_DARK}; background: transparent;")
        return lbl

    def _ui_prestamo(self):
        f = self.tab_prestamo
        outer = QVBoxLayout(f)
        outer.setContentsMargins(20, 18, 20, 12)
        outer.setSpacing(14)

        card, layout = self._card()
        outer.addWidget(card)

        layout.addWidget(self._seccion_titulo("Nuevo préstamo"))

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(12)
        layout.addLayout(form)

        self.matricula_edit = QLineEdit()
        self.matricula_edit.setFixedWidth(220)
        self.matricula_edit.setPlaceholderText("Ej. 1234567")
        self.salon_edit = QLineEdit()
        self.salon_edit.setFixedWidth(220)
        self.salon_edit.setPlaceholderText("Ej. A-201")

        lbl_mat = QLabel(f"Matrícula ({Config.MATRICULA_MIN}–{Config.MATRICULA_MAX} dígitos)")
        form.addWidget(lbl_mat, 0, 0)
        form.addWidget(self.matricula_edit, 0, 1)

        lbl_salon = QLabel("Aula")
        form.addWidget(lbl_salon, 1, 0)
        form.addWidget(self.salon_edit, 1, 1)

        lbl_eq = QLabel("Laptop")
        form.addWidget(lbl_eq, 2, 0)
        self.cb_equipo = QComboBox()
        self.cb_equipo.setFixedWidth(220)
        form.addWidget(self.cb_equipo, 2, 1)

        lbl_proy = QLabel("Proyector")
        form.addWidget(lbl_proy, 3, 0)
        self.cb_proy = QComboBox()
        self.cb_proy.setFixedWidth(220)
        form.addWidget(self.cb_proy, 3, 1)

        lbl_extras = QLabel("Extras")
        form.addWidget(lbl_extras, 4, 0, Qt.AlignmentFlag.AlignTop)

        extras_widget = QWidget()
        extras_layout = QHBoxLayout(extras_widget)
        extras_layout.setContentsMargins(0, 0, 0, 0)
        extras_layout.setSpacing(16)
        self.extras_checks: dict[str, QCheckBox] = {}
        for extra in Config.EXTRAS:
            chk = QCheckBox(extra.title())
            self.extras_checks[extra] = chk
            extras_layout.addWidget(chk)
        extras_layout.addStretch()
        form.addWidget(extras_widget, 4, 1)

        aviso = QLabel(
            "⚠  Selecciona al menos una laptop o un proyector.\n"
            "Un alumno no puede tener dos préstamos activos simultáneos."
        )
        aviso.setFont(QFont(Theme.FONT_FAMILY, 9))
        aviso.setStyleSheet(f"color: {Theme.FG_MUTED}; background: transparent;")
        layout.addWidget(aviso)

        self.btn_prestar = QPushButton("Registrar préstamo")
        self.btn_prestar.setObjectName("btnPrimary")
        self.btn_prestar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_prestar.clicked.connect(self.prestar)
        layout.addWidget(self.btn_prestar, alignment=Qt.AlignmentFlag.AlignLeft)

        outer.addStretch()

        creditos_font = QFont(Theme.FONT_FAMILY, 8)
        creditos_font.setItalic(True)
        creditos = QLabel("CREADO EN LDDI")
        creditos.setFont(creditos_font)
        creditos.setStyleSheet(f"color: {Theme.FG_HINT}; background: transparent;")
        outer.addWidget(creditos, alignment=Qt.AlignmentFlag.AlignHCenter)

    def _ui_devolucion(self):
        f = self.tab_devolucion
        outer = QVBoxLayout(f)
        outer.setContentsMargins(20, 18, 20, 12)
        outer.setSpacing(14)

        card, layout = self._card()
        outer.addWidget(card)

        layout.addWidget(self._seccion_titulo("Registrar devolución"))

        layout.addWidget(QLabel("Selecciona el préstamo activo:"))

        self.cb_devolucion = QComboBox()
        layout.addWidget(self.cb_devolucion)
        self.cb_devolucion.currentTextChanged.connect(self._mostrar_info_prestamo)

        self.btn_devolver = QPushButton("Registrar devolución")
        self.btn_devolver.setObjectName("btnSecondary")
        self.btn_devolver.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_devolver.clicked.connect(self.devolver)
        layout.addWidget(self.btn_devolver, alignment=Qt.AlignmentFlag.AlignLeft)

        self.info_group = QGroupBox("Información del préstamo")
        info_layout = QVBoxLayout(self.info_group)
        self.info_label = QLabel("Selecciona un préstamo para ver detalles.")
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet("background: transparent;")
        info_layout.addWidget(self.info_label)
        outer.addWidget(self.info_group)
        outer.addStretch()

    def _ui_historial(self):
        f = self.tab_historial
        outer = QVBoxLayout(f)
        outer.setContentsMargins(20, 18, 20, 12)
        outer.setSpacing(10)

        header_row = QHBoxLayout()
        header_row.addWidget(self._seccion_titulo("Préstamos activos"))
        header_row.addStretch()
        header_row.addWidget(QLabel("Buscar matrícula:"))
        self.busq_edit = QLineEdit()
        self.busq_edit.setFixedWidth(160)
        self.busq_edit.setPlaceholderText("Ej. 1234567")
        self.busq_edit.textChanged.connect(self._refresh)
        header_row.addWidget(self.busq_edit)
        outer.addLayout(header_row)

        cols = ("Matrícula", "Aula", "Laptop", "Proyector", "Extras", "Salida", "Límite", "Estado")
        self.tabla = QTableWidget(0, len(cols))
        self.tabla.setHorizontalHeaderLabels(cols)
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabla.verticalHeader().setVisible(False)
        self.tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabla.setAlternatingRowColors(False)
        self.tabla.setShowGrid(True)
        outer.addWidget(self.tabla)

    # ----------------------------------------------------------
    # LÓGICA INTERNA
    # ----------------------------------------------------------

    def _set_busy(self, ocupado: bool):
        """Deshabilita botones durante operaciones para evitar doble clic."""
        self.btn_prestar.setEnabled(not ocupado)
        self.btn_devolver.setEnabled(not ocupado)
        if ocupado:
            self.lbl_estado.setText("⏳  Guardando...")
            self.lbl_estado.setStyleSheet(
                f"background-color: {Theme.STATUS_BG}; color: #555555; "
                f"padding: 8px 16px; border-top: 1px solid {Theme.BORDER};"
            )
        QApplication.processEvents()

    def _refresh(self):
        activos = self.db.activos()
        ahora   = datetime.datetime.now()
        filtro  = self.busq_edit.text().strip() if hasattr(self, "busq_edit") else ""

        equipos_ocupados = {p["equipo"]    for p in activos}
        proy_ocupados    = {p["proyector"] for p in activos}

        # Poblar comboboxes con equipos disponibles, preservando selección actual
        equipo_actual = self.cb_equipo.currentText()
        proy_actual   = self.cb_proy.currentText()

        valores_equipo = [Config.SIN_LAPTOP]    + [e for e in Config.EQUIPOS    if e not in equipos_ocupados]
        valores_proy   = [Config.SIN_PROYECTOR] + [p for p in Config.PROYECTORES if p not in proy_ocupados]

        self.cb_equipo.blockSignals(True)
        self.cb_equipo.clear()
        self.cb_equipo.addItems(valores_equipo)
        self.cb_equipo.setCurrentText(equipo_actual if equipo_actual in valores_equipo else Config.SIN_LAPTOP)
        self.cb_equipo.blockSignals(False)

        self.cb_proy.blockSignals(True)
        self.cb_proy.clear()
        self.cb_proy.addItems(valores_proy)
        self.cb_proy.setCurrentText(proy_actual if proy_actual in valores_proy else Config.SIN_PROYECTOR)
        self.cb_proy.blockSignals(False)

        # Opciones de devolución
        seleccion_actual = self.cb_devolucion.currentText()
        self.devoluciones_map.clear()
        opciones = []
        for p in activos:
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = int((limite - ahora).total_seconds() / 60)
            aviso = (
                f"⚠ {restante} min" if 0 < restante <= Config.ALERTA_MINUTOS
                else ("🔴 VENCIDO" if restante <= 0 else "")
            )
            txt = f'{p["matricula"]} → {p["equipo"] or "-"} / {p["proyector"] or "-"}  {aviso}'
            opciones.append(txt)
            self.devoluciones_map[txt] = p["id"]

        self.cb_devolucion.blockSignals(True)
        self.cb_devolucion.clear()
        self.cb_devolucion.addItems(opciones)
        if seleccion_actual in opciones:
            self.cb_devolucion.setCurrentText(seleccion_actual)
        else:
            self.cb_devolucion.setCurrentText("")
            self.info_label.setText("Selecciona un préstamo para ver detalles.")
        self.cb_devolucion.blockSignals(False)

        if not opciones:
            self.info_label.setText("Sin préstamos activos.")

        # Tabla con filtro por matrícula
        filas = [p for p in activos if not filtro or filtro in p["matricula"]]
        self.tabla.setRowCount(len(filas))

        color_map = {
            "vencido":    QColor(Theme.LATE_BG),
            "por_vencer": QColor(Theme.WARN_BG),
            "ok":         QColor(Theme.OK_BG),
        }

        for row, p in enumerate(filas):
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = (limite - ahora).total_seconds() / 60
            tag = (
                "vencido"    if restante <= 0
                else "por_vencer" if restante <= Config.ALERTA_MINUTOS
                else "ok"
            )
            valores = (
                p["matricula"],
                p["salon"],
                p["equipo"]    or "-",
                p["proyector"] or "-",
                p.get("extras", "") or "-",
                p["hora_salida"][11:16],
                p["hora_limite"][11:16],
                "VENCIDO" if restante <= 0 else "PRESTADO",
            )
            color = color_map[tag]
            for col, valor in enumerate(valores):
                item = QTableWidgetItem(str(valor))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setBackground(QBrush(color))
                self.tabla.setItem(row, col, item)

    def _monitorear(self):
        """Revisa vencimientos cada minuto; notifica una sola vez por préstamo."""
        activos     = self.db.activos()
        ahora       = datetime.datetime.now()
        ids_activos = {p["id"] for p in activos}

        # Limpiar IDs ya devueltos del set de notificados
        self._notificados &= ids_activos

        for p in activos:
            limite   = datetime.datetime.fromisoformat(p["hora_limite"])
            restante = (limite - ahora).total_seconds() / 60
            if restante <= 0 and p["id"] not in self._notificados:
                self._notificados.add(p["id"])
                QMessageBox.warning(
                    self,
                    "Préstamo vencido",
                    f"El préstamo de la matrícula {p['matricula']}\n"
                    f"({p['equipo'] or '-'} / {p['proyector'] or '-'}) ha vencido.",
                )

        self._refresh()
        QTimer.singleShot(60_000, self._monitorear)

    def _mostrar_info_prestamo(self, _texto: str = None):
        sel = self.cb_devolucion.currentText()
        pid = self.devoluciones_map.get(sel)
        if not pid:
            return
        p = next((x for x in self.db.activos() if x["id"] == pid), None)
        if not p:
            return

        limite   = datetime.datetime.fromisoformat(p["hora_limite"])
        ahora    = datetime.datetime.now()
        restante = int((limite - ahora).total_seconds() / 60)
        estado_txt = f"⏱ {restante} min restantes" if restante > 0 else "🔴 TIEMPO VENCIDO"

        self.info_label.setText(
            f"Matrícula: {p['matricula']}    Aula: {p['salon']}\n"
            f"Laptop: {p['equipo'] or 'N/A'}    Proyector: {p['proyector'] or 'N/A'}\n"
            f"Extras: {p.get('extras', '') or 'N/A'}\n"
            f"Salida: {p['hora_salida'][11:16]}    "
            f"Límite: {p['hora_limite'][11:16]}    {estado_txt}"
        )

    # ----------------------------------------------------------
    # ACCIONES
    # ----------------------------------------------------------

    def prestar(self):
        matricula = self.matricula_edit.text().strip()
        salon     = self.salon_edit.text().strip()
        equipo    = self.cb_equipo.currentText()
        proy      = self.cb_proy.currentText()

        equipo = "" if equipo == Config.SIN_LAPTOP    else equipo
        proy   = "" if proy   == Config.SIN_PROYECTOR else proy
        extras = [k for k, v in self.extras_checks.items() if v.isChecked()]

        error = self.service.validar(matricula, salon, equipo, proy)
        if error:
            QMessageBox.critical(self, "Error", error)
            return

        resumen = (
            f"Matrícula: {matricula}\n"
            f"Aula: {salon}\n"
            f"Laptop: {equipo or 'N/A'}\n"
            f"Proyector: {proy or 'N/A'}\n"
            f"Extras: {', '.join(extras) or 'N/A'}\n\n"
            "¿Confirmar préstamo?"
        )
        resp = QMessageBox.question(
            self, "Confirmar préstamo", resumen,
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )
        if resp != QMessageBox.StandardButton.Ok:
            return

        self._set_busy(True)
        try:
            prestamo = self.service.crear_prestamo(matricula, salon, equipo, proy, extras)
            self.excel.actualizar_o_agregar(prestamo)
        finally:
            self._set_busy(False)

        # Limpiar formulario
        self.matricula_edit.clear()
        self.salon_edit.clear()
        for chk in self.extras_checks.values():
            chk.setChecked(False)

        limite = datetime.datetime.fromisoformat(prestamo["hora_limite"])
        self.lbl_estado.setText(f"✔  Préstamo registrado. Límite: {limite.strftime('%H:%M')}")
        self.lbl_estado.setStyleSheet(
            f"background-color: {Theme.STATUS_BG}; color: {Theme.PRIMARY_DARK}; "
            f"padding: 8px 16px; border-top: 1px solid {Theme.BORDER};"
        )
        self._refresh()

    def devolver(self):
        seleccion = self.cb_devolucion.currentText()
        if not seleccion:
            QMessageBox.critical(self, "Error", "Selecciona un préstamo activo.")
            return

        prestamo_id = self.devoluciones_map.get(seleccion)
        if not prestamo_id:
            QMessageBox.critical(self, "Error", "Préstamo no encontrado.")
            return

        resp = QMessageBox.question(
            self, "Confirmar devolución", "¿Registrar la devolución de este equipo?",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
        )
        if resp != QMessageBox.StandardButton.Ok:
            return

        self._set_busy(True)
        try:
            p = self.service.registrar_devolucion(prestamo_id)
        finally:
            self._set_busy(False)

        if not p:
            QMessageBox.critical(self, "Error", "No se encontró el préstamo en la base de datos.")
            return

        self.excel.actualizar_o_agregar(p)

        if p["estado"] == "DEVUELTO TARDE":
            QMessageBox.warning(
                self, "Entrega tardía", "El equipo fue devuelto fuera del tiempo límite."
            )

        self.cb_devolucion.setCurrentText("")
        self.info_label.setText("Selecciona un préstamo para ver detalles.")
        self.lbl_estado.setText(f"✔  Devolución registrada ({p['estado']})")
        self.lbl_estado.setStyleSheet(
            f"background-color: {Theme.STATUS_BG}; color: {Theme.PRIMARY}; "
            f"padding: 8px 16px; border-top: 1px solid {Theme.BORDER};"
        )
        self._refresh()


# =============================================================
# MAIN
# =============================================================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())