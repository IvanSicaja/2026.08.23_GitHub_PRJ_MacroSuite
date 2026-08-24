#!/usr/bin/env python3
"""
MacroSuite — Nutrition Planning Software
Dark Apple-style UI · Excel database · Automatic backup · Persistent settings.
"""

import sys
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from dataclasses import dataclass, field
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QTableWidget, QTableWidgetItem, QPushButton, QLabel,
    QLineEdit, QDialog, QFormLayout, QComboBox, QDoubleSpinBox,
    QMessageBox, QCompleter, QHeaderView, QListWidget, QFrame,
    QStatusBar, QFileDialog, QAbstractItemView, QGridLayout,
    QInputDialog,
)
from PySide6.QtCore import Qt, QTimer, Signal, QSettings, QSize, QEvent, QObject, QPointF
from PySide6.QtGui import (
    QFont as QFontGui, QColor, QAction, QPixmap, QPainter, QPolygonF,
)


# ━━━━━━━━━━━━━━━━━━━━━━ CONSTANTS ━━━━━━━━━━━━━━━━━━━━━━

APP_TITLE = "MacroSuite — Nutrition Planning Software"

NUTRITION_FIELDS = [
    ("energy_kj",    "Energy (kJ)"),
    ("energy_kcal",  "Energy (kcal)"),
    ("fat",          "Fat (g)"),
    ("saturated_fat","Sat. Fat (g)"),
    ("carbohydrate", "Carbs (g)"),
    ("sugars",       "Sugars (g)"),
    ("fibre",        "Fibre (g)"),
    ("protein",      "Protein (g)"),
    ("salt",         "Salt (g)"),
]
NUTRITION_KEYS   = [k for k, _ in NUTRITION_FIELDS]
NUTRITION_LABELS = [l for _, l in NUTRITION_FIELDS]

# Excel column order for nutrition (col 7+): user has kcal in G, kJ in H
EXCEL_COL_ORDER = [
    "energy_kcal",   # G (col 7)
    "energy_kj",     # H (col 8)
    "fat",           # I (col 9)
    "saturated_fat", # J (col 10)
    "carbohydrate",  # K (col 11)
    "sugars",        # L (col 12)
    "fibre",         # M (col 13)
    "protein",       # N (col 14)
    "salt",          # O (col 15)
]

SETTINGS_ORG     = "MacroSuite"
SETTINGS_APP     = "MacroSuite"
SETTINGS_LAST_DB = "last_database_path"

# ── Calorie text colors (light, comfortable on dark background) ──
CAL_TEXT_GREEN  = QColor(125, 205, 145)   # 0–149 kcal  — soft green
CAL_TEXT_YELLOW = QColor(215, 200, 95)    # 150–399 kcal — soft gold
CAL_TEXT_ORANGE = QColor(215, 140, 90)    # ≥400 kcal   — soft orange

# ── Theme colors ──
C_BG        = "#1c1c1e"
C_CARD      = "#2c2c2e"
C_CARD2     = "#323234"
C_BORDER    = "#38383a"
C_BORDER2   = "#48484a"
C_TEXT       = "#f5f5f7"
C_TEXT2      = "#98989d"
C_TEXT3      = "#636366"
C_ACCENT     = "#0a84ff"
C_ACCENT_HV  = "#409cff"
C_GREEN      = "#30d158"
C_RED        = "#ff453a"
C_TOTAL_BG   = "#1a3a5c"
C_PER100_BG  = "#2a2a2c"


# ━━━━━━━━━━━━━━━ AUTO-SELECT ON FOCUS ━━━━━━━━━━━━━━━━

class SelectAllOnFocus(QObject):
    """
    Application-wide event filter:
    1. First click / Tab into a field → select all (typing replaces content).
    2. Second click on same field → normal cursor (for editing).
    3. Ctrl+Space / Alt+Down → open combo dropdown.
    4. Tab while completer is open → cycle suggestions.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._just_focused = set()   # widget ids that just gained focus

    def eventFilter(self, obj, event):
        et = event.type()

        # ── Keyboard shortcuts ──
        if et == QEvent.KeyPress:
            # Ctrl+Space / Alt+Down → open combo dropdown
            ctrl_space = event.key() == Qt.Key_Space and event.modifiers() & Qt.ControlModifier
            alt_down   = event.key() == Qt.Key_Down  and event.modifiers() & Qt.AltModifier
            if ctrl_space or alt_down:
                combo = None
                if isinstance(obj, QComboBox):
                    combo = obj
                elif isinstance(obj, QLineEdit) and isinstance(obj.parent(), QComboBox):
                    combo = obj.parent()
                if combo:
                    combo.showPopup()
                    return True

            # Tab cycles through visible completer suggestions
            if event.key() == Qt.Key_Tab and isinstance(obj, QLineEdit):
                comp = obj.completer()
                if comp and comp.popup() and comp.popup().isVisible():
                    popup = comp.popup()
                    model = comp.completionModel()
                    if model.rowCount() > 0:
                        cur = popup.currentIndex()
                        nxt = (cur.row() + 1) if cur.isValid() and cur.row() >= 0 else 0
                        if nxt >= model.rowCount():
                            nxt = 0
                        idx = model.index(nxt, 0)
                        popup.setCurrentIndex(idx)
                        text = idx.data()
                        if text:
                            obj.setText(text)
                    return True

        # ── Select-all on focus (works for mouse AND keyboard) ──
        # Handle QComboBox by targeting its internal lineEdit
        target = obj
        if isinstance(obj, QComboBox) and obj.isEditable() and obj.lineEdit():
            target = obj.lineEdit()

        if not isinstance(target, (QDoubleSpinBox, QLineEdit)):
            return super().eventFilter(obj, event)
        if isinstance(target, QLineEdit) and target.isReadOnly():
            return super().eventFilter(obj, event)

        oid = id(target)

        if et == QEvent.FocusIn:
            # Widget just gained focus — mark it and select all
            self._just_focused.add(oid)
            QTimer.singleShot(0, lambda o=target: self._sel(o))

        elif et == QEvent.MouseButtonRelease:
            if oid in self._just_focused:
                # FIRST click after focus gain — re-select (overrides cursor placement)
                self._just_focused.discard(oid)
                QTimer.singleShot(0, lambda o=target: self._sel(o))
            # Second+ click: oid not in set → normal cursor behavior

        elif et == QEvent.FocusOut:
            self._just_focused.discard(oid)

        return super().eventFilter(obj, event)

    @staticmethod
    def _sel(obj):
        if isinstance(obj, QDoubleSpinBox):
            obj.selectAll()
        elif isinstance(obj, QLineEdit) and not obj.isReadOnly():
            obj.selectAll()


# ━━━━━━━━━━━━━━━━ ARROW ICON HELPER ━━━━━━━━━━━━━━━━━

def _create_arrow_image() -> str:
    """Create a tiny down-arrow PNG at runtime for combo-box styling."""
    path = os.path.join(tempfile.gettempdir(), "macrosuite_arrow.png")
    pm = QPixmap(14, 10)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(C_TEXT2))
    p.drawPolygon(QPolygonF([QPointF(1, 1), QPointF(13, 1), QPointF(7, 9)]))
    p.end()
    pm.save(path, "PNG")
    return path


# ━━━━━━━━━━━━━━━━━━ APPLE DARK THEME ━━━━━━━━━━━━━━━━━━

# {ARROW} is replaced at startup with the actual path
DARK_STYLE_TEMPLATE = """
* {{
    font-family: -apple-system, "SF Pro Text", "Helvetica Neue", Arial, sans-serif;
}}
QMainWindow {{ background-color: {bg}; }}
QWidget     {{ background-color: {bg}; color: {tx}; }}

QFrame#dbBar {{
    background-color: {cd}; border-bottom: 1px solid {bd}; padding: 8px 16px;
}}
QFrame#dbBar QLabel {{ color: {tx2}; font-size: 12px; background: transparent; }}
QFrame#dbBar QLabel#dbPath {{ color: {tx}; font-size: 13px; font-weight: 500; background: transparent; }}

QTabWidget::pane {{ border: none; background: {bg}; }}
QTabBar {{ background: {bg}; }}
QTabBar::tab {{
    background: transparent; color: {tx2};
    padding: 12px 32px; border: none;
    border-bottom: 2px solid transparent;
    font-size: 13px; font-weight: 600;
}}
QTabBar::tab:selected {{ color: {ac}; border-bottom: 2px solid {ac}; }}
QTabBar::tab:hover:!selected {{ color: #c7c7cc; }}

QTableWidget {{
    gridline-color: {bd}; background-color: {cd};
    alternate-background-color: {cd2};
    selection-background-color: {ac}33; selection-color: {tx};
    border: 1px solid {bd}; border-radius: 10px;
    font-size: 12px; outline: none;
}}
QTableWidget::item {{ padding: 5px 10px; border: none; }}
QTableWidget::item:selected {{ background-color: {ac}44; }}
QHeaderView::section {{
    background-color: #3a3a3c; color: {tx2};
    padding: 8px 10px; border: none;
    border-right: 1px solid {bd2}; border-bottom: 1px solid {bd2};
    font-weight: 600; font-size: 11px;
}}
QHeaderView::section:first {{ border-top-left-radius: 10px; }}
QHeaderView::section:last  {{ border-top-right-radius: 10px; border-right: none; }}

QPushButton {{
    background-color: {ac}; color: #ffffff; border: none;
    padding: 8px 22px; border-radius: 8px;
    font-weight: 600; font-size: 13px; min-height: 18px;
}}
QPushButton:hover    {{ background-color: {achv}; }}
QPushButton:pressed  {{ background-color: #0071e3; }}
QPushButton:disabled {{ background-color: {bd2}; color: {tx3}; }}
QPushButton[class="danger"]       {{ background-color: {rd}; }}
QPushButton[class="danger"]:hover {{ background-color: #ff6961; }}
QPushButton[class="secondary"]       {{ background-color: #3a3a3c; color: {tx}; border: 1px solid {bd2}; }}
QPushButton[class="secondary"]:hover {{ background-color: {bd2}; }}
QPushButton[class="dbButton"]       {{ background: #3a3a3c; color: {tx}; border: 1px solid {bd2}; padding: 6px 16px; font-size: 12px; border-radius: 6px; }}
QPushButton[class="dbButton"]:hover {{ background: {bd2}; }}

QLineEdit {{
    padding: 8px 14px; border: 1px solid {bd2};
    border-radius: 8px; background: #3a3a3c; color: {tx}; font-size: 13px;
    selection-background-color: {ac};
}}
QLineEdit:focus {{ border-color: {ac}; }}
QLineEdit::placeholder {{ color: {tx3}; }}

QComboBox {{
    padding: 8px 36px 8px 14px;
    border: 1px solid {bd2}; border-radius: 8px;
    background: #3a3a3c; color: {tx}; font-size: 13px;
    min-width: 180px;
}}
QComboBox:focus {{ border-color: {ac}; }}
QComboBox::drop-down {{
    subcontrol-origin: padding; subcontrol-position: center right;
    width: 32px; border: none; border-left: 1px solid {bd2};
}}
QComboBox::down-arrow {{
    image: url({arrow});
    width: 14px; height: 10px;
}}
QComboBox QAbstractItemView {{
    background: #3a3a3c; color: {tx};
    border: 1px solid {bd2}; border-radius: 8px;
    selection-background-color: {ac}; outline: none;
}}

QDoubleSpinBox, QSpinBox {{
    padding: 8px 14px; border: 1px solid {bd2};
    border-radius: 8px; background: #3a3a3c; color: {tx}; font-size: 13px;
}}
QDoubleSpinBox:focus, QSpinBox:focus {{ border-color: {ac}; }}

QListWidget {{
    border: 1px solid {bd}; border-radius: 10px;
    background: {cd}; font-size: 13px; outline: none;
}}
QListWidget::item {{
    padding: 12px 16px; border-bottom: 1px solid {bd}; border-radius: 0;
}}
QListWidget::item:selected {{ background-color: {ac}33; color: {ac}; }}
QListWidget::item:hover:!selected {{ background-color: #3a3a3c; }}
QListWidget::item:last {{ border-bottom: none; }}

QLabel {{ background: transparent; }}
QLabel#sectionTitle {{ font-size: 20px; font-weight: 700; color: {tx}; }}

QStatusBar {{
    background: {cd}; border-top: 1px solid {bd}; color: {tx3}; font-size: 12px;
}}
QFrame#separator {{ background-color: {bd}; max-height: 1px; }}
QScrollBar:vertical {{
    background: {cd}; width: 8px; border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {tx3}; border-radius: 4px; min-height: 30px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {cd}; height: 8px; border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {tx3}; border-radius: 4px; min-width: 30px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QMessageBox {{ background-color: {cd}; }}
QMessageBox QLabel {{ color: {tx}; }}
QDialog {{ background-color: {cd}; }}
QInputDialog {{ background-color: {cd}; }}
"""


def _build_stylesheet(arrow_path: str) -> str:
    return DARK_STYLE_TEMPLATE.format(
        bg=C_BG, cd=C_CARD, cd2=C_CARD2, bd=C_BORDER, bd2=C_BORDER2,
        tx=C_TEXT, tx2=C_TEXT2, tx3=C_TEXT3,
        ac=C_ACCENT, achv=C_ACCENT_HV, rd=C_RED,
        arrow=arrow_path.replace("\\", "/"),
    )


# ━━━━━━━━━━━━━━━━━━━━ DATA CLASSES ━━━━━━━━━━━━━━━━━━━━

@dataclass
class Ingredient:
    id: int
    name: str
    brand: str = ""
    product_name: str = ""
    energy_kj: float = 0.0
    energy_kcal: float = 0.0
    fat: float = 0.0
    saturated_fat: float = 0.0
    carbohydrate: float = 0.0
    sugars: float = 0.0
    fibre: float = 0.0
    protein: float = 0.0
    salt: float = 0.0
    package_size: str = ""

    def scaled_nutrition(self, grams: float) -> Dict[str, float]:
        f = grams / 100.0
        return {k: round(getattr(self, k) * f, 2) for k in NUTRITION_KEYS}


@dataclass
class MealIngredient:
    ingredient_id: int
    amount_grams: float


@dataclass
class Meal:
    name: str
    items: List[MealIngredient] = field(default_factory=list)


@dataclass
class MenuEntry:
    item_type: str
    item_name: str
    amount: float


@dataclass
class Menu:
    name: str
    items: List[MenuEntry] = field(default_factory=list)


# ━━━━━━━━━━━━━━━━━━ DATABASE MANAGER ━━━━━━━━━━━━━━━━━━

class DatabaseManager:
    """
    Excel I/O — preserves original formatting.
    Ingredients sheet: only VALUES are updated, all styles/formatting untouched.
    Meals/Menus sheets: managed by the app (recreated on save).
    """

    INGREDIENTS_SHEET = "Ingrediants"
    MEALS_SHEET       = "Meals"
    MENUS_SHEET       = "Menues"

    MEAL_HEADERS = [
        "MealName", "IngredientID", "IngredientName", "AmountGrams",
        "Energy (kJ)", "Energy (kcal)", "Fat (g)", "Saturated Fat (g)",
        "Carbohydrate (g)", "Sugars (g)", "Fibre (g)", "Protein (g)", "Salt (g)",
    ]
    MENU_HEADERS = [
        "MenuName", "ItemType", "ItemName", "Amount",
        "Energy (kJ)", "Energy (kcal)", "Fat (g)", "Saturated Fat (g)",
        "Carbohydrate (g)", "Sugars (g)", "Fibre (g)", "Protein (g)", "Salt (g)",
    ]

    def __init__(self, path: str):
        self.path = Path(path)
        backup = self.path.parent / f"{self.path.stem}_backup{self.path.suffix}"
        if not backup.exists():
            shutil.copy2(self.path, backup)

    # ── Read ──

    def load_ingredients(self) -> Dict[int, Ingredient]:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        ws = wb[self.INGREDIENTS_SHEET]
        result = {}
        for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
            rid, name = row[0], row[2]
            if rid is None or not name:
                continue
            result[int(rid)] = Ingredient(
                id=int(rid), name=str(name).strip(),
                brand=str(row[3] or "").strip(),
                product_name=str(row[4] or "").strip(),
                energy_kcal=_f(row[6]),  energy_kj=_f(row[7]),
                fat=_f(row[8]),         saturated_fat=_f(row[9]),
                carbohydrate=_f(row[10]), sugars=_f(row[11]),
                fibre=_f(row[12]),      protein=_f(row[13]),
                salt=_f(row[14]),
                package_size=str(row[15] or "").strip(),
            )
        wb.close()
        return result

    def load_meals(self) -> Dict[str, Meal]:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        ws = wb[self.MEALS_SHEET]
        meals: Dict[str, Meal] = OrderedDict()
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if not header or header[0] != "MealName":
            for row in ws.iter_rows(min_row=2, max_col=17, values_only=True):
                n = row[2] if len(row) > 2 else None
                if n and str(n).strip():
                    name = str(n).strip()
                    if name not in meals:
                        meals[name] = Meal(name=name)
            wb.close()
            return meals
        for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
            name = row[0]
            if not name:
                continue
            name = str(name).strip()
            if name not in meals:
                meals[name] = Meal(name=name)
            if row[1] is not None and row[3] is not None:
                meals[name].items.append(
                    MealIngredient(ingredient_id=int(row[1]), amount_grams=_f(row[3]))
                )
        wb.close()
        return meals

    def load_menus(self) -> Dict[str, Menu]:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        ws = wb[self.MENUS_SHEET]
        menus: Dict[str, Menu] = OrderedDict()
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if not header or header[0] != "MenuName":
            for row in ws.iter_rows(min_row=2, max_col=18, values_only=True):
                n = row[2] if len(row) > 2 else None
                if n and str(n).strip():
                    name = str(n).strip()
                    if name not in menus:
                        menus[name] = Menu(name=name)
            wb.close()
            return menus
        for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
            name = row[0]
            if not name:
                continue
            name = str(name).strip()
            if name not in menus:
                menus[name] = Menu(name=name)
            if row[1] and row[2] and row[3] is not None:
                menus[name].items.append(MenuEntry(
                    item_type=str(row[1]).strip().lower(),
                    item_name=str(row[2]).strip(),
                    amount=_f(row[3]),
                ))
        wb.close()
        return menus

    # ── Write (ABSOLUTE ZERO formatting changes) ──
    # Key insight: ws.cell(r,c) CREATES a cell with default (white) formatting
    # if none exists. We must NEVER access cells we don't need to write to.

    @staticmethod
    def _safe_write(ws, row, col, value):
        """Write value to cell. For empty/None values, only clear existing cells.
        NEVER creates a new cell for an empty value."""
        if value is not None and value != "":
            ws.cell(row, col).value = value
        elif (row, col) in ws._cells:
            ws._cells[(row, col)].value = None
        # else: cell doesn't exist and value is empty → do nothing

    @staticmethod
    def _clear_existing_data(ws, min_row=2):
        """Clear values of ONLY cells that actually exist and have data.
        Never creates new cells, never touches formatting."""
        for (r, c), cell in list(ws._cells.items()):
            if r >= min_row and cell.value is not None:
                cell.value = None

    @staticmethod
    def _copy_style_to(ws, row, col, style):
        """After writing to a new cell, apply the sheet's data style."""
        if style is not None and (row, col) in ws._cells:
            ws._cells[(row, col)]._style = style

    def save_all(self, ingredients, meals, menus):
        wb = openpyxl.load_workbook(self.path)

        # ── INGREDIENTS ──
        # • Each ingredient stays in its original row
        # • Column B is NEVER touched
        # • Empty cells are NEVER created
        ws = wb[self.INGREDIENTS_SHEET]

        # Map existing IDs → row numbers (only read, don't create cells)
        id_to_row: Dict[int, int] = {}
        for (r, c), cell in ws._cells.items():
            if r >= 2 and c == 1 and cell.value is not None:
                try:
                    id_to_row[int(cell.value)] = r
                except (ValueError, TypeError):
                    pass

        used_rows = set(id_to_row.values())
        written_ids = set()

        for ing in ingredients.values():
            if ing.id in id_to_row:
                r = id_to_row[ing.id]
            else:
                r = 2
                while r in used_rows:
                    r += 1
                used_rows.add(r)
                self._safe_write(ws, r, 1, ing.id)

            written_ids.add(ing.id)
            # Write data — use _safe_write to avoid creating cells for empty values
            self._safe_write(ws, r, 3, ing.name)
            self._safe_write(ws, r, 4, ing.brand or None)
            self._safe_write(ws, r, 5, ing.product_name or None)
            self._safe_write(ws, r, 6, "per 100 g")
            # Write nutrition in Excel column order (G=kcal, H=kJ, I=fat, ...)
            for c, key in enumerate(EXCEL_COL_ORDER):
                val = getattr(ing, key)
                self._safe_write(ws, r, 7 + c, val if val else None)
            self._safe_write(ws, r, 16, ing.package_size or None)

        # Clear deleted ingredients — only columns that exist
        for old_id, old_row in id_to_row.items():
            if old_id not in written_ids:
                for col in range(3, 17):
                    if (old_row, col) in ws._cells:
                        ws._cells[(old_row, col)].value = None

        # ── MEALS ──
        ws2 = wb[self.MEALS_SHEET]

        # Capture a data cell's style to apply to new cells
        data_style = None
        for (r, c), cell in ws2._cells.items():
            if r >= 2:
                data_style = cell._style
                break

        # Update headers — preserve each cell's own formatting
        for c, h in enumerate(self.MEAL_HEADERS, 1):
            if (1, c) in ws2._cells:
                ws2._cells[(1, c)].value = h  # existing cell: update value only
            else:
                ws2.cell(1, c).value = h  # new cell: unavoidable creation

        # Clear ONLY existing data cells
        self._clear_existing_data(ws2, min_row=2)

        # Write meal data, applying the original data style
        r = 2
        for meal in meals.values():
            if not meal.items:
                ws2.cell(r, 1).value = meal.name
                self._copy_style_to(ws2, r, 1, data_style)
                r += 1
                continue
            for mi in meal.items:
                ws2.cell(r, 1).value = meal.name
                ws2.cell(r, 2).value = mi.ingredient_id
                ing = ingredients.get(mi.ingredient_id)
                ws2.cell(r, 3).value = ing.name if ing else "?"
                ws2.cell(r, 4).value = mi.amount_grams
                if ing:
                    nutr = ing.scaled_nutrition(mi.amount_grams)
                    for c, key in enumerate(EXCEL_COL_ORDER):
                        ws2.cell(r, 5 + c).value = nutr[key]
                # Apply original data style to all written cells in this row
                for c in range(1, 5 + len(NUTRITION_KEYS)):
                    self._copy_style_to(ws2, r, c, data_style)
                r += 1

        # ── MENUS ──
        ws3 = wb[self.MENUS_SHEET]

        data_style3 = None
        for (r, c), cell in ws3._cells.items():
            if r >= 2:
                data_style3 = cell._style
                break

        for c, h in enumerate(self.MENU_HEADERS, 1):
            if (1, c) in ws3._cells:
                ws3._cells[(1, c)].value = h
            else:
                ws3.cell(1, c).value = h

        self._clear_existing_data(ws3, min_row=2)

        r = 2
        for menu in menus.values():
            if not menu.items:
                ws3.cell(r, 1).value = menu.name
                self._copy_style_to(ws3, r, 1, data_style3)
                r += 1
                continue
            for entry in menu.items:
                ws3.cell(r, 1).value = menu.name
                ws3.cell(r, 2).value = entry.item_type
                ws3.cell(r, 3).value = entry.item_name
                ws3.cell(r, 4).value = entry.amount
                for c in range(1, 5):
                    self._copy_style_to(ws3, r, c, data_style3)
                r += 1

        # Atomic write
        tmp = self.path.with_suffix(".tmp.xlsx")
        try:
            wb.save(tmp); wb.close()
            if self.path.exists():
                self.path.unlink()
            tmp.rename(self.path)
        except Exception:
            wb.close()
            if tmp.exists():
                tmp.unlink()
            raise


def _write_headers(ws, headers):
    hf = Font(name="Arial", bold=True, size=11, color="CCCCCC")
    hfill = PatternFill("solid", fgColor="3A3A3C")
    ha = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = hf; cell.fill = hfill; cell.alignment = ha


def _f(val) -> float:
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        try:
            return float(str(val).replace(",", ".").strip())
        except (ValueError, TypeError):
            return 0.0


# ━━━━━━━━━━━━━━━━━━ NUTRITION CALC ━━━━━━━━━━━━━━━━━━━━

class NutritionCalc:
    @staticmethod
    def meal_totals(meal, ingredients):
        total_g = 0.0
        totals = {k: 0.0 for k in NUTRITION_KEYS}
        for mi in meal.items:
            ing = ingredients.get(mi.ingredient_id)
            if not ing:
                continue
            total_g += mi.amount_grams
            for k, v in ing.scaled_nutrition(mi.amount_grams).items():
                totals[k] += v
        return round(total_g, 1), {k: round(v, 2) for k, v in totals.items()}

    @staticmethod
    def per100(total_g, totals):
        if total_g <= 0:
            return {k: 0.0 for k in NUTRITION_KEYS}
        return {k: round(v / total_g * 100, 2) for k, v in totals.items()}

    @staticmethod
    def menu_totals(menu, meals, ingredients):
        total_g = 0.0
        totals = {k: 0.0 for k in NUTRITION_KEYS}
        for entry in menu.items:
            if entry.item_type == "ingredient":
                ing = _find_ing(entry.item_name, ingredients)
                if ing:
                    total_g += entry.amount
                    for k, v in ing.scaled_nutrition(entry.amount).items():
                        totals[k] += v
            elif entry.item_type == "meal":
                m = meals.get(entry.item_name)
                if m:
                    mg, mt = NutritionCalc.meal_totals(m, ingredients)
                    total_g += mg * entry.amount
                    for k in NUTRITION_KEYS:
                        totals[k] += mt[k] * entry.amount
        return round(total_g, 1), {k: round(v, 2) for k, v in totals.items()}


def _find_ing(name, ingredients):
    nl = name.lower()
    for ing in ingredients.values():
        if ing.name.lower() == nl:
            return ing
    return None


# ━━━━━━━━━━━━━━━━━━ TABLE HELPERS ━━━━━━━━━━━━━━━━━━━━

def _calorie_text_color(kcal_per100: float) -> QColor:
    """Return a light text color based on energy density per 100 g."""
    if kcal_per100 < 150:
        return CAL_TEXT_GREEN
    elif kcal_per100 < 400:
        return CAL_TEXT_YELLOW
    else:
        return CAL_TEXT_ORANGE


class NumericItem(QTableWidgetItem):
    """Table item that sorts numerically instead of alphabetically."""
    def __lt__(self, other):
        v1 = self.data(Qt.UserRole)
        v2 = other.data(Qt.UserRole) if other else None
        if v1 is not None and v2 is not None:
            try:
                return float(v1) < float(v2)
            except (TypeError, ValueError):
                pass
        return super().__lt__(other)


def _num_item(value, suffix=""):
    if isinstance(value, float):
        text = f"{value:.2f}" if value != int(value) else str(int(value))
    else:
        text = str(value)
    item = NumericItem(text + suffix)
    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    try:
        item.setData(Qt.UserRole, float(value))
    except (ValueError, TypeError):
        pass
    return item


def _color_row_text(table: QTableWidget, row: int, color: QColor):
    """Apply a text (foreground) color to every cell in a row."""
    for col in range(table.columnCount()):
        item = table.item(row, col)
        if item:
            item.setForeground(color)


def _make_total_item(text, is_label=False):
    item = QTableWidgetItem(str(text))
    item.setFlags(Qt.ItemIsEnabled)
    item.setBackground(QColor(C_TOTAL_BG))
    item.setForeground(QColor("#7abaff"))
    f = item.font(); f.setBold(True); f.setPointSize(12); item.setFont(f)
    item.setTextAlignment((Qt.AlignLeft if is_label else Qt.AlignRight) | Qt.AlignVCenter)
    return item


def _make_per100_item(text, is_label=False):
    item = QTableWidgetItem(str(text))
    item.setFlags(Qt.ItemIsEnabled)
    item.setBackground(QColor(C_PER100_BG))
    item.setForeground(QColor(C_TEXT2))
    f = item.font(); f.setItalic(True); item.setFont(f)
    item.setTextAlignment((Qt.AlignLeft if is_label else Qt.AlignRight) | Qt.AlignVCenter)
    return item


def _fmt(v):
    if v == 0:
        return "0"
    if v == int(v):
        return str(int(v))
    return f"{v:.2f}"


# ━━━━━━━━━━━━━━━━━━━━━ DIALOGS ━━━━━━━━━━━━━━━━━━━━━━━

class IngredientDialog(QDialog):
    """Add / Edit ingredient — with autocomplete from existing data."""

    def __init__(self, parent=None, ingredient=None, next_id=1,
                 existing: Optional[Dict[int, "Ingredient"]] = None):
        super().__init__(parent)
        self.setWindowTitle("Edit Ingredient" if ingredient else "New Ingredient")
        self.setMinimumWidth(480)
        self.ingredient = ingredient

        lay = QFormLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(24, 24, 24, 24)

        # Gather completions from existing data
        names = sorted(set(i.name for i in (existing or {}).values()))
        brands = sorted(set(i.brand for i in (existing or {}).values() if i.brand))
        products = sorted(set(i.product_name for i in (existing or {}).values() if i.product_name))
        sizes = sorted(set(i.package_size for i in (existing or {}).values() if i.package_size))

        self.name_edit = QLineEdit(ingredient.name if ingredient else "")
        self.name_edit.setPlaceholderText("e.g. Chicken breast")
        self._add_completer(self.name_edit, names)

        self.brand_edit = QLineEdit(ingredient.brand if ingredient else "")
        self.brand_edit.setPlaceholderText("e.g. Organic Farm")
        self._add_completer(self.brand_edit, brands)

        self.product_edit = QLineEdit(ingredient.product_name if ingredient else "")
        self.product_edit.setPlaceholderText("e.g. Bio Hähnchenbrust")
        self._add_completer(self.product_edit, products)

        self.package_edit = QLineEdit(ingredient.package_size if ingredient else "")
        self.package_edit.setPlaceholderText("e.g. 500")
        self._add_completer(self.package_edit, sizes)

        lay.addRow("Name *", self.name_edit)
        lay.addRow("Brand", self.brand_edit)
        lay.addRow("Product Name", self.product_edit)
        lay.addRow("Package Size", self.package_edit)

        sep = QFrame(); sep.setObjectName("separator"); sep.setFrameShape(QFrame.HLine)
        lay.addRow(sep)

        lbl = QLabel("Nutrition per 100 g")
        lbl.setStyleSheet(f"font-weight: 700; font-size: 14px; color: {C_ACCENT};")
        lay.addRow(lbl)

        self.spins: Dict[str, QDoubleSpinBox] = {}
        for key, label in NUTRITION_FIELDS:
            s = QDoubleSpinBox()
            s.setRange(0, 99999); s.setDecimals(2); s.setSingleStep(0.1)
            if ingredient:
                s.setValue(getattr(ingredient, key))
            self.spins[key] = s
            lay.addRow(label, s)

        btns = QHBoxLayout(); btns.addStretch()
        cancel = QPushButton("Cancel"); cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        save = QPushButton("Save"); save.clicked.connect(self.accept); save.setDefault(True)
        btns.addWidget(cancel); btns.addWidget(save)
        lay.addRow(btns)

        self._next_id = next_id

    @staticmethod
    def _add_completer(edit: QLineEdit, items: list):
        c = QCompleter(items)
        c.setCaseSensitivity(Qt.CaseInsensitive)
        c.setFilterMode(Qt.MatchContains)
        c.popup().setStyleSheet(
            f"background: #3a3a3c; color: {C_TEXT}; border: 1px solid {C_BORDER2};"
            f"selection-background-color: {C_ACCENT}; outline: none; font-size: 13px;"
        )
        edit.setCompleter(c)

    def get_ingredient(self):
        name = self.name_edit.text().strip()
        if not name:
            return None
        return Ingredient(
            id=self.ingredient.id if self.ingredient else self._next_id,
            name=name, brand=self.brand_edit.text().strip(),
            product_name=self.product_edit.text().strip(),
            package_size=self.package_edit.text().strip(),
            **{k: s.value() for k, s in self.spins.items()},
        )


class IngredientPickerDialog(QDialog):
    """Pick ingredient by Brand → Name → Product (auto-filled)."""

    def __init__(self, parent, ingredients: Dict[int, Ingredient]):
        super().__init__(parent)
        self.setWindowTitle("Select Ingredient")
        self.setMinimumWidth(520)
        self.ingredients = ingredients
        self._all = sorted(ingredients.values(), key=lambda i: i.name.lower())

        lay = QFormLayout(self)
        lay.setSpacing(12); lay.setContentsMargins(24, 24, 24, 24)

        self.brand_combo = QComboBox()
        self.brand_combo.setEditable(True)
        self.brand_combo.setInsertPolicy(QComboBox.NoInsert)
        brands = sorted(set(i.brand for i in self._all if i.brand))
        self.brand_combo.addItem("— All brands —")
        self.brand_combo.addItems(brands)
        self._set_completer(self.brand_combo, ["— All brands —"] + brands)
        self.brand_combo.currentIndexChanged.connect(self._on_brand)
        lay.addRow("Brand", self.brand_combo)

        self.name_combo = QComboBox()
        self.name_combo.setEditable(True)
        self.name_combo.setInsertPolicy(QComboBox.NoInsert)
        self.name_combo.currentIndexChanged.connect(self._on_name)
        lay.addRow("Name", self.name_combo)

        self.product_label = QLineEdit("")
        self.product_label.setReadOnly(True)
        self.product_label.setStyleSheet(f"background: {C_BG}; color: {C_TEXT2}; border: 1px solid {C_BORDER};")
        lay.addRow("Product", self.product_label)

        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0.1, 99999)
        self.amount_spin.setDecimals(1)
        self.amount_spin.setValue(100)
        lay.addRow("Amount (g)", self.amount_spin)

        btns = QHBoxLayout(); btns.addStretch()
        cancel = QPushButton("Cancel"); cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        add = QPushButton("Add"); add.clicked.connect(self.accept); add.setDefault(True)
        btns.addWidget(cancel); btns.addWidget(add)
        lay.addRow(btns)

        self._populate_names()

    @staticmethod
    def _set_completer(combo, items):
        c = QCompleter(items)
        c.setCaseSensitivity(Qt.CaseInsensitive)
        c.setFilterMode(Qt.MatchContains)
        c.popup().setStyleSheet(
            f"background: #3a3a3c; color: {C_TEXT}; border: 1px solid {C_BORDER2};"
            f"selection-background-color: {C_ACCENT}; outline: none; font-size: 13px;"
        )
        combo.setCompleter(c)

    def _filtered(self):
        bt = self.brand_combo.currentText().strip()
        if bt == "— All brands —" or not bt:
            return self._all
        bl = bt.lower()
        return [i for i in self._all if i.brand.lower() == bl]

    def _populate_names(self):
        self.name_combo.blockSignals(True)
        self.name_combo.clear()
        names = [i.name for i in self._filtered()]
        self.name_combo.addItems(names)
        self._set_completer(self.name_combo, names)
        self.name_combo.blockSignals(False)
        if names:
            self.name_combo.setCurrentIndex(0)
            self._on_name(0)

    def _on_brand(self, idx):
        self._populate_names()

    def _on_name(self, idx):
        nt = self.name_combo.currentText().strip()
        for ing in self._all:
            if ing.name == nt:
                self.product_label.setText(ing.product_name or "—")
                return
        self.product_label.setText("—")

    def get_result(self):
        nt = self.name_combo.currentText().strip()
        for ing in self._all:
            if ing.name == nt:
                return ing.id, self.amount_spin.value()
        return None


class AddMenuItemDialog(QDialog):
    """Pick ingredient (Brand→Name→Product) or meal to add to a menu."""

    def __init__(self, parent, ingredients, meals):
        super().__init__(parent)
        self.setWindowTitle("Add Item to Menu")
        self.setMinimumWidth(520)
        self.ingredients = ingredients
        self.meals = meals
        self._all_ings = sorted(ingredients.values(), key=lambda i: i.name.lower())

        lay = QFormLayout(self)
        lay.setSpacing(12); lay.setContentsMargins(24, 24, 24, 24)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["Ingredient", "Meal"])
        self.type_combo.currentIndexChanged.connect(self._on_type)
        lay.addRow("Type", self.type_combo)

        self.brand_combo = QComboBox()
        self.brand_combo.setEditable(True)
        self.brand_combo.setInsertPolicy(QComboBox.NoInsert)
        self.brand_combo.currentIndexChanged.connect(self._on_brand)
        self.brand_lbl = QLabel("Brand")
        lay.addRow(self.brand_lbl, self.brand_combo)

        self.name_combo = QComboBox()
        self.name_combo.setEditable(True)
        self.name_combo.setInsertPolicy(QComboBox.NoInsert)
        self.name_combo.currentIndexChanged.connect(self._on_name)
        self.name_lbl = QLabel("Name")
        lay.addRow(self.name_lbl, self.name_combo)

        self.product_field = QLineEdit("")
        self.product_field.setReadOnly(True)
        self.product_field.setStyleSheet(f"background: {C_BG}; color: {C_TEXT2}; border: 1px solid {C_BORDER};")
        self.product_lbl = QLabel("Product")
        lay.addRow(self.product_lbl, self.product_field)

        self.meal_combo = QComboBox()
        self.meal_combo.setEditable(True)
        self.meal_combo.setInsertPolicy(QComboBox.NoInsert)
        self.meal_lbl = QLabel("Meal")
        lay.addRow(self.meal_lbl, self.meal_combo)

        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0.01, 99999)
        self.amount_spin.setDecimals(1)
        self.amount_spin.setValue(100)
        self.amount_lbl_w = QLabel("Amount (g)")
        lay.addRow(self.amount_lbl_w, self.amount_spin)

        btns = QHBoxLayout(); btns.addStretch()
        cancel = QPushButton("Cancel"); cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        add = QPushButton("Add"); add.clicked.connect(self.accept); add.setDefault(True)
        btns.addWidget(cancel); btns.addWidget(add)
        lay.addRow(btns)

        self._on_type(0)

    @staticmethod
    def _set_completer(combo, items):
        c = QCompleter(items)
        c.setCaseSensitivity(Qt.CaseInsensitive)
        c.setFilterMode(Qt.MatchContains)
        c.popup().setStyleSheet(
            f"background: #3a3a3c; color: {C_TEXT}; border: 1px solid {C_BORDER2};"
            f"selection-background-color: {C_ACCENT}; outline: none; font-size: 13px;"
        )
        combo.setCompleter(c)

    def _on_type(self, idx):
        is_ing = idx == 0
        for w in (self.brand_combo, self.brand_lbl, self.name_combo,
                  self.name_lbl, self.product_field, self.product_lbl):
            w.setVisible(is_ing)
        self.meal_combo.setVisible(not is_ing)
        self.meal_lbl.setVisible(not is_ing)

        if is_ing:
            self.amount_lbl_w.setText("Amount (g)")
            self.amount_spin.setValue(100)
            self._populate_brands()
        else:
            self.amount_lbl_w.setText("Servings (×)")
            self.amount_spin.setValue(1.0)
            self.meal_combo.clear()
            names = sorted(self.meals.keys())
            self.meal_combo.addItems(names)
            self._set_completer(self.meal_combo, names)

    def _populate_brands(self):
        self.brand_combo.blockSignals(True)
        self.brand_combo.clear()
        brands = sorted(set(i.brand for i in self._all_ings if i.brand))
        self.brand_combo.addItem("— All brands —")
        self.brand_combo.addItems(brands)
        self._set_completer(self.brand_combo, ["— All brands —"] + brands)
        self.brand_combo.blockSignals(False)
        self._populate_names()

    def _on_brand(self, idx):
        self._populate_names()

    def _populate_names(self):
        self.name_combo.blockSignals(True)
        self.name_combo.clear()
        bt = self.brand_combo.currentText().strip()
        if bt == "— All brands —" or not bt:
            filtered = self._all_ings
        else:
            bl = bt.lower()
            filtered = [i for i in self._all_ings if i.brand.lower() == bl]
        names = [i.name for i in filtered]
        self.name_combo.addItems(names)
        self._set_completer(self.name_combo, names)
        self.name_combo.blockSignals(False)
        if names:
            self.name_combo.setCurrentIndex(0)
            self._on_name(0)

    def _on_name(self, idx):
        nt = self.name_combo.currentText().strip()
        for ing in self._all_ings:
            if ing.name == nt:
                self.product_field.setText(ing.product_name or "—")
                return
        self.product_field.setText("—")

    def get_result(self):
        if self.type_combo.currentIndex() == 0:
            nt = self.name_combo.currentText().strip()
            for ing in self._all_ings:
                if ing.name == nt:
                    return MenuEntry("ingredient", ing.name, self.amount_spin.value())
            QMessageBox.warning(self, "Not Found", f"Ingredient '{nt}' not found.")
            return None
        else:
            mt = self.meal_combo.currentText().strip()
            if mt not in self.meals:
                QMessageBox.warning(self, "Not Found", f"Meal '{mt}' not found.")
                return None
            return MenuEntry("meal", mt, self.amount_spin.value())


# ━━━━━━━━━━━━━━━━━ INGREDIENTS TAB ━━━━━━━━━━━━━━━━━━━

class IngredientsTab(QWidget):
    data_changed = Signal()

    def __init__(self):
        super().__init__()
        self.ingredients: Dict[int, Ingredient] = {}
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20); lay.setSpacing(14)

        hdr = QHBoxLayout()
        t = QLabel("Ingredients"); t.setObjectName("sectionTitle")
        hdr.addWidget(t); hdr.addStretch()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…"); self.search.setFixedWidth(280)
        self.search.textChanged.connect(self._filter)
        hdr.addWidget(self.search)
        add = QPushButton("＋  Add Ingredient"); add.clicked.connect(self._add)
        hdr.addWidget(add)
        lay.addLayout(hdr)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)  # we handle colors manually
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        self.table.doubleClicked.connect(self._edit)
        cols = ["ID", "Name", "Brand", "Product"] + NUTRITION_LABELS + ["Pkg Size (g)"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        lay.addWidget(self.table)

        btns = QHBoxLayout(); btns.addStretch()
        eb = QPushButton("Edit"); eb.setProperty("class", "secondary"); eb.clicked.connect(self._edit)
        db = QPushButton("Delete"); db.setProperty("class", "danger"); db.clicked.connect(self._delete)
        btns.addWidget(eb); btns.addWidget(db)
        lay.addLayout(btns)

    def load(self, ingredients):
        self.ingredients = ingredients; self._populate()

    def _populate(self, ft=""):
        ft = ft.lower()
        self.table.setSortingEnabled(False)
        items = [i for i in self.ingredients.values()
                 if not ft or ft in i.name.lower() or ft in i.brand.lower()]
        self.table.setRowCount(len(items))
        for r, ing in enumerate(sorted(items, key=lambda x: x.name.lower())):
            self.table.setItem(r, 0, _num_item(ing.id))
            self.table.setItem(r, 1, QTableWidgetItem(ing.name))
            self.table.setItem(r, 2, QTableWidgetItem(ing.brand))
            self.table.setItem(r, 3, QTableWidgetItem(ing.product_name))
            for c, key in enumerate(NUTRITION_KEYS):
                self.table.setItem(r, 4 + c, _num_item(getattr(ing, key)))
            self.table.setItem(r, 4 + len(NUTRITION_KEYS), QTableWidgetItem(ing.package_size))
            # Color row by calorie density
            _color_row_text(self.table, r, _calorie_text_color(ing.energy_kcal))
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def _filter(self, text): self._populate(text)
    def _next_id(self): return max(self.ingredients.keys(), default=0) + 1

    def _add(self):
        d = IngredientDialog(self, next_id=self._next_id(), existing=self.ingredients)
        if d.exec() == QDialog.Accepted:
            ing = d.get_ingredient()
            if ing:
                self.ingredients[ing.id] = ing
                self._populate(self.search.text())
                self.data_changed.emit()

    def _edit(self):
        row = self.table.currentRow()
        if row < 0: return
        iid = int(self.table.item(row, 0).text())
        ing = self.ingredients.get(iid)
        if not ing: return
        d = IngredientDialog(self, ingredient=ing, existing=self.ingredients)
        if d.exec() == QDialog.Accepted:
            u = d.get_ingredient()
            if u:
                self.ingredients[u.id] = u
                self._populate(self.search.text())
                self.data_changed.emit()

    def _delete(self):
        row = self.table.currentRow()
        if row < 0: return
        iid = int(self.table.item(row, 0).text())
        name = self.table.item(row, 1).text()
        if QMessageBox.question(
            self, "Delete Ingredient",
            f"Delete '{name}'?\nIt will be removed from all meals using it.",
        ) == QMessageBox.Yes:
            del self.ingredients[iid]
            self._populate(self.search.text())
            self.data_changed.emit()


# ━━━━━━━━━━━━━━━━━━━ MEALS TAB ━━━━━━━━━━━━━━━━━━━━━━

LEFT_PANEL_WIDTH = 320

class MealsTab(QWidget):
    data_changed = Signal()

    def __init__(self, ing_tab):
        super().__init__()
        self.ing_tab = ing_tab
        self.meals: Dict[str, Meal] = {}
        self._build()

    @property
    def ingredients(self): return self.ing_tab.ingredients

    def _build(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20); lay.setSpacing(14)

        left = QVBoxLayout(); left.setSpacing(10)
        t = QLabel("Meals"); t.setObjectName("sectionTitle"); left.addWidget(t)
        self.search = QLineEdit(); self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self._filter_list); left.addWidget(self.search)
        self.meal_list = QListWidget()
        self.meal_list.currentItemChanged.connect(self._on_select)
        left.addWidget(self.meal_list)
        lb = QHBoxLayout(); lb.setSpacing(6)
        a = QPushButton("＋ New"); a.clicked.connect(self._add_meal)
        r = QPushButton("Rename"); r.setProperty("class", "secondary"); r.clicked.connect(self._rename_meal)
        d = QPushButton("Delete"); d.setProperty("class", "danger"); d.clicked.connect(self._delete_meal)
        lb.addWidget(a); lb.addWidget(r); lb.addWidget(d)
        left.addLayout(lb)
        lw = QWidget(); lw.setLayout(left); lw.setFixedWidth(LEFT_PANEL_WIDTH)
        lay.addWidget(lw)

        right = QVBoxLayout(); right.setSpacing(10)
        self.title_lbl = QLabel("Select a meal"); self.title_lbl.setObjectName("sectionTitle")
        right.addWidget(self.title_lbl)

        self.detail = QTableWidget()
        self.detail.setAlternatingRowColors(True)
        self.detail.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.detail.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.detail.verticalHeader().setVisible(False)
        self.detail.horizontalHeader().setStretchLastSection(True)
        cols = ["Ingredient", "Amount (g)"] + NUTRITION_LABELS
        self.detail.setColumnCount(len(cols))
        self.detail.setHorizontalHeaderLabels(cols)
        right.addWidget(self.detail)

        rb = QHBoxLayout()
        ai = QPushButton("＋ Add Ingredient"); ai.clicked.connect(self._add_ing)
        ea = QPushButton("Edit Amount"); ea.setProperty("class", "secondary"); ea.clicked.connect(self._edit_amount)
        ri = QPushButton("Remove"); ri.setProperty("class", "danger"); ri.clicked.connect(self._remove_ing)
        rb.addWidget(ai); rb.addWidget(ea); rb.addStretch(); rb.addWidget(ri)
        right.addLayout(rb)
        lay.addLayout(right, 1)

    def load(self, meals): self.meals = meals; self._populate_list()

    def _populate_list(self, ft=""):
        ft = ft.lower()
        self.meal_list.blockSignals(True)
        cur = self.meal_list.currentItem()
        cn = cur.text() if cur else None
        self.meal_list.clear()
        for n in sorted(self.meals, key=str.lower):
            if ft and ft not in n.lower(): continue
            self.meal_list.addItem(n)
        if cn:
            found = self.meal_list.findItems(cn, Qt.MatchExactly)
            if found: self.meal_list.setCurrentItem(found[0])
        self.meal_list.blockSignals(False)
        if self.meal_list.currentItem():
            self._on_select(self.meal_list.currentItem())

    def _filter_list(self, t): self._populate_list(t)
    def _cur(self):
        i = self.meal_list.currentItem()
        return self.meals.get(i.text()) if i else None

    def _on_select(self, cur, prev=None):
        meal = self._cur()
        if not meal:
            self.title_lbl.setText("Select a meal"); self.detail.setRowCount(0); return
        self.title_lbl.setText(meal.name); self._refresh()

    def _refresh(self):
        meal = self._cur()
        if not meal: return
        n = len(meal.items)
        self.detail.setSortingEnabled(False)
        self.detail.setRowCount(n + 2)

        for r, mi in enumerate(meal.items):
            ing = self.ingredients.get(mi.ingredient_id)
            self.detail.setItem(r, 0, QTableWidgetItem(ing.name if ing else f"[ID {mi.ingredient_id}]"))
            self.detail.setItem(r, 1, _num_item(mi.amount_grams))
            if ing:
                sc = ing.scaled_nutrition(mi.amount_grams)
                for c, key in enumerate(NUTRITION_KEYS):
                    self.detail.setItem(r, 2 + c, _num_item(sc[key]))
                _color_row_text(self.detail, r, _calorie_text_color(ing.energy_kcal))

        # TOTAL
        tg, tt = NutritionCalc.meal_totals(meal, self.ingredients)
        self.detail.setItem(n, 0, _make_total_item("TOTAL", True))
        self.detail.setItem(n, 1, _make_total_item(_fmt(tg)))
        for c, key in enumerate(NUTRITION_KEYS):
            self.detail.setItem(n, 2 + c, _make_total_item(_fmt(tt[key])))

        # PER 100g
        p = NutritionCalc.per100(tg, tt)
        self.detail.setItem(n+1, 0, _make_per100_item("per 100 g", True))
        self.detail.setItem(n+1, 1, _make_per100_item("100"))
        for c, key in enumerate(NUTRITION_KEYS):
            self.detail.setItem(n+1, 2 + c, _make_per100_item(_fmt(p[key])))

        self.detail.resizeColumnsToContents()

    def _add_meal(self):
        n, ok = QInputDialog.getText(self, "New Meal", "Meal name:")
        n = n.strip() if ok else ""
        if not n: return
        if n in self.meals: QMessageBox.warning(self, "Exists", f"'{n}' already exists."); return
        self.meals[n] = Meal(name=n); self._populate_list()
        found = self.meal_list.findItems(n, Qt.MatchExactly)
        if found: self.meal_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _rename_meal(self):
        m = self._cur()
        if not m: return
        nn, ok = QInputDialog.getText(self, "Rename", "New name:", text=m.name)
        nn = nn.strip() if ok else ""
        if not nn or nn == m.name: return
        if nn in self.meals: QMessageBox.warning(self, "Exists", f"'{nn}' already exists."); return
        del self.meals[m.name]; m.name = nn; self.meals[nn] = m; self._populate_list()
        found = self.meal_list.findItems(nn, Qt.MatchExactly)
        if found: self.meal_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _delete_meal(self):
        m = self._cur()
        if not m: return
        if QMessageBox.question(self, "Delete", f"Delete meal '{m.name}'?") == QMessageBox.Yes:
            del self.meals[m.name]; self._populate_list(); self.data_changed.emit()

    def _add_ing(self):
        m = self._cur()
        if not m: QMessageBox.information(self, "No Meal", "Select or create a meal first."); return
        d = IngredientPickerDialog(self, self.ingredients)
        if d.exec() == QDialog.Accepted:
            r = d.get_result()
            if r:
                m.items.append(MealIngredient(ingredient_id=r[0], amount_grams=r[1]))
                self._refresh(); self.data_changed.emit()

    def _edit_amount(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items): return
        mi = m.items[row]
        amt, ok = QInputDialog.getDouble(self, "Edit Amount", "Amount (g):", mi.amount_grams, 0.1, 99999, 1)
        if ok: mi.amount_grams = amt; self._refresh(); self.data_changed.emit()

    def _remove_ing(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items): return
        m.items.pop(row); self._refresh(); self.data_changed.emit()


# ━━━━━━━━━━━━━━━━━━━ MENUS TAB ━━━━━━━━━━━━━━━━━━━━━━

class MenusTab(QWidget):
    data_changed = Signal()

    def __init__(self, ing_tab, meals_tab):
        super().__init__()
        self.ing_tab = ing_tab; self.meals_tab = meals_tab
        self.menus: Dict[str, Menu] = {}
        self._build()

    @property
    def ingredients(self): return self.ing_tab.ingredients
    @property
    def meals(self): return self.meals_tab.meals

    def _build(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20); lay.setSpacing(14)

        left = QVBoxLayout(); left.setSpacing(10)
        t = QLabel("Menus"); t.setObjectName("sectionTitle"); left.addWidget(t)
        self.search = QLineEdit(); self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self._filter_list); left.addWidget(self.search)
        self.menu_list = QListWidget()
        self.menu_list.currentItemChanged.connect(self._on_select)
        left.addWidget(self.menu_list)
        lb = QHBoxLayout(); lb.setSpacing(6)
        a = QPushButton("＋ New"); a.clicked.connect(self._add_menu)
        r = QPushButton("Rename"); r.setProperty("class", "secondary"); r.clicked.connect(self._rename_menu)
        d = QPushButton("Delete"); d.setProperty("class", "danger"); d.clicked.connect(self._delete_menu)
        lb.addWidget(a); lb.addWidget(r); lb.addWidget(d)
        left.addLayout(lb)
        lw = QWidget(); lw.setLayout(left); lw.setFixedWidth(LEFT_PANEL_WIDTH)
        lay.addWidget(lw)

        right = QVBoxLayout(); right.setSpacing(10)
        self.title_lbl = QLabel("Select a menu"); self.title_lbl.setObjectName("sectionTitle")
        right.addWidget(self.title_lbl)

        self.detail = QTableWidget()
        self.detail.setAlternatingRowColors(True)
        self.detail.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.detail.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.detail.verticalHeader().setVisible(False)
        self.detail.horizontalHeader().setStretchLastSection(True)
        cols = ["Type", "Name", "Amount", "Weight (g)"] + NUTRITION_LABELS
        self.detail.setColumnCount(len(cols))
        self.detail.setHorizontalHeaderLabels(cols)
        right.addWidget(self.detail)

        rb = QHBoxLayout()
        ai = QPushButton("＋ Add Item"); ai.clicked.connect(self._add_item)
        ea = QPushButton("Edit Amount"); ea.setProperty("class", "secondary"); ea.clicked.connect(self._edit_amount)
        ri = QPushButton("Remove"); ri.setProperty("class", "danger"); ri.clicked.connect(self._remove_item)
        rb.addWidget(ai); rb.addWidget(ea); rb.addStretch(); rb.addWidget(ri)
        right.addLayout(rb)
        lay.addLayout(right, 1)

    def load(self, menus): self.menus = menus; self._populate_list()

    def _populate_list(self, ft=""):
        ft = ft.lower()
        self.menu_list.blockSignals(True)
        cur = self.menu_list.currentItem()
        cn = cur.text() if cur else None
        self.menu_list.clear()
        for n in sorted(self.menus, key=str.lower):
            if ft and ft not in n.lower(): continue
            self.menu_list.addItem(n)
        if cn:
            found = self.menu_list.findItems(cn, Qt.MatchExactly)
            if found: self.menu_list.setCurrentItem(found[0])
        self.menu_list.blockSignals(False)
        if self.menu_list.currentItem():
            self._on_select(self.menu_list.currentItem())

    def _filter_list(self, t): self._populate_list(t)
    def _cur(self):
        i = self.menu_list.currentItem()
        return self.menus.get(i.text()) if i else None

    def _on_select(self, cur, prev=None):
        menu = self._cur()
        if not menu:
            self.title_lbl.setText("Select a menu"); self.detail.setRowCount(0); return
        self.title_lbl.setText(menu.name); self._refresh()

    def _refresh(self):
        menu = self._cur()
        if not menu: return
        ni = len(menu.items)
        self.detail.setSortingEnabled(False)
        self.detail.setRowCount(ni + 2)

        for r, entry in enumerate(menu.items):
            self.detail.setItem(r, 0, QTableWidgetItem(entry.item_type.capitalize()))
            self.detail.setItem(r, 1, QTableWidgetItem(entry.item_name))

            kcal_for_color = 0.0
            if entry.item_type == "ingredient":
                self.detail.setItem(r, 2, _num_item(entry.amount))
                self.detail.setItem(r, 3, _num_item(entry.amount))
                ing = _find_ing(entry.item_name, self.ingredients)
                if ing:
                    kcal_for_color = ing.energy_kcal
                    sc = ing.scaled_nutrition(entry.amount)
                    for c, key in enumerate(NUTRITION_KEYS):
                        self.detail.setItem(r, 4 + c, _num_item(sc[key]))
            else:
                self.detail.setItem(r, 2, _num_item(entry.amount, " ×"))
                meal = self.meals.get(entry.item_name)
                if meal:
                    mg, mt = NutritionCalc.meal_totals(meal, self.ingredients)
                    self.detail.setItem(r, 3, _num_item(round(mg * entry.amount, 1)))
                    for c, key in enumerate(NUTRITION_KEYS):
                        self.detail.setItem(r, 4 + c, _num_item(round(mt[key] * entry.amount, 2)))
                    # Use meal's per-100g kcal for color
                    p100 = NutritionCalc.per100(mg, mt)
                    kcal_for_color = p100.get("energy_kcal", 0)

            _color_row_text(self.detail, r, _calorie_text_color(kcal_for_color))

        # TOTAL
        tg, tt = NutritionCalc.menu_totals(menu, self.meals, self.ingredients)
        self.detail.setItem(ni, 0, _make_total_item("TOTAL", True))
        self.detail.setItem(ni, 1, _make_total_item(""))
        self.detail.setItem(ni, 2, _make_total_item(""))
        self.detail.setItem(ni, 3, _make_total_item(_fmt(tg)))
        for c, key in enumerate(NUTRITION_KEYS):
            self.detail.setItem(ni, 4 + c, _make_total_item(_fmt(tt[key])))

        # PER 100g
        p = NutritionCalc.per100(tg, tt)
        self.detail.setItem(ni+1, 0, _make_per100_item("per 100 g", True))
        self.detail.setItem(ni+1, 1, _make_per100_item(""))
        self.detail.setItem(ni+1, 2, _make_per100_item(""))
        self.detail.setItem(ni+1, 3, _make_per100_item("100"))
        for c, key in enumerate(NUTRITION_KEYS):
            self.detail.setItem(ni+1, 4 + c, _make_per100_item(_fmt(p[key])))

        self.detail.resizeColumnsToContents()

    def _add_menu(self):
        n, ok = QInputDialog.getText(self, "New Menu", "Menu name:")
        n = n.strip() if ok else ""
        if not n: return
        if n in self.menus: QMessageBox.warning(self, "Exists", f"'{n}' already exists."); return
        self.menus[n] = Menu(name=n); self._populate_list()
        found = self.menu_list.findItems(n, Qt.MatchExactly)
        if found: self.menu_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _rename_menu(self):
        m = self._cur()
        if not m: return
        nn, ok = QInputDialog.getText(self, "Rename", "New name:", text=m.name)
        nn = nn.strip() if ok else ""
        if not nn or nn == m.name: return
        if nn in self.menus: QMessageBox.warning(self, "Exists", f"'{nn}' already exists."); return
        del self.menus[m.name]; m.name = nn; self.menus[nn] = m; self._populate_list()
        found = self.menu_list.findItems(nn, Qt.MatchExactly)
        if found: self.menu_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _delete_menu(self):
        m = self._cur()
        if not m: return
        if QMessageBox.question(self, "Delete", f"Delete menu '{m.name}'?") == QMessageBox.Yes:
            del self.menus[m.name]; self._populate_list(); self.data_changed.emit()

    def _add_item(self):
        m = self._cur()
        if not m: QMessageBox.information(self, "No Menu", "Select or create a menu first."); return
        d = AddMenuItemDialog(self, self.ingredients, self.meals)
        if d.exec() == QDialog.Accepted:
            e = d.get_result()
            if e: m.items.append(e); self._refresh(); self.data_changed.emit()

    def _edit_amount(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items): return
        entry = m.items[row]
        label = "Amount (g):" if entry.item_type == "ingredient" else "Servings (×):"
        amt, ok = QInputDialog.getDouble(self, "Edit Amount", label, entry.amount, 0.01, 99999, 2)
        if ok: entry.amount = amt; self._refresh(); self.data_changed.emit()

    def _remove_item(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items): return
        m.items.pop(row); self._refresh(); self.data_changed.emit()


# ━━━━━━━━━━━━━━━━━━ MAIN WINDOW ━━━━━━━━━━━━━━━━━━━━━

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1240, 780)

        self.db: Optional[DatabaseManager] = None
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)

        central = QWidget(); self.setCentralWidget(central)
        self.root = QVBoxLayout(central)
        self.root.setContentsMargins(0, 0, 0, 0); self.root.setSpacing(0)

        # Database bar
        self.db_bar = QFrame(); self.db_bar.setObjectName("dbBar"); self.db_bar.setFixedHeight(52)
        db_lay = QHBoxLayout(self.db_bar)
        db_lay.setContentsMargins(20, 0, 20, 0); db_lay.setSpacing(12)
        self.db_icon = QLabel("◉")
        self.db_icon.setStyleSheet(f"font-size: 16px; color: {C_TEXT3}; background: transparent;")
        db_lay.addWidget(self.db_icon)
        db_label = QLabel("Database")
        db_label.setStyleSheet(f"color: {C_TEXT2}; font-size: 12px; font-weight: 600; background: transparent;")
        db_lay.addWidget(db_label)
        self.db_path_label = QLabel("No database selected")
        self.db_path_label.setObjectName("dbPath")
        db_lay.addWidget(self.db_path_label, 1)
        browse_btn = QPushButton("Change…")
        browse_btn.setProperty("class", "dbButton"); browse_btn.setFixedWidth(100)
        browse_btn.clicked.connect(self._browse_db)
        db_lay.addWidget(browse_btn)
        self.root.addWidget(self.db_bar)

        # Tabs
        self.ing_tab = IngredientsTab()
        self.meals_tab = MealsTab(self.ing_tab)
        self.menus_tab = MenusTab(self.ing_tab, self.meals_tab)
        self.tabs = QTabWidget()
        self.tabs.addTab(self.ing_tab, "   Ingredients   ")
        self.tabs.addTab(self.meals_tab, "   Meals   ")
        self.tabs.addTab(self.menus_tab, "   Menus   ")
        self.tabs.currentChanged.connect(self._on_tab)
        self.root.addWidget(self.tabs, 1)

        self.status = QStatusBar(); self.setStatusBar(self.status)

        self._save_timer = QTimer()
        self._save_timer.setSingleShot(True); self._save_timer.setInterval(2000)
        self._save_timer.timeout.connect(self._do_save)
        self.ing_tab.data_changed.connect(self._schedule_save)
        self.meals_tab.data_changed.connect(self._schedule_save)
        self.menus_tab.data_changed.connect(self._schedule_save)

        mb = self.menuBar()
        mb.setStyleSheet(f"QMenuBar {{ background: {C_CARD}; color: {C_TEXT}; }}"
                         f"QMenuBar::item:selected {{ background: {C_BORDER2}; }}"
                         f"QMenu {{ background: #3a3a3c; color: {C_TEXT}; border: 1px solid {C_BORDER2}; }}"
                         f"QMenu::item:selected {{ background: {C_ACCENT}; }}")
        fm = mb.addMenu("File")
        sa = QAction("Save Now", self); sa.setShortcut("Ctrl+S"); sa.triggered.connect(self._do_save); fm.addAction(sa)
        fm.addSeparator()
        qa = QAction("Quit", self); qa.setShortcut("Ctrl+Q"); qa.triggered.connect(self.close); fm.addAction(qa)

        # Ctrl+Plus shortcut → open Add dialog for current tab
        from PySide6.QtGui import QKeySequence, QShortcut
        add_sc = QShortcut(QKeySequence("Ctrl++"), self)
        add_sc.activated.connect(self._on_add_shortcut)
        # Also handle Ctrl+= (Plus without Shift on some keyboards)
        add_sc2 = QShortcut(QKeySequence("Ctrl+="), self)
        add_sc2.activated.connect(self._on_add_shortcut)

        # Ctrl+Tab / Ctrl+Shift+Tab → cycle tabs
        tab_fwd = QShortcut(QKeySequence("Ctrl+Tab"), self)
        tab_fwd.activated.connect(lambda: self.tabs.setCurrentIndex(
            (self.tabs.currentIndex() + 1) % self.tabs.count()))
        tab_bwd = QShortcut(QKeySequence("Ctrl+Shift+Tab"), self)
        tab_bwd.activated.connect(lambda: self.tabs.setCurrentIndex(
            (self.tabs.currentIndex() - 1) % self.tabs.count()))

        last = self.settings.value(SETTINGS_LAST_DB, "")
        if last and Path(last).is_file():
            self._open_db(last)
        else:
            self.status.showMessage("Select a database to get started")

    def _browse_db(self):
        start = str(Path(self.db.path).parent) if self.db else str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, "Select Nutrition Database", start, "Excel Files (*.xlsx);;All Files (*)")
        if path: self._open_db(path)

    def _open_db(self, path):
        try:
            self.db = DatabaseManager(path)
            self.settings.setValue(SETTINGS_LAST_DB, path)
            self.ing_tab.load(self.db.load_ingredients())
            self.meals_tab.load(self.db.load_meals())
            self.menus_tab.load(self.db.load_menus())
            display = path if len(path) <= 65 else str(Path(Path(path).parts[0], "…", *Path(path).parts[-2:]))
            self.db_path_label.setText(display); self.db_path_label.setToolTip(path)
            self.db_icon.setStyleSheet(f"font-size: 16px; color: {C_GREEN}; background: transparent;")
            self.db_icon.setText("◉")
            self.status.showMessage(f"Loaded — {len(self.ing_tab.ingredients)} ingredients")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open database:\n\n{e}")

    def _schedule_save(self):
        if not self.db: return
        self.status.showMessage("Unsaved changes…"); self._save_timer.start()

    def _do_save(self):
        if not self.db: return
        try:
            self.db.save_all(self.ing_tab.ingredients, self.meals_tab.meals, self.menus_tab.menus)
            self.status.showMessage(f"Saved  ·  {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            self.status.showMessage(f"SAVE ERROR: {e}")
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n\n{e}\n\nYour backup is intact.")

    def _on_tab(self, idx):
        if idx == 1: self.meals_tab._refresh()
        elif idx == 2: self.menus_tab._refresh()

    def _on_add_shortcut(self):
        """Ctrl+Plus → trigger the Add button on the current tab."""
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.ing_tab._add()
        elif idx == 1:
            self.meals_tab._add_ing()
        elif idx == 2:
            self.menus_tab._add_item()

    def closeEvent(self, event):
        if self.db:
            try:
                self.db.save_all(self.ing_tab.ingredients, self.meals_tab.meals, self.menus_tab.menus)
            except Exception as e:
                if QMessageBox.warning(self, "Save Error", f"Could not save:\n{e}\n\nQuit anyway?",
                                       QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                    event.ignore(); return
        event.accept()


# ━━━━━━━━━━━━━━━━━━━━ ENTRY POINT ━━━━━━━━━━━━━━━━━━━━

def main():
    app = QApplication(sys.argv)
    app.setFont(QFontGui("SF Pro Text", 12))

    # Create combo-box arrow icon and build stylesheet
    arrow_path = _create_arrow_image()
    app.setStyleSheet(_build_stylesheet(arrow_path))

    app.setApplicationName(SETTINGS_APP)
    app.setOrganizationName(SETTINGS_ORG)

    # Auto-select text on focus for all inputs
    focus_filter = SelectAllOnFocus(app)
    app.installEventFilter(focus_filter)

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()