#!/usr/bin/env python3
"""
Nutrition Planner — Professional meal & menu planning application.
Dark Apple-style UI · Excel database · Automatic backup · Persistent settings.

Usage:  python nutrition_planner.py
"""

import sys
import os
import shutil
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
from PySide6.QtCore import Qt, QTimer, Signal, QSettings, QSize
from PySide6.QtGui import QFont as QFontGui, QColor, QAction


# ━━━━━━━━━━━━━━━━━━━━━━ CONSTANTS ━━━━━━━━━━━━━━━━━━━━━━

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

SETTINGS_ORG  = "NutritionPlanner"
SETTINGS_APP  = "NutritionPlanner"
SETTINGS_LAST_DB = "last_database_path"

# ━━━━━━━━━━━━━━━━━━ APPLE DARK THEME ━━━━━━━━━━━━━━━━━━

DARK_STYLE = """
/* ── Global ── */
* {
    font-family: -apple-system, "SF Pro Text", "Helvetica Neue", Arial, sans-serif;
}
QMainWindow {
    background-color: #1c1c1e;
}
QWidget {
    background-color: #1c1c1e;
    color: #f5f5f7;
}

/* ── Top database bar ── */
QFrame#dbBar {
    background-color: #2c2c2e;
    border-bottom: 1px solid #38383a;
    padding: 8px 16px;
}
QFrame#dbBar QLabel {
    color: #98989d;
    font-size: 12px;
    background: transparent;
}
QFrame#dbBar QLabel#dbPath {
    color: #f5f5f7;
    font-size: 13px;
    font-weight: 500;
}

/* ── Tabs ── */
QTabWidget::pane {
    border: none;
    background: #1c1c1e;
}
QTabBar {
    background: #1c1c1e;
}
QTabBar::tab {
    background: transparent;
    color: #98989d;
    padding: 12px 32px;
    margin: 0;
    border: none;
    border-bottom: 2px solid transparent;
    font-size: 13px;
    font-weight: 600;
}
QTabBar::tab:selected {
    color: #0a84ff;
    border-bottom: 2px solid #0a84ff;
}
QTabBar::tab:hover:!selected {
    color: #c7c7cc;
}

/* ── Tables ── */
QTableWidget {
    gridline-color: #38383a;
    background-color: #2c2c2e;
    alternate-background-color: #323234;
    selection-background-color: #0a84ff33;
    selection-color: #f5f5f7;
    border: 1px solid #38383a;
    border-radius: 10px;
    font-size: 12px;
    outline: none;
}
QTableWidget::item {
    padding: 5px 10px;
    border: none;
}
QTableWidget::item:selected {
    background-color: #0a84ff44;
}
QHeaderView::section {
    background-color: #3a3a3c;
    color: #98989d;
    padding: 8px 10px;
    border: none;
    border-right: 1px solid #48484a;
    border-bottom: 1px solid #48484a;
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
}
QHeaderView::section:first {
    border-top-left-radius: 10px;
}
QHeaderView::section:last {
    border-top-right-radius: 10px;
    border-right: none;
}

/* ── Buttons ── */
QPushButton {
    background-color: #0a84ff;
    color: #ffffff;
    border: none;
    padding: 8px 22px;
    border-radius: 8px;
    font-weight: 600;
    font-size: 13px;
    min-height: 18px;
}
QPushButton:hover {
    background-color: #409cff;
}
QPushButton:pressed {
    background-color: #0071e3;
}
QPushButton:disabled {
    background-color: #48484a;
    color: #636366;
}
QPushButton[class="danger"] {
    background-color: #ff453a;
}
QPushButton[class="danger"]:hover {
    background-color: #ff6961;
}
QPushButton[class="danger"]:pressed {
    background-color: #d93025;
}
QPushButton[class="secondary"] {
    background-color: #3a3a3c;
    color: #f5f5f7;
    border: 1px solid #48484a;
}
QPushButton[class="secondary"]:hover {
    background-color: #48484a;
}
QPushButton[class="ghost"] {
    background-color: transparent;
    color: #0a84ff;
    border: none;
    padding: 6px 14px;
}
QPushButton[class="ghost"]:hover {
    background-color: #0a84ff22;
    border-radius: 6px;
}
QPushButton[class="dbButton"] {
    background-color: #3a3a3c;
    color: #f5f5f7;
    border: 1px solid #48484a;
    padding: 6px 16px;
    font-size: 12px;
    border-radius: 6px;
}
QPushButton[class="dbButton"]:hover {
    background-color: #48484a;
}

/* ── Inputs ── */
QLineEdit {
    padding: 8px 14px;
    border: 1px solid #48484a;
    border-radius: 8px;
    background: #3a3a3c;
    color: #f5f5f7;
    font-size: 13px;
    selection-background-color: #0a84ff;
}
QLineEdit:focus {
    border-color: #0a84ff;
}
QLineEdit::placeholder {
    color: #636366;
}
QComboBox {
    padding: 8px 14px;
    border: 1px solid #48484a;
    border-radius: 8px;
    background: #3a3a3c;
    color: #f5f5f7;
    font-size: 13px;
    min-width: 180px;
}
QComboBox:focus {
    border-color: #0a84ff;
}
QComboBox::drop-down {
    border: none;
    width: 28px;
}
QComboBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #98989d;
    margin-right: 8px;
}
QComboBox QAbstractItemView {
    background: #3a3a3c;
    color: #f5f5f7;
    border: 1px solid #48484a;
    border-radius: 8px;
    selection-background-color: #0a84ff;
    outline: none;
}
QDoubleSpinBox, QSpinBox {
    padding: 8px 14px;
    border: 1px solid #48484a;
    border-radius: 8px;
    background: #3a3a3c;
    color: #f5f5f7;
    font-size: 13px;
}
QDoubleSpinBox:focus, QSpinBox:focus {
    border-color: #0a84ff;
}

/* ── List ── */
QListWidget {
    border: 1px solid #38383a;
    border-radius: 10px;
    background: #2c2c2e;
    font-size: 13px;
    outline: none;
}
QListWidget::item {
    padding: 12px 16px;
    border-bottom: 1px solid #38383a;
    border-radius: 0;
}
QListWidget::item:selected {
    background-color: #0a84ff33;
    color: #0a84ff;
}
QListWidget::item:hover:!selected {
    background-color: #3a3a3c;
}
QListWidget::item:last {
    border-bottom: none;
}

/* ── Misc ── */
QLabel#sectionTitle {
    font-size: 20px;
    font-weight: 700;
    color: #f5f5f7;
    background: transparent;
}
QLabel#subtitle {
    font-size: 12px;
    color: #98989d;
    background: transparent;
}
QLabel {
    background: transparent;
}
QStatusBar {
    background: #2c2c2e;
    border-top: 1px solid #38383a;
    color: #636366;
    font-size: 12px;
}
QFrame#summaryCard {
    background: #2c2c2e;
    border: 1px solid #38383a;
    border-radius: 12px;
    padding: 14px;
}
QFrame#separator {
    background-color: #38383a;
    max-height: 1px;
}
QScrollBar:vertical {
    background: #2c2c2e;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #636366;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: #2c2c2e;
    height: 8px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal {
    background: #636366;
    border-radius: 4px;
    min-width: 30px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}
QMessageBox {
    background-color: #2c2c2e;
}
QMessageBox QLabel {
    color: #f5f5f7;
}
QDialog {
    background-color: #2c2c2e;
}
QFormLayout {
    background: transparent;
}
QInputDialog {
    background-color: #2c2c2e;
}
"""


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

    def nutrition_dict(self) -> Dict[str, float]:
        return {k: getattr(self, k) for k in NUTRITION_KEYS}

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
    item_type: str      # "ingredient" | "meal"
    item_name: str
    amount: float       # grams for ingredients, servings for meals


@dataclass
class Menu:
    name: str
    items: List[MenuEntry] = field(default_factory=list)


# ━━━━━━━━━━━━━━━━━━ DATABASE MANAGER ━━━━━━━━━━━━━━━━━━

class DatabaseManager:
    """
    Excel I/O with industry-standard backup strategy:

    1. On first open → one-time backup created next to the DB file
       (e.g. data.xlsx → data_backup.xlsx). Never overwritten.
    2. Every change auto-saves to the main file via debounced timer (2 s).
    3. Writes are atomic: write to .tmp, then rename over original.
    4. If app crashes mid-write, only the .tmp is lost — original intact.
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
        self._ensure_backup()

    def _ensure_backup(self):
        """Create ONE backup next to the file — only if it doesn't exist yet."""
        backup = self.path.parent / f"{self.path.stem}_backup{self.path.suffix}"
        if not backup.exists():
            shutil.copy2(self.path, backup)

    # ── Read ──────────────────────────────────────────────

    def load_ingredients(self) -> Dict[int, Ingredient]:
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        ws = wb[self.INGREDIENTS_SHEET]
        ingredients = {}
        for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
            rid, name = row[0], row[2]
            if rid is None or not name:
                continue
            ingredients[int(rid)] = Ingredient(
                id=int(rid), name=str(name).strip(),
                brand=str(row[3] or "").strip(),
                product_name=str(row[4] or "").strip(),
                energy_kj=_f(row[6]),   energy_kcal=_f(row[7]),
                fat=_f(row[8]),         saturated_fat=_f(row[9]),
                carbohydrate=_f(row[10]), sugars=_f(row[11]),
                fibre=_f(row[12]),      protein=_f(row[13]),
                salt=_f(row[14]),
                package_size=str(row[15] or "").strip(),
            )
        wb.close()
        return ingredients

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

    # ── Write (atomic) ───────────────────────────────────

    def save_all(self, ingredients: Dict[int, Ingredient],
                 meals: Dict[str, Meal], menus: Dict[str, Menu]):
        wb = openpyxl.Workbook()

        # — Ingredients —
        ws = wb.active
        ws.title = self.INGREDIENTS_SHEET
        ING_H = [
            "ID", "human verified", "Ingredient Name [ENG]", "Brand",
            "Branded Product Name [Original language]", "Basis",
            "Energy (kJ)", "Energy (kcal)", "Fat (g)", "Saturated Fat (g)",
            "Carbohydrate (g)", "Sugars (g)", "Fibre (g)", "Protein (g)",
            "Salt (g)", "Package Size (g)",
        ]
        self._write_headers(ws, ING_H)
        for r, ing in enumerate(sorted(ingredients.values(), key=lambda i: i.id), 2):
            ws.cell(r, 1, ing.id)
            ws.cell(r, 2, 3)
            ws.cell(r, 3, ing.name)
            ws.cell(r, 4, ing.brand)
            ws.cell(r, 5, ing.product_name)
            ws.cell(r, 6, "per 100 g")
            for c, key in enumerate(NUTRITION_KEYS):
                ws.cell(r, 7 + c, getattr(ing, key))
            ws.cell(r, 16, ing.package_size)

        # — Meals —
        ws2 = wb.create_sheet(self.MEALS_SHEET)
        self._write_headers(ws2, self.MEAL_HEADERS)
        r = 2
        for meal in meals.values():
            if not meal.items:
                ws2.cell(r, 1, meal.name); r += 1; continue
            for mi in meal.items:
                ws2.cell(r, 1, meal.name)
                ws2.cell(r, 2, mi.ingredient_id)
                ing = ingredients.get(mi.ingredient_id)
                ws2.cell(r, 3, ing.name if ing else "?")
                ws2.cell(r, 4, mi.amount_grams)
                if ing:
                    nutr = ing.scaled_nutrition(mi.amount_grams)
                    for c, key in enumerate(NUTRITION_KEYS):
                        ws2.cell(r, 5 + c, nutr[key])
                r += 1

        # — Menus —
        ws3 = wb.create_sheet(self.MENUS_SHEET)
        self._write_headers(ws3, self.MENU_HEADERS)
        r = 2
        for menu in menus.values():
            if not menu.items:
                ws3.cell(r, 1, menu.name); r += 1; continue
            for entry in menu.items:
                ws3.cell(r, 1, menu.name)
                ws3.cell(r, 2, entry.item_type)
                ws3.cell(r, 3, entry.item_name)
                ws3.cell(r, 4, entry.amount)
                r += 1

        # Atomic: write tmp → rename
        tmp = self.path.with_suffix(".tmp.xlsx")
        try:
            wb.save(tmp)
            wb.close()
            if self.path.exists():
                self.path.unlink()
            tmp.rename(self.path)
        except Exception:
            wb.close()
            if tmp.exists():
                tmp.unlink()
            raise

    @staticmethod
    def _write_headers(ws, headers):
        hf = Font(name="Arial", bold=True, size=11, color="CCCCCC")
        hfill = PatternFill("solid", fgColor="3A3A3C")
        ha = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c, h in enumerate(headers, 1):
            cell = ws.cell(1, c, h)
            cell.font = hf
            cell.fill = hfill
            cell.alignment = ha


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
    def meal_totals(meal: Meal, ingredients: Dict[int, Ingredient]):
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
    def per100(total_g: float, totals: Dict[str, float]):
        if total_g <= 0:
            return {k: 0.0 for k in NUTRITION_KEYS}
        return {k: round(v / total_g * 100, 2) for k, v in totals.items()}

    @staticmethod
    def menu_totals(menu: Menu, meals: Dict[str, Meal], ingredients: Dict[int, Ingredient]):
        total_g = 0.0
        totals = {k: 0.0 for k in NUTRITION_KEYS}
        for entry in menu.items:
            if entry.item_type == "ingredient":
                ing = _find_ing_by_name(entry.item_name, ingredients)
                if ing:
                    total_g += entry.amount
                    for k, v in ing.scaled_nutrition(entry.amount).items():
                        totals[k] += v
            elif entry.item_type == "meal":
                meal = meals.get(entry.item_name)
                if meal:
                    mg, mt = NutritionCalc.meal_totals(meal, ingredients)
                    total_g += mg * entry.amount
                    for k in NUTRITION_KEYS:
                        totals[k] += mt[k] * entry.amount
        return round(total_g, 1), {k: round(v, 2) for k, v in totals.items()}


def _find_ing_by_name(name: str, ingredients: Dict[int, Ingredient]):
    nl = name.lower()
    for ing in ingredients.values():
        if ing.name.lower() == nl:
            return ing
    return None


# ━━━━━━━━━━━━━━━━━━━━━ DIALOGS ━━━━━━━━━━━━━━━━━━━━━━━

class IngredientDialog(QDialog):
    def __init__(self, parent=None, ingredient=None, next_id=1):
        super().__init__(parent)
        self.setWindowTitle("Edit Ingredient" if ingredient else "New Ingredient")
        self.setMinimumWidth(480)
        self.ingredient = ingredient

        lay = QFormLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(24, 24, 24, 24)

        self.name_edit = QLineEdit(ingredient.name if ingredient else "")
        self.name_edit.setPlaceholderText("e.g. Chicken breast")
        self.brand_edit = QLineEdit(ingredient.brand if ingredient else "")
        self.brand_edit.setPlaceholderText("e.g. Organic Farm")
        self.product_edit = QLineEdit(ingredient.product_name if ingredient else "")
        self.package_edit = QLineEdit(ingredient.package_size if ingredient else "")

        lay.addRow("Name *", self.name_edit)
        lay.addRow("Brand", self.brand_edit)
        lay.addRow("Product Name", self.product_edit)
        lay.addRow("Package Size", self.package_edit)

        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.HLine)
        lay.addRow(sep)

        lbl = QLabel("Nutrition per 100 g")
        lbl.setStyleSheet("font-weight: 700; font-size: 14px; color: #0a84ff;")
        lay.addRow(lbl)

        self.spins: Dict[str, QDoubleSpinBox] = {}
        for key, label in NUTRITION_FIELDS:
            s = QDoubleSpinBox()
            s.setRange(0, 99999)
            s.setDecimals(2)
            s.setSingleStep(0.1)
            if ingredient:
                s.setValue(getattr(ingredient, key))
            self.spins[key] = s
            lay.addRow(label, s)

        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Cancel")
        cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        save = QPushButton("Save")
        save.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(save)
        lay.addRow(btns)

        self._next_id = next_id

    def get_ingredient(self):
        name = self.name_edit.text().strip()
        if not name:
            return None
        return Ingredient(
            id=self.ingredient.id if self.ingredient else self._next_id,
            name=name,
            brand=self.brand_edit.text().strip(),
            product_name=self.product_edit.text().strip(),
            package_size=self.package_edit.text().strip(),
            **{k: s.value() for k, s in self.spins.items()},
        )


class AddMealIngredientDialog(QDialog):
    def __init__(self, parent, ingredients: Dict[int, Ingredient]):
        super().__init__(parent)
        self.setWindowTitle("Add Ingredient to Meal")
        self.setMinimumWidth(440)
        self.ingredients = ingredients

        lay = QFormLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(24, 24, 24, 24)

        self.combo = QComboBox()
        self.combo.setEditable(True)
        self.combo.setInsertPolicy(QComboBox.NoInsert)
        names = sorted(
            [f"{i.name}  ({i.brand})" if i.brand else i.name for i in ingredients.values()]
        )
        self.combo.addItems(names)
        c = QCompleter(names)
        c.setCaseSensitivity(Qt.CaseInsensitive)
        c.setFilterMode(Qt.MatchContains)
        self.combo.setCompleter(c)
        lay.addRow("Ingredient", self.combo)

        self.amount = QDoubleSpinBox()
        self.amount.setRange(0.1, 99999)
        self.amount.setDecimals(1)
        self.amount.setValue(100)
        self.amount.setSuffix(" g")
        lay.addRow("Amount", self.amount)

        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Cancel")
        cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        add = QPushButton("Add")
        add.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(add)
        lay.addRow(btns)

    def get_result(self):
        text = self.combo.currentText().strip()
        for ing in self.ingredients.values():
            display = f"{ing.name}  ({ing.brand})" if ing.brand else ing.name
            if display == text or ing.name.lower() == text.lower():
                return ing.id, self.amount.value()
        return None


class AddMenuItemDialog(QDialog):
    def __init__(self, parent, ingredients, meals):
        super().__init__(parent)
        self.setWindowTitle("Add Item to Menu")
        self.setMinimumWidth(460)
        self.ingredients = ingredients
        self.meals = meals

        lay = QFormLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(24, 24, 24, 24)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["Ingredient", "Meal"])
        self.type_combo.currentIndexChanged.connect(self._on_type)
        lay.addRow("Type", self.type_combo)

        self.item_combo = QComboBox()
        self.item_combo.setEditable(True)
        self.item_combo.setInsertPolicy(QComboBox.NoInsert)
        lay.addRow("Item", self.item_combo)

        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0.01, 99999)
        self.amount_spin.setDecimals(1)
        self.amount_spin.setValue(100)
        self.amount_lbl = QLabel("Amount (g)")
        lay.addRow(self.amount_lbl, self.amount_spin)

        btns = QHBoxLayout()
        btns.addStretch()
        cancel = QPushButton("Cancel")
        cancel.setProperty("class", "secondary")
        cancel.clicked.connect(self.reject)
        add = QPushButton("Add")
        add.clicked.connect(self.accept)
        btns.addWidget(cancel)
        btns.addWidget(add)
        lay.addRow(btns)

        self._on_type(0)

    def _on_type(self, idx):
        self.item_combo.clear()
        if idx == 0:
            names = sorted(i.name for i in self.ingredients.values())
            self.item_combo.addItems(names)
            self.amount_lbl.setText("Amount (g)")
            self.amount_spin.setSuffix(" g")
            self.amount_spin.setValue(100)
        else:
            names = sorted(self.meals.keys())
            self.item_combo.addItems(names)
            self.amount_lbl.setText("Servings")
            self.amount_spin.setSuffix(" ×")
            self.amount_spin.setValue(1.0)
        c = QCompleter([self.item_combo.itemText(i) for i in range(self.item_combo.count())])
        c.setCaseSensitivity(Qt.CaseInsensitive)
        c.setFilterMode(Qt.MatchContains)
        self.item_combo.setCompleter(c)

    def get_result(self):
        t = "ingredient" if self.type_combo.currentIndex() == 0 else "meal"
        name = self.item_combo.currentText().strip()
        if not name:
            return None
        if t == "ingredient" and not _find_ing_by_name(name, self.ingredients):
            QMessageBox.warning(self, "Not Found", f"Ingredient '{name}' not found.\nAdd it in the Ingredients tab first.")
            return None
        if t == "meal" and name not in self.meals:
            QMessageBox.warning(self, "Not Found", f"Meal '{name}' not found.\nCreate it in the Meals tab first.")
            return None
        return MenuEntry(item_type=t, item_name=name, amount=self.amount_spin.value())


# ━━━━━━━━━━━━━━━━━━ SUMMARY WIDGET ━━━━━━━━━━━━━━━━━━━

class NutritionSummary(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("summaryCard")
        g = QGridLayout(self)
        g.setSpacing(6)
        self._labels: Dict[str, QLabel] = {}

        title = QLabel("Nutrition Summary")
        title.setStyleSheet("font-weight: 700; font-size: 13px; color: #0a84ff; background: transparent;")
        g.addWidget(title, 0, 0, 1, 4)

        self.weight_lbl = QLabel("Total: 0 g")
        self.weight_lbl.setStyleSheet("font-weight: 600; color: #98989d; background: transparent;")
        g.addWidget(self.weight_lbl, 0, 4, 1, 4, Qt.AlignRight)

        col = 0
        for key, label in NUTRITION_FIELDS:
            if key == "energy_kj":
                continue
            vl = QLabel("0")
            vl.setAlignment(Qt.AlignCenter)
            vl.setStyleSheet("font-weight: 700; font-size: 15px; color: #f5f5f7; background: transparent;")
            nl = QLabel(label.replace(" (g)", "").replace(" (kcal)", ""))
            nl.setAlignment(Qt.AlignCenter)
            nl.setStyleSheet("font-size: 10px; color: #636366; background: transparent;")
            g.addWidget(vl, 1, col)
            g.addWidget(nl, 2, col)
            self._labels[key] = vl
            col += 1

    def update(self, total_g, totals):
        self.weight_lbl.setText(f"Total: {total_g:.0f} g")
        for k, lbl in self._labels.items():
            lbl.setText(f"{totals.get(k, 0):.1f}")


# ━━━━━━━━━━━━━━━━━━ TABLE HELPERS ━━━━━━━━━━━━━━━━━━━━

def _num_item(value, suffix=""):
    if isinstance(value, float):
        text = f"{value:.2f}" if value != int(value) else str(int(value))
    else:
        text = str(value)
    item = QTableWidgetItem(text + suffix)
    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
    try:
        item.setData(Qt.UserRole, float(value))
    except (ValueError, TypeError):
        pass
    return item


# ━━━━━━━━━━━━━━━━━ INGREDIENTS TAB ━━━━━━━━━━━━━━━━━━━

class IngredientsTab(QWidget):
    data_changed = Signal()

    def __init__(self):
        super().__init__()
        self.ingredients: Dict[int, Ingredient] = {}
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(14)

        hdr = QHBoxLayout()
        t = QLabel("Ingredients")
        t.setObjectName("sectionTitle")
        hdr.addWidget(t)
        hdr.addStretch()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…")
        self.search.setFixedWidth(280)
        self.search.textChanged.connect(self._filter)
        hdr.addWidget(self.search)
        add = QPushButton("＋  Add Ingredient")
        add.clicked.connect(self._add)
        hdr.addWidget(add)
        lay.addLayout(hdr)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)
        self.table.doubleClicked.connect(self._edit)
        cols = ["ID", "Name", "Brand", "Product"] + NUTRITION_LABELS + ["Pkg Size"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        lay.addWidget(self.table)

        btns = QHBoxLayout()
        btns.addStretch()
        eb = QPushButton("Edit")
        eb.setProperty("class", "secondary")
        eb.clicked.connect(self._edit)
        db = QPushButton("Delete")
        db.setProperty("class", "danger")
        db.clicked.connect(self._delete)
        btns.addWidget(eb)
        btns.addWidget(db)
        lay.addLayout(btns)

    def load(self, ingredients):
        self.ingredients = ingredients
        self._populate()

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
        self.table.resizeColumnsToContents()
        self.table.setSortingEnabled(True)

    def _filter(self, text):
        self._populate(text)

    def _next_id(self):
        return max(self.ingredients.keys(), default=0) + 1

    def _add(self):
        d = IngredientDialog(self, next_id=self._next_id())
        if d.exec() == QDialog.Accepted:
            ing = d.get_ingredient()
            if ing:
                self.ingredients[ing.id] = ing
                self._populate(self.search.text())
                self.data_changed.emit()

    def _edit(self):
        row = self.table.currentRow()
        if row < 0:
            return
        iid = int(self.table.item(row, 0).text())
        ing = self.ingredients.get(iid)
        if not ing:
            return
        d = IngredientDialog(self, ingredient=ing)
        if d.exec() == QDialog.Accepted:
            u = d.get_ingredient()
            if u:
                self.ingredients[u.id] = u
                self._populate(self.search.text())
                self.data_changed.emit()

    def _delete(self):
        row = self.table.currentRow()
        if row < 0:
            return
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

class MealsTab(QWidget):
    data_changed = Signal()

    def __init__(self, ing_tab: IngredientsTab):
        super().__init__()
        self.ing_tab = ing_tab
        self.meals: Dict[str, Meal] = {}
        self._build()

    @property
    def ingredients(self):
        return self.ing_tab.ingredients

    def _build(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(14)

        # Left
        left = QVBoxLayout()
        left.setSpacing(10)
        t = QLabel("Meals")
        t.setObjectName("sectionTitle")
        left.addWidget(t)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self._filter_list)
        left.addWidget(self.search)
        self.meal_list = QListWidget()
        self.meal_list.currentItemChanged.connect(self._on_select)
        left.addWidget(self.meal_list)
        lb = QHBoxLayout()
        lb.setSpacing(6)
        a = QPushButton("＋ New")
        a.clicked.connect(self._add_meal)
        r = QPushButton("Rename")
        r.setProperty("class", "secondary")
        r.clicked.connect(self._rename_meal)
        d = QPushButton("Delete")
        d.setProperty("class", "danger")
        d.clicked.connect(self._delete_meal)
        lb.addWidget(a)
        lb.addWidget(r)
        lb.addWidget(d)
        left.addLayout(lb)
        lw = QWidget()
        lw.setLayout(left)
        lw.setFixedWidth(270)
        lay.addWidget(lw)

        # Right
        right = QVBoxLayout()
        right.setSpacing(10)
        self.title_lbl = QLabel("Select a meal")
        self.title_lbl.setObjectName("sectionTitle")
        right.addWidget(self.title_lbl)
        self.summary = NutritionSummary()
        right.addWidget(self.summary)
        self.per100_lbl = QLabel("")
        self.per100_lbl.setStyleSheet("font-size: 11px; color: #636366; padding: 2px 4px; background: transparent;")
        right.addWidget(self.per100_lbl)

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
        ai = QPushButton("＋ Add Ingredient")
        ai.clicked.connect(self._add_ing)
        ea = QPushButton("Edit Amount")
        ea.setProperty("class", "secondary")
        ea.clicked.connect(self._edit_amount)
        ri = QPushButton("Remove")
        ri.setProperty("class", "danger")
        ri.clicked.connect(self._remove_ing)
        rb.addWidget(ai)
        rb.addWidget(ea)
        rb.addStretch()
        rb.addWidget(ri)
        right.addLayout(rb)
        lay.addLayout(right, 1)

    def load(self, meals):
        self.meals = meals
        self._populate_list()

    def _populate_list(self, ft=""):
        ft = ft.lower()
        self.meal_list.blockSignals(True)
        cur = self.meal_list.currentItem()
        cur_name = cur.text() if cur else None
        self.meal_list.clear()
        for n in sorted(self.meals, key=str.lower):
            if ft and ft not in n.lower():
                continue
            self.meal_list.addItem(n)
        if cur_name:
            found = self.meal_list.findItems(cur_name, Qt.MatchExactly)
            if found:
                self.meal_list.setCurrentItem(found[0])
        self.meal_list.blockSignals(False)
        if self.meal_list.currentItem():
            self._on_select(self.meal_list.currentItem())

    def _filter_list(self, t):
        self._populate_list(t)

    def _cur(self):
        i = self.meal_list.currentItem()
        return self.meals.get(i.text()) if i else None

    def _on_select(self, cur, prev=None):
        meal = self._cur()
        if not meal:
            self.title_lbl.setText("Select a meal")
            self.detail.setRowCount(0)
            self.summary.update(0, {k: 0 for k in NUTRITION_KEYS})
            self.per100_lbl.setText("")
            return
        self.title_lbl.setText(meal.name)
        self._refresh()

    def _refresh(self):
        meal = self._cur()
        if not meal:
            return
        self.detail.setRowCount(len(meal.items))
        for r, mi in enumerate(meal.items):
            ing = self.ingredients.get(mi.ingredient_id)
            self.detail.setItem(r, 0, QTableWidgetItem(ing.name if ing else f"[ID {mi.ingredient_id}]"))
            self.detail.setItem(r, 1, _num_item(mi.amount_grams))
            if ing:
                sc = ing.scaled_nutrition(mi.amount_grams)
                for c, key in enumerate(NUTRITION_KEYS):
                    self.detail.setItem(r, 2 + c, _num_item(sc[key]))
        self.detail.resizeColumnsToContents()
        tg, tt = NutritionCalc.meal_totals(meal, self.ingredients)
        self.summary.update(tg, tt)
        p = NutritionCalc.per100(tg, tt)
        parts = [f"{NUTRITION_FIELDS[i][1]}: {p[NUTRITION_KEYS[i]]:.1f}"
                 for i in range(len(NUTRITION_KEYS)) if NUTRITION_KEYS[i] != "energy_kj"]
        self.per100_lbl.setText("Per 100 g:  " + "  ·  ".join(parts))

    def _add_meal(self):
        n, ok = QInputDialog.getText(self, "New Meal", "Meal name:")
        n = n.strip() if ok else ""
        if not n:
            return
        if n in self.meals:
            QMessageBox.warning(self, "Exists", f"'{n}' already exists.")
            return
        self.meals[n] = Meal(name=n)
        self._populate_list()
        found = self.meal_list.findItems(n, Qt.MatchExactly)
        if found:
            self.meal_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _rename_meal(self):
        m = self._cur()
        if not m:
            return
        nn, ok = QInputDialog.getText(self, "Rename", "New name:", text=m.name)
        nn = nn.strip() if ok else ""
        if not nn or nn == m.name:
            return
        if nn in self.meals:
            QMessageBox.warning(self, "Exists", f"'{nn}' already exists.")
            return
        del self.meals[m.name]
        m.name = nn
        self.meals[nn] = m
        self._populate_list()
        found = self.meal_list.findItems(nn, Qt.MatchExactly)
        if found:
            self.meal_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _delete_meal(self):
        m = self._cur()
        if not m:
            return
        if QMessageBox.question(self, "Delete", f"Delete meal '{m.name}'?") == QMessageBox.Yes:
            del self.meals[m.name]
            self._populate_list()
            self.data_changed.emit()

    def _add_ing(self):
        m = self._cur()
        if not m:
            QMessageBox.information(self, "No Meal", "Select or create a meal first.")
            return
        d = AddMealIngredientDialog(self, self.ingredients)
        if d.exec() == QDialog.Accepted:
            r = d.get_result()
            if r:
                m.items.append(MealIngredient(ingredient_id=r[0], amount_grams=r[1]))
                self._refresh()
                self.data_changed.emit()

    def _edit_amount(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items):
            return
        mi = m.items[row]
        amt, ok = QInputDialog.getDouble(self, "Edit Amount", "Amount (g):", mi.amount_grams, 0.1, 99999, 1)
        if ok:
            mi.amount_grams = amt
            self._refresh()
            self.data_changed.emit()

    def _remove_ing(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items):
            return
        m.items.pop(row)
        self._refresh()
        self.data_changed.emit()


# ━━━━━━━━━━━━━━━━━━━ MENUS TAB ━━━━━━━━━━━━━━━━━━━━━━

class MenusTab(QWidget):
    data_changed = Signal()

    def __init__(self, ing_tab: IngredientsTab, meals_tab: MealsTab):
        super().__init__()
        self.ing_tab = ing_tab
        self.meals_tab = meals_tab
        self.menus: Dict[str, Menu] = {}
        self._build()

    @property
    def ingredients(self):
        return self.ing_tab.ingredients

    @property
    def meals(self):
        return self.meals_tab.meals

    def _build(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(14)

        # Left
        left = QVBoxLayout()
        left.setSpacing(10)
        t = QLabel("Menus")
        t.setObjectName("sectionTitle")
        left.addWidget(t)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search…")
        self.search.textChanged.connect(self._filter_list)
        left.addWidget(self.search)
        self.menu_list = QListWidget()
        self.menu_list.currentItemChanged.connect(self._on_select)
        left.addWidget(self.menu_list)
        lb = QHBoxLayout()
        lb.setSpacing(6)
        a = QPushButton("＋ New")
        a.clicked.connect(self._add_menu)
        r = QPushButton("Rename")
        r.setProperty("class", "secondary")
        r.clicked.connect(self._rename_menu)
        d = QPushButton("Delete")
        d.setProperty("class", "danger")
        d.clicked.connect(self._delete_menu)
        lb.addWidget(a)
        lb.addWidget(r)
        lb.addWidget(d)
        left.addLayout(lb)
        lw = QWidget()
        lw.setLayout(left)
        lw.setFixedWidth(270)
        lay.addWidget(lw)

        # Right
        right = QVBoxLayout()
        right.setSpacing(10)
        self.title_lbl = QLabel("Select a menu")
        self.title_lbl.setObjectName("sectionTitle")
        right.addWidget(self.title_lbl)
        self.summary = NutritionSummary()
        right.addWidget(self.summary)
        self.per100_lbl = QLabel("")
        self.per100_lbl.setStyleSheet("font-size: 11px; color: #636366; padding: 2px 4px; background: transparent;")
        right.addWidget(self.per100_lbl)

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
        ai = QPushButton("＋ Add Item")
        ai.clicked.connect(self._add_item)
        ea = QPushButton("Edit Amount")
        ea.setProperty("class", "secondary")
        ea.clicked.connect(self._edit_amount)
        ri = QPushButton("Remove")
        ri.setProperty("class", "danger")
        ri.clicked.connect(self._remove_item)
        rb.addWidget(ai)
        rb.addWidget(ea)
        rb.addStretch()
        rb.addWidget(ri)
        right.addLayout(rb)
        lay.addLayout(right, 1)

    def load(self, menus):
        self.menus = menus
        self._populate_list()

    def _populate_list(self, ft=""):
        ft = ft.lower()
        self.menu_list.blockSignals(True)
        cur = self.menu_list.currentItem()
        cur_name = cur.text() if cur else None
        self.menu_list.clear()
        for n in sorted(self.menus, key=str.lower):
            if ft and ft not in n.lower():
                continue
            self.menu_list.addItem(n)
        if cur_name:
            found = self.menu_list.findItems(cur_name, Qt.MatchExactly)
            if found:
                self.menu_list.setCurrentItem(found[0])
        self.menu_list.blockSignals(False)
        if self.menu_list.currentItem():
            self._on_select(self.menu_list.currentItem())

    def _filter_list(self, t):
        self._populate_list(t)

    def _cur(self):
        i = self.menu_list.currentItem()
        return self.menus.get(i.text()) if i else None

    def _on_select(self, cur, prev=None):
        menu = self._cur()
        if not menu:
            self.title_lbl.setText("Select a menu")
            self.detail.setRowCount(0)
            self.summary.update(0, {k: 0 for k in NUTRITION_KEYS})
            self.per100_lbl.setText("")
            return
        self.title_lbl.setText(menu.name)
        self._refresh()

    def _refresh(self):
        menu = self._cur()
        if not menu:
            return
        self.detail.setRowCount(len(menu.items))
        for r, entry in enumerate(menu.items):
            self.detail.setItem(r, 0, QTableWidgetItem(entry.item_type.capitalize()))
            self.detail.setItem(r, 1, QTableWidgetItem(entry.item_name))

            if entry.item_type == "ingredient":
                self.detail.setItem(r, 2, _num_item(entry.amount, " g"))
                self.detail.setItem(r, 3, _num_item(entry.amount, " g"))
                ing = _find_ing_by_name(entry.item_name, self.ingredients)
                if ing:
                    sc = ing.scaled_nutrition(entry.amount)
                    for c, key in enumerate(NUTRITION_KEYS):
                        self.detail.setItem(r, 4 + c, _num_item(sc[key]))
            else:
                self.detail.setItem(r, 2, _num_item(entry.amount, " ×"))
                meal = self.meals.get(entry.item_name)
                if meal:
                    mg, mt = NutritionCalc.meal_totals(meal, self.ingredients)
                    self.detail.setItem(r, 3, _num_item(round(mg * entry.amount, 1), " g"))
                    for c, key in enumerate(NUTRITION_KEYS):
                        self.detail.setItem(r, 4 + c, _num_item(round(mt[key] * entry.amount, 2)))
        self.detail.resizeColumnsToContents()

        tg, tt = NutritionCalc.menu_totals(menu, self.meals, self.ingredients)
        self.summary.update(tg, tt)
        p = NutritionCalc.per100(tg, tt)
        parts = [f"{NUTRITION_FIELDS[i][1]}: {p[NUTRITION_KEYS[i]]:.1f}"
                 for i in range(len(NUTRITION_KEYS)) if NUTRITION_KEYS[i] != "energy_kj"]
        self.per100_lbl.setText("Per 100 g:  " + "  ·  ".join(parts))

    def _add_menu(self):
        n, ok = QInputDialog.getText(self, "New Menu", "Menu name:")
        n = n.strip() if ok else ""
        if not n:
            return
        if n in self.menus:
            QMessageBox.warning(self, "Exists", f"'{n}' already exists.")
            return
        self.menus[n] = Menu(name=n)
        self._populate_list()
        found = self.menu_list.findItems(n, Qt.MatchExactly)
        if found:
            self.menu_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _rename_menu(self):
        m = self._cur()
        if not m:
            return
        nn, ok = QInputDialog.getText(self, "Rename", "New name:", text=m.name)
        nn = nn.strip() if ok else ""
        if not nn or nn == m.name:
            return
        if nn in self.menus:
            QMessageBox.warning(self, "Exists", f"'{nn}' already exists.")
            return
        del self.menus[m.name]
        m.name = nn
        self.menus[nn] = m
        self._populate_list()
        found = self.menu_list.findItems(nn, Qt.MatchExactly)
        if found:
            self.menu_list.setCurrentItem(found[0])
        self.data_changed.emit()

    def _delete_menu(self):
        m = self._cur()
        if not m:
            return
        if QMessageBox.question(self, "Delete", f"Delete menu '{m.name}'?") == QMessageBox.Yes:
            del self.menus[m.name]
            self._populate_list()
            self.data_changed.emit()

    def _add_item(self):
        m = self._cur()
        if not m:
            QMessageBox.information(self, "No Menu", "Select or create a menu first.")
            return
        d = AddMenuItemDialog(self, self.ingredients, self.meals)
        if d.exec() == QDialog.Accepted:
            e = d.get_result()
            if e:
                m.items.append(e)
                self._refresh()
                self.data_changed.emit()

    def _edit_amount(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items):
            return
        entry = m.items[row]
        label = "Amount (g):" if entry.item_type == "ingredient" else "Servings:"
        amt, ok = QInputDialog.getDouble(self, "Edit Amount", label, entry.amount, 0.01, 99999, 2)
        if ok:
            entry.amount = amt
            self._refresh()
            self.data_changed.emit()

    def _remove_item(self):
        m = self._cur()
        row = self.detail.currentRow()
        if not m or row < 0 or row >= len(m.items):
            return
        m.items.pop(row)
        self._refresh()
        self.data_changed.emit()


# ━━━━━━━━━━━━━━━━━━ MAIN WINDOW ━━━━━━━━━━━━━━━━━━━━━

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nutrition Planner")
        self.resize(1240, 780)

        self.db: Optional[DatabaseManager] = None
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)

        # Central container
        central = QWidget()
        self.setCentralWidget(central)
        self.root_layout = QVBoxLayout(central)
        self.root_layout.setContentsMargins(0, 0, 0, 0)
        self.root_layout.setSpacing(0)

        # ── Database bar (top) ──
        self.db_bar = QFrame()
        self.db_bar.setObjectName("dbBar")
        self.db_bar.setFixedHeight(52)
        db_lay = QHBoxLayout(self.db_bar)
        db_lay.setContentsMargins(20, 0, 20, 0)
        db_lay.setSpacing(12)

        db_icon = QLabel("◉")
        db_icon.setStyleSheet("font-size: 16px; color: #30d158; background: transparent;")
        db_lay.addWidget(db_icon)
        self.db_icon = db_icon

        db_label = QLabel("Database")
        db_label.setStyleSheet("color: #98989d; font-size: 12px; font-weight: 600; background: transparent;")
        db_lay.addWidget(db_label)

        self.db_path_label = QLabel("No database selected")
        self.db_path_label.setObjectName("dbPath")
        self.db_path_label.setStyleSheet("color: #f5f5f7; font-size: 13px; font-weight: 500; background: transparent;")
        db_lay.addWidget(self.db_path_label, 1)

        browse_btn = QPushButton("Change…")
        browse_btn.setProperty("class", "dbButton")
        browse_btn.setFixedWidth(100)
        browse_btn.clicked.connect(self._browse_db)
        db_lay.addWidget(browse_btn)

        self.root_layout.addWidget(self.db_bar)

        # ── Tabs ──
        self.ing_tab = IngredientsTab()
        self.meals_tab = MealsTab(self.ing_tab)
        self.menus_tab = MenusTab(self.ing_tab, self.meals_tab)

        self.tabs = QTabWidget()
        self.tabs.addTab(self.ing_tab, "   Ingredients   ")
        self.tabs.addTab(self.meals_tab, "   Meals   ")
        self.tabs.addTab(self.menus_tab, "   Menus   ")
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.root_layout.addWidget(self.tabs, 1)

        # ── Status bar ──
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # ── Auto-save timer (2 s debounce — industry standard) ──
        self._save_timer = QTimer()
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(2000)
        self._save_timer.timeout.connect(self._do_save)

        self.ing_tab.data_changed.connect(self._schedule_save)
        self.meals_tab.data_changed.connect(self._schedule_save)
        self.menus_tab.data_changed.connect(self._schedule_save)

        # ── Menu bar ──
        mb = self.menuBar()
        mb.setStyleSheet("QMenuBar { background: #2c2c2e; color: #f5f5f7; }"
                         "QMenuBar::item:selected { background: #48484a; }"
                         "QMenu { background: #3a3a3c; color: #f5f5f7; border: 1px solid #48484a; }"
                         "QMenu::item:selected { background: #0a84ff; }")
        fm = mb.addMenu("File")
        sa = QAction("Save Now", self)
        sa.setShortcut("Ctrl+S")
        sa.triggered.connect(self._do_save)
        fm.addAction(sa)
        fm.addSeparator()
        qa = QAction("Quit", self)
        qa.setShortcut("Ctrl+Q")
        qa.triggered.connect(self.close)
        fm.addAction(qa)

        # ── Load last database ──
        last = self.settings.value(SETTINGS_LAST_DB, "")
        if last and Path(last).is_file():
            self._open_db(last)
        else:
            self.status.showMessage("Select a database to get started")
            self._set_db_indicator(False)

    # ── Database management ──

    def _browse_db(self):
        start = str(Path(self.db.path).parent) if self.db else str(Path.home())
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Nutrition Database", start,
            "Excel Files (*.xlsx);;All Files (*)",
        )
        if path:
            self._open_db(path)

    def _open_db(self, path: str):
        try:
            self.db = DatabaseManager(path)
            self.settings.setValue(SETTINGS_LAST_DB, path)

            self.ing_tab.load(self.db.load_ingredients())
            self.meals_tab.load(self.db.load_meals())
            self.menus_tab.load(self.db.load_menus())

            display = self._shorten_path(path)
            self.db_path_label.setText(display)
            self.db_path_label.setToolTip(path)
            self._set_db_indicator(True)
            self.status.showMessage(f"Loaded — {len(self.ing_tab.ingredients)} ingredients")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open database:\n\n{e}")
            self._set_db_indicator(False)

    def _set_db_indicator(self, connected: bool):
        if connected:
            self.db_icon.setText("◉")
            self.db_icon.setStyleSheet("font-size: 16px; color: #30d158; background: transparent;")
        else:
            self.db_icon.setText("◯")
            self.db_icon.setStyleSheet("font-size: 16px; color: #636366; background: transparent;")

    @staticmethod
    def _shorten_path(p: str, max_len=60) -> str:
        if len(p) <= max_len:
            return p
        parts = Path(p).parts
        return str(Path(parts[0], "…", *parts[-2:]))

    # ── Save ──

    def _schedule_save(self):
        if not self.db:
            return
        self.status.showMessage("Unsaved changes…")
        self._save_timer.start()

    def _do_save(self):
        if not self.db:
            return
        try:
            self.db.save_all(
                self.ing_tab.ingredients,
                self.meals_tab.meals,
                self.menus_tab.menus,
            )
            self.status.showMessage(f"Saved  ·  {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            self.status.showMessage(f"SAVE ERROR: {e}")
            QMessageBox.critical(self, "Save Error",
                                 f"Failed to save:\n\n{e}\n\nYour backup is intact.")

    def _on_tab_changed(self, idx):
        if idx == 1:
            self.meals_tab._refresh()
        elif idx == 2:
            self.menus_tab._refresh()

    def closeEvent(self, event):
        if self.db:
            try:
                self.db.save_all(
                    self.ing_tab.ingredients,
                    self.meals_tab.meals,
                    self.menus_tab.menus,
                )
            except Exception as e:
                if QMessageBox.warning(
                    self, "Save Error",
                    f"Could not save:\n{e}\n\nQuit anyway?",
                    QMessageBox.Yes | QMessageBox.No,
                ) == QMessageBox.No:
                    event.ignore()
                    return
        event.accept()


# ━━━━━━━━━━━━━━━━━━━━ ENTRY POINT ━━━━━━━━━━━━━━━━━━━━

def main():
    app = QApplication(sys.argv)
    app.setFont(QFontGui("SF Pro Text", 12))
    app.setStyleSheet(DARK_STYLE)
    app.setApplicationName(SETTINGS_APP)
    app.setOrganizationName(SETTINGS_ORG)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()