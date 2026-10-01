import sys
import sqlite3
import os
import shutil
import csv
import re
try:
    import winreg
except ImportError:
    winreg = None
import configparser
from datetime import date, datetime
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
import openpyxl
from pathlib import Path

CURRENCY_OPTIONS = {
    "USD": {
        "label": "US Dollar",
        "symbol": "$",
        "thousands_sep": ",",
        "decimal_sep": ".",
        "decimals": 2,
        "grouping": "western",
        "symbol_position": "prefix",
        "space_between": False,
    },
    "IDR": {
        "label": "Indonesian Rupiah",
        "symbol": "Rp",
        "thousands_sep": ".",
        "decimal_sep": ",",
        "decimals": 0,
        "grouping": "western",
        "symbol_position": "prefix",
        "space_between": True,
    },
}
DEFAULT_CURRENCY_CODE = "USD"
DEFAULT_LANGUAGE_CODE = "en_US"
DEFAULT_APP_ICON_PATH = r"E:\MY PROJECT\app.ico"


DEFAULT_THEME_CODE = "dark"
THEME_ORDER = ("light", "dark", "blue", "purple", "red", "orange", "classic")

THEME_OPTIONS = {
    "light": {
        "label": "LIGHT",
        "window": "#f3f6fa", "panel": "#ffffff", "panel_alt": "#f8fafc",
        "surface": "#e8eef5", "surface_hover": "#dce6f1", "surface_pressed": "#cbd8e6",
        "input": "#ffffff", "text": "#17202a", "muted": "#5f6b7a",
        "border": "#c4cfdb", "border_strong": "#9eacbc", "accent": "#2563eb",
        "accent_hover": "#1d4ed8", "accent_pressed": "#1e40af", "accent_text": "#ffffff",
        "tab_selected": "#dbeafe", "header": "#e5ebf2", "selection": "#2563eb",
        "disabled_bg": "#e5e9ef", "disabled_text": "#8a94a3", "menu": "#ffffff",
        "scroll_bg": "#e7ecf2", "scroll_handle": "#aab5c2", "tooltip": "#1f2937",
        "tooltip_text": "#ffffff", "danger": "#c62828", "success": "#2e7d32",
        "warning": "#b26a00", "neutral": "#657384", "radius": "4px", "group_radius": "6px",
    },
    "dark": {
        "label": "DARK",
        "window": "#0f141a", "panel": "#111820", "panel_alt": "#111a23",
        "surface": "#263445", "surface_hover": "#32455c", "surface_pressed": "#1c2938",
        "input": "#0b1117", "text": "#e6edf3", "muted": "#aeb8c4",
        "border": "#364250", "border_strong": "#4d6075", "accent": "#1f6feb",
        "accent_hover": "#3d8bfd", "accent_pressed": "#1557b0", "accent_text": "#ffffff",
        "tab_selected": "#243447", "header": "#1b2735", "selection": "#1f6feb",
        "disabled_bg": "#121820", "disabled_text": "#697586", "menu": "#161d26",
        "scroll_bg": "#10171f", "scroll_handle": "#3b4858", "tooltip": "#243447",
        "tooltip_text": "#ffffff", "danger": "#b4232c", "success": "#238636",
        "warning": "#9e6a03", "neutral": "#46515f", "radius": "4px", "group_radius": "6px",
    },
    "blue": {
        "label": "BLUE",
        "window": "#071522", "panel": "#0b1d2e", "panel_alt": "#0d2438",
        "surface": "#123653", "surface_hover": "#194b70", "surface_pressed": "#0c2941",
        "input": "#06111c", "text": "#e8f4ff", "muted": "#a9c5dc",
        "border": "#24506f", "border_strong": "#39769e", "accent": "#2196f3",
        "accent_hover": "#42a5f5", "accent_pressed": "#1475bd", "accent_text": "#ffffff",
        "tab_selected": "#154365", "header": "#10304a", "selection": "#1976d2",
        "disabled_bg": "#10202e", "disabled_text": "#66859d", "menu": "#0d2335",
        "scroll_bg": "#091a28", "scroll_handle": "#2c5f82", "tooltip": "#16466a",
        "tooltip_text": "#ffffff", "danger": "#c0394b", "success": "#168f63",
        "warning": "#aa7700", "neutral": "#41657c", "radius": "4px", "group_radius": "6px",
    },
    "purple": {
        "label": "PURPLE",
        "window": "#160f21", "panel": "#1d142b", "panel_alt": "#241833",
        "surface": "#44265f", "surface_hover": "#593278", "surface_pressed": "#351d4b",
        "input": "#100b18", "text": "#f2eaff", "muted": "#c6afd8",
        "border": "#5b3a70", "border_strong": "#7b5293", "accent": "#9c4dcc",
        "accent_hover": "#b264df", "accent_pressed": "#75369f", "accent_text": "#ffffff",
        "tab_selected": "#4a2864", "header": "#342047", "selection": "#8e44ad",
        "disabled_bg": "#21172c", "disabled_text": "#826d92", "menu": "#21162f",
        "scroll_bg": "#180f22", "scroll_handle": "#654079", "tooltip": "#4b2963",
        "tooltip_text": "#ffffff", "danger": "#bd334f", "success": "#2b8f68",
        "warning": "#a76c00", "neutral": "#63506f", "radius": "4px", "group_radius": "6px",
    },
    "red": {
        "label": "RED",
        "window": "#1b0d10", "panel": "#241115", "panel_alt": "#2c151a",
        "surface": "#5a252d", "surface_hover": "#74313b", "surface_pressed": "#461b22",
        "input": "#13090b", "text": "#fff0f1", "muted": "#d9afb4",
        "border": "#71343d", "border_strong": "#934955", "accent": "#e53935",
        "accent_hover": "#ef5350", "accent_pressed": "#b71c1c", "accent_text": "#ffffff",
        "tab_selected": "#612832", "header": "#451e25", "selection": "#d32f2f",
        "disabled_bg": "#281417", "disabled_text": "#926d72", "menu": "#291317",
        "scroll_bg": "#1c0d10", "scroll_handle": "#74323b", "tooltip": "#642a33",
        "tooltip_text": "#ffffff", "danger": "#d32f2f", "success": "#2e7d55",
        "warning": "#a96e00", "neutral": "#715158", "radius": "4px", "group_radius": "6px",
    },
    "orange": {
        "label": "ORANGE",
        "window": "#1a120a", "panel": "#24180d", "panel_alt": "#2c1d10",
        "surface": "#5b3818", "surface_hover": "#754a20", "surface_pressed": "#462a12",
        "input": "#120c07", "text": "#fff4e8", "muted": "#d7b997",
        "border": "#714823", "border_strong": "#966234", "accent": "#f57c00",
        "accent_hover": "#fb8c00", "accent_pressed": "#c65f00", "accent_text": "#ffffff",
        "tab_selected": "#633d1a", "header": "#472c14", "selection": "#ef6c00",
        "disabled_bg": "#281b10", "disabled_text": "#92785e", "menu": "#291b0f",
        "scroll_bg": "#1c1209", "scroll_handle": "#76502c", "tooltip": "#68401b",
        "tooltip_text": "#ffffff", "danger": "#c43b32", "success": "#2d8756",
        "warning": "#d87800", "neutral": "#715d49", "radius": "4px", "group_radius": "6px",
    },
    "classic": {
        "label": "CLASSIC",
        "window": "#d4d0c8", "panel": "#ece9d8", "panel_alt": "#f5f3e8",
        "surface": "#ece9d8", "surface_hover": "#f7f5ec", "surface_pressed": "#c8c4bb",
        "input": "#ffffff", "text": "#000000", "muted": "#404040",
        "border": "#9a9a9a", "border_strong": "#6f6f6f", "accent": "#0a64ad",
        "accent_hover": "#1976bd", "accent_pressed": "#064b82", "accent_text": "#ffffff",
        "tab_selected": "#ffffff", "header": "#d4d0c8", "selection": "#0a64ad",
        "disabled_bg": "#d4d0c8", "disabled_text": "#777777", "menu": "#ece9d8",
        "scroll_bg": "#d4d0c8", "scroll_handle": "#a0a0a0", "tooltip": "#ffffe1",
        "tooltip_text": "#000000", "danger": "#b00020", "success": "#2e7d32",
        "warning": "#9a6100", "neutral": "#777777", "radius": "1px", "group_radius": "2px",
    },
}


def get_theme_config(theme_code=None):
    code = str(theme_code or DEFAULT_THEME_CODE).strip().lower()
    return THEME_OPTIONS.get(code, THEME_OPTIONS[DEFAULT_THEME_CODE])


def build_theme_stylesheet(theme_code=None):
    """Build a fully opaque control theme over transparent layout containers."""
    p = get_theme_config(theme_code)
    return f"""
        /* Layout-only widgets must stay transparent so nested panels do not
           cover each other with the window color. Real surfaces are styled
           explicitly below. */
        QWidget {{
            color: {p['text']};
            background-color: transparent;
            font-size: 8pt;
        }}
        QMainWindow, QDialog {{
            background-color: {p['window']};
            color: {p['text']};
        }}
        QMainWindow > QWidget, QDialog > QWidget,
        QTabWidget, QStackedWidget {{
            background-color: {p['window']};
        }}
        QFrame, QSplitter, QDialogButtonBox {{ background-color: transparent; }}
        QSplitter::handle {{ background-color: {p['border']}; }}
        QSplitter::handle:hover {{ background-color: {p['accent']}; }}

        QMenuBar {{
            background-color: {p['menu']}; color: {p['text']};
            border-bottom: 1px solid {p['border']};
        }}
        /* Menu entries use an opaque base and opaque hover/pressed states.
           This prevents the native style from drawing translucent overlays. */
        QMenuBar::item {{
            background-color: {p['menu']}; color: {p['text']};
            padding: 4px 8px;
        }}
        QMenuBar::item:selected, QMenu::item:selected {{
            background-color: {p['surface_hover']}; color: {p['text']};
        }}
        QMenuBar::item:pressed, QMenu::item:pressed {{
            background-color: {p['surface_pressed']}; color: {p['text']};
        }}
        QMenu {{
            background-color: {p['menu']}; color: {p['text']};
            border: 1px solid {p['border_strong']}; padding: 3px;
        }}
        QMenu::item {{
            padding: 5px 24px 5px 9px;
            background-color: {p['menu']}; color: {p['text']};
        }}
        QMenu::item:disabled {{ color: {p['disabled_text']}; }}
        QMenu::separator {{
            height: 1px; background-color: {p['border']}; margin: 4px 7px;
        }}

        QTabWidget::pane {{
            border: 1px solid {p['border']}; background-color: {p['window']};
        }}
        QTabBar {{ background-color: transparent; }}
        QTabBar::tab {{
            background-color: {p['menu']}; color: {p['muted']};
            border: 1px solid {p['border']};
            padding: 7px 14px; margin-right: 2px;
        }}
        QTabBar::tab:selected {{
            background-color: {p['tab_selected']}; color: {p['text']};
            border-bottom: 2px solid {p['accent']};
        }}
        QTabBar::tab:hover:!selected {{
            background-color: {p['surface_hover']}; color: {p['text']};
            border-color: {p['accent']};
        }}
        QTabBar::tab:pressed:!selected {{
            background-color: {p['surface_pressed']}; color: {p['text']};
        }}

        QWidget#inputPanel, QWidget#tablePanel {{
            background-color: {p['panel']};
        }}
        QScrollArea {{ border: 0; background-color: transparent; }}
        QScrollArea > QWidget > QWidget {{ background-color: transparent; }}
        QScrollArea#inputPanelScroll,
        QScrollArea#inputPanelScroll > QWidget > QWidget {{
            background-color: {p['panel']};
        }}

        QGroupBox {{
            font-weight: 600; border: 1px solid {p['border']};
            border-radius: {p['group_radius']}; margin-top: 8px;
            padding-top: 8px; background-color: {p['panel']};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; left: 9px; padding: 0 5px;
            color: {p['accent']}; background-color: {p['panel']};
        }}
        QLabel, QCheckBox, QRadioButton {{ background-color: transparent; }}
        QCheckBox:hover, QRadioButton:hover {{
            background-color: {p['surface_hover']}; color: {p['text']};
            border-radius: {p['radius']};
        }}
        QCheckBox:pressed, QRadioButton:pressed {{
            background-color: {p['surface_pressed']}; color: {p['text']};
        }}
        QLabel#accentLabel {{ color: {p['accent']}; }}
        QLabel#successLabel {{ color: {p['success']}; }}
        QLabel#alertLabel {{ color: {p['danger']}; font-weight: bold; }}
        QFrame[frameShape="4"], QFrame[frameShape="5"] {{
            color: {p['border']}; background-color: transparent;
        }}

        QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QDateEdit,
        QSpinBox, QDoubleSpinBox, QListWidget, QTreeWidget {{
            min-height: 22px; padding: 3px 6px;
            border: 1px solid {p['border_strong']};
            border-radius: {p['radius']};
            background-color: {p['input']}; color: {p['text']};
            selection-background-color: {p['selection']};
            selection-color: {p['accent_text']};
        }}
        QTextEdit, QPlainTextEdit {{ padding: 4px 6px; }}
        /* Keep field surfaces opaque while the pointer is over them. */
        QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover,
        QComboBox:hover, QDateEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover,
        QListWidget:hover, QTreeWidget:hover {{
            background-color: {p['input']};
            border-color: {p['accent']};
        }}
        QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus,
        QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
        QListWidget:focus, QTreeWidget:focus {{
            border: 1px solid {p['accent']};
        }}
        QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled,
        QComboBox:disabled, QDateEdit:disabled, QSpinBox:disabled,
        QDoubleSpinBox:disabled {{
            color: {p['disabled_text']};
            background-color: {p['disabled_bg']};
        }}

        /* Keep a dedicated, visible arrow/button zone on every dropdown. */
        QComboBox {{ padding-right: 32px; }}
        QComboBox::drop-down, QDateEdit::drop-down {{
            subcontrol-origin: padding; subcontrol-position: top right;
            width: 28px; border-left: 1px solid {p['border']};
            background-color: {p['surface']};
            border-top-right-radius: {p['radius']};
            border-bottom-right-radius: {p['radius']};
        }}
        QComboBox::drop-down:hover, QDateEdit::drop-down:hover {{
            background-color: {p['surface_hover']};
        }}
        QComboBox::down-arrow, QDateEdit::down-arrow {{
            width: 10px; height: 7px;
        }}
        QComboBox::down-arrow:on, QDateEdit::down-arrow:on {{ top: 1px; }}

        /* Read-only calculated dates do not need a right-side control. */
        QDateEdit[hideDateButton="true"] {{ padding-right: 6px; }}
        QDateEdit[hideDateButton="true"]::drop-down {{
            width: 0; border: 0; background: transparent;
        }}
        QDateEdit[hideDateButton="true"]::down-arrow {{
            width: 0; height: 0;
        }}
        QComboBox QAbstractItemView, QAbstractItemView {{
            background-color: {p['panel']}; color: {p['text']};
            border: 1px solid {p['border_strong']};
            selection-background-color: {p['selection']};
            selection-color: {p['accent_text']};
            outline: 0; padding: 3px;
        }}
        QComboBox QAbstractItemView::item {{
            min-height: 26px; padding: 4px 8px; border: 0;
            background-color: {p['panel']}; color: {p['text']};
        }}
        QAbstractItemView::item:hover {{
            background-color: {p['surface_hover']}; color: {p['text']};
        }}
        QAbstractItemView::item:selected,
        QAbstractItemView::item:selected:hover {{
            background-color: {p['selection']}; color: {p['accent_text']};
        }}

        /* Numeric fields use Qt's PlusMinus button symbols; these rules make
           the two button areas clear in every theme. */
        QSpinBox, QDoubleSpinBox,
        QDateEdit[contrastSpinButtons="true"] {{ padding-right: 30px; }}
        QSpinBox::up-button, QDoubleSpinBox::up-button,
        QDateEdit[contrastSpinButtons="true"]::up-button {{
            subcontrol-origin: border; subcontrol-position: top right;
            width: 24px; border-left: 1px solid {p['border']};
            border-bottom: 1px solid {p['border']};
            background-color: {p['surface']};
            border-top-right-radius: {p['radius']};
        }}
        QSpinBox::down-button, QDoubleSpinBox::down-button,
        QDateEdit[contrastSpinButtons="true"]::down-button {{
            subcontrol-origin: border; subcontrol-position: bottom right;
            width: 24px; border-left: 1px solid {p['border']};
            background-color: {p['surface']};
            border-bottom-right-radius: {p['radius']};
        }}
        QSpinBox::up-button:hover, QSpinBox::down-button:hover,
        QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover,
        QDateEdit[contrastSpinButtons="true"]::up-button:hover,
        QDateEdit[contrastSpinButtons="true"]::down-button:hover {{
            background-color: {p['surface_hover']};
        }}
        QSpinBox::up-button:pressed, QSpinBox::down-button:pressed,
        QDoubleSpinBox::up-button:pressed, QDoubleSpinBox::down-button:pressed,
        QDateEdit[contrastSpinButtons="true"]::up-button:pressed,
        QDateEdit[contrastSpinButtons="true"]::down-button:pressed {{
            background-color: {p['surface_pressed']};
        }}
        QSpinBox::up-arrow, QSpinBox::down-arrow,
        QDoubleSpinBox::up-arrow, QDoubleSpinBox::down-arrow,
        QDateEdit[contrastSpinButtons="true"]::up-arrow,
        QDateEdit[contrastSpinButtons="true"]::down-arrow {{
            width: 12px; height: 12px;
        }}
        QSpinBox::up-button:disabled, QSpinBox::down-button:disabled,
        QDoubleSpinBox::up-button:disabled, QDoubleSpinBox::down-button:disabled,
        QDateEdit[contrastSpinButtons="true"]::up-button:disabled,
        QDateEdit[contrastSpinButtons="true"]::down-button:disabled {{
            background-color: {p['disabled_bg']};
            border-color: {p['border']};
        }}

        /* Numeric inputs marked with hideSpinButtons remain keyboard-editable
           but do not display the increment/decrement button column. */
        QSpinBox[hideSpinButtons="true"],
        QDoubleSpinBox[hideSpinButtons="true"] {{
            padding-right: 6px;
        }}
        QSpinBox[hideSpinButtons="true"]::up-button,
        QSpinBox[hideSpinButtons="true"]::down-button,
        QDoubleSpinBox[hideSpinButtons="true"]::up-button,
        QDoubleSpinBox[hideSpinButtons="true"]::down-button {{
            width: 0; height: 0;
            border: 0; background: transparent;
        }}
        QSpinBox[hideSpinButtons="true"]::up-arrow,
        QSpinBox[hideSpinButtons="true"]::down-arrow,
        QDoubleSpinBox[hideSpinButtons="true"]::up-arrow,
        QDoubleSpinBox[hideSpinButtons="true"]::down-arrow {{
            width: 0; height: 0;
        }}

        QPushButton, QToolButton {{
            min-height: 25px; padding: 3px 9px;
            border: 1px solid {p['border_strong']};
            border-radius: {p['radius']};
            background-color: {p['surface']}; color: {p['text']};
            font-weight: 600;
        }}
        QPushButton:hover, QToolButton:hover {{
            background-color: {p['surface_hover']}; border-color: {p['accent']};
        }}
        QPushButton:pressed, QToolButton:pressed {{
            background-color: {p['surface_pressed']};
        }}
        QPushButton:disabled, QToolButton:disabled {{
            background-color: {p['disabled_bg']}; color: {p['disabled_text']};
            border-color: {p['border']};
        }}
        QPushButton#primaryButton {{
            background-color: {p['accent']}; border-color: {p['accent_hover']};
            color: {p['accent_text']};
        }}
        QPushButton#primaryButton:hover {{
            background-color: {p['accent_hover']}; color: {p['accent_text']};
        }}
        QPushButton#primaryButton:pressed {{
            background-color: {p['accent_pressed']}; color: {p['accent_text']};
        }}
        QPushButton#successButton {{ background-color: {p['success']}; color: #ffffff; }}
        QPushButton#warningButton {{ background-color: {p['warning']}; color: #ffffff; }}
        QPushButton#dangerButton {{ background-color: {p['danger']}; color: #ffffff; }}
        QPushButton#neutralButton {{ background-color: {p['neutral']}; color: #ffffff; }}
        /* Preserve semantic button colors on hover instead of falling back to
           a platform-generated translucent highlight. */
        QPushButton#successButton:hover {{
            background-color: {p['success']}; color: #ffffff;
            border-color: {p['accent_text']};
        }}
        QPushButton#warningButton:hover {{
            background-color: {p['warning']}; color: #ffffff;
            border-color: {p['accent_text']};
        }}
        QPushButton#dangerButton:hover {{
            background-color: {p['danger']}; color: #ffffff;
            border-color: {p['accent_text']};
        }}
        QPushButton#neutralButton:hover {{
            background-color: {p['neutral']}; color: #ffffff;
            border-color: {p['accent_text']};
        }}
        QPushButton#secondaryButton {{
            background-color: {p['surface']}; border-color: {p['border_strong']};
        }}

        QTableWidget, QTableView {{
            background-color: {p['input']};
            alternate-background-color: {p['panel_alt']};
            gridline-color: {p['border']}; border: 1px solid {p['border']};
            selection-background-color: {p['selection']};
            selection-color: {p['accent_text']};
        }}
        QTableWidget::item, QTableView::item {{ padding: 4px; }}
        QHeaderView {{ background-color: {p['header']}; }}
        QHeaderView::section, QTableCornerButton::section {{
            background-color: {p['header']}; color: {p['text']};
            padding: 5px; border: 0;
            border-right: 1px solid {p['border']};
            border-bottom: 1px solid {p['border']}; font-weight: 600;
        }}
        QHeaderView::section:hover, QTableCornerButton::section:hover {{
            background-color: {p['surface_hover']}; color: {p['text']};
        }}
        QHeaderView::section:pressed {{
            background-color: {p['surface_pressed']}; color: {p['text']};
        }}

        QScrollBar:vertical {{
            background-color: {p['scroll_bg']}; width: 12px; margin: 0;
        }}
        QScrollBar::handle:vertical {{
            background-color: {p['scroll_handle']}; min-height: 28px;
            border-radius: 5px;
        }}
        QScrollBar:horizontal {{
            background-color: {p['scroll_bg']}; height: 12px; margin: 0;
        }}
        QScrollBar::handle:horizontal {{
            background-color: {p['scroll_handle']}; min-width: 28px;
            border-radius: 5px;
        }}
        QScrollBar::handle:hover {{ background-color: {p['accent']}; }}
        QScrollBar::handle:pressed {{ background-color: {p['accent_pressed']}; }}
        QScrollBar::add-line, QScrollBar::sub-line {{
            width: 0; height: 0; background-color: {p['scroll_bg']};
        }}
        QScrollBar::add-page, QScrollBar::sub-page {{
            background-color: {p['scroll_bg']};
        }}

        QCalendarWidget {{ background-color: {p['panel']}; color: {p['text']}; }}
        QCalendarWidget QWidget {{ background-color: {p['panel']}; color: {p['text']}; }}
        QCalendarWidget QAbstractItemView {{
            background-color: {p['input']}; color: {p['text']};
            selection-background-color: {p['selection']};
            selection-color: {p['accent_text']};
        }}

        QStatusBar {{
            background-color: {p['menu']}; color: {p['muted']};
            border-top: 1px solid {p['border']};
        }}
        QStatusBar::item {{ border: 0; }}
        QToolTip {{
            background-color: {p['tooltip']}; color: {p['tooltip_text']};
            border: 1px solid {p['border_strong']}; padding: 4px;
        }}
        QProgressBar {{
            border: 1px solid {p['border_strong']};
            border-radius: {p['radius']};
            background-color: {p['input']}; color: {p['text']};
            text-align: center;
        }}
        QProgressBar::chunk {{ background-color: {p['accent']}; }}
    """


class ContrastSpinBoxStyle(QProxyStyle):
    """Draw extra-bold, high-contrast plus/minus glyphs for numeric spin boxes."""

    MIN_GLYPH_PEN_WIDTH = 3.0
    MAX_GLYPH_PEN_WIDTH = 4.5
    GLYPH_PEN_RATIO = 0.22
    GLYPH_ARM_RATIO = 0.31

    def polish(self, target):
        """Use PlusMinus symbols for numeric fields and opted-in date filters.

        Date editors keep their normal calendar behavior unless the widget has
        the ``contrastSpinButtons`` property enabled. This limits the bold
        plus/minus controls to the requested FROM/TO filter fields.
        """
        result = super().polish(target)
        contrast_date_edit = (
            isinstance(target, QDateEdit)
            and bool(target.property("contrastSpinButtons"))
        )
        if isinstance(target, (QSpinBox, QDoubleSpinBox)):
            if bool(target.property("hideSpinButtons")):
                target.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
            else:
                target.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.PlusMinus)
        elif contrast_date_edit:
            target.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.PlusMinus)
        return result

    @staticmethod
    def _relative_luminance(color):
        channels = []
        for value in (color.redF(), color.greenF(), color.blueF()):
            channels.append(
                value / 12.92
                if value <= 0.04045
                else ((value + 0.055) / 1.055) ** 2.4
            )
        return (
            (0.2126 * channels[0])
            + (0.7152 * channels[1])
            + (0.0722 * channels[2])
        )

    @classmethod
    def _contrast_glyph_color(cls, background):
        """Return black or white, whichever has the stronger WCAG contrast."""
        luminance = cls._relative_luminance(background)
        contrast_with_white = 1.05 / (luminance + 0.05)
        contrast_with_black = (luminance + 0.05) / 0.05
        return (
            QColor("#ffffff")
            if contrast_with_white >= contrast_with_black
            else QColor("#000000")
        )

    def drawPrimitive(self, element, option, painter, widget=None):
        plus_element = QStyle.PrimitiveElement.PE_IndicatorSpinPlus
        minus_element = QStyle.PrimitiveElement.PE_IndicatorSpinMinus

        if element not in (plus_element, minus_element):
            return super().drawPrimitive(element, option, painter, widget)

        app = QApplication.instance()
        theme_code = (
            app.property("themeCode")
            if app is not None
            else DEFAULT_THEME_CODE
        )
        theme = get_theme_config(theme_code)
        state = option.state

        if not (state & QStyle.StateFlag.State_Enabled):
            glyph_color = QColor(theme["disabled_text"])
        else:
            if state & QStyle.StateFlag.State_Sunken:
                button_background = QColor(theme["surface_pressed"])
            elif state & QStyle.StateFlag.State_MouseOver:
                button_background = QColor(theme["surface_hover"])
            else:
                button_background = QColor(theme["surface"])

            glyph_color = self._contrast_glyph_color(button_background)

        # Leave a small safe margin so the thick round-ended glyph never
        # touches the spin-button border, including at high-DPI scaling.
        rect = option.rect.adjusted(3, 3, -3, -3)
        minimum_side = min(rect.width(), rect.height())
        if minimum_side <= 4:
            return

        center = rect.center()
        arm = max(4, round(minimum_side * self.GLYPH_ARM_RATIO))
        pen_width = max(
            self.MIN_GLYPH_PEN_WIDTH,
            min(
                self.MAX_GLYPH_PEN_WIDTH,
                minimum_side * self.GLYPH_PEN_RATIO,
            ),
        )

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        pen = QPen(glyph_color)
        pen.setWidthF(pen_width)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)

        painter.drawLine(
            center.x() - arm,
            center.y(),
            center.x() + arm,
            center.y(),
        )
        if element == plus_element:
            painter.drawLine(
                center.x(),
                center.y() - arm,
                center.x(),
                center.y() + arm,
            )

        painter.restore()

def apply_application_theme(app, theme_code=None):
    """Apply a complete, cross-platform application theme and return its code."""
    code = str(theme_code or DEFAULT_THEME_CODE).strip().lower()
    if code not in THEME_OPTIONS:
        code = DEFAULT_THEME_CODE
    p = THEME_OPTIONS[code]

    # Store the theme before installing the proxy style so glyph drawing
    # immediately uses the correct hover/pressed background colors.
    app.setProperty("themeCode", code)

    try:
        available = {name.lower(): name for name in QStyleFactory.keys()}
        preferred = "windows" if code == "classic" else "fusion"
        style_name = available.get(preferred) or available.get("fusion")
        if style_name:
            base_style = QStyleFactory.create(style_name)
            if base_style is not None:
                app.setStyle(ContrastSpinBoxStyle(base_style))
    except Exception:
        pass

    palette = QPalette()
    role_colors = {
        QPalette.ColorRole.Window: p["window"],
        QPalette.ColorRole.WindowText: p["text"],
        QPalette.ColorRole.Base: p["input"],
        QPalette.ColorRole.AlternateBase: p["panel_alt"],
        QPalette.ColorRole.ToolTipBase: p["tooltip"],
        QPalette.ColorRole.ToolTipText: p["tooltip_text"],
        QPalette.ColorRole.Text: p["text"],
        QPalette.ColorRole.Button: p["surface"],
        QPalette.ColorRole.ButtonText: p["text"],
        QPalette.ColorRole.BrightText: p["danger"],
        QPalette.ColorRole.Highlight: p["selection"],
        QPalette.ColorRole.HighlightedText: p["accent_text"],
        QPalette.ColorRole.Link: p["accent"],
        QPalette.ColorRole.PlaceholderText: p["muted"],
    }
    for role, color in role_colors.items():
        palette.setColor(role, QColor(color))
    app.setPalette(palette)
    app.setStyleSheet(build_theme_stylesheet(code))
    app.setProperty("themeCode", code)
    return code

def create_calendar_icon(size=18):
    """Create a crisp theme-aware calendar icon without external image files."""
    app = QApplication.instance()
    theme_code = app.property("themeCode") if app is not None else DEFAULT_THEME_CODE
    theme = get_theme_config(theme_code)

    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    margin = max(1, round(size * 0.12))
    rect = QRectF(margin, margin + 2, size - (margin * 2), size - margin - 3)
    pen = QPen(QColor(theme["text"]))
    pen.setWidthF(max(1.4, size * 0.09))
    painter.setPen(pen)
    painter.setBrush(Qt.BrushStyle.NoBrush)
    painter.drawRoundedRect(rect, 2, 2)

    header_y = rect.top() + max(4, size * 0.28)
    painter.drawLine(QPointF(rect.left(), header_y), QPointF(rect.right(), header_y))
    for x in (rect.left() + rect.width() * 0.3, rect.left() + rect.width() * 0.7):
        painter.drawLine(QPointF(x, rect.top() - 2), QPointF(x, rect.top() + 3))

    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor(theme["accent"]))
    dot = max(1.5, size * 0.09)
    for row in range(2):
        for col in range(3):
            cx = rect.left() + rect.width() * (0.24 + col * 0.26)
            cy = header_y + rect.height() * (0.24 + row * 0.26)
            painter.drawEllipse(QPointF(cx, cy), dot, dot)

    painter.end()
    return QIcon(pixmap)


def get_runtime_app_dir():
    """Return the folder that should contain external app files."""
    if getattr(sys, "frozen", False) or "__compiled__" in globals():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def get_app_file_path(filename):
    return os.path.join(get_runtime_app_dir(), filename)

LANGUAGE_OPTIONS = {
    "en_US": {
        "label": "English (United States)",
        "native": "English",
        "currency": "USD",
        "thousands_sep": ",",
        "decimal_sep": ".",
        "grouping": "western",
        "date_format": "MM-dd-yyyy",
        "python_date_format": "%m-%d-%Y",
        "rtl": False,
    },
    "id_ID": {
        "label": "Indonesian (Indonesia)",
        "native": "Bahasa Indonesia",
        "currency": "IDR",
        "thousands_sep": ".",
        "decimal_sep": ",",
        "grouping": "western",
        "date_format": "dd-MM-yyyy",
        "python_date_format": "%d-%m-%Y",
        "rtl": False,
    },
}

UI_TRANSLATIONS = {
    "id_ID": {
        "FILE": "BERKAS",
        "IMPORT": "IMPOR",
        "EXPORT": "EKSPOR",
        "OR": "ATAU",
        "TO": "KE",
        "EXCEL": "EXCEL",
        "CSV": "CSV",
        "BACKUP": "CADANGKAN",
        "RESTORE": "PULIHKAN",
        "EXIT": "KELUAR",
        "SETTINGS": "PENGATURAN",
        "DATABASE": "DATABASE",
        "PATH": "LOKASI",
        "CURRENCY": "MATA UANG",
        "LANGUAGE": "BAHASA",
        "LOCALE": "LOKAL",
        "HELP": "BANTUAN",
        "ABOUT": "TENTANG",
        "QUICK": "CEPAT",
        "NEW": "BARU",
        "SAVE": "SIMPAN",
        "SEARCH": "CARI",
        "READY": "SIAP",
        "USER DATA": "DATA PENGGUNA",
        "SUPPLIERS": "PEMASOK",
        "REPORTS": "LAPORAN",
        "WARRANTY TRACKING": "PELACAKAN GARANSI",
        "PARTS": "SUKU CADANG",
        "CUSTOMER INFORMATION": "INFORMASI PELANGGAN",
        "DEVICE INFORMATION": "INFORMASI PERANGKAT",
        "SERVICE INFORMATION": "INFORMASI SERVIS",
        "DATE & WARRANTY": "TANGGAL & GARANSI",
        "PARTS & SUPPLIER": "SUKU CADANG & PEMASOK",
        "ADD SUPPLIER": "TAMBAH PEMASOK",
        "REPORT OPTIONS": "OPSI LAPORAN",
        "WARRANTY FILTER": "FILTER GARANSI",
        "PART INFORMATION": "INFORMASI SUKU CADANG",
        "FILTER PARTS": "FILTER SUKU CADANG",
        "NAME": "NAMA",
        "PHONE": "TELEPON",
        "ADDRESS": "ALAMAT",
        "DEVICES": "PERANGKAT",
        "DEVICE": "PERANGKAT",
        "BRAND": "MEREK",
        "MODEL": "MODEL",
        "ISSUE": "MASALAH",
        "DESCRIPTION": "DESKRIPSI",
        "STATUS": "STATUS",
        "TECHNICIAN": "TEKNISI",
        "PRICE": "HARGA",
        "SERVICE": "SERVIS",
        "DATE": "TANGGAL",
        "SERVICE DATE": "TANGGAL SERVIS",
        "WARRANTY": "GARANSI",
        "WARRANTY END": "AKHIR GARANSI",
        "GOOGLE MAP": "GOOGLE MAP",
        "PART": "SUKU CADANG",
        "SUPPLIER": "PEMASOK",
        "PURCHASE INVOICE": "INVOICE PEMBELIAN",
        "CONTACT": "KONTAK",
        "EMAIL": "EMAIL",
        "FROM": "DARI",
        "REPORT TYPE": "JENIS LAPORAN",
        "TYPE": "JENIS",
        "QUANTITY": "JUMLAH",
        "UNIT PRICE": "HARGA SATUAN",
        "MIN STOCK": "STOK MINIMUM",
        "LOCATION": "LOKASI",
        "NOTES": "CATATAN",
        "CUSTOMER": "PELANGGAN",
        "WARRANTY PERIOD": "PERIODE GARANSI",
        "END DATE": "TANGGAL AKHIR",
        "DAYS LEFT": "SISA HARI",
        "CLEAR": "BERSIHKAN",
        "EDIT": "EDIT",
        "DELETE": "HAPUS",
        "REFRESH": "SEGARKAN",
        "EXPORT SELECTED": "EKSPOR TERPILIH",
        "ADD": "TAMBAH",
        "ADD PART": "TAMBAH SUKU CADANG",
        "GENERATE REPORT": "BUAT LAPORAN",
        "SEND REMINDER": "KIRIM PENGINGAT",
        "EDIT CONTACT": "EDIT KONTAK",
        "TABLE VIEW": "TAMPILAN TABEL",
        "TEXT VIEW": "TAMPILAN TEKS",
        "CLOSE": "TUTUP",
        "BROWSE": "TELUSURI",
        "CANCEL": "BATAL",
        "USE DEFAULT": "GUNAKAN DEFAULT",
        "AUTO BACKUP ON EXIT": "CADANGKAN OTOMATIS SAAT KELUAR",
        "NAME CUSTOMER": "NAMA PELANGGAN",
        "NOMOR PHONE": "NOMOR TELEPON",
        "ADDRESS LENGKAP": "ALAMAT LENGKAP",
        "MODEL DEVICE": "MODEL PERANGKAT",
        "SERVICE DESCRIPTION": "DESKRIPSI SERVIS",
        "TECHNICIAN NAME": "NAMA TEKNISI",
        "LINK GOOGLE MAPS": "LINK GOOGLE MAPS",
        "NOMOR INVOICE": "NOMOR INVOICE",
        "SEARCH BY NAME, PHONE, DEVICE, OR ISSUE...": "CARI BERDASARKAN NAMA, TELEPON, PERANGKAT, ATAU MASALAH...",
        "SUPPLIER NAME": "NAMA PEMASOK",
        "CONTACT NAME": "NAMA KONTAK",
        "SUPPLIER ADDRESS": "ALAMAT PEMASOK",
        "PART NAME": "NAMA SUKU CADANG",
        "STORAGE LOCATION": "LOKASI PENYIMPANAN",
        "ADDITIONAL NOTES": "CATATAN TAMBAHAN",
        "LANGUAGE / LOCALE": "BAHASA / LOKAL",
        "SELECT LANGUAGE / LOCALE": "PILIH BAHASA / LOKAL",
        "SELECT APPLICATION LANGUAGE AND LOCALE:": "PILIH BAHASA DAN LOKAL APLIKASI:",
    },
    "es_MX": {
        "FILE": "ARCHIVO", "IMPORT": "IMPORTAR", "EXPORT": "EXPORTAR", "OR": "O", "TO": "A",
        "BACKUP": "COPIA", "RESTORE": "RESTAURAR", "EXIT": "SALIR", "SETTINGS": "AJUSTES",
        "DATABASE": "BASE DE DATOS", "PATH": "RUTA", "CURRENCY": "MONEDA", "LANGUAGE": "IDIOMA",
        "LOCALE": "REGIÓN", "HELP": "AYUDA", "ABOUT": "ACERCA DE", "QUICK": "RÁPIDO",
        "NEW": "NUEVO", "SAVE": "GUARDAR", "SEARCH": "BUSCAR", "READY": "LISTO",
        "USER DATA": "DATOS DE USUARIO", "SUPPLIERS": "PROVEEDORES", "REPORTS": "INFORMES",
        "WARRANTY TRACKING": "SEGUIMIENTO DE GARANTÍA", "PARTS": "PIEZAS",
        "CUSTOMER INFORMATION": "INFORMACIÓN DEL CLIENTE", "DEVICE INFORMATION": "INFORMACIÓN DEL EQUIPO",
        "SERVICE INFORMATION": "INFORMACIÓN DEL SERVICIO", "DATE & WARRANTY": "FECHA Y GARANTÍA",
        "PARTS & SUPPLIER": "PIEZAS Y PROVEEDOR", "ADD SUPPLIER": "AGREGAR PROVEEDOR",
        "REPORT OPTIONS": "OPCIONES DE INFORME", "WARRANTY FILTER": "FILTRO DE GARANTÍA",
        "PART INFORMATION": "INFORMACIÓN DE PIEZA", "FILTER PARTS": "FILTRAR PIEZAS",
        "NAME": "NOMBRE", "PHONE": "TELÉFONO", "ADDRESS": "DIRECCIÓN", "DEVICES": "EQUIPOS",
        "DEVICE": "EQUIPO", "BRAND": "MARCA", "MODEL": "MODELO", "ISSUE": "PROBLEMA",
        "DESCRIPTION": "DESCRIPCIÓN", "STATUS": "ESTADO", "TECHNICIAN": "TÉCNICO",
        "PRICE": "PRECIO", "SERVICE DATE": "FECHA DE SERVICIO", "WARRANTY": "GARANTÍA",
        "WARRANTY END": "FIN DE GARANTÍA", "PART": "PIEZA", "SUPPLIER": "PROVEEDOR",
        "PURCHASE INVOICE": "FACTURA DE COMPRA", "CONTACT": "CONTACTO", "EMAIL": "EMAIL",
        "FROM": "DESDE", "REPORT TYPE": "TIPO DE INFORME", "TYPE": "TIPO",
        "QUANTITY": "CANTIDAD", "UNIT PRICE": "PRECIO UNITARIO", "MIN STOCK": "STOCK MÍNIMO",
        "LOCATION": "UBICACIÓN", "NOTES": "NOTAS", "CUSTOMER": "CLIENTE",
        "WARRANTY PERIOD": "PERIODO DE GARANTÍA", "END DATE": "FECHA FINAL", "DAYS LEFT": "DÍAS RESTANTES",
        "CLEAR": "LIMPIAR", "EDIT": "EDITAR", "DELETE": "ELIMINAR", "REFRESH": "ACTUALIZAR",
        "EXPORT SELECTED": "EXPORTAR SELECCIÓN", "ADD": "AGREGAR", "ADD PART": "AGREGAR PIEZA",
        "GENERATE REPORT": "GENERAR INFORME", "SEND REMINDER": "ENVIAR RECORDATORIO",
        "EDIT CONTACT": "EDITAR CONTACTO", "TABLE VIEW": "VISTA DE TABLA", "TEXT VIEW": "VISTA DE TEXTO",
        "CLOSE": "CERRAR", "BROWSE": "EXAMINAR", "CANCEL": "CANCELAR", "USE DEFAULT": "USAR PREDETERMINADO",
        "AUTO BACKUP ON EXIT": "COPIA AUTOMÁTICA AL SALIR", "LANGUAGE / LOCALE": "IDIOMA / REGIÓN",
    },
    "fr_FR": {
        "FILE": "FICHIER", "IMPORT": "IMPORTER", "EXPORT": "EXPORTER", "OR": "OU", "TO": "VERS",
        "BACKUP": "SAUVEGARDE", "RESTORE": "RESTAURER", "EXIT": "QUITTER", "SETTINGS": "PARAMÈTRES",
        "DATABASE": "BASE DE DONNÉES", "PATH": "CHEMIN", "CURRENCY": "DEVISE", "LANGUAGE": "LANGUE",
        "LOCALE": "RÉGION", "HELP": "AIDE", "ABOUT": "À PROPOS", "QUICK": "RAPIDE",
        "NEW": "NOUVEAU", "SAVE": "ENREGISTRER", "SEARCH": "RECHERCHER", "READY": "PRÊT",
        "USER DATA": "DONNÉES UTILISATEUR", "SUPPLIERS": "FOURNISSEURS", "REPORTS": "RAPPORTS",
        "WARRANTY TRACKING": "SUIVI DE GARANTIE", "PARTS": "PIÈCES",
        "CUSTOMER INFORMATION": "INFORMATIONS CLIENT", "DEVICE INFORMATION": "INFORMATIONS APPAREIL",
        "SERVICE INFORMATION": "INFORMATIONS SERVICE", "DATE & WARRANTY": "DATE ET GARANTIE",
        "PARTS & SUPPLIER": "PIÈCES ET FOURNISSEUR", "ADD SUPPLIER": "AJOUTER FOURNISSEUR",
        "REPORT OPTIONS": "OPTIONS DE RAPPORT", "WARRANTY FILTER": "FILTRE DE GARANTIE",
        "PART INFORMATION": "INFORMATIONS PIÈCE", "FILTER PARTS": "FILTRER LES PIÈCES",
        "NAME": "NOM", "PHONE": "TÉLÉPHONE", "ADDRESS": "ADRESSE", "DEVICES": "APPAREILS",
        "DEVICE": "APPAREIL", "BRAND": "MARQUE", "MODEL": "MODÈLE", "ISSUE": "PROBLÈME",
        "DESCRIPTION": "DESCRIPTION", "STATUS": "STATUT", "TECHNICIAN": "TECHNICIEN",
        "PRICE": "PRIX", "SERVICE DATE": "DATE DE SERVICE", "WARRANTY": "GARANTIE",
        "WARRANTY END": "FIN DE GARANTIE", "PART": "PIÈCE", "SUPPLIER": "FOURNISSEUR",
        "PURCHASE INVOICE": "FACTURE D'ACHAT", "CONTACT": "CONTACT", "EMAIL": "EMAIL",
        "FROM": "DU", "REPORT TYPE": "TYPE DE RAPPORT", "TYPE": "TYPE", "QUANTITY": "QUANTITÉ",
        "UNIT PRICE": "PRIX UNITAIRE", "MIN STOCK": "STOCK MINIMUM", "LOCATION": "EMPLACEMENT",
        "NOTES": "NOTES", "CUSTOMER": "CLIENT", "WARRANTY PERIOD": "PÉRIODE DE GARANTIE",
        "END DATE": "DATE DE FIN", "DAYS LEFT": "JOURS RESTANTS", "CLEAR": "EFFACER",
        "EDIT": "MODIFIER", "DELETE": "SUPPRIMER", "REFRESH": "ACTUALISER",
        "EXPORT SELECTED": "EXPORTER LA SÉLECTION", "ADD": "AJOUTER", "ADD PART": "AJOUTER PIÈCE",
        "GENERATE REPORT": "GÉNÉRER RAPPORT", "SEND REMINDER": "ENVOYER RAPPEL",
        "EDIT CONTACT": "MODIFIER CONTACT", "TABLE VIEW": "VUE TABLEAU", "TEXT VIEW": "VUE TEXTE",
        "CLOSE": "FERMER", "BROWSE": "PARCOURIR", "CANCEL": "ANNULER", "USE DEFAULT": "UTILISER PAR DÉFAUT",
        "AUTO BACKUP ON EXIT": "SAUVEGARDE AUTO À LA SORTIE", "LANGUAGE / LOCALE": "LANGUE / RÉGION",
    },
    "de_DE": {
        "FILE": "DATEI", "IMPORT": "IMPORTIEREN", "EXPORT": "EXPORTIEREN", "OR": "ODER", "TO": "NACH",
        "BACKUP": "SICHERUNG", "RESTORE": "WIEDERHERSTELLEN", "EXIT": "BEENDEN", "SETTINGS": "EINSTELLUNGEN",
        "DATABASE": "DATENBANK", "PATH": "PFAD", "CURRENCY": "WÄHRUNG", "LANGUAGE": "SPRACHE",
        "LOCALE": "REGION", "HELP": "HILFE", "ABOUT": "INFO", "NEW": "NEU", "SAVE": "SPEICHERN",
        "SEARCH": "SUCHE", "READY": "BEREIT", "USER DATA": "BENUTZERDATEN", "SUPPLIERS": "LIEFERANTEN",
        "REPORTS": "BERICHTE", "WARRANTY TRACKING": "GARANTIEVERFOLGUNG", "PARTS": "TEILE",
        "CUSTOMER INFORMATION": "KUNDENINFORMATIONEN", "DEVICE INFORMATION": "GERÄTEINFORMATIONEN",
        "SERVICE INFORMATION": "SERVICEINFORMATIONEN", "DATE & WARRANTY": "DATUM & GARANTIE",
        "PARTS & SUPPLIER": "TEILE & LIEFERANT", "ADD SUPPLIER": "LIEFERANT HINZUFÜGEN",
        "REPORT OPTIONS": "BERICHTSOPTIONEN", "WARRANTY FILTER": "GARANTIEFILTER",
        "PART INFORMATION": "TEILEINFORMATIONEN", "FILTER PARTS": "TEILE FILTERN",
        "NAME": "NAME", "PHONE": "TELEFON", "ADDRESS": "ADRESSE", "DEVICES": "GERÄTE",
        "DEVICE": "GERÄT", "BRAND": "MARKE", "MODEL": "MODELL", "ISSUE": "PROBLEM",
        "DESCRIPTION": "BESCHREIBUNG", "STATUS": "STATUS", "TECHNICIAN": "TECHNIKER",
        "PRICE": "PREIS", "SERVICE DATE": "SERVICEDATUM", "WARRANTY": "GARANTIE",
        "WARRANTY END": "GARANTIEENDE", "PART": "TEIL", "SUPPLIER": "LIEFERANT",
        "PURCHASE INVOICE": "EINKAUFSRECHNUNG", "CONTACT": "KONTAKT", "EMAIL": "EMAIL",
        "FROM": "VON", "REPORT TYPE": "BERICHTSTYP", "TYPE": "TYP", "QUANTITY": "MENGE",
        "UNIT PRICE": "EINZELPREIS", "MIN STOCK": "MINDESTBESTAND", "LOCATION": "STANDORT",
        "NOTES": "NOTIZEN", "CUSTOMER": "KUNDE", "WARRANTY PERIOD": "GARANTIEZEITRAUM",
        "END DATE": "ENDDATUM", "DAYS LEFT": "TAGE ÜBRIG", "CLEAR": "LEEREN", "EDIT": "BEARBEITEN",
        "DELETE": "LÖSCHEN", "REFRESH": "AKTUALISIEREN", "EXPORT SELECTED": "AUSWAHL EXPORTIEREN",
        "ADD": "HINZUFÜGEN", "ADD PART": "TEIL HINZUFÜGEN", "GENERATE REPORT": "BERICHT ERSTELLEN",
        "SEND REMINDER": "ERINNERUNG SENDEN", "EDIT CONTACT": "KONTAKT BEARBEITEN",
        "TABLE VIEW": "TABELLENANSICHT", "TEXT VIEW": "TEXTANSICHT", "CLOSE": "SCHLIESSEN",
        "BROWSE": "DURCHSUCHEN", "CANCEL": "ABBRECHEN", "USE DEFAULT": "STANDARD VERWENDEN",
        "AUTO BACKUP ON EXIT": "AUTOMATISCHE SICHERUNG BEIM BEENDEN", "LANGUAGE / LOCALE": "SPRACHE / REGION",
    },
    "pt_BR": {
        "FILE": "ARQUIVO", "IMPORT": "IMPORTAR", "EXPORT": "EXPORTAR", "OR": "OU", "TO": "PARA",
        "BACKUP": "BACKUP", "RESTORE": "RESTAURAR", "EXIT": "SAIR", "SETTINGS": "CONFIGURAÇÕES",
        "DATABASE": "BANCO DE DADOS", "PATH": "CAMINHO", "CURRENCY": "MOEDA", "LANGUAGE": "IDIOMA",
        "LOCALE": "LOCALIDADE", "HELP": "AJUDA", "ABOUT": "SOBRE", "NEW": "NOVO", "SAVE": "SALVAR",
        "SEARCH": "PESQUISAR", "READY": "PRONTO", "USER DATA": "DADOS DO USUÁRIO", "SUPPLIERS": "FORNECEDORES",
        "REPORTS": "RELATÓRIOS", "WARRANTY TRACKING": "RASTREAMENTO DE GARANTIA", "PARTS": "PEÇAS",
        "CUSTOMER INFORMATION": "INFORMAÇÕES DO CLIENTE", "DEVICE INFORMATION": "INFORMAÇÕES DO DISPOSITIVO",
        "SERVICE INFORMATION": "INFORMAÇÕES DO SERVIÇO", "DATE & WARRANTY": "DATA E GARANTIA",
        "PARTS & SUPPLIER": "PEÇAS E FORNECEDOR", "ADD SUPPLIER": "ADICIONAR FORNECEDOR",
        "REPORT OPTIONS": "OPÇÕES DE RELATÓRIO", "WARRANTY FILTER": "FILTRO DE GARANTIA",
        "PART INFORMATION": "INFORMAÇÕES DA PEÇA", "FILTER PARTS": "FILTRAR PEÇAS",
        "NAME": "NOME", "PHONE": "TELEFONE", "ADDRESS": "ENDEREÇO", "DEVICES": "DISPOSITIVOS",
        "DEVICE": "DISPOSITIVO", "BRAND": "MARCA", "MODEL": "MODELO", "ISSUE": "PROBLEMA",
        "DESCRIPTION": "DESCRIÇÃO", "STATUS": "STATUS", "TECHNICIAN": "TÉCNICO", "PRICE": "PREÇO",
        "SERVICE DATE": "DATA DO SERVIÇO", "WARRANTY": "GARANTIA", "WARRANTY END": "FIM DA GARANTIA",
        "PART": "PEÇA", "SUPPLIER": "FORNECEDOR", "PURCHASE INVOICE": "NOTA DE COMPRA",
        "CONTACT": "CONTATO", "EMAIL": "EMAIL", "FROM": "DE", "REPORT TYPE": "TIPO DE RELATÓRIO",
        "TYPE": "TIPO", "QUANTITY": "QUANTIDADE", "UNIT PRICE": "PREÇO UNITÁRIO",
        "MIN STOCK": "ESTOQUE MÍNIMO", "LOCATION": "LOCALIZAÇÃO", "NOTES": "NOTAS",
        "CUSTOMER": "CLIENTE", "WARRANTY PERIOD": "PERÍODO DE GARANTIA", "END DATE": "DATA FINAL",
        "DAYS LEFT": "DIAS RESTANTES", "CLEAR": "LIMPAR", "EDIT": "EDITAR", "DELETE": "EXCLUIR",
        "REFRESH": "ATUALIZAR", "EXPORT SELECTED": "EXPORTAR SELECIONADOS", "ADD": "ADICIONAR",
        "ADD PART": "ADICIONAR PEÇA", "GENERATE REPORT": "GERAR RELATÓRIO", "SEND REMINDER": "ENVIAR LEMBRETE",
        "EDIT CONTACT": "EDITAR CONTATO", "TABLE VIEW": "VISÃO DE TABELA", "TEXT VIEW": "VISÃO DE TEXTO",
        "CLOSE": "FECHAR", "BROWSE": "PROCURAR", "CANCEL": "CANCELAR", "USE DEFAULT": "USAR PADRÃO",
        "AUTO BACKUP ON EXIT": "BACKUP AUTOMÁTICO AO SAIR", "LANGUAGE / LOCALE": "IDIOMA / LOCALIDADE",
    },
    "it_IT": {
        "FILE": "FILE", "IMPORT": "IMPORTA", "EXPORT": "ESPORTA", "OR": "O", "TO": "IN",
        "BACKUP": "BACKUP", "RESTORE": "RIPRISTINA", "EXIT": "ESCI", "SETTINGS": "IMPOSTAZIONI",
        "DATABASE": "DATABASE", "PATH": "PERCORSO", "CURRENCY": "VALUTA", "LANGUAGE": "LINGUA",
        "LOCALE": "AREA", "HELP": "AIUTO", "ABOUT": "INFORMAZIONI", "NEW": "NUOVO", "SAVE": "SALVA",
        "SEARCH": "CERCA", "READY": "PRONTO", "USER DATA": "DATI UTENTE", "SUPPLIERS": "FORNITORI",
        "REPORTS": "REPORT", "WARRANTY TRACKING": "TRACCIAMENTO GARANZIA", "PARTS": "RICAMBI",
        "CUSTOMER INFORMATION": "INFORMAZIONI CLIENTE", "DEVICE INFORMATION": "INFORMAZIONI DISPOSITIVO",
        "SERVICE INFORMATION": "INFORMAZIONI SERVIZIO", "DATE & WARRANTY": "DATA E GARANZIA",
        "PARTS & SUPPLIER": "RICAMBI E FORNITORE", "ADD SUPPLIER": "AGGIUNGI FORNITORE",
        "REPORT OPTIONS": "OPZIONI REPORT", "WARRANTY FILTER": "FILTRO GARANZIA",
        "PART INFORMATION": "INFORMAZIONI RICAMBIO", "FILTER PARTS": "FILTRA RICAMBI",
        "NAME": "NOME", "PHONE": "TELEFONO", "ADDRESS": "INDIRIZZO", "DEVICES": "DISPOSITIVI",
        "DEVICE": "DISPOSITIVO", "BRAND": "MARCA", "MODEL": "MODELLO", "ISSUE": "PROBLEMA",
        "DESCRIPTION": "DESCRIZIONE", "STATUS": "STATO", "TECHNICIAN": "TECNICO", "PRICE": "PREZZO",
        "SERVICE DATE": "DATA SERVIZIO", "WARRANTY": "GARANZIA", "WARRANTY END": "FINE GARANZIA",
        "PART": "RICAMBIO", "SUPPLIER": "FORNITORE", "PURCHASE INVOICE": "FATTURA ACQUISTO",
        "CONTACT": "CONTATTO", "EMAIL": "EMAIL", "FROM": "DA", "REPORT TYPE": "TIPO REPORT",
        "TYPE": "TIPO", "QUANTITY": "QUANTITÀ", "UNIT PRICE": "PREZZO UNITARIO",
        "MIN STOCK": "SCORTA MINIMA", "LOCATION": "POSIZIONE", "NOTES": "NOTE", "CUSTOMER": "CLIENTE",
        "WARRANTY PERIOD": "PERIODO GARANZIA", "END DATE": "DATA FINE", "DAYS LEFT": "GIORNI RIMASTI",
        "CLEAR": "PULISCI", "EDIT": "MODIFICA", "DELETE": "ELIMINA", "REFRESH": "AGGIORNA",
        "EXPORT SELECTED": "ESPORTA SELEZIONATI", "ADD": "AGGIUNGI", "ADD PART": "AGGIUNGI RICAMBIO",
        "GENERATE REPORT": "GENERA REPORT", "SEND REMINDER": "INVIA PROMEMORIA",
        "EDIT CONTACT": "MODIFICA CONTATTO", "TABLE VIEW": "VISTA TABELLA", "TEXT VIEW": "VISTA TESTO",
        "CLOSE": "CHIUDI", "BROWSE": "SFOGLIA", "CANCEL": "ANNULLA", "USE DEFAULT": "USA PREDEFINITO",
        "AUTO BACKUP ON EXIT": "BACKUP AUTOMATICO ALL'USCITA", "LANGUAGE / LOCALE": "LINGUA / AREA",
    },
    "nl_NL": {
        "FILE": "BESTAND", "IMPORT": "IMPORTEREN", "EXPORT": "EXPORTEREN", "OR": "OF", "TO": "NAAR",
        "BACKUP": "BACK-UP", "RESTORE": "HERSTELLEN", "EXIT": "AFSLUITEN", "SETTINGS": "INSTELLINGEN",
        "DATABASE": "DATABASE", "PATH": "PAD", "CURRENCY": "VALUTA", "LANGUAGE": "TAAL",
        "LOCALE": "REGIO", "HELP": "HELP", "ABOUT": "OVER", "NEW": "NIEUW", "SAVE": "OPSLAAN",
        "SEARCH": "ZOEKEN", "READY": "GEREED", "USER DATA": "GEBRUIKERSGEGEVENS", "SUPPLIERS": "LEVERANCIERS",
        "REPORTS": "RAPPORTEN", "WARRANTY TRACKING": "GARANTIEBEWAKING", "PARTS": "ONDERDELEN",
        "CUSTOMER INFORMATION": "KLANTINFORMATIE", "DEVICE INFORMATION": "APPARAATINFORMATIE",
        "SERVICE INFORMATION": "SERVICE-INFORMATIE", "DATE & WARRANTY": "DATUM & GARANTIE",
        "PARTS & SUPPLIER": "ONDERDELEN & LEVERANCIER", "ADD SUPPLIER": "LEVERANCIER TOEVOEGEN",
        "REPORT OPTIONS": "RAPPORTOPTIES", "WARRANTY FILTER": "GARANTIEFILTER",
        "PART INFORMATION": "ONDERDEELINFORMATIE", "FILTER PARTS": "ONDERDELEN FILTEREN",
        "NAME": "NAAM", "PHONE": "TELEFOON", "ADDRESS": "ADRES", "DEVICES": "APPARATEN",
        "DEVICE": "APPARAAT", "BRAND": "MERK", "MODEL": "MODEL", "ISSUE": "PROBLEEM",
        "DESCRIPTION": "BESCHRIJVING", "STATUS": "STATUS", "TECHNICIAN": "MONTEUR", "PRICE": "PRIJS",
        "SERVICE DATE": "SERVICEDATUM", "WARRANTY": "GARANTIE", "WARRANTY END": "EINDE GARANTIE",
        "PART": "ONDERDEEL", "SUPPLIER": "LEVERANCIER", "PURCHASE INVOICE": "INKOOPFACTUUR",
        "CONTACT": "CONTACT", "EMAIL": "EMAIL", "FROM": "VAN", "REPORT TYPE": "RAPPORTTYPE",
        "TYPE": "TYPE", "QUANTITY": "AANTAL", "UNIT PRICE": "STUKPRIJS", "MIN STOCK": "MINIMUMVOORRAAD",
        "LOCATION": "LOCATIE", "NOTES": "NOTITIES", "CUSTOMER": "KLANT",
        "WARRANTY PERIOD": "GARANTIEPERIODE", "END DATE": "EINDDATUM", "DAYS LEFT": "DAGEN OVER",
        "CLEAR": "WISSEN", "EDIT": "BEWERKEN", "DELETE": "VERWIJDEREN", "REFRESH": "VERVERSEN",
        "EXPORT SELECTED": "SELECTIE EXPORTEREN", "ADD": "TOEVOEGEN", "ADD PART": "ONDERDEEL TOEVOEGEN",
        "GENERATE REPORT": "RAPPORT MAKEN", "SEND REMINDER": "HERINNERING STUREN",
        "EDIT CONTACT": "CONTACT BEWERKEN", "TABLE VIEW": "TABELWEERGAVE", "TEXT VIEW": "TEKSTWEERGAVE",
        "CLOSE": "SLUITEN", "BROWSE": "BLADEREN", "CANCEL": "ANNULEREN", "USE DEFAULT": "STANDAARD GEBRUIKEN",
        "AUTO BACKUP ON EXIT": "AUTOMATISCHE BACK-UP BIJ AFSLUITEN", "LANGUAGE / LOCALE": "TAAL / REGIO",
    },
    "zh_CN": {
        "FILE": "文件", "IMPORT": "导入", "EXPORT": "导出", "OR": "或", "TO": "到",
        "BACKUP": "备份", "RESTORE": "恢复", "EXIT": "退出", "SETTINGS": "设置",
        "DATABASE": "数据库", "PATH": "路径", "CURRENCY": "货币", "LANGUAGE": "语言",
        "LOCALE": "区域", "HELP": "帮助", "ABOUT": "关于", "NEW": "新建", "SAVE": "保存",
        "SEARCH": "搜索", "READY": "就绪", "USER DATA": "用户数据", "SUPPLIERS": "供应商",
        "REPORTS": "报表", "WARRANTY TRACKING": "保修跟踪", "PARTS": "配件",
        "CUSTOMER INFORMATION": "客户信息", "DEVICE INFORMATION": "设备信息", "SERVICE INFORMATION": "服务信息",
        "DATE & WARRANTY": "日期和保修", "PARTS & SUPPLIER": "配件和供应商", "ADD SUPPLIER": "添加供应商",
        "REPORT OPTIONS": "报表选项", "WARRANTY FILTER": "保修筛选", "PART INFORMATION": "配件信息",
        "FILTER PARTS": "筛选配件", "NAME": "姓名", "PHONE": "电话", "ADDRESS": "地址",
        "DEVICES": "设备", "DEVICE": "设备", "BRAND": "品牌", "MODEL": "型号", "ISSUE": "问题",
        "DESCRIPTION": "描述", "STATUS": "状态", "TECHNICIAN": "技术员", "PRICE": "价格",
        "SERVICE DATE": "服务日期", "WARRANTY": "保修", "WARRANTY END": "保修结束",
        "PART": "配件", "SUPPLIER": "供应商", "PURCHASE INVOICE": "采购发票", "CONTACT": "联系人",
        "EMAIL": "电子邮件", "FROM": "从", "REPORT TYPE": "报表类型", "TYPE": "类型",
        "QUANTITY": "数量", "UNIT PRICE": "单价", "MIN STOCK": "最低库存", "LOCATION": "位置",
        "NOTES": "备注", "CUSTOMER": "客户", "WARRANTY PERIOD": "保修期限", "END DATE": "结束日期",
        "DAYS LEFT": "剩余天数", "CLEAR": "清除", "EDIT": "编辑", "DELETE": "删除", "REFRESH": "刷新",
        "EXPORT SELECTED": "导出所选", "ADD": "添加", "ADD PART": "添加配件", "GENERATE REPORT": "生成报表",
        "SEND REMINDER": "发送提醒", "EDIT CONTACT": "编辑联系人", "TABLE VIEW": "表格视图", "TEXT VIEW": "文本视图",
        "CLOSE": "关闭", "BROWSE": "浏览", "CANCEL": "取消", "USE DEFAULT": "使用默认值",
        "AUTO BACKUP ON EXIT": "退出时自动备份", "LANGUAGE / LOCALE": "语言 / 区域",
    },
    "ja_JP": {
        "FILE": "ファイル", "IMPORT": "インポート", "EXPORT": "エクスポート", "OR": "または", "TO": "へ",
        "BACKUP": "バックアップ", "RESTORE": "復元", "EXIT": "終了", "SETTINGS": "設定",
        "DATABASE": "データベース", "PATH": "パス", "CURRENCY": "通貨", "LANGUAGE": "言語",
        "LOCALE": "地域", "HELP": "ヘルプ", "ABOUT": "情報", "NEW": "新規", "SAVE": "保存",
        "SEARCH": "検索", "READY": "準備完了", "USER DATA": "ユーザーデータ", "SUPPLIERS": "仕入先",
        "REPORTS": "レポート", "WARRANTY TRACKING": "保証管理", "PARTS": "部品",
        "CUSTOMER INFORMATION": "顧客情報", "DEVICE INFORMATION": "機器情報", "SERVICE INFORMATION": "サービス情報",
        "DATE & WARRANTY": "日付と保証", "PARTS & SUPPLIER": "部品と仕入先", "ADD SUPPLIER": "仕入先を追加",
        "REPORT OPTIONS": "レポートオプション", "WARRANTY FILTER": "保証フィルター", "PART INFORMATION": "部品情報",
        "FILTER PARTS": "部品を絞り込み", "NAME": "名前", "PHONE": "電話", "ADDRESS": "住所",
        "DEVICES": "機器", "DEVICE": "機器", "BRAND": "ブランド", "MODEL": "モデル", "ISSUE": "問題",
        "DESCRIPTION": "説明", "STATUS": "状態", "TECHNICIAN": "技術者", "PRICE": "価格",
        "SERVICE DATE": "サービス日", "WARRANTY": "保証", "WARRANTY END": "保証終了",
        "PART": "部品", "SUPPLIER": "仕入先", "PURCHASE INVOICE": "購入請求書", "CONTACT": "連絡先",
        "EMAIL": "メール", "FROM": "開始", "REPORT TYPE": "レポート種別", "TYPE": "種類",
        "QUANTITY": "数量", "UNIT PRICE": "単価", "MIN STOCK": "最小在庫", "LOCATION": "場所",
        "NOTES": "メモ", "CUSTOMER": "顧客", "WARRANTY PERIOD": "保証期間", "END DATE": "終了日",
        "DAYS LEFT": "残り日数", "CLEAR": "クリア", "EDIT": "編集", "DELETE": "削除", "REFRESH": "更新",
        "EXPORT SELECTED": "選択をエクスポート", "ADD": "追加", "ADD PART": "部品を追加", "GENERATE REPORT": "レポート作成",
        "SEND REMINDER": "リマインダー送信", "EDIT CONTACT": "連絡先を編集", "TABLE VIEW": "表ビュー", "TEXT VIEW": "テキストビュー",
        "CLOSE": "閉じる", "BROWSE": "参照", "CANCEL": "キャンセル", "USE DEFAULT": "既定値を使用",
        "AUTO BACKUP ON EXIT": "終了時に自動バックアップ", "LANGUAGE / LOCALE": "言語 / 地域",
    },
    "ko_KR": {
        "FILE": "파일", "IMPORT": "가져오기", "EXPORT": "내보내기", "OR": "또는", "TO": "로",
        "BACKUP": "백업", "RESTORE": "복원", "EXIT": "종료", "SETTINGS": "설정",
        "DATABASE": "데이터베이스", "PATH": "경로", "CURRENCY": "통화", "LANGUAGE": "언어",
        "LOCALE": "지역", "HELP": "도움말", "ABOUT": "정보", "NEW": "새로 만들기", "SAVE": "저장",
        "SEARCH": "검색", "READY": "준비됨", "USER DATA": "사용자 데이터", "SUPPLIERS": "공급업체",
        "REPORTS": "보고서", "WARRANTY TRACKING": "보증 추적", "PARTS": "부품",
        "CUSTOMER INFORMATION": "고객 정보", "DEVICE INFORMATION": "장치 정보", "SERVICE INFORMATION": "서비스 정보",
        "DATE & WARRANTY": "날짜 및 보증", "PARTS & SUPPLIER": "부품 및 공급업체", "ADD SUPPLIER": "공급업체 추가",
        "REPORT OPTIONS": "보고서 옵션", "WARRANTY FILTER": "보증 필터", "PART INFORMATION": "부품 정보",
        "FILTER PARTS": "부품 필터", "NAME": "이름", "PHONE": "전화", "ADDRESS": "주소",
        "DEVICES": "장치", "DEVICE": "장치", "BRAND": "브랜드", "MODEL": "모델", "ISSUE": "문제",
        "DESCRIPTION": "설명", "STATUS": "상태", "TECHNICIAN": "기술자", "PRICE": "가격",
        "SERVICE DATE": "서비스 날짜", "WARRANTY": "보증", "WARRANTY END": "보증 종료",
        "PART": "부품", "SUPPLIER": "공급업체", "PURCHASE INVOICE": "구매 송장", "CONTACT": "연락처",
        "EMAIL": "이메일", "FROM": "시작", "REPORT TYPE": "보고서 유형", "TYPE": "유형",
        "QUANTITY": "수량", "UNIT PRICE": "단가", "MIN STOCK": "최소 재고", "LOCATION": "위치",
        "NOTES": "메모", "CUSTOMER": "고객", "WARRANTY PERIOD": "보증 기간", "END DATE": "종료일",
        "DAYS LEFT": "남은 일수", "CLEAR": "지우기", "EDIT": "편집", "DELETE": "삭제", "REFRESH": "새로 고침",
        "EXPORT SELECTED": "선택 내보내기", "ADD": "추가", "ADD PART": "부품 추가", "GENERATE REPORT": "보고서 생성",
        "SEND REMINDER": "알림 보내기", "EDIT CONTACT": "연락처 편집", "TABLE VIEW": "표 보기", "TEXT VIEW": "텍스트 보기",
        "CLOSE": "닫기", "BROWSE": "찾아보기", "CANCEL": "취소", "USE DEFAULT": "기본값 사용",
        "AUTO BACKUP ON EXIT": "종료 시 자동 백업", "LANGUAGE / LOCALE": "언어 / 지역",
    },
}

UI_TRANSLATIONS.update({
    code: {
        **UI_TRANSLATIONS.get(code, {}),
        "LANGUAGE / LOCALE": UI_TRANSLATIONS.get(code, {}).get("LANGUAGE / LOCALE", "LANGUAGE / LOCALE"),
    }
    for code in LANGUAGE_OPTIONS
})

UI_TRANSLATIONS.update({
    "ar_SA": {
        "FILE": "ملف", "IMPORT": "استيراد", "EXPORT": "تصدير", "OR": "أو", "TO": "إلى",
        "BACKUP": "نسخ احتياطي", "RESTORE": "استعادة", "EXIT": "خروج", "SETTINGS": "الإعدادات",
        "DATABASE": "قاعدة البيانات", "PATH": "المسار", "CURRENCY": "العملة", "LANGUAGE": "اللغة",
        "LOCALE": "المنطقة", "HELP": "مساعدة", "ABOUT": "حول", "NEW": "جديد", "SAVE": "حفظ",
        "SEARCH": "بحث", "READY": "جاهز", "USER DATA": "بيانات المستخدم", "SUPPLIERS": "الموردون",
        "REPORTS": "التقارير", "WARRANTY TRACKING": "تتبع الضمان", "PARTS": "القطع",
        "CUSTOMER INFORMATION": "معلومات العميل", "DEVICE INFORMATION": "معلومات الجهاز",
        "SERVICE INFORMATION": "معلومات الخدمة", "DATE & WARRANTY": "التاريخ والضمان",
        "PARTS & SUPPLIER": "القطع والمورد", "ADD SUPPLIER": "إضافة مورد",
        "REPORT OPTIONS": "خيارات التقرير", "WARRANTY FILTER": "مرشح الضمان",
        "PART INFORMATION": "معلومات القطعة", "FILTER PARTS": "تصفية القطع",
        "NAME": "الاسم", "PHONE": "الهاتف", "ADDRESS": "العنوان", "DEVICES": "الأجهزة",
        "DEVICE": "الجهاز", "BRAND": "العلامة", "MODEL": "الطراز", "ISSUE": "المشكلة",
        "DESCRIPTION": "الوصف", "STATUS": "الحالة", "TECHNICIAN": "الفني", "PRICE": "السعر",
        "SERVICE DATE": "تاريخ الخدمة", "WARRANTY": "الضمان", "WARRANTY END": "نهاية الضمان",
        "PART": "قطعة", "SUPPLIER": "مورد", "PURCHASE INVOICE": "فاتورة الشراء",
        "CONTACT": "جهة الاتصال", "EMAIL": "البريد الإلكتروني", "FROM": "من",
        "REPORT TYPE": "نوع التقرير", "TYPE": "النوع", "QUANTITY": "الكمية",
        "UNIT PRICE": "سعر الوحدة", "MIN STOCK": "الحد الأدنى للمخزون", "LOCATION": "الموقع",
        "NOTES": "ملاحظات", "CUSTOMER": "العميل", "WARRANTY PERIOD": "مدة الضمان",
        "END DATE": "تاريخ الانتهاء", "DAYS LEFT": "الأيام المتبقية", "CLEAR": "مسح",
        "EDIT": "تعديل", "DELETE": "حذف", "REFRESH": "تحديث", "EXPORT SELECTED": "تصدير المحدد",
        "ADD": "إضافة", "ADD PART": "إضافة قطعة", "GENERATE REPORT": "إنشاء تقرير",
        "SEND REMINDER": "إرسال تذكير", "EDIT CONTACT": "تعديل جهة الاتصال",
        "TABLE VIEW": "عرض الجدول", "TEXT VIEW": "عرض النص", "CLOSE": "إغلاق",
        "BROWSE": "استعراض", "CANCEL": "إلغاء", "USE DEFAULT": "استخدام الافتراضي",
        "AUTO BACKUP ON EXIT": "نسخ احتياطي تلقائي عند الخروج", "LANGUAGE / LOCALE": "اللغة / المنطقة",
    },
    "hi_IN": {
        "FILE": "फ़ाइल", "IMPORT": "आयात", "EXPORT": "निर्यात", "OR": "या", "TO": "को",
        "BACKUP": "बैकअप", "RESTORE": "पुनर्स्थापित", "EXIT": "बाहर निकलें", "SETTINGS": "सेटिंग्स",
        "DATABASE": "डेटाबेस", "PATH": "पथ", "CURRENCY": "मुद्रा", "LANGUAGE": "भाषा",
        "LOCALE": "स्थान", "HELP": "सहायता", "ABOUT": "के बारे में", "NEW": "नया", "SAVE": "सहेजें",
        "SEARCH": "खोजें", "READY": "तैयार", "USER DATA": "उपयोगकर्ता डेटा", "SUPPLIERS": "आपूर्तिकर्ता",
        "REPORTS": "रिपोर्ट", "WARRANTY TRACKING": "वारंटी ट्रैकिंग", "PARTS": "पार्ट्स",
        "CUSTOMER INFORMATION": "ग्राहक जानकारी", "DEVICE INFORMATION": "डिवाइस जानकारी",
        "SERVICE INFORMATION": "सेवा जानकारी", "DATE & WARRANTY": "तारीख और वारंटी",
        "PARTS & SUPPLIER": "पार्ट्स और आपूर्तिकर्ता", "ADD SUPPLIER": "आपूर्तिकर्ता जोड़ें",
        "REPORT OPTIONS": "रिपोर्ट विकल्प", "WARRANTY FILTER": "वारंटी फ़िल्टर",
        "PART INFORMATION": "पार्ट जानकारी", "FILTER PARTS": "पार्ट्स फ़िल्टर",
        "NAME": "नाम", "PHONE": "फ़ोन", "ADDRESS": "पता", "DEVICES": "डिवाइस",
        "DEVICE": "डिवाइस", "BRAND": "ब्रांड", "MODEL": "मॉडल", "ISSUE": "समस्या",
        "DESCRIPTION": "विवरण", "STATUS": "स्थिति", "TECHNICIAN": "तकनीशियन", "PRICE": "मूल्य",
        "SERVICE DATE": "सेवा तारीख", "WARRANTY": "वारंटी", "WARRANTY END": "वारंटी समाप्ति",
        "PART": "पार्ट", "SUPPLIER": "आपूर्तिकर्ता", "PURCHASE INVOICE": "खरीद चालान",
        "CONTACT": "संपर्क", "EMAIL": "ईमेल", "FROM": "से", "REPORT TYPE": "रिपोर्ट प्रकार",
        "TYPE": "प्रकार", "QUANTITY": "मात्रा", "UNIT PRICE": "इकाई मूल्य", "MIN STOCK": "न्यूनतम स्टॉक",
        "LOCATION": "स्थान", "NOTES": "नोट्स", "CUSTOMER": "ग्राहक", "WARRANTY PERIOD": "वारंटी अवधि",
        "END DATE": "अंतिम तारीख", "DAYS LEFT": "शेष दिन", "CLEAR": "साफ़ करें", "EDIT": "संपादित करें",
        "DELETE": "हटाएं", "REFRESH": "रीफ़्रेश", "EXPORT SELECTED": "चयनित निर्यात",
        "ADD": "जोड़ें", "ADD PART": "पार्ट जोड़ें", "GENERATE REPORT": "रिपोर्ट बनाएं",
        "SEND REMINDER": "अनुस्मारक भेजें", "EDIT CONTACT": "संपर्क संपादित करें",
        "TABLE VIEW": "तालिका दृश्य", "TEXT VIEW": "पाठ दृश्य", "CLOSE": "बंद करें",
        "BROWSE": "ब्राउज़", "CANCEL": "रद्द करें", "USE DEFAULT": "डिफ़ॉल्ट उपयोग करें",
        "AUTO BACKUP ON EXIT": "बाहर निकलते समय ऑटो बैकअप", "LANGUAGE / LOCALE": "भाषा / स्थान",
    },
    "bn_BD": {
        "FILE": "ফাইল", "IMPORT": "আমদানি", "EXPORT": "রপ্তানি", "OR": "অথবা", "TO": "এ",
        "BACKUP": "ব্যাকআপ", "RESTORE": "পুনরুদ্ধার", "EXIT": "প্রস্থান", "SETTINGS": "সেটিংস",
        "DATABASE": "ডাটাবেস", "PATH": "পথ", "CURRENCY": "মুদ্রা", "LANGUAGE": "ভাষা",
        "LOCALE": "লোকেল", "HELP": "সহায়তা", "ABOUT": "সম্পর্কে", "NEW": "নতুন", "SAVE": "সংরক্ষণ",
        "SEARCH": "অনুসন্ধান", "READY": "প্রস্তুত", "USER DATA": "ব্যবহারকারীর ডেটা", "SUPPLIERS": "সরবরাহকারী",
        "REPORTS": "রিপোর্ট", "WARRANTY TRACKING": "ওয়ারেন্টি ট্র্যাকিং", "PARTS": "পার্টস",
        "CUSTOMER INFORMATION": "গ্রাহকের তথ্য", "DEVICE INFORMATION": "ডিভাইস তথ্য",
        "SERVICE INFORMATION": "সেবা তথ্য", "DATE & WARRANTY": "তারিখ ও ওয়ারেন্টি",
        "PARTS & SUPPLIER": "পার্টস ও সরবরাহকারী", "ADD SUPPLIER": "সরবরাহকারী যোগ করুন",
        "REPORT OPTIONS": "রিপোর্ট অপশন", "WARRANTY FILTER": "ওয়ারেন্টি ফিল্টার",
        "PART INFORMATION": "পার্ট তথ্য", "FILTER PARTS": "পার্টস ফিল্টার",
        "NAME": "নাম", "PHONE": "ফোন", "ADDRESS": "ঠিকানা", "DEVICES": "ডিভাইস",
        "DEVICE": "ডিভাইস", "BRAND": "ব্র্যান্ড", "MODEL": "মডেল", "ISSUE": "সমস্যা",
        "DESCRIPTION": "বিবরণ", "STATUS": "অবস্থা", "TECHNICIAN": "টেকনিশিয়ান", "PRICE": "মূল্য",
        "SERVICE DATE": "সেবার তারিখ", "WARRANTY": "ওয়ারেন্টি", "WARRANTY END": "ওয়ারেন্টি শেষ",
        "PART": "পার্ট", "SUPPLIER": "সরবরাহকারী", "PURCHASE INVOICE": "ক্রয় ইনভয়েস",
        "CONTACT": "যোগাযোগ", "EMAIL": "ইমেইল", "FROM": "থেকে", "REPORT TYPE": "রিপোর্ট ধরন",
        "TYPE": "ধরন", "QUANTITY": "পরিমাণ", "UNIT PRICE": "একক মূল্য", "MIN STOCK": "ন্যূনতম স্টক",
        "LOCATION": "অবস্থান", "NOTES": "নোট", "CUSTOMER": "গ্রাহক", "WARRANTY PERIOD": "ওয়ারেন্টি সময়কাল",
        "END DATE": "শেষ তারিখ", "DAYS LEFT": "বাকি দিন", "CLEAR": "মুছুন", "EDIT": "সম্পাদনা",
        "DELETE": "মুছে ফেলুন", "REFRESH": "রিফ্রেশ", "EXPORT SELECTED": "নির্বাচিত রপ্তানি",
        "ADD": "যোগ করুন", "ADD PART": "পার্ট যোগ করুন", "GENERATE REPORT": "রিপোর্ট তৈরি",
        "SEND REMINDER": "রিমাইন্ডার পাঠান", "EDIT CONTACT": "যোগাযোগ সম্পাদনা",
        "TABLE VIEW": "টেবিল ভিউ", "TEXT VIEW": "টেক্সট ভিউ", "CLOSE": "বন্ধ", "BROWSE": "ব্রাউজ",
        "CANCEL": "বাতিল", "USE DEFAULT": "ডিফল্ট ব্যবহার", "AUTO BACKUP ON EXIT": "প্রস্থানে অটো ব্যাকআপ",
        "LANGUAGE / LOCALE": "ভাষা / লোকেল",
    },
    "ru_RU": {
        "FILE": "ФАЙЛ", "IMPORT": "ИМПОРТ", "EXPORT": "ЭКСПОРТ", "OR": "ИЛИ", "TO": "В",
        "BACKUP": "РЕЗЕРВНАЯ КОПИЯ", "RESTORE": "ВОССТАНОВИТЬ", "EXIT": "ВЫХОД", "SETTINGS": "НАСТРОЙКИ",
        "DATABASE": "БАЗА ДАННЫХ", "PATH": "ПУТЬ", "CURRENCY": "ВАЛЮТА", "LANGUAGE": "ЯЗЫК",
        "LOCALE": "РЕГИОН", "HELP": "ПОМОЩЬ", "ABOUT": "О ПРОГРАММЕ", "NEW": "НОВЫЙ", "SAVE": "СОХРАНИТЬ",
        "SEARCH": "ПОИСК", "READY": "ГОТОВО", "USER DATA": "ДАННЫЕ ПОЛЬЗОВАТЕЛЯ", "SUPPLIERS": "ПОСТАВЩИКИ",
        "REPORTS": "ОТЧЕТЫ", "WARRANTY TRACKING": "ОТСЛЕЖИВАНИЕ ГАРАНТИИ", "PARTS": "ЗАПЧАСТИ",
        "CUSTOMER INFORMATION": "ИНФОРМАЦИЯ О КЛИЕНТЕ", "DEVICE INFORMATION": "ИНФОРМАЦИЯ ОБ УСТРОЙСТВЕ",
        "SERVICE INFORMATION": "ИНФОРМАЦИЯ О СЕРВИСЕ", "DATE & WARRANTY": "ДАТА И ГАРАНТИЯ",
        "PARTS & SUPPLIER": "ЗАПЧАСТИ И ПОСТАВЩИК", "ADD SUPPLIER": "ДОБАВИТЬ ПОСТАВЩИКА",
        "REPORT OPTIONS": "ПАРАМЕТРЫ ОТЧЕТА", "WARRANTY FILTER": "ФИЛЬТР ГАРАНТИИ",
        "PART INFORMATION": "ИНФОРМАЦИЯ О ЗАПЧАСТИ", "FILTER PARTS": "ФИЛЬТР ЗАПЧАСТЕЙ",
        "NAME": "ИМЯ", "PHONE": "ТЕЛЕФОН", "ADDRESS": "АДРЕС", "DEVICES": "УСТРОЙСТВА",
        "DEVICE": "УСТРОЙСТВО", "BRAND": "БРЕНД", "MODEL": "МОДЕЛЬ", "ISSUE": "ПРОБЛЕМА",
        "DESCRIPTION": "ОПИСАНИЕ", "STATUS": "СТАТУС", "TECHNICIAN": "ТЕХНИК", "PRICE": "ЦЕНА",
        "SERVICE DATE": "ДАТА СЕРВИСА", "WARRANTY": "ГАРАНТИЯ", "WARRANTY END": "КОНЕЦ ГАРАНТИИ",
        "PART": "ЗАПЧАСТЬ", "SUPPLIER": "ПОСТАВЩИК", "PURCHASE INVOICE": "СЧЕТ ПОКУПКИ",
        "CONTACT": "КОНТАКТ", "EMAIL": "EMAIL", "FROM": "С", "REPORT TYPE": "ТИП ОТЧЕТА",
        "TYPE": "ТИП", "QUANTITY": "КОЛИЧЕСТВО", "UNIT PRICE": "ЦЕНА ЗА ЕД.", "MIN STOCK": "МИН. ЗАПАС",
        "LOCATION": "МЕСТО", "NOTES": "ЗАМЕТКИ", "CUSTOMER": "КЛИЕНТ", "WARRANTY PERIOD": "ПЕРИОД ГАРАНТИИ",
        "END DATE": "ДАТА ОКОНЧАНИЯ", "DAYS LEFT": "ДНЕЙ ОСТАЛОСЬ", "CLEAR": "ОЧИСТИТЬ",
        "EDIT": "ИЗМЕНИТЬ", "DELETE": "УДАЛИТЬ", "REFRESH": "ОБНОВИТЬ", "EXPORT SELECTED": "ЭКСПОРТ ВЫБРАННОГО",
        "ADD": "ДОБАВИТЬ", "ADD PART": "ДОБАВИТЬ ЗАПЧАСТЬ", "GENERATE REPORT": "СОЗДАТЬ ОТЧЕТ",
        "SEND REMINDER": "ОТПРАВИТЬ НАПОМИНАНИЕ", "EDIT CONTACT": "ИЗМЕНИТЬ КОНТАКТ",
        "TABLE VIEW": "ТАБЛИЦА", "TEXT VIEW": "ТЕКСТ", "CLOSE": "ЗАКРЫТЬ", "BROWSE": "ОБЗОР",
        "CANCEL": "ОТМЕНА", "USE DEFAULT": "ПО УМОЛЧАНИЮ", "AUTO BACKUP ON EXIT": "АВТО-БЭКАП ПРИ ВЫХОДЕ",
        "LANGUAGE / LOCALE": "ЯЗЫК / РЕГИОН",
    },
    "tr_TR": {
        "FILE": "DOSYA", "IMPORT": "İÇE AKTAR", "EXPORT": "DIŞA AKTAR", "OR": "VEYA", "TO": "HEDEF",
        "BACKUP": "YEDEKLE", "RESTORE": "GERİ YÜKLE", "EXIT": "ÇIKIŞ", "SETTINGS": "AYARLAR",
        "DATABASE": "VERİTABANI", "PATH": "YOL", "CURRENCY": "PARA BİRİMİ", "LANGUAGE": "DİL",
        "LOCALE": "YEREL", "HELP": "YARDIM", "ABOUT": "HAKKINDA", "NEW": "YENİ", "SAVE": "KAYDET",
        "SEARCH": "ARA", "READY": "HAZIR", "USER DATA": "KULLANICI VERİSİ", "SUPPLIERS": "TEDARİKÇİLER",
        "REPORTS": "RAPORLAR", "WARRANTY TRACKING": "GARANTİ TAKİBİ", "PARTS": "PARÇALAR",
        "CUSTOMER INFORMATION": "MÜŞTERİ BİLGİSİ", "DEVICE INFORMATION": "CİHAZ BİLGİSİ",
        "SERVICE INFORMATION": "SERVİS BİLGİSİ", "DATE & WARRANTY": "TARİH VE GARANTİ",
        "PARTS & SUPPLIER": "PARÇALAR VE TEDARİKÇİ", "ADD SUPPLIER": "TEDARİKÇİ EKLE",
        "REPORT OPTIONS": "RAPOR SEÇENEKLERİ", "WARRANTY FILTER": "GARANTİ FİLTRESİ",
        "PART INFORMATION": "PARÇA BİLGİSİ", "FILTER PARTS": "PARÇALARI FİLTRELE",
        "NAME": "AD", "PHONE": "TELEFON", "ADDRESS": "ADRES", "DEVICES": "CİHAZLAR",
        "DEVICE": "CİHAZ", "BRAND": "MARKA", "MODEL": "MODEL", "ISSUE": "SORUN",
        "DESCRIPTION": "AÇIKLAMA", "STATUS": "DURUM", "TECHNICIAN": "TEKNİSYEN", "PRICE": "FİYAT",
        "SERVICE DATE": "SERVİS TARİHİ", "WARRANTY": "GARANTİ", "WARRANTY END": "GARANTİ BİTİŞİ",
        "PART": "PARÇA", "SUPPLIER": "TEDARİKÇİ", "PURCHASE INVOICE": "SATIN ALMA FATURASI",
        "CONTACT": "KONTAK", "EMAIL": "EMAIL", "FROM": "BAŞLANGIÇ", "REPORT TYPE": "RAPOR TÜRÜ",
        "TYPE": "TÜR", "QUANTITY": "MİKTAR", "UNIT PRICE": "BİRİM FİYAT", "MIN STOCK": "MİN. STOK",
        "LOCATION": "KONUM", "NOTES": "NOTLAR", "CUSTOMER": "MÜŞTERİ", "WARRANTY PERIOD": "GARANTİ SÜRESİ",
        "END DATE": "BİTİŞ TARİHİ", "DAYS LEFT": "KALAN GÜN", "CLEAR": "TEMİZLE", "EDIT": "DÜZENLE",
        "DELETE": "SİL", "REFRESH": "YENİLE", "EXPORT SELECTED": "SEÇİLİLERİ DIŞA AKTAR",
        "ADD": "EKLE", "ADD PART": "PARÇA EKLE", "GENERATE REPORT": "RAPOR OLUŞTUR",
        "SEND REMINDER": "HATIRLATICI GÖNDER", "EDIT CONTACT": "KONTAĞI DÜZENLE",
        "TABLE VIEW": "TABLO GÖRÜNÜMÜ", "TEXT VIEW": "METİN GÖRÜNÜMÜ", "CLOSE": "KAPAT",
        "BROWSE": "GÖZAT", "CANCEL": "İPTAL", "USE DEFAULT": "VARSAYILAN", "AUTO BACKUP ON EXIT": "ÇIKIŞTA OTOMATİK YEDEKLE",
        "LANGUAGE / LOCALE": "DİL / YEREL",
    },
    "vi_VN": {
        "FILE": "TỆP", "IMPORT": "NHẬP", "EXPORT": "XUẤT", "OR": "HOẶC", "TO": "ĐẾN",
        "BACKUP": "SAO LƯU", "RESTORE": "KHÔI PHỤC", "EXIT": "THOÁT", "SETTINGS": "CÀI ĐẶT",
        "DATABASE": "CƠ SỞ DỮ LIỆU", "PATH": "ĐƯỜNG DẪN", "CURRENCY": "TIỀN TỆ", "LANGUAGE": "NGÔN NGỮ",
        "LOCALE": "KHU VỰC", "HELP": "TRỢ GIÚP", "ABOUT": "GIỚI THIỆU", "NEW": "MỚI", "SAVE": "LƯU",
        "SEARCH": "TÌM KIẾM", "READY": "SẴN SÀNG", "USER DATA": "DỮ LIỆU NGƯỜI DÙNG", "SUPPLIERS": "NHÀ CUNG CẤP",
        "REPORTS": "BÁO CÁO", "WARRANTY TRACKING": "THEO DÕI BẢO HÀNH", "PARTS": "LINH KIỆN",
        "CUSTOMER INFORMATION": "THÔNG TIN KHÁCH HÀNG", "DEVICE INFORMATION": "THÔNG TIN THIẾT BỊ",
        "SERVICE INFORMATION": "THÔNG TIN DỊCH VỤ", "DATE & WARRANTY": "NGÀY & BẢO HÀNH",
        "PARTS & SUPPLIER": "LINH KIỆN & NHÀ CUNG CẤP", "ADD SUPPLIER": "THÊM NHÀ CUNG CẤP",
        "REPORT OPTIONS": "TÙY CHỌN BÁO CÁO", "WARRANTY FILTER": "LỌC BẢO HÀNH",
        "PART INFORMATION": "THÔNG TIN LINH KIỆN", "FILTER PARTS": "LỌC LINH KIỆN",
        "NAME": "TÊN", "PHONE": "ĐIỆN THOẠI", "ADDRESS": "ĐỊA CHỈ", "DEVICES": "THIẾT BỊ",
        "DEVICE": "THIẾT BỊ", "BRAND": "THƯƠNG HIỆU", "MODEL": "MODEL", "ISSUE": "SỰ CỐ",
        "DESCRIPTION": "MÔ TẢ", "STATUS": "TRẠNG THÁI", "TECHNICIAN": "KỸ THUẬT VIÊN", "PRICE": "GIÁ",
        "SERVICE DATE": "NGÀY DỊCH VỤ", "WARRANTY": "BẢO HÀNH", "WARRANTY END": "HẾT BẢO HÀNH",
        "PART": "LINH KIỆN", "SUPPLIER": "NHÀ CUNG CẤP", "PURCHASE INVOICE": "HÓA ĐƠN MUA",
        "CONTACT": "LIÊN HỆ", "EMAIL": "EMAIL", "FROM": "TỪ", "REPORT TYPE": "LOẠI BÁO CÁO",
        "TYPE": "LOẠI", "QUANTITY": "SỐ LƯỢNG", "UNIT PRICE": "ĐƠN GIÁ", "MIN STOCK": "TỒN TỐI THIỂU",
        "LOCATION": "VỊ TRÍ", "NOTES": "GHI CHÚ", "CUSTOMER": "KHÁCH HÀNG", "WARRANTY PERIOD": "THỜI HẠN BẢO HÀNH",
        "END DATE": "NGÀY KẾT THÚC", "DAYS LEFT": "NGÀY CÒN LẠI", "CLEAR": "XÓA", "EDIT": "SỬA",
        "DELETE": "XÓA", "REFRESH": "LÀM MỚI", "EXPORT SELECTED": "XUẤT MỤC CHỌN",
        "ADD": "THÊM", "ADD PART": "THÊM LINH KIỆN", "GENERATE REPORT": "TẠO BÁO CÁO",
        "SEND REMINDER": "GỬI NHẮC NHỞ", "EDIT CONTACT": "SỬA LIÊN HỆ",
        "TABLE VIEW": "DẠNG BẢNG", "TEXT VIEW": "DẠNG VĂN BẢN", "CLOSE": "ĐÓNG",
        "BROWSE": "DUYỆT", "CANCEL": "HỦY", "USE DEFAULT": "DÙNG MẶC ĐỊNH",
        "AUTO BACKUP ON EXIT": "TỰ ĐỘNG SAO LƯU KHI THOÁT", "LANGUAGE / LOCALE": "NGÔN NGỮ / KHU VỰC",
    },
    "th_TH": {
        "FILE": "ไฟล์", "IMPORT": "นำเข้า", "EXPORT": "ส่งออก", "OR": "หรือ", "TO": "ไปยัง",
        "BACKUP": "สำรองข้อมูล", "RESTORE": "กู้คืน", "EXIT": "ออก", "SETTINGS": "การตั้งค่า",
        "DATABASE": "ฐานข้อมูล", "PATH": "เส้นทาง", "CURRENCY": "สกุลเงิน", "LANGUAGE": "ภาษา",
        "LOCALE": "ภูมิภาค", "HELP": "ช่วยเหลือ", "ABOUT": "เกี่ยวกับ", "NEW": "ใหม่", "SAVE": "บันทึก",
        "SEARCH": "ค้นหา", "READY": "พร้อม", "USER DATA": "ข้อมูลผู้ใช้", "SUPPLIERS": "ซัพพลายเออร์",
        "REPORTS": "รายงาน", "WARRANTY TRACKING": "ติดตามประกัน", "PARTS": "อะไหล่",
        "CUSTOMER INFORMATION": "ข้อมูลลูกค้า", "DEVICE INFORMATION": "ข้อมูลอุปกรณ์",
        "SERVICE INFORMATION": "ข้อมูลบริการ", "DATE & WARRANTY": "วันที่และประกัน",
        "PARTS & SUPPLIER": "อะไหล่และซัพพลายเออร์", "ADD SUPPLIER": "เพิ่มซัพพลายเออร์",
        "REPORT OPTIONS": "ตัวเลือกรายงาน", "WARRANTY FILTER": "ตัวกรองประกัน",
        "PART INFORMATION": "ข้อมูลอะไหล่", "FILTER PARTS": "กรองอะไหล่",
        "NAME": "ชื่อ", "PHONE": "โทรศัพท์", "ADDRESS": "ที่อยู่", "DEVICES": "อุปกรณ์",
        "DEVICE": "อุปกรณ์", "BRAND": "แบรนด์", "MODEL": "รุ่น", "ISSUE": "ปัญหา",
        "DESCRIPTION": "คำอธิบาย", "STATUS": "สถานะ", "TECHNICIAN": "ช่าง", "PRICE": "ราคา",
        "SERVICE DATE": "วันที่บริการ", "WARRANTY": "ประกัน", "WARRANTY END": "สิ้นสุดประกัน",
        "PART": "อะไหล่", "SUPPLIER": "ซัพพลายเออร์", "PURCHASE INVOICE": "ใบแจ้งหนี้ซื้อ",
        "CONTACT": "ผู้ติดต่อ", "EMAIL": "อีเมล", "FROM": "จาก", "REPORT TYPE": "ประเภทรายงาน",
        "TYPE": "ประเภท", "QUANTITY": "จำนวน", "UNIT PRICE": "ราคาต่อหน่วย", "MIN STOCK": "สต็อกขั้นต่ำ",
        "LOCATION": "ตำแหน่ง", "NOTES": "หมายเหตุ", "CUSTOMER": "ลูกค้า", "WARRANTY PERIOD": "ระยะประกัน",
        "END DATE": "วันที่สิ้นสุด", "DAYS LEFT": "วันที่เหลือ", "CLEAR": "ล้าง", "EDIT": "แก้ไข",
        "DELETE": "ลบ", "REFRESH": "รีเฟรช", "EXPORT SELECTED": "ส่งออกที่เลือก",
        "ADD": "เพิ่ม", "ADD PART": "เพิ่มอะไหล่", "GENERATE REPORT": "สร้างรายงาน",
        "SEND REMINDER": "ส่งเตือน", "EDIT CONTACT": "แก้ไขผู้ติดต่อ", "TABLE VIEW": "มุมมองตาราง",
        "TEXT VIEW": "มุมมองข้อความ", "CLOSE": "ปิด", "BROWSE": "เรียกดู", "CANCEL": "ยกเลิก",
        "USE DEFAULT": "ใช้ค่าเริ่มต้น", "AUTO BACKUP ON EXIT": "สำรองอัตโนมัติเมื่อออก",
        "LANGUAGE / LOCALE": "ภาษา / ภูมิภาค",
    },
    "ms_MY": {
        "FILE": "FAIL", "IMPORT": "IMPORT", "EXPORT": "EKSPORT", "OR": "ATAU", "TO": "KE",
        "BACKUP": "SANDARAN", "RESTORE": "PULIHKAN", "EXIT": "KELUAR", "SETTINGS": "TETAPAN",
        "DATABASE": "PANGKALAN DATA", "PATH": "LALUAN", "CURRENCY": "MATA WANG", "LANGUAGE": "BAHASA",
        "LOCALE": "LOKAL", "HELP": "BANTUAN", "ABOUT": "PERIHAL", "NEW": "BAHARU", "SAVE": "SIMPAN",
        "SEARCH": "CARI", "READY": "SEDIA", "USER DATA": "DATA PENGGUNA", "SUPPLIERS": "PEMBEKAL",
        "REPORTS": "LAPORAN", "WARRANTY TRACKING": "JEJAK JAMINAN", "PARTS": "ALAT GANTI",
        "CUSTOMER INFORMATION": "MAKLUMAT PELANGGAN", "DEVICE INFORMATION": "MAKLUMAT PERANTI",
        "SERVICE INFORMATION": "MAKLUMAT SERVIS", "DATE & WARRANTY": "TARIKH & JAMINAN",
        "PARTS & SUPPLIER": "ALAT GANTI & PEMBEKAL", "ADD SUPPLIER": "TAMBAH PEMBEKAL",
        "REPORT OPTIONS": "PILIHAN LAPORAN", "WARRANTY FILTER": "PENAPIS JAMINAN",
        "PART INFORMATION": "MAKLUMAT ALAT GANTI", "FILTER PARTS": "TAPIS ALAT GANTI",
        "NAME": "NAMA", "PHONE": "TELEFON", "ADDRESS": "ALAMAT", "DEVICES": "PERANTI",
        "DEVICE": "PERANTI", "BRAND": "JENAMA", "MODEL": "MODEL", "ISSUE": "MASALAH",
        "DESCRIPTION": "PENERANGAN", "STATUS": "STATUS", "TECHNICIAN": "JURUTEKNIK", "PRICE": "HARGA",
        "SERVICE DATE": "TARIKH SERVIS", "WARRANTY": "JAMINAN", "WARRANTY END": "TAMAT JAMINAN",
        "PART": "ALAT GANTI", "SUPPLIER": "PEMBEKAL", "PURCHASE INVOICE": "INVOIS BELIAN",
        "CONTACT": "KONTAK", "EMAIL": "EMAIL", "FROM": "DARI", "REPORT TYPE": "JENIS LAPORAN",
        "TYPE": "JENIS", "QUANTITY": "KUANTITI", "UNIT PRICE": "HARGA UNIT", "MIN STOCK": "STOK MINIMUM",
        "LOCATION": "LOKASI", "NOTES": "NOTA", "CUSTOMER": "PELANGGAN", "WARRANTY PERIOD": "TEMPOH JAMINAN",
        "END DATE": "TARIKH TAMAT", "DAYS LEFT": "BAKI HARI", "CLEAR": "KOSONGKAN", "EDIT": "EDIT",
        "DELETE": "PADAM", "REFRESH": "SEGAR SEMULA", "EXPORT SELECTED": "EKSPORT PILIHAN",
        "ADD": "TAMBAH", "ADD PART": "TAMBAH ALAT GANTI", "GENERATE REPORT": "JANA LAPORAN",
        "SEND REMINDER": "HANTAR PERINGATAN", "EDIT CONTACT": "EDIT KONTAK", "TABLE VIEW": "PAPARAN JADUAL",
        "TEXT VIEW": "PAPARAN TEKS", "CLOSE": "TUTUP", "BROWSE": "SEMAK IMBAS", "CANCEL": "BATAL",
        "USE DEFAULT": "GUNA LALAI", "AUTO BACKUP ON EXIT": "SANDARAN AUTO SEMASA KELUAR",
        "LANGUAGE / LOCALE": "BAHASA / LOKAL",
    },
    "fil_PH": {
        "FILE": "FILE", "IMPORT": "IMPORT", "EXPORT": "EXPORT", "OR": "O", "TO": "SA",
        "BACKUP": "BACKUP", "RESTORE": "IBALIK", "EXIT": "LUMABAS", "SETTINGS": "SETTINGS",
        "DATABASE": "DATABASE", "PATH": "PATH", "CURRENCY": "PERA", "LANGUAGE": "WIKA",
        "LOCALE": "LOCALE", "HELP": "TULONG", "ABOUT": "TUNGKOL", "NEW": "BAGO", "SAVE": "I-SAVE",
        "SEARCH": "HANAP", "READY": "HANDA", "USER DATA": "DATA NG USER", "SUPPLIERS": "SUPPLIERS",
        "REPORTS": "REPORTS", "WARRANTY TRACKING": "PAGSUBAYBAY NG WARRANTY", "PARTS": "PIYESA",
        "CUSTOMER INFORMATION": "IMPORMASYON NG CUSTOMER", "DEVICE INFORMATION": "IMPORMASYON NG DEVICE",
        "SERVICE INFORMATION": "IMPORMASYON NG SERBISYO", "DATE & WARRANTY": "PETSA AT WARRANTY",
        "PARTS & SUPPLIER": "PIYESA AT SUPPLIER", "ADD SUPPLIER": "MAGDAGDAG NG SUPPLIER",
        "REPORT OPTIONS": "OPSYON NG REPORT", "WARRANTY FILTER": "FILTER NG WARRANTY",
        "PART INFORMATION": "IMPORMASYON NG PIYESA", "FILTER PARTS": "I-FILTER ANG PIYESA",
        "NAME": "PANGALAN", "PHONE": "TELEPONO", "ADDRESS": "ADDRESS", "DEVICES": "DEVICES",
        "DEVICE": "DEVICE", "BRAND": "BRAND", "MODEL": "MODEL", "ISSUE": "ISYU",
        "DESCRIPTION": "DESCRIPTION", "STATUS": "STATUS", "TECHNICIAN": "TECHNICIAN", "PRICE": "PRESYO",
        "SERVICE DATE": "PETSA NG SERBISYO", "WARRANTY": "WARRANTY", "WARRANTY END": "TAPOS NG WARRANTY",
        "PART": "PIYESA", "SUPPLIER": "SUPPLIER", "PURCHASE INVOICE": "PURCHASE INVOICE",
        "CONTACT": "CONTACT", "EMAIL": "EMAIL", "FROM": "MULA", "REPORT TYPE": "URI NG REPORT",
        "TYPE": "URI", "QUANTITY": "DAMI", "UNIT PRICE": "UNIT PRICE", "MIN STOCK": "MIN STOCK",
        "LOCATION": "LOKASYON", "NOTES": "NOTES", "CUSTOMER": "CUSTOMER", "WARRANTY PERIOD": "PANAHON NG WARRANTY",
        "END DATE": "PETSA NG TAPOS", "DAYS LEFT": "ARAW NA NATITIRA", "CLEAR": "CLEAR", "EDIT": "EDIT",
        "DELETE": "DELETE", "REFRESH": "REFRESH", "EXPORT SELECTED": "EXPORT SELECTED",
        "ADD": "ADD", "ADD PART": "ADD PART", "GENERATE REPORT": "GUMAWA NG REPORT",
        "SEND REMINDER": "MAGPADALA NG PAALALA", "EDIT CONTACT": "EDIT CONTACT",
        "TABLE VIEW": "TABLE VIEW", "TEXT VIEW": "TEXT VIEW", "CLOSE": "ISARA",
        "BROWSE": "BROWSE", "CANCEL": "KANSELA", "USE DEFAULT": "GAMITIN ANG DEFAULT",
        "AUTO BACKUP ON EXIT": "AUTO BACKUP SA PAGLABAS", "LANGUAGE / LOCALE": "WIKA / LOCALE",
    },
})

EXTRA_UI_TRANSLATIONS = {
    "id_ID": {
        "SELECT": "PILIH", "SELECTED": "TERPILIH", "ROW": "BARIS", "ROWS": "BARIS",
        "RECORD": "DATA", "RECORDS": "DATA", "DATA": "DATA", "COMPLETE": "SELESAI",
        "SUCCESSFUL": "BERHASIL", "SUCCESSFULLY": "BERHASIL", "FAILED": "GAGAL",
        "IMPORTED": "DIIMPOR", "EXPORTED": "DIEKSPOR", "UPDATED": "DIPERBARUI",
        "ADDED": "DITAMBAH", "DELETED": "DIHAPUS", "LOADED": "DIMUAT", "FORM": "FORMULIR",
        "CLEARED": "DIBERSIHKAN", "REQUIRED": "WAJIB", "INVALID": "TIDAK VALID",
        "DRIVE": "DRIVE", "FOLDER": "FOLDER", "FILE": "BERKAS", "FOUND": "DITEMUKAN",
        "NOT": "TIDAK", "OPEN": "BUKA", "EXISTING": "YANG ADA", "CREATE": "BUAT",
        "REBUILD": "BANGUN ULANG", "DUMMY": "DUMMY", "CATEGORY": "KATEGORI",
        "EMPTY": "KOSONG", "VALUE": "NILAI", "VALUES": "NILAI", "PROTECTED": "DILINDUNGI",
        "DUPLICATE": "DUPLIKAT", "EXISTS": "SUDAH ADA", "CURRENT": "SAAT INI",
        "CONTINUE": "LANJUTKAN", "CHOOSE": "PILIH", "ANOTHER": "LAIN", "REPLACE": "MENGGANTI",
        "ACTION": "TINDAKAN", "UNDONE": "DIBATALKAN", "WARNING": "PERINGATAN",
        "COPIED": "DISALIN", "CLIPBOARD": "CLIPBOARD", "LINK": "LINK", "CUSTOMERS": "PELANGGAN",
        "REMINDER": "PENGINGAT", "REMINDERS": "PENGINGAT", "SENT": "TERKIRIM",
        "LOGGED": "DICATAT", "PERIOD": "PERIODE", "GENERATED": "DIBUAT",
        "SUMMARY": "RINGKASAN", "TOTAL": "TOTAL", "SERVICES": "SERVIS", "REVENUE": "PENDAPATAN",
        "AVG": "RATA-RATA", "MONTH": "BULAN", "ACTIVE": "AKTIF", "EXPIRED": "KEDALUWARSA",
        "EXPIRING": "AKAN BERAKHIR", "SOON": "SEGERA", "WAITING": "MENUNGGU",
        "PROGRESS": "PROSES", "REPAIRED": "DIPERBAIKI", "PENDING": "TERTUNDA",
        "UNIQUE": "UNIK", "LOW": "RENDAH", "STOCK": "STOK", "YES": "YA", "NO": "TIDAK",
        "APPLICATION": "APLIKASI", "FIRST-TIME": "PERTAMA KALI", "SETUP": "PENGATURAN",
        "CURRENT DATE": "TANGGAL SAAT INI", "ALL FILES": "SEMUA BERKAS", "DATABASE FILES": "BERKAS DATABASE",
    },
    "es_MX": {
        "SELECT": "SELECCIONAR", "SELECTED": "SELECCIONADO", "ROW": "FILA", "ROWS": "FILAS",
        "RECORD": "REGISTRO", "RECORDS": "REGISTROS", "DATA": "DATOS", "COMPLETE": "COMPLETO",
        "SUCCESSFUL": "EXITOSO", "SUCCESSFULLY": "CORRECTAMENTE", "FAILED": "FALLÓ",
        "IMPORTED": "IMPORTADO", "EXPORTED": "EXPORTADO", "UPDATED": "ACTUALIZADO",
        "ADDED": "AGREGADO", "DELETED": "ELIMINADO", "LOADED": "CARGADO", "FORM": "FORMULARIO",
        "CLEARED": "LIMPIO", "REQUIRED": "OBLIGATORIO", "INVALID": "INVÁLIDO", "DRIVE": "UNIDAD",
        "FOLDER": "CARPETA", "FILE": "ARCHIVO", "FOUND": "ENCONTRADO", "NOT": "NO",
        "OPEN": "ABRIR", "EXISTING": "EXISTENTE", "CREATE": "CREAR", "REBUILD": "RECONSTRUIR",
        "DUMMY": "DUMMY", "CATEGORY": "CATEGORÍA", "EMPTY": "VACÍO", "VALUE": "VALOR",
        "VALUES": "VALORES", "PROTECTED": "PROTEGIDO", "DUPLICATE": "DUPLICADO", "EXISTS": "YA EXISTE",
        "CURRENT": "ACTUAL", "CONTINUE": "CONTINUAR", "CHOOSE": "ELEGIR", "ANOTHER": "OTRA",
        "REPLACE": "REEMPLAZAR", "ACTION": "ACCIÓN", "UNDONE": "DESHACER", "WARNING": "ADVERTENCIA",
        "COPIED": "COPIADO", "CLIPBOARD": "PORTAPAPELES", "LINK": "ENLACE", "CUSTOMERS": "CLIENTES",
        "REMINDER": "RECORDATORIO", "REMINDERS": "RECORDATORIOS", "SENT": "ENVIADO",
        "LOGGED": "REGISTRADO", "PERIOD": "PERIODO", "GENERATED": "GENERADO", "SUMMARY": "RESUMEN",
        "TOTAL": "TOTAL", "SERVICES": "SERVICIOS", "REVENUE": "INGRESOS", "AVG": "PROMEDIO",
        "MONTH": "MES", "ACTIVE": "ACTIVO", "EXPIRED": "VENCIDO", "EXPIRING": "POR VENCER",
        "SOON": "PRONTO", "WAITING": "ESPERANDO", "PROGRESS": "PROGRESO", "REPAIRED": "REPARADO",
        "PENDING": "PENDIENTE", "UNIQUE": "ÚNICO", "LOW": "BAJO", "STOCK": "STOCK",
        "YES": "SÍ", "NO": "NO", "APPLICATION": "APLICACIÓN", "FIRST-TIME": "PRIMERA VEZ",
        "SETUP": "CONFIGURACIÓN", "CURRENT DATE": "FECHA ACTUAL", "ALL FILES": "TODOS LOS ARCHIVOS",
        "DATABASE FILES": "ARCHIVOS DE BASE DE DATOS",
    },
    "fr_FR": {
        "SELECT": "SÉLECTIONNER", "SELECTED": "SÉLECTIONNÉ", "ROW": "LIGNE", "ROWS": "LIGNES",
        "RECORD": "ENREGISTREMENT", "RECORDS": "ENREGISTREMENTS", "DATA": "DONNÉES",
        "COMPLETE": "TERMINÉ", "SUCCESSFUL": "RÉUSSI", "SUCCESSFULLY": "AVEC SUCCÈS", "FAILED": "ÉCHEC",
        "IMPORTED": "IMPORTÉ", "EXPORTED": "EXPORTÉ", "UPDATED": "MIS À JOUR", "ADDED": "AJOUTÉ",
        "DELETED": "SUPPRIMÉ", "LOADED": "CHARGÉ", "FORM": "FORMULAIRE", "CLEARED": "EFFACÉ",
        "REQUIRED": "REQUIS", "INVALID": "INVALIDE", "DRIVE": "LECTEUR", "FOLDER": "DOSSIER",
        "FILE": "FICHIER", "FOUND": "TROUVÉ", "NOT": "NON", "OPEN": "OUVRIR", "EXISTING": "EXISTANT",
        "CREATE": "CRÉER", "REBUILD": "RECONSTRUIRE", "CATEGORY": "CATÉGORIE", "EMPTY": "VIDE",
        "VALUE": "VALEUR", "VALUES": "VALEURS", "PROTECTED": "PROTÉGÉ", "DUPLICATE": "DOUBLON",
        "EXISTS": "EXISTE DÉJÀ", "CURRENT": "ACTUEL", "CONTINUE": "CONTINUER", "CHOOSE": "CHOISIR",
        "ANOTHER": "AUTRE", "REPLACE": "REMPLACER", "ACTION": "ACTION", "UNDONE": "ANNULÉ",
        "WARNING": "AVERTISSEMENT", "COPIED": "COPIÉ", "CLIPBOARD": "PRESSE-PAPIERS",
        "LINK": "LIEN", "CUSTOMERS": "CLIENTS", "REMINDER": "RAPPEL", "REMINDERS": "RAPPELS",
        "SENT": "ENVOYÉ", "LOGGED": "ENREGISTRÉ", "PERIOD": "PÉRIODE", "GENERATED": "GÉNÉRÉ",
        "SUMMARY": "RÉSUMÉ", "TOTAL": "TOTAL", "SERVICES": "SERVICES", "REVENUE": "REVENU",
        "AVG": "MOYENNE", "MONTH": "MOIS", "ACTIVE": "ACTIF", "EXPIRED": "EXPIRÉ",
        "EXPIRING": "EXPIRE", "SOON": "BIENTÔT", "WAITING": "EN ATTENTE", "PROGRESS": "PROGRÈS",
        "REPAIRED": "RÉPARÉ", "PENDING": "EN ATTENTE", "UNIQUE": "UNIQUE", "LOW": "FAIBLE",
        "STOCK": "STOCK", "YES": "OUI", "NO": "NON", "APPLICATION": "APPLICATION",
        "FIRST-TIME": "PREMIÈRE FOIS", "SETUP": "CONFIGURATION", "CURRENT DATE": "DATE ACTUELLE",
        "ALL FILES": "TOUS LES FICHIERS", "DATABASE FILES": "FICHIERS DE BASE DE DONNÉES",
    },
    "de_DE": {
        "SELECT": "AUSWÄHLEN", "SELECTED": "AUSGEWÄHLT", "ROW": "ZEILE", "ROWS": "ZEILEN",
        "RECORD": "DATENSATZ", "RECORDS": "DATENSÄTZE", "DATA": "DATEN", "COMPLETE": "ABGESCHLOSSEN",
        "SUCCESSFUL": "ERFOLGREICH", "SUCCESSFULLY": "ERFOLGREICH", "FAILED": "FEHLGESCHLAGEN",
        "IMPORTED": "IMPORTIERT", "EXPORTED": "EXPORTIERT", "UPDATED": "AKTUALISIERT",
        "ADDED": "HINZUGEFÜGT", "DELETED": "GELÖSCHT", "LOADED": "GELADEN", "FORM": "FORMULAR",
        "CLEARED": "GELEERT", "REQUIRED": "ERFORDERLICH", "INVALID": "UNGÜLTIG", "DRIVE": "LAUFWERK",
        "FOLDER": "ORDNER", "FILE": "DATEI", "FOUND": "GEFUNDEN", "NOT": "NICHT",
        "OPEN": "ÖFFNEN", "EXISTING": "VORHANDEN", "CREATE": "ERSTELLEN", "REBUILD": "NEU ERSTELLEN",
        "CATEGORY": "KATEGORIE", "EMPTY": "LEER", "VALUE": "WERT", "VALUES": "WERTE",
        "PROTECTED": "GESCHÜTZT", "DUPLICATE": "DUPLIKAT", "EXISTS": "EXISTIERT BEREITS",
        "CURRENT": "AKTUELL", "CONTINUE": "FORTFAHREN", "CHOOSE": "WÄHLEN", "ANOTHER": "ANDEREN",
        "REPLACE": "ERSETZEN", "ACTION": "AKTION", "UNDONE": "RÜCKGÄNGIG", "WARNING": "WARNUNG",
        "COPIED": "KOPIERT", "CLIPBOARD": "ZWISCHENABLAGE", "LINK": "LINK", "CUSTOMERS": "KUNDEN",
        "REMINDER": "ERINNERUNG", "REMINDERS": "ERINNERUNGEN", "SENT": "GESENDET",
        "LOGGED": "PROTOKOLLIERT", "PERIOD": "ZEITRAUM", "GENERATED": "ERSTELLT",
        "SUMMARY": "ZUSAMMENFASSUNG", "TOTAL": "GESAMT", "SERVICES": "SERVICES",
        "REVENUE": "UMSATZ", "AVG": "DURCHSCHNITT", "MONTH": "MONAT", "ACTIVE": "AKTIV",
        "EXPIRED": "ABGELAUFEN", "EXPIRING": "LÄUFT AB", "SOON": "BALD", "WAITING": "WARTEND",
        "PROGRESS": "FORTSCHRITT", "REPAIRED": "REPARIERT", "PENDING": "AUSSTEHEND",
        "UNIQUE": "EINDEUTIG", "LOW": "NIEDRIG", "STOCK": "BESTAND", "YES": "JA", "NO": "NEIN",
        "APPLICATION": "ANWENDUNG", "FIRST-TIME": "ERSTMALS", "SETUP": "EINRICHTUNG",
        "CURRENT DATE": "AKTUELLES DATUM", "ALL FILES": "ALLE DATEIEN", "DATABASE FILES": "DATENBANKDATEIEN",
    },
    "pt_BR": {
        "SELECT": "SELECIONAR", "SELECTED": "SELECIONADO", "ROW": "LINHA", "ROWS": "LINHAS",
        "RECORD": "REGISTRO", "RECORDS": "REGISTROS", "DATA": "DADOS", "COMPLETE": "CONCLUÍDO",
        "SUCCESSFUL": "BEM-SUCEDIDO", "SUCCESSFULLY": "COM SUCESSO", "FAILED": "FALHOU",
        "IMPORTED": "IMPORTADO", "EXPORTED": "EXPORTADO", "UPDATED": "ATUALIZADO",
        "ADDED": "ADICIONADO", "DELETED": "EXCLUÍDO", "LOADED": "CARREGADO", "FORM": "FORMULÁRIO",
        "CLEARED": "LIMPO", "REQUIRED": "OBRIGATÓRIO", "INVALID": "INVÁLIDO", "DRIVE": "UNIDADE",
        "FOLDER": "PASTA", "FILE": "ARQUIVO", "FOUND": "ENCONTRADO", "NOT": "NÃO",
        "OPEN": "ABRIR", "EXISTING": "EXISTENTE", "CREATE": "CRIAR", "REBUILD": "RECRIAR",
        "CATEGORY": "CATEGORIA", "EMPTY": "VAZIO", "VALUE": "VALOR", "VALUES": "VALORES",
        "PROTECTED": "PROTEGIDO", "DUPLICATE": "DUPLICADO", "EXISTS": "JÁ EXISTE",
        "CURRENT": "ATUAL", "CONTINUE": "CONTINUAR", "CHOOSE": "ESCOLHER", "ANOTHER": "OUTRA",
        "REPLACE": "SUBSTITUIR", "ACTION": "AÇÃO", "UNDONE": "DESFEITO", "WARNING": "AVISO",
        "COPIED": "COPIADO", "CLIPBOARD": "ÁREA DE TRANSFERÊNCIA", "LINK": "LINK", "CUSTOMERS": "CLIENTES",
        "REMINDER": "LEMBRETE", "REMINDERS": "LEMBRETES", "SENT": "ENVIADO", "LOGGED": "REGISTRADO",
        "PERIOD": "PERÍODO", "GENERATED": "GERADO", "SUMMARY": "RESUMO", "TOTAL": "TOTAL",
        "SERVICES": "SERVIÇOS", "REVENUE": "RECEITA", "AVG": "MÉDIA", "MONTH": "MÊS",
        "ACTIVE": "ATIVO", "EXPIRED": "EXPIRADO", "EXPIRING": "EXPIRANDO", "SOON": "EM BREVE",
        "WAITING": "AGUARDANDO", "PROGRESS": "PROGRESSO", "REPAIRED": "REPARADO",
        "PENDING": "PENDENTE", "UNIQUE": "ÚNICO", "LOW": "BAIXO", "STOCK": "ESTOQUE",
        "YES": "SIM", "NO": "NÃO", "APPLICATION": "APLICAÇÃO", "FIRST-TIME": "PRIMEIRA VEZ",
        "SETUP": "CONFIGURAÇÃO", "CURRENT DATE": "DATA ATUAL", "ALL FILES": "TODOS OS ARQUIVOS",
        "DATABASE FILES": "ARQUIVOS DE BANCO DE DADOS",
    },
}

_GENERIC_EXTRA = {
    "it_IT": {"SELECT": "SELEZIONA", "SELECTED": "SELEZIONATO", "ROW": "RIGA", "ROWS": "RIGHE", "RECORD": "RECORD", "RECORDS": "RECORD", "DATA": "DATI", "COMPLETE": "COMPLETO", "SUCCESSFUL": "RIUSCITO", "SUCCESSFULLY": "CON SUCCESSO", "FAILED": "FALLITO", "IMPORTED": "IMPORTATO", "EXPORTED": "ESPORTATO", "UPDATED": "AGGIORNATO", "ADDED": "AGGIUNTO", "DELETED": "ELIMINATO", "LOADED": "CARICATO", "FORM": "MODULO", "CLEARED": "PULITO", "REQUIRED": "RICHIESTO", "INVALID": "NON VALIDO", "DRIVE": "UNITÀ", "FOLDER": "CARTELLA", "FOUND": "TROVATO", "NOT": "NON", "OPEN": "APRI", "EXISTING": "ESISTENTE", "CREATE": "CREA", "REBUILD": "RICOSTRUISCI", "CATEGORY": "CATEGORIA", "EMPTY": "VUOTO", "VALUE": "VALORE", "VALUES": "VALORI", "CURRENT": "CORRENTE", "WARNING": "AVVISO", "COPIED": "COPIATO", "CLIPBOARD": "APPUNTI", "CUSTOMERS": "CLIENTI", "REMINDER": "PROMEMORIA", "REMINDERS": "PROMEMORIA", "SENT": "INVIATO", "PERIOD": "PERIODO", "GENERATED": "GENERATO", "SUMMARY": "RIEPILOGO", "TOTAL": "TOTALE", "SERVICES": "SERVIZI", "REVENUE": "RICAVI", "AVG": "MEDIA", "MONTH": "MESE", "ACTIVE": "ATTIVO", "EXPIRED": "SCADUTO", "EXPIRING": "IN SCADENZA", "SOON": "PRESTO", "WAITING": "IN ATTESA", "PROGRESS": "AVANZAMENTO", "PENDING": "IN SOSPESO", "LOW": "BASSO", "STOCK": "SCORTA", "YES": "SÌ", "NO": "NO", "APPLICATION": "APPLICAZIONE", "SETUP": "CONFIGURAZIONE"},
    "nl_NL": {"SELECT": "SELECTEREN", "SELECTED": "GESELECTEERD", "ROW": "RIJ", "ROWS": "RIJEN", "RECORD": "RECORD", "RECORDS": "RECORDS", "DATA": "GEGEVENS", "COMPLETE": "VOLTOOID", "SUCCESSFUL": "GESLAAGD", "SUCCESSFULLY": "SUCCESVOL", "FAILED": "MISLUKT", "IMPORTED": "GEÏMPORTEERD", "EXPORTED": "GEËXPORTEERD", "UPDATED": "BIJGEWERKT", "ADDED": "TOEGEVOEGD", "DELETED": "VERWIJDERD", "LOADED": "GELADEN", "FORM": "FORMULIER", "CLEARED": "GEWIST", "REQUIRED": "VEREIST", "INVALID": "ONGELDIG", "DRIVE": "SCHIJF", "FOLDER": "MAP", "FOUND": "GEVONDEN", "NOT": "NIET", "OPEN": "OPENEN", "EXISTING": "BESTAAND", "CREATE": "MAKEN", "REBUILD": "OPNIEUW MAKEN", "CATEGORY": "CATEGORIE", "EMPTY": "LEEG", "VALUE": "WAARDE", "VALUES": "WAARDEN", "CURRENT": "HUIDIG", "WARNING": "WAARSCHUWING", "COPIED": "GEKOPIEERD", "CLIPBOARD": "KLEMBORD", "CUSTOMERS": "KLANTEN", "REMINDER": "HERINNERING", "REMINDERS": "HERINNERINGEN", "SENT": "VERZONDEN", "PERIOD": "PERIODE", "GENERATED": "GEGENEREERD", "SUMMARY": "SAMENVATTING", "TOTAL": "TOTAAL", "SERVICES": "SERVICES", "REVENUE": "OMZET", "AVG": "GEMIDDELD", "MONTH": "MAAND", "ACTIVE": "ACTIEF", "EXPIRED": "VERLOPEN", "EXPIRING": "VERLOOPT", "SOON": "BINNENKORT", "WAITING": "WACHTEND", "PROGRESS": "VOORTGANG", "PENDING": "IN AFWACHTING", "LOW": "LAAG", "STOCK": "VOORRAAD", "YES": "JA", "NO": "NEE", "APPLICATION": "APPLICATIE", "SETUP": "INSTALLATIE"},
    "zh_CN": {"SELECT": "选择", "SELECTED": "已选择", "ROW": "行", "ROWS": "行", "RECORD": "记录", "RECORDS": "记录", "DATA": "数据", "COMPLETE": "完成", "SUCCESSFUL": "成功", "SUCCESSFULLY": "成功", "FAILED": "失败", "IMPORTED": "已导入", "EXPORTED": "已导出", "UPDATED": "已更新", "ADDED": "已添加", "DELETED": "已删除", "LOADED": "已加载", "FORM": "表单", "CLEARED": "已清除", "REQUIRED": "必填", "INVALID": "无效", "DRIVE": "驱动器", "FOLDER": "文件夹", "FOUND": "找到", "NOT": "未", "OPEN": "打开", "EXISTING": "现有", "CREATE": "创建", "REBUILD": "重建", "CATEGORY": "类别", "EMPTY": "空", "VALUE": "值", "VALUES": "值", "CURRENT": "当前", "WARNING": "警告", "COPIED": "已复制", "CLIPBOARD": "剪贴板", "CUSTOMERS": "客户", "REMINDER": "提醒", "REMINDERS": "提醒", "SENT": "已发送", "PERIOD": "期间", "GENERATED": "已生成", "SUMMARY": "摘要", "TOTAL": "总计", "SERVICES": "服务", "REVENUE": "收入", "AVG": "平均", "MONTH": "月份", "ACTIVE": "有效", "EXPIRED": "已过期", "EXPIRING": "即将过期", "SOON": "很快", "WAITING": "等待", "PROGRESS": "进度", "PENDING": "待处理", "LOW": "低", "STOCK": "库存", "YES": "是", "NO": "否", "APPLICATION": "应用程序", "SETUP": "设置"},
    "ja_JP": {"SELECT": "選択", "SELECTED": "選択済み", "ROW": "行", "ROWS": "行", "RECORD": "レコード", "RECORDS": "レコード", "DATA": "データ", "COMPLETE": "完了", "SUCCESSFUL": "成功", "SUCCESSFULLY": "成功", "FAILED": "失敗", "IMPORTED": "インポート済み", "EXPORTED": "エクスポート済み", "UPDATED": "更新済み", "ADDED": "追加済み", "DELETED": "削除済み", "LOADED": "読み込み済み", "FORM": "フォーム", "CLEARED": "クリア済み", "REQUIRED": "必須", "INVALID": "無効", "DRIVE": "ドライブ", "FOLDER": "フォルダー", "FOUND": "見つかりました", "NOT": "未", "OPEN": "開く", "EXISTING": "既存", "CREATE": "作成", "REBUILD": "再構築", "CATEGORY": "カテゴリ", "EMPTY": "空", "VALUE": "値", "VALUES": "値", "CURRENT": "現在", "WARNING": "警告", "COPIED": "コピー済み", "CLIPBOARD": "クリップボード", "CUSTOMERS": "顧客", "REMINDER": "リマインダー", "REMINDERS": "リマインダー", "SENT": "送信済み", "PERIOD": "期間", "GENERATED": "生成済み", "SUMMARY": "概要", "TOTAL": "合計", "SERVICES": "サービス", "REVENUE": "収益", "AVG": "平均", "MONTH": "月", "ACTIVE": "有効", "EXPIRED": "期限切れ", "EXPIRING": "期限間近", "SOON": "まもなく", "WAITING": "待機中", "PROGRESS": "進行中", "PENDING": "保留", "LOW": "低", "STOCK": "在庫", "YES": "はい", "NO": "いいえ", "APPLICATION": "アプリケーション", "SETUP": "セットアップ"},
    "ko_KR": {"SELECT": "선택", "SELECTED": "선택됨", "ROW": "행", "ROWS": "행", "RECORD": "레코드", "RECORDS": "레코드", "DATA": "데이터", "COMPLETE": "완료", "SUCCESSFUL": "성공", "SUCCESSFULLY": "성공적으로", "FAILED": "실패", "IMPORTED": "가져옴", "EXPORTED": "내보냄", "UPDATED": "업데이트됨", "ADDED": "추가됨", "DELETED": "삭제됨", "LOADED": "로드됨", "FORM": "양식", "CLEARED": "지워짐", "REQUIRED": "필수", "INVALID": "잘못됨", "DRIVE": "드라이브", "FOLDER": "폴더", "FOUND": "찾음", "NOT": "아님", "OPEN": "열기", "EXISTING": "기존", "CREATE": "만들기", "REBUILD": "다시 만들기", "CATEGORY": "범주", "EMPTY": "비어 있음", "VALUE": "값", "VALUES": "값", "CURRENT": "현재", "WARNING": "경고", "COPIED": "복사됨", "CLIPBOARD": "클립보드", "CUSTOMERS": "고객", "REMINDER": "알림", "REMINDERS": "알림", "SENT": "전송됨", "PERIOD": "기간", "GENERATED": "생성됨", "SUMMARY": "요약", "TOTAL": "합계", "SERVICES": "서비스", "REVENUE": "매출", "AVG": "평균", "MONTH": "월", "ACTIVE": "활성", "EXPIRED": "만료됨", "EXPIRING": "만료 예정", "SOON": "곧", "WAITING": "대기", "PROGRESS": "진행", "PENDING": "보류", "LOW": "낮음", "STOCK": "재고", "YES": "예", "NO": "아니요", "APPLICATION": "애플리케이션", "SETUP": "설정"},
}

for _code, _items in {**EXTRA_UI_TRANSLATIONS, **_GENERIC_EXTRA}.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)

_MORE_GENERIC_EXTRA = {
    "ar_SA": {"SELECT": "اختيار", "SELECTED": "محدد", "ROW": "صف", "ROWS": "صفوف", "RECORD": "سجل", "RECORDS": "سجلات", "DATA": "بيانات", "COMPLETE": "مكتمل", "SUCCESSFUL": "ناجح", "SUCCESSFULLY": "بنجاح", "FAILED": "فشل", "IMPORTED": "تم الاستيراد", "EXPORTED": "تم التصدير", "UPDATED": "تم التحديث", "ADDED": "تمت الإضافة", "DELETED": "تم الحذف", "LOADED": "تم التحميل", "FORM": "نموذج", "CLEARED": "تم المسح", "REQUIRED": "مطلوب", "INVALID": "غير صالح", "DRIVE": "محرك", "FOLDER": "مجلد", "FOUND": "موجود", "NOT": "ليس", "OPEN": "فتح", "EXISTING": "موجود", "CREATE": "إنشاء", "REBUILD": "إعادة بناء", "CATEGORY": "فئة", "EMPTY": "فارغ", "VALUE": "قيمة", "VALUES": "قيم", "CURRENT": "حالي", "WARNING": "تحذير", "COPIED": "تم النسخ", "CLIPBOARD": "الحافظة", "CUSTOMERS": "عملاء", "REMINDER": "تذكير", "REMINDERS": "تذكيرات", "SENT": "مرسل", "PERIOD": "الفترة", "GENERATED": "تم الإنشاء", "SUMMARY": "ملخص", "TOTAL": "الإجمالي", "SERVICES": "خدمات", "REVENUE": "الإيراد", "AVG": "متوسط", "MONTH": "شهر", "ACTIVE": "نشط", "EXPIRED": "منتهي", "EXPIRING": "ينتهي", "SOON": "قريبًا", "WAITING": "انتظار", "PROGRESS": "تقدم", "PENDING": "معلق", "LOW": "منخفض", "STOCK": "مخزون", "YES": "نعم", "NO": "لا", "APPLICATION": "تطبيق", "SETUP": "إعداد"},
    "hi_IN": {"SELECT": "चुनें", "SELECTED": "चयनित", "ROW": "पंक्ति", "ROWS": "पंक्तियाँ", "RECORD": "रिकॉर्ड", "RECORDS": "रिकॉर्ड", "DATA": "डेटा", "COMPLETE": "पूर्ण", "SUCCESSFUL": "सफल", "SUCCESSFULLY": "सफलतापूर्वक", "FAILED": "विफल", "IMPORTED": "आयातित", "EXPORTED": "निर्यातित", "UPDATED": "अपडेट किया गया", "ADDED": "जोड़ा गया", "DELETED": "हटाया गया", "LOADED": "लोड किया गया", "FORM": "फ़ॉर्म", "CLEARED": "साफ़ किया गया", "REQUIRED": "आवश्यक", "INVALID": "अमान्य", "DRIVE": "ड्राइव", "FOLDER": "फ़ोल्डर", "FOUND": "मिला", "NOT": "नहीं", "OPEN": "खोलें", "EXISTING": "मौजूदा", "CREATE": "बनाएं", "REBUILD": "फिर बनाएं", "CATEGORY": "श्रेणी", "EMPTY": "खाली", "VALUE": "मान", "VALUES": "मान", "CURRENT": "वर्तमान", "WARNING": "चेतावनी", "COPIED": "कॉपी किया गया", "CLIPBOARD": "क्लिपबोर्ड", "CUSTOMERS": "ग्राहक", "REMINDER": "अनुस्मारक", "REMINDERS": "अनुस्मारक", "SENT": "भेजा गया", "PERIOD": "अवधि", "GENERATED": "जनरेट किया गया", "SUMMARY": "सारांश", "TOTAL": "कुल", "SERVICES": "सेवाएँ", "REVENUE": "राजस्व", "AVG": "औसत", "MONTH": "माह", "ACTIVE": "सक्रिय", "EXPIRED": "समाप्त", "EXPIRING": "समाप्त हो रहा", "SOON": "जल्द", "WAITING": "प्रतीक्षा", "PROGRESS": "प्रगति", "PENDING": "लंबित", "LOW": "कम", "STOCK": "स्टॉक", "YES": "हाँ", "NO": "नहीं", "APPLICATION": "एप्लिकेशन", "SETUP": "सेटअप"},
    "bn_BD": {"SELECT": "নির্বাচন", "SELECTED": "নির্বাচিত", "ROW": "সারি", "ROWS": "সারি", "RECORD": "রেকর্ড", "RECORDS": "রেকর্ড", "DATA": "ডেটা", "COMPLETE": "সম্পূর্ণ", "SUCCESSFUL": "সফল", "SUCCESSFULLY": "সফলভাবে", "FAILED": "ব্যর্থ", "IMPORTED": "আমদানি হয়েছে", "EXPORTED": "রপ্তানি হয়েছে", "UPDATED": "আপডেট হয়েছে", "ADDED": "যোগ হয়েছে", "DELETED": "মুছে গেছে", "LOADED": "লোড হয়েছে", "FORM": "ফর্ম", "CLEARED": "মুছে ফেলা", "REQUIRED": "আবশ্যক", "INVALID": "অবৈধ", "DRIVE": "ড্রাইভ", "FOLDER": "ফোল্ডার", "FOUND": "পাওয়া গেছে", "NOT": "নয়", "OPEN": "খুলুন", "EXISTING": "বিদ্যমান", "CREATE": "তৈরি", "REBUILD": "পুনর্নির্মাণ", "CATEGORY": "বিভাগ", "EMPTY": "খালি", "VALUE": "মান", "VALUES": "মান", "CURRENT": "বর্তমান", "WARNING": "সতর্কতা", "COPIED": "কপি হয়েছে", "CLIPBOARD": "ক্লিপবোর্ড", "CUSTOMERS": "গ্রাহক", "REMINDER": "রিমাইন্ডার", "REMINDERS": "রিমাইন্ডার", "SENT": "পাঠানো হয়েছে", "PERIOD": "সময়কাল", "GENERATED": "তৈরি হয়েছে", "SUMMARY": "সারাংশ", "TOTAL": "মোট", "SERVICES": "সেবা", "REVENUE": "আয়", "AVG": "গড়", "MONTH": "মাস", "ACTIVE": "সক্রিয়", "EXPIRED": "মেয়াদোত্তীর্ণ", "EXPIRING": "মেয়াদ শেষ হচ্ছে", "SOON": "শীঘ্রই", "WAITING": "অপেক্ষা", "PROGRESS": "অগ্রগতি", "PENDING": "মুলতুবি", "LOW": "কম", "STOCK": "স্টক", "YES": "হ্যাঁ", "NO": "না", "APPLICATION": "অ্যাপ্লিকেশন", "SETUP": "সেটআপ"},
    "ru_RU": {"SELECT": "ВЫБРАТЬ", "SELECTED": "ВЫБРАНО", "ROW": "СТРОКА", "ROWS": "СТРОКИ", "RECORD": "ЗАПИСЬ", "RECORDS": "ЗАПИСИ", "DATA": "ДАННЫЕ", "COMPLETE": "ЗАВЕРШЕНО", "SUCCESSFUL": "УСПЕШНО", "SUCCESSFULLY": "УСПЕШНО", "FAILED": "СБОЙ", "IMPORTED": "ИМПОРТИРОВАНО", "EXPORTED": "ЭКСПОРТИРОВАНО", "UPDATED": "ОБНОВЛЕНО", "ADDED": "ДОБАВЛЕНО", "DELETED": "УДАЛЕНО", "LOADED": "ЗАГРУЖЕНО", "FORM": "ФОРМА", "CLEARED": "ОЧИЩЕНО", "REQUIRED": "ОБЯЗАТЕЛЬНО", "INVALID": "НЕДОПУСТИМО", "DRIVE": "ДИСК", "FOLDER": "ПАПКА", "FOUND": "НАЙДЕНО", "NOT": "НЕ", "OPEN": "ОТКРЫТЬ", "EXISTING": "СУЩЕСТВУЮЩИЙ", "CREATE": "СОЗДАТЬ", "REBUILD": "ПЕРЕСОЗДАТЬ", "CATEGORY": "КАТЕГОРИЯ", "EMPTY": "ПУСТО", "VALUE": "ЗНАЧЕНИЕ", "VALUES": "ЗНАЧЕНИЯ", "CURRENT": "ТЕКУЩИЙ", "WARNING": "ПРЕДУПРЕЖДЕНИЕ", "COPIED": "СКОПИРОВАНО", "CLIPBOARD": "БУФЕР ОБМЕНА", "CUSTOMERS": "КЛИЕНТЫ", "REMINDER": "НАПОМИНАНИЕ", "REMINDERS": "НАПОМИНАНИЯ", "SENT": "ОТПРАВЛЕНО", "PERIOD": "ПЕРИОД", "GENERATED": "СОЗДАНО", "SUMMARY": "СВОДКА", "TOTAL": "ИТОГО", "SERVICES": "УСЛУГИ", "REVENUE": "ВЫРУЧКА", "AVG": "СРЕДНЕЕ", "MONTH": "МЕСЯЦ", "ACTIVE": "АКТИВНО", "EXPIRED": "ИСТЕКЛО", "EXPIRING": "ИСТЕКАЕТ", "SOON": "СКОРО", "WAITING": "ОЖИДАНИЕ", "PROGRESS": "ПРОГРЕСС", "PENDING": "ОЖИДАЕТ", "LOW": "НИЗКИЙ", "STOCK": "ЗАПАС", "YES": "ДА", "NO": "НЕТ", "APPLICATION": "ПРИЛОЖЕНИЕ", "SETUP": "НАСТРОЙКА"},
    "tr_TR": {"SELECT": "SEÇ", "SELECTED": "SEÇİLİ", "ROW": "SATIR", "ROWS": "SATIR", "RECORD": "KAYIT", "RECORDS": "KAYITLAR", "DATA": "VERİ", "COMPLETE": "TAMAM", "SUCCESSFUL": "BAŞARILI", "SUCCESSFULLY": "BAŞARIYLA", "FAILED": "BAŞARISIZ", "IMPORTED": "İÇE AKTARILDI", "EXPORTED": "DIŞA AKTARILDI", "UPDATED": "GÜNCELLENDİ", "ADDED": "EKLENDİ", "DELETED": "SİLİNDİ", "LOADED": "YÜKLENDİ", "FORM": "FORM", "CLEARED": "TEMİZLENDİ", "REQUIRED": "GEREKLİ", "INVALID": "GEÇERSİZ", "DRIVE": "SÜRÜCÜ", "FOLDER": "KLASÖR", "FOUND": "BULUNDU", "NOT": "DEĞİL", "OPEN": "AÇ", "EXISTING": "MEVCUT", "CREATE": "OLUŞTUR", "REBUILD": "YENİDEN OLUŞTUR", "CATEGORY": "KATEGORİ", "EMPTY": "BOŞ", "VALUE": "DEĞER", "VALUES": "DEĞERLER", "CURRENT": "GEÇERLİ", "WARNING": "UYARI", "COPIED": "KOPYALANDI", "CLIPBOARD": "PANO", "CUSTOMERS": "MÜŞTERİLER", "REMINDER": "HATIRLATICI", "REMINDERS": "HATIRLATICILAR", "SENT": "GÖNDERİLDİ", "PERIOD": "DÖNEM", "GENERATED": "OLUŞTURULDU", "SUMMARY": "ÖZET", "TOTAL": "TOPLAM", "SERVICES": "SERVİSLER", "REVENUE": "GELİR", "AVG": "ORTALAMA", "MONTH": "AY", "ACTIVE": "AKTİF", "EXPIRED": "SÜRESİ DOLDU", "EXPIRING": "SÜRESİ DOLUYOR", "SOON": "YAKINDA", "WAITING": "BEKLİYOR", "PROGRESS": "İLERLEME", "PENDING": "BEKLEMEDE", "LOW": "DÜŞÜK", "STOCK": "STOK", "YES": "EVET", "NO": "HAYIR", "APPLICATION": "UYGULAMA", "SETUP": "KURULUM"},
    "vi_VN": {"SELECT": "CHỌN", "SELECTED": "ĐÃ CHỌN", "ROW": "DÒNG", "ROWS": "DÒNG", "RECORD": "BẢN GHI", "RECORDS": "BẢN GHI", "DATA": "DỮ LIỆU", "COMPLETE": "HOÀN TẤT", "SUCCESSFUL": "THÀNH CÔNG", "SUCCESSFULLY": "THÀNH CÔNG", "FAILED": "THẤT BẠI", "IMPORTED": "ĐÃ NHẬP", "EXPORTED": "ĐÃ XUẤT", "UPDATED": "ĐÃ CẬP NHẬT", "ADDED": "ĐÃ THÊM", "DELETED": "ĐÃ XÓA", "LOADED": "ĐÃ TẢI", "FORM": "BIỂU MẪU", "CLEARED": "ĐÃ XÓA", "REQUIRED": "BẮT BUỘC", "INVALID": "KHÔNG HỢP LỆ", "DRIVE": "Ổ ĐĨA", "FOLDER": "THƯ MỤC", "FOUND": "TÌM THẤY", "NOT": "KHÔNG", "OPEN": "MỞ", "EXISTING": "HIỆN CÓ", "CREATE": "TẠO", "REBUILD": "TẠO LẠI", "CATEGORY": "DANH MỤC", "EMPTY": "TRỐNG", "VALUE": "GIÁ TRỊ", "VALUES": "GIÁ TRỊ", "CURRENT": "HIỆN TẠI", "WARNING": "CẢNH BÁO", "COPIED": "ĐÃ SAO CHÉP", "CLIPBOARD": "CLIPBOARD", "CUSTOMERS": "KHÁCH HÀNG", "REMINDER": "NHẮC NHỞ", "REMINDERS": "NHẮC NHỞ", "SENT": "ĐÃ GỬI", "PERIOD": "KỲ", "GENERATED": "ĐÃ TẠO", "SUMMARY": "TÓM TẮT", "TOTAL": "TỔNG", "SERVICES": "DỊCH VỤ", "REVENUE": "DOANH THU", "AVG": "TRUNG BÌNH", "MONTH": "THÁNG", "ACTIVE": "ĐANG HOẠT ĐỘNG", "EXPIRED": "HẾT HẠN", "EXPIRING": "SẮP HẾT HẠN", "SOON": "SỚM", "WAITING": "ĐANG CHỜ", "PROGRESS": "TIẾN ĐỘ", "PENDING": "ĐANG CHỜ", "LOW": "THẤP", "STOCK": "TỒN KHO", "YES": "CÓ", "NO": "KHÔNG", "APPLICATION": "ỨNG DỤNG", "SETUP": "THIẾT LẬP"},
    "th_TH": {"SELECT": "เลือก", "SELECTED": "เลือกแล้ว", "ROW": "แถว", "ROWS": "แถว", "RECORD": "ระเบียน", "RECORDS": "ระเบียน", "DATA": "ข้อมูล", "COMPLETE": "เสร็จสิ้น", "SUCCESSFUL": "สำเร็จ", "SUCCESSFULLY": "สำเร็จ", "FAILED": "ล้มเหลว", "IMPORTED": "นำเข้าแล้ว", "EXPORTED": "ส่งออกแล้ว", "UPDATED": "อัปเดตแล้ว", "ADDED": "เพิ่มแล้ว", "DELETED": "ลบแล้ว", "LOADED": "โหลดแล้ว", "FORM": "ฟอร์ม", "CLEARED": "ล้างแล้ว", "REQUIRED": "จำเป็น", "INVALID": "ไม่ถูกต้อง", "DRIVE": "ไดรฟ์", "FOLDER": "โฟลเดอร์", "FOUND": "พบ", "NOT": "ไม่", "OPEN": "เปิด", "EXISTING": "ที่มีอยู่", "CREATE": "สร้าง", "REBUILD": "สร้างใหม่", "CATEGORY": "หมวดหมู่", "EMPTY": "ว่าง", "VALUE": "ค่า", "VALUES": "ค่า", "CURRENT": "ปัจจุบัน", "WARNING": "คำเตือน", "COPIED": "คัดลอกแล้ว", "CLIPBOARD": "คลิปบอร์ด", "CUSTOMERS": "ลูกค้า", "REMINDER": "การเตือน", "REMINDERS": "การเตือน", "SENT": "ส่งแล้ว", "PERIOD": "ช่วงเวลา", "GENERATED": "สร้างแล้ว", "SUMMARY": "สรุป", "TOTAL": "รวม", "SERVICES": "บริการ", "REVENUE": "รายได้", "AVG": "เฉลี่ย", "MONTH": "เดือน", "ACTIVE": "ใช้งาน", "EXPIRED": "หมดอายุ", "EXPIRING": "ใกล้หมดอายุ", "SOON": "เร็วๆ นี้", "WAITING": "รอ", "PROGRESS": "ความคืบหน้า", "PENDING": "รอดำเนินการ", "LOW": "ต่ำ", "STOCK": "สต็อก", "YES": "ใช่", "NO": "ไม่", "APPLICATION": "แอปพลิเคชัน", "SETUP": "ตั้งค่า"},
    "ms_MY": {"SELECT": "PILIH", "SELECTED": "DIPILIH", "ROW": "BARIS", "ROWS": "BARIS", "RECORD": "REKOD", "RECORDS": "REKOD", "DATA": "DATA", "COMPLETE": "SELESAI", "SUCCESSFUL": "BERJAYA", "SUCCESSFULLY": "BERJAYA", "FAILED": "GAGAL", "IMPORTED": "DIIMPORT", "EXPORTED": "DIEKSPORT", "UPDATED": "DIKEMAS KINI", "ADDED": "DITAMBAH", "DELETED": "DIPADAM", "LOADED": "DIMUAT", "FORM": "BORANG", "CLEARED": "DIKOSONGKAN", "REQUIRED": "DIPERLUKAN", "INVALID": "TIDAK SAH", "DRIVE": "PEMACU", "FOLDER": "FOLDER", "FOUND": "DITEMUI", "NOT": "TIDAK", "OPEN": "BUKA", "EXISTING": "SEDIA ADA", "CREATE": "CIPTA", "REBUILD": "BINA SEMULA", "CATEGORY": "KATEGORI", "EMPTY": "KOSONG", "VALUE": "NILAI", "VALUES": "NILAI", "CURRENT": "SEMASA", "WARNING": "AMARAN", "COPIED": "DISALIN", "CLIPBOARD": "PAPAN KLIP", "CUSTOMERS": "PELANGGAN", "REMINDER": "PERINGATAN", "REMINDERS": "PERINGATAN", "SENT": "DIHANTAR", "PERIOD": "TEMPOH", "GENERATED": "DIJANA", "SUMMARY": "RINGKASAN", "TOTAL": "JUMLAH", "SERVICES": "SERVIS", "REVENUE": "HASIL", "AVG": "PURATA", "MONTH": "BULAN", "ACTIVE": "AKTIF", "EXPIRED": "TAMAT TEMPOH", "EXPIRING": "AKAN TAMAT", "SOON": "SEGERA", "WAITING": "MENUNGGU", "PROGRESS": "KEMAJUAN", "PENDING": "TERTANGGUH", "LOW": "RENDAH", "STOCK": "STOK", "YES": "YA", "NO": "TIDAK", "APPLICATION": "APLIKASI", "SETUP": "TETAPAN"},
    "fil_PH": {"SELECT": "PILIIN", "SELECTED": "NAPILI", "ROW": "HANAY", "ROWS": "HANAY", "RECORD": "RECORD", "RECORDS": "RECORDS", "DATA": "DATA", "COMPLETE": "TAPOS", "SUCCESSFUL": "MATAGUMPAY", "SUCCESSFULLY": "MATAGUMPAY", "FAILED": "NABIGO", "IMPORTED": "NA-IMPORT", "EXPORTED": "NA-EXPORT", "UPDATED": "NA-UPDATE", "ADDED": "NAIDAGDAG", "DELETED": "NABURA", "LOADED": "NA-LOAD", "FORM": "FORM", "CLEARED": "NALINIS", "REQUIRED": "KINAKAILANGAN", "INVALID": "HINDI WASTO", "DRIVE": "DRIVE", "FOLDER": "FOLDER", "FOUND": "NATAGPUAN", "NOT": "HINDI", "OPEN": "BUKSAN", "EXISTING": "UMIIRAL", "CREATE": "GUMAWA", "REBUILD": "GAWING MULI", "CATEGORY": "KATEGORYA", "EMPTY": "WALANG LAMAN", "VALUE": "HALAGA", "VALUES": "MGA HALAGA", "CURRENT": "KASALUKUYAN", "WARNING": "BABALA", "COPIED": "NAKOPYA", "CLIPBOARD": "CLIPBOARD", "CUSTOMERS": "CUSTOMERS", "REMINDER": "PAALALA", "REMINDERS": "MGA PAALALA", "SENT": "NAIPADALA", "PERIOD": "PANAHON", "GENERATED": "NAGAWA", "SUMMARY": "BUOD", "TOTAL": "KABUUAN", "SERVICES": "SERBISYO", "REVENUE": "KITA", "AVG": "AVERAGE", "MONTH": "BUWAN", "ACTIVE": "AKTIBO", "EXPIRED": "EXPIRED", "EXPIRING": "MAG-EEXPIRE", "SOON": "MALAPIT", "WAITING": "NAGHIHINTAY", "PROGRESS": "PROGRESO", "PENDING": "NAKABINBIN", "LOW": "MABABA", "STOCK": "STOCK", "YES": "OO", "NO": "HINDI", "APPLICATION": "APPLICATION", "SETUP": "SETUP"},
}

for _code, _items in _MORE_GENERIC_EXTRA.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)

_SELECTION_TRANSLATIONS = {
    "id_ID": {"SELECTION": "PILIHAN", "NO SELECTION": "TIDAK ADA PILIHAN"},
    "es_MX": {"SELECTION": "SELECCIÓN", "NO SELECTION": "SIN SELECCIÓN"},
    "fr_FR": {"SELECTION": "SÉLECTION", "NO SELECTION": "AUCUNE SÉLECTION"},
    "de_DE": {"SELECTION": "AUSWAHL", "NO SELECTION": "KEINE AUSWAHL"},
    "pt_BR": {"SELECTION": "SELEÇÃO", "NO SELECTION": "NENHUMA SELEÇÃO"},
    "it_IT": {"SELECTION": "SELEZIONE", "NO SELECTION": "NESSUNA SELEZIONE"},
    "nl_NL": {"SELECTION": "SELECTIE", "NO SELECTION": "GEEN SELECTIE"},
    "zh_CN": {"SELECTION": "选择", "NO SELECTION": "未选择"},
    "ja_JP": {"SELECTION": "選択", "NO SELECTION": "未選択"},
    "ko_KR": {"SELECTION": "선택", "NO SELECTION": "선택 없음"},
    "ar_SA": {"SELECTION": "تحديد", "NO SELECTION": "لا يوجد تحديد"},
    "hi_IN": {"SELECTION": "चयन", "NO SELECTION": "कोई चयन नहीं"},
    "bn_BD": {"SELECTION": "নির্বাচন", "NO SELECTION": "কোনো নির্বাচন নেই"},
    "ru_RU": {"SELECTION": "ВЫБОР", "NO SELECTION": "НЕТ ВЫБОРА"},
    "tr_TR": {"SELECTION": "SEÇİM", "NO SELECTION": "SEÇİM YOK"},
    "vi_VN": {"SELECTION": "LỰA CHỌN", "NO SELECTION": "CHƯA CHỌN"},
    "th_TH": {"SELECTION": "การเลือก", "NO SELECTION": "ไม่มีการเลือก"},
    "ms_MY": {"SELECTION": "PILIHAN", "NO SELECTION": "TIADA PILIHAN"},
    "fil_PH": {"SELECTION": "SELECTION", "NO SELECTION": "WALANG NAPILI"},
}

for _code, _items in _SELECTION_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)

_COPY_TRANSLATIONS = {
    "id_ID": "SALIN", "es_MX": "COPIAR", "fr_FR": "COPIER", "de_DE": "KOPIEREN",
    "pt_BR": "COPIAR", "it_IT": "COPIA", "nl_NL": "KOPIËREN", "zh_CN": "复制",
    "ja_JP": "コピー", "ko_KR": "복사", "ar_SA": "نسخ", "hi_IN": "कॉपी",
    "bn_BD": "কপি", "ru_RU": "КОПИРОВАТЬ", "tr_TR": "KOPYALA", "vi_VN": "SAO CHÉP",
    "th_TH": "คัดลอก", "ms_MY": "SALIN", "fil_PH": "KOPYAHIN",
}

for _code, _copy_text in _COPY_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update({"COPY": _copy_text})


# Supplemental full-UI localization added for all 20 locale options.
# Covers scattered labels, buttons, hints, dummy-data dialogs, status text, and sentence fragments.
_FULL_UI_SUPPLEMENTAL_TRANSLATIONS = {'ar_SA': {'A': '',
           'ACTION': 'إجراء',
           'ALL': 'كل',
           'ALL FIELDS': 'كل حقول',
           'AN': '',
           'ANOTHER': 'آخر',
           'ARE': 'هي',
           'AS': 'كما هو',
           'AVAILABLE': 'متاح',
           'BACKUP': 'نسخة احتياطية',
           'CHANGES': 'التغييرات',
           'CLICK': 'انقر',
           'COMPUTER': 'كمبيوتر',
           'CONTACT': 'جهة الاتصال',
           'CONTINUE': 'متابعة',
           'CREATE': 'إنشاء',
           'CSV': 'CSV',
           'CURRENT': 'الحالي',
           'Computer Service Manager 3.1': 'كمبيوتر خدمة مدير 3.1',
           'DATA': 'بيانات',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'هل',
           'DRIVE': 'محرك',
           'DUMMY': 'تجريبي',
           'ENTER': 'أدخل',
           'EXCEL': 'إكسل',
           'EXISTING': 'موجود',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'حقول',
           'FILE': 'ملف',
           'FILES': 'ملفات',
           'FOLDER': 'مجلد',
           'FOR': 'لـ',
           'FORM': 'نموذج',
           'FORMAT': 'تنسيق',
           'FROM': 'من',
           'IMMEDIATELY': 'فورًا',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'إدخال',
           'IS': 'هو',
           'ITEM': 'عنصر',
           'ITEMS': 'عناصر',
           'KIND': 'نوع',
           'LOCATION': 'موقع',
           'MANAGER': 'مدير',
           'MESSAGE': 'رسالة',
           'MODE': 'وضع',
           'MULTIPLE': 'متعددة',
           'NEW': 'جديد',
           'NO': 'لا',
           'OPEN': 'فتح',
           'PART': 'قطعة',
           'PATHS': 'مسارات',
           'PLEASE': 'يرجى',
           'PROVIDED': 'مقدم',
           'REPLACE': 'استبدال',
           'REQUIRED': 'مطلوب',
           'RESERVED': 'محفوظة',
           'RESTORED': 'تمت الاستعادة',
           'RESTORING': 'استعادة',
           'RIGHTS': 'حقوق',
           'ROW': 'صف',
           'ROWS': 'صفوف',
           'SAVED': 'محفوظ',
           'SELECT': 'اختر',
           'SEND': 'إرسال',
           'SEPARATE': 'افصل',
           'SERVICE': 'خدمة',
           'SERVICES': 'خدمات',
           'SOFTWARE': 'برنامج',
           'SUCCESSFULLY': 'بنجاح',
           'SUPPLIER': 'مورد',
           'THE': '',
           'THEN': 'ثم',
           'THIS': 'هذا',
           'TYPE': 'اكتب',
           'TYPE NEW VALUE': 'اكتب جديد VALUE',
           'UNDONE': 'تراجع',
           'UNSUPPORTED': 'غير مدعوم',
           'UPDATE': 'تحديث',
           'USE': 'استخدم',
           'USING': 'باستخدام',
           'WANT': 'تريد',
           'WARRANTIES': 'ضمانات',
           'WAS': 'كان',
           'WILL': 'سوف',
           'WITH': 'مع',
           'YES': 'نعم',
           'YOU': 'أنت'},
 'bn_BD': {'A': '',
           'ACTION': 'কর্ম',
           'ALL': 'সব',
           'ALL FIELDS': 'সব ক্ষেত্র',
           'AN': '',
           'ANOTHER': 'অন্য',
           'ARE': 'হয়',
           'AS': 'যেমন আছে',
           'AVAILABLE': 'উপলভ্য',
           'BACKUP': 'ব্যাকআপ',
           'CHANGES': 'পরিবর্তন',
           'CLICK': 'ক্লিক করুন',
           'COMPUTER': 'কম্পিউটার',
           'CONTACT': 'যোগাযোগ',
           'CONTINUE': 'চালিয়ে যান',
           'CREATE': 'তৈরি করুন',
           'CSV': 'CSV',
           'CURRENT': 'বর্তমান',
           'Computer Service Manager 3.1': 'কম্পিউটার সেবা ম্যানেজার 3.1',
           'DATA': 'ডেটা',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'আপনি কি',
           'DRIVE': 'ড্রাইভ',
           'DUMMY': 'ডামি',
           'ENTER': 'প্রবেশ করুন',
           'EXCEL': 'এক্সেল',
           'EXISTING': 'বিদ্যমান',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'ক্ষেত্র',
           'FILE': 'ফাইল',
           'FILES': 'ফাইল',
           'FOLDER': 'ফোল্ডার',
           'FOR': 'জন্য',
           'FORM': 'ফর্ম',
           'FORMAT': 'ফরম্যাট',
           'FROM': 'থেকে',
           'IMMEDIATELY': 'তাৎক্ষণিকভাবে',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'ইনপুট',
           'IS': 'হয়',
           'ITEM': 'আইটেম',
           'ITEMS': 'আইটেম',
           'KIND': 'ধরন',
           'LOCATION': 'অবস্থান',
           'MANAGER': 'ম্যানেজার',
           'MESSAGE': 'বার্তা',
           'MODE': 'মোড',
           'MULTIPLE': 'একাধিক',
           'NEW': 'নতুন',
           'NO': 'না',
           'OPEN': 'খুলুন',
           'PART': 'পার্ট',
           'PATHS': 'পথ',
           'PLEASE': 'অনুগ্রহ করে',
           'PROVIDED': 'প্রদত্ত',
           'REPLACE': 'প্রতিস্থাপন',
           'REQUIRED': 'আবশ্যক',
           'RESERVED': 'সংরক্ষিত',
           'RESTORED': 'পুনরুদ্ধার হয়েছে',
           'RESTORING': 'পুনরুদ্ধার',
           'RIGHTS': 'অধিকার',
           'ROW': 'সারি',
           'ROWS': 'সারি',
           'SAVED': 'সংরক্ষিত',
           'SELECT': 'নির্বাচন করুন',
           'SEND': 'পাঠান',
           'SEPARATE': 'আলাদা করুন',
           'SERVICE': 'সেবা',
           'SERVICES': 'সেবা',
           'SOFTWARE': 'সফটওয়্যার',
           'SUCCESSFULLY': 'সফলভাবে',
           'SUPPLIER': 'সরবরাহকারী',
           'THE': '',
           'THEN': 'তারপর',
           'THIS': 'এই',
           'TYPE': 'টাইপ করুন',
           'TYPE NEW VALUE': 'টাইপ করুন নতুন VALUE',
           'UNDONE': 'পূর্বাবস্থায়',
           'UNSUPPORTED': 'অসমর্থিত',
           'UPDATE': 'আপডেট',
           'USE': 'ব্যবহার করুন',
           'USING': 'ব্যবহার করে',
           'WANT': 'চান',
           'WARRANTIES': 'ওয়ারেন্টি',
           'WAS': 'ছিল',
           'WILL': 'করবে',
           'WITH': 'সহ',
           'YES': 'হ্যাঁ',
           'YOU': 'আপনি'},
 'de_DE': {'A': '',
           'ACTION': 'AKTION',
           'ALL': 'ALLE',
           'ALL FIELDS': 'ALLE FELDER',
           'AN': '',
           'ANOTHER': 'ANDEREN',
           'ARE': 'WERDEN',
           'AS': 'WIE',
           'AVAILABLE': 'VERFÜGBAR',
           'BACKUP': 'SICHERUNG',
           'CHANGES': 'ÄNDERUNGEN',
           'CLICK': 'KLICKEN',
           'COMPUTER': 'COMPUTER',
           'CONTACT': 'KONTAKT',
           'CONTINUE': 'FORTFAHREN',
           'CREATE': 'ERSTELLEN',
           'CSV': 'CSV',
           'CURRENT': 'AKTUELL',
           'Computer Service Manager 3.1': 'COMPUTER SERVICE MANAGER 3.1',
           'DATA': 'DATEN',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'MÖCHTEN',
           'DRIVE': 'LAUFWERK',
           'DUMMY': 'DUMMY',
           'ENTER': 'EINGEBEN',
           'EXCEL': 'EXCEL',
           'EXISTING': 'VORHANDEN',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'FELDER',
           'FILE': 'DATEI',
           'FILES': 'DATEIEN',
           'FOLDER': 'ORDNER',
           'FOR': 'FÜR',
           'FORM': 'FORMULAR',
           'FORMAT': 'FORMAT',
           'FROM': 'VON',
           'IMMEDIATELY': 'SOFORT',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'EINGABE',
           'IS': 'IST',
           'ITEM': 'ELEMENT',
           'ITEMS': 'ELEMENTE',
           'KIND': 'ART',
           'LOCATION': 'STANDORT',
           'MANAGER': 'MANAGER',
           'MESSAGE': 'NACHRICHT',
           'MODE': 'MODUS',
           'MULTIPLE': 'MEHRERE',
           'NEW': 'NEU',
           'NO': 'NEIN',
           'OPEN': 'ÖFFNEN',
           'PART': 'TEIL',
           'PATHS': 'PFADE',
           'PLEASE': 'BITTE',
           'PROVIDED': 'BEREITGESTELLT',
           'REPLACE': 'ERSETZEN',
           'REQUIRED': 'ERFORDERLICH',
           'RESERVED': 'VORBEHALTEN',
           'RESTORED': 'WIEDERHERGESTELLT',
           'RESTORING': 'WIEDERHERSTELLEN',
           'RIGHTS': 'RECHTE',
           'ROW': 'ZEILE',
           'ROWS': 'ZEILEN',
           'SAVED': 'GESPEICHERT',
           'SELECT': 'AUSWÄHLEN',
           'SEND': 'SENDEN',
           'SEPARATE': 'TRENNEN',
           'SERVICE': 'SERVICE',
           'SERVICES': 'SERVICES',
           'SOFTWARE': 'SOFTWARE',
           'SUCCESSFULLY': 'ERFOLGREICH',
           'SUPPLIER': 'LIEFERANT',
           'THE': '',
           'THEN': 'DANN',
           'THIS': 'DIESE',
           'TYPE': 'EINGEBEN',
           'TYPE NEW VALUE': 'EINGEBEN NEU VALUE',
           'UNDONE': 'RÜCKGÄNGIG',
           'UNSUPPORTED': 'NICHT UNTERSTÜTZT',
           'UPDATE': 'AKTUALISIEREN',
           'USE': 'VERWENDEN',
           'USING': 'MIT',
           'WANT': 'MÖCHTEN',
           'WARRANTIES': 'GARANTIEN',
           'WAS': 'WURDE',
           'WILL': 'WIRD',
           'WITH': 'MIT',
           'YES': 'JA',
           'YOU': 'SIE'},
 'es_MX': {'A': '',
           'ACTION': 'ACCIÓN',
           'ALL': 'TODO',
           'ALL FIELDS': 'TODO CAMPOS',
           'AN': '',
           'ANOTHER': 'OTRA',
           'ARE': 'SE',
           'AS': 'COMO',
           'AVAILABLE': 'DISPONIBLE',
           'BACKUP': 'COPIA',
           'CHANGES': 'CAMBIOS',
           'CLICK': 'HAGA CLIC',
           'COMPUTER': 'COMPUTADORA',
           'CONTACT': 'CONTACTO',
           'CONTINUE': 'CONTINUAR',
           'CREATE': 'CREAR',
           'CSV': 'CSV',
           'CURRENT': 'ACTUAL',
           'Computer Service Manager 3.1': 'COMPUTADORA SERVICIO GESTOR 3.1',
           'DATA': 'DATOS',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': '¿',
           'DRIVE': 'UNIDAD',
           'DUMMY': 'DUMMY',
           'ENTER': 'INTRODUZCA',
           'EXCEL': 'EXCEL',
           'EXISTING': 'EXISTENTE',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'CAMPOS',
           'FILE': 'ARCHIVO',
           'FILES': 'ARCHIVOS',
           'FOLDER': 'CARPETA',
           'FOR': 'PARA',
           'FORM': 'FORMULARIO',
           'FORMAT': 'FORMATO',
           'FROM': 'DESDE',
           'IMMEDIATELY': 'INMEDIATAMENTE',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'ENTRADA',
           'IS': 'ES',
           'ITEM': 'ELEMENTO',
           'ITEMS': 'ELEMENTOS',
           'KIND': 'TIPO',
           'LOCATION': 'UBICACIÓN',
           'MANAGER': 'GESTOR',
           'MESSAGE': 'MENSAJE',
           'MODE': 'MODO',
           'MULTIPLE': 'MÚLTIPLES',
           'NEW': 'NUEVO',
           'NO': 'NO',
           'OPEN': 'ABRIR',
           'PART': 'PIEZA',
           'PATHS': 'RUTAS',
           'PLEASE': 'POR FAVOR',
           'PROVIDED': 'PROPORCIONADA',
           'REPLACE': 'REEMPLAZAR',
           'REQUIRED': 'OBLIGATORIO',
           'RESERVED': 'RESERVADOS',
           'RESTORED': 'RESTAURADA',
           'RESTORING': 'RESTAURAR',
           'RIGHTS': 'DERECHOS',
           'ROW': 'FILA',
           'ROWS': 'FILAS',
           'SAVED': 'GUARDADO',
           'SELECT': 'SELECCIONE',
           'SEND': 'ENVIAR',
           'SEPARATE': 'SEPARAR',
           'SERVICE': 'SERVICIO',
           'SERVICES': 'SERVICIOS',
           'SOFTWARE': 'SOFTWARE',
           'SUCCESSFULLY': 'CORRECTAMENTE',
           'SUPPLIER': 'PROVEEDOR',
           'THE': '',
           'THEN': 'LUEGO',
           'THIS': 'ESTA',
           'TYPE': 'ESCRIBA',
           'TYPE NEW VALUE': 'ESCRIBA NUEVO VALUE',
           'UNDONE': 'DESHACER',
           'UNSUPPORTED': 'NO COMPATIBLE',
           'UPDATE': 'ACTUALIZAR',
           'USE': 'USAR',
           'USING': 'USANDO',
           'WANT': 'QUIERE',
           'WARRANTIES': 'GARANTÍAS',
           'WAS': 'FUE',
           'WILL': 'VA A',
           'WITH': 'CON',
           'YES': 'SÍ',
           'YOU': 'USTED'},
 'fil_PH': {'ALL': 'LAHAT',
            'ALL FIELDS': 'LAHAT FIELD',
            'AS': 'AS',
            'AVAILABLE': 'AVAILABLE',
            'CHANGES': 'MGA PAGBABAGO',
            'CLICK': 'I-CLICK',
            'COMPUTER': 'COMPUTER',
            'CSV': 'CSV',
            'Computer Service Manager 3.1': 'COMPUTER SERBISYO MANAGER 3.1',
            'DELETE SELECTED': 'DELETE SELECTED',
            'DUMMY': 'DUMMY',
            'ENTER': 'ILAGAY',
            'EXCEL': 'EXCEL',
            'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
            'FIELDS': 'FIELD',
            'FILE': 'FILE',
            'FILES': 'FILES',
            'FOR': 'PARA SA',
            'FORMAT': 'FORMAT',
            'IMMEDIATELY': 'AGAD',
            'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
            'INF': 'INF',
            'INPUT': 'INPUT',
            'IS': 'AY',
            'ITEMS': 'ITEMS',
            'KIND': 'URI',
            'LOCATION': 'LOKASYON',
            'MANAGER': 'MANAGER',
            'MESSAGE': 'MENSAHE',
            'MODE': 'MODE',
            'MULTIPLE': 'MARAMI',
            'NEW': 'BAGO',
            'PATHS': 'PATHS',
            'PLEASE': 'PAKI',
            'PROVIDED': 'IBINIGAY',
            'REQUIRED': 'KAILANGAN',
            'RESERVED': 'NAKALAAN',
            'RESTORED': 'NA-RESTORE',
            'RIGHTS': 'KARAPATAN',
            'SAVED': 'NA-SAVE',
            'SEPARATE': 'PAGHIWALAYIN',
            'SERVICE': 'SERBISYO',
            'SOFTWARE': 'SOFTWARE',
            'THEN': 'PAGKATAPOS',
            'THIS': 'ITO',
            'TYPE': 'I-TYPE',
            'TYPE NEW VALUE': 'I-TYPE BAGO VALUE',
            'UNSUPPORTED': 'HINDI SUPORTADO',
            'UPDATE': 'I-UPDATE',
            'USE': 'GAMITIN',
            'USING': 'GAMIT',
            'WANT': 'GUSTO',
            'WARRANTIES': 'WARRANTIES',
            'WAS': 'AY',
            'WITH': 'MAY',
            'YOU': 'IKAW'},
 'fr_FR': {'A': '',
           'ACTION': 'ACTION',
           'ALL': 'TOUS',
           'ALL FIELDS': 'TOUS CHAMPS',
           'AN': '',
           'ANOTHER': 'AUTRE',
           'ARE': 'SONT',
           'AS': 'TEL QUEL',
           'AVAILABLE': 'DISPONIBLE',
           'BACKUP': 'SAUVEGARDE',
           'CHANGES': 'MODIFICATIONS',
           'CLICK': 'CLIQUER',
           'COMPUTER': 'ORDINATEUR',
           'CONTACT': 'CONTACT',
           'CONTINUE': 'CONTINUER',
           'CREATE': 'CRÉER',
           'CSV': 'CSV',
           'CURRENT': 'ACTUEL',
           'Computer Service Manager 3.1': 'ORDINATEUR SERVICE GESTIONNAIRE 3.1',
           'DATA': 'DONNÉES',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'VOULEZ',
           'DRIVE': 'LECTEUR',
           'DUMMY': 'DUMMY',
           'ENTER': 'SAISIR',
           'EXCEL': 'EXCEL',
           'EXISTING': 'EXISTANT',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'CHAMPS',
           'FILE': 'FICHIER',
           'FILES': 'FICHIERS',
           'FOLDER': 'DOSSIER',
           'FOR': 'POUR',
           'FORM': 'FORMULAIRE',
           'FORMAT': 'FORMAT',
           'FROM': 'DE',
           'IMMEDIATELY': 'IMMÉDIATEMENT',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'SAISIE',
           'IS': 'EST',
           'ITEM': 'ÉLÉMENT',
           'ITEMS': 'ÉLÉMENTS',
           'KIND': 'TYPE',
           'LOCATION': 'EMPLACEMENT',
           'MANAGER': 'GESTIONNAIRE',
           'MESSAGE': 'MESSAGE',
           'MODE': 'MODE',
           'MULTIPLE': 'PLUSIEURS',
           'NEW': 'NOUVEAU',
           'NO': 'NON',
           'OPEN': 'OUVRIR',
           'PART': 'PIÈCE',
           'PATHS': 'CHEMINS',
           'PLEASE': 'VEUILLEZ',
           'PROVIDED': 'FOURNIE',
           'REPLACE': 'REMPLACER',
           'REQUIRED': 'REQUIS',
           'RESERVED': 'RÉSERVÉS',
           'RESTORED': 'RESTAURÉE',
           'RESTORING': 'RESTAURATION',
           'RIGHTS': 'DROITS',
           'ROW': 'LIGNE',
           'ROWS': 'LIGNES',
           'SAVED': 'ENREGISTRÉ',
           'SELECT': 'SÉLECTIONNER',
           'SEND': 'ENVOYER',
           'SEPARATE': 'SÉPARER',
           'SERVICE': 'SERVICE',
           'SERVICES': 'SERVICES',
           'SOFTWARE': 'LOGICIEL',
           'SUCCESSFULLY': 'AVEC SUCCÈS',
           'SUPPLIER': 'FOURNISSEUR',
           'THE': '',
           'THEN': 'PUIS',
           'THIS': 'CETTE',
           'TYPE': 'SAISIR',
           'TYPE NEW VALUE': 'SAISIR NOUVEAU VALUE',
           'UNDONE': 'ANNULÉ',
           'UNSUPPORTED': 'NON PRIS EN CHARGE',
           'UPDATE': 'METTRE À JOUR',
           'USE': 'UTILISER',
           'USING': 'EN UTILISANT',
           'WANT': 'VOULEZ',
           'WARRANTIES': 'GARANTIES',
           'WAS': 'A ÉTÉ',
           'WILL': 'VA',
           'WITH': 'AVEC',
           'YES': 'OUI',
           'YOU': 'VOUS'},
 'hi_IN': {'A': '',
           'ACTION': 'क्रिया',
           'ALL': 'सभी',
           'ALL FIELDS': 'सभी फ़ील्ड',
           'AN': '',
           'ANOTHER': 'दूसरा',
           'ARE': 'हैं',
           'AS': 'जैसा है',
           'AVAILABLE': 'उपलब्ध',
           'BACKUP': 'बैकअप',
           'CHANGES': 'बदलाव',
           'CLICK': 'क्लिक करें',
           'COMPUTER': 'कंप्यूटर',
           'CONTACT': 'संपर्क',
           'CONTINUE': 'जारी रखें',
           'CREATE': 'बनाएं',
           'CSV': 'CSV',
           'CURRENT': 'वर्तमान',
           'Computer Service Manager 3.1': 'कंप्यूटर सेवा प्रबंधक 3.1',
           'DATA': 'डेटा',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'क्या',
           'DRIVE': 'ड्राइव',
           'DUMMY': 'डमी',
           'ENTER': 'दर्ज करें',
           'EXCEL': 'एक्सेल',
           'EXISTING': 'मौजूदा',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'फ़ील्ड',
           'FILE': 'फ़ाइल',
           'FILES': 'फ़ाइलें',
           'FOLDER': 'फ़ोल्डर',
           'FOR': 'के लिए',
           'FORM': 'फ़ॉर्म',
           'FORMAT': 'प्रारूप',
           'FROM': 'से',
           'IMMEDIATELY': 'तुरंत',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'इनपुट',
           'IS': 'है',
           'ITEM': 'आइटम',
           'ITEMS': 'आइटम',
           'KIND': 'प्रकार',
           'LOCATION': 'स्थान',
           'MANAGER': 'प्रबंधक',
           'MESSAGE': 'संदेश',
           'MODE': 'मोड',
           'MULTIPLE': 'कई',
           'NEW': 'नया',
           'NO': 'नहीं',
           'OPEN': 'खोलें',
           'PART': 'पार्ट',
           'PATHS': 'पथ',
           'PLEASE': 'कृपया',
           'PROVIDED': 'प्रदान किया गया',
           'REPLACE': 'बदलें',
           'REQUIRED': 'आवश्यक',
           'RESERVED': 'सुरक्षित',
           'RESTORED': 'पुनर्स्थापित',
           'RESTORING': 'पुनर्स्थापना',
           'RIGHTS': 'अधिकार',
           'ROW': 'पंक्ति',
           'ROWS': 'पंक्तियाँ',
           'SAVED': 'सहेजा गया',
           'SELECT': 'चुनें',
           'SEND': 'भेजें',
           'SEPARATE': 'अलग करें',
           'SERVICE': 'सेवा',
           'SERVICES': 'सेवाएँ',
           'SOFTWARE': 'सॉफ़्टवेयर',
           'SUCCESSFULLY': 'सफलतापूर्वक',
           'SUPPLIER': 'आपूर्तिकर्ता',
           'THE': '',
           'THEN': 'फिर',
           'THIS': 'यह',
           'TYPE': 'टाइप करें',
           'TYPE NEW VALUE': 'टाइप करें नया VALUE',
           'UNDONE': 'पूर्ववत',
           'UNSUPPORTED': 'समर्थित नहीं',
           'UPDATE': 'अपडेट',
           'USE': 'उपयोग करें',
           'USING': 'का उपयोग करके',
           'WANT': 'चाहते हैं',
           'WARRANTIES': 'वारंटियाँ',
           'WAS': 'था',
           'WILL': 'होगा',
           'WITH': 'साथ',
           'YES': 'हाँ',
           'YOU': 'आप'},
 'id_ID': {'A': '',
           'ACTION': 'TINDAKAN',
           'ALL': 'SEMUA',
           'ALL FIELDS': 'SEMUA KOLOM',
           'AN': '',
           'ANOTHER': 'LAIN',
           'ARE': 'AKAN',
           'AS': 'SEBAGAIMANA',
           'AVAILABLE': 'TERSEDIA',
           'BACKUP': 'CADANGAN',
           'CHANGES': 'PERUBAHAN',
           'CLICK': 'KLIK',
           'COMPUTER': 'KOMPUTER',
           'CONTACT': 'KONTAK',
           'CONTINUE': 'LANJUTKAN',
           'CREATE': 'BUAT',
           'CSV': 'CSV',
           'CURRENT': 'SAAT INI',
           'Computer Service Manager 3.1': 'KOMPUTER SERVIS MANAJER 3.1',
           'DATA': 'DATA',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'APAKAH',
           'DRIVE': 'DRIVE',
           'DUMMY': 'DUMMY',
           'ENTER': 'MASUKKAN',
           'EXCEL': 'EXCEL',
           'EXISTING': 'YANG ADA',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'KOLOM',
           'FILE': 'BERKAS',
           'FILES': 'BERKAS',
           'FOLDER': 'FOLDER',
           'FOR': 'UNTUK',
           'FORM': 'FORMULIR',
           'FORMAT': 'FORMAT',
           'FROM': 'DARI',
           'IMMEDIATELY': 'LANGSUNG',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'INPUT',
           'IS': 'ADALAH',
           'ITEM': 'ITEM',
           'ITEMS': 'ITEM',
           'KIND': 'JENIS',
           'LOCATION': 'LOKASI',
           'MANAGER': 'MANAJER',
           'MESSAGE': 'PESAN',
           'MODE': 'MODE',
           'MULTIPLE': 'BEBERAPA',
           'NEW': 'BARU',
           'NO': 'TIDAK',
           'OPEN': 'BUKA',
           'PART': 'SUKU CADANG',
           'PATHS': 'LOKASI',
           'PLEASE': 'SILAKAN',
           'PROVIDED': 'DISEDIAKAN',
           'REPLACE': 'MENGGANTI',
           'REQUIRED': 'WAJIB',
           'RESERVED': 'DILINDUNGI',
           'RESTORED': 'DIPULIHKAN',
           'RESTORING': 'PEMULIHAN',
           'RIGHTS': 'HAK',
           'ROW': 'BARIS',
           'ROWS': 'BARIS',
           'SAVED': 'DISIMPAN',
           'SELECT': 'PILIH',
           'SEND': 'KIRIM',
           'SEPARATE': 'PISAHKAN',
           'SERVICE': 'SERVIS',
           'SERVICES': 'SERVIS',
           'SOFTWARE': 'PERANGKAT LUNAK',
           'SUCCESSFULLY': 'BERHASIL',
           'SUPPLIER': 'PEMASOK',
           'THE': '',
           'THEN': 'LALU',
           'THIS': 'INI',
           'TYPE': 'KETIK',
           'TYPE NEW VALUE': 'KETIK BARU VALUE',
           'UNDONE': 'DIBATALKAN',
           'UNSUPPORTED': 'TIDAK DIDUKUNG',
           'UPDATE': 'PERBARUI',
           'USE': 'GUNAKAN',
           'USING': 'MENGGUNAKAN',
           'WANT': 'INGIN',
           'WARRANTIES': 'GARANSI',
           'WAS': 'TELAH',
           'WILL': 'AKAN',
           'WITH': 'DENGAN',
           'YES': 'YA',
           'YOU': 'ANDA'},
 'it_IT': {'A': '',
           'ACTION': 'AZIONE',
           'ALL': 'TUTTI',
           'ALL FIELDS': 'TUTTI CAMPI',
           'AN': '',
           'ANOTHER': 'ALTRA',
           'ARE': 'SONO',
           'AS': 'COSÌ COMÈ',
           'AVAILABLE': 'DISPONIBILE',
           'BACKUP': 'BACKUP',
           'CHANGES': 'MODIFICHE',
           'CLICK': 'CLICCA',
           'COMPUTER': 'COMPUTER',
           'CONTACT': 'CONTATTO',
           'CONTINUE': 'CONTINUA',
           'CREATE': 'CREA',
           'CSV': 'CSV',
           'CURRENT': 'CORRENTE',
           'Computer Service Manager 3.1': 'COMPUTER SERVIZIO GESTORE 3.1',
           'DATA': 'DATI',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'VUOI',
           'DRIVE': 'UNITÀ',
           'DUMMY': 'DUMMY',
           'ENTER': 'INSERISCI',
           'EXCEL': 'EXCEL',
           'EXISTING': 'ESISTENTE',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'CAMPI',
           'FILE': 'FILE',
           'FILES': 'FILE',
           'FOLDER': 'CARTELLA',
           'FOR': 'PER',
           'FORM': 'MODULO',
           'FORMAT': 'FORMATO',
           'FROM': 'DA',
           'IMMEDIATELY': 'IMMEDIATAMENTE',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'INPUT',
           'IS': 'È',
           'ITEM': 'ELEMENTO',
           'ITEMS': 'ELEMENTI',
           'KIND': 'TIPO',
           'LOCATION': 'POSIZIONE',
           'MANAGER': 'GESTORE',
           'MESSAGE': 'MESSAGGIO',
           'MODE': 'MODALITÀ',
           'MULTIPLE': 'MULTIPLI',
           'NEW': 'NUOVO',
           'NO': 'NO',
           'OPEN': 'APRI',
           'PART': 'RICAMBIO',
           'PATHS': 'PERCORSI',
           'PLEASE': 'PER FAVORE',
           'PROVIDED': 'FORNITA',
           'REPLACE': 'SOSTITUIRE',
           'REQUIRED': 'RICHIESTO',
           'RESERVED': 'RISERVATI',
           'RESTORED': 'RIPRISTINATO',
           'RESTORING': 'RIPRISTINO',
           'RIGHTS': 'DIRITTI',
           'ROW': 'RIGA',
           'ROWS': 'RIGHE',
           'SAVED': 'SALVATO',
           'SELECT': 'SELEZIONA',
           'SEND': 'INVIA',
           'SEPARATE': 'SEPARA',
           'SERVICE': 'SERVIZIO',
           'SERVICES': 'SERVIZI',
           'SOFTWARE': 'SOFTWARE',
           'SUCCESSFULLY': 'CON SUCCESSO',
           'SUPPLIER': 'FORNITORE',
           'THE': '',
           'THEN': 'POI',
           'THIS': 'QUESTA',
           'TYPE': 'DIGITA',
           'TYPE NEW VALUE': 'DIGITA NUOVO VALUE',
           'UNDONE': 'ANNULLATO',
           'UNSUPPORTED': 'NON SUPPORTATO',
           'UPDATE': 'AGGIORNA',
           'USE': 'USA',
           'USING': 'USANDO',
           'WANT': 'VUOI',
           'WARRANTIES': 'GARANZIE',
           'WAS': 'È STATO',
           'WILL': 'SARÀ',
           'WITH': 'CON',
           'YES': 'SÌ',
           'YOU': 'TU'},
 'ja_JP': {'A': '',
           'ACTION': '操作',
           'ALL': 'すべて',
           'ALL FIELDS': 'すべて フィールド',
           'AN': '',
           'ANOTHER': '別の',
           'ARE': 'は',
           'AS': 'そのまま',
           'AVAILABLE': '利用可能',
           'BACKUP': 'バックアップ',
           'CHANGES': '変更',
           'CLICK': 'クリック',
           'COMPUTER': 'コンピューター',
           'CONTACT': '連絡先',
           'CONTINUE': '続行',
           'CREATE': '作成',
           'CSV': 'CSV',
           'CURRENT': '現在',
           'Computer Service Manager 3.1': 'コンピューター サービス マネージャー 3.1',
           'DATA': 'データ',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'しますか',
           'DRIVE': 'ドライブ',
           'DUMMY': 'ダミー',
           'ENTER': '入力',
           'EXCEL': 'Excel',
           'EXISTING': '既存',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'フィールド',
           'FILE': 'ファイル',
           'FILES': 'ファイル',
           'FOLDER': 'フォルダー',
           'FOR': 'のため',
           'FORM': 'フォーム',
           'FORMAT': '形式',
           'FROM': 'から',
           'IMMEDIATELY': 'すぐに',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': '入力',
           'IS': 'です',
           'ITEM': '項目',
           'ITEMS': '項目',
           'KIND': '種類',
           'LOCATION': '場所',
           'MANAGER': 'マネージャー',
           'MESSAGE': 'メッセージ',
           'MODE': 'モード',
           'MULTIPLE': '複数',
           'NEW': '新規',
           'NO': 'いいえ',
           'OPEN': '開く',
           'PART': '部品',
           'PATHS': 'パス',
           'PLEASE': 'してください',
           'PROVIDED': '提供',
           'REPLACE': '置き換える',
           'REQUIRED': '必須',
           'RESERVED': '留保',
           'RESTORED': '復元済み',
           'RESTORING': '復元',
           'RIGHTS': '権利',
           'ROW': '行',
           'ROWS': '行',
           'SAVED': '保存済み',
           'SELECT': '選択',
           'SEND': '送信',
           'SEPARATE': '区切る',
           'SERVICE': 'サービス',
           'SERVICES': 'サービス',
           'SOFTWARE': 'ソフトウェア',
           'SUCCESSFULLY': '成功',
           'SUPPLIER': '仕入先',
           'THE': '',
           'THEN': '次に',
           'THIS': 'この',
           'TYPE': '入力',
           'TYPE NEW VALUE': '入力 新規 VALUE',
           'UNDONE': '元に戻す',
           'UNSUPPORTED': '未対応',
           'UPDATE': '更新',
           'USE': '使用',
           'USING': '使用して',
           'WANT': 'したい',
           'WARRANTIES': '保証',
           'WAS': 'されました',
           'WILL': 'します',
           'WITH': 'で',
           'YES': 'はい',
           'YOU': 'あなた'},
 'ko_KR': {'A': '',
           'ACTION': '작업',
           'ALL': '모두',
           'ALL FIELDS': '모두 필드',
           'AN': '',
           'ANOTHER': '다른',
           'ARE': '은',
           'AS': '그대로',
           'AVAILABLE': '사용 가능',
           'BACKUP': '백업',
           'CHANGES': '변경 사항',
           'CLICK': '클릭',
           'COMPUTER': '컴퓨터',
           'CONTACT': '연락처',
           'CONTINUE': '계속',
           'CREATE': '만들기',
           'CSV': 'CSV',
           'CURRENT': '현재',
           'Computer Service Manager 3.1': '컴퓨터 서비스 관리자 3.1',
           'DATA': '데이터',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': '하시겠습니까',
           'DRIVE': '드라이브',
           'DUMMY': '더미',
           'ENTER': '입력',
           'EXCEL': 'Excel',
           'EXISTING': '기존',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': '필드',
           'FILE': '파일',
           'FILES': '파일',
           'FOLDER': '폴더',
           'FOR': '위해',
           'FORM': '양식',
           'FORMAT': '형식',
           'FROM': '에서',
           'IMMEDIATELY': '즉시',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': '입력',
           'IS': '입니다',
           'ITEM': '항목',
           'ITEMS': '항목',
           'KIND': '종류',
           'LOCATION': '위치',
           'MANAGER': '관리자',
           'MESSAGE': '메시지',
           'MODE': '모드',
           'MULTIPLE': '여러',
           'NEW': '새로 만들기',
           'NO': '아니요',
           'OPEN': '열기',
           'PART': '부품',
           'PATHS': '경로',
           'PLEASE': '제발',
           'PROVIDED': '제공됨',
           'REPLACE': '교체',
           'REQUIRED': '필수',
           'RESERVED': '보유',
           'RESTORED': '복원됨',
           'RESTORING': '복원',
           'RIGHTS': '권리',
           'ROW': '행',
           'ROWS': '행',
           'SAVED': '저장됨',
           'SELECT': '선택',
           'SEND': '보내기',
           'SEPARATE': '구분',
           'SERVICE': '서비스',
           'SERVICES': '서비스',
           'SOFTWARE': '소프트웨어',
           'SUCCESSFULLY': '성공적으로',
           'SUPPLIER': '공급업체',
           'THE': '',
           'THEN': '그런 다음',
           'THIS': '이',
           'TYPE': '입력',
           'TYPE NEW VALUE': '입력 새로 만들기 VALUE',
           'UNDONE': '실행 취소',
           'UNSUPPORTED': '지원되지 않음',
           'UPDATE': '업데이트',
           'USE': '사용',
           'USING': '사용하여',
           'WANT': '원함',
           'WARRANTIES': '보증',
           'WAS': '되었습니다',
           'WILL': '예정',
           'WITH': '와',
           'YES': '예',
           'YOU': '사용자'},
 'ms_MY': {'ALL': 'SEMUA',
           'ALL FIELDS': 'SEMUA MEDAN',
           'AS': 'SEBAGAIMANA',
           'AVAILABLE': 'TERSEDIA',
           'CHANGES': 'PERUBAHAN',
           'CLICK': 'KLIK',
           'COMPUTER': 'KOMPUTER',
           'CSV': 'CSV',
           'Computer Service Manager 3.1': 'KOMPUTER SERVIS PENGURUS 3.1',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DUMMY': 'DUMMY',
           'ENTER': 'MASUKKAN',
           'EXCEL': 'EXCEL',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'MEDAN',
           'FILE': 'FAIL',
           'FILES': 'FAIL',
           'FOR': 'UNTUK',
           'FORMAT': 'FORMAT',
           'IMMEDIATELY': 'SERTA-MERTA',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'INPUT',
           'IS': 'ADALAH',
           'ITEMS': 'ITEM',
           'KIND': 'JENIS',
           'LOCATION': 'LOKASI',
           'MANAGER': 'PENGURUS',
           'MESSAGE': 'MESEJ',
           'MODE': 'MOD',
           'MULTIPLE': 'BERBILANG',
           'NEW': 'BAHARU',
           'PATHS': 'LALUAN',
           'PLEASE': 'SILA',
           'PROVIDED': 'DISEDIAKAN',
           'REQUIRED': 'DIPERLUKAN',
           'RESERVED': 'TERPELIHARA',
           'RESTORED': 'DIPULIHKAN',
           'RIGHTS': 'HAK',
           'SAVED': 'DISIMPAN',
           'SEPARATE': 'ASINGKAN',
           'SERVICE': 'SERVIS',
           'SOFTWARE': 'PERISIAN',
           'THEN': 'KEMUDIAN',
           'THIS': 'INI',
           'TYPE': 'TAIP',
           'TYPE NEW VALUE': 'TAIP BAHARU VALUE',
           'UNSUPPORTED': 'TIDAK DISOKONG',
           'UPDATE': 'KEMAS KINI',
           'USE': 'GUNAKAN',
           'USING': 'MENGGUNAKAN',
           'WANT': 'MAHU',
           'WARRANTIES': 'WARANTI',
           'WAS': 'TELAH',
           'WITH': 'DENGAN',
           'YOU': 'ANDA'},
 'nl_NL': {'A': '',
           'ACTION': 'ACTIE',
           'ALL': 'ALLE',
           'ALL FIELDS': 'ALLE VELDEN',
           'AN': '',
           'ANOTHER': 'ANDERE',
           'ARE': 'WORDEN',
           'AS': 'ZOALS',
           'AVAILABLE': 'BESCHIKBAAR',
           'BACKUP': 'BACK-UP',
           'CHANGES': 'WIJZIGINGEN',
           'CLICK': 'KLIK',
           'COMPUTER': 'COMPUTER',
           'CONTACT': 'CONTACT',
           'CONTINUE': 'DOORGAAN',
           'CREATE': 'MAKEN',
           'CSV': 'CSV',
           'CURRENT': 'HUIDIG',
           'Computer Service Manager 3.1': 'COMPUTER SERVICE BEHEERDER 3.1',
           'DATA': 'GEGEVENS',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'WILT',
           'DRIVE': 'SCHIJF',
           'DUMMY': 'DUMMY',
           'ENTER': 'VOER IN',
           'EXCEL': 'EXCEL',
           'EXISTING': 'BESTAAND',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'VELDEN',
           'FILE': 'BESTAND',
           'FILES': 'BESTANDEN',
           'FOLDER': 'MAP',
           'FOR': 'VOOR',
           'FORM': 'FORMULIER',
           'FORMAT': 'FORMAAT',
           'FROM': 'VAN',
           'IMMEDIATELY': 'DIRECT',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'INVOER',
           'IS': 'IS',
           'ITEM': 'ITEM',
           'ITEMS': 'ITEMS',
           'KIND': 'SOORT',
           'LOCATION': 'LOCATIE',
           'MANAGER': 'BEHEERDER',
           'MESSAGE': 'BERICHT',
           'MODE': 'MODUS',
           'MULTIPLE': 'MEERDERE',
           'NEW': 'NIEUW',
           'NO': 'NEE',
           'OPEN': 'OPENEN',
           'PART': 'ONDERDEEL',
           'PATHS': 'PADEN',
           'PLEASE': 'GELIEVE',
           'PROVIDED': 'GELEVERD',
           'REPLACE': 'VERVANGEN',
           'REQUIRED': 'VEREIST',
           'RESERVED': 'VOORBEHOUDEN',
           'RESTORED': 'HERSTELD',
           'RESTORING': 'HERSTELLEN',
           'RIGHTS': 'RECHTEN',
           'ROW': 'RIJ',
           'ROWS': 'RIJEN',
           'SAVED': 'OPGESLAGEN',
           'SELECT': 'SELECTEER',
           'SEND': 'VERSTUREN',
           'SEPARATE': 'SCHEIDEN',
           'SERVICE': 'SERVICE',
           'SERVICES': 'SERVICES',
           'SOFTWARE': 'SOFTWARE',
           'SUCCESSFULLY': 'SUCCESVOL',
           'SUPPLIER': 'LEVERANCIER',
           'THE': '',
           'THEN': 'DAARNA',
           'THIS': 'DEZE',
           'TYPE': 'TYP',
           'TYPE NEW VALUE': 'TYP NIEUW VALUE',
           'UNDONE': 'ONGEDAAN',
           'UNSUPPORTED': 'NIET ONDERSTEUND',
           'UPDATE': 'BIJWERKEN',
           'USE': 'GEBRUIK',
           'USING': 'MET',
           'WANT': 'WILT',
           'WARRANTIES': 'GARANTIES',
           'WAS': 'WAS',
           'WILL': 'ZAL',
           'WITH': 'MET',
           'YES': 'JA',
           'YOU': 'U'},
 'pt_BR': {'A': '',
           'ACTION': 'AÇÃO',
           'ALL': 'TODOS',
           'ALL FIELDS': 'TODOS CAMPOS',
           'AN': '',
           'ANOTHER': 'OUTRA',
           'ARE': 'SÃO',
           'AS': 'COMO',
           'AVAILABLE': 'DISPONÍVEL',
           'BACKUP': 'BACKUP',
           'CHANGES': 'ALTERAÇÕES',
           'CLICK': 'CLIQUE',
           'COMPUTER': 'COMPUTADOR',
           'CONTACT': 'CONTATO',
           'CONTINUE': 'CONTINUAR',
           'CREATE': 'CRIAR',
           'CSV': 'CSV',
           'CURRENT': 'ATUAL',
           'Computer Service Manager 3.1': 'COMPUTADOR SERVIÇO GERENCIADOR 3.1',
           'DATA': 'DADOS',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'VOCÊ',
           'DRIVE': 'UNIDADE',
           'DUMMY': 'DUMMY',
           'ENTER': 'INSIRA',
           'EXCEL': 'EXCEL',
           'EXISTING': 'EXISTENTE',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'CAMPOS',
           'FILE': 'ARQUIVO',
           'FILES': 'ARQUIVOS',
           'FOLDER': 'PASTA',
           'FOR': 'PARA',
           'FORM': 'FORMULÁRIO',
           'FORMAT': 'FORMATO',
           'FROM': 'DE',
           'IMMEDIATELY': 'IMEDIATAMENTE',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'ENTRADA',
           'IS': 'É',
           'ITEM': 'ITEM',
           'ITEMS': 'ITENS',
           'KIND': 'TIPO',
           'LOCATION': 'LOCALIZAÇÃO',
           'MANAGER': 'GERENCIADOR',
           'MESSAGE': 'MENSAGEM',
           'MODE': 'MODO',
           'MULTIPLE': 'MÚLTIPLOS',
           'NEW': 'NOVO',
           'NO': 'NÃO',
           'OPEN': 'ABRIR',
           'PART': 'PEÇA',
           'PATHS': 'CAMINHOS',
           'PLEASE': 'POR FAVOR',
           'PROVIDED': 'FORNECIDA',
           'REPLACE': 'SUBSTITUIR',
           'REQUIRED': 'OBRIGATÓRIO',
           'RESERVED': 'RESERVADOS',
           'RESTORED': 'RESTAURADO',
           'RESTORING': 'RESTAURAR',
           'RIGHTS': 'DIREITOS',
           'ROW': 'LINHA',
           'ROWS': 'LINHAS',
           'SAVED': 'SALVO',
           'SELECT': 'SELECIONE',
           'SEND': 'ENVIAR',
           'SEPARATE': 'SEPARAR',
           'SERVICE': 'SERVIÇO',
           'SERVICES': 'SERVIÇOS',
           'SOFTWARE': 'SOFTWARE',
           'SUCCESSFULLY': 'COM SUCESSO',
           'SUPPLIER': 'FORNECEDOR',
           'THE': '',
           'THEN': 'DEPOIS',
           'THIS': 'ESTE',
           'TYPE': 'DIGITE',
           'TYPE NEW VALUE': 'DIGITE NOVO VALUE',
           'UNDONE': 'DESFEITO',
           'UNSUPPORTED': 'NÃO SUPORTADO',
           'UPDATE': 'ATUALIZAR',
           'USE': 'USAR',
           'USING': 'USANDO',
           'WANT': 'DESEJA',
           'WARRANTIES': 'GARANTIAS',
           'WAS': 'FOI',
           'WILL': 'IRÁ',
           'WITH': 'COM',
           'YES': 'SIM',
           'YOU': 'VOCÊ'},
 'ru_RU': {'A': '',
           'ACTION': 'ДЕЙСТВИЕ',
           'ALL': 'ВСЕ',
           'ALL FIELDS': 'ВСЕ ПОЛЯ',
           'AN': '',
           'ANOTHER': 'ДРУГОЙ',
           'ARE': 'БУДУТ',
           'AS': 'КАК',
           'AVAILABLE': 'ДОСТУПЕН',
           'BACKUP': 'РЕЗЕРВНАЯ КОПИЯ',
           'CHANGES': 'ИЗМЕНЕНИЯ',
           'CLICK': 'НАЖМИТЕ',
           'COMPUTER': 'КОМПЬЮТЕР',
           'CONTACT': 'КОНТАКТ',
           'CONTINUE': 'ПРОДОЛЖИТЬ',
           'CREATE': 'СОЗДАТЬ',
           'CSV': 'CSV',
           'CURRENT': 'ТЕКУЩИЙ',
           'Computer Service Manager 3.1': 'КОМПЬЮТЕР СЕРВИС МЕНЕДЖЕР 3.1',
           'DATA': 'ДАННЫЕ',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': 'ВЫ',
           'DRIVE': 'ДИСК',
           'DUMMY': 'DUMMY',
           'ENTER': 'ВВЕДИТЕ',
           'EXCEL': 'EXCEL',
           'EXISTING': 'СУЩЕСТВУЮЩИЙ',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'ПОЛЯ',
           'FILE': 'ФАЙЛ',
           'FILES': 'ФАЙЛЫ',
           'FOLDER': 'ПАПКА',
           'FOR': 'ДЛЯ',
           'FORM': 'ФОРМА',
           'FORMAT': 'ФОРМАТ',
           'FROM': 'ИЗ',
           'IMMEDIATELY': 'СРАЗУ',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'ВВОД',
           'IS': 'ЕСТЬ',
           'ITEM': 'ЭЛЕМЕНТ',
           'ITEMS': 'ЭЛЕМЕНТЫ',
           'KIND': 'ВИД',
           'LOCATION': 'МЕСТО',
           'MANAGER': 'МЕНЕДЖЕР',
           'MESSAGE': 'СООБЩЕНИЕ',
           'MODE': 'РЕЖИМ',
           'MULTIPLE': 'НЕСКОЛЬКО',
           'NEW': 'НОВЫЙ',
           'NO': 'НЕТ',
           'OPEN': 'ОТКРЫТЬ',
           'PART': 'ЗАПЧАСТЬ',
           'PATHS': 'ПУТИ',
           'PLEASE': 'ПОЖАЛУЙСТА',
           'PROVIDED': 'ПРЕДОСТАВЛЕНО',
           'REPLACE': 'ЗАМЕНИТЬ',
           'REQUIRED': 'ОБЯЗАТЕЛЬНО',
           'RESERVED': 'ЗАЩИЩЕНЫ',
           'RESTORED': 'ВОССТАНОВЛЕНО',
           'RESTORING': 'ВОССТАНОВЛЕНИЕ',
           'RIGHTS': 'ПРАВА',
           'ROW': 'СТРОКА',
           'ROWS': 'СТРОКИ',
           'SAVED': 'СОХРАНЕНО',
           'SELECT': 'ВЫБЕРИТЕ',
           'SEND': 'ОТПРАВИТЬ',
           'SEPARATE': 'РАЗДЕЛИТЬ',
           'SERVICE': 'СЕРВИС',
           'SERVICES': 'УСЛУГИ',
           'SOFTWARE': 'ПРОГРАММА',
           'SUCCESSFULLY': 'УСПЕШНО',
           'SUPPLIER': 'ПОСТАВЩИК',
           'THE': '',
           'THEN': 'ЗАТЕМ',
           'THIS': 'ЭТОТ',
           'TYPE': 'ВВЕДИТЕ',
           'TYPE NEW VALUE': 'ВВЕДИТЕ НОВЫЙ VALUE',
           'UNDONE': 'ОТМЕНЕНО',
           'UNSUPPORTED': 'НЕ ПОДДЕРЖИВАЕТСЯ',
           'UPDATE': 'ОБНОВИТЬ',
           'USE': 'ИСПОЛЬЗОВАТЬ',
           'USING': 'ИСПОЛЬЗУЯ',
           'WANT': 'ХОТИТЕ',
           'WARRANTIES': 'ГАРАНТИИ',
           'WAS': 'БЫЛ',
           'WILL': 'БУДЕТ',
           'WITH': 'С',
           'YES': 'ДА',
           'YOU': 'ВЫ'},
 'th_TH': {'ALL': 'ทั้งหมด',
           'ALL FIELDS': 'ทั้งหมด ฟิลด์',
           'AS': 'ตามสภาพ',
           'AVAILABLE': 'พร้อมใช้งาน',
           'CHANGES': 'การเปลี่ยนแปลง',
           'CLICK': 'คลิก',
           'COMPUTER': 'คอมพิวเตอร์',
           'CSV': 'CSV',
           'Computer Service Manager 3.1': 'คอมพิวเตอร์ บริการ ตัวจัดการ 3.1',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DUMMY': 'ดัมมี่',
           'ENTER': 'ป้อน',
           'EXCEL': 'EXCEL',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'ฟิลด์',
           'FILE': 'ไฟล์',
           'FILES': 'ไฟล์',
           'FOR': 'สำหรับ',
           'FORMAT': 'รูปแบบ',
           'IMMEDIATELY': 'ทันที',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'อินพุต',
           'IS': 'คือ',
           'ITEMS': 'รายการ',
           'KIND': 'ชนิด',
           'LOCATION': 'ตำแหน่ง',
           'MANAGER': 'ตัวจัดการ',
           'MESSAGE': 'ข้อความ',
           'MODE': 'โหมด',
           'MULTIPLE': 'หลาย',
           'NEW': 'ใหม่',
           'PATHS': 'พาธ',
           'PLEASE': 'โปรด',
           'PROVIDED': 'ให้ไว้',
           'REQUIRED': 'จำเป็น',
           'RESERVED': 'สงวนไว้',
           'RESTORED': 'กู้คืนแล้ว',
           'RIGHTS': 'สิทธิ์',
           'SAVED': 'บันทึกแล้ว',
           'SEPARATE': 'คั่น',
           'SERVICE': 'บริการ',
           'SOFTWARE': 'ซอฟต์แวร์',
           'THEN': 'จากนั้น',
           'THIS': 'นี้',
           'TYPE': 'พิมพ์',
           'TYPE NEW VALUE': 'พิมพ์ ใหม่ VALUE',
           'UNSUPPORTED': 'ไม่รองรับ',
           'UPDATE': 'อัปเดต',
           'USE': 'ใช้',
           'USING': 'โดยใช้',
           'WANT': 'ต้องการ',
           'WARRANTIES': 'การรับประกัน',
           'WAS': 'ถูก',
           'WITH': 'กับ',
           'YOU': 'คุณ'},
 'tr_TR': {'ALL': 'TÜMÜ',
           'ALL FIELDS': 'TÜMÜ ALANLAR',
           'AS': 'OLDUĞU GİBİ',
           'AVAILABLE': 'MEVCUT',
           'CHANGES': 'DEĞİŞİKLİKLER',
           'CLICK': 'TIKLA',
           'COMPUTER': 'BİLGİSAYAR',
           'CSV': 'CSV',
           'Computer Service Manager 3.1': 'BİLGİSAYAR SERVİS YÖNETİCİ 3.1',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': '',
           'DUMMY': 'DUMMY',
           'ENTER': 'GİRİN',
           'EXCEL': 'EXCEL',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'ALANLAR',
           'FILE': 'DOSYA',
           'FILES': 'DOSYALAR',
           'FOR': 'İÇİN',
           'FORMAT': 'FORMAT',
           'IMMEDIATELY': 'HEMEN',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'GİRİŞ',
           'IS': 'DİR',
           'ITEMS': 'ÖĞELER',
           'KIND': 'TÜR',
           'LOCATION': 'KONUM',
           'MANAGER': 'YÖNETİCİ',
           'MESSAGE': 'MESAJ',
           'MODE': 'MOD',
           'MULTIPLE': 'ÇOKLU',
           'NEW': 'YENİ',
           'PATHS': 'YOLLAR',
           'PLEASE': 'LÜTFEN',
           'PROVIDED': 'SAĞLANIR',
           'REQUIRED': 'GEREKLİ',
           'RESERVED': 'SAKLI',
           'RESTORED': 'GERİ YÜKLENDİ',
           'RIGHTS': 'HAKLAR',
           'SAVED': 'KAYDEDİLDİ',
           'SEPARATE': 'AYIR',
           'SERVICE': 'SERVİS',
           'SOFTWARE': 'YAZILIM',
           'THEN': 'SONRA',
           'THIS': 'BU',
           'TYPE': 'YAZIN',
           'TYPE NEW VALUE': 'YAZIN YENİ VALUE',
           'UNSUPPORTED': 'DESTEKLENMEYEN',
           'UPDATE': 'GÜNCELLE',
           'USE': 'KULLAN',
           'USING': 'KULLANARAK',
           'WANT': 'İSTİYOR',
           'WARRANTIES': 'GARANTİLER',
           'WAS': 'BULUNDU',
           'WITH': 'İLE',
           'YOU': 'SİZ'},
 'vi_VN': {'ALL': 'TẤT CẢ',
           'ALL FIELDS': 'TẤT CẢ TRƯỜNG',
           'AS': 'NHƯ',
           'AVAILABLE': 'CÓ SẴN',
           'CHANGES': 'THAY ĐỔI',
           'CLICK': 'NHẤP',
           'COMPUTER': 'MÁY TÍNH',
           'CSV': 'CSV',
           'Computer Service Manager 3.1': 'MÁY TÍNH DỊCH VỤ TRÌNH QUẢN LÝ 3.1',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DUMMY': 'DUMMY',
           'ENTER': 'NHẬP',
           'EXCEL': 'EXCEL',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': 'TRƯỜNG',
           'FILE': 'TỆP',
           'FILES': 'TỆP',
           'FOR': 'CHO',
           'FORMAT': 'ĐỊNH DẠNG',
           'IMMEDIATELY': 'NGAY',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': 'ĐẦU VÀO',
           'IS': 'LÀ',
           'ITEMS': 'MỤC',
           'KIND': 'LOẠI',
           'LOCATION': 'VỊ TRÍ',
           'MANAGER': 'TRÌNH QUẢN LÝ',
           'MESSAGE': 'TIN NHẮN',
           'MODE': 'CHẾ ĐỘ',
           'MULTIPLE': 'NHIỀU',
           'NEW': 'MỚI',
           'PATHS': 'ĐƯỜNG DẪN',
           'PLEASE': 'VUI LÒNG',
           'PROVIDED': 'ĐƯỢC CUNG CẤP',
           'REQUIRED': 'BẮT BUỘC',
           'RESERVED': 'BẢO LƯU',
           'RESTORED': 'ĐÃ KHÔI PHỤC',
           'RIGHTS': 'QUYỀN',
           'SAVED': 'ĐÃ LƯU',
           'SEPARATE': 'TÁCH',
           'SERVICE': 'DỊCH VỤ',
           'SOFTWARE': 'PHẦN MỀM',
           'THEN': 'RỒI',
           'THIS': 'NÀY',
           'TYPE': 'NHẬP',
           'TYPE NEW VALUE': 'NHẬP MỚI VALUE',
           'UNSUPPORTED': 'KHÔNG HỖ TRỢ',
           'UPDATE': 'CẬP NHẬT',
           'USE': 'SỬ DỤNG',
           'USING': 'BẰNG',
           'WANT': 'MUỐN',
           'WARRANTIES': 'BẢO HÀNH',
           'WAS': 'ĐÃ',
           'WITH': 'VỚI',
           'YOU': 'BẠN'},
 'zh_CN': {'A': '',
           'ACTION': '操作',
           'ALL': '全部',
           'ALL FIELDS': '全部 字段',
           'AN': '',
           'ANOTHER': '另一个',
           'ARE': '已',
           'AS': '按原样',
           'AVAILABLE': '可用',
           'BACKUP': '备份',
           'CHANGES': '更改',
           'CLICK': '点击',
           'COMPUTER': '计算机',
           'CONTACT': '联系人',
           'CONTINUE': '继续',
           'CREATE': '创建',
           'CSV': 'CSV',
           'CURRENT': '当前',
           'Computer Service Manager 3.1': '计算机 服务 管理器 3.1',
           'DATA': '数据',
           'DELETE SELECTED': 'DELETE SELECTED',
           'DO': '是否',
           'DRIVE': '驱动器',
           'DUMMY': '示例',
           'ENTER': '输入',
           'EXCEL': 'Excel',
           'EXISTING': '现有',
           'EXPORT TO EXCEL': 'EXPORT TO EXCEL',
           'FIELDS': '字段',
           'FILE': '文件',
           'FILES': '文件',
           'FOLDER': '文件夹',
           'FOR': '为',
           'FORM': '表单',
           'FORMAT': '格式',
           'FROM': '从',
           'IMMEDIATELY': '立即',
           'IMPORT EXCEL OR CSV': 'IMPORT EXCEL OR CSV',
           'INF': 'INF',
           'INPUT': '输入',
           'IS': '是',
           'ITEM': '项目',
           'ITEMS': '项目',
           'KIND': '类型',
           'LOCATION': '位置',
           'MANAGER': '管理器',
           'MESSAGE': '消息',
           'MODE': '模式',
           'MULTIPLE': '多个',
           'NEW': '新建',
           'NO': '否',
           'OPEN': '打开',
           'PART': '配件',
           'PATHS': '路径',
           'PLEASE': '请',
           'PROVIDED': '提供',
           'REPLACE': '替换',
           'REQUIRED': '必需',
           'RESERVED': '保留',
           'RESTORED': '已恢复',
           'RESTORING': '恢复',
           'RIGHTS': '权利',
           'ROW': '行',
           'ROWS': '行',
           'SAVED': '保存',
           'SELECT': '选择',
           'SEND': '发送',
           'SEPARATE': '分隔',
           'SERVICE': '服务',
           'SERVICES': '服务',
           'SOFTWARE': '软件',
           'SUCCESSFULLY': '成功',
           'SUPPLIER': '供应商',
           'THE': '',
           'THEN': '然后',
           'THIS': '此',
           'TYPE': '输入',
           'TYPE NEW VALUE': '输入 新建 VALUE',
           'UNDONE': '撤销',
           'UNSUPPORTED': '不支持',
           'UPDATE': '更新',
           'USE': '使用',
           'USING': '使用',
           'WANT': '想要',
           'WARRANTIES': '保修',
           'WAS': '已',
           'WILL': '将',
           'WITH': '用',
           'YES': '是',
           'YOU': '您'}}

for _code, _items in _FULL_UI_SUPPLEMENTAL_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)


_THEME_UI_TRANSLATIONS = {
    "id_ID": {"THEME": "TEMA", "LIGHT": "TERANG", "DARK": "GELAP", "BLUE": "BIRU", "PURPLE": "UNGU", "RED": "MERAH", "ORANGE": "ORANYE", "CLASSIC": "KLASIK"},
    "es_MX": {"THEME": "TEMA", "LIGHT": "CLARO", "DARK": "OSCURO", "BLUE": "AZUL", "PURPLE": "MORADO", "RED": "ROJO", "ORANGE": "NARANJA", "CLASSIC": "CLÁSICO"},
    "fr_FR": {"THEME": "THÈME", "LIGHT": "CLAIR", "DARK": "SOMBRE", "BLUE": "BLEU", "PURPLE": "VIOLET", "RED": "ROUGE", "ORANGE": "ORANGE", "CLASSIC": "CLASSIQUE"},
    "de_DE": {"THEME": "DESIGN", "LIGHT": "HELL", "DARK": "DUNKEL", "BLUE": "BLAU", "PURPLE": "VIOLETT", "RED": "ROT", "ORANGE": "ORANGE", "CLASSIC": "KLASSISCH"},
    "pt_BR": {"THEME": "TEMA", "LIGHT": "CLARO", "DARK": "ESCURO", "BLUE": "AZUL", "PURPLE": "ROXO", "RED": "VERMELHO", "ORANGE": "LARANJA", "CLASSIC": "CLÁSSICO"},
    "it_IT": {"THEME": "TEMA", "LIGHT": "CHIARO", "DARK": "SCURO", "BLUE": "BLU", "PURPLE": "VIOLA", "RED": "ROSSO", "ORANGE": "ARANCIONE", "CLASSIC": "CLASSICO"},
    "nl_NL": {"THEME": "THEMA", "LIGHT": "LICHT", "DARK": "DONKER", "BLUE": "BLAUW", "PURPLE": "PAARS", "RED": "ROOD", "ORANGE": "ORANJE", "CLASSIC": "KLASSIEK"},
    "zh_CN": {"THEME": "主题", "LIGHT": "浅色", "DARK": "深色", "BLUE": "蓝色", "PURPLE": "紫色", "RED": "红色", "ORANGE": "橙色", "CLASSIC": "经典"},
    "ja_JP": {"THEME": "テーマ", "LIGHT": "ライト", "DARK": "ダーク", "BLUE": "ブルー", "PURPLE": "パープル", "RED": "レッド", "ORANGE": "オレンジ", "CLASSIC": "クラシック"},
    "ko_KR": {"THEME": "테마", "LIGHT": "라이트", "DARK": "다크", "BLUE": "블루", "PURPLE": "퍼플", "RED": "레드", "ORANGE": "오렌지", "CLASSIC": "클래식"},
    "ar_SA": {"THEME": "السمة", "LIGHT": "فاتح", "DARK": "داكن", "BLUE": "أزرق", "PURPLE": "بنفسجي", "RED": "أحمر", "ORANGE": "برتقالي", "CLASSIC": "كلاسيكي"},
    "hi_IN": {"THEME": "थीम", "LIGHT": "हल्का", "DARK": "गहरा", "BLUE": "नीला", "PURPLE": "बैंगनी", "RED": "लाल", "ORANGE": "नारंगी", "CLASSIC": "क्लासिक"},
    "bn_BD": {"THEME": "থিম", "LIGHT": "হালকা", "DARK": "গাঢ়", "BLUE": "নীল", "PURPLE": "বেগুনি", "RED": "লাল", "ORANGE": "কমলা", "CLASSIC": "ক্লাসিক"},
    "ru_RU": {"THEME": "ТЕМА", "LIGHT": "СВЕТЛАЯ", "DARK": "ТЁМНАЯ", "BLUE": "СИНЯЯ", "PURPLE": "ФИОЛЕТОВАЯ", "RED": "КРАСНАЯ", "ORANGE": "ОРАНЖЕВАЯ", "CLASSIC": "КЛАССИЧЕСКАЯ"},
    "tr_TR": {"THEME": "TEMA", "LIGHT": "AÇIK", "DARK": "KOYU", "BLUE": "MAVİ", "PURPLE": "MOR", "RED": "KIRMIZI", "ORANGE": "TURUNCU", "CLASSIC": "KLASİK"},
    "vi_VN": {"THEME": "GIAO DIỆN", "LIGHT": "SÁNG", "DARK": "TỐI", "BLUE": "XANH DƯƠNG", "PURPLE": "TÍM", "RED": "ĐỎ", "ORANGE": "CAM", "CLASSIC": "CỔ ĐIỂN"},
    "th_TH": {"THEME": "ธีม", "LIGHT": "สว่าง", "DARK": "มืด", "BLUE": "สีน้ำเงิน", "PURPLE": "สีม่วง", "RED": "สีแดง", "ORANGE": "สีส้ม", "CLASSIC": "คลาสสิก"},
    "ms_MY": {"THEME": "TEMA", "LIGHT": "CERAH", "DARK": "GELAP", "BLUE": "BIRU", "PURPLE": "UNGU", "RED": "MERAH", "ORANGE": "JINGGA", "CLASSIC": "KLASIK"},
    "fil_PH": {"THEME": "TEMA", "LIGHT": "MALIWANAG", "DARK": "MADILIM", "BLUE": "ASUL", "PURPLE": "LILA", "RED": "PULA", "ORANGE": "KAHEL", "CLASSIC": "KLASIKO"},
}
for _code, _items in _THEME_UI_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)


# Exact translations for supplier/sparepart import actions and feedback.
# These complete phrases avoid awkward word order in languages where the
# token-by-token translation fallback is not natural.
_IMPORT_DATA_FEATURE_TRANSLATIONS = {
    "id_ID": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPOR DATA PENGGUNA DARI EXCEL ATAU CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPOR DATA PEMASOK DARI EXCEL ATAU CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPOR DATA SUKU CADANG DARI EXCEL ATAU CSV",
        "SKIPPED": "DILEWATI",
        "MISSING REQUIRED COLUMNS": "KOLOM WAJIB TIDAK DITEMUKAN",
    },
    "es_MX": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPORTAR DATOS DE USUARIO DESDE EXCEL O CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPORTAR PROVEEDORES DESDE EXCEL O CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPORTAR PIEZAS DESDE EXCEL O CSV",
        "SKIPPED": "OMITIDAS",
        "MISSING REQUIRED COLUMNS": "FALTAN COLUMNAS OBLIGATORIAS",
    },
    "fr_FR": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPORTER LES DONNÉES UTILISATEUR DEPUIS EXCEL OU CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPORTER LES FOURNISSEURS DEPUIS EXCEL OU CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPORTER LES PIÈCES DEPUIS EXCEL OU CSV",
        "SKIPPED": "IGNORÉES",
        "MISSING REQUIRED COLUMNS": "COLONNES OBLIGATOIRES MANQUANTES",
    },
    "de_DE": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "BENUTZERDATEN AUS EXCEL ODER CSV IMPORTIEREN",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "LIEFERANTEN AUS EXCEL ODER CSV IMPORTIEREN",
        "IMPORT PARTS FROM EXCEL OR CSV": "TEILE AUS EXCEL ODER CSV IMPORTIEREN",
        "SKIPPED": "ÜBERSPRUNGEN",
        "MISSING REQUIRED COLUMNS": "ERFORDERLICHE SPALTEN FEHLEN",
    },
    "pt_BR": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPORTAR DADOS DO USUÁRIO DO EXCEL OU CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPORTAR FORNECEDORES DO EXCEL OU CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPORTAR PEÇAS DO EXCEL OU CSV",
        "SKIPPED": "IGNORADAS",
        "MISSING REQUIRED COLUMNS": "COLUNAS OBRIGATÓRIAS AUSENTES",
    },
    "it_IT": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPORTA DATI UTENTE DA EXCEL O CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPORTA FORNITORI DA EXCEL O CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPORTA RICAMBI DA EXCEL O CSV",
        "SKIPPED": "SALTATE",
        "MISSING REQUIRED COLUMNS": "COLONNE OBBLIGATORIE MANCANTI",
    },
    "nl_NL": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "GEBRUIKERSGEGEVENS IMPORTEREN UIT EXCEL OF CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "LEVERANCIERS IMPORTEREN UIT EXCEL OF CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "ONDERDELEN IMPORTEREN UIT EXCEL OF CSV",
        "SKIPPED": "OVERGESLAGEN",
        "MISSING REQUIRED COLUMNS": "VEREISTE KOLOMMEN ONTBREKEN",
    },
    "zh_CN": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "从 EXCEL 或 CSV 导入用户数据",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "从 EXCEL 或 CSV 导入供应商",
        "IMPORT PARTS FROM EXCEL OR CSV": "从 EXCEL 或 CSV 导入配件",
        "SKIPPED": "已跳过",
        "MISSING REQUIRED COLUMNS": "缺少必填列",
    },
    "ja_JP": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "EXCEL または CSV からユーザーデータをインポート",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "EXCEL または CSV から仕入先をインポート",
        "IMPORT PARTS FROM EXCEL OR CSV": "EXCEL または CSV から部品をインポート",
        "SKIPPED": "スキップ",
        "MISSING REQUIRED COLUMNS": "必須列がありません",
    },
    "ko_KR": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "EXCEL 또는 CSV에서 사용자 데이터 가져오기",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "EXCEL 또는 CSV에서 공급업체 가져오기",
        "IMPORT PARTS FROM EXCEL OR CSV": "EXCEL 또는 CSV에서 부품 가져오기",
        "SKIPPED": "건너뜀",
        "MISSING REQUIRED COLUMNS": "필수 열이 없습니다",
    },
    "ar_SA": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "استيراد بيانات المستخدم من EXCEL أو CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "استيراد الموردين من EXCEL أو CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "استيراد القطع من EXCEL أو CSV",
        "SKIPPED": "تم تخطيها",
        "MISSING REQUIRED COLUMNS": "الأعمدة المطلوبة مفقودة",
    },
    "hi_IN": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "EXCEL या CSV से उपयोगकर्ता डेटा आयात करें",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "EXCEL या CSV से आपूर्तिकर्ता आयात करें",
        "IMPORT PARTS FROM EXCEL OR CSV": "EXCEL या CSV से पार्ट्स आयात करें",
        "SKIPPED": "छोड़ी गईं",
        "MISSING REQUIRED COLUMNS": "आवश्यक कॉलम मौजूद नहीं हैं",
    },
    "bn_BD": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "EXCEL বা CSV থেকে ব্যবহারকারীর ডেটা আমদানি করুন",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "EXCEL বা CSV থেকে সরবরাহকারী আমদানি করুন",
        "IMPORT PARTS FROM EXCEL OR CSV": "EXCEL বা CSV থেকে পার্টস আমদানি করুন",
        "SKIPPED": "এড়ানো হয়েছে",
        "MISSING REQUIRED COLUMNS": "প্রয়োজনীয় কলাম অনুপস্থিত",
    },
    "ru_RU": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "ИМПОРТ ДАННЫХ ПОЛЬЗОВАТЕЛЕЙ ИЗ EXCEL ИЛИ CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "ИМПОРТ ПОСТАВЩИКОВ ИЗ EXCEL ИЛИ CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "ИМПОРТ ЗАПЧАСТЕЙ ИЗ EXCEL ИЛИ CSV",
        "SKIPPED": "ПРОПУЩЕНО",
        "MISSING REQUIRED COLUMNS": "ОТСУТСТВУЮТ ОБЯЗАТЕЛЬНЫЕ СТОЛБЦЫ",
    },
    "tr_TR": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "EXCEL VEYA CSV'DEN KULLANICI VERİSİ İÇE AKTAR",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "EXCEL VEYA CSV'DEN TEDARİKÇİLERİ İÇE AKTAR",
        "IMPORT PARTS FROM EXCEL OR CSV": "EXCEL VEYA CSV'DEN PARÇALARI İÇE AKTAR",
        "SKIPPED": "ATLANDI",
        "MISSING REQUIRED COLUMNS": "GEREKLİ SÜTUNLAR EKSİK",
    },
    "vi_VN": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "NHẬP DỮ LIỆU NGƯỜI DÙNG TỪ EXCEL HOẶC CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "NHẬP NHÀ CUNG CẤP TỪ EXCEL HOẶC CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "NHẬP LINH KIỆN TỪ EXCEL HOẶC CSV",
        "SKIPPED": "ĐÃ BỎ QUA",
        "MISSING REQUIRED COLUMNS": "THIẾU CỘT BẮT BUỘC",
    },
    "th_TH": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "นำเข้าข้อมูลผู้ใช้จาก EXCEL หรือ CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "นำเข้าซัพพลายเออร์จาก EXCEL หรือ CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "นำเข้าอะไหล่จาก EXCEL หรือ CSV",
        "SKIPPED": "ข้ามแล้ว",
        "MISSING REQUIRED COLUMNS": "ไม่มีคอลัมน์ที่จำเป็น",
    },
    "ms_MY": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "IMPORT DATA PENGGUNA DARI EXCEL ATAU CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "IMPORT PEMBEKAL DARI EXCEL ATAU CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "IMPORT ALAT GANTI DARI EXCEL ATAU CSV",
        "SKIPPED": "DILANGKAU",
        "MISSING REQUIRED COLUMNS": "LAJUR WAJIB TIADA",
    },
    "fil_PH": {
        "IMPORT USER DATA FROM EXCEL OR CSV": "MAG-IMPORT NG USER DATA MULA SA EXCEL O CSV",
        "IMPORT SUPPLIERS FROM EXCEL OR CSV": "MAG-IMPORT NG MGA SUPPLIER MULA SA EXCEL O CSV",
        "IMPORT PARTS FROM EXCEL OR CSV": "MAG-IMPORT NG MGA PIYESA MULA SA EXCEL O CSV",
        "SKIPPED": "NILAKTAWAN",
        "MISSING REQUIRED COLUMNS": "KULANG ANG MGA KINAKAILANGANG COLUMN",
    },
}
for _code, _items in _IMPORT_DATA_FEATURE_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)


_COMBINED_IMPORT_FEATURE_TRANSLATIONS = {
    "id_ID": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPOR PEMASOK & SUKU CADANG DARI EXCEL ATAU CSV",
        "RECORD TYPE": "JENIS DATA",
    },
    "es_MX": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPORTAR PROVEEDORES Y PIEZAS DESDE EXCEL O CSV",
        "RECORD TYPE": "TIPO DE REGISTRO",
    },
    "fr_FR": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPORTER LES FOURNISSEURS ET LES PIÈCES DEPUIS EXCEL OU CSV",
        "RECORD TYPE": "TYPE D’ENREGISTREMENT",
    },
    "de_DE": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "LIEFERANTEN UND TEILE AUS EXCEL ODER CSV IMPORTIEREN",
        "RECORD TYPE": "DATENSATZTYP",
    },
    "pt_BR": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPORTAR FORNECEDORES E PEÇAS DO EXCEL OU CSV",
        "RECORD TYPE": "TIPO DE REGISTRO",
    },
    "it_IT": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPORTA FORNITORI E RICAMBI DA EXCEL O CSV",
        "RECORD TYPE": "TIPO DI RECORD",
    },
    "nl_NL": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "LEVERANCIERS EN ONDERDELEN IMPORTEREN UIT EXCEL OF CSV",
        "RECORD TYPE": "RECORDTYPE",
    },
    "zh_CN": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "从 EXCEL 或 CSV 导入供应商和配件",
        "RECORD TYPE": "记录类型",
    },
    "ja_JP": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "EXCEL または CSV から仕入先と部品をインポート",
        "RECORD TYPE": "レコード種別",
    },
    "ko_KR": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "EXCEL 또는 CSV에서 공급업체와 부품 가져오기",
        "RECORD TYPE": "레코드 유형",
    },
    "ar_SA": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "استيراد الموردين والقطع من EXCEL أو CSV",
        "RECORD TYPE": "نوع السجل",
    },
    "hi_IN": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "EXCEL या CSV से आपूर्तिकर्ता और पार्ट्स आयात करें",
        "RECORD TYPE": "रिकॉर्ड प्रकार",
    },
    "bn_BD": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "EXCEL বা CSV থেকে সরবরাহকারী ও পার্টস আমদানি করুন",
        "RECORD TYPE": "রেকর্ডের ধরন",
    },
    "ru_RU": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "ИМПОРТ ПОСТАВЩИКОВ И ЗАПЧАСТЕЙ ИЗ EXCEL ИЛИ CSV",
        "RECORD TYPE": "ТИП ЗАПИСИ",
    },
    "tr_TR": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "EXCEL VEYA CSV'DEN TEDARİKÇİLERİ VE PARÇALARI İÇE AKTAR",
        "RECORD TYPE": "KAYIT TÜRÜ",
    },
    "vi_VN": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "NHẬP NHÀ CUNG CẤP VÀ LINH KIỆN TỪ EXCEL HOẶC CSV",
        "RECORD TYPE": "LOẠI BẢN GHI",
    },
    "th_TH": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "นำเข้าซัพพลายเออร์และอะไหล่จาก EXCEL หรือ CSV",
        "RECORD TYPE": "ประเภทระเบียน",
    },
    "ms_MY": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "IMPORT PEMBEKAL DAN ALAT GANTI DARI EXCEL ATAU CSV",
        "RECORD TYPE": "JENIS REKOD",
    },
    "fil_PH": {
        "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV": "MAG-IMPORT NG MGA SUPPLIER AT PIYESA MULA SA EXCEL O CSV",
        "RECORD TYPE": "URI NG RECORD",
    },
}
for _code, _items in _COMBINED_IMPORT_FEATURE_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)


# Exact translations for the external Dummy Creator menu action. Keeping this
# as a complete phrase avoids unnatural word order from token-by-token fallback.
_DUMMY_CREATOR_ACTION_TRANSLATIONS = {
    "id_ID": "BUKA PEMBUAT DATA DUMMY",
    "es_MX": "ABRIR CREADOR DE DATOS DE PRUEBA",
    "fr_FR": "OUVRIR LE CRÉATEUR DE DONNÉES FACTICES",
    "de_DE": "DUMMY-DATEN-ERSTELLER ÖFFNEN",
    "pt_BR": "ABRIR CRIADOR DE DADOS FICTÍCIOS",
    "it_IT": "APRI GENERATORE DI DATI DI PROVA",
    "nl_NL": "DUMMYGEGEVENSMAKER OPENEN",
    "zh_CN": "打开示例数据创建器",
    "ja_JP": "ダミーデータ作成ツールを開く",
    "ko_KR": "더미 데이터 생성기 열기",
    "ar_SA": "فتح منشئ البيانات التجريبية",
    "hi_IN": "डमी डेटा क्रिएटर खोलें",
    "bn_BD": "ডামি ডেটা ক্রিয়েটর খুলুন",
    "ru_RU": "ОТКРЫТЬ ГЕНЕРАТОР ТЕСТОВЫХ ДАННЫХ",
    "tr_TR": "TEST VERİSİ OLUŞTURUCUSUNU AÇ",
    "vi_VN": "MỞ TRÌNH TẠO DỮ LIỆU MẪU",
    "th_TH": "เปิดเครื่องมือสร้างข้อมูลตัวอย่าง",
    "ms_MY": "BUKA PENCIPTA DATA CONTOH",
    "fil_PH": "BUKSAN ANG TAGAGAWA NG DUMMY DATA",
}
for _code, _translated_text in _DUMMY_CREATOR_ACTION_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {})["OPEN DUMMY CREATOR"] = _translated_text


# Build exact menu/action phrases from the current per-language word tables so
# compound actions do not fall back to English words.
for _code in LANGUAGE_OPTIONS:
    if _code == DEFAULT_LANGUAGE_CODE:
        continue
    _t = UI_TRANSLATIONS.setdefault(_code, {})
    def _w(key, table=_t):
        return table.get(key, key)
    _t.update({
        "IMPORT EXCEL OR CSV": f"{_w('IMPORT')} EXCEL {_w('OR')} CSV",
        "EXPORT TO EXCEL": f"{_w('EXPORT')} {_w('TO')} EXCEL",
        "EXPORT SUPPLIERS TO EXCEL": f"{_w('EXPORT')} {_w('SUPPLIERS')} {_w('TO')} EXCEL",
        "EXPORT PARTS TO EXCEL": f"{_w('EXPORT')} {_w('PARTS')} {_w('TO')} EXCEL",
        "DATABASE PATH": f"{_w('DATABASE')} {_w('PATH')}",
        "DATABASE PATH SETTINGS": f"{_w('DATABASE')} {_w('PATH')} {_w('SETTINGS')}",
        "BACKUP FOLDER": f"{_w('BACKUP')} {_w('FOLDER')}",
        "SELECT FILE TO IMPORT": f"{_w('SELECT')} {_w('FILE')} {_w('TO')} {_w('IMPORT')}",
        "SELECT BACKUP FILE": f"{_w('SELECT')} {_w('BACKUP')} {_w('FILE')}",
        "SELECT DATABASE FILE": f"{_w('SELECT')} {_w('DATABASE')} {_w('FILE')}",
        "SELECT DATABASE FOLDER": f"{_w('SELECT')} {_w('DATABASE')} {_w('FOLDER')}",
        "SELECT FOLDER": f"{_w('SELECT')} {_w('FOLDER')}",
        "OPEN EXISTING DATABASE": f"{_w('OPEN')} {_w('EXISTING')} {_w('DATABASE')}",
        "CREATE NEW DATABASE": f"{_w('CREATE')} {_w('NEW')} {_w('DATABASE')}",
        "CREATE / REBUILD DUMMY.INI": f"{_w('CREATE')} / {_w('REBUILD')} {_w('DUMMY')}.INI",
        "CREATE DUMMY.INI": f"{_w('CREATE')} {_w('DUMMY')}.INI",
        "EDIT DUMMY DATA": f"{_w('EDIT')} {_w('DUMMY')} {_w('DATA')}",
        "DUMMY.INI SAVED": f"{_w('DUMMY')}.INI {_w('SAVED')}",
        "DELETE SELECTED": f"{_w('DELETE')} {_w('SELECTED')}",
        "TYPE NEW VALUE": f"{_w('TYPE')} {_w('NEW')} {_w('VALUE')}",
        "TYPE NEW VALUE...": f"{_w('TYPE')} {_w('NEW')} {_w('VALUE')}...",
        "SEND WARRANTY REMINDER": f"{_w('SEND')} {_w('WARRANTY')} {_w('REMINDER')}",
        "REMINDER MESSAGE (SMS/WHATSAPP)": f"{_w('REMINDER')} {_w('MESSAGE')} (SMS/WHATSAPP)",
        "COMPUTER SERVICE MANAGER v4.4": f"{_w('COMPUTER')} {_w('SERVICE')} {_w('MANAGER')} v4.4",
        "ALL RIGHTS RESERVED": f"{_w('ALL')} {_w('RIGHTS')} {_w('RESERVED')}",
    })
del _w


def get_currency_config(currency_code=None):
    code = str(currency_code or DEFAULT_CURRENCY_CODE).strip().upper()
    return CURRENCY_OPTIONS.get(code, CURRENCY_OPTIONS[DEFAULT_CURRENCY_CODE])


def get_language_config(language_code=None):
    code = str(language_code or DEFAULT_LANGUAGE_CODE).strip()
    return LANGUAGE_OPTIONS.get(code, LANGUAGE_OPTIONS[DEFAULT_LANGUAGE_CODE])


_UI_TRANSLATION_PHRASE_CACHE = {}
_UI_TRANSLATION_PROTECTED_RE = re.compile(
    r"(?:https?://[^\s<>]+|mailto:[^\s<>]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}|"
    r"[A-Za-z]:[\\/][^\n\r\t<>\"']+|(?:[A-Za-z0-9_.-]+\.(?:db|sqlite3?|xlsx?|csv|ini|exe|py|chm|hhp|ico))|"
    r"\[[A-Za-z_][A-Za-z0-9_ ]*\]|<[^>]+>)",
    flags=re.IGNORECASE,
)


def _protect_translation_fragments(text):
    """Mask paths, links, addresses, file names, placeholders, and HTML tags."""
    protected = []

    def _mask(match):
        marker = f"\uf000{len(protected)}\uf001"
        protected.append(match.group(0))
        return marker

    return _UI_TRANSLATION_PROTECTED_RE.sub(_mask, text), protected


def _restore_translation_fragments(text, protected):
    for index, value in enumerate(protected):
        text = text.replace(f"\uf000{index}\uf001", value)
    return text


def _translate_known_phrases(text, translations, language_code):
    """Replace the longest known English phrases inside a dynamic message."""
    cache_key = (language_code, id(translations), len(translations))
    candidates = _UI_TRANSLATION_PHRASE_CACHE.get(cache_key)
    if candidates is None:
        candidates = []
        for source, target in translations.items():
            source = str(source)
            if not target or source == str(target):
                continue
            if "\n" in source or "\r" in source:
                continue
            # Single words are handled by the token fallback below. Restricting
            # this pass to phrases avoids altering names and runtime values.
            if len(source.split()) < 2 and not source.endswith(":"):
                continue
            if not re.search(r"[A-Za-z]", source):
                continue
            candidates.append((source, str(target)))
        candidates.sort(key=lambda item: len(item[0]), reverse=True)
        _UI_TRANSLATION_PHRASE_CACHE.clear()
        _UI_TRANSLATION_PHRASE_CACHE[cache_key] = candidates

    result = text
    changed = False
    for source, target in candidates:
        pattern = re.escape(source)
        if source[:1].isalnum():
            pattern = r"(?<![\w])" + pattern
        if source[-1:].isalnum():
            pattern += r"(?![\w])"
        result, count = re.subn(pattern, lambda _match, value=target: value, result, flags=re.IGNORECASE)
        changed = changed or bool(count)
    return result, changed


def translate_ui_text(text, language_code=None):
    """Translate UI text using exact phrases, dynamic patterns, then word fallback.

    Exact and pattern matching happen before line-by-line processing so complete
    dialog messages remain natural instead of becoming literal word-for-word
    translations. Unknown values, file paths, names, and exception details are
    preserved unchanged.
    """
    if text is None:
        return text

    source_text = str(text)
    if not source_text.strip():
        return source_text

    code = str(language_code or DEFAULT_LANGUAGE_CODE).strip()
    translations = UI_TRANSLATIONS.get(code, {})
    if code == DEFAULT_LANGUAGE_CODE and not translations:
        return source_text

    # Prefer complete-message translations, including multi-line strings.
    exact = translations.get(source_text) or translations.get(source_text.upper())
    if exact is not None:
        return exact

    # Translate messages containing runtime values such as paths, counts, names,
    # or exception details while keeping those values intact.
    for pattern, replacement in globals().get("UI_TRANSLATION_PATTERNS", {}).get(code, []):
        if re.fullmatch(pattern, source_text, flags=re.IGNORECASE | re.DOTALL):
            return re.sub(pattern, replacement, source_text, count=1, flags=re.IGNORECASE | re.DOTALL)

    # Preserve report-heading decorations such as "=== TITLE ===" while
    # translating the title as one complete phrase.
    decorated = re.fullmatch(r"(\s*[=*_#~]{2,}\s*)(.*?)(\s*[=*_#~]{2,}\s*)", source_text)
    if decorated and decorated.group(2).strip():
        return f"{decorated.group(1)}{translate_ui_text(decorated.group(2).strip(), code)}{decorated.group(3)}"

    if "\n" in source_text:
        return "\n".join(translate_ui_text(line, code) for line in source_text.split("\n"))

    leading = source_text[:len(source_text) - len(source_text.lstrip())]
    trailing = source_text[len(source_text.rstrip()):]
    core = source_text.strip()
    colon = ":" if core.endswith(":") else ""
    if colon:
        core = core[:-1].strip()

    ellipsis = "..." if core.endswith("...") else ""
    if ellipsis:
        core = core[:-3].strip()

    # Preserve small icon/pictogram prefixes and translate only the label part.
    icon_prefix = ""
    parts = core.split(" ", 1)
    if len(parts) == 2 and not any(ch.isalnum() for ch in parts[0]):
        icon_prefix = parts[0] + " "
        core = parts[1].strip()

    upper_core = core.upper()
    translated = translations.get(core) or translations.get(upper_core)
    if translated is None:
        protected_core, protected_values = _protect_translation_fragments(core)
        phrase_text, phrase_changed = _translate_known_phrases(protected_core, translations, code)
        tokenized = re.split(r"(\s+|/|&|-|:|,|\.|!|\?|\(|\)|\[|\]|≤|<|>)", phrase_text)
        translated_parts = []
        token_changed = False
        for token in tokenized:
            lookup = translations.get(token) or translations.get(token.upper())
            if lookup is not None:
                translated_parts.append(str(lookup))
                token_changed = token_changed or str(lookup) != token
            else:
                translated_parts.append(token)
        translated = "".join(translated_parts) if (phrase_changed or token_changed) else protected_core
        translated = _restore_translation_fragments(translated, protected_values)

    return f"{leading}{icon_prefix}{translated}{ellipsis}{colon}{trailing}"

def format_number_value(value, language_code=None, decimals=0):
    cfg = get_language_config(language_code)
    return format_localized_number(
        value,
        thousands_sep=cfg.get("thousands_sep", ","),
        decimal_sep=cfg.get("decimal_sep", "."),
        decimals=decimals,
        grouping=cfg.get("grouping", "western"),
    )


def format_localized_number(value, thousands_sep=",", decimal_sep=".", decimals=2, grouping="western"):
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = 0.0

    negative = number < 0
    number = abs(number)
    decimals = max(0, int(decimals))
    base_text = f"{number:.{decimals}f}"

    if "." in base_text:
        integer_part, fractional_part = base_text.split(".", 1)
    else:
        integer_part, fractional_part = base_text, ""

    if grouping == "indian":
        if len(integer_part) > 3:
            last_three = integer_part[-3:]
            remaining = integer_part[:-3]
            grouped_chunks = []
            while len(remaining) > 2:
                grouped_chunks.insert(0, remaining[-2:])
                remaining = remaining[:-2]
            if remaining:
                grouped_chunks.insert(0, remaining)
            integer_part = thousands_sep.join(grouped_chunks + [last_three])
        else:
            pass
    else:
        integer_part = f"{int(integer_part):,}"
        if thousands_sep != ",":
            integer_part = integer_part.replace(",", thousands_sep)

    if decimals > 0:
        formatted = f"{integer_part}{decimal_sep}{fractional_part}"
    else:
        formatted = integer_part

    if negative:
        formatted = f"-{formatted}"
    return formatted


def format_currency_value(value, currency_code=None):
    cfg = get_currency_config(currency_code)
    number_text = format_localized_number(
        value,
        thousands_sep=cfg["thousands_sep"],
        decimal_sep=cfg["decimal_sep"],
        decimals=cfg["decimals"],
        grouping=cfg.get("grouping", "western"),
    )
    symbol = cfg["symbol"]
    spacer = " " if cfg.get("space_between") else ""

    if cfg.get("symbol_position", "prefix") == "suffix":
        return f"{number_text}{spacer}{symbol}"
    return f"{symbol}{spacer}{number_text}"


def parse_currency_value(text, currency_code=None):
    if text is None:
        return 0.0

    raw_text = str(text).strip()
    if not raw_text:
        return 0.0

    cfg = get_currency_config(currency_code)
    expected_decimals = int(cfg.get("decimals", 2))

    tokens = set()
    for code, data in CURRENCY_OPTIONS.items():
        tokens.add(code)
        tokens.add(code.lower())
        tokens.add(data.get("symbol", ""))
        tokens.add(str(data.get("symbol", "")).lower())
    tokens.update({" ", " "})

    cleaned = raw_text
    for token in sorted(tokens, key=len, reverse=True):
        if token:
            cleaned = cleaned.replace(token, "")

    cleaned = cleaned.replace("(", "-").replace(")", "")
    cleaned = "".join(ch for ch in cleaned if ch.isdigit() or ch in ".,-")

    if not cleaned or cleaned in {"-", ".", ","}:
        return 0.0

    if cleaned.count("-") > 1:
        cleaned = cleaned.replace("-", "")
    if "-" in cleaned and not cleaned.startswith("-"):
        cleaned = "-" + cleaned.replace("-", "")

    decimal_sep = None
    thousands_sep = None

    if "." in cleaned and "," in cleaned:
        decimal_sep = "." if cleaned.rfind(".") > cleaned.rfind(",") else ","
        thousands_sep = "," if decimal_sep == "." else "."
    elif "." in cleaned:
        if cleaned.count(".") > 1:
            thousands_sep = "."
        else:
            digits_after = len(cleaned.rsplit(".", 1)[1])
            if expected_decimals == 0:
                thousands_sep = "."
            elif digits_after <= expected_decimals:
                decimal_sep = "."
            elif digits_after == 3:
                thousands_sep = "."
            else:
                decimal_sep = "."
    elif "," in cleaned:
        if cleaned.count(",") > 1:
            thousands_sep = ","
        else:
            digits_after = len(cleaned.rsplit(",", 1)[1])
            if expected_decimals == 0:
                thousands_sep = ","
            elif digits_after <= expected_decimals:
                decimal_sep = ","
            elif digits_after == 3:
                thousands_sep = ","
            else:
                decimal_sep = ","

    normalized = cleaned
    if thousands_sep:
        normalized = normalized.replace(thousands_sep, "")
    if decimal_sep and decimal_sep != ".":
        normalized = normalized.replace(decimal_sep, ".")

    try:
        return float(normalized)
    except (TypeError, ValueError):
        fallback = "".join(ch for ch in normalized if ch.isdigit() or ch in ".-")
        try:
            return float(fallback)
        except (TypeError, ValueError):
            return 0.0


_LOCALIZATION_HOOKS_INSTALLED = False
_ORIGINAL_QMESSAGEBOX_STATIC = {}
_ORIGINAL_QMESSAGEBOX_INSTANCE = {}
_ORIGINAL_QFILEDIALOG_STATIC = {}
_ORIGINAL_QINPUTDIALOG_STATIC = {}
_ORIGINAL_QSTATUSBAR_SHOWMESSAGE = None


def _find_translation_owner(parent=None):
    checked = set()

    def walk(obj):
        while obj is not None and id(obj) not in checked:
            checked.add(id(obj))
            if hasattr(obj, "t") and callable(obj.t):
                return obj
            try:
                obj = obj.parent()
            except Exception:
                break
        return None

    owner = walk(parent)
    if owner is not None:
        return owner

    app = QApplication.instance()
    if app is None:
        return None

    for candidate in (app.activeModalWidget(), app.activeWindow(), app.focusWidget()):
        owner = walk(candidate)
        if owner is not None:
            return owner

    try:
        for widget in app.topLevelWidgets():
            owner = walk(widget)
            if owner is not None:
                return owner
    except Exception:
        pass
    return None


def _localize_runtime_text(parent, text):
    if text is None:
        return text
    owner = _find_translation_owner(parent)
    if owner is not None:
        try:
            return owner.t(text)
        except Exception:
            pass
    return translate_ui_text(text, DEFAULT_LANGUAGE_CODE)


def _localize_file_filter(parent, text):
    if not text:
        return text
    parts = []
    for chunk in str(text).split(";;"):
        if "(" in chunk:
            label, pattern = chunk.split("(", 1)
            parts.append(f"{_localize_runtime_text(parent, label.strip())} ({pattern}")
        else:
            parts.append(_localize_runtime_text(parent, chunk))
    return ";;".join(parts)


def install_localization_runtime_hooks():
    """Translate scattered Qt dialogs/status strings without rewriting every call site."""
    global _LOCALIZATION_HOOKS_INSTALLED
    global _ORIGINAL_QSTATUSBAR_SHOWMESSAGE
    if _LOCALIZATION_HOOKS_INSTALLED:
        return
    _LOCALIZATION_HOOKS_INSTALLED = True

    for name in ("information", "warning", "critical", "question"):
        original = getattr(QMessageBox, name)
        _ORIGINAL_QMESSAGEBOX_STATIC[name] = original

        def make_wrapper(method_name, method):
            def wrapper(parent, title, text, *args, **kwargs):
                return method(
                    parent,
                    _localize_runtime_text(parent, title),
                    _localize_runtime_text(parent, text),
                    *args,
                    **kwargs
                )
            return wrapper

        setattr(QMessageBox, name, staticmethod(make_wrapper(name, original)))

    for name in ("setWindowTitle", "setText", "setInformativeText", "setDetailedText"):
        original = getattr(QMessageBox, name)
        _ORIGINAL_QMESSAGEBOX_INSTANCE[name] = original

        def make_instance_wrapper(method):
            def wrapper(self, text, *args, **kwargs):
                return method(self, _localize_runtime_text(self, text), *args, **kwargs)
            return wrapper

        setattr(QMessageBox, name, make_instance_wrapper(original))

    original_add_button = QMessageBox.addButton
    _ORIGINAL_QMESSAGEBOX_INSTANCE["addButton"] = original_add_button

    def add_button_wrapper(self, *args):
        if args and isinstance(args[0], str):
            args = (_localize_runtime_text(self, args[0]),) + args[1:]
        return original_add_button(self, *args)

    QMessageBox.addButton = add_button_wrapper

    for name in ("getOpenFileName", "getSaveFileName", "getExistingDirectory"):
        original = getattr(QFileDialog, name)
        _ORIGINAL_QFILEDIALOG_STATIC[name] = original

        def make_file_wrapper(method_name, method):
            def wrapper(*args, **kwargs):
                args = list(args)
                parent = args[0] if args else kwargs.get("parent")
                if len(args) > 1:
                    args[1] = _localize_runtime_text(parent, args[1])
                elif "caption" in kwargs:
                    kwargs["caption"] = _localize_runtime_text(parent, kwargs["caption"])
                if method_name in {"getOpenFileName", "getSaveFileName"}:
                    if len(args) > 3:
                        args[3] = _localize_file_filter(parent, args[3])
                    elif "filter" in kwargs:
                        kwargs["filter"] = _localize_file_filter(parent, kwargs["filter"])
                return method(*args, **kwargs)
            return wrapper

        setattr(QFileDialog, name, staticmethod(make_file_wrapper(name, original)))

    original_get_item = QInputDialog.getItem
    _ORIGINAL_QINPUTDIALOG_STATIC["getItem"] = original_get_item

    def get_item_wrapper(parent, title, label, items, *args, **kwargs):
        return original_get_item(
            parent,
            _localize_runtime_text(parent, title),
            _localize_runtime_text(parent, label),
            items,
            *args,
            **kwargs
        )

    QInputDialog.getItem = staticmethod(get_item_wrapper)

    _ORIGINAL_QSTATUSBAR_SHOWMESSAGE = QStatusBar.showMessage

    def show_message_wrapper(self, message, *args, **kwargs):
        return _ORIGINAL_QSTATUSBAR_SHOWMESSAGE(self, _localize_runtime_text(self, message), *args, **kwargs)

    QStatusBar.showMessage = show_message_wrapper



# Final broad UI phrase coverage: exact strings that appear in labels, buttons,
# hints, dialogs, status messages, and dummy-data tools.
_FULL_UI_COVERAGE_TRANSLATIONS = {
    "id_ID": {
        "CREATED BY": "DIBUAT OLEH", "BY": "OLEH", "ERROR": "KESALAHAN", "EMAIL": "SUREL",
        "INVOICE": "FAKTUR", "NOMOR": "NOMOR", "NUMBER": "NOMOR", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "TAUTAN GOOGLE MAPS", "NOMOR INVOICE": "NOMOR FAKTUR",
        "STATUS": "STATUS", "EDIT": "UBAH", "EDIT DUMMY DATA": "UBAH DATA DUMMY",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "BUAT LAPORAN TERLEBIH DAHULU SEBELUM MENGEKSPOR",
        "SERVICE KOMPUTER PANGGILAN BOGOR": "LAYANAN KOMPUTER PANGGILAN BOGOR",
    },
    "es_MX": {
        "CREATED BY": "CREADO POR", "BY": "POR", "ERROR": "ERROR", "EMAIL": "CORREO ELECTRÓNICO",
        "INVOICE": "FACTURA", "NOMOR": "NÚMERO", "NUMBER": "NÚMERO", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "ENLACE DE GOOGLE MAPS", "NOMOR INVOICE": "NÚMERO DE FACTURA",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "GENERE EL INFORME PRIMERO ANTES DE EXPORTAR",
    },
    "fr_FR": {
        "CREATED BY": "CRÉÉ PAR", "BY": "PAR", "ERROR": "ERREUR", "EMAIL": "E-MAIL",
        "INVOICE": "FACTURE", "NOMOR": "NUMÉRO", "NUMBER": "NUMÉRO", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "LIEN GOOGLE MAPS", "NOMOR INVOICE": "NUMÉRO DE FACTURE",
        "ADDITIONAL NOTES": "NOTES SUPPLÉMENTAIRES", "APPLICATION CLOSED": "APPLICATION FERMÉE",
        "SERVICE DESCRIPTION": "DESCRIPTION DU SERVICE", "SERVICE KOMPUTER PANGGILAN BOGOR": "SERVICE INFORMATIQUE À DOMICILE À BOGOR",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "GÉNÉREZ D'ABORD LE RAPPORT AVANT D'EXPORTER",
    },
    "de_DE": {
        "CREATED BY": "ERSTELLT VON", "BY": "VON", "ERROR": "FEHLER", "EMAIL": "E-MAIL",
        "INVOICE": "RECHNUNG", "NOMOR": "NUMMER", "NUMBER": "NUMMER", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "GOOGLE-MAPS-LINK", "NOMOR INVOICE": "RECHNUNGSNUMMER",
        "STATUS": "STATUS", "SERVICE KOMPUTER PANGGILAN BOGOR": "COMPUTER-SERVICE VOR ORT IN BOGOR",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "ERST BERICHT ERSTELLEN, DANN EXPORTIEREN",
    },
    "pt_BR": {
        "CREATED BY": "CRIADO POR", "BY": "POR", "ERROR": "ERRO", "EMAIL": "E-MAIL",
        "INVOICE": "FATURA", "NOMOR": "NÚMERO", "NUMBER": "NÚMERO", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "LINK DO GOOGLE MAPS", "NOMOR INVOICE": "NÚMERO DA FATURA", "STATUS": "SITUAÇÃO",
        "BACKUP": "CÓPIA DE SEGURANÇA", "GENERATE REPORT FIRST BEFORE EXPORTING": "GERE O RELATÓRIO PRIMEIRO ANTES DE EXPORTAR",
    },
    "it_IT": {
        "CREATED BY": "CREATO DA", "BY": "DA", "ERROR": "ERRORE", "EMAIL": "E-MAIL",
        "INVOICE": "FATTURA", "NOMOR": "NUMERO", "NUMBER": "NUMERO", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "LINK GOOGLE MAPS", "NOMOR INVOICE": "NUMERO FATTURA", "BACKUP": "COPIA DI SICUREZZA",
        "FILE": "ARCHIVIO", "GENERATE REPORT FIRST BEFORE EXPORTING": "GENERA PRIMA IL REPORT PRIMA DI ESPORTARE",
    },
    "nl_NL": {
        "CREATED BY": "GEMAAKT DOOR", "BY": "DOOR", "ERROR": "FOUT", "EMAIL": "E-MAIL",
        "INVOICE": "FACTUUR", "NOMOR": "NUMMER", "NUMBER": "NUMMER", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "GOOGLE MAPS-LINK", "NOMOR INVOICE": "FACTUURNUMMER", "HELP": "HULP", "STATUS": "STATUS",
        "SERVICE KOMPUTER PANGGILAN BOGOR": "COMPUTERSERVICE AAN HUIS IN BOGOR",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "GENEREER EERST HET RAPPORT VOORDAT U EXPORTEERT",
    },
    "zh_CN": {
        "CREATED BY": "创建者", "BY": "由", "ERROR": "错误", "EMAIL": "电子邮件",
        "INVOICE": "发票", "NOMOR": "编号", "NUMBER": "编号", "GOOGLE MAPS": "谷歌地图",
        "LINK GOOGLE MAPS": "谷歌地图链接", "NOMOR INVOICE": "发票编号",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "请先生成报表再导出",
    },
    "ja_JP": {
        "CREATED BY": "作成者", "BY": "作成", "ERROR": "エラー", "EMAIL": "メール",
        "INVOICE": "請求書", "NOMOR": "番号", "NUMBER": "番号", "GOOGLE MAPS": "Google マップ",
        "LINK GOOGLE MAPS": "Google マップリンク", "NOMOR INVOICE": "請求書番号",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "エクスポートする前に先にレポートを生成してください",
    },
    "ko_KR": {
        "CREATED BY": "작성자", "BY": "작성", "ERROR": "오류", "EMAIL": "이메일",
        "INVOICE": "송장", "NOMOR": "번호", "NUMBER": "번호", "GOOGLE MAPS": "Google 지도",
        "LINK GOOGLE MAPS": "Google 지도 링크", "NOMOR INVOICE": "송장 번호",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "내보내기 전에 먼저 보고서를 생성하세요",
    },
    "ar_SA": {
        "CREATED BY": "تم الإنشاء بواسطة", "BY": "بواسطة", "ERROR": "خطأ", "EMAIL": "البريد الإلكتروني",
        "INVOICE": "فاتورة", "NOMOR": "رقم", "NUMBER": "رقم", "GOOGLE MAPS": "خرائط Google",
        "LINK GOOGLE MAPS": "رابط خرائط Google", "NOMOR INVOICE": "رقم الفاتورة",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "أنشئ التقرير أولاً قبل التصدير",
    },
    "hi_IN": {
        "CREATED BY": "द्वारा बनाया गया", "BY": "द्वारा", "ERROR": "त्रुटि", "EMAIL": "ईमेल",
        "INVOICE": "चालान", "NOMOR": "नंबर", "NUMBER": "नंबर", "GOOGLE MAPS": "Google मैप्स",
        "LINK GOOGLE MAPS": "Google मैप्स लिंक", "NOMOR INVOICE": "चालान नंबर",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "निर्यात करने से पहले रिपोर्ट बनाएं",
    },
    "bn_BD": {
        "CREATED BY": "তৈরি করেছেন", "BY": "দ্বারা", "ERROR": "ত্রুটি", "EMAIL": "ইমেইল",
        "INVOICE": "চালান", "NOMOR": "নম্বর", "NUMBER": "নম্বর", "GOOGLE MAPS": "Google Maps",
        "LINK GOOGLE MAPS": "Google Maps লিংক", "NOMOR INVOICE": "চালান নম্বর",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "রপ্তানির আগে প্রথমে রিপোর্ট তৈরি করুন",
    },
    "ru_RU": {
        "CREATED BY": "СОЗДАНО", "BY": "ОТ", "ERROR": "ОШИБКА", "EMAIL": "ЭЛ. ПОЧТА",
        "INVOICE": "СЧЕТ", "NOMOR": "НОМЕР", "NUMBER": "НОМЕР", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "ССЫЛКА GOOGLE MAPS", "NOMOR INVOICE": "НОМЕР СЧЕТА",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "СНАЧАЛА СОЗДАЙТЕ ОТЧЕТ ПЕРЕД ЭКСПОРТОМ",
    },
    "tr_TR": {
        "CREATED BY": "OLUŞTURAN", "BY": "TARAFINDAN", "ERROR": "HATA", "EMAIL": "E-POSTA",
        "INVOICE": "FATURA", "NOMOR": "NUMARA", "NUMBER": "NUMARA", "GOOGLE MAPS": "GOOGLE HARİTALAR",
        "LINK GOOGLE MAPS": "GOOGLE HARİTALAR BAĞLANTISI", "NOMOR INVOICE": "FATURA NUMARASI",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "DIŞA AKTARMADAN ÖNCE RAPOR OLUŞTURUN",
    },
    "vi_VN": {
        "CREATED BY": "TẠO BỞI", "BY": "BỞI", "ERROR": "LỖI", "EMAIL": "EMAIL",
        "INVOICE": "HÓA ĐƠN", "NOMOR": "SỐ", "NUMBER": "SỐ", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "LIÊN KẾT GOOGLE MAPS", "NOMOR INVOICE": "SỐ HÓA ĐƠN",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "HÃY TẠO BÁO CÁO TRƯỚC KHI XUẤT",
    },
    "th_TH": {
        "CREATED BY": "สร้างโดย", "BY": "โดย", "ERROR": "ข้อผิดพลาด", "EMAIL": "อีเมล",
        "INVOICE": "ใบแจ้งหนี้", "NOMOR": "หมายเลข", "NUMBER": "หมายเลข", "GOOGLE MAPS": "Google Maps",
        "LINK GOOGLE MAPS": "ลิงก์ Google Maps", "NOMOR INVOICE": "หมายเลขใบแจ้งหนี้",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "โปรดสร้างรายงานก่อนส่งออก",
    },
    "ms_MY": {
        "CREATED BY": "DICIPTA OLEH", "BY": "OLEH", "ERROR": "RALAT", "EMAIL": "E-MEL",
        "INVOICE": "INVOIS", "NOMOR": "NOMBOR", "NUMBER": "NOMBOR", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "PAUTAN GOOGLE MAPS", "NOMOR INVOICE": "NOMBOR INVOIS", "STATUS": "STATUS",
        "EDIT": "UBAH", "EDIT DUMMY DATA": "UBAH DATA DUMMY",
        "GENERATE REPORT FIRST BEFORE EXPORTING": "JANA LAPORAN DAHULU SEBELUM MENGEKSPORT",
    },
    "fil_PH": {
        "FILE": "TALAKASAN", "IMPORT": "MAG-IMPORT", "EXPORT": "MAG-EXPORT", "BACKUP": "KOPYANG PANG-SEGURIDAD",
        "SETTINGS": "MGA SETTING", "REPORTS": "MGA ULAT", "SUPPLIERS": "MGA SUPPLIER", "ADD": "MAGDAGDAG",
        "REFRESH": "I-REFRESH", "BROWSE": "MAG-BROWSE", "EDIT": "I-EDIT", "DELETE": "BURAHIN",
        "CREATED BY": "GINAWA NI", "BY": "NI", "ERROR": "ERROR", "EMAIL": "EMAIL",
        "INVOICE": "INVOICE", "NOMOR": "NUMERO", "NUMBER": "NUMERO", "GOOGLE MAPS": "GOOGLE MAPS",
        "LINK GOOGLE MAPS": "LINK NG GOOGLE MAPS", "NOMOR INVOICE": "NUMERO NG INVOICE", "STATUS": "KATAYUAN",
        "ADDRESS LENGKAP": "KUMPLETONG ADDRESS", "MODEL DEVICE": "MODEL NG DEVICE", "ADDITIONAL NOTES": "KARAGDAGANG TALA",
        "APPLICATION CLOSED": "NAISARA ANG APPLICATION", "DATABASE PATH": "PATH NG DATABASE", "DATABASE PATH SETTINGS": "SETTING NG PATH NG DATABASE",
        "DATABASE PATH:": "PATH NG DATABASE:", "BACKUP FOLDER:": "BACKUP FOLDER:", "BRAND:": "BRAND:",
        "CONFIRM DELETE": "KUMPIRMAHIN ANG PAGBURA", "EDIT DUMMY DATA": "I-EDIT ANG DUMMY DATA",
        "FIRST-TIME DATABASE SETUP": "UNANG SETUP NG DATABASE", "SUPPLIER ADDRESS": "ADDRESS NG SUPPLIER",
        "SUPPLIERS REFRESHED": "NA-REFRESH ANG MGA SUPPLIER", "GENERATE REPORT FIRST BEFORE EXPORTING": "GUMAWA MUNA NG REPORT BAGO MAG-EXPORT",
        "✏️ EDIT": "✏️ I-EDIT", "✏️ EDIT CONTACT": "✏️ I-EDIT ANG CONTACT", "❌ DELETE": "❌ BURAHIN", "➕ ADD": "➕ MAGDAGDAG",
        "➕ ADD PART": "➕ MAGDAGDAG NG PIYESA", "📄 TEXT VIEW": "📄 TEXT VIEW", "📋 TABLE VIEW": "📋 TABLE VIEW", "📤 EXPORT SELECTED": "📤 I-EXPORT ANG NAPILI",
        "🔃 REFRESH": "🔃 I-REFRESH", "🔄 REFRESH": "🔄 I-REFRESH", "🗑️ CLEAR": "🗑️ LINISIN",
    },
}
for _code, _items in _FULL_UI_COVERAGE_TRANSLATIONS.items():
    UI_TRANSLATIONS.setdefault(_code, {}).update(_items)


# German phrase-level corrections requested in TRANSLATE.txt.
# Exact phrase mappings are used so the UI does not fall back to mixed
# word-by-word translations such as "RECHNUNG NUMMER" or "GERÄT TYPES".
_GERMAN_UI_REFINEMENTS = {
    "INVOICE NUMBER": "RECHNUNGSNUMMER",
    "EXPIRY DATE FROM": "ABLAUFDATUM VON",
    "EXPIRY DATE FROM:": "ABLAUFDATUM VON:",
    "PART NAME": "TEILENAME",
    "DEVICE TYPES": "GERÄTE-TYPEN",
    "BRANDS": "MARKEN",
    "MODELS": "MODELLE",
    "ISSUES": "FEHLER",
    "STATUSES": "STATUS",
    "TECHNICIANS": "TECHNIKER",
    "WARRANTY PERIODS": "GARANTIE-ZEITEN",
    "PART TYPES": "TEILE-TYPEN",
}
UI_TRANSLATIONS.setdefault("de_DE", {}).update(_GERMAN_UI_REFINEMENTS)


# Indonesian phrase-level corrections. These override literal word-by-word
# results with clear, natural interface language and cover all primary screens,
# dialogs, report headings, filters, tooltips, and common status values.
_INDONESIAN_UI_REFINEMENTS = {
    # Main window, menus, tabs, and common controls
    "BACKUP DATABASE": "CADANGKAN DATABASE",
    "RESTORE DATABASE": "PULIHKAN DATABASE",
    "RELOAD DROPDOWN DATA": "MUAT ULANG DATA PILIHAN",
    "CUSTOMER NAME": "NAMA PELANGGAN",
    "PHONE NUMBER": "NOMOR TELEPON",
    "FULL ADDRESS": "ALAMAT LENGKAP",
    "DEVICE MODEL": "MODEL PERANGKAT",
    "INVOICE NUMBER": "NOMOR FAKTUR",
    "INVOICE": "FAKTUR",
    "EDIT": "UBAH",
    "EDIT CONTACT": "UBAH KONTAK",
    "✏️ EDIT": "✏️ UBAH",
    "✏️ EDIT CONTACT": "✏️ UBAH KONTAK",
    "EDIT MODE": "MODE UBAH",
    "REFRESH": "MUAT ULANG",
    "BROWSE": "TELUSURI",
    "USE DEFAULT": "GUNAKAN BAWAAN",
    "AUTO BACKUP ON EXIT": "CADANGKAN OTOMATIS SAAT KELUAR",
    "WARRANTY FILTER": "FILTER GARANSI",
    "FILTER PARTS": "FILTER SUKU CADANG",
    "TABLE VIEW": "TAMPILAN TABEL",
    "TEXT VIEW": "TAMPILAN TEKS",
    "CURRENT DATE": "TANGGAL HARI INI",
    "SET SERVICE DATE TO TODAY": "ATUR TANGGAL SERVIS KE HARI INI",
    "RESET WARRANTY RANGE FROM TODAY": "ATUR ULANG RENTANG GARANSI MULAI HARI INI",
    "COPY": "SALIN",
    "COPIED": "TERSALIN",
    "READY": "SIAP",

    # First-run database setup and path settings
    "FIRST-TIME DATABASE SETUP": "PENGATURAN DATABASE PERTAMA KALI",
    "WHAT DO YOU WANT TO DO WITH THE DATABASE?": "APA YANG INGIN ANDA LAKUKAN PADA DATABASE?",
    "OPEN EXISTING DATABASE": "BUKA DATABASE YANG SUDAH ADA",
    "CREATE NEW DATABASE": "BUAT DATABASE BARU",
    "NO DATABASE SELECTED": "TIDAK ADA DATABASE YANG DIPILIH",
    "NO EXISTING DATABASE WAS SELECTED. THE APPLICATION WILL CONTINUE WITH NEW DATABASE CREATION.": "TIDAK ADA DATABASE LAMA YANG DIPILIH. APLIKASI AKAN MELANJUTKAN PEMBUATAN DATABASE BARU.",
    "DATABASE LOCATION REQUIRED": "LOKASI DATABASE DIPERLUKAN",
    "DRIVE D: WAS NOT FOUND.\nPLEASE CREATE OR SELECT A FOLDER FOR THE DATABASE.": "DRIVE D: TIDAK DITEMUKAN.\nBUAT ATAU PILIH FOLDER UNTUK MENYIMPAN DATABASE.",
    "SET DATABASE LOCATION": "ATUR LOKASI DATABASE",
    "FIRST-TIME SETUP:\nPLEASE CHOOSE A FOLDER TO SAVE THE NEW DATABASE.": "PENGATURAN PERTAMA KALI:\nPILIH FOLDER UNTUK MENYIMPAN DATABASE BARU.",
    "SELECT DATABASE FOLDER": "PILIH FOLDER DATABASE",
    "DATABASE PATH SETTINGS": "PENGATURAN LOKASI DATABASE",
    "DATABASE PATH:": "LOKASI DATABASE:",
    "BACKUP FOLDER:": "FOLDER CADANGAN:",
    "IMPORT/EXPORT FOLDER:": "FOLDER IMPOR/EKSPOR:",
    "USE D:\\BACKUP (DEFAULT)": "GUNAKAN D:\\BACKUP (BAWAAN)",
    "DRIVE NOT FOUND": "DRIVE TIDAK DITEMUKAN",
    "DRIVE D: IS NOT AVAILABLE. PLEASE CHOOSE ANOTHER FOLDER.": "DRIVE D: TIDAK TERSEDIA. PILIH FOLDER LAIN.",
    "SELECT DATABASE FILE": "PILIH BERKAS DATABASE",
    "SELECT FOLDER": "PILIH FOLDER",
    "INVALID PATH": "LOKASI TIDAK VALID",
    "DATABASE PATH CANNOT BE EMPTY": "LOKASI DATABASE TIDAK BOLEH KOSONG",
    "BACKUP FOLDER CANNOT BE EMPTY": "FOLDER CADANGAN TIDAK BOLEH KOSONG",
    "IMPORT/EXPORT FOLDER CANNOT BE EMPTY": "FOLDER IMPOR/EKSPOR TIDAK BOLEH KOSONG",
    "INVALID DRIVE": "DRIVE TIDAK VALID",
    "THE DRIVE FOR DATABASE PATH DOES NOT EXIST.": "DRIVE UNTUK LOKASI DATABASE TIDAK TERSEDIA.",
    "THE DRIVE FOR BACKUP FOLDER DOES NOT EXIST.": "DRIVE UNTUK FOLDER CADANGAN TIDAK TERSEDIA.",
    "THE DRIVE FOR IMPORT/EXPORT FOLDER DOES NOT EXIST.": "DRIVE UNTUK FOLDER IMPOR/EKSPOR TIDAK TERSEDIA.",
    "SETTINGS SAVED": "PENGATURAN TERSIMPAN",
    "DATABASE PATHS UPDATED SUCCESSFULLY!": "LOKASI DATABASE BERHASIL DIPERBARUI!",

    # Import, export, backup, and restore dialogs
    "UNSUPPORTED FORMAT": "FORMAT TIDAK DIDUKUNG",
    "PLEASE SELECT EXCEL OR CSV FILE": "PILIH BERKAS EXCEL ATAU CSV.",
    "IMPORT FAILED": "IMPOR GAGAL",
    "EXPORT FAILED": "EKSPOR GAGAL",
    "BACKUP SUCCESSFUL": "PENCADANGAN BERHASIL",
    "BACKUP FAILED": "PENCADANGAN GAGAL",
    "CONFIRM RESTORE": "KONFIRMASI PEMULIHAN",
    "WARNING: RESTORING WILL REPLACE ALL CURRENT DATA WITH BACKUP DATA.\nTHIS ACTION CANNOT BE UNDONE.\n\nDO YOU WANT TO CONTINUE?": "PERINGATAN: PEMULIHAN AKAN MENGGANTI SEMUA DATA SAAT INI DENGAN DATA CADANGAN.\nTINDAKAN INI TIDAK DAPAT DIBATALKAN.\n\nAPAKAH ANDA INGIN MELANJUTKAN?",
    "SELECT BACKUP FILE": "PILIH BERKAS CADANGAN",
    "RESTORE SUCCESSFUL": "PEMULIHAN BERHASIL",
    "DATABASE RESTORED SUCCESSFULLY FROM BACKUP!": "DATABASE BERHASIL DIPULIHKAN DARI CADANGAN!",
    "DATABASE RESTORED FROM BACKUP": "DATABASE DIPULIHKAN DARI CADANGAN",
    "RESTORE FAILED": "PEMULIHAN GAGAL",
    "IMPORT COMPLETE": "IMPOR SELESAI",
    "EXPORT SUCCESSFUL": "EKSPOR BERHASIL",
    "NO SELECTION": "BELUM ADA PILIHAN",
    "SELECT ROWS TO EXPORT": "PILIH BARIS YANG AKAN DIEKSPOR",
    "NO DATA": "TIDAK ADA DATA",
    "GENERATE REPORT FIRST BEFORE EXPORTING": "BUAT LAPORAN TERLEBIH DAHULU SEBELUM MENGEKSPOR",
    "REPORT EXPORTED": "LAPORAN BERHASIL DIEKSPOR",
    "REPORT EXPORT FAILED": "EKSPOR LAPORAN GAGAL",
    "EXPORT COMPLETE": "EKSPOR SELESAI",
    "EXPORT ERROR": "KESALAHAN EKSPOR",
    "SAVE EXPORT FILE": "SIMPAN BERKAS EKSPOR",
    "SAVE REPORT EXPORT": "SIMPAN EKSPOR LAPORAN",
    "EXCEL FILES": "BERKAS EXCEL",
    "CSV FILES": "BERKAS CSV",
    "DATA HAS BEEN EXPORTED TO:": "DATA BERHASIL DIEKSPOR KE:",
    "THE DRIVE FOR EXPORT PATH DOES NOT EXIST.": "DRIVE UNTUK LOKASI EKSPOR TIDAK TERSEDIA.",
    "SELECT DEFAULT CURRENCY FOR THE APPLICATION:": "PILIH MATA UANG BAWAAN APLIKASI:",
    "SELECT APPLICATION CURRENCY:": "PILIH MATA UANG APLIKASI:",

    # Validation and record operations
    "ERROR": "KESALAHAN",
    "NAME IS REQUIRED": "NAMA WAJIB DIISI",
    "SUPPLIER NAME IS REQUIRED": "NAMA PEMASOK WAJIB DIISI",
    "PART NAME IS REQUIRED": "NAMA SUKU CADANG WAJIB DIISI",
    "FORM CLEARED": "FORMULIR DIBERSIHKAN",
    "SELECT A CONTACT TO EDIT": "PILIH KONTAK YANG AKAN DIUBAH",
    "EDIT MODE: UPDATE FIELDS THEN CLICK SAVE": "MODE UBAH: PERBARUI DATA, LALU KLIK SIMPAN",
    "CONFIRM DELETE": "KONFIRMASI PENGHAPUSAN",
    "SELECT A CONTACT TO DELETE": "PILIH KONTAK YANG AKAN DIHAPUS",
    "SELECT A SUPPLIER TO EDIT": "PILIH PEMASOK YANG AKAN DIUBAH",
    "EDIT MODE: UPDATE FIELDS THEN CLICK UPDATE": "MODE UBAH: PERBARUI DATA, LALU KLIK PERBARUI",
    "SELECT A SUPPLIER TO UPDATE": "PILIH PEMASOK YANG AKAN DIPERBARUI",
    "CANNOT DETERMINE SELECTED SUPPLIER": "PEMASOK YANG DIPILIH TIDAK DAPAT DITENTUKAN",
    "SELECT A SUPPLIER TO DELETE": "PILIH PEMASOK YANG AKAN DIHAPUS",
    "SELECT A PART TO EDIT": "PILIH SUKU CADANG YANG AKAN DIUBAH",
    "SELECT A PART TO UPDATE": "PILIH SUKU CADANG YANG AKAN DIPERBARUI",
    "SELECT A PART TO DELETE": "PILIH SUKU CADANG YANG AKAN DIHAPUS",
    "SUPPLIERS REFRESHED": "DATA PEMASOK DIMUAT ULANG",
    "PARTS REFRESHED": "DATA SUKU CADANG DIMUAT ULANG",

    # Warranty screen and reminders
    "SELECT A WARRANTY ROW TO EDIT": "PILIH BARIS GARANSI YANG AKAN DIUBAH",
    "NOT FOUND": "TIDAK DITEMUKAN",
    "SELECT WARRANTIES TO SEND REMINDERS": "PILIH DATA GARANSI YANG AKAN DIKIRIMI PENGINGAT",
    "SEND WARRANTY REMINDER": "KIRIM PENGINGAT GARANSI",
    "REMINDER MESSAGE (SMS/WHATSAPP)": "PESAN PENGINGAT (SMS/WHATSAPP)",
    "SEND REMINDERS": "KIRIM PENGINGAT",
    "REMINDERS SENT": "PENGINGAT TERKIRIM",
    "CUSTOMERS SELECTED FOR REMINDER": "PELANGGAN TERPILIH UNTUK PENGINGAT",
    "CUSTOMERS SELECTED FOR REMINDER:": "PELANGGAN TERPILIH UNTUK PENGINGAT:",
    "TOTAL SELECTED": "TOTAL TERPILIH",
    "NAME:": "NAMA:",
    "PHONE:": "TELEPON:",
    "WARRANTY STATUS:": "STATUS GARANSI:",
    "END DATE:": "TANGGAL BERAKHIR:",
    "DEAR [CUSTOMER],": "YTH. [PELANGGAN],",
    "BOGOR ON-CALL COMPUTER SERVICE REMINDS YOU THAT YOUR DEVICE REPAIR WARRANTY WILL EXPIRE ON:": "LAYANAN KOMPUTER PANGGILAN BOGOR MENGINGATKAN BAHWA GARANSI PERBAIKAN PERANGKAT ANDA AKAN BERAKHIR PADA:",
    "PLEASE CONTACT US IF YOU NEED FURTHER SERVICE.": "HUBUNGI KAMI APABILA ANDA MEMERLUKAN LAYANAN LEBIH LANJUT.",
    "THANK YOU,": "TERIMA KASIH,",
    "ACTIVE": "AKTIF",
    "EXPIRED": "KEDALUWARSA",
    "EXPIRING SOON": "SEGERA BERAKHIR",
    "EXPIRING SOON (< 7 DAYS)": "SEGERA BERAKHIR (< 7 HARI)",
    "EXPIRING SOON (≤ 7 DAYS)": "SEGERA BERAKHIR (≤ 7 HARI)",
    "EXPIRING SOON (≤ 30 DAYS)": "SEGERA BERAKHIR (≤ 30 HARI)",
    "EXPIRING SOON (≤30 days)": "SEGERA BERAKHIR (≤ 30 HARI)",
    "NO WARRANTY": "TANPA GARANSI",
    "N/A": "TIDAK TERSEDIA",
    "SOON": "SEGERA",
    "WARNING": "PERINGATAN",

    # Search and system combo choices
    "ALL": "SEMUA",
    "ALL FIELDS": "SEMUA KOLOM",
    "SERVICE SUMMARY": "RINGKASAN SERVIS",
    "FINANCIAL REPORT": "LAPORAN KEUANGAN",
    "TECHNICIAN PERFORMANCE": "KINERJA TEKNISI",
    "DEVICE BRAND STATISTICS": "STATISTIK MEREK PERANGKAT",
    "WARRANTY SUMMARY": "RINGKASAN GARANSI",

    # Report tables and report text
    "TOTAL SERVICES": "TOTAL SERVIS",
    "COMPLETED": "SELESAI",
    "IN REPAIR": "DALAM PERBAIKAN",
    "WAITING PARTS": "MENUNGGU SUKU CADANG",
    "WAITING FOR PARTS": "MENUNGGU SUKU CADANG",
    "TOTAL REVENUE": "TOTAL PENDAPATAN",
    "AVG REVENUE/SERVICE": "RATA-RATA PENDAPATAN/SERVIS",
    "IN PROGRESS": "SEDANG DIPROSES",
    "AVG REVENUE": "RATA-RATA PENDAPATAN",
    "TOTAL DEVICES": "TOTAL PERANGKAT",
    "UNIQUE MODELS": "MODEL UNIK",
    "REPAIRED": "SELESAI DIPERBAIKI",
    "AVG REPAIR COST": "RATA-RATA BIAYA PERBAIKAN",
    "TOTAL CUSTOMERS": "TOTAL PELANGGAN",
    "SERVICE SUMMARY REPORT": "LAPORAN RINGKASAN SERVIS",
    "TECHNICIAN PERFORMANCE REPORT": "LAPORAN KINERJA TEKNISI",
    "WARRANTY SUMMARY REPORT": "LAPORAN RINGKASAN GARANSI",
    "PERIOD": "PERIODE",
    "GENERATED": "DIBUAT",
    "TOTAL ACTIVE WARRANTIES": "TOTAL GARANSI AKTIF",
    "TOTAL EXPIRING SOON": "TOTAL SEGERA BERAKHIR",
    "TOTAL EXPIRED": "TOTAL KEDALUWARSA",
    "AVG PER SERVICE": "RATA-RATA PER SERVIS",
    "REVENUE": "PENDAPATAN",
    "REPAIR COST": "BIAYA PERBAIKAN",

    # About dialog
    "COMPUTER SERVICE MANAGER v4.4": "COMPUTER SERVICE MANAGER v4.4",
    "CREATED BY:": "DIBUAT OLEH:",
    "ALL RIGHTS RESERVED": "SEMUA HAK DILINDUNGI",
    "⚠️ WARNING!": "⚠️ PERINGATAN!",
    "THIS SOFTWARE IS PROVIDED AS-IS": "PERANGKAT LUNAK INI DISEDIAKAN APA ADANYA",
    "NO WARRANTY OF ANY KIND IS PROVIDED": "TIDAK ADA JAMINAN DALAM BENTUK APA PUN",
}
UI_TRANSLATIONS.setdefault("id_ID", {}).update(_INDONESIAN_UI_REFINEMENTS)

# Complete Indonesian coverage for the dashboard, authentication, account
# management, session identity, and role-based access messages.
_INDONESIAN_SECURITY_DASHBOARD_TRANSLATIONS = {
    "ACCOUNT": "AKUN",
    "SECURE LOGIN": "LOGIN AMAN",
    "SIGN IN TO CONTINUE": "MASUK UNTUK MELANJUTKAN",
    "LOGIN": "MASUK",
    "LOGOUT": "KELUAR DARI AKUN",
    "USERNAME": "NAMA PENGGUNA",
    "PASSWORD": "KATA SANDI",
    "CURRENT PASSWORD": "KATA SANDI SAAT INI",
    "NEW PASSWORD": "KATA SANDI BARU",
    "CONFIRM PASSWORD": "KONFIRMASI KATA SANDI",
    "SHOW PASSWORD": "TAMPILKAN KATA SANDI",
    "SHOW PASSWORDS": "TAMPILKAN KATA SANDI",
    "DISPLAY NAME": "NAMA TAMPILAN",
    "ROLE": "PERAN",
    "ADMIN": "ADMIN",
    "ADMINISTRATOR": "ADMINISTRATOR",
    "STAFF": "STAF",
    "CHANGE PASSWORD": "UBAH KATA SANDI",
    "CHANGE PASSWORD FOR": "UBAH KATA SANDI UNTUK",
    "USER MANAGEMENT": "MANAJEMEN PENGGUNA",
    "ADD USER": "TAMBAH PENGGUNA",
    "ADD APPLICATION USER": "TAMBAH PENGGUNA APLIKASI",
    "RESET PASSWORD": "ATUR ULANG KATA SANDI",
    "RESET USER PASSWORD": "ATUR ULANG KATA SANDI PENGGUNA",
    "ENABLE / DISABLE": "AKTIFKAN / NONAKTIFKAN",
    "UNLOCK": "BUKA KUNCI",
    "LOCKED UNTIL": "TERKUNCI SAMPAI",
    "LAST LOGIN": "LOGIN TERAKHIR",
    "FAILED": "GAGAL",
    "CREATED": "DIBUAT",
    "SIGNED IN": "SEDANG MASUK",
    "SIGNED IN AS": "MASUK SEBAGAI",
    "SECURE ADMINISTRATOR SETUP": "PENGATURAN ADMINISTRATOR AMAN",
    "CREATE THE FIRST ADMINISTRATOR ACCOUNT": "BUAT AKUN ADMINISTRATOR PERTAMA",
    "CREATE ADMINISTRATOR": "BUAT ADMINISTRATOR",
    "NO DEFAULT PASSWORD IS USED. PASSWORDS ARE STORED ONLY AS UNIQUE, SALTED SCRYPT HASHES. KEEP THIS ADMINISTRATOR PASSWORD SAFE.": "TIDAK ADA KATA SANDI BAWAAN. KATA SANDI HANYA DISIMPAN SEBAGAI HASH SCRYPT UNIK DENGAN SALT. SIMPAN KATA SANDI ADMINISTRATOR INI DENGAN AMAN.",
    "MINIMUM 4 CHARACTERS. ANY LETTER, NUMBER, SPACE, OR SYMBOL FORMAT IS ALLOWED.": "MINIMAL 4 KARAKTER. HURUF, ANGKA, SPASI, DAN SIMBOL DIPERBOLEHKAN.",
    "ADMINISTRATOR REQUIRED": "MEMERLUKAN ADMINISTRATOR",
    "ADMINISTRATOR ONLY": "KHUSUS ADMINISTRATOR",
    "THIS ACTION IS RESTRICTED TO ADMINISTRATOR ACCOUNTS.": "TINDAKAN INI HANYA DAPAT DILAKUKAN OLEH AKUN ADMINISTRATOR.",
    "STAFF ACCESS IS LIMITED": "AKSES STAF DIBATASI",
    "STAFF CAN ADD NEW SERVICE RECORDS AND VIEW OPERATIONAL DATA. EDITING, DELETING, IMPORT, EXPORT, BACKUP, ADMINISTRATIVE SETTINGS, USER MANAGEMENT, SUPPLIER CHANGES, PART CHANGES, REPORTS, AND REMINDERS REQUIRE AN ADMINISTRATOR.": "STAF DAPAT MENAMBAH DATA SERVIS BARU DAN MELIHAT DATA OPERASIONAL. PENGUBAHAN, PENGHAPUSAN, IMPOR, EKSPOR, PENCADANGAN, PENGATURAN ADMINISTRATIF, MANAJEMEN PENGGUNA, PERUBAHAN PEMASOK, PERUBAHAN SUKU CADANG, LAPORAN, DAN PENGINGAT MEMERLUKAN ADMINISTRATOR.",
    "DASHBOARD REVIEW": "RINGKASAN DASBOR",
    "REFRESH DASHBOARD": "MUAT ULANG DASBOR",
    "TOTAL SERVICE RECORDS": "TOTAL DATA SERVIS",
    "TOTAL SERVICE VALUE": "TOTAL NILAI SERVIS",
    "ACTIVE WARRANTIES": "GARANSI AKTIF",
    "EXPIRING IN 30 DAYS": "BERAKHIR DALAM 30 HARI",
    "LOW-STOCK PARTS": "SUKU CADANG STOK RENDAH",
    "RECENT SERVICE RECORDS": "DATA SERVIS TERBARU",
    "SERVICE STATUS GRAPHIC": "GRAFIK STATUS SERVIS",
    "RECENT SECURITY AND DATA ACTIVITY": "AKTIVITAS KEAMANAN DAN DATA TERBARU",
    "MY RECENT ACTIVITY": "AKTIVITAS TERBARU SAYA",
    "COUNT": "JUMLAH",
    "TIME": "WAKTU",
    "USER": "PENGGUNA",
    "ACTION": "TINDAKAN",
    "ITEM": "ITEM",
    "DETAILS": "RINCIAN",
    "NO STATUS DATA": "BELUM ADA DATA STATUS",
    "INVALID USERNAME": "NAMA PENGGUNA TIDAK VALID",
    "USERNAME MUST BE 3-32 CHARACTERS USING LETTERS, NUMBERS, DOT, DASH, OR UNDERSCORE.": "NAMA PENGGUNA HARUS 3-32 KARAKTER DAN HANYA BOLEH BERISI HURUF, ANGKA, TITIK, TANDA HUBUNG, ATAU GARIS BAWAH.",
    "WEAK PASSWORD": "KATA SANDI TERLALU LEMAH",
    "PASSWORD MUST CONTAIN AT LEAST 4 CHARACTERS.": "KATA SANDI HARUS BERISI MINIMAL 4 KARAKTER.",
    "PASSWORD IS TOO LONG.": "KATA SANDI TERLALU PANJANG.",
    "PASSWORD MISMATCH": "KATA SANDI TIDAK SAMA",
    "THE PASSWORDS DO NOT MATCH.": "KATA SANDI DAN KONFIRMASINYA TIDAK SAMA.",
    "USERNAME EXISTS": "NAMA PENGGUNA SUDAH ADA",
    "THAT USERNAME ALREADY EXISTS.": "NAMA PENGGUNA TERSEBUT SUDAH TERDAFTAR.",
    "SETUP FAILED": "PENGATURAN GAGAL",
    "ENTER BOTH USERNAME AND PASSWORD.": "MASUKKAN NAMA PENGGUNA DAN KATA SANDI.",
    "INVALID USERNAME OR PASSWORD.": "NAMA PENGGUNA ATAU KATA SANDI SALAH.",
    "THIS ACCOUNT IS DISABLED. CONTACT AN ADMINISTRATOR.": "AKUN INI DINONAKTIFKAN. HUBUNGI ADMINISTRATOR.",
    "LOGIN SUCCESSFUL.": "LOGIN BERHASIL.",
    "PASSWORD NOT CHANGED": "KATA SANDI TIDAK DIUBAH",
    "PASSWORD CHANGED": "KATA SANDI DIUBAH",
    "YOUR PASSWORD WAS CHANGED SUCCESSFULLY.": "KATA SANDI ANDA BERHASIL DIUBAH.",
    "CURRENT PASSWORD IS INCORRECT.": "KATA SANDI SAAT INI SALAH.",
    "USER NOT CREATED": "PENGGUNA TIDAK DIBUAT",
    "SET A NEW PASSWORD FOR": "ATUR KATA SANDI BARU UNTUK",
    "PASSWORD RESET": "KATA SANDI DIATUR ULANG",
    "THE PASSWORD WAS RESET.": "KATA SANDI BERHASIL DIATUR ULANG.",
    "SELECT A USER FIRST.": "PILIH PENGGUNA TERLEBIH DAHULU.",
    "ACTION BLOCKED": "TINDAKAN DIBLOKIR",
    "USER NOT FOUND.": "PENGGUNA TIDAK DITEMUKAN.",
    "YOU CANNOT DISABLE YOUR CURRENT ACCOUNT.": "ANDA TIDAK DAPAT MENONAKTIFKAN AKUN YANG SEDANG DIGUNAKAN.",
    "AT LEAST ONE ACTIVE ADMINISTRATOR IS REQUIRED.": "MINIMAL SATU ADMINISTRATOR AKTIF HARUS TERSEDIA.",
    "USER STATUS UPDATED.": "STATUS PENGGUNA BERHASIL DIPERBARUI.",
    "A NEW SUPPLIER MUST BE CREATED BY AN ADMINISTRATOR.": "PEMASOK BARU HARUS DIBUAT OLEH ADMINISTRATOR.",
    "A NEW PART MUST BE CREATED BY AN ADMINISTRATOR.": "SUKU CADANG BARU HARUS DIBUAT OLEH ADMINISTRATOR.",
}
UI_TRANSLATIONS.setdefault("id_ID", {}).update(_INDONESIAN_SECURITY_DASHBOARD_TRANSLATIONS)


# Final Indonesian phrase-level completion.  Every entry below is a complete
# sentence or interface phrase, preventing the word-by-word fallback from
# producing mixed Indonesian/English text.
_INDONESIAN_FINAL_COMPLETION_TRANSLATIONS = {
    # Product identity, help, and about dialog
    "Computer Service Manager 4.4 Secure": "PENGELOLA SERVIS KOMPUTER 4.4 — AMAN",
    "COMPUTER SERVICE MANAGER": "PENGELOLA SERVIS KOMPUTER",
    "COMPUTER SERVICE MANAGER v4.4": "PENGELOLA SERVIS KOMPUTER v4.4",
    "CURRENT TAB HELP": "BANTUAN TAB AKTIF",
    "HELP CONTENTS": "ISI BANTUAN",
    "OPEN APPLICATION HELP CONTENTS": "BUKA ISI BANTUAN APLIKASI",
    "OPEN HELP FOR THE ACTIVE TAB": "BUKA BANTUAN UNTUK TAB AKTIF",
    "EDIT DUMMY DATA": "UBAH DATA CONTOH",
    "OPEN DUMMY CREATOR": "BUKA PEMBUAT DATA CONTOH",
    "OPEN DUMMY DATA EDITOR": "BUKA PENYUNTING DATA CONTOH",
    "MULTILINGUAL DUMMY CREATOR OPENED": "PEMBUAT DATA CONTOH MULTIBAHASA DIBUKA",
    "DUMMY DATA EDITOR OPENED": "PENYUNTING DATA CONTOH DIBUKA",
    "PLACE DUMMY_CREATOR.EXE OR DUMMY_CREATOR.PY BESIDE THE APPLICATION.": "LETAKKAN DUMMY_CREATOR.EXE ATAU DUMMY_CREATOR.PY DI FOLDER YANG SAMA DENGAN APLIKASI.",
    "BUILD HELP.HHP AND COPY HELP.CHM TO THE APPLICATION FOLDER.": "KOMPILASI HELP.HHP, LALU SALIN HELP.CHM KE FOLDER APLIKASI.",
    "DONATION:<br><a href=\"https://paypal.me/rahfie\">paypal.me/rahfie</a>": "DONASI:<br><a href=\"https://paypal.me/rahfie\">paypal.me/rahfie</a>",
    "COPYRIGHT © ECOMTECH 2026": "HAK CIPTA © ECOMTECH 2026",

    # Database, paths, import, export, and file dialogs
    "DATABASE PATH": "LOKASI DATABASE",
    "SELECT FILE TO IMPORT": "PILIH BERKAS UNTUK DIIMPOR",
    "SAVE EXPORT FILE": "SIMPAN BERKAS EKSPOR",
    "OPEN AN EXISTING DATABASE OR CREATE A NEW DATABASE FOR THIS APPLICATION.": "BUKA DATABASE YANG SUDAH ADA ATAU BUAT DATABASE BARU UNTUK APLIKASI INI.",
    "PLEASE SELECT AN XLSX OR CSV FILE": "PILIH BERKAS XLSX ATAU CSV.",
    "LEGACY .XLS FILES ARE NOT SUPPORTED. SAVE THE FILE AS .XLSX OR CSV FIRST.": "BERKAS .XLS LAMA TIDAK DIDUKUNG. SIMPAN TERLEBIH DAHULU SEBAGAI .XLSX ATAU CSV.",
    "UNKNOWN TABLE TO EXPORT.": "TABEL YANG AKAN DIEKSPOR TIDAK DIKENALI.",
    "THE DRIVE FOR EXPORT PATH DOES NOT EXIST.": "DRIVE UNTUK LOKASI EKSPOR TIDAK TERSEDIA.",
    "ID (ROWID)": "ID (ROWID)",
    "CREATED DATE": "TANGGAL DIBUAT",
    "QTY": "JUMLAH",
    "EXPORTED TO:": "BERHASIL DIEKSPOR KE:",
    "EXPORTED": "BERHASIL DIEKSPOR",
    "TO EXCEL": "KE EXCEL",
    "FAILED TO EXPORT": "GAGAL MENGEKSPOR",
    "GOOGLE MAP LINK EMPTY": "TAUTAN GOOGLE MAPS KOSONG",
    "GOOGLE MAP LINK COPIED": "TAUTAN GOOGLE MAPS BERHASIL DISALIN",
    "GOOGLE MAP LINK EMPTY (CLIPBOARD CLEARED)": "TAUTAN GOOGLE MAPS KOSONG (PAPAN KLIP DIKOSONGKAN)",
    "COPIED TO CLIPBOARD": "BERHASIL DISALIN KE PAPAN KLIP",
    "COPY PHONE": "SALIN NOMOR TELEPON",
    "COPY ADDRESS": "SALIN ALAMAT",
    "COPY GOOGLE MAP": "SALIN TAUTAN GOOGLE MAPS",
    "OPEN GOOGLE MAP": "BUKA GOOGLE MAPS",
    "COPY INVOICE": "SALIN FAKTUR",

    # Form labels, validation, and natural Indonesian phrasing
    "TYPE": "JENIS",
    "TYPE:": "JENIS:",
    "PURCHASE INVOICE": "FAKTUR PEMBELIAN",
    "PURCHASE INVOICE:": "FAKTUR PEMBELIAN:",
    "EXPIRY DATE FROM": "TANGGAL KEDALUWARSA MULAI",
    "EXPIRY DATE FROM:": "TANGGAL KEDALUWARSA MULAI:",
    "CALCULATED AUTOMATICALLY: SERVICE DATE + WARRANTY": "DIHITUNG OTOMATIS: TANGGAL SERVIS + MASA GARANSI",
    "FILTER PARTS": "SARING SUKU CADANG",
    "WARRANTY FILTER": "PENYARING GARANSI",
    "INVALID DATE RANGE": "RENTANG TANGGAL TIDAK VALID",
    "FROM DATE CANNOT BE AFTER TO DATE": "TANGGAL MULAI TIDAK BOLEH MELEWATI TANGGAL AKHIR.",
    "INVALID WARRANTY DATE RANGE: FROM IS AFTER TO": "RENTANG TANGGAL GARANSI TIDAK VALID: TANGGAL MULAI MELEWATI TANGGAL AKHIR",
    "A SUPPLIER WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.": "PEMASOK DENGAN NAMA INI SUDAH ADA. PILIH PEMASOK TERSEBUT, LALU KLIK SIMPAN UNTUK MEMPERBARUINYA.",
    "ANOTHER SUPPLIER ALREADY USES THIS NAME.": "NAMA INI SUDAH DIGUNAKAN OLEH PEMASOK LAIN.",
    "A PART WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.": "SUKU CADANG DENGAN NAMA INI SUDAH ADA. PILIH SUKU CADANG TERSEBUT, LALU KLIK SIMPAN UNTUK MEMPERBARUINYA.",
    "ANOTHER PART ALREADY USES THIS NAME.": "NAMA INI SUDAH DIGUNAKAN OLEH SUKU CADANG LAIN.",
    "DUPLICATE SUPPLIER": "PEMASOK DUPLIKAT",
    "DUPLICATE PART": "SUKU CADANG DUPLIKAT",
    "CONTACT:<br><a href=\"mailto:e-comtech@mail.com\">e-comtech@mail.com</a> / <a href=\"mailto:rahfie27@gmail.com\">rahfie27@gmail.com</a>": "KONTAK:<br><a href=\"mailto:e-comtech@mail.com\">e-comtech@mail.com</a> / <a href=\"mailto:rahfie27@gmail.com\">rahfie27@gmail.com</a>",
    "REMINDER MESSAGE TEMPLATE": "FORMAT PESAN PENGINGAT",
    "LOG WARRANTY REMINDER": "CATAT PENGINGAT GARANSI",
    "LOG REMINDER": "CATAT PENGINGAT",
    "LOG REMINDERS": "CATAT PENGINGAT",
    "📝 LOG REMINDER": "📝 CATAT PENGINGAT",
    "📝 LOG REMINDERS": "📝 CATAT PENGINGAT",
    "SELECT WARRANTIES TO LOG REMINDERS": "PILIH DATA GARANSI YANG PENGINGATNYA AKAN DICATAT",
    "REMINDERS LOGGED": "PENGINGAT BERHASIL DICATAT",
    "ITEM": "OBJEK",
    "GOOGLE MAP": "GOOGLE MAPS",
    "GOOGLE MAP:": "GOOGLE MAPS:",

    # Account and security messages not covered by the original account pass
    "INVALID USER ROLE.": "PERAN PENGGUNA TIDAK VALID.",
    "ADMINISTRATOR REQUIRED.": "TINDAKAN INI MEMERLUKAN ADMINISTRATOR.",
    "PASSWORD CHANGED.": "KATA SANDI BERHASIL DIUBAH.",
    "USER STATUS UPDATED.": "STATUS PENGGUNA BERHASIL DIPERBARUI.",
    "YES": "YA",
    "NO": "TIDAK",

    # Audit-log action codes, entity names, and fixed details
    "SYSTEM": "SISTEM",
    "CREATE": "BUAT",
    "UPDATE": "PERBARUI",
    "DELETE": "HAPUS",
    "APPLICATION_OPEN": "APLIKASI DIBUKA",
    "APPLICATION_EXIT": "APLIKASI DITUTUP",
    "ACCESS_DENIED": "AKSES DITOLAK",
    "USER_CREATED": "PENGGUNA DIBUAT",
    "LOGIN_SUCCESS": "LOGIN BERHASIL",
    "LOGIN_FAILED": "LOGIN GAGAL",
    "LOGIN_BLOCKED": "LOGIN DIBLOKIR",
    "PASSWORD_CHANGE_FAILED": "PERUBAHAN KATA SANDI GAGAL",
    "PASSWORD_CHANGED": "KATA SANDI DIUBAH",
    "PASSWORD_RESET": "KATA SANDI DIATUR ULANG",
    "USER_ENABLED": "PENGGUNA DIAKTIFKAN",
    "USER_DISABLED": "PENGGUNA DINONAKTIFKAN",
    "USER_UNLOCKED": "KUNCI PENGGUNA DIBUKA",
    "DATABASE_RESTORED": "DATABASE DIPULIHKAN",
    "CONTACT": "KONTAK",
    "SUPPLIER": "PEMASOK",
    "PART": "SUKU CADANG",
    "MAIN WINDOW OPENED": "JENDELA UTAMA DIBUKA",
    "ADMINISTRATOR ACTION BLOCKED": "TINDAKAN ADMINISTRATOR DIBLOKIR",
    "FIRST ADMINISTRATOR SESSION": "SESI ADMINISTRATOR PERTAMA",
    "UNKNOWN OR INVALID CREDENTIALS": "KREDENSIAL TIDAK DIKENALI ATAU TIDAK VALID",
    "ACCOUNT DISABLED": "AKUN DINONAKTIFKAN",
    "ACCOUNT LOCKED": "AKUN TERKUNCI",
    "INVALID CREDENTIALS": "KREDENSIAL TIDAK VALID",
    "USER REQUESTED LOGOUT": "PENGGUNA MEMINTA KELUAR DARI AKUN",
    "WINDOW CLOSED": "JENDELA DITUTUP",
    "DASHBOARD REFRESH FAILED": "PEMUATAN ULANG DASBOR GAGAL",
    "NO SMS OR WHATSAPP MESSAGE WAS SENT.": "TIDAK ADA PESAN SMS ATAU WHATSAPP YANG DIKIRIM.",
}
UI_TRANSLATIONS.setdefault("id_ID", {}).update(_INDONESIAN_FINAL_COMPLETION_TRANSLATIONS)


# Complete static language catalogs for every supported non-English locale.
# Existing hand-refined entries take precedence; this layer fills only missing text.
_UI_TRANSLATION_COMPLETION_KEYS = ('A NEW PART MUST BE CREATED BY AN ADMINISTRATOR.',
 'A NEW SUPPLIER MUST BE CREATED BY AN ADMINISTRATOR.',
 'A PART WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.',
 'A SUPPLIER WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.',
 'ACCESS_DENIED',
 'ACCOUNT',
 'ACCOUNT DISABLED',
 'ACCOUNT LOCKED',
 'ACTION BLOCKED',
 'ACTIVE WARRANTIES',
 'ADD APPLICATION USER',
 'ADD USER',
 'ADDITIONAL NOTES',
 'ADDRESS LENGKAP',
 'ADMIN',
 'ADMINISTRATOR',
 'ADMINISTRATOR ACTION BLOCKED',
 'ADMINISTRATOR ONLY',
 'ADMINISTRATOR REQUIRED',
 'ADMINISTRATOR REQUIRED.',
 'ANOTHER PART ALREADY USES THIS NAME.',
 'ANOTHER SUPPLIER ALREADY USES THIS NAME.',
 'APPLICATION CLOSED',
 'APPLICATION_EXIT',
 'APPLICATION_OPEN',
 'AT LEAST ONE ACTIVE ADMINISTRATOR IS REQUIRED.',
 'AVG PER SERVICE',
 'AVG REPAIR COST',
 'AVG REVENUE',
 'AVG REVENUE/SERVICE',
 'BACKUP DATABASE',
 'BACKUP FAILED',
 'BACKUP FOLDER CANNOT BE EMPTY',
 'BACKUP FOLDER:',
 'BACKUP SUCCESSFUL',
 'BOGOR ON-CALL COMPUTER SERVICE REMINDS YOU THAT YOUR DEVICE REPAIR WARRANTY WILL EXPIRE ON:',
 'BRAND:',
 'BRANDS',
 'BUILD HELP.HHP AND COPY HELP.CHM TO THE APPLICATION FOLDER.',
 'CALCULATED AUTOMATICALLY: SERVICE DATE + WARRANTY',
 'CANNOT DETERMINE SELECTED SUPPLIER',
 'CHANGE PASSWORD',
 'CHANGE PASSWORD FOR',
 'COMPLETED',
 'COMPUTER SERVICE MANAGER',
 'CONFIRM DELETE',
 'CONFIRM PASSWORD',
 'CONFIRM RESTORE',
 'CONTACT NAME',
 'CONTACT:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
 'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
 'COPIED TO CLIPBOARD',
 'COPY ADDRESS',
 'COPY GOOGLE MAP',
 'COPY INVOICE',
 'COPY PHONE',
 'COPYRIGHT © ECOMTECH 2026',
 'COUNT',
 'CREATE ADMINISTRATOR',
 'CREATE THE FIRST ADMINISTRATOR ACCOUNT',
 'CREATED',
 'CREATED BY:',
 'CREATED DATE',
 'CSV FILES',
 'CURRENT PASSWORD',
 'CURRENT PASSWORD IS INCORRECT.',
 'CURRENT TAB HELP',
 'CUSTOMER NAME',
 'CUSTOMERS SELECTED FOR REMINDER',
 'CUSTOMERS SELECTED FOR REMINDER:',
 'Computer Service Manager 4.4 Secure',
 'DASHBOARD REFRESH FAILED',
 'DASHBOARD REVIEW',
 'DATA HAS BEEN EXPORTED TO:',
 'DATABASE LOCATION REQUIRED',
 'DATABASE PATH CANNOT BE EMPTY',
 'DATABASE PATH:',
 'DATABASE PATHS UPDATED SUCCESSFULLY!',
 'DATABASE RESTORED FROM BACKUP',
 'DATABASE RESTORED SUCCESSFULLY FROM BACKUP!',
 'DATABASE_RESTORED',
 'DATE',
 'DEAR [CUSTOMER],',
 'DETAILS',
 'DEVICE BRAND STATISTICS',
 'DEVICE MODEL',
 'DEVICE TYPES',
 'DISPLAY NAME',
 'DONATION:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
 'DRIVE D: IS NOT AVAILABLE. PLEASE CHOOSE ANOTHER FOLDER.',
 'DRIVE D: WAS NOT FOUND.\nPLEASE CREATE OR SELECT A FOLDER FOR THE DATABASE.',
 'DRIVE NOT FOUND',
 'DUMMY DATA EDITOR OPENED',
 'DUPLICATE PART',
 'DUPLICATE SUPPLIER',
 'EDIT MODE',
 'EDIT MODE: UPDATE FIELDS THEN CLICK SAVE',
 'EDIT MODE: UPDATE FIELDS THEN CLICK UPDATE',
 'ENABLE / DISABLE',
 'END DATE:',
 'ENTER BOTH USERNAME AND PASSWORD.',
 'EXCEL FILES',
 'EXPIRING IN 30 DAYS',
 'EXPIRING SOON',
 'EXPIRING SOON (< 7 DAYS)',
 'EXPIRING SOON (≤ 30 DAYS)',
 'EXPIRING SOON (≤ 7 DAYS)',
 'EXPIRING SOON (≤30 days)',
 'EXPIRY DATE FROM',
 'EXPIRY DATE FROM:',
 'EXPORT COMPLETE',
 'EXPORT ERROR',
 'EXPORT FAILED',
 'EXPORT SUCCESSFUL',
 'EXPORTED TO:',
 'FAILED TO EXPORT',
 'FINANCIAL REPORT',
 'FIRST ADMINISTRATOR SESSION',
 'FIRST-TIME DATABASE SETUP',
 'FIRST-TIME SETUP:\nPLEASE CHOOSE A FOLDER TO SAVE THE NEW DATABASE.',
 'FORM CLEARED',
 'FROM DATE CANNOT BE AFTER TO DATE',
 'FULL ADDRESS',
 'GOOGLE MAP',
 'GOOGLE MAP LINK COPIED',
 'GOOGLE MAP LINK EMPTY',
 'GOOGLE MAP LINK EMPTY (CLIPBOARD CLEARED)',
 'GOOGLE MAP:',
 'HELP CONTENTS',
 'ID (ROWID)',
 'IMPORT COMPLETE',
 'IMPORT FAILED',
 'IMPORT/EXPORT FOLDER CANNOT BE EMPTY',
 'IMPORT/EXPORT FOLDER:',
 'IN PROGRESS',
 'IN REPAIR',
 'INVALID CREDENTIALS',
 'INVALID DATE RANGE',
 'INVALID DRIVE',
 'INVALID PATH',
 'INVALID USER ROLE.',
 'INVALID USERNAME',
 'INVALID USERNAME OR PASSWORD.',
 'INVALID WARRANTY DATE RANGE: FROM IS AFTER TO',
 'INVOICE NUMBER',
 'ISSUES',
 'LAST LOGIN',
 'LEGACY .XLS FILES ARE NOT SUPPORTED. SAVE THE FILE AS .XLSX OR CSV FIRST.',
 'LOCKED UNTIL',
 'LOG REMINDER',
 'LOG REMINDERS',
 'LOG WARRANTY REMINDER',
 'LOGIN',
 'LOGIN SUCCESSFUL.',
 'LOGIN_BLOCKED',
 'LOGIN_FAILED',
 'LOGIN_SUCCESS',
 'LOGOUT',
 'LOW-STOCK PARTS',
 'MAIN WINDOW OPENED',
 'MINIMUM 4 CHARACTERS. ANY LETTER, NUMBER, SPACE, OR SYMBOL FORMAT IS ALLOWED.',
 'MODEL DEVICE',
 'MODELS',
 'MULTILINGUAL DUMMY CREATOR OPENED',
 'MY RECENT ACTIVITY',
 'N/A',
 'NAME CUSTOMER',
 'NAME IS REQUIRED',
 'NAME:',
 'NEW PASSWORD',
 'NO DATA',
 'NO DATABASE SELECTED',
 'NO DEFAULT PASSWORD IS USED. PASSWORDS ARE STORED ONLY AS UNIQUE, SALTED SCRYPT HASHES. KEEP THIS '
 'ADMINISTRATOR PASSWORD SAFE.',
 'NO EXISTING DATABASE WAS SELECTED. THE APPLICATION WILL CONTINUE WITH NEW DATABASE CREATION.',
 'NO SMS OR WHATSAPP MESSAGE WAS SENT.',
 'NO STATUS DATA',
 'NO WARRANTY',
 'NO WARRANTY OF ANY KIND IS PROVIDED',
 'NOMOR PHONE',
 'NOT FOUND',
 'OPEN AN EXISTING DATABASE OR CREATE A NEW DATABASE FOR THIS APPLICATION.',
 'OPEN APPLICATION HELP CONTENTS',
 'OPEN DUMMY DATA EDITOR',
 'OPEN GOOGLE MAP',
 'OPEN HELP FOR THE ACTIVE TAB',
 'PART NAME',
 'PART NAME IS REQUIRED',
 'PART TYPES',
 'PARTS REFRESHED',
 'PASSWORD',
 'PASSWORD CHANGED',
 'PASSWORD CHANGED.',
 'PASSWORD IS TOO LONG.',
 'PASSWORD MISMATCH',
 'PASSWORD MUST CONTAIN AT LEAST 4 CHARACTERS.',
 'PASSWORD NOT CHANGED',
 'PASSWORD RESET',
 'PASSWORD_CHANGED',
 'PASSWORD_CHANGE_FAILED',
 'PASSWORD_RESET',
 'PHONE NUMBER',
 'PHONE:',
 'PLACE DUMMY_CREATOR.EXE OR DUMMY_CREATOR.PY BESIDE THE APPLICATION.',
 'PLEASE CONTACT US IF YOU NEED FURTHER SERVICE.',
 'PLEASE SELECT AN XLSX OR CSV FILE',
 'PLEASE SELECT EXCEL OR CSV FILE',
 'PURCHASE INVOICE:',
 'QTY',
 'RECENT SECURITY AND DATA ACTIVITY',
 'RECENT SERVICE RECORDS',
 'REFRESH DASHBOARD',
 'RELOAD DROPDOWN DATA',
 'REMINDER MESSAGE TEMPLATE',
 'REMINDERS LOGGED',
 'REMINDERS SENT',
 'REPAIR COST',
 'REPORT EXPORT FAILED',
 'REPORT EXPORTED',
 'RESET PASSWORD',
 'RESET USER PASSWORD',
 'RESET WARRANTY RANGE FROM TODAY',
 'RESTORE DATABASE',
 'RESTORE FAILED',
 'RESTORE SUCCESSFUL',
 'ROLE',
 'SAVE EXPORT FILE',
 'SAVE REPORT EXPORT',
 'SEARCH BY NAME, PHONE, DEVICE, OR ISSUE...',
 'SECURE ADMINISTRATOR SETUP',
 'SECURE LOGIN',
 'SELECT A CONTACT TO DELETE',
 'SELECT A CONTACT TO EDIT',
 'SELECT A PART TO DELETE',
 'SELECT A PART TO EDIT',
 'SELECT A PART TO UPDATE',
 'SELECT A SUPPLIER TO DELETE',
 'SELECT A SUPPLIER TO EDIT',
 'SELECT A SUPPLIER TO UPDATE',
 'SELECT A USER FIRST.',
 'SELECT A WARRANTY ROW TO EDIT',
 'SELECT APPLICATION CURRENCY:',
 'SELECT APPLICATION LANGUAGE AND LOCALE:',
 'SELECT DEFAULT CURRENCY FOR THE APPLICATION:',
 'SELECT LANGUAGE / LOCALE',
 'SELECT ROWS TO EXPORT',
 'SELECT WARRANTIES TO LOG REMINDERS',
 'SELECT WARRANTIES TO SEND REMINDERS',
 'SEND REMINDERS',
 'SERVICE DESCRIPTION',
 'SERVICE KOMPUTER PANGGILAN BOGOR',
 'SERVICE STATUS GRAPHIC',
 'SERVICE SUMMARY',
 'SERVICE SUMMARY REPORT',
 'SET A NEW PASSWORD FOR',
 'SET DATABASE LOCATION',
 'SET SERVICE DATE TO TODAY',
 'SETTINGS SAVED',
 'SETUP FAILED',
 'SHOW PASSWORD',
 'SHOW PASSWORDS',
 'SIGN IN TO CONTINUE',
 'SIGNED IN',
 'SIGNED IN AS',
 'STAFF',
 'STAFF ACCESS IS LIMITED',
 'STAFF CAN ADD NEW SERVICE RECORDS AND VIEW OPERATIONAL DATA. EDITING, DELETING, IMPORT, EXPORT, BACKUP, '
 'ADMINISTRATIVE SETTINGS, USER MANAGEMENT, SUPPLIER CHANGES, PART CHANGES, REPORTS, AND REMINDERS REQUIRE '
 'AN ADMINISTRATOR.',
 'STATUSES',
 'STORAGE LOCATION',
 'SUPPLIER ADDRESS',
 'SUPPLIER NAME',
 'SUPPLIER NAME IS REQUIRED',
 'SUPPLIERS REFRESHED',
 'SYSTEM',
 'TECHNICIAN NAME',
 'TECHNICIAN PERFORMANCE',
 'TECHNICIAN PERFORMANCE REPORT',
 'TECHNICIANS',
 'THANK YOU,',
 'THAT USERNAME ALREADY EXISTS.',
 'THE DRIVE FOR BACKUP FOLDER DOES NOT EXIST.',
 'THE DRIVE FOR DATABASE PATH DOES NOT EXIST.',
 'THE DRIVE FOR EXPORT PATH DOES NOT EXIST.',
 'THE DRIVE FOR IMPORT/EXPORT FOLDER DOES NOT EXIST.',
 'THE PASSWORD WAS RESET.',
 'THE PASSWORDS DO NOT MATCH.',
 'THIS ACCOUNT IS DISABLED. CONTACT AN ADMINISTRATOR.',
 'THIS ACTION IS RESTRICTED TO ADMINISTRATOR ACCOUNTS.',
 'THIS SOFTWARE IS PROVIDED AS-IS',
 'TIME',
 'TO EXCEL',
 'TOTAL ACTIVE WARRANTIES',
 'TOTAL CUSTOMERS',
 'TOTAL DEVICES',
 'TOTAL EXPIRED',
 'TOTAL EXPIRING SOON',
 'TOTAL REVENUE',
 'TOTAL SELECTED',
 'TOTAL SERVICE RECORDS',
 'TOTAL SERVICE VALUE',
 'TOTAL SERVICES',
 'TYPE:',
 'UNIQUE MODELS',
 'UNKNOWN OR INVALID CREDENTIALS',
 'UNKNOWN TABLE TO EXPORT.',
 'UNLOCK',
 'UNSUPPORTED FORMAT',
 'USE D:\\BACKUP (DEFAULT)',
 'USER',
 'USER MANAGEMENT',
 'USER NOT CREATED',
 'USER NOT FOUND.',
 'USER REQUESTED LOGOUT',
 'USER STATUS UPDATED.',
 'USERNAME',
 'USERNAME EXISTS',
 'USERNAME MUST BE 3-32 CHARACTERS USING LETTERS, NUMBERS, DOT, DASH, OR UNDERSCORE.',
 'USER_CREATED',
 'USER_DISABLED',
 'USER_ENABLED',
 'USER_UNLOCKED',
 'WAITING FOR PARTS',
 'WAITING PARTS',
 'WARNING: RESTORING WILL REPLACE ALL CURRENT DATA WITH BACKUP DATA.\n'
 'THIS ACTION CANNOT BE UNDONE.\n'
 '\n'
 'DO YOU WANT TO CONTINUE?',
 'WARRANTY PERIODS',
 'WARRANTY STATUS:',
 'WARRANTY SUMMARY',
 'WARRANTY SUMMARY REPORT',
 'WEAK PASSWORD',
 'WHAT DO YOU WANT TO DO WITH THE DATABASE?',
 'WINDOW CLOSED',
 'YOU CANNOT DISABLE YOUR CURRENT ACCOUNT.',
 'YOUR PASSWORD WAS CHANGED SUCCESSFULLY.',
 '⚠️ WARNING!',
 '✏️ EDIT',
 '✏️ EDIT CONTACT',
 '❌ DELETE',
 '➕ ADD',
 '➕ ADD PART',
 '📄 TEXT VIEW',
 '📋 TABLE VIEW',
 '📝 LOG REMINDER',
 '📝 LOG REMINDERS',
 '📤 EXPORT SELECTED',
 '🔃 REFRESH',
 '🔄 REFRESH',
 '🗑️ CLEAR')
_UI_TRANSLATION_COMPLETIONS = {'es_MX': ['Un administrador debe crear una pieza nueva.',
           'Un administrador debe crear un proveedor nuevo.',
           'Ya existe una pieza con este nombre. Selecciónela y use Guardar para actualizarla.',
           'Ya existe un proveedor con este nombre. Selecciónelo y use Guardar para actualizarlo.',
           'ACCESO_DENEGADO',
           'CUENTA',
           'CUENTA DESHABILITADA',
           'CUENTA BLOQUEADA',
           'ACCIÓN BLOQUEADA',
           'GARANTÍAS ACTIVAS',
           'AGREGAR USUARIO DE LA APLICACIÓN',
           'AGREGAR USUARIO',
           'NOTAS ADICIONALES',
           'DIRECCIÓN COMPLETA',
           'ADMIN',
           'ADMINISTRADOR',
           'ACCIÓN DE ADMINISTRADOR BLOQUEADA',
           'SOLO ADMINISTRADORES',
           'SE REQUIERE ADMINISTRADOR',
           'SE REQUIERE ADMINISTRADOR.',
           'Otra pieza ya utiliza este nombre.',
           'Otro proveedor ya utiliza este nombre.',
           'APLICACIÓN CERRADA',
           'SALIDA_DE_LA_APLICACIÓN',
           'APERTURA_DE_LA_APLICACIÓN',
           'Se requiere al menos un administrador activo.',
           'PROMEDIO POR SERVICIO',
           'COSTO PROMEDIO DE REPARACIÓN',
           'INGRESO PROMEDIO',
           'INGRESO PROMEDIO/SERVICIO',
           'RESPALDAR BASE DE DATOS',
           'ERROR DE RESPALDO',
           'LA CARPETA DE RESPALDO NO PUEDE ESTAR VACÍA',
           'CARPETA DE RESPALDO:',
           'RESPALDO CORRECTO',
           'El servicio informático a domicilio de Bogor le recuerda que la garantía de reparación de su dispositivo '
           'vencerá el:',
           'MARCA:',
           'MARCAS',
           'Compile HELP.HHP y copie HELP.CHM en la carpeta de la aplicación.',
           'CALCULADO AUTOMÁTICAMENTE: FECHA DE SERVICIO + GARANTÍA',
           'NO SE PUEDE DETERMINAR EL PROVEEDOR SELECCIONADO',
           'CAMBIAR CONTRASEÑA',
           'CAMBIAR CONTRASEÑA DE',
           'COMPLETADO',
           'ADMINISTRADOR DE SERVICIO INFORMÁTICO',
           'CONFIRMAR ELIMINACIÓN',
           'CONFIRMAR CONTRASEÑA',
           'CONFIRMAR RESTAURACIÓN',
           'NOMBRE DEL CONTACTO',
           'CONTACTO:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'COPIADO AL PORTAPAPELES',
           'COPIAR DIRECCIÓN',
           'COPIAR MAPA DE GOOGLE',
           'COPIAR FACTURA',
           'COPIAR TELÉFONO',
           'DERECHOS DE AUTOR © ECOMTECH 2026',
           'CANTIDAD',
           'CREAR ADMINISTRADOR',
           'CREAR LA PRIMERA CUENTA DE ADMINISTRADOR',
           'CREADO',
           'CREADO POR:',
           'FECHA DE CREACIÓN',
           'ARCHIVOS CSV',
           'CONTRASEÑA ACTUAL',
           'LA CONTRASEÑA ACTUAL ES INCORRECTA.',
           'AYUDA DE LA PESTAÑA ACTUAL',
           'NOMBRE DEL CLIENTE',
           'CLIENTES SELECCIONADOS PARA RECORDATORIO',
           'CLIENTES SELECCIONADOS PARA RECORDATORIO:',
           'Administrador de Servicio Informático 4.4 Seguro',
           'ERROR AL ACTUALIZAR EL PANEL',
           'REVISIÓN DEL PANEL',
           'LOS DATOS SE HAN EXPORTADO A:',
           'SE REQUIERE LA UBICACIÓN DE LA BASE DE DATOS',
           'LA RUTA DE LA BASE DE DATOS NO PUEDE ESTAR VACÍA',
           'RUTA DE LA BASE DE DATOS:',
           '¡LAS RUTAS DE LA BASE DE DATOS SE ACTUALIZARON CORRECTAMENTE!',
           'BASE DE DATOS RESTAURADA DESDE EL RESPALDO',
           '¡BASE DE DATOS RESTAURADA CORRECTAMENTE DESDE EL RESPALDO!',
           'BASE_DE_DATOS_RESTAURADA',
           'FECHA',
           'ESTIMADO/A [CUSTOMER],',
           'DETALLES',
           'ESTADÍSTICAS DE MARCAS DE DISPOSITIVOS',
           'MODELO DEL DISPOSITIVO',
           'TIPOS DE DISPOSITIVO',
           'NOMBRE PARA MOSTRAR',
           'DONACIÓN:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'La unidad D: no está disponible. Elija otra carpeta.',
           'No se encontró la unidad D:.\nCree o seleccione una carpeta para la base de datos.',
           'UNIDAD NO ENCONTRADA',
           'EDITOR DE DATOS DE PRUEBA ABIERTO',
           'PIEZA DUPLICADA',
           'PROVEEDOR DUPLICADO',
           'MODO DE EDICIÓN',
           'MODO DE EDICIÓN: ACTUALICE LOS CAMPOS Y HAGA CLIC EN GUARDAR',
           'MODO DE EDICIÓN: ACTUALICE LOS CAMPOS Y HAGA CLIC EN ACTUALIZAR',
           'HABILITAR / DESHABILITAR',
           'FECHA DE FIN:',
           'INTRODUZCA EL NOMBRE DE USUARIO Y LA CONTRASEÑA.',
           'ARCHIVOS DE EXCEL',
           'VENCE EN 30 DÍAS',
           'PRÓXIMO A VENCER',
           'PRÓXIMO A VENCER (< 7 DÍAS)',
           'PRÓXIMO A VENCER (≤ 30 DÍAS)',
           'PRÓXIMO A VENCER (≤ 7 DÍAS)',
           'PRÓXIMO A VENCER (≤30 días)',
           'FECHA DE VENCIMIENTO DESDE',
           'FECHA DE VENCIMIENTO DESDE:',
           'EXPORTACIÓN COMPLETA',
           'ERROR DE EXPORTACIÓN',
           'ERROR AL EXPORTAR',
           'EXPORTACIÓN CORRECTA',
           'EXPORTADO A:',
           'NO SE PUDO EXPORTAR',
           'INFORME FINANCIERO',
           'PRIMERA SESIÓN DEL ADMINISTRADOR',
           'CONFIGURACIÓN INICIAL DE LA BASE DE DATOS',
           'CONFIGURACIÓN INICIAL:\nElija una carpeta para guardar la nueva base de datos.',
           'FORMULARIO LIMPIADO',
           'LA FECHA DESDE NO PUEDE SER POSTERIOR A LA FECHA HASTA',
           'DIRECCIÓN COMPLETA',
           'MAPA DE GOOGLE',
           'ENLACE DE GOOGLE MAPS COPIADO',
           'ENLACE DE GOOGLE MAPS VACÍO',
           'ENLACE DE GOOGLE MAPS VACÍO (PORTAPAPELES LIMPIADO)',
           'MAPA DE GOOGLE:',
           'CONTENIDO DE LA AYUDA',
           'ID (ROWID)',
           'IMPORTACIÓN COMPLETA',
           'ERROR AL IMPORTAR',
           'LA CARPETA DE IMPORTACIÓN/EXPORTACIÓN NO PUEDE ESTAR VACÍA',
           'CARPETA DE IMPORTACIÓN/EXPORTACIÓN:',
           'EN PROGRESO',
           'EN REPARACIÓN',
           'CREDENCIALES NO VÁLIDAS',
           'INTERVALO DE FECHAS NO VÁLIDO',
           'UNIDAD NO VÁLIDA',
           'RUTA NO VÁLIDA',
           'ROL DE USUARIO NO VÁLIDO.',
           'NOMBRE DE USUARIO NO VÁLIDO',
           'NOMBRE DE USUARIO O CONTRASEÑA NO VÁLIDOS.',
           'INTERVALO DE GARANTÍA NO VÁLIDO: LA FECHA DESDE ES POSTERIOR A LA FECHA HASTA',
           'NÚMERO DE FACTURA',
           'PROBLEMAS',
           'ÚLTIMO INICIO DE SESIÓN',
           'LOS ARCHIVOS .XLS ANTIGUOS NO SON COMPATIBLES. GUARDE PRIMERO EL ARCHIVO COMO .XLSX O CSV.',
           'BLOQUEADO HASTA',
           'REGISTRAR RECORDATORIO',
           'REGISTRAR RECORDATORIOS',
           'REGISTRAR RECORDATORIO DE GARANTÍA',
           'INICIAR SESIÓN',
           'INICIO DE SESIÓN CORRECTO.',
           'INICIO_DE_SESIÓN_BLOQUEADO',
           'ERROR_DE_INICIO_DE_SESIÓN',
           'INICIO_DE_SESIÓN_CORRECTO',
           'CERRAR SESIÓN',
           'PIEZAS CON POCO STOCK',
           'VENTANA PRINCIPAL ABIERTA',
           'MÍNIMO 4 CARACTERES. SE PERMITE CUALQUIER FORMATO CON LETRAS, NÚMEROS, ESPACIOS O SÍMBOLOS.',
           'MODELO DEL DISPOSITIVO',
           'MODELOS',
           'CREADOR MULTILINGÜE DE DATOS DE PRUEBA ABIERTO',
           'MI ACTIVIDAD RECIENTE',
           'N/D',
           'NOMBRE DEL CLIENTE',
           'EL NOMBRE ES OBLIGATORIO',
           'NOMBRE:',
           'NUEVA CONTRASEÑA',
           'SIN DATOS',
           'NO SE HA SELECCIONADO UNA BASE DE DATOS',
           'NO SE USA NINGUNA CONTRASEÑA PREDETERMINADA. LAS CONTRASEÑAS SOLO SE GUARDAN COMO HASHES SCRYPT ÚNICOS, '
           'CON SAL. MANTENGA SEGURA ESTA CONTRASEÑA DE ADMINISTRADOR.',
           'NO SE SELECCIONÓ NINGUNA BASE DE DATOS EXISTENTE. LA APLICACIÓN CONTINUARÁ CREANDO UNA BASE DE DATOS '
           'NUEVA.',
           'NO SE ENVIÓ NINGÚN MENSAJE SMS NI DE WHATSAPP.',
           'SIN DATOS DE ESTADO',
           'SIN GARANTÍA',
           'NO SE OFRECE NINGÚN TIPO DE GARANTÍA',
           'NÚMERO DE TELÉFONO',
           'NO ENCONTRADO',
           'ABRA UNA BASE DE DATOS EXISTENTE O CREE UNA NUEVA PARA ESTA APLICACIÓN.',
           'ABRIR EL CONTENIDO DE AYUDA DE LA APLICACIÓN',
           'ABRIR EL EDITOR DE DATOS DE PRUEBA',
           'ABRIR MAPA DE GOOGLE',
           'ABRIR LA AYUDA DE LA PESTAÑA ACTIVA',
           'NOMBRE DE LA PIEZA',
           'EL NOMBRE DE LA PIEZA ES OBLIGATORIO',
           'TIPOS DE PIEZA',
           'PIEZAS ACTUALIZADAS',
           'CONTRASEÑA',
           'CONTRASEÑA CAMBIADA',
           'CONTRASEÑA CAMBIADA.',
           'LA CONTRASEÑA ES DEMASIADO LARGA.',
           'LAS CONTRASEÑAS NO COINCIDEN',
           'LA CONTRASEÑA DEBE CONTENER AL MENOS 4 CARACTERES.',
           'CONTRASEÑA SIN CAMBIOS',
           'CONTRASEÑA RESTABLECIDA',
           'CONTRASEÑA_CAMBIADA',
           'ERROR_AL_CAMBIAR_CONTRASEÑA',
           'CONTRASEÑA_RESTABLECIDA',
           'NÚMERO DE TELÉFONO',
           'TELÉFONO:',
           'COLOQUE DUMMY_CREATOR.EXE O DUMMY_CREATOR.PY JUNTO A LA APLICACIÓN.',
           'PÓNGASE EN CONTACTO CON NOSOTROS SI NECESITA MÁS SERVICIO.',
           'SELECCIONE UN ARCHIVO XLSX O CSV',
           'SELECCIONE UN ARCHIVO DE EXCEL O CSV',
           'FACTURA DE COMPRA:',
           'CANT.',
           'ACTIVIDAD RECIENTE DE SEGURIDAD Y DATOS',
           'REGISTROS DE SERVICIO RECIENTES',
           'ACTUALIZAR PANEL',
           'RECARGAR DATOS DE LAS LISTAS DESPLEGABLES',
           'PLANTILLA DE MENSAJE DE RECORDATORIO',
           'RECORDATORIOS REGISTRADOS',
           'RECORDATORIOS ENVIADOS',
           'COSTO DE REPARACIÓN',
           'ERROR AL EXPORTAR EL INFORME',
           'INFORME EXPORTADO',
           'RESTABLECER CONTRASEÑA',
           'RESTABLECER CONTRASEÑA DEL USUARIO',
           'RESTABLECER EL INTERVALO DE GARANTÍA DESDE HOY',
           'RESTAURAR BASE DE DATOS',
           'ERROR AL RESTAURAR',
           'RESTAURACIÓN CORRECTA',
           'ROL',
           'GUARDAR ARCHIVO DE EXPORTACIÓN',
           'GUARDAR EXPORTACIÓN DEL INFORME',
           'BUSCAR POR NOMBRE, TELÉFONO, DISPOSITIVO O PROBLEMA...',
           'CONFIGURACIÓN SEGURA DEL ADMINISTRADOR',
           'INICIO DE SESIÓN SEGURO',
           'SELECCIONE UN CONTACTO PARA ELIMINARLO',
           'SELECCIONE UN CONTACTO PARA EDITARLO',
           'SELECCIONE UNA PIEZA PARA ELIMINARLA',
           'SELECCIONE UNA PIEZA PARA EDITARLA',
           'SELECCIONE UNA PIEZA PARA ACTUALIZARLA',
           'SELECCIONE UN PROVEEDOR PARA ELIMINARLO',
           'SELECCIONE UN PROVEEDOR PARA EDITARLO',
           'SELECCIONE UN PROVEEDOR PARA ACTUALIZARLO',
           'SELECCIONE PRIMERO UN USUARIO.',
           'SELECCIONE UNA FILA DE GARANTÍA PARA EDITARLA',
           'SELECCIONE LA MONEDA DE LA APLICACIÓN:',
           'SELECCIONE EL IDIOMA Y LA REGIÓN DE LA APLICACIÓN:',
           'SELECCIONE LA MONEDA PREDETERMINADA DE LA APLICACIÓN:',
           'SELECCIONAR IDIOMA / REGIÓN',
           'SELECCIONE FILAS PARA EXPORTAR',
           'SELECCIONE GARANTÍAS PARA REGISTRAR RECORDATORIOS',
           'SELECCIONE GARANTÍAS PARA ENVIAR RECORDATORIOS',
           'ENVIAR RECORDATORIOS',
           'DESCRIPCIÓN DEL SERVICIO',
           'SERVICIO INFORMÁTICO A DOMICILIO EN BOGOR',
           'GRÁFICO DE ESTADO DEL SERVICIO',
           'RESUMEN DEL SERVICIO',
           'INFORME RESUMIDO DEL SERVICIO',
           'ESTABLECER UNA NUEVA CONTRASEÑA PARA',
           'ESTABLECER LA UBICACIÓN DE LA BASE DE DATOS',
           'ESTABLECER LA FECHA DE SERVICIO EN HOY',
           'AJUSTES GUARDADOS',
           'ERROR DE CONFIGURACIÓN',
           'MOSTRAR CONTRASEÑA',
           'MOSTRAR CONTRASEÑAS',
           'INICIE SESIÓN PARA CONTINUAR',
           'SESIÓN INICIADA',
           'SESIÓN INICIADA COMO',
           'PERSONAL',
           'EL ACCESO DEL PERSONAL ES LIMITADO',
           'EL PERSONAL PUEDE AGREGAR NUEVOS REGISTROS DE SERVICIO Y VER DATOS OPERATIVOS. LA EDICIÓN, ELIMINACIÓN, '
           'IMPORTACIÓN, EXPORTACIÓN, RESPALDO, AJUSTES ADMINISTRATIVOS, GESTIÓN DE USUARIOS, CAMBIOS DE PROVEEDORES, '
           'CAMBIOS DE PIEZAS, INFORMES Y RECORDATORIOS REQUIEREN UN ADMINISTRADOR.',
           'ESTADOS',
           'UBICACIÓN DE ALMACENAMIENTO',
           'DIRECCIÓN DEL PROVEEDOR',
           'NOMBRE DEL PROVEEDOR',
           'EL NOMBRE DEL PROVEEDOR ES OBLIGATORIO',
           'PROVEEDORES ACTUALIZADOS',
           'SISTEMA',
           'NOMBRE DEL TÉCNICO',
           'RENDIMIENTO DEL TÉCNICO',
           'INFORME DE RENDIMIENTO DEL TÉCNICO',
           'TÉCNICOS',
           'GRACIAS,',
           'ESE NOMBRE DE USUARIO YA EXISTE.',
           'LA UNIDAD DE LA CARPETA DE RESPALDO NO EXISTE.',
           'LA UNIDAD DE LA RUTA DE LA BASE DE DATOS NO EXISTE.',
           'LA UNIDAD DE LA RUTA DE EXPORTACIÓN NO EXISTE.',
           'LA UNIDAD DE LA CARPETA DE IMPORTACIÓN/EXPORTACIÓN NO EXISTE.',
           'LA CONTRASEÑA FUE RESTABLECIDA.',
           'LAS CONTRASEÑAS NO COINCIDEN.',
           'ESTA CUENTA ESTÁ DESHABILITADA. CONTACTE A UN ADMINISTRADOR.',
           'ESTA ACCIÓN ESTÁ RESTRINGIDA A CUENTAS DE ADMINISTRADOR.',
           'ESTE SOFTWARE SE PROPORCIONA TAL CUAL',
           'HORA',
           'A EXCEL',
           'TOTAL DE GARANTÍAS ACTIVAS',
           'TOTAL DE CLIENTES',
           'TOTAL DE DISPOSITIVOS',
           'TOTAL VENCIDO',
           'TOTAL PRÓXIMO A VENCER',
           'INGRESOS TOTALES',
           'TOTAL SELECCIONADO',
           'TOTAL DE REGISTROS DE SERVICIO',
           'VALOR TOTAL DEL SERVICIO',
           'TOTAL DE SERVICIOS',
           'TIPO:',
           'MODELOS ÚNICOS',
           'CREDENCIALES DESCONOCIDAS O NO VÁLIDAS',
           'TABLA DESCONOCIDA PARA EXPORTAR.',
           'DESBLOQUEAR',
           'FORMATO NO COMPATIBLE',
           'USAR D:\\BACKUP (PREDETERMINADO)',
           'USUARIO',
           'GESTIÓN DE USUARIOS',
           'USUARIO NO CREADO',
           'USUARIO NO ENCONTRADO.',
           'EL USUARIO SOLICITÓ CERRAR SESIÓN',
           'ESTADO DEL USUARIO ACTUALIZADO.',
           'NOMBRE DE USUARIO',
           'EL NOMBRE DE USUARIO YA EXISTE',
           'EL NOMBRE DE USUARIO DEBE TENER ENTRE 3 Y 32 CARACTERES Y USAR LETRAS, NÚMEROS, PUNTO, GUION O GUION BAJO.',
           'USUARIO_CREADO',
           'USUARIO_DESHABILITADO',
           'USUARIO_HABILITADO',
           'USUARIO_DESBLOQUEADO',
           'ESPERANDO PIEZAS',
           'PIEZAS EN ESPERA',
           'ADVERTENCIA: LA RESTAURACIÓN REEMPLAZARÁ TODOS LOS DATOS ACTUALES CON LOS DATOS DEL RESPALDO.\n'
           'ESTA ACCIÓN NO SE PUEDE DESHACER.\n'
           '\n'
           '¿DESEA CONTINUAR?',
           'PERIODOS DE GARANTÍA',
           'ESTADO DE LA GARANTÍA:',
           'RESUMEN DE GARANTÍA',
           'INFORME RESUMIDO DE GARANTÍA',
           'CONTRASEÑA DÉBIL',
           '¿QUÉ DESEA HACER CON LA BASE DE DATOS?',
           'VENTANA CERRADA',
           'NO PUEDE DESHABILITAR SU CUENTA ACTUAL.',
           'SU CONTRASEÑA SE CAMBIÓ CORRECTAMENTE.',
           '⚠️ ¡ADVERTENCIA!',
           '✏️ EDITAR',
           '✏️ EDITAR CONTACTO',
           '❌ ELIMINAR',
           '➕ AGREGAR',
           '➕ AGREGAR PIEZA',
           '📄 VISTA DE TEXTO',
           '📋 VISTA DE TABLA',
           '📝 REGISTRAR RECORDATORIO',
           '📝 REGISTRAR RECORDATORIOS',
           '📤 EXPORTAR SELECCIÓN',
           '🔃 ACTUALIZAR',
           '🔄 ACTUALIZAR',
           '🗑️ LIMPIAR'],
 'fr_FR': ['Une nouvelle pièce doit être créée par un administrateur.',
           'Un nouveau fournisseur doit être créé par un administrateur.',
           'Une pièce portant ce nom existe déjà. Sélectionnez-la et utilisez Enregistrer pour la mettre à jour.',
           'Un fournisseur portant ce nom existe déjà. Sélectionnez-le et utilisez Enregistrer pour le mettre à jour.',
           'ACCÈS_REFUSÉ',
           'COMPTE',
           'COMPTE DÉSACTIVÉ',
           'COMPTE VERROUILLÉ',
           'ACTION BLOQUÉE',
           'GARANTIES ACTIVES',
           'AJOUTER UN UTILISATEUR DE L’APPLICATION',
           'AJOUTER UN UTILISATEUR',
           'NOTES SUPPLÉMENTAIRES',
           'ADRESSE COMPLÈTE',
           'ADMIN',
           'ADMINISTRATEUR',
           'ACTION D’ADMINISTRATEUR BLOQUÉE',
           'ADMINISTRATEUR UNIQUEMENT',
           'ADMINISTRATEUR REQUIS',
           'ADMINISTRATEUR REQUIS.',
           'Une autre pièce utilise déjà ce nom.',
           'Un autre fournisseur utilise déjà ce nom.',
           'APPLICATION FERMÉE',
           'FERMETURE_DE_L_APPLICATION',
           'OUVERTURE_DE_L_APPLICATION',
           'Au moins un administrateur actif est requis.',
           'MOYENNE PAR SERVICE',
           'COÛT MOYEN DE RÉPARATION',
           'REVENU MOYEN',
           'REVENU MOYEN/SERVICE',
           'SAUVEGARDER LA BASE DE DONNÉES',
           'ÉCHEC DE LA SAUVEGARDE',
           'LE DOSSIER DE SAUVEGARDE NE PEUT PAS ÊTRE VIDE',
           'DOSSIER DE SAUVEGARDE :',
           'SAUVEGARDE RÉUSSIE',
           'Le service informatique à domicile de Bogor vous rappelle que la garantie de réparation de votre appareil '
           'expirera le :',
           'MARQUE :',
           'MARQUES',
           'Compilez HELP.HHP et copiez HELP.CHM dans le dossier de l’application.',
           'CALCULÉ AUTOMATIQUEMENT : DATE DE SERVICE + GARANTIE',
           'IMPOSSIBLE DE DÉTERMINER LE FOURNISSEUR SÉLECTIONNÉ',
           'MODIFIER LE MOT DE PASSE',
           'MODIFIER LE MOT DE PASSE DE',
           'TERMINÉ',
           'GESTIONNAIRE DE SERVICE INFORMATIQUE',
           'CONFIRMER LA SUPPRESSION',
           'CONFIRMER LE MOT DE PASSE',
           'CONFIRMER LA RESTAURATION',
           'NOM DU CONTACT',
           'CONTACT :<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'COPIÉ DANS LE PRESSE-PAPIERS',
           'COPIER L’ADRESSE',
           'COPIER LA CARTE GOOGLE',
           'COPIER LA FACTURE',
           'COPIER LE TÉLÉPHONE',
           'DROITS D’AUTEUR © ECOMTECH 2026',
           'NOMBRE',
           'CRÉER UN ADMINISTRATEUR',
           'CRÉER LE PREMIER COMPTE ADMINISTRATEUR',
           'CRÉÉ',
           'CRÉÉ PAR :',
           'DATE DE CRÉATION',
           'FICHIERS CSV',
           'MOT DE PASSE ACTUEL',
           'LE MOT DE PASSE ACTUEL EST INCORRECT.',
           'AIDE DE L’ONGLET ACTUEL',
           'NOM DU CLIENT',
           'CLIENTS SÉLECTIONNÉS POUR LE RAPPEL',
           'CLIENTS SÉLECTIONNÉS POUR LE RAPPEL :',
           'Gestionnaire de service informatique 4.4 sécurisé',
           'ÉCHEC DE L’ACTUALISATION DU TABLEAU DE BORD',
           'REVUE DU TABLEAU DE BORD',
           'LES DONNÉES ONT ÉTÉ EXPORTÉES VERS :',
           'EMPLACEMENT DE LA BASE DE DONNÉES REQUIS',
           'LE CHEMIN DE LA BASE DE DONNÉES NE PEUT PAS ÊTRE VIDE',
           'CHEMIN DE LA BASE DE DONNÉES :',
           'LES CHEMINS DE LA BASE DE DONNÉES ONT ÉTÉ MIS À JOUR AVEC SUCCÈS !',
           'BASE DE DONNÉES RESTAURÉE DEPUIS LA SAUVEGARDE',
           'BASE DE DONNÉES RESTAURÉE AVEC SUCCÈS DEPUIS LA SAUVEGARDE !',
           'BASE_DE_DONNÉES_RESTAURÉE',
           'DATE',
           'CHER/CHÈRE [CUSTOMER],',
           'DÉTAILS',
           'STATISTIQUES DES MARQUES D’APPAREILS',
           'MODÈLE DE L’APPAREIL',
           'TYPES D’APPAREILS',
           'NOM D’AFFICHAGE',
           'DON :<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Le lecteur D: n’est pas disponible. Veuillez choisir un autre dossier.',
           'Le lecteur D: est introuvable.\nVeuillez créer ou sélectionner un dossier pour la base de données.',
           'LECTEUR INTROUVABLE',
           'ÉDITEUR DE DONNÉES FACTICES OUVERT',
           'PIÈCE EN DOUBLE',
           'FOURNISSEUR EN DOUBLE',
           'MODE ÉDITION',
           'MODE ÉDITION : MODIFIEZ LES CHAMPS PUIS CLIQUEZ SUR ENREGISTRER',
           'MODE ÉDITION : MODIFIEZ LES CHAMPS PUIS CLIQUEZ SUR METTRE À JOUR',
           'ACTIVER / DÉSACTIVER',
           'DATE DE FIN :',
           'SAISISSEZ LE NOM D’UTILISATEUR ET LE MOT DE PASSE.',
           'FICHIERS EXCEL',
           'EXPIRE DANS 30 JOURS',
           'EXPIRE BIENTÔT',
           'EXPIRE BIENTÔT (< 7 JOURS)',
           'EXPIRE BIENTÔT (≤ 30 JOURS)',
           'EXPIRE BIENTÔT (≤ 7 JOURS)',
           'EXPIRE BIENTÔT (≤30 jours)',
           'DATE D’EXPIRATION À PARTIR DU',
           'DATE D’EXPIRATION À PARTIR DU :',
           'EXPORTATION TERMINÉE',
           'ERREUR D’EXPORTATION',
           'ÉCHEC DE L’EXPORTATION',
           'EXPORTATION RÉUSSIE',
           'EXPORTÉ VERS :',
           'ÉCHEC DE L’EXPORTATION',
           'RAPPORT FINANCIER',
           'PREMIÈRE SESSION ADMINISTRATEUR',
           'CONFIGURATION INITIALE DE LA BASE DE DONNÉES',
           'CONFIGURATION INITIALE :\nVeuillez choisir un dossier pour enregistrer la nouvelle base de données.',
           'FORMULAIRE EFFACÉ',
           'LA DATE DE DÉBUT NE PEUT PAS ÊTRE POSTÉRIEURE À LA DATE DE FIN',
           'ADRESSE COMPLÈTE',
           'CARTE GOOGLE',
           'LIEN GOOGLE MAPS COPIÉ',
           'LIEN GOOGLE MAPS VIDE',
           'LIEN GOOGLE MAPS VIDE (PRESSE-PAPIERS EFFACÉ)',
           'CARTE GOOGLE :',
           'CONTENU DE L’AIDE',
           'ID (ROWID)',
           'IMPORTATION TERMINÉE',
           'ÉCHEC DE L’IMPORTATION',
           'LE DOSSIER D’IMPORTATION/EXPORTATION NE PEUT PAS ÊTRE VIDE',
           'DOSSIER D’IMPORTATION/EXPORTATION :',
           'EN COURS',
           'EN RÉPARATION',
           'IDENTIFIANTS INVALIDES',
           'PLAGE DE DATES INVALIDE',
           'LECTEUR INVALIDE',
           'CHEMIN INVALIDE',
           'RÔLE UTILISATEUR INVALIDE.',
           'NOM D’UTILISATEUR INVALIDE',
           'NOM D’UTILISATEUR OU MOT DE PASSE INVALIDE.',
           'PLAGE DE DATES DE GARANTIE INVALIDE : LA DATE DE DÉBUT EST POSTÉRIEURE À LA DATE DE FIN',
           'NUMÉRO DE FACTURE',
           'PROBLÈMES',
           'DERNIÈRE CONNEXION',
           'LES ANCIENS FICHIERS .XLS NE SONT PAS PRIS EN CHARGE. ENREGISTREZ D’ABORD LE FICHIER AU FORMAT .XLSX OU '
           'CSV.',
           'VERROUILLÉ JUSQU’À',
           'ENREGISTRER UN RAPPEL',
           'ENREGISTRER LES RAPPELS',
           'ENREGISTRER UN RAPPEL DE GARANTIE',
           'CONNEXION',
           'CONNEXION RÉUSSIE.',
           'CONNEXION_BLOQUÉE',
           'ÉCHEC_DE_CONNEXION',
           'CONNEXION_RÉUSSIE',
           'DÉCONNEXION',
           'PIÈCES À STOCK FAIBLE',
           'FENÊTRE PRINCIPALE OUVERTE',
           '4 CARACTÈRES MINIMUM. TOUT FORMAT DE LETTRES, CHIFFRES, ESPACES OU SYMBOLES EST AUTORISÉ.',
           'MODÈLE DE L’APPAREIL',
           'MODÈLES',
           'CRÉATEUR MULTILINGUE DE DONNÉES FACTICES OUVERT',
           'MON ACTIVITÉ RÉCENTE',
           'S/O',
           'NOM DU CLIENT',
           'LE NOM EST REQUIS',
           'NOM :',
           'NOUVEAU MOT DE PASSE',
           'AUCUNE DONNÉE',
           'AUCUNE BASE DE DONNÉES SÉLECTIONNÉE',
           'AUCUN MOT DE PASSE PAR DÉFAUT N’EST UTILISÉ. LES MOTS DE PASSE SONT STOCKÉS UNIQUEMENT SOUS FORME DE '
           'HACHAGES SCRYPT UNIQUES ET SALÉS. CONSERVEZ CE MOT DE PASSE ADMINISTRATEUR EN LIEU SÛR.',
           'AUCUNE BASE DE DONNÉES EXISTANTE N’A ÉTÉ SÉLECTIONNÉE. L’APPLICATION VA CONTINUER AVEC LA CRÉATION D’UNE '
           'NOUVELLE BASE DE DONNÉES.',
           'AUCUN MESSAGE SMS OU WHATSAPP N’A ÉTÉ ENVOYÉ.',
           'AUCUNE DONNÉE D’ÉTAT',
           'AUCUNE GARANTIE',
           'AUCUNE GARANTIE D’AUCUNE SORTE N’EST FOURNIE',
           'NUMÉRO DE TÉLÉPHONE',
           'INTROUVABLE',
           'OUVREZ UNE BASE DE DONNÉES EXISTANTE OU CRÉEZ-EN UNE NOUVELLE POUR CETTE APPLICATION.',
           'OUVRIR LE CONTENU DE L’AIDE DE L’APPLICATION',
           'OUVRIR L’ÉDITEUR DE DONNÉES FACTICES',
           'OUVRIR LA CARTE GOOGLE',
           'OUVRIR L’AIDE DE L’ONGLET ACTIF',
           'NOM DE LA PIÈCE',
           'LE NOM DE LA PIÈCE EST REQUIS',
           'TYPES DE PIÈCES',
           'PIÈCES ACTUALISÉES',
           'MOT DE PASSE',
           'MOT DE PASSE MODIFIÉ',
           'MOT DE PASSE MODIFIÉ.',
           'LE MOT DE PASSE EST TROP LONG.',
           'LES MOTS DE PASSE NE CORRESPONDENT PAS',
           'LE MOT DE PASSE DOIT CONTENIR AU MOINS 4 CARACTÈRES.',
           'MOT DE PASSE NON MODIFIÉ',
           'MOT DE PASSE RÉINITIALISÉ',
           'MOT_DE_PASSE_MODIFIÉ',
           'ÉCHEC_DE_MODIFICATION_DU_MOT_DE_PASSE',
           'MOT_DE_PASSE_RÉINITIALISÉ',
           'NUMÉRO DE TÉLÉPHONE',
           'TÉLÉPHONE :',
           'PLACEZ DUMMY_CREATOR.EXE OU DUMMY_CREATOR.PY À CÔTÉ DE L’APPLICATION.',
           'VEUILLEZ NOUS CONTACTER SI VOUS AVEZ BESOIN D’UN SERVICE SUPPLÉMENTAIRE.',
           'VEUILLEZ SÉLECTIONNER UN FICHIER XLSX OU CSV',
           'VEUILLEZ SÉLECTIONNER UN FICHIER EXCEL OU CSV',
           'FACTURE D’ACHAT :',
           'QTÉ',
           'ACTIVITÉ RÉCENTE DE SÉCURITÉ ET DE DONNÉES',
           'ENREGISTREMENTS DE SERVICE RÉCENTS',
           'ACTUALISER LE TABLEAU DE BORD',
           'RECHARGER LES DONNÉES DES LISTES DÉROULANTES',
           'MODÈLE DE MESSAGE DE RAPPEL',
           'RAPPELS ENREGISTRÉS',
           'RAPPELS ENVOYÉS',
           'COÛT DE RÉPARATION',
           'ÉCHEC DE L’EXPORTATION DU RAPPORT',
           'RAPPORT EXPORTÉ',
           'RÉINITIALISER LE MOT DE PASSE',
           'RÉINITIALISER LE MOT DE PASSE UTILISATEUR',
           'RÉINITIALISER LA PLAGE DE GARANTIE À PARTIR D’AUJOURD’HUI',
           'RESTAURER LA BASE DE DONNÉES',
           'ÉCHEC DE LA RESTAURATION',
           'RESTAURATION RÉUSSIE',
           'RÔLE',
           'ENREGISTRER LE FICHIER D’EXPORTATION',
           'ENREGISTRER L’EXPORTATION DU RAPPORT',
           'RECHERCHER PAR NOM, TÉLÉPHONE, APPAREIL OU PROBLÈME...',
           'CONFIGURATION SÉCURISÉE DE L’ADMINISTRATEUR',
           'CONNEXION SÉCURISÉE',
           'SÉLECTIONNEZ UN CONTACT À SUPPRIMER',
           'SÉLECTIONNEZ UN CONTACT À MODIFIER',
           'SÉLECTIONNEZ UNE PIÈCE À SUPPRIMER',
           'SÉLECTIONNEZ UNE PIÈCE À MODIFIER',
           'SÉLECTIONNEZ UNE PIÈCE À METTRE À JOUR',
           'SÉLECTIONNEZ UN FOURNISSEUR À SUPPRIMER',
           'SÉLECTIONNEZ UN FOURNISSEUR À MODIFIER',
           'SÉLECTIONNEZ UN FOURNISSEUR À METTRE À JOUR',
           'SÉLECTIONNEZ D’ABORD UN UTILISATEUR.',
           'SÉLECTIONNEZ UNE LIGNE DE GARANTIE À MODIFIER',
           'SÉLECTIONNEZ LA DEVISE DE L’APPLICATION :',
           'SÉLECTIONNEZ LA LANGUE ET LA RÉGION DE L’APPLICATION :',
           'SÉLECTIONNEZ LA DEVISE PAR DÉFAUT DE L’APPLICATION :',
           'SÉLECTIONNER LA LANGUE / RÉGION',
           'SÉLECTIONNEZ LES LIGNES À EXPORTER',
           'SÉLECTIONNEZ LES GARANTIES POUR ENREGISTRER LES RAPPELS',
           'SÉLECTIONNEZ LES GARANTIES POUR ENVOYER LES RAPPELS',
           'ENVOYER LES RAPPELS',
           'DESCRIPTION DU SERVICE',
           'SERVICE INFORMATIQUE À DOMICILE À BOGOR',
           'GRAPHIQUE D’ÉTAT DU SERVICE',
           'RÉSUMÉ DU SERVICE',
           'RAPPORT RÉCAPITULATIF DU SERVICE',
           'DÉFINIR UN NOUVEAU MOT DE PASSE POUR',
           'DÉFINIR L’EMPLACEMENT DE LA BASE DE DONNÉES',
           'DÉFINIR LA DATE DE SERVICE À AUJOURD’HUI',
           'PARAMÈTRES ENREGISTRÉS',
           'ÉCHEC DE LA CONFIGURATION',
           'AFFICHER LE MOT DE PASSE',
           'AFFICHER LES MOTS DE PASSE',
           'CONNECTEZ-VOUS POUR CONTINUER',
           'CONNECTÉ',
           'CONNECTÉ EN TANT QUE',
           'PERSONNEL',
           'L’ACCÈS DU PERSONNEL EST LIMITÉ',
           'LE PERSONNEL PEUT AJOUTER DE NOUVEAUX ENREGISTREMENTS DE SERVICE ET CONSULTER LES DONNÉES OPÉRATIONNELLES. '
           'LA MODIFICATION, LA SUPPRESSION, L’IMPORTATION, L’EXPORTATION, LA SAUVEGARDE, LES PARAMÈTRES '
           'ADMINISTRATIFS, LA GESTION DES UTILISATEURS, LES MODIFICATIONS DES FOURNISSEURS, LES MODIFICATIONS DES '
           'PIÈCES, LES RAPPORTS ET LES RAPPELS NÉCESSITENT UN ADMINISTRATEUR.',
           'ÉTATS',
           'EMPLACEMENT DE STOCKAGE',
           'ADRESSE DU FOURNISSEUR',
           'NOM DU FOURNISSEUR',
           'LE NOM DU FOURNISSEUR EST REQUIS',
           'FOURNISSEURS ACTUALISÉS',
           'SYSTÈME',
           'NOM DU TECHNICIEN',
           'PERFORMANCES DU TECHNICIEN',
           'RAPPORT DE PERFORMANCE DU TECHNICIEN',
           'TECHNICIENS',
           'MERCI,',
           'CE NOM D’UTILISATEUR EXISTE DÉJÀ.',
           'LE LECTEUR DU DOSSIER DE SAUVEGARDE N’EXISTE PAS.',
           'LE LECTEUR DU CHEMIN DE LA BASE DE DONNÉES N’EXISTE PAS.',
           'LE LECTEUR DU CHEMIN D’EXPORTATION N’EXISTE PAS.',
           'LE LECTEUR DU DOSSIER D’IMPORTATION/EXPORTATION N’EXISTE PAS.',
           'LE MOT DE PASSE A ÉTÉ RÉINITIALISÉ.',
           'LES MOTS DE PASSE NE CORRESPONDENT PAS.',
           'CE COMPTE EST DÉSACTIVÉ. CONTACTEZ UN ADMINISTRATEUR.',
           'CETTE ACTION EST RÉSERVÉE AUX COMPTES ADMINISTRATEUR.',
           'CE LOGICIEL EST FOURNI TEL QUEL',
           'HEURE',
           'VERS EXCEL',
           'TOTAL DES GARANTIES ACTIVES',
           'TOTAL DES CLIENTS',
           'TOTAL DES APPAREILS',
           'TOTAL EXPIRÉ',
           'TOTAL EXPIRANT BIENTÔT',
           'REVENU TOTAL',
           'TOTAL SÉLECTIONNÉ',
           'TOTAL DES ENREGISTREMENTS DE SERVICE',
           'VALEUR TOTALE DU SERVICE',
           'TOTAL DES SERVICES',
           'TYPE :',
           'MODÈLES UNIQUES',
           'IDENTIFIANTS INCONNUS OU INVALIDES',
           'TABLE INCONNUE À EXPORTER.',
           'DÉVERROUILLER',
           'FORMAT NON PRIS EN CHARGE',
           'UTILISER D:\\BACKUP (PAR DÉFAUT)',
           'UTILISATEUR',
           'GESTION DES UTILISATEURS',
           'UTILISATEUR NON CRÉÉ',
           'UTILISATEUR INTROUVABLE.',
           'L’UTILISATEUR A DEMANDÉ LA DÉCONNEXION',
           'ÉTAT DE L’UTILISATEUR MIS À JOUR.',
           'NOM D’UTILISATEUR',
           'LE NOM D’UTILISATEUR EXISTE',
           'LE NOM D’UTILISATEUR DOIT COMPORTER DE 3 À 32 CARACTÈRES ET UTILISER DES LETTRES, DES CHIFFRES, UN POINT, '
           'UN TIRET OU UN TRAIT DE SOULIGNEMENT.',
           'UTILISATEUR_CRÉÉ',
           'UTILISATEUR_DÉSACTIVÉ',
           'UTILISATEUR_ACTIVÉ',
           'UTILISATEUR_DÉVERROUILLÉ',
           'EN ATTENTE DE PIÈCES',
           'PIÈCES EN ATTENTE',
           'AVERTISSEMENT : LA RESTAURATION REMPLACERA TOUTES LES DONNÉES ACTUELLES PAR LES DONNÉES DE SAUVEGARDE.\n'
           'CETTE ACTION EST IRRÉVERSIBLE.\n'
           '\n'
           'VOULEZ-VOUS CONTINUER ?',
           'PÉRIODES DE GARANTIE',
           'ÉTAT DE LA GARANTIE :',
           'RÉSUMÉ DE LA GARANTIE',
           'RAPPORT RÉCAPITULATIF DE GARANTIE',
           'MOT DE PASSE FAIBLE',
           'QUE VOULEZ-VOUS FAIRE AVEC LA BASE DE DONNÉES ?',
           'FENÊTRE FERMÉE',
           'VOUS NE POUVEZ PAS DÉSACTIVER VOTRE COMPTE ACTUEL.',
           'VOTRE MOT DE PASSE A ÉTÉ MODIFIÉ AVEC SUCCÈS.',
           '⚠️ AVERTISSEMENT !',
           '✏️ MODIFIER',
           '✏️ MODIFIER LE CONTACT',
           '❌ SUPPRIMER',
           '➕ AJOUTER',
           '➕ AJOUTER UNE PIÈCE',
           '📄 VUE TEXTE',
           '📋 VUE TABLEAU',
           '📝 ENREGISTRER UN RAPPEL',
           '📝 ENREGISTRER LES RAPPELS',
           '📤 EXPORTER LA SÉLECTION',
           '🔃 ACTUALISER',
           '🔄 ACTUALISER',
           '🗑️ EFFACER'],
 'de_DE': ['Ein neues Teil muss von einem Administrator erstellt werden.',
           'Ein neuer Lieferant muss von einem Administrator erstellt werden.',
           'Ein Teil mit diesem Namen ist bereits vorhanden. Wählen Sie es aus und verwenden Sie Speichern, um es zu '
           'aktualisieren.',
           'Ein Lieferant mit diesem Namen ist bereits vorhanden. Wählen Sie ihn aus und verwenden Sie Speichern, um '
           'ihn zu aktualisieren.',
           'ZUGRIFF_VERWEIGERT',
           'KONTO',
           'KONTO DEAKTIVIERT',
           'KONTO GESPERRT',
           'AKTION BLOCKIERT',
           'AKTIVE GARANTIEN',
           'ANWENDUNGSBENUTZER HINZUFÜGEN',
           'BENUTZER HINZUFÜGEN',
           'ZUSÄTZLICHE NOTIZEN',
           'VOLLSTÄNDIGE ADRESSE',
           'ADMIN',
           'ADMINISTRATOR',
           'ADMINISTRATORAKTION BLOCKIERT',
           'NUR ADMINISTRATOR',
           'ADMINISTRATOR ERFORDERLICH',
           'ADMINISTRATOR ERFORDERLICH.',
           'Ein anderes Teil verwendet diesen Namen bereits.',
           'Ein anderer Lieferant verwendet diesen Namen bereits.',
           'ANWENDUNG GESCHLOSSEN',
           'ANWENDUNG_BEENDET',
           'ANWENDUNG_GEÖFFNET',
           'Mindestens ein aktiver Administrator ist erforderlich.',
           'DURCHSCHNITT PRO SERVICE',
           'DURCHSCHNITTLICHE REPARATURKOSTEN',
           'DURCHSCHNITTLICHER UMSATZ',
           'DURCHSCHNITTSUMSATZ/SERVICE',
           'DATENBANK SICHERN',
           'SICHERUNG FEHLGESCHLAGEN',
           'DER SICHERUNGSORDNER DARF NICHT LEER SEIN',
           'SICHERUNGSORDNER:',
           'SICHERUNG ERFOLGREICH',
           'Der Computer-Service vor Ort in Bogor erinnert Sie daran, dass die Reparaturgarantie Ihres Geräts am '
           'folgenden Datum abläuft:',
           'MARKE:',
           'MARKEN',
           'Erstellen Sie HELP.HHP und kopieren Sie HELP.CHM in den Anwendungsordner.',
           'AUTOMATISCH BERECHNET: SERVICEDATUM + GARANTIE',
           'AUSGEWÄHLTER LIEFERANT KANN NICHT ERMITTELT WERDEN',
           'PASSWORT ÄNDERN',
           'PASSWORT ÄNDERN FÜR',
           'ABGESCHLOSSEN',
           'COMPUTER-SERVICE-MANAGER',
           'LÖSCHEN BESTÄTIGEN',
           'PASSWORT BESTÄTIGEN',
           'WIEDERHERSTELLUNG BESTÄTIGEN',
           'KONTAKTNAME',
           'KONTAKT:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'IN DIE ZWISCHENABLAGE KOPIERT',
           'ADRESSE KOPIEREN',
           'GOOGLE-KARTE KOPIEREN',
           'RECHNUNG KOPIEREN',
           'TELEFONNUMMER KOPIEREN',
           'URHEBERRECHT © ECOMTECH 2026',
           'ANZAHL',
           'ADMINISTRATOR ERSTELLEN',
           'ERSTES ADMINISTRATORKONTO ERSTELLEN',
           'ERSTELLT',
           'ERSTELLT VON:',
           'ERSTELLDATUM',
           'CSV-DATEIEN',
           'AKTUELLES PASSWORT',
           'DAS AKTUELLE PASSWORT IST FALSCH.',
           'HILFE FÜR AKTUELLEN TAB',
           'KUNDENNAME',
           'FÜR ERINNERUNG AUSGEWÄHLTE KUNDEN',
           'FÜR ERINNERUNG AUSGEWÄHLTE KUNDEN:',
           'Computer-Service-Manager 4.4 Sicher',
           'DASHBOARD-AKTUALISIERUNG FEHLGESCHLAGEN',
           'DASHBOARD-ÜBERSICHT',
           'DATEN WURDEN EXPORTIERT NACH:',
           'DATENBANKSPEICHERORT ERFORDERLICH',
           'DER DATENBANKPFAD DARF NICHT LEER SEIN',
           'DATENBANKPFAD:',
           'DATENBANKPFADE ERFOLGREICH AKTUALISIERT!',
           'DATENBANK AUS SICHERUNG WIEDERHERGESTELLT',
           'DATENBANK ERFOLGREICH AUS SICHERUNG WIEDERHERGESTELLT!',
           'DATENBANK_WIEDERHERGESTELLT',
           'DATUM',
           'GUTEN TAG [CUSTOMER],',
           'DETAILS',
           'GERÄTEMARKENSTATISTIK',
           'GERÄTEMODELL',
           'GERÄTETYPEN',
           'ANZEIGENAME',
           'SPENDE:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Laufwerk D: ist nicht verfügbar. Bitte wählen Sie einen anderen Ordner.',
           'Laufwerk D: wurde nicht gefunden.\nBitte erstellen oder wählen Sie einen Ordner für die Datenbank.',
           'LAUFWERK NICHT GEFUNDEN',
           'DUMMY-DATENEDITOR GEÖFFNET',
           'DOPPELTES TEIL',
           'DOPPELTER LIEFERANT',
           'BEARBEITUNGSMODUS',
           'BEARBEITUNGSMODUS: FELDER AKTUALISIEREN UND DANN AUF SPEICHERN KLICKEN',
           'BEARBEITUNGSMODUS: FELDER AKTUALISIEREN UND DANN AUF AKTUALISIEREN KLICKEN',
           'AKTIVIEREN / DEAKTIVIEREN',
           'ENDDATUM:',
           'GEBEN SIE BENUTZERNAME UND PASSWORT EIN.',
           'EXCEL-DATEIEN',
           'LÄUFT IN 30 TAGEN AB',
           'LÄUFT BALD AB',
           'LÄUFT BALD AB (< 7 TAGE)',
           'LÄUFT BALD AB (≤ 30 TAGE)',
           'LÄUFT BALD AB (≤ 7 TAGE)',
           'LÄUFT BALD AB (≤30 Tage)',
           'ABLAUFDATUM VON',
           'ABLAUFDATUM VON:',
           'EXPORT ABGESCHLOSSEN',
           'EXPORTFEHLER',
           'EXPORT FEHLGESCHLAGEN',
           'EXPORT ERFOLGREICH',
           'EXPORTIERT NACH:',
           'EXPORT FEHLGESCHLAGEN',
           'FINANZBERICHT',
           'ERSTE ADMINISTRATORSITZUNG',
           'ERSTMALIGE DATENBANKEINRICHTUNG',
           'ERSTMALIGE EINRICHTUNG:\nBitte wählen Sie einen Ordner zum Speichern der neuen Datenbank.',
           'FORMULAR GELEERT',
           'DAS VON-DATUM DARF NICHT NACH DEM BIS-DATUM LIEGEN',
           'VOLLSTÄNDIGE ADRESSE',
           'GOOGLE-KARTE',
           'GOOGLE-MAPS-LINK KOPIERT',
           'GOOGLE-MAPS-LINK LEER',
           'GOOGLE-MAPS-LINK LEER (ZWISCHENABLAGE GELEERT)',
           'GOOGLE-KARTE:',
           'HILFEINHALT',
           'ID (ROWID)',
           'IMPORT ABGESCHLOSSEN',
           'IMPORT FEHLGESCHLAGEN',
           'DER IMPORT-/EXPORTORDNER DARF NICHT LEER SEIN',
           'IMPORT-/EXPORTORDNER:',
           'IN BEARBEITUNG',
           'IN REPARATUR',
           'UNGÜLTIGE ANMELDEDATEN',
           'UNGÜLTIGER DATUMSBEREICH',
           'UNGÜLTIGES LAUFWERK',
           'UNGÜLTIGER PFAD',
           'UNGÜLTIGE BENUTZERROLLE.',
           'UNGÜLTIGER BENUTZERNAME',
           'UNGÜLTIGER BENUTZERNAME ODER UNGÜLTIGES PASSWORT.',
           'UNGÜLTIGER GARANTIEDATUMSBEREICH: VON LIEGT NACH BIS',
           'RECHNUNGSNUMMER',
           'PROBLEME',
           'LETZTE ANMELDUNG',
           'ÄLTERE .XLS-DATEIEN WERDEN NICHT UNTERSTÜTZT. SPEICHERN SIE DIE DATEI ZUERST ALS .XLSX ODER CSV.',
           'GESPERRT BIS',
           'ERINNERUNG PROTOKOLLIEREN',
           'ERINNERUNGEN PROTOKOLLIEREN',
           'GARANTIEERINNERUNG PROTOKOLLIEREN',
           'ANMELDEN',
           'ANMELDUNG ERFOLGREICH.',
           'ANMELDUNG_BLOCKIERT',
           'ANMELDUNG_FEHLGESCHLAGEN',
           'ANMELDUNG_ERFOLGREICH',
           'ABMELDEN',
           'TEILE MIT NIEDRIGEM BESTAND',
           'HAUPTFENSTER GEÖFFNET',
           'MINDESTENS 4 ZEICHEN. JEDE KOMBINATION AUS BUCHSTABEN, ZAHLEN, LEERZEICHEN ODER SYMBOLEN IST ZULÄSSIG.',
           'GERÄTEMODELL',
           'MODELLE',
           'MEHRSPRACHIGER DUMMY-DATEN-ERSTELLER GEÖFFNET',
           'MEINE LETZTEN AKTIVITÄTEN',
           'K. A.',
           'KUNDENNAME',
           'NAME IST ERFORDERLICH',
           'NAME:',
           'NEUES PASSWORT',
           'KEINE DATEN',
           'KEINE DATENBANK AUSGEWÄHLT',
           'ES WIRD KEIN STANDARDPASSWORT VERWENDET. PASSWÖRTER WERDEN NUR ALS EINDEUTIGE, GESALZENE SCRYPT-HASHES '
           'GESPEICHERT. BEWAHREN SIE DIESES ADMINISTRATORPASSWORT SICHER AUF.',
           'ES WURDE KEINE VORHANDENE DATENBANK AUSGEWÄHLT. DIE ANWENDUNG FÄHRT MIT DER ERSTELLUNG EINER NEUEN '
           'DATENBANK FORT.',
           'ES WURDE KEINE SMS- ODER WHATSAPP-NACHRICHT GESENDET.',
           'KEINE STATUSDATEN',
           'KEINE GARANTIE',
           'ES WIRD KEINERLEI GARANTIE GEWÄHRT',
           'TELEFONNUMMER',
           'NICHT GEFUNDEN',
           'ÖFFNEN SIE EINE VORHANDENE DATENBANK ODER ERSTELLEN SIE EINE NEUE DATENBANK FÜR DIESE ANWENDUNG.',
           'HILFEINHALT DER ANWENDUNG ÖFFNEN',
           'DUMMY-DATENEDITOR ÖFFNEN',
           'GOOGLE-KARTE ÖFFNEN',
           'HILFE FÜR DEN AKTIVEN TAB ÖFFNEN',
           'TEILENAME',
           'TEILENAME IST ERFORDERLICH',
           'TEILETYPEN',
           'TEILE AKTUALISIERT',
           'PASSWORT',
           'PASSWORT GEÄNDERT',
           'PASSWORT GEÄNDERT.',
           'DAS PASSWORT IST ZU LANG.',
           'PASSWÖRTER STIMMEN NICHT ÜBEREIN',
           'DAS PASSWORT MUSS MINDESTENS 4 ZEICHEN ENTHALTEN.',
           'PASSWORT NICHT GEÄNDERT',
           'PASSWORT ZURÜCKGESETZT',
           'PASSWORT_GEÄNDERT',
           'PASSWORTÄNDERUNG_FEHLGESCHLAGEN',
           'PASSWORT_ZURÜCKGESETZT',
           'TELEFONNUMMER',
           'TELEFON:',
           'LEGEN SIE DUMMY_CREATOR.EXE ODER DUMMY_CREATOR.PY NEBEN DER ANWENDUNG AB.',
           'BITTE KONTAKTIEREN SIE UNS, WENN SIE WEITEREN SERVICE BENÖTIGEN.',
           'BITTE WÄHLEN SIE EINE XLSX- ODER CSV-DATEI',
           'BITTE WÄHLEN SIE EINE EXCEL- ODER CSV-DATEI',
           'EINKAUFSRECHNUNG:',
           'MENGE',
           'LETZTE SICHERHEITS- UND DATENAKTIVITÄTEN',
           'LETZTE SERVICEEINTRÄGE',
           'DASHBOARD AKTUALISIEREN',
           'DROPDOWN-DATEN NEU LADEN',
           'VORLAGE FÜR ERINNERUNGSNACHRICHT',
           'ERINNERUNGEN PROTOKOLLIERT',
           'ERINNERUNGEN GESENDET',
           'REPARATURKOSTEN',
           'BERICHTSEXPORT FEHLGESCHLAGEN',
           'BERICHT EXPORTIERT',
           'PASSWORT ZURÜCKSETZEN',
           'BENUTZEPASSWORT ZURÜCKSETZEN',
           'GARANTIEZEITRAUM AB HEUTE ZURÜCKSETZEN',
           'DATENBANK WIEDERHERSTELLEN',
           'WIEDERHERSTELLUNG FEHLGESCHLAGEN',
           'WIEDERHERSTELLUNG ERFOLGREICH',
           'ROLLE',
           'EXPORTDATEI SPEICHERN',
           'BERICHTSEXPORT SPEICHERN',
           'NACH NAME, TELEFON, GERÄT ODER PROBLEM SUCHEN...',
           'SICHERE ADMINISTRATOREINRICHTUNG',
           'SICHERE ANMELDUNG',
           'WÄHLEN SIE EINEN KONTAKT ZUM LÖSCHEN AUS',
           'WÄHLEN SIE EINEN KONTAKT ZUM BEARBEITEN AUS',
           'WÄHLEN SIE EIN TEIL ZUM LÖSCHEN AUS',
           'WÄHLEN SIE EIN TEIL ZUM BEARBEITEN AUS',
           'WÄHLEN SIE EIN TEIL ZUM AKTUALISIEREN AUS',
           'WÄHLEN SIE EINEN LIEFERANTEN ZUM LÖSCHEN AUS',
           'WÄHLEN SIE EINEN LIEFERANTEN ZUM BEARBEITEN AUS',
           'WÄHLEN SIE EINEN LIEFERANTEN ZUM AKTUALISIEREN AUS',
           'WÄHLEN SIE ZUERST EINEN BENUTZER AUS.',
           'WÄHLEN SIE EINE GARANTIEZEILE ZUM BEARBEITEN AUS',
           'ANWENDUNGSWÄHRUNG AUSWÄHLEN:',
           'ANWENDUNGSSPRACHE UND REGION AUSWÄHLEN:',
           'STANDARDWÄHRUNG DER ANWENDUNG AUSWÄHLEN:',
           'SPRACHE / REGION AUSWÄHLEN',
           'ZEILEN ZUM EXPORTIEREN AUSWÄHLEN',
           'GARANTIEN ZUM PROTOKOLLIEREN VON ERINNERUNGEN AUSWÄHLEN',
           'GARANTIEN ZUM SENDEN VON ERINNERUNGEN AUSWÄHLEN',
           'ERINNERUNGEN SENDEN',
           'SERVICEBESCHREIBUNG',
           'COMPUTER-SERVICE VOR ORT IN BOGOR',
           'SERVICESTATUS-GRAFIK',
           'SERVICEZUSAMMENFASSUNG',
           'BERICHT ZUR SERVICEZUSAMMENFASSUNG',
           'NEUES PASSWORT FESTLEGEN FÜR',
           'DATENBANKSPEICHERORT FESTLEGEN',
           'SERVICEDATUM AUF HEUTE SETZEN',
           'EINSTELLUNGEN GESPEICHERT',
           'EINRICHTUNG FEHLGESCHLAGEN',
           'PASSWORT ANZEIGEN',
           'PASSWÖRTER ANZEIGEN',
           'ZUM FORTFAHREN ANMELDEN',
           'ANGEMELDET',
           'ANGEMELDET ALS',
           'MITARBEITER',
           'MITARBEITERZUGRIFF IST EINGESCHRÄNKT',
           'MITARBEITER KÖNNEN NEUE SERVICEEINTRÄGE HINZUFÜGEN UND BETRIEBSDATEN ANZEIGEN. BEARBEITEN, LÖSCHEN, '
           'IMPORT, EXPORT, SICHERUNG, ADMINISTRATIVE EINSTELLUNGEN, BENUTZERVERWALTUNG, LIEFERANTENÄNDERUNGEN, '
           'TEILEÄNDERUNGEN, BERICHTE UND ERINNERUNGEN ERFORDERN EINEN ADMINISTRATOR.',
           'STATUSWERTE',
           'LAGERORT',
           'LIEFERANTENADRESSE',
           'LIEFERANTENNAME',
           'LIEFERANTENNAME IST ERFORDERLICH',
           'LIEFERANTEN AKTUALISIERT',
           'SYSTEM',
           'TECHNIKERNAME',
           'TECHNIKERLEISTUNG',
           'BERICHT ZUR TECHNIKERLEISTUNG',
           'TECHNIKER',
           'VIELEN DANK,',
           'DIESER BENUTZERNAME EXISTIERT BEREITS.',
           'DAS LAUFWERK FÜR DEN SICHERUNGSORDNER EXISTIERT NICHT.',
           'DAS LAUFWERK FÜR DEN DATENBANKPFAD EXISTIERT NICHT.',
           'DAS LAUFWERK FÜR DEN EXPORTPFAD EXISTIERT NICHT.',
           'DAS LAUFWERK FÜR DEN IMPORT-/EXPORTORDNER EXISTIERT NICHT.',
           'DAS PASSWORT WURDE ZURÜCKGESETZT.',
           'DIE PASSWÖRTER STIMMEN NICHT ÜBEREIN.',
           'DIESES KONTO IST DEAKTIVIERT. WENDEN SIE SICH AN EINEN ADMINISTRATOR.',
           'DIESE AKTION IST AUF ADMINISTRATORKONTEN BESCHRÄNKT.',
           'DIESE SOFTWARE WIRD OHNE GEWÄHR BEREITGESTELLT',
           'ZEIT',
           'NACH EXCEL',
           'AKTIVE GARANTIEN GESAMT',
           'KUNDEN GESAMT',
           'GERÄTE GESAMT',
           'ABGELAUFEN GESAMT',
           'BALD ABLAUFEND GESAMT',
           'GESAMTUMSATZ',
           'AUSGEWÄHLT GESAMT',
           'SERVICEEINTRÄGE GESAMT',
           'GESAMTSERVICEWERT',
           'SERVICES GESAMT',
           'TYP:',
           'EINDEUTIGE MODELLE',
           'UNBEKANNTE ODER UNGÜLTIGE ANMELDEDATEN',
           'UNBEKANNTE TABELLE FÜR EXPORT.',
           'ENTSPERREN',
           'NICHT UNTERSTÜTZTES FORMAT',
           'D:\\BACKUP VERWENDEN (STANDARD)',
           'BENUTZER',
           'BENUTZERVERWALTUNG',
           'BENUTZER NICHT ERSTELLT',
           'BENUTZER NICHT GEFUNDEN.',
           'BENUTZER HAT ABMELDUNG ANGEFORDERT',
           'BENUTZERSTATUS AKTUALISIERT.',
           'BENUTZERNAME',
           'BENUTZERNAME EXISTIERT',
           'DER BENUTZERNAME MUSS 3–32 ZEICHEN LANG SEIN UND DARF BUCHSTABEN, ZAHLEN, PUNKT, BINDESTRICH ODER '
           'UNTERSTRICH ENTHALTEN.',
           'BENUTZER_ERSTELLT',
           'BENUTZER_DEAKTIVIERT',
           'BENUTZER_AKTIVIERT',
           'BENUTZER_ENTSPERRT',
           'WARTEN AUF TEILE',
           'WARTENDE TEILE',
           'WARNUNG: BEIM WIEDERHERSTELLEN WERDEN ALLE AKTUELLEN DATEN DURCH DIE SICHERUNGSDATEN ERSETZT.\n'
           'DIESE AKTION KANN NICHT RÜCKGÄNGIG GEMACHT WERDEN.\n'
           '\n'
           'MÖCHTEN SIE FORTFAHREN?',
           'GARANTIEZEITRÄUME',
           'GARANTIESTATUS:',
           'GARANTIEZUSAMMENFASSUNG',
           'BERICHT ZUR GARANTIEZUSAMMENFASSUNG',
           'SCHWACHES PASSWORT',
           'WAS MÖCHTEN SIE MIT DER DATENBANK TUN?',
           'FENSTER GESCHLOSSEN',
           'SIE KÖNNEN IHR AKTUELLES KONTO NICHT DEAKTIVIEREN.',
           'IHR PASSWORT WURDE ERFOLGREICH GEÄNDERT.',
           '⚠️ WARNUNG!',
           '✏️ BEARBEITEN',
           '✏️ KONTAKT BEARBEITEN',
           '❌ LÖSCHEN',
           '➕ HINZUFÜGEN',
           '➕ TEIL HINZUFÜGEN',
           '📄 TEXTANSICHT',
           '📋 TABELLENANSICHT',
           '📝 ERINNERUNG PROTOKOLLIEREN',
           '📝 ERINNERUNGEN PROTOKOLLIEREN',
           '📤 AUSWAHL EXPORTIEREN',
           '🔃 AKTUALISIEREN',
           '🔄 AKTUALISIEREN',
           '🗑️ LEEREN'],
 'pt_BR': ['Uma nova peça deve ser criada por um administrador.',
           'Um novo fornecedor deve ser criado por um administrador.',
           'Já existe uma peça com este nome. Selecione-a e use Salvar para atualizá-la.',
           'Já existe um fornecedor com este nome. Selecione-o e use Salvar para atualizá-lo.',
           'ACESSO_NEGADO',
           'CONTA',
           'CONTA DESATIVADA',
           'CONTA BLOQUEADA',
           'AÇÃO BLOQUEADA',
           'GARANTIAS ATIVAS',
           'ADICIONAR USUÁRIO DO APLICATIVO',
           'ADICIONAR USUÁRIO',
           'OBSERVAÇÕES ADICIONAIS',
           'ENDEREÇO COMPLETO',
           'ADMIN',
           'ADMINISTRADOR',
           'AÇÃO DO ADMINISTRADOR BLOQUEADA',
           'SOMENTE ADMINISTRADOR',
           'ADMINISTRADOR NECESSÁRIO',
           'ADMINISTRADOR NECESSÁRIO.',
           'Outra peça já usa este nome.',
           'Outro fornecedor já usa este nome.',
           'APLICATIVO FECHADO',
           'SAÍDA_DO_APLICATIVO',
           'ABERTURA_DO_APLICATIVO',
           'É necessário pelo menos um administrador ativo.',
           'MÉDIA POR SERVIÇO',
           'CUSTO MÉDIO DE REPARO',
           'RECEITA MÉDIA',
           'RECEITA MÉDIA/SERVIÇO',
           'FAZER BACKUP DO BANCO DE DADOS',
           'FALHA NO BACKUP',
           'A pasta de backup não pode ficar vazia',
           'PASTA DE BACKUP:',
           'BACKUP CONCLUÍDO',
           'O Serviço de Computador em Domicílio de Bogor lembra que a garantia do reparo do seu dispositivo vencerá '
           'em:',
           'MARCA:',
           'MARCAS',
           'Compile HELP.HHP e copie HELP.CHM para a pasta do aplicativo.',
           'Calculado automaticamente: data do serviço + garantia',
           'Não foi possível determinar o fornecedor selecionado',
           'ALTERAR SENHA',
           'ALTERAR SENHA DE',
           'CONCLUÍDO',
           'GERENCIADOR DE SERVIÇOS DE COMPUTADOR',
           'CONFIRMAR EXCLUSÃO',
           'CONFIRMAR SENHA',
           'CONFIRMAR RESTAURAÇÃO',
           'NOME DO CONTATO',
           'CONTATO:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'COPIADO PARA A ÁREA DE TRANSFERÊNCIA',
           'COPIAR ENDEREÇO',
           'COPIAR GOOGLE MAP',
           'COPIAR NOTA FISCAL',
           'COPIAR TELEFONE',
           'DIREITOS AUTORAIS © ECOMTECH 2026',
           'CONTAGEM',
           'CRIAR ADMINISTRADOR',
           'CRIAR A PRIMEIRA CONTA DE ADMINISTRADOR',
           'CRIADO',
           'CRIADO POR:',
           'DATA DE CRIAÇÃO',
           'ARQUIVOS CSV',
           'SENHA ATUAL',
           'A senha atual está incorreta.',
           'AJUDA DA ABA ATUAL',
           'NOME DO CLIENTE',
           'CLIENTES SELECIONADOS PARA LEMBRETE',
           'CLIENTES SELECIONADOS PARA LEMBRETE:',
           'Gerenciador de Serviços de Computador 4.4 Seguro',
           'FALHA AO ATUALIZAR O PAINEL',
           'VISÃO GERAL DO PAINEL',
           'Os dados foram exportados para:',
           'LOCAL DO BANCO DE DADOS NECESSÁRIO',
           'O caminho do banco de dados não pode ficar vazio',
           'CAMINHO DO BANCO DE DADOS:',
           'Caminhos do banco de dados atualizados com sucesso!',
           'BANCO DE DADOS RESTAURADO DO BACKUP',
           'Banco de dados restaurado com sucesso a partir do backup!',
           'BANCO_DE_DADOS_RESTAURADO',
           'DATA',
           'Prezado(a) [CUSTOMER],',
           'DETALHES',
           'ESTATÍSTICAS DE MARCAS DE DISPOSITIVOS',
           'MODELO DO DISPOSITIVO',
           'TIPOS DE DISPOSITIVO',
           'NOME DE EXIBIÇÃO',
           'DOAÇÃO:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'A unidade D: não está disponível. Escolha outra pasta.',
           'A unidade D: não foi encontrada.\nCrie ou selecione uma pasta para o banco de dados.',
           'UNIDADE NÃO ENCONTRADA',
           'EDITOR DE DADOS FICTÍCIOS ABERTO',
           'PEÇA DUPLICADA',
           'FORNECEDOR DUPLICADO',
           'MODO DE EDIÇÃO',
           'MODO DE EDIÇÃO: ATUALIZE OS CAMPOS E CLIQUE EM SALVAR',
           'MODO DE EDIÇÃO: ATUALIZE OS CAMPOS E CLIQUE EM ATUALIZAR',
           'ATIVAR / DESATIVAR',
           'DATA FINAL:',
           'Digite o nome de usuário e a senha.',
           'ARQUIVOS EXCEL',
           'VENCE EM 30 DIAS',
           'VENCENDO EM BREVE',
           'VENCENDO EM BREVE (< 7 DIAS)',
           'VENCENDO EM BREVE (≤ 30 DIAS)',
           'VENCENDO EM BREVE (≤ 7 DIAS)',
           'VENCENDO EM BREVE (≤30 dias)',
           'DATA DE VENCIMENTO A PARTIR DE',
           'DATA DE VENCIMENTO A PARTIR DE:',
           'EXPORTAÇÃO CONCLUÍDA',
           'ERRO DE EXPORTAÇÃO',
           'FALHA NA EXPORTAÇÃO',
           'EXPORTAÇÃO CONCLUÍDA',
           'EXPORTADO PARA:',
           'Falha ao exportar',
           'RELATÓRIO FINANCEIRO',
           'PRIMEIRA SESSÃO DO ADMINISTRADOR',
           'CONFIGURAÇÃO INICIAL DO BANCO DE DADOS',
           'CONFIGURAÇÃO INICIAL:\nEscolha uma pasta para salvar o novo banco de dados.',
           'FORMULÁRIO LIMPO',
           'A data inicial não pode ser posterior à data final',
           'ENDEREÇO COMPLETO',
           'GOOGLE MAP',
           'LINK DO GOOGLE MAP COPIADO',
           'LINK DO GOOGLE MAP VAZIO',
           'LINK DO GOOGLE MAP VAZIO (ÁREA DE TRANSFERÊNCIA LIMPA)',
           'GOOGLE MAP:',
           'CONTEÚDO DA AJUDA',
           'ID (ROWID)',
           'IMPORTAÇÃO CONCLUÍDA',
           'FALHA NA IMPORTAÇÃO',
           'A pasta de importação/exportação não pode ficar vazia',
           'PASTA DE IMPORTAÇÃO/EXPORTAÇÃO:',
           'EM ANDAMENTO',
           'EM REPARO',
           'CREDENCIAIS INVÁLIDAS',
           'INTERVALO DE DATAS INVÁLIDO',
           'UNIDADE INVÁLIDA',
           'CAMINHO INVÁLIDO',
           'Função de usuário inválida.',
           'NOME DE USUÁRIO INVÁLIDO',
           'Nome de usuário ou senha inválidos.',
           'Intervalo de datas de garantia inválido: a data inicial é posterior à final',
           'NÚMERO DA NOTA FISCAL',
           'PROBLEMAS',
           'ÚLTIMO LOGIN',
           'Arquivos .XLS antigos não são compatíveis. Salve o arquivo como .XLSX ou CSV primeiro.',
           'BLOQUEADO ATÉ',
           'REGISTRAR LEMBRETE',
           'REGISTRAR LEMBRETES',
           'REGISTRAR LEMBRETE DE GARANTIA',
           'ENTRAR',
           'Login realizado com sucesso.',
           'LOGIN_BLOQUEADO',
           'FALHA_NO_LOGIN',
           'LOGIN_BEM_SUCEDIDO',
           'SAIR',
           'PEÇAS COM ESTOQUE BAIXO',
           'JANELA PRINCIPAL ABERTA',
           'Mínimo de 4 caracteres. É permitido qualquer formato com letras, números, espaços ou símbolos.',
           'MODELO DO DISPOSITIVO',
           'MODELOS',
           'CRIADOR DE DADOS FICTÍCIOS MULTILÍNGUE ABERTO',
           'MINHA ATIVIDADE RECENTE',
           'N/D',
           'NOME DO CLIENTE',
           'O NOME É OBRIGATÓRIO',
           'NOME:',
           'NOVA SENHA',
           'SEM DADOS',
           'NENHUM BANCO DE DADOS SELECIONADO',
           'Nenhuma senha padrão é usada. As senhas são armazenadas apenas como hashes scrypt exclusivos, com salt. '
           'Mantenha esta senha de administrador em segurança.',
           'Nenhum banco de dados existente foi selecionado. O aplicativo continuará com a criação de um novo banco de '
           'dados.',
           'Nenhuma mensagem SMS ou WhatsApp foi enviada.',
           'SEM DADOS DE STATUS',
           'SEM GARANTIA',
           'Nenhuma garantia de qualquer tipo é fornecida',
           'NÚMERO DE TELEFONE',
           'NÃO ENCONTRADO',
           'Abra um banco de dados existente ou crie um novo banco de dados para este aplicativo.',
           'ABRIR O CONTEÚDO DA AJUDA DO APLICATIVO',
           'ABRIR O EDITOR DE DADOS FICTÍCIOS',
           'ABRIR GOOGLE MAP',
           'ABRIR A AJUDA DA ABA ATIVA',
           'NOME DA PEÇA',
           'O NOME DA PEÇA É OBRIGATÓRIO',
           'TIPOS DE PEÇA',
           'PEÇAS ATUALIZADAS',
           'SENHA',
           'SENHA ALTERADA',
           'Senha alterada.',
           'A senha é muito longa.',
           'SENHAS NÃO COINCIDEM',
           'A senha deve conter pelo menos 4 caracteres.',
           'SENHA NÃO ALTERADA',
           'SENHA REDEFINIDA',
           'SENHA_ALTERADA',
           'FALHA_AO_ALTERAR_SENHA',
           'SENHA_REDEFINIDA',
           'NÚMERO DE TELEFONE',
           'TELEFONE:',
           'Coloque DUMMY_CREATOR.EXE ou DUMMY_CREATOR.PY ao lado do aplicativo.',
           'Entre em contato conosco se precisar de mais serviços.',
           'Selecione um arquivo XLSX ou CSV',
           'Selecione um arquivo Excel ou CSV',
           'NOTA FISCAL DE COMPRA:',
           'QTD.',
           'ATIVIDADE RECENTE DE SEGURANÇA E DADOS',
           'REGISTROS DE SERVIÇO RECENTES',
           'ATUALIZAR PAINEL',
           'RECARREGAR DADOS DAS LISTAS',
           'MODELO DE MENSAGEM DE LEMBRETE',
           'LEMBRETES REGISTRADOS',
           'LEMBRETES ENVIADOS',
           'CUSTO DO REPARO',
           'FALHA AO EXPORTAR RELATÓRIO',
           'RELATÓRIO EXPORTADO',
           'REDEFINIR SENHA',
           'REDEFINIR SENHA DO USUÁRIO',
           'REDEFINIR INTERVALO DE GARANTIA A PARTIR DE HOJE',
           'RESTAURAR BANCO DE DADOS',
           'FALHA NA RESTAURAÇÃO',
           'RESTAURAÇÃO CONCLUÍDA',
           'FUNÇÃO',
           'SALVAR ARQUIVO DE EXPORTAÇÃO',
           'SALVAR EXPORTAÇÃO DO RELATÓRIO',
           'PESQUISAR POR NOME, TELEFONE, DISPOSITIVO OU PROBLEMA...',
           'CONFIGURAÇÃO SEGURA DO ADMINISTRADOR',
           'LOGIN SEGURO',
           'Selecione um contato para excluir',
           'Selecione um contato para editar',
           'Selecione uma peça para excluir',
           'Selecione uma peça para editar',
           'Selecione uma peça para atualizar',
           'Selecione um fornecedor para excluir',
           'Selecione um fornecedor para editar',
           'Selecione um fornecedor para atualizar',
           'Selecione um usuário primeiro.',
           'Selecione uma linha de garantia para editar',
           'SELECIONE A MOEDA DO APLICATIVO:',
           'SELECIONE O IDIOMA E A LOCALIDADE DO APLICATIVO:',
           'SELECIONE A MOEDA PADRÃO DO APLICATIVO:',
           'SELECIONAR IDIOMA / LOCALIDADE',
           'SELECIONAR LINHAS PARA EXPORTAR',
           'SELECIONAR GARANTIAS PARA REGISTRAR LEMBRETES',
           'SELECIONAR GARANTIAS PARA ENVIAR LEMBRETES',
           'ENVIAR LEMBRETES',
           'DESCRIÇÃO DO SERVIÇO',
           'SERVIÇO DE COMPUTADOR EM DOMICÍLIO DE BOGOR',
           'GRÁFICO DE STATUS DOS SERVIÇOS',
           'RESUMO DOS SERVIÇOS',
           'RELATÓRIO DE RESUMO DOS SERVIÇOS',
           'Definir uma nova senha para',
           'DEFINIR LOCAL DO BANCO DE DADOS',
           'DEFINIR A DATA DO SERVIÇO COMO HOJE',
           'CONFIGURAÇÕES SALVAS',
           'FALHA NA CONFIGURAÇÃO',
           'MOSTRAR SENHA',
           'MOSTRAR SENHAS',
           'ENTRE PARA CONTINUAR',
           'SESSÃO INICIADA',
           'SESSÃO INICIADA COMO',
           'FUNCIONÁRIO',
           'O ACESSO DO FUNCIONÁRIO É LIMITADO',
           'Funcionários podem adicionar novos registros de serviço e visualizar dados operacionais. Edição, exclusão, '
           'importação, exportação, backup, configurações administrativas, gerenciamento de usuários, alterações de '
           'fornecedores, alterações de peças, relatórios e lembretes exigem um administrador.',
           'STATUS',
           'LOCAL DE ARMAZENAMENTO',
           'ENDEREÇO DO FORNECEDOR',
           'NOME DO FORNECEDOR',
           'O NOME DO FORNECEDOR É OBRIGATÓRIO',
           'FORNECEDORES ATUALIZADOS',
           'SISTEMA',
           'NOME DO TÉCNICO',
           'DESEMPENHO DOS TÉCNICOS',
           'RELATÓRIO DE DESEMPENHO DOS TÉCNICOS',
           'TÉCNICOS',
           'Obrigado,',
           'Esse nome de usuário já existe.',
           'A unidade da pasta de backup não existe.',
           'A unidade do caminho do banco de dados não existe.',
           'A unidade do caminho de exportação não existe.',
           'A unidade da pasta de importação/exportação não existe.',
           'A senha foi redefinida.',
           'As senhas não coincidem.',
           'Esta conta está desativada. Entre em contato com um administrador.',
           'Esta ação é restrita a contas de administrador.',
           'Este software é fornecido no estado em que se encontra',
           'HORA',
           'PARA EXCEL',
           'TOTAL DE GARANTIAS ATIVAS',
           'TOTAL DE CLIENTES',
           'TOTAL DE DISPOSITIVOS',
           'TOTAL DE GARANTIAS VENCIDAS',
           'TOTAL VENCENDO EM BREVE',
           'RECEITA TOTAL',
           'TOTAL SELECIONADO',
           'TOTAL DE REGISTROS DE SERVIÇO',
           'VALOR TOTAL DOS SERVIÇOS',
           'TOTAL DE SERVIÇOS',
           'TIPO:',
           'MODELOS ÚNICOS',
           'CREDENCIAIS DESCONHECIDAS OU INVÁLIDAS',
           'Tabela desconhecida para exportação.',
           'DESBLOQUEAR',
           'FORMATO NÃO COMPATÍVEL',
           'USAR D:\\BACKUP (PADRÃO)',
           'USUÁRIO',
           'GERENCIAMENTO DE USUÁRIOS',
           'USUÁRIO NÃO CRIADO',
           'Usuário não encontrado.',
           'USUÁRIO SOLICITOU SAÍDA',
           'Status do usuário atualizado.',
           'NOME DE USUÁRIO',
           'NOME DE USUÁRIO JÁ EXISTE',
           'O nome de usuário deve ter de 3 a 32 caracteres e usar letras, números, ponto, hífen ou sublinhado.',
           'USUÁRIO_CRIADO',
           'USUÁRIO_DESATIVADO',
           'USUÁRIO_ATIVADO',
           'USUÁRIO_DESBLOQUEADO',
           'AGUARDANDO PEÇAS',
           'AGUARDANDO PEÇAS',
           'ATENÇÃO: A RESTAURAÇÃO SUBSTITUIRÁ TODOS OS DADOS ATUAIS PELOS DADOS DO BACKUP.\n'
           'ESTA AÇÃO NÃO PODE SER DESFEITA.\n'
           '\n'
           'DESEJA CONTINUAR?',
           'PERÍODOS DE GARANTIA',
           'STATUS DA GARANTIA:',
           'RESUMO DAS GARANTIAS',
           'RELATÓRIO DE RESUMO DAS GARANTIAS',
           'SENHA FRACA',
           'O que você deseja fazer com o banco de dados?',
           'JANELA FECHADA',
           'Você não pode desativar sua conta atual.',
           'Sua senha foi alterada com sucesso.',
           '⚠️ ATENÇÃO!',
           '✏️ EDITAR',
           '✏️ EDITAR CONTATO',
           '❌ EXCLUIR',
           '➕ ADICIONAR',
           '➕ ADICIONAR PEÇA',
           '📄 VISUALIZAÇÃO EM TEXTO',
           '📋 VISUALIZAÇÃO EM TABELA',
           '📝 REGISTRAR LEMBRETE',
           '📝 REGISTRAR LEMBRETES',
           '📤 EXPORTAR SELECIONADOS',
           '🔃 ATUALIZAR',
           '🔄 ATUALIZAR',
           '🗑️ LIMPAR'],
 'it_IT': ['Un nuovo ricambio deve essere creato da un amministratore.',
           'Un nuovo fornitore deve essere creato da un amministratore.',
           'Esiste già un ricambio con questo nome. Selezionalo e usa Salva per aggiornarlo.',
           'Esiste già un fornitore con questo nome. Selezionalo e usa Salva per aggiornarlo.',
           'ACCESSO_NEGATO',
           'ACCOUNT',
           'ACCOUNT DISABILITATO',
           'ACCOUNT BLOCCATO',
           'AZIONE BLOCCATA',
           'GARANZIE ATTIVE',
           "AGGIUNGI UTENTE DELL'APPLICAZIONE",
           'AGGIUNGI UTENTE',
           'NOTE AGGIUNTIVE',
           'INDIRIZZO COMPLETO',
           'ADMIN',
           'AMMINISTRATORE',
           "AZIONE DELL'AMMINISTRATORE BLOCCATA",
           'SOLO AMMINISTRATORE',
           'AMMINISTRATORE NECESSARIO',
           'AMMINISTRATORE NECESSARIO.',
           'Un altro ricambio utilizza già questo nome.',
           'Un altro fornitore utilizza già questo nome.',
           'APPLICAZIONE CHIUSA',
           'USCITA_APPLICAZIONE',
           'APERTURA_APPLICAZIONE',
           'È necessario almeno un amministratore attivo.',
           'MEDIA PER INTERVENTO',
           'COSTO MEDIO DI RIPARAZIONE',
           'RICAVO MEDIO',
           'RICAVO MEDIO/INTERVENTO',
           'BACKUP DEL DATABASE',
           'BACKUP NON RIUSCITO',
           'La cartella di backup non può essere vuota',
           'CARTELLA DI BACKUP:',
           'BACKUP COMPLETATO',
           'Il Servizio Computer a Domicilio di Bogor ti ricorda che la garanzia di riparazione del dispositivo scadrà '
           'il:',
           'MARCA:',
           'MARCHE',
           "Compila HELP.HHP e copia HELP.CHM nella cartella dell'applicazione.",
           "Calcolato automaticamente: data dell'intervento + garanzia",
           'Impossibile determinare il fornitore selezionato',
           'CAMBIA PASSWORD',
           'CAMBIA PASSWORD PER',
           'COMPLETATO',
           'GESTORE ASSISTENZA COMPUTER',
           'CONFERMA ELIMINAZIONE',
           'CONFERMA PASSWORD',
           'CONFERMA RIPRISTINO',
           'NOME DEL CONTATTO',
           'CONTATTO:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'COPIATO NEGLI APPUNTI',
           'COPIA INDIRIZZO',
           'COPIA GOOGLE MAP',
           'COPIA FATTURA',
           'COPIA TELEFONO',
           'COPYRIGHT © ECOMTECH 2026',
           'CONTEGGIO',
           'CREA AMMINISTRATORE',
           'CREA IL PRIMO ACCOUNT AMMINISTRATORE',
           'CREATO',
           'CREATO DA:',
           'DATA DI CREAZIONE',
           'FILE CSV',
           'PASSWORD ATTUALE',
           'La password attuale non è corretta.',
           'GUIDA DELLA SCHEDA CORRENTE',
           'NOME DEL CLIENTE',
           'CLIENTI SELEZIONATI PER IL PROMEMORIA',
           'CLIENTI SELEZIONATI PER IL PROMEMORIA:',
           'Gestore Assistenza Computer 4.4 Sicuro',
           'AGGIORNAMENTO DASHBOARD NON RIUSCITO',
           'RIEPILOGO DASHBOARD',
           'I dati sono stati esportati in:',
           'POSIZIONE DEL DATABASE OBBLIGATORIA',
           'Il percorso del database non può essere vuoto',
           'PERCORSO DEL DATABASE:',
           'Percorsi del database aggiornati correttamente!',
           'DATABASE RIPRISTINATO DAL BACKUP',
           'Database ripristinato correttamente dal backup!',
           'DATABASE_RIPRISTINATO',
           'DATA',
           'Gentile [CUSTOMER],',
           'DETTAGLI',
           'STATISTICHE MARCHE DEI DISPOSITIVI',
           'MODELLO DEL DISPOSITIVO',
           'TIPI DI DISPOSITIVO',
           'NOME VISUALIZZATO',
           'DONAZIONE:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           "L'unità D: non è disponibile. Scegli un'altra cartella.",
           "L'unità D: non è stata trovata.\nCrea o seleziona una cartella per il database.",
           'UNITÀ NON TROVATA',
           'EDITOR DATI DI PROVA APERTO',
           'RICAMBIO DUPLICATO',
           'FORNITORE DUPLICATO',
           'MODALITÀ MODIFICA',
           'MODALITÀ MODIFICA: AGGIORNA I CAMPI E FAI CLIC SU SALVA',
           'MODALITÀ MODIFICA: AGGIORNA I CAMPI E FAI CLIC SU AGGIORNA',
           'ABILITA / DISABILITA',
           'DATA DI FINE:',
           'Inserisci nome utente e password.',
           'FILE EXCEL',
           'IN SCADENZA TRA 30 GIORNI',
           'IN SCADENZA A BREVE',
           'IN SCADENZA A BREVE (< 7 GIORNI)',
           'IN SCADENZA A BREVE (≤ 30 GIORNI)',
           'IN SCADENZA A BREVE (≤ 7 GIORNI)',
           'IN SCADENZA A BREVE (≤30 giorni)',
           'DATA DI SCADENZA DA',
           'DATA DI SCADENZA DA:',
           'ESPORTAZIONE COMPLETATA',
           'ERRORE DI ESPORTAZIONE',
           'ESPORTAZIONE NON RIUSCITA',
           'ESPORTAZIONE COMPLETATA',
           'ESPORTATO IN:',
           'Esportazione non riuscita',
           'REPORT FINANZIARIO',
           "PRIMA SESSIONE DELL'AMMINISTRATORE",
           'CONFIGURAZIONE INIZIALE DEL DATABASE',
           'CONFIGURAZIONE INIZIALE:\nScegli una cartella in cui salvare il nuovo database.',
           'MODULO AZZERATO',
           'La data iniziale non può essere successiva alla data finale',
           'INDIRIZZO COMPLETO',
           'GOOGLE MAP',
           'LINK GOOGLE MAP COPIATO',
           'LINK GOOGLE MAP VUOTO',
           'LINK GOOGLE MAP VUOTO (APPUNTI CANCELLATI)',
           'GOOGLE MAP:',
           'CONTENUTI DELLA GUIDA',
           'ID (ROWID)',
           'IMPORTAZIONE COMPLETATA',
           'IMPORTAZIONE NON RIUSCITA',
           'La cartella di importazione/esportazione non può essere vuota',
           'CARTELLA DI IMPORTAZIONE/ESPORTAZIONE:',
           'IN CORSO',
           'IN RIPARAZIONE',
           'CREDENZIALI NON VALIDE',
           'INTERVALLO DI DATE NON VALIDO',
           'UNITÀ NON VALIDA',
           'PERCORSO NON VALIDO',
           'Ruolo utente non valido.',
           'NOME UTENTE NON VALIDO',
           'Nome utente o password non validi.',
           'Intervallo date garanzia non valido: la data iniziale è successiva alla data finale',
           'NUMERO FATTURA',
           'PROBLEMI',
           'ULTIMO ACCESSO',
           'I file .XLS obsoleti non sono supportati. Salva prima il file come .XLSX o CSV.',
           'BLOCCATO FINO A',
           'REGISTRA PROMEMORIA',
           'REGISTRA PROMEMORIA',
           'REGISTRA PROMEMORIA GARANZIA',
           'ACCEDI',
           'Accesso riuscito.',
           'ACCESSO_BLOCCATO',
           'ACCESSO_NON_RIUSCITO',
           'ACCESSO_RIUSCITO',
           'ESCI',
           'RICAMBI CON SCORTE BASSE',
           'FINESTRA PRINCIPALE APERTA',
           'Minimo 4 caratteri. È consentito qualsiasi formato con lettere, numeri, spazi o simboli.',
           'MODELLO DEL DISPOSITIVO',
           'MODELLI',
           'CREATORE DATI DI PROVA MULTILINGUE APERTO',
           'LE MIE ATTIVITÀ RECENTI',
           'N/D',
           'NOME DEL CLIENTE',
           'IL NOME È OBBLIGATORIO',
           'NOME:',
           'NUOVA PASSWORD',
           'NESSUN DATO',
           'NESSUN DATABASE SELEZIONATO',
           'Non viene usata alcuna password predefinita. Le password vengono archiviate solo come hash scrypt univoci '
           'con salt. Conserva al sicuro questa password di amministratore.',
           "Non è stato selezionato alcun database esistente. L'applicazione continuerà creando un nuovo database.",
           'Non è stato inviato alcun messaggio SMS o WhatsApp.',
           'NESSUN DATO SULLO STATO',
           'NESSUNA GARANZIA',
           'Non viene fornita alcuna garanzia di alcun tipo',
           'NUMERO DI TELEFONO',
           'NON TROVATO',
           'Apri un database esistente o creane uno nuovo per questa applicazione.',
           "APRI I CONTENUTI DELLA GUIDA DELL'APPLICAZIONE",
           "APRI L'EDITOR DATI DI PROVA",
           'APRI GOOGLE MAP',
           'APRI LA GUIDA DELLA SCHEDA ATTIVA',
           'NOME DEL RICAMBIO',
           'IL NOME DEL RICAMBIO È OBBLIGATORIO',
           'TIPI DI RICAMBIO',
           'RICAMBI AGGIORNATI',
           'PASSWORD',
           'PASSWORD MODIFICATA',
           'Password modificata.',
           'La password è troppo lunga.',
           'LE PASSWORD NON CORRISPONDONO',
           'La password deve contenere almeno 4 caratteri.',
           'PASSWORD NON MODIFICATA',
           'PASSWORD REIMPOSTATA',
           'PASSWORD_MODIFICATA',
           'MODIFICA_PASSWORD_NON_RIUSCITA',
           'PASSWORD_REIMPOSTATA',
           'NUMERO DI TELEFONO',
           'TELEFONO:',
           "Posiziona DUMMY_CREATOR.EXE o DUMMY_CREATOR.PY accanto all'applicazione.",
           'Contattaci se hai bisogno di ulteriore assistenza.',
           'Seleziona un file XLSX o CSV',
           'Seleziona un file Excel o CSV',
           'FATTURA DI ACQUISTO:',
           'QTÀ',
           'ATTIVITÀ RECENTI DI SICUREZZA E DATI',
           'REGISTRI DI ASSISTENZA RECENTI',
           'AGGIORNA DASHBOARD',
           'RICARICA DATI DEGLI ELENCHI',
           'MODELLO DEL MESSAGGIO DI PROMEMORIA',
           'PROMEMORIA REGISTRATI',
           'PROMEMORIA INVIATI',
           'COSTO DI RIPARAZIONE',
           'ESPORTAZIONE REPORT NON RIUSCITA',
           'REPORT ESPORTATO',
           'REIMPOSTA PASSWORD',
           'REIMPOSTA PASSWORD UTENTE',
           'REIMPOSTA INTERVALLO GARANZIA DA OGGI',
           'RIPRISTINA DATABASE',
           'RIPRISTINO NON RIUSCITO',
           'RIPRISTINO COMPLETATO',
           'RUOLO',
           'SALVA FILE DI ESPORTAZIONE',
           'SALVA ESPORTAZIONE REPORT',
           'CERCA PER NOME, TELEFONO, DISPOSITIVO O PROBLEMA...',
           "CONFIGURAZIONE SICURA DELL'AMMINISTRATORE",
           'ACCESSO SICURO',
           'Seleziona un contatto da eliminare',
           'Seleziona un contatto da modificare',
           'Seleziona un ricambio da eliminare',
           'Seleziona un ricambio da modificare',
           'Seleziona un ricambio da aggiornare',
           'Seleziona un fornitore da eliminare',
           'Seleziona un fornitore da modificare',
           'Seleziona un fornitore da aggiornare',
           'Seleziona prima un utente.',
           'Seleziona una riga di garanzia da modificare',
           "SELEZIONA LA VALUTA DELL'APPLICAZIONE:",
           "SELEZIONA LINGUA E IMPOSTAZIONI LOCALI DELL'APPLICAZIONE:",
           "SELEZIONA LA VALUTA PREDEFINITA DELL'APPLICAZIONE:",
           'SELEZIONA LINGUA / IMPOSTAZIONI LOCALI',
           'SELEZIONA RIGHE DA ESPORTARE',
           'SELEZIONA GARANZIE PER REGISTRARE PROMEMORIA',
           'SELEZIONA GARANZIE PER INVIARE PROMEMORIA',
           'INVIA PROMEMORIA',
           "DESCRIZIONE DELL'INTERVENTO",
           'SERVIZIO COMPUTER A DOMICILIO DI BOGOR',
           'GRAFICO DELLO STATO DEGLI INTERVENTI',
           'RIEPILOGO DEGLI INTERVENTI',
           'REPORT RIEPILOGATIVO DEGLI INTERVENTI',
           'Imposta una nuova password per',
           'IMPOSTA POSIZIONE DEL DATABASE',
           "IMPOSTA LA DATA DELL'INTERVENTO A OGGI",
           'IMPOSTAZIONI SALVATE',
           'CONFIGURAZIONE NON RIUSCITA',
           'MOSTRA PASSWORD',
           'MOSTRA PASSWORD',
           'ACCEDI PER CONTINUARE',
           'ACCESSO ESEGUITO',
           'ACCESSO ESEGUITO COME',
           'OPERATORE',
           "L'ACCESSO DELL'OPERATORE È LIMITATO",
           'Gli operatori possono aggiungere nuovi registri di assistenza e visualizzare i dati operativi. Modifiche, '
           'eliminazioni, importazione, esportazione, backup, impostazioni amministrative, gestione utenti, modifiche '
           'dei fornitori, modifiche dei ricambi, report e promemoria richiedono un amministratore.',
           'STATI',
           'POSIZIONE DI ARCHIVIAZIONE',
           'INDIRIZZO DEL FORNITORE',
           'NOME DEL FORNITORE',
           'IL NOME DEL FORNITORE È OBBLIGATORIO',
           'FORNITORI AGGIORNATI',
           'SISTEMA',
           'NOME DEL TECNICO',
           'PRESTAZIONI DEI TECNICI',
           'REPORT SULLE PRESTAZIONI DEI TECNICI',
           'TECNICI',
           'Grazie,',
           'Questo nome utente esiste già.',
           "L'unità della cartella di backup non esiste.",
           "L'unità del percorso del database non esiste.",
           "L'unità del percorso di esportazione non esiste.",
           "L'unità della cartella di importazione/esportazione non esiste.",
           'La password è stata reimpostata.',
           'Le password non corrispondono.',
           'Questo account è disabilitato. Contatta un amministratore.',
           'Questa azione è riservata agli account amministratore.',
           "Questo software viene fornito così com'è",
           'ORA',
           'IN EXCEL',
           'TOTALE GARANZIE ATTIVE',
           'TOTALE CLIENTI',
           'TOTALE DISPOSITIVI',
           'TOTALE SCADUTI',
           'TOTALE IN SCADENZA A BREVE',
           'RICAVO TOTALE',
           'TOTALE SELEZIONATO',
           'TOTALE REGISTRI DI ASSISTENZA',
           'VALORE TOTALE DEGLI INTERVENTI',
           'TOTALE INTERVENTI',
           'TIPO:',
           'MODELLI UNIVOCI',
           'CREDENZIALI SCONOSCIUTE O NON VALIDE',
           'Tabella sconosciuta da esportare.',
           'SBLOCCA',
           'FORMATO NON SUPPORTATO',
           'USA D:\\BACKUP (PREDEFINITO)',
           'UTENTE',
           'GESTIONE UTENTI',
           'UTENTE NON CREATO',
           'Utente non trovato.',
           "L'UTENTE HA RICHIESTO LA DISCONNESSIONE",
           'Stato utente aggiornato.',
           'NOME UTENTE',
           'NOME UTENTE GIÀ ESISTENTE',
           'Il nome utente deve contenere da 3 a 32 caratteri tra lettere, numeri, punto, trattino o trattino basso.',
           'UTENTE_CREATO',
           'UTENTE_DISABILITATO',
           'UTENTE_ABILITATO',
           'UTENTE_SBLOCCATO',
           'IN ATTESA DI RICAMBI',
           'IN ATTESA DI RICAMBI',
           'ATTENZIONE: IL RIPRISTINO SOSTITUIRÀ TUTTI I DATI ATTUALI CON I DATI DEL BACKUP.\n'
           'QUESTA AZIONE NON PUÒ ESSERE ANNULLATA.\n'
           '\n'
           'VUOI CONTINUARE?',
           'PERIODI DI GARANZIA',
           'STATO DELLA GARANZIA:',
           'RIEPILOGO GARANZIE',
           'REPORT RIEPILOGATIVO GARANZIE',
           'PASSWORD DEBOLE',
           'Cosa vuoi fare con il database?',
           'FINESTRA CHIUSA',
           'Non puoi disabilitare il tuo account corrente.',
           'La password è stata modificata correttamente.',
           '⚠️ ATTENZIONE!',
           '✏️ MODIFICA',
           '✏️ MODIFICA CONTATTO',
           '❌ ELIMINA',
           '➕ AGGIUNGI',
           '➕ AGGIUNGI RICAMBIO',
           '📄 VISUALIZZAZIONE TESTO',
           '📋 VISUALIZZAZIONE TABELLA',
           '📝 REGISTRA PROMEMORIA',
           '📝 REGISTRA PROMEMORIA',
           '📤 ESPORTA SELEZIONATI',
           '🔃 AGGIORNA',
           '🔄 AGGIORNA',
           '🗑️ CANCELLA'],
 'nl_NL': ['Een nieuw onderdeel moet door een beheerder worden aangemaakt.',
           'Een nieuwe leverancier moet door een beheerder worden aangemaakt.',
           'Er bestaat al een onderdeel met deze naam. Selecteer het en gebruik Opslaan om het bij te werken.',
           'Er bestaat al een leverancier met deze naam. Selecteer deze en gebruik Opslaan om deze bij te werken.',
           'TOEGANG_GEWEIGERD',
           'ACCOUNT',
           'ACCOUNT UITGESCHAKELD',
           'ACCOUNT GEBLOKKEERD',
           'ACTIE GEBLOKKEERD',
           'ACTIEVE GARANTIES',
           'APPLICATIEGEBRUIKER TOEVOEGEN',
           'GEBRUIKER TOEVOEGEN',
           'AANVULLENDE OPMERKINGEN',
           'VOLLEDIG ADRES',
           'BEHEER',
           'BEHEERDER',
           'BEHEERDERSACTIE GEBLOKKEERD',
           'ALLEEN BEHEERDER',
           'BEHEERDER VEREIST',
           'BEHEERDER VEREIST.',
           'Een ander onderdeel gebruikt deze naam al.',
           'Een andere leverancier gebruikt deze naam al.',
           'APPLICATIE GESLOTEN',
           'APPLICATIE_AFGESLOTEN',
           'APPLICATIE_GEOPEND',
           'Er is ten minste één actieve beheerder vereist.',
           'GEMIDDELD PER SERVICE',
           'GEMIDDELDE REPARATIEKOSTEN',
           'GEMIDDELDE OMZET',
           'GEMIDDELDE OMZET/SERVICE',
           'DATABASE BACK-UP MAKEN',
           'BACK-UP MISLUKT',
           'De back-upmap mag niet leeg zijn',
           'BACK-UPMAP:',
           'BACK-UP GESLAAGD',
           'Bogor Computer Service aan Huis herinnert u eraan dat de reparatiegarantie van uw apparaat verloopt op:',
           'MERK:',
           'MERKEN',
           'Bouw HELP.HHP en kopieer HELP.CHM naar de toepassingsmap.',
           'Automatisch berekend: servicedatum + garantie',
           'Kan de geselecteerde leverancier niet bepalen',
           'WACHTWOORD WIJZIGEN',
           'WACHTWOORD WIJZIGEN VOOR',
           'VOLTOOID',
           'COMPUTERSERVICEBEHEER',
           'VERWIJDEREN BEVESTIGEN',
           'WACHTWOORD BEVESTIGEN',
           'HERSTEL BEVESTIGEN',
           'NAAM CONTACTPERSOON',
           'CONTACT:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'GEKOPIEERD NAAR KLEMBORD',
           'ADRES KOPIËREN',
           'GOOGLE MAP KOPIËREN',
           'FACTUUR KOPIËREN',
           'TELEFOON KOPIËREN',
           'AUTEURSRECHT © ECOMTECH 2026',
           'AANTAL',
           'BEHEERDER AANMAKEN',
           'HET EERSTE BEHEERDERSACCOUNT AANMAKEN',
           'AANGEMAAKT',
           'AANGEMAAKT DOOR:',
           'AANMAAKDATUM',
           'CSV-BESTANDEN',
           'HUIDIG WACHTWOORD',
           'Het huidige wachtwoord is onjuist.',
           'HELP VOOR HUIDIG TABBLAD',
           'KLANTNAAM',
           'KLANTEN GESELECTEERD VOOR HERINNERING',
           'KLANTEN GESELECTEERD VOOR HERINNERING:',
           'Veilig Computerservicebeheer 4.4',
           'DASHBOARD VERNIEUWEN MISLUKT',
           'DASHBOARDOVERZICHT',
           'Gegevens zijn geëxporteerd naar:',
           'DATABASELOCATIE VEREIST',
           'Het databasepad mag niet leeg zijn',
           'DATABASEPAD:',
           'Databasepaden zijn bijgewerkt!',
           'DATABASE HERSTELD VANUIT BACK-UP',
           'Database is hersteld vanuit de back-up!',
           'DATABASE_HERSTELD',
           'DATUM',
           'Geachte [CUSTOMER],',
           'DETAILS',
           'STATISTIEKEN APPARAATMERKEN',
           'APPARAATMODEL',
           'APPARAATTYPEN',
           'WEERGAVENAAM',
           'DONATIE:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Station D: is niet beschikbaar. Kies een andere map.',
           'Station D: is niet gevonden.\nMaak of selecteer een map voor de database.',
           'STATION NIET GEVONDEN',
           'TESTGEGEVENSEDITOR GEOPEND',
           'DUBBEL ONDERDEEL',
           'DUBBELE LEVERANCIER',
           'BEWERKINGSMODUS',
           'BEWERKINGSMODUS: WERK VELDEN BIJ EN KLIK OP OPSLAAN',
           'BEWERKINGSMODUS: WERK VELDEN BIJ EN KLIK OP BIJWERKEN',
           'INSCHAKELEN / UITSCHAKELEN',
           'EINDDATUM:',
           'Voer zowel gebruikersnaam als wachtwoord in.',
           'EXCEL-BESTANDEN',
           'VERLOOPT OVER 30 DAGEN',
           'VERLOOPT BINNENKORT',
           'VERLOOPT BINNENKORT (< 7 DAGEN)',
           'VERLOOPT BINNENKORT (≤ 30 DAGEN)',
           'VERLOOPT BINNENKORT (≤ 7 DAGEN)',
           'VERLOOPT BINNENKORT (≤30 dagen)',
           'VERVALDATUM VANAF',
           'VERVALDATUM VANAF:',
           'EXPORT VOLTOOID',
           'EXPORTFOUT',
           'EXPORT MISLUKT',
           'EXPORT GESLAAGD',
           'GEËXPORTEERD NAAR:',
           'Exporteren mislukt',
           'FINANCIEEL RAPPORT',
           'EERSTE BEHEERDERSESSIE',
           'EERSTE DATABASECONFIGURATIE',
           'EERSTE CONFIGURATIE:\nKies een map om de nieuwe database op te slaan.',
           'FORMULIER GEWIST',
           'Van-datum mag niet na tot-datum liggen',
           'VOLLEDIG ADRES',
           'GOOGLE MAP',
           'GOOGLE MAP-LINK GEKOPIEERD',
           'GOOGLE MAP-LINK IS LEEG',
           'GOOGLE MAP-LINK IS LEEG (KLEMBORD GEWIST)',
           'GOOGLE MAP:',
           'HELPINHOUD',
           'ID (ROWID)',
           'IMPORT VOLTOOID',
           'IMPORT MISLUKT',
           'De import-/exportmap mag niet leeg zijn',
           'IMPORT-/EXPORTMAP:',
           'BEZIG',
           'IN REPARATIE',
           'ONGELDIGE AANMELDINGSGEGEVENS',
           'ONGELDIG DATUMBEREIK',
           'ONGELDIG STATION',
           'ONGELDIG PAD',
           'Ongeldige gebruikersrol.',
           'ONGELDIGE GEBRUIKERSNAAM',
           'Ongeldige gebruikersnaam of ongeldig wachtwoord.',
           'Ongeldig garantiedatumbereik: van-datum ligt na tot-datum',
           'FACTUURNUMMER',
           'PROBLEMEN',
           'LAATSTE AANMELDING',
           'Oude .XLS-bestanden worden niet ondersteund. Sla het bestand eerst op als .XLSX of CSV.',
           'GEBLOKKEERD TOT',
           'HERINNERING REGISTREREN',
           'HERINNERINGEN REGISTREREN',
           'GARANTIEHERINNERING REGISTREREN',
           'AANMELDEN',
           'Aanmelding geslaagd.',
           'AANMELDING_GEBLOKKEERD',
           'AANMELDING_MISLUKT',
           'AANMELDING_GESLAAGD',
           'AFMELDEN',
           'ONDERDELEN MET LAGE VOORRAAD',
           'HOOFDVENSTER GEOPEND',
           'Minimaal 4 tekens. Elke combinatie van letters, cijfers, spaties of symbolen is toegestaan.',
           'APPARAATMODEL',
           'MODELLEN',
           'MEERTALIGE TESTGEGEVENSGENERATOR GEOPEND',
           'MIJN RECENTE ACTIVITEIT',
           'N.V.T.',
           'KLANTNAAM',
           'NAAM IS VERPLICHT',
           'NAAM:',
           'NIEUW WACHTWOORD',
           'GEEN GEGEVENS',
           'GEEN DATABASE GESELECTEERD',
           'Er wordt geen standaardwachtwoord gebruikt. Wachtwoorden worden alleen opgeslagen als unieke, gezouten '
           'scrypt-hashes. Bewaar dit beheerderswachtwoord veilig.',
           'Er is geen bestaande database geselecteerd. De toepassing gaat verder met het maken van een nieuwe '
           'database.',
           'Er is geen sms- of WhatsApp-bericht verzonden.',
           'GEEN STATUSGEGEVENS',
           'GEEN GARANTIE',
           'Er wordt geen enkele garantie verleend',
           'TELEFOONNUMMER',
           'NIET GEVONDEN',
           'Open een bestaande database of maak een nieuwe database voor deze toepassing.',
           'HELPINHOUD VAN TOEPASSING OPENEN',
           'TESTGEGEVENSEDITOR OPENEN',
           'GOOGLE MAP OPENEN',
           'HELP VOOR ACTIEF TABBLAD OPENEN',
           'ONDERDEELNAAM',
           'ONDERDEELNAAM IS VERPLICHT',
           'ONDERDEELTYPEN',
           'ONDERDELEN VERNIEUWD',
           'WACHTWOORD',
           'WACHTWOORD GEWIJZIGD',
           'Wachtwoord gewijzigd.',
           'Het wachtwoord is te lang.',
           'WACHTWOORDEN KOMEN NIET OVEREEN',
           'Het wachtwoord moet minimaal 4 tekens bevatten.',
           'WACHTWOORD NIET GEWIJZIGD',
           'WACHTWOORD OPNIEUW INGESTELD',
           'WACHTWOORD_GEWIJZIGD',
           'WACHTWOORDWIJZIGING_MISLUKT',
           'WACHTWOORD_OPNIEUW_INGESTELD',
           'TELEFOONNUMMER',
           'TELEFOON:',
           'Plaats DUMMY_CREATOR.EXE of DUMMY_CREATOR.PY naast de toepassing.',
           'Neem contact met ons op als u verdere service nodig hebt.',
           'Selecteer een XLSX- of CSV-bestand',
           'Selecteer een Excel- of CSV-bestand',
           'INKOOPFACTUUR:',
           'AANTAL',
           'RECENTE BEVEILIGINGS- EN GEGEVENSACTIVITEIT',
           'RECENTE SERVICEREGISTRATIES',
           'DASHBOARD VERNIEUWEN',
           'KEUZELIJSTGEGEVENS OPNIEUW LADEN',
           'SJABLOON VOOR HERINNERINGSBERICHT',
           'HERINNERINGEN GEREGISTREERD',
           'HERINNERINGEN VERZONDEN',
           'REPARATIEKOSTEN',
           'RAPPORTEXPORT MISLUKT',
           'RAPPORT GEËXPORTEERD',
           'WACHTWOORD OPNIEUW INSTELLEN',
           'GEBRUIKERSWACHTWOORD OPNIEUW INSTELLEN',
           'GARANTIEBEREIK VANAF VANDAAG OPNIEUW INSTELLEN',
           'DATABASE HERSTELLEN',
           'HERSTEL MISLUKT',
           'HERSTEL GESLAAGD',
           'ROL',
           'EXPORTBESTAND OPSLAAN',
           'RAPPORTEXPORT OPSLAAN',
           'ZOEKEN OP NAAM, TELEFOON, APPARAAT OF PROBLEEM...',
           'VEILIGE BEHEERDERSCONFIGURATIE',
           'VEILIG AANMELDEN',
           'Selecteer een contactpersoon om te verwijderen',
           'Selecteer een contactpersoon om te bewerken',
           'Selecteer een onderdeel om te verwijderen',
           'Selecteer een onderdeel om te bewerken',
           'Selecteer een onderdeel om bij te werken',
           'Selecteer een leverancier om te verwijderen',
           'Selecteer een leverancier om te bewerken',
           'Selecteer een leverancier om bij te werken',
           'Selecteer eerst een gebruiker.',
           'Selecteer een garantieregel om te bewerken',
           'SELECTEER APPLICATIEVALUTA:',
           'SELECTEER APPLICATIETAAL EN LANDINSTELLING:',
           'SELECTEER STANDAARDVALUTA VOOR DE TOEPASSING:',
           'TAAL / LANDINSTELLING SELECTEREN',
           'RIJEN SELECTEREN OM TE EXPORTEREN',
           'GARANTIES SELECTEREN OM HERINNERINGEN TE REGISTREREN',
           'GARANTIES SELECTEREN OM HERINNERINGEN TE VERZENDEN',
           'HERINNERINGEN VERZENDEN',
           'SERVICEBESCHRIJVING',
           'BOGOR COMPUTERSERVICE AAN HUIS',
           'GRAFIEK SERVICESTATUS',
           'SERVICEOVERZICHT',
           'RAPPORT SERVICEOVERZICHT',
           'Een nieuw wachtwoord instellen voor',
           'DATABASELOCATIE INSTELLEN',
           'SERVICEDATUM INSTELLEN OP VANDAAG',
           'INSTELLINGEN OPGESLAGEN',
           'CONFIGURATIE MISLUKT',
           'WACHTWOORD TONEN',
           'WACHTWOORDEN TONEN',
           'MELD U AAN OM DOOR TE GAAN',
           'AANGEMELD',
           'AANGEMELD ALS',
           'MEDEWERKER',
           'TOEGANG VOOR MEDEWERKERS IS BEPERKT',
           'Medewerkers kunnen nieuwe serviceregistraties toevoegen en operationele gegevens bekijken. Bewerken, '
           'verwijderen, importeren, exporteren, back-ups, beheerdersinstellingen, gebruikersbeheer, wijzigingen aan '
           'leveranciers, wijzigingen aan onderdelen, rapporten en herinneringen vereisen een beheerder.',
           'STATUSSEN',
           'OPSLAGLOCATIE',
           'ADRES LEVERANCIER',
           'NAAM LEVERANCIER',
           'NAAM LEVERANCIER IS VERPLICHT',
           'LEVERANCIERS VERNIEUWD',
           'SYSTEEM',
           'NAAM TECHNICUS',
           'PRESTATIES TECHNICI',
           'RAPPORT PRESTATIES TECHNICI',
           'TECHNICI',
           'Dank u,',
           'Die gebruikersnaam bestaat al.',
           'Het station voor de back-upmap bestaat niet.',
           'Het station voor het databasepad bestaat niet.',
           'Het station voor het exportpad bestaat niet.',
           'Het station voor de import-/exportmap bestaat niet.',
           'Het wachtwoord is opnieuw ingesteld.',
           'De wachtwoorden komen niet overeen.',
           'Dit account is uitgeschakeld. Neem contact op met een beheerder.',
           'Deze actie is beperkt tot beheerdersaccounts.',
           'Deze software wordt geleverd zoals deze is',
           'TIJD',
           'NAAR EXCEL',
           'TOTAAL ACTIEVE GARANTIES',
           'TOTAAL KLANTEN',
           'TOTAAL APPARATEN',
           'TOTAAL VERLOPEN',
           'TOTAAL VERLOOPT BINNENKORT',
           'TOTALE OMZET',
           'TOTAAL GESELECTEERD',
           'TOTAAL SERVICEREGISTRATIES',
           'TOTALE SERVICEWAARDE',
           'TOTAAL SERVICES',
           'TYPE:',
           'UNIEKE MODELLEN',
           'ONBEKENDE OF ONGELDIGE AANMELDINGSGEGEVENS',
           'Onbekende tabel om te exporteren.',
           'DEBLOKKEREN',
           'NIET-ONDERSTEUND FORMAAT',
           'D:\\BACKUP GEBRUIKEN (STANDAARD)',
           'GEBRUIKER',
           'GEBRUIKERSBEHEER',
           'GEBRUIKER NIET AANGEMAAKT',
           'Gebruiker niet gevonden.',
           'GEBRUIKER HEEFT AFMELDEN AANGEVRAAGD',
           'Gebruikersstatus bijgewerkt.',
           'GEBRUIKERSNAAM',
           'GEBRUIKERSNAAM BESTAAT AL',
           'De gebruikersnaam moet 3-32 tekens bevatten en mag letters, cijfers, punten, streepjes of '
           'onderstrepingstekens gebruiken.',
           'GEBRUIKER_AANGEMAAKT',
           'GEBRUIKER_UITGESCHAKELD',
           'GEBRUIKER_INGESCHAKELD',
           'GEBRUIKER_GEDEBLOKKEERD',
           'WACHTEN OP ONDERDELEN',
           'WACHT OP ONDERDELEN',
           'WAARSCHUWING: HERSTELLEN VERVANGT ALLE HUIDIGE GEGEVENS DOOR DE BACK-UPGEGEVENS.\n'
           'DEZE ACTIE KAN NIET ONGEDAAN WORDEN GEMAAKT.\n'
           '\n'
           'WILT U DOORGAAN?',
           'GARANTIEPERIODEN',
           'GARANTIESTATUS:',
           'GARANTIEOVERZICHT',
           'RAPPORT GARANTIEOVERZICHT',
           'ZWAK WACHTWOORD',
           'Wat wilt u met de database doen?',
           'VENSTER GESLOTEN',
           'U kunt uw huidige account niet uitschakelen.',
           'Uw wachtwoord is gewijzigd.',
           '⚠️ WAARSCHUWING!',
           '✏️ BEWERKEN',
           '✏️ CONTACT BEWERKEN',
           '❌ VERWIJDEREN',
           '➕ TOEVOEGEN',
           '➕ ONDERDEEL TOEVOEGEN',
           '📄 TEKSTWEERGAVE',
           '📋 TABELWEERGAVE',
           '📝 HERINNERING REGISTREREN',
           '📝 HERINNERINGEN REGISTREREN',
           '📤 SELECTIE EXPORTEREN',
           '🔃 VERNIEUWEN',
           '🔄 VERNIEUWEN',
           '🗑️ WISSEN'],
 'zh_CN': ['新零件必须由管理员创建。',
           '新供应商必须由管理员创建。',
           '已存在同名零件。请选择该零件并使用“保存”进行更新。',
           '已存在同名供应商。请选择该供应商并使用“保存”进行更新。',
           '访问被拒绝',
           '账户',
           '账户已禁用',
           '账户已锁定',
           '操作已阻止',
           '有效保修',
           '添加应用程序用户',
           '添加用户',
           '附加备注',
           '完整地址',
           '管理员',
           '管理员',
           '管理员操作已阻止',
           '仅限管理员',
           '需要管理员权限',
           '需要管理员权限。',
           '已有其他零件使用此名称。',
           '已有其他供应商使用此名称。',
           '应用程序已关闭',
           '应用程序退出',
           '应用程序打开',
           '必须至少保留一个有效管理员。',
           '每次服务平均值',
           '平均维修成本',
           '平均收入',
           '平均收入/服务',
           '备份数据库',
           '备份失败',
           '备份文件夹不能为空',
           '备份文件夹：',
           '备份成功',
           '茂物上门电脑服务提醒您，设备维修保修将于以下日期到期：',
           '品牌：',
           '品牌',
           '请编译 HELP.HHP，并将 HELP.CHM 复制到应用程序文件夹。',
           '自动计算：服务日期 + 保修期',
           '无法确定所选供应商',
           '更改密码',
           '为以下用户更改密码',
           '已完成',
           '电脑服务管理器',
           '确认删除',
           '确认密码',
           '确认还原',
           '联系人姓名',
           '联系方式：<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           '已复制到剪贴板',
           '复制地址',
           '复制 Google 地图',
           '复制发票',
           '复制电话',
           '版权所有 © ECOMTECH 2026',
           '数量',
           '创建管理员',
           '创建第一个管理员账户',
           '已创建',
           '创建者：',
           '创建日期',
           'CSV 文件',
           '当前密码',
           '当前密码不正确。',
           '当前选项卡帮助',
           '客户姓名',
           '已选择用于提醒的客户',
           '已选择用于提醒的客户：',
           '安全电脑服务管理器 4.4',
           '仪表板刷新失败',
           '仪表板概览',
           '数据已导出到：',
           '需要指定数据库位置',
           '数据库路径不能为空',
           '数据库路径：',
           '数据库路径已成功更新！',
           '数据库已从备份还原',
           '已成功从备份还原数据库！',
           '数据库已还原',
           '日期',
           '尊敬的 [CUSTOMER]：',
           '详细信息',
           '设备品牌统计',
           '设备型号',
           '设备类型',
           '显示名称',
           '捐赠：<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'D: 驱动器不可用。请选择其他文件夹。',
           '未找到 D: 驱动器。\n请为数据库创建或选择一个文件夹。',
           '未找到驱动器',
           '虚拟数据编辑器已打开',
           '重复零件',
           '重复供应商',
           '编辑模式',
           '编辑模式：更新字段，然后单击“保存”',
           '编辑模式：更新字段，然后单击“更新”',
           '启用 / 禁用',
           '结束日期：',
           '请输入用户名和密码。',
           'Excel 文件',
           '30 天后到期',
           '即将到期',
           '即将到期（< 7 天）',
           '即将到期（≤ 30 天）',
           '即将到期（≤ 7 天）',
           '即将到期（≤30 天）',
           '到期日期起始',
           '到期日期起始：',
           '导出完成',
           '导出错误',
           '导出失败',
           '导出成功',
           '已导出到：',
           '导出失败',
           '财务报告',
           '首次管理员会话',
           '首次数据库设置',
           '首次设置：\n请选择用于保存新数据库的文件夹。',
           '表单已清空',
           '起始日期不能晚于结束日期',
           '完整地址',
           'Google 地图',
           'Google 地图链接已复制',
           'Google 地图链接为空',
           'Google 地图链接为空（剪贴板已清空）',
           'Google 地图：',
           '帮助内容',
           'ID（ROWID）',
           '导入完成',
           '导入失败',
           '导入/导出文件夹不能为空',
           '导入/导出文件夹：',
           '进行中',
           '维修中',
           '凭据无效',
           '日期范围无效',
           '驱动器无效',
           '路径无效',
           '用户角色无效。',
           '用户名无效',
           '用户名或密码无效。',
           '保修日期范围无效：起始日期晚于结束日期',
           '发票编号',
           '问题',
           '上次登录',
           '不支持旧版 .XLS 文件。请先将文件另存为 .XLSX 或 CSV。',
           '锁定至',
           '记录提醒',
           '记录提醒',
           '记录保修提醒',
           '登录',
           '登录成功。',
           '登录被阻止',
           '登录失败',
           '登录成功',
           '退出登录',
           '低库存零件',
           '主窗口已打开',
           '至少 4 个字符。允许任意字母、数字、空格或符号组合。',
           '设备型号',
           '型号',
           '多语言虚拟数据创建器已打开',
           '我的最近活动',
           '不适用',
           '客户姓名',
           '姓名为必填项',
           '姓名：',
           '新密码',
           '无数据',
           '未选择数据库',
           '不使用任何默认密码。密码仅以唯一的加盐 scrypt 哈希形式存储。请妥善保管此管理员密码。',
           '未选择现有数据库。应用程序将继续创建新数据库。',
           '未发送任何短信或 WhatsApp 消息。',
           '无状态数据',
           '无保修',
           '不提供任何形式的保证',
           '电话号码',
           '未找到',
           '打开现有数据库或为此应用程序创建新数据库。',
           '打开应用程序帮助内容',
           '打开虚拟数据编辑器',
           '打开 Google 地图',
           '打开当前选项卡的帮助',
           '零件名称',
           '零件名称为必填项',
           '零件类型',
           '零件已刷新',
           '密码',
           '密码已更改',
           '密码已更改。',
           '密码过长。',
           '密码不匹配',
           '密码必须至少包含 4 个字符。',
           '密码未更改',
           '密码已重置',
           '密码已更改',
           '密码更改失败',
           '密码已重置',
           '电话号码',
           '电话：',
           '请将 DUMMY_CREATOR.EXE 或 DUMMY_CREATOR.PY 放在应用程序旁边。',
           '如需进一步服务，请联系我们。',
           '请选择 XLSX 或 CSV 文件',
           '请选择 Excel 或 CSV 文件',
           '采购发票：',
           '数量',
           '最近的安全和数据活动',
           '最近的服务记录',
           '刷新仪表板',
           '重新加载下拉列表数据',
           '提醒消息模板',
           '提醒已记录',
           '提醒已发送',
           '维修成本',
           '报告导出失败',
           '报告已导出',
           '重置密码',
           '重置用户密码',
           '从今天起重置保修范围',
           '还原数据库',
           '还原失败',
           '还原成功',
           '角色',
           '保存导出文件',
           '保存报告导出',
           '按姓名、电话、设备或问题搜索...',
           '安全管理员设置',
           '安全登录',
           '请选择要删除的联系人',
           '请选择要编辑的联系人',
           '请选择要删除的零件',
           '请选择要编辑的零件',
           '请选择要更新的零件',
           '请选择要删除的供应商',
           '请选择要编辑的供应商',
           '请选择要更新的供应商',
           '请先选择用户。',
           '请选择要编辑的保修记录行',
           '选择应用程序货币：',
           '选择应用程序语言和区域设置：',
           '选择应用程序默认货币：',
           '选择语言 / 区域设置',
           '选择要导出的行',
           '选择要记录提醒的保修',
           '选择要发送提醒的保修',
           '发送提醒',
           '服务说明',
           '茂物上门电脑服务',
           '服务状态图表',
           '服务摘要',
           '服务摘要报告',
           '为以下用户设置新密码',
           '设置数据库位置',
           '将服务日期设为今天',
           '设置已保存',
           '设置失败',
           '显示密码',
           '显示密码',
           '请登录以继续',
           '已登录',
           '登录身份',
           '员工',
           '员工访问权限受限',
           '员工可以添加新的服务记录并查看运营数据。编辑、删除、导入、导出、备份、管理设置、用户管理、供应商更改、零件更改、报告和提醒均需要管理员权限。',
           '状态',
           '存储位置',
           '供应商地址',
           '供应商名称',
           '供应商名称为必填项',
           '供应商已刷新',
           '系统',
           '技术员姓名',
           '技术员绩效',
           '技术员绩效报告',
           '技术员',
           '谢谢，',
           '该用户名已存在。',
           '备份文件夹所在的驱动器不存在。',
           '数据库路径所在的驱动器不存在。',
           '导出路径所在的驱动器不存在。',
           '导入/导出文件夹所在的驱动器不存在。',
           '密码已重置。',
           '密码不匹配。',
           '此账户已禁用。请联系管理员。',
           '此操作仅限管理员账户。',
           '本软件按原样提供',
           '时间',
           '导出到 Excel',
           '有效保修总数',
           '客户总数',
           '设备总数',
           '已过期总数',
           '即将到期总数',
           '总收入',
           '已选总数',
           '服务记录总数',
           '服务总价值',
           '服务总数',
           '类型：',
           '唯一型号',
           '未知或无效的凭据',
           '要导出的表未知。',
           '解锁',
           '不支持的格式',
           '使用 D:\\BACKUP（默认）',
           '用户',
           '用户管理',
           '用户未创建',
           '未找到用户。',
           '用户请求退出登录',
           '用户状态已更新。',
           '用户名',
           '用户名已存在',
           '用户名必须为 3–32 个字符，只能使用字母、数字、句点、连字符或下划线。',
           '用户已创建',
           '用户已禁用',
           '用户已启用',
           '用户已解锁',
           '等待零件',
           '等待零件',
           '警告：还原操作将使用备份数据替换所有当前数据。\n此操作无法撤销。\n\n是否继续？',
           '保修期',
           '保修状态：',
           '保修摘要',
           '保修摘要报告',
           '弱密码',
           '您希望如何处理数据库？',
           '窗口已关闭',
           '不能禁用当前账户。',
           '密码已成功更改。',
           '⚠️ 警告！',
           '✏️ 编辑',
           '✏️ 编辑联系人',
           '❌ 删除',
           '➕ 添加',
           '➕ 添加零件',
           '📄 文本视图',
           '📋 表格视图',
           '📝 记录提醒',
           '📝 记录提醒',
           '📤 导出所选项',
           '🔃 刷新',
           '🔄 刷新',
           '🗑️ 清空'],
 'ja_JP': ['新しい部品は管理者が作成する必要があります。',
           '新しい仕入先は管理者が作成する必要があります。',
           '同じ名前の部品が既に存在します。選択して［保存］を使用し、更新してください。',
           '同じ名前の仕入先が既に存在します。選択して［保存］を使用し、更新してください。',
           'アクセス拒否',
           'アカウント',
           'アカウント無効',
           'アカウントロック',
           '操作がブロックされました',
           '有効な保証',
           'アプリケーションユーザーを追加',
           'ユーザーを追加',
           '追加メモ',
           '完全な住所',
           '管理者',
           '管理者',
           '管理者の操作がブロックされました',
           '管理者のみ',
           '管理者が必要です',
           '管理者が必要です。',
           '別の部品が既にこの名前を使用しています。',
           '別の仕入先が既にこの名前を使用しています。',
           'アプリケーションを終了しました',
           'アプリケーション終了',
           'アプリケーション起動',
           '少なくとも1人の有効な管理者が必要です。',
           'サービス1件あたりの平均',
           '平均修理費用',
           '平均売上',
           '平均売上／サービス',
           'データベースをバックアップ',
           'バックアップ失敗',
           'バックアップフォルダーを空にすることはできません',
           'バックアップフォルダー：',
           'バックアップ成功',
           'ボゴール訪問コンピューターサービスから、デバイス修理保証の有効期限をお知らせします：',
           'ブランド：',
           'ブランド',
           'HELP.HHP をビルドし、HELP.CHM をアプリケーションフォルダーにコピーしてください。',
           '自動計算：サービス日 + 保証期間',
           '選択した仕入先を特定できません',
           'パスワードを変更',
           '次のユーザーのパスワードを変更',
           '完了',
           'コンピューターサービス管理',
           '削除の確認',
           'パスワードの確認',
           '復元の確認',
           '連絡先名',
           '連絡先：<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'クリップボードにコピーしました',
           '住所をコピー',
           'Google マップをコピー',
           '請求書をコピー',
           '電話番号をコピー',
           '著作権 © ECOMTECH 2026',
           '件数',
           '管理者を作成',
           '最初の管理者アカウントを作成',
           '作成済み',
           '作成者：',
           '作成日',
           'CSV ファイル',
           '現在のパスワード',
           '現在のパスワードが正しくありません。',
           '現在のタブのヘルプ',
           '顧客名',
           'リマインダー対象として選択された顧客',
           'リマインダー対象として選択された顧客：',
           'セキュア コンピューターサービス管理 4.4',
           'ダッシュボードの更新に失敗しました',
           'ダッシュボード概要',
           'データのエクスポート先：',
           'データベースの場所が必要です',
           'データベースパスを空にすることはできません',
           'データベースパス：',
           'データベースパスを更新しました！',
           'バックアップからデータベースを復元しました',
           'バックアップからデータベースを復元しました！',
           'データベース復元済み',
           '日付',
           '[CUSTOMER] 様',
           '詳細',
           'デバイスブランド統計',
           'デバイスモデル',
           'デバイスの種類',
           '表示名',
           '寄付：<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'ドライブ D: は使用できません。別のフォルダーを選択してください。',
           'ドライブ D: が見つかりません。\nデータベース用のフォルダーを作成または選択してください。',
           'ドライブが見つかりません',
           'ダミーデータエディターを開きました',
           '部品が重複しています',
           '仕入先が重複しています',
           '編集モード',
           '編集モード：フィールドを更新して［保存］をクリックしてください',
           '編集モード：フィールドを更新して［更新］をクリックしてください',
           '有効化 / 無効化',
           '終了日：',
           'ユーザー名とパスワードを入力してください。',
           'Excel ファイル',
           '30日後に期限切れ',
           'まもなく期限切れ',
           'まもなく期限切れ（7日未満）',
           'まもなく期限切れ（30日以内）',
           'まもなく期限切れ（7日以内）',
           'まもなく期限切れ（30日以内）',
           '有効期限の開始日',
           '有効期限の開始日：',
           'エクスポート完了',
           'エクスポートエラー',
           'エクスポート失敗',
           'エクスポート成功',
           'エクスポート先：',
           'エクスポートできませんでした',
           '財務レポート',
           '初回管理者セッション',
           '初回データベース設定',
           '初回設定：\n新しいデータベースを保存するフォルダーを選択してください。',
           'フォームをクリアしました',
           '開始日を終了日より後にすることはできません',
           '完全な住所',
           'Google マップ',
           'Google マップのリンクをコピーしました',
           'Google マップのリンクが空です',
           'Google マップのリンクが空です（クリップボードをクリアしました）',
           'Google マップ：',
           'ヘルプの内容',
           'ID（ROWID）',
           'インポート完了',
           'インポート失敗',
           'インポート／エクスポートフォルダーを空にすることはできません',
           'インポート／エクスポートフォルダー：',
           '進行中',
           '修理中',
           '認証情報が無効です',
           '日付範囲が無効です',
           'ドライブが無効です',
           'パスが無効です',
           'ユーザーの役割が無効です。',
           'ユーザー名が無効です',
           'ユーザー名またはパスワードが無効です。',
           '保証の日付範囲が無効です：開始日が終了日より後です',
           '請求書番号',
           '問題',
           '最終ログイン',
           '旧形式の .XLS ファイルはサポートされていません。先に .XLSX または CSV として保存してください。',
           'ロック期限',
           'リマインダーを記録',
           'リマインダーを記録',
           '保証リマインダーを記録',
           'ログイン',
           'ログインしました。',
           'ログインがブロックされました',
           'ログイン失敗',
           'ログイン成功',
           'ログアウト',
           '在庫不足の部品',
           'メインウィンドウを開きました',
           '4文字以上。文字、数字、空白、記号を任意に使用できます。',
           'デバイスモデル',
           'モデル',
           '多言語ダミーデータ作成ツールを開きました',
           '最近のアクティビティ',
           '該当なし',
           '顧客名',
           '名前は必須です',
           '名前：',
           '新しいパスワード',
           'データなし',
           'データベースが選択されていません',
           '既定のパスワードは使用されません。パスワードは固有のソルト付き scrypt ハッシュとしてのみ保存されます。この管理者パスワードを安全に保管してください。',
           '既存のデータベースが選択されていません。アプリケーションは新しいデータベースの作成を続行します。',
           'SMS または WhatsApp メッセージは送信されていません。',
           'ステータスデータなし',
           '保証なし',
           'いかなる種類の保証も提供されません',
           '電話番号',
           '見つかりません',
           '既存のデータベースを開くか、このアプリケーション用の新しいデータベースを作成してください。',
           'アプリケーションのヘルプ内容を開く',
           'ダミーデータエディターを開く',
           'Google マップを開く',
           'アクティブなタブのヘルプを開く',
           '部品名',
           '部品名は必須です',
           '部品の種類',
           '部品を更新しました',
           'パスワード',
           'パスワードを変更しました',
           'パスワードを変更しました。',
           'パスワードが長すぎます。',
           'パスワードが一致しません',
           'パスワードは4文字以上である必要があります。',
           'パスワードは変更されませんでした',
           'パスワードをリセットしました',
           'パスワード変更済み',
           'パスワード変更失敗',
           'パスワードリセット済み',
           '電話番号',
           '電話：',
           'DUMMY_CREATOR.EXE または DUMMY_CREATOR.PY をアプリケーションと同じ場所に置いてください。',
           '追加のサービスが必要な場合はお問い合わせください。',
           'XLSX または CSV ファイルを選択してください',
           'Excel または CSV ファイルを選択してください',
           '仕入請求書：',
           '数量',
           '最近のセキュリティおよびデータアクティビティ',
           '最近のサービス記録',
           'ダッシュボードを更新',
           'ドロップダウンデータを再読み込み',
           'リマインダーメッセージのテンプレート',
           'リマインダーを記録しました',
           'リマインダーを送信しました',
           '修理費用',
           'レポートのエクスポートに失敗しました',
           'レポートをエクスポートしました',
           'パスワードをリセット',
           'ユーザーのパスワードをリセット',
           '保証範囲を今日からにリセット',
           'データベースを復元',
           '復元失敗',
           '復元成功',
           '役割',
           'エクスポートファイルを保存',
           'レポートのエクスポートを保存',
           '名前、電話番号、デバイス、または問題で検索...',
           '安全な管理者設定',
           '安全なログイン',
           '削除する連絡先を選択してください',
           '編集する連絡先を選択してください',
           '削除する部品を選択してください',
           '編集する部品を選択してください',
           '更新する部品を選択してください',
           '削除する仕入先を選択してください',
           '編集する仕入先を選択してください',
           '更新する仕入先を選択してください',
           '先にユーザーを選択してください。',
           '編集する保証行を選択してください',
           'アプリケーションの通貨を選択：',
           'アプリケーションの言語とロケールを選択：',
           'アプリケーションの既定通貨を選択：',
           '言語 / ロケールを選択',
           'エクスポートする行を選択',
           'リマインダーを記録する保証を選択',
           'リマインダーを送信する保証を選択',
           'リマインダーを送信',
           'サービスの説明',
           'ボゴール訪問コンピューターサービス',
           'サービスステータスグラフ',
           'サービス概要',
           'サービス概要レポート',
           '次のユーザーに新しいパスワードを設定',
           'データベースの場所を設定',
           'サービス日を今日に設定',
           '設定を保存しました',
           '設定失敗',
           'パスワードを表示',
           'パスワードを表示',
           '続行するにはログインしてください',
           'ログイン中',
           'ログインユーザー',
           'スタッフ',
           'スタッフのアクセスは制限されています',
           'スタッフは新しいサービス記録の追加と運用データの表示ができます。編集、削除、インポート、エクスポート、バックアップ、管理設定、ユーザー管理、仕入先の変更、部品の変更、レポート、リマインダーには管理者権限が必要です。',
           'ステータス',
           '保存場所',
           '仕入先住所',
           '仕入先名',
           '仕入先名は必須です',
           '仕入先を更新しました',
           'システム',
           '技術者名',
           '技術者の実績',
           '技術者実績レポート',
           '技術者',
           'ありがとうございます。',
           'そのユーザー名は既に存在します。',
           'バックアップフォルダーのドライブが存在しません。',
           'データベースパスのドライブが存在しません。',
           'エクスポートパスのドライブが存在しません。',
           'インポート／エクスポートフォルダーのドライブが存在しません。',
           'パスワードをリセットしました。',
           'パスワードが一致しません。',
           'このアカウントは無効です。管理者に連絡してください。',
           'この操作は管理者アカウントに制限されています。',
           '本ソフトウェアは現状のまま提供されます',
           '時刻',
           'Excel へ',
           '有効な保証の合計',
           '顧客合計',
           'デバイス合計',
           '期限切れ合計',
           'まもなく期限切れの合計',
           '総売上',
           '選択件数合計',
           'サービス記録合計',
           'サービス総額',
           'サービス合計',
           '種類：',
           '固有モデル',
           '不明または無効な認証情報',
           'エクスポート対象のテーブルが不明です。',
           'ロック解除',
           'サポートされていない形式',
           'D:\\BACKUP を使用（既定）',
           'ユーザー',
           'ユーザー管理',
           'ユーザーは作成されませんでした',
           'ユーザーが見つかりません。',
           'ユーザーがログアウトを要求しました',
           'ユーザーステータスを更新しました。',
           'ユーザー名',
           'ユーザー名は既に存在します',
           'ユーザー名は3～32文字で、文字、数字、ピリオド、ハイフン、またはアンダースコアを使用してください。',
           'ユーザー作成済み',
           'ユーザー無効',
           'ユーザー有効',
           'ユーザーロック解除済み',
           '部品待ち',
           '部品待ち',
           '警告：復元すると、現在のすべてのデータがバックアップデータに置き換えられます。\nこの操作は元に戻せません。\n\n続行しますか？',
           '保証期間',
           '保証ステータス：',
           '保証概要',
           '保証概要レポート',
           '脆弱なパスワード',
           'データベースに対して何を行いますか？',
           'ウィンドウを閉じました',
           '現在のアカウントを無効にすることはできません。',
           'パスワードを変更しました。',
           '⚠️ 警告！',
           '✏️ 編集',
           '✏️ 連絡先を編集',
           '❌ 削除',
           '➕ 追加',
           '➕ 部品を追加',
           '📄 テキスト表示',
           '📋 テーブル表示',
           '📝 リマインダーを記録',
           '📝 リマインダーを記録',
           '📤 選択項目をエクスポート',
           '🔃 更新',
           '🔄 更新',
           '🗑️ クリア'],
 'ko_KR': ['새 부품은 관리자가 생성해야 합니다.',
           '새 공급업체는 관리자가 생성해야 합니다.',
           '같은 이름의 부품이 이미 있습니다. 해당 부품을 선택하고 저장을 사용하여 업데이트하세요.',
           '같은 이름의 공급업체가 이미 있습니다. 해당 공급업체를 선택하고 저장을 사용하여 업데이트하세요.',
           '접근_거부',
           '계정',
           '계정 비활성화됨',
           '계정 잠김',
           '작업 차단됨',
           '유효한 보증',
           '애플리케이션 사용자 추가',
           '사용자 추가',
           '추가 메모',
           '전체 주소',
           '관리자',
           '관리자',
           '관리자 작업 차단됨',
           '관리자 전용',
           '관리자 권한 필요',
           '관리자 권한이 필요합니다.',
           '다른 부품에서 이미 이 이름을 사용하고 있습니다.',
           '다른 공급업체에서 이미 이 이름을 사용하고 있습니다.',
           '애플리케이션 종료됨',
           '애플리케이션_종료',
           '애플리케이션_실행',
           '활성 관리자 계정이 최소 하나 필요합니다.',
           '서비스당 평균',
           '평균 수리 비용',
           '평균 수익',
           '평균 수익/서비스',
           '데이터베이스 백업',
           '백업 실패',
           '백업 폴더는 비워 둘 수 없습니다',
           '백업 폴더:',
           '백업 성공',
           '보고르 방문 컴퓨터 서비스에서 기기 수리 보증 만료일을 알려드립니다:',
           '브랜드:',
           '브랜드',
           'HELP.HHP를 빌드하고 HELP.CHM을 애플리케이션 폴더에 복사하세요.',
           '자동 계산: 서비스 날짜 + 보증 기간',
           '선택한 공급업체를 확인할 수 없습니다',
           '비밀번호 변경',
           '다음 사용자의 비밀번호 변경',
           '완료됨',
           '컴퓨터 서비스 관리자',
           '삭제 확인',
           '비밀번호 확인',
           '복원 확인',
           '연락처 이름',
           '연락처:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           '클립보드에 복사됨',
           '주소 복사',
           'Google 지도 복사',
           '송장 복사',
           '전화번호 복사',
           '저작권 © ECOMTECH 2026',
           '개수',
           '관리자 생성',
           '첫 번째 관리자 계정 생성',
           '생성됨',
           '생성자:',
           '생성 날짜',
           'CSV 파일',
           '현재 비밀번호',
           '현재 비밀번호가 올바르지 않습니다.',
           '현재 탭 도움말',
           '고객 이름',
           '알림 대상으로 선택된 고객',
           '알림 대상으로 선택된 고객:',
           '보안 컴퓨터 서비스 관리자 4.4',
           '대시보드 새로 고침 실패',
           '대시보드 개요',
           '데이터 내보내기 위치:',
           '데이터베이스 위치 필요',
           '데이터베이스 경로는 비워 둘 수 없습니다',
           '데이터베이스 경로:',
           '데이터베이스 경로가 성공적으로 업데이트되었습니다!',
           '백업에서 데이터베이스 복원됨',
           '백업에서 데이터베이스를 성공적으로 복원했습니다!',
           '데이터베이스_복원됨',
           '날짜',
           '[CUSTOMER] 고객님께,',
           '세부 정보',
           '기기 브랜드 통계',
           '기기 모델',
           '기기 유형',
           '표시 이름',
           '후원:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'D: 드라이브를 사용할 수 없습니다. 다른 폴더를 선택하세요.',
           'D: 드라이브를 찾을 수 없습니다.\n데이터베이스용 폴더를 만들거나 선택하세요.',
           '드라이브를 찾을 수 없음',
           '더미 데이터 편집기 열림',
           '중복 부품',
           '중복 공급업체',
           '편집 모드',
           '편집 모드: 필드를 수정한 다음 저장을 클릭하세요',
           '편집 모드: 필드를 수정한 다음 업데이트를 클릭하세요',
           '활성화 / 비활성화',
           '종료 날짜:',
           '사용자 이름과 비밀번호를 모두 입력하세요.',
           'Excel 파일',
           '30일 후 만료',
           '곧 만료',
           '곧 만료 (< 7일)',
           '곧 만료 (≤ 30일)',
           '곧 만료 (≤ 7일)',
           '곧 만료 (≤30일)',
           '만료 시작 날짜',
           '만료 시작 날짜:',
           '내보내기 완료',
           '내보내기 오류',
           '내보내기 실패',
           '내보내기 성공',
           '내보낸 위치:',
           '내보내기에 실패했습니다',
           '재무 보고서',
           '첫 관리자 세션',
           '최초 데이터베이스 설정',
           '최초 설정:\n새 데이터베이스를 저장할 폴더를 선택하세요.',
           '양식 지워짐',
           '시작 날짜는 종료 날짜보다 늦을 수 없습니다',
           '전체 주소',
           'Google 지도',
           'Google 지도 링크 복사됨',
           'Google 지도 링크가 비어 있음',
           'Google 지도 링크가 비어 있음 (클립보드 지워짐)',
           'Google 지도:',
           '도움말 목차',
           'ID (ROWID)',
           '가져오기 완료',
           '가져오기 실패',
           '가져오기/내보내기 폴더는 비워 둘 수 없습니다',
           '가져오기/내보내기 폴더:',
           '진행 중',
           '수리 중',
           '잘못된 자격 증명',
           '잘못된 날짜 범위',
           '잘못된 드라이브',
           '잘못된 경로',
           '잘못된 사용자 역할입니다.',
           '잘못된 사용자 이름',
           '사용자 이름 또는 비밀번호가 올바르지 않습니다.',
           '잘못된 보증 날짜 범위: 시작 날짜가 종료 날짜보다 늦습니다',
           '송장 번호',
           '문제',
           '마지막 로그인',
           '이전 .XLS 파일은 지원되지 않습니다. 먼저 파일을 .XLSX 또는 CSV로 저장하세요.',
           '잠금 해제 시간',
           '알림 기록',
           '알림 기록',
           '보증 알림 기록',
           '로그인',
           '로그인에 성공했습니다.',
           '로그인_차단됨',
           '로그인_실패',
           '로그인_성공',
           '로그아웃',
           '재고 부족 부품',
           '기본 창 열림',
           '최소 4자입니다. 문자, 숫자, 공백 또는 기호를 자유롭게 사용할 수 있습니다.',
           '기기 모델',
           '모델',
           '다국어 더미 데이터 생성기 열림',
           '내 최근 활동',
           '해당 없음',
           '고객 이름',
           '이름은 필수입니다',
           '이름:',
           '새 비밀번호',
           '데이터 없음',
           '선택된 데이터베이스 없음',
           '기본 비밀번호는 사용하지 않습니다. 비밀번호는 고유한 솔트 적용 scrypt 해시로만 저장됩니다. 이 관리자 비밀번호를 안전하게 보관하세요.',
           '기존 데이터베이스가 선택되지 않았습니다. 애플리케이션이 새 데이터베이스 생성을 계속합니다.',
           'SMS 또는 WhatsApp 메시지가 전송되지 않았습니다.',
           '상태 데이터 없음',
           '보증 없음',
           '어떠한 형태의 보증도 제공되지 않습니다',
           '전화번호',
           '찾을 수 없음',
           '기존 데이터베이스를 열거나 이 애플리케이션용 새 데이터베이스를 만드세요.',
           '애플리케이션 도움말 목차 열기',
           '더미 데이터 편집기 열기',
           'Google 지도 열기',
           '활성 탭 도움말 열기',
           '부품 이름',
           '부품 이름은 필수입니다',
           '부품 유형',
           '부품 새로 고침됨',
           '비밀번호',
           '비밀번호 변경됨',
           '비밀번호가 변경되었습니다.',
           '비밀번호가 너무 깁니다.',
           '비밀번호 불일치',
           '비밀번호는 최소 4자를 포함해야 합니다.',
           '비밀번호 변경되지 않음',
           '비밀번호 재설정됨',
           '비밀번호_변경됨',
           '비밀번호_변경_실패',
           '비밀번호_재설정됨',
           '전화번호',
           '전화:',
           'DUMMY_CREATOR.EXE 또는 DUMMY_CREATOR.PY를 애플리케이션 옆에 두세요.',
           '추가 서비스가 필요하면 문의해 주세요.',
           'XLSX 또는 CSV 파일을 선택하세요',
           'Excel 또는 CSV 파일을 선택하세요',
           '구매 송장:',
           '수량',
           '최근 보안 및 데이터 활동',
           '최근 서비스 기록',
           '대시보드 새로 고침',
           '드롭다운 데이터 다시 불러오기',
           '알림 메시지 템플릿',
           '알림 기록됨',
           '알림 전송됨',
           '수리 비용',
           '보고서 내보내기 실패',
           '보고서 내보냄',
           '비밀번호 재설정',
           '사용자 비밀번호 재설정',
           '오늘부터 보증 범위 재설정',
           '데이터베이스 복원',
           '복원 실패',
           '복원 성공',
           '역할',
           '내보내기 파일 저장',
           '보고서 내보내기 저장',
           '이름, 전화번호, 기기 또는 문제로 검색...',
           '안전한 관리자 설정',
           '보안 로그인',
           '삭제할 연락처를 선택하세요',
           '편집할 연락처를 선택하세요',
           '삭제할 부품을 선택하세요',
           '편집할 부품을 선택하세요',
           '업데이트할 부품을 선택하세요',
           '삭제할 공급업체를 선택하세요',
           '편집할 공급업체를 선택하세요',
           '업데이트할 공급업체를 선택하세요',
           '먼저 사용자를 선택하세요.',
           '편집할 보증 행을 선택하세요',
           '애플리케이션 통화 선택:',
           '애플리케이션 언어 및 로캘 선택:',
           '애플리케이션 기본 통화 선택:',
           '언어 / 로캘 선택',
           '내보낼 행 선택',
           '알림을 기록할 보증 선택',
           '알림을 보낼 보증 선택',
           '알림 보내기',
           '서비스 설명',
           '보고르 방문 컴퓨터 서비스',
           '서비스 상태 그래프',
           '서비스 요약',
           '서비스 요약 보고서',
           '다음 사용자의 새 비밀번호 설정',
           '데이터베이스 위치 설정',
           '서비스 날짜를 오늘로 설정',
           '설정 저장됨',
           '설정 실패',
           '비밀번호 표시',
           '비밀번호 표시',
           '계속하려면 로그인하세요',
           '로그인됨',
           '로그인 사용자',
           '직원',
           '직원 접근 권한이 제한됨',
           '직원은 새 서비스 기록을 추가하고 운영 데이터를 볼 수 있습니다. 편집, 삭제, 가져오기, 내보내기, 백업, 관리자 설정, 사용자 관리, 공급업체 변경, 부품 변경, 보고서 및 알림에는 '
           '관리자 권한이 필요합니다.',
           '상태',
           '저장 위치',
           '공급업체 주소',
           '공급업체 이름',
           '공급업체 이름은 필수입니다',
           '공급업체 새로 고침됨',
           '시스템',
           '기술자 이름',
           '기술자 실적',
           '기술자 실적 보고서',
           '기술자',
           '감사합니다.',
           '해당 사용자 이름이 이미 있습니다.',
           '백업 폴더의 드라이브가 존재하지 않습니다.',
           '데이터베이스 경로의 드라이브가 존재하지 않습니다.',
           '내보내기 경로의 드라이브가 존재하지 않습니다.',
           '가져오기/내보내기 폴더의 드라이브가 존재하지 않습니다.',
           '비밀번호가 재설정되었습니다.',
           '비밀번호가 일치하지 않습니다.',
           '이 계정은 비활성화되었습니다. 관리자에게 문의하세요.',
           '이 작업은 관리자 계정으로 제한됩니다.',
           '이 소프트웨어는 있는 그대로 제공됩니다',
           '시간',
           'Excel로',
           '전체 유효 보증',
           '전체 고객',
           '전체 기기',
           '전체 만료',
           '전체 곧 만료',
           '총수익',
           '전체 선택됨',
           '전체 서비스 기록',
           '총 서비스 금액',
           '전체 서비스',
           '유형:',
           '고유 모델',
           '알 수 없거나 잘못된 자격 증명',
           '내보낼 수 없는 알 수 없는 테이블입니다.',
           '잠금 해제',
           '지원되지 않는 형식',
           'D:\\BACKUP 사용 (기본값)',
           '사용자',
           '사용자 관리',
           '사용자가 생성되지 않음',
           '사용자를 찾을 수 없습니다.',
           '사용자가 로그아웃을 요청함',
           '사용자 상태가 업데이트되었습니다.',
           '사용자 이름',
           '사용자 이름이 이미 존재함',
           '사용자 이름은 3~32자이며 문자, 숫자, 마침표, 하이픈 또는 밑줄을 사용해야 합니다.',
           '사용자_생성됨',
           '사용자_비활성화됨',
           '사용자_활성화됨',
           '사용자_잠금해제됨',
           '부품 대기 중',
           '부품 대기 중',
           '경고: 복원하면 현재 모든 데이터가 백업 데이터로 대체됩니다.\n이 작업은 실행 취소할 수 없습니다.\n\n계속하시겠습니까?',
           '보증 기간',
           '보증 상태:',
           '보증 요약',
           '보증 요약 보고서',
           '취약한 비밀번호',
           '데이터베이스에 대해 무엇을 하시겠습니까?',
           '창 닫힘',
           '현재 계정은 비활성화할 수 없습니다.',
           '비밀번호가 성공적으로 변경되었습니다.',
           '⚠️ 경고!',
           '✏️ 편집',
           '✏️ 연락처 편집',
           '❌ 삭제',
           '➕ 추가',
           '➕ 부품 추가',
           '📄 텍스트 보기',
           '📋 표 보기',
           '📝 알림 기록',
           '📝 알림 기록',
           '📤 선택 항목 내보내기',
           '🔃 새로 고침',
           '🔄 새로 고침',
           '🗑️ 지우기'],
 'ar_SA': ['يجب أن ينشئ المسؤول قطعة جديدة.',
           'يجب أن ينشئ المسؤول مورّدًا جديدًا.',
           'توجد قطعة بهذا الاسم بالفعل. حدّدها واستخدم حفظ لتحديثها.',
           'يوجد مورّد بهذا الاسم بالفعل. حدّده واستخدم حفظ لتحديثه.',
           'تم_رفض_الوصول',
           'الحساب',
           'الحساب معطّل',
           'الحساب مقفل',
           'تم حظر الإجراء',
           'الضمانات النشطة',
           'إضافة مستخدم للتطبيق',
           'إضافة مستخدم',
           'ملاحظات إضافية',
           'العنوان الكامل',
           'مسؤول',
           'المسؤول',
           'تم حظر إجراء المسؤول',
           'للمسؤول فقط',
           'يلزم مسؤول',
           'يلزم مسؤول.',
           'تستخدم قطعة أخرى هذا الاسم بالفعل.',
           'يستخدم مورّد آخر هذا الاسم بالفعل.',
           'تم إغلاق التطبيق',
           'خروج_التطبيق',
           'فتح_التطبيق',
           'يجب توفر مسؤول نشط واحد على الأقل.',
           'المتوسط لكل خدمة',
           'متوسط تكلفة الإصلاح',
           'متوسط الإيراد',
           'متوسط الإيراد/الخدمة',
           'نسخ قاعدة البيانات احتياطيًا',
           'فشل النسخ الاحتياطي',
           'لا يمكن أن يكون مجلد النسخ الاحتياطي فارغًا',
           'مجلد النسخ الاحتياطي:',
           'نجح النسخ الاحتياطي',
           'تذكّرك خدمة صيانة الكمبيوتر المنزلية في بوغور بأن ضمان إصلاح جهازك سينتهي في:',
           'العلامة التجارية:',
           'العلامات التجارية',
           'أنشئ HELP.HHP وانسخ HELP.CHM إلى مجلد التطبيق.',
           'يُحسب تلقائيًا: تاريخ الخدمة + مدة الضمان',
           'تعذّر تحديد المورّد المحدد',
           'تغيير كلمة المرور',
           'تغيير كلمة المرور لـ',
           'مكتمل',
           'مدير خدمات الكمبيوتر',
           'تأكيد الحذف',
           'تأكيد كلمة المرور',
           'تأكيد الاستعادة',
           'اسم جهة الاتصال',
           'للتواصل:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'تم النسخ إلى الحافظة',
           'نسخ العنوان',
           'نسخ خريطة Google',
           'نسخ الفاتورة',
           'نسخ الهاتف',
           'حقوق النشر © ECOMTECH 2026',
           'العدد',
           'إنشاء مسؤول',
           'إنشاء أول حساب مسؤول',
           'تم الإنشاء',
           'أنشأه:',
           'تاريخ الإنشاء',
           'ملفات CSV',
           'كلمة المرور الحالية',
           'كلمة المرور الحالية غير صحيحة.',
           'مساعدة علامة التبويب الحالية',
           'اسم العميل',
           'العملاء المحددون للتذكير',
           'العملاء المحددون للتذكير:',
           'مدير خدمات الكمبيوتر الآمن 4.4',
           'فشل تحديث لوحة المعلومات',
           'نظرة عامة على لوحة المعلومات',
           'تم تصدير البيانات إلى:',
           'يلزم تحديد موقع قاعدة البيانات',
           'لا يمكن أن يكون مسار قاعدة البيانات فارغًا',
           'مسار قاعدة البيانات:',
           'تم تحديث مسارات قاعدة البيانات بنجاح!',
           'تمت استعادة قاعدة البيانات من النسخة الاحتياطية',
           'تمت استعادة قاعدة البيانات من النسخة الاحتياطية بنجاح!',
           'تمت_استعادة_قاعدة_البيانات',
           'التاريخ',
           'عزيزي/عزيزتي [CUSTOMER]،',
           'التفاصيل',
           'إحصاءات العلامات التجارية للأجهزة',
           'طراز الجهاز',
           'أنواع الأجهزة',
           'اسم العرض',
           'تبرع:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'محرك الأقراص D: غير متاح. يُرجى اختيار مجلد آخر.',
           'لم يتم العثور على محرك الأقراص D:.\nيُرجى إنشاء مجلد لقاعدة البيانات أو تحديده.',
           'لم يتم العثور على محرك الأقراص',
           'تم فتح محرر البيانات الوهمية',
           'قطعة مكررة',
           'مورّد مكرر',
           'وضع التحرير',
           'وضع التحرير: حدّث الحقول ثم انقر على حفظ',
           'وضع التحرير: حدّث الحقول ثم انقر على تحديث',
           'تمكين / تعطيل',
           'تاريخ الانتهاء:',
           'أدخل اسم المستخدم وكلمة المرور.',
           'ملفات Excel',
           'ينتهي خلال 30 يومًا',
           'ينتهي قريبًا',
           'ينتهي قريبًا (< 7 أيام)',
           'ينتهي قريبًا (≤ 30 يومًا)',
           'ينتهي قريبًا (≤ 7 أيام)',
           'ينتهي قريبًا (≤30 يومًا)',
           'تاريخ انتهاء الصلاحية من',
           'تاريخ انتهاء الصلاحية من:',
           'اكتمل التصدير',
           'خطأ في التصدير',
           'فشل التصدير',
           'نجح التصدير',
           'تم التصدير إلى:',
           'تعذّر التصدير',
           'التقرير المالي',
           'جلسة المسؤول الأولى',
           'الإعداد الأول لقاعدة البيانات',
           'الإعداد لأول مرة:\nيُرجى اختيار مجلد لحفظ قاعدة البيانات الجديدة.',
           'تم مسح النموذج',
           'لا يمكن أن يكون تاريخ البداية بعد تاريخ النهاية',
           'العنوان الكامل',
           'خريطة Google',
           'تم نسخ رابط خريطة Google',
           'رابط خريطة Google فارغ',
           'رابط خريطة Google فارغ (تم مسح الحافظة)',
           'خريطة Google:',
           'محتويات المساعدة',
           'المعرّف (ROWID)',
           'اكتمل الاستيراد',
           'فشل الاستيراد',
           'لا يمكن أن يكون مجلد الاستيراد/التصدير فارغًا',
           'مجلد الاستيراد/التصدير:',
           'قيد التنفيذ',
           'قيد الإصلاح',
           'بيانات الاعتماد غير صالحة',
           'نطاق التاريخ غير صالح',
           'محرك الأقراص غير صالح',
           'المسار غير صالح',
           'دور المستخدم غير صالح.',
           'اسم المستخدم غير صالح',
           'اسم المستخدم أو كلمة المرور غير صحيحة.',
           'نطاق تاريخ الضمان غير صالح: تاريخ البداية بعد تاريخ النهاية',
           'رقم الفاتورة',
           'المشكلات',
           'آخر تسجيل دخول',
           'ملفات .XLS القديمة غير مدعومة. احفظ الملف أولًا بصيغة .XLSX أو CSV.',
           'مقفل حتى',
           'تسجيل تذكير',
           'تسجيل التذكيرات',
           'تسجيل تذكير الضمان',
           'تسجيل الدخول',
           'تم تسجيل الدخول بنجاح.',
           'تم_حظر_تسجيل_الدخول',
           'فشل_تسجيل_الدخول',
           'نجح_تسجيل_الدخول',
           'تسجيل الخروج',
           'قطع منخفضة المخزون',
           'تم فتح النافذة الرئيسية',
           '4 أحرف على الأقل. يُسمح بأي صيغة تتضمن أحرفًا أو أرقامًا أو مسافات أو رموزًا.',
           'طراز الجهاز',
           'الطُرز',
           'تم فتح منشئ البيانات الوهمية متعدد اللغات',
           'نشاطي الأخير',
           'غير متاح',
           'اسم العميل',
           'الاسم مطلوب',
           'الاسم:',
           'كلمة المرور الجديدة',
           'لا توجد بيانات',
           'لم يتم تحديد قاعدة بيانات',
           'لا تُستخدم أي كلمة مرور افتراضية. تُخزّن كلمات المرور فقط كتجزئات scrypt فريدة ومملحة. احتفظ بكلمة مرور '
           'المسؤول هذه بأمان.',
           'لم يتم تحديد قاعدة بيانات موجودة. سيواصل التطبيق إنشاء قاعدة بيانات جديدة.',
           'لم يتم إرسال أي رسالة SMS أو WhatsApp.',
           'لا توجد بيانات حالة',
           'بلا ضمان',
           'لا يتم تقديم أي ضمان من أي نوع',
           'رقم الهاتف',
           'غير موجود',
           'افتح قاعدة بيانات موجودة أو أنشئ قاعدة بيانات جديدة لهذا التطبيق.',
           'فتح محتويات مساعدة التطبيق',
           'فتح محرر البيانات الوهمية',
           'فتح خريطة Google',
           'فتح المساعدة لعلامة التبويب النشطة',
           'اسم القطعة',
           'اسم القطعة مطلوب',
           'أنواع القطع',
           'تم تحديث القطع',
           'كلمة المرور',
           'تم تغيير كلمة المرور',
           'تم تغيير كلمة المرور.',
           'كلمة المرور طويلة جدًا.',
           'كلمتا المرور غير متطابقتين',
           'يجب أن تحتوي كلمة المرور على 4 أحرف على الأقل.',
           'لم تتغير كلمة المرور',
           'تمت إعادة تعيين كلمة المرور',
           'تم_تغيير_كلمة_المرور',
           'فشل_تغيير_كلمة_المرور',
           'تمت_إعادة_تعيين_كلمة_المرور',
           'رقم الهاتف',
           'الهاتف:',
           'ضع DUMMY_CREATOR.EXE أو DUMMY_CREATOR.PY بجانب التطبيق.',
           'يُرجى التواصل معنا إذا احتجت إلى خدمة إضافية.',
           'يُرجى تحديد ملف XLSX أو CSV',
           'يُرجى تحديد ملف Excel أو CSV',
           'فاتورة الشراء:',
           'الكمية',
           'نشاط الأمان والبيانات الأخير',
           'سجلات الخدمة الأخيرة',
           'تحديث لوحة المعلومات',
           'إعادة تحميل بيانات القوائم المنسدلة',
           'قالب رسالة التذكير',
           'تم تسجيل التذكيرات',
           'تم إرسال التذكيرات',
           'تكلفة الإصلاح',
           'فشل تصدير التقرير',
           'تم تصدير التقرير',
           'إعادة تعيين كلمة المرور',
           'إعادة تعيين كلمة مرور المستخدم',
           'إعادة تعيين نطاق الضمان بدءًا من اليوم',
           'استعادة قاعدة البيانات',
           'فشلت الاستعادة',
           'نجحت الاستعادة',
           'الدور',
           'حفظ ملف التصدير',
           'حفظ تصدير التقرير',
           'البحث بالاسم أو الهاتف أو الجهاز أو المشكلة...',
           'إعداد مسؤول آمن',
           'تسجيل دخول آمن',
           'حدّد جهة اتصال لحذفها',
           'حدّد جهة اتصال لتحريرها',
           'حدّد قطعة لحذفها',
           'حدّد قطعة لتحريرها',
           'حدّد قطعة لتحديثها',
           'حدّد مورّدًا لحذفه',
           'حدّد مورّدًا لتحريره',
           'حدّد مورّدًا لتحديثه',
           'حدّد مستخدمًا أولًا.',
           'حدّد صف ضمان لتحريره',
           'تحديد عملة التطبيق:',
           'تحديد لغة التطبيق وإعداداته المحلية:',
           'تحديد العملة الافتراضية للتطبيق:',
           'تحديد اللغة / الإعدادات المحلية',
           'تحديد الصفوف المراد تصديرها',
           'تحديد الضمانات لتسجيل التذكيرات',
           'تحديد الضمانات لإرسال التذكيرات',
           'إرسال التذكيرات',
           'وصف الخدمة',
           'خدمة صيانة الكمبيوتر المنزلية في بوغور',
           'مخطط حالة الخدمة',
           'ملخص الخدمة',
           'تقرير ملخص الخدمة',
           'تعيين كلمة مرور جديدة لـ',
           'تعيين موقع قاعدة البيانات',
           'تعيين تاريخ الخدمة إلى اليوم',
           'تم حفظ الإعدادات',
           'فشل الإعداد',
           'إظهار كلمة المرور',
           'إظهار كلمات المرور',
           'سجّل الدخول للمتابعة',
           'تم تسجيل الدخول',
           'تم تسجيل الدخول باسم',
           'موظف',
           'وصول الموظف محدود',
           'يمكن للموظفين إضافة سجلات خدمة جديدة وعرض البيانات التشغيلية. يتطلب التحرير والحذف والاستيراد والتصدير '
           'والنسخ الاحتياطي والإعدادات الإدارية وإدارة المستخدمين وتغييرات المورّدين وتغييرات القطع والتقارير '
           'والتذكيرات حساب مسؤول.',
           'الحالات',
           'موقع التخزين',
           'عنوان المورّد',
           'اسم المورّد',
           'اسم المورّد مطلوب',
           'تم تحديث المورّدين',
           'النظام',
           'اسم الفني',
           'أداء الفنيين',
           'تقرير أداء الفنيين',
           'الفنيون',
           'شكرًا لك،',
           'اسم المستخدم هذا موجود بالفعل.',
           'محرك الأقراص الخاص بمجلد النسخ الاحتياطي غير موجود.',
           'محرك الأقراص الخاص بمسار قاعدة البيانات غير موجود.',
           'محرك الأقراص الخاص بمسار التصدير غير موجود.',
           'محرك الأقراص الخاص بمجلد الاستيراد/التصدير غير موجود.',
           'تمت إعادة تعيين كلمة المرور.',
           'كلمتا المرور غير متطابقتين.',
           'هذا الحساب معطّل. تواصل مع أحد المسؤولين.',
           'هذا الإجراء مقيّد بحسابات المسؤولين.',
           'يُقدّم هذا البرنامج كما هو',
           'الوقت',
           'إلى Excel',
           'إجمالي الضمانات النشطة',
           'إجمالي العملاء',
           'إجمالي الأجهزة',
           'إجمالي المنتهية',
           'إجمالي المنتهية قريبًا',
           'إجمالي الإيرادات',
           'إجمالي المحدد',
           'إجمالي سجلات الخدمة',
           'إجمالي قيمة الخدمة',
           'إجمالي الخدمات',
           'النوع:',
           'طُرز فريدة',
           'بيانات اعتماد مجهولة أو غير صالحة',
           'جدول غير معروف للتصدير.',
           'إلغاء القفل',
           'تنسيق غير مدعوم',
           'استخدام D:\\BACKUP (افتراضي)',
           'المستخدم',
           'إدارة المستخدمين',
           'لم يتم إنشاء المستخدم',
           'لم يتم العثور على المستخدم.',
           'طلب المستخدم تسجيل الخروج',
           'تم تحديث حالة المستخدم.',
           'اسم المستخدم',
           'اسم المستخدم موجود',
           'يجب أن يتكون اسم المستخدم من 3 إلى 32 حرفًا باستخدام الأحرف أو الأرقام أو النقطة أو الشرطة أو الشرطة '
           'السفلية.',
           'تم_إنشاء_المستخدم',
           'تم_تعطيل_المستخدم',
           'تم_تمكين_المستخدم',
           'تم_إلغاء_قفل_المستخدم',
           'في انتظار القطع',
           'انتظار القطع',
           'تحذير: ستستبدل الاستعادة جميع البيانات الحالية ببيانات النسخة الاحتياطية.\n'
           'لا يمكن التراجع عن هذا الإجراء.\n'
           '\n'
           'هل تريد المتابعة؟',
           'فترات الضمان',
           'حالة الضمان:',
           'ملخص الضمان',
           'تقرير ملخص الضمان',
           'كلمة مرور ضعيفة',
           'ماذا تريد أن تفعل بقاعدة البيانات؟',
           'تم إغلاق النافذة',
           'لا يمكنك تعطيل حسابك الحالي.',
           'تم تغيير كلمة مرورك بنجاح.',
           '⚠️ تحذير!',
           '✏️ تحرير',
           '✏️ تحرير جهة الاتصال',
           '❌ حذف',
           '➕ إضافة',
           '➕ إضافة قطعة',
           '📄 عرض نصي',
           '📋 عرض جدولي',
           '📝 تسجيل تذكير',
           '📝 تسجيل التذكيرات',
           '📤 تصدير المحدد',
           '🔃 تحديث',
           '🔄 تحديث',
           '🗑️ مسح'],
 'hi_IN': ['नया पार्ट केवल प्रशासक द्वारा बनाया जाना चाहिए।',
           'नया आपूर्तिकर्ता केवल प्रशासक द्वारा बनाया जाना चाहिए।',
           'इस नाम का पार्ट पहले से मौजूद है। उसे चुनें और अपडेट करने के लिए सहेजें का उपयोग करें।',
           'इस नाम का आपूर्तिकर्ता पहले से मौजूद है। उसे चुनें और अपडेट करने के लिए सहेजें का उपयोग करें।',
           'पहुँच_अस्वीकृत',
           'खाता',
           'खाता अक्षम',
           'खाता लॉक',
           'कार्रवाई अवरुद्ध',
           'सक्रिय वारंटियाँ',
           'एप्लिकेशन उपयोगकर्ता जोड़ें',
           'उपयोगकर्ता जोड़ें',
           'अतिरिक्त टिप्पणियाँ',
           'पूरा पता',
           'एडमिन',
           'प्रशासक',
           'प्रशासक की कार्रवाई अवरुद्ध',
           'केवल प्रशासक',
           'प्रशासक आवश्यक',
           'प्रशासक आवश्यक है।',
           'कोई अन्य पार्ट पहले से इस नाम का उपयोग कर रहा है।',
           'कोई अन्य आपूर्तिकर्ता पहले से इस नाम का उपयोग कर रहा है।',
           'एप्लिकेशन बंद हुआ',
           'एप्लिकेशन_बंद',
           'एप्लिकेशन_खुला',
           'कम से कम एक सक्रिय प्रशासक आवश्यक है।',
           'प्रति सेवा औसत',
           'औसत मरम्मत लागत',
           'औसत आय',
           'औसत आय/सेवा',
           'डेटाबेस का बैकअप लें',
           'बैकअप विफल',
           'बैकअप फ़ोल्डर खाली नहीं हो सकता',
           'बैकअप फ़ोल्डर:',
           'बैकअप सफल',
           'बोगोर ऑन-कॉल कंप्यूटर सेवा आपको याद दिलाती है कि आपके डिवाइस की मरम्मत वारंटी इस दिन समाप्त होगी:',
           'ब्रांड:',
           'ब्रांड',
           'HELP.HHP बनाएँ और HELP.CHM को एप्लिकेशन फ़ोल्डर में कॉपी करें।',
           'स्वतः गणना: सेवा तिथि + वारंटी अवधि',
           'चुने गए आपूर्तिकर्ता का निर्धारण नहीं हो सका',
           'पासवर्ड बदलें',
           'इसके लिए पासवर्ड बदलें',
           'पूर्ण',
           'कंप्यूटर सेवा प्रबंधक',
           'हटाने की पुष्टि करें',
           'पासवर्ड की पुष्टि करें',
           'पुनर्स्थापना की पुष्टि करें',
           'संपर्क नाम',
           'संपर्क:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'क्लिपबोर्ड पर कॉपी किया गया',
           'पता कॉपी करें',
           'Google मानचित्र कॉपी करें',
           'चालान कॉपी करें',
           'फ़ोन कॉपी करें',
           'कॉपीराइट © ECOMTECH 2026',
           'गणना',
           'प्रशासक बनाएँ',
           'पहला प्रशासक खाता बनाएँ',
           'बनाया गया',
           'बनाने वाला:',
           'बनाने की तिथि',
           'CSV फ़ाइलें',
           'वर्तमान पासवर्ड',
           'वर्तमान पासवर्ड गलत है।',
           'वर्तमान टैब सहायता',
           'ग्राहक का नाम',
           'रिमाइंडर के लिए चुने गए ग्राहक',
           'रिमाइंडर के लिए चुने गए ग्राहक:',
           'सुरक्षित कंप्यूटर सेवा प्रबंधक 4.4',
           'डैशबोर्ड रीफ़्रेश विफल',
           'डैशबोर्ड अवलोकन',
           'डेटा यहाँ निर्यात किया गया है:',
           'डेटाबेस स्थान आवश्यक',
           'डेटाबेस पथ खाली नहीं हो सकता',
           'डेटाबेस पथ:',
           'डेटाबेस पथ सफलतापूर्वक अपडेट किए गए!',
           'डेटाबेस बैकअप से पुनर्स्थापित हुआ',
           'डेटाबेस बैकअप से सफलतापूर्वक पुनर्स्थापित हुआ!',
           'डेटाबेस_पुनर्स्थापित',
           'तिथि',
           'प्रिय [CUSTOMER],',
           'विवरण',
           'डिवाइस ब्रांड आँकड़े',
           'डिवाइस मॉडल',
           'डिवाइस प्रकार',
           'प्रदर्शन नाम',
           'दान:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'ड्राइव D: उपलब्ध नहीं है। कृपया कोई अन्य फ़ोल्डर चुनें।',
           'ड्राइव D: नहीं मिली।\nकृपया डेटाबेस के लिए फ़ोल्डर बनाएँ या चुनें।',
           'ड्राइव नहीं मिली',
           'डमी डेटा संपादक खुला',
           'डुप्लिकेट पार्ट',
           'डुप्लिकेट आपूर्तिकर्ता',
           'संपादन मोड',
           'संपादन मोड: फ़ील्ड अपडेट करें और सहेजें पर क्लिक करें',
           'संपादन मोड: फ़ील्ड अपडेट करें और अपडेट पर क्लिक करें',
           'सक्षम / अक्षम',
           'समाप्ति तिथि:',
           'उपयोगकर्ता नाम और पासवर्ड दोनों दर्ज करें।',
           'Excel फ़ाइलें',
           '30 दिनों में समाप्त',
           'जल्द समाप्त होने वाली',
           'जल्द समाप्त होने वाली (< 7 दिन)',
           'जल्द समाप्त होने वाली (≤ 30 दिन)',
           'जल्द समाप्त होने वाली (≤ 7 दिन)',
           'जल्द समाप्त होने वाली (≤30 दिन)',
           'समाप्ति तिथि से',
           'समाप्ति तिथि से:',
           'निर्यात पूर्ण',
           'निर्यात त्रुटि',
           'निर्यात विफल',
           'निर्यात सफल',
           'यहाँ निर्यात किया गया:',
           'निर्यात करने में विफल',
           'वित्तीय रिपोर्ट',
           'पहला प्रशासक सत्र',
           'पहली बार डेटाबेस सेटअप',
           'पहली बार सेटअप:\nनया डेटाबेस सहेजने के लिए फ़ोल्डर चुनें।',
           'फ़ॉर्म साफ़ किया गया',
           'प्रारंभ तिथि समाप्ति तिथि के बाद नहीं हो सकती',
           'पूरा पता',
           'Google मानचित्र',
           'Google मानचित्र लिंक कॉपी किया गया',
           'Google मानचित्र लिंक खाली है',
           'Google मानचित्र लिंक खाली है (क्लिपबोर्ड साफ़ किया गया)',
           'Google मानचित्र:',
           'सहायता सामग्री',
           'ID (ROWID)',
           'आयात पूर्ण',
           'आयात विफल',
           'आयात/निर्यात फ़ोल्डर खाली नहीं हो सकता',
           'आयात/निर्यात फ़ोल्डर:',
           'प्रगति पर',
           'मरम्मत में',
           'अमान्य क्रेडेंशियल',
           'अमान्य तिथि सीमा',
           'अमान्य ड्राइव',
           'अमान्य पथ',
           'अमान्य उपयोगकर्ता भूमिका।',
           'अमान्य उपयोगकर्ता नाम',
           'अमान्य उपयोगकर्ता नाम या पासवर्ड।',
           'अमान्य वारंटी तिथि सीमा: प्रारंभ तिथि समाप्ति तिथि के बाद है',
           'चालान संख्या',
           'समस्याएँ',
           'अंतिम लॉगिन',
           'पुरानी .XLS फ़ाइलें समर्थित नहीं हैं। पहले फ़ाइल को .XLSX या CSV के रूप में सहेजें।',
           'यहाँ तक लॉक',
           'रिमाइंडर लॉग करें',
           'रिमाइंडर लॉग करें',
           'वारंटी रिमाइंडर लॉग करें',
           'लॉगिन',
           'लॉगिन सफल।',
           'लॉगिन_अवरुद्ध',
           'लॉगिन_विफल',
           'लॉगिन_सफल',
           'लॉगआउट',
           'कम स्टॉक वाले पार्ट',
           'मुख्य विंडो खुली',
           'कम से कम 4 अक्षर। अक्षर, अंक, रिक्त स्थान या प्रतीकों का कोई भी प्रारूप अनुमत है।',
           'डिवाइस मॉडल',
           'मॉडल',
           'बहुभाषी डमी डेटा निर्माता खुला',
           'मेरी हाल की गतिविधि',
           'लागू नहीं',
           'ग्राहक का नाम',
           'नाम आवश्यक है',
           'नाम:',
           'नया पासवर्ड',
           'कोई डेटा नहीं',
           'कोई डेटाबेस चयनित नहीं',
           'कोई डिफ़ॉल्ट पासवर्ड उपयोग नहीं किया जाता। पासवर्ड केवल अद्वितीय, सॉल्टेड scrypt हैश के रूप में संग्रहीत '
           'होते हैं। इस प्रशासक पासवर्ड को सुरक्षित रखें।',
           'कोई मौजूदा डेटाबेस चयनित नहीं किया गया। एप्लिकेशन नया डेटाबेस बनाना जारी रखेगा।',
           'कोई SMS या WhatsApp संदेश नहीं भेजा गया।',
           'कोई स्थिति डेटा नहीं',
           'कोई वारंटी नहीं',
           'किसी भी प्रकार की कोई गारंटी प्रदान नहीं की जाती',
           'फ़ोन नंबर',
           'नहीं मिला',
           'मौजूदा डेटाबेस खोलें या इस एप्लिकेशन के लिए नया डेटाबेस बनाएँ।',
           'एप्लिकेशन सहायता सामग्री खोलें',
           'डमी डेटा संपादक खोलें',
           'Google मानचित्र खोलें',
           'सक्रिय टैब के लिए सहायता खोलें',
           'पार्ट का नाम',
           'पार्ट का नाम आवश्यक है',
           'पार्ट के प्रकार',
           'पार्ट रीफ़्रेश किए गए',
           'पासवर्ड',
           'पासवर्ड बदला गया',
           'पासवर्ड बदल दिया गया।',
           'पासवर्ड बहुत लंबा है।',
           'पासवर्ड मेल नहीं खाते',
           'पासवर्ड में कम से कम 4 अक्षर होने चाहिए।',
           'पासवर्ड नहीं बदला गया',
           'पासवर्ड रीसेट किया गया',
           'पासवर्ड_बदला_गया',
           'पासवर्ड_बदलना_विफल',
           'पासवर्ड_रीसेट',
           'फ़ोन नंबर',
           'फ़ोन:',
           'DUMMY_CREATOR.EXE या DUMMY_CREATOR.PY को एप्लिकेशन के पास रखें।',
           'यदि आपको आगे सेवा चाहिए तो कृपया हमसे संपर्क करें।',
           'कृपया XLSX या CSV फ़ाइल चुनें',
           'कृपया Excel या CSV फ़ाइल चुनें',
           'खरीद चालान:',
           'मात्रा',
           'हाल की सुरक्षा और डेटा गतिविधि',
           'हाल के सेवा रिकॉर्ड',
           'डैशबोर्ड रीफ़्रेश करें',
           'ड्रॉपडाउन डेटा पुनः लोड करें',
           'रिमाइंडर संदेश टेम्पलेट',
           'रिमाइंडर लॉग किए गए',
           'रिमाइंडर भेजे गए',
           'मरम्मत लागत',
           'रिपोर्ट निर्यात विफल',
           'रिपोर्ट निर्यात की गई',
           'पासवर्ड रीसेट करें',
           'उपयोगकर्ता पासवर्ड रीसेट करें',
           'आज से वारंटी सीमा रीसेट करें',
           'डेटाबेस पुनर्स्थापित करें',
           'पुनर्स्थापना विफल',
           'पुनर्स्थापना सफल',
           'भूमिका',
           'निर्यात फ़ाइल सहेजें',
           'रिपोर्ट निर्यात सहेजें',
           'नाम, फ़ोन, डिवाइस या समस्या से खोजें...',
           'सुरक्षित प्रशासक सेटअप',
           'सुरक्षित लॉगिन',
           'हटाने के लिए संपर्क चुनें',
           'संपादित करने के लिए संपर्क चुनें',
           'हटाने के लिए पार्ट चुनें',
           'संपादित करने के लिए पार्ट चुनें',
           'अपडेट करने के लिए पार्ट चुनें',
           'हटाने के लिए आपूर्तिकर्ता चुनें',
           'संपादित करने के लिए आपूर्तिकर्ता चुनें',
           'अपडेट करने के लिए आपूर्तिकर्ता चुनें',
           'पहले उपयोगकर्ता चुनें।',
           'संपादित करने के लिए वारंटी पंक्ति चुनें',
           'एप्लिकेशन मुद्रा चुनें:',
           'एप्लिकेशन भाषा और लोकेल चुनें:',
           'एप्लिकेशन के लिए डिफ़ॉल्ट मुद्रा चुनें:',
           'भाषा / लोकेल चुनें',
           'निर्यात करने के लिए पंक्तियाँ चुनें',
           'रिमाइंडर लॉग करने के लिए वारंटियाँ चुनें',
           'रिमाइंडर भेजने के लिए वारंटियाँ चुनें',
           'रिमाइंडर भेजें',
           'सेवा विवरण',
           'बोगोर ऑन-कॉल कंप्यूटर सेवा',
           'सेवा स्थिति ग्राफ़',
           'सेवा सारांश',
           'सेवा सारांश रिपोर्ट',
           'इसके लिए नया पासवर्ड सेट करें',
           'डेटाबेस स्थान सेट करें',
           'सेवा तिथि आज पर सेट करें',
           'सेटिंग्स सहेजी गईं',
           'सेटअप विफल',
           'पासवर्ड दिखाएँ',
           'पासवर्ड दिखाएँ',
           'जारी रखने के लिए साइन इन करें',
           'साइन इन किया गया',
           'इस रूप में साइन इन',
           'कर्मचारी',
           'कर्मचारी पहुँच सीमित है',
           'कर्मचारी नए सेवा रिकॉर्ड जोड़ सकते हैं और परिचालन डेटा देख सकते हैं। संपादन, हटाना, आयात, निर्यात, बैकअप, '
           'प्रशासनिक सेटिंग्स, उपयोगकर्ता प्रबंधन, आपूर्तिकर्ता परिवर्तन, पार्ट परिवर्तन, रिपोर्ट और रिमाइंडर के लिए '
           'प्रशासक आवश्यक है।',
           'स्थितियाँ',
           'भंडारण स्थान',
           'आपूर्तिकर्ता का पता',
           'आपूर्तिकर्ता का नाम',
           'आपूर्तिकर्ता का नाम आवश्यक है',
           'आपूर्तिकर्ता रीफ़्रेश किए गए',
           'सिस्टम',
           'तकनीशियन का नाम',
           'तकनीशियन प्रदर्शन',
           'तकनीशियन प्रदर्शन रिपोर्ट',
           'तकनीशियन',
           'धन्यवाद,',
           'यह उपयोगकर्ता नाम पहले से मौजूद है।',
           'बैकअप फ़ोल्डर की ड्राइव मौजूद नहीं है।',
           'डेटाबेस पथ की ड्राइव मौजूद नहीं है।',
           'निर्यात पथ की ड्राइव मौजूद नहीं है।',
           'आयात/निर्यात फ़ोल्डर की ड्राइव मौजूद नहीं है।',
           'पासवर्ड रीसेट किया गया।',
           'पासवर्ड मेल नहीं खाते।',
           'यह खाता अक्षम है। प्रशासक से संपर्क करें।',
           'यह कार्रवाई केवल प्रशासक खातों तक सीमित है।',
           'यह सॉफ़्टवेयर जैसा है वैसा ही प्रदान किया जाता है',
           'समय',
           'Excel में',
           'कुल सक्रिय वारंटियाँ',
           'कुल ग्राहक',
           'कुल डिवाइस',
           'कुल समाप्त',
           'कुल जल्द समाप्त होने वाली',
           'कुल आय',
           'कुल चयनित',
           'कुल सेवा रिकॉर्ड',
           'कुल सेवा मूल्य',
           'कुल सेवाएँ',
           'प्रकार:',
           'अद्वितीय मॉडल',
           'अज्ञात या अमान्य क्रेडेंशियल',
           'निर्यात के लिए अज्ञात तालिका।',
           'अनलॉक',
           'असमर्थित प्रारूप',
           'D:\\BACKUP उपयोग करें (डिफ़ॉल्ट)',
           'उपयोगकर्ता',
           'उपयोगकर्ता प्रबंधन',
           'उपयोगकर्ता नहीं बनाया गया',
           'उपयोगकर्ता नहीं मिला।',
           'उपयोगकर्ता ने लॉगआउट का अनुरोध किया',
           'उपयोगकर्ता स्थिति अपडेट की गई।',
           'उपयोगकर्ता नाम',
           'उपयोगकर्ता नाम मौजूद है',
           'उपयोगकर्ता नाम 3-32 अक्षरों का होना चाहिए और उसमें अक्षर, अंक, बिंदु, डैश या अंडरस्कोर का उपयोग किया जा '
           'सकता है।',
           'उपयोगकर्ता_बनाया_गया',
           'उपयोगकर्ता_अक्षम',
           'उपयोगकर्ता_सक्षम',
           'उपयोगकर्ता_अनलॉक',
           'पार्ट की प्रतीक्षा',
           'पार्ट की प्रतीक्षा',
           'चेतावनी: पुनर्स्थापना सभी वर्तमान डेटा को बैकअप डेटा से बदल देगी।\n'
           'यह कार्रवाई पूर्ववत नहीं की जा सकती।\n'
           '\n'
           'क्या आप जारी रखना चाहते हैं?',
           'वारंटी अवधियाँ',
           'वारंटी स्थिति:',
           'वारंटी सारांश',
           'वारंटी सारांश रिपोर्ट',
           'कमज़ोर पासवर्ड',
           'आप डेटाबेस के साथ क्या करना चाहते हैं?',
           'विंडो बंद',
           'आप अपना वर्तमान खाता अक्षम नहीं कर सकते।',
           'आपका पासवर्ड सफलतापूर्वक बदल दिया गया है।',
           '⚠️ चेतावनी!',
           '✏️ संपादित करें',
           '✏️ संपर्क संपादित करें',
           '❌ हटाएँ',
           '➕ जोड़ें',
           '➕ पार्ट जोड़ें',
           '📄 पाठ दृश्य',
           '📋 तालिका दृश्य',
           '📝 रिमाइंडर लॉग करें',
           '📝 रिमाइंडर लॉग करें',
           '📤 चयनित निर्यात करें',
           '🔃 रीफ़्रेश',
           '🔄 रीफ़्रेश',
           '🗑️ साफ़ करें'],
 'bn_BD': ['নতুন যন্ত্রাংশ একজন প্রশাসককে তৈরি করতে হবে।',
           'নতুন সরবরাহকারী একজন প্রশাসককে তৈরি করতে হবে।',
           'এই নামে একটি যন্ত্রাংশ আগে থেকেই আছে। সেটি নির্বাচন করে হালনাগাদ করতে সংরক্ষণ ব্যবহার করুন।',
           'এই নামে একটি সরবরাহকারী আগে থেকেই আছে। সেটি নির্বাচন করে হালনাগাদ করতে সংরক্ষণ ব্যবহার করুন।',
           'অ্যাক্সেস_প্রত্যাখ্যাত',
           'অ্যাকাউন্ট',
           'অ্যাকাউন্ট নিষ্ক্রিয়',
           'অ্যাকাউন্ট লক করা',
           'কাজটি অবরুদ্ধ',
           'সক্রিয় ওয়ারেন্টি',
           'অ্যাপ্লিকেশন ব্যবহারকারী যোগ করুন',
           'ব্যবহারকারী যোগ করুন',
           'অতিরিক্ত নোট',
           'সম্পূর্ণ ঠিকানা',
           'অ্যাডমিন',
           'প্রশাসক',
           'প্রশাসকের কাজ অবরুদ্ধ',
           'শুধু প্রশাসক',
           'প্রশাসক প্রয়োজন',
           'প্রশাসক প্রয়োজন।',
           'অন্য একটি যন্ত্রাংশ ইতিমধ্যে এই নাম ব্যবহার করছে।',
           'অন্য একটি সরবরাহকারী ইতিমধ্যে এই নাম ব্যবহার করছে।',
           'অ্যাপ্লিকেশন বন্ধ হয়েছে',
           'অ্যাপ্লিকেশন_প্রস্থান',
           'অ্যাপ্লিকেশন_খোলা',
           'অন্তত একজন সক্রিয় প্রশাসক প্রয়োজন।',
           'প্রতি সার্ভিসের গড়',
           'গড় মেরামত খরচ',
           'গড় আয়',
           'গড় আয়/সার্ভিস',
           'ডেটাবেস ব্যাকআপ করুন',
           'ব্যাকআপ ব্যর্থ',
           'ব্যাকআপ ফোল্ডার খালি রাখা যাবে না',
           'ব্যাকআপ ফোল্ডার:',
           'ব্যাকআপ সফল',
           'বোগোর অন-কল কম্পিউটার সার্ভিস আপনাকে স্মরণ করিয়ে দিচ্ছে যে আপনার ডিভাইসের মেরামত ওয়ারেন্টি শেষ হবে:',
           'ব্র্যান্ড:',
           'ব্র্যান্ডসমূহ',
           'HELP.HHP তৈরি করুন এবং HELP.CHM অ্যাপ্লিকেশন ফোল্ডারে কপি করুন।',
           'স্বয়ংক্রিয় হিসাব: সার্ভিসের তারিখ + ওয়ারেন্টির মেয়াদ',
           'নির্বাচিত সরবরাহকারী নির্ধারণ করা যায়নি',
           'পাসওয়ার্ড পরিবর্তন করুন',
           'এর জন্য পাসওয়ার্ড পরিবর্তন করুন',
           'সম্পন্ন',
           'কম্পিউটার সার্ভিস ম্যানেজার',
           'মুছে ফেলা নিশ্চিত করুন',
           'পাসওয়ার্ড নিশ্চিত করুন',
           'পুনরুদ্ধার নিশ্চিত করুন',
           'যোগাযোগের নাম',
           'যোগাযোগ:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'ক্লিপবোর্ডে কপি করা হয়েছে',
           'ঠিকানা কপি করুন',
           'Google মানচিত্র কপি করুন',
           'চালান কপি করুন',
           'ফোন কপি করুন',
           'কপিরাইট © ECOMTECH 2026',
           'সংখ্যা',
           'প্রশাসক তৈরি করুন',
           'প্রথম প্রশাসক অ্যাকাউন্ট তৈরি করুন',
           'তৈরি হয়েছে',
           'তৈরি করেছেন:',
           'তৈরির তারিখ',
           'CSV ফাইল',
           'বর্তমান পাসওয়ার্ড',
           'বর্তমান পাসওয়ার্ডটি ভুল।',
           'বর্তমান ট্যাবের সহায়তা',
           'গ্রাহকের নাম',
           'স্মরণিকার জন্য নির্বাচিত গ্রাহক',
           'স্মরণিকার জন্য নির্বাচিত গ্রাহক:',
           'নিরাপদ কম্পিউটার সার্ভিস ম্যানেজার 4.4',
           'ড্যাশবোর্ড রিফ্রেশ ব্যর্থ',
           'ড্যাশবোর্ড পর্যালোচনা',
           'ডেটা এখানে রপ্তানি করা হয়েছে:',
           'ডেটাবেসের অবস্থান প্রয়োজন',
           'ডেটাবেসের পথ খালি রাখা যাবে না',
           'ডেটাবেসের পথ:',
           'ডেটাবেসের পথ সফলভাবে হালনাগাদ হয়েছে!',
           'ব্যাকআপ থেকে ডেটাবেস পুনরুদ্ধার হয়েছে',
           'ব্যাকআপ থেকে ডেটাবেস সফলভাবে পুনরুদ্ধার হয়েছে!',
           'ডেটাবেস_পুনরুদ্ধার',
           'তারিখ',
           'প্রিয় [CUSTOMER],',
           'বিস্তারিত',
           'ডিভাইস ব্র্যান্ডের পরিসংখ্যান',
           'ডিভাইস মডেল',
           'ডিভাইসের ধরন',
           'প্রদর্শনের নাম',
           'অনুদান:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'ড্রাইভ D: উপলভ্য নয়। অন্য ফোল্ডার বেছে নিন।',
           'ড্রাইভ D: পাওয়া যায়নি।\nডেটাবেসের জন্য একটি ফোল্ডার তৈরি বা নির্বাচন করুন।',
           'ড্রাইভ পাওয়া যায়নি',
           'ডামি ডেটা সম্পাদক খোলা হয়েছে',
           'ডুপ্লিকেট যন্ত্রাংশ',
           'ডুপ্লিকেট সরবরাহকারী',
           'সম্পাদনা মোড',
           'সম্পাদনা মোড: ঘরগুলো হালনাগাদ করে সংরক্ষণে ক্লিক করুন',
           'সম্পাদনা মোড: ঘরগুলো হালনাগাদ করে হালনাগাদে ক্লিক করুন',
           'সক্রিয় / নিষ্ক্রিয়',
           'শেষ তারিখ:',
           'ব্যবহারকারীর নাম ও পাসওয়ার্ড উভয়ই লিখুন।',
           'Excel ফাইল',
           '৩০ দিনের মধ্যে মেয়াদ শেষ',
           'শিগগির মেয়াদ শেষ',
           'শিগগির মেয়াদ শেষ (< ৭ দিন)',
           'শিগগির মেয়াদ শেষ (≤ ৩০ দিন)',
           'শিগগির মেয়াদ শেষ (≤ ৭ দিন)',
           'শিগগির মেয়াদ শেষ (≤৩০ দিন)',
           'মেয়াদ শেষের তারিখ শুরু',
           'মেয়াদ শেষের তারিখ শুরু:',
           'রপ্তানি সম্পন্ন',
           'রপ্তানি ত্রুটি',
           'রপ্তানি ব্যর্থ',
           'রপ্তানি সফল',
           'এখানে রপ্তানি হয়েছে:',
           'রপ্তানি করা যায়নি',
           'আর্থিক প্রতিবেদন',
           'প্রথম প্রশাসক সেশন',
           'প্রথমবার ডেটাবেস সেটআপ',
           'প্রথমবার সেটআপ:\nনতুন ডেটাবেস সংরক্ষণের জন্য একটি ফোল্ডার বেছে নিন।',
           'ফর্ম পরিষ্কার করা হয়েছে',
           'শুরুর তারিখ শেষের তারিখের পরে হতে পারবে না',
           'সম্পূর্ণ ঠিকানা',
           'Google মানচিত্র',
           'Google মানচিত্রের লিংক কপি হয়েছে',
           'Google মানচিত্রের লিংক খালি',
           'Google মানচিত্রের লিংক খালি (ক্লিপবোর্ড পরিষ্কার করা হয়েছে)',
           'Google মানচিত্র:',
           'সহায়তা বিষয়বস্তু',
           'ID (ROWID)',
           'আমদানি সম্পন্ন',
           'আমদানি ব্যর্থ',
           'আমদানি/রপ্তানি ফোল্ডার খালি রাখা যাবে না',
           'আমদানি/রপ্তানি ফোল্ডার:',
           'চলমান',
           'মেরামত চলছে',
           'অবৈধ পরিচয়পত্র',
           'অবৈধ তারিখ সীমা',
           'অবৈধ ড্রাইভ',
           'অবৈধ পথ',
           'অবৈধ ব্যবহারকারী ভূমিকা।',
           'অবৈধ ব্যবহারকারীর নাম',
           'অবৈধ ব্যবহারকারীর নাম বা পাসওয়ার্ড।',
           'অবৈধ ওয়ারেন্টি তারিখ সীমা: শুরুর তারিখ শেষের তারিখের পরে',
           'চালান নম্বর',
           'সমস্যাসমূহ',
           'সর্বশেষ লগইন',
           'পুরোনো .XLS ফাইল সমর্থিত নয়। আগে ফাইলটি .XLSX বা CSV হিসেবে সংরক্ষণ করুন।',
           'যতক্ষণ পর্যন্ত লক',
           'স্মরণিকা লগ করুন',
           'স্মরণিকাগুলো লগ করুন',
           'ওয়ারেন্টি স্মরণিকা লগ করুন',
           'লগইন',
           'লগইন সফল।',
           'লগইন_অবরুদ্ধ',
           'লগইন_ব্যর্থ',
           'লগইন_সফল',
           'লগআউট',
           'কম মজুদের যন্ত্রাংশ',
           'মূল উইন্ডো খোলা হয়েছে',
           'কমপক্ষে ৪ অক্ষর। অক্ষর, সংখ্যা, ফাঁকা স্থান বা প্রতীকের যেকোনো বিন্যাস অনুমোদিত।',
           'ডিভাইস মডেল',
           'মডেলসমূহ',
           'বহুভাষিক ডামি ডেটা নির্মাতা খোলা হয়েছে',
           'আমার সাম্প্রতিক কার্যকলাপ',
           'প্রযোজ্য নয়',
           'গ্রাহকের নাম',
           'নাম আবশ্যক',
           'নাম:',
           'নতুন পাসওয়ার্ড',
           'কোনো ডেটা নেই',
           'কোনো ডেটাবেস নির্বাচিত নয়',
           'কোনো পূর্বনির্ধারিত পাসওয়ার্ড ব্যবহার করা হয় না। পাসওয়ার্ড শুধু স্বতন্ত্র, সল্টযুক্ত scrypt হ্যাশ '
           'হিসেবে সংরক্ষিত হয়। এই প্রশাসক পাসওয়ার্ড নিরাপদে রাখুন।',
           'কোনো বিদ্যমান ডেটাবেস নির্বাচন করা হয়নি। অ্যাপ্লিকেশন নতুন ডেটাবেস তৈরি করা চালিয়ে যাবে।',
           'কোনো SMS বা WhatsApp বার্তা পাঠানো হয়নি।',
           'কোনো অবস্থা ডেটা নেই',
           'কোনো ওয়ারেন্টি নেই',
           'কোনো ধরনের নিশ্চয়তা প্রদান করা হয় না',
           'ফোন নম্বর',
           'পাওয়া যায়নি',
           'একটি বিদ্যমান ডেটাবেস খুলুন অথবা এই অ্যাপ্লিকেশনের জন্য নতুন ডেটাবেস তৈরি করুন।',
           'অ্যাপ্লিকেশনের সহায়তা বিষয়বস্তু খুলুন',
           'ডামি ডেটা সম্পাদক খুলুন',
           'Google মানচিত্র খুলুন',
           'সক্রিয় ট্যাবের সহায়তা খুলুন',
           'যন্ত্রাংশের নাম',
           'যন্ত্রাংশের নাম আবশ্যক',
           'যন্ত্রাংশের ধরন',
           'যন্ত্রাংশ রিফ্রেশ হয়েছে',
           'পাসওয়ার্ড',
           'পাসওয়ার্ড পরিবর্তিত',
           'পাসওয়ার্ড পরিবর্তিত হয়েছে।',
           'পাসওয়ার্ড অনেক দীর্ঘ।',
           'পাসওয়ার্ড মিলছে না',
           'পাসওয়ার্ডে কমপক্ষে ৪ অক্ষর থাকতে হবে।',
           'পাসওয়ার্ড পরিবর্তিত হয়নি',
           'পাসওয়ার্ড রিসেট হয়েছে',
           'পাসওয়ার্ড_পরিবর্তিত',
           'পাসওয়ার্ড_পরিবর্তন_ব্যর্থ',
           'পাসওয়ার্ড_রিসেট',
           'ফোন নম্বর',
           'ফোন:',
           'DUMMY_CREATOR.EXE অথবা DUMMY_CREATOR.PY অ্যাপ্লিকেশনের পাশে রাখুন।',
           'আরও সেবার প্রয়োজন হলে আমাদের সঙ্গে যোগাযোগ করুন।',
           'একটি XLSX বা CSV ফাইল নির্বাচন করুন',
           'একটি Excel বা CSV ফাইল নির্বাচন করুন',
           'ক্রয় চালান:',
           'পরিমাণ',
           'সাম্প্রতিক নিরাপত্তা ও ডেটা কার্যকলাপ',
           'সাম্প্রতিক সার্ভিস রেকর্ড',
           'ড্যাশবোর্ড রিফ্রেশ করুন',
           'ড্রপডাউন ডেটা পুনরায় লোড করুন',
           'স্মরণিকা বার্তার টেমপ্লেট',
           'স্মরণিকা লগ করা হয়েছে',
           'স্মরণিকা পাঠানো হয়েছে',
           'মেরামত খরচ',
           'প্রতিবেদন রপ্তানি ব্যর্থ',
           'প্রতিবেদন রপ্তানি হয়েছে',
           'পাসওয়ার্ড রিসেট করুন',
           'ব্যবহারকারীর পাসওয়ার্ড রিসেট করুন',
           'আজ থেকে ওয়ারেন্টির সীমা রিসেট করুন',
           'ডেটাবেস পুনরুদ্ধার করুন',
           'পুনরুদ্ধার ব্যর্থ',
           'পুনরুদ্ধার সফল',
           'ভূমিকা',
           'রপ্তানি ফাইল সংরক্ষণ করুন',
           'প্রতিবেদন রপ্তানি সংরক্ষণ করুন',
           'নাম, ফোন, ডিভাইস বা সমস্যা দিয়ে খুঁজুন...',
           'নিরাপদ প্রশাসক সেটআপ',
           'নিরাপদ লগইন',
           'মুছতে একটি যোগাযোগ নির্বাচন করুন',
           'সম্পাদনা করতে একটি যোগাযোগ নির্বাচন করুন',
           'মুছতে একটি যন্ত্রাংশ নির্বাচন করুন',
           'সম্পাদনা করতে একটি যন্ত্রাংশ নির্বাচন করুন',
           'হালনাগাদ করতে একটি যন্ত্রাংশ নির্বাচন করুন',
           'মুছতে একটি সরবরাহকারী নির্বাচন করুন',
           'সম্পাদনা করতে একটি সরবরাহকারী নির্বাচন করুন',
           'হালনাগাদ করতে একটি সরবরাহকারী নির্বাচন করুন',
           'প্রথমে একজন ব্যবহারকারী নির্বাচন করুন।',
           'সম্পাদনা করতে একটি ওয়ারেন্টি সারি নির্বাচন করুন',
           'অ্যাপ্লিকেশনের মুদ্রা নির্বাচন করুন:',
           'অ্যাপ্লিকেশনের ভাষা ও লোকেল নির্বাচন করুন:',
           'অ্যাপ্লিকেশনের পূর্বনির্ধারিত মুদ্রা নির্বাচন করুন:',
           'ভাষা / লোকেল নির্বাচন করুন',
           'রপ্তানি করার সারি নির্বাচন করুন',
           'স্মরণিকা লগ করার ওয়ারেন্টি নির্বাচন করুন',
           'স্মরণিকা পাঠানোর ওয়ারেন্টি নির্বাচন করুন',
           'স্মরণিকা পাঠান',
           'সার্ভিসের বিবরণ',
           'বোগোর অন-কল কম্পিউটার সার্ভিস',
           'সার্ভিস অবস্থার গ্রাফ',
           'সার্ভিস সারাংশ',
           'সার্ভিস সারাংশ প্রতিবেদন',
           'এর জন্য নতুন পাসওয়ার্ড নির্ধারণ করুন',
           'ডেটাবেসের অবস্থান নির্ধারণ করুন',
           'সার্ভিসের তারিখ আজ নির্ধারণ করুন',
           'সেটিংস সংরক্ষিত হয়েছে',
           'সেটআপ ব্যর্থ',
           'পাসওয়ার্ড দেখান',
           'পাসওয়ার্ডগুলো দেখান',
           'চালিয়ে যেতে সাইন ইন করুন',
           'সাইন ইন করা হয়েছে',
           'যে নামে সাইন ইন',
           'কর্মী',
           'কর্মীদের প্রবেশাধিকার সীমিত',
           'কর্মীরা নতুন সার্ভিস রেকর্ড যোগ করতে এবং পরিচালনাগত ডেটা দেখতে পারেন। সম্পাদনা, মুছে ফেলা, আমদানি, '
           'রপ্তানি, ব্যাকআপ, প্রশাসনিক সেটিংস, ব্যবহারকারী ব্যবস্থাপনা, সরবরাহকারী পরিবর্তন, যন্ত্রাংশ পরিবর্তন, '
           'প্রতিবেদন ও স্মরণিকার জন্য প্রশাসক প্রয়োজন।',
           'অবস্থাসমূহ',
           'সংরক্ষণের অবস্থান',
           'সরবরাহকারীর ঠিকানা',
           'সরবরাহকারীর নাম',
           'সরবরাহকারীর নাম আবশ্যক',
           'সরবরাহকারীরা রিফ্রেশ হয়েছে',
           'সিস্টেম',
           'প্রযুক্তিবিদের নাম',
           'প্রযুক্তিবিদের কর্মদক্ষতা',
           'প্রযুক্তিবিদের কর্মদক্ষতা প্রতিবেদন',
           'প্রযুক্তিবিদরা',
           'ধন্যবাদ,',
           'এই ব্যবহারকারীর নাম আগে থেকেই আছে।',
           'ব্যাকআপ ফোল্ডারের ড্রাইভটি নেই।',
           'ডেটাবেস পথের ড্রাইভটি নেই।',
           'রপ্তানি পথের ড্রাইভটি নেই।',
           'আমদানি/রপ্তানি ফোল্ডারের ড্রাইভটি নেই।',
           'পাসওয়ার্ড রিসেট করা হয়েছে।',
           'পাসওয়ার্ড মিলছে না।',
           'এই অ্যাকাউন্ট নিষ্ক্রিয়। একজন প্রশাসকের সঙ্গে যোগাযোগ করুন।',
           'এই কাজটি শুধু প্রশাসক অ্যাকাউন্টের জন্য।',
           'এই সফটওয়্যারটি যেমন আছে তেমনই প্রদান করা হয়',
           'সময়',
           'Excel-এ',
           'মোট সক্রিয় ওয়ারেন্টি',
           'মোট গ্রাহক',
           'মোট ডিভাইস',
           'মোট মেয়াদোত্তীর্ণ',
           'মোট শিগগির মেয়াদ শেষ',
           'মোট আয়',
           'মোট নির্বাচিত',
           'মোট সার্ভিস রেকর্ড',
           'মোট সার্ভিস মূল্য',
           'মোট সার্ভিস',
           'ধরন:',
           'অনন্য মডেল',
           'অজানা বা অবৈধ পরিচয়পত্র',
           'রপ্তানির জন্য অজানা টেবিল।',
           'আনলক',
           'অসমর্থিত বিন্যাস',
           'D:\\BACKUP ব্যবহার করুন (পূর্বনির্ধারিত)',
           'ব্যবহারকারী',
           'ব্যবহারকারী ব্যবস্থাপনা',
           'ব্যবহারকারী তৈরি হয়নি',
           'ব্যবহারকারী পাওয়া যায়নি।',
           'ব্যবহারকারী লগআউটের অনুরোধ করেছে',
           'ব্যবহারকারীর অবস্থা হালনাগাদ হয়েছে।',
           'ব্যবহারকারীর নাম',
           'ব্যবহারকারীর নাম বিদ্যমান',
           'ব্যবহারকারীর নাম ৩-৩২ অক্ষরের হতে হবে এবং অক্ষর, সংখ্যা, বিন্দু, ড্যাশ বা আন্ডারস্কোর ব্যবহার করা যাবে।',
           'ব্যবহারকারী_তৈরি',
           'ব্যবহারকারী_নিষ্ক্রিয়',
           'ব্যবহারকারী_সক্রিয়',
           'ব্যবহারকারী_আনলক',
           'যন্ত্রাংশের অপেক্ষায়',
           'যন্ত্রাংশের অপেক্ষায়',
           'সতর্কতা: পুনরুদ্ধার করলে বর্তমান সব ডেটা ব্যাকআপ ডেটা দিয়ে প্রতিস্থাপিত হবে।\n'
           'এই কাজটি পূর্বাবস্থায় ফেরানো যাবে না।\n'
           '\n'
           'আপনি কি চালিয়ে যেতে চান?',
           'ওয়ারেন্টির মেয়াদসমূহ',
           'ওয়ারেন্টির অবস্থা:',
           'ওয়ারেন্টি সারাংশ',
           'ওয়ারেন্টি সারাংশ প্রতিবেদন',
           'দুর্বল পাসওয়ার্ড',
           'আপনি ডেটাবেসের সঙ্গে কী করতে চান?',
           'উইন্ডো বন্ধ হয়েছে',
           'আপনি আপনার বর্তমান অ্যাকাউন্ট নিষ্ক্রিয় করতে পারবেন না।',
           'আপনার পাসওয়ার্ড সফলভাবে পরিবর্তন করা হয়েছে।',
           '⚠️ সতর্কতা!',
           '✏️ সম্পাদনা',
           '✏️ যোগাযোগ সম্পাদনা',
           '❌ মুছুন',
           '➕ যোগ করুন',
           '➕ যন্ত্রাংশ যোগ করুন',
           '📄 পাঠ্য দৃশ্য',
           '📋 টেবিল দৃশ্য',
           '📝 স্মরণিকা লগ করুন',
           '📝 স্মরণিকাগুলো লগ করুন',
           '📤 নির্বাচিতগুলো রপ্তানি করুন',
           '🔃 রিফ্রেশ',
           '🔄 রিফ্রেশ',
           '🗑️ পরিষ্কার করুন'],
 'ru_RU': ['Новая деталь должна быть создана администратором.',
           'Новый поставщик должен быть создан администратором.',
           'Деталь с таким названием уже существует. Выберите её и нажмите «Сохранить», чтобы обновить.',
           'Поставщик с таким названием уже существует. Выберите его и нажмите «Сохранить», чтобы обновить.',
           'ДОСТУП_ЗАПРЕЩЁН',
           'УЧЁТНАЯ ЗАПИСЬ',
           'УЧЁТНАЯ ЗАПИСЬ ОТКЛЮЧЕНА',
           'УЧЁТНАЯ ЗАПИСЬ ЗАБЛОКИРОВАНА',
           'ДЕЙСТВИЕ ЗАБЛОКИРОВАНО',
           'ДЕЙСТВУЮЩИЕ ГАРАНТИИ',
           'ДОБАВИТЬ ПОЛЬЗОВАТЕЛЯ ПРИЛОЖЕНИЯ',
           'ДОБАВИТЬ ПОЛЬЗОВАТЕЛЯ',
           'ДОПОЛНИТЕЛЬНЫЕ ПРИМЕЧАНИЯ',
           'ПОЛНЫЙ АДРЕС',
           'АДМИН',
           'АДМИНИСТРАТОР',
           'ДЕЙСТВИЕ АДМИНИСТРАТОРА ЗАБЛОКИРОВАНО',
           'ТОЛЬКО ДЛЯ АДМИНИСТРАТОРА',
           'ТРЕБУЕТСЯ АДМИНИСТРАТОР',
           'ТРЕБУЕТСЯ АДМИНИСТРАТОР.',
           'Другая деталь уже использует это название.',
           'Другой поставщик уже использует это название.',
           'ПРИЛОЖЕНИЕ ЗАКРЫТО',
           'ВЫХОД_ИЗ_ПРИЛОЖЕНИЯ',
           'ПРИЛОЖЕНИЕ_ОТКРЫТО',
           'Требуется хотя бы один активный администратор.',
           'СРЕДНЕЕ НА ОДНО ОБСЛУЖИВАНИЕ',
           'СРЕДНЯЯ СТОИМОСТЬ РЕМОНТА',
           'СРЕДНЯЯ ВЫРУЧКА',
           'СРЕДНЯЯ ВЫРУЧКА/ОБСЛУЖИВАНИЕ',
           'РЕЗЕРВНОЕ КОПИРОВАНИЕ БАЗЫ ДАННЫХ',
           'ОШИБКА РЕЗЕРВНОГО КОПИРОВАНИЯ',
           'Папка резервного копирования не может быть пустой',
           'ПАПКА РЕЗЕРВНОГО КОПИРОВАНИЯ:',
           'РЕЗЕРВНОЕ КОПИРОВАНИЕ ВЫПОЛНЕНО',
           'Выездная компьютерная служба Богор напоминает, что гарантия на ремонт вашего устройства истекает:',
           'БРЕНД:',
           'БРЕНДЫ',
           'Соберите HELP.HHP и скопируйте HELP.CHM в папку приложения.',
           'Рассчитывается автоматически: дата обслуживания + срок гарантии',
           'Не удалось определить выбранного поставщика',
           'ИЗМЕНИТЬ ПАРОЛЬ',
           'ИЗМЕНИТЬ ПАРОЛЬ ДЛЯ',
           'ЗАВЕРШЕНО',
           'МЕНЕДЖЕР КОМПЬЮТЕРНОГО СЕРВИСА',
           'ПОДТВЕРДИТЬ УДАЛЕНИЕ',
           'ПОДТВЕРДИТЬ ПАРОЛЬ',
           'ПОДТВЕРДИТЬ ВОССТАНОВЛЕНИЕ',
           'ИМЯ КОНТАКТА',
           'КОНТАКТ:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'СКОПИРОВАНО В БУФЕР ОБМЕНА',
           'КОПИРОВАТЬ АДРЕС',
           'КОПИРОВАТЬ GOOGLE КАРТЫ',
           'КОПИРОВАТЬ СЧЁТ',
           'КОПИРОВАТЬ ТЕЛЕФОН',
           'АВТОРСКИЕ ПРАВА © ECOMTECH 2026',
           'КОЛИЧЕСТВО',
           'СОЗДАТЬ АДМИНИСТРАТОРА',
           'СОЗДАТЬ ПЕРВУЮ УЧЁТНУЮ ЗАПИСЬ АДМИНИСТРАТОРА',
           'СОЗДАНО',
           'СОЗДАЛ:',
           'ДАТА СОЗДАНИЯ',
           'ФАЙЛЫ CSV',
           'ТЕКУЩИЙ ПАРОЛЬ',
           'Текущий пароль неверен.',
           'СПРАВКА ТЕКУЩЕЙ ВКЛАДКИ',
           'ИМЯ КЛИЕНТА',
           'КЛИЕНТЫ, ВЫБРАННЫЕ ДЛЯ НАПОМИНАНИЯ',
           'КЛИЕНТЫ, ВЫБРАННЫЕ ДЛЯ НАПОМИНАНИЯ:',
           'Защищённый менеджер компьютерного сервиса 4.4',
           'НЕ УДАЛОСЬ ОБНОВИТЬ ПАНЕЛЬ',
           'ОБЗОР ПАНЕЛИ',
           'Данные экспортированы в:',
           'ТРЕБУЕТСЯ РАСПОЛОЖЕНИЕ БАЗЫ ДАННЫХ',
           'Путь к базе данных не может быть пустым',
           'ПУТЬ К БАЗЕ ДАННЫХ:',
           'Пути к базе данных успешно обновлены!',
           'БАЗА ДАННЫХ ВОССТАНОВЛЕНА ИЗ РЕЗЕРВНОЙ КОПИИ',
           'База данных успешно восстановлена из резервной копии!',
           'БАЗА_ДАННЫХ_ВОССТАНОВЛЕНА',
           'ДАТА',
           'Уважаемый(ая) [CUSTOMER],',
           'ПОДРОБНОСТИ',
           'СТАТИСТИКА БРЕНДОВ УСТРОЙСТВ',
           'МОДЕЛЬ УСТРОЙСТВА',
           'ТИПЫ УСТРОЙСТВ',
           'ОТОБРАЖАЕМОЕ ИМЯ',
           'ПОЖЕРТВОВАНИЕ:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Диск D: недоступен. Выберите другую папку.',
           'Диск D: не найден.\nСоздайте или выберите папку для базы данных.',
           'ДИСК НЕ НАЙДЕН',
           'РЕДАКТОР ТЕСТОВЫХ ДАННЫХ ОТКРЫТ',
           'ДУБЛИКАТ ДЕТАЛИ',
           'ДУБЛИКАТ ПОСТАВЩИКА',
           'РЕЖИМ РЕДАКТИРОВАНИЯ',
           'РЕЖИМ РЕДАКТИРОВАНИЯ: ОБНОВИТЕ ПОЛЯ И НАЖМИТЕ «СОХРАНИТЬ»',
           'РЕЖИМ РЕДАКТИРОВАНИЯ: ОБНОВИТЕ ПОЛЯ И НАЖМИТЕ «ОБНОВИТЬ»',
           'ВКЛЮЧИТЬ / ОТКЛЮЧИТЬ',
           'ДАТА ОКОНЧАНИЯ:',
           'Введите имя пользователя и пароль.',
           'ФАЙЛЫ EXCEL',
           'ИСТЕКАЕТ ЧЕРЕЗ 30 ДНЕЙ',
           'СКОРО ИСТЕКАЕТ',
           'СКОРО ИСТЕКАЕТ (< 7 ДНЕЙ)',
           'СКОРО ИСТЕКАЕТ (≤ 30 ДНЕЙ)',
           'СКОРО ИСТЕКАЕТ (≤ 7 ДНЕЙ)',
           'СКОРО ИСТЕКАЕТ (≤30 дней)',
           'ДАТА ИСТЕЧЕНИЯ С',
           'ДАТА ИСТЕЧЕНИЯ С:',
           'ЭКСПОРТ ЗАВЕРШЁН',
           'ОШИБКА ЭКСПОРТА',
           'ЭКСПОРТ НЕ ВЫПОЛНЕН',
           'ЭКСПОРТ ВЫПОЛНЕН',
           'ЭКСПОРТИРОВАНО В:',
           'Не удалось выполнить экспорт',
           'ФИНАНСОВЫЙ ОТЧЁТ',
           'ПЕРВЫЙ СЕАНС АДМИНИСТРАТОРА',
           'ПЕРВОНАЧАЛЬНАЯ НАСТРОЙКА БАЗЫ ДАННЫХ',
           'ПЕРВОНАЧАЛЬНАЯ НАСТРОЙКА:\nВыберите папку для сохранения новой базы данных.',
           'ФОРМА ОЧИЩЕНА',
           'Начальная дата не может быть позже конечной',
           'ПОЛНЫЙ АДРЕС',
           'GOOGLE КАРТЫ',
           'ССЫЛКА GOOGLE КАРТ СКОПИРОВАНА',
           'ССЫЛКА GOOGLE КАРТ ПУСТА',
           'ССЫЛКА GOOGLE КАРТ ПУСТА (БУФЕР ОБМЕНА ОЧИЩЕН)',
           'GOOGLE КАРТЫ:',
           'СОДЕРЖАНИЕ СПРАВКИ',
           'ID (ROWID)',
           'ИМПОРТ ЗАВЕРШЁН',
           'ИМПОРТ НЕ ВЫПОЛНЕН',
           'Папка импорта/экспорта не может быть пустой',
           'ПАПКА ИМПОРТА/ЭКСПОРТА:',
           'В ПРОЦЕССЕ',
           'В РЕМОНТЕ',
           'НЕДОПУСТИМЫЕ УЧЁТНЫЕ ДАННЫЕ',
           'НЕДОПУСТИМЫЙ ДИАПАЗОН ДАТ',
           'НЕДОПУСТИМЫЙ ДИСК',
           'НЕДОПУСТИМЫЙ ПУТЬ',
           'Недопустимая роль пользователя.',
           'НЕДОПУСТИМОЕ ИМЯ ПОЛЬЗОВАТЕЛЯ',
           'Неверное имя пользователя или пароль.',
           'Недопустимый диапазон дат гарантии: начальная дата позже конечной',
           'НОМЕР СЧЁТА',
           'НЕИСПРАВНОСТИ',
           'ПОСЛЕДНИЙ ВХОД',
           'Устаревшие файлы .XLS не поддерживаются. Сначала сохраните файл как .XLSX или CSV.',
           'ЗАБЛОКИРОВАНО ДО',
           'ЗАПИСАТЬ НАПОМИНАНИЕ',
           'ЗАПИСАТЬ НАПОМИНАНИЯ',
           'ЗАПИСАТЬ НАПОМИНАНИЕ О ГАРАНТИИ',
           'ВОЙТИ',
           'Вход выполнен успешно.',
           'ВХОД_ЗАБЛОКИРОВАН',
           'ОШИБКА_ВХОДА',
           'ВХОД_ВЫПОЛНЕН',
           'ВЫЙТИ',
           'ДЕТАЛИ С НИЗКИМ ОСТАТКОМ',
           'ГЛАВНОЕ ОКНО ОТКРЫТО',
           'Минимум 4 символа. Допускаются любые буквы, цифры, пробелы и символы.',
           'МОДЕЛЬ УСТРОЙСТВА',
           'МОДЕЛИ',
           'МНОГОЯЗЫЧНЫЙ ГЕНЕРАТОР ТЕСТОВЫХ ДАННЫХ ОТКРЫТ',
           'МОИ НЕДАВНИЕ ДЕЙСТВИЯ',
           'Н/Д',
           'ИМЯ КЛИЕНТА',
           'ИМЯ ОБЯЗАТЕЛЬНО',
           'ИМЯ:',
           'НОВЫЙ ПАРОЛЬ',
           'НЕТ ДАННЫХ',
           'БАЗА ДАННЫХ НЕ ВЫБРАНА',
           'Пароль по умолчанию не используется. Пароли хранятся только как уникальные хеши scrypt с солью. Храните '
           'этот пароль администратора в безопасности.',
           'Существующая база данных не выбрана. Приложение продолжит создание новой базы данных.',
           'SMS- или WhatsApp-сообщение не отправлялось.',
           'НЕТ ДАННЫХ О СТАТУСЕ',
           'БЕЗ ГАРАНТИИ',
           'Никакие гарантии не предоставляются',
           'НОМЕР ТЕЛЕФОНА',
           'НЕ НАЙДЕНО',
           'Откройте существующую базу данных или создайте новую базу данных для этого приложения.',
           'ОТКРЫТЬ СОДЕРЖАНИЕ СПРАВКИ ПРИЛОЖЕНИЯ',
           'ОТКРЫТЬ РЕДАКТОР ТЕСТОВЫХ ДАННЫХ',
           'ОТКРЫТЬ GOOGLE КАРТЫ',
           'ОТКРЫТЬ СПРАВКУ ДЛЯ АКТИВНОЙ ВКЛАДКИ',
           'НАЗВАНИЕ ДЕТАЛИ',
           'НАЗВАНИЕ ДЕТАЛИ ОБЯЗАТЕЛЬНО',
           'ТИПЫ ДЕТАЛЕЙ',
           'ДЕТАЛИ ОБНОВЛЕНЫ',
           'ПАРОЛЬ',
           'ПАРОЛЬ ИЗМЕНЁН',
           'Пароль изменён.',
           'Пароль слишком длинный.',
           'ПАРОЛИ НЕ СОВПАДАЮТ',
           'Пароль должен содержать не менее 4 символов.',
           'ПАРОЛЬ НЕ ИЗМЕНЁН',
           'ПАРОЛЬ СБРОШЕН',
           'ПАРОЛЬ_ИЗМЕНЁН',
           'ОШИБКА_ИЗМЕНЕНИЯ_ПАРОЛЯ',
           'ПАРОЛЬ_СБРОШЕН',
           'НОМЕР ТЕЛЕФОНА',
           'ТЕЛЕФОН:',
           'Поместите DUMMY_CREATOR.EXE или DUMMY_CREATOR.PY рядом с приложением.',
           'Свяжитесь с нами, если вам потребуется дополнительное обслуживание.',
           'Выберите файл XLSX или CSV',
           'Выберите файл Excel или CSV',
           'СЧЁТ НА ПОКУПКУ:',
           'КОЛ-ВО',
           'НЕДАВНИЕ СОБЫТИЯ БЕЗОПАСНОСТИ И ДАННЫХ',
           'НЕДАВНИЕ ЗАПИСИ ОБ ОБСЛУЖИВАНИИ',
           'ОБНОВИТЬ ПАНЕЛЬ',
           'ПЕРЕЗАГРУЗИТЬ ДАННЫЕ СПИСКОВ',
           'ШАБЛОН СООБЩЕНИЯ-НАПОМИНАНИЯ',
           'НАПОМИНАНИЯ ЗАПИСАНЫ',
           'НАПОМИНАНИЯ ОТПРАВЛЕНЫ',
           'СТОИМОСТЬ РЕМОНТА',
           'НЕ УДАЛОСЬ ЭКСПОРТИРОВАТЬ ОТЧЁТ',
           'ОТЧЁТ ЭКСПОРТИРОВАН',
           'СБРОСИТЬ ПАРОЛЬ',
           'СБРОСИТЬ ПАРОЛЬ ПОЛЬЗОВАТЕЛЯ',
           'СБРОСИТЬ ДИАПАЗОН ГАРАНТИИ ОТ СЕГОДНЯШНЕГО ДНЯ',
           'ВОССТАНОВИТЬ БАЗУ ДАННЫХ',
           'ВОССТАНОВЛЕНИЕ НЕ ВЫПОЛНЕНО',
           'ВОССТАНОВЛЕНИЕ ВЫПОЛНЕНО',
           'РОЛЬ',
           'СОХРАНИТЬ ФАЙЛ ЭКСПОРТА',
           'СОХРАНИТЬ ЭКСПОРТ ОТЧЁТА',
           'ПОИСК ПО ИМЕНИ, ТЕЛЕФОНУ, УСТРОЙСТВУ ИЛИ НЕИСПРАВНОСТИ...',
           'БЕЗОПАСНАЯ НАСТРОЙКА АДМИНИСТРАТОРА',
           'БЕЗОПАСНЫЙ ВХОД',
           'Выберите контакт для удаления',
           'Выберите контакт для редактирования',
           'Выберите деталь для удаления',
           'Выберите деталь для редактирования',
           'Выберите деталь для обновления',
           'Выберите поставщика для удаления',
           'Выберите поставщика для редактирования',
           'Выберите поставщика для обновления',
           'Сначала выберите пользователя.',
           'Выберите строку гарантии для редактирования',
           'ВЫБЕРИТЕ ВАЛЮТУ ПРИЛОЖЕНИЯ:',
           'ВЫБЕРИТЕ ЯЗЫК И ЛОКАЛЬ ПРИЛОЖЕНИЯ:',
           'ВЫБЕРИТЕ ВАЛЮТУ ПРИЛОЖЕНИЯ ПО УМОЛЧАНИЮ:',
           'ВЫБРАТЬ ЯЗЫК / ЛОКАЛЬ',
           'ВЫБРАТЬ СТРОКИ ДЛЯ ЭКСПОРТА',
           'ВЫБРАТЬ ГАРАНТИИ ДЛЯ ЗАПИСИ НАПОМИНАНИЙ',
           'ВЫБРАТЬ ГАРАНТИИ ДЛЯ ОТПРАВКИ НАПОМИНАНИЙ',
           'ОТПРАВИТЬ НАПОМИНАНИЯ',
           'ОПИСАНИЕ ОБСЛУЖИВАНИЯ',
           'ВЫЕЗДНАЯ КОМПЬЮТЕРНАЯ СЛУЖБА БОГОР',
           'ГРАФИК СТАТУСА ОБСЛУЖИВАНИЯ',
           'СВОДКА ОБСЛУЖИВАНИЯ',
           'ОТЧЁТ СО СВОДКОЙ ОБСЛУЖИВАНИЯ',
           'Установить новый пароль для',
           'ЗАДАТЬ РАСПОЛОЖЕНИЕ БАЗЫ ДАННЫХ',
           'УСТАНОВИТЬ ДАТУ ОБСЛУЖИВАНИЯ НА СЕГОДНЯ',
           'НАСТРОЙКИ СОХРАНЕНЫ',
           'ОШИБКА НАСТРОЙКИ',
           'ПОКАЗАТЬ ПАРОЛЬ',
           'ПОКАЗАТЬ ПАРОЛИ',
           'ВОЙДИТЕ, ЧТОБЫ ПРОДОЛЖИТЬ',
           'ВХОД ВЫПОЛНЕН',
           'ВХОД ВЫПОЛНЕН КАК',
           'СОТРУДНИК',
           'ДОСТУП СОТРУДНИКА ОГРАНИЧЕН',
           'Сотрудники могут добавлять новые записи об обслуживании и просматривать рабочие данные. Редактирование, '
           'удаление, импорт, экспорт, резервное копирование, административные настройки, управление пользователями, '
           'изменения поставщиков, изменения деталей, отчёты и напоминания требуют прав администратора.',
           'СТАТУСЫ',
           'МЕСТО ХРАНЕНИЯ',
           'АДРЕС ПОСТАВЩИКА',
           'НАЗВАНИЕ ПОСТАВЩИКА',
           'НАЗВАНИЕ ПОСТАВЩИКА ОБЯЗАТЕЛЬНО',
           'ПОСТАВЩИКИ ОБНОВЛЕНЫ',
           'СИСТЕМА',
           'ИМЯ ТЕХНИКА',
           'ЭФФЕКТИВНОСТЬ ТЕХНИКОВ',
           'ОТЧЁТ ОБ ЭФФЕКТИВНОСТИ ТЕХНИКОВ',
           'ТЕХНИКИ',
           'Спасибо,',
           'Это имя пользователя уже существует.',
           'Диск для папки резервного копирования не существует.',
           'Диск для пути к базе данных не существует.',
           'Диск для пути экспорта не существует.',
           'Диск для папки импорта/экспорта не существует.',
           'Пароль был сброшен.',
           'Пароли не совпадают.',
           'Эта учётная запись отключена. Обратитесь к администратору.',
           'Это действие доступно только учётным записям администраторов.',
           'Это программное обеспечение предоставляется «как есть»',
           'ВРЕМЯ',
           'В EXCEL',
           'ВСЕГО ДЕЙСТВУЮЩИХ ГАРАНТИЙ',
           'ВСЕГО КЛИЕНТОВ',
           'ВСЕГО УСТРОЙСТВ',
           'ВСЕГО ИСТЁКШИХ',
           'ВСЕГО СКОРО ИСТЕКАЮЩИХ',
           'ОБЩАЯ ВЫРУЧКА',
           'ВСЕГО ВЫБРАНО',
           'ВСЕГО ЗАПИСЕЙ ОБ ОБСЛУЖИВАНИИ',
           'ОБЩАЯ СТОИМОСТЬ ОБСЛУЖИВАНИЯ',
           'ВСЕГО ОБСЛУЖИВАНИЙ',
           'ТИП:',
           'УНИКАЛЬНЫЕ МОДЕЛИ',
           'НЕИЗВЕСТНЫЕ ИЛИ НЕДОПУСТИМЫЕ УЧЁТНЫЕ ДАННЫЕ',
           'Неизвестная таблица для экспорта.',
           'РАЗБЛОКИРОВАТЬ',
           'НЕПОДДЕРЖИВАЕМЫЙ ФОРМАТ',
           'ИСПОЛЬЗОВАТЬ D:\\BACKUP (ПО УМОЛЧАНИЮ)',
           'ПОЛЬЗОВАТЕЛЬ',
           'УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ',
           'ПОЛЬЗОВАТЕЛЬ НЕ СОЗДАН',
           'Пользователь не найден.',
           'ПОЛЬЗОВАТЕЛЬ ЗАПРОСИЛ ВЫХОД',
           'Статус пользователя обновлён.',
           'ИМЯ ПОЛЬЗОВАТЕЛЯ',
           'ИМЯ ПОЛЬЗОВАТЕЛЯ УЖЕ СУЩЕСТВУЕТ',
           'Имя пользователя должно содержать 3–32 символа: буквы, цифры, точку, дефис или подчёркивание.',
           'ПОЛЬЗОВАТЕЛЬ_СОЗДАН',
           'ПОЛЬЗОВАТЕЛЬ_ОТКЛЮЧЕН',
           'ПОЛЬЗОВАТЕЛЬ_ВКЛЮЧЕН',
           'ПОЛЬЗОВАТЕЛЬ_РАЗБЛОКИРОВАН',
           'ОЖИДАНИЕ ДЕТАЛЕЙ',
           'ОЖИДАЕТ ДЕТАЛЕЙ',
           'ВНИМАНИЕ: ПРИ ВОССТАНОВЛЕНИИ ВСЕ ТЕКУЩИЕ ДАННЫЕ БУДУТ ЗАМЕНЕНЫ ДАННЫМИ ИЗ РЕЗЕРВНОЙ КОПИИ.\n'
           'ЭТО ДЕЙСТВИЕ НЕЛЬЗЯ ОТМЕНИТЬ.\n'
           '\n'
           'ПРОДОЛЖИТЬ?',
           'ГАРАНТИЙНЫЕ ПЕРИОДЫ',
           'СТАТУС ГАРАНТИИ:',
           'СВОДКА ПО ГАРАНТИЯМ',
           'ОТЧЁТ СО СВОДКОЙ ПО ГАРАНТИЯМ',
           'СЛАБЫЙ ПАРОЛЬ',
           'Что вы хотите сделать с базой данных?',
           'ОКНО ЗАКРЫТО',
           'Нельзя отключить текущую учётную запись.',
           'Ваш пароль успешно изменён.',
           '⚠️ ВНИМАНИЕ!',
           '✏️ ИЗМЕНИТЬ',
           '✏️ ИЗМЕНИТЬ КОНТАКТ',
           '❌ УДАЛИТЬ',
           '➕ ДОБАВИТЬ',
           '➕ ДОБАВИТЬ ДЕТАЛЬ',
           '📄 ТЕКСТОВЫЙ ВИД',
           '📋 ТАБЛИЧНЫЙ ВИД',
           '📝 ЗАПИСАТЬ НАПОМИНАНИЕ',
           '📝 ЗАПИСАТЬ НАПОМИНАНИЯ',
           '📤 ЭКСПОРТИРОВАТЬ ВЫБРАННОЕ',
           '🔃 ОБНОВИТЬ',
           '🔄 ОБНОВИТЬ',
           '🗑️ ОЧИСТИТЬ'],
 'tr_TR': ['Yeni bir parça yönetici tarafından oluşturulmalıdır.',
           'Yeni bir tedarikçi yönetici tarafından oluşturulmalıdır.',
           "Bu ada sahip bir parça zaten var. Parçayı seçin ve güncellemek için Kaydet'i kullanın.",
           "Bu ada sahip bir tedarikçi zaten var. Tedarikçiyi seçin ve güncellemek için Kaydet'i kullanın.",
           'ERİŞİM_REDDEDİLDİ',
           'HESAP',
           'HESAP DEVRE DIŞI',
           'HESAP KİLİTLİ',
           'İŞLEM ENGELLENDİ',
           'ETKİN GARANTİLER',
           'UYGULAMA KULLANICISI EKLE',
           'KULLANICI EKLE',
           'EK NOTLAR',
           'TAM ADRES',
           'YÖNETİCİ',
           'YÖNETİCİ',
           'YÖNETİCİ İŞLEMİ ENGELLENDİ',
           'YALNIZCA YÖNETİCİ',
           'YÖNETİCİ GEREKLİ',
           'YÖNETİCİ GEREKLİ.',
           'Başka bir parça bu adı zaten kullanıyor.',
           'Başka bir tedarikçi bu adı zaten kullanıyor.',
           'UYGULAMA KAPATILDI',
           'UYGULAMADAN_ÇIKIŞ',
           'UYGULAMA_AÇILDI',
           'En az bir etkin yönetici gereklidir.',
           'SERVİS BAŞINA ORTALAMA',
           'ORTALAMA ONARIM MALİYETİ',
           'ORTALAMA GELİR',
           'ORTALAMA GELİR/SERVİS',
           'VERİTABANINI YEDEKLE',
           'YEDEKLEME BAŞARISIZ',
           'Yedekleme klasörü boş olamaz',
           'YEDEKLEME KLASÖRÜ:',
           'YEDEKLEME BAŞARILI',
           'Bogor Yerinde Bilgisayar Servisi, cihazınızın onarım garantisinin şu tarihte sona ereceğini hatırlatır:',
           'MARKA:',
           'MARKALAR',
           'HELP.HHP dosyasını derleyin ve HELP.CHM dosyasını uygulama klasörüne kopyalayın.',
           'Otomatik hesaplanır: servis tarihi + garanti süresi',
           'Seçilen tedarikçi belirlenemedi',
           'PAROLAYI DEĞİŞTİR',
           'ŞUNUN PAROLASINI DEĞİŞTİR',
           'TAMAMLANDI',
           'BİLGİSAYAR SERVİS YÖNETİCİSİ',
           'SİLMEYİ ONAYLA',
           'PAROLAYI ONAYLA',
           'GERİ YÜKLEMEYİ ONAYLA',
           'İLETİŞİM ADI',
           'İLETİŞİM:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'PANOYA KOPYALANDI',
           'ADRESİ KOPYALA',
           "GOOGLE HARİTALAR'I KOPYALA",
           'FATURAYI KOPYALA',
           'TELEFONU KOPYALA',
           'TELİF HAKKI © ECOMTECH 2026',
           'SAYI',
           'YÖNETİCİ OLUŞTUR',
           'İLK YÖNETİCİ HESABINI OLUŞTUR',
           'OLUŞTURULDU',
           'OLUŞTURAN:',
           'OLUŞTURULMA TARİHİ',
           'CSV DOSYALARI',
           'GEÇERLİ PAROLA',
           'Geçerli parola yanlış.',
           'GEÇERLİ SEKME YARDIMI',
           'MÜŞTERİ ADI',
           'HATIRLATICI İÇİN SEÇİLEN MÜŞTERİLER',
           'HATIRLATICI İÇİN SEÇİLEN MÜŞTERİLER:',
           'Güvenli Bilgisayar Servis Yöneticisi 4.4',
           'PANO YENİLEME BAŞARISIZ',
           'PANO GENEL BAKIŞI',
           'Veriler şuraya aktarıldı:',
           'VERİTABANI KONUMU GEREKLİ',
           'Veritabanı yolu boş olamaz',
           'VERİTABANI YOLU:',
           'Veritabanı yolları başarıyla güncellendi!',
           'VERİTABANI YEDEKTEN GERİ YÜKLENDİ',
           'Veritabanı yedekten başarıyla geri yüklendi!',
           'VERİTABANI_GERİ_YÜKLENDİ',
           'TARİH',
           'Sayın [CUSTOMER],',
           'AYRINTILAR',
           'CİHAZ MARKASI İSTATİSTİKLERİ',
           'CİHAZ MODELİ',
           'CİHAZ TÜRLERİ',
           'GÖRÜNEN AD',
           'BAĞIŞ:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'D: sürücüsü kullanılamıyor. Lütfen başka bir klasör seçin.',
           'D: sürücüsü bulunamadı.\nLütfen veritabanı için bir klasör oluşturun veya seçin.',
           'SÜRÜCÜ BULUNAMADI',
           'SAHTE VERİ DÜZENLEYİCİ AÇILDI',
           'YİNELENEN PARÇA',
           'YİNELENEN TEDARİKÇİ',
           'DÜZENLEME MODU',
           "DÜZENLEME MODU: ALANLARI GÜNCELLEYİN VE KAYDET'E TIKLAYIN",
           "DÜZENLEME MODU: ALANLARI GÜNCELLEYİN VE GÜNCELLE'YE TIKLAYIN",
           'ETKİNLEŞTİR / DEVRE DIŞI BIRAK',
           'BİTİŞ TARİHİ:',
           'Kullanıcı adı ve parolayı girin.',
           'EXCEL DOSYALARI',
           '30 GÜN İÇİNDE SONA ERİYOR',
           'YAKINDA SONA ERİYOR',
           'YAKINDA SONA ERİYOR (< 7 GÜN)',
           'YAKINDA SONA ERİYOR (≤ 30 GÜN)',
           'YAKINDA SONA ERİYOR (≤ 7 GÜN)',
           'YAKINDA SONA ERİYOR (≤30 gün)',
           'SONA ERME TARİHİ BAŞLANGICI',
           'SONA ERME TARİHİ BAŞLANGICI:',
           'AKTARMA TAMAMLANDI',
           'AKTARMA HATASI',
           'AKTARMA BAŞARISIZ',
           'AKTARMA BAŞARILI',
           'ŞURAYA AKTARILDI:',
           'Aktarma başarısız',
           'FİNANSAL RAPOR',
           'İLK YÖNETİCİ OTURUMU',
           'İLK VERİTABANI KURULUMU',
           'İLK KURULUM:\nYeni veritabanını kaydetmek için bir klasör seçin.',
           'FORM TEMİZLENDİ',
           'Başlangıç tarihi bitiş tarihinden sonra olamaz',
           'TAM ADRES',
           'GOOGLE HARİTALAR',
           'GOOGLE HARİTALAR BAĞLANTISI KOPYALANDI',
           'GOOGLE HARİTALAR BAĞLANTISI BOŞ',
           'GOOGLE HARİTALAR BAĞLANTISI BOŞ (PANO TEMİZLENDİ)',
           'GOOGLE HARİTALAR:',
           'YARDIM İÇERİĞİ',
           'ID (ROWID)',
           'İÇE AKTARMA TAMAMLANDI',
           'İÇE AKTARMA BAŞARISIZ',
           'İçe/dışa aktarma klasörü boş olamaz',
           'İÇE/DIŞA AKTARMA KLASÖRÜ:',
           'DEVAM EDİYOR',
           'ONARIMDA',
           'GEÇERSİZ KİMLİK BİLGİLERİ',
           'GEÇERSİZ TARİH ARALIĞI',
           'GEÇERSİZ SÜRÜCÜ',
           'GEÇERSİZ YOL',
           'Geçersiz kullanıcı rolü.',
           'GEÇERSİZ KULLANICI ADI',
           'Geçersiz kullanıcı adı veya parola.',
           'Geçersiz garanti tarih aralığı: başlangıç tarihi bitiş tarihinden sonra',
           'FATURA NUMARASI',
           'SORUNLAR',
           'SON GİRİŞ',
           'Eski .XLS dosyaları desteklenmez. Önce dosyayı .XLSX veya CSV olarak kaydedin.',
           'ŞU ZAMANA KADAR KİLİTLİ',
           'HATIRLATICIYI KAYDET',
           'HATIRLATICILARI KAYDET',
           'GARANTİ HATIRLATICISINI KAYDET',
           'GİRİŞ',
           'Giriş başarılı.',
           'GİRİŞ_ENGELLENDİ',
           'GİRİŞ_BAŞARISIZ',
           'GİRİŞ_BAŞARILI',
           'ÇIKIŞ',
           'DÜŞÜK STOKLU PARÇALAR',
           'ANA PENCERE AÇILDI',
           'En az 4 karakter. Harf, sayı, boşluk veya simgelerden oluşan her biçime izin verilir.',
           'CİHAZ MODELİ',
           'MODELLER',
           'ÇOK DİLLİ SAHTE VERİ OLUŞTURUCU AÇILDI',
           'SON ETKİNLİĞİM',
           'YOK',
           'MÜŞTERİ ADI',
           'AD GEREKLİ',
           'AD:',
           'YENİ PAROLA',
           'VERİ YOK',
           'VERİTABANI SEÇİLMEDİ',
           'Varsayılan parola kullanılmaz. Parolalar yalnızca benzersiz, tuzlanmış scrypt karmaları olarak saklanır. '
           'Bu yönetici parolasını güvenli tutun.',
           'Mevcut veritabanı seçilmedi. Uygulama yeni bir veritabanı oluşturmaya devam edecek.',
           'Hiçbir SMS veya WhatsApp iletisi gönderilmedi.',
           'DURUM VERİSİ YOK',
           'GARANTİ YOK',
           'Hiçbir tür garanti verilmez',
           'TELEFON NUMARASI',
           'BULUNAMADI',
           'Mevcut bir veritabanını açın veya bu uygulama için yeni bir veritabanı oluşturun.',
           'UYGULAMA YARDIM İÇERİĞİNİ AÇ',
           'SAHTE VERİ DÜZENLEYİCİYİ AÇ',
           "GOOGLE HARİTALAR'I AÇ",
           'ETKİN SEKME YARDIMINI AÇ',
           'PARÇA ADI',
           'PARÇA ADI GEREKLİ',
           'PARÇA TÜRLERİ',
           'PARÇALAR YENİLENDİ',
           'PAROLA',
           'PAROLA DEĞİŞTİRİLDİ',
           'Parola değiştirildi.',
           'Parola çok uzun.',
           'PAROLALAR EŞLEŞMİYOR',
           'Parola en az 4 karakter içermelidir.',
           'PAROLA DEĞİŞTİRİLMEDİ',
           'PAROLA SIFIRLANDI',
           'PAROLA_DEĞİŞTİRİLDİ',
           'PAROLA_DEĞİŞTİRME_BAŞARISIZ',
           'PAROLA_SIFIRLANDI',
           'TELEFON NUMARASI',
           'TELEFON:',
           'DUMMY_CREATOR.EXE veya DUMMY_CREATOR.PY dosyasını uygulamanın yanına yerleştirin.',
           'Daha fazla hizmete ihtiyacınız varsa lütfen bizimle iletişime geçin.',
           'Lütfen bir XLSX veya CSV dosyası seçin',
           'Lütfen bir Excel veya CSV dosyası seçin',
           'SATIN ALMA FATURASI:',
           'MİKTAR',
           'SON GÜVENLİK VE VERİ ETKİNLİĞİ',
           'SON SERVİS KAYITLARI',
           'PANOYU YENİLE',
           'AÇILIR LİSTE VERİLERİNİ YENİDEN YÜKLE',
           'HATIRLATICI İLETİSİ ŞABLONU',
           'HATIRLATICILAR KAYDEDİLDİ',
           'HATIRLATICILAR GÖNDERİLDİ',
           'ONARIM MALİYETİ',
           'RAPOR AKTARMA BAŞARISIZ',
           'RAPOR AKTARILDI',
           'PAROLAYI SIFIRLA',
           'KULLANICI PAROLASINI SIFIRLA',
           'GARANTİ ARALIĞINI BUGÜNDEN İTİBAREN SIFIRLA',
           'VERİTABANINI GERİ YÜKLE',
           'GERİ YÜKLEME BAŞARISIZ',
           'GERİ YÜKLEME BAŞARILI',
           'ROL',
           'AKTARMA DOSYASINI KAYDET',
           'RAPOR AKTARMASINI KAYDET',
           'AD, TELEFON, CİHAZ VEYA SORUNA GÖRE ARA...',
           'GÜVENLİ YÖNETİCİ KURULUMU',
           'GÜVENLİ GİRİŞ',
           'Silmek için bir kişi seçin',
           'Düzenlemek için bir kişi seçin',
           'Silmek için bir parça seçin',
           'Düzenlemek için bir parça seçin',
           'Güncellemek için bir parça seçin',
           'Silmek için bir tedarikçi seçin',
           'Düzenlemek için bir tedarikçi seçin',
           'Güncellemek için bir tedarikçi seçin',
           'Önce bir kullanıcı seçin.',
           'Düzenlemek için bir garanti satırı seçin',
           'UYGULAMA PARA BİRİMİNİ SEÇİN:',
           'UYGULAMA DİLİNİ VE YEREL AYARINI SEÇİN:',
           'UYGULAMA İÇİN VARSAYILAN PARA BİRİMİNİ SEÇİN:',
           'DİL / YEREL AYAR SEÇİN',
           'AKTARILACAK SATIRLARI SEÇİN',
           'HATIRLATICI KAYDETMEK İÇİN GARANTİLERİ SEÇİN',
           'HATIRLATICI GÖNDERMEK İÇİN GARANTİLERİ SEÇİN',
           'HATIRLATICILARI GÖNDER',
           'SERVİS AÇIKLAMASI',
           'BOGOR YERİNDE BİLGİSAYAR SERVİSİ',
           'SERVİS DURUM GRAFİĞİ',
           'SERVİS ÖZETİ',
           'SERVİS ÖZET RAPORU',
           'Şunun için yeni parola belirle',
           'VERİTABANI KONUMUNU AYARLA',
           'SERVİS TARİHİNİ BUGÜN OLARAK AYARLA',
           'AYARLAR KAYDEDİLDİ',
           'KURULUM BAŞARISIZ',
           'PAROLAYI GÖSTER',
           'PAROLALARI GÖSTER',
           'DEVAM ETMEK İÇİN OTURUM AÇIN',
           'OTURUM AÇILDI',
           'ŞU OLARAK OTURUM AÇILDI',
           'PERSONEL',
           'PERSONEL ERİŞİMİ SINIRLI',
           'Personel yeni servis kayıtları ekleyebilir ve operasyonel verileri görüntüleyebilir. Düzenleme, silme, içe '
           'aktarma, dışa aktarma, yedekleme, yönetim ayarları, kullanıcı yönetimi, tedarikçi değişiklikleri, parça '
           'değişiklikleri, raporlar ve hatırlatıcılar için yönetici gerekir.',
           'DURUMLAR',
           'DEPOLAMA KONUMU',
           'TEDARİKÇİ ADRESİ',
           'TEDARİKÇİ ADI',
           'TEDARİKÇİ ADI GEREKLİ',
           'TEDARİKÇİLER YENİLENDİ',
           'SİSTEM',
           'TEKNİSYEN ADI',
           'TEKNİSYEN PERFORMANSI',
           'TEKNİSYEN PERFORMANS RAPORU',
           'TEKNİSYENLER',
           'Teşekkür ederiz,',
           'Bu kullanıcı adı zaten var.',
           'Yedekleme klasörünün sürücüsü mevcut değil.',
           'Veritabanı yolunun sürücüsü mevcut değil.',
           'Aktarma yolunun sürücüsü mevcut değil.',
           'İçe/dışa aktarma klasörünün sürücüsü mevcut değil.',
           'Parola sıfırlandı.',
           'Parolalar eşleşmiyor.',
           'Bu hesap devre dışı. Bir yöneticiyle iletişime geçin.',
           'Bu işlem yalnızca yönetici hesaplarıyla sınırlıdır.',
           'Bu yazılım olduğu gibi sağlanır',
           'SAAT',
           "EXCEL'E",
           'TOPLAM ETKİN GARANTİ',
           'TOPLAM MÜŞTERİ',
           'TOPLAM CİHAZ',
           'TOPLAM SONA ERMİŞ',
           'TOPLAM YAKINDA SONA ERECEK',
           'TOPLAM GELİR',
           'TOPLAM SEÇİLEN',
           'TOPLAM SERVİS KAYDI',
           'TOPLAM SERVİS DEĞERİ',
           'TOPLAM SERVİS',
           'TÜR:',
           'BENZERSİZ MODELLER',
           'BİLİNMEYEN VEYA GEÇERSİZ KİMLİK BİLGİLERİ',
           'Aktarılacak tablo bilinmiyor.',
           'KİLİDİ AÇ',
           'DESTEKLENMEYEN BİÇİM',
           'D:\\BACKUP KULLAN (VARSAYILAN)',
           'KULLANICI',
           'KULLANICI YÖNETİMİ',
           'KULLANICI OLUŞTURULMADI',
           'Kullanıcı bulunamadı.',
           'KULLANICI ÇIKIŞ İSTEDİ',
           'Kullanıcı durumu güncellendi.',
           'KULLANICI ADI',
           'KULLANICI ADI VAR',
           'Kullanıcı adı 3-32 karakter olmalı ve harf, sayı, nokta, tire veya alt çizgi kullanmalıdır.',
           'KULLANICI_OLUŞTURULDU',
           'KULLANICI_DEVRE_DIŞI',
           'KULLANICI_ETKİN',
           'KULLANICI_KİLİDİ_AÇILDI',
           'PARÇA BEKLENİYOR',
           'PARÇA BEKLİYOR',
           'UYARI: GERİ YÜKLEME TÜM GEÇERLİ VERİLERİ YEDEK VERİLERİYLE DEĞİŞTİRECEKTİR.\n'
           'BU İŞLEM GERİ ALINAMAZ.\n'
           '\n'
           'DEVAM ETMEK İSTİYOR MUSUNUZ?',
           'GARANTİ SÜRELERİ',
           'GARANTİ DURUMU:',
           'GARANTİ ÖZETİ',
           'GARANTİ ÖZET RAPORU',
           'ZAYIF PAROLA',
           'Veritabanıyla ne yapmak istiyorsunuz?',
           'PENCERE KAPATILDI',
           'Geçerli hesabınızı devre dışı bırakamazsınız.',
           'Parolanız başarıyla değiştirildi.',
           '⚠️ UYARI!',
           '✏️ DÜZENLE',
           '✏️ KİŞİYİ DÜZENLE',
           '❌ SİL',
           '➕ EKLE',
           '➕ PARÇA EKLE',
           '📄 METİN GÖRÜNÜMÜ',
           '📋 TABLO GÖRÜNÜMÜ',
           '📝 HATIRLATICIYI KAYDET',
           '📝 HATIRLATICILARI KAYDET',
           '📤 SEÇİLENLERİ AKTAR',
           '🔃 YENİLE',
           '🔄 YENİLE',
           '🗑️ TEMİZLE'],
 'vi_VN': ['Linh kiện mới phải do quản trị viên tạo.',
           'Nhà cung cấp mới phải do quản trị viên tạo.',
           'Đã tồn tại linh kiện có tên này. Hãy chọn linh kiện đó và dùng Lưu để cập nhật.',
           'Đã tồn tại nhà cung cấp có tên này. Hãy chọn nhà cung cấp đó và dùng Lưu để cập nhật.',
           'TRUY_CẬP_BỊ_TỪ_CHỐI',
           'TÀI KHOẢN',
           'TÀI KHOẢN ĐÃ TẮT',
           'TÀI KHOẢN BỊ KHÓA',
           'THAO TÁC BỊ CHẶN',
           'BẢO HÀNH ĐANG HIỆU LỰC',
           'THÊM NGƯỜI DÙNG ỨNG DỤNG',
           'THÊM NGƯỜI DÙNG',
           'GHI CHÚ BỔ SUNG',
           'ĐỊA CHỈ ĐẦY ĐỦ',
           'QUẢN TRỊ',
           'QUẢN TRỊ VIÊN',
           'THAO TÁC QUẢN TRỊ VIÊN BỊ CHẶN',
           'CHỈ DÀNH CHO QUẢN TRỊ VIÊN',
           'CẦN QUẢN TRỊ VIÊN',
           'CẦN QUẢN TRỊ VIÊN.',
           'Một linh kiện khác đã sử dụng tên này.',
           'Một nhà cung cấp khác đã sử dụng tên này.',
           'ỨNG DỤNG ĐÃ ĐÓNG',
           'THOÁT_ỨNG_DỤNG',
           'MỞ_ỨNG_DỤNG',
           'Phải có ít nhất một quản trị viên đang hoạt động.',
           'TRUNG BÌNH MỖI DỊCH VỤ',
           'CHI PHÍ SỬA CHỮA TRUNG BÌNH',
           'DOANH THU TRUNG BÌNH',
           'DOANH THU TRUNG BÌNH/DỊCH VỤ',
           'SAO LƯU CƠ SỞ DỮ LIỆU',
           'SAO LƯU THẤT BẠI',
           'Thư mục sao lưu không được để trống',
           'THƯ MỤC SAO LƯU:',
           'SAO LƯU THÀNH CÔNG',
           'Dịch vụ máy tính tận nơi Bogor xin nhắc rằng bảo hành sửa chữa thiết bị của bạn sẽ hết hạn vào:',
           'THƯƠNG HIỆU:',
           'THƯƠNG HIỆU',
           'Biên dịch HELP.HHP và sao chép HELP.CHM vào thư mục ứng dụng.',
           'Tự động tính: ngày dịch vụ + thời hạn bảo hành',
           'Không thể xác định nhà cung cấp đã chọn',
           'ĐỔI MẬT KHẨU',
           'ĐỔI MẬT KHẨU CHO',
           'ĐÃ HOÀN THÀNH',
           'TRÌNH QUẢN LÝ DỊCH VỤ MÁY TÍNH',
           'XÁC NHẬN XÓA',
           'XÁC NHẬN MẬT KHẨU',
           'XÁC NHẬN KHÔI PHỤC',
           'TÊN LIÊN HỆ',
           'LIÊN HỆ:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'ĐÃ SAO CHÉP VÀO BẢNG TẠM',
           'SAO CHÉP ĐỊA CHỈ',
           'SAO CHÉP GOOGLE MAP',
           'SAO CHÉP HÓA ĐƠN',
           'SAO CHÉP ĐIỆN THOẠI',
           'BẢN QUYỀN © ECOMTECH 2026',
           'SỐ LƯỢNG',
           'TẠO QUẢN TRỊ VIÊN',
           'TẠO TÀI KHOẢN QUẢN TRỊ VIÊN ĐẦU TIÊN',
           'ĐÃ TẠO',
           'NGƯỜI TẠO:',
           'NGÀY TẠO',
           'TỆP CSV',
           'MẬT KHẨU HIỆN TẠI',
           'Mật khẩu hiện tại không đúng.',
           'TRỢ GIÚP THẺ HIỆN TẠI',
           'TÊN KHÁCH HÀNG',
           'KHÁCH HÀNG ĐÃ CHỌN ĐỂ NHẮC',
           'KHÁCH HÀNG ĐÃ CHỌN ĐỂ NHẮC:',
           'Trình quản lý dịch vụ máy tính bảo mật 4.4',
           'LÀM MỚI BẢNG ĐIỀU KHIỂN THẤT BẠI',
           'TỔNG QUAN BẢNG ĐIỀU KHIỂN',
           'Dữ liệu đã được xuất tới:',
           'CẦN VỊ TRÍ CƠ SỞ DỮ LIỆU',
           'Đường dẫn cơ sở dữ liệu không được để trống',
           'ĐƯỜNG DẪN CƠ SỞ DỮ LIỆU:',
           'Đã cập nhật đường dẫn cơ sở dữ liệu thành công!',
           'CƠ SỞ DỮ LIỆU ĐÃ KHÔI PHỤC TỪ BẢN SAO LƯU',
           'Đã khôi phục cơ sở dữ liệu từ bản sao lưu thành công!',
           'CƠ_SỞ_DỮ_LIỆU_ĐÃ_KHÔI_PHỤC',
           'NGÀY',
           'Kính gửi [CUSTOMER],',
           'CHI TIẾT',
           'THỐNG KÊ THƯƠNG HIỆU THIẾT BỊ',
           'MẪU THIẾT BỊ',
           'LOẠI THIẾT BỊ',
           'TÊN HIỂN THỊ',
           'ỦNG HỘ:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Ổ D: không khả dụng. Vui lòng chọn thư mục khác.',
           'Không tìm thấy ổ D:.\nVui lòng tạo hoặc chọn thư mục cho cơ sở dữ liệu.',
           'KHÔNG TÌM THẤY Ổ ĐĨA',
           'ĐÃ MỞ TRÌNH SỬA DỮ LIỆU MẪU',
           'TRÙNG LINH KIỆN',
           'TRÙNG NHÀ CUNG CẤP',
           'CHẾ ĐỘ CHỈNH SỬA',
           'CHẾ ĐỘ CHỈNH SỬA: CẬP NHẬT CÁC TRƯỜNG RỒI BẤM LƯU',
           'CHẾ ĐỘ CHỈNH SỬA: CẬP NHẬT CÁC TRƯỜNG RỒI BẤM CẬP NHẬT',
           'BẬT / TẮT',
           'NGÀY KẾT THÚC:',
           'Nhập cả tên người dùng và mật khẩu.',
           'TỆP EXCEL',
           'HẾT HẠN TRONG 30 NGÀY',
           'SẮP HẾT HẠN',
           'SẮP HẾT HẠN (< 7 NGÀY)',
           'SẮP HẾT HẠN (≤ 30 NGÀY)',
           'SẮP HẾT HẠN (≤ 7 NGÀY)',
           'SẮP HẾT HẠN (≤30 ngày)',
           'NGÀY HẾT HẠN TỪ',
           'NGÀY HẾT HẠN TỪ:',
           'XUẤT HOÀN TẤT',
           'LỖI XUẤT',
           'XUẤT THẤT BẠI',
           'XUẤT THÀNH CÔNG',
           'ĐÃ XUẤT TỚI:',
           'Không thể xuất',
           'BÁO CÁO TÀI CHÍNH',
           'PHIÊN QUẢN TRỊ VIÊN ĐẦU TIÊN',
           'THIẾT LẬP CƠ SỞ DỮ LIỆU LẦN ĐẦU',
           'THIẾT LẬP LẦN ĐẦU:\nVui lòng chọn thư mục để lưu cơ sở dữ liệu mới.',
           'ĐÃ XÓA BIỂU MẪU',
           'Ngày bắt đầu không được sau ngày kết thúc',
           'ĐỊA CHỈ ĐẦY ĐỦ',
           'GOOGLE MAP',
           'ĐÃ SAO CHÉP LIÊN KẾT GOOGLE MAP',
           'LIÊN KẾT GOOGLE MAP TRỐNG',
           'LIÊN KẾT GOOGLE MAP TRỐNG (ĐÃ XÓA BẢNG TẠM)',
           'GOOGLE MAP:',
           'NỘI DUNG TRỢ GIÚP',
           'ID (ROWID)',
           'NHẬP HOÀN TẤT',
           'NHẬP THẤT BẠI',
           'Thư mục nhập/xuất không được để trống',
           'THƯ MỤC NHẬP/XUẤT:',
           'ĐANG THỰC HIỆN',
           'ĐANG SỬA CHỮA',
           'THÔNG TIN ĐĂNG NHẬP KHÔNG HỢP LỆ',
           'KHOẢNG NGÀY KHÔNG HỢP LỆ',
           'Ổ ĐĨA KHÔNG HỢP LỆ',
           'ĐƯỜNG DẪN KHÔNG HỢP LỆ',
           'Vai trò người dùng không hợp lệ.',
           'TÊN NGƯỜI DÙNG KHÔNG HỢP LỆ',
           'Tên người dùng hoặc mật khẩu không hợp lệ.',
           'Khoảng ngày bảo hành không hợp lệ: ngày bắt đầu sau ngày kết thúc',
           'SỐ HÓA ĐƠN',
           'SỰ CỐ',
           'ĐĂNG NHẬP LẦN CUỐI',
           'Không hỗ trợ tệp .XLS cũ. Trước tiên hãy lưu tệp dưới dạng .XLSX hoặc CSV.',
           'BỊ KHÓA ĐẾN',
           'GHI NHẬT KÝ NHẮC NHỞ',
           'GHI NHẬT KÝ CÁC NHẮC NHỞ',
           'GHI NHẬT KÝ NHẮC BẢO HÀNH',
           'ĐĂNG NHẬP',
           'Đăng nhập thành công.',
           'ĐĂNG_NHẬP_BỊ_CHẶN',
           'ĐĂNG_NHẬP_THẤT_BẠI',
           'ĐĂNG_NHẬP_THÀNH_CÔNG',
           'ĐĂNG XUẤT',
           'LINH KIỆN SẮP HẾT HÀNG',
           'ĐÃ MỞ CỬA SỔ CHÍNH',
           'Tối thiểu 4 ký tự. Cho phép mọi dạng chữ cái, số, khoảng trắng hoặc ký hiệu.',
           'MẪU THIẾT BỊ',
           'CÁC MẪU',
           'ĐÃ MỞ TRÌNH TẠO DỮ LIỆU MẪU ĐA NGÔN NGỮ',
           'HOẠT ĐỘNG GẦN ĐÂY CỦA TÔI',
           'KHÔNG ÁP DỤNG',
           'TÊN KHÁCH HÀNG',
           'BẮT BUỘC NHẬP TÊN',
           'TÊN:',
           'MẬT KHẨU MỚI',
           'KHÔNG CÓ DỮ LIỆU',
           'CHƯA CHỌN CƠ SỞ DỮ LIỆU',
           'Không sử dụng mật khẩu mặc định. Mật khẩu chỉ được lưu dưới dạng băm scrypt duy nhất có thêm salt. Hãy giữ '
           'an toàn mật khẩu quản trị viên này.',
           'Chưa chọn cơ sở dữ liệu hiện có. Ứng dụng sẽ tiếp tục tạo cơ sở dữ liệu mới.',
           'Không có tin nhắn SMS hoặc WhatsApp nào được gửi.',
           'KHÔNG CÓ DỮ LIỆU TRẠNG THÁI',
           'KHÔNG BẢO HÀNH',
           'Không cung cấp bất kỳ hình thức bảo đảm nào',
           'SỐ ĐIỆN THOẠI',
           'KHÔNG TÌM THẤY',
           'Mở cơ sở dữ liệu hiện có hoặc tạo cơ sở dữ liệu mới cho ứng dụng này.',
           'MỞ NỘI DUNG TRỢ GIÚP ỨNG DỤNG',
           'MỞ TRÌNH SỬA DỮ LIỆU MẪU',
           'MỞ GOOGLE MAP',
           'MỞ TRỢ GIÚP CHO THẺ ĐANG HOẠT ĐỘNG',
           'TÊN LINH KIỆN',
           'BẮT BUỘC NHẬP TÊN LINH KIỆN',
           'LOẠI LINH KIỆN',
           'ĐÃ LÀM MỚI LINH KIỆN',
           'MẬT KHẨU',
           'ĐÃ ĐỔI MẬT KHẨU',
           'Mật khẩu đã được đổi.',
           'Mật khẩu quá dài.',
           'MẬT KHẨU KHÔNG KHỚP',
           'Mật khẩu phải có ít nhất 4 ký tự.',
           'CHƯA ĐỔI MẬT KHẨU',
           'ĐÃ ĐẶT LẠI MẬT KHẨU',
           'MẬT_KHẨU_ĐÃ_ĐỔI',
           'ĐỔI_MẬT_KHẨU_THẤT_BẠI',
           'MẬT_KHẨU_ĐÃ_ĐẶT_LẠI',
           'SỐ ĐIỆN THOẠI',
           'ĐIỆN THOẠI:',
           'Đặt DUMMY_CREATOR.EXE hoặc DUMMY_CREATOR.PY cạnh ứng dụng.',
           'Vui lòng liên hệ với chúng tôi nếu bạn cần thêm dịch vụ.',
           'Vui lòng chọn tệp XLSX hoặc CSV',
           'Vui lòng chọn tệp Excel hoặc CSV',
           'HÓA ĐƠN MUA HÀNG:',
           'SỐ LƯỢNG',
           'HOẠT ĐỘNG BẢO MẬT VÀ DỮ LIỆU GẦN ĐÂY',
           'HỒ SƠ DỊCH VỤ GẦN ĐÂY',
           'LÀM MỚI BẢNG ĐIỀU KHIỂN',
           'TẢI LẠI DỮ LIỆU DANH SÁCH',
           'MẪU TIN NHẮN NHẮC NHỞ',
           'ĐÃ GHI NHẬT KÝ NHẮC NHỞ',
           'ĐÃ GỬI NHẮC NHỞ',
           'CHI PHÍ SỬA CHỮA',
           'XUẤT BÁO CÁO THẤT BẠI',
           'ĐÃ XUẤT BÁO CÁO',
           'ĐẶT LẠI MẬT KHẨU',
           'ĐẶT LẠI MẬT KHẨU NGƯỜI DÙNG',
           'ĐẶT LẠI KHOẢNG BẢO HÀNH TỪ HÔM NAY',
           'KHÔI PHỤC CƠ SỞ DỮ LIỆU',
           'KHÔI PHỤC THẤT BẠI',
           'KHÔI PHỤC THÀNH CÔNG',
           'VAI TRÒ',
           'LƯU TỆP XUẤT',
           'LƯU BẢN XUẤT BÁO CÁO',
           'TÌM THEO TÊN, ĐIỆN THOẠI, THIẾT BỊ HOẶC SỰ CỐ...',
           'THIẾT LẬP QUẢN TRỊ VIÊN AN TOÀN',
           'ĐĂNG NHẬP AN TOÀN',
           'Chọn một liên hệ để xóa',
           'Chọn một liên hệ để chỉnh sửa',
           'Chọn một linh kiện để xóa',
           'Chọn một linh kiện để chỉnh sửa',
           'Chọn một linh kiện để cập nhật',
           'Chọn một nhà cung cấp để xóa',
           'Chọn một nhà cung cấp để chỉnh sửa',
           'Chọn một nhà cung cấp để cập nhật',
           'Hãy chọn người dùng trước.',
           'Chọn một dòng bảo hành để chỉnh sửa',
           'CHỌN TIỀN TỆ ỨNG DỤNG:',
           'CHỌN NGÔN NGỮ VÀ NGÔN NGỮ KHU VỰC CỦA ỨNG DỤNG:',
           'CHỌN TIỀN TỆ MẶC ĐỊNH CHO ỨNG DỤNG:',
           'CHỌN NGÔN NGỮ / KHU VỰC',
           'CHỌN CÁC DÒNG ĐỂ XUẤT',
           'CHỌN BẢO HÀNH ĐỂ GHI NHẬT KÝ NHẮC NHỞ',
           'CHỌN BẢO HÀNH ĐỂ GỬI NHẮC NHỞ',
           'GỬI NHẮC NHỞ',
           'MÔ TẢ DỊCH VỤ',
           'DỊCH VỤ MÁY TÍNH TẬN NƠI BOGOR',
           'BIỂU ĐỒ TRẠNG THÁI DỊCH VỤ',
           'TÓM TẮT DỊCH VỤ',
           'BÁO CÁO TÓM TẮT DỊCH VỤ',
           'Đặt mật khẩu mới cho',
           'ĐẶT VỊ TRÍ CƠ SỞ DỮ LIỆU',
           'ĐẶT NGÀY DỊCH VỤ LÀ HÔM NAY',
           'ĐÃ LƯU CÀI ĐẶT',
           'THIẾT LẬP THẤT BẠI',
           'HIỂN THỊ MẬT KHẨU',
           'HIỂN THỊ MẬT KHẨU',
           'ĐĂNG NHẬP ĐỂ TIẾP TỤC',
           'ĐÃ ĐĂNG NHẬP',
           'ĐĂNG NHẬP VỚI TƯ CÁCH',
           'NHÂN VIÊN',
           'QUYỀN TRUY CẬP CỦA NHÂN VIÊN BỊ HẠN CHẾ',
           'Nhân viên có thể thêm hồ sơ dịch vụ mới và xem dữ liệu vận hành. Việc chỉnh sửa, xóa, nhập, xuất, sao lưu, '
           'cài đặt quản trị, quản lý người dùng, thay đổi nhà cung cấp, thay đổi linh kiện, báo cáo và nhắc nhở cần '
           'quyền quản trị viên.',
           'TRẠNG THÁI',
           'VỊ TRÍ LƯU TRỮ',
           'ĐỊA CHỈ NHÀ CUNG CẤP',
           'TÊN NHÀ CUNG CẤP',
           'BẮT BUỘC NHẬP TÊN NHÀ CUNG CẤP',
           'ĐÃ LÀM MỚI NHÀ CUNG CẤP',
           'HỆ THỐNG',
           'TÊN KỸ THUẬT VIÊN',
           'HIỆU SUẤT KỸ THUẬT VIÊN',
           'BÁO CÁO HIỆU SUẤT KỸ THUẬT VIÊN',
           'KỸ THUẬT VIÊN',
           'Xin cảm ơn,',
           'Tên người dùng đó đã tồn tại.',
           'Ổ đĩa của thư mục sao lưu không tồn tại.',
           'Ổ đĩa của đường dẫn cơ sở dữ liệu không tồn tại.',
           'Ổ đĩa của đường dẫn xuất không tồn tại.',
           'Ổ đĩa của thư mục nhập/xuất không tồn tại.',
           'Mật khẩu đã được đặt lại.',
           'Mật khẩu không khớp.',
           'Tài khoản này đã bị tắt. Hãy liên hệ quản trị viên.',
           'Thao tác này chỉ dành cho tài khoản quản trị viên.',
           'Phần mềm này được cung cấp nguyên trạng',
           'THỜI GIAN',
           'SANG EXCEL',
           'TỔNG BẢO HÀNH ĐANG HIỆU LỰC',
           'TỔNG KHÁCH HÀNG',
           'TỔNG THIẾT BỊ',
           'TỔNG ĐÃ HẾT HẠN',
           'TỔNG SẮP HẾT HẠN',
           'TỔNG DOANH THU',
           'TỔNG ĐÃ CHỌN',
           'TỔNG HỒ SƠ DỊCH VỤ',
           'TỔNG GIÁ TRỊ DỊCH VỤ',
           'TỔNG DỊCH VỤ',
           'LOẠI:',
           'MẪU DUY NHẤT',
           'THÔNG TIN ĐĂNG NHẬP KHÔNG XÁC ĐỊNH HOẶC KHÔNG HỢP LỆ',
           'Không xác định bảng để xuất.',
           'MỞ KHÓA',
           'ĐỊNH DẠNG KHÔNG ĐƯỢC HỖ TRỢ',
           'DÙNG D:\\BACKUP (MẶC ĐỊNH)',
           'NGƯỜI DÙNG',
           'QUẢN LÝ NGƯỜI DÙNG',
           'CHƯA TẠO NGƯỜI DÙNG',
           'Không tìm thấy người dùng.',
           'NGƯỜI DÙNG YÊU CẦU ĐĂNG XUẤT',
           'Đã cập nhật trạng thái người dùng.',
           'TÊN NGƯỜI DÙNG',
           'TÊN NGƯỜI DÙNG ĐÃ TỒN TẠI',
           'Tên người dùng phải dài 3-32 ký tự và chỉ dùng chữ cái, số, dấu chấm, dấu gạch ngang hoặc dấu gạch dưới.',
           'NGƯỜI_DÙNG_ĐÃ_TẠO',
           'NGƯỜI_DÙNG_ĐÃ_TẮT',
           'NGƯỜI_DÙNG_ĐÃ_BẬT',
           'NGƯỜI_DÙNG_ĐÃ_MỞ_KHÓA',
           'ĐANG CHỜ LINH KIỆN',
           'CHỜ LINH KIỆN',
           'CẢNH BÁO: KHÔI PHỤC SẼ THAY THẾ TOÀN BỘ DỮ LIỆU HIỆN TẠI BẰNG DỮ LIỆU SAO LƯU.\n'
           'KHÔNG THỂ HOÀN TÁC THAO TÁC NÀY.\n'
           '\n'
           'BẠN CÓ MUỐN TIẾP TỤC KHÔNG?',
           'THỜI HẠN BẢO HÀNH',
           'TRẠNG THÁI BẢO HÀNH:',
           'TÓM TẮT BẢO HÀNH',
           'BÁO CÁO TÓM TẮT BẢO HÀNH',
           'MẬT KHẨU YẾU',
           'Bạn muốn làm gì với cơ sở dữ liệu?',
           'CỬA SỔ ĐÃ ĐÓNG',
           'Bạn không thể tắt tài khoản hiện tại của mình.',
           'Mật khẩu của bạn đã được đổi thành công.',
           '⚠️ CẢNH BÁO!',
           '✏️ CHỈNH SỬA',
           '✏️ CHỈNH SỬA LIÊN HỆ',
           '❌ XÓA',
           '➕ THÊM',
           '➕ THÊM LINH KIỆN',
           '📄 DẠNG VĂN BẢN',
           '📋 DẠNG BẢNG',
           '📝 GHI NHẬT KÝ NHẮC NHỞ',
           '📝 GHI NHẬT KÝ CÁC NHẮC NHỞ',
           '📤 XUẤT MỤC ĐÃ CHỌN',
           '🔃 LÀM MỚI',
           '🔄 LÀM MỚI',
           '🗑️ XÓA SẠCH'],
 'th_TH': ['ผู้ดูแลระบบต้องเป็นผู้สร้างอะไหล่ใหม่',
           'ผู้ดูแลระบบต้องเป็นผู้สร้างผู้จำหน่ายรายใหม่',
           'มีอะไหล่ชื่อนี้อยู่แล้ว โปรดเลือกอะไหล่นั้นและใช้ บันทึก เพื่ออัปเดต',
           'มีผู้จำหน่ายชื่อนี้อยู่แล้ว โปรดเลือกผู้จำหน่ายนั้นและใช้ บันทึก เพื่ออัปเดต',
           'ปฏิเสธ_การเข้าถึง',
           'บัญชี',
           'บัญชีถูกปิดใช้งาน',
           'บัญชีถูกล็อก',
           'การดำเนินการถูกบล็อก',
           'การรับประกันที่ยังมีผล',
           'เพิ่มผู้ใช้แอปพลิเคชัน',
           'เพิ่มผู้ใช้',
           'หมายเหตุเพิ่มเติม',
           'ที่อยู่แบบเต็ม',
           'ผู้ดูแล',
           'ผู้ดูแลระบบ',
           'การดำเนินการของผู้ดูแลระบบถูกบล็อก',
           'เฉพาะผู้ดูแลระบบ',
           'ต้องใช้สิทธิ์ผู้ดูแลระบบ',
           'ต้องใช้สิทธิ์ผู้ดูแลระบบ',
           'มีอะไหล่อื่นใช้ชื่อนี้อยู่แล้ว',
           'มีผู้จำหน่ายอื่นใช้ชื่อนี้อยู่แล้ว',
           'ปิดแอปพลิเคชันแล้ว',
           'ออกจาก_แอปพลิเคชัน',
           'เปิด_แอปพลิเคชัน',
           'ต้องมีผู้ดูแลระบบที่ใช้งานอยู่อย่างน้อยหนึ่งคน',
           'ค่าเฉลี่ยต่อบริการ',
           'ค่าซ่อมเฉลี่ย',
           'รายรับเฉลี่ย',
           'รายรับเฉลี่ย/บริการ',
           'สำรองฐานข้อมูล',
           'สำรองข้อมูลล้มเหลว',
           'โฟลเดอร์สำรองข้อมูลต้องไม่ว่าง',
           'โฟลเดอร์สำรองข้อมูล:',
           'สำรองข้อมูลสำเร็จ',
           'บริการคอมพิวเตอร์นอกสถานที่โบกอร์ขอแจ้งเตือนว่าการรับประกันการซ่อมอุปกรณ์ของคุณจะหมดอายุในวันที่:',
           'ยี่ห้อ:',
           'ยี่ห้อ',
           'สร้าง HELP.HHP และคัดลอก HELP.CHM ไปยังโฟลเดอร์แอปพลิเคชัน',
           'คำนวณอัตโนมัติ: วันที่ให้บริการ + ระยะเวลารับประกัน',
           'ไม่สามารถระบุผู้จำหน่ายที่เลือกได้',
           'เปลี่ยนรหัสผ่าน',
           'เปลี่ยนรหัสผ่านสำหรับ',
           'เสร็จสมบูรณ์',
           'ตัวจัดการบริการคอมพิวเตอร์',
           'ยืนยันการลบ',
           'ยืนยันรหัสผ่าน',
           'ยืนยันการคืนค่า',
           'ชื่อผู้ติดต่อ',
           'ติดต่อ:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'คัดลอกไปยังคลิปบอร์ดแล้ว',
           'คัดลอกที่อยู่',
           'คัดลอก Google Maps',
           'คัดลอกใบแจ้งหนี้',
           'คัดลอกโทรศัพท์',
           'ลิขสิทธิ์ © ECOMTECH 2026',
           'จำนวน',
           'สร้างผู้ดูแลระบบ',
           'สร้างบัญชีผู้ดูแลระบบบัญชีแรก',
           'สร้างแล้ว',
           'สร้างโดย:',
           'วันที่สร้าง',
           'ไฟล์ CSV',
           'รหัสผ่านปัจจุบัน',
           'รหัสผ่านปัจจุบันไม่ถูกต้อง',
           'วิธีใช้แท็บปัจจุบัน',
           'ชื่อลูกค้า',
           'ลูกค้าที่เลือกสำหรับการแจ้งเตือน',
           'ลูกค้าที่เลือกสำหรับการแจ้งเตือน:',
           'ตัวจัดการบริการคอมพิวเตอร์แบบปลอดภัย 4.4',
           'รีเฟรชแดชบอร์ดล้มเหลว',
           'ภาพรวมแดชบอร์ด',
           'ส่งออกข้อมูลไปยัง:',
           'ต้องระบุตำแหน่งฐานข้อมูล',
           'เส้นทางฐานข้อมูลต้องไม่ว่าง',
           'เส้นทางฐานข้อมูล:',
           'อัปเดตเส้นทางฐานข้อมูลสำเร็จ!',
           'คืนค่าฐานข้อมูลจากข้อมูลสำรองแล้ว',
           'คืนค่าฐานข้อมูลจากข้อมูลสำรองสำเร็จ!',
           'คืนค่า_ฐานข้อมูลแล้ว',
           'วันที่',
           'เรียน [CUSTOMER],',
           'รายละเอียด',
           'สถิติยี่ห้ออุปกรณ์',
           'รุ่นอุปกรณ์',
           'ประเภทอุปกรณ์',
           'ชื่อที่แสดง',
           'บริจาค:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'ไม่สามารถใช้ไดรฟ์ D: ได้ โปรดเลือกโฟลเดอร์อื่น',
           'ไม่พบไดรฟ์ D:\nโปรดสร้างหรือเลือกโฟลเดอร์สำหรับฐานข้อมูล',
           'ไม่พบไดรฟ์',
           'เปิดตัวแก้ไขข้อมูลจำลองแล้ว',
           'อะไหล่ซ้ำ',
           'ผู้จำหน่ายซ้ำ',
           'โหมดแก้ไข',
           'โหมดแก้ไข: อัปเดตช่องข้อมูลแล้วคลิก บันทึก',
           'โหมดแก้ไข: อัปเดตช่องข้อมูลแล้วคลิก อัปเดต',
           'เปิดใช้งาน / ปิดใช้งาน',
           'วันที่สิ้นสุด:',
           'กรอกทั้งชื่อผู้ใช้และรหัสผ่าน',
           'ไฟล์ Excel',
           'หมดอายุภายใน 30 วัน',
           'ใกล้หมดอายุ',
           'ใกล้หมดอายุ (< 7 วัน)',
           'ใกล้หมดอายุ (≤ 30 วัน)',
           'ใกล้หมดอายุ (≤ 7 วัน)',
           'ใกล้หมดอายุ (≤30 วัน)',
           'วันที่หมดอายุตั้งแต่',
           'วันที่หมดอายุตั้งแต่:',
           'ส่งออกเสร็จสมบูรณ์',
           'ข้อผิดพลาดในการส่งออก',
           'ส่งออกล้มเหลว',
           'ส่งออกสำเร็จ',
           'ส่งออกไปยัง:',
           'ไม่สามารถส่งออกได้',
           'รายงานทางการเงิน',
           'เซสชันผู้ดูแลระบบครั้งแรก',
           'ตั้งค่าฐานข้อมูลครั้งแรก',
           'การตั้งค่าครั้งแรก:\nโปรดเลือกโฟลเดอร์สำหรับบันทึกฐานข้อมูลใหม่',
           'ล้างแบบฟอร์มแล้ว',
           'วันที่เริ่มต้นต้องไม่อยู่หลังวันที่สิ้นสุด',
           'ที่อยู่แบบเต็ม',
           'Google Maps',
           'คัดลอกลิงก์ Google Maps แล้ว',
           'ลิงก์ Google Maps ว่าง',
           'ลิงก์ Google Maps ว่าง (ล้างคลิปบอร์ดแล้ว)',
           'Google Maps:',
           'เนื้อหาวิธีใช้',
           'ID (ROWID)',
           'นำเข้าเสร็จสมบูรณ์',
           'นำเข้าล้มเหลว',
           'โฟลเดอร์นำเข้า/ส่งออกต้องไม่ว่าง',
           'โฟลเดอร์นำเข้า/ส่งออก:',
           'กำลังดำเนินการ',
           'กำลังซ่อม',
           'ข้อมูลรับรองไม่ถูกต้อง',
           'ช่วงวันที่ไม่ถูกต้อง',
           'ไดรฟ์ไม่ถูกต้อง',
           'เส้นทางไม่ถูกต้อง',
           'บทบาทผู้ใช้ไม่ถูกต้อง',
           'ชื่อผู้ใช้ไม่ถูกต้อง',
           'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง',
           'ช่วงวันที่รับประกันไม่ถูกต้อง: วันที่เริ่มต้นอยู่หลังวันที่สิ้นสุด',
           'หมายเลขใบแจ้งหนี้',
           'ปัญหา',
           'เข้าสู่ระบบล่าสุด',
           'ไม่รองรับไฟล์ .XLS รุ่นเก่า โปรดบันทึกไฟล์เป็น .XLSX หรือ CSV ก่อน',
           'ล็อกจนถึง',
           'บันทึกการแจ้งเตือน',
           'บันทึกการแจ้งเตือน',
           'บันทึกการแจ้งเตือนการรับประกัน',
           'เข้าสู่ระบบ',
           'เข้าสู่ระบบสำเร็จ',
           'การเข้าสู่ระบบ_ถูกบล็อก',
           'การเข้าสู่ระบบ_ล้มเหลว',
           'การเข้าสู่ระบบ_สำเร็จ',
           'ออกจากระบบ',
           'อะไหล่สต็อกต่ำ',
           'เปิดหน้าต่างหลักแล้ว',
           'อย่างน้อย 4 อักขระ อนุญาตให้ใช้ตัวอักษร ตัวเลข ช่องว่าง หรือสัญลักษณ์ได้ทุกรูปแบบ',
           'รุ่นอุปกรณ์',
           'รุ่น',
           'เปิดตัวสร้างข้อมูลจำลองหลายภาษาแล้ว',
           'กิจกรรมล่าสุดของฉัน',
           'ไม่มี',
           'ชื่อลูกค้า',
           'ต้องระบุชื่อ',
           'ชื่อ:',
           'รหัสผ่านใหม่',
           'ไม่มีข้อมูล',
           'ไม่ได้เลือกฐานข้อมูล',
           'ไม่มีการใช้รหัสผ่านเริ่มต้น รหัสผ่านจะถูกจัดเก็บเป็นแฮช scrypt แบบใส่ salt ที่ไม่ซ้ำกันเท่านั้น '
           'โปรดเก็บรหัสผ่านผู้ดูแลระบบนี้ให้ปลอดภัย',
           'ไม่ได้เลือกฐานข้อมูลที่มีอยู่ แอปพลิเคชันจะดำเนินการสร้างฐานข้อมูลใหม่ต่อไป',
           'ไม่มีการส่งข้อความ SMS หรือ WhatsApp',
           'ไม่มีข้อมูลสถานะ',
           'ไม่มีการรับประกัน',
           'ไม่มีการรับประกันใด ๆ ทั้งสิ้น',
           'หมายเลขโทรศัพท์',
           'ไม่พบ',
           'เปิดฐานข้อมูลที่มีอยู่หรือสร้างฐานข้อมูลใหม่สำหรับแอปพลิเคชันนี้',
           'เปิดเนื้อหาวิธีใช้ของแอปพลิเคชัน',
           'เปิดตัวแก้ไขข้อมูลจำลอง',
           'เปิด Google Maps',
           'เปิดวิธีใช้สำหรับแท็บที่ใช้งานอยู่',
           'ชื่ออะไหล่',
           'ต้องระบุชื่ออะไหล่',
           'ประเภทอะไหล่',
           'รีเฟรชอะไหล่แล้ว',
           'รหัสผ่าน',
           'เปลี่ยนรหัสผ่านแล้ว',
           'เปลี่ยนรหัสผ่านแล้ว',
           'รหัสผ่านยาวเกินไป',
           'รหัสผ่านไม่ตรงกัน',
           'รหัสผ่านต้องมีอย่างน้อย 4 อักขระ',
           'ไม่ได้เปลี่ยนรหัสผ่าน',
           'รีเซ็ตรหัสผ่านแล้ว',
           'เปลี่ยน_รหัสผ่านแล้ว',
           'เปลี่ยน_รหัสผ่าน_ล้มเหลว',
           'รีเซ็ต_รหัสผ่านแล้ว',
           'หมายเลขโทรศัพท์',
           'โทรศัพท์:',
           'วาง DUMMY_CREATOR.EXE หรือ DUMMY_CREATOR.PY ไว้ข้างแอปพลิเคชัน',
           'โปรดติดต่อเราหากคุณต้องการบริการเพิ่มเติม',
           'โปรดเลือกไฟล์ XLSX หรือ CSV',
           'โปรดเลือกไฟล์ Excel หรือ CSV',
           'ใบแจ้งหนี้การซื้อ:',
           'จำนวน',
           'กิจกรรมความปลอดภัยและข้อมูลล่าสุด',
           'บันทึกบริการล่าสุด',
           'รีเฟรชแดชบอร์ด',
           'โหลดข้อมูลรายการแบบเลื่อนลงใหม่',
           'แม่แบบข้อความแจ้งเตือน',
           'บันทึกการแจ้งเตือนแล้ว',
           'ส่งการแจ้งเตือนแล้ว',
           'ค่าซ่อม',
           'ส่งออกรายงานล้มเหลว',
           'ส่งออกรายงานแล้ว',
           'รีเซ็ตรหัสผ่าน',
           'รีเซ็ตรหัสผ่านผู้ใช้',
           'รีเซ็ตช่วงการรับประกันตั้งแต่วันนี้',
           'คืนค่าฐานข้อมูล',
           'คืนค่าล้มเหลว',
           'คืนค่าสำเร็จ',
           'บทบาท',
           'บันทึกไฟล์ส่งออก',
           'บันทึกการส่งออกรายงาน',
           'ค้นหาตามชื่อ โทรศัพท์ อุปกรณ์ หรือปัญหา...',
           'การตั้งค่าผู้ดูแลระบบที่ปลอดภัย',
           'เข้าสู่ระบบอย่างปลอดภัย',
           'เลือกผู้ติดต่อที่จะลบ',
           'เลือกผู้ติดต่อที่จะแก้ไข',
           'เลือกอะไหล่ที่จะลบ',
           'เลือกอะไหล่ที่จะแก้ไข',
           'เลือกอะไหล่ที่จะอัปเดต',
           'เลือกผู้จำหน่ายที่จะลบ',
           'เลือกผู้จำหน่ายที่จะแก้ไข',
           'เลือกผู้จำหน่ายที่จะอัปเดต',
           'โปรดเลือกผู้ใช้ก่อน',
           'เลือกแถวการรับประกันที่จะแก้ไข',
           'เลือกสกุลเงินของแอปพลิเคชัน:',
           'เลือกภาษาและโลแคลของแอปพลิเคชัน:',
           'เลือกสกุลเงินเริ่มต้นของแอปพลิเคชัน:',
           'เลือกภาษา / โลแคล',
           'เลือกแถวที่จะส่งออก',
           'เลือกการรับประกันเพื่อบันทึกการแจ้งเตือน',
           'เลือกการรับประกันเพื่อส่งการแจ้งเตือน',
           'ส่งการแจ้งเตือน',
           'คำอธิบายบริการ',
           'บริการคอมพิวเตอร์นอกสถานที่โบกอร์',
           'กราฟสถานะบริการ',
           'สรุปบริการ',
           'รายงานสรุปบริการ',
           'ตั้งรหัสผ่านใหม่สำหรับ',
           'ตั้งค่าตำแหน่งฐานข้อมูล',
           'ตั้งวันที่ให้บริการเป็นวันนี้',
           'บันทึกการตั้งค่าแล้ว',
           'ตั้งค่าล้มเหลว',
           'แสดงรหัสผ่าน',
           'แสดงรหัสผ่าน',
           'เข้าสู่ระบบเพื่อดำเนินการต่อ',
           'เข้าสู่ระบบแล้ว',
           'เข้าสู่ระบบในชื่อ',
           'พนักงาน',
           'การเข้าถึงของพนักงานถูกจำกัด',
           'พนักงานสามารถเพิ่มบันทึกบริการใหม่และดูข้อมูลการดำเนินงานได้ การแก้ไข การลบ การนำเข้า การส่งออก '
           'การสำรองข้อมูล การตั้งค่าผู้ดูแลระบบ การจัดการผู้ใช้ การเปลี่ยนแปลงผู้จำหน่าย การเปลี่ยนแปลงอะไหล่ รายงาน '
           'และการแจ้งเตือน ต้องใช้สิทธิ์ผู้ดูแลระบบ',
           'สถานะ',
           'ตำแหน่งจัดเก็บ',
           'ที่อยู่ผู้จำหน่าย',
           'ชื่อผู้จำหน่าย',
           'ต้องระบุชื่อผู้จำหน่าย',
           'รีเฟรชผู้จำหน่ายแล้ว',
           'ระบบ',
           'ชื่อช่างเทคนิค',
           'ประสิทธิภาพของช่างเทคนิค',
           'รายงานประสิทธิภาพของช่างเทคนิค',
           'ช่างเทคนิค',
           'ขอบคุณ',
           'ชื่อผู้ใช้นี้มีอยู่แล้ว',
           'ไม่มีไดรฟ์สำหรับโฟลเดอร์สำรองข้อมูล',
           'ไม่มีไดรฟ์สำหรับเส้นทางฐานข้อมูล',
           'ไม่มีไดรฟ์สำหรับเส้นทางส่งออก',
           'ไม่มีไดรฟ์สำหรับโฟลเดอร์นำเข้า/ส่งออก',
           'รีเซ็ตรหัสผ่านแล้ว',
           'รหัสผ่านไม่ตรงกัน',
           'บัญชีนี้ถูกปิดใช้งาน โปรดติดต่อผู้ดูแลระบบ',
           'การดำเนินการนี้จำกัดเฉพาะบัญชีผู้ดูแลระบบ',
           'ซอฟต์แวร์นี้จัดให้ตามสภาพที่เป็นอยู่',
           'เวลา',
           'ไปยัง Excel',
           'การรับประกันที่ยังมีผลทั้งหมด',
           'ลูกค้าทั้งหมด',
           'อุปกรณ์ทั้งหมด',
           'หมดอายุทั้งหมด',
           'ใกล้หมดอายุทั้งหมด',
           'รายรับรวม',
           'ที่เลือกทั้งหมด',
           'บันทึกบริการทั้งหมด',
           'มูลค่าบริการรวม',
           'บริการทั้งหมด',
           'ประเภท:',
           'รุ่นที่ไม่ซ้ำกัน',
           'ข้อมูลรับรองไม่ทราบหรือไม่ถูกต้อง',
           'ไม่รู้จักตารางที่จะส่งออก',
           'ปลดล็อก',
           'รูปแบบที่ไม่รองรับ',
           'ใช้ D:\\BACKUP (ค่าเริ่มต้น)',
           'ผู้ใช้',
           'การจัดการผู้ใช้',
           'ไม่ได้สร้างผู้ใช้',
           'ไม่พบผู้ใช้',
           'ผู้ใช้ขอออกจากระบบ',
           'อัปเดตสถานะผู้ใช้แล้ว',
           'ชื่อผู้ใช้',
           'ชื่อผู้ใช้มีอยู่แล้ว',
           'ชื่อผู้ใช้ต้องมี 3-32 อักขระ และใช้ได้เฉพาะตัวอักษร ตัวเลข จุด ขีดกลาง หรือขีดล่าง',
           'สร้าง_ผู้ใช้แล้ว',
           'ปิดใช้งาน_ผู้ใช้แล้ว',
           'เปิดใช้งาน_ผู้ใช้แล้ว',
           'ปลดล็อก_ผู้ใช้แล้ว',
           'กำลังรออะไหล่',
           'รออะไหล่',
           'คำเตือน: การคืนค่าจะแทนที่ข้อมูลปัจจุบันทั้งหมดด้วยข้อมูลสำรอง\n'
           'ไม่สามารถยกเลิกการดำเนินการนี้ได้\n'
           '\n'
           'คุณต้องการดำเนินการต่อหรือไม่?',
           'ระยะเวลารับประกัน',
           'สถานะการรับประกัน:',
           'สรุปการรับประกัน',
           'รายงานสรุปการรับประกัน',
           'รหัสผ่านไม่รัดกุม',
           'คุณต้องการทำอะไรกับฐานข้อมูล?',
           'ปิดหน้าต่างแล้ว',
           'คุณไม่สามารถปิดใช้งานบัญชีปัจจุบันของคุณได้',
           'เปลี่ยนรหัสผ่านของคุณสำเร็จแล้ว',
           '⚠️ คำเตือน!',
           '✏️ แก้ไข',
           '✏️ แก้ไขผู้ติดต่อ',
           '❌ ลบ',
           '➕ เพิ่ม',
           '➕ เพิ่มอะไหล่',
           '📄 มุมมองข้อความ',
           '📋 มุมมองตาราง',
           '📝 บันทึกการแจ้งเตือน',
           '📝 บันทึกการแจ้งเตือน',
           '📤 ส่งออกที่เลือก',
           '🔃 รีเฟรช',
           '🔄 รีเฟรช',
           '🗑️ ล้าง'],
 'ms_MY': ['Alat ganti baharu mesti dicipta oleh pentadbir.',
           'Pembekal baharu mesti dicipta oleh pentadbir.',
           'Alat ganti dengan nama ini sudah wujud. Pilih alat ganti tersebut dan gunakan Simpan untuk mengemas '
           'kininya.',
           'Pembekal dengan nama ini sudah wujud. Pilih pembekal tersebut dan gunakan Simpan untuk mengemas kininya.',
           'AKSES_DITOLAK',
           'AKAUN',
           'AKAUN DILUMPUHKAN',
           'AKAUN DIKUNCI',
           'TINDAKAN DISEKAT',
           'WARANTI AKTIF',
           'TAMBAH PENGGUNA APLIKASI',
           'TAMBAH PENGGUNA',
           'NOTA TAMBAHAN',
           'ALAMAT PENUH',
           'PENTADBIR',
           'PENTADBIR',
           'TINDAKAN PENTADBIR DISEKAT',
           'PENTADBIR SAHAJA',
           'PENTADBIR DIPERLUKAN',
           'PENTADBIR DIPERLUKAN.',
           'Alat ganti lain sudah menggunakan nama ini.',
           'Pembekal lain sudah menggunakan nama ini.',
           'APLIKASI DITUTUP',
           'KELUAR_APLIKASI',
           'APLIKASI_DIBUKA',
           'Sekurang-kurangnya seorang pentadbir aktif diperlukan.',
           'PURATA SETIAP SERVIS',
           'PURATA KOS PEMBAIKAN',
           'PURATA HASIL',
           'PURATA HASIL/SERVIS',
           'SANDARKAN PANGKALAN DATA',
           'SANDARAN GAGAL',
           'Folder sandaran tidak boleh kosong',
           'FOLDER SANDARAN:',
           'SANDARAN BERJAYA',
           'Servis Komputer Panggilan Bogor mengingatkan bahawa waranti pembaikan peranti anda akan tamat pada:',
           'JENAMA:',
           'JENAMA',
           'Bina HELP.HHP dan salin HELP.CHM ke folder aplikasi.',
           'Dikira secara automatik: tarikh servis + tempoh waranti',
           'Tidak dapat menentukan pembekal yang dipilih',
           'TUKAR KATA LALUAN',
           'TUKAR KATA LALUAN UNTUK',
           'SELESAI',
           'PENGURUS SERVIS KOMPUTER',
           'SAHKAN PEMADAMAN',
           'SAHKAN KATA LALUAN',
           'SAHKAN PEMULIHAN',
           'NAMA HUBUNGAN',
           'HUBUNGI:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
           'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
           'DISALIN KE PAPAN KLIP',
           'SALIN ALAMAT',
           'SALIN GOOGLE MAP',
           'SALIN INVOIS',
           'SALIN TELEFON',
           'HAK CIPTA © ECOMTECH 2026',
           'BILANGAN',
           'CIPTA PENTADBIR',
           'CIPTA AKAUN PENTADBIR PERTAMA',
           'DICIPTA',
           'DICIPTA OLEH:',
           'TARIKH DICIPTA',
           'FAIL CSV',
           'KATA LALUAN SEMASA',
           'Kata laluan semasa tidak betul.',
           'BANTUAN TAB SEMASA',
           'NAMA PELANGGAN',
           'PELANGGAN DIPILIH UNTUK PERINGATAN',
           'PELANGGAN DIPILIH UNTUK PERINGATAN:',
           'Pengurus Servis Komputer Selamat 4.4',
           'SEGAR SEMULA PAPAN PEMUKA GAGAL',
           'GAMBARAN KESELURUHAN PAPAN PEMUKA',
           'Data telah dieksport ke:',
           'LOKASI PANGKALAN DATA DIPERLUKAN',
           'Laluan pangkalan data tidak boleh kosong',
           'LALUAN PANGKALAN DATA:',
           'Laluan pangkalan data berjaya dikemas kini!',
           'PANGKALAN DATA DIPULIHKAN DARIPADA SANDARAN',
           'Pangkalan data berjaya dipulihkan daripada sandaran!',
           'PANGKALAN_DATA_DIPULIHKAN',
           'TARIKH',
           'Yang dihormati [CUSTOMER],',
           'BUTIRAN',
           'STATISTIK JENAMA PERANTI',
           'MODEL PERANTI',
           'JENIS PERANTI',
           'NAMA PAPARAN',
           'SUMBANGAN:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
           'Pemacu D: tidak tersedia. Sila pilih folder lain.',
           'Pemacu D: tidak ditemui.\nSila cipta atau pilih folder untuk pangkalan data.',
           'PEMACU TIDAK DITEMUI',
           'EDITOR DATA CONTOH DIBUKA',
           'ALAT GANTI PENDUA',
           'PEMBEKAL PENDUA',
           'MOD SUNTING',
           'MOD SUNTING: KEMAS KINI MEDAN KEMUDIAN KLIK SIMPAN',
           'MOD SUNTING: KEMAS KINI MEDAN KEMUDIAN KLIK KEMAS KINI',
           'DAYAKAN / LUMPUHKAN',
           'TARIKH TAMAT:',
           'Masukkan nama pengguna dan kata laluan.',
           'FAIL EXCEL',
           'TAMAT DALAM 30 HARI',
           'AKAN TAMAT TIDAK LAMA LAGI',
           'AKAN TAMAT TIDAK LAMA LAGI (< 7 HARI)',
           'AKAN TAMAT TIDAK LAMA LAGI (≤ 30 HARI)',
           'AKAN TAMAT TIDAK LAMA LAGI (≤ 7 HARI)',
           'AKAN TAMAT TIDAK LAMA LAGI (≤30 hari)',
           'TARIKH TAMAT DARI',
           'TARIKH TAMAT DARI:',
           'EKSPORT SELESAI',
           'RALAT EKSPORT',
           'EKSPORT GAGAL',
           'EKSPORT BERJAYA',
           'DIEKSPORT KE:',
           'Gagal mengeksport',
           'LAPORAN KEWANGAN',
           'SESI PENTADBIR PERTAMA',
           'PERSEDIAAN PANGKALAN DATA KALI PERTAMA',
           'PERSEDIAAN KALI PERTAMA:\nSila pilih folder untuk menyimpan pangkalan data baharu.',
           'BORANG DIKOSONGKAN',
           'Tarikh mula tidak boleh selepas tarikh tamat',
           'ALAMAT PENUH',
           'GOOGLE MAP',
           'PAUTAN GOOGLE MAP DISALIN',
           'PAUTAN GOOGLE MAP KOSONG',
           'PAUTAN GOOGLE MAP KOSONG (PAPAN KLIP DIKOSONGKAN)',
           'GOOGLE MAP:',
           'KANDUNGAN BANTUAN',
           'ID (ROWID)',
           'IMPORT SELESAI',
           'IMPORT GAGAL',
           'Folder import/eksport tidak boleh kosong',
           'FOLDER IMPORT/EKSPORT:',
           'SEDANG DIJALANKAN',
           'DALAM PEMBAIKAN',
           'KELAYAKAN TIDAK SAH',
           'JULAT TARIKH TIDAK SAH',
           'PEMACU TIDAK SAH',
           'LALUAN TIDAK SAH',
           'Peranan pengguna tidak sah.',
           'NAMA PENGGUNA TIDAK SAH',
           'Nama pengguna atau kata laluan tidak sah.',
           'Julat tarikh waranti tidak sah: tarikh mula selepas tarikh tamat',
           'NOMBOR INVOIS',
           'MASALAH',
           'LOG MASUK TERAKHIR',
           'Fail .XLS lama tidak disokong. Simpan fail sebagai .XLSX atau CSV dahulu.',
           'DIKUNCI SEHINGGA',
           'LOG PERINGATAN',
           'LOG PERINGATAN',
           'LOG PERINGATAN WARANTI',
           'LOG MASUK',
           'Log masuk berjaya.',
           'LOG_MASUK_DISEKAT',
           'LOG_MASUK_GAGAL',
           'LOG_MASUK_BERJAYA',
           'LOG KELUAR',
           'ALAT GANTI STOK RENDAH',
           'TETINGKAP UTAMA DIBUKA',
           'Minimum 4 aksara. Sebarang format huruf, nombor, ruang atau simbol dibenarkan.',
           'MODEL PERANTI',
           'MODEL',
           'PENCIPTA DATA CONTOH PELBAGAI BAHASA DIBUKA',
           'AKTIVITI TERKINI SAYA',
           'TIDAK BERKENAAN',
           'NAMA PELANGGAN',
           'NAMA DIPERLUKAN',
           'NAMA:',
           'KATA LALUAN BAHARU',
           'TIADA DATA',
           'TIADA PANGKALAN DATA DIPILIH',
           'Tiada kata laluan lalai digunakan. Kata laluan disimpan hanya sebagai cincangan scrypt unik dengan salt. '
           'Simpan kata laluan pentadbir ini dengan selamat.',
           'Tiada pangkalan data sedia ada dipilih. Aplikasi akan meneruskan penciptaan pangkalan data baharu.',
           'Tiada mesej SMS atau WhatsApp dihantar.',
           'TIADA DATA STATUS',
           'TIADA WARANTI',
           'Tiada jaminan dalam apa-apa bentuk diberikan',
           'NOMBOR TELEFON',
           'TIDAK DITEMUI',
           'Buka pangkalan data sedia ada atau cipta pangkalan data baharu untuk aplikasi ini.',
           'BUKA KANDUNGAN BANTUAN APLIKASI',
           'BUKA EDITOR DATA CONTOH',
           'BUKA GOOGLE MAP',
           'BUKA BANTUAN UNTUK TAB AKTIF',
           'NAMA ALAT GANTI',
           'NAMA ALAT GANTI DIPERLUKAN',
           'JENIS ALAT GANTI',
           'ALAT GANTI DISEGAR SEMULA',
           'KATA LALUAN',
           'KATA LALUAN DITUKAR',
           'Kata laluan telah ditukar.',
           'Kata laluan terlalu panjang.',
           'KATA LALUAN TIDAK SEPADAN',
           'Kata laluan mesti mengandungi sekurang-kurangnya 4 aksara.',
           'KATA LALUAN TIDAK DITUKAR',
           'KATA LALUAN DITETAPKAN SEMULA',
           'KATA_LALUAN_DITUKAR',
           'PERTUKARAN_KATA_LALUAN_GAGAL',
           'KATA_LALUAN_DITETAPKAN_SEMULA',
           'NOMBOR TELEFON',
           'TELEFON:',
           'Letakkan DUMMY_CREATOR.EXE atau DUMMY_CREATOR.PY di sebelah aplikasi.',
           'Sila hubungi kami jika anda memerlukan perkhidmatan lanjut.',
           'Sila pilih fail XLSX atau CSV',
           'Sila pilih fail Excel atau CSV',
           'INVOIS PEMBELIAN:',
           'KUANTITI',
           'AKTIVITI KESELAMATAN DAN DATA TERKINI',
           'REKOD SERVIS TERKINI',
           'SEGAR SEMULA PAPAN PEMUKA',
           'MUAT SEMULA DATA SENARAI',
           'TEMPLAT MESEJ PERINGATAN',
           'PERINGATAN DILOG',
           'PERINGATAN DIHANTAR',
           'KOS PEMBAIKAN',
           'EKSPORT LAPORAN GAGAL',
           'LAPORAN DIEKSPORT',
           'TETAPKAN SEMULA KATA LALUAN',
           'TETAPKAN SEMULA KATA LALUAN PENGGUNA',
           'TETAPKAN SEMULA JULAT WARANTI DARI HARI INI',
           'PULIHKAN PANGKALAN DATA',
           'PEMULIHAN GAGAL',
           'PEMULIHAN BERJAYA',
           'PERANAN',
           'SIMPAN FAIL EKSPORT',
           'SIMPAN EKSPORT LAPORAN',
           'CARI MENGIKUT NAMA, TELEFON, PERANTI ATAU MASALAH...',
           'PERSEDIAAN PENTADBIR SELAMAT',
           'LOG MASUK SELAMAT',
           'Pilih hubungan untuk dipadam',
           'Pilih hubungan untuk disunting',
           'Pilih alat ganti untuk dipadam',
           'Pilih alat ganti untuk disunting',
           'Pilih alat ganti untuk dikemas kini',
           'Pilih pembekal untuk dipadam',
           'Pilih pembekal untuk disunting',
           'Pilih pembekal untuk dikemas kini',
           'Pilih pengguna dahulu.',
           'Pilih baris waranti untuk disunting',
           'PILIH MATA WANG APLIKASI:',
           'PILIH BAHASA DAN LOKAL APLIKASI:',
           'PILIH MATA WANG LALAI UNTUK APLIKASI:',
           'PILIH BAHASA / LOKAL',
           'PILIH BARIS UNTUK DIEKSPORT',
           'PILIH WARANTI UNTUK LOG PERINGATAN',
           'PILIH WARANTI UNTUK HANTAR PERINGATAN',
           'HANTAR PERINGATAN',
           'PENERANGAN SERVIS',
           'SERVIS KOMPUTER PANGGILAN BOGOR',
           'GRAF STATUS SERVIS',
           'RINGKASAN SERVIS',
           'LAPORAN RINGKASAN SERVIS',
           'Tetapkan kata laluan baharu untuk',
           'TETAPKAN LOKASI PANGKALAN DATA',
           'TETAPKAN TARIKH SERVIS KEPADA HARI INI',
           'TETAPAN DISIMPAN',
           'PERSEDIAAN GAGAL',
           'TUNJUKKAN KATA LALUAN',
           'TUNJUKKAN KATA LALUAN',
           'LOG MASUK UNTUK TERUSKAN',
           'TELAH LOG MASUK',
           'LOG MASUK SEBAGAI',
           'KAKITANGAN',
           'AKSES KAKITANGAN TERHAD',
           'Kakitangan boleh menambah rekod servis baharu dan melihat data operasi. Penyuntingan, pemadaman, import, '
           'eksport, sandaran, tetapan pentadbiran, pengurusan pengguna, perubahan pembekal, perubahan alat ganti, '
           'laporan dan peringatan memerlukan pentadbir.',
           'STATUS',
           'LOKASI SIMPANAN',
           'ALAMAT PEMBEKAL',
           'NAMA PEMBEKAL',
           'NAMA PEMBEKAL DIPERLUKAN',
           'PEMBEKAL DISEGAR SEMULA',
           'SISTEM',
           'NAMA JURUTEKNIK',
           'PRESTASI JURUTEKNIK',
           'LAPORAN PRESTASI JURUTEKNIK',
           'JURUTEKNIK',
           'Terima kasih,',
           'Nama pengguna itu sudah wujud.',
           'Pemacu untuk folder sandaran tidak wujud.',
           'Pemacu untuk laluan pangkalan data tidak wujud.',
           'Pemacu untuk laluan eksport tidak wujud.',
           'Pemacu untuk folder import/eksport tidak wujud.',
           'Kata laluan telah ditetapkan semula.',
           'Kata laluan tidak sepadan.',
           'Akaun ini dilumpuhkan. Hubungi pentadbir.',
           'Tindakan ini terhad kepada akaun pentadbir.',
           'Perisian ini disediakan sebagaimana adanya',
           'MASA',
           'KE EXCEL',
           'JUMLAH WARANTI AKTIF',
           'JUMLAH PELANGGAN',
           'JUMLAH PERANTI',
           'JUMLAH TAMAT TEMPOH',
           'JUMLAH AKAN TAMAT TIDAK LAMA LAGI',
           'JUMLAH HASIL',
           'JUMLAH DIPILIH',
           'JUMLAH REKOD SERVIS',
           'JUMLAH NILAI SERVIS',
           'JUMLAH SERVIS',
           'JENIS:',
           'MODEL UNIK',
           'KELAYAKAN TIDAK DIKENALI ATAU TIDAK SAH',
           'Jadual tidak dikenali untuk dieksport.',
           'BUKA KUNCI',
           'FORMAT TIDAK DISOKONG',
           'GUNAKAN D:\\BACKUP (LALAI)',
           'PENGGUNA',
           'PENGURUSAN PENGGUNA',
           'PENGGUNA TIDAK DICIPTA',
           'Pengguna tidak ditemui.',
           'PENGGUNA MEMINTA LOG KELUAR',
           'Status pengguna dikemas kini.',
           'NAMA PENGGUNA',
           'NAMA PENGGUNA SUDAH WUJUD',
           'Nama pengguna mesti 3-32 aksara dan menggunakan huruf, nombor, titik, sengkang atau garis bawah.',
           'PENGGUNA_DICIPTA',
           'PENGGUNA_DILUMPUHKAN',
           'PENGGUNA_DIDAYAKAN',
           'PENGGUNA_DIBUKA_KUNCI',
           'MENUNGGU ALAT GANTI',
           'MENUNGGU ALAT GANTI',
           'AMARAN: PEMULIHAN AKAN MENGGANTIKAN SEMUA DATA SEMASA DENGAN DATA SANDARAN.\n'
           'TINDAKAN INI TIDAK BOLEH DIBATALKAN.\n'
           '\n'
           'ADAKAH ANDA MAHU TERUSKAN?',
           'TEMPOH WARANTI',
           'STATUS WARANTI:',
           'RINGKASAN WARANTI',
           'LAPORAN RINGKASAN WARANTI',
           'KATA LALUAN LEMAH',
           'Apakah yang anda mahu lakukan dengan pangkalan data?',
           'TETINGKAP DITUTUP',
           'Anda tidak boleh melumpuhkan akaun semasa anda.',
           'Kata laluan anda berjaya ditukar.',
           '⚠️ AMARAN!',
           '✏️ SUNTING',
           '✏️ SUNTING HUBUNGAN',
           '❌ PADAM',
           '➕ TAMBAH',
           '➕ TAMBAH ALAT GANTI',
           '📄 PAPARAN TEKS',
           '📋 PAPARAN JADUAL',
           '📝 LOG PERINGATAN',
           '📝 LOG PERINGATAN',
           '📤 EKSPORT YANG DIPILIH',
           '🔃 SEGAR SEMULA',
           '🔄 SEGAR SEMULA',
           '🗑️ KOSONGKAN'],
 'fil_PH': ['Dapat gumawa ng bagong piyesa ang isang administrator.',
            'Dapat gumawa ng bagong supplier ang isang administrator.',
            'May piyesang ganito na ang pangalan. Piliin ito at gamitin ang I-save upang i-update.',
            'May supplier na ganito na ang pangalan. Piliin ito at gamitin ang I-save upang i-update.',
            'TINANGGIHAN_ANG_ACCESS',
            'ACCOUNT',
            'NA-DISABLE ANG ACCOUNT',
            'NA-LOCK ANG ACCOUNT',
            'HINARANG ANG PAGKILOS',
            'MGA AKTIBONG WARRANTY',
            'MAGDAGDAG NG USER NG APPLICATION',
            'MAGDAGDAG NG USER',
            'MGA KARAGDAGANG TALA',
            'KUMPLETONG ADDRESS',
            'ADMIN',
            'ADMINISTRATOR',
            'HINARANG ANG PAGKILOS NG ADMINISTRATOR',
            'PARA SA ADMINISTRATOR LAMANG',
            'KAILANGAN ANG ADMINISTRATOR',
            'KAILANGAN ANG ADMINISTRATOR.',
            'May ibang piyesang gumagamit na ng pangalang ito.',
            'May ibang supplier na gumagamit na ng pangalang ito.',
            'ISINARA ANG APPLICATION',
            'PAGLABAS_SA_APPLICATION',
            'BINUKSAN_ANG_APPLICATION',
            'Kailangang may kahit isang aktibong administrator.',
            'AVERAGE BAWAT SERBISYO',
            'AVERAGE NA GASTOS SA PAGKUKUMPUNI',
            'AVERAGE NA KITA',
            'AVERAGE NA KITA/SERBISYO',
            'I-BACKUP ANG DATABASE',
            'NABIGO ANG BACKUP',
            'Hindi maaaring walang laman ang backup folder',
            'BACKUP FOLDER:',
            'MATAGUMPAY ANG BACKUP',
            'Ipinapaalala ng Bogor On-Call Computer Service na mag-e-expire ang warranty sa pagkukumpuni ng iyong '
            'device sa:',
            'BRAND:',
            'MGA BRAND',
            'I-build ang HELP.HHP at kopyahin ang HELP.CHM sa application folder.',
            'Awtomatikong kinakalkula: petsa ng serbisyo + panahon ng warranty',
            'Hindi matukoy ang napiling supplier',
            'PALITAN ANG PASSWORD',
            'PALITAN ANG PASSWORD PARA KAY',
            'TAPOS NA',
            'COMPUTER SERVICE MANAGER',
            'KUMPIRMAHIN ANG PAG-DELETE',
            'KUMPIRMAHIN ANG PASSWORD',
            'KUMPIRMAHIN ANG PAG-RESTORE',
            'PANGALAN NG CONTACT',
            'CONTACT:<br><a href="mailto:e-comtech@mail.com">e-comtech@mail.com</a> / <a '
            'href="mailto:rahfie27@gmail.com">rahfie27@gmail.com</a>',
            'NAKOPYA SA CLIPBOARD',
            'KOPYAHIN ANG ADDRESS',
            'KOPYAHIN ANG GOOGLE MAP',
            'KOPYAHIN ANG INVOICE',
            'KOPYAHIN ANG TELEPONO',
            'COPYRIGHT © ECOMTECH 2026',
            'BILANG',
            'GUMAWA NG ADMINISTRATOR',
            'GUMAWA NG UNANG ADMINISTRATOR ACCOUNT',
            'NAGAWA',
            'GINAWA NI:',
            'PETSA NG PAGKAKAGAWA',
            'MGA CSV FILE',
            'KASALUKUYANG PASSWORD',
            'Mali ang kasalukuyang password.',
            'TULONG PARA SA KASALUKUYANG TAB',
            'PANGALAN NG CUSTOMER',
            'MGA CUSTOMER NA PINILI PARA SA PAALALA',
            'MGA CUSTOMER NA PINILI PARA SA PAALALA:',
            'Secure Computer Service Manager 4.4',
            'NABIGO ANG PAG-REFRESH NG DASHBOARD',
            'PANGKALAHATANG-TANAW NG DASHBOARD',
            'Na-export ang data sa:',
            'KAILANGAN ANG LOKASYON NG DATABASE',
            'Hindi maaaring walang laman ang database path',
            'DATABASE PATH:',
            'Matagumpay na na-update ang mga database path!',
            'NA-RESTORE ANG DATABASE MULA SA BACKUP',
            'Matagumpay na na-restore ang database mula sa backup!',
            'NA_RESTORE_ANG_DATABASE',
            'PETSA',
            'Minamahal na [CUSTOMER],',
            'MGA DETALYE',
            'ESTADISTIKA NG MGA BRAND NG DEVICE',
            'MODEL NG DEVICE',
            'MGA URI NG DEVICE',
            'DISPLAY NAME',
            'DONASYON:<br><a href="https://paypal.me/rahfie">paypal.me/rahfie</a>',
            'Hindi available ang drive D:. Pumili ng ibang folder.',
            'Hindi nakita ang drive D:.\nGumawa o pumili ng folder para sa database.',
            'HINDI NAKITA ANG DRIVE',
            'BINUKSAN ANG DUMMY DATA EDITOR',
            'DUPLIKADONG PIYESA',
            'DUPLIKADONG SUPPLIER',
            'EDIT MODE',
            'EDIT MODE: I-UPDATE ANG MGA FIELD AT I-CLICK ANG I-SAVE',
            'EDIT MODE: I-UPDATE ANG MGA FIELD AT I-CLICK ANG I-UPDATE',
            'I-ENABLE / I-DISABLE',
            'PETSA NG PAGTATAPOS:',
            'Ilagay ang username at password.',
            'MGA EXCEL FILE',
            'MAG-E-EXPIRE SA LOOB NG 30 ARAW',
            'MALAPIT NANG MAG-EXPIRE',
            'MALAPIT NANG MAG-EXPIRE (< 7 ARAW)',
            'MALAPIT NANG MAG-EXPIRE (≤ 30 ARAW)',
            'MALAPIT NANG MAG-EXPIRE (≤ 7 ARAW)',
            'MALAPIT NANG MAG-EXPIRE (≤30 araw)',
            'PETSA NG PAG-EXPIRE MULA',
            'PETSA NG PAG-EXPIRE MULA:',
            'TAPOS NA ANG PAG-EXPORT',
            'ERROR SA PAG-EXPORT',
            'NABIGO ANG PAG-EXPORT',
            'MATAGUMPAY ANG PAG-EXPORT',
            'NA-EXPORT SA:',
            'Nabigong mag-export',
            'ULAT PINANSYAL',
            'UNANG SESSION NG ADMINISTRATOR',
            'UNANG SETUP NG DATABASE',
            'UNANG SETUP:\nPumili ng folder kung saan ise-save ang bagong database.',
            'NALINIS ANG FORM',
            'Hindi maaaring mas huli ang petsa mula kaysa petsa hanggang',
            'KUMPLETONG ADDRESS',
            'GOOGLE MAP',
            'NAKOPYA ANG GOOGLE MAP LINK',
            'WALANG LAMAN ANG GOOGLE MAP LINK',
            'WALANG LAMAN ANG GOOGLE MAP LINK (NALINIS ANG CLIPBOARD)',
            'GOOGLE MAP:',
            'NILALAMAN NG TULONG',
            'ID (ROWID)',
            'TAPOS NA ANG PAG-IMPORT',
            'NABIGO ANG PAG-IMPORT',
            'Hindi maaaring walang laman ang import/export folder',
            'IMPORT/EXPORT FOLDER:',
            'ISINASAGAWA',
            'INAAYOS',
            'HINDI WASTONG CREDENTIALS',
            'HINDI WASTONG SAKLAW NG PETSA',
            'HINDI WASTONG DRIVE',
            'HINDI WASTONG PATH',
            'Hindi wastong tungkulin ng user.',
            'HINDI WASTONG USERNAME',
            'Hindi wastong username o password.',
            'Hindi wastong saklaw ng petsa ng warranty: mas huli ang petsa mula kaysa petsa hanggang',
            'NUMERO NG INVOICE',
            'MGA ISYU',
            'HULING PAG-LOGIN',
            'Hindi sinusuportahan ang mga lumang .XLS file. I-save muna ang file bilang .XLSX o CSV.',
            'NA-LOCK HANGGANG',
            'I-LOG ANG PAALALA',
            'I-LOG ANG MGA PAALALA',
            'I-LOG ANG PAALALA SA WARRANTY',
            'MAG-LOGIN',
            'Matagumpay ang pag-login.',
            'HINARANG_ANG_PAG_LOGIN',
            'NABIGO_ANG_PAG_LOGIN',
            'MATAGUMPAY_ANG_PAG_LOGIN',
            'MAG-LOGOUT',
            'MGA PIYESANG MABABA ANG STOCK',
            'BINUKSAN ANG MAIN WINDOW',
            'Hindi bababa sa 4 na character. Pinapayagan ang anumang kombinasyon ng letra, numero, espasyo, o simbolo.',
            'MODEL NG DEVICE',
            'MGA MODEL',
            'BINUKSAN ANG MULTILINGUAL DUMMY DATA CREATOR',
            'MGA KAMAKAILANG AKTIBIDAD KO',
            'HINDI NAAANGKOP',
            'PANGALAN NG CUSTOMER',
            'KAILANGAN ANG PANGALAN',
            'PANGALAN:',
            'BAGONG PASSWORD',
            'WALANG DATA',
            'WALANG NAPILING DATABASE',
            'Walang ginagamit na default na password. Ang mga password ay iniimbak lamang bilang natatangi at may salt '
            'na scrypt hash. Panatilihing ligtas ang password ng administrator na ito.',
            'Walang napiling umiiral na database. Magpapatuloy ang application sa paggawa ng bagong database.',
            'Walang ipinadalang SMS o WhatsApp message.',
            'WALANG STATUS DATA',
            'WALANG WARRANTY',
            'Walang anumang uri ng garantiya ang ibinibigay',
            'NUMERO NG TELEPONO',
            'HINDI NATAGPUAN',
            'Magbukas ng umiiral na database o gumawa ng bagong database para sa application na ito.',
            'BUKSAN ANG NILALAMAN NG TULONG NG APPLICATION',
            'BUKSAN ANG DUMMY DATA EDITOR',
            'BUKSAN ANG GOOGLE MAP',
            'BUKSAN ANG TULONG PARA SA AKTIBONG TAB',
            'PANGALAN NG PIYESA',
            'KAILANGAN ANG PANGALAN NG PIYESA',
            'MGA URI NG PIYESA',
            'NA-REFRESH ANG MGA PIYESA',
            'PASSWORD',
            'NAPALITAN ANG PASSWORD',
            'Napalitan ang password.',
            'Masyadong mahaba ang password.',
            'HINDI MAGKATUGMA ANG MGA PASSWORD',
            'Dapat may hindi bababa sa 4 na character ang password.',
            'HINDI NAPALITAN ANG PASSWORD',
            'NA-RESET ANG PASSWORD',
            'NAPALITAN_ANG_PASSWORD',
            'NABIGO_ANG_PAGPALIT_NG_PASSWORD',
            'NA_RESET_ANG_PASSWORD',
            'NUMERO NG TELEPONO',
            'TELEPONO:',
            'Ilagay ang DUMMY_CREATOR.EXE o DUMMY_CREATOR.PY sa tabi ng application.',
            'Makipag-ugnayan sa amin kung kailangan mo ng karagdagang serbisyo.',
            'Pumili ng XLSX o CSV file',
            'Pumili ng Excel o CSV file',
            'INVOICE NG PAGBILI:',
            'DAMI',
            'KAMAKAILANG AKTIBIDAD SA SEGURIDAD AT DATA',
            'MGA KAMAKAILANG RECORD NG SERBISYO',
            'I-REFRESH ANG DASHBOARD',
            'I-RELOAD ANG DATA NG MGA DROPDOWN',
            'TEMPLATE NG MENSAHE NG PAALALA',
            'NA-LOG ANG MGA PAALALA',
            'NAIPADALA ANG MGA PAALALA',
            'GASTOS SA PAGKUKUMPUNI',
            'NABIGO ANG PAG-EXPORT NG ULAT',
            'NA-EXPORT ANG ULAT',
            'I-RESET ANG PASSWORD',
            'I-RESET ANG PASSWORD NG USER',
            'I-RESET ANG SAKLAW NG WARRANTY MULA NGAYON',
            'I-RESTORE ANG DATABASE',
            'NABIGO ANG PAG-RESTORE',
            'MATAGUMPAY ANG PAG-RESTORE',
            'TUNGKULIN',
            'I-SAVE ANG EXPORT FILE',
            'I-SAVE ANG PAG-EXPORT NG ULAT',
            'MAGHANAP AYON SA PANGALAN, TELEPONO, DEVICE, O ISYU...',
            'LIGTAS NA SETUP NG ADMINISTRATOR',
            'LIGTAS NA PAG-LOGIN',
            'Pumili ng contact na ide-delete',
            'Pumili ng contact na ie-edit',
            'Pumili ng piyesang ide-delete',
            'Pumili ng piyesang ie-edit',
            'Pumili ng piyesang ia-update',
            'Pumili ng supplier na ide-delete',
            'Pumili ng supplier na ie-edit',
            'Pumili ng supplier na ia-update',
            'Pumili muna ng user.',
            'Pumili ng warranty row na ie-edit',
            'PILIIN ANG CURRENCY NG APPLICATION:',
            'PILIIN ANG WIKA AT LOCALE NG APPLICATION:',
            'PILIIN ANG DEFAULT NA CURRENCY NG APPLICATION:',
            'PILIIN ANG WIKA / LOCALE',
            'PILIIN ANG MGA ROW NA IE-EXPORT',
            'PILIIN ANG MGA WARRANTY NA LALAGYAN NG PAALALA',
            'PILIIN ANG MGA WARRANTY NA PADADALHAN NG PAALALA',
            'IPADALA ANG MGA PAALALA',
            'PAGLALARAWAN NG SERBISYO',
            'BOGOR ON-CALL COMPUTER SERVICE',
            'GRAPH NG STATUS NG SERBISYO',
            'BUOD NG SERBISYO',
            'ULAT NG BUOD NG SERBISYO',
            'Magtakda ng bagong password para kay',
            'ITAKDA ANG LOKASYON NG DATABASE',
            'ITAKDA ANG PETSA NG SERBISYO SA NGAYON',
            'NA-SAVE ANG MGA SETTING',
            'NABIGO ANG SETUP',
            'IPAKITA ANG PASSWORD',
            'IPAKITA ANG MGA PASSWORD',
            'MAG-SIGN IN UPANG MAGPATULOY',
            'NAKA-SIGN IN',
            'NAKA-SIGN IN BILANG',
            'STAFF',
            'LIMITADO ANG ACCESS NG STAFF',
            'Maaaring magdagdag ang staff ng mga bagong record ng serbisyo at tingnan ang operational data. Ang '
            'pag-edit, pag-delete, pag-import, pag-export, backup, mga setting ng administrator, pamamahala ng user, '
            'mga pagbabago sa supplier, mga pagbabago sa piyesa, mga ulat, at mga paalala ay nangangailangan ng '
            'administrator.',
            'MGA STATUS',
            'LOKASYON NG IMBAKAN',
            'ADDRESS NG SUPPLIER',
            'PANGALAN NG SUPPLIER',
            'KAILANGAN ANG PANGALAN NG SUPPLIER',
            'NA-REFRESH ANG MGA SUPPLIER',
            'SYSTEM',
            'PANGALAN NG TECHNICIAN',
            'PAGGANAP NG MGA TECHNICIAN',
            'ULAT NG PAGGANAP NG MGA TECHNICIAN',
            'MGA TECHNICIAN',
            'Salamat,',
            'May ganitong username na.',
            'Walang drive para sa backup folder.',
            'Walang drive para sa database path.',
            'Walang drive para sa export path.',
            'Walang drive para sa import/export folder.',
            'Na-reset ang password.',
            'Hindi magkatugma ang mga password.',
            'Naka-disable ang account na ito. Makipag-ugnayan sa administrator.',
            'Para lamang sa mga administrator account ang pagkilos na ito.',
            'Ibinibigay ang software na ito nang walang anumang garantiya',
            'ORAS',
            'SA EXCEL',
            'KABUUANG AKTIBONG WARRANTY',
            'KABUUANG CUSTOMER',
            'KABUUANG DEVICE',
            'KABUUANG NAG-EXPIRE',
            'KABUUANG MALAPIT NANG MAG-EXPIRE',
            'KABUUANG KITA',
            'KABUUANG NAPILI',
            'KABUUANG RECORD NG SERBISYO',
            'KABUUANG HALAGA NG SERBISYO',
            'KABUUANG SERBISYO',
            'URI:',
            'MGA NATATANGING MODEL',
            'HINDI KILALA O HINDI WASTONG CREDENTIALS',
            'Hindi kilalang table para sa pag-export.',
            'I-UNLOCK',
            'HINDI SINUSUPORTAHANG FORMAT',
            'GAMITIN ANG D:\\BACKUP (DEFAULT)',
            'USER',
            'PAMAMAHALA NG USER',
            'HINDI NAGAWA ANG USER',
            'Hindi natagpuan ang user.',
            'HINILING NG USER ANG PAG-LOGOUT',
            'Na-update ang status ng user.',
            'USERNAME',
            'UMIIRAL NA ANG USERNAME',
            'Dapat 3-32 character ang username at gumamit lamang ng mga letra, numero, tuldok, gitling, o underscore.',
            'NAGAWA_ANG_USER',
            'NA_DISABLE_ANG_USER',
            'NA_ENABLE_ANG_USER',
            'NA_UNLOCK_ANG_USER',
            'NAGHIHINTAY NG MGA PIYESA',
            'NAGHIHINTAY NG PIYESA',
            'BABALA: PAPALITAN NG PAG-RESTORE ANG LAHAT NG KASALUKUYANG DATA NG DATA MULA SA BACKUP.\n'
            'HINDI MAAARING I-UNDO ANG PAGKILOS NA ITO.\n'
            '\n'
            'GUSTO MO BANG MAGPATULOY?',
            'MGA PANAHON NG WARRANTY',
            'STATUS NG WARRANTY:',
            'BUOD NG WARRANTY',
            'ULAT NG BUOD NG WARRANTY',
            'MAHINANG PASSWORD',
            'Ano ang gusto mong gawin sa database?',
            'ISINARA ANG WINDOW',
            'Hindi mo maaaring i-disable ang kasalukuyan mong account.',
            'Matagumpay na napalitan ang iyong password.',
            '⚠️ BABALA!',
            '✏️ I-EDIT',
            '✏️ I-EDIT ANG CONTACT',
            '❌ I-DELETE',
            '➕ IDAGDAG',
            '➕ MAGDAGDAG NG PIYESA',
            '📄 TEXT VIEW',
            '📋 TABLE VIEW',
            '📝 I-LOG ANG PAALALA',
            '📝 I-LOG ANG MGA PAALALA',
            '📤 I-EXPORT ANG NAPILI',
            '🔃 I-REFRESH',
            '🔄 I-REFRESH',
            '🗑️ LINISIN']}
for _language_code, _translated_values in _UI_TRANSLATION_COMPLETIONS.items():
    _language_table = UI_TRANSLATIONS.setdefault(_language_code, {})
    for _source_key, _translated_value in zip(_UI_TRANSLATION_COMPLETION_KEYS, _translated_values):
        _language_table.setdefault(_source_key, _translated_value)
UI_TRANSLATIONS.setdefault("id_ID", {}).update({'APPLICATION CLOSED': 'APLIKASI DITUTUP',
 'BRAND:': 'MEREK:',
 'BRANDS': 'MEREK',
 'DEVICE TYPES': 'JENIS PERANGKAT',
 'ISSUES': 'MASALAH',
 'MODELS': 'MODEL',
 'PART TYPES': 'JENIS SUKU CADANG',
 'STATUSES': 'STATUS',
 'TECHNICIANS': 'TEKNISI',
 'WARRANTY PERIODS': 'PERIODE GARANSI',
 '❌ DELETE': '❌ HAPUS',
 '➕ ADD': '➕ TAMBAH',
 '➕ ADD PART': '➕ TAMBAH SUKU CADANG',
 '📄 TEXT VIEW': '📄 TAMPILAN TEKS',
 '📋 TABLE VIEW': '📋 TAMPILAN TABEL',
 '📤 EXPORT SELECTED': '📤 EKSPOR TERPILIH',
 '🔃 REFRESH': '🔃 MUAT ULANG',
 '🔄 REFRESH': '🔄 MUAT ULANG',
 '🗑️ CLEAR': '🗑️ BERSIHKAN'})
_UI_TRANSLATION_EXTRAS = {'de_DE': {'QUICK': 'SCHNELL'},
 'pt_BR': {'QUICK': 'RÁPIDO'},
 'it_IT': {'ALL FILES': 'TUTTI I FILE',
           'CHOOSE': 'SCEGLI',
           'CURRENT DATE': 'DATA CORRENTE',
           'DATABASE FILES': 'FILE DI DATABASE',
           'DUPLICATE': 'DUPLICATO',
           'EXISTS': 'ESISTE',
           'FIRST-TIME': 'PRIMO AVVIO',
           'LINK': 'COLLEGAMENTO',
           'LOGGED': 'REGISTRATO',
           'PROTECTED': 'PROTETTO',
           'QUICK': 'RAPIDO',
           'REPAIRED': 'RIPARATO',
           'UNIQUE': 'UNIVOCO'},
 'nl_NL': {'ALL FILES': 'ALLE BESTANDEN',
           'CHOOSE': 'KIEZEN',
           'CURRENT DATE': 'HUIDIGE DATUM',
           'DATABASE FILES': 'DATABASEBESTANDEN',
           'DUPLICATE': 'DUBBEL',
           'EXISTS': 'BESTAAT',
           'FIRST-TIME': 'EERSTE KEER',
           'LINK': 'KOPPELING',
           'LOGGED': 'GEREGISTREERD',
           'PROTECTED': 'BEVEILIGD',
           'QUICK': 'SNEL',
           'REPAIRED': 'GEREPAREERD',
           'UNIQUE': 'UNIEK'},
 'zh_CN': {'ALL FILES': '所有文件',
           'CHOOSE': '选择',
           'CURRENT DATE': '当前日期',
           'DATABASE FILES': '数据库文件',
           'DUPLICATE': '重复',
           'EXISTS': '已存在',
           'FIRST-TIME': '首次',
           'LINK': '链接',
           'LOGGED': '已记录',
           'PROTECTED': '受保护',
           'QUICK': '快速',
           'REPAIRED': '已修复',
           'UNIQUE': '唯一'},
 'ja_JP': {'ALL FILES': 'すべてのファイル',
           'CHOOSE': '選択',
           'CURRENT DATE': '現在の日付',
           'DATABASE FILES': 'データベースファイル',
           'DUPLICATE': '重複',
           'EXISTS': '存在します',
           'FIRST-TIME': '初回',
           'LINK': 'リンク',
           'LOGGED': '記録済み',
           'PROTECTED': '保護済み',
           'QUICK': 'クイック',
           'REPAIRED': '修理済み',
           'UNIQUE': '固有'},
 'ko_KR': {'ALL FILES': '모든 파일',
           'CHOOSE': '선택',
           'CURRENT DATE': '현재 날짜',
           'DATABASE FILES': '데이터베이스 파일',
           'DUPLICATE': '중복',
           'EXISTS': '존재함',
           'FIRST-TIME': '최초',
           'LINK': '링크',
           'LOGGED': '기록됨',
           'PROTECTED': '보호됨',
           'QUICK': '빠른',
           'REPAIRED': '수리 완료',
           'UNIQUE': '고유'},
 'ar_SA': {'ALL FILES': 'كل الملفات',
           'CHOOSE': 'اختيار',
           'CURRENT DATE': 'التاريخ الحالي',
           'DATABASE FILES': 'ملفات قاعدة البيانات',
           'DUPLICATE': 'مكرر',
           'EXISTS': 'موجود',
           'FIRST-TIME': 'المرة الأولى',
           'LINK': 'رابط',
           'LOGGED': 'مسجل',
           'PROTECTED': 'محمي',
           'QUICK': 'سريع',
           'REPAIRED': 'تم الإصلاح',
           'UNIQUE': 'فريد'},
 'hi_IN': {'ALL FILES': 'सभी फ़ाइलें',
           'CHOOSE': 'चुनें',
           'CURRENT DATE': 'वर्तमान तिथि',
           'DATABASE FILES': 'डेटाबेस फ़ाइलें',
           'DUPLICATE': 'डुप्लिकेट',
           'EXISTS': 'मौजूद है',
           'FIRST-TIME': 'पहली बार',
           'LINK': 'लिंक',
           'LOGGED': 'लॉग किया गया',
           'PROTECTED': 'सुरक्षित',
           'QUICK': 'त्वरित',
           'REPAIRED': 'मरम्मत किया गया',
           'UNIQUE': 'अद्वितीय'},
 'bn_BD': {'ALL FILES': 'সব ফাইল',
           'CHOOSE': 'নির্বাচন করুন',
           'CURRENT DATE': 'বর্তমান তারিখ',
           'DATABASE FILES': 'ডেটাবেস ফাইল',
           'DUPLICATE': 'ডুপ্লিকেট',
           'EXISTS': 'বিদ্যমান',
           'FIRST-TIME': 'প্রথমবার',
           'LINK': 'লিংক',
           'LOGGED': 'লগ করা হয়েছে',
           'PROTECTED': 'সুরক্ষিত',
           'QUICK': 'দ্রুত',
           'REPAIRED': 'মেরামত করা হয়েছে',
           'UNIQUE': 'অনন্য'},
 'ru_RU': {'ALL FILES': 'ВСЕ ФАЙЛЫ',
           'CHOOSE': 'ВЫБРАТЬ',
           'CURRENT DATE': 'ТЕКУЩАЯ ДАТА',
           'DATABASE FILES': 'ФАЙЛЫ БАЗЫ ДАННЫХ',
           'DUPLICATE': 'ДУБЛИКАТ',
           'EXISTS': 'СУЩЕСТВУЕТ',
           'FIRST-TIME': 'ПЕРВЫЙ ЗАПУСК',
           'LINK': 'ССЫЛКА',
           'LOGGED': 'ЗАПИСАНО',
           'PROTECTED': 'ЗАЩИЩЕНО',
           'QUICK': 'БЫСТРО',
           'REPAIRED': 'ОТРЕМОНТИРОВАНО',
           'UNIQUE': 'УНИКАЛЬНЫЙ'},
 'tr_TR': {'A': 'BİR',
           'ACTION': 'EYLEM',
           'ALL FILES': 'TÜM DOSYALAR',
           'AN': 'BİR',
           'ANOTHER': 'BAŞKA',
           'ARE': 'BULUNUYOR',
           'CHOOSE': 'SEÇ',
           'CONTINUE': 'DEVAM ET',
           'CURRENT DATE': 'GEÇERLİ TARİH',
           'DATABASE FILES': 'VERİTABANI DOSYALARI',
           'DUPLICATE': 'YİNELENEN',
           'EXISTS': 'VAR',
           'FIRST-TIME': 'İLK KEZ',
           'ITEM': 'ÖĞE',
           'LINK': 'BAĞLANTI',
           'LOGGED': 'KAYDEDİLDİ',
           'PROTECTED': 'KORUMALI',
           'QUICK': 'HIZLI',
           'REPAIRED': 'ONARILDI',
           'REPLACE': 'DEĞİŞTİR',
           'RESTORING': 'GERİ YÜKLENİYOR',
           'SEND': 'GÖNDER',
           'THE': 'BU',
           'UNDONE': 'GERİ ALINMIŞ',
           'UNIQUE': 'BENZERSİZ',
           'WILL': 'OLACAK'},
 'vi_VN': {'A': 'MỘT',
           'ACTION': 'HÀNH ĐỘNG',
           'ALL FILES': 'TẤT CẢ TỆP',
           'AN': 'MỘT',
           'ANOTHER': 'KHÁC',
           'ARE': 'LÀ',
           'CHOOSE': 'CHỌN',
           'CONTINUE': 'TIẾP TỤC',
           'CURRENT DATE': 'NGÀY HIỆN TẠI',
           'DATABASE FILES': 'TỆP CƠ SỞ DỮ LIỆU',
           'DO': 'THỰC HIỆN',
           'DUPLICATE': 'TRÙNG',
           'EXISTS': 'TỒN TẠI',
           'FIRST-TIME': 'LẦN ĐẦU',
           'ITEM': 'MỤC',
           'LINK': 'LIÊN KẾT',
           'LOGGED': 'ĐÃ GHI NHẬT KÝ',
           'PROTECTED': 'ĐƯỢC BẢO VỆ',
           'QUICK': 'NHANH',
           'REPAIRED': 'ĐÃ SỬA',
           'REPLACE': 'THAY THẾ',
           'RESTORING': 'ĐANG KHÔI PHỤC',
           'SEND': 'GỬI',
           'THE': 'ĐÓ',
           'UNDONE': 'HOÀN TÁC',
           'UNIQUE': 'DUY NHẤT',
           'WILL': 'SẼ'},
 'th_TH': {'A': 'หนึ่ง',
           'ACTION': 'การดำเนินการ',
           'ALL FILES': 'ไฟล์ทั้งหมด',
           'AN': 'หนึ่ง',
           'ANOTHER': 'อื่น',
           'ARE': 'เป็น',
           'CHOOSE': 'เลือก',
           'CONTINUE': 'ดำเนินการต่อ',
           'CURRENT DATE': 'วันที่ปัจจุบัน',
           'DATABASE FILES': 'ไฟล์ฐานข้อมูล',
           'DO': 'ทำ',
           'DUPLICATE': 'ซ้ำ',
           'EXISTS': 'มีอยู่',
           'FIRST-TIME': 'ครั้งแรก',
           'ITEM': 'รายการ',
           'LINK': 'ลิงก์',
           'LOGGED': 'บันทึกแล้ว',
           'PROTECTED': 'มีการป้องกัน',
           'QUICK': 'ด่วน',
           'REPAIRED': 'ซ่อมแล้ว',
           'REPLACE': 'แทนที่',
           'RESTORING': 'กำลังคืนค่า',
           'SEND': 'ส่ง',
           'THE': 'นั้น',
           'UNDONE': 'ย้อนกลับ',
           'UNIQUE': 'ไม่ซ้ำกัน',
           'WILL': 'จะ'},
 'ms_MY': {'A': 'SATU',
           'ACTION': 'TINDAKAN',
           'ALL FILES': 'SEMUA FAIL',
           'AN': 'SATU',
           'ANOTHER': 'LAIN',
           'ARE': 'ADALAH',
           'CHOOSE': 'PILIH',
           'CONTINUE': 'TERUSKAN',
           'CURRENT DATE': 'TARIKH SEMASA',
           'DATABASE FILES': 'FAIL PANGKALAN DATA',
           'DO': 'LAKUKAN',
           'DUPLICATE': 'PENDUA',
           'EXISTS': 'WUJUD',
           'FIRST-TIME': 'KALI PERTAMA',
           'ITEM': 'ITEM',
           'LINK': 'PAUTAN',
           'LOGGED': 'DILOG',
           'PROTECTED': 'DILINDUNGI',
           'QUICK': 'PANTAS',
           'REPAIRED': 'DIBAIKI',
           'REPLACE': 'GANTIKAN',
           'RESTORING': 'SEDANG MEMULIHKAN',
           'SEND': 'HANTAR',
           'THE': 'ITU',
           'UNDONE': 'DIBATALKAN',
           'UNIQUE': 'UNIK',
           'WILL': 'AKAN'},
 'fil_PH': {'A': 'ISANG',
            'ACTION': 'PAGKILOS',
            'ALL FILES': 'LAHAT NG FILE',
            'AN': 'ISANG',
            'ANOTHER': 'IBA',
            'ARE': 'AY',
            'CHOOSE': 'PUMILI',
            'CONTINUE': 'MAGPATULOY',
            'CURRENT DATE': 'KASALUKUYANG PETSA',
            'DATABASE FILES': 'MGA DATABASE FILE',
            'DO': 'GAWIN',
            'DUPLICATE': 'DUPLIKADO',
            'EXISTS': 'UMIIRAL',
            'FIRST-TIME': 'UNANG BESES',
            'ITEM': 'ITEM',
            'LINK': 'LINK',
            'LOGGED': 'NA-LOG',
            'PROTECTED': 'PROTEKTADO',
            'QUICK': 'MABILIS',
            'REPAIRED': 'NAAYOS',
            'REPLACE': 'PALITAN',
            'RESTORING': 'NAGRE-RESTORE',
            'SEND': 'IPADALA',
            'THE': 'ANG',
            'UNDONE': 'I-UNDO',
            'UNIQUE': 'NATATANGI',
            'WILL': 'GAGAWIN'}}
for _language_code, _extra_items in _UI_TRANSLATION_EXTRAS.items():
    UI_TRANSLATIONS.setdefault(_language_code, {}).update(_extra_items)
del _extra_items
del _language_code, _translated_values, _language_table, _source_key, _translated_value


# Dynamic Indonesian messages. Captured values (paths, names, counts, errors,
# dates, and currency text) are deliberately left untouched.
UI_TRANSLATION_PATTERNS = {
    "id_ID": [
        (r"FAILED TO CREATE ADMINISTRATOR:\n(.+)", r"GAGAL MEMBUAT ADMINISTRATOR:\n\1"),
        (r"ACCOUNT LOCKED\. TRY AGAIN IN ABOUT (\d+) MINUTE\(S\)\.", r"AKUN TERKUNCI. COBA LAGI DALAM SEKITAR \1 MENIT."),
        (r"TOO MANY FAILED ATTEMPTS\. ACCOUNT LOCKED FOR (\d+) MINUTES\.", r"TERLALU BANYAK PERCOBAAN GAGAL. AKUN DIKUNCI SELAMA \1 MENIT."),
        (r"INVALID USERNAME OR PASSWORD\. (\d+) ATTEMPT\(S\) REMAIN BEFORE LOCKOUT\.", r"NAMA PENGGUNA ATAU KATA SANDI SALAH. TERSISA \1 PERCOBAAN SEBELUM AKUN DIKUNCI."),
        (r"DROPDOWN DATA RELOADED: (.+)", r"DATA PILIHAN DIMUAT ULANG: \1"),
        (r"DUMMY\.INI NOT FOUND - RUN DUMMY_DATA_EDITOR\.PY: (.+)", r"DUMMY.INI TIDAK DITEMUKAN - JALANKAN DUMMY_DATA_EDITOR.PY: \1"),
        (r"FAILED TO IMPORT DATA:\n(.+)", r"GAGAL MENGIMPOR DATA:\n\1"),
        (r"EXPORT ERROR: (.+)", r"KESALAHAN EKSPOR: \1"),
        (r"DATA HAS BEEN EXPORTED TO:\n(.+)", r"DATA BERHASIL DIEKSPOR KE:\n\1"),
        (r"BACKED UP TO (.+)", r"BERHASIL DICADANGKAN KE \1"),
        (r"EXPORT FAILED: (.+)", r"EKSPOR GAGAL: \1"),
        (r"DATABASE BACKED UP SUCCESSFULLY TO:\n(.+)", r"DATABASE BERHASIL DICADANGKAN KE:\n\1"),
        (r"FAILED TO BACKUP DATABASE:\n(.+)", r"GAGAL MENCADANGKAN DATABASE:\n\1"),
        (r"FAILED TO RESTORE DATABASE:\n(.+)", r"GAGAL MEMULIHKAN DATABASE:\n\1"),
        (r"SUCCESSFULLY IMPORTED (\d+) RECORDS FROM EXCEL\n(\d+) ROWS SKIPPED", r"BERHASIL MENGIMPOR \1 DATA DARI EXCEL\n\2 BARIS DILEWATI"),
        (r"SUCCESSFULLY IMPORTED (\d+) RECORDS FROM CSV\n(\d+) ROWS SKIPPED", r"BERHASIL MENGIMPOR \1 DATA DARI CSV\n\2 BARIS DILEWATI"),
        (r"IMPORTED (\d+) RECORDS", r"\1 DATA BERHASIL DIIMPOR"),
        (r"DATA EXPORTED SUCCESSFULLY TO:\n(.+)", r"DATA BERHASIL DIEKSPOR KE:\n\1"),
        (r"EXPORTED TO (.+)", r"BERHASIL DIEKSPOR KE \1"),
        (r"FAILED TO EXPORT DATA:\n(.+)", r"GAGAL MENGEKSPOR DATA:\n\1"),
        (r"SELECTED DATA EXPORTED TO:\n(.+)", r"DATA TERPILIH BERHASIL DIEKSPOR KE:\n\1"),
        (r"EXPORTED (\d+) ROWS", r"\1 BARIS BERHASIL DIEKSPOR"),
        (r"REPORT EXPORTED TO:\n(.+)", r"LAPORAN BERHASIL DIEKSPOR KE:\n\1"),
        (r"REPORT EXPORTED TO (.+)", r"LAPORAN BERHASIL DIEKSPOR KE \1"),
        (r"FAILED TO EXPORT REPORT:\n(.+)", r"GAGAL MENGEKSPOR LAPORAN:\n\1"),
        (r"WARNING: FAILED TO REGISTER SUPPLIER/PART: (.+)", r"PERINGATAN: GAGAL MENDAFTARKAN PEMASOK/SUKU CADANG: \1"),
        (r"OPENED CONTACT FROM WARRANTY: (.+)", r"KONTAK DARI DATA GARANSI DIBUKA: \1"),
        (r"CONTACT NOT FOUND IN TABLE: (.+)", r"KONTAK TIDAK DITEMUKAN DI TABEL: \1"),
        (r"SELECTED SUPPLIER: (.+)", r"PEMASOK TERPILIH: \1"),
        (r"UPDATED SUPPLIER: (.+)", r"PEMASOK BERHASIL DIPERBARUI: \1"),
        (r"DELETE SUPPLIER '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!", r"HAPUS PEMASOK '\1'?\n\nTINDAKAN INI TIDAK DAPAT DIBATALKAN!"),
        (r"DELETED SUPPLIER: (.+)", r"PEMASOK BERHASIL DIHAPUS: \1"),
        (r"ADDED PART: (.+)", r"SUKU CADANG BERHASIL DITAMBAHKAN: \1"),
        (r"UPDATED PART: (.+)", r"SUKU CADANG BERHASIL DIPERBARUI: \1"),
        (r"DELETE PART '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!", r"HAPUS SUKU CADANG '\1'?\n\nTINDAKAN INI TIDAK DAPAT DIBATALKAN!"),
        (r"DELETED PART: (.+)", r"SUKU CADANG BERHASIL DIHAPUS: \1"),
        (r"DELETE '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!", r"HAPUS '\1'?\n\nTINDAKAN INI TIDAK DAPAT DIBATALKAN!"),
        (r"DELETED (.+)", r"\1 BERHASIL DIHAPUS"),
        (r"ADDED (.+)", r"\1 BERHASIL DITAMBAHKAN"),
        (r"REMINDERS SENT TO (\d+) CUSTOMERS\n\nREMINDERS HAVE BEEN LOGGED IN THE DATABASE\.", r"PENGINGAT TERKIRIM KEPADA \1 PELANGGAN\n\nRIWAYAT PENGINGAT TELAH DISIMPAN DI DATABASE."),
        (r"REMINDERS SENT TO (\d+) CUSTOMERS", r"PENGINGAT TERKIRIM KEPADA \1 PELANGGAN"),
        (r"⚠️ WARNING: (\d+) PARTS HAVE LOW STOCK!", r"⚠️ PERINGATAN: STOK \1 SUKU CADANG MENIPIS!"),
        (r"(-?\d+) \(EXPIRED\)", r"\1 (KEDALUWARSA)"),
        (r"(-?\d+) \(SOON\)", r"\1 (SEGERA)"),
        (r"(-?\d+) \(WARNING\)", r"\1 (PERINGATAN)"),
        (r"PERIOD: (.+) TO (.+)", r"PERIODE: \1 SAMPAI \2"),
        (r"GENERATED: (.+)", r"DIBUAT: \1"),
        (r"CURRENT DATE: (.+)", r"TANGGAL HARI INI: \1"),
        (r"TOTAL SERVICES: (.+)", r"TOTAL SERVIS: \1"),
        (r"TOTAL REVENUE: (.+)", r"TOTAL PENDAPATAN: \1"),
        (r"TOTAL ACTIVE WARRANTIES: (.+)", r"TOTAL GARANSI AKTIF: \1"),
        (r"TOTAL EXPIRING SOON: (.+)", r"TOTAL SEGERA BERAKHIR: \1"),
        (r"TOTAL EXPIRED: (.+)", r"TOTAL KEDALUWARSA: \1"),
        (r"COMPLETED: (.+)", r"SELESAI: \1"),
        (r"IN REPAIR: (.+)", r"DALAM PERBAIKAN: \1"),
        (r"IN PROGRESS: (.+)", r"SEDANG DIPROSES: \1"),
        (r"WAITING FOR PARTS: (.+)", r"MENUNGGU SUKU CADANG: \1"),
        (r"REVENUE: (.+)", r"PENDAPATAN: \1"),
        (r"AVG PER SERVICE: (.+)", r"RATA-RATA PER SERVIS: \1"),
        (r"TOTAL DEVICES: (.+)", r"TOTAL PERANGKAT: \1"),
        (r"UNIQUE MODELS: (.+)", r"MODEL UNIK: \1"),
        (r"REPAIRED: (.+)", r"SELESAI DIPERBAIKI: \1"),
        (r"PENDING: (.+)", r"TERTUNDA: \1"),
        (r"AVG REPAIR COST: (.+)", r"RATA-RATA BIAYA PERBAIKAN: \1"),
        (r"TOTAL CUSTOMERS: (.+)", r"TOTAL PELANGGAN: \1"),
        (r"ACTIVE: (.+)", r"AKTIF: \1"),
        (r"EXPIRING SOON: (.+)", r"SEGERA BERAKHIR: \1"),
        (r"EXPIRED: (.+)", r"KEDALUWARSA: \1"),
    ]
}

# Additional dynamic phrases used by status bars, export helpers, account logs,
# dashboard audit rows, and reminder logging.
UI_TRANSLATION_PATTERNS.setdefault("id_ID", []).extend([
    (r"SIGNED IN AS (.+)", r"MASUK SEBAGAI \1"),
    (r"DUMMY\.INI NOT FOUND - OPEN SETTINGS > EDIT DUMMY DATA: (.+)", r"DUMMY.INI TIDAK DITEMUKAN - BUKA PENGATURAN > UBAH DATA CONTOH: \1"),
    (r"DASHBOARD REFRESH FAILED: (.+)", r"PEMUATAN ULANG DASBOR GAGAL: \1"),
    (r"REMINDERS LOGGED FOR (\d+) CUSTOMERS\.\n\nNO SMS OR WHATSAPP MESSAGE WAS SENT\.", r"PENGINGAT UNTUK \1 PELANGGAN BERHASIL DICATAT.\n\nTIDAK ADA PESAN SMS ATAU WHATSAPP YANG DIKIRIM."),
    (r"REMINDERS LOGGED FOR (\d+) CUSTOMERS", r"PENGINGAT UNTUK \1 PELANGGAN BERHASIL DICATAT"),
    (r"EXPORTED TO:\n(.+)", r"BERHASIL DIEKSPOR KE:\n\1"),
    (r"EXPORTED SUPPLIERS TO EXCEL", r"DATA PEMASOK BERHASIL DIEKSPOR KE EXCEL"),
    (r"EXPORTED SPAREPARTS TO EXCEL", r"DATA SUKU CADANG BERHASIL DIEKSPOR KE EXCEL"),
    (r"FAILED TO EXPORT SUPPLIERS:\n(.+)", r"GAGAL MENGEKSPOR DATA PEMASOK:\n\1"),
    (r"FAILED TO EXPORT SPAREPARTS:\n(.+)", r"GAGAL MENGEKSPOR DATA SUKU CADANG:\n\1"),
    (r"SAVE SUPPLIERS EXPORT", r"SIMPAN EKSPOR DATA PEMASOK"),
    (r"SAVE SPAREPARTS EXPORT", r"SIMPAN EKSPOR DATA SUKU CADANG"),
    (r"USERNAME=(.+); ROLE=ADMIN", r"NAMA PENGGUNA=\1; PERAN=ADMIN"),
    (r"USERNAME=(.+); ROLE=STAFF", r"NAMA PENGGUNA=\1; PERAN=STAF"),
    (r"ROLE=ADMIN", r"PERAN=ADMIN"),
    (r"ROLE=STAFF", r"PERAN=STAF"),
    (r"USERNAME=(.+)", r"NAMA PENGGUNA=\1"),
])


# Complete dynamic-message localization for every non-English locale.
# Full-sentence templates are used where word order or grammar differs, while
# paths, names, dates, counts, error details, and file formats remain unchanged.
_MULTILINGUAL_DYNAMIC_TEMPLATES = {'es_MX': {'failed_admin': 'NO SE PUDO CREAR EL ADMINISTRADOR:\\n\\1',
           'locked': 'CUENTA BLOQUEADA. INTÉNTELO DE NUEVO EN APROXIMADAMENTE \\1 MINUTO(S).',
           'too_many': 'DEMASIADOS INTENTOS FALLIDOS. LA CUENTA SE BLOQUEÓ DURANTE \\1 MINUTOS.',
           'attempts': 'NOMBRE DE USUARIO O CONTRASEÑA NO VÁLIDOS. QUEDAN \\1 INTENTO(S) ANTES DEL BLOQUEO.',
           'cannot_undo': 'ESTA ACCIÓN NO SE PUEDE DESHACER.',
           'imported': 'SE IMPORTARON CORRECTAMENTE \\1 REGISTROS DESDE {source_format}\\nSE OMITIERON \\2 '
                       'FILAS',
           'sent_logged': 'SE ENVIARON RECORDATORIOS A \\1 CLIENTES\\n\\nLOS RECORDATORIOS SE REGISTRARON EN '
                          'LA BASE DE DATOS.',
           'sent': 'SE ENVIARON RECORDATORIOS A \\1 CLIENTES',
           'low_stock': '⚠️ ADVERTENCIA: ¡\\1 PIEZAS TIENEN POCO STOCK!',
           'period': 'PERÍODO: \\1 A \\2',
           'logged_no_sms': 'SE REGISTRARON RECORDATORIOS PARA \\1 CLIENTES.\\n\\nNO SE ENVIÓ NINGÚN MENSAJE '
                            'SMS NI DE WHATSAPP.',
           'logged': 'SE REGISTRARON RECORDATORIOS PARA \\1 CLIENTES'},
 'fr_FR': {'failed_admin': 'ÉCHEC DE LA CRÉATION DE L’ADMINISTRATEUR :\\n\\1',
           'locked': 'COMPTE VERROUILLÉ. RÉESSAYEZ DANS ENVIRON \\1 MINUTE(S).',
           'too_many': 'TROP DE TENTATIVES ÉCHOUÉES. COMPTE VERROUILLÉ PENDANT \\1 MINUTES.',
           'attempts': 'NOM D’UTILISATEUR OU MOT DE PASSE INCORRECT. IL RESTE \\1 TENTATIVE(S) AVANT LE '
                       'VERROUILLAGE.',
           'cannot_undo': 'CETTE ACTION EST IRRÉVERSIBLE.',
           'imported': 'IMPORTATION RÉUSSIE DE \\1 ENREGISTREMENTS DEPUIS {source_format}\\n\\2 LIGNES '
                       'IGNORÉES',
           'sent_logged': 'RAPPELS ENVOYÉS À \\1 CLIENTS\\n\\nLES RAPPELS ONT ÉTÉ ENREGISTRÉS DANS LA BASE '
                          'DE DONNÉES.',
           'sent': 'RAPPELS ENVOYÉS À \\1 CLIENTS',
           'low_stock': '⚠️ AVERTISSEMENT : \\1 PIÈCES ONT UN STOCK FAIBLE !',
           'period': 'PÉRIODE : \\1 À \\2',
           'logged_no_sms': 'RAPPELS ENREGISTRÉS POUR \\1 CLIENTS.\\n\\nAUCUN MESSAGE SMS OU WHATSAPP N’A '
                            'ÉTÉ ENVOYÉ.',
           'logged': 'RAPPELS ENREGISTRÉS POUR \\1 CLIENTS'},
 'de_DE': {'failed_admin': 'ADMINISTRATOR KONNTE NICHT ERSTELLT WERDEN:\\n\\1',
           'locked': 'KONTO GESPERRT. VERSUCHEN SIE ES IN ETWA \\1 MINUTE(N) ERNEUT.',
           'too_many': 'ZU VIELE FEHLGESCHLAGENE VERSUCHE. KONTO FÜR \\1 MINUTEN GESPERRT.',
           'attempts': 'BENUTZERNAME ODER PASSWORT UNGÜLTIG. NOCH \\1 VERSUCH(E) BIS ZUR SPERRUNG.',
           'cannot_undo': 'DIESE AKTION KANN NICHT RÜCKGÄNGIG GEMACHT WERDEN.',
           'imported': 'ERFOLGREICH \\1 DATENSÄTZE AUS {source_format} IMPORTIERT\\n\\2 ZEILEN ÜBERSPRUNGEN',
           'sent_logged': 'ERINNERUNGEN AN \\1 KUNDEN GESENDET\\n\\nDIE ERINNERUNGEN WURDEN IN DER DATENBANK '
                          'PROTOKOLLIERT.',
           'sent': 'ERINNERUNGEN AN \\1 KUNDEN GESENDET',
           'low_stock': '⚠️ WARNUNG: \\1 TEILE HABEN EINEN NIEDRIGEN LAGERBESTAND!',
           'period': 'ZEITRAUM: \\1 BIS \\2',
           'logged_no_sms': 'ERINNERUNGEN FÜR \\1 KUNDEN PROTOKOLLIERT.\\n\\nES WURDE KEINE SMS- ODER '
                            'WHATSAPP-NACHRICHT GESENDET.',
           'logged': 'ERINNERUNGEN FÜR \\1 KUNDEN PROTOKOLLIERT'},
 'pt_BR': {'failed_admin': 'FALHA AO CRIAR O ADMINISTRADOR:\\n\\1',
           'locked': 'CONTA BLOQUEADA. TENTE NOVAMENTE EM CERCA DE \\1 MINUTO(S).',
           'too_many': 'MUITAS TENTATIVAS SEM SUCESSO. CONTA BLOQUEADA POR \\1 MINUTOS.',
           'attempts': 'NOME DE USUÁRIO OU SENHA INVÁLIDOS. RESTAM \\1 TENTATIVA(S) ANTES DO BLOQUEIO.',
           'cannot_undo': 'ESTA AÇÃO NÃO PODE SER DESFEITA.',
           'imported': '\\1 REGISTROS IMPORTADOS COM SUCESSO DO {source_format}\\n\\2 LINHAS IGNORADAS',
           'sent_logged': 'LEMBRETES ENVIADOS PARA \\1 CLIENTES\\n\\nOS LEMBRETES FORAM REGISTRADOS NO BANCO '
                          'DE DADOS.',
           'sent': 'LEMBRETES ENVIADOS PARA \\1 CLIENTES',
           'low_stock': '⚠️ ATENÇÃO: \\1 PEÇAS ESTÃO COM ESTOQUE BAIXO!',
           'period': 'PERÍODO: \\1 A \\2',
           'logged_no_sms': 'LEMBRETES REGISTRADOS PARA \\1 CLIENTES.\\n\\nNENHUMA MENSAGEM SMS OU WHATSAPP '
                            'FOI ENVIADA.',
           'logged': 'LEMBRETES REGISTRADOS PARA \\1 CLIENTES'},
 'it_IT': {'failed_admin': 'IMPOSSIBILE CREARE L’AMMINISTRATORE:\\n\\1',
           'locked': 'ACCOUNT BLOCCATO. RIPROVA TRA CIRCA \\1 MINUTO/I.',
           'too_many': 'TROPPI TENTATIVI NON RIUSCITI. ACCOUNT BLOCCATO PER \\1 MINUTI.',
           'attempts': 'NOME UTENTE O PASSWORD NON VALIDI. RESTANO \\1 TENTATIVI PRIMA DEL BLOCCO.',
           'cannot_undo': 'QUESTA AZIONE NON PUÒ ESSERE ANNULLATA.',
           'imported': 'IMPORTATI CORRETTAMENTE \\1 RECORD DA {source_format}\\n\\2 RIGHE IGNORATE',
           'sent_logged': 'PROMEMORIA INVIATI A \\1 CLIENTI\\n\\nI PROMEMORIA SONO STATI REGISTRATI NEL '
                          'DATABASE.',
           'sent': 'PROMEMORIA INVIATI A \\1 CLIENTI',
           'low_stock': '⚠️ ATTENZIONE: \\1 RICAMBI HANNO SCORTE BASSE!',
           'period': 'PERIODO: \\1 - \\2',
           'logged_no_sms': 'PROMEMORIA REGISTRATI PER \\1 CLIENTI.\\n\\nNON È STATO INVIATO ALCUN MESSAGGIO '
                            'SMS O WHATSAPP.',
           'logged': 'PROMEMORIA REGISTRATI PER \\1 CLIENTI'},
 'nl_NL': {'failed_admin': 'BEHEERDER KON NIET WORDEN AANGEMAAKT:\\n\\1',
           'locked': 'ACCOUNT GEBLOKKEERD. PROBEER HET OVER ONGEVEER \\1 MINU(U)T(EN) OPNIEUW.',
           'too_many': 'TE VEEL MISLUKTE POGINGEN. ACCOUNT \\1 MINUTEN GEBLOKKEERD.',
           'attempts': 'ONGELDIGE GEBRUIKERSNAAM OF WACHTWOORD. NOG \\1 POGING(EN) VOOR BLOKKERING.',
           'cannot_undo': 'DEZE ACTIE KAN NIET ONGEDAAN WORDEN GEMAAKT.',
           'imported': '\\1 RECORDS UIT {source_format} GEÏMPORTEERD\\n\\2 RIJEN OVERGESLAGEN',
           'sent_logged': 'HERINNERINGEN VERZONDEN AAN \\1 KLANTEN\\n\\nDE HERINNERINGEN ZIJN IN DE DATABASE '
                          'VASTGELEGD.',
           'sent': 'HERINNERINGEN VERZONDEN AAN \\1 KLANTEN',
           'low_stock': '⚠️ WAARSCHUWING: \\1 ONDERDELEN HEBBEN EEN LAGE VOORRAAD!',
           'period': 'PERIODE: \\1 TOT \\2',
           'logged_no_sms': 'HERINNERINGEN VASTGELEGD VOOR \\1 KLANTEN.\\n\\nER IS GEEN SMS- OF '
                            'WHATSAPP-BERICHT VERZONDEN.',
           'logged': 'HERINNERINGEN VASTGELEGD VOOR \\1 KLANTEN'},
 'zh_CN': {'failed_admin': '创建管理员失败：\\n\\1',
           'locked': '账户已锁定。请约 \\1 分钟后重试。',
           'too_many': '失败尝试次数过多。账户已锁定 \\1 分钟。',
           'attempts': '用户名或密码无效。锁定前还可尝试 \\1 次。',
           'cannot_undo': '此操作无法撤销。',
           'imported': '已成功从 {source_format} 导入 \\1 条记录\\n跳过 \\2 行',
           'sent_logged': '已向 \\1 位客户发送提醒\\n\\n提醒已记录到数据库。',
           'sent': '已向 \\1 位客户发送提醒',
           'low_stock': '⚠️ 警告：\\1 个零件库存不足！',
           'period': '期间：\\1 至 \\2',
           'logged_no_sms': '已为 \\1 位客户记录提醒。\\n\\n未发送任何短信或 WhatsApp 消息。',
           'logged': '已为 \\1 位客户记录提醒'},
 'ja_JP': {'failed_admin': '管理者を作成できませんでした：\\n\\1',
           'locked': 'アカウントがロックされています。約 \\1 分後にもう一度お試しください。',
           'too_many': '失敗回数が多すぎます。アカウントは \\1 分間ロックされました。',
           'attempts': 'ユーザー名またはパスワードが無効です。ロックまであと \\1 回試行できます。',
           'cannot_undo': 'この操作は元に戻せません。',
           'imported': '{source_format} から \\1 件のレコードをインポートしました\\n\\2 行をスキップしました',
           'sent_logged': '\\1 人の顧客にリマインダーを送信しました\\n\\nリマインダーをデータベースに記録しました。',
           'sent': '\\1 人の顧客にリマインダーを送信しました',
           'low_stock': '⚠️ 警告：\\1 個の部品の在庫が少なくなっています！',
           'period': '期間：\\1 ～ \\2',
           'logged_no_sms': '\\1 人の顧客のリマインダーを記録しました。\\n\\nSMS または WhatsApp メッセージは送信されていません。',
           'logged': '\\1 人の顧客のリマインダーを記録しました'},
 'ko_KR': {'failed_admin': '관리자를 만들지 못했습니다:\\n\\1',
           'locked': '계정이 잠겼습니다. 약 \\1분 후 다시 시도하세요.',
           'too_many': '실패한 시도가 너무 많습니다. 계정이 \\1분 동안 잠겼습니다.',
           'attempts': '사용자 이름 또는 비밀번호가 잘못되었습니다. 잠기기 전 \\1회 더 시도할 수 있습니다.',
           'cannot_undo': '이 작업은 실행 취소할 수 없습니다.',
           'imported': '{source_format}에서 \\1개 기록을 가져왔습니다\\n\\2개 행을 건너뛰었습니다',
           'sent_logged': '\\1명의 고객에게 알림을 보냈습니다\\n\\n알림을 데이터베이스에 기록했습니다.',
           'sent': '\\1명의 고객에게 알림을 보냈습니다',
           'low_stock': '⚠️ 경고: \\1개 부품의 재고가 부족합니다!',
           'period': '기간: \\1 ~ \\2',
           'logged_no_sms': '\\1명의 고객에 대한 알림을 기록했습니다.\\n\\nSMS 또는 WhatsApp 메시지는 전송되지 않았습니다.',
           'logged': '\\1명의 고객에 대한 알림을 기록했습니다'},
 'ar_SA': {'failed_admin': 'فشل إنشاء المسؤول:\\n\\1',
           'locked': 'الحساب مقفل. حاول مرة أخرى بعد نحو \\1 دقيقة.',
           'too_many': 'محاولات فاشلة كثيرة جدًا. تم قفل الحساب لمدة \\1 دقيقة.',
           'attempts': 'اسم المستخدم أو كلمة المرور غير صحيحة. تبقت \\1 محاولة قبل القفل.',
           'cannot_undo': 'لا يمكن التراجع عن هذا الإجراء.',
           'imported': 'تم استيراد \\1 سجل بنجاح من {source_format}\\nتم تخطي \\2 صف',
           'sent_logged': 'تم إرسال تذكيرات إلى \\1 عميل\\n\\nتم تسجيل التذكيرات في قاعدة البيانات.',
           'sent': 'تم إرسال تذكيرات إلى \\1 عميل',
           'low_stock': '⚠️ تحذير: مخزون \\1 قطعة منخفض!',
           'period': 'الفترة: من \\1 إلى \\2',
           'logged_no_sms': 'تم تسجيل تذكيرات لـ \\1 عميل.\\n\\nلم يتم إرسال أي رسالة SMS أو WhatsApp.',
           'logged': 'تم تسجيل تذكيرات لـ \\1 عميل'},
 'hi_IN': {'failed_admin': 'व्यवस्थापक नहीं बनाया जा सका:\\n\\1',
           'locked': 'खाता लॉक है। लगभग \\1 मिनट बाद फिर प्रयास करें।',
           'too_many': 'बहुत अधिक असफल प्रयास। खाता \\1 मिनट के लिए लॉक कर दिया गया है।',
           'attempts': 'उपयोगकर्ता नाम या पासवर्ड गलत है। लॉक होने से पहले \\1 प्रयास शेष है।',
           'cannot_undo': 'यह कार्रवाई पूर्ववत नहीं की जा सकती।',
           'imported': '{source_format} से \\1 रिकॉर्ड सफलतापूर्वक आयात किए गए\\n\\2 पंक्तियाँ छोड़ी गईं',
           'sent_logged': '\\1 ग्राहकों को रिमाइंडर भेजे गए\\n\\nरिमाइंडर डेटाबेस में दर्ज किए गए हैं।',
           'sent': '\\1 ग्राहकों को रिमाइंडर भेजे गए',
           'low_stock': '⚠️ चेतावनी: \\1 पार्ट का स्टॉक कम है!',
           'period': 'अवधि: \\1 से \\2',
           'logged_no_sms': '\\1 ग्राहकों के लिए रिमाइंडर दर्ज किए गए।\\n\\nकोई SMS या WhatsApp संदेश नहीं '
                            'भेजा गया।',
           'logged': '\\1 ग्राहकों के लिए रिमाइंडर दर्ज किए गए'},
 'bn_BD': {'failed_admin': 'প্রশাসক তৈরি করা যায়নি:\\n\\1',
           'locked': 'অ্যাকাউন্ট লক করা হয়েছে। প্রায় \\1 মিনিট পরে আবার চেষ্টা করুন।',
           'too_many': 'অনেক বেশি ব্যর্থ চেষ্টা। অ্যাকাউন্ট \\1 মিনিটের জন্য লক করা হয়েছে।',
           'attempts': 'ব্যবহারকারীর নাম বা পাসওয়ার্ড ভুল। লক হওয়ার আগে \\1টি চেষ্টা বাকি।',
           'cannot_undo': 'এই কাজটি পূর্বাবস্থায় ফেরানো যাবে না।',
           'imported': '{source_format} থেকে \\1টি রেকর্ড সফলভাবে আমদানি হয়েছে\\n\\2টি সারি বাদ দেওয়া '
                       'হয়েছে',
           'sent_logged': '\\1 জন গ্রাহককে স্মরণিকা পাঠানো হয়েছে\\n\\nস্মরণিকাগুলি ডেটাবেসে নথিভুক্ত '
                          'হয়েছে।',
           'sent': '\\1 জন গ্রাহককে স্মরণিকা পাঠানো হয়েছে',
           'low_stock': '⚠️ সতর্কতা: \\1টি যন্ত্রাংশের মজুদ কম!',
           'period': 'সময়কাল: \\1 থেকে \\2',
           'logged_no_sms': '\\1 জন গ্রাহকের জন্য স্মরণিকা নথিভুক্ত হয়েছে।\\n\\nকোনো SMS বা WhatsApp বার্তা '
                            'পাঠানো হয়নি।',
           'logged': '\\1 জন গ্রাহকের জন্য স্মরণিকা নথিভুক্ত হয়েছে'},
 'ru_RU': {'failed_admin': 'НЕ УДАЛОСЬ СОЗДАТЬ АДМИНИСТРАТОРА:\\n\\1',
           'locked': 'УЧЁТНАЯ ЗАПИСЬ ЗАБЛОКИРОВАНА. ПОВТОРИТЕ ПОПЫТКУ ПРИМЕРНО ЧЕРЕЗ \\1 МИНУТУ(Ы).',
           'too_many': 'СЛИШКОМ МНОГО НЕУДАЧНЫХ ПОПЫТОК. УЧЁТНАЯ ЗАПИСЬ ЗАБЛОКИРОВАНА НА \\1 МИНУТ.',
           'attempts': 'НЕВЕРНОЕ ИМЯ ПОЛЬЗОВАТЕЛЯ ИЛИ ПАРОЛЬ. ДО БЛОКИРОВКИ ОСТАЛОСЬ ПОПЫТОК: \\1.',
           'cannot_undo': 'ЭТО ДЕЙСТВИЕ НЕЛЬЗЯ ОТМЕНИТЬ.',
           'imported': 'УСПЕШНО ИМПОРТИРОВАНО \\1 ЗАПИСЕЙ ИЗ {source_format}\\nПРОПУЩЕНО СТРОК: \\2',
           'sent_logged': 'НАПОМИНАНИЯ ОТПРАВЛЕНЫ \\1 КЛИЕНТАМ\\n\\nНАПОМИНАНИЯ ЗАПИСАНЫ В БАЗУ ДАННЫХ.',
           'sent': 'НАПОМИНАНИЯ ОТПРАВЛЕНЫ \\1 КЛИЕНТАМ',
           'low_stock': '⚠️ ПРЕДУПРЕЖДЕНИЕ: У \\1 ДЕТАЛЕЙ НИЗКИЙ ОСТАТОК!',
           'period': 'ПЕРИОД: С \\1 ПО \\2',
           'logged_no_sms': 'НАПОМИНАНИЯ ЗАПИСАНЫ ДЛЯ \\1 КЛИЕНТОВ.\\n\\nSMS ИЛИ WHATSAPP-СООБЩЕНИЯ НЕ '
                            'ОТПРАВЛЯЛИСЬ.',
           'logged': 'НАПОМИНАНИЯ ЗАПИСАНЫ ДЛЯ \\1 КЛИЕНТОВ'},
 'tr_TR': {'failed_admin': 'YÖNETİCİ OLUŞTURULAMADI:\\n\\1',
           'locked': 'HESAP KİLİTLİ. YAKLAŞIK \\1 DAKİKA SONRA TEKRAR DENEYİN.',
           'too_many': 'ÇOK FAZLA BAŞARISIZ DENEME. HESAP \\1 DAKİKA KİLİTLENDİ.',
           'attempts': 'KULLANICI ADI VEYA PAROLA GEÇERSİZ. KİLİTLENMEDEN ÖNCE \\1 DENEME KALDI.',
           'cannot_undo': 'BU İŞLEM GERİ ALINAMAZ.',
           'imported': '{source_format} DOSYASINDAN \\1 KAYIT BAŞARIYLA İÇE AKTARILDI\\n\\2 SATIR ATLANDI',
           'sent_logged': '\\1 MÜŞTERİYE HATIRLATICI GÖNDERİLDİ\\n\\nHATIRLATICILAR VERİTABANINA KAYDEDİLDİ.',
           'sent': '\\1 MÜŞTERİYE HATIRLATICI GÖNDERİLDİ',
           'low_stock': '⚠️ UYARI: \\1 PARÇANIN STOĞU DÜŞÜK!',
           'period': 'DÖNEM: \\1 - \\2',
           'logged_no_sms': '\\1 MÜŞTERİ İÇİN HATIRLATICI KAYDEDİLDİ.\\n\\nSMS VEYA WHATSAPP MESAJI '
                            'GÖNDERİLMEDİ.',
           'logged': '\\1 MÜŞTERİ İÇİN HATIRLATICI KAYDEDİLDİ'},
 'vi_VN': {'failed_admin': 'KHÔNG THỂ TẠO QUẢN TRỊ VIÊN:\\n\\1',
           'locked': 'TÀI KHOẢN BỊ KHÓA. HÃY THỬ LẠI SAU KHOẢNG \\1 PHÚT.',
           'too_many': 'QUÁ NHIỀU LẦN THỬ THẤT BẠI. TÀI KHOẢN BỊ KHÓA TRONG \\1 PHÚT.',
           'attempts': 'TÊN NGƯỜI DÙNG HOẶC MẬT KHẨU KHÔNG HỢP LỆ. CÒN \\1 LẦN THỬ TRƯỚC KHI BỊ KHÓA.',
           'cannot_undo': 'KHÔNG THỂ HOÀN TÁC THAO TÁC NÀY.',
           'imported': 'ĐÃ NHẬP THÀNH CÔNG \\1 BẢN GHI TỪ {source_format}\\nĐÃ BỎ QUA \\2 DÒNG',
           'sent_logged': 'ĐÃ GỬI NHẮC NHỞ CHO \\1 KHÁCH HÀNG\\n\\nCÁC NHẮC NHỞ ĐÃ ĐƯỢC GHI VÀO CƠ SỞ DỮ '
                          'LIỆU.',
           'sent': 'ĐÃ GỬI NHẮC NHỞ CHO \\1 KHÁCH HÀNG',
           'low_stock': '⚠️ CẢNH BÁO: \\1 LINH KIỆN SẮP HẾT HÀNG!',
           'period': 'KỲ: \\1 ĐẾN \\2',
           'logged_no_sms': 'ĐÃ GHI NHẮC NHỞ CHO \\1 KHÁCH HÀNG.\\n\\nKHÔNG CÓ TIN NHẮN SMS HOẶC WHATSAPP '
                            'NÀO ĐƯỢC GỬI.',
           'logged': 'ĐÃ GHI NHẮC NHỞ CHO \\1 KHÁCH HÀNG'},
 'th_TH': {'failed_admin': 'ไม่สามารถสร้างผู้ดูแลระบบได้:\\n\\1',
           'locked': 'บัญชีถูกล็อก โปรดลองอีกครั้งในประมาณ \\1 นาที',
           'too_many': 'มีความพยายามล้มเหลวมากเกินไป บัญชีถูกล็อกเป็นเวลา \\1 นาที',
           'attempts': 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง เหลืออีก \\1 ครั้งก่อนบัญชีถูกล็อก',
           'cannot_undo': 'ไม่สามารถยกเลิกการดำเนินการนี้ได้',
           'imported': 'นำเข้า \\1 ระเบียนจาก {source_format} สำเร็จ\\nข้าม \\2 แถว',
           'sent_logged': 'ส่งการแจ้งเตือนให้ลูกค้า \\1 รายแล้ว\\n\\nบันทึกการแจ้งเตือนลงในฐานข้อมูลแล้ว',
           'sent': 'ส่งการแจ้งเตือนให้ลูกค้า \\1 รายแล้ว',
           'low_stock': '⚠️ คำเตือน: อะไหล่ \\1 รายการมีสต็อกต่ำ!',
           'period': 'ช่วงเวลา: \\1 ถึง \\2',
           'logged_no_sms': 'บันทึกการแจ้งเตือนสำหรับลูกค้า \\1 รายแล้ว\\n\\nไม่มีการส่งข้อความ SMS หรือ '
                            'WhatsApp',
           'logged': 'บันทึกการแจ้งเตือนสำหรับลูกค้า \\1 รายแล้ว'},
 'ms_MY': {'failed_admin': 'GAGAL MENCIPTA PENTADBIR:\\n\\1',
           'locked': 'AKAUN DIKUNCI. CUBA LAGI DALAM KIRA-KIRA \\1 MINIT.',
           'too_many': 'TERLALU BANYAK PERCUBAAN GAGAL. AKAUN DIKUNCI SELAMA \\1 MINIT.',
           'attempts': 'NAMA PENGGUNA ATAU KATA LALUAN TIDAK SAH. TINGGAL \\1 PERCUBAAN SEBELUM DIKUNCI.',
           'cannot_undo': 'TINDAKAN INI TIDAK BOLEH DIBATALKAN.',
           'imported': 'BERJAYA MENGIMPORT \\1 REKOD DARIPADA {source_format}\\n\\2 BARIS DILANGKAU',
           'sent_logged': 'PERINGATAN DIHANTAR KEPADA \\1 PELANGGAN\\n\\nPERINGATAN TELAH DIREKODKAN DALAM '
                          'PANGKALAN DATA.',
           'sent': 'PERINGATAN DIHANTAR KEPADA \\1 PELANGGAN',
           'low_stock': '⚠️ AMARAN: \\1 ALAT GANTI MEMPUNYAI STOK RENDAH!',
           'period': 'TEMPOH: \\1 HINGGA \\2',
           'logged_no_sms': 'PERINGATAN DIREKODKAN UNTUK \\1 PELANGGAN.\\n\\nTIADA MESEJ SMS ATAU WHATSAPP '
                            'DIHANTAR.',
           'logged': 'PERINGATAN DIREKODKAN UNTUK \\1 PELANGGAN'},
 'fil_PH': {'failed_admin': 'HINDI NAGAWANG LUMIKHA NG ADMINISTRATOR:\\n\\1',
            'locked': 'NA-LOCK ANG ACCOUNT. SUBUKANG MULI PAGKALIPAS NG HUMIGIT-KUMULANG \\1 MINUTO.',
            'too_many': 'NAPAKARAMING NABIGONG PAGTATANGKA. NA-LOCK ANG ACCOUNT NANG \\1 MINUTO.',
            'attempts': 'HINDI WASTO ANG USERNAME O PASSWORD. MAY \\1 PAGSUBOK PA BAGO MA-LOCK.',
            'cannot_undo': 'HINDI MAAARING I-UNDO ANG PAGKILOS NA ITO.',
            'imported': 'MATAGUMPAY NA NA-IMPORT ANG \\1 RECORD MULA SA {source_format}\\n\\2 ROW ANG '
                        'NILAKTAWAN',
            'sent_logged': 'NAIPADALA ANG MGA PAALALA SA \\1 CUSTOMER\\n\\nNAITALA ANG MGA PAALALA SA '
                           'DATABASE.',
            'sent': 'NAIPADALA ANG MGA PAALALA SA \\1 CUSTOMER',
            'low_stock': '⚠️ BABALA: MABABA ANG STOCK NG \\1 PIYESA!',
            'period': 'PANAHON: \\1 HANGGANG \\2',
            'logged_no_sms': 'NAITALA ANG MGA PAALALA PARA SA \\1 CUSTOMER.\\n\\nWALANG IPINADALANG SMS O '
                             'WHATSAPP MESSAGE.',
            'logged': 'NAITALA ANG MGA PAALALA PARA SA \\1 CUSTOMER'}}

def _dynamic_ui_translation(language_code, source_key, fallback=None):
    value = UI_TRANSLATIONS.get(language_code, {}).get(source_key)
    return str(value if value is not None else (fallback if fallback is not None else source_key))

def _dynamic_ui_label(language_code, source_key, fallback=None):
    return _dynamic_ui_translation(language_code, source_key, fallback).rstrip(" :：")

def _build_multilingual_dynamic_patterns(language_code, template):
    t = lambda key, fallback=None: _dynamic_ui_translation(language_code, key, fallback)
    label = lambda key, fallback=None: _dynamic_ui_label(language_code, key, fallback)
    colon = lambda key, fallback=None: label(key, fallback) + ":"
    cannot_undo = template["cannot_undo"]
    return [
        (r"FAILED TO CREATE ADMINISTRATOR:\n(.+)", template["failed_admin"]),
        (r"ACCOUNT LOCKED\. TRY AGAIN IN ABOUT (\d+) MINUTE\(S\)\.", template["locked"]),
        (r"TOO MANY FAILED ATTEMPTS\. ACCOUNT LOCKED FOR (\d+) MINUTES\.", template["too_many"]),
        (r"INVALID USERNAME OR PASSWORD\. (\d+) ATTEMPT\(S\) REMAIN BEFORE LOCKOUT\.", template["attempts"]),
        (r"DROPDOWN DATA RELOADED: (.+)", colon("RELOAD DROPDOWN DATA") + r" \1"),
        (r"DUMMY\.INI NOT FOUND - RUN DUMMY_DATA_EDITOR\.PY: (.+)",
         f"DUMMY.INI {t('NOT FOUND')} - {t('OPEN DUMMY DATA EDITOR')}: " + r"\1"),
        (r"FAILED TO IMPORT DATA:\n(.+)", colon("IMPORT FAILED") + r"\n\1"),
        (r"EXPORT ERROR: (.+)", colon("EXPORT ERROR") + r" \1"),
        (r"DATA HAS BEEN EXPORTED TO:\n(.+)", colon("DATA HAS BEEN EXPORTED TO:") + r"\n\1"),
        (r"BACKED UP TO (.+)", label("BACKUP SUCCESSFUL") + r": \1"),
        (r"EXPORT FAILED: (.+)", colon("EXPORT FAILED") + r" \1"),
        (r"DATABASE BACKED UP SUCCESSFULLY TO:\n(.+)", colon("BACKUP SUCCESSFUL") + r"\n\1"),
        (r"FAILED TO BACKUP DATABASE:\n(.+)", colon("BACKUP FAILED") + r"\n\1"),
        (r"FAILED TO RESTORE DATABASE:\n(.+)", colon("RESTORE FAILED") + r"\n\1"),
        (r"SUCCESSFULLY IMPORTED (\d+) RECORDS FROM EXCEL\n(\d+) ROWS SKIPPED",
         template["imported"].format(source_format="EXCEL")),
        (r"SUCCESSFULLY IMPORTED (\d+) RECORDS FROM CSV\n(\d+) ROWS SKIPPED",
         template["imported"].format(source_format="CSV")),
        (r"IMPORTED (\d+) RECORDS", label("IMPORT COMPLETE") + r": \1 " + t("RECORDS")),
        (r"DATA EXPORTED SUCCESSFULLY TO:\n(.+)", colon("DATA HAS BEEN EXPORTED TO:") + r"\n\1"),
        (r"EXPORTED TO (.+)", label("EXPORTED TO:") + r" \1"),
        (r"FAILED TO EXPORT DATA:\n(.+)", colon("EXPORT FAILED") + r"\n\1"),
        (r"SELECTED DATA EXPORTED TO:\n(.+)",
         f"{t('SELECTED')} — {label('DATA HAS BEEN EXPORTED TO:')}:" + r"\n\1"),
        (r"EXPORTED (\d+) ROWS", label("EXPORTED") + r" \1 " + t("ROWS")),
        (r"REPORT EXPORTED TO:\n(.+)", colon("REPORT EXPORTED") + r"\n\1"),
        (r"REPORT EXPORTED TO (.+)", label("REPORT EXPORTED") + r": \1"),
        (r"FAILED TO EXPORT REPORT:\n(.+)", colon("REPORT EXPORT FAILED") + r"\n\1"),
        (r"WARNING: FAILED TO REGISTER SUPPLIER/PART: (.+)",
         f"{label('WARNING')}: {t('FAILED')} — {t('SUPPLIER')}/{t('PART')}: " + r"\1"),
        (r"OPENED CONTACT FROM WARRANTY: (.+)", f"{t('CONTACT')} — {t('WARRANTY')}: " + r"\1"),
        (r"CONTACT NOT FOUND IN TABLE: (.+)", f"{t('CONTACT')} {t('NOT FOUND')}: " + r"\1"),
        (r"SELECTED SUPPLIER: (.+)", f"{t('SELECTED')} {t('SUPPLIER')}: " + r"\1"),
        (r"UPDATED SUPPLIER: (.+)", f"{t('UPDATED')} {t('SUPPLIER')}: " + r"\1"),
        (r"DELETE SUPPLIER '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!",
         f"{t('DELETE')} {t('SUPPLIER')} '" + r"\1" + f"'?\n\n{cannot_undo}"),
        (r"DELETED SUPPLIER: (.+)", f"{t('DELETED')} {t('SUPPLIER')}: " + r"\1"),
        (r"ADDED PART: (.+)", f"{t('ADDED')} {t('PART')}: " + r"\1"),
        (r"UPDATED PART: (.+)", f"{t('UPDATED')} {t('PART')}: " + r"\1"),
        (r"DELETE PART '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!",
         f"{t('DELETE')} {t('PART')} '" + r"\1" + f"'?\n\n{cannot_undo}"),
        (r"DELETED PART: (.+)", f"{t('DELETED')} {t('PART')}: " + r"\1"),
        (r"DELETE '(.+)'\?\n\nTHIS ACTION CANNOT BE UNDONE!",
         t("DELETE") + " '" + r"\1" + f"'?\n\n{cannot_undo}"),
        (r"DELETED (.+)", t("DELETED") + r": \1"),
        (r"ADDED (.+)", t("ADDED") + r": \1"),
        (r"REMINDERS SENT TO (\d+) CUSTOMERS\n\nREMINDERS HAVE BEEN LOGGED IN THE DATABASE\.", template["sent_logged"]),
        (r"REMINDERS SENT TO (\d+) CUSTOMERS", template["sent"]),
        (r"⚠️ WARNING: (\d+) PARTS HAVE LOW STOCK!", template["low_stock"]),
        (r"(-?\d+) \(EXPIRED\)", r"\1 (" + t("EXPIRED") + ")"),
        (r"(-?\d+) \(SOON\)", r"\1 (" + t("EXPIRING SOON") + ")"),
        (r"(-?\d+) \(WARNING\)", r"\1 (" + t("WARNING") + ")"),
        (r"PERIOD: (.+) TO (.+)", template["period"]),
        (r"GENERATED: (.+)", colon("GENERATED") + r" \1"),
        (r"CURRENT DATE: (.+)", colon("CURRENT DATE") + r" \1"),
        (r"TOTAL SERVICES: (.+)", colon("TOTAL SERVICES") + r" \1"),
        (r"TOTAL REVENUE: (.+)", colon("TOTAL REVENUE") + r" \1"),
        (r"TOTAL ACTIVE WARRANTIES: (.+)", colon("TOTAL ACTIVE WARRANTIES") + r" \1"),
        (r"TOTAL EXPIRING SOON: (.+)", colon("TOTAL EXPIRING SOON") + r" \1"),
        (r"TOTAL EXPIRED: (.+)", colon("TOTAL EXPIRED") + r" \1"),
        (r"COMPLETED: (.+)", colon("COMPLETED") + r" \1"),
        (r"IN REPAIR: (.+)", colon("IN REPAIR") + r" \1"),
        (r"IN PROGRESS: (.+)", colon("IN PROGRESS") + r" \1"),
        (r"WAITING FOR PARTS: (.+)", colon("WAITING FOR PARTS") + r" \1"),
        (r"REVENUE: (.+)", colon("REVENUE") + r" \1"),
        (r"AVG PER SERVICE: (.+)", colon("AVG PER SERVICE") + r" \1"),
        (r"TOTAL DEVICES: (.+)", colon("TOTAL DEVICES") + r" \1"),
        (r"UNIQUE MODELS: (.+)", colon("UNIQUE MODELS") + r" \1"),
        (r"REPAIRED: (.+)", colon("REPAIRED") + r" \1"),
        (r"PENDING: (.+)", colon("PENDING") + r" \1"),
        (r"AVG REPAIR COST: (.+)", colon("AVG REPAIR COST") + r" \1"),
        (r"TOTAL CUSTOMERS: (.+)", colon("TOTAL CUSTOMERS") + r" \1"),
        (r"ACTIVE: (.+)", colon("ACTIVE") + r" \1"),
        (r"EXPIRING SOON: (.+)", colon("EXPIRING SOON") + r" \1"),
        (r"EXPIRED: (.+)", colon("EXPIRED") + r" \1"),
        (r"SIGNED IN AS ADMINISTRATOR", t("SIGNED IN AS") + " " + t("ADMINISTRATOR")),
        (r"SIGNED IN AS STAFF", t("SIGNED IN AS") + " " + t("STAFF")),
        (r"SIGNED IN AS (.+)", t("SIGNED IN AS") + r" \1"),
        (r"DUMMY\.INI NOT FOUND - OPEN SETTINGS > EDIT DUMMY DATA: (.+)",
         f"DUMMY.INI {t('NOT FOUND')} - {t('SETTINGS')} > {t('EDIT DUMMY DATA')}: " + r"\1"),
        (r"DASHBOARD REFRESH FAILED: (.+)", colon("DASHBOARD REFRESH FAILED") + r" \1"),
        (r"REMINDERS LOGGED FOR (\d+) CUSTOMERS\.\n\nNO SMS OR WHATSAPP MESSAGE WAS SENT\.", template["logged_no_sms"]),
        (r"REMINDERS LOGGED FOR (\d+) CUSTOMERS", template["logged"]),
        (r"EXPORTED TO:\n(.+)", colon("EXPORTED TO:") + r"\n\1"),
        (r"EXPORTED SUPPLIERS TO EXCEL", t("EXPORT SUPPLIERS TO EXCEL")),
        (r"EXPORTED SPAREPARTS TO EXCEL", t("EXPORT PARTS TO EXCEL")),
        (r"FAILED TO EXPORT SUPPLIERS:\n(.+)",
         f"{label('EXPORT FAILED')} — {t('SUPPLIERS')}:" + r"\n\1"),
        (r"FAILED TO EXPORT SPAREPARTS:\n(.+)",
         f"{label('EXPORT FAILED')} — {t('PARTS')}:" + r"\n\1"),
        (r"SAVE SUPPLIERS EXPORT", f"{t('SAVE')} — {t('EXPORT SUPPLIERS TO EXCEL')}"),
        (r"SAVE SPAREPARTS EXPORT", f"{t('SAVE')} — {t('EXPORT PARTS TO EXCEL')}"),
        (r"USERNAME=(.+); ROLE=ADMIN", f"{t('USERNAME')}=" + r"\1" + f"; {t('ROLE')}={t('ADMINISTRATOR')}"),
        (r"USERNAME=(.+); ROLE=STAFF", f"{t('USERNAME')}=" + r"\1" + f"; {t('ROLE')}={t('STAFF')}"),
        (r"ROLE=ADMIN", f"{t('ROLE')}={t('ADMINISTRATOR')}"),
        (r"ROLE=STAFF", f"{t('ROLE')}={t('STAFF')}"),
        (r"USERNAME=(.+)", f"{t('USERNAME')}=" + r"\1"),
    ]

for _language_code, _dynamic_template in _MULTILINGUAL_DYNAMIC_TEMPLATES.items():
    UI_TRANSLATION_PATTERNS.setdefault(_language_code, []).extend(
        _build_multilingual_dynamic_patterns(_language_code, _dynamic_template)
    )
del _language_code, _dynamic_template

# Keep only the two supported language/localization catalogs in memory.
SUPPORTED_LANGUAGE_CODES = ("en_US", "id_ID")
SUPPORTED_CURRENCY_CODES = ("USD", "IDR")
LANGUAGE_OPTIONS = {
    code: LANGUAGE_OPTIONS[code]
    for code in SUPPORTED_LANGUAGE_CODES
}
CURRENCY_OPTIONS = {
    code: CURRENCY_OPTIONS[code]
    for code in SUPPORTED_CURRENCY_CODES
}
UI_TRANSLATIONS = {
    "id_ID": UI_TRANSLATIONS.get("id_ID", {})
}
UI_TRANSLATION_PATTERNS = {
    "id_ID": UI_TRANSLATION_PATTERNS.get("id_ID", [])
}


class ReliableComboBox(QComboBox):
    """Combo box with a readable popup, search completion, and safe wheel behavior."""

    def __init__(self, parent=None):
        super().__init__(parent)
        popup_view = QListView(self)
        popup_view.setUniformItemSizes(True)
        popup_view.setTextElideMode(Qt.TextElideMode.ElideNone)
        popup_view.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        popup_view.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setView(popup_view)
        self.setMaxVisibleItems(18)
        self.setMinimumContentsLength(12)
        self.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)

    def setEditable(self, editable):
        super().setEditable(editable)
        if editable:
            self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            completer = self.completer()
            if completer is not None:
                completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                completer.setFilterMode(Qt.MatchFlag.MatchContains)
                completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                completer.popup().setUniformItemSizes(True)

    def showPopup(self):
        font_metrics = self.fontMetrics()
        widest = self.width()
        for index in range(self.count()):
            widest = max(widest, font_metrics.horizontalAdvance(self.itemText(index)) + 56)
        available_width = QApplication.primaryScreen().availableGeometry().width() if QApplication.primaryScreen() else 1200
        popup_width = min(widest, max(360, int(available_width * 0.70)))
        self.view().setMinimumWidth(popup_width)
        super().showPopup()
        popup_window = self.view().window()
        if popup_window is not None:
            popup_window.setMinimumWidth(popup_width)

    def paintEvent(self, event):
        super().paintEvent(event)

        # Draw a small theme-aware chevron over the dropdown sub-control. This
        # keeps the arrow visible with every Qt platform style and every theme.
        option = QStyleOptionComboBox()
        self.initStyleOption(option)
        arrow_rect = self.style().subControlRect(
            QStyle.ComplexControl.CC_ComboBox,
            option,
            QStyle.SubControl.SC_ComboBoxArrow,
            self,
        )
        if not arrow_rect.isValid() or arrow_rect.width() < 4:
            width = 28
            if self.layoutDirection() == Qt.LayoutDirection.RightToLeft:
                arrow_rect = QRect(0, 0, width, self.height())
            else:
                arrow_rect = QRect(self.width() - width, 0, width, self.height())

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        color_group = (
            QPalette.ColorGroup.Active
            if self.isEnabled()
            else QPalette.ColorGroup.Disabled
        )
        arrow_color = self.palette().color(color_group, QPalette.ColorRole.Text)
        pen = QPen(arrow_color)
        pen.setWidthF(1.7)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)

        center = arrow_rect.center()
        half_width = max(3, min(5, arrow_rect.width() // 5))
        vertical_offset = 1 if self.view().isVisible() else 0
        y = center.y() - 1 + vertical_offset
        path = QPainterPath()
        path.moveTo(center.x() - half_width, y - 2)
        path.lineTo(center.x(), y + 2)
        path.lineTo(center.x() + half_width, y - 2)
        painter.drawPath(path)

    def wheelEvent(self, event):
        # Prevent accidental value changes while scrolling a tab/form.
        if self.hasFocus():
            super().wheelEvent(event)
        else:
            event.ignore()


class BuiltInDummyDataEditor(QDialog):
    """Internal editor for the external dummy.ini dropdown data file."""

    CATEGORY_LABELS = {
        "device_types": "DEVICE TYPES",
        "brands": "BRANDS",
        "models": "MODELS",
        "issues": "ISSUES",
        "statuses": "STATUSES",
        "technicians": "TECHNICIANS",
        "warranty_periods": "WARRANTY PERIODS",
        "parts": "PARTS",
        "sparepart_types": "PART TYPES",
    }

    def __init__(self, manager):
        super().__init__(manager)
        self.manager = manager
        self._t = getattr(manager, "t", lambda value: value)
        self.setWindowTitle(self._t("EDIT DUMMY DATA"))
        self.setModal(False)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.resize(860, 560)
        self.setMinimumSize(720, 460)

        self.data = {
            key: list(manager.dropdown_data.get(key, []))
            for key in manager.EDITABLE_DROPDOWN_CATEGORIES
        }

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        description = QLabel(
            self._t("EDIT DROPDOWN VALUES USED BY THE APPLICATION. "
                    "CHANGES ARE SAVED TO DUMMY.INI BESIDE THE APPLICATION.")
        )
        description.setWordWrap(True)
        root.addWidget(description)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        category_panel = QWidget()
        category_layout = QVBoxLayout(category_panel)
        category_layout.setContentsMargins(0, 0, 0, 0)
        category_layout.setSpacing(5)
        category_layout.addWidget(QLabel(self._t("CATEGORY")))

        self.category_list = QListWidget()
        self.category_list.setMinimumWidth(210)
        self.category_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        for key in manager.EDITABLE_DROPDOWN_CATEGORIES:
            label = self.CATEGORY_LABELS.get(key, key.replace("_", " ").upper())
            item = QListWidgetItem(self._t(label))
            item.setData(Qt.ItemDataRole.UserRole, key)
            self.category_list.addItem(item)
        category_layout.addWidget(self.category_list, 1)
        splitter.addWidget(category_panel)

        values_panel = QWidget()
        values_layout = QVBoxLayout(values_panel)
        values_layout.setContentsMargins(0, 0, 0, 0)
        values_layout.setSpacing(6)

        self.category_title = QLabel()
        self.category_title.setObjectName("accentLabel")
        self.category_title.setStyleSheet("font-size: 11pt; font-weight: 700;")
        values_layout.addWidget(self.category_title)

        self.value_list = QListWidget()
        self.value_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.value_list.setAlternatingRowColors(True)
        self.value_list.itemDoubleClicked.connect(self.copy_selected_value_to_input)
        values_layout.addWidget(self.value_list, 1)

        order_buttons = QHBoxLayout()
        self.move_up_btn = QPushButton("▲ " + self._t("UP"))
        self.move_down_btn = QPushButton("▼ " + self._t("DOWN"))
        self.delete_btn = QPushButton("❌ " + self._t("DELETE SELECTED"))
        self.delete_btn.setObjectName("dangerButton")
        self.move_up_btn.clicked.connect(lambda: self.move_selected_value(-1))
        self.move_down_btn.clicked.connect(lambda: self.move_selected_value(1))
        self.delete_btn.clicked.connect(self.delete_selected_value)
        order_buttons.addWidget(self.move_up_btn)
        order_buttons.addWidget(self.move_down_btn)
        order_buttons.addStretch(1)
        order_buttons.addWidget(self.delete_btn)
        values_layout.addLayout(order_buttons)

        input_row = QHBoxLayout()
        self.value_input = QLineEdit()
        self.value_input.setPlaceholderText(self._t("TYPE NEW VALUE..."))
        self.value_input.returnPressed.connect(self.add_value)
        self.add_btn = QPushButton("➕ " + self._t("ADD"))
        self.add_btn.setObjectName("successButton")
        self.update_btn = QPushButton("✏️ " + self._t("UPDATE"))
        self.add_btn.clicked.connect(self.add_value)
        self.update_btn.clicked.connect(self.update_selected_value)
        input_row.addWidget(self.value_input, 1)
        input_row.addWidget(self.add_btn)
        input_row.addWidget(self.update_btn)
        values_layout.addLayout(input_row)
        splitter.addWidget(values_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([230, 600])

        footer = QHBoxLayout()
        self.path_label = QLabel(manager.dropdown_data_path)
        self.path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.path_label.setToolTip(manager.dropdown_data_path)
        footer.addWidget(self.path_label, 1)

        reload_btn = QPushButton("🔄 " + self._t("RELOAD"))
        save_btn = QPushButton("💾 " + self._t("SAVE"))
        close_btn = QPushButton(self._t("CLOSE"))
        save_btn.setObjectName("primaryButton")
        reload_btn.clicked.connect(self.reload_from_disk)
        save_btn.clicked.connect(self.save_to_disk)
        close_btn.clicked.connect(self.close)
        footer.addWidget(reload_btn)
        footer.addWidget(save_btn)
        footer.addWidget(close_btn)
        root.addLayout(footer)

        self.category_list.currentItemChanged.connect(self.populate_current_category)
        self.value_list.currentItemChanged.connect(self.on_value_selected)
        if self.category_list.count():
            self.category_list.setCurrentRow(0)

    def current_category_key(self):
        item = self.category_list.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item is not None else None

    def populate_current_category(self, current=None, previous=None):
        del previous
        key = current.data(Qt.ItemDataRole.UserRole) if current is not None else self.current_category_key()
        if not key:
            return
        label = self.CATEGORY_LABELS.get(key, key.replace("_", " ").upper())
        self.category_title.setText(self._t(label))
        self.value_list.clear()
        self.value_list.addItems(self.data.get(key, []))
        self.value_input.clear()
        if self.value_list.count():
            self.value_list.setCurrentRow(0)

    def on_value_selected(self, current, previous=None):
        del previous
        if current is not None:
            self.value_input.setText(current.text())

    def copy_selected_value_to_input(self, item):
        if item is not None:
            self.value_input.setText(item.text())
            self.value_input.setFocus()
            self.value_input.selectAll()

    def normalized_input_value(self):
        return str(self.value_input.text()).strip().upper()

    def add_value(self):
        key = self.current_category_key()
        value = self.normalized_input_value()
        if not key or not value:
            return
        existing = {item.casefold() for item in self.data.get(key, [])}
        if value.casefold() in existing:
            QMessageBox.information(self, self._t("DUPLICATE VALUE"), self._t("VALUE ALREADY EXISTS."))
            return
        self.data.setdefault(key, []).append(value)
        self.populate_current_category(self.category_list.currentItem())
        self.value_list.setCurrentRow(self.value_list.count() - 1)
        self.value_input.clear()
        self.value_input.setFocus()

    def update_selected_value(self):
        key = self.current_category_key()
        row = self.value_list.currentRow()
        value = self.normalized_input_value()
        if not key or row < 0 or not value:
            return
        values = self.data.get(key, [])
        duplicate = any(
            index != row and item.casefold() == value.casefold()
            for index, item in enumerate(values)
        )
        if duplicate:
            QMessageBox.information(self, self._t("DUPLICATE VALUE"), self._t("VALUE ALREADY EXISTS."))
            return
        values[row] = value
        self.populate_current_category(self.category_list.currentItem())
        self.value_list.setCurrentRow(row)

    def delete_selected_value(self):
        key = self.current_category_key()
        row = self.value_list.currentRow()
        if not key or row < 0:
            return
        del self.data[key][row]
        self.populate_current_category(self.category_list.currentItem())
        if self.value_list.count():
            self.value_list.setCurrentRow(min(row, self.value_list.count() - 1))

    def move_selected_value(self, offset):
        key = self.current_category_key()
        row = self.value_list.currentRow()
        if not key or row < 0:
            return
        target = row + int(offset)
        values = self.data.get(key, [])
        if target < 0 or target >= len(values):
            return
        values[row], values[target] = values[target], values[row]
        self.populate_current_category(self.category_list.currentItem())
        self.value_list.setCurrentRow(target)

    def reload_from_disk(self):
        self.manager.load_dropdown_data()
        self.data = {
            key: list(self.manager.dropdown_data.get(key, []))
            for key in self.manager.EDITABLE_DROPDOWN_CATEGORIES
        }
        self.populate_current_category(self.category_list.currentItem())
        self.manager.set_status(self._t("DROPDOWN DATA RELOADED"))

    def save_to_disk(self):
        parser = configparser.ConfigParser(interpolation=None)
        parser.optionxform = str
        for key in self.manager.EDITABLE_DROPDOWN_CATEGORIES:
            parser[key.upper()] = {"VALUES": "|".join(self.data.get(key, []))}

        target = Path(self.manager.dropdown_data_path)
        temporary = target.with_suffix(target.suffix + ".tmp")
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            with temporary.open("w", encoding="utf-8", newline="") as stream:
                parser.write(stream)
            os.replace(temporary, target)
        except OSError as exc:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
            QMessageBox.critical(
                self,
                self._t("SAVE FAILED"),
                self._t("FAILED TO SAVE DUMMY.INI") + f":\n{exc}\n\n{target}",
            )
            return False

        self.manager.reload_dropdown_data(show_message=False)
        self.manager.set_status(self._t("DUMMY.INI SAVED") + f": {target}")
        QMessageBox.information(
            self,
            self._t("DUMMY.INI SAVED"),
            self._t("DROPDOWN DATA SAVED SUCCESSFULLY.") + f"\n\n{target}",
        )
        return True






class ServiceManager(QMainWindow):
    """Main application with internal editing, external dummy creator, and CHM help."""

    EDITABLE_DROPDOWN_CATEGORIES = (
        "device_types",
        "brands",
        "models",
        "issues",
        "statuses",
        "technicians",
        "warranty_periods",
        "parts",
        "sparepart_types",
    )

    SYSTEM_DROPDOWN_OPTIONS = {
        "search_fields": [
            "ALL FIELDS", "NAME", "PHONE", "DEVICES", "BRAND", "MODEL",
            "ISSUE", "STATUS", "TECHNICIAN", "PART", "SUPPLIER", "INVOICE",
        ],
        "report_types": [
            "SERVICE SUMMARY",
            "FINANCIAL REPORT",
            "TECHNICIAN PERFORMANCE",
            "DEVICE BRAND STATISTICS",
            "WARRANTY SUMMARY",
        ],
        "warranty_filters": [
            "ALL",
            "ACTIVE",
            "EXPIRING SOON (≤ 7 DAYS)",
            "EXPIRING SOON (≤ 30 DAYS)",
            "EXPIRED",
            "NO WARRANTY",
        ],
    }

    def __init__(self):
        super().__init__()
        self.apply_app_icon()
        self.conn = None
        self.cursor = None
        # The application now opens directly in unrestricted local-access mode.
        self.current_user = {
            "username": "local",
            "display_name": "Local User",
            "role": "admin",
            "is_active": 1,
        }
        self.default_db_path = r"D:\\BACKUP\\DATABASE USER.db"
        self.default_backup_dir = r"D:\\BACKUP\\"
        self.app_dir = get_runtime_app_dir()
        self.settings_ini_path = os.path.join(self.app_dir, "setting.ini")
        self.dropdown_data_path = os.path.join(self.app_dir, "dummy.ini")
        self.dropdown_data = {key: [] for key in self.EDITABLE_DROPDOWN_CATEGORIES}
        self.dropdown_watcher = None
        self.settings = {}
        self.is_first_run = False
        self.selected_contact_row_id = None
        self.selected_supplier_row_id = None
        self.selected_supplier_original_name = None
        self.selected_sparepart_row_id = None
        self._is_shutting_down = False
        self._dummy_editor_dialog = None

        # Load configuration and theme before building widgets.
        self.load_settings()
        self.apply_theme_setting_to_ui()
        self.ensure_database_location_ready()
        self.load_dropdown_data()
        self.create_database()

        self.initUI()
        self.apply_theme_setting_to_ui()
        self.apply_language_setting_to_ui()
        self.apply_currency_setting_to_ui()
        self.setup_dropdown_file_watcher()
        QTimer.singleShot(0, self.prompt_locale_on_launch_if_needed)

        # Load database-backed values after all combo boxes exist.
        self.load_contacts()
        self.load_suppliers_for_combo()
        self.refresh_dropdown_combos()
        self.apply_role_permissions()

        self.warranty_timer = QTimer(self)
        self.warranty_timer.timeout.connect(self.update_warranty_status)
        self.warranty_timer.start(60000)

        self.audit_event("APPLICATION_OPEN", details="MAIN WINDOW OPENED")

    @staticmethod
    def normalize_dropdown_values(values):
        """Normalize dropdown values without embedding any sample data in this app."""
        normalized = []
        seen = set()
        for value in values or []:
            item = str(value).strip().upper()
            marker = item.casefold()
            if item and marker not in seen:
                normalized.append(item)
                seen.add(marker)
        return normalized

    def load_dropdown_data(self):
        """Read editable dropdown values from dummy.ini without creating or modifying it."""
        self.dropdown_data = {key: [] for key in self.EDITABLE_DROPDOWN_CATEGORIES}
        if not os.path.isfile(self.dropdown_data_path):
            return False

        parser = configparser.ConfigParser(interpolation=None)
        parser.optionxform = str
        try:
            parser.read(self.dropdown_data_path, encoding="utf-8")
        except (OSError, configparser.Error):
            return False

        legacy = parser["DUMMY"] if parser.has_section("DUMMY") else None
        for key in self.EDITABLE_DROPDOWN_CATEGORIES:
            values = []
            if legacy is not None:
                raw = legacy.get(key, "")
                values = str(raw).split("|") if str(raw).strip() else []
            else:
                section_name = key.upper()
                if parser.has_section(section_name):
                    section = parser[section_name]
                    raw = section.get("VALUES", "")
                    if str(raw).strip():
                        values = str(raw).split("|")
                    else:
                        indexed = []
                        for option_name, option_value in section.items():
                            if option_name.upper().startswith("ITEM_") and str(option_value).strip():
                                indexed.append((option_name.upper(), str(option_value)))
                        indexed.sort(key=lambda item: item[0])
                        values = [value for _name, value in indexed]
            self.dropdown_data[key] = self.normalize_dropdown_values(values)
        return True

    def database_distinct_values(self, table_name, column_name):
        allowed = {
            ("spareparts", "part_name"),
            ("spareparts", "part_type"),
            ("spareparts", "brand"),
            ("contacts", "devices"),
            ("contacts", "merek"),
            ("contacts", "model"),
            ("contacts", "masalah"),
            ("contacts", "status"),
            ("contacts", "teknisi"),
        }
        if (table_name, column_name) not in allowed or self.cursor is None:
            return []
        try:
            self.cursor.execute(
                f"SELECT DISTINCT {column_name} FROM {table_name} "
                f"WHERE {column_name} IS NOT NULL AND TRIM({column_name}) <> '' "
                f"ORDER BY {column_name}"
            )
            return self.normalize_dropdown_values(row[0] for row in self.cursor.fetchall())
        except sqlite3.Error:
            return []

    def combo_values(self, key):
        """Return stable system choices or external editable choices plus database values."""
        if key in self.SYSTEM_DROPDOWN_OPTIONS:
            return list(self.SYSTEM_DROPDOWN_OPTIONS[key])

        external = list(self.dropdown_data.get(key, []))
        database_values = []
        if key == "device_types":
            database_values = self.database_distinct_values("contacts", "devices")
        elif key == "brands":
            database_values = (
                self.database_distinct_values("contacts", "merek")
                + self.database_distinct_values("spareparts", "brand")
            )
        elif key == "models":
            database_values = self.database_distinct_values("contacts", "model")
        elif key == "issues":
            database_values = self.database_distinct_values("contacts", "masalah")
        elif key == "statuses":
            database_values = self.database_distinct_values("contacts", "status")
        elif key == "technicians":
            database_values = self.database_distinct_values("contacts", "teknisi")
        elif key == "parts":
            database_values = self.database_distinct_values("spareparts", "part_name")
        elif key == "sparepart_types":
            database_values = self.database_distinct_values("spareparts", "part_type")
        elif key == "sparepart_filter_types":
            return self.normalize_dropdown_values(
                ["ALL"]
                + self.dropdown_data.get("sparepart_types", [])
                + self.database_distinct_values("spareparts", "part_type")
            )
        elif key == "sparepart_filter_brands":
            return self.normalize_dropdown_values(
                ["ALL"]
                + self.dropdown_data.get("brands", [])
                + self.database_distinct_values("spareparts", "brand")
            )
        return self.normalize_dropdown_values(external + database_values)

    def add_combo_values(self, combo, values, translate_display=False):
        for value in values or []:
            source = str(value)
            display = self.t(source) if translate_display else source
            combo.addItem(display, source)

    def set_combo_values(self, combo, values, editable=False, translate_display=False, current_value=None):
        blocker = QSignalBlocker(combo)
        combo.clear()
        self.add_combo_values(combo, values, translate_display=translate_display)
        combo.setEditable(editable)
        if editable:
            combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        if current_value not in (None, ""):
            index = combo.findData(str(current_value))
            if index < 0:
                index = combo.findText(str(current_value), Qt.MatchFlag.MatchFixedString)
            if index >= 0:
                combo.setCurrentIndex(index)
            elif editable:
                combo.setCurrentText(str(current_value))
        elif combo.count() > 0:
            combo.setCurrentIndex(0)
        del blocker

    @staticmethod
    def combo_source_value(combo):
        value = combo.currentData()
        if value is None:
            value = combo.currentText()
        return str(value).strip()

    def refresh_dropdown_combos(self):
        """Refresh every combo box while preserving user selections."""
        combo_configs = [
            ("device_type", "device_types", True),
            ("brand_combo", "brands", True),
            ("model_input", "models", True),
            ("issue_combo", "issues", True),
            ("status_combo", "statuses", True),
            ("technician_input", "technicians", True),
            ("warranty_combo", "warranty_periods", True),
            ("parts_combo", "parts", True),
            ("search_combo", "search_fields", False),
            ("report_type", "report_types", False),
            ("warranty_filter_combo", "warranty_filters", False),
            ("sparepart_type", "sparepart_types", True),
            ("sparepart_brand", "brands", True),
            ("sparepart_filter_type", "sparepart_filter_types", False),
            ("sparepart_filter_brand", "sparepart_filter_brands", False),
        ]
        for attr_name, key, editable in combo_configs:
            combo = getattr(self, attr_name, None)
            if combo is None:
                continue
            current_value = self.combo_source_value(combo)
            self.set_combo_values(
                combo,
                self.combo_values(key),
                editable=editable,
                translate_display=True,
                current_value=current_value,
            )

    def refresh_translated_combo_items(self):
        self.refresh_dropdown_combos()

    def reload_dropdown_data(self, show_message=True):
        if show_message and not self.is_admin():
            self.require_admin()
            return False
        loaded = self.load_dropdown_data()
        self.refresh_dropdown_combos()
        self.setup_dropdown_file_watcher()
        if show_message and hasattr(self, "status_bar"):
            if loaded:
                self.status_bar.showMessage(f"DROPDOWN DATA RELOADED: {self.dropdown_data_path}")
            else:
                self.status_bar.showMessage(
                    f"DUMMY.INI NOT FOUND - OPEN SETTINGS > EDIT DUMMY DATA: {self.dropdown_data_path}"
                )

    def setup_dropdown_file_watcher(self):
        if self.dropdown_watcher is None:
            self.dropdown_watcher = QFileSystemWatcher(self)
            self.dropdown_watcher.fileChanged.connect(self.on_dropdown_file_changed)
            self.dropdown_watcher.directoryChanged.connect(self.on_dropdown_directory_changed)

        watched_files = set(self.dropdown_watcher.files())
        if os.path.isfile(self.dropdown_data_path) and self.dropdown_data_path not in watched_files:
            self.dropdown_watcher.addPath(self.dropdown_data_path)

        # Watching the application directory also detects the first creation of
        # dummy.ini and editors that save by atomically replacing the file.
        watched_directories = set(self.dropdown_watcher.directories())
        if os.path.isdir(self.app_dir) and self.app_dir not in watched_directories:
            self.dropdown_watcher.addPath(self.app_dir)

    def on_dropdown_file_changed(self, _path):
        # Editors commonly save by replacing the file, so re-add the watcher after reload.
        QTimer.singleShot(200, lambda: self.reload_dropdown_data(show_message=False))

    def on_dropdown_directory_changed(self, _path):
        if os.path.isfile(self.dropdown_data_path):
            QTimer.singleShot(250, lambda: self.reload_dropdown_data(show_message=False))
    def initUI(self):
        self.setWindowTitle('Computer Service Manager 4.4')
        self.load_window_geometry()
        self.setMinimumSize(1180, 720)
        
        # Menu Bar
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('FILE')
        
        self.import_action = QAction('IMPORT USER DATA FROM EXCEL OR CSV', self)
        self.import_action.triggered.connect(self.import_data_menu)
        file_menu.addAction(self.import_action)

        self.import_reference_action = QAction(
            'IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV', self
        )
        self.import_reference_action.triggered.connect(
            self.import_suppliers_and_spareparts_data
        )
        file_menu.addAction(self.import_reference_action)
        
        self.export_action = QAction('EXPORT TO EXCEL', self)
        self.export_action.triggered.connect(self.export_data_menu)
        file_menu.addAction(self.export_action)


        self.export_suppliers_action = QAction('EXPORT SUPPLIERS TO EXCEL', self)
        self.export_suppliers_action.triggered.connect(lambda: self.export_table_to_excel('suppliers'))
        file_menu.addAction(self.export_suppliers_action)

        self.export_spareparts_action = QAction('EXPORT PARTS TO EXCEL', self)
        self.export_spareparts_action.triggered.connect(lambda: self.export_table_to_excel('spareparts'))
        file_menu.addAction(self.export_spareparts_action)

        
        file_menu.addSeparator()
        
        self.backup_action = QAction('BACKUP DATABASE', self)
        self.backup_action.triggered.connect(self.backup_database)
        file_menu.addAction(self.backup_action)
        
        self.restore_action = QAction('RESTORE DATABASE', self)
        self.restore_action.triggered.connect(self.restore_database)
        file_menu.addAction(self.restore_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction('EXIT', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Settings menu
        settings_menu = menubar.addMenu('SETTINGS')
        
        self.db_path_action = QAction('DATABASE PATH', self)
        self.db_path_action.triggered.connect(self.set_database_path)
        settings_menu.addAction(self.db_path_action)

        self.currency_action = QAction('CURRENCY', self)
        self.currency_action.triggered.connect(self.change_currency_setting)
        settings_menu.addAction(self.currency_action)

        language_action = QAction('LANGUAGE / LOCALE', self)
        language_action.triggered.connect(self.change_language_setting)
        settings_menu.addAction(language_action)

        theme_menu = settings_menu.addMenu('THEME')
        self.theme_action_group = QActionGroup(self)
        self.theme_action_group.setExclusive(True)
        self.theme_actions = {}
        current_theme = self.get_theme_code()
        for theme_code in THEME_ORDER:
            theme_data = THEME_OPTIONS[theme_code]
            theme_action = QAction(theme_data["label"], self)
            theme_action.setCheckable(True)
            theme_action.setData(theme_code)
            theme_action.setChecked(theme_code == current_theme)
            theme_action.triggered.connect(
                lambda checked=False, code=theme_code: self.set_theme(code)
            )
            self.theme_action_group.addAction(theme_action)
            theme_menu.addAction(theme_action)
            self.theme_actions[theme_code] = theme_action

        self.reload_dropdown_action = QAction('RELOAD DROPDOWN DATA', self)
        self.reload_dropdown_action.triggered.connect(self.reload_dropdown_data)
        settings_menu.addAction(self.reload_dropdown_action)

        self.create_dummy_action = QAction('OPEN DUMMY CREATOR', self)
        self.create_dummy_action.setToolTip('OPEN DUMMY CREATOR')
        self.create_dummy_action.triggered.connect(self.open_external_dummy_creator)
        settings_menu.addAction(self.create_dummy_action)

        settings_menu.addSeparator()
        self.dummy_editor_action = QAction('EDIT DUMMY DATA', self)
        self.dummy_editor_action.setToolTip('OPEN DUMMY DATA EDITOR')
        self.dummy_editor_action.triggered.connect(self.open_dummy_data_editor)
        settings_menu.addAction(self.dummy_editor_action)
        


        # Help menu
        help_menu = menubar.addMenu('HELP')

        help_contents_action = QAction('HELP CONTENTS', self)
        help_contents_action.setStatusTip('OPEN APPLICATION HELP CONTENTS')
        help_contents_action.triggered.connect(
            lambda checked=False: self.open_help_contents("index.html")
        )
        help_menu.addAction(help_contents_action)

        current_help_action = QAction('CURRENT TAB HELP', self)
        current_help_action.setShortcut(QKeySequence('F1'))
        current_help_action.setStatusTip('OPEN HELP FOR THE ACTIVE TAB')
        current_help_action.triggered.connect(self.open_current_tab_help)
        help_menu.addAction(current_help_action)

        help_menu.addSeparator()
        about_action = QAction('ABOUT', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        # Central widget with tabs


        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)
        
        self.contacts_tab = QWidget()
        self.suppliers_tab = QWidget()
        self.reports_tab = QWidget()
        self.warranty_tab = QWidget()  # New warranty tracking tab
        self.spareparts_tab = QWidget()  # New spareparts tab
        
        self.tab_widget.addTab(self.contacts_tab, "USER DATA")
        self.tab_widget.addTab(self.suppliers_tab, "SUPPLIERS")
        self.tab_widget.addTab(self.reports_tab, "REPORTS")
        self.tab_widget.addTab(self.warranty_tab, "WARRANTY TRACKING")
        self.tab_widget.addTab(self.spareparts_tab, "PARTS")  # Add spareparts tab
        
        self.setup_contacts_tab()
        self.setup_suppliers_tab()
        self.setup_reports_tab()
        self.setup_warranty_tab()
        self.setup_spareparts_tab()  # Setup spareparts tab
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(self.t("READY"))
        self.apply_role_permissions()

    # ----------------------------
    # Small UI helpers (refactor)
    # ----------------------------
    def set_status(self, message: str):
        try:
            if hasattr(self, "status_bar") and self.status_bar:
                self.status_bar.showMessage(message)
        except Exception:
            pass

    def ui_info(self, title: str, message: str):
        QMessageBox.information(self, title, message)

    def ui_warn(self, title: str, message: str):
        QMessageBox.warning(self, title, message)

    def ui_error(self, title: str, message: str):
        QMessageBox.critical(self, title, message)

    def open_url(self, url: str):
        """Open URL in default browser (safe no-op if empty)."""
        url = (url or "").strip()
        if not url:
            return
        if not (url.lower().startswith("http://") or url.lower().startswith("https://")):
            url = "https://" + url
        QDesktopServices.openUrl(QUrl(url))

    def apply_app_icon(self):
        """Apply the requested application/window icon when the .ico file exists."""
        icon_path = get_app_file_path("app.ico")
        if not os.path.exists(icon_path):
            icon_path = DEFAULT_APP_ICON_PATH
        try:
            if icon_path and os.path.exists(icon_path):
                icon = QIcon(icon_path)
                self.setWindowIcon(icon)
                app = QApplication.instance()
                if app is not None:
                    app.setWindowIcon(icon)
        except Exception:
            pass

    def get_theme_code(self):
        code = str(self.settings.get('theme_code', DEFAULT_THEME_CODE)).strip().lower()
        return code if code in THEME_OPTIONS else DEFAULT_THEME_CODE

    def get_theme_label(self, theme_code=None):
        code = str(theme_code or self.get_theme_code()).strip().lower()
        data = THEME_OPTIONS.get(code, THEME_OPTIONS[DEFAULT_THEME_CODE])
        return self.t(data["label"])

    def apply_theme_setting_to_ui(self):
        app = QApplication.instance()
        code = self.get_theme_code()
        if app is not None:
            apply_application_theme(app, code)
        for theme_code, action in getattr(self, "theme_actions", {}).items():
            action.setChecked(theme_code == code)

        # Rebuild programmatically drawn calendar icons after a theme change so
        # they always keep the correct foreground and accent colors.
        for button_name in (
            "current_service_date_btn",
            "from_date_icon_btn",
            "to_date_icon_btn",
            "warranty_from_date_icon_btn",
            "warranty_to_date_icon_btn",
        ):
            button = getattr(self, button_name, None)
            if isinstance(button, QAbstractButton):
                button.setIcon(create_calendar_icon())
                button.setIconSize(QSize(18, 18))
        return code

    def set_theme(self, theme_code):
        code = str(theme_code or DEFAULT_THEME_CODE).strip().lower()
        if code not in THEME_OPTIONS:
            code = DEFAULT_THEME_CODE
        self.settings['theme_code'] = code
        self.save_settings()
        self.apply_theme_setting_to_ui()
        if hasattr(self, "status_bar") and self.status_bar:
            self.status_bar.showMessage(f"{self.t('THEME')}: {self.get_theme_label(code)}")
        return True

    def get_language_code(self):
        code = str(self.settings.get('language_code', DEFAULT_LANGUAGE_CODE)).strip()
        if code not in LANGUAGE_OPTIONS:
            code = DEFAULT_LANGUAGE_CODE
        return code

    def get_language_label(self):
        cfg = get_language_config(self.get_language_code())
        return f"{cfg['native']} / {cfg['label']}"

    def get_language_currency_code(self, language_code=None):
        cfg = get_language_config(language_code or self.get_language_code())
        currency_code = str(cfg.get("currency", DEFAULT_CURRENCY_CODE)).strip().upper()
        return currency_code if currency_code in CURRENCY_OPTIONS else DEFAULT_CURRENCY_CODE

    def t(self, text):
        return translate_ui_text(text, self.get_language_code())

    def format_number(self, value, decimals=0):
        return format_number_value(value, self.get_language_code(), decimals=decimals)

    def format_date_for_display(self, db_date):
        if not db_date:
            return ""
        try:
            date_obj = datetime.strptime(str(db_date).split()[0], "%Y-%m-%d")
            return date_obj.strftime(get_language_config(self.get_language_code()).get("python_date_format", "%d-%m-%Y"))
        except Exception:
            return str(db_date)

    def _remember_source_text(self, obj, prop_name, current_text):
        source_prop = f"_source_{prop_name}"
        source_text = obj.property(source_prop)
        if source_text is None:
            source_text = current_text
            obj.setProperty(source_prop, source_text)
        return source_text

    def _translate_action(self, action):
        if action is None or action.isSeparator():
            return
        text = action.text()
        if text:
            source_text = self._remember_source_text(action, "text", text)
            action.setText(self.t(source_text))
        for prop_name, getter, setter in (
            ("tooltip", action.toolTip, action.setToolTip),
            ("status_tip", action.statusTip, action.setStatusTip),
            ("whats_this", action.whatsThis, action.setWhatsThis),
        ):
            value = getter()
            if value:
                source_value = self._remember_source_text(action, prop_name, value)
                setter(self.t(source_value))
        if action.menu() is not None:
            self._translate_menu(action.menu())

    def _translate_menu(self, menu):
        if menu is None:
            return
        title = menu.title()
        if title:
            source_title = self._remember_source_text(menu, "title", title)
            menu.setTitle(self.t(source_title))
        for action in menu.actions():
            self._translate_action(action)

    def _translate_widget_texts(self, root_widget):
        widgets = [root_widget] if isinstance(root_widget, QWidget) else []
        if isinstance(root_widget, QWidget):
            widgets.extend(root_widget.findChildren(QWidget))

        for widget in widgets:
            window_title = widget.windowTitle()
            if window_title:
                source_window_title = self._remember_source_text(widget, "window_title", window_title)
                widget.setWindowTitle(self.t(source_window_title))

            if isinstance(widget, QGroupBox):
                title = widget.title()
                if title:
                    source_title = self._remember_source_text(widget, "title", title)
                    widget.setTitle(self.t(source_title))

            if isinstance(widget, QAbstractButton):
                text = widget.text()
                if text:
                    source_text = self._remember_source_text(widget, "text", text)
                    widget.setText(self.t(source_text))
            elif isinstance(widget, QLabel):
                text = widget.text()
                if text:
                    source_text = self._remember_source_text(widget, "text", text)
                    widget.setText(self.t(source_text))

            if isinstance(widget, (QLineEdit, QTextEdit)):
                placeholder = widget.placeholderText()
                if placeholder:
                    source_placeholder = self._remember_source_text(widget, "placeholder", placeholder)
                    widget.setPlaceholderText(self.t(source_placeholder))

            tooltip = widget.toolTip()
            if tooltip:
                source_tooltip = self._remember_source_text(widget, "tooltip", tooltip)
                widget.setToolTip(self.t(source_tooltip))
            status_tip = widget.statusTip()
            if status_tip:
                source_status_tip = self._remember_source_text(widget, "status_tip", status_tip)
                widget.setStatusTip(self.t(source_status_tip))
            whats_this = widget.whatsThis()
            if whats_this:
                source_whats_this = self._remember_source_text(widget, "whats_this", whats_this)
                widget.setWhatsThis(self.t(source_whats_this))

            if isinstance(widget, QComboBox):
                try:
                    placeholder = widget.placeholderText()
                    if placeholder:
                        source_placeholder = self._remember_source_text(widget, "placeholder", placeholder)
                        widget.setPlaceholderText(self.t(source_placeholder))
                except Exception:
                    pass

    def _translate_tabs(self, root_widget=None):
        root = root_widget if isinstance(root_widget, QWidget) else self
        for tab_widget in root.findChildren(QTabWidget):
            for index in range(tab_widget.count()):
                prop_name = f"_source_tab_{index}"
                source_text = tab_widget.property(prop_name)
                if source_text is None:
                    source_text = tab_widget.tabText(index)
                    tab_widget.setProperty(prop_name, source_text)
                tab_widget.setTabText(index, self.t(source_text))

    def _translate_table_headers(self, root_widget=None):
        source_role = Qt.ItemDataRole.UserRole + 730
        root = root_widget if isinstance(root_widget, QWidget) else self
        tables = [root] if isinstance(root, QTableWidget) else []
        tables.extend(root.findChildren(QTableWidget))
        for table in tables:
            for col in range(table.columnCount()):
                item = table.horizontalHeaderItem(col)
                if item is None:
                    continue
                source_text = item.data(source_role)
                if source_text is None:
                    source_text = item.text()
                    item.setData(source_role, source_text)
                item.setText(self.t(source_text))

    def apply_language_setting_to_ui(self, root_widget=None):
        cfg = get_language_config(self.get_language_code())
        app = QApplication.instance()
        if app is not None:
            app.setLayoutDirection(Qt.LayoutDirection.RightToLeft if cfg.get("rtl") else Qt.LayoutDirection.LeftToRight)

        date_format = cfg.get("date_format", "dd-MM-yyyy")
        for widget_name in (
            "service_date", "warranty_end_date", "from_date", "to_date",
            "warranty_from_date", "warranty_to_date",
        ):
            widget = getattr(self, widget_name, None)
            if isinstance(widget, QDateEdit):
                widget.setDisplayFormat(date_format)

        root = root_widget or self
        self._translate_widget_texts(root)
        self._translate_tabs(root)
        self._translate_table_headers(root)

        if root_widget is None:
            if self.menuBar() is not None:
                for action in self.menuBar().actions():
                    self._translate_action(action)
            for action in self.findChildren(QAction):
                self._translate_action(action)
            self.refresh_translated_combo_items()

    def refresh_language_sensitive_views(self):
        try:
            self.load_contacts()
        except Exception:
            pass
        try:
            self.load_warranty_data()
        except Exception:
            pass
        try:
            self.load_spareparts()
        except Exception:
            pass
        try:
            if hasattr(self, "report_table") and self.report_table.columnCount() > 0:
                self.generate_report()
        except Exception:
            pass

    def prompt_locale_on_launch_if_needed(self):
        raw_language_code = str(self.settings.get('language_code', '')).strip()
        if self.is_first_run or not raw_language_code:
            self.select_language_setting(first_run=True)
        else:
            self.apply_language_setting_to_ui()

        if not str(self.settings.get('currency_code', '')).strip():
            self.settings['currency_code'] = self.get_language_currency_code()
            self.save_settings()
        self.apply_currency_setting_to_ui()

    def change_language_setting(self):
        self.select_language_setting(first_run=False)

    def select_language_setting(self, first_run=False):
        labels = [
            f"{code} - {data['native']} / {data['label']} ({data['currency']})"
            for code, data in LANGUAGE_OPTIONS.items()
        ]
        codes = list(LANGUAGE_OPTIONS.keys())
        current_code = self.get_language_code()
        current_index = codes.index(current_code) if current_code in codes else 0
        message = "SELECT APPLICATION LANGUAGE AND LOCALE:"

        selected_label, ok = QInputDialog.getItem(
            self,
            self.t("SELECT LANGUAGE / LOCALE"),
            self.t(message),
            labels,
            current_index,
            False
        )

        if not ok:
            if first_run and not str(self.settings.get('language_code', '')).strip():
                self.settings['language_code'] = DEFAULT_LANGUAGE_CODE
                self.settings['currency_code'] = self.get_language_currency_code(DEFAULT_LANGUAGE_CODE)
                self.save_settings()
                self.apply_language_setting_to_ui()
                self.apply_currency_setting_to_ui()
            return False

        selected_code = selected_label.split(" - ", 1)[0].strip()
        if selected_code not in LANGUAGE_OPTIONS:
            selected_code = DEFAULT_LANGUAGE_CODE

        self.settings['language_code'] = selected_code
        self.settings['currency_code'] = self.get_language_currency_code(selected_code)
        self.save_settings()
        self.apply_language_setting_to_ui()
        self.apply_currency_setting_to_ui()
        self.refresh_currency_views()
        self.refresh_language_sensitive_views()
        self.apply_role_permissions()
        if hasattr(self, "status_bar") and self.status_bar:
            self.status_bar.showMessage(
                f"{self.t('LANGUAGE')}: {self.get_language_label()} | {self.t('CURRENCY')}: {self.get_currency_label()}"
            )

        if not first_run:
            QMessageBox.information(
                self,
                self.t("LANGUAGE / LOCALE"),
                f"{self.get_language_label()}\n{self.get_currency_label()}"
            )
        return True

    def get_currency_code(self):
        code = str(self.settings.get('currency_code', DEFAULT_CURRENCY_CODE)).strip().upper()
        if code not in CURRENCY_OPTIONS:
            code = DEFAULT_CURRENCY_CODE
        return code

    def get_currency_symbol(self):
        return get_currency_config(self.get_currency_code())["symbol"]

    def get_currency_label(self):
        option = get_currency_config(self.get_currency_code())
        return f"{option['label']} ({self.get_currency_code()})"


    def apply_currency_setting_to_ui(self):
        code = self.get_currency_code()
        for widget_name in ("price_input", "sparepart_price"):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                if hasattr(widget, "set_currency_code"):
                    widget.set_currency_code(code)
                elif hasattr(widget, "set_currency_symbol"):
                    widget.set_currency_symbol(self.get_currency_symbol())

    def refresh_currency_views(self):
        try:
            self.load_contacts()
        except Exception:
            pass

        try:
            self.load_spareparts()
        except Exception:
            pass

        try:
            if hasattr(self, "report_table") and self.report_table.columnCount() > 0:
                self.generate_report()
        except Exception:
            pass


    def change_currency_setting(self):
        if not self.require_admin():
            return
        self.select_currency_setting(first_run=False)

    def select_currency_setting(self, first_run=False):
        labels = [
            f"{code} - {data['label']} ({data['symbol']})"
            for code, data in CURRENCY_OPTIONS.items()
        ]
        codes = list(CURRENCY_OPTIONS.keys())
        current_code = self.get_currency_code()
        current_index = codes.index(current_code) if current_code in codes else 0
        message = "SELECT DEFAULT CURRENCY FOR THE APPLICATION:" if first_run else "SELECT APPLICATION CURRENCY:"

        selected_label, ok = QInputDialog.getItem(
            self,
            self.t("CURRENCY"),
            self.t(message),
            labels,
            current_index,
            False
        )

        if not ok:
            if first_run and not str(self.settings.get('currency_code', '')).strip():
                self.settings['currency_code'] = DEFAULT_CURRENCY_CODE
                self.save_settings()
                self.apply_currency_setting_to_ui()
            return False

        selected_code = selected_label.split(" - ", 1)[0].strip().upper()
        if selected_code not in CURRENCY_OPTIONS:
            selected_code = DEFAULT_CURRENCY_CODE

        self.settings['currency_code'] = selected_code
        self.save_settings()
        self.apply_currency_setting_to_ui()
        self.refresh_currency_views()

        if hasattr(self, "status_bar") and self.status_bar:
            self.status_bar.showMessage(f"{self.t('CURRENCY')}: {self.get_currency_label()}")

        if not first_run:
            QMessageBox.information(
                self,
                self.t("CURRENCY"),
                f"{self.get_currency_label()}"
            )
        return True

    def format_currency(self, value):
        """Format numeric value using the selected currency display."""
        return format_currency_value(value, self.get_currency_code())

    # ----------------------------
    # Added useful features
    # ----------------------------
    def export_table_to_excel(self, table_name: str):
        """Export a DB table (suppliers/spareparts) to a chosen Excel file."""
        if not self.require_admin():
            return
        table_name = (table_name or "").strip().lower()
        if table_name not in {"suppliers", "spareparts"}:
            self.ui_warn("EXPORT", "UNKNOWN TABLE TO EXPORT.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = self.get_export_file_path(
            f"{table_name}_export_{timestamp}.xlsx",
            f"SAVE {table_name.upper()} EXPORT"
        )
        if not file_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = table_name.upper()

            if table_name == "suppliers":
                headers = ["ID (ROWID)", "NAME", "CONTACT", "PHONE", "EMAIL", "ADDRESS", "CREATED DATE"]
                sql = """SELECT rowid, name, contact_person, phone, email, address, created_date
                         FROM suppliers ORDER BY created_date DESC"""
            else:
                headers = ["ID (ROWID)", "PART NAME", "TYPE", "BRAND", "QTY", "UNIT PRICE",
                           "SUPPLIER", "MIN STOCK", "LOCATION", "NOTES", "CREATED DATE"]
                sql = """SELECT rowid, part_name, part_type, brand, quantity, unit_price,
                         supplier, min_stock, location, notes, created_date
                         FROM spareparts ORDER BY created_date DESC"""

            for col, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=col, value=self.t(h))
                try:
                    c.font = openpyxl.styles.Font(bold=True)
                except Exception:
                    pass

            self.cursor.execute(sql)
            rows = self.cursor.fetchall()
            for r_i, row in enumerate(rows, 2):
                for c_i, val in enumerate(row, 1):
                    ws.cell(row=r_i, column=c_i, value=val)

            for column_cells in ws.columns:
                length = 0
                col_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        if cell.value is not None:
                            length = max(length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 60)

            wb.save(file_path)
            self.ui_info("EXPORT COMPLETE", f"EXPORTED TO:\n{file_path}")
            self.set_status(f"EXPORTED {table_name} TO EXCEL")
        except Exception as e:
            self.ui_error("EXPORT ERROR", f"FAILED TO EXPORT {table_name}:\n{e}")
            self.set_status(f"EXPORT FAILED: {e}")

    def open_current_maps(self):
        """Open the Google Maps link from the form or selected table row."""
        url = ""
        try:
            url = (self.maps_input.text() or "").strip()
        except Exception:
            pass

        if not url:
            try:
                row = self.contact_table.currentRow()
                if row >= 0:
                    item = self.contact_table.item(row, 14)  # GOOGLE MAP column
                    url = (item.text() if item else "").strip()
            except Exception:
                url = ""

        if not url:
            self.set_status("GOOGLE MAP LINK EMPTY")
            return
        self.open_url(url)

    
    def copy_current_maps(self):
        """Copy the Google Maps link from the form or selected table row to clipboard.
        If empty, the clipboard will be cleared (no warning dialog).
        """
        url = ""
        try:
            url = (self.maps_input.text() or "").strip()
        except Exception:
            pass

        if not url:
            try:
                row = self.contact_table.currentRow()
                if row >= 0:
                    item = self.contact_table.item(row, 14)  # GOOGLE MAP column
                    url = (item.text() if item else "").strip()
            except Exception:
                url = ""

        try:
            QApplication.clipboard().setText(url)
            if url:
                self.set_status("GOOGLE MAP LINK COPIED")
            else:
                self.set_status("GOOGLE MAP LINK EMPTY (CLIPBOARD CLEARED)")
        except Exception:
            pass

    def contact_table_double_click(self, row: int, col: int):
        """Double click: open map if clicking map column."""
        try:
            if col == 14:
                item = self.contact_table.item(row, col)
                if item:
                    self.open_url(item.text().strip())
        except Exception:
            pass

    def show_contact_context_menu(self, pos):
        """Right click menu for contact table (copy/open links)."""
        try:
            row = self.contact_table.rowAt(pos.y())
            if row < 0:
                return
            menu = QMenu(self)

            def copy_col(c):
                it = self.contact_table.item(row, c)
                if it and it.text():
                    QApplication.clipboard().setText(it.text())
                    self.set_status("COPIED TO CLIPBOARD")

            act_copy_phone = QAction(self.t("COPY PHONE"), self)
            act_copy_phone.triggered.connect(lambda: copy_col(1))
            menu.addAction(act_copy_phone)

            act_copy_addr = QAction(self.t("COPY ADDRESS"), self)
            act_copy_addr.triggered.connect(lambda: copy_col(2))
            menu.addAction(act_copy_addr)

            menu.addSeparator()

            act_copy_map = QAction(self.t("COPY GOOGLE MAP"), self)


            act_copy_map.triggered.connect(lambda: copy_col(14))


            menu.addAction(act_copy_map)



            act_open_map = QAction(self.t("OPEN GOOGLE MAP"), self)
            act_open_map.triggered.connect(lambda: self.open_url((self.contact_table.item(row, 14).text() if self.contact_table.item(row, 14) else "")))
            menu.addAction(act_open_map)

            act_copy_invoice = QAction(self.t("COPY INVOICE"), self)
            act_copy_invoice.triggered.connect(lambda: copy_col(17))
            menu.addAction(act_copy_invoice)

            menu.exec(self.contact_table.viewport().mapToGlobal(pos))
        except Exception:
            pass

    def _stop_runtime_services(self):
        """Stop timers and other app-owned runtime services."""
        for timer_name in ("warranty_timer",):
            timer = getattr(self, timer_name, None)
            if timer is not None:
                try:
                    if timer.isActive():
                        timer.stop()
                    timer.deleteLater()
                except RuntimeError:
                    pass
                except Exception:
                    pass
                setattr(self, timer_name, None)

    def _close_database_connection(self):
        """Commit and close the SQLite connection exactly once."""
        conn = getattr(self, "conn", None)
        if conn is None:
            self.cursor = None
            return

        try:
            conn.commit()
        except sqlite3.Error:
            pass
        except Exception:
            pass

        try:
            conn.close()
        except Exception:
            pass
        finally:
            self.conn = None
            self.cursor = None

    def shutdown(self):
        """Save state and release app resources. Safe to call more than once."""
        if getattr(self, "_is_shutting_down", False):
            return
        self._is_shutting_down = True

        try:
            self.save_window_geometry()
        except Exception:
            pass

        if self.settings.get('auto_backup_on_exit', True):
            try:
                if self.conn:
                    self.conn.commit()
                self.backup_database(silent=True)
            except Exception:
                pass

        self._stop_runtime_services()
        self._close_database_connection()


    

    def import_data_menu(self):
        """Menu handler for the supported XLSX/CSV importer."""
        if not self.require_admin():
            return
        return self.import_data()

    def export_data_menu(self):
        """Menu handler for the standard Excel exporter."""
        if not self.require_admin():
            return
        return self.export_data()


    def read_settings_parser(self):
        parser = configparser.ConfigParser()
        parser.optionxform = str
        if os.path.exists(self.settings_ini_path):
            parser.read(self.settings_ini_path, encoding="utf-8")
        return parser

    def write_settings_parser(self, parser):
        os.makedirs(os.path.dirname(self.settings_ini_path), exist_ok=True)
        with open(self.settings_ini_path, "w", encoding="utf-8") as f:
            parser.write(f)

    def migrate_registry_settings_to_ini(self):
        """Migrate legacy registry settings into setting.ini if available."""
        if os.path.exists(self.settings_ini_path) or winreg is None:
            return False

        migrated = False
        parser = configparser.ConfigParser()
        parser.optionxform = str
        parser["SETTINGS"] = {}
        parser["WINDOW"] = {}

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Service Manager\Settings")
            parser["SETTINGS"]["DatabasePath"] = str(winreg.QueryValueEx(key, "DatabasePath")[0])
            parser["SETTINGS"]["BackupDir"] = str(winreg.QueryValueEx(key, "BackupDir")[0])
            try:
                parser["SETTINGS"]["AutoBackupOnExit"] = "1" if bool(winreg.QueryValueEx(key, "AutoBackupOnExit")[0]) else "0"
            except Exception:
                parser["SETTINGS"]["AutoBackupOnExit"] = "1"
            try:
                parser["SETTINGS"]["DbLocationInitialized"] = "1" if bool(winreg.QueryValueEx(key, "DbLocationInitialized")[0]) else "0"
            except Exception:
                parser["SETTINGS"]["DbLocationInitialized"] = "0"
            winreg.CloseKey(key)
            migrated = True
        except Exception:
            pass

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Service Manager")
            parser["WINDOW"]["WindowX"] = str(winreg.QueryValueEx(key, "WindowX")[0])
            parser["WINDOW"]["WindowY"] = str(winreg.QueryValueEx(key, "WindowY")[0])
            parser["WINDOW"]["WindowWidth"] = str(winreg.QueryValueEx(key, "WindowWidth")[0])
            parser["WINDOW"]["WindowHeight"] = str(winreg.QueryValueEx(key, "WindowHeight")[0])
            winreg.CloseKey(key)
            migrated = True
        except Exception:
            pass

        if migrated:
            self.write_settings_parser(parser)

        return migrated

    def load_window_geometry(self):
        """Load window position and size from setting.ini next to the app."""
        try:
            parser = self.read_settings_parser()
            if parser.has_section("WINDOW"):
                x = parser.getint("WINDOW", "WindowX", fallback=50)
                y = parser.getint("WINDOW", "WindowY", fallback=50)
                width = parser.getint("WINDOW", "WindowWidth", fallback=1400)
                height = parser.getint("WINDOW", "WindowHeight", fallback=800)
                self.setGeometry(x, y, width, height)
            else:
                self.setGeometry(50, 50, 1400, 800)
        except Exception:
            self.setGeometry(50, 50, 1400, 800)

    def save_window_geometry(self):
        """Save window position and size to setting.ini next to the app."""
        try:
            parser = self.read_settings_parser()
            if not parser.has_section("WINDOW"):
                parser.add_section("WINDOW")
            geometry = self.geometry()
            parser["WINDOW"]["WindowX"] = str(geometry.x())
            parser["WINDOW"]["WindowY"] = str(geometry.y())
            parser["WINDOW"]["WindowWidth"] = str(geometry.width())
            parser["WINDOW"]["WindowHeight"] = str(geometry.height())
            self.write_settings_parser(parser)
        except Exception as e:
            print(f"Failed to save window geometry: {e}")

    def load_settings(self):
        """Load settings from setting.ini next to the app."""
        self.is_first_run = False

        defaults = {
            'db_path': self.default_db_path,
            'backup_dir': self.default_backup_dir,
            'import_export_dir': str(Path.home() / "Downloads"),
            'auto_backup_on_exit': True,
            'db_location_initialized': False,
            'currency_code': '',
            'language_code': '',
            'theme_code': DEFAULT_THEME_CODE,
        }

        migrated = self.migrate_registry_settings_to_ini()
        ini_exists = os.path.exists(self.settings_ini_path)

        if not ini_exists and not migrated:
            self.is_first_run = True
            self.settings = defaults.copy()
            self.save_settings()
            return

        parser = self.read_settings_parser()
        settings_section = parser["SETTINGS"] if parser.has_section("SETTINGS") else {}

        self.settings['db_path'] = settings_section.get("DatabasePath", defaults['db_path'])
        self.settings['backup_dir'] = settings_section.get("BackupDir", defaults['backup_dir'])
        self.settings['import_export_dir'] = settings_section.get("ImportExportDir", defaults['import_export_dir'])
        self.settings['auto_backup_on_exit'] = str(settings_section.get("AutoBackupOnExit", "1")).strip().lower() in {"1", "true", "yes", "on"}
        self.settings['db_location_initialized'] = str(settings_section.get("DbLocationInitialized", "0")).strip().lower() in {"1", "true", "yes", "on"}

        if parser.has_section("SETTINGS") and parser.has_option("SETTINGS", "CurrencyCode"):
            raw_currency_code = settings_section.get("CurrencyCode", "")
        else:
            raw_currency_code = ""
        self.settings['currency_code'] = str(raw_currency_code).strip().upper()

        if parser.has_section("SETTINGS") and parser.has_option("SETTINGS", "LanguageCode"):
            raw_language_code = settings_section.get("LanguageCode", "")
        else:
            raw_language_code = ""
        self.settings['language_code'] = str(raw_language_code).strip()

        raw_theme_code = settings_section.get("ThemeCode", defaults['theme_code'])
        self.settings['theme_code'] = str(raw_theme_code).strip().lower()

        if not self.settings['db_path']:
            self.settings['db_path'] = defaults['db_path']
        if not self.settings['backup_dir']:
            self.settings['backup_dir'] = defaults['backup_dir']
        if not self.settings['import_export_dir'] or not self.drive_exists_for_path(self.settings['import_export_dir']):
            self.settings['import_export_dir'] = defaults['import_export_dir']
        if self.settings['currency_code'] and self.settings['currency_code'] not in CURRENCY_OPTIONS:
            self.settings['currency_code'] = ''
        if self.settings['language_code'] and self.settings['language_code'] not in LANGUAGE_OPTIONS:
            self.settings['language_code'] = ''
        if self.settings['theme_code'] not in THEME_OPTIONS:
            self.settings['theme_code'] = DEFAULT_THEME_CODE

    def save_settings(self):
        """Save settings to setting.ini next to the app."""
        try:
            parser = self.read_settings_parser()
            if not parser.has_section("SETTINGS"):
                parser.add_section("SETTINGS")
            parser["SETTINGS"]["DatabasePath"] = self.settings.get('db_path', self.default_db_path)
            parser["SETTINGS"]["BackupDir"] = self.settings.get('backup_dir', self.default_backup_dir)
            parser["SETTINGS"]["ImportExportDir"] = self.settings.get('import_export_dir', str(Path.home() / "Downloads"))
            parser["SETTINGS"]["AutoBackupOnExit"] = "1" if self.settings.get('auto_backup_on_exit', True) else "0"
            parser["SETTINGS"]["DbLocationInitialized"] = "1" if self.settings.get('db_location_initialized', False) else "0"
            parser["SETTINGS"]["CurrencyCode"] = str(self.settings.get('currency_code', '')).strip().upper()
            parser["SETTINGS"]["LanguageCode"] = str(self.settings.get('language_code', '')).strip()
            parser["SETTINGS"]["ThemeCode"] = str(self.settings.get('theme_code', DEFAULT_THEME_CODE)).strip().lower()
            self.write_settings_parser(parser)
        except Exception as e:
            print(f"Failed to save settings: {e}")

    def drive_exists_for_path(self, path):
        """Return True when the drive/root for a path exists."""
        path = (path or "").strip()
        if not path:
            return False
        drive, _ = os.path.splitdrive(path)
        if drive:
            return os.path.exists(drive + "\\")
        return True

    def get_import_export_dir(self):
        """Return the best folder for import/export dialogs."""
        candidates = [
            self.settings.get('import_export_dir', ''),
            self.settings.get('backup_dir', ''),
            os.path.dirname(self.settings.get('db_path', '')),
            str(Path.home() / "Downloads"),
            str(Path.home() / "Documents"),
            str(Path.home()),
        ]

        for candidate in candidates:
            candidate = os.path.normpath(str(candidate).strip()) if candidate else ""
            if not candidate or not self.drive_exists_for_path(candidate):
                continue
            if os.path.isfile(candidate):
                candidate = os.path.dirname(candidate)
            if os.path.isdir(candidate):
                return candidate

        return str(Path.home())

    def remember_import_export_path(self, file_path):
        """Persist the folder used by the last import/export action."""
        folder = os.path.dirname(os.path.normpath(file_path or ""))
        if folder and self.drive_exists_for_path(folder):
            self.settings['import_export_dir'] = folder
            self.save_settings()

    def get_import_file_path(self, title="SELECT FILE TO IMPORT"):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            title,
            self.get_import_export_dir(),
            "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.remember_import_export_path(file_path)
        return file_path

    def get_export_file_path(self, default_filename, title="SAVE EXPORT FILE"):
        initial_path = os.path.join(self.get_import_export_dir(), default_filename)
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            title,
            initial_path,
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if not file_path:
            return ""

        file_path = os.path.normpath(file_path)
        root, ext = os.path.splitext(file_path)
        if ext.lower() != ".xlsx":
            file_path = root + ".xlsx" if ext else file_path + ".xlsx"

        if not self.drive_exists_for_path(file_path):
            self.ui_warn("INVALID PATH", "THE DRIVE FOR EXPORT PATH DOES NOT EXIST.")
            return ""

        export_dir = os.path.dirname(file_path)
        if export_dir:
            os.makedirs(export_dir, exist_ok=True)

        self.remember_import_export_path(file_path)
        return file_path

    def candidate_database_filenames(self):
        return [
            "DATABASE USER.db",
            "database.db",
            "DATABASE.db",
            "service_manager.db",
        ]

    def find_existing_database_path(self):
        """Find an existing database file in common locations before creating a new one."""
        configured_db = os.path.normpath(str(self.settings.get('db_path', '')).strip())
        if configured_db and os.path.isfile(configured_db):
            return configured_db

        search_roots = []
        for candidate in [
            self.app_dir,
            os.path.dirname(configured_db) if configured_db else '',
            self.settings.get('backup_dir', ''),
            self.default_backup_dir,
            os.path.join(str(Path.home()), 'Documents', 'ServiceManagerData'),
            os.path.join(str(Path.home()), 'Documents'),
            os.path.join(str(Path.home()), 'Downloads'),
        ]:
            candidate = os.path.normpath(str(candidate).strip()) if candidate else ''
            if candidate and candidate not in search_roots and os.path.isdir(candidate):
                search_roots.append(candidate)

        filenames = {name.lower() for name in self.candidate_database_filenames()}

        for root in search_roots:
            for filename in self.candidate_database_filenames():
                full_path = os.path.join(root, filename)
                if os.path.isfile(full_path):
                    return os.path.normpath(full_path)

            try:
                for current_root, dirs, files in os.walk(root):
                    depth = os.path.relpath(current_root, root).count(os.sep)
                    if depth > 2:
                        dirs[:] = []
                        continue
                    for file_name in files:
                        if file_name.lower() in filenames:
                            return os.path.normpath(os.path.join(current_root, file_name))
            except Exception:
                continue

        return ""

    def apply_database_path(self, db_path):
        """Normalize and persist an existing database path."""
        if not db_path:
            return False

        db_path = os.path.normpath(db_path)
        db_dir = os.path.dirname(db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)

        self.settings['db_path'] = db_path
        self.settings['backup_dir'] = db_dir or self.settings.get('backup_dir', self.default_backup_dir)
        self.settings['db_location_initialized'] = True
        self.save_settings()
        return True

    def get_database_fallback_dir(self):
        """Return a safe fallback folder for database dialogs."""
        fallback_dir = self.default_backup_dir if self.drive_exists_for_path(self.default_db_path) else str(Path.home() / "Documents" / "ServiceManagerData")
        try:
            os.makedirs(fallback_dir, exist_ok=True)
        except Exception:
            fallback_dir = str(Path.home())
        return os.path.normpath(fallback_dir)

    def prompt_open_existing_database(self):
        """Ask the user to select an existing SQLite database file."""
        fallback_dir = self.get_database_fallback_dir()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "OPEN EXISTING DATABASE",
            fallback_dir,
            "Database Files (*.db *.sqlite *.sqlite3);;All Files (*)"
        )

        if not file_path:
            return False

        return self.apply_database_path(file_path)

    def prompt_open_or_create_database_first_launch(self):
        """At first launch ask the user to open an existing database or create a new one."""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("FIRST-TIME DATABASE SETUP")
        msg_box.setIcon(QMessageBox.Icon.Question)
        msg_box.setText("WHAT DO YOU WANT TO DO WITH THE DATABASE?")
        msg_box.setInformativeText("OPEN AN EXISTING DATABASE OR CREATE A NEW DATABASE FOR THIS APPLICATION.")

        open_btn = msg_box.addButton("OPEN EXISTING DATABASE", QMessageBox.ButtonRole.AcceptRole)
        create_btn = msg_box.addButton("CREATE NEW DATABASE", QMessageBox.ButtonRole.ActionRole)
        msg_box.setDefaultButton(create_btn)
        msg_box.exec()

        if msg_box.clickedButton() == open_btn:
            if self.prompt_open_existing_database():
                return True

            QMessageBox.information(
                self,
                "NO DATABASE SELECTED",
                "NO EXISTING DATABASE WAS SELECTED. THE APPLICATION WILL CONTINUE WITH NEW DATABASE CREATION."
            )

        self.prompt_for_database_directory(first_run=True, missing_drive=False)
        return True

    def prompt_for_database_directory(self, first_run=False, missing_drive=False):
        """Ask the user where the database should be stored."""
        fallback_dir = self.get_database_fallback_dir()

        if missing_drive:
            QMessageBox.information(
                self,
                "DATABASE LOCATION REQUIRED",
                "DRIVE D: WAS NOT FOUND.\nPLEASE CREATE OR SELECT A FOLDER FOR THE DATABASE."
            )
        elif first_run:
            QMessageBox.information(
                self,
                "SET DATABASE LOCATION",
                "FIRST-TIME SETUP:\nPLEASE CHOOSE A FOLDER TO SAVE THE NEW DATABASE."
            )

        selected_dir = QFileDialog.getExistingDirectory(
            self,
            "SELECT DATABASE FOLDER",
            fallback_dir,
            QFileDialog.Option.ShowDirsOnly
        )

        if not selected_dir:
            selected_dir = fallback_dir

        selected_dir = os.path.normpath(selected_dir)
        try:
            os.makedirs(selected_dir, exist_ok=True)
        except Exception:
            selected_dir = os.path.normpath(str(Path.home() / "Documents" / "ServiceManagerData"))
            os.makedirs(selected_dir, exist_ok=True)

        self.settings['db_path'] = os.path.join(selected_dir, "DATABASE USER.db")
        self.settings['backup_dir'] = selected_dir
        self.settings['db_location_initialized'] = True
        self.save_settings()

    def ensure_database_location_ready(self):
        """Ensure the configured database location is available before connecting."""
        db_path = self.settings.get('db_path', '')
        backup_dir = self.settings.get('backup_dir', '')
        missing_drive = (not self.drive_exists_for_path(db_path)) or (backup_dir and not self.drive_exists_for_path(backup_dir))

        if self.is_first_run:
            self.prompt_open_or_create_database_first_launch()
        else:
            existing_db = self.find_existing_database_path()
            if existing_db:
                self.apply_database_path(existing_db)
            elif not self.settings.get('db_location_initialized', False) or missing_drive:
                self.prompt_for_database_directory(first_run=False, missing_drive=missing_drive)

        db_dir = os.path.dirname(self.settings.get('db_path', ''))
        backup_dir = self.settings.get('backup_dir') or db_dir
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        if backup_dir:
            os.makedirs(backup_dir, exist_ok=True)

    def create_database(self):
        """Create SQLite database and tables"""
        db_path = self.settings['db_path']
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()
        
        # Contacts table based on Excel structure
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS contacts (
                nama TEXT NOT NULL,
                telepon TEXT,
                alamat TEXT,
                devices TEXT,
                merek TEXT,
                model TEXT,
                masalah TEXT,
                deskripsi TEXT,
                status TEXT,
                teknisi TEXT,
                harga REAL,
                tanggal_service DATE,
                garansi TEXT,
                berakhir_garansi DATE,
                google_map TEXT,
                sparepart TEXT,
                supplier TEXT,
                invoice_purchasing TEXT,
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Suppliers table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS suppliers (
                name TEXT NOT NULL,
                contact_person TEXT,
                phone TEXT,
                email TEXT,
                address TEXT,
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Warranty reminders table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS warranty_reminders (
                customer_name TEXT NOT NULL,
                phone TEXT,
                reminder_sent_date DATE,
                next_reminder_date DATE,
                status TEXT DEFAULT 'PENDING'
            )
        ''')
        
        # Spareparts inventory table
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS spareparts (
                part_name TEXT NOT NULL,
                part_type TEXT,
                brand TEXT,
                quantity INTEGER DEFAULT 0,
                unit_price REAL DEFAULT 0,
                supplier TEXT,
                min_stock INTEGER DEFAULT 5,
                location TEXT,
                notes TEXT,
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_time TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                username TEXT NOT NULL DEFAULT 'SYSTEM',
                action TEXT NOT NULL,
                entity_type TEXT,
                entity_id TEXT,
                details TEXT
            )
        ''')

        # Attribute automatic data-change audit entries to the local desktop session.
        self.conn.create_function("current_app_user", 0, self.current_username)

        # Indexes (speed up search/filter)
        try:
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_nama ON contacts(nama)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_telepon ON contacts(telepon)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_status ON contacts(status)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_tanggal_service ON contacts(tanggal_service)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_berakhir_garansi ON contacts(berakhir_garansi)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_suppliers_name ON suppliers(name)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_spareparts_part_name ON spareparts(part_name)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_spareparts_type ON spareparts(part_type)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_spareparts_brand ON spareparts(brand)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_event_time ON audit_log(event_time)")
            self.cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_username ON audit_log(username)")
        except Exception:
            pass


        # Automatic data-change audit trail for the three operational tables.
        audit_triggers = {
            "audit_contacts_insert": """
                CREATE TRIGGER IF NOT EXISTS audit_contacts_insert AFTER INSERT ON contacts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'CREATE', 'CONTACT', NEW.rowid, NEW.nama);
                END
            """,
            "audit_contacts_update": """
                CREATE TRIGGER IF NOT EXISTS audit_contacts_update AFTER UPDATE ON contacts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'UPDATE', 'CONTACT', NEW.rowid, NEW.nama);
                END
            """,
            "audit_contacts_delete": """
                CREATE TRIGGER IF NOT EXISTS audit_contacts_delete AFTER DELETE ON contacts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'DELETE', 'CONTACT', OLD.rowid, OLD.nama);
                END
            """,
            "audit_suppliers_insert": """
                CREATE TRIGGER IF NOT EXISTS audit_suppliers_insert AFTER INSERT ON suppliers
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'CREATE', 'SUPPLIER', NEW.rowid, NEW.name);
                END
            """,
            "audit_suppliers_update": """
                CREATE TRIGGER IF NOT EXISTS audit_suppliers_update AFTER UPDATE ON suppliers
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'UPDATE', 'SUPPLIER', NEW.rowid, NEW.name);
                END
            """,
            "audit_suppliers_delete": """
                CREATE TRIGGER IF NOT EXISTS audit_suppliers_delete AFTER DELETE ON suppliers
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'DELETE', 'SUPPLIER', OLD.rowid, OLD.name);
                END
            """,
            "audit_parts_insert": """
                CREATE TRIGGER IF NOT EXISTS audit_parts_insert AFTER INSERT ON spareparts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'CREATE', 'PART', NEW.rowid, NEW.part_name);
                END
            """,
            "audit_parts_update": """
                CREATE TRIGGER IF NOT EXISTS audit_parts_update AFTER UPDATE ON spareparts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'UPDATE', 'PART', NEW.rowid, NEW.part_name);
                END
            """,
            "audit_parts_delete": """
                CREATE TRIGGER IF NOT EXISTS audit_parts_delete AFTER DELETE ON spareparts
                BEGIN
                    INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                    VALUES (current_app_user(), 'DELETE', 'PART', OLD.rowid, OLD.part_name);
                END
            """,
        }
        for trigger_sql in audit_triggers.values():
            self.cursor.execute(trigger_sql)


        self.recalculate_database_warranty_dates()
        self.conn.commit()
    
    # ------------------------------------------------------------------
    # Local access and operational audit
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Local access and operational audit
    # ------------------------------------------------------------------
    def current_username(self):
        return "local"

    def is_admin(self):
        # Login and role restrictions were removed. The local desktop user has
        # unrestricted access to all operational features.
        return True

    def require_admin(self):
        return True

    def audit_event(self, action, entity_type="", entity_id="", details="", actor=None):
        if self.cursor is None or self.conn is None:
            return
        username = str(actor or self.current_username()).strip().lower() or "local"
        try:
            self.cursor.execute(
                """
                INSERT INTO audit_log(username, action, entity_type, entity_id, details)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    username,
                    str(action or "").strip().upper(),
                    str(entity_type or "").strip().upper(),
                    str(entity_id or "").strip(),
                    str(details or "").strip()[:1000],
                ),
            )
            self.conn.commit()
        except sqlite3.Error:
            pass

    def apply_role_permissions(self):
        """Keep all operational controls enabled in local-access mode."""
        unrestricted_actions = (
            "import_action", "import_reference_action", "export_action",
            "export_suppliers_action", "export_spareparts_action",
            "backup_action", "restore_action", "db_path_action",
            "currency_action", "reload_dropdown_action",
            "create_dummy_action", "dummy_editor_action",
        )
        for name in unrestricted_actions:
            action = getattr(self, name, None)
            if action is not None:
                action.setEnabled(True)
                original_tooltip = action.property("_permission_original_tooltip")
                if original_tooltip:
                    action.setToolTip(self.t(str(original_tooltip)))

        unrestricted_buttons = (
            "edit_contact_btn", "delete_btn", "export_selected_btn",
            "export_report_btn", "add_supplier_btn", "edit_supplier_btn",
            "update_supplier_btn", "delete_supplier_btn",
            "send_reminder_btn", "edit_warranty_contact_btn",
            "add_sparepart_btn", "edit_sparepart_btn",
            "update_sparepart_btn", "delete_sparepart_btn",
        )
        for name in unrestricted_buttons:
            button = getattr(self, name, None)
            if button is not None:
                button.setEnabled(True)
                original_tooltip = button.property("_permission_original_tooltip")
                if original_tooltip:
                    button.setToolTip(self.t(str(original_tooltip)))

        tab_widget = getattr(self, "tab_widget", None)
        reports_tab = getattr(self, "reports_tab", None)
        if tab_widget is not None and reports_tab is not None:
            report_index = tab_widget.indexOf(reports_tab)
            if report_index >= 0:
                tab_widget.setTabEnabled(report_index, True)
                tab_widget.setTabToolTip(report_index, "")

        for combo_name in ("supplier_combo", "parts_combo"):
            combo = getattr(self, combo_name, None)
            if isinstance(combo, QComboBox):
                combo.setEditable(True)




    def set_database_path(self):
        """Dialog to set database and backup paths"""
        if not self.require_admin():
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("DATABASE PATH SETTINGS")
        dialog.setFixedSize(500, 440)
        
        layout = QVBoxLayout(dialog)
        
        # Database path
        db_layout = QHBoxLayout()
        db_label = QLabel("DATABASE PATH:")
        self.db_path_edit = QLineEdit(self.settings['db_path'])
        db_browse_btn = QPushButton("BROWSE")
        db_browse_btn.clicked.connect(lambda: self.browse_path(self.db_path_edit, is_file=True))
        
        db_layout.addWidget(db_label)
        db_layout.addWidget(self.db_path_edit)
        db_layout.addWidget(db_browse_btn)
        
        # Backup path
        backup_layout = QHBoxLayout()
        backup_label = QLabel("BACKUP FOLDER:")
        self.backup_path_edit = QLineEdit(self.settings['backup_dir'])
        backup_browse_btn = QPushButton("BROWSE")
        backup_browse_btn.clicked.connect(lambda: self.browse_path(self.backup_path_edit, is_file=False))
        
        backup_layout.addWidget(backup_label)
        backup_layout.addWidget(self.backup_path_edit)
        backup_layout.addWidget(backup_browse_btn)

        # Import/export folder
        transfer_layout = QHBoxLayout()
        transfer_label = QLabel("IMPORT/EXPORT FOLDER:")
        self.transfer_path_edit = QLineEdit(self.settings.get('import_export_dir', self.get_import_export_dir()))
        transfer_browse_btn = QPushButton("BROWSE")
        transfer_browse_btn.clicked.connect(lambda: self.browse_path(self.transfer_path_edit, is_file=False))

        transfer_layout.addWidget(transfer_label)
        transfer_layout.addWidget(self.transfer_path_edit)
        transfer_layout.addWidget(transfer_browse_btn)

        # Quick default (D:\\BACKUP)
        default_layout = QHBoxLayout()
        use_master_btn = QPushButton("USE D:\\BACKUP (DEFAULT)")
        use_master_btn.setMinimumHeight(26)
        use_master_btn.clicked.connect(self.set_master_default_paths)
        default_layout.addStretch()
        default_layout.addWidget(use_master_btn)
        default_layout.addStretch()

        # Buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton(self.t("SAVE"))
        save_btn.clicked.connect(lambda: self.save_path_settings(dialog))
        cancel_btn = QPushButton(self.t("CANCEL"))
        cancel_btn.clicked.connect(dialog.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(db_layout)
        layout.addLayout(backup_layout)
        layout.addLayout(transfer_layout)
        layout.addLayout(default_layout)

        # Options
        self.auto_backup_chk = QCheckBox("AUTO BACKUP ON EXIT")
        self.auto_backup_chk.setChecked(bool(self.settings.get('auto_backup_on_exit', True)))
        layout.addWidget(self.auto_backup_chk)
        layout.addStretch()
        layout.addLayout(button_layout)
        
        self.apply_language_setting_to_ui(dialog)
        dialog.exec()
    
    def set_master_default_paths(self):
        """Set default DB/backup location to D:\\BACKUP"""
        if not self.drive_exists_for_path(self.default_db_path):
            QMessageBox.warning(self, "DRIVE NOT FOUND", "DRIVE D: IS NOT AVAILABLE. PLEASE CHOOSE ANOTHER FOLDER.")
            return
        if hasattr(self, 'db_path_edit'):
            self.db_path_edit.setText(self.default_db_path)
        if hasattr(self, 'backup_path_edit'):
            self.backup_path_edit.setText(self.default_backup_dir)
        if hasattr(self, 'transfer_path_edit'):
            self.transfer_path_edit.setText(self.default_backup_dir)

    def browse_path(self, edit_widget, is_file=False):
        if is_file:
            path, _ = QFileDialog.getSaveFileName(self, "SELECT DATABASE FILE", 
                                                 edit_widget.text(), 
                                                 "Database Files (*.db);;All Files (*)")
        else:
            path = QFileDialog.getExistingDirectory(self, "SELECT FOLDER", 
                                                   edit_widget.text())
        if path:
            edit_widget.setText(path)
    
    def save_path_settings(self, dialog):
        db_path = self.db_path_edit.text().strip()
        backup_dir = self.backup_path_edit.text().strip()
        transfer_dir = self.transfer_path_edit.text().strip() if hasattr(self, 'transfer_path_edit') else self.get_import_export_dir()

        if not db_path:
            QMessageBox.warning(self, "INVALID PATH", "DATABASE PATH CANNOT BE EMPTY")
            return
        if not backup_dir:
            QMessageBox.warning(self, "INVALID PATH", "BACKUP FOLDER CANNOT BE EMPTY")
            return
        if not transfer_dir:
            QMessageBox.warning(self, "INVALID PATH", "IMPORT/EXPORT FOLDER CANNOT BE EMPTY")
            return
        if not self.drive_exists_for_path(db_path):
            QMessageBox.warning(self, "INVALID DRIVE", "THE DRIVE FOR DATABASE PATH DOES NOT EXIST.")
            return
        if not self.drive_exists_for_path(backup_dir):
            QMessageBox.warning(self, "INVALID DRIVE", "THE DRIVE FOR BACKUP FOLDER DOES NOT EXIST.")
            return
        if not self.drive_exists_for_path(transfer_dir):
            QMessageBox.warning(self, "INVALID DRIVE", "THE DRIVE FOR IMPORT/EXPORT FOLDER DOES NOT EXIST.")
            return

        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        os.makedirs(backup_dir, exist_ok=True)
        os.makedirs(transfer_dir, exist_ok=True)

        self.apply_database_path(db_path)
        self.settings['backup_dir'] = backup_dir
        self.settings['import_export_dir'] = transfer_dir
        self.settings['db_location_initialized'] = True
        if hasattr(self, 'auto_backup_chk'):
            self.settings['auto_backup_on_exit'] = bool(self.auto_backup_chk.isChecked())
        self.save_settings()
        
        # Close current connection and reconnect to new database
        if self.conn:
            self.conn.close()
        
        self.create_database()
        self.load_contacts()
        self.load_suppliers()
        self.load_suppliers_for_combo()
        self.load_warranty_data()
        self.load_spareparts()
        
        dialog.accept()
        QMessageBox.information(self, "SETTINGS SAVED", 
                               "DATABASE PATHS UPDATED SUCCESSFULLY!")
    
    def setup_contacts_tab(self):
        """Create a two-panel customer editor with a scrollable left form."""
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        # Left panel is hosted by a scroll area so all fields remain reachable
        # on smaller displays and when Windows text scaling is enabled.
        left_scroll = QScrollArea()
        left_scroll.setObjectName("inputPanelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        left_scroll.setMinimumWidth(405)
        left_scroll.setMaximumWidth(540)

        left_panel = QWidget()
        left_panel.setObjectName("inputPanel")
        left_panel.setMinimumWidth(370)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 6, 8, 6)
        left_layout.setSpacing(4)

        # Customer Info
        info_group = QGroupBox("CUSTOMER INFORMATION")
        info_layout = QFormLayout(info_group)
        info_layout.setContentsMargins(8, 8, 8, 6)
        info_layout.setHorizontalSpacing(8)
        info_layout.setVerticalSpacing(4)

        self.name_input = UpperCaseLineEdit()
        self.name_input.setPlaceholderText("CUSTOMER NAME")
        self.phone_input = UpperCaseLineEdit()
        self.phone_input.setPlaceholderText("PHONE NUMBER")
        self.address_input = UpperCaseTextEdit()
        self.address_input.setFixedHeight(44)
        self.address_input.setPlaceholderText("FULL ADDRESS")

        info_layout.addRow("NAME:", self.name_input)
        info_layout.addRow("PHONE:", self.phone_input)
        info_layout.addRow("ADDRESS:", self.address_input)

        # Device Info
        device_group = QGroupBox("DEVICE INFORMATION")
        device_layout = QFormLayout(device_group)
        device_layout.setContentsMargins(8, 8, 8, 6)
        device_layout.setHorizontalSpacing(8)
        device_layout.setVerticalSpacing(4)

        self.device_type = ReliableComboBox()
        self.device_type.addItems(self.combo_values("device_types"))
        self.device_type.setEditable(True)

        self.brand_combo = ReliableComboBox()
        self.brand_combo.addItems(self.combo_values("brands"))
        self.brand_combo.setEditable(True)

        self.model_input = ReliableComboBox()
        self.model_input.addItems(self.combo_values("models"))
        self.model_input.setEditable(True)
        self.model_input.lineEdit().setPlaceholderText("DEVICE MODEL")

        device_layout.addRow("DEVICES:", self.device_type)
        device_layout.addRow("BRAND:", self.brand_combo)
        device_layout.addRow("MODEL:", self.model_input)

        # Service Info
        service_group = QGroupBox("SERVICE INFORMATION")
        service_layout = QFormLayout(service_group)
        service_layout.setContentsMargins(8, 8, 8, 6)
        service_layout.setHorizontalSpacing(8)
        service_layout.setVerticalSpacing(4)

        self.issue_combo = ReliableComboBox()
        self.issue_combo.addItems(self.combo_values("issues"))
        self.issue_combo.setEditable(True)

        self.description_input = UpperCaseTextEdit()
        self.description_input.setFixedHeight(44)
        self.description_input.setPlaceholderText("SERVICE DESCRIPTION")

        self.status_combo = ReliableComboBox()
        self.status_combo.addItems(self.combo_values("statuses"))
        self.status_combo.setEditable(True)

        self.technician_input = ReliableComboBox()
        self.technician_input.addItems(self.combo_values("technicians"))
        self.technician_input.setEditable(True)
        self.technician_input.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.technician_input.setPlaceholderText("TECHNICIAN NAME")

        self.price_input = PriceSpinBox()
        self.price_input.setRange(0, 1000000000)
        self.price_input.setSingleStep(10000)

        service_layout.addRow("ISSUE:", self.issue_combo)
        service_layout.addRow("DESCRIPTION:", self.description_input)
        service_layout.addRow("STATUS:", self.status_combo)
        service_layout.addRow("TECHNICIAN:", self.technician_input)
        service_layout.addRow("PRICE:", self.price_input)

        # Warranty and Date
        date_warranty_group = QGroupBox("DATE & WARRANTY")
        date_warranty_layout = QFormLayout(date_warranty_group)
        date_warranty_layout.setContentsMargins(8, 8, 8, 6)
        date_warranty_layout.setHorizontalSpacing(8)
        date_warranty_layout.setVerticalSpacing(4)

        self.service_date = QDateEdit()
        self.service_date.setCalendarPopup(True)
        self.service_date.setDate(QDate.currentDate())
        self.service_date.setDisplayFormat("dd-MM-yyyy")
        self.service_date.dateChanged.connect(self.update_warranty_end_date)

        self.current_service_date_btn = QPushButton("CURRENT DATE")
        self.current_service_date_btn.setIcon(create_calendar_icon())
        self.current_service_date_btn.setIconSize(QSize(18, 18))
        self.current_service_date_btn.setObjectName("secondaryButton")
        self.current_service_date_btn.setToolTip("SET SERVICE DATE TO TODAY")
        self.current_service_date_btn.clicked.connect(self.set_service_date_to_current)

        service_date_row = QWidget()
        service_date_row_layout = QHBoxLayout(service_date_row)
        service_date_row_layout.setContentsMargins(0, 0, 0, 0)
        service_date_row_layout.setSpacing(5)
        service_date_row_layout.addWidget(self.service_date, 1)
        service_date_row_layout.addWidget(self.current_service_date_btn)

        self.warranty_combo = ReliableComboBox()
        self.warranty_combo.addItems(self.combo_values("warranty_periods"))
        self.warranty_combo.setEditable(True)
        self.warranty_combo.currentTextChanged.connect(self.update_warranty_end_date)

        self.warranty_end_date = QDateEdit()
        self.warranty_end_date.setProperty("hideDateButton", True)
        self.warranty_end_date.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.warranty_end_date.setCalendarPopup(False)
        self.warranty_end_date.setReadOnly(True)
        self.warranty_end_date.setDate(QDate.currentDate())
        self.warranty_end_date.setDisplayFormat("dd-MM-yyyy")
        self.warranty_end_date.setToolTip("CALCULATED AUTOMATICALLY: SERVICE DATE + WARRANTY")
        self.update_warranty_end_date()

        date_warranty_layout.addRow("SERVICE DATE:", service_date_row)
        date_warranty_layout.addRow("WARRANTY:", self.warranty_combo)
        date_warranty_layout.addRow("WARRANTY END:", self.warranty_end_date)

        # Parts and Supplier Info
        parts_group = QGroupBox("PARTS & SUPPLIER")
        parts_layout = QFormLayout(parts_group)
        parts_layout.setContentsMargins(8, 8, 8, 6)
        parts_layout.setHorizontalSpacing(8)
        parts_layout.setVerticalSpacing(4)

        self.parts_combo = ReliableComboBox()
        self.parts_combo.addItems(self.combo_values("parts"))
        self.parts_combo.setEditable(True)

        self.supplier_combo = ReliableComboBox()
        self.supplier_combo.addItem("", None)
        self.supplier_combo.setEditable(True)

        self.maps_input = QLineEdit()
        self.maps_input.setPlaceholderText("LINK GOOGLE MAPS")

        maps_row = QHBoxLayout()
        maps_row.setContentsMargins(0, 0, 0, 0)
        maps_row.setSpacing(5)
        maps_row.addWidget(self.maps_input, 1)
        maps_copy_btn = QPushButton("COPY")
        maps_copy_btn.setObjectName("secondaryButton")
        maps_copy_btn.setMinimumWidth(68)
        maps_copy_btn.clicked.connect(self.copy_current_maps)
        maps_row.addWidget(maps_copy_btn)
        maps_open_btn = QPushButton("OPEN")
        maps_open_btn.setObjectName("primaryButton")
        maps_open_btn.setMinimumWidth(68)
        maps_open_btn.clicked.connect(self.open_current_maps)
        maps_row.addWidget(maps_open_btn)
        maps_widget = QWidget()
        maps_widget.setLayout(maps_row)

        parts_layout.addRow("PART:", self.parts_combo)
        parts_layout.addRow("SUPPLIER:", self.supplier_combo)
        parts_layout.addRow("GOOGLE MAP:", maps_widget)

        # Invoice
        invoice_group = QGroupBox("INVOICE")
        invoice_layout = QFormLayout(invoice_group)
        invoice_layout.setContentsMargins(8, 8, 8, 6)
        invoice_layout.setVerticalSpacing(4)

        self.invoice_input = UpperCaseLineEdit()
        self.invoice_input.setPlaceholderText("INVOICE NUMBER")
        invoice_layout.addRow("PURCHASE INVOICE:", self.invoice_input)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.setSpacing(5)

        self.save_btn = QPushButton("💾 SAVE")
        self.save_btn.setObjectName("successButton")
        self.save_btn.clicked.connect(self.save_contact)

        self.clear_btn = QPushButton("🗑️ CLEAR")
        self.clear_btn.setObjectName("dangerButton")
        self.clear_btn.clicked.connect(self.clear_form)

        self.edit_contact_btn = QPushButton("✏️ EDIT")
        self.edit_contact_btn.setObjectName("neutralButton")
        self.edit_contact_btn.clicked.connect(self.edit_selected_contact)

        self.delete_btn = QPushButton("❌ DELETE")
        self.delete_btn.setObjectName("warningButton")
        self.delete_btn.clicked.connect(self.delete_contact)

        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.clear_btn)
        button_layout.addWidget(self.edit_contact_btn)
        button_layout.addWidget(self.delete_btn)

        left_layout.addWidget(info_group)
        left_layout.addWidget(device_group)
        left_layout.addWidget(service_group)
        left_layout.addWidget(date_warranty_group)
        left_layout.addWidget(parts_group)
        left_layout.addWidget(invoice_group)
        left_layout.addLayout(button_layout)
        left_layout.addStretch(1)

        # Right Panel - Contact List
        right_panel = QWidget()
        right_panel.setObjectName("tablePanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 8, 6)
        right_layout.setSpacing(8)

        search_group = QGroupBox("SEARCH")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(8, 8, 8, 6)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("SEARCH BY NAME, PHONE, DEVICE, OR ISSUE...")
        self.search_input.textChanged.connect(self.filter_contacts)

        self.search_combo = ReliableComboBox()
        self.add_combo_values(self.search_combo, self.combo_values("search_fields"), translate_display=True)
        self.search_combo.setEditable(False)

        search_layout.addWidget(self.search_input, 3)
        search_layout.addWidget(self.search_combo, 1)

        self.contact_table = QTableWidget()
        self.contact_table.setAlternatingRowColors(True)
        self.contact_table.setColumnCount(18)
        self.contact_table.setHorizontalHeaderLabels([
            "NAME", "PHONE", "ADDRESS", "DEVICES", "BRAND", "MODEL",
            "ISSUE", "DESCRIPTION", "STATUS", "TECHNICIAN", "PRICE",
            "SERVICE DATE", "WARRANTY", "WARRANTY END",
            "GOOGLE MAP", "PART", "SUPPLIER", "INVOICE"
        ])

        column_widths = [150, 100, 200, 100, 100, 100, 200, 200, 100,
                         100, 100, 100, 100, 120, 150, 150, 150, 150]
        for column, width in enumerate(column_widths):
            self.contact_table.setColumnWidth(column, width)

        self.contact_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.contact_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.contact_table.cellClicked.connect(self.load_selected_contact)
        self.contact_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.contact_table.customContextMenuRequested.connect(self.show_contact_context_menu)
        self.contact_table.cellDoubleClicked.connect(self.contact_table_double_click)

        table_buttons = QHBoxLayout()
        refresh_btn = QPushButton("🔄 REFRESH")
        refresh_btn.setObjectName("primaryButton")
        refresh_btn.clicked.connect(self.load_contacts)

        self.export_selected_btn = QPushButton("📤 EXPORT SELECTED")
        self.export_selected_btn.setObjectName("warningButton")
        self.export_selected_btn.clicked.connect(self.export_selected)

        table_buttons.addWidget(refresh_btn)
        table_buttons.addWidget(self.export_selected_btn)
        table_buttons.addStretch()

        right_layout.addWidget(search_group)
        right_layout.addWidget(self.contact_table, 1)
        right_layout.addLayout(table_buttons)

        left_scroll.setWidget(left_panel)
        splitter.addWidget(left_scroll)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([445, 955])

        contacts_layout = QHBoxLayout(self.contacts_tab)
        contacts_layout.setContentsMargins(0, 0, 0, 0)
        contacts_layout.addWidget(splitter)

    def set_service_date_to_current(self):
        """Set the service date to today and recalculate warranty expiry."""
        self.service_date.setDate(QDate.currentDate())
        self.update_warranty_end_date()

    @staticmethod
    def calculate_warranty_end_date(start_date, warranty_period):
        """Return warranty expiry as service date plus the warranty duration."""
        if not isinstance(start_date, QDate) or not start_date.isValid():
            return QDate.currentDate()

        period = re.sub(r"\s+", " ", str(warranty_period or "").strip().upper())
        if period in {"", "0", "NONE", "NO WARRANTY", "TANPA GARANSI", "TIDAK ADA GARANSI"}:
            return start_date

        patterns = (
            (r"(\d+)\s*(?:DAY|DAYS|HARI)\b", "days"),
            (r"(\d+)\s*(?:WEEK|WEEKS|MINGGU)\b", "weeks"),
            (r"(\d+)\s*(?:MONTH|MONTHS|BULAN)\b", "months"),
            (r"(\d+)\s*(?:YEAR|YEARS|TAHUN)\b", "years"),
        )
        for pattern, unit in patterns:
            match = re.search(pattern, period)
            if match:
                amount = int(match.group(1))
                if unit == "days":
                    return start_date.addDays(amount)
                if unit == "weeks":
                    return start_date.addDays(amount * 7)
                if unit == "months":
                    return start_date.addMonths(amount)
                return start_date.addYears(amount)

        if period.isdigit():
            return start_date.addDays(int(period))
        return start_date

    def calculate_warranty_end_date_text(self, service_date_value, warranty_period):
        """Calculate a database-ready expiry; no-warranty records use an empty date."""
        normalized_period = re.sub(
            r"\s+", " ", str(warranty_period or "").strip().upper()
        )
        if normalized_period in {
            "", "0", "NONE", "NO WARRANTY", "TANPA GARANSI", "TIDAK ADA GARANSI"
        }:
            return ""

        service_text = str(service_date_value or "").strip()
        service_date = QDate.fromString(service_text, "yyyy-MM-dd")
        if not service_date.isValid():
            service_date = QDate.fromString(self.parse_date(service_text), "yyyy-MM-dd")
        if not service_date.isValid():
            return ""
        return self.calculate_warranty_end_date(
            service_date, normalized_period
        ).toString("yyyy-MM-dd")

    def recalculate_database_warranty_dates(self):
        """Synchronize stored expiry dates, including clearing no-warranty dates."""
        if self.cursor is None:
            return 0
        updated = 0
        try:
            rows = self.cursor.execute(
                "SELECT rowid, tanggal_service, garansi, berakhir_garansi FROM contacts"
            ).fetchall()
            for row_id, service_date, warranty_period, stored_end_date in rows:
                calculated = self.calculate_warranty_end_date_text(
                    service_date, warranty_period
                )
                if calculated != (stored_end_date or ""):
                    self.cursor.execute(
                        "UPDATE contacts SET berakhir_garansi=? WHERE rowid=?",
                        (calculated, row_id),
                    )
                    updated += 1
        except sqlite3.Error:
            return 0
        return updated

    def update_warranty_end_date(self):
        """Always set warranty end to service date plus warranty duration."""
        warranty_period = self.combo_source_value(self.warranty_combo)
        self.warranty_end_date.setDate(
            self.calculate_warranty_end_date(self.service_date.date(), warranty_period)
        )

    def setup_suppliers_tab(self):
        """Create supplier input and supplier table as two resizable panels."""
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        left_panel = QWidget()
        left_panel.setObjectName("inputPanel")
        left_panel.setMinimumWidth(330)
        left_panel.setMaximumWidth(480)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 6, 8, 6)
        left_layout.setSpacing(8)

        form_group = QGroupBox("ADD SUPPLIER")
        form_layout = QFormLayout(form_group)
        form_layout.setContentsMargins(8, 10, 8, 8)
        form_layout.setHorizontalSpacing(8)
        form_layout.setVerticalSpacing(6)

        self.supplier_name = UpperCaseLineEdit()
        self.supplier_name.setPlaceholderText("SUPPLIER NAME")
        self.supplier_contact = UpperCaseLineEdit()
        self.supplier_contact.setPlaceholderText("CONTACT NAME")
        self.supplier_phone = UpperCaseLineEdit()
        self.supplier_phone.setPlaceholderText("PHONE NUMBER")
        self.supplier_email = QLineEdit()
        self.supplier_email.setPlaceholderText("EMAIL")
        self.supplier_address = UpperCaseTextEdit()
        self.supplier_address.setFixedHeight(64)
        self.supplier_address.setPlaceholderText("SUPPLIER ADDRESS")

        form_layout.addRow("NAME:", self.supplier_name)
        form_layout.addRow("CONTACT:", self.supplier_contact)
        form_layout.addRow("PHONE:", self.supplier_phone)
        form_layout.addRow("EMAIL:", self.supplier_email)
        form_layout.addRow("ADDRESS:", self.supplier_address)

        supplier_buttons = QGridLayout()
        supplier_buttons.setHorizontalSpacing(6)
        supplier_buttons.setVerticalSpacing(6)

        self.add_supplier_btn = QPushButton("➕ ADD")
        self.add_supplier_btn.setObjectName("successButton")
        self.add_supplier_btn.clicked.connect(self.add_supplier)

        self.edit_supplier_btn = QPushButton("✏️ EDIT")
        self.edit_supplier_btn.setObjectName("neutralButton")
        self.edit_supplier_btn.clicked.connect(self.edit_selected_supplier)

        self.update_supplier_btn = QPushButton("💾 SAVE")
        self.update_supplier_btn.setObjectName("primaryButton")
        self.update_supplier_btn.clicked.connect(self.update_supplier)

        self.delete_supplier_btn = QPushButton("❌ DELETE")
        self.delete_supplier_btn.setObjectName("warningButton")
        self.delete_supplier_btn.clicked.connect(self.delete_supplier)

        self.clear_supplier_btn = QPushButton("🗑️ CLEAR")
        self.clear_supplier_btn.setObjectName("dangerButton")
        self.clear_supplier_btn.clicked.connect(self.clear_supplier_form)

        self.refresh_supplier_btn = QPushButton("🔃 REFRESH")
        self.refresh_supplier_btn.setObjectName("secondaryButton")
        self.refresh_supplier_btn.clicked.connect(self.refresh_suppliers_tab)

        supplier_buttons.addWidget(self.add_supplier_btn, 0, 0)
        supplier_buttons.addWidget(self.edit_supplier_btn, 0, 1)
        supplier_buttons.addWidget(self.update_supplier_btn, 1, 0)
        supplier_buttons.addWidget(self.delete_supplier_btn, 1, 1)
        supplier_buttons.addWidget(self.clear_supplier_btn, 2, 0)
        supplier_buttons.addWidget(self.refresh_supplier_btn, 2, 1)

        left_layout.addWidget(form_group)
        left_layout.addLayout(supplier_buttons)
        left_layout.addStretch(1)

        right_panel = QWidget()
        right_panel.setObjectName("tablePanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 8, 6)

        self.suppliers_table = QTableWidget()
        self.suppliers_table.setAlternatingRowColors(True)
        self.suppliers_table.setColumnCount(5)
        self.suppliers_table.setHorizontalHeaderLabels([
            "NAME", "CONTACT", "PHONE", "EMAIL", "ADDRESS"
        ])
        self.suppliers_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.suppliers_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.suppliers_table.cellClicked.connect(self.load_selected_supplier)
        self.suppliers_table.horizontalHeader().setStretchLastSection(True)
        right_layout.addWidget(self.suppliers_table)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([390, 1010])

        suppliers_layout = QHBoxLayout(self.suppliers_tab)
        suppliers_layout.setContentsMargins(0, 0, 0, 0)
        suppliers_layout.addWidget(splitter)

    def create_date_picker_row(self, date_edit, tooltip="OPEN CALENDAR"):
        """Return a date editor row with a dedicated calendar icon button."""
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(5)

        calendar_button = QToolButton()
        calendar_button.setIcon(create_calendar_icon())
        calendar_button.setIconSize(QSize(18, 18))
        calendar_button.setObjectName("datePickerButton")
        calendar_button.setToolTip(tooltip)
        calendar_button.setAccessibleName(tooltip)
        calendar_button.setFixedWidth(32)
        calendar_button.clicked.connect(
            lambda _checked=False, edit=date_edit, button=calendar_button:
                self.show_date_picker_menu(edit, button)
        )

        row_layout.addWidget(date_edit, 1)
        row_layout.addWidget(calendar_button)
        return row, calendar_button

    def show_date_picker_menu(self, date_edit, calendar_button):
        """Open a themed calendar below the requested date icon button."""
        if not isinstance(date_edit, QDateEdit):
            return

        menu = QMenu(calendar_button)
        calendar = QCalendarWidget(menu)
        calendar.setGridVisible(True)
        calendar.setMinimumDate(date_edit.minimumDate())
        calendar.setMaximumDate(date_edit.maximumDate())
        calendar.setSelectedDate(date_edit.date())

        calendar_action = QWidgetAction(menu)
        calendar_action.setDefaultWidget(calendar)
        menu.addAction(calendar_action)

        def select_date(selected_date):
            date_edit.setDate(selected_date)
            menu.close()
            date_edit.setFocus()

        calendar.clicked.connect(select_date)
        menu.exec(calendar_button.mapToGlobal(QPoint(0, calendar_button.height())))

    def setup_reports_tab(self):
        """Create report options and results as two resizable panels."""
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        left_panel = QWidget()
        left_panel.setObjectName("inputPanel")
        left_panel.setMinimumWidth(320)
        left_panel.setMaximumWidth(460)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 6, 8, 6)

        controls_group = QGroupBox("REPORT OPTIONS")
        controls_layout = QVBoxLayout(controls_group)
        controls_layout.setContentsMargins(8, 10, 8, 8)
        controls_layout.setSpacing(8)

        date_form = QFormLayout()
        date_form.setHorizontalSpacing(8)
        date_form.setVerticalSpacing(6)

        self.from_date = QDateEdit()
        self.from_date.setProperty("hideDateButton", True)
        self.from_date.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.from_date.setCalendarPopup(False)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.from_date.setDisplayFormat("dd-MM-yyyy")

        self.to_date = QDateEdit()
        self.to_date.setProperty("hideDateButton", True)
        self.to_date.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.to_date.setCalendarPopup(False)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.setDisplayFormat("dd-MM-yyyy")

        self.from_date_row, self.from_date_icon_btn = self.create_date_picker_row(
            self.from_date, "SELECT REPORT START DATE"
        )
        self.to_date_row, self.to_date_icon_btn = self.create_date_picker_row(
            self.to_date, "SELECT REPORT END DATE"
        )

        self.report_type = ReliableComboBox()
        self.add_combo_values(self.report_type, self.combo_values("report_types"), translate_display=True)
        self.report_type.setEditable(False)

        date_form.addRow("FROM:", self.from_date_row)
        date_form.addRow("TO:", self.to_date_row)
        date_form.addRow("REPORT TYPE:", self.report_type)

        button_layout = QVBoxLayout()
        self.generate_report_btn = QPushButton("📊 GENERATE REPORT")
        self.generate_report_btn.setObjectName("primaryButton")
        self.generate_report_btn.clicked.connect(self.generate_report)

        self.export_report_btn = QPushButton("📤 EXPORT TO EXCEL")
        self.export_report_btn.setObjectName("successButton")
        self.export_report_btn.clicked.connect(self.export_report)

        button_layout.addWidget(self.generate_report_btn)
        button_layout.addWidget(self.export_report_btn)

        controls_layout.addLayout(date_form)
        controls_layout.addLayout(button_layout)
        left_layout.addWidget(controls_group)
        left_layout.addStretch(1)

        right_panel = QWidget()
        right_panel.setObjectName("tablePanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 8, 6)

        self.report_table = QTableWidget()
        self.report_table.setAlternatingRowColors(True)
        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        self.report_text.setFont(QFont("Courier New", 10))

        self.report_view_tabs = QTabWidget()
        self.report_view_tabs.addTab(self.report_table, "📋 TABLE VIEW")
        self.report_view_tabs.addTab(self.report_text, "📄 TEXT VIEW")
        right_layout.addWidget(self.report_view_tabs)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([370, 1030])

        reports_layout = QHBoxLayout(self.reports_tab)
        reports_layout.setContentsMargins(0, 0, 0, 0)
        reports_layout.addWidget(splitter)

    def setup_warranty_tab(self):
        """Create warranty filters and warranty table as two resizable panels."""
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        left_panel = QWidget()
        left_panel.setObjectName("inputPanel")
        left_panel.setMinimumWidth(330)
        left_panel.setMaximumWidth(470)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 6, 8, 6)

        filter_group = QGroupBox("WARRANTY FILTER")
        filter_layout = QVBoxLayout(filter_group)
        filter_layout.setContentsMargins(8, 10, 8, 8)
        filter_layout.setSpacing(8)

        date_form = QFormLayout()
        date_form.setHorizontalSpacing(8)
        date_form.setVerticalSpacing(6)

        self.warranty_from_date = QDateEdit()
        self.warranty_from_date.setProperty("hideDateButton", True)
        self.warranty_from_date.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.warranty_from_date.setCalendarPopup(False)
        self.warranty_from_date.setDate(QDate.currentDate())
        self.warranty_from_date.setDisplayFormat("dd-MM-yyyy")

        self.warranty_to_date = QDateEdit()
        self.warranty_to_date.setProperty("hideDateButton", True)
        self.warranty_to_date.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.warranty_to_date.setCalendarPopup(False)
        self.warranty_to_date.setDate(QDate.currentDate().addDays(30))
        self.warranty_to_date.setDisplayFormat("dd-MM-yyyy")

        self.warranty_from_date_row, self.warranty_from_date_icon_btn = self.create_date_picker_row(
            self.warranty_from_date, "SELECT EXPIRY START DATE"
        )
        self.warranty_to_date_row, self.warranty_to_date_icon_btn = self.create_date_picker_row(
            self.warranty_to_date, "SELECT EXPIRY END DATE"
        )

        self.warranty_filter_combo = ReliableComboBox()
        self.add_combo_values(self.warranty_filter_combo, self.combo_values("warranty_filters"), translate_display=True)
        self.warranty_filter_combo.setEditable(False)
        self.warranty_filter_combo.currentTextChanged.connect(self.load_warranty_data)

        date_form.addRow("EXPIRY DATE FROM:", self.warranty_from_date_row)
        date_form.addRow("TO:", self.warranty_to_date_row)
        date_form.addRow("STATUS:", self.warranty_filter_combo)

        self.current_warranty_date_btn = QPushButton("📅 CURRENT DATE")
        self.current_warranty_date_btn.setObjectName("secondaryButton")
        self.current_warranty_date_btn.setToolTip("RESET WARRANTY RANGE FROM TODAY")
        self.current_warranty_date_btn.clicked.connect(self.set_warranty_filter_to_current_date)

        self.refresh_warranty_btn = QPushButton("🔄 REFRESH")
        self.refresh_warranty_btn.setObjectName("primaryButton")
        self.refresh_warranty_btn.clicked.connect(self.load_warranty_data)

        self.send_reminder_btn = QPushButton("📝 LOG REMINDER")
        self.send_reminder_btn.setObjectName("warningButton")
        self.send_reminder_btn.clicked.connect(self.send_warranty_reminder)

        self.edit_warranty_contact_btn = QPushButton("✏️ EDIT CONTACT")
        self.edit_warranty_contact_btn.setObjectName("neutralButton")
        self.edit_warranty_contact_btn.clicked.connect(self.edit_contact_from_warranty)

        action_layout = QGridLayout()
        action_layout.setHorizontalSpacing(6)
        action_layout.setVerticalSpacing(6)
        action_layout.addWidget(self.current_warranty_date_btn, 0, 0, 1, 2)
        action_layout.addWidget(self.refresh_warranty_btn, 1, 0)
        action_layout.addWidget(self.send_reminder_btn, 1, 1)
        action_layout.addWidget(self.edit_warranty_contact_btn, 2, 0, 1, 2)

        filter_layout.addLayout(date_form)
        filter_layout.addLayout(action_layout)
        left_layout.addWidget(filter_group)
        left_layout.addStretch(1)

        right_panel = QWidget()
        right_panel.setObjectName("tablePanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 8, 6)

        self.warranty_table = QTableWidget()
        self.warranty_table.setAlternatingRowColors(True)
        self.warranty_table.setColumnCount(9)
        self.warranty_table.setHorizontalHeaderLabels([
            "CUSTOMER", "PHONE", "DEVICE", "SERVICE DATE",
            "WARRANTY PERIOD", "END DATE", "DAYS LEFT",
            "STATUS", "TECHNICIAN"
        ])
        self.warranty_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.warranty_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.warranty_table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)

        column_widths = [150, 120, 150, 100, 110, 100, 80, 120, 100]
        for column, width in enumerate(column_widths):
            self.warranty_table.setColumnWidth(column, width)
        self.warranty_table.horizontalHeader().setStretchLastSection(True)
        right_layout.addWidget(self.warranty_table)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([390, 1010])

        warranty_layout = QHBoxLayout(self.warranty_tab)
        warranty_layout.setContentsMargins(0, 0, 0, 0)
        warranty_layout.addWidget(splitter)

    def set_warranty_filter_to_current_date(self):
        """Reset the warranty filter to today through the next 30 days."""
        today = QDate.currentDate()
        self.warranty_from_date.setDate(today)
        self.warranty_to_date.setDate(today.addDays(30))
        self.load_warranty_data()

    def edit_contact_from_warranty(self):
        if not self.require_admin():
            return
        row = self.warranty_table.currentRow()
        row_id = self._table_row_id(self.warranty_table, row)
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A WARRANTY ROW TO EDIT")
            return

        name_item = self.warranty_table.item(row, 0)
        customer_name = name_item.text().strip().upper() if name_item else ""
        self.tab_widget.setCurrentWidget(self.contacts_tab)
        # Reload the unfiltered table so the exact row can always be located.
        self.load_contacts()
        for contact_row in range(self.contact_table.rowCount()):
            if self._table_row_id(self.contact_table, contact_row) == row_id:
                self.contact_table.selectRow(contact_row)
                self.load_selected_contact(contact_row, 0)
                self.status_bar.showMessage(
                    f"OPENED CONTACT FROM WARRANTY: {customer_name}"
                )
                return
        QMessageBox.warning(
            self, "NOT FOUND", f"CONTACT NOT FOUND IN TABLE: {customer_name}"
        )

    def setup_spareparts_tab(self):
        """Create spare-part input/filter controls and inventory table as two panels."""
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        left_panel = QWidget()
        left_panel.setObjectName("inputPanel")
        left_panel.setMinimumWidth(350)
        left_panel.setMaximumWidth(500)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 6, 8, 6)
        left_layout.setSpacing(8)

        form_group = QGroupBox("PART INFORMATION")
        form_layout = QFormLayout(form_group)
        form_layout.setContentsMargins(8, 10, 8, 8)
        form_layout.setHorizontalSpacing(8)
        form_layout.setVerticalSpacing(5)

        self.sparepart_name = UpperCaseLineEdit()
        self.sparepart_name.setPlaceholderText("PART NAME")

        self.sparepart_type = ReliableComboBox()
        self.sparepart_type.addItems(self.combo_values("sparepart_types"))
        self.sparepart_type.setEditable(True)

        self.sparepart_brand = ReliableComboBox()
        self.sparepart_brand.addItems(self.combo_values("brands"))
        self.sparepart_brand.setEditable(True)

        self.sparepart_quantity = QSpinBox()
        self.sparepart_quantity.setProperty("hideSpinButtons", True)
        self.sparepart_quantity.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.sparepart_quantity.setRange(0, 10000)
        self.sparepart_quantity.setValue(0)

        self.sparepart_price = PriceSpinBox()
        self.sparepart_price.setRange(0, 1000000000)
        self.sparepart_price.setSingleStep(10000)

        self.sparepart_supplier = ReliableComboBox()
        self.sparepart_supplier.addItem("", None)
        self.sparepart_supplier.setEditable(True)

        self.sparepart_min_stock = QSpinBox()
        self.sparepart_min_stock.setProperty("hideSpinButtons", True)
        self.sparepart_min_stock.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.sparepart_min_stock.setRange(1, 1000)
        self.sparepart_min_stock.setValue(5)

        self.sparepart_location = UpperCaseLineEdit()
        self.sparepart_location.setPlaceholderText("STORAGE LOCATION")

        self.sparepart_notes = UpperCaseTextEdit()
        self.sparepart_notes.setFixedHeight(54)
        self.sparepart_notes.setPlaceholderText("ADDITIONAL NOTES")

        form_layout.addRow("PART NAME:", self.sparepart_name)
        form_layout.addRow("TYPE:", self.sparepart_type)
        form_layout.addRow("BRAND:", self.sparepart_brand)
        form_layout.addRow("QUANTITY:", self.sparepart_quantity)
        form_layout.addRow("UNIT PRICE:", self.sparepart_price)
        form_layout.addRow("SUPPLIER:", self.sparepart_supplier)
        form_layout.addRow("MIN STOCK:", self.sparepart_min_stock)
        form_layout.addRow("LOCATION:", self.sparepart_location)
        form_layout.addRow("NOTES:", self.sparepart_notes)

        sparepart_buttons = QGridLayout()
        sparepart_buttons.setHorizontalSpacing(6)
        sparepart_buttons.setVerticalSpacing(6)

        self.edit_sparepart_btn = QPushButton("✏️ EDIT")
        self.edit_sparepart_btn.setObjectName("neutralButton")
        self.edit_sparepart_btn.clicked.connect(self.edit_selected_sparepart)

        self.add_sparepart_btn = QPushButton("➕ ADD PART")
        self.add_sparepart_btn.setObjectName("successButton")
        self.add_sparepart_btn.clicked.connect(self.add_sparepart)

        self.update_sparepart_btn = QPushButton("💾 SAVE")
        self.update_sparepart_btn.setObjectName("primaryButton")
        self.update_sparepart_btn.clicked.connect(self.update_sparepart)

        self.clear_sparepart_btn = QPushButton("🗑️ CLEAR")
        self.clear_sparepart_btn.setObjectName("dangerButton")
        self.clear_sparepart_btn.clicked.connect(self.clear_sparepart_form)

        self.delete_sparepart_btn = QPushButton("❌ DELETE")
        self.delete_sparepart_btn.setObjectName("warningButton")
        self.delete_sparepart_btn.clicked.connect(self.delete_sparepart)

        self.refresh_sparepart_btn = QPushButton("🔃 REFRESH")
        self.refresh_sparepart_btn.setObjectName("secondaryButton")
        self.refresh_sparepart_btn.clicked.connect(self.refresh_spareparts_tab)

        sparepart_buttons.addWidget(self.add_sparepart_btn, 0, 0)
        sparepart_buttons.addWidget(self.edit_sparepart_btn, 0, 1)
        sparepart_buttons.addWidget(self.update_sparepart_btn, 1, 0)
        sparepart_buttons.addWidget(self.clear_sparepart_btn, 1, 1)
        sparepart_buttons.addWidget(self.delete_sparepart_btn, 2, 0)
        sparepart_buttons.addWidget(self.refresh_sparepart_btn, 2, 1)

        filter_group = QGroupBox("FILTER PARTS")
        filter_layout = QFormLayout(filter_group)
        filter_layout.setContentsMargins(8, 10, 8, 8)
        filter_layout.setHorizontalSpacing(8)
        filter_layout.setVerticalSpacing(6)

        self.sparepart_filter_type = ReliableComboBox()
        self.add_combo_values(self.sparepart_filter_type, self.combo_values("sparepart_filter_types"), translate_display=True)
        self.sparepart_filter_type.setEditable(True)
        self.sparepart_filter_type.currentTextChanged.connect(self.load_spareparts)

        self.sparepart_filter_brand = ReliableComboBox()
        self.add_combo_values(self.sparepart_filter_brand, self.combo_values("sparepart_filter_brands"), translate_display=True)
        self.sparepart_filter_brand.setEditable(True)
        self.sparepart_filter_brand.currentTextChanged.connect(self.load_spareparts)

        filter_layout.addRow("TYPE:", self.sparepart_filter_type)
        filter_layout.addRow("BRAND:", self.sparepart_filter_brand)

        self.low_stock_label = QLabel()
        self.low_stock_label.setObjectName("alertLabel")
        self.low_stock_label.setWordWrap(True)

        left_layout.addWidget(form_group)
        left_layout.addLayout(sparepart_buttons)
        left_layout.addWidget(filter_group)
        left_layout.addWidget(self.low_stock_label)
        left_layout.addStretch(1)

        right_panel = QWidget()
        right_panel.setObjectName("tablePanel")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 6, 8, 6)

        self.spareparts_table = QTableWidget()
        self.spareparts_table.setAlternatingRowColors(True)
        self.spareparts_table.setColumnCount(9)
        self.spareparts_table.setHorizontalHeaderLabels([
            "NAME", "TYPE", "BRAND", "QUANTITY", "PRICE",
            "SUPPLIER", "MIN STOCK", "LOCATION", "NOTES"
        ])
        self.spareparts_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.spareparts_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.spareparts_table.cellClicked.connect(self.load_selected_sparepart)

        column_widths = [150, 100, 100, 80, 100, 120, 80, 120, 180]
        for column, width in enumerate(column_widths):
            self.spareparts_table.setColumnWidth(column, width)
        self.spareparts_table.horizontalHeader().setStretchLastSection(True)
        right_layout.addWidget(self.spareparts_table)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([420, 980])

        spareparts_layout = QHBoxLayout(self.spareparts_tab)
        spareparts_layout.setContentsMargins(0, 0, 0, 0)
        spareparts_layout.addWidget(splitter)

        self.load_spareparts()

    def refresh_suppliers_tab(self):
        """Refresh suppliers table and reload supplier combo boxes."""
        try:
            self.load_suppliers()
        except Exception:
            pass
        try:
            self.load_suppliers_for_combo()
        except Exception:
            pass
        try:
            self.status_bar.showMessage("SUPPLIERS REFRESHED")
        except Exception:
            pass
    
    def refresh_spareparts_tab(self):
        """Refresh spareparts table and reload supplier list used by spareparts."""
        try:
            self.load_suppliers_for_combo()
        except Exception:
            pass
        try:
            self.load_spareparts()
        except Exception:
            pass
        try:
            self.status_bar.showMessage("PARTS REFRESHED")
        except Exception:
            pass
    
    def backup_database(self, silent: bool = False):
        """Create a backup of the database"""
        if not silent and not self.require_admin():
            return False
        try:
            if self.conn:
                self.conn.commit()

            # Get backup directory
            backup_dir = self.settings['backup_dir']
            os.makedirs(backup_dir, exist_ok=True)
            
            # Create backup filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"User_Database_PC_Backup_{timestamp}.db"
            backup_path = os.path.join(backup_dir, backup_filename)
            
            # Copy database file
            shutil.copy2(self.settings['db_path'], backup_path)
            
            if not silent:
                QMessageBox.information(self, "BACKUP SUCCESSFUL", 
                                       f"DATABASE BACKED UP SUCCESSFULLY TO:\n{backup_path}")
            self.set_status(f"BACKED UP TO {backup_path}")
            
        except Exception as e:
            if not silent:
                QMessageBox.critical(self, "BACKUP FAILED", 
                                    f"FAILED TO BACKUP DATABASE:\n{str(e)}")
            else:
                print(f"Backup failed during shutdown: {e}")
            self.set_status("BACKUP FAILED")
    
    def restore_database(self):
        """Validate and restore a SQLite backup without leaving a broken connection."""
        if not self.require_admin():
            return
        reply = QMessageBox.question(
            self, "CONFIRM RESTORE",
            "WARNING: RESTORING WILL REPLACE ALL CURRENT DATA WITH BACKUP DATA.\n"
            "THIS ACTION CANNOT BE UNDONE.\n\n"
            "DO YOU WANT TO CONTINUE?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        backup_file, _ = QFileDialog.getOpenFileName(
            self,
            "SELECT BACKUP FILE",
            self.settings["backup_dir"],
            "Database Files (*.db);;All Files (*)",
        )
        if not backup_file:
            return

        db_path = self.settings["db_path"]
        temp_backup = db_path + ".tmp"
        try:
            # Reject corrupt or unrelated files before replacing the live database.
            with sqlite3.connect(backup_file) as check_conn:
                quick_check = check_conn.execute("PRAGMA quick_check").fetchone()
                if not quick_check or str(quick_check[0]).lower() != "ok":
                    raise sqlite3.DatabaseError("BACKUP DATABASE FAILED SQLITE QUICK_CHECK")
                tables = {
                    row[0]
                    for row in check_conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
                if "contacts" not in tables:
                    raise sqlite3.DatabaseError("SELECTED FILE IS NOT A VALID APPLICATION BACKUP")

            if self.conn:
                self.conn.close()
                self.conn = None
                self.cursor = None

            if os.path.exists(db_path):
                shutil.copy2(db_path, temp_backup)
            shutil.copy2(backup_file, db_path)

            # Recreate the connection and ensure all current tables/indexes exist.
            self.create_database()
            self.load_contacts()
            self.load_suppliers()
            self.load_suppliers_for_combo()
            self.load_warranty_data()
            self.load_spareparts()
            self.audit_event("DATABASE_RESTORED", details=os.path.basename(backup_file))

            if os.path.exists(temp_backup):
                os.remove(temp_backup)

            QMessageBox.information(
                self, "RESTORE SUCCESSFUL", "DATABASE RESTORED SUCCESSFULLY FROM BACKUP!"
            )
            self.status_bar.showMessage("DATABASE RESTORED FROM BACKUP")
        except (OSError, sqlite3.Error) as exc:
            try:
                if self.conn:
                    self.conn.close()
            except sqlite3.Error:
                pass
            self.conn = None
            self.cursor = None

            rollback_error = None
            if os.path.exists(temp_backup):
                try:
                    shutil.copy2(temp_backup, db_path)
                    os.remove(temp_backup)
                except OSError as rollback_exc:
                    rollback_error = rollback_exc

            try:
                self.create_database()
            except (OSError, sqlite3.Error) as reconnect_exc:
                rollback_error = rollback_error or reconnect_exc

            details = str(exc)
            if rollback_error:
                details += f"\n\nROLLBACK/RECONNECT ERROR: {rollback_error}"
            QMessageBox.critical(
                self, "RESTORE FAILED", f"FAILED TO RESTORE DATABASE:\n{details}"
            )
            self.status_bar.showMessage("RESTORE FAILED")
    
    @staticmethod
    def _normalize_import_header(value):
        """Normalize spreadsheet headers without losing non-Latin text."""
        if value is None:
            return ""
        text = str(value).strip().upper().replace("：", ":")
        text = re.sub(r"[\r\n\t]+", " ", text)
        text = re.sub(r"[_-]+", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip(" :")

    def _import_header_aliases(self, table_name):
        """Return localized and common header aliases for reference imports."""
        table_name = str(table_name or "").strip().lower()
        if table_name == "suppliers":
            fields = {
                "NAME": ("NAME", "SUPPLIER", "SUPPLIER NAME"),
                "CONTACT": ("CONTACT", "CONTACT PERSON"),
                "PHONE": ("PHONE", "PHONE NUMBER", "TELEPHONE"),
                "EMAIL": ("EMAIL", "E-MAIL"),
                "ADDRESS": ("ADDRESS", "SUPPLIER ADDRESS"),
            }
        elif table_name == "spareparts":
            fields = {
                "PART NAME": ("PART NAME", "PART", "SPAREPART", "SPARE PART", "ITEM NAME"),
                "TYPE": ("TYPE", "PART TYPE", "CATEGORY"),
                "BRAND": ("BRAND",),
                "QTY": ("QTY", "QUANTITY", "STOCK"),
                "UNIT PRICE": ("UNIT PRICE", "PRICE", "COST"),
                "SUPPLIER": ("SUPPLIER", "SUPPLIER NAME"),
                "MIN STOCK": ("MIN STOCK", "MINIMUM STOCK", "MIN QTY", "REORDER LEVEL"),
                "LOCATION": ("LOCATION", "STORAGE LOCATION"),
                "NOTES": ("NOTES", "NOTE", "DESCRIPTION"),
            }
        else:
            raise ValueError("UNKNOWN IMPORT TABLE")

        aliases = {}
        for canonical, source_aliases in fields.items():
            for source in source_aliases:
                aliases[self._normalize_import_header(source)] = canonical
                for language_code in LANGUAGE_OPTIONS:
                    translated = translate_ui_text(source, language_code)
                    aliases[self._normalize_import_header(translated)] = canonical

        # Columns emitted by the exporter but intentionally not imported.
        for ignored in ("ID", "ID (ROWID)", "ROWID", "CREATED DATE", "CREATED AT"):
            aliases[self._normalize_import_header(ignored)] = ""
            for language_code in LANGUAGE_OPTIONS:
                aliases[self._normalize_import_header(
                    translate_ui_text(ignored, language_code)
                )] = ""
        return aliases

    def _canonical_import_headers(self, raw_headers, table_name):
        aliases = self._import_header_aliases(table_name)
        canonical = []
        for header in raw_headers:
            normalized = self._normalize_import_header(header)
            canonical.append(aliases.get(normalized, ""))

        required = ["NAME"] if table_name == "suppliers" else ["PART NAME"]
        missing = [field for field in required if field not in canonical]
        if missing:
            raise ValueError(
                f"{self.t('MISSING REQUIRED COLUMNS')}: "
                + ", ".join(self.t(field) for field in missing)
            )
        return canonical

    def _read_reference_import_rows(self, file_path, table_name):
        """Read localized XLSX/CSV headers and return canonical row dictionaries."""
        lower_path = str(file_path).lower()
        if lower_path.endswith(".xlsx"):
            workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            try:
                sheet = workbook.active
                iterator = sheet.iter_rows(values_only=True)
                try:
                    raw_headers = next(iterator)
                except StopIteration:
                    raise ValueError("FILE IS EMPTY") from None
                headers = self._canonical_import_headers(raw_headers, table_name)
                rows = [tuple(row) for row in iterator]
            finally:
                workbook.close()
            return rows, headers, "EXCEL"

        if lower_path.endswith(".csv"):
            last_error = None
            for encoding in ("utf-8-sig", "cp1252"):
                try:
                    with open(file_path, "r", encoding=encoding, newline="") as handle:
                        all_rows = list(csv.reader(handle))
                    break
                except UnicodeDecodeError as exc:
                    last_error = exc
            else:
                raise last_error or ValueError("FAILED TO READ CSV")

            if not all_rows:
                raise ValueError("CSV FILE IS EMPTY")
            headers = self._canonical_import_headers(all_rows[0], table_name)
            return [tuple(row) for row in all_rows[1:]], headers, "CSV"

        if lower_path.endswith(".xls"):
            raise ValueError(
                "LEGACY .XLS FILES ARE NOT SUPPORTED. SAVE THE FILE AS .XLSX OR CSV FIRST."
            )
        raise ValueError("PLEASE SELECT AN XLSX OR CSV FILE")

    @staticmethod
    def _import_text(value, uppercase=False):
        if value is None:
            return ""
        if isinstance(value, datetime):
            text = value.isoformat(sep=" ")
        elif isinstance(value, date):
            text = value.isoformat()
        else:
            text = str(value).strip()
        return text.upper() if uppercase else text

    @staticmethod
    def _parse_import_number(value, integer=False, default=0):
        """Parse spreadsheet numbers with either comma or dot separators."""
        if value in (None, ""):
            return int(default) if integer else float(default)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            number = float(value)
            return int(round(number)) if integer else number

        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        cleaned = re.sub(r"[^0-9,.-]", "", text)
        if cleaned in {"", "-", ".", ","}:
            return int(default) if integer else float(default)

        sign = "-" if cleaned.startswith("-") else ""
        cleaned = cleaned.lstrip("-").replace("-", "")
        comma = cleaned.rfind(",")
        dot = cleaned.rfind(".")

        if comma >= 0 and dot >= 0:
            decimal_sep = "," if comma > dot else "."
            thousands_sep = "." if decimal_sep == "," else ","
            normalized = cleaned.replace(thousands_sep, "").replace(decimal_sep, ".")
        elif comma >= 0 or dot >= 0:
            sep = "," if comma >= 0 else "."
            parts = cleaned.split(sep)
            # One or more 3-digit groups normally indicate thousands grouping.
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) >= 1):
                normalized = "".join(parts)
            else:
                normalized = ".".join(parts)
        else:
            normalized = cleaned

        try:
            number = float(sign + normalized)
        except ValueError:
            number = float(default)
        return int(round(number)) if integer else number

    def _rows_as_import_dicts(self, rows, headers):
        for row in rows:
            if not any(value not in (None, "") and str(value).strip() for value in row):
                yield None
                continue
            data = {}
            for index, canonical in enumerate(headers):
                if not canonical:
                    continue
                data[canonical] = row[index] if index < len(row) else ""
            yield data

    def _combined_import_header_aliases(self):
        """Return canonical aliases for one-file supplier and part imports."""
        fields = {
            "RECORD TYPE": (
                "RECORD TYPE", "DATA TYPE", "ENTITY TYPE", "ENTITY", "RECORD",
            ),
            "NAME": (
                "NAME", "SUPPLIER NAME", "COMPANY NAME", "VENDOR NAME",
            ),
            "CONTACT": ("CONTACT", "CONTACT PERSON"),
            "PHONE": ("PHONE", "PHONE NUMBER", "TELEPHONE"),
            "EMAIL": ("EMAIL", "E-MAIL"),
            "ADDRESS": ("ADDRESS", "SUPPLIER ADDRESS"),
            "PART NAME": (
                "PART NAME", "PART", "SPAREPART", "SPARE PART", "ITEM NAME",
            ),
            "TYPE": ("TYPE", "PART TYPE", "CATEGORY"),
            "BRAND": ("BRAND",),
            "QTY": ("QTY", "QUANTITY", "STOCK"),
            "UNIT PRICE": ("UNIT PRICE", "PRICE", "COST"),
            "SUPPLIER": ("SUPPLIER", "PART SUPPLIER", "VENDOR"),
            "MIN STOCK": (
                "MIN STOCK", "MINIMUM STOCK", "MIN QTY", "REORDER LEVEL",
            ),
            "LOCATION": ("LOCATION", "STORAGE LOCATION"),
            "NOTES": ("NOTES", "NOTE", "DESCRIPTION"),
        }

        aliases = {}
        for canonical, source_aliases in fields.items():
            for source in source_aliases:
                aliases[self._normalize_import_header(source)] = canonical
                for language_code in LANGUAGE_OPTIONS:
                    translated = translate_ui_text(source, language_code)
                    aliases[self._normalize_import_header(translated)] = canonical

        for ignored in ("ID", "ID (ROWID)", "ROWID", "CREATED DATE", "CREATED AT"):
            aliases[self._normalize_import_header(ignored)] = ""
            for language_code in LANGUAGE_OPTIONS:
                aliases[self._normalize_import_header(
                    translate_ui_text(ignored, language_code)
                )] = ""
        return aliases

    def _canonical_combined_import_headers(self, raw_headers):
        aliases = self._combined_import_header_aliases()
        canonical = []
        for header in raw_headers:
            normalized = self._normalize_import_header(header)
            canonical.append(aliases.get(normalized, ""))

        required = ("RECORD TYPE", "NAME", "PART NAME")
        missing = [field for field in required if field not in canonical]
        if missing:
            raise ValueError(
                f"{self.t('MISSING REQUIRED COLUMNS')}: "
                + ", ".join(self.t(field) for field in missing)
            )
        return canonical

    def _read_combined_reference_import_rows(self, file_path):
        """Read a single XLSX/CSV containing supplier and sparepart rows."""
        lower_path = str(file_path).lower()
        if lower_path.endswith(".xlsx"):
            workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            try:
                sheet = workbook.active
                iterator = sheet.iter_rows(values_only=True)
                try:
                    raw_headers = next(iterator)
                except StopIteration:
                    raise ValueError("FILE IS EMPTY") from None
                headers = self._canonical_combined_import_headers(raw_headers)
                rows = [tuple(row) for row in iterator]
            finally:
                workbook.close()
            return rows, headers, "EXCEL"

        if lower_path.endswith(".csv"):
            last_error = None
            for encoding in ("utf-8-sig", "cp1252"):
                try:
                    with open(file_path, "r", encoding=encoding, newline="") as handle:
                        all_rows = list(csv.reader(handle))
                    break
                except UnicodeDecodeError as exc:
                    last_error = exc
            else:
                raise last_error or ValueError("FAILED TO READ CSV")

            if not all_rows:
                raise ValueError("CSV FILE IS EMPTY")
            headers = self._canonical_combined_import_headers(all_rows[0])
            return [tuple(row) for row in all_rows[1:]], headers, "CSV"

        if lower_path.endswith(".xls"):
            raise ValueError(
                "LEGACY .XLS FILES ARE NOT SUPPORTED. SAVE THE FILE AS .XLSX OR CSV FIRST."
            )
        raise ValueError("PLEASE SELECT AN XLSX OR CSV FILE")

    def _combined_record_type(self, value):
        normalized = self._normalize_import_header(value)
        supplier_aliases = {
            "SUPPLIER", "SUPPLIERS", "VENDOR", "VENDORS", "SUPPLIER RECORD",
        }
        part_aliases = {
            "PART", "PARTS", "SPAREPART", "SPAREPARTS", "SPARE PART",
            "SPARE PARTS", "PART RECORD",
        }
        for language_code in LANGUAGE_OPTIONS:
            for source in ("SUPPLIER", "SUPPLIERS"):
                supplier_aliases.add(self._normalize_import_header(
                    translate_ui_text(source, language_code)
                ))
            for source in ("PART", "PARTS"):
                part_aliases.add(self._normalize_import_header(
                    translate_ui_text(source, language_code)
                ))
        if normalized in supplier_aliases:
            return "suppliers"
        if normalized in part_aliases:
            return "spareparts"
        return ""

    def import_suppliers_and_spareparts_data(self):
        """Import suppliers and spareparts together from one XLSX/CSV file."""
        if not self.require_admin():
            return

        file_path = self.get_import_file_path(
            "IMPORT SUPPLIERS & PARTS FROM EXCEL OR CSV"
        )
        if not file_path:
            return

        supplier_counts = {"added": 0, "updated": 0, "skipped": 0}
        part_counts = {"added": 0, "updated": 0, "skipped": 0}
        unknown_skipped = 0

        try:
            rows, headers, source_format = self._read_combined_reference_import_rows(
                file_path
            )
            supplier_rows = []
            part_rows = []
            for data in self._rows_as_import_dicts(rows, headers):
                if data is None:
                    unknown_skipped += 1
                    continue
                row_type = self._combined_record_type(data.get("RECORD TYPE"))
                if row_type == "suppliers":
                    supplier_rows.append(data)
                elif row_type == "spareparts":
                    part_rows.append(data)
                else:
                    unknown_skipped += 1

            with self.conn:
                # Suppliers are imported first so part supplier names already exist.
                for data in supplier_rows:
                    name = self._import_text(data.get("NAME"), uppercase=True)
                    if not name:
                        supplier_counts["skipped"] += 1
                        continue
                    values = (
                        self._import_text(data.get("CONTACT"), uppercase=True),
                        self._import_text(data.get("PHONE"), uppercase=True),
                        self._import_text(data.get("EMAIL"), uppercase=False),
                        self._import_text(data.get("ADDRESS"), uppercase=True),
                    )
                    existing = self.cursor.execute(
                        """
                        SELECT rowid, contact_person, phone, email, address
                        FROM suppliers
                        WHERE UPPER(TRIM(name))=? LIMIT 1
                        """,
                        (name,),
                    ).fetchone()
                    if existing:
                        current_values = existing[1:]
                        field_names = ("CONTACT", "PHONE", "EMAIL", "ADDRESS")
                        merged_values = tuple(
                            values[index] if field_name in data else current_values[index]
                            for index, field_name in enumerate(field_names)
                        )
                        self.cursor.execute(
                            """
                            UPDATE suppliers
                            SET contact_person=?, phone=?, email=?, address=?
                            WHERE rowid=?
                            """,
                            merged_values + (existing[0],),
                        )
                        supplier_counts["updated"] += 1
                    else:
                        self.cursor.execute(
                            """
                            INSERT INTO suppliers
                                (name, contact_person, phone, email, address)
                            VALUES (?, ?, ?, ?, ?)
                            """,
                            (name,) + values,
                        )
                        supplier_counts["added"] += 1

                for data in part_rows:
                    part_name = self._import_text(data.get("PART NAME"), uppercase=True)
                    if not part_name:
                        part_counts["skipped"] += 1
                        continue
                    values = (
                        self._import_text(data.get("TYPE"), uppercase=True),
                        self._import_text(data.get("BRAND"), uppercase=True),
                        self._parse_import_number(data.get("QTY"), integer=True, default=0),
                        self._parse_import_number(data.get("UNIT PRICE"), integer=False, default=0),
                        self._import_text(data.get("SUPPLIER"), uppercase=True),
                        self._parse_import_number(data.get("MIN STOCK"), integer=True, default=5),
                        self._import_text(data.get("LOCATION"), uppercase=True),
                        self._import_text(data.get("NOTES"), uppercase=True),
                    )
                    existing = self.cursor.execute(
                        """
                        SELECT rowid, part_type, brand, quantity, unit_price,
                               supplier, min_stock, location, notes
                        FROM spareparts
                        WHERE UPPER(TRIM(part_name))=? LIMIT 1
                        """,
                        (part_name,),
                    ).fetchone()
                    if existing:
                        current_values = existing[1:]
                        field_names = (
                            "TYPE", "BRAND", "QTY", "UNIT PRICE", "SUPPLIER",
                            "MIN STOCK", "LOCATION", "NOTES",
                        )
                        merged_values = tuple(
                            values[index] if field_name in data else current_values[index]
                            for index, field_name in enumerate(field_names)
                        )
                        self.cursor.execute(
                            """
                            UPDATE spareparts SET
                                part_type=?, brand=?, quantity=?, unit_price=?, supplier=?,
                                min_stock=?, location=?, notes=?
                            WHERE rowid=?
                            """,
                            merged_values + (existing[0],),
                        )
                        part_counts["updated"] += 1
                    else:
                        self.cursor.execute(
                            """
                            INSERT INTO spareparts
                                (part_name, part_type, brand, quantity, unit_price,
                                 supplier, min_stock, location, notes)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (part_name,) + values,
                        )
                        part_counts["added"] += 1

            self.load_suppliers()
            self.load_suppliers_for_combo()
            self.load_spareparts()
            self.refresh_dropdown_combos()

            total_skipped = (
                supplier_counts["skipped"]
                + part_counts["skipped"]
                + unknown_skipped
            )
            self.audit_event(
                "COMBINED_REFERENCE_DATA_IMPORTED",
                entity_type="SUPPLIERS_AND_PARTS",
                details=(
                    f"FORMAT={source_format}; "
                    f"SUPPLIERS_ADDED={supplier_counts['added']}; "
                    f"SUPPLIERS_UPDATED={supplier_counts['updated']}; "
                    f"PARTS_ADDED={part_counts['added']}; "
                    f"PARTS_UPDATED={part_counts['updated']}; "
                    f"SKIPPED={total_skipped}; "
                    f"FILE={os.path.basename(file_path)}"
                ),
            )

            summary = "\n".join([
                f"{self.t('SUPPLIERS')} ({source_format})",
                f"  {self.t('ADDED')}: {self.format_number(supplier_counts['added'])}",
                f"  {self.t('UPDATED')}: {self.format_number(supplier_counts['updated'])}",
                f"{self.t('PARTS')} ({source_format})",
                f"  {self.t('ADDED')}: {self.format_number(part_counts['added'])}",
                f"  {self.t('UPDATED')}: {self.format_number(part_counts['updated'])}",
                f"{self.t('SKIPPED')}: {self.format_number(total_skipped)}",
            ])
            QMessageBox.information(self, self.t("IMPORT COMPLETE"), summary)
            self.status_bar.showMessage(
                f"{self.t('IMPORT COMPLETE')}: "
                f"{self.t('SUPPLIERS')} {self.format_number(supplier_counts['added'])}, "
                f"{self.t('PARTS')} {self.format_number(part_counts['added'])}, "
                f"{self.t('SKIPPED')} {self.format_number(total_skipped)}"
            )
        except (
            OSError,
            csv.Error,
            openpyxl.utils.exceptions.InvalidFileException,
            sqlite3.Error,
            ValueError,
        ) as exc:
            try:
                self.conn.rollback()
            except Exception:
                pass
            QMessageBox.critical(
                self,
                self.t("IMPORT FAILED"),
                f"{self.t('FAILED TO IMPORT DATA:')}\n{exc}",
            )
            self.status_bar.showMessage(self.t("IMPORT FAILED"))

    def import_suppliers_data(self):
        """Import or update supplier records from an XLSX/CSV file."""
        return self.import_reference_data("suppliers")

    def import_spareparts_data(self):
        """Import or update sparepart records from an XLSX/CSV file."""
        return self.import_reference_data("spareparts")

    def import_reference_data(self, table_name):
        """Import suppliers or spareparts, updating duplicate names safely."""
        if not self.require_admin():
            return

        table_name = str(table_name or "").strip().lower()
        if table_name not in {"suppliers", "spareparts"}:
            self.ui_error("IMPORT FAILED", "UNKNOWN IMPORT TABLE")
            return

        entity_key = "SUPPLIERS" if table_name == "suppliers" else "PARTS"
        file_path = self.get_import_file_path(
            f"IMPORT {entity_key} FROM EXCEL OR CSV"
        )
        if not file_path:
            return

        try:
            rows, headers, source_format = self._read_reference_import_rows(
                file_path, table_name
            )
            added = 0
            updated = 0
            skipped = 0

            with self.conn:
                for data in self._rows_as_import_dicts(rows, headers):
                    if data is None:
                        skipped += 1
                        continue

                    if table_name == "suppliers":
                        name = self._import_text(data.get("NAME"), uppercase=True)
                        if not name:
                            skipped += 1
                            continue
                        values = (
                            self._import_text(data.get("CONTACT"), uppercase=True),
                            self._import_text(data.get("PHONE"), uppercase=True),
                            self._import_text(data.get("EMAIL"), uppercase=False),
                            self._import_text(data.get("ADDRESS"), uppercase=True),
                        )
                        existing = self.cursor.execute(
                            """
                            SELECT rowid, contact_person, phone, email, address
                            FROM suppliers
                            WHERE UPPER(TRIM(name))=? LIMIT 1
                            """,
                            (name,),
                        ).fetchone()
                        if existing:
                            # Preserve existing values when an optional column is
                            # absent from the import file. An explicitly blank cell
                            # still clears that field, which matches spreadsheet intent.
                            current_values = existing[1:]
                            field_names = ("CONTACT", "PHONE", "EMAIL", "ADDRESS")
                            merged_values = tuple(
                                values[index] if field_name in data else current_values[index]
                                for index, field_name in enumerate(field_names)
                            )
                            self.cursor.execute(
                                """
                                UPDATE suppliers
                                SET contact_person=?, phone=?, email=?, address=?
                                WHERE rowid=?
                                """,
                                merged_values + (existing[0],),
                            )
                            updated += 1
                        else:
                            self.cursor.execute(
                                """
                                INSERT INTO suppliers
                                    (name, contact_person, phone, email, address)
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (name,) + values,
                            )
                            added += 1
                    else:
                        part_name = self._import_text(data.get("PART NAME"), uppercase=True)
                        if not part_name:
                            skipped += 1
                            continue
                        values = (
                            self._import_text(data.get("TYPE"), uppercase=True),
                            self._import_text(data.get("BRAND"), uppercase=True),
                            self._parse_import_number(data.get("QTY"), integer=True, default=0),
                            self._parse_import_number(data.get("UNIT PRICE"), integer=False, default=0),
                            self._import_text(data.get("SUPPLIER"), uppercase=True),
                            self._parse_import_number(data.get("MIN STOCK"), integer=True, default=5),
                            self._import_text(data.get("LOCATION"), uppercase=True),
                            self._import_text(data.get("NOTES"), uppercase=True),
                        )
                        existing = self.cursor.execute(
                            """
                            SELECT rowid, part_type, brand, quantity, unit_price,
                                   supplier, min_stock, location, notes
                            FROM spareparts
                            WHERE UPPER(TRIM(part_name))=? LIMIT 1
                            """,
                            (part_name,),
                        ).fetchone()
                        if existing:
                            current_values = existing[1:]
                            field_names = (
                                "TYPE", "BRAND", "QTY", "UNIT PRICE", "SUPPLIER",
                                "MIN STOCK", "LOCATION", "NOTES",
                            )
                            merged_values = tuple(
                                values[index] if field_name in data else current_values[index]
                                for index, field_name in enumerate(field_names)
                            )
                            self.cursor.execute(
                                """
                                UPDATE spareparts SET
                                    part_type=?, brand=?, quantity=?, unit_price=?, supplier=?,
                                    min_stock=?, location=?, notes=?
                                WHERE rowid=?
                                """,
                                merged_values + (existing[0],),
                            )
                            updated += 1
                        else:
                            self.cursor.execute(
                                """
                                INSERT INTO spareparts
                                    (part_name, part_type, brand, quantity, unit_price,
                                     supplier, min_stock, location, notes)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """,
                                (part_name,) + values,
                            )
                            added += 1

            if table_name == "suppliers":
                self.load_suppliers()
                self.load_suppliers_for_combo()
            else:
                self.load_spareparts()
                self.refresh_dropdown_combos()

            self.audit_event(
                "REFERENCE_DATA_IMPORTED",
                entity_type=table_name.upper(),
                details=(
                    f"FORMAT={source_format}; ADDED={added}; UPDATED={updated}; "
                    f"SKIPPED={skipped}; FILE={os.path.basename(file_path)}"
                ),
            )

            summary = "\n".join([
                f"{self.t(entity_key)}: {source_format}",
                f"{self.t('ADDED')}: {self.format_number(added)}",
                f"{self.t('UPDATED')}: {self.format_number(updated)}",
                f"{self.t('SKIPPED')}: {self.format_number(skipped)}",
            ])
            QMessageBox.information(self, self.t("IMPORT COMPLETE"), summary)
            self.status_bar.showMessage(
                f"{self.t('IMPORT COMPLETE')}: {self.t(entity_key)} - "
                f"{self.t('ADDED')} {self.format_number(added)}, "
                f"{self.t('UPDATED')} {self.format_number(updated)}, "
                f"{self.t('SKIPPED')} {self.format_number(skipped)}"
            )
        except (
            OSError,
            csv.Error,
            openpyxl.utils.exceptions.InvalidFileException,
            sqlite3.Error,
            ValueError,
        ) as exc:
            try:
                self.conn.rollback()
            except Exception:
                pass
            QMessageBox.critical(
                self,
                self.t("IMPORT FAILED"),
                f"{self.t('FAILED TO IMPORT DATA:')}\n{exc}",
            )
            self.status_bar.showMessage(self.t("IMPORT FAILED"))

    def import_data(self):
        """Import supported .xlsx or CSV data files."""
        if not self.require_admin():
            return
        file_path = self.get_import_file_path()
        if not file_path:
            return

        try:
            lower_path = file_path.lower()
            if lower_path.endswith(".xlsx"):
                self.import_excel(file_path)
            elif lower_path.endswith(".csv"):
                self.import_csv(file_path)
            elif lower_path.endswith(".xls"):
                QMessageBox.warning(
                    self,
                    "UNSUPPORTED FORMAT",
                    "LEGACY .XLS FILES ARE NOT SUPPORTED. SAVE THE FILE AS .XLSX OR CSV FIRST.",
                )
            else:
                QMessageBox.warning(
                    self, "UNSUPPORTED FORMAT", "PLEASE SELECT AN XLSX OR CSV FILE"
                )
        except (OSError, csv.Error, openpyxl.utils.exceptions.InvalidFileException, sqlite3.Error, ValueError) as exc:
            QMessageBox.critical(
                self, "IMPORT FAILED", f"FAILED TO IMPORT DATA:\n{exc}"
            )
            self.status_bar.showMessage("IMPORT FAILED")
    
    def import_excel(self, file_path):
        """Import customer rows from an .xlsx workbook."""
        workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        try:
            sheet = workbook.active
            headers = [
                str(cell.value).strip().upper() if cell.value is not None else ""
                for cell in sheet[1]
            ]

            imported_count = 0
            skipped_count = 0
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not any(value not in (None, "") for value in row):
                    skipped_count += 1
                    continue

                data_dict = {}
                for index, header in enumerate(headers):
                    value = row[index] if index < len(row) else None
                    if value is None:
                        data_dict[header] = ""
                        continue

                    if header == "SERVICE DATE" and isinstance(value, (date, datetime)):
                        cleaned = value.strftime("%Y-%m-%d")
                    else:
                        cleaned = str(value).strip()

                    # URLs can contain case-sensitive path/query components.
                    data_dict[header] = cleaned if header == "GOOGLE MAP" else cleaned.upper()

                nama = data_dict.get("NAME", "").strip()
                if not nama:
                    skipped_count += 1
                    continue

                service_date_value = self.parse_date(data_dict.get("SERVICE DATE", ""))
                warranty_value = data_dict.get("WARRANTY", "")
                warranty_end_value = self.calculate_warranty_end_date_text(
                    service_date_value, warranty_value
                )

                try:
                    self.cursor.execute(
                        """
                        INSERT INTO contacts
                        (nama, telepon, alamat, devices, merek, model,
                         masalah, deskripsi, status, teknisi, harga,
                         tanggal_service, garansi, berakhir_garansi,
                         google_map, sparepart, supplier, invoice_purchasing)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            nama,
                            data_dict.get("PHONE", ""),
                            data_dict.get("ADDRESS", ""),
                            data_dict.get("DEVICES", ""),
                            data_dict.get("BRAND", ""),
                            data_dict.get("MODEL", ""),
                            data_dict.get("ISSUE", ""),
                            data_dict.get("DESCRIPTION", ""),
                            data_dict.get("STATUS", ""),
                            data_dict.get("TECHNICIAN", ""),
                            self.parse_price(data_dict.get("PRICE", "0")),
                            service_date_value,
                            warranty_value,
                            warranty_end_value,
                            data_dict.get("GOOGLE MAP", ""),
                            data_dict.get("PART", ""),
                            data_dict.get("SUPPLIER", ""),
                            data_dict.get("PURCHASE INVOICE", ""),
                        ),
                    )
                    imported_count += 1
                except (sqlite3.Error, ValueError) as exc:
                    print(f"Error importing row {nama}: {exc}")
                    skipped_count += 1

            self.conn.commit()
        finally:
            workbook.close()

        self.load_contacts()
        self.load_warranty_data()
        QMessageBox.information(
            self,
            "IMPORT COMPLETE",
            f"SUCCESSFULLY IMPORTED {imported_count} RECORDS FROM EXCEL\n"
            f"{skipped_count} ROWS SKIPPED",
        )
        self.status_bar.showMessage(f"IMPORTED {imported_count} RECORDS")
    
    def import_csv(self, file_path):
        """Import customer rows from a UTF-8 CSV file."""
        with open(file_path, "r", encoding="utf-8-sig", newline="") as file:
            csv_reader = csv.reader(file)
            try:
                headers = [header.strip().upper() for header in next(csv_reader)]
            except StopIteration:
                raise ValueError("CSV FILE IS EMPTY") from None

            imported_count = 0
            skipped_count = 0
            for row in csv_reader:
                if not any(value.strip() for value in row):
                    skipped_count += 1
                    continue

                data_dict = {}
                for index, header in enumerate(headers):
                    cleaned = row[index].strip() if index < len(row) else ""
                    data_dict[header] = cleaned if header == "GOOGLE MAP" else cleaned.upper()

                nama = data_dict.get("NAME", "").strip()
                if not nama:
                    skipped_count += 1
                    continue

                service_date_value = self.parse_date(data_dict.get("SERVICE DATE", ""))
                warranty_value = data_dict.get("WARRANTY", "")
                warranty_end_value = self.calculate_warranty_end_date_text(
                    service_date_value, warranty_value
                )

                try:
                    self.cursor.execute(
                        """
                        INSERT INTO contacts
                        (nama, telepon, alamat, devices, merek, model,
                         masalah, deskripsi, status, teknisi, harga,
                         tanggal_service, garansi, berakhir_garansi,
                         google_map, sparepart, supplier, invoice_purchasing)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            nama,
                            data_dict.get("PHONE", ""),
                            data_dict.get("ADDRESS", ""),
                            data_dict.get("DEVICES", ""),
                            data_dict.get("BRAND", ""),
                            data_dict.get("MODEL", ""),
                            data_dict.get("ISSUE", ""),
                            data_dict.get("DESCRIPTION", ""),
                            data_dict.get("STATUS", ""),
                            data_dict.get("TECHNICIAN", ""),
                            self.parse_price(data_dict.get("PRICE", "0")),
                            service_date_value,
                            warranty_value,
                            warranty_end_value,
                            data_dict.get("GOOGLE MAP", ""),
                            data_dict.get("PART", ""),
                            data_dict.get("SUPPLIER", ""),
                            data_dict.get("PURCHASE INVOICE", ""),
                        ),
                    )
                    imported_count += 1
                except (sqlite3.Error, ValueError) as exc:
                    print(f"Error importing row {nama}: {exc}")
                    skipped_count += 1

        self.conn.commit()
        self.load_contacts()
        self.load_warranty_data()
        QMessageBox.information(
            self,
            "IMPORT COMPLETE",
            f"SUCCESSFULLY IMPORTED {imported_count} RECORDS FROM CSV\n"
            f"{skipped_count} ROWS SKIPPED",
        )
        self.status_bar.showMessage(f"IMPORTED {imported_count} RECORDS")
    
    def export_data(self):
        """Export all data to a chosen Excel file"""
        if not self.require_admin():
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Data_User_Export_{timestamp}.xlsx"
        file_path = self.get_export_file_path(filename, "SAVE DATA EXPORT")
        if not file_path:
            return
        
        try:
            # Create new workbook
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "USER DATA"
            
            # Write headers based on Excel template
            headers = [
                "NAME", "PHONE", "ADDRESS", "DEVICES", "BRAND", "MODEL", 
                "ISSUE", "DESCRIPTION", "STATUS", "TECHNICIAN", "PRICE", 
                "SERVICE DATE", "WARRANTY", "WARRANTY END", 
                "GOOGLE MAP", "PART", "SUPPLIER", "PURCHASE INVOICE"
            ]
            
            for col, header in enumerate(headers, 1):
                sheet.cell(row=1, column=col, value=self.t(header))
                sheet.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
            
            # Get data from database
            self.cursor.execute('SELECT * FROM contacts ORDER BY created_date DESC')
            contacts = self.cursor.fetchall()
            
            # Write data
            for row_idx, contact in enumerate(contacts, 2):
                for col_idx, value in enumerate(contact[:18], 1):  # Skip created_date
                    cell = sheet.cell(row=row_idx, column=col_idx, value=value)
                    
                    # Format price column
                    if col_idx == 11 and value:  # PRICE column
                        if isinstance(value, (int, float)):
                            cell.number_format = '#,##0'
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except (TypeError, ValueError):
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Save workbook
            workbook.save(file_path)
            
            QMessageBox.information(self, "EXPORT SUCCESSFUL", 
                                   f"DATA EXPORTED SUCCESSFULLY TO:\n{file_path}")
            self.status_bar.showMessage(f"EXPORTED TO {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "EXPORT FAILED", 
                                f"FAILED TO EXPORT DATA:\n{str(e)}")
            self.status_bar.showMessage("EXPORT FAILED")
    
    def export_selected(self):
        """Export selected rows to Excel"""
        if not self.require_admin():
            return
        selected_rows = set()
        for item in self.contact_table.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            QMessageBox.warning(self, "NO SELECTION", "SELECT ROWS TO EXPORT")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Data_User_Selected_{timestamp}.xlsx"
        file_path = self.get_export_file_path(filename, "SAVE SELECTED DATA EXPORT")
        if not file_path:
            return
        
        try:
            # Create new workbook
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "SELECTED DATA"
            
            # Write headers
            headers = []
            for col in range(self.contact_table.columnCount()):
                headers.append(self.contact_table.horizontalHeaderItem(col).text())
            
            for col, header in enumerate(headers, 1):
                sheet.cell(row=1, column=col, value=header)
                sheet.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
            
            # Write selected data
            row_idx = 2
            for row in selected_rows:
                for col in range(self.contact_table.columnCount()):
                    item = self.contact_table.item(row, col)
                    value = item.text() if item else ""
                    sheet.cell(row=row_idx, column=col+1, value=value)
                row_idx += 1
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except (TypeError, ValueError):
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            workbook.save(file_path)
            
            QMessageBox.information(self, "EXPORT SUCCESSFUL", 
                                   f"SELECTED DATA EXPORTED TO:\n{file_path}")
            self.status_bar.showMessage(f"EXPORTED {len(selected_rows)} ROWS")
            
        except Exception as e:
            QMessageBox.critical(self, "EXPORT FAILED", 
                                f"FAILED TO EXPORT DATA:\n{str(e)}")
            self.status_bar.showMessage("EXPORT FAILED")
    
    def export_report(self):
        """Export current report to Excel"""
        if not self.require_admin():
            return
        if self.report_table.rowCount() == 0:
            QMessageBox.warning(self, "NO DATA", 
                               "GENERATE REPORT FIRST BEFORE EXPORTING")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_type = self.combo_source_value(self.report_type).replace(" ", "_")
        filename = f"Report_{report_type}_{timestamp}.xlsx"
        file_path = self.get_export_file_path(filename, "SAVE REPORT EXPORT")
        if not file_path:
            return
        
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = report_type[:31]  # Excel sheet name max 31 chars
            
            # Write headers
            headers = []
            for col in range(self.report_table.columnCount()):
                headers.append(self.report_table.horizontalHeaderItem(col).text())
            
            for col, header in enumerate(headers, 1):
                sheet.cell(row=1, column=col, value=header)
                sheet.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
            
            # Write data
            for row in range(self.report_table.rowCount()):
                for col in range(self.report_table.columnCount()):
                    item = self.report_table.item(row, col)
                    if item:
                        sheet.cell(row=row+2, column=col+1, value=item.text())
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except (TypeError, ValueError):
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            workbook.save(file_path)
            
            QMessageBox.information(self, "REPORT EXPORTED", 
                                   f"REPORT EXPORTED TO:\n{file_path}")
            self.status_bar.showMessage(f"REPORT EXPORTED TO {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "EXPORT FAILED", 
                                f"FAILED TO EXPORT REPORT:\n{str(e)}")
            self.status_bar.showMessage("REPORT EXPORT FAILED")
    
    def parse_price(self, price_str):
        """Parse price string to float"""
        return parse_currency_value(price_str, self.get_currency_code())
    
    def parse_date(self, date_value):
        """Parse common spreadsheet/local date values to YYYY-MM-DD."""
        if date_value in (None, ""):
            return ""
        if isinstance(date_value, datetime):
            return date_value.date().isoformat()
        if isinstance(date_value, date):
            return date_value.isoformat()

        date_text = str(date_value).strip()
        if not date_text:
            return ""
        preferred = get_language_config(
            self.settings.get("language_code", DEFAULT_LANGUAGE_CODE)
        ).get("python_date_format")
        formats = [
            preferred,
            "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S",
            "%Y/%m/%d",
            "%Y/%m/%d %H:%M:%S",
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%y",
            "%d/%m/%y",
            "%d.%m.%Y",
            "%d %b %Y",
            "%d %B %Y",
        ]
        for fmt in dict.fromkeys(fmt for fmt in formats if fmt):
            try:
                return datetime.strptime(date_text, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

        match = re.search(r"(?<!\d)(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?!\d)", date_text)
        if not match:
            return ""
        day, month, year = (int(part) for part in match.groups())
        if year < 100:
            year += 2000 if year < 50 else 1900
        try:
            return date(year, month, day).isoformat()
        except ValueError:
            return ""
    
    def load_suppliers_for_combo(self):
        """Load suppliers into combo boxes"""
        if not self.conn:
            return
            
        self.cursor.execute('SELECT DISTINCT name FROM suppliers ORDER BY name')
        suppliers = self.cursor.fetchall()
        
        # Clear and reload supplier combo in contacts tab
        self.supplier_combo.clear()
        self.supplier_combo.addItem("", None)
        
        # Also load for spareparts tab
        self.sparepart_supplier.clear()
        self.sparepart_supplier.addItem("", None)
        
        for supplier, in suppliers:
            self.supplier_combo.addItem(supplier, supplier)
            self.sparepart_supplier.addItem(supplier, supplier)
    
    def save_contact(self):
        if self.selected_contact_row_id is not None and not self.is_admin():
            self.require_admin()
            return
        nama = self.name_input.text().strip().upper()
        if not nama:
            QMessageBox.warning(self, "ERROR", "NAME IS REQUIRED")
            self.name_input.setFocus()
            return

        supplier = self.supplier_combo.currentData()
        if not supplier and self.supplier_combo.currentText():
            supplier = self.supplier_combo.currentText().strip().upper()
        sparepart = self.parts_combo.currentText().strip().upper()
        supplier = supplier or None
        sparepart = sparepart or None

        try:
            if supplier:
                exists = self.cursor.execute(
                    "SELECT 1 FROM suppliers WHERE UPPER(TRIM(name))=? LIMIT 1",
                    (supplier.upper(),),
                ).fetchone()
                if not exists:
                    if not self.is_admin():
                        QMessageBox.warning(
                            self,
                            "ADMINISTRATOR REQUIRED",
                            "A NEW SUPPLIER MUST BE CREATED BY AN ADMINISTRATOR.",
                        )
                        return
                    self.cursor.execute("INSERT INTO suppliers (name) VALUES (?)", (supplier,))
            if sparepart:
                exists = self.cursor.execute(
                    "SELECT 1 FROM spareparts WHERE UPPER(TRIM(part_name))=? LIMIT 1",
                    (sparepart.upper(),),
                ).fetchone()
                if not exists:
                    if not self.is_admin():
                        QMessageBox.warning(
                            self,
                            "ADMINISTRATOR REQUIRED",
                            "A NEW PART MUST BE CREATED BY AN ADMINISTRATOR.",
                        )
                        return
                    self.cursor.execute(
                        "INSERT INTO spareparts (part_name, supplier) VALUES (?, ?)",
                        (sparepart, supplier),
                    )
        except sqlite3.Error as exc:
            self.status_bar.showMessage(
                f"WARNING: FAILED TO REGISTER SUPPLIER/PART: {exc}"
            )

        warranty_value = self.combo_source_value(self.warranty_combo).strip().upper()
        service_date_str = self.service_date.date().toString("yyyy-MM-dd")
        warranty_end_date_str = self.calculate_warranty_end_date_text(
            service_date_str, warranty_value
        )
        values = (
            nama,
            self.phone_input.text().strip().upper(),
            self.address_input.toPlainText().strip().upper(),
            self.device_type.currentText(),
            self.brand_combo.currentText(),
            self.model_input.currentText().strip().upper(),
            self.issue_combo.currentText(),
            self.description_input.toPlainText().strip().upper(),
            self.status_combo.currentText(),
            self.technician_input.currentText().strip().upper(),
            self.price_input.value(),
            service_date_str,
            warranty_value,
            warranty_end_date_str,
            self.maps_input.text().strip(),
            sparepart,
            supplier,
            self.invoice_input.text().strip().upper(),
        )

        if self.selected_contact_row_id is not None:
            self.cursor.execute(
                """
                UPDATE contacts SET
                    nama=?, telepon=?, alamat=?, devices=?, merek=?, model=?,
                    masalah=?, deskripsi=?, status=?, teknisi=?, harga=?,
                    tanggal_service=?, garansi=?, berakhir_garansi=?,
                    google_map=?, sparepart=?, supplier=?, invoice_purchasing=?
                WHERE rowid=?
                """,
                values + (self.selected_contact_row_id,),
            )
            message = f"UPDATED {nama}"
        else:
            self.cursor.execute(
                """
                INSERT INTO contacts
                    (nama, telepon, alamat, devices, merek, model,
                     masalah, deskripsi, status, teknisi, harga,
                     tanggal_service, garansi, berakhir_garansi,
                     google_map, sparepart, supplier, invoice_purchasing)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                values,
            )
            message = f"SAVED {nama}"

        self.conn.commit()
        self.load_suppliers_for_combo()
        self.load_spareparts()
        self.load_contacts()
        self.load_warranty_data()
        self.clear_form()
        self.status_bar.showMessage(message)
    
    def load_selected_contact(self, row, column):
        if not self.is_admin():
            return
        row_id = self._table_row_id(self.contact_table, row)
        if row_id is None:
            return

        self.cursor.execute(
            """
            SELECT nama, telepon, alamat, devices, merek, model,
                   masalah, deskripsi, status, teknisi, harga,
                   tanggal_service, garansi, berakhir_garansi,
                   google_map, sparepart, supplier, invoice_purchasing
            FROM contacts WHERE rowid=?
            """,
            (row_id,),
        )
        contact = self.cursor.fetchone()
        if not contact:
            return

        self.selected_contact_row_id = row_id
        self.name_input.setText(contact[0] or "")
        self.phone_input.setText(contact[1] or "")
        self.address_input.setText(contact[2] or "")

        for combo, value in (
            (self.device_type, contact[3]),
            (self.brand_combo, contact[4]),
            (self.issue_combo, contact[6]),
            (self.status_combo, contact[8]),
        ):
            value = value or ""
            if value:
                index = combo.findText(value)
                if index >= 0:
                    combo.setCurrentIndex(index)
                else:
                    combo.addItem(value)
                    combo.setCurrentText(value)
            else:
                combo.setCurrentIndex(0)

        self.model_input.setCurrentText(contact[5] or "")
        self.description_input.setText(contact[7] or "")
        self.technician_input.setCurrentText(contact[9] or "")
        self.price_input.setValue(float(contact[10]) if contact[10] else 0)

        service_date = QDate.fromString(contact[11] or "", "yyyy-MM-dd")
        self.service_date.setDate(
            service_date if service_date.isValid() else QDate.currentDate()
        )

        warranty = contact[12] or ""
        if warranty:
            index = self.warranty_combo.findText(warranty)
            if index >= 0:
                self.warranty_combo.setCurrentIndex(index)
            else:
                self.warranty_combo.addItem(warranty)
                self.warranty_combo.setCurrentText(warranty)
        else:
            self.warranty_combo.setCurrentIndex(0)
        self.update_warranty_end_date()

        self.maps_input.setText(contact[14] or "")
        sparepart = contact[15] or ""
        if sparepart:
            index = self.parts_combo.findText(sparepart)
            if index >= 0:
                self.parts_combo.setCurrentIndex(index)
            else:
                self.parts_combo.addItem(sparepart)
                self.parts_combo.setCurrentText(sparepart)
        else:
            self.parts_combo.setCurrentIndex(0)

        supplier = contact[16] or None
        if supplier:
            index = self.supplier_combo.findData(supplier)
            if index >= 0:
                self.supplier_combo.setCurrentIndex(index)
            else:
                self.supplier_combo.addItem(supplier, supplier)
                self.supplier_combo.setCurrentText(supplier)
        else:
            self.supplier_combo.setCurrentIndex(0)
        self.invoice_input.setText(contact[17] or "")
    
    @staticmethod
    def _table_row_id(table, row):
        """Return the SQLite rowid stored on a table row, or None."""
        if row < 0:
            return None
        item = table.item(row, 0)
        if item is None:
            return None
        value = item.data(Qt.ItemDataRole.UserRole)
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    def _populate_contacts_table(self, contacts):
        """Populate the contacts grid once for both full and filtered queries."""
        self.contact_table.setRowCount(len(contacts))
        status_colors = {
            "NEW": QColor(255, 255, 200),
            "DIAGNOSING": QColor(173, 216, 230),
            "IN REPAIR": QColor(255, 218, 185),
            "WAITING FOR PARTS": QColor(255, 182, 193),
            "COMPLETED": QColor(144, 238, 144),
            "CANCELLED": QColor(220, 220, 220),
        }

        for row, contact in enumerate(contacts):
            row_id, *visible_values = contact
            for col, data in enumerate(visible_values):
                if col == 10:
                    value = self.format_currency(data)
                elif col in (11, 13):
                    value = self.format_date_for_display(data) if data else ""
                else:
                    value = str(data) if data is not None else ""

                item = QTableWidgetItem(value)
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_id)
                if col == 8 and data:
                    status = str(data).upper()
                    if status in status_colors:
                        item.setBackground(status_colors[status])
                        item.setForeground(QColor(0, 0, 0))
                self.contact_table.setItem(row, col, item)

    def load_contacts(self):
        if not self.conn:
            return
        self.cursor.execute(
            """
            SELECT rowid,
                   nama, telepon, alamat, devices, merek, model,
                   masalah, deskripsi, status, teknisi, harga,
                   tanggal_service, garansi, berakhir_garansi,
                   google_map, sparepart, supplier, invoice_purchasing
            FROM contacts
            ORDER BY created_date DESC, rowid DESC
            """
        )
        contacts = self.cursor.fetchall()
        self._populate_contacts_table(contacts)
        self.set_status(f"LOADED {len(contacts)} CONTACTS")

    def filter_contacts(self):
        search_text = self.search_input.text().strip().lower()
        search_field = self.combo_source_value(self.search_combo)
        if not search_text:
            self.load_contacts()
            return

        select_clause = """
            SELECT rowid,
                   nama, telepon, alamat, devices, merek, model,
                   masalah, deskripsi, status, teknisi, harga,
                   tanggal_service, garansi, berakhir_garansi,
                   google_map, sparepart, supplier, invoice_purchasing
            FROM contacts
        """
        if search_field == "ALL FIELDS":
            query = select_clause + """
                WHERE LOWER(COALESCE(nama, '')) LIKE ?
                   OR LOWER(COALESCE(telepon, '')) LIKE ?
                   OR LOWER(COALESCE(alamat, '')) LIKE ?
                   OR LOWER(COALESCE(devices, '')) LIKE ?
                   OR LOWER(COALESCE(merek, '')) LIKE ?
                   OR LOWER(COALESCE(model, '')) LIKE ?
                   OR LOWER(COALESCE(masalah, '')) LIKE ?
                   OR LOWER(COALESCE(deskripsi, '')) LIKE ?
                   OR LOWER(COALESCE(status, '')) LIKE ?
                   OR LOWER(COALESCE(teknisi, '')) LIKE ?
                   OR LOWER(COALESCE(sparepart, '')) LIKE ?
                   OR LOWER(COALESCE(supplier, '')) LIKE ?
                   OR LOWER(COALESCE(invoice_purchasing, '')) LIKE ?
                ORDER BY created_date DESC, rowid DESC
            """
            params = [f"%{search_text}%"] * 13
        else:
            field_map = {
                "NAME": "nama", "PHONE": "telepon", "DEVICES": "devices",
                "BRAND": "merek", "MODEL": "model", "ISSUE": "masalah",
                "STATUS": "status", "TECHNICIAN": "teknisi", "PART": "sparepart",
                "SUPPLIER": "supplier", "INVOICE": "invoice_purchasing",
            }
            field = field_map.get(search_field, "nama")
            query = (
                select_clause
                + f" WHERE LOWER(COALESCE({field}, '')) LIKE ?"
                + " ORDER BY created_date DESC, rowid DESC"
            )
            params = [f"%{search_text}%"]

        self.cursor.execute(query, params)
        self._populate_contacts_table(self.cursor.fetchall())
    
    def clear_form(self):
        self.selected_contact_row_id = None
        self.name_input.clear()
        self.phone_input.clear()
        self.address_input.clear()
        self.device_type.setCurrentIndex(0)
        self.brand_combo.setCurrentIndex(0)
        self.model_input.setCurrentText("")
        self.issue_combo.setCurrentIndex(0)
        self.description_input.clear()
        self.status_combo.setCurrentIndex(0)
        self.technician_input.setCurrentIndex(0)
        self.price_input.setValue(0)
        self.service_date.setDate(QDate.currentDate())
        self.warranty_combo.setCurrentIndex(0)
        self.update_warranty_end_date()
        self.maps_input.clear()
        self.parts_combo.setCurrentIndex(0)
        self.supplier_combo.setCurrentIndex(0)
        self.invoice_input.clear()
        self.contact_table.clearSelection()
        self.status_bar.showMessage("FORM CLEARED")
    
    def edit_selected_contact(self):
        if not self.require_admin():
            return
        selected_row = self.contact_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A CONTACT TO EDIT")
            return
        self.load_selected_contact(selected_row, 0)
        self.name_input.setFocus()
        self.status_bar.showMessage("EDIT MODE: UPDATE FIELDS THEN CLICK SAVE")

    def delete_contact(self):
        if not self.require_admin():
            return
        selected_row = self.contact_table.currentRow()
        row_id = self._table_row_id(self.contact_table, selected_row)
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A CONTACT TO DELETE")
            return

        name_item = self.contact_table.item(selected_row, 0)
        phone_item = self.contact_table.item(selected_row, 1)
        contact_name = name_item.text() if name_item else ""
        contact_phone = phone_item.text() if phone_item else ""
        reply = QMessageBox.question(
            self,
            "CONFIRM DELETE",
            f"DELETE '{contact_name}'?\n\nTHIS ACTION CANNOT BE UNDONE!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.cursor.execute(
            "DELETE FROM warranty_reminders WHERE customer_name=? AND COALESCE(phone, '')=?",
            (contact_name, contact_phone),
        )
        self.cursor.execute("DELETE FROM contacts WHERE rowid=?", (row_id,))
        self.conn.commit()
        self.status_bar.showMessage(f"DELETED {contact_name}")
        self.load_contacts()
        self.load_warranty_data()
        self.clear_form()
    
    def add_supplier(self):
        if not self.require_admin():
            return
        supplier_name = self.supplier_name.text().strip().upper()
        if not supplier_name:
            QMessageBox.warning(self, "ERROR", "SUPPLIER NAME IS REQUIRED")
            self.supplier_name.setFocus()
            return

        duplicate = self.cursor.execute(
            "SELECT rowid FROM suppliers WHERE UPPER(TRIM(name))=? LIMIT 1",
            (supplier_name,),
        ).fetchone()
        if duplicate:
            QMessageBox.warning(
                self,
                "DUPLICATE SUPPLIER",
                "A SUPPLIER WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.",
            )
            return

        self.cursor.execute(
            """
            INSERT INTO suppliers (name, contact_person, phone, email, address)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                supplier_name,
                self.supplier_contact.text().strip().upper(),
                self.supplier_phone.text().strip().upper(),
                self.supplier_email.text().strip(),
                self.supplier_address.toPlainText().strip().upper(),
            ),
        )
        self.conn.commit()
        self.load_suppliers()
        self.load_suppliers_for_combo()
        self.clear_supplier_form()
        self.status_bar.showMessage(f"ADDED {supplier_name}")
    
    def load_suppliers(self):
        if not self.conn:
            return
        self.cursor.execute(
            """
            SELECT rowid, name, contact_person, phone, email, address
            FROM suppliers ORDER BY name, rowid
            """
        )
        suppliers = self.cursor.fetchall()
        self.suppliers_table.setRowCount(len(suppliers))
        for row, supplier in enumerate(suppliers):
            row_id, *values = supplier
            for col, data in enumerate(values):
                item = QTableWidgetItem(str(data) if data is not None else "")
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_id)
                self.suppliers_table.setItem(row, col, item)
    
    def clear_supplier_form(self):
        self.supplier_name.clear()
        self.supplier_contact.clear()
        self.supplier_phone.clear()
        self.supplier_email.clear()
        self.supplier_address.clear()
        self.selected_supplier_row_id = None
        self.selected_supplier_original_name = None
        self.suppliers_table.clearSelection()
    

    def load_selected_supplier(self, row, column):
        """Load the exact selected supplier row into the form."""
        row_id = self._table_row_id(self.suppliers_table, row)
        if row_id is None:
            return
        self.selected_supplier_row_id = row_id
        name_item = self.suppliers_table.item(row, 0)
        self.selected_supplier_original_name = (
            name_item.text().strip().upper() if name_item else ""
        )
        self.supplier_name.setText(self.selected_supplier_original_name)
        self.supplier_contact.setText(
            self.suppliers_table.item(row, 1).text() if self.suppliers_table.item(row, 1) else ""
        )
        self.supplier_phone.setText(
            self.suppliers_table.item(row, 2).text() if self.suppliers_table.item(row, 2) else ""
        )
        self.supplier_email.setText(
            self.suppliers_table.item(row, 3).text() if self.suppliers_table.item(row, 3) else ""
        )
        self.supplier_address.setText(
            self.suppliers_table.item(row, 4).text() if self.suppliers_table.item(row, 4) else ""
        )
        self.status_bar.showMessage(
            f"SELECTED SUPPLIER: {self.selected_supplier_original_name}"
        )

    def edit_selected_supplier(self):
        if not self.require_admin():
            return
        row = self.suppliers_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A SUPPLIER TO EDIT")
            return
        self.load_selected_supplier(row, 0)
        self.supplier_name.setFocus()
        self.status_bar.showMessage("EDIT MODE: UPDATE FIELDS THEN CLICK UPDATE")

    def update_supplier(self):
        if not self.require_admin():
            return
        row_id = self.selected_supplier_row_id
        if row_id is None:
            row_id = self._table_row_id(self.suppliers_table, self.suppliers_table.currentRow())
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A SUPPLIER TO UPDATE")
            return

        new_name = self.supplier_name.text().strip().upper()
        if not new_name:
            QMessageBox.warning(self, "ERROR", "SUPPLIER NAME IS REQUIRED")
            self.supplier_name.setFocus()
            return
        current_name_row = self.cursor.execute(
            "SELECT name FROM suppliers WHERE rowid=?", (row_id,)
        ).fetchone()
        current_name = str(current_name_row[0] or "").strip().upper() if current_name_row else ""
        if new_name != current_name:
            duplicate = self.cursor.execute(
                """
                SELECT rowid FROM suppliers
                WHERE UPPER(TRIM(name))=? AND rowid<>? LIMIT 1
                """,
                (new_name, row_id),
            ).fetchone()
            if duplicate:
                QMessageBox.warning(
                    self, "DUPLICATE SUPPLIER", "ANOTHER SUPPLIER ALREADY USES THIS NAME."
                )
                return

        self.cursor.execute(
            """
            UPDATE suppliers SET name=?, contact_person=?, phone=?, email=?, address=?
            WHERE rowid=?
            """,
            (
                new_name,
                self.supplier_contact.text().strip().upper(),
                self.supplier_phone.text().strip().upper(),
                self.supplier_email.text().strip(),
                self.supplier_address.toPlainText().strip().upper(),
                row_id,
            ),
        )
        self.conn.commit()
        self.load_suppliers()
        self.load_suppliers_for_combo()
        self.clear_supplier_form()
        self.status_bar.showMessage(f"UPDATED SUPPLIER: {new_name}")

    def delete_supplier(self):
        if not self.require_admin():
            return
        row = self.suppliers_table.currentRow()
        row_id = self._table_row_id(self.suppliers_table, row)
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A SUPPLIER TO DELETE")
            return
        item = self.suppliers_table.item(row, 0)
        supplier_name = item.text().strip().upper() if item else ""
        reply = QMessageBox.question(
            self,
            "CONFIRM DELETE",
            f"DELETE SUPPLIER '{supplier_name}'?\n\nTHIS ACTION CANNOT BE UNDONE!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.cursor.execute("DELETE FROM suppliers WHERE rowid=?", (row_id,))
        self.conn.commit()
        self.load_suppliers()
        self.load_suppliers_for_combo()
        self.clear_supplier_form()
        self.status_bar.showMessage(f"DELETED SUPPLIER: {supplier_name}")

    def load_warranty_data(self):
        """Load warranty data using exact contact row IDs and validated dates."""
        if not self.conn:
            return
        if self.warranty_from_date.date() > self.warranty_to_date.date():
            self.warranty_table.setRowCount(0)
            self.status_bar.showMessage("INVALID WARRANTY DATE RANGE: FROM IS AFTER TO")
            return

        filter_type = self.combo_source_value(self.warranty_filter_combo)
        from_date = self.warranty_from_date.date().toString("yyyy-MM-dd")
        to_date = self.warranty_to_date.date().toString("yyyy-MM-dd")
        base_select = """
            SELECT rowid, nama, telepon, devices,
                   tanggal_service, garansi, berakhir_garansi, status, teknisi
            FROM contacts
        """
        params = []
        if filter_type == "NO WARRANTY":
            query = base_select + """
                WHERE COALESCE(TRIM(garansi), '') IN ('', 'NO WARRANTY')
                ORDER BY created_date DESC, rowid DESC
            """
        else:
            query = base_select + """
                WHERE COALESCE(TRIM(garansi), '') NOT IN ('', 'NO WARRANTY')
                  AND COALESCE(TRIM(berakhir_garansi), '') <> ''
                  AND berakhir_garansi BETWEEN ? AND ?
            """
            params.extend([from_date, to_date])
            if filter_type == "ACTIVE":
                query += " AND berakhir_garansi >= date('now')"
            elif filter_type == "EXPIRING SOON (≤ 7 DAYS)":
                query += " AND berakhir_garansi BETWEEN date('now') AND date('now', '+7 days')"
            elif filter_type == "EXPIRING SOON (≤ 30 DAYS)":
                query += " AND berakhir_garansi BETWEEN date('now') AND date('now', '+30 days')"
            elif filter_type == "EXPIRED":
                query += " AND berakhir_garansi < date('now')"
            query += " ORDER BY berakhir_garansi ASC, rowid ASC"

        self.cursor.execute(query, params)
        warranties = self.cursor.fetchall()
        self.warranty_table.setRowCount(len(warranties))
        today = datetime.now().date()

        for row, warranty in enumerate(warranties):
            row_id, nama, phone, device, service_date, warranty_period, end_date, status, technician = warranty
            days_left = "N/A"
            end_date_obj = None
            if end_date and warranty_period and warranty_period != "NO WARRANTY":
                try:
                    end_date_obj = datetime.strptime(end_date, "%Y-%m-%d").date()
                    days_left = str((end_date_obj - today).days)
                except (TypeError, ValueError):
                    end_date_obj = None

            warranty_status = "ACTIVE"
            status_color = QColor(144, 238, 144)
            if not warranty_period or warranty_period == "NO WARRANTY":
                warranty_status = "NO WARRANTY"
                status_color = QColor(211, 211, 211)
            elif end_date_obj:
                days_left_int = (end_date_obj - today).days
                if days_left_int < 0:
                    warranty_status = "EXPIRED"
                    status_color = QColor(255, 99, 71)
                elif days_left_int <= 7:
                    warranty_status = "EXPIRING SOON (≤ 7 DAYS)"
                    status_color = QColor(255, 165, 0)
                elif days_left_int <= 30:
                    warranty_status = "EXPIRING SOON (≤ 30 DAYS)"
                    status_color = QColor(255, 215, 0)

            data_items = [
                nama,
                phone,
                device,
                self.format_date_for_display(service_date) if service_date else "",
                self.t(warranty_period if warranty_period else "N/A"),
                self.format_date_for_display(end_date) if end_date else "",
                days_left,
                self.t(warranty_status),
                technician if technician else self.t("N/A"),
            ]
            for col, data in enumerate(data_items):
                item = QTableWidgetItem(str(data) if data not in (None, "") else "N/A")
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_id)
                if col == 7:
                    item.setBackground(status_color)
                    if warranty_status == "EXPIRED":
                        item.setForeground(QColor(255, 255, 255))
                    elif warranty_status == "NO WARRANTY":
                        item.setForeground(QColor(100, 100, 100))
                elif col == 6 and data != "N/A":
                    try:
                        days = int(data)
                        if days < 0:
                            item.setForeground(QColor(255, 0, 0))
                            item.setText(self.t(f"{days} (EXPIRED)"))
                        elif days <= 7:
                            item.setForeground(QColor(255, 140, 0))
                            item.setText(self.t(f"{days} (SOON)"))
                        elif days <= 30:
                            item.setForeground(QColor(255, 165, 0))
                            item.setText(self.t(f"{days} (WARNING)"))
                        else:
                            item.setForeground(QColor(0, 128, 0))
                    except (TypeError, ValueError):
                        pass
                self.warranty_table.setItem(row, col, item)
    
    def update_warranty_status(self):
        """Update warranty status colors in real-time"""
        self.load_warranty_data()
    
    def send_warranty_reminder(self):
        """Preview and log reminders; this application has no SMS/WhatsApp sender."""
        if not self.require_admin():
            return
        selected_rows = self.warranty_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "NO SELECTION", "SELECT WARRANTIES TO LOG REMINDERS")
            return

        customers_to_remind = []
        for index in selected_rows:
            row = index.row()
            customers_to_remind.append(
                {
                    "name": self.warranty_table.item(row, 0).text(),
                    "phone": self.warranty_table.item(row, 1).text(),
                    "status": self.warranty_table.item(row, 7).text(),
                    "end_date": self.warranty_table.item(row, 5).text(),
                }
            )

        reminder_dialog = QDialog(self)
        reminder_dialog.setWindowTitle("LOG WARRANTY REMINDER")
        reminder_dialog.setFixedSize(600, 400)
        layout = QVBoxLayout(reminder_dialog)

        customer_list = QTextEdit()
        customer_list.setReadOnly(True)
        customer_list.setFont(QFont("Courier New", 10))
        customer_list.setPlainText(
            self.t("CUSTOMERS SELECTED FOR REMINDER:") + "\n" + "=" * 50 + "\n"
        )
        for customer in customers_to_remind:
            customer_list.append(f"{self.t('Name:')} {customer['name']}")
            customer_list.append(f"{self.t('Phone:')} {customer['phone']}")
            customer_list.append(
                f"{self.t('Warranty Status:')} {self.t(customer['status'])}"
            )
            customer_list.append(f"{self.t('End Date:')} {customer['end_date']}")
            customer_list.append("-" * 50 + "\n")

        layout.addWidget(
            QLabel(
                f"{self.t('TOTAL SELECTED:')} {self.format_number(len(customers_to_remind))} "
                f"{self.t('CUSTOMERS')}"
            )
        )
        layout.addWidget(customer_list)

        message_group = QGroupBox("REMINDER MESSAGE TEMPLATE")
        message_layout = QVBoxLayout()
        self.reminder_message = QTextEdit()
        self.reminder_message.setPlainText(
            self.t(
                "Dear [CUSTOMER],\n\n"
                "Bogor On-Call Computer Service reminds you that your device repair warranty "
                "will expire on:\n"
                "📅 [END DATE]\n\n"
                "Please contact us if you need further service.\n\n"
                "Thank you,\n"
                "Bogor On-Call Computer Service\n"
                "📞 0812-3456-7890"
            )
        )
        message_layout.addWidget(self.reminder_message)
        message_group.setLayout(message_layout)
        layout.addWidget(message_group)

        button_layout = QHBoxLayout()
        log_btn = QPushButton("📝 LOG REMINDERS")
        log_btn.setObjectName("successButton")
        log_btn.clicked.connect(
            lambda: self.confirm_send_reminders(reminder_dialog, customers_to_remind)
        )
        cancel_btn = QPushButton(self.t("CANCEL"))
        cancel_btn.setObjectName("dangerButton")
        cancel_btn.clicked.connect(reminder_dialog.reject)
        button_layout.addStretch()
        button_layout.addWidget(log_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        self.apply_language_setting_to_ui(reminder_dialog)
        reminder_dialog.exec()
    
    def confirm_send_reminders(self, dialog, customers):
        """Record reminder activity in the local database."""
        for customer in customers:
            self.cursor.execute(
                """
                INSERT INTO warranty_reminders
                    (customer_name, phone, reminder_sent_date, next_reminder_date, status)
                VALUES (?, ?, date('now'), date('now', '+7 days'), 'LOGGED')
                """,
                (customer["name"], customer["phone"]),
            )
        self.conn.commit()
        QMessageBox.information(
            self,
            "REMINDERS LOGGED",
            f"REMINDERS LOGGED FOR {len(customers)} CUSTOMERS.\n\n"
            "NO SMS OR WHATSAPP MESSAGE WAS SENT.",
        )
        dialog.accept()
        self.status_bar.showMessage(f"REMINDERS LOGGED FOR {len(customers)} CUSTOMERS")
    
    # Spareparts management methods
    def add_sparepart(self):
        if not self.require_admin():
            return
        part_name = self.sparepart_name.text().strip().upper()
        if not part_name:
            QMessageBox.warning(self, "ERROR", "PART NAME IS REQUIRED")
            self.sparepart_name.setFocus()
            return
        duplicate = self.cursor.execute(
            "SELECT rowid FROM spareparts WHERE UPPER(TRIM(part_name))=? LIMIT 1",
            (part_name,),
        ).fetchone()
        if duplicate:
            QMessageBox.warning(
                self,
                "DUPLICATE PART",
                "A PART WITH THIS NAME ALREADY EXISTS. SELECT IT AND USE SAVE TO UPDATE IT.",
            )
            return

        supplier = self.sparepart_supplier.currentData()
        if not supplier and self.sparepart_supplier.currentText():
            supplier = self.sparepart_supplier.currentText().strip().upper()
        self.cursor.execute(
            """
            INSERT INTO spareparts
                (part_name, part_type, brand, quantity, unit_price, supplier,
                 min_stock, location, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                part_name,
                self.sparepart_type.currentText(),
                self.sparepart_brand.currentText(),
                self.sparepart_quantity.value(),
                self.sparepart_price.value(),
                supplier,
                self.sparepart_min_stock.value(),
                self.sparepart_location.text().strip().upper(),
                self.sparepart_notes.toPlainText().strip().upper(),
            ),
        )
        self.conn.commit()
        self.load_spareparts()
        self.clear_sparepart_form()
        self.status_bar.showMessage(f"ADDED PART: {part_name}")
    
    def edit_selected_sparepart(self):
        if not self.require_admin():
            return
        row = self.spareparts_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A PART TO EDIT")
            return
        self.load_selected_sparepart(row, 0)
        self.sparepart_name.setFocus()
        self.status_bar.showMessage("EDIT MODE: UPDATE FIELDS THEN CLICK UPDATE")

    def update_sparepart(self):
        if not self.require_admin():
            return
        row_id = self.selected_sparepart_row_id
        if row_id is None:
            row_id = self._table_row_id(self.spareparts_table, self.spareparts_table.currentRow())
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A PART TO UPDATE")
            return

        new_part_name = self.sparepart_name.text().strip().upper()
        if not new_part_name:
            QMessageBox.warning(self, "ERROR", "PART NAME IS REQUIRED")
            return
        current_name_row = self.cursor.execute(
            "SELECT part_name FROM spareparts WHERE rowid=?", (row_id,)
        ).fetchone()
        current_name = str(current_name_row[0] or "").strip().upper() if current_name_row else ""
        if new_part_name != current_name:
            duplicate = self.cursor.execute(
                """
                SELECT rowid FROM spareparts
                WHERE UPPER(TRIM(part_name))=? AND rowid<>? LIMIT 1
                """,
                (new_part_name, row_id),
            ).fetchone()
            if duplicate:
                QMessageBox.warning(
                    self, "DUPLICATE PART", "ANOTHER PART ALREADY USES THIS NAME."
                )
                return

        supplier = self.sparepart_supplier.currentData()
        if not supplier and self.sparepart_supplier.currentText():
            supplier = self.sparepart_supplier.currentText().strip().upper()
        self.cursor.execute(
            """
            UPDATE spareparts SET
                part_name=?, part_type=?, brand=?, quantity=?, unit_price=?, supplier=?,
                min_stock=?, location=?, notes=?
            WHERE rowid=?
            """,
            (
                new_part_name,
                self.sparepart_type.currentText(),
                self.sparepart_brand.currentText(),
                self.sparepart_quantity.value(),
                self.sparepart_price.value(),
                supplier,
                self.sparepart_min_stock.value(),
                self.sparepart_location.text().strip().upper(),
                self.sparepart_notes.toPlainText().strip().upper(),
                row_id,
            ),
        )
        self.conn.commit()
        self.load_spareparts()
        self.clear_sparepart_form()
        self.status_bar.showMessage(f"UPDATED PART: {new_part_name}")
    
    def delete_sparepart(self):
        if not self.require_admin():
            return
        row = self.spareparts_table.currentRow()
        row_id = self._table_row_id(self.spareparts_table, row)
        if row_id is None:
            QMessageBox.warning(self, "NO SELECTION", "SELECT A PART TO DELETE")
            return
        item = self.spareparts_table.item(row, 0)
        part_name = item.text() if item else ""
        reply = QMessageBox.question(
            self,
            "CONFIRM DELETE",
            f"DELETE PART '{part_name}'?\n\nTHIS ACTION CANNOT BE UNDONE!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.cursor.execute("DELETE FROM spareparts WHERE rowid=?", (row_id,))
        self.conn.commit()
        self.load_spareparts()
        self.clear_sparepart_form()
        self.status_bar.showMessage(f"DELETED PART: {part_name}")
    
    def load_spareparts(self):
        if not self.conn:
            return
        filter_type = self.combo_source_value(self.sparepart_filter_type)
        filter_brand = self.combo_source_value(self.sparepart_filter_brand)
        query = """
            SELECT rowid, part_name, part_type, brand, quantity, unit_price,
                   supplier, min_stock, location, notes
            FROM spareparts WHERE 1=1
        """
        params = []
        if filter_type != "ALL":
            query += " AND part_type=?"
            params.append(filter_type)
        if filter_brand != "ALL":
            query += " AND brand=?"
            params.append(filter_brand)
        query += " ORDER BY part_name, rowid"
        self.cursor.execute(query, params)
        spareparts = self.cursor.fetchall()
        self.spareparts_table.setRowCount(len(spareparts))
        low_stock_count = 0

        for row, sparepart in enumerate(spareparts):
            row_id, part_name, part_type, brand, quantity, unit_price, supplier, min_stock, location, notes = sparepart
            quantity = int(quantity or 0)
            min_stock = int(min_stock or 0)
            if quantity < min_stock:
                low_stock_count += 1
            data_items = [
                part_name,
                part_type or "",
                brand or "",
                self.format_number(quantity),
                self.format_currency(unit_price),
                supplier or "",
                self.format_number(min_stock),
                location or "",
                notes or "",
            ]
            for col, data in enumerate(data_items):
                item = QTableWidgetItem(str(data))
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, row_id)
                if col == 3:
                    if quantity < min_stock:
                        item.setForeground(QColor(255, 0, 0))
                        item.setBackground(QColor(255, 240, 240))
                    elif quantity == min_stock:
                        item.setForeground(QColor(255, 165, 0))
                self.spareparts_table.setItem(row, col, item)

        if low_stock_count > 0:
            self.low_stock_label.setText(
                self.t(
                    f"⚠️ WARNING: {self.format_number(low_stock_count)} PARTS HAVE LOW STOCK!"
                )
            )
        else:
            self.low_stock_label.clear()
    
    def load_selected_sparepart(self, row, column):
        row_id = self._table_row_id(self.spareparts_table, row)
        if row_id is None:
            return
        self.cursor.execute(
            """
            SELECT part_name, part_type, brand, quantity, unit_price,
                   supplier, min_stock, location, notes
            FROM spareparts WHERE rowid=?
            """,
            (row_id,),
        )
        sparepart = self.cursor.fetchone()
        if not sparepart:
            return
        self.selected_sparepart_row_id = row_id
        self.sparepart_name.setText(sparepart[0] or "")
        for combo, value in (
            (self.sparepart_type, sparepart[1]),
            (self.sparepart_brand, sparepart[2]),
        ):
            value = value or ""
            if value:
                index = combo.findText(value)
                if index >= 0:
                    combo.setCurrentIndex(index)
                else:
                    combo.addItem(value)
                    combo.setCurrentText(value)
            else:
                combo.setCurrentIndex(0)
        self.sparepart_quantity.setValue(int(sparepart[3] or 0))
        self.sparepart_price.setValue(float(sparepart[4] or 0))
        supplier = sparepart[5] or None
        if supplier:
            index = self.sparepart_supplier.findData(supplier)
            if index >= 0:
                self.sparepart_supplier.setCurrentIndex(index)
            else:
                self.sparepart_supplier.addItem(supplier, supplier)
                self.sparepart_supplier.setCurrentText(supplier)
        else:
            self.sparepart_supplier.setCurrentIndex(0)
        self.sparepart_min_stock.setValue(int(sparepart[6] or 5))
        self.sparepart_location.setText(sparepart[7] or "")
        self.sparepart_notes.setText(sparepart[8] or "")
    
    def clear_sparepart_form(self):
        self.selected_sparepart_row_id = None
        self.sparepart_name.clear()
        self.sparepart_type.setCurrentIndex(0)
        self.sparepart_brand.setCurrentIndex(0)
        self.sparepart_quantity.setValue(0)
        self.sparepart_price.setValue(0)
        self.sparepart_supplier.setCurrentIndex(0)
        self.sparepart_min_stock.setValue(5)
        self.sparepart_location.clear()
        self.sparepart_notes.clear()
        self.spareparts_table.clearSelection()
    
    def generate_report(self):
        report_type = self.combo_source_value(self.report_type)
        if self.from_date.date() > self.to_date.date():
            QMessageBox.warning(self, "INVALID DATE RANGE", "FROM DATE CANNOT BE AFTER TO DATE")
            return

        from_date = self.from_date.date().toString("yyyy-MM-dd")
        to_date = self.to_date.date().toString("yyyy-MM-dd")
        display_from_date = self.format_date_for_display(from_date)
        display_to_date = self.format_date_for_display(to_date)
        now = datetime.now()
        py_date_format = get_language_config(self.get_language_code()).get("python_date_format", "%d-%m-%Y")
        generated_stamp = f"{now.strftime(py_date_format)} {now.strftime('%H:%M:%S')}"
        current_date_display = now.strftime(py_date_format)
        
        if report_type == "SERVICE SUMMARY":
            self.cursor.execute('''
                SELECT 
                    status,
                    COUNT(*) as total_services,
                    SUM(CASE WHEN status = 'COMPLETED' THEN 1 ELSE 0 END) as completed,
                    SUM(CASE WHEN status = 'IN REPAIR' THEN 1 ELSE 0 END) as in_repair,
                    SUM(CASE WHEN status = 'WAITING FOR PARTS' THEN 1 ELSE 0 END) as waiting_parts,
                    SUM(harga) as total_revenue
                FROM contacts
                WHERE tanggal_service BETWEEN ? AND ?
                GROUP BY status
                ORDER BY 
                    CASE status 
                        WHEN 'COMPLETED' THEN 1
                        WHEN 'IN REPAIR' THEN 2
                        WHEN 'WAITING FOR PARTS' THEN 3
                        WHEN 'DIAGNOSING' THEN 4
                        WHEN 'NEW' THEN 5
                        WHEN 'CANCELLED' THEN 6
                        ELSE 7
                    END
            ''', (from_date, to_date))
            
            data = self.cursor.fetchall()
            
            self.report_table.setColumnCount(6)
            self.report_table.setHorizontalHeaderLabels([
                "STATUS", "TOTAL SERVICES", "COMPLETED", "IN REPAIR", 
                "WAITING PARTS", "TOTAL REVENUE"
            ])
            self.report_table.setRowCount(len(data))
            
            total_revenue = 0
            total_services = 0
            
            for row, item in enumerate(data):
                status, services, completed, in_repair, waiting_parts, revenue = item
                total_revenue += revenue if revenue else 0
                total_services += services if services else 0
                
                self.report_table.setItem(row, 0, QTableWidgetItem(self.t(status)))
                self.report_table.setItem(row, 1, QTableWidgetItem(self.format_number(services)))
                self.report_table.setItem(row, 2, QTableWidgetItem(self.format_number(completed)))
                self.report_table.setItem(row, 3, QTableWidgetItem(self.format_number(in_repair)))
                self.report_table.setItem(row, 4, QTableWidgetItem(self.format_number(waiting_parts)))
                
                revenue_item = QTableWidgetItem(self.format_currency(revenue))
                self.report_table.setItem(row, 5, revenue_item)
            
            report_text = "=== SERVICE SUMMARY REPORT ===\n"
            report_text += f"Period: {display_from_date} to {display_to_date}\n"
            report_text += f"Generated: {generated_stamp}\n"
            report_text += "="*50 + "\n\n"
            
            for item in data:
                status, services, completed, in_repair, waiting_parts, revenue = item
                report_text += f"{status}:\n"
                report_text += f"  Total Services: {services}\n"
                report_text += f"  Completed: {completed}\n"
                report_text += f"  In Repair: {in_repair}\n"
                report_text += f"  Waiting for Parts: {waiting_parts}\n"
                report_text += f"  Revenue: {self.format_currency(revenue)}\n\n"
            
            report_text += "="*50 + "\n"
            report_text += f"TOTAL SERVICES: {total_services}\n"
            report_text += f"TOTAL REVENUE: {self.format_currency(total_revenue)}\n"
            
            self.report_text.setText(self.t(report_text))
            
        elif report_type == "FINANCIAL REPORT":
            self.cursor.execute('''
                SELECT 
                    strftime('%Y-%m', tanggal_service) as month,
                    COUNT(*) as total_services,
                    SUM(harga) as total_revenue,
                    AVG(harga) as avg_revenue_per_service
                FROM contacts
                WHERE tanggal_service BETWEEN ? AND ?
                    AND harga > 0
                GROUP BY strftime('%Y-%m', tanggal_service)
                ORDER BY month DESC
            ''', (from_date, to_date))
            
            data = self.cursor.fetchall()
            
            self.report_table.setColumnCount(4)
            self.report_table.setHorizontalHeaderLabels([
                "MONTH", "TOTAL SERVICES", "TOTAL REVENUE", "AVG REVENUE/SERVICE"
            ])
            self.report_table.setRowCount(len(data))
            
            for row, item in enumerate(data):
                month, services, revenue, avg_revenue = item
                
                self.report_table.setItem(row, 0, QTableWidgetItem(month))
                self.report_table.setItem(row, 1, QTableWidgetItem(self.format_number(services)))
                
                revenue_item = QTableWidgetItem(self.format_currency(revenue))
                self.report_table.setItem(row, 2, revenue_item)
                
                avg_item = QTableWidgetItem(self.format_currency(avg_revenue))
                self.report_table.setItem(row, 3, avg_item)
            
            report_text = "=== FINANCIAL REPORT ===\n"
            report_text += f"Period: {display_from_date} to {display_to_date}\n"
            report_text += f"Generated: {generated_stamp}\n"
            report_text += "="*50 + "\n\n"
            
            for item in data:
                month, services, revenue, avg_revenue = item
                report_text += f"{month}:\n"
                report_text += f"  Total Services: {services}\n"
                report_text += f"  Total Revenue: {self.format_currency(revenue)}\n"
                report_text += f"  Avg per Service: {self.format_currency(avg_revenue)}\n\n"
            
            self.report_text.setText(self.t(report_text))
            
        elif report_type == "TECHNICIAN PERFORMANCE":
            self.cursor.execute('''
                SELECT 
                    teknisi,
                    COUNT(*) as total_services,
                    SUM(CASE WHEN status = 'COMPLETED' THEN 1 ELSE 0 END) as completed,
                    SUM(CASE WHEN status IN ('IN REPAIR', 'WAITING FOR PARTS') THEN 1 ELSE 0 END) as in_progress,
                    SUM(harga) as total_revenue,
                    ROUND(AVG(harga), 0) as avg_revenue
                FROM contacts
                WHERE tanggal_service BETWEEN ? AND ?
                    AND teknisi IS NOT NULL AND teknisi != ''
                GROUP BY teknisi
                ORDER BY total_services DESC, total_revenue DESC
            ''', (from_date, to_date))
            
            data = self.cursor.fetchall()
            
            self.report_table.setColumnCount(6)
            self.report_table.setHorizontalHeaderLabels([
                "TECHNICIAN", "TOTAL SERVICES", "COMPLETED", 
                "IN PROGRESS", "TOTAL REVENUE", "AVG REVENUE"
            ])
            self.report_table.setRowCount(len(data))
            
            for row, item in enumerate(data):
                technician, services, completed, in_progress, revenue, avg_revenue = item
                
                self.report_table.setItem(row, 0, QTableWidgetItem(technician))
                self.report_table.setItem(row, 1, QTableWidgetItem(self.format_number(services)))
                self.report_table.setItem(row, 2, QTableWidgetItem(self.format_number(completed)))
                self.report_table.setItem(row, 3, QTableWidgetItem(self.format_number(in_progress)))
                
                revenue_item = QTableWidgetItem(self.format_currency(revenue))
                self.report_table.setItem(row, 4, revenue_item)
                
                avg_item = QTableWidgetItem(self.format_currency(avg_revenue))
                self.report_table.setItem(row, 5, avg_item)
            
            report_text = "=== TECHNICIAN PERFORMANCE REPORT ===\n"
            report_text += f"Period: {display_from_date} to {display_to_date}\n"
            report_text += f"Generated: {generated_stamp}\n"
            report_text += "="*50 + "\n\n"
            
            for item in data:
                technician, services, completed, in_progress, revenue, avg_revenue = item
                completion_rate = (completed / services * 100) if services > 0 else 0
                
                report_text += f"{technician}:\n"
                report_text += f"  Total Services: {services}\n"
                report_text += f"  Completed: {completed} ({completion_rate:.1f}%)\n"
                report_text += f"  In Progress: {in_progress}\n"
                report_text += f"  Total Revenue: {self.format_currency(revenue)}\n"
                report_text += f"  Avg per Service: {self.format_currency(avg_revenue)}\n\n"
            
            self.report_text.setText(self.t(report_text))
            
        elif report_type == "DEVICE BRAND STATISTICS":
            self.cursor.execute('''
                SELECT 
                    merek,
                    COUNT(*) as total_devices,
                    COUNT(DISTINCT model) as unique_models,
                    SUM(CASE WHEN status = 'COMPLETED' THEN 1 ELSE 0 END) as repaired,
                    SUM(CASE WHEN status IN ('IN REPAIR', 'WAITING FOR PARTS') THEN 1 ELSE 0 END) as pending,
                    ROUND(AVG(harga), 0) as avg_repair_cost
                FROM contacts
                WHERE tanggal_service BETWEEN ? AND ?
                    AND merek IS NOT NULL AND merek != ''
                GROUP BY merek
                ORDER BY total_devices DESC
            ''', (from_date, to_date))
            
            data = self.cursor.fetchall()
            
            self.report_table.setColumnCount(6)
            self.report_table.setHorizontalHeaderLabels([
                "BRAND", "TOTAL DEVICES", "UNIQUE MODELS", 
                "REPAIRED", "PENDING", "AVG REPAIR COST"
            ])
            self.report_table.setRowCount(len(data))
            
            for row, item in enumerate(data):
                brand, devices, models, repaired, pending, avg_cost = item
                
                self.report_table.setItem(row, 0, QTableWidgetItem(brand))
                self.report_table.setItem(row, 1, QTableWidgetItem(self.format_number(devices)))
                self.report_table.setItem(row, 2, QTableWidgetItem(self.format_number(models)))
                self.report_table.setItem(row, 3, QTableWidgetItem(self.format_number(repaired)))
                self.report_table.setItem(row, 4, QTableWidgetItem(self.format_number(pending)))
                
                cost_item = QTableWidgetItem(self.format_currency(avg_cost))
                self.report_table.setItem(row, 5, cost_item)
            
            report_text = "=== DEVICE BRAND STATISTICS ===\n"
            report_text += f"Period: {display_from_date} to {display_to_date}\n"
            report_text += f"Generated: {generated_stamp}\n"
            report_text += "="*50 + "\n\n"
            
            for item in data:
                brand, devices, models, repaired, pending, avg_cost = item
                repair_rate = (repaired / devices * 100) if devices > 0 else 0
                
                report_text += f"{brand}:\n"
                report_text += f"  Total Devices: {devices}\n"
                report_text += f"  Unique Models: {models}\n"
                report_text += f"  Repaired: {repaired} ({repair_rate:.1f}%)\n"
                report_text += f"  Pending: {pending}\n"
                report_text += f"  Avg Repair Cost: {self.format_currency(avg_cost)}\n\n"
            
            self.report_text.setText(self.t(report_text))
            
        elif report_type == "WARRANTY SUMMARY":
            self.cursor.execute('''
                SELECT 
                    garansi,
                    COUNT(*) as total_customers,
                    SUM(CASE 
                        WHEN COALESCE(TRIM(garansi), '') NOT IN ('', 'NO WARRANTY') AND COALESCE(TRIM(berakhir_garansi), '') <> '' AND berakhir_garansi >= date('now') THEN 1 
                        ELSE 0 
                    END) as active_warranties,
                    SUM(CASE 
                        WHEN COALESCE(TRIM(garansi), '') NOT IN ('', 'NO WARRANTY') AND COALESCE(TRIM(berakhir_garansi), '') <> '' AND berakhir_garansi BETWEEN date('now') AND date('now', '+30 days') THEN 1 
                        ELSE 0 
                    END) as expiring_soon,
                    SUM(CASE 
                        WHEN COALESCE(TRIM(garansi), '') NOT IN ('', 'NO WARRANTY') AND COALESCE(TRIM(berakhir_garansi), '') <> '' AND berakhir_garansi < date('now') THEN 1 
                        ELSE 0 
                    END) as expired
                FROM contacts
                WHERE tanggal_service BETWEEN ? AND ?
                GROUP BY garansi
                ORDER BY 
                    CASE garansi 
                        WHEN 'NO WARRANTY' THEN 1
                        WHEN '30 DAYS' THEN 2
                        WHEN '60 DAYS' THEN 3
                        WHEN '90 DAYS' THEN 4
                        WHEN '6 MONTHS' THEN 5
                        WHEN '1 YEAR' THEN 6
                        ELSE 7
                    END
            ''', (from_date, to_date))
            
            data = self.cursor.fetchall()
            
            self.report_table.setColumnCount(5)
            self.report_table.setHorizontalHeaderLabels([
                "WARRANTY PERIOD", "TOTAL CUSTOMERS", "ACTIVE", 
                "EXPIRING SOON (≤30 days)", "EXPIRED"
            ])
            self.report_table.setRowCount(len(data))
            
            for row, item in enumerate(data):
                warranty, total, active, expiring, expired = item
                
                self.report_table.setItem(row, 0, QTableWidgetItem(self.t(warranty if warranty else "NO WARRANTY")))
                self.report_table.setItem(row, 1, QTableWidgetItem(self.format_number(total)))
                self.report_table.setItem(row, 2, QTableWidgetItem(self.format_number(active)))
                self.report_table.setItem(row, 3, QTableWidgetItem(self.format_number(expiring)))
                self.report_table.setItem(row, 4, QTableWidgetItem(self.format_number(expired)))
            
            report_text = "=== WARRANTY SUMMARY REPORT ===\n"
            report_text += f"Period: {display_from_date} to {display_to_date}\n"
            report_text += f"Generated: {generated_stamp}\n"
            report_text += f"Current Date: {current_date_display}\n"
            report_text += "="*50 + "\n\n"
            
            total_active = 0
            total_expiring = 0
            total_expired = 0
            
            for item in data:
                warranty, total, active, expiring, expired = item
                warranty_display = warranty if warranty else "NO WARRANTY"
                
                total_active += active
                total_expiring += expiring
                total_expired += expired
                
                report_text += f"{warranty_display}:\n"
                report_text += f"  Total Customers: {total}\n"
                report_text += f"  Active: {active}\n"
                report_text += f"  Expiring Soon: {expiring}\n"
                report_text += f"  Expired: {expired}\n\n"
            
            report_text += "="*50 + "\n"
            report_text += "SUMMARY:\n"
            report_text += f"  Total Active Warranties: {total_active}\n"
            report_text += f"  Total Expiring Soon: {total_expiring}\n"
            report_text += f"  Total Expired: {total_expired}\n"
            
            self.report_text.setText(self.t(report_text))

        self._translate_table_headers()
    
    def _start_detached_program(self, program_path, arguments=None):
        """Start an external program without blocking the main application."""
        program_path = os.path.abspath(program_path)
        arguments = list(arguments or [])
        working_directory = os.path.dirname(program_path) or self.app_dir

        try:
            result = QProcess.startDetached(program_path, arguments, working_directory)
            if isinstance(result, tuple):
                return bool(result[0])
            return bool(result)
        except Exception:
            return False

    def _find_external_python_interpreter(self):
        """Return an interpreter command for Dummy_Creator.py when available."""
        if not (getattr(sys, "frozen", False) or "__compiled__" in globals()):
            executable = os.path.abspath(sys.executable)
            if os.path.isfile(executable):
                return executable, []

        candidates = (
            ("pythonw.exe", []),
            ("pyw.exe", ["-3"]),
            ("python.exe", []),
            ("py.exe", ["-3"]),
        ) if sys.platform.startswith("win") else (
            ("python3", []),
            ("python", []),
        )
        for command, prefix_arguments in candidates:
            resolved = shutil.which(command)
            if resolved:
                return resolved, list(prefix_arguments)
        return None, []

    def open_external_dummy_creator(self, checked=False):
        """Launch Dummy_Creator.exe or Dummy_Creator.py beside the application."""
        if not self.require_admin():
            return False
        del checked
        creator_arguments = [
            "--output", self.dropdown_data_path,
            "--language", self.get_language_code(),
        ]

        executable_path = get_app_file_path("Dummy_Creator.exe")
        if os.path.isfile(executable_path):
            if self._start_detached_program(executable_path, creator_arguments):
                self.set_status(self.t("MULTILINGUAL DUMMY CREATOR OPENED"))
                return True

        script_path = get_app_file_path("Dummy_Creator.py")
        if os.path.isfile(script_path):
            interpreter, prefix_arguments = self._find_external_python_interpreter()
            if interpreter:
                arguments = list(prefix_arguments) + [script_path] + creator_arguments
                if self._start_detached_program(interpreter, arguments):
                    self.set_status(self.t("MULTILINGUAL DUMMY CREATOR OPENED"))
                    return True

        expected = (
            f"{executable_path}\n"
            f"{script_path}"
        )
        self.ui_warn(
            self.t("DUMMY CREATOR NOT FOUND"),
            self.t("PLACE DUMMY_CREATOR.EXE OR DUMMY_CREATOR.PY BESIDE THE APPLICATION.")
            + "\n\n"
            + self.t("EXPECTED FILE")
            + f":\n{expected}\n\n"
            + self.t("FOR A COMPILED APPLICATION, DUMMY_CREATOR.EXE IS RECOMMENDED."),
        )
        return False

    def open_dummy_data_editor(self):
        """Open the built-in dummy.ini editor; no second executable is required."""
        if not self.require_admin():
            return False
        dialog = self._dummy_editor_dialog
        if dialog is not None and dialog.isVisible():
            dialog.raise_()
            dialog.activateWindow()
            return True

        dialog = BuiltInDummyDataEditor(self)
        self._dummy_editor_dialog = dialog
        dialog.finished.connect(lambda _result: setattr(self, "_dummy_editor_dialog", None))
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        self.set_status(self.t("DUMMY DATA EDITOR OPENED"))
        return True

    def current_help_topic(self):
        """Return the CHM topic associated with the active main tab."""
        topics = {
            0: "user_data.html",
            1: "suppliers.html",
            2: "reports.html",
            3: "warranty.html",
            4: "parts.html",
        }
        index = self.tab_widget.currentIndex() if hasattr(self, "tab_widget") else 0
        return topics.get(index, "index.html")

    def open_current_tab_help(self, checked=False):
        del checked
        return self.open_help_contents(self.current_help_topic())

    def open_help_contents(self, topic="index.html"):
        """Open a topic from Help.chm beside the compiled application."""
        if isinstance(topic, bool) or not topic:
            topic = "index.html"
        topic = str(topic).replace("\\", "/").lstrip("/")
        help_path = get_app_file_path("Help.chm")

        opened = False
        if os.path.isfile(help_path):
            if sys.platform.startswith("win"):
                windows_dir = os.environ.get("WINDIR", r"C:\Windows")
                html_help_viewer = os.path.join(windows_dir, "hh.exe")
                target = f"mk:@MSITStore:{help_path}::/{topic}"
                if os.path.isfile(html_help_viewer):
                    opened = self._start_detached_program(html_help_viewer, [target])

                if not opened:
                    try:
                        os.startfile(help_path)  # type: ignore[attr-defined]
                        opened = True
                    except (AttributeError, OSError):
                        opened = False
            else:
                opened = QDesktopServices.openUrl(QUrl.fromLocalFile(help_path))

        # Development/deployment fallback: the HTML source can be placed in
        # either Help or Help_Project beside the executable.
        if not opened:
            for folder_name in ("Help", "Help_Project"):
                html_topic = get_app_file_path(os.path.join(folder_name, topic))
                if not os.path.isfile(html_topic):
                    html_topic = get_app_file_path(os.path.join(folder_name, "index.html"))
                if os.path.isfile(html_topic):
                    opened = QDesktopServices.openUrl(QUrl.fromLocalFile(html_topic))
                    if opened:
                        break

        if not opened:
            self.ui_warn(
                self.t("HELP FILE NOT FOUND"),
                self.t("HELP.CHM WAS NOT FOUND")
                + "\n\n"
                + self.t("EXPECTED FILE")
                + f":\n{help_path}\n\n"
                + self.t("BUILD HELP.HHP AND COPY HELP.CHM TO THE APPLICATION FOLDER."),
            )
            return False

        self.set_status(self.t("HELP OPENED"))
        return True

    def show_about(self):
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle("ABOUT")
        about_dialog.setFixedSize(540, 485)

        layout = QVBoxLayout(about_dialog)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(10)

        title = QLabel("COMPUTER SERVICE MANAGER v4.4")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setObjectName("accentLabel")
        title.setStyleSheet(
            "font-size: 18px;"
            "font-weight: bold;"
            "margin-bottom: 5px;"
        )

        subtitle = QLabel("SERVICE KOMPUTER PANGGILAN BOGOR")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet(
            "font-size: 14px;"
            "font-weight: bold;"
            "margin-bottom: 10px;"
        )

        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)

        created_by = QLabel("CREATED BY:")
        created_by.setAlignment(Qt.AlignmentFlag.AlignCenter)
        created_by.setStyleSheet(
            "font-size: 12px;"
            "margin-bottom: 5px;"
        )

        author = QLabel("RAHFIE27")
        author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        author.setObjectName("successLabel")
        author.setStyleSheet(
            "font-size: 16px;"
            "font-weight: bold;"
            "margin-bottom: 8px;"
        )

        contact_label = QLabel(
            'CONTACT:<br>'
            '<a href="mailto:e-comtech@mail.com">'
            'e-comtech@mail.com'
            '</a> / '
            '<a href="mailto:rahfie27@gmail.com">'
            'rahfie27@gmail.com'
            '</a>'
        )
        contact_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        contact_label.setTextFormat(Qt.TextFormat.RichText)
        contact_label.setOpenExternalLinks(True)
        contact_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextBrowserInteraction
        )
        contact_label.setStyleSheet(
            "font-size: 11px;"
            "margin-bottom: 6px;"
        )

        donation_label = QLabel(
            'DONATION:<br>'
            '<a href="https://paypal.me/rahfie">'
            'paypal.me/rahfie'
            '</a>'
        )
        donation_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        donation_label.setTextFormat(Qt.TextFormat.RichText)
        donation_label.setOpenExternalLinks(True)
        donation_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextBrowserInteraction
        )
        donation_label.setStyleSheet(
            "font-size: 11px;"
            "margin-bottom: 10px;"
        )

        copyright_label = QLabel("COPYRIGHT © ECOMTECH 2026")
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        copyright_label.setStyleSheet(
            "font-size: 10px;"
            "margin-top: 10px;"
            "margin-bottom: 5px;"
        )

        rights_label = QLabel("ALL RIGHTS RESERVED")
        rights_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rights_label.setStyleSheet(
            "font-size: 10px;"
            "margin-bottom: 15px;"
        )

        warning = QLabel("⚠️ WARNING!")
        warning.setAlignment(Qt.AlignmentFlag.AlignCenter)
        warning.setObjectName("alertLabel")
        warning.setStyleSheet(
            "font-size: 12px;"
            "font-weight: bold;"
            "margin-bottom: 5px;"
        )

        disclaimer = QLabel("THIS SOFTWARE IS PROVIDED AS-IS")
        disclaimer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        disclaimer.setStyleSheet(
            "font-size: 10px;"
            "margin-bottom: 5px;"
        )

        no_warranty = QLabel("NO WARRANTY OF ANY KIND IS PROVIDED")
        no_warranty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        no_warranty.setStyleSheet(
            "font-size: 9px;"
            "font-style: italic;"
        )

        close_btn = QPushButton("CLOSE")
        close_btn.clicked.connect(about_dialog.accept)
        close_btn.setFixedWidth(100)
        close_btn.setObjectName("primaryButton")

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(separator)
        layout.addWidget(created_by)
        layout.addWidget(author)
        layout.addWidget(contact_label)
        layout.addWidget(donation_label)
        layout.addWidget(copyright_label)
        layout.addWidget(rights_label)
        layout.addWidget(warning)
        layout.addWidget(disclaimer)
        layout.addWidget(no_warranty)
        layout.addStretch()
        layout.addWidget(
            close_btn,
            alignment=Qt.AlignmentFlag.AlignCenter
        )

        self.apply_language_setting_to_ui(about_dialog)
        about_dialog.exec()
    
    def closeEvent(self, event):
        if self.current_user:
            self.audit_event("APPLICATION_EXIT", details="WINDOW CLOSED")
        self.shutdown()
        self.set_status("APPLICATION CLOSED")
        event.accept()
        app = QApplication.instance()
        if app is not None:
            QTimer.singleShot(0, app.quit)


class UpperCaseLineEdit(QLineEdit):
    """Custom QLineEdit that converts text to uppercase"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
    def focusOutEvent(self, event):
        text = self.text()
        if text:
            self.setText(text.upper())
        super().focusOutEvent(event)


class LowerCaseLineEdit(QLineEdit):
    """Custom QLineEdit that converts text to lowercase (useful for URLs)."""
    def __init__(self, parent=None):
        super().__init__(parent)

    def focusOutEvent(self, event):
        text = self.text()
        if text:
            self.setText(text.lower())
        super().focusOutEvent(event)


class UpperCaseTextEdit(QTextEdit):
    """Custom QTextEdit that converts text to uppercase"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
    def focusOutEvent(self, event):
        text = self.toPlainText()
        if text:
            self.setPlainText(text.upper())
        super().focusOutEvent(event)


class PriceSpinBox(QDoubleSpinBox):
    """Custom QDoubleSpinBox for price input with localized currency formatting"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setProperty("hideSpinButtons", True)
        self.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.currency_code = DEFAULT_CURRENCY_CODE
        self.apply_currency_rules()
        self.setMaximum(999999999)
        self.setSingleStep(10000)

    def apply_currency_rules(self):
        cfg = get_currency_config(self.currency_code)
        self.setDecimals(int(cfg.get("decimals", 0)))

    def set_currency_code(self, currency_code):
        code = str(currency_code or DEFAULT_CURRENCY_CODE).strip().upper()
        self.currency_code = code if code in CURRENCY_OPTIONS else DEFAULT_CURRENCY_CODE
        self.apply_currency_rules()
        line_edit = self.lineEdit()
        if line_edit is not None:
            line_edit.setText(self.textFromValue(self.value()))

    def set_currency_symbol(self, symbol):
        symbol = str(symbol or "").strip()
        matched_code = None
        for code, data in CURRENCY_OPTIONS.items():
            if data.get("symbol") == symbol:
                matched_code = code
                break
        self.set_currency_code(matched_code or DEFAULT_CURRENCY_CODE)

    def textFromValue(self, value):
        return format_currency_value(value, self.currency_code)

    def valueFromText(self, text):
        return parse_currency_value(text, self.currency_code)

    def focusOutEvent(self, event):
        current_value = self.value()
        self.setValue(current_value)
        super().focusOutEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    app.setStyle('Fusion')
    icon_path = get_app_file_path("app.ico")
    if not os.path.exists(icon_path):
        icon_path = DEFAULT_APP_ICON_PATH
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    install_localization_runtime_hooks()
    
    # Set application font
    font = QFont()
    font.setFamily("Segoe UI")
    font.setPointSize(8)
    app.setFont(font)
    
    # Apply a safe default until ServiceManager loads the saved theme.
    apply_application_theme(app, DEFAULT_THEME_CODE)

    window = ServiceManager()
    app.aboutToQuit.connect(window.shutdown)
    window.show()

    exit_code = app.exec()
    try:
        window.shutdown()
        window.deleteLater()
        app.processEvents()
    except RuntimeError:
        pass
    return exit_code

if __name__ == '__main__':
    sys.exit(main())
