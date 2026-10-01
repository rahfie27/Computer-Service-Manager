#!/usr/bin/env python3
"""Multilingual dummy.ini generator for Computer Service Manager 4.1.

The PyQt6 interface contains generation controls only. User preferences are
stored under the current user's Windows Registry. ``--write-only`` remains
available for creating the file without PyQt6.
"""

from __future__ import annotations

import argparse
import configparser
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List

try:
    import winreg
except ImportError:  # Available on Windows only.
    winreg = None  # type: ignore[assignment]

QT_IMPORT_ERROR = None
try:
    from PyQt6.QtCore import QByteArray, Qt
    from PyQt6.QtGui import QAction, QActionGroup, QColor, QIcon, QPalette
    from PyQt6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFormLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QSizePolicy,
        QStatusBar,
        QStyleFactory,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # GUI-only dependency; --write-only still works.
    QT_IMPORT_ERROR = exc
    QMainWindow = object  # type: ignore[assignment,misc]

APP_TITLE = "Dummy-INI-Ersteller"
DEFAULT_LANGUAGE_CODE = "de_DE"

DEFAULT_THEME_CODE = "dark"
THEME_ORDER = ("light", "dark", "blue", "purple", "red", "orange", "classic")

THEME_OPTIONS = {
    "light": {
        "label": "LIGHT",
        "window": "#f3f6fa", "panel": "#ffffff", "panel_alt": "#f8fafc",
        "surface": "#e8eef5", "surface_hover": "#dce6f1", "surface_pressed": "#cbd8e6",
        "input": "#ffffff", "text": "#17202a", "muted": "#5f6b7a",
        "border": "#c4cfdb", "border_strong": "#9eacbc", "accent": "#2563eb",
        "accent_hover": "#1d4ed8", "accent_pressed": "#1e40af", "accent_text": "#ffffff", "selection": "#2563eb",
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
        "accent_hover": "#3d8bfd", "accent_pressed": "#1557b0", "accent_text": "#ffffff", "selection": "#1f6feb",
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
        "accent_hover": "#42a5f5", "accent_pressed": "#1475bd", "accent_text": "#ffffff", "selection": "#1976d2",
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
        "accent_hover": "#b264df", "accent_pressed": "#75369f", "accent_text": "#ffffff", "selection": "#8e44ad",
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
        "accent_hover": "#ef5350", "accent_pressed": "#b71c1c", "accent_text": "#ffffff", "selection": "#d32f2f",
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
        "accent_hover": "#fb8c00", "accent_pressed": "#c65f00", "accent_text": "#ffffff", "selection": "#ef6c00",
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
        "accent_hover": "#1976bd", "accent_pressed": "#064b82", "accent_text": "#ffffff", "selection": "#0a64ad",
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
    p = get_theme_config(theme_code)
    return f"""
        QWidget {{
            color: {p['text']};
            background-color: {p['window']};
            font-size: 8pt;
        }}
        QMainWindow, QDialog {{ background-color: {p['window']}; }}
        QMenuBar {{
            background-color: {p['menu']}; color: {p['text']};
            border-bottom: 1px solid {p['border']};
        }}
        QMenuBar::item {{ background: transparent; padding: 4px 8px; }}
        QMenuBar::item:selected, QMenu::item:selected {{
            background-color: {p['surface_hover']}; color: {p['text']};
        }}
        QMenu {{
            background-color: {p['menu']}; color: {p['text']};
            border: 1px solid {p['border_strong']}; padding: 3px;
        }}
        QMenu::item {{ padding: 5px 24px 5px 9px; }}
        QMenu::item:disabled {{ color: {p['disabled_text']}; }}
        QGroupBox {{
            font-weight: 600; border: 1px solid {p['border']}; border-radius: {p['group_radius']};
            margin-top: 8px; padding-top: 8px; background-color: {p['panel']};
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; left: 9px; padding: 0 5px; color: {p['accent']};
        }}
        QLabel {{ background: transparent; }}
        QLabel#accentLabel {{ color: {p['accent']}; }}
        QLabel#successLabel {{ color: {p['success']}; }}
        QLabel#alertLabel {{ color: {p['danger']}; font-weight: bold; }}
        QFrame[frameShape="4"], QFrame[frameShape="5"] {{ color: {p['border']}; }}
        QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QDateEdit,
        QSpinBox, QDoubleSpinBox, QListWidget {{
            min-height: 22px; padding: 3px 6px; border: 1px solid {p['border_strong']};
            border-radius: {p['radius']}; background-color: {p['input']}; color: {p['text']};
            selection-background-color: {p['selection']}; selection-color: {p['accent_text']};
        }}
        QTextEdit, QPlainTextEdit {{ padding: 4px 6px; }}
        QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus,
        QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
        QListWidget:focus {{ border: 1px solid {p['accent']}; }}
        QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled,
        QDateEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled {{
            color: {p['disabled_text']}; background-color: {p['disabled_bg']};
        }}
        QComboBox::drop-down, QDateEdit::drop-down {{
            subcontrol-origin: padding; subcontrol-position: top right; width: 28px;
            border-left: 1px solid {p['border']}; background-color: {p['surface']};
        }}
        QComboBox::drop-down:hover, QDateEdit::drop-down:hover {{ background-color: {p['surface_hover']}; }}
        QComboBox QAbstractItemView, QAbstractItemView {{
            background-color: {p['panel']}; color: {p['text']}; border: 1px solid {p['border_strong']};
            selection-background-color: {p['selection']}; selection-color: {p['accent_text']};
            outline: 0; padding: 3px;
        }}
        QComboBox QAbstractItemView::item {{ min-height: 26px; padding: 4px 8px; border: 0; }}
        QPushButton, QToolButton {{
            min-height: 25px; padding: 3px 9px; border: 1px solid {p['border_strong']};
            border-radius: {p['radius']}; background-color: {p['surface']}; color: {p['text']};
            font-weight: 600;
        }}
        QPushButton:hover, QToolButton:hover {{
            background-color: {p['surface_hover']}; border-color: {p['accent']};
        }}
        QPushButton:pressed, QToolButton:pressed {{ background-color: {p['surface_pressed']}; }}
        QPushButton:disabled, QToolButton:disabled {{
            background-color: {p['disabled_bg']}; color: {p['disabled_text']}; border-color: {p['border']};
        }}
        QPushButton#primaryButton {{ background-color: {p['accent']}; border-color: {p['accent_hover']}; color: {p['accent_text']}; }}
        QPushButton#primaryButton:hover {{ background-color: {p['accent_hover']}; }}
        QPushButton#primaryButton:pressed {{ background-color: {p['accent_pressed']}; }}
        QPushButton#successButton {{ background-color: {p['success']}; color: #ffffff; }}
        QPushButton#warningButton {{ background-color: {p['warning']}; color: #ffffff; }}
        QPushButton#dangerButton {{ background-color: {p['danger']}; color: #ffffff; }}
        QPushButton#neutralButton {{ background-color: {p['neutral']}; color: #ffffff; }}
        QPushButton#secondaryButton {{ background-color: {p['surface']}; border-color: {p['border_strong']}; }}
        QScrollBar:vertical {{ background: {p['scroll_bg']}; width: 12px; margin: 0; }}
        QScrollBar::handle:vertical {{ background: {p['scroll_handle']}; min-height: 28px; border-radius: 5px; }}
        QScrollBar:horizontal {{ background: {p['scroll_bg']}; height: 12px; margin: 0; }}
        QScrollBar::handle:horizontal {{ background: {p['scroll_handle']}; min-width: 28px; border-radius: 5px; }}
        QScrollBar::handle:hover {{ background: {p['accent']}; }}
        QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
        QStatusBar {{
            background-color: {p['menu']}; color: {p['muted']}; border-top: 1px solid {p['border']};
        }}
        QToolTip {{
            background-color: {p['tooltip']}; color: {p['tooltip_text']};
            border: 1px solid {p['border_strong']}; padding: 4px;
        }}
        QProgressBar {{
            border: 1px solid {p['border_strong']}; border-radius: {p['radius']};
            background: {p['input']}; color: {p['text']}; text-align: center;
        }}
        QProgressBar::chunk {{ background-color: {p['accent']}; }}
    """


def apply_application_theme(app, theme_code=None):
    """Apply a complete, cross-platform application theme and return its code."""
    code = str(theme_code or DEFAULT_THEME_CODE).strip().lower()
    if code not in THEME_OPTIONS:
        code = DEFAULT_THEME_CODE
    p = THEME_OPTIONS[code]

    try:
        available = {name.lower(): name for name in QStyleFactory.keys()}
        preferred = "windows" if code == "classic" else "fusion"
        style_name = available.get(preferred) or available.get("fusion")
        if style_name:
            app.setStyle(style_name)
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

CATEGORY_ORDER = (
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

LANGUAGE_OPTIONS = {
    "en_US": "English (United States)",
    "id_ID": "Bahasa Indonesia (Indonesia)",
    "es_MX": "Español (México)",
    "fr_FR": "Français (France)",
    "de_DE": "Deutsch (Deutschland)",
    "pt_BR": "Português (Brasil)",
    "it_IT": "Italiano (Italia)",
    "nl_NL": "Nederlands (Nederland)",
    "zh_CN": "简体中文 (中国)",
    "ja_JP": "日本語 (日本)",
    "ko_KR": "한국어 (대한민국)",
    "ar_SA": "العربية (السعودية)",
    "hi_IN": "हिन्दी (भारत)",
    "bn_BD": "বাংলা (বাংলাদেশ)",
    "ru_RU": "Русский (Россия)",
    "tr_TR": "Türkçe (Türkiye)",
    "vi_VN": "Tiếng Việt (Việt Nam)",
    "th_TH": "ไทย (ประเทศไทย)",
    "ms_MY": "Bahasa Melayu (Malaysia)",
    "fil_PH": "Filipino (Pilipinas)",
}

COMMON_DATA = {
    "brands": [
        "ACER", "APPLE", "ASUS", "DELL", "HP", "LENOVO", "MSI",
        "SAMSUNG", "SONY", "TOSHIBA",
    ],
    "models": [
        "ASPIRE 5", "IDEAPAD 5", "LATITUDE 5420", "MACBOOK AIR",
        "PAVILION 15", "ROG STRIX", "THINKPAD T14", "VIVOBOOK 14",
    ],
}

LOCALIZED_DATA: Dict[str, Dict[str, List[str]]] = {
    "en_US": {
        "device_types": ["DESKTOP", "LAPTOP", "ALL-IN-ONE", "TABLET", "SMARTPHONE", "PRINTER"],
        "issues": ["NO POWER", "SLOW PERFORMANCE", "OVERHEATING", "BLUE SCREEN", "BROKEN DISPLAY", "BATTERY NOT CHARGING", "NETWORK PROBLEM", "DATA RECOVERY"],
        "statuses": ["RECEIVED", "DIAGNOSIS", "WAITING FOR PART", "IN REPAIR", "READY", "COMPLETED", "CANCELLED"],
        "technicians": ["ALEX", "CHRIS", "JORDAN", "MORGAN", "TAYLOR"],
        "warranty_periods": ["NO WARRANTY", "7 DAYS", "14 DAYS", "30 DAYS", "3 MONTHS", "6 MONTHS", "1 YEAR"],
        "parts": ["BATTERY", "CHARGER", "DISPLAY", "FAN", "KEYBOARD", "MAINBOARD", "RAM", "SSD"],
        "sparepart_types": ["POWER", "DISPLAY", "STORAGE", "MEMORY", "COOLING", "INPUT DEVICE", "NETWORK", "MAINBOARD"],
    },
    "id_ID": {
        "device_types": ["KOMPUTER DESKTOP", "LAPTOP", "ALL-IN-ONE", "TABLET", "PONSEL", "PRINTER"],
        "issues": ["MATI TOTAL", "KINERJA LAMBAT", "TERLALU PANAS", "LAYAR BIRU", "LAYAR RUSAK", "BATERAI TIDAK MENGISI", "MASALAH JARINGAN", "PEMULIHAN DATA"],
        "statuses": ["DITERIMA", "DIAGNOSIS", "MENUNGGU SUKU CADANG", "DALAM PERBAIKAN", "SIAP", "SELESAI", "DIBATALKAN"],
        "technicians": ["ANDI", "BUDI", "CITRA", "DEWI", "EKO"],
        "warranty_periods": ["TANPA GARANSI", "7 HARI", "14 HARI", "30 HARI", "3 BULAN", "6 BULAN", "1 TAHUN"],
        "parts": ["BATERAI", "PENGISI DAYA", "LAYAR", "KIPAS", "KEYBOARD", "MAINBOARD", "RAM", "SSD"],
        "sparepart_types": ["DAYA", "LAYAR", "PENYIMPANAN", "MEMORI", "PENDINGIN", "PERANGKAT INPUT", "JARINGAN", "MAINBOARD"],
    },
    "es_MX": {
        "device_types": ["COMPUTADORA DE ESCRITORIO", "PORTÁTIL", "TODO EN UNO", "TABLETA", "TELÉFONO", "IMPRESORA"],
        "issues": ["NO ENCIENDE", "RENDIMIENTO LENTO", "SOBRECALENTAMIENTO", "PANTALLA AZUL", "PANTALLA ROTA", "BATERÍA NO CARGA", "PROBLEMA DE RED", "RECUPERACIÓN DE DATOS"],
        "statuses": ["RECIBIDO", "DIAGNÓSTICO", "ESPERANDO PIEZA", "EN REPARACIÓN", "LISTO", "COMPLETADO", "CANCELADO"],
        "technicians": ["CARLOS", "DANIELA", "ELENA", "JAVIER", "MIGUEL"],
        "warranty_periods": ["SIN GARANTÍA", "7 DÍAS", "14 DÍAS", "30 DÍAS", "3 MESES", "6 MESES", "1 AÑO"],
        "parts": ["BATERÍA", "CARGADOR", "PANTALLA", "VENTILADOR", "TECLADO", "TARJETA MADRE", "RAM", "SSD"],
        "sparepart_types": ["ENERGÍA", "PANTALLA", "ALMACENAMIENTO", "MEMORIA", "REFRIGERACIÓN", "DISPOSITIVO DE ENTRADA", "RED", "TARJETA MADRE"],
    },
    "fr_FR": {
        "device_types": ["ORDINATEUR DE BUREAU", "ORDINATEUR PORTABLE", "TOUT-EN-UN", "TABLETTE", "TÉLÉPHONE", "IMPRIMANTE"],
        "issues": ["NE S'ALLUME PAS", "PERFORMANCES LENTES", "SURCHAUFFE", "ÉCRAN BLEU", "ÉCRAN CASSÉ", "BATTERIE NON CHARGÉE", "PROBLÈME RÉSEAU", "RÉCUPÉRATION DE DONNÉES"],
        "statuses": ["REÇU", "DIAGNOSTIC", "EN ATTENTE DE PIÈCE", "EN RÉPARATION", "PRÊT", "TERMINÉ", "ANNULÉ"],
        "technicians": ["ALAIN", "CAMILLE", "ÉLODIE", "JULIEN", "SOPHIE"],
        "warranty_periods": ["SANS GARANTIE", "7 JOURS", "14 JOURS", "30 JOURS", "3 MOIS", "6 MOIS", "1 AN"],
        "parts": ["BATTERIE", "CHARGEUR", "ÉCRAN", "VENTILATEUR", "CLAVIER", "CARTE MÈRE", "RAM", "SSD"],
        "sparepart_types": ["ALIMENTATION", "AFFICHAGE", "STOCKAGE", "MÉMOIRE", "REFROIDISSEMENT", "PÉRIPHÉRIQUE D'ENTRÉE", "RÉSEAU", "CARTE MÈRE"],
    },
    "de_DE": {
        "device_types": ["DESKTOP-PC", "LAPTOP", "ALL-IN-ONE", "TABLET", "SMARTPHONE", "DRUCKER"],
        "issues": ["KEIN STROM", "LANGSAME LEISTUNG", "ÜBERHITZUNG", "BLAUER BILDSCHIRM", "DISPLAY DEFEKT", "AKKU LÄDT NICHT", "NETZWERKPROBLEM", "DATENWIEDERHERSTELLUNG"],
        "statuses": ["ANGENOMMEN", "DIAGNOSE", "WARTET AUF ERSATZTEIL", "IN REPARATUR", "BEREIT", "ABGESCHLOSSEN", "STORNIERT"],
        "technicians": ["ANNA", "FELIX", "HANNA", "LUKAS", "MAX"],
        "warranty_periods": ["KEINE GARANTIE", "7 TAGE", "14 TAGE", "30 TAGE", "3 MONATE", "6 MONATE", "1 JAHR"],
        "parts": ["AKKU", "LADEGERÄT", "DISPLAY", "LÜFTER", "TASTATUR", "MAINBOARD", "RAM", "SSD"],
        "sparepart_types": ["STROMVERSORGUNG", "DISPLAY", "SPEICHER", "ARBEITSSPEICHER", "KÜHLUNG", "EINGABEGERÄT", "NETZWERK", "MAINBOARD"],
    },
    "pt_BR": {
        "device_types": ["COMPUTADOR DE MESA", "NOTEBOOK", "TUDO EM UM", "TABLET", "CELULAR", "IMPRESSORA"],
        "issues": ["NÃO LIGA", "DESEMPENHO LENTO", "SUPERAQUECIMENTO", "TELA AZUL", "TELA QUEBRADA", "BATERIA NÃO CARREGA", "PROBLEMA DE REDE", "RECUPERAÇÃO DE DADOS"],
        "statuses": ["RECEBIDO", "DIAGNÓSTICO", "AGUARDANDO PEÇA", "EM REPARO", "PRONTO", "CONCLUÍDO", "CANCELADO"],
        "technicians": ["ANA", "BRUNO", "CARLA", "DIEGO", "RAFAEL"],
        "warranty_periods": ["SEM GARANTIA", "7 DIAS", "14 DIAS", "30 DIAS", "3 MESES", "6 MESES", "1 ANO"],
        "parts": ["BATERIA", "CARREGADOR", "TELA", "VENTOINHA", "TECLADO", "PLACA-MÃE", "RAM", "SSD"],
        "sparepart_types": ["ENERGIA", "TELA", "ARMAZENAMENTO", "MEMÓRIA", "REFRIGERAÇÃO", "DISPOSITIVO DE ENTRADA", "REDE", "PLACA-MÃE"],
    },
    "it_IT": {
        "device_types": ["COMPUTER DESKTOP", "PORTATILE", "ALL-IN-ONE", "TABLET", "SMARTPHONE", "STAMPANTE"],
        "issues": ["NON SI ACCENDE", "PRESTAZIONI LENTE", "SURRISCALDAMENTO", "SCHERMATA BLU", "SCHERMO ROTTO", "BATTERIA NON SI CARICA", "PROBLEMA DI RETE", "RECUPERO DATI"],
        "statuses": ["RICEVUTO", "DIAGNOSI", "IN ATTESA DEL RICAMBIO", "IN RIPARAZIONE", "PRONTO", "COMPLETATO", "ANNULLATO"],
        "technicians": ["ANDREA", "CHIARA", "GIULIA", "LUCA", "MARCO"],
        "warranty_periods": ["NESSUNA GARANZIA", "7 GIORNI", "14 GIORNI", "30 GIORNI", "3 MESI", "6 MESI", "1 ANNO"],
        "parts": ["BATTERIA", "CARICATORE", "SCHERMO", "VENTOLA", "TASTIERA", "SCHEDA MADRE", "RAM", "SSD"],
        "sparepart_types": ["ALIMENTAZIONE", "SCHERMO", "ARCHIVIAZIONE", "MEMORIA", "RAFFREDDAMENTO", "DISPOSITIVO DI INPUT", "RETE", "SCHEDA MADRE"],
    },
    "nl_NL": {
        "device_types": ["DESKTOP", "LAPTOP", "ALL-IN-ONE", "TABLET", "SMARTPHONE", "PRINTER"],
        "issues": ["GEEN STROOM", "TRAGE PRESTATIES", "OVERVERHITTING", "BLAUW SCHERM", "KAPOT SCHERM", "ACCU LAADT NIET", "NETWERKPROBLEEM", "GEGEVENSHERSTEL"],
        "statuses": ["ONTVANGEN", "DIAGNOSE", "WACHT OP ONDERDEEL", "IN REPARATIE", "GEREED", "VOLTOOID", "GEANNULEERD"],
        "technicians": ["DAAN", "EMMA", "LARS", "NOA", "SOPHIE"],
        "warranty_periods": ["GEEN GARANTIE", "7 DAGEN", "14 DAGEN", "30 DAGEN", "3 MAANDEN", "6 MAANDEN", "1 JAAR"],
        "parts": ["ACCU", "OPLADER", "SCHERM", "VENTILATOR", "TOETSENBORD", "MOEDERBORD", "RAM", "SSD"],
        "sparepart_types": ["VOEDING", "BEELDSCHERM", "OPSLAG", "GEHEUGEN", "KOELING", "INVOERAPPARAAT", "NETWERK", "MOEDERBORD"],
    },
    "zh_CN": {
        "device_types": ["台式电脑", "笔记本电脑", "一体机", "平板电脑", "智能手机", "打印机"],
        "issues": ["无法开机", "运行缓慢", "过热", "蓝屏", "屏幕损坏", "电池无法充电", "网络问题", "数据恢复"],
        "statuses": ["已接收", "检测中", "等待配件", "维修中", "可取件", "已完成", "已取消"],
        "technicians": ["陈伟", "李娜", "王强", "张敏", "赵磊"],
        "warranty_periods": ["无保修", "7天", "14天", "30天", "3个月", "6个月", "1年"],
        "parts": ["电池", "充电器", "显示屏", "风扇", "键盘", "主板", "内存", "固态硬盘"],
        "sparepart_types": ["电源", "显示", "存储", "内存", "散热", "输入设备", "网络", "主板"],
    },
    "ja_JP": {
        "device_types": ["デスクトップ", "ノートパソコン", "一体型パソコン", "タブレット", "スマートフォン", "プリンター"],
        "issues": ["電源が入らない", "動作が遅い", "過熱", "ブルースクリーン", "画面破損", "バッテリーが充電できない", "ネットワーク障害", "データ復旧"],
        "statuses": ["受付済み", "診断中", "部品待ち", "修理中", "受け取り可能", "完了", "キャンセル"],
        "technicians": ["佐藤", "鈴木", "高橋", "田中", "山本"],
        "warranty_periods": ["保証なし", "7日", "14日", "30日", "3か月", "6か月", "1年"],
        "parts": ["バッテリー", "充電器", "ディスプレイ", "ファン", "キーボード", "マザーボード", "RAM", "SSD"],
        "sparepart_types": ["電源", "ディスプレイ", "ストレージ", "メモリ", "冷却", "入力デバイス", "ネットワーク", "マザーボード"],
    },
    "ko_KR": {
        "device_types": ["데스크톱", "노트북", "일체형 PC", "태블릿", "스마트폰", "프린터"],
        "issues": ["전원 불량", "성능 저하", "과열", "블루스크린", "화면 파손", "배터리 충전 불가", "네트워크 문제", "데이터 복구"],
        "statuses": ["접수", "진단 중", "부품 대기", "수리 중", "수령 가능", "완료", "취소"],
        "technicians": ["김민준", "박서준", "이서연", "정지훈", "최유진"],
        "warranty_periods": ["보증 없음", "7일", "14일", "30일", "3개월", "6개월", "1년"],
        "parts": ["배터리", "충전기", "디스플레이", "팬", "키보드", "메인보드", "RAM", "SSD"],
        "sparepart_types": ["전원", "디스플레이", "저장장치", "메모리", "냉각", "입력장치", "네트워크", "메인보드"],
    },
    "ar_SA": {
        "device_types": ["حاسوب مكتبي", "حاسوب محمول", "حاسوب متكامل", "جهاز لوحي", "هاتف ذكي", "طابعة"],
        "issues": ["لا يعمل", "أداء بطيء", "ارتفاع الحرارة", "شاشة زرقاء", "شاشة مكسورة", "البطارية لا تشحن", "مشكلة شبكة", "استعادة بيانات"],
        "statuses": ["تم الاستلام", "قيد التشخيص", "بانتظار القطعة", "قيد الإصلاح", "جاهز", "مكتمل", "ملغي"],
        "technicians": ["أحمد", "خالد", "سارة", "عمر", "نورة"],
        "warranty_periods": ["بدون ضمان", "7 أيام", "14 يوماً", "30 يوماً", "3 أشهر", "6 أشهر", "سنة واحدة"],
        "parts": ["بطارية", "شاحن", "شاشة", "مروحة", "لوحة مفاتيح", "لوحة أم", "ذاكرة RAM", "قرص SSD"],
        "sparepart_types": ["طاقة", "شاشة", "تخزين", "ذاكرة", "تبريد", "جهاز إدخال", "شبكة", "لوحة أم"],
    },
    "hi_IN": {
        "device_types": ["डेस्कटॉप", "लैपटॉप", "ऑल-इन-वन", "टैबलेट", "स्मार्टफोन", "प्रिंटर"],
        "issues": ["बिजली नहीं", "धीमा प्रदर्शन", "अधिक गर्म होना", "ब्लू स्क्रीन", "टूटी स्क्रीन", "बैटरी चार्ज नहीं होती", "नेटवर्क समस्या", "डेटा रिकवरी"],
        "statuses": ["प्राप्त", "जाँच में", "पार्ट की प्रतीक्षा", "मरम्मत में", "तैयार", "पूर्ण", "रद्द"],
        "technicians": ["अमित", "नेहा", "राहुल", "सोनिया", "विकास"],
        "warranty_periods": ["कोई वारंटी नहीं", "7 दिन", "14 दिन", "30 दिन", "3 महीने", "6 महीने", "1 वर्ष"],
        "parts": ["बैटरी", "चार्जर", "डिस्प्ले", "पंखा", "कीबोर्ड", "मदरबोर्ड", "RAM", "SSD"],
        "sparepart_types": ["पावर", "डिस्प्ले", "स्टोरेज", "मेमोरी", "कूलिंग", "इनपुट डिवाइस", "नेटवर्क", "मदरबोर्ड"],
    },
    "bn_BD": {
        "device_types": ["ডেস্কটপ", "ল্যাপটপ", "অল-ইন-ওয়ান", "ট্যাবলেট", "স্মার্টফোন", "প্রিন্টার"],
        "issues": ["বিদ্যুৎ নেই", "ধীর গতি", "অতিরিক্ত গরম", "ব্লু স্ক্রিন", "স্ক্রিন ভাঙা", "ব্যাটারি চার্জ হয় না", "নেটওয়ার্ক সমস্যা", "ডেটা পুনরুদ্ধার"],
        "statuses": ["গ্রহণ করা হয়েছে", "পরীক্ষাধীন", "যন্ত্রাংশের অপেক্ষা", "মেরামত চলছে", "প্রস্তুত", "সম্পন্ন", "বাতিল"],
        "technicians": ["আরিফ", "করিম", "নুসরাত", "রুবেল", "সুমাইয়া"],
        "warranty_periods": ["ওয়ারেন্টি নেই", "7 দিন", "14 দিন", "30 দিন", "3 মাস", "6 মাস", "1 বছর"],
        "parts": ["ব্যাটারি", "চার্জার", "ডিসপ্লে", "ফ্যান", "কীবোর্ড", "মাদারবোর্ড", "RAM", "SSD"],
        "sparepart_types": ["পাওয়ার", "ডিসপ্লে", "স্টোরেজ", "মেমোরি", "কুলিং", "ইনপুট ডিভাইস", "নেটওয়ার্ক", "মাদারবোর্ড"],
    },
    "ru_RU": {
        "device_types": ["НАСТОЛЬНЫЙ ПК", "НОУТБУК", "МОНОБЛОК", "ПЛАНШЕТ", "СМАРТФОН", "ПРИНТЕР"],
        "issues": ["НЕ ВКЛЮЧАЕТСЯ", "МЕДЛЕННАЯ РАБОТА", "ПЕРЕГРЕВ", "СИНИЙ ЭКРАН", "РАЗБИТ ЭКРАН", "БАТАРЕЯ НЕ ЗАРЯЖАЕТСЯ", "ПРОБЛЕМА С СЕТЬЮ", "ВОССТАНОВЛЕНИЕ ДАННЫХ"],
        "statuses": ["ПРИНЯТО", "ДИАГНОСТИКА", "ОЖИДАНИЕ ДЕТАЛИ", "В РЕМОНТЕ", "ГОТОВО", "ЗАВЕРШЕНО", "ОТМЕНЕНО"],
        "technicians": ["АЛЕКСЕЙ", "АННА", "ДМИТРИЙ", "ИРИНА", "МИХАИЛ"],
        "warranty_periods": ["БЕЗ ГАРАНТИИ", "7 ДНЕЙ", "14 ДНЕЙ", "30 ДНЕЙ", "3 МЕСЯЦА", "6 МЕСЯЦЕВ", "1 ГОД"],
        "parts": ["БАТАРЕЯ", "ЗАРЯДНОЕ УСТРОЙСТВО", "ДИСПЛЕЙ", "ВЕНТИЛЯТОР", "КЛАВИАТУРА", "МАТЕРИНСКАЯ ПЛАТА", "RAM", "SSD"],
        "sparepart_types": ["ПИТАНИЕ", "ДИСПЛЕЙ", "ХРАНИЛИЩЕ", "ПАМЯТЬ", "ОХЛАЖДЕНИЕ", "УСТРОЙСТВО ВВОДА", "СЕТЬ", "МАТЕРИНСКАЯ ПЛАТА"],
    },
    "tr_TR": {
        "device_types": ["MASAÜSTÜ", "DİZÜSTÜ", "HEPSİ BİR ARADA", "TABLET", "AKILLI TELEFON", "YAZICI"],
        "issues": ["GÜÇ YOK", "YAVAŞ PERFORMANS", "AŞIRI ISINMA", "MAVİ EKRAN", "KIRIK EKRAN", "BATARYA ŞARJ OLMUYOR", "AĞ SORUNU", "VERİ KURTARMA"],
        "statuses": ["TESLİM ALINDI", "TEŞHİS", "PARÇA BEKLENİYOR", "ONARIMDA", "HAZIR", "TAMAMLANDI", "İPTAL EDİLDİ"],
        "technicians": ["AHMET", "AYŞE", "EMRE", "MEHMET", "ZEYNEP"],
        "warranty_periods": ["GARANTİ YOK", "7 GÜN", "14 GÜN", "30 GÜN", "3 AY", "6 AY", "1 YIL"],
        "parts": ["BATARYA", "ŞARJ CİHAZI", "EKRAN", "FAN", "KLAVYE", "ANAKART", "RAM", "SSD"],
        "sparepart_types": ["GÜÇ", "EKRAN", "DEPOLAMA", "BELLEK", "SOĞUTMA", "GİRİŞ AYGITI", "AĞ", "ANAKART"],
    },
    "vi_VN": {
        "device_types": ["MÁY TÍNH ĐỂ BÀN", "MÁY TÍNH XÁCH TAY", "MÁY TÍNH TẤT CẢ TRONG MỘT", "MÁY TÍNH BẢNG", "ĐIỆN THOẠI", "MÁY IN"],
        "issues": ["KHÔNG LÊN NGUỒN", "HOẠT ĐỘNG CHẬM", "QUÁ NHIỆT", "MÀN HÌNH XANH", "VỠ MÀN HÌNH", "PIN KHÔNG SẠC", "LỖI MẠNG", "KHÔI PHỤC DỮ LIỆU"],
        "statuses": ["ĐÃ NHẬN", "ĐANG CHẨN ĐOÁN", "CHỜ LINH KIỆN", "ĐANG SỬA", "SẴN SÀNG", "HOÀN TẤT", "ĐÃ HỦY"],
        "technicians": ["AN", "BÌNH", "HÙNG", "LAN", "MINH"],
        "warranty_periods": ["KHÔNG BẢO HÀNH", "7 NGÀY", "14 NGÀY", "30 NGÀY", "3 THÁNG", "6 THÁNG", "1 NĂM"],
        "parts": ["PIN", "BỘ SẠC", "MÀN HÌNH", "QUẠT", "BÀN PHÍM", "BO MẠCH CHỦ", "RAM", "SSD"],
        "sparepart_types": ["NGUỒN", "MÀN HÌNH", "LƯU TRỮ", "BỘ NHỚ", "LÀM MÁT", "THIẾT BỊ NHẬP", "MẠNG", "BO MẠCH CHỦ"],
    },
    "th_TH": {
        "device_types": ["คอมพิวเตอร์ตั้งโต๊ะ", "แล็ปท็อป", "ออลอินวัน", "แท็บเล็ต", "สมาร์ตโฟน", "เครื่องพิมพ์"],
        "issues": ["เปิดไม่ติด", "ทำงานช้า", "ร้อนเกินไป", "จอฟ้า", "หน้าจอแตก", "แบตเตอรี่ไม่ชาร์จ", "ปัญหาเครือข่าย", "กู้คืนข้อมูล"],
        "statuses": ["รับเครื่องแล้ว", "กำลังตรวจสอบ", "รออะไหล่", "กำลังซ่อม", "พร้อมรับ", "เสร็จสิ้น", "ยกเลิก"],
        "technicians": ["กิตติ", "ณัฐ", "พิม", "มานะ", "สุรีย์"],
        "warranty_periods": ["ไม่มีประกัน", "7 วัน", "14 วัน", "30 วัน", "3 เดือน", "6 เดือน", "1 ปี"],
        "parts": ["แบตเตอรี่", "ที่ชาร์จ", "หน้าจอ", "พัดลม", "แป้นพิมพ์", "เมนบอร์ด", "RAM", "SSD"],
        "sparepart_types": ["ไฟเลี้ยง", "จอภาพ", "ที่เก็บข้อมูล", "หน่วยความจำ", "ระบายความร้อน", "อุปกรณ์ป้อนข้อมูล", "เครือข่าย", "เมนบอร์ด"],
    },
    "ms_MY": {
        "device_types": ["KOMPUTER MEJA", "KOMPUTER RIBA", "SEMUA DALAM SATU", "TABLET", "TELEFON PINTAR", "PENCETAK"],
        "issues": ["TIADA KUASA", "PRESTASI PERLAHAN", "TERLALU PANAS", "SKRIN BIRU", "SKRIN PECAH", "BATERI TIDAK MENGECAS", "MASALAH RANGKAIAN", "PEMULIHAN DATA"],
        "statuses": ["DITERIMA", "DIAGNOSIS", "MENUNGGU ALAT GANTI", "DALAM PEMBAIKAN", "SEDIA", "SELESAI", "DIBATALKAN"],
        "technicians": ["AHMAD", "FARAH", "HAKIM", "NADIA", "ZUL"],
        "warranty_periods": ["TIADA WARANTI", "7 HARI", "14 HARI", "30 HARI", "3 BULAN", "6 BULAN", "1 TAHUN"],
        "parts": ["BATERI", "PENGECAS", "SKRIN", "KIPAS", "PAPAN KEKUNCI", "PAPAN INDUK", "RAM", "SSD"],
        "sparepart_types": ["KUASA", "PAPARAN", "STORAN", "MEMORI", "PENYEJUKAN", "PERANTI INPUT", "RANGKAIAN", "PAPAN INDUK"],
    },
    "fil_PH": {
        "device_types": ["DESKTOP", "LAPTOP", "ALL-IN-ONE", "TABLET", "SMARTPHONE", "PRINTER"],
        "issues": ["WALANG KURYENTE", "MABAGAL NA TAKBO", "SOBRANG INIT", "BLUE SCREEN", "SIRANG SCREEN", "HINDI NAGCHA-CHARGE ANG BATERYA", "PROBLEMA SA NETWORK", "PAG-RECOVER NG DATA"],
        "statuses": ["NATANGGAP", "SINUSURI", "NAGHIHINTAY NG PIYESA", "GINAGAWA", "HANDA NA", "TAPOS NA", "KINANSELA"],
        "technicians": ["ANGELO", "JOSE", "MARIA", "PAOLO", "ROSE"],
        "warranty_periods": ["WALANG WARRANTY", "7 ARAW", "14 ARAW", "30 ARAW", "3 BUWAN", "6 BUWAN", "1 TAON"],
        "parts": ["BATERYA", "CHARGER", "DISPLAY", "BENTILADOR", "KEYBOARD", "MOTHERBOARD", "RAM", "SSD"],
        "sparepart_types": ["POWER", "DISPLAY", "STORAGE", "MEMORY", "COOLING", "INPUT DEVICE", "NETWORK", "MOTHERBOARD"],
    },
}

# Compact UI translations. Missing entries deliberately fall back to English.
UI_TEXT = {
    "en_US": { "configuration": "CONFIGURATION", "actions": "ACTIONS", "file": "FILE", "settings": "SETTINGS", "theme": "THEME", "title": "Multilingual Dummy Data Creator", "ui_language": "Interface language", "data_language": "Dummy data language", "target": "Target dummy.ini", "browse": "Browse", "generate": "Create dummy.ini", "open_folder": "Open folder", "close": "Close", "replace": "Replace existing file", "ready": "Ready", "saved": "dummy.ini created successfully", "confirm_replace": "The target file already exists. Replace it?", "error": "Error", "select_target": "Select dummy.ini target"},
    "id_ID": {"title": "Pembuat Data Dummy Multibahasa", "ui_language": "Bahasa antarmuka", "data_language": "Bahasa data dummy", "target": "Tujuan dummy.ini", "browse": "Telusuri", "generate": "Buat dummy.ini", "open_folder": "Buka folder", "close": "Tutup", "replace": "Timpa file yang ada", "ready": "Siap", "saved": "dummy.ini berhasil dibuat", "confirm_replace": "File tujuan sudah ada. Timpa file tersebut?", "error": "Kesalahan", "select_target": "Pilih tujuan dummy.ini"},
    "es_MX": {"title": "Creador de datos de prueba multilingüe", "ui_language": "Idioma de interfaz", "data_language": "Idioma de datos", "target": "Destino dummy.ini", "browse": "Examinar", "generate": "Crear dummy.ini", "open_folder": "Abrir carpeta", "close": "Cerrar", "replace": "Reemplazar archivo existente", "ready": "Listo", "saved": "dummy.ini creado correctamente", "confirm_replace": "El archivo ya existe. ¿Reemplazarlo?", "error": "Error", "select_target": "Seleccionar destino dummy.ini"},
    "fr_FR": {"title": "Créateur de données fictives multilingues", "ui_language": "Langue de l'interface", "data_language": "Langue des données", "target": "Cible dummy.ini", "browse": "Parcourir", "generate": "Créer dummy.ini", "open_folder": "Ouvrir le dossier", "close": "Fermer", "replace": "Remplacer le fichier existant", "ready": "Prêt", "saved": "dummy.ini créé avec succès", "confirm_replace": "Le fichier existe déjà. Le remplacer ?", "error": "Erreur", "select_target": "Sélectionner la cible dummy.ini"},
    "de_DE": {
        "configuration": "KONFIGURATION",
        "actions": "AKTIONEN",
        "file": "DATEI",
        "settings": "EINSTELLUNGEN",
        "theme": "DESIGN",
        "title": "Mehrsprachiger Dummy-Daten-Ersteller",
        "ui_language": "Oberflächensprache",
        "data_language": "Sprache der Dummy-Daten",
        "target": "Zieldatei dummy.ini",
        "browse": "Durchsuchen",
        "generate": "dummy.ini erstellen",
        "open_folder": "Ordner öffnen",
        "close": "Schließen",
        "replace": "Vorhandene Datei ersetzen",
        "ready": "Bereit",
        "saved": "dummy.ini wurde erfolgreich erstellt",
        "confirm_replace": "Die Zieldatei existiert bereits. Soll sie ersetzt werden?",
        "error": "Fehler",
        "select_target": "Ziel für dummy.ini auswählen",
        "ini_files": "INI-Dateien",
        "all_files": "Alle Dateien",
    },
    "pt_BR": {"title": "Criador de dados fictícios multilíngue", "ui_language": "Idioma da interface", "data_language": "Idioma dos dados", "target": "Destino dummy.ini", "browse": "Procurar", "generate": "Criar dummy.ini", "open_folder": "Abrir pasta", "close": "Fechar", "replace": "Substituir arquivo existente", "ready": "Pronto", "saved": "dummy.ini criado com sucesso", "confirm_replace": "O arquivo já existe. Substituir?", "error": "Erro", "select_target": "Selecionar destino dummy.ini"},
    "zh_CN": {"title": "多语言示例数据生成器", "ui_language": "界面语言", "data_language": "示例数据语言", "target": "dummy.ini 目标", "browse": "浏览", "generate": "创建 dummy.ini", "open_folder": "打开文件夹", "close": "关闭", "replace": "替换现有文件", "ready": "就绪", "saved": "dummy.ini 创建成功", "confirm_replace": "目标文件已存在，是否替换？", "error": "错误", "select_target": "选择 dummy.ini 目标"},
    "ja_JP": {"title": "多言語ダミーデータ作成ツール", "ui_language": "表示言語", "data_language": "ダミーデータ言語", "target": "dummy.ini の保存先", "browse": "参照", "generate": "dummy.ini を作成", "open_folder": "フォルダーを開く", "close": "閉じる", "replace": "既存ファイルを置換", "ready": "準備完了", "saved": "dummy.ini を作成しました", "confirm_replace": "ファイルは既に存在します。置換しますか？", "error": "エラー", "select_target": "dummy.ini の保存先を選択"},
    "ko_KR": {"title": "다국어 더미 데이터 생성기", "ui_language": "인터페이스 언어", "data_language": "더미 데이터 언어", "target": "dummy.ini 대상", "browse": "찾아보기", "generate": "dummy.ini 생성", "open_folder": "폴더 열기", "close": "닫기", "replace": "기존 파일 교체", "ready": "준비됨", "saved": "dummy.ini 생성 완료", "confirm_replace": "대상 파일이 이미 있습니다. 교체할까요?", "error": "오류", "select_target": "dummy.ini 대상 선택"},
    "ar_SA": {"title": "منشئ بيانات تجريبية متعدد اللغات", "ui_language": "لغة الواجهة", "data_language": "لغة البيانات", "target": "مسار dummy.ini", "browse": "استعراض", "generate": "إنشاء dummy.ini", "open_folder": "فتح المجلد", "close": "إغلاق", "replace": "استبدال الملف الموجود", "ready": "جاهز", "saved": "تم إنشاء dummy.ini بنجاح", "confirm_replace": "الملف موجود بالفعل. هل تريد استبداله؟", "error": "خطأ", "select_target": "اختر مسار dummy.ini"},
    "ru_RU": {"title": "Создание многоязычных тестовых данных", "ui_language": "Язык интерфейса", "data_language": "Язык данных", "target": "Файл dummy.ini", "browse": "Обзор", "generate": "Создать dummy.ini", "open_folder": "Открыть папку", "close": "Закрыть", "replace": "Заменить существующий файл", "ready": "Готово", "saved": "dummy.ini успешно создан", "confirm_replace": "Файл уже существует. Заменить?", "error": "Ошибка", "select_target": "Выберите файл dummy.ini"},
    "tr_TR": {"title": "Çok Dilli Örnek Veri Oluşturucu", "ui_language": "Arayüz dili", "data_language": "Örnek veri dili", "target": "dummy.ini hedefi", "browse": "Gözat", "generate": "dummy.ini oluştur", "open_folder": "Klasörü aç", "close": "Kapat", "replace": "Mevcut dosyayı değiştir", "ready": "Hazır", "saved": "dummy.ini başarıyla oluşturuldu", "confirm_replace": "Hedef dosya zaten var. Değiştirilsin mi?", "error": "Hata", "select_target": "dummy.ini hedefini seç"},
    "vi_VN": {"title": "Trình tạo dữ liệu mẫu đa ngôn ngữ", "ui_language": "Ngôn ngữ giao diện", "data_language": "Ngôn ngữ dữ liệu", "target": "Đích dummy.ini", "browse": "Duyệt", "generate": "Tạo dummy.ini", "open_folder": "Mở thư mục", "close": "Đóng", "replace": "Thay thế tệp hiện có", "ready": "Sẵn sàng", "saved": "Đã tạo dummy.ini thành công", "confirm_replace": "Tệp đích đã tồn tại. Thay thế?", "error": "Lỗi", "select_target": "Chọn đích dummy.ini"},
    "th_TH": {"title": "เครื่องมือสร้างข้อมูลตัวอย่างหลายภาษา", "ui_language": "ภาษาอินเทอร์เฟซ", "data_language": "ภาษาข้อมูลตัวอย่าง", "target": "ปลายทาง dummy.ini", "browse": "เรียกดู", "generate": "สร้าง dummy.ini", "open_folder": "เปิดโฟลเดอร์", "close": "ปิด", "replace": "แทนที่ไฟล์เดิม", "ready": "พร้อม", "saved": "สร้าง dummy.ini สำเร็จ", "confirm_replace": "มีไฟล์ปลายทางอยู่แล้ว ต้องการแทนที่หรือไม่?", "error": "ข้อผิดพลาด", "select_target": "เลือกปลายทาง dummy.ini"},
    "ms_MY": {"title": "Pencipta Data Dummy Berbilang Bahasa", "ui_language": "Bahasa antara muka", "data_language": "Bahasa data dummy", "target": "Destinasi dummy.ini", "browse": "Semak imbas", "generate": "Cipta dummy.ini", "open_folder": "Buka folder", "close": "Tutup", "replace": "Gantikan fail sedia ada", "ready": "Sedia", "saved": "dummy.ini berjaya dicipta", "confirm_replace": "Fail sasaran sudah wujud. Gantikannya?", "error": "Ralat", "select_target": "Pilih destinasi dummy.ini"},
    "fil_PH": {"title": "Multilingual Dummy Data Creator", "ui_language": "Wika ng interface", "data_language": "Wika ng dummy data", "target": "Target na dummy.ini", "browse": "Mag-browse", "generate": "Gumawa ng dummy.ini", "open_folder": "Buksan ang folder", "close": "Isara", "replace": "Palitan ang kasalukuyang file", "ready": "Handa", "saved": "Matagumpay na nagawa ang dummy.ini", "confirm_replace": "Mayroon nang target file. Palitan ito?", "error": "Error", "select_target": "Piliin ang target na dummy.ini"},
}

THEME_LABELS = {
    "de_DE": {
        "light": "HELL",
        "dark": "DUNKEL",
        "blue": "BLAU",
        "purple": "LILA",
        "red": "ROT",
        "orange": "ORANGE",
        "classic": "KLASSISCH",
    },
}


def normalize_locale(locale: str | None) -> str:
    value = str(locale or "").strip()
    return value if value in LANGUAGE_OPTIONS else DEFAULT_LANGUAGE_CODE


def get_dummy_data(locale: str) -> Dict[str, List[str]]:
    locale = normalize_locale(locale)
    localized = LOCALIZED_DATA.get(locale, LOCALIZED_DATA["en_US"])
    result: Dict[str, List[str]] = {}
    for category in CATEGORY_ORDER:
        source = localized.get(category, COMMON_DATA.get(category, LOCALIZED_DATA["en_US"].get(category, [])))
        # Preserve script/case for readability; the main app normalizes values when loading.
        result[category] = [str(value).strip() for value in source if str(value).strip()]
    return result


def write_dummy_ini(
    target: Path,
    locale: str,
    replace: bool = False,
) -> Path:
    target = target.expanduser().resolve()
    if target.exists() and not replace:
        raise FileExistsError(str(target))

    source_data = get_dummy_data(locale)
    data: Dict[str, List[str]] = {}
    for category in CATEGORY_ORDER:
        values = source_data.get(category, [])
        data[category] = [str(value).strip() for value in values if str(value).strip()]

    parser = configparser.ConfigParser(interpolation=None)
    parser.optionxform = str
    parser["META"] = {
        "FORMAT_VERSION": "1",
        "LANGUAGE_CODE": normalize_locale(locale),
        "LANGUAGE_NAME": LANGUAGE_OPTIONS[normalize_locale(locale)],
        "GENERATOR": "Dummy_Creator.py",
    }
    for category in CATEGORY_ORDER:
        parser[category.upper()] = {"VALUES": "|".join(data[category])}

    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(
        prefix=target.name + ".",
        suffix=".tmp",
        dir=str(target.parent),
        text=True,
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as stream:
            parser.write(stream)
        os.replace(temporary, target)
    except Exception:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    return target


def open_folder(path: Path) -> None:
    folder = path.expanduser().resolve()
    if folder.is_file() or folder.suffix:
        folder = folder.parent
    folder.mkdir(parents=True, exist_ok=True)
    if sys.platform.startswith("win"):
        os.startfile(str(folder))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(folder)])
    else:
        subprocess.Popen(["xdg-open", str(folder)])

def get_runtime_app_dir() -> Path:
    """Return the directory containing the source file or compiled executable."""
    if getattr(sys, "frozen", False) or "__compiled__" in globals():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_app_file_path(filename: str) -> Path:
    return get_runtime_app_dir() / filename


REGISTRY_PATH = r"Software\Computer Service Manager 4.1\Dummy Creator"
REGISTRY_DEFAULTS = {
    "UiLanguage": DEFAULT_LANGUAGE_CODE,
    "DataLanguage": DEFAULT_LANGUAGE_CODE,
    "TargetPath": "",
    "ReplaceExisting": 1,
    "ThemeCode": DEFAULT_THEME_CODE,
    "WindowGeometry": b"",
    "WindowState": b"",
}


def read_registry_settings() -> dict[str, object]:
    """Load every persisted GUI setting from HKCU on Windows."""
    settings = dict(REGISTRY_DEFAULTS)
    if winreg is None:
        return settings

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH) as key:
            for name in settings:
                try:
                    value, _value_type = winreg.QueryValueEx(key, name)
                    settings[name] = value
                except FileNotFoundError:
                    continue
    except OSError:
        pass
    return settings


def write_registry_settings(values: dict[str, object]) -> None:
    """Write GUI settings to HKCU without creating a settings file."""
    if winreg is None:
        return

    try:
        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            REGISTRY_PATH,
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            for name, value in values.items():
                if isinstance(value, bool):
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, int(value))
                elif isinstance(value, int):
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
                elif isinstance(value, (bytes, bytearray)):
                    winreg.SetValueEx(key, name, 0, winreg.REG_BINARY, bytes(value))
                else:
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
    except OSError:
        # Registry restrictions must not prevent the generator from running.
        pass


def write_registry_setting(name: str, value: object) -> None:
    write_registry_settings({name: value})


def resolve_theme(
    theme_code: str | None = None,
    settings: dict[str, object] | None = None,
) -> str:
    code = str(theme_code or "").strip().lower()
    if code in THEME_OPTIONS:
        return code

    stored = str(
        (settings or {}).get("ThemeCode", DEFAULT_THEME_CODE)
    ).strip().lower()
    return stored if stored in THEME_OPTIONS else DEFAULT_THEME_CODE


class DummyCreatorWindow(QMainWindow):
    """Compact generator-only window with registry-backed settings."""

    def __init__(
        self,
        output: Path,
        ui_locale: str,
        data_locale: str,
        theme_code: str,
        replace_existing: bool,
        registry_settings: dict[str, object],
    ):
        super().__init__()
        self.ui_locale = normalize_locale(ui_locale)
        self.data_locale = normalize_locale(data_locale)
        self.theme_code = resolve_theme(theme_code, registry_settings)
        self.theme_actions: dict[str, QAction] = {}
        self._saved_geometry = registry_settings.get("WindowGeometry", b"")
        self._saved_state = registry_settings.get("WindowState", b"")

        self.setWindowTitle(APP_TITLE)
        self.resize(680, 430)
        self.setMinimumSize(560, 360)
        self.apply_app_icon()
        self.init_ui(output, replace_existing)
        self.apply_theme(self.theme_code, persist=False)
        self.apply_texts()
        self.restore_saved_window_state()

    def apply_app_icon(self) -> None:
        icon_path = get_app_file_path("app.ico")
        try:
            if icon_path.exists():
                icon = QIcon(str(icon_path))
                self.setWindowIcon(icon)
                app = QApplication.instance()
                if app is not None:
                    app.setWindowIcon(icon)
        except Exception:
            pass

    def init_ui(self, output: Path, replace_existing: bool) -> None:
        self.create_menu_bar()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(12)

        self.configuration_group = QGroupBox()
        form = QFormLayout(self.configuration_group)
        form.setContentsMargins(14, 18, 14, 14)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)
        form.setFieldGrowthPolicy(
            QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow
        )

        self.ui_combo = QComboBox()
        self.ui_combo.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Fixed,
        )
        self.populate_language_combo(self.ui_combo, self.ui_locale)
        self.ui_combo.currentIndexChanged.connect(self.on_ui_language_changed)
        self.ui_language_label = QLabel()
        form.addRow(self.ui_language_label, self.ui_combo)

        self.data_combo = QComboBox()
        self.data_combo.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Fixed,
        )
        self.populate_language_combo(self.data_combo, self.data_locale)
        self.data_combo.currentIndexChanged.connect(self.on_data_language_changed)
        self.data_language_label = QLabel()
        form.addRow(self.data_language_label, self.data_combo)

        target_widget = QWidget()
        target_layout = QHBoxLayout(target_widget)
        target_layout.setContentsMargins(0, 0, 0, 0)
        target_layout.setSpacing(8)
        self.target_entry = QLineEdit(str(output))
        self.target_entry.editingFinished.connect(self.persist_target_path)
        self.browse_button = QPushButton()
        self.browse_button.setObjectName("secondaryButton")
        self.browse_button.clicked.connect(self.browse_target)
        target_layout.addWidget(self.target_entry, 1)
        target_layout.addWidget(self.browse_button)
        self.target_label = QLabel()
        form.addRow(self.target_label, target_widget)

        self.replace_check = QCheckBox()
        self.replace_check.setChecked(bool(replace_existing))
        self.replace_check.toggled.connect(
            lambda checked: write_registry_setting("ReplaceExisting", checked)
        )
        form.addRow(QLabel(""), self.replace_check)
        main_layout.addWidget(self.configuration_group)

        self.actions_group = QGroupBox()
        actions_layout = QHBoxLayout(self.actions_group)
        actions_layout.setContentsMargins(14, 18, 14, 14)
        actions_layout.setSpacing(8)

        self.generate_button = QPushButton()
        self.generate_button.setObjectName("primaryButton")
        self.generate_button.clicked.connect(self.generate)
        self.generate_button.setDefault(True)

        self.open_button = QPushButton()
        self.open_button.setObjectName("secondaryButton")
        self.open_button.clicked.connect(self.open_target_folder)

        self.close_button = QPushButton()
        self.close_button.setObjectName("neutralButton")
        self.close_button.clicked.connect(self.close)

        actions_layout.addWidget(self.generate_button, 1)
        actions_layout.addWidget(self.open_button)
        actions_layout.addWidget(self.close_button)
        main_layout.addWidget(self.actions_group)
        main_layout.addStretch(1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def create_menu_bar(self) -> None:
        menubar = self.menuBar()

        self.file_menu = menubar.addMenu("FILE")
        self.generate_action = QAction("CREATE DUMMY.INI", self)
        self.generate_action.triggered.connect(self.generate)
        self.file_menu.addAction(self.generate_action)

        self.open_folder_action = QAction("OPEN FOLDER", self)
        self.open_folder_action.triggered.connect(self.open_target_folder)
        self.file_menu.addAction(self.open_folder_action)
        self.file_menu.addSeparator()

        self.exit_action = QAction("EXIT", self)
        self.exit_action.triggered.connect(self.close)
        self.file_menu.addAction(self.exit_action)

        self.settings_menu = menubar.addMenu("SETTINGS")
        self.theme_menu = self.settings_menu.addMenu("THEME")
        self.theme_action_group = QActionGroup(self)
        self.theme_action_group.setExclusive(True)

        for code in THEME_ORDER:
            action = QAction(THEME_OPTIONS[code]["label"], self)
            action.setCheckable(True)
            action.setData(code)
            action.setChecked(code == self.theme_code)
            action.triggered.connect(
                lambda checked=False, selected=code: self.apply_theme(selected)
            )
            self.theme_action_group.addAction(action)
            self.theme_menu.addAction(action)
            self.theme_actions[code] = action

    @staticmethod
    def populate_language_combo(combo: QComboBox, current_locale: str) -> None:
        combo.blockSignals(True)
        combo.clear()
        for code, label in LANGUAGE_OPTIONS.items():
            combo.addItem(label, code)
        index = combo.findData(normalize_locale(current_locale))
        combo.setCurrentIndex(index if index >= 0 else 0)
        combo.blockSignals(False)

    def tr(self, key: str) -> str:
        locale = normalize_locale(self.ui_locale)
        return UI_TEXT.get(locale, {}).get(
            key,
            UI_TEXT["en_US"].get(key, key),
        )

    def set_status(self, message: str) -> None:
        self.status_bar.showMessage(str(message))

    def current_combo_locale(self, combo: QComboBox) -> str:
        return normalize_locale(combo.currentData())

    def on_ui_language_changed(self, _index: int = -1) -> None:
        self.ui_locale = self.current_combo_locale(self.ui_combo)
        write_registry_setting("UiLanguage", self.ui_locale)
        self.apply_texts()

    def on_data_language_changed(self, _index: int = -1) -> None:
        self.data_locale = self.current_combo_locale(self.data_combo)
        write_registry_setting("DataLanguage", self.data_locale)
        self.set_status(self.tr("ready"))

    def apply_texts(self) -> None:
        self.setWindowTitle(self.tr("title"))
        self.configuration_group.setTitle(self.tr("configuration"))
        self.actions_group.setTitle(self.tr("actions"))
        self.ui_language_label.setText(self.tr("ui_language"))
        self.data_language_label.setText(self.tr("data_language"))
        self.target_label.setText(self.tr("target"))
        self.browse_button.setText(self.tr("browse"))
        self.replace_check.setText(self.tr("replace"))
        self.generate_button.setText(self.tr("generate"))
        self.open_button.setText(self.tr("open_folder"))
        self.close_button.setText(self.tr("close"))

        self.file_menu.setTitle(self.tr("file"))
        self.settings_menu.setTitle(self.tr("settings"))
        self.theme_menu.setTitle(self.tr("theme"))
        theme_labels = THEME_LABELS.get(normalize_locale(self.ui_locale), {})
        for code, action in self.theme_actions.items():
            action.setText(
                theme_labels.get(code, THEME_OPTIONS[code]["label"])
            )
        self.generate_action.setText(self.tr("generate"))
        self.open_folder_action.setText(self.tr("open_folder"))
        self.exit_action.setText(self.tr("close"))

        rtl = self.ui_locale == "ar_SA"
        direction = (
            Qt.LayoutDirection.RightToLeft
            if rtl
            else Qt.LayoutDirection.LeftToRight
        )
        self.setLayoutDirection(direction)
        self.set_status(self.tr("ready"))

    def apply_theme(self, theme_code: str, persist: bool = True) -> None:
        code = str(theme_code or DEFAULT_THEME_CODE).strip().lower()
        if code not in THEME_OPTIONS:
            code = DEFAULT_THEME_CODE
        app = QApplication.instance()
        if app is not None:
            apply_application_theme(app, code)
        self.theme_code = code
        for item_code, action in self.theme_actions.items():
            action.setChecked(item_code == code)
        if persist:
            write_registry_setting("ThemeCode", code)
            self.set_status(
                f"{self.tr('theme')}: {THEME_OPTIONS[code]['label']}"
            )

    def browse_target(self) -> None:
        current = Path(
            self.target_entry.text().strip() or "dummy.ini"
        ).expanduser()
        selected, _selected_filter = QFileDialog.getSaveFileName(
            self,
            self.tr("select_target"),
            str(current),
            f"{self.tr('ini_files')} (*.ini);;"
            f"{self.tr('all_files')} (*.*)",
        )
        if selected:
            self.target_entry.setText(selected)
            self.persist_target_path()

    def normalized_target(self) -> Path:
        target = Path(
            self.target_entry.text().strip() or "dummy.ini"
        ).expanduser()
        if target.suffix.lower() != ".ini":
            target = target.with_suffix(".ini")
        return target

    def persist_target_path(self) -> None:
        target = self.normalized_target()
        self.target_entry.setText(str(target))
        write_registry_setting("TargetPath", str(target))

    def save_all_settings(self) -> None:
        write_registry_settings(
            {
                "UiLanguage": self.ui_locale,
                "DataLanguage": self.data_locale,
                "TargetPath": str(self.normalized_target()),
                "ReplaceExisting": self.replace_check.isChecked(),
                "ThemeCode": self.theme_code,
                "WindowGeometry": bytes(self.saveGeometry()),
                "WindowState": bytes(self.saveState()),
            }
        )

    def restore_saved_window_state(self) -> None:
        try:
            if (
                isinstance(self._saved_geometry, (bytes, bytearray))
                and self._saved_geometry
            ):
                self.restoreGeometry(
                    QByteArray(bytes(self._saved_geometry))
                )
            if (
                isinstance(self._saved_state, (bytes, bytearray))
                and self._saved_state
            ):
                self.restoreState(QByteArray(bytes(self._saved_state)))
        except Exception:
            pass

    def generate(self) -> None:
        try:
            target = self.normalized_target()
            self.target_entry.setText(str(target))

            replace = self.replace_check.isChecked()
            if target.exists() and not replace:
                answer = QMessageBox.question(
                    self,
                    self.tr("title"),
                    self.tr("confirm_replace"),
                    QMessageBox.StandardButton.Yes
                    | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
                replace = True

            written = write_dummy_ini(
                target,
                self.data_locale,
                replace=replace,
            )
            self.target_entry.setText(str(written))
            self.save_all_settings()
            self.set_status(f"{self.tr('saved')}: {written}")
            QMessageBox.information(
                self,
                self.tr("title"),
                f"{self.tr('saved')}\n\n{written}",
            )
        except Exception as exc:
            self.set_status(str(exc))
            QMessageBox.critical(self, self.tr("error"), str(exc))

    def open_target_folder(self) -> None:
        try:
            open_folder(self.normalized_target())
        except Exception as exc:
            QMessageBox.critical(self, self.tr("error"), str(exc))

    def closeEvent(self, event) -> None:  # type: ignore[no-untyped-def]
        self.save_all_settings()
        super().closeEvent(event)


def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        default="",
        help="Target path for dummy.ini",
    )
    parser.add_argument(
        "--language",
        default="",
        help="Initial UI and data language; defaults to the registry value",
    )
    parser.add_argument(
        "--theme",
        default="",
        help="Initial theme code; defaults to the registry value",
    )
    parser.add_argument(
        "--write-only",
        action="store_true",
        help="Create the file without opening the graphical interface",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Replace an existing target in write-only mode",
    )
    return parser.parse_args(argv)


def default_output_path() -> Path:
    return get_runtime_app_dir() / "dummy.ini"


def main(argv: List[str] | None = None) -> int:
    args = parse_args(argv)
    settings = read_registry_settings()

    data_locale = normalize_locale(
        args.language
        or str(settings.get("DataLanguage", DEFAULT_LANGUAGE_CODE))
    )
    stored_target = str(settings.get("TargetPath", "")).strip()
    target_value = args.output or stored_target
    output = (
        Path(target_value).expanduser()
        if target_value
        else default_output_path()
    )

    if args.write_only:
        try:
            written = write_dummy_ini(
                output,
                data_locale,
                replace=args.force,
            )
        except FileExistsError:
            print(f"Target file already exists: {output}", file=sys.stderr)
            return 2
        except Exception as exc:
            print(f"Could not create dummy.ini: {exc}", file=sys.stderr)
            return 1
        print(written)
        return 0

    if QT_IMPORT_ERROR is not None:
        raise SystemExit(
            "PyQt6 is required to open the graphical interface: "
            f"{QT_IMPORT_ERROR}. Use --write-only for command-line generation."
        )

    ui_locale = normalize_locale(
        args.language
        or str(settings.get("UiLanguage", DEFAULT_LANGUAGE_CODE))
    )
    theme_code = resolve_theme(args.theme, settings)
    replace_existing = bool(
        int(settings.get("ReplaceExisting", 1) or 0)
    )

    app = QApplication.instance() or QApplication(sys.argv[:1])
    apply_application_theme(app, theme_code)
    window = DummyCreatorWindow(
        output,
        ui_locale,
        data_locale,
        theme_code,
        replace_existing,
        settings,
    )
    window.show()
    return int(app.exec())


if __name__ == "__main__":
    raise SystemExit(main())
