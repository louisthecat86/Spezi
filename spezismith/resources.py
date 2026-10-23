"""Ressourcen- und Konfigurations-Helfer (PyInstaller-tauglich, portabel)."""

from __future__ import annotations

import json
import os
import sys

# Ressourcen, die in die EXE eingebettet werden
TEMPLATE_DE = os.path.join("templates", "vorlage_de.xlsx")
TEMPLATE_EN = os.path.join("templates", "vorlage_en.xlsx")
LOGO_PNG = os.path.join("assets", "logo.png")


def resource_path(relative: str) -> str:
    """Pfad zu einer mitgepackten Ressource (im EXE-Container oder Repo)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        candidate = os.path.join(base, relative)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(app_root(), relative)


def app_root() -> str:
    """Repo-/Projektwurzel im Entwicklungsmodus."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def config_path() -> str:
    """
    Konfiguration liegt portabel direkt neben der EXE
    (bzw. im Projektverzeichnis im Entwicklungsmodus).
    """
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), "spezismith_config.json")
    return os.path.join(app_root(), "spezismith_config.json")


def load_config() -> dict:
    try:
        with open(config_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_config(cfg: dict) -> None:
    try:
        with open(config_path(), "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass  # Konfiguration ist optional - nie deshalb abstuerzen
