from __future__ import annotations
import streamlit as st
import pandas as pd
import io
import os
import json
import hashlib
import secrets
import textwrap
import zlib
import math
from dataclasses import dataclass
from typing import Optional, Dict, List, Tuple
from datetime import date, datetime
from pathlib import Path
import streamlit.components.v1 as components

# =========================================================
# CONFIGURACIÓN
# =========================================================
APP_NAME = "LipidCare 2026 Pro"
AUTOR_APP = "Ricardo Daniel Olano, Especialista en Cardiología y en Hipertensión Arterial"
PREVENT_URL = "https://professional.heart.org/en/guidelines-and-statements/prevent-calculator"
PAHO_HEARTS_URL = "https://www.paho.org/cardioapp/web/"

DATA_DIR = Path(os.environ.get("LIPIDCARE_DATA_DIR", ".lipidcare_data"))
DATA_DIR.mkdir(exist_ok=True)
USERS_FILE = DATA_DIR / "users.json"
HISTORIAL_FILE = DATA_DIR / "historial.json"

DEFAULT_ADMIN_USER = "admin"
DEFAULT_ADMIN_PASS = "admin1234"

# =========================================================
# MOTOR DE CÁLCULO PREVENT (MÓDULO INTEGRADO SIN PARÁMETROS 'tc')
# =========================================================
PREVENT_AVAILABLE = False
PREVENT_IMPORT_ERROR = ""
try:
    import pyprevent
    PREVENT_AVAILABLE = True
except Exception as e:
    PREVENT_IMPORT_ERROR = repr(e)

# =========================================================
# MOTOR EXCEL
# =========================================================
EXCEL_ENGINE = "openpyxl"

st.set_page_config(page_title=APP_NAME, page_icon="🫀", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# ESTILOS INSTITUCIONALES SEMAFORIZADOS
# =========================================================
st.markdown("""
<style>
html, body, [class*="css"] { color:#111827 !important; }
.main {background:#F8FAFC;}
.block-container {padding-top:1rem; padding-bottom:2rem;}
section[data-testid="stSidebar"] { background:#F1F5F9 !important; }
.hero { background: linear-gradient(135deg,#0B4F8A 0%,#123C69 55%,#0F766E 100%); padding:28px 34px; border-radius:26px; color:white !important; box-shadow:0 14px 34px rgba(11,79,138,.25); margin-bottom:18px; }
.hero h1 {font-size:2.35rem; margin:0 0 8px 0; font-weight:900; color:white !important;}
.hero p {font-size:1rem; opacity:.96; margin:0; color:white !important;}
.card {background:white; border-radius:22px; padding:20px 22px; box-shadow:0 8px 24px rgba(15,23,42,.07); border:1px solid #E5E7EB; margin-bottom:16px;}
.badge {display:inline-block; padding:6px 11px; border-radius:999px; font-weight:800; font-size:.82rem; margin:2px 4px 2px 0;}
.badge-green {background:#BBF7D0; color:#16A34A !important;}
.badge-yellow {background:#FEF08A; color:#CA8A04 !important;}
.badge-orange {background:#FED7AA; color:#EA580C !important;}
.badge-red {background:#FECACA; color:#DC2626 !important;}
.badge-blue {background:#BFDBFE; color:#2563EB !important;}
.badge-gray {background:#E5E7EB; color:#6B7280 !important;}
.alert-red {border-left:6px solid #B91C1C; background:#FEF2F2; padding:14px 16px; border-radius:14px; margin-bottom:15px;}
.alert-green {border-left:6px solid #0F766E; background:#ECFDF5; padding:14px 16px; border-radius:14px; margin-bottom:15px;}
.alert-orange {border-left:6px solid #EA580C; background:#FFF7ED; padding:14px 16px; border-radius:14px; margin-bottom:15px;}
.semaforo-card{background:#FFFFFF; border:1px solid #CBD5E1; border-radius:18px; padding:14px 15px; box-shadow:0 4px 14px rgba(15,23,42,.05); min-height:116px; margin-bottom:10px;}
.semaforo-title{font-size:.88rem;color:#334155 !important;font-weight:800;margin-bottom:4px;}
.semaforo-value{font-size:1.28rem;color:#111827 !important;font-weight:900;margin-bottom:6px;}
.semaforo-ref{font-size:.78rem;color:#475569 !important;}
.user-bar {background:#0F172A; color:white !important; padding:10px 18px; border-radius:14px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center; font-weight:800;}
.rx-card {background:#FFFFFF; border:2px solid #0B4F8A; border-radius:18px; padding:18px 20px; margin-bottom:14px;}
.rx-title {color:#0B4F8A !important; font-weight:900; font-size:1.1rem; margin-bottom:10px;}
.rx-drug {background:#EFF6FF; border-left:5px solid #0B4F8A; padding:10px 14px; border-radius:10px; margin:6px 0; color:#0B4F8A !important; font-weight:800;}
</style>
""", unsafe_allow_html=True)

# =========================================================
# UTILIDADES DE TEXTO Y COMPONENTES VISUALES
# =========================================================
def badge_html(texto: str, color: str = "blue") -> str:
    cls = {"green":"badge-green", "yellow":"badge-yellow", "orange":"badge-orange",
           "red":"badge-red", "blue":"badge-blue", "gray":"badge-gray"}.get(color, "badge-blue")
    return f'<span class="badge {cls}">{texto}</span>'

def semaforo_item(nombre: str, valor, unidad: str, color: str, interpretacion: str, referencia: str):
    if valor is None: valor_txt = "No calculado"
    elif isinstance(valor, float): valor_txt = f"{valor:.2f} {unidad}"
    else: valor_txt = f"{valor} {unidad}"
    st.markdown(f'<div class="semaforo-card"><div class="semaforo-title">{nombre}</div><div class="semaforo-value">{valor_txt}</div>{badge_html(interpretacion, color)}<div class="semaforo-ref">Referencia: {referencia}</div></div>', unsafe_allow_html=True)

# =========================================================
# SEGURIDAD Y BASE DE DATOS LOCAL
