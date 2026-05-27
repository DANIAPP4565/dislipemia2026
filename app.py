from __future__ import annotations
import streamlit as st
import pandas as pd
import io
import os
import json
import hashlib
import secrets
import textwrap
from dataclasses import dataclass, asdict
from typing import Optional, Dict, List, Tuple
from datetime import datetime
from pathlib import Path

# =========================================================
# CONFIGURACIÓN E INICIALIZACIÓN
# =========================================================
APP_NAME = "LipidCare 2026 Pro"
AUTOR_APP = "Ricardo Daniel Olano, Especialista en Cardiología y en Hipertensión Arterial"

DATA_DIR = Path(os.environ.get("LIPIDCARE_DATA_DIR", ".lipidcare_data"))
DATA_DIR.mkdir(exist_ok=True)
USERS_FILE = DATA_DIR / "users.json"
HISTORIAL_FILE = DATA_DIR / "historial.json"

# =========================================================
# MOTOR DE CÁLCULO PREVENT (CORREGIDO)
# =========================================================
PREVENT_AVAILABLE = False
PREVENT_IMPORT_ERROR = ""
try:
    import pyprevent
    PREVENT_AVAILABLE = True
except Exception as e:
    PREVENT_IMPORT_ERROR = repr(e)

# =========================================================
# MODELO DE DATOS CLÍNICOS
# =========================================================
@dataclass
class PacienteData:
    id_registro: str
    usuario: str
    fecha: str
    nombre: str
    edad: int
    sexo: str  # 'Masculino' o 'Femenino'
    presion_sistolica: int
    tratamiento_hta: bool
    colesterol_total: int
    hdl: int
    ldl_actual: int
    diabetes: bool
    tabaquismo: bool
    egfr: int
    antecedente_infarto: bool  # Define Prevención Primaria vs Secundaria
    
    # Resultados calculados
    prevent_10: Optional[float] = None
    prevent_30: Optional[float] = None
    ops_hearts_riesgo: str = "No determinado"
    categoria_riesgo_final: str = "Bajo"
    meta_ldl: str = "Ver indicaciones"
    indicacion_tratamiento: str = ""

# =========================================================
# GESTIÓN DE SEGURIDAD Y PERSISTENCIA (JSON/EXCEL)
# =========================================================
def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

def cargar_usuarios() -> dict:
    if not USERS_FILE.exists():
        admin_default = {"admin": {"password": hash_password("admin1234"), "rol": "admin"}}
        USERS_FILE.write_text(json.dumps(admin_default, indent=4))
        return admin_default
    try:
        return json.loads(USERS_FILE.read_text())
    except:
        return {}

def guardar_usuario(username: str, password_raw: str, rol: str = "medico") -> bool:
    usuarios = cargar_usuarios()
    if username in usuarios:
        return False
    usuarios[username] = {"password": hash_password(password_raw), "rol": rol}
    USERS_FILE.write_text(json.dumps(usuarios, indent=4))
    return True

def cargar_historial() -> list:
    if not HISTORIAL_FILE.exists():
        return []
    try:
        return json.loads(HISTORIAL_FILE.read_text())
    except:
        return []

def guardar_registro(p: PacienteData):
    historial = cargar_historial()
    historial.append(asdict(p))
    HISTORIAL_FILE.write_text(json.dumps(historial, indent=4), encoding="utf-8")

# =========================================================
# LÓGICA CLÍNICA INTERNA (OPS HEARTS + CORRELACIÓN AHA 2026)
# =========================================================
def calcular_ops_hearts(p: PacienteData) -> str:
    """Estratificación simplificada según Tablas de Riesgo OPS HEARTS para las Américas."""
    if p.antecedente_infarto:
        return "Muy Alto"
    
    puntos = 0
    if p.edad >= 60: puntos += 2
    elif p.edad >= 50: puntos += 1
    if p.tabaquismo: puntos += 2
    if p.diabetes: puntos += 2
    if p.presion_sistolica >= 160: puntos += 2
    elif p.presion_sistolica >= 140: puntos += 1
    if p.colesterol_total >= 240: puntos += 2
    elif p.colesterol_total >= 200: puntos += 1
    
    if puntos >= 6: return "Alto / Muy Alto"
    elif puntos >= 3: return "Moderado"
    return "Bajo"

def procesar_evaluacion_completa(p: PacienteData) -> PacienteData:
    # 1. Ejecutar Score PREVENT mediante el módulo integrado (Solución al TypeError)
    if PREVENT_AVAILABLE:
        try:
            genero_py = "female" if p.sexo == "Femenino" else "male"
            
            # Pasamos estrictamente los parámetros validados numéricos sin abreviaciones ambiguas
            res = pyprevent.calculate_risk(
                age=int(p.edad),
                sex=genero_py,
                sbp=int(p.presion_sistolica),
                bp_med=1 if p.tratamiento_hta else 0,
                tot_chol=int(p.colesterol_total),
                hdl_chol=int(p.hdl),
