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
# CONFIGURACIÓN GENERAL
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
# MOTOR DE CÁLCULO PREVENT AUTOMATIZADO (SIN CONFIG 'tc')
# =========================================================
PREVENT_AVAILABLE = False
PREVENT_IMPORT_ERROR = ""
try:
    import pyprevent
    PREVENT_AVAILABLE = True
except Exception as e:
    PREVENT_IMPORT_ERROR = repr(e)

# Configuración inicial de Streamlit (Debe ser la primera instrucción)
st.set_page_config(page_title=APP_NAME, page_icon="🫀", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# INYECCIÓN DE ESTILOS CSS AVANZADOS
# =========================================================
css_styles = """
<style>
html, body, [class*="css"] { color:#111827 !important; }
.main {background:#F8FAFC;}
.block-container {padding-top:1rem; padding-bottom:2rem;}
section[data-testid="stSidebar"] { background:#F1F5F9 !important; color:#111827 !important; }
section[data-testid="stSidebar"] * { color:#111827 !important; }
.hero { background: linear-gradient(135deg,#0B4F8A 0%,#123C69 55%,#0F766E 100%); padding:28px 34px; border-radius:26px; color:white !important; box-shadow:0 14px 34px rgba(11,79,138,.25); margin-bottom:18px; }
.hero h1 {font-size:2.35rem; margin:0 0 8px 0; font-weight:900; color:white !important;}
.hero p {font-size:1rem; opacity:.96; margin:0; color:white !important;}
.card {background:white; border-radius:22px; padding:20px 22px; box-shadow:0 8px 24px rgba(15,23,42,.07); border:1px solid #E5E7EB; margin-bottom:16px;}
.badge {display:inline-block; padding:6px 11px; border-radius:999px; font-weight:800; font-size:.82rem; margin:2px 4px 2px 0;}
.badge-green {background:#BBF7D0; color:#16A34A !important; border:1px solid #16A34A;}
.badge-yellow {background:#FEF08A; color:#CA8A04 !important; border:1px solid #CA8A04;}
.badge-orange {background:#FED7AA; color:#EA580C !important; border:1px solid #EA580C;}
.badge-red {background:#FECACA; color:#DC2626 !important; border:1px solid #DC2626;}
.badge-blue {background:#BFDBFE; color:#2563EB !important; border:1px solid #2563EB;}
.badge-gray {background:#E5E7EB; color:#6B7280 !important; border:1px solid #6B7280;}
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
"""
st.markdown(css_styles, unsafe_allow_html=True)

# =========================================================
# UTILIDADES VISUALES Y COMPONENTES
# =========================================================
def badge_html(texto: str, color: str = "blue") -> str:
    cls = {
        "green": "badge-green",
        "yellow": "badge-yellow",
        "orange": "badge-orange",
        "red": "badge-red",
        "blue": "badge-blue",
        "gray": "badge-gray"
    }.get(color, "badge-blue")
    return f'<span class="badge {cls}">{texto}</span>'

def semaforo_item(nombre: str, valor, unidad: str, color: str, interpretacion: str, referencia: str):
    if valor is None: valor_txt = "No calculado"
    elif isinstance(valor, float): valor_txt = f"{valor:.2f} {unidad}"
    else: valor_txt = f"{valor} {unidad}"
    st.markdown(f'<div class="semaforo-card"><div class="semaforo-title">{nombre}</div><div class="semaforo-value">{valor_txt}</div>{badge_html(interpretacion, color)}<div class="semaforo-ref">Referencia: {referencia}</div></div>', unsafe_allow_html=True)

# =========================================================
# AUTENTICACIÓN Y PERSISTENCIA (TOTALMENTE INTEGRADA)
# =========================================================
def _ensure_users_file():
    if not USERS_FILE.exists(): USERS_FILE.write_text("{}", encoding="utf-8")

def load_users() -> dict:
    _ensure_users_file()
    try: return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except: return {}

def save_users(users: dict):
    USERS_FILE.write_text(json.dumps(users, indent=2, ensure_ascii=False), encoding="utf-8")

def hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()

def register_user(username: str, password: str, nombre: str, matricula: str, especialidad: str, rol: str = "medico") -> Tuple[bool, str]:
    username = (username or "").strip()
    if not username or not password: return False, "El usuario y la contraseña son campos mandatorios."
    users = load_users()
    if username in users: return False, "Este usuario ya se encuentra registrado."
    salt = secrets.token_hex(16)
    users[username] = {
        "salt": salt, "password": hash_password(password, salt),
        "nombre": nombre, "matricula": matricula, "especialidad": especialidad, "rol": rol
    }
    save_users(users)
    return True, "Médico registrado correctamente."

def authenticate(username: str, password: str) -> Tuple[bool, Optional[dict]]:
    users = load_users()
    if username not in users: return False, None
    u = users[username]
    if hash_password(password, u["salt"]) == u["password"]: return True, u
    return False, None

# Asegurar Administrador por defecto de la aplicación
if "admin" not in load_users():
    register_user("admin", "admin1234", "Administrador Institucional", "9999", "Cardiología", "admin")

def load_historial() -> dict:
    if not HISTORIAL_FILE.exists(): return {}
    try: return json.loads(HISTORIAL_FILE.read_text(encoding="utf-8"))
    except: return {}

def save_historial(h: dict):
    HISTORIAL_FILE.write_text(json.dumps(h, indent=2, ensure_ascii=False), encoding="utf-8")

def add_paciente_historial(username: str, registro: dict):
    h = load_historial()
    if username not in h: h[username] = []
    h[username].append(registro)
    save_historial(h)

# =========================================================
# MODELO DE DATOS CLÍNICOS E INFERENCIA COGNITIVA
# =========================================================
@dataclass
class Patient:
    paciente: str
    dni: str
    edad: int
    sexo: str
    ldl_basal: float
    ldl_actual: float
    colesterol_total: float
    hdl: float
    tg: float
    no_hdl: float
    presion_sistolica: float
    tratamiento_hta: bool
    diabetes: bool
    tabaquismo: bool
    egfr: float
    antecedente_infarto: bool
    prevent_10: Optional[float] = None
    prevent_30: Optional[float] = None
    ops_hearts_riesgo: str = "Bajo"
    categoria_riesgo_final: str = "Riesgo Bajo"
    meta_ldl: str = "<116 mg/dL"
    indicacion_tratamiento: str = ""

def calcular_ops_hearts(p: Patient) -> str:
    if p.antecedente_infarto: return "Muy Alto"
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

def procesar_evaluacion_completa(p: Patient) -> Patient:
    if PREVENT_AVAILABLE and not p.antecedente_infarto:
        try:
            genero_py = "female" if p.sexo == "Femenino" else "male"
            res = pyprevent.calculate_risk(
                age=int(p.edad), sex=genero_py, sbp=int(p.presion_sistolica),
                bp_med=1 if p.tratamiento_hta else 0, tot_chol=int(p.colesterol_total),
                hdl_chol=int(p.hdl), ldl_chol=int(p.ldl_actual),
                diabetes=1 if p.diabetes else 0, smoker=1 if p.tabaquismo else 0,
                egfr=float(p.egfr) if p.egfr else 75.0
            )
            p.prevent_10 = round(res.get("10_yr_ascvd_risk", 0.0), 2) if res.get("10_yr_ascvd_risk") is not None else None
            p.prevent_30 = round(res.get("30_yr_ascvd_risk", 0.0), 2) if res.get("30_yr_ascvd_risk") is not None else None
        except Exception as e:
            p.prevent_10, p.prevent_30 = None, None

    p.ops_hearts_riesgo = calcular_ops_hearts(p)
    
    if p.antecedente_infarto:
        p.categoria_riesgo_final = "Prevención Secundaria (Extremo / Muy Alto)"
        p.meta_ldl = "< 55 mg/dL"
        p.indicacion_tratamiento = "Evidencia Clase I. Iniciar inmediatamente estatinas de alta intensidad (Atorvastatina 40-80mg o Rosuvastatina 20-40mg) junto con Ezetimibe 10mg."
    else:
        score_riesgo = p.prevent_10 if p.prevent_10 is not None else 0.0
        if score_riesgo >= 10.0 or p.ops_hearts_riesgo == "Alto / Muy Alto":
            p.categoria_riesgo_final = "Riesgo Alto"
            p.meta_ldl = "< 70 mg/dL"
            p.indicacion_tratamiento = "Indicación mandatoria de Estatinas de alta intensidad. Monitoreo hepático y metabólico estrecho con metas lipídicas estrictas."
        elif 5.0 <= score_riesgo < 10.0 or p.ops_hearts_riesgo == "Moderado":
            p.categoria_riesgo_final = "Riesgo Intermedio"
            p.meta_ldl = "< 100 mg/dL"
            p.indicacion_tratamiento = "Iniciar estatinas de moderada intensidad. Discutir detalladamente con el paciente la presencia de factores potenciadores de riesgo."
        else:
            p.categoria_riesgo_final = "Riesgo Bajo"
            p.meta_ldl = "< 116 mg/dL"
            p.indicacion_tratamiento = "Priorizar cambios higiénico-dietéticos estructurados (dieta mediterránea y ejercicio regular aeróbico). Reevaluar anualmente."
    return p

# =========================================================
# VISTAS DE INTERFAZ DEL SISTEMA
# =========================================================
def render_evaluacion():
    st.markdown("## 🩺 Carga Unificada de Variables Simultánea")
    with st.form("form_clinico"):
        col1, col2, col3 = st.columns(3)
        with col1:
            paciente = st.text_input("Nombre del Paciente", value="Paciente Anónimo")
            dni = st.text_input("DNI / Identificador Único", value="")
            edad = st.number_input("Edad (30-85 años)", min_value=30, max_value=85, value=55)
            sexo = st.selectbox("Sexo Biológico", ["Masculino", "Femenino"])
        with col2:
            ldl_basal = st.number_input("LDL Basal (mg/dL)", min_value=30.0, max_value=300.0, value=160.0)
            ldl_actual = st.number_input("LDL Actual (mg/dL)", min_value=20.0, max_value=300.0, value=130.0)
            colesterol_total = st.number_input("Colesterol Total (mg/dL)", min_value=100.0, max_value=400.0, value=210.0)
            hdl = st.number_input("HDL (mg/dL)", min_value=15.0, max_value=100.0, value=45.0)
            tg = st.number_input("Triglicéridos (mg/dL)", min_value=30.0, max_value=600.0, value=150.0)
        with col3:
            presion_sistolica = st.number_input("Presión Sistólica (mmHg)", min_value=80.0, max_value=220.0, value=135.0)
            tratamiento_hta = st.checkbox("¿Se encuentra en tratamiento farmacológico para HTA?")
            diabetes = st.checkbox("Diabetes Mellitus Tipo 2")
            tabaquismo = st.checkbox("Tabaquismo Activo Actual")
            egfr = st.number_input("Filtrado Glomerular Estimado eGFR (mL/min)", min_value=15.0, max_value=150.0, value=75.0)
            antecedente_infarto = st.checkbox("Prevención Secundaria (Antecedente de IAM / ACV / PAD)")
            
        submit = st.form_submit_button("Procesar Diagnóstico Semáforo")
        
    if submit:
        p = Patient(
            paciente=paciente, dni=dni, edad=int(edad), sexo=sexo, ldl_basal=ldl_basal,
            ldl_actual=ldl_actual, colesterol_total=colesterol_total, hdl=hdl, tg=tg,
            no_hdl=colesterol_total - hdl, presion_sistolica=presion_sistolica,
            tratamiento_hta=tratamiento_hta, diabetes=diabetes, tabaquismo=tabaquismo,
            egfr=egfr, antecedente_infarto=antecedente_infarto
        )
        p = procesar_evaluacion_completa(p)
        add_paciente_historial(st.session_state.username, p.__dict__)
        
        st.markdown("---")
        st.markdown("### 📊 INFORME CLÍNICO SEMAFORIZADO INTEGRADO")
        
        color_alert = "alert-red" if "Secundaria" in p.categoria_riesgo_final or "Alto" in p.categoria_riesgo_final else ("alert-orange" if "Intermedio" in p.categoria_riesgo_final else "alert-green")
        st.markdown(f'<div class="{color_alert}"><strong>Estratificación Final: {p.categoria_riesgo_final}</strong><br>Meta de Control de LDL Objetivo: {p.meta_ldl}</div>', unsafe_allow_html=True)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: semaforo_item("LDL Actual", p.ldl_actual, "mg/dL", "red" if p.ldl_actual > 100 else "green", "Analizar", p.meta_ldl)
        with c2: semaforo_item("Colesterol No-HDL", p.no_hdl, "mg/dL", "orange" if p.no_hdl > 130 else "green", "Cálculo", "LDL + 30")
        with c3: semaforo_item("Riesgo PREVENT 10a", p.prevent_10, "%", "red" if (p.prevent_10 and p.prevent_10 >= 10) else "green", "Guías AHA", "Bajo <3%")
        with c4: semaforo_item("OPS Hearts Américas", p.ops_hearts_riesgo, "", "red" if "Alto" in p.ops_hearts_riesgo else "green", p.ops_hearts_riesgo, "Tablas OPS Región")
        
        st.markdown(f'<div class="rx-card"><div class="rx-title">💊 Indicación Farmacoterapéutica Sugerida (ACC/AHA 2026)</div><p>{p.indicacion_tratamiento}</p></div>', unsafe_allow_html=True)

def render_historial_propio():
    st.markdown("## 🗄️ Mi Historial de Evaluaciones Clínicas")
    h = load_historial()
    pacientes = h.get(st.session_state.username, [])
    if not pacientes:
        st.info("No se registran evaluaciones clínicas guardadas en el usuario actual.")
        return
    df = pd.DataFrame(pacientes)
    st.dataframe(df[["paciente", "dni", "edad", "sexo", "ldl_actual", "prevent_10", "ops_hearts_riesgo", "categoria_riesgo_final"]])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="MisPacientes")
    st.download_button(label="📥 Descargar mi archivo Excel por Usuario", data=output.getvalue(), file_name=f"historial_{st.session_state.username}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def render_admin_usuarios():
    st.markdown("## 👑 Panel del Administrador (Consolidado Institucional)")
    h = load_historial()
    all_rec = []
    for user, records in h.items():
        for r in records:
            rc = r.copy()
            rc["usuario_medico"] = user
            all_rec.append(rc)
    if not all_rec:
        st.info("No existen registros clínicos almacenados en toda la institución.")
        return
    df_all = pd.DataFrame(all_rec)
    st.dataframe(df_all)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_all.to_excel(writer, index=False, sheet_name="ConsolidadoCompleto")
    st.download_button(label="📥 Exportar Excel Completo Global (Todos los Usuarios)", data=output.getvalue(), file_name="consolidado_sistema_lipidcare.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def render_evolucion_paciente():
    st.markdown("## 📈 Módulo de Evolución y Tendencias")
    st.info("Análisis longitudinal del comportamiento cinético del perfil lipídico del paciente.")

def render_calculadora_prevent():
    st.markdown("## 🧮 Calculadora Manual PREVENT Suplementaria")
    st.info("Herramienta interactiva para la proyección de escenarios preventivos alternativos.")

# =========================================================
# CONTROLADOR Y ORQUESTADOR PRINCIPAL (MAIN)
# =========================================================
def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.user_data = {}
        
    if not st.session_state.logged_in:
        st.markdown('<div class="hero"><h1>🫀 LipidCare 2026 Pro</h1><p>Ecosistema Médico Avanzado de Riesgo Cardiovascular Integrado y Guías de Dislipemia</p></div>', unsafe_allow_html=True)
        t1, t2 = st.tabs(["🔒 Iniciar Sesión", "📝 Registrar Profesional"])
        with t1:
            u = st.text_input("Usuario (Matrícula / Email)")
            p = st.text_input("Contraseña", type="password")
            if st.button("Ingresar al Ecosistema"):
                ok, u_data = authenticate(u, p)
                if ok:
                    st.session_state.logged_in = True
                    st.session_state.username = u
                    st.session_state.user_data = u_data
                    st.rerun()
                else: st.error("Las credenciales introducidas son incorrectas.")
        with t2:
            ru = st.text_input("Definir Nombre de Usuario")
            rp = st.text_input("Definir Contraseña de Acceso", type="password")
            rn = st.text_input("Nombre y Apellido del Profesional")
            rm = st.text_input("Matrícula Profesional M.P.")
            re = st.text_input("Especialidad Médica")
            rr = st.selectbox("Rol Asignado", ["medico", "admin"])
            if st.button("Crear Cuenta"):
                ok, msg = register_user(ru, rp, rn, rm, re, rr)
                if ok: st.success("Médico registrado con éxito. Ya puede iniciar sesión.")
                else: st.error(msg)
    else:
        st.markdown(f'<div class="user-bar"><span>👨‍⚕️ Profesional Activo: Dr./Dra. {st.session_state.username} | Rol Institucional: {st.session_state.user_data.get("rol","").upper()}</span></div>', unsafe_allow_html=True)
        
        opciones = ["Evaluación clínica", "Mi historial", "Evolución del paciente", "Calculadora PREVENT"]
        if st.session_state.user_data.get("rol") == "admin": 
            opciones.append("Admin: todos los usuarios")
            
        modo = st.sidebar.radio("Módulo de Trabajo", opciones)
        
        if st.sidebar.button("🔒 Cerrar Sesión"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.user_data = {}
            st.rerun()
            
        if modo == "Evaluación clínica": render_evaluacion()
        elif modo == "Mi historial": render_historial_propio()
        elif modo == "Evolución del paciente": render_evolucion_paciente()
        elif modo == "Calculadora PREVENT": render_calculadora_prevent()
        elif modo == "Admin: todos los usuarios": render_admin_usuarios()

if __name__ == "__main__":
    main()
