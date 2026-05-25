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
from dataclasses import dataclass
from typing import Optional, Dict, List, Tuple
from datetime import date, datetime
from pathlib import Path
import streamlit.components.v1 as components

# =========================================================
# Configuracion
# =========================================================
APP_NAME = "LipidCare 2026 Pro"
AUTOR_APP = "Ricardo Daniel Olano, Especialista en Cardiologia y en Hipertension Arterial"
PREVENT_URL = "https://professional.heart.org/en/guidelines-and-statements/prevent-calculator"

DATA_DIR = Path(os.environ.get("LIPIDCARE_DATA_DIR", ".lipidcare_data"))
DATA_DIR.mkdir(exist_ok=True)
USERS_FILE = DATA_DIR / "users.json"
HISTORIAL_FILE = DATA_DIR / "historial.json"

# Admin por defecto (puede registrarse desde la pantalla de registro tambien)
DEFAULT_ADMIN_USER = "admin"
DEFAULT_ADMIN_PASS = "admin1234"

# =========================================================
# Motor PDF (interno con soporte de rectangulos coloreados)
# =========================================================
PDF_ENGINE = "interno_sin_dependencias"
PDF_IMPORT_ERROR = ""
try:
    from fpdf import FPDF  # opcional
    PDF_ENGINE_FPDF_AVAILABLE = True
except Exception as e:
    FPDF = None
    PDF_ENGINE_FPDF_AVAILABLE = False
    PDF_IMPORT_ERROR = repr(e)

# =========================================================
# Motor Excel
# =========================================================
EXCEL_ENGINE = None
EXCEL_IMPORT_ERROR = ""
try:
    import openpyxl  # noqa: F401
    EXCEL_ENGINE = "openpyxl"
except Exception as e1:
    try:
        import xlsxwriter  # noqa: F401
        EXCEL_ENGINE = "xlsxwriter"
    except Exception as e2:
        EXCEL_IMPORT_ERROR = f"openpyxl: {repr(e1)} | xlsxwriter: {repr(e2)}"

st.set_page_config(page_title=APP_NAME, page_icon="🫀", layout="wide", initial_sidebar_state="expanded")

# =========================================================
# Estilos
# =========================================================
st.markdown('''
<style>
html, body, [class*="css"] { color:#111827 !important; }
.main {background:#F8FAFC;}
.block-container {padding-top:1rem; padding-bottom:2rem;}
section[data-testid="stSidebar"] { background:#F1F5F9 !important; color:#111827 !important; }
section[data-testid="stSidebar"] * { color:#111827 !important; }
section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea, section[data-testid="stSidebar"] select { background:#FFFFFF !important; color:#111827 !important; }
.hero { background: linear-gradient(135deg,#0B4F8A 0%,#123C69 55%,#0F766E 100%); padding:28px 34px; border-radius:26px; color:white !important; box-shadow:0 14px 34px rgba(11,79,138,.25); margin-bottom:18px; }
.hero h1 {font-size:2.35rem; margin:0 0 8px 0; font-weight:900; color:white !important;}
.hero p {font-size:1rem; opacity:.96; margin:0; color:white !important;}
.card {background:white; border-radius:22px; padding:20px 22px; box-shadow:0 8px 24px rgba(15,23,42,.07); border:1px solid #E5E7EB; margin-bottom:16px; color:#111827 !important;}
.badge {display:inline-block; padding:6px 11px; border-radius:999px; font-weight:800; font-size:.82rem; margin:2px 4px 2px 0;}
.badge-green {background:#BBF7D0; color:#111827 !important; border:1px solid #16A34A;}
.badge-yellow {background:#FEF08A; color:#111827 !important; border:1px solid #CA8A04;}
.badge-orange {background:#FED7AA; color:#111827 !important; border:1px solid #EA580C;}
.badge-red {background:#FECACA; color:#111827 !important; border:1px solid #DC2626;}
.badge-blue {background:#BFDBFE; color:#111827 !important; border:1px solid #2563EB;}
.badge-gray {background:#E5E7EB; color:#111827 !important; border:1px solid #6B7280;}
.alert-red {border-left:6px solid #B91C1C; background:#FEF2F2; color:#111827 !important; padding:14px 16px; border-radius:14px;}
.alert-green {border-left:6px solid #0F766E; background:#ECFDF5; color:#111827 !important; padding:14px 16px; border-radius:14px;}
.alert-orange {border-left:6px solid #EA580C; background:#FFF7ED; color:#111827 !important; padding:14px 16px; border-radius:14px;}
.alert-blue {border-left:6px solid #2563EB; background:#EFF6FF; color:#111827 !important; padding:14px 16px; border-radius:14px;}
.semaforo-card{background:#FFFFFF; border:1px solid #CBD5E1; border-radius:18px; padding:14px 15px; box-shadow:0 4px 14px rgba(15,23,42,.05); min-height:116px; margin-bottom:10px; color:#111827 !important;}
.summary-card{background:#FFFFFF; border:1px solid #CBD5E1; border-radius:20px; padding:18px 18px; box-shadow:0 6px 18px rgba(15,23,42,.07); min-height:132px; margin-bottom:12px; color:#111827 !important;}
.summary-title{font-size:.82rem;color:#334155 !important;font-weight:900;text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;}
.summary-value{font-size:1.45rem;color:#111827 !important;font-weight:900;line-height:1.15;margin-bottom:8px;}
.summary-caption{font-size:.84rem;color:#475569 !important;font-weight:700;margin-top:6px;}
.semaforo-title{font-size:.88rem;color:#334155 !important;font-weight:800;margin-bottom:4px;}
.semaforo-value{font-size:1.28rem;color:#111827 !important;font-weight:900;margin-bottom:6px;}
.semaforo-ref{font-size:.78rem;color:#475569 !important;}
.stButton button, .stDownloadButton button, a[data-testid="stLinkButton"] {background:#0B4F8A !important; color:white !important; border:1px solid #0B4F8A !important; font-weight:800 !important;}
.stButton button:hover, .stDownloadButton button:hover, a[data-testid="stLinkButton"]:hover {background:#123C69 !important; color:white !important;}
.footer {color:#64748B !important; font-size:.82rem; margin-top:18px;}
.user-bar {background:#0F172A; color:white !important; padding:10px 18px; border-radius:14px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center; font-weight:800;}
.user-bar * { color:white !important; }
.rx-card {background:#FFFFFF; border:2px solid #0B4F8A; border-radius:18px; padding:18px 20px; margin-bottom:14px; box-shadow:0 6px 18px rgba(11,79,138,.12);}
.rx-title {color:#0B4F8A !important; font-weight:900; font-size:1.1rem; margin-bottom:10px;}
.rx-drug {background:#EFF6FF; border-left:5px solid #0B4F8A; padding:10px 14px; border-radius:10px; margin:6px 0; color:#0B4F8A !important; font-weight:800;}
</style>
''', unsafe_allow_html=True)

# =========================================================
# Utilidades de texto
# =========================================================
def safe_text(txt) -> str:
    if txt is None:
        return ""
    txt = str(txt)
    replacements = {
        "≥": ">=", "≤": "<=", "–": "-", "—": "-", "“": '"', "”": '"', "’": "'",
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N", "°": " grados ", "²": "2", "ü": "u", "Ü": "U", "ı": "i"
    }
    for a, b in replacements.items():
        txt = txt.replace(a, b)
    return txt

def badge_html(texto: str, color: str = "blue") -> str:
    cls = {"green":"badge-green", "yellow":"badge-yellow", "orange":"badge-orange",
           "red":"badge-red", "blue":"badge-blue", "gray":"badge-gray"}.get(color, "badge-blue")
    return f'<span class="badge {cls}">{texto}</span>'

def resumen_card(titulo: str, valor: str, badge_texto: str = "", color: str = "blue", caption: str = ""):
    badge = badge_html(badge_texto, color) if badge_texto else ""
    st.markdown(f'''<div class="summary-card"><div class="summary-title">{titulo}</div><div class="summary-value">{valor}</div>{badge}<div class="summary-caption">{caption}</div></div>''', unsafe_allow_html=True)

def semaforo_item(nombre: str, valor, unidad: str, color: str, interpretacion: str, referencia: str):
    if valor is None: valor_txt = "No informado"
    elif isinstance(valor, float): valor_txt = f"{valor:.1f} {unidad}"
    else: valor_txt = f"{valor} {unidad}"
    st.markdown(f'''<div class="semaforo-card"><div class="semaforo-title">{nombre}</div><div class="semaforo-value">{valor_txt}</div>{badge_html(interpretacion, color)}<div class="semaforo-ref">Referencia: {referencia}</div></div>''', unsafe_allow_html=True)

# =========================================================
# Autenticacion
# =========================================================
def _ensure_users_file():
    if not USERS_FILE.exists():
        USERS_FILE.write_text("{}", encoding="utf-8")

def load_users() -> Dict[str, dict]:
    _ensure_users_file()
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_users(users: Dict[str, dict]):
    USERS_FILE.write_text(json.dumps(users, indent=2, ensure_ascii=False), encoding="utf-8")

def hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()

def register_user(username: str, password: str, nombre: str, matricula: str, especialidad: str, rol: str = "medico") -> Tuple[bool, str]:
    username = (username or "").strip()
    if not username or not password:
        return False, "Usuario y contrasena son obligatorios."
    if len(password) < 6:
        return False, "La contrasena debe tener al menos 6 caracteres."
    users = load_users()
    if username in users:
        return False, "El usuario ya existe."
    salt = secrets.token_hex(16)
    users[username] = {
        "salt": salt,
        "password": hash_password(password, salt),
        "nombre": (nombre or "").strip(),
        "matricula": (matricula or "").strip(),
        "especialidad": (especialidad or "").strip(),
        "rol": rol,
        "creado": datetime.now().isoformat(timespec="seconds")
    }
    save_users(users)
    return True, "Usuario registrado correctamente. Ya puede iniciar sesion."

def authenticate(username: str, password: str) -> Tuple[bool, Optional[dict]]:
    users = load_users()
    if username not in users:
        return False, None
    u = users[username]
    if hash_password(password, u["salt"]) == u["password"]:
        return True, u
    return False, None

def ensure_default_admin():
    users = load_users()
    if DEFAULT_ADMIN_USER not in users:
        register_user(DEFAULT_ADMIN_USER, DEFAULT_ADMIN_PASS,
                      "Administrador del sistema", "ADMIN", "Administracion", rol="admin")

ensure_default_admin()

# =========================================================
# Historial persistente por usuario
# =========================================================
def load_historial() -> Dict[str, list]:
    if not HISTORIAL_FILE.exists():
        return {}
    try:
        return json.loads(HISTORIAL_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_historial(historial: Dict[str, list]):
    HISTORIAL_FILE.write_text(json.dumps(historial, indent=2, ensure_ascii=False, default=str), encoding="utf-8")

def add_paciente_historial(username: str, registro: dict):
    h = load_historial()
    if username not in h:
        h[username] = []
    h[username].append(registro)
    save_historial(h)

def borrar_historial_usuario(username: str):
    h = load_historial()
    h[username] = []
    save_historial(h)

def borrar_paciente(username: str, idx: int):
    h = load_historial()
    if username in h and 0 <= idx < len(h[username]):
        h[username].pop(idx)
        save_historial(h)

# =========================================================
# Modelo de datos
# =========================================================
@dataclass
class Patient:
    paciente: str
    dni: str
    medico: str
    matricula: str
    edad: int
    sexo: str
    ldl_basal: float
    ldl_actual: float
    hdl: float
    tg: float
    colesterol_total: float
    no_hdl: float
    lpa_valor: Optional[float]
    lpa_unidad: str
    apob: Optional[float]
    diabetes: bool
    ckd: bool
    egfr: Optional[float]
    presion_sistolica: float
    hta: bool
    tratamiento_hta: bool
    tabaquismo: bool
    inflamacion_cronica: bool
    antecedente_familiar: bool
    menopausia_precoz: bool
    preeclampsia: bool
    ascvd: bool
    iam: bool
    acv: bool
    pad: bool
    revascularizacion: bool
    fh_sospecha: bool
    cac: Optional[int]
    prevent_10: Optional[float]
    prevent_30: Optional[float]
    estatina: str
    dosis_estatina: str
    ezetimibe: bool
    pcsk9: bool
    inclisiran: bool
    bempedoico: bool
    intolerancia_sams: bool
    observaciones: str

# =========================================================
# Logica clinica (mejorada para prevencion primaria)
# =========================================================
def pct_reduccion(ldl_basal: float, ldl_actual: float) -> Optional[float]:
    if ldl_basal and ldl_basal > 0 and ldl_actual >= 0:
        return round((ldl_basal - ldl_actual) / ldl_basal * 100, 1)
    return None

def lpa_alta(valor: Optional[float], unidad: str) -> bool:
    if valor is None:
        return False
    return valor >= 50 if unidad == "mg/dL" else valor >= 125

def parse_meta_ldl(meta_txt: str) -> Optional[float]:
    if "55" in meta_txt: return 55
    if "70" in meta_txt: return 70
    if "100" in meta_txt: return 100
    if "116" in meta_txt: return 116
    return None

def clasificar_prevent(riesgo: Optional[float]) -> str:
    if riesgo is None:
        return "No informado"
    if riesgo < 5: return "Bajo"
    if riesgo < 7.5: return "Limitrofe"
    if riesgo < 20: return "Intermedio"
    return "Alto"

def determinar_perfil(p: Patient) -> Dict[str, str]:
    """Definicion clara de perfiles segun ACC/AHA + ESC adaptado."""
    eventos = sum([p.iam, p.acv, p.pad, p.revascularizacion])
    biologico_alto = p.diabetes or p.ckd or p.fh_sospecha
    if p.ascvd:
        perfil = "Prevencion secundaria"
        riesgo = "Muy alto riesgo" if eventos >= 1 or biologico_alto else "Alto riesgo secundario"
    else:
        # Prevencion primaria - clarificada
        if p.ldl_actual >= 190 or p.fh_sospecha:
            perfil = "Prevencion primaria"
            riesgo = "Alto riesgo primario"  # LDL severamente elevado o FH
        elif p.diabetes and 40 <= p.edad <= 75:
            perfil = "Prevencion primaria"
            # diabetes con potenciadores adicionales = muy alto riesgo
            potenciadores_dm = sum([p.ckd, p.hta, p.tabaquismo, p.antecedente_familiar,
                                     lpa_alta(p.lpa_valor, p.lpa_unidad),
                                     (p.cac is not None and p.cac >= 100)])
            riesgo = "Alto riesgo primario" if potenciadores_dm >= 1 or p.edad >= 50 else "Riesgo intermedio"
        elif p.ckd:
            perfil = "Prevencion primaria"
            riesgo = "Alto riesgo primario"
        else:
            perfil = "Prevencion primaria"
            cat = clasificar_prevent(p.prevent_10)
            if cat == "Alto": riesgo = "Alto riesgo primario"
            elif cat == "Intermedio": riesgo = "Riesgo intermedio"
            elif cat == "Limitrofe": riesgo = "Riesgo limitrofe"
            elif cat == "Bajo": riesgo = "Bajo riesgo"
            else: riesgo = "Riesgo no estimado"
    return {"perfil": perfil, "riesgo": riesgo}

def metas_lipidicas(p: Patient) -> Dict[str, str]:
    info = determinar_perfil(p)
    riesgo_lower = info["riesgo"].lower()
    if info["perfil"] == "Prevencion secundaria" and "muy alto" in riesgo_lower:
        return {"ldl":"<55 mg/dL", "no_hdl":"<85 mg/dL", "reduccion":">=50%"}
    if info["perfil"] == "Prevencion secundaria":
        return {"ldl":"<70 mg/dL", "no_hdl":"<100 mg/dL", "reduccion":">=50%"}
    if "alto" in riesgo_lower:
        return {"ldl":"<70 mg/dL", "no_hdl":"<100 mg/dL", "reduccion":">=50%"}
    if "intermedio" in riesgo_lower:
        return {"ldl":"<100 mg/dL", "no_hdl":"<130 mg/dL", "reduccion":"30-49% o mayor si potenciadores"}
    if "limitrofe" in riesgo_lower:
        return {"ldl":"<116 mg/dL", "no_hdl":"<145 mg/dL", "reduccion":"segun potenciadores/CAC"}
    return {"ldl":"<116 mg/dL", "no_hdl":"<145 mg/dL", "reduccion":"estilo de vida"}

def estado_meta(p: Patient) -> Dict[str, object]:
    metas = metas_lipidicas(p)
    meta_ldl = parse_meta_ldl(metas["ldl"])
    reduccion = pct_reduccion(p.ldl_basal, p.ldl_actual)
    cumple_ldl = meta_ldl is not None and p.ldl_actual < meta_ldl
    requiere_50 = metas["reduccion"] == ">=50%"
    cumple_red = True if not requiere_50 else (reduccion is not None and reduccion >= 50)
    if cumple_ldl and cumple_red:
        return {"texto":"En meta", "color":"green", "reduccion": reduccion if reduccion is not None else "No calculable"}
    if meta_ldl is not None and p.ldl_actual <= meta_ldl + 20:
        return {"texto":"Cerca de meta", "color":"orange", "reduccion": reduccion if reduccion is not None else "No calculable"}
    return {"texto":"Fuera de meta", "color":"red", "reduccion": reduccion if reduccion is not None else "No calculable"}

# === Clasificadores semaforicos ===
def clasificar_ldl_vs_meta(ldl: float, meta_txt: str):
    meta = parse_meta_ldl(meta_txt)
    if meta is None: return "gray", "Individualizar", "segun perfil clinico"
    if ldl < meta: return "green", "En meta", f"<{meta:.0f} mg/dL"
    if ldl <= meta + 20: return "orange", "Cerca de meta", f"<{meta:.0f} mg/dL"
    return "red", "Fuera de meta", f"<{meta:.0f} mg/dL"

def clasificar_no_hdl(no_hdl: float, meta_txt: str):
    if "85" in meta_txt: meta = 85
    elif "100" in meta_txt: meta = 100
    elif "130" in meta_txt: meta = 130
    elif "145" in meta_txt: meta = 145
    else: return "gray", "Individualizar", "segun perfil clinico"
    if no_hdl < meta: return "green", "En meta", f"<{meta} mg/dL"
    if no_hdl <= meta + 20: return "orange", "Cerca de meta", f"<{meta} mg/dL"
    return "red", "Fuera de meta", f"<{meta} mg/dL"

def clasificar_tg(tg: float):
    if tg < 150: return "green", "Normal", "<150 mg/dL"
    if tg < 175: return "yellow", "Limitrofe", "150-174 mg/dL"
    if tg < 500: return "orange", "Elevado", "175-499 mg/dL"
    return "red", "Muy elevado", ">=500 mg/dL"

def clasificar_hdl(hdl: float, sexo: str):
    bajo = 40 if sexo == "Masculino" else 50
    if hdl >= 60: return "green", "Protector", ">=60 mg/dL"
    if hdl >= bajo: return "blue", "Aceptable", f">={bajo} mg/dL"
    return "red", "Bajo", f"<{bajo} mg/dL"

def clasificar_lpa(valor: Optional[float], unidad: str):
    if valor is None: return "gray", "No medida", "medir una vez en la vida"
    umbral = 50 if unidad == "mg/dL" else 125
    if valor < umbral: return "green", "No elevada", f"<{umbral} {unidad}"
    return "red", "Elevada", f">={umbral} {unidad}"

def clasificar_apob(valor: Optional[float]):
    if valor is None: return "gray", "No medida", "util si TG altos/diabetes/CKD"
    if valor < 90: return "green", "Optima", "<90 mg/dL"
    if valor < 130: return "orange", "Elevada", "90-129 mg/dL"
    return "red", "Muy elevada", ">=130 mg/dL"

def clasificar_egfr(valor: Optional[float]):
    if valor is None: return "gray", "No aplica", "si CKD, cargar eGFR"
    if valor >= 60: return "green", "Preservado", ">=60"
    if valor >= 30: return "orange", "Disminuido", "30-59"
    return "red", "Muy disminuido", "<30"

def clasificar_cac_valor(valor: Optional[int]):
    if valor is None: return "gray", "No disponible", "usar si duda clinica"
    if valor == 0: return "green", "CAC 0", "0"
    if valor < 100: return "orange", "CAC positivo", "1-99"
    return "red", "CAC alto", ">=100"

def clasificar_prevent_color(riesgo: Optional[float]):
    cat = clasificar_prevent(riesgo)
    if cat == "Bajo": return "green", cat, "<5%"
    if cat == "Limitrofe": return "yellow", cat, "5-<7.5%"
    if cat == "Intermedio": return "orange", cat, "7.5-<20%"
    if cat == "Alto": return "red", cat, ">=20%"
    return "gray", cat, "30-79 anos"

# =========================================================
# Decision farmacologica clara - PREVENCION PRIMARIA
# =========================================================
def decidir_tratamiento_primaria(p: Patient) -> Dict[str, object]:
    """
    Devuelve una decision clara para prevencion primaria:
    - indica si requiere farmaco SI/NO
    - droga sugerida (clase + ejemplo + dosis)
    - intensidad
    - justificacion
    """
    info = determinar_perfil(p)
    riesgo_cat = clasificar_prevent(p.prevent_10)

    # Caso 1: ASCVD = no aplica (esto es prev. primaria)
    if p.ascvd:
        return {
            "requiere_farmaco": True,
            "intensidad": "Alta",
            "droga_principal": "Atorvastatina 40-80 mg/dia o Rosuvastatina 20-40 mg/dia",
            "justificacion": "Paciente con ASCVD establecida (prevencion secundaria). La indicacion farmacologica es obligatoria.",
            "color": "red"
        }

    # Caso 2: LDL >=190 mg/dL o sospecha de FH
    if p.ldl_actual >= 190 or p.fh_sospecha:
        return {
            "requiere_farmaco": True,
            "intensidad": "Alta",
            "droga_principal": "Atorvastatina 40-80 mg/dia o Rosuvastatina 20-40 mg/dia",
            "complementaria": "Si no alcanza meta a 4-12 semanas: agregar Ezetimibe 10 mg/dia. Si persiste: PCSK9 mAb (alirocumab/evolocumab).",
            "justificacion": "LDL-C >=190 mg/dL o sospecha de hipercolesterolemia familiar. Indicacion farmacologica de clase I, sin necesidad de calculo de riesgo.",
            "color": "red"
        }

    # Caso 3: Diabetes 40-75 anos
    if p.diabetes and 40 <= p.edad <= 75:
        potenciadores_dm = sum([p.ckd, p.hta, p.tabaquismo, p.antecedente_familiar,
                                 lpa_alta(p.lpa_valor, p.lpa_unidad),
                                 (p.cac is not None and p.cac >= 100)])
        if potenciadores_dm >= 1 or p.edad >= 50:
            return {
                "requiere_farmaco": True,
                "intensidad": "Alta",
                "droga_principal": "Atorvastatina 40-80 mg/dia o Rosuvastatina 20-40 mg/dia",
                "complementaria": "Si LDL-C persiste >=70 mg/dL: agregar Ezetimibe 10 mg/dia.",
                "justificacion": "Diabetes con potenciadores adicionales o edad >=50 anos. Recomendacion clase I de estatina de alta intensidad.",
                "color": "red"
            }
        else:
            return {
                "requiere_farmaco": True,
                "intensidad": "Moderada",
                "droga_principal": "Atorvastatina 10-20 mg/dia o Rosuvastatina 5-10 mg/dia",
                "justificacion": "Diabetes en adulto 40-49 anos sin otros potenciadores: estatina de moderada intensidad como minimo.",
                "color": "orange"
            }

    # Caso 4: Enfermedad renal cronica (sin dialisis)
    if p.ckd and p.egfr is not None and p.egfr < 60 and p.egfr >= 30:
        return {
            "requiere_farmaco": True,
            "intensidad": "Moderada-alta",
            "droga_principal": "Atorvastatina 20-40 mg/dia (preferida en CKD por menor excrecion renal)",
            "justificacion": "Enfermedad renal cronica (eGFR 30-59): indicacion de estatina por riesgo CV elevado.",
            "color": "orange"
        }

    # Caso 5: Edad < 40 sin factores de alto riesgo
    if p.edad < 40 and not (p.fh_sospecha or p.ldl_actual >= 190 or lpa_alta(p.lpa_valor, p.lpa_unidad)):
        return {
            "requiere_farmaco": False,
            "intensidad": "Ninguna",
            "droga_principal": "No indicada de inicio. Priorizar estilo de vida.",
            "justificacion": "Edad <40 anos sin LDL severamente elevado, sin FH y sin Lp(a) elevada. Reevaluar en 3-5 anos.",
            "color": "green"
        }

    # Caso 6: PREVENT 10 anos - cuatro tiers
    if p.prevent_10 is None:
        return {
            "requiere_farmaco": False,
            "intensidad": "Por definir",
            "droga_principal": "Calcular PREVENT primero. Use el modulo PREVENT oficial.",
            "justificacion": "Sin estimacion de riesgo PREVENT no es posible decidir farmacoterapia en prevencion primaria estandar.",
            "color": "gray"
        }

    if riesgo_cat == "Alto":  # >=20%
        return {
            "requiere_farmaco": True,
            "intensidad": "Alta",
            "droga_principal": "Atorvastatina 40-80 mg/dia o Rosuvastatina 20-40 mg/dia",
            "complementaria": "Agregar Ezetimibe 10 mg/dia si LDL-C >=70 mg/dL a 4-12 semanas.",
            "justificacion": f"PREVENT 10 anos = {p.prevent_10}% (alto riesgo, >=20%). Indicacion clase I de estatina de alta intensidad.",
            "color": "red"
        }

    if riesgo_cat == "Intermedio":  # 7.5 a <20%
        # buscar potenciadores y CAC
        pots = potenciadores_riesgo(p)
        if p.cac is not None and p.cac >= 100:
            return {
                "requiere_farmaco": True,
                "intensidad": "Alta",
                "droga_principal": "Atorvastatina 40-80 mg/dia o Rosuvastatina 20-40 mg/dia",
                "justificacion": f"PREVENT intermedio ({p.prevent_10}%) + CAC >=100: tratar como alto riesgo.",
                "color": "red"
            }
        if p.cac is not None and p.cac == 0 and not p.diabetes and not p.tabaquismo and not p.fh_sospecha:
            return {
                "requiere_farmaco": False,
                "intensidad": "Diferir",
                "droga_principal": "Diferir estatina; reforzar estilo de vida y reevaluar en 3-5 anos.",
                "justificacion": f"PREVENT intermedio ({p.prevent_10}%) con CAC = 0 sin diabetes/tabaquismo/FH: bajo riesgo real, puede diferirse el tratamiento.",
                "color": "green"
            }
        if pots:
            return {
                "requiere_farmaco": True,
                "intensidad": "Moderada-alta",
                "droga_principal": "Atorvastatina 20-40 mg/dia o Rosuvastatina 10-20 mg/dia",
                "complementaria": "Reevaluar respuesta a 4-12 semanas; si no logra >=30-50% de reduccion, intensificar o agregar Ezetimibe.",
                "justificacion": f"PREVENT intermedio ({p.prevent_10}%) con potenciadores: {', '.join(pots[:4])}. Indicacion firme de estatina.",
                "color": "orange"
            }
        return {
            "requiere_farmaco": True,
            "intensidad": "Moderada",
            "droga_principal": "Atorvastatina 10-20 mg/dia o Rosuvastatina 5-10 mg/dia",
            "justificacion": f"PREVENT intermedio ({p.prevent_10}%) sin CAC disponible: estatina de moderada intensidad con decision compartida.",
            "color": "orange"
        }

    if riesgo_cat == "Limitrofe":  # 5 a <7.5%
        pots = potenciadores_riesgo(p)
        if p.cac is not None and p.cac >= 100:
            return {
                "requiere_farmaco": True,
                "intensidad": "Moderada",
                "droga_principal": "Atorvastatina 10-20 mg/dia o Rosuvastatina 5-10 mg/dia",
                "justificacion": f"PREVENT limitrofe ({p.prevent_10}%) + CAC >=100: indicar estatina.",
                "color": "orange"
            }
        if pots and len(pots) >= 1:
            return {
                "requiere_farmaco": True,
                "intensidad": "Moderada",
                "droga_principal": "Atorvastatina 10-20 mg/dia o Rosuvastatina 5-10 mg/dia",
                "justificacion": f"PREVENT limitrofe ({p.prevent_10}%) con potenciadores: {', '.join(pots[:3])}. Considerar estatina.",
                "color": "orange"
            }
        return {
            "requiere_farmaco": False,
            "intensidad": "Estilo de vida",
            "droga_principal": "Sin indicacion farmacologica de inicio. Intensificar estilo de vida y reevaluar.",
            "justificacion": f"PREVENT limitrofe ({p.prevent_10}%) sin potenciadores ni CAC alto: priorizar estilo de vida.",
            "color": "yellow"
        }

    # Bajo riesgo
    return {
        "requiere_farmaco": False,
        "intensidad": "Estilo de vida",
        "droga_principal": "Sin indicacion farmacologica. Estilo de vida cardioprotector.",
        "justificacion": f"PREVENT bajo ({p.prevent_10}%): priorizar dieta, actividad fisica, control de peso, no tabaco.",
        "color": "green"
    }

def potenciadores_riesgo(p: Patient) -> List[str]:
    out = []
    if lpa_alta(p.lpa_valor, p.lpa_unidad): out.append(f"Lp(a) elevada ({p.lpa_valor:.0f} {p.lpa_unidad})")
    if p.apob is not None and p.apob >= 130: out.append(f"ApoB elevada ({p.apob:.0f} mg/dL)")
    if p.diabetes: out.append("diabetes")
    if p.ckd: out.append("enfermedad renal cronica")
    if p.hta: out.append("hipertension arterial")
    if p.tabaquismo: out.append("tabaquismo activo")
    if p.inflamacion_cronica: out.append("inflamacion cronica")
    if p.antecedente_familiar: out.append("antecedente familiar de ASCVD prematura")
    if p.menopausia_precoz: out.append("menopausia precoz")
    if p.preeclampsia: out.append("antecedente de preeclampsia")
    if p.fh_sospecha: out.append("sospecha de hipercolesterolemia familiar")
    if p.cac is not None and p.cac > 0: out.append(f"CAC positivo ({p.cac})")
    if p.tg >= 175: out.append(f"hipertrigliceridemia ({p.tg:.0f} mg/dL)")
    return out

def recomendaciones_cac(p: Patient) -> str:
    if p.cac is None: return "CAC no disponible. Usar si hay incertidumbre en prevencion primaria limitrofe/intermedia."
    if p.cac == 0: return "CAC = 0: puede apoyar diferir o reducir intensidad en escenarios seleccionados, excepto diabetes, tabaquismo, FH, ASCVD familiar prematura u otros riesgos mayores."
    if 1 <= p.cac < 100: return "CAC 1-99: evidencia aterosclerotica subclinica; favorece estatina, especialmente con edad >55 anos o potenciadores."
    return "CAC >=100: favorece estatina e intensificacion para alcanzar meta de LDL-C."

def plan_farmacologico_completo(p: Patient) -> List[str]:
    info = determinar_perfil(p)
    estado = estado_meta(p)
    plan = []
    decision = decidir_tratamiento_primaria(p)

    if p.intolerancia_sams:
        plan.append("CONFIRMAR SAMS: evaluar temporalidad, CK si corresponde, interacciones, hipotiroidismo/deficiencia de vitamina D; realizar pausa y reexposicion con otra estatina, dosis baja o dias alternos.")
        plan.append("Usar Ezetimibe 10 mg/dia como primer no estatinico. Considerar Acido bempedoico 180 mg/dia si persiste LDL-C sobre meta.")
    else:
        plan.append(f"INDICACION PRINCIPAL: {decision['droga_principal']}")
        if "complementaria" in decision:
            plan.append(f"COMPLEMENTO: {decision['complementaria']}")

    if estado["texto"] != "En meta":
        if not p.ezetimibe and not p.intolerancia_sams:
            plan.append("ESCALADO 1: Si no alcanza meta a 4-12 semanas con estatina sola, agregar Ezetimibe 10 mg/dia (reduce LDL adicional ~20-25%).")
        if not (p.pcsk9 or p.inclisiran) and (info["perfil"] == "Prevencion secundaria" or "Alto" in info["riesgo"]):
            plan.append("ESCALADO 2: Si persiste sobre meta con estatina + ezetimibe en 4-12 semanas, considerar PCSK9 monoclonal (alirocumab 75-150 mg SC c/2 semanas o evolocumab 140 mg SC c/2 semanas o 420 mg c/mes). Inclisiran 284 mg SC dia 0, 90 y luego c/6 meses si se prioriza adherencia.")
        if not p.bempedoico:
            plan.append("ALTERNATIVA: Acido bempedoico 180 mg/dia, especialmente en intolerancia a estatinas o necesidad adicional de reduccion de LDL-C (~17-25%).")

    if p.lpa_valor is None:
        plan.append("LABORATORIO: Solicitar Lp(a) al menos una vez en la vida (estratificador adicional).")
    elif lpa_alta(p.lpa_valor, p.lpa_unidad):
        plan.append("Lp(a) ELEVADA: intensificar control de LDL-C y todos los factores de riesgo; en prevencion secundaria considerar PCSK9 si no alcanza meta.")

    if p.apob is None and (p.diabetes or p.ckd or p.tg >= 150):
        plan.append("LABORATORIO: Solicitar ApoB para evaluar riesgo residual por discordancia con LDL-C.")

    if not p.ascvd and p.cac is None and clasificar_prevent(p.prevent_10) in ["Limitrofe", "Intermedio"]:
        plan.append("IMAGEN: Considerar score de calcio coronario (CAC) para reclasificacion en duda clinica.")

    return plan

# =========================================================
# Datos para semaforizacion
# =========================================================
def semaforo_items_data(p: Patient, metas: Dict[str, str]) -> List[Dict[str, str]]:
    base = [
        ("LDL-C actual", p.ldl_actual, "mg/dL", *clasificar_ldl_vs_meta(p.ldl_actual, metas["ldl"])),
        ("No-HDL-C", p.no_hdl, "mg/dL", *clasificar_no_hdl(p.no_hdl, metas["no_hdl"])),
        ("Trigliceridos", p.tg, "mg/dL", *clasificar_tg(p.tg)),
        ("HDL-C", p.hdl, "mg/dL", *clasificar_hdl(p.hdl, p.sexo)),
        ("Lp(a)", p.lpa_valor, p.lpa_unidad, *clasificar_lpa(p.lpa_valor, p.lpa_unidad)),
        ("ApoB", p.apob, "mg/dL", *clasificar_apob(p.apob)),
        ("eGFR", p.egfr, "ml/min/1.73m2", *clasificar_egfr(p.egfr)),
        ("CAC", p.cac, "Agatston", *clasificar_cac_valor(p.cac)),
    ]
    if p.prevent_10 is not None:
        base.append(("PREVENT 10 anos", p.prevent_10, "%", *clasificar_prevent_color(p.prevent_10)))
    if p.prevent_30 is not None:
        color30 = "green" if p.prevent_30 < 15 else "orange" if p.prevent_30 < 30 else "red"
        interp30 = "Bajo largo plazo" if p.prevent_30 < 15 else "Intermedio largo plazo" if p.prevent_30 < 30 else "Alto largo plazo"
        base.append(("PREVENT 30 anos", p.prevent_30, "%", color30, interp30, "orientativo"))

    items = []
    for nombre, valor, unidad, color, interp, ref in base:
        if valor is None: valor_txt = "No informado"
        elif isinstance(valor, float): valor_txt = f"{valor:.1f} {unidad}"
        else: valor_txt = f"{valor} {unidad}"
        items.append({"indicador": nombre, "valor": valor_txt, "color": color, "interpretacion": interp, "referencia": ref})
    return items

def mostrar_panel_bioquimico(p: Patient, metas: Dict[str, str]):
    st.subheader("Semaforizacion bioquimica y de riesgo")
    items_data = semaforo_items_data(p, metas)
    cols = st.columns(4)
    for i, it in enumerate(items_data):
        with cols[i % 4]:
            # extract numeric
            valor_str = it["valor"].split(" ")[0] if it["valor"] != "No informado" else None
            try:
                valor_num = float(valor_str) if valor_str else None
            except Exception:
                valor_num = None
            unidad = it["valor"].split(" ", 1)[1] if " " in it["valor"] else ""
            semaforo_item(it["indicador"], valor_num if valor_num is not None else None, unidad,
                          it["color"], it["interpretacion"], it["referencia"])

# =========================================================
# PDF interno con SOPORTE DE GRAFICOS COLOREADOS
# =========================================================
COLOR_LIGHT_PDF = {
    "green": (0.78, 0.97, 0.85),
    "yellow": (1.00, 0.96, 0.62),
    "orange": (0.99, 0.88, 0.74),
    "red": (0.99, 0.83, 0.83),
    "blue": (0.80, 0.89, 1.00),
    "gray": (0.93, 0.93, 0.95)
}
COLOR_DARK_PDF = {
    "green": (0.06, 0.55, 0.25),
    "yellow": (0.79, 0.54, 0.02),
    "orange": (0.92, 0.34, 0.05),
    "red": (0.79, 0.10, 0.10),
    "blue": (0.13, 0.32, 0.65),
    "gray": (0.42, 0.45, 0.5)
}

def _pdf_escape_text(text: str) -> str:
    text = safe_text(text)
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

class PDFBuilder:
    """Constructor PDF minimo con soporte de texto y rectangulos coloreados."""
    def __init__(self, title: str = APP_NAME):
        self.width = 595
        self.height = 842
        self.margin_x = 40
        self.margin_y = 40
        self.pages: List[List[str]] = []
        self.current: List[str] = []
        self.y = self.height - self.margin_y
        self.title = title
        self.page_count = 0
        self._draw_header()

    def _draw_header(self):
        self.page_count += 1
        # banner azul
        self._rect_raw(0, self.height - 60, self.width, 60, COLOR_DARK_PDF["blue"])
        self._text_raw(self.margin_x, self.height - 35, self.title, 16, (1, 1, 1))
        self._text_raw(self.margin_x, self.height - 52, AUTOR_APP, 8, (1, 1, 1))
        self._text_raw(self.width - 130, self.height - 52, f"Fecha: {date.today().isoformat()}", 8, (1, 1, 1))
        self.y = self.height - 80

    def _rect_raw(self, x: float, y: float, w: float, h: float, fill_rgb: Tuple[float, float, float], stroke_rgb: Optional[Tuple[float, float, float]] = None, line_width: float = 0.7):
        r, g, b = fill_rgb
        self.current.append(f"{r:.3f} {g:.3f} {b:.3f} rg")
        self.current.append(f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re f")
        if stroke_rgb:
            sr, sg, sb = stroke_rgb
            self.current.append(f"{sr:.3f} {sg:.3f} {sb:.3f} RG")
            self.current.append(f"{line_width} w")
            self.current.append(f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re S")

    def _text_raw(self, x: float, y: float, text: str, size: int = 10, color_rgb: Tuple[float, float, float] = (0, 0, 0), bold: bool = False):
        r, g, b = color_rgb
        font = "F2" if bold else "F1"
        self.current.append(f"BT {r:.3f} {g:.3f} {b:.3f} rg /{font} {size} Tf {x:.2f} {y:.2f} Td ({_pdf_escape_text(text)}) Tj ET")

    def new_page(self):
        if self.current:
            self.pages.append(self.current)
        self.current = []
        self._draw_header()

    def ensure_space(self, h: float):
        if self.y - h < self.margin_y + 20:
            self.new_page()

    def heading(self, text: str, size: int = 13, color: Tuple[float, float, float] = (0.04, 0.30, 0.54)):
        self.ensure_space(size + 14)
        self.y -= size + 4
        # underline bar
        self._rect_raw(self.margin_x, self.y - 3, 4, size + 2, color)
        self._text_raw(self.margin_x + 10, self.y, text, size, color, bold=True)
        self.y -= 8

    def text(self, text: str, size: int = 10, color: Tuple[float, float, float] = (0.10, 0.10, 0.10), indent: float = 0):
        usable = self.width - 2 * self.margin_x - indent
        # Cada caracter aprox size*0.55 px
        chars_per_line = max(20, int(usable / (size * 0.50)))
        for raw_line in safe_text(text).split("\n"):
            wrapped = textwrap.wrap(raw_line, width=chars_per_line) if raw_line.strip() else [""]
            for line in wrapped:
                self.ensure_space(size + 2)
                self._text_raw(self.margin_x + indent, self.y - size, line, size, color)
                self.y -= size + 2

    def bullet(self, text: str, size: int = 10, color: Tuple[float, float, float] = (0.10, 0.10, 0.10)):
        self.ensure_space(size + 4)
        self._text_raw(self.margin_x + 8, self.y - size, "*", size, color, bold=True)
        self.text(text, size=size, color=color, indent=20)

    def colored_box(self, x: float, y: float, w: float, h: float, color_name: str,
                    label: str, value: str, status: str, reference: str):
        fill = COLOR_LIGHT_PDF.get(color_name, COLOR_LIGHT_PDF["gray"])
        border = COLOR_DARK_PDF.get(color_name, COLOR_DARK_PDF["gray"])
        # caja
        self._rect_raw(x, y, w, h, fill, border, 1.0)
        # banda lateral de color fuerte
        self._rect_raw(x, y, 5, h, border)
        # textos
        self._text_raw(x + 12, y + h - 16, label, 9, (0.20, 0.22, 0.26))
        self._text_raw(x + 12, y + h - 36, value, 13, (0.05, 0.05, 0.05), bold=True)
        self._text_raw(x + 12, y + h - 52, status, 9, border, bold=True)
        self._text_raw(x + 12, y + 8, f"Ref: {reference}", 7, (0.30, 0.32, 0.36))

    def grid_semaforo(self, items: List[Dict], cols: int = 2, box_h: float = 70, gap: float = 8):
        usable = self.width - 2 * self.margin_x
        box_w = (usable - gap * (cols - 1)) / cols
        for i, it in enumerate(items):
            col = i % cols
            if col == 0:
                self.ensure_space(box_h + 6)
                self.y -= box_h
                row_y = self.y
            x = self.margin_x + col * (box_w + gap)
            self.colored_box(x, row_y, box_w, box_h, it["color"],
                             it["indicador"], it["valor"], it["interpretacion"], it["referencia"])
            if col == cols - 1:
                self.y -= 4
        # Si quedo fila incompleta
        if len(items) % cols != 0:
            self.y -= 4

    def barra_meta_ldl(self, ldl_actual: float, ldl_basal: float, meta_ldl: Optional[float]):
        """Barra didactica que muestra LDL actual contra meta y basal."""
        self.ensure_space(70)
        self.y -= 10
        usable = self.width - 2 * self.margin_x
        bar_x = self.margin_x
        bar_y = self.y - 28
        bar_h = 22
        max_val = max(ldl_basal, ldl_actual, (meta_ldl or 0) + 50, 200)
        # fondo
        self._rect_raw(bar_x, bar_y, usable, bar_h, (0.95, 0.95, 0.97), (0.80, 0.82, 0.86), 0.5)
        # zona verde (meta cumplida)
        if meta_ldl is not None:
            zona_w = (meta_ldl / max_val) * usable
            self._rect_raw(bar_x, bar_y, zona_w, bar_h, COLOR_LIGHT_PDF["green"])
            # zona naranja
            zona_w2 = ((meta_ldl + 20) / max_val) * usable
            self._rect_raw(bar_x + zona_w, bar_y, zona_w2 - zona_w, bar_h, COLOR_LIGHT_PDF["orange"])
            # zona roja
            self._rect_raw(bar_x + zona_w2, bar_y, usable - zona_w2, bar_h, COLOR_LIGHT_PDF["red"])
            # linea de meta
            meta_x = bar_x + (meta_ldl / max_val) * usable
            self._rect_raw(meta_x - 1, bar_y - 4, 2, bar_h + 8, COLOR_DARK_PDF["green"])
            self._text_raw(meta_x - 18, bar_y + bar_h + 8, f"Meta {meta_ldl:.0f}", 7, COLOR_DARK_PDF["green"], bold=True)
        # marcador LDL actual
        ldl_x = bar_x + (ldl_actual / max_val) * usable
        self._rect_raw(ldl_x - 2, bar_y - 6, 4, bar_h + 12, (0.05, 0.05, 0.05))
        self._text_raw(ldl_x - 24, bar_y - 12, f"Actual {ldl_actual:.0f}", 8, (0.05, 0.05, 0.05), bold=True)
        # marcador LDL basal (si difiere)
        if ldl_basal > 0 and abs(ldl_basal - ldl_actual) > 5:
            bas_x = bar_x + (ldl_basal / max_val) * usable
            self._rect_raw(bas_x - 1, bar_y - 4, 2, bar_h + 8, (0.40, 0.40, 0.45))
            self._text_raw(bas_x - 22, bar_y + bar_h + 22, f"Basal {ldl_basal:.0f}", 7, (0.40, 0.40, 0.45))
        self._text_raw(self.margin_x, bar_y - 22, "LDL-C: zonas verde (meta), naranja (cerca), roja (fuera)", 8, (0.30, 0.32, 0.36))
        self.y = bar_y - 32

    def barra_riesgo_prevent(self, prevent_10: Optional[float]):
        if prevent_10 is None:
            return
        self.ensure_space(60)
        self.y -= 10
        usable = self.width - 2 * self.margin_x
        bar_x = self.margin_x
        bar_y = self.y - 24
        bar_h = 20
        max_val = 30  # %
        # zonas
        z1 = (5 / max_val) * usable    # bajo
        z2 = (7.5 / max_val) * usable  # limitrofe
        z3 = (20 / max_val) * usable   # intermedio
        self._rect_raw(bar_x, bar_y, z1, bar_h, COLOR_LIGHT_PDF["green"])
        self._rect_raw(bar_x + z1, bar_y, z2 - z1, bar_h, COLOR_LIGHT_PDF["yellow"])
        self._rect_raw(bar_x + z2, bar_y, z3 - z2, bar_h, COLOR_LIGHT_PDF["orange"])
        self._rect_raw(bar_x + z3, bar_y, usable - z3, bar_h, COLOR_LIGHT_PDF["red"])
        # marcador
        val = min(prevent_10, max_val)
        mx = bar_x + (val / max_val) * usable
        self._rect_raw(mx - 2, bar_y - 6, 4, bar_h + 12, (0.05, 0.05, 0.05))
        self._text_raw(mx - 22, bar_y - 12, f"PREVENT {prevent_10:.1f}%", 8, (0.05, 0.05, 0.05), bold=True)
        # rotulos
        self._text_raw(bar_x, bar_y - 22, "Riesgo PREVENT 10 anos: <5% bajo / 5-7.5% limitrofe / 7.5-20% intermedio / >=20% alto", 8, (0.30, 0.32, 0.36))
        self.y = bar_y - 28

    def line_chart(self, points: List[Tuple[str, float]], meta: Optional[float] = None,
                   title: str = "Evolucion temporal del LDL-C", y_label: str = "LDL-C (mg/dL)",
                   secundaria: Optional[List[Tuple[str, float]]] = None,
                   secundaria_label: str = "LDL basal"):
        """Dibuja un grafico de lineas con marcadores y opcional linea de meta."""
        self.ensure_space(220)
        self.y -= 24
        chart_x = self.margin_x + 50
        chart_h = 150
        chart_y = self.y - chart_h
        chart_w = self.width - 2 * self.margin_x - 60
        # Titulo
        self._text_raw(self.margin_x, chart_y + chart_h + 18, title, 11, (0.05, 0.05, 0.05), bold=True)
        # Fondo
        self._rect_raw(chart_x, chart_y, chart_w, chart_h, (1, 1, 1), (0.70, 0.72, 0.78), 0.5)
        if not points:
            self._text_raw(chart_x + chart_w / 2 - 50, chart_y + chart_h / 2, "Sin datos suficientes", 10, (0.40, 0.42, 0.46))
            self.y = chart_y - 18
            return
        # Escala
        all_vals = [v for _, v in points]
        if secundaria:
            all_vals += [v for _, v in secundaria]
        if meta is not None:
            all_vals.append(meta)
        y_max = max(all_vals) * 1.20 if all_vals else 200
        y_min = 0
        y_range = max(y_max - y_min, 1)
        # Gridlines y ticks Y
        for frac in [0, 0.25, 0.5, 0.75, 1.0]:
            v = y_min + frac * y_range
            ty = chart_y + frac * chart_h
            self._text_raw(chart_x - 38, ty - 3, f"{v:.0f}", 7, (0.40, 0.42, 0.46))
            self.current.append(f"0.88 0.89 0.92 RG 0.3 w {chart_x:.2f} {ty:.2f} m {chart_x + chart_w:.2f} {ty:.2f} l S")
        # Linea de meta (segmentos)
        if meta is not None:
            my = chart_y + ((meta - y_min) / y_range) * chart_h
            seg, gap, x = 7, 4, chart_x
            r, g, b = COLOR_DARK_PDF["green"]
            while x < chart_x + chart_w:
                self._rect_raw(x, my - 0.6, min(seg, chart_x + chart_w - x), 1.2, (r, g, b))
                x += seg + gap
            self._text_raw(chart_x + chart_w - 70, my + 3, f"Meta {meta:.0f} mg/dL", 7, (r, g, b), bold=True)
        # Posiciones X
        n = len(points)
        if n == 1:
            x_positions = [chart_x + chart_w / 2]
        else:
            x_positions = [chart_x + i * (chart_w / (n - 1)) for i in range(n)]
        # Coordenadas
        coords_main = [(x_positions[i], chart_y + ((points[i][1] - y_min) / y_range) * chart_h) for i in range(n)]
        # Linea principal (LDL actual)
        if len(coords_main) > 1:
            r, g, b = COLOR_DARK_PDF["blue"]
            path = f"{r:.3f} {g:.3f} {b:.3f} RG 1.6 w {coords_main[0][0]:.2f} {coords_main[0][1]:.2f} m"
            for cx, cy in coords_main[1:]:
                path += f" {cx:.2f} {cy:.2f} l"
            path += " S"
            self.current.append(path)
        # Serie secundaria (LDL basal historico)
        if secundaria and len(secundaria) == n:
            coords_sec = [(x_positions[i], chart_y + ((secundaria[i][1] - y_min) / y_range) * chart_h) for i in range(n)]
            if len(coords_sec) > 1:
                r, g, b = (0.55, 0.55, 0.60)
                path = f"{r:.3f} {g:.3f} {b:.3f} RG 1.0 w {coords_sec[0][0]:.2f} {coords_sec[0][1]:.2f} m"
                for cx, cy in coords_sec[1:]:
                    path += f" {cx:.2f} {cy:.2f} l"
                path += " S"
                self.current.append(path)
            for cx, cy in coords_sec:
                self._rect_raw(cx - 2.5, cy - 2.5, 5, 5, (0.55, 0.55, 0.60))
        # Marcadores principales
        rb, gb, bb = COLOR_DARK_PDF["blue"]
        for i, (cx, cy) in enumerate(coords_main):
            self._rect_raw(cx - 4, cy - 4, 8, 8, (rb, gb, bb))
            self._rect_raw(cx - 2, cy - 2, 4, 4, (1, 1, 1))
            # Valor encima del marcador
            self._text_raw(cx - 10, cy + 7, f"{points[i][1]:.0f}", 7, (rb, gb, bb), bold=True)
        # Etiquetas X (fechas)
        for i, (label, _) in enumerate(points):
            cx = x_positions[i]
            short = label[:10]
            self._text_raw(cx - 22, chart_y - 12, short, 6, (0.30, 0.32, 0.36))
        # Leyenda
        leg_y = chart_y + chart_h + 4
        # primary
        self._rect_raw(chart_x + 5, leg_y, 12, 3, COLOR_DARK_PDF["blue"])
        self._text_raw(chart_x + 22, leg_y - 2, "LDL-C actual", 7, (0.20, 0.22, 0.26))
        if secundaria:
            self._rect_raw(chart_x + 90, leg_y, 12, 3, (0.55, 0.55, 0.60))
            self._text_raw(chart_x + 107, leg_y - 2, secundaria_label, 7, (0.20, 0.22, 0.26))
        if meta is not None:
            self._rect_raw(chart_x + 170, leg_y, 12, 3, COLOR_DARK_PDF["green"])
            self._text_raw(chart_x + 187, leg_y - 2, f"Meta {meta:.0f}", 7, COLOR_DARK_PDF["green"], bold=True)
        # Eje X label
        self._text_raw(chart_x + chart_w / 2 - 30, chart_y - 24, "Fecha de evaluacion", 7, (0.30, 0.32, 0.36))
        self.y = chart_y - 32

    def build(self) -> bytes:
        if self.current:
            self.pages.append(self.current)
        kids_refs, page_objs = [], []
        next_obj = 5  # 1=cat, 2=pages, 3=fontF1, 4=fontF2 -> luego pages
        for page_lines in self.pages:
            content = "\n".join(page_lines).encode("latin-1", errors="replace")
            compressed = zlib.compress(content)
            content_obj_num = next_obj + 1
            page_obj = (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {self.width} {self.height}] "
                        f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {content_obj_num} 0 R >>").encode()
            content_obj = b"<< /Length " + str(len(compressed)).encode() + b" /Filter /FlateDecode >>\nstream\n" + compressed + b"\nendstream"
            page_objs.extend([page_obj, content_obj])
            kids_refs.append(f"{next_obj} 0 R")
            next_obj += 2
        catalog = b"<< /Type /Catalog /Pages 2 0 R >>"
        pages = f"<< /Type /Pages /Kids [{' '.join(kids_refs)}] /Count {len(self.pages)} >>".encode()
        font1 = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        font2 = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"
        all_objs = [catalog, pages, font1, font2] + page_objs
        out = bytearray()
        out.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for i, obj in enumerate(all_objs, start=1):
            offsets.append(len(out))
            out.extend(f"{i} 0 obj\n".encode())
            out.extend(obj)
            out.extend(b"\nendobj\n")
        xref_pos = len(out)
        out.extend(f"xref\n0 {len(all_objs)+1}\n".encode())
        out.extend(b"0000000000 65535 f \n")
        for off in offsets[1:]:
            out.extend(f"{off:010d} 00000 n \n".encode())
        out.extend(f"trailer\n<< /Size {len(all_objs)+1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF".encode())
        return bytes(out)

# =========================================================
# Generadores de PDF graficos
# =========================================================
def pdf_informe_medico(p: Patient) -> bytes:
    info = determinar_perfil(p)
    metas = metas_lipidicas(p)
    estado = estado_meta(p)
    decision = decidir_tratamiento_primaria(p)
    plan = plan_farmacologico_completo(p)
    items = semaforo_items_data(p, metas)
    pots = potenciadores_riesgo(p)
    meta_ldl = parse_meta_ldl(metas["ldl"])

    pdf = PDFBuilder(title=f"{APP_NAME} - Informe medico")

    # Datos paciente
    pdf.heading("Datos del paciente")
    pdf.text(f"Paciente: {p.paciente or 'No informado'}    DNI/ID: {p.dni or 'No informado'}")
    pdf.text(f"Edad: {p.edad} anos    Sexo: {p.sexo}    PAS: {p.presion_sistolica:.0f} mmHg")
    pdf.text(f"Medico: {p.medico or 'No informado'}    Matricula: {p.matricula or 'No informada'}")

    # Perfil
    pdf.heading("Perfil de riesgo")
    pdf.text(f"{info['perfil']} - {info['riesgo']}", size=11)
    pdf.text(f"Variables PREVENT sincronizadas: edad {p.edad} anos, sexo {p.sexo}, colesterol total {p.colesterol_total:.0f} mg/dL, HDL {p.hdl:.0f} mg/dL, PAS {p.presion_sistolica:.0f} mmHg, diabetes {'si' if p.diabetes else 'no'}, tabaquismo {'si' if p.tabaquismo else 'no'}, eGFR {p.egfr if p.egfr is not None else 'no informado'}.", size=9)
    if p.prevent_10 is not None:
        pdf.text(f"PREVENT 10 anos: {p.prevent_10}% ({clasificar_prevent(p.prevent_10)})")
    if p.prevent_30 is not None:
        pdf.text(f"PREVENT 30 anos: {p.prevent_30}%")
    pdf.barra_riesgo_prevent(p.prevent_10)

    # Metas
    pdf.heading("Metas lipidicas")
    pdf.text(f"LDL-C objetivo: {metas['ldl']}    No-HDL-C: {metas['no_hdl']}    Reduccion: {metas['reduccion']}")
    pdf.text(f"Estado actual: {estado['texto']}", size=11)
    pdf.barra_meta_ldl(p.ldl_actual, p.ldl_basal, meta_ldl)

    # Semaforizacion grafica
    pdf.heading("Semaforizacion bioquimica y de riesgo")
    pdf.grid_semaforo(items, cols=2, box_h=72, gap=10)

    # Potenciadores
    pdf.heading("Potenciadores / reclasificadores")
    if pots:
        for pot in pots:
            pdf.bullet(pot)
    else:
        pdf.text("No registrados.")
    pdf.text(recomendaciones_cac(p))

    # DECISION FARMACOLOGICA
    pdf.heading("Decision farmacologica recomendada")
    color_dec = decision.get("color", "blue")
    fill = COLOR_LIGHT_PDF.get(color_dec, COLOR_LIGHT_PDF["blue"])
    border = COLOR_DARK_PDF.get(color_dec, COLOR_DARK_PDF["blue"])
    pdf.ensure_space(110)
    pdf.y -= 90
    box_y = pdf.y
    pdf._rect_raw(pdf.margin_x, box_y, pdf.width - 2 * pdf.margin_x, 90, fill, border, 1.2)
    pdf._rect_raw(pdf.margin_x, box_y, 6, 90, border)
    req_txt = "REQUIERE FARMACO" if decision["requiere_farmaco"] else "NO REQUIERE FARMACO DE INICIO"
    pdf._text_raw(pdf.margin_x + 14, box_y + 70, req_txt, 12, border, bold=True)
    pdf._text_raw(pdf.margin_x + 14, box_y + 54, f"Intensidad: {decision['intensidad']}", 10, (0.10, 0.10, 0.10), bold=True)
    pdf._text_raw(pdf.margin_x + 14, box_y + 38, f"Droga sugerida: {safe_text(decision['droga_principal'])[:90]}", 9, (0.10, 0.10, 0.10))
    if "complementaria" in decision:
        pdf._text_raw(pdf.margin_x + 14, box_y + 22, f"Complemento: {safe_text(decision['complementaria'])[:95]}", 8, (0.20, 0.22, 0.26))
    pdf._text_raw(pdf.margin_x + 14, box_y + 8, f"Justificacion: {safe_text(decision['justificacion'])[:100]}", 8, (0.20, 0.22, 0.26))
    pdf.y = box_y - 6

    # Plan completo
    pdf.heading("Plan farmacologico y de seguimiento")
    for item in plan:
        pdf.bullet(item, size=9)

    # Seguimiento
    pdf.heading("Seguimiento")
    pdf.text("Repetir perfil lipidico en 4-12 semanas luego del inicio o ajuste terapeutico. Luego controlar cada 3-12 meses segun estabilidad y adherencia.")
    if p.observaciones:
        pdf.heading("Observaciones")
        pdf.text(p.observaciones)

    pdf.heading("Aviso")
    pdf.text("Herramienta de soporte a la decision clinica. No sustituye juicio medico ni guias/regulaciones locales.", size=8, color=(0.40, 0.42, 0.45))

    return pdf.build()

def pdf_informe_paciente(p: Patient) -> bytes:
    info = determinar_perfil(p)
    metas = metas_lipidicas(p)
    estado = estado_meta(p)
    decision = decidir_tratamiento_primaria(p)
    items = semaforo_items_data(p, metas)
    meta_ldl = parse_meta_ldl(metas["ldl"])

    pdf = PDFBuilder(title=f"{APP_NAME} - Informe para el paciente")

    pdf.heading("Datos personales")
    pdf.text(f"Paciente: {p.paciente or 'No informado'}")
    pdf.text(f"Edad: {p.edad} anos")
    pdf.text(f"Medico: {p.medico or 'No informado'}")

    pdf.heading("Su nivel de riesgo")
    pdf.text(f"Su perfil corresponde a: {info['perfil']} - {info['riesgo']}.", size=11)
    pdf.text("Este nivel de riesgo determina las metas de tratamiento que su medico recomienda.")

    pdf.heading("Su LDL-C (colesterol malo)")
    pdf.text(f"Valor actual: {p.ldl_actual:.0f} mg/dL    Meta sugerida: {metas['ldl']}", size=11)
    pdf.text(f"Estado: {estado['texto']}.")
    pdf.barra_meta_ldl(p.ldl_actual, p.ldl_basal, meta_ldl)

    pdf.heading("Resultados de su laboratorio")
    pdf.grid_semaforo(items[:6], cols=2, box_h=68, gap=10)

    pdf.heading("Que significa la semaforizacion")
    pdf.text("Verde: dentro de los valores recomendados.")
    pdf.text("Naranja/Amarillo: cerca pero no en meta. Hay margen de mejora.")
    pdf.text("Rojo: fuera de meta. Es necesario actuar.")

    pdf.heading("Que indica su medico")
    color_dec = decision.get("color", "blue")
    fill = COLOR_LIGHT_PDF.get(color_dec, COLOR_LIGHT_PDF["blue"])
    border = COLOR_DARK_PDF.get(color_dec, COLOR_DARK_PDF["blue"])
    pdf.ensure_space(80)
    pdf.y -= 60
    box_y = pdf.y
    pdf._rect_raw(pdf.margin_x, box_y, pdf.width - 2 * pdf.margin_x, 60, fill, border, 1.0)
    pdf._rect_raw(pdf.margin_x, box_y, 6, 60, border)
    req_txt = "Necesita medicacion" if decision["requiere_farmaco"] else "Por ahora, no necesita medicacion"
    pdf._text_raw(pdf.margin_x + 14, box_y + 42, req_txt, 12, border, bold=True)
    pdf._text_raw(pdf.margin_x + 14, box_y + 24, safe_text(decision['droga_principal'])[:90], 9, (0.10, 0.10, 0.10))
    pdf._text_raw(pdf.margin_x + 14, box_y + 8, safe_text(decision['justificacion'])[:100], 8, (0.20, 0.22, 0.26))
    pdf.y = box_y - 6

    pdf.heading("Recomendaciones de estilo de vida")
    pdf.bullet("Alimentacion cardioprotectora: vegetales, frutas, granos integrales, pescado, aceite de oliva. Reducir ultraprocesados y grasas saturadas.")
    pdf.bullet("Actividad fisica: al menos 150 minutos por semana de intensidad moderada o 75 minutos de alta intensidad.")
    pdf.bullet("No fumar ni vapear. Evitar humo de segunda mano.")
    pdf.bullet("Controlar peso, presion arterial y glucemia.")
    pdf.bullet("Cumplir la medicacion indicada y no suspenderla sin consultar.")
    pdf.bullet("Repetir el laboratorio en 4 a 12 semanas si hubo inicio o cambio de tratamiento.")

    pdf.heading("Aviso")
    pdf.text("Este informe simplificado no reemplaza la explicacion personalizada de su medico.", size=8, color=(0.40, 0.42, 0.45))

    return pdf.build()

# =========================================================
# Evolucion temporal del paciente
# =========================================================
def _patient_key(registro: dict) -> str:
    """Clave para agrupar evaluaciones del mismo paciente."""
    dni = (registro.get("dni") or "").strip()
    nombre = (registro.get("paciente") or "").strip()
    return f"DNI:{dni}" if dni else f"NOM:{nombre.lower()}"

def listar_pacientes_unicos(username: str) -> List[Dict]:
    """Devuelve lista de pacientes unicos del usuario con cantidad de evaluaciones."""
    historial = load_historial().get(username, [])
    grupos: Dict[str, Dict] = {}
    for r in historial:
        k = _patient_key(r)
        if not k or k in ("DNI:", "NOM:"):
            continue
        if k not in grupos:
            grupos[k] = {
                "key": k,
                "dni": r.get("dni") or "",
                "paciente": r.get("paciente") or "",
                "evaluaciones": []
            }
        grupos[k]["evaluaciones"].append(r)
    out = []
    for k, g in grupos.items():
        g["evaluaciones"].sort(key=lambda x: x.get("fecha", ""))
        g["n"] = len(g["evaluaciones"])
        g["primera"] = g["evaluaciones"][0].get("fecha", "")
        g["ultima"] = g["evaluaciones"][-1].get("fecha", "")
        ldls = [e.get("ldl_actual") for e in g["evaluaciones"] if isinstance(e.get("ldl_actual"), (int, float))]
        g["ldl_inicial"] = ldls[0] if ldls else None
        g["ldl_ultimo"] = ldls[-1] if ldls else None
        if g["ldl_inicial"] and g["ldl_ultimo"] and g["ldl_inicial"] > 0:
            g["delta_pct"] = round((g["ldl_inicial"] - g["ldl_ultimo"]) / g["ldl_inicial"] * 100, 1)
        else:
            g["delta_pct"] = None
        out.append(g)
    out.sort(key=lambda x: x["paciente"].lower() or x["dni"])
    return out

def evolucion_dataframe(evaluaciones: List[dict]) -> pd.DataFrame:
    if not evaluaciones:
        return pd.DataFrame()
    df = pd.DataFrame(evaluaciones)
    if "fecha" in df.columns:
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
        df = df.sort_values("fecha").reset_index(drop=True)
    cols_pref = ["fecha", "ldl_actual", "ldl_basal", "no_hdl", "tg", "hdl",
                 "estado_meta", "meta_ldl", "reduccion_ldl", "intensidad_recomendada",
                 "droga_recomendada"]
    cols_present = [c for c in cols_pref if c in df.columns]
    other = [c for c in df.columns if c not in cols_present]
    return df[cols_present + other]

def pdf_evolucion_paciente(username: str, paciente_key: str) -> Optional[bytes]:
    """Genera PDF con grafico de evolucion temporal de un paciente."""
    pacientes = listar_pacientes_unicos(username)
    g = next((x for x in pacientes if x["key"] == paciente_key), None)
    if g is None:
        return None
    evals = g["evaluaciones"]
    if not evals:
        return None
    pdf = PDFBuilder(title=f"{APP_NAME} - Evolucion del paciente")

    pdf.heading("Datos del paciente")
    pdf.text(f"Paciente: {g['paciente'] or 'No informado'}    DNI/ID: {g['dni'] or 'No informado'}")
    pdf.text(f"Cantidad de evaluaciones registradas: {g['n']}")
    pdf.text(f"Primera evaluacion: {str(g['primera'])[:10]}    Ultima evaluacion: {str(g['ultima'])[:10]}")
    if g["delta_pct"] is not None:
        signo = "reduccion" if g["delta_pct"] >= 0 else "aumento"
        pdf.text(f"Cambio LDL-C desde primera a ultima evaluacion: {abs(g['delta_pct']):.1f}% ({signo}) - de {g['ldl_inicial']:.0f} a {g['ldl_ultimo']:.0f} mg/dL.", size=11)

    # Grafico de evolucion
    pdf.heading("Evolucion temporal del LDL-C")
    puntos = []
    puntos_basal = []
    for e in evals:
        f = str(e.get("fecha", ""))[:10]
        ldl_a = e.get("ldl_actual")
        ldl_b = e.get("ldl_basal")
        if isinstance(ldl_a, (int, float)):
            puntos.append((f, float(ldl_a)))
            puntos_basal.append((f, float(ldl_b) if isinstance(ldl_b, (int, float)) else float(ldl_a)))
    # Meta a partir de la ultima evaluacion
    meta_str = evals[-1].get("meta_ldl", "")
    meta_val = None
    for token in ["55", "70", "100", "116"]:
        if token in str(meta_str):
            meta_val = float(token); break
    pdf.line_chart(puntos, meta=meta_val, title="LDL-C a lo largo del tiempo",
                   secundaria=puntos_basal if any(b[1] != p[1] for b, p in zip(puntos_basal, puntos)) else None,
                   secundaria_label="LDL basal pre-tto")

    # Tabla resumen de evaluaciones
    pdf.heading("Tabla de evaluaciones")
    headers = ["Fecha", "LDL act", "LDL bas", "No-HDL", "TG", "Estado", "Meta"]
    pdf.ensure_space(20)
    pdf.y -= 14
    col_w = [70, 55, 55, 55, 45, 75, 60]
    cx = pdf.margin_x
    # cabecera
    pdf._rect_raw(cx, pdf.y, sum(col_w), 16, COLOR_DARK_PDF["blue"])
    for i, h in enumerate(headers):
        pdf._text_raw(cx + 4, pdf.y + 4, h, 8, (1, 1, 1), bold=True)
        cx += col_w[i]
    pdf.y -= 4
    # filas
    for idx, e in enumerate(evals):
        pdf.ensure_space(14)
        pdf.y -= 14
        cx = pdf.margin_x
        bg = (0.97, 0.98, 0.99) if idx % 2 == 0 else (1, 1, 1)
        pdf._rect_raw(cx, pdf.y, sum(col_w), 14, bg, (0.85, 0.87, 0.90), 0.3)
        ldl_act = e.get("ldl_actual")
        ldl_bas = e.get("ldl_basal")
        no_hdl = e.get("no_hdl")
        tg = e.get("tg")
        estado = str(e.get("estado_meta", ""))[:14]
        meta = str(e.get("meta_ldl", ""))[:10]
        valores = [
            str(e.get("fecha", ""))[:10],
            f"{ldl_act:.0f}" if isinstance(ldl_act, (int, float)) else "-",
            f"{ldl_bas:.0f}" if isinstance(ldl_bas, (int, float)) else "-",
            f"{no_hdl:.0f}" if isinstance(no_hdl, (int, float)) else "-",
            f"{tg:.0f}" if isinstance(tg, (int, float)) else "-",
            estado, meta
        ]
        for i, v in enumerate(valores):
            color = (0.05, 0.05, 0.05)
            if i == 5:  # estado
                if "En meta" in v: color = COLOR_DARK_PDF["green"]
                elif "Cerca" in v: color = COLOR_DARK_PDF["orange"]
                elif "Fuera" in v: color = COLOR_DARK_PDF["red"]
            pdf._text_raw(cx + 4, pdf.y + 4, v, 8, color)
            cx += col_w[i]

    # Ultima decision
    ultima = evals[-1]
    pdf.heading("Ultima recomendacion farmacologica registrada")
    intensidad = ultima.get("intensidad_recomendada", "No registrada")
    droga = ultima.get("droga_recomendada", "No registrada")
    requiere = "SI" if ultima.get("requiere_farmaco") else "NO"
    pdf.text(f"Requiere farmaco: {requiere}   Intensidad: {intensidad}")
    pdf.text(f"Droga sugerida: {droga}", size=10)
    just = ultima.get("justificacion_decision", "")
    if just:
        pdf.text(f"Justificacion: {just}", size=9)

    pdf.heading("Aviso")
    pdf.text("Herramienta de soporte clinico. La interpretacion final corresponde al medico tratante.", size=8, color=(0.40, 0.42, 0.45))
    return pdf.build()

# =========================================================
# Texto plano para descargas TXT (mantenido)
# =========================================================
def nota_clinica(p: Patient) -> str:
    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    pots, plan = potenciadores_riesgo(p), plan_farmacologico_completo(p)
    decision = decidir_tratamiento_primaria(p)
    red = estado["reduccion"]
    semaforo_txt = "\n".join([f"- {it['indicador']}: {it['valor']} | {it['interpretacion']} | Referencia/meta: {it['referencia']}" for it in semaforo_items_data(p, metas)])
    return f"""
{APP_NAME}
Autor: {AUTOR_APP}
Fecha: {date.today().isoformat()}
Paciente: {p.paciente or 'No informado'}
DNI/ID: {p.dni or 'No informado'}
Medico: {p.medico or 'No informado'}
Matricula: {p.matricula or 'No informada'}

PERFIL DE RIESGO
{perfil['perfil']} - {perfil['riesgo']}.
Presion sistolica para PREVENT: {p.presion_sistolica:.0f} mmHg.
Riesgo PREVENT 10 anos: {p.prevent_10 if p.prevent_10 is not None else 'No aplica/no informado'}% ({clasificar_prevent(p.prevent_10)}).
Riesgo PREVENT 30 anos: {p.prevent_30 if p.prevent_30 is not None else 'No informado'}%.

DECISION FARMACOLOGICA
Requiere farmaco: {'SI' if decision['requiere_farmaco'] else 'NO'}
Intensidad: {decision['intensidad']}
Droga sugerida: {decision['droga_principal']}
{('Complemento: ' + decision['complementaria']) if 'complementaria' in decision else ''}
Justificacion: {decision['justificacion']}

LIPIDOS
Colesterol total: {p.colesterol_total:.0f} mg/dL.
LDL-C basal: {p.ldl_basal:.0f} mg/dL.
LDL-C actual: {p.ldl_actual:.0f} mg/dL.
Reduccion LDL-C: {red if isinstance(red, str) else str(red) + '%'}.
HDL-C: {p.hdl:.0f} mg/dL.
No-HDL-C: {p.no_hdl:.0f} mg/dL.
Trigliceridos: {p.tg:.0f} mg/dL.
Lp(a): {str(p.lpa_valor) + ' ' + p.lpa_unidad if p.lpa_valor is not None else 'No informada'}.
ApoB: {str(p.apob) + ' mg/dL' if p.apob is not None else 'No informada'}.

METAS RECOMENDADAS
LDL-C: {metas['ldl']}.
No-HDL-C: {metas['no_hdl']}.
Reduccion recomendada: {metas['reduccion']}.
Estado actual: {estado['texto']}.

SEMAFORIZACION BIOQUIMICA Y DE RIESGO
{semaforo_txt}

POTENCIADORES / RECLASIFICADORES
{', '.join(pots) if pots else 'No registrados'}.

CAC: {recomendaciones_cac(p)}

PLAN FARMACOLOGICO
""" + "\n".join([f"- {x}" for x in plan]) + f"""

SEGUIMIENTO
Repetir perfil lipidico en 4-12 semanas luego de inicio o ajuste. Luego cada 3-12 meses.

OBSERVACIONES
{p.observaciones or 'Sin observaciones adicionales.'}

Aviso: herramienta de soporte a la decision clinica.
"""

# =========================================================
# Exportacion de datos
# =========================================================
def make_row(p: Patient, decision: dict) -> Dict[str, object]:
    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    return {
        "fecha": datetime.now().isoformat(timespec="seconds"),
        "paciente": p.paciente, "dni": p.dni, "medico": p.medico, "matricula": p.matricula,
        "edad": p.edad, "sexo": p.sexo,
        "perfil": perfil["perfil"], "riesgo": perfil["riesgo"],
        "presion_sistolica_prevent": p.presion_sistolica,
        "prevent_10": p.prevent_10, "prevent_30": p.prevent_30,
        "colesterol_total": p.colesterol_total, "hdl": p.hdl,
        "ldl_basal": p.ldl_basal, "ldl_actual": p.ldl_actual,
        "reduccion_ldl": estado["reduccion"], "meta_ldl": metas["ldl"],
        "estado_meta": estado["texto"],
        "no_hdl": p.no_hdl, "tg": p.tg, "lpa": p.lpa_valor, "lpa_unidad": p.lpa_unidad, "apob": p.apob,
        "diabetes": p.diabetes, "ckd": p.ckd, "egfr": p.egfr, "hta": p.hta, "tratamiento_hta": p.tratamiento_hta, "tabaquismo": p.tabaquismo,
        "inflamacion_cronica": p.inflamacion_cronica, "antecedente_familiar": p.antecedente_familiar,
        "menopausia_precoz": p.menopausia_precoz, "preeclampsia": p.preeclampsia,
        "ascvd": p.ascvd, "iam": p.iam, "acv": p.acv, "pad": p.pad, "revascularizacion": p.revascularizacion,
        "fh_sospecha": p.fh_sospecha, "cac": p.cac,
        "estatina": p.estatina, "dosis_estatina": p.dosis_estatina,
        "ezetimibe": p.ezetimibe, "pcsk9": p.pcsk9, "inclisiran": p.inclisiran,
        "bempedoico": p.bempedoico, "sams": p.intolerancia_sams,
        "requiere_farmaco": decision["requiere_farmaco"],
        "intensidad_recomendada": decision["intensidad"],
        "droga_recomendada": decision["droga_principal"],
        "justificacion_decision": decision["justificacion"],
        "observaciones": p.observaciones,
    }

def excel_bytes_from_df(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    if EXCEL_ENGINE is None:
        raise RuntimeError("Motor Excel no disponible.")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine=EXCEL_ENGINE) as writer:
        clean = sheet_name[:31] or "Datos"
        df.to_excel(writer, index=False, sheet_name=clean)
        ws = writer.sheets[clean]
        for i, col in enumerate(df.columns):
            width = min(max(len(str(col)) + 2, 12), 42)
            if EXCEL_ENGINE == "xlsxwriter":
                ws.set_column(i, i, width)
            else:
                from openpyxl.utils import get_column_letter
                ws.column_dimensions[get_column_letter(i + 1)].width = width
    return output.getvalue()

def excel_historial_multiusuario_bytes() -> bytes:
    if EXCEL_ENGINE is None:
        raise RuntimeError("Motor Excel no disponible.")
    historiales = load_historial()
    users = load_users()
    output = io.BytesIO()
    all_rows = []
    with pd.ExcelWriter(output, engine=EXCEL_ENGINE) as writer:
        resumen = []
        for usuario, registros in historiales.items():
            if not registros: continue
            df = pd.DataFrame(registros)
            df.insert(0, "usuario", usuario)
            all_rows.append(df)
            nombre_visible = users.get(usuario, {}).get("nombre", usuario)
            resumen.append({"usuario_login": usuario, "nombre": nombre_visible,
                            "matricula": users.get(usuario, {}).get("matricula", ""),
                            "cantidad_pacientes": len(df)})
            hoja = usuario.replace("/", "-").replace("\\", "-").replace("*", "-").replace("?", "-").replace(":", "-")[:31]
            df.to_excel(writer, index=False, sheet_name=hoja or "Usuario")
        if resumen:
            pd.DataFrame(resumen).to_excel(writer, index=False, sheet_name="Resumen")
        else:
            pd.DataFrame(columns=["usuario_login", "nombre", "matricula", "cantidad_pacientes"]).to_excel(writer, index=False, sheet_name="Resumen")
        if all_rows:
            consolidado = pd.concat(all_rows, ignore_index=True)
            consolidado.to_excel(writer, index=False, sheet_name="Consolidado")
    return output.getvalue()

# =========================================================
# UI - Login
# =========================================================
def render_login():
    st.markdown(f'''<div class="hero"><h1>🫀 {APP_NAME}</h1><p>Acceda al sistema con su usuario o registre uno nuevo. Cada profesional gestiona su propio historial de pacientes.</p></div>''', unsafe_allow_html=True)
    tab_login, tab_reg = st.tabs(["Iniciar sesion", "Registrarse"])
    with tab_login:
        with st.form("login_form"):
            u = st.text_input("Usuario")
            pw = st.text_input("Contrasena", type="password")
            ok = st.form_submit_button("Ingresar")
            if ok:
                valid, user = authenticate(u.strip(), pw)
                if valid:
                    st.session_state.authenticated = True
                    st.session_state.username = u.strip()
                    st.session_state.user_data = user
                    st.success(f"Bienvenido, {user.get('nombre', u)}")
                    st.rerun()
                else:
                    st.error("Usuario o contrasena incorrectos.")
        st.caption(f"Usuario administrador por defecto: {DEFAULT_ADMIN_USER} / {DEFAULT_ADMIN_PASS} (cambiar en produccion).")
    with tab_reg:
        with st.form("reg_form"):
            new_u = st.text_input("Nuevo usuario")
            new_pw = st.text_input("Nueva contrasena (>=6 caracteres)", type="password")
            nombre = st.text_input("Nombre completo")
            mat = st.text_input("Matricula")
            esp = st.text_input("Especialidad", value="Cardiologia")
            ok2 = st.form_submit_button("Registrarse")
            if ok2:
                ok_ok, msg = register_user(new_u, new_pw, nombre, mat, esp, rol="medico")
                if ok_ok: st.success(msg)
                else: st.error(msg)


# =========================================================
# Sincronizacion PREVENT desde barra lateral
# =========================================================
def _ss_get(name: str, default=None):
    return st.session_state.get(name, default)

def _prevent_bool_txt(v: bool) -> str:
    return "Si" if bool(v) else "No"

def get_prevent_inputs_from_state() -> Dict[str, object]:
    """Devuelve las variables PREVENT tomadas desde los mismos widgets de la barra lateral."""
    return {
        "Edad": _ss_get("prev_edad", 55),
        "Sexo": _ss_get("prev_sexo", "Masculino"),
        "Colesterol total": _ss_get("prev_colesterol_total", 220.0),
        "HDL-C": _ss_get("prev_hdl", 45.0),
        "Presion sistolica": _ss_get("prev_pas", 130.0),
        "Tratamiento antihipertensivo": _prevent_bool_txt(_ss_get("prev_tratamiento_hta", False)),
        "Diabetes": _prevent_bool_txt(_ss_get("prev_diabetes", False)),
        "Tabaquismo activo": _prevent_bool_txt(_ss_get("prev_tabaquismo", False)),
        "eGFR": _ss_get("prev_egfr", 75.0),
        "Estatina actual": _ss_get("prev_estatina", "Ninguna"),
        "ASCVD clinica": _prevent_bool_txt(_ss_get("prev_ascvd", False)),
    }

def render_prevent_sync_panel(show_editor: bool = False):
    """Panel unico de variables PREVENT sincronizadas con Evaluacion clinica."""
    st.markdown("### Variables sincronizadas para PREVENT")
    st.caption("Estos datos se completan automaticamente desde la barra vertical de ingreso clinico. Si se editan aqui, quedan sincronizados para la evaluacion.")

    if show_editor:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.number_input("Edad", 18, 100, int(_ss_get("prev_edad", 55)), key="prev_edad")
            st.selectbox("Sexo", ["Masculino", "Femenino"], key="prev_sexo")
            st.number_input("Presion sistolica / PAS (mmHg)", 70.0, 260.0, float(_ss_get("prev_pas", 130.0)), 1.0, key="prev_pas")
        with c2:
            st.number_input("Colesterol total (mg/dL)", 0.0, 600.0, float(_ss_get("prev_colesterol_total", 220.0)), 1.0, key="prev_colesterol_total")
            st.number_input("HDL-C (mg/dL)", 0.0, 150.0, float(_ss_get("prev_hdl", 45.0)), 1.0, key="prev_hdl")
            st.number_input("eGFR ml/min/1.73m2", 0.0, 150.0, float(_ss_get("prev_egfr", 75.0)), 1.0, key="prev_egfr")
        with c3:
            st.checkbox("Diabetes", key="prev_diabetes")
            st.checkbox("Tabaquismo activo", key="prev_tabaquismo")
            st.checkbox("Tratamiento antihipertensivo", key="prev_tratamiento_hta")
            st.checkbox("ASCVD clinica establecida", key="prev_ascvd")

    datos = get_prevent_inputs_from_state()
    df_prev = pd.DataFrame([datos])
    st.dataframe(df_prev, use_container_width=True, hide_index=True)

    ascvd = bool(_ss_get("prev_ascvd", False))
    edad = int(_ss_get("prev_edad", 55) or 55)
    if ascvd or edad < 30 or edad > 79:
        st.warning("PREVENT se utiliza para prevencion primaria en adultos de 30 a 79 anos sin ASCVD clinica establecida.")
    else:
        st.success("Variables basicas listas para cargar/calcular PREVENT en prevencion primaria.")

    st.caption("Por seguridad metodologica, los porcentajes PREVENT 10 y 30 anos se guardan como resultado oficial transcripto, pero las variables de entrada ya no se duplican: quedan sincronizadas en toda la app.")

# =========================================================
# UI - Aplicacion principal
# =========================================================
def render_user_bar():
    user = st.session_state.user_data
    nombre = user.get("nombre", st.session_state.username)
    rol = user.get("rol", "medico")
    cols = st.columns([6, 1])
    with cols[0]:
        st.markdown(f'<div class="user-bar"><span>Sesion: {nombre} ({st.session_state.username}) - Rol: {rol}</span><span>{date.today().isoformat()}</span></div>', unsafe_allow_html=True)
    with cols[1]:
        if st.button("Cerrar sesion"):
            for k in ["authenticated", "username", "user_data"]:
                st.session_state.pop(k, None)
            st.rerun()

def render_evaluacion():
    with st.sidebar:
        st.header("Ingreso clinico")
        paciente = st.text_input("Paciente", "")
        dni = st.text_input("DNI / ID", "")
        # auto-relleno
        u = st.session_state.user_data
        medico = st.text_input("Medico", u.get("nombre", "") or "")
        matricula = st.text_input("Matricula", u.get("matricula", "") or "")
        st.subheader("Demograficos")
        edad = st.number_input("Edad", 18, 100, 55, key="prev_edad")
        sexo = st.selectbox("Sexo", ["Masculino", "Femenino"], key="prev_sexo")
        st.subheader("Lipidos")
        colesterol_total = st.number_input("Colesterol total (mg/dL)", 0.0, 600.0, 220.0, 1.0, key="prev_colesterol_total")
        hdl = st.number_input("HDL-C (mg/dL)", 0.0, 150.0, 45.0, 1.0, key="prev_hdl")
        tg = st.number_input("Trigliceridos (mg/dL)", 0.0, 1000.0, 150.0, 1.0)
        ldl_basal = st.number_input("LDL-C basal/pretratamiento (mg/dL)", 0.0, 500.0, 160.0, 1.0)
        ldl_actual = st.number_input("LDL-C actual (mg/dL)", 0.0, 500.0, 120.0, 1.0)
        no_hdl_calc = max(colesterol_total - hdl, 0)
        no_hdl = st.number_input("No-HDL-C (mg/dL)", 0.0, 600.0, float(no_hdl_calc), 1.0)
        mide_lpa = st.checkbox("Lp(a) disponible")
        lpa_valor, lpa_unidad = None, "nmol/L"
        if mide_lpa:
            lpa_unidad = st.selectbox("Unidad Lp(a)", ["nmol/L", "mg/dL"])
            lpa_valor = st.number_input("Lp(a)", 0.0, 600.0, 100.0, 1.0)
        mide_apob = st.checkbox("ApoB disponible")
        apob = st.number_input("ApoB (mg/dL)", 0.0, 300.0, 100.0, 1.0) if mide_apob else None
        st.subheader("Comorbilidades y potenciadores")
        diabetes = st.checkbox("Diabetes", key="prev_diabetes")
        ckd = st.checkbox("Enfermedad renal cronica")
        egfr = st.number_input("eGFR ml/min/1.73m2 para PREVENT", 0.0, 150.0, 75.0, 1.0, key="prev_egfr")
        hta = st.checkbox("Hipertension arterial", key="prev_hta")
        presion_sistolica = st.number_input("Presion sistolica / PAS para PREVENT (mmHg)", 70.0, 260.0, 130.0, 1.0, key="prev_pas")
        tratamiento_hta = st.checkbox("Tratamiento antihipertensivo", key="prev_tratamiento_hta")
        tabaquismo = st.checkbox("Tabaquismo activo", key="prev_tabaquismo")
        inflamacion_cronica = st.checkbox("Inflamacion cronica")
        antecedente_familiar = st.checkbox("ASCVD prematura familiar")
        menopausia_precoz = st.checkbox("Menopausia precoz") if sexo == "Femenino" else False
        preeclampsia = st.checkbox("Antecedente de preeclampsia") if sexo == "Femenino" else False
        fh_sospecha = st.checkbox("Sospecha de hipercolesterolemia familiar")
        st.subheader("Historia cardiovascular")
        ascvd = st.checkbox("ASCVD clinica establecida", key="prev_ascvd")
        iam = st.checkbox("IAM previo") if ascvd else False
        acv = st.checkbox("ACV/AIT previo") if ascvd else False
        pad = st.checkbox("Enfermedad arterial periferica") if ascvd else False
        revascularizacion = st.checkbox("Revascularizacion previa") if ascvd else False
        st.subheader("Imagen y PREVENT")
        st.caption("Las variables requeridas para PREVENT se toman automaticamente de la barra vertical. Solo transcriba el resultado final si usa la calculadora oficial.")
        tiene_cac = st.checkbox("CAC disponible")
        cac = st.number_input("CAC Agatston", 0, 5000, 0, 1) if tiene_cac else None
        if ascvd or edad < 30 or edad > 79:
            st.warning("PREVENT esta disenado para prevencion primaria en adultos 30-79 anos sin ASCVD clinica.")
        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            prev10 = st.number_input("PREVENT 10 anos (%)", 0.0, 100.0, 0.0, 0.1)
        with col_prev2:
            prev30 = st.number_input("PREVENT 30 anos (%)", 0.0, 100.0, 0.0, 0.1)
        prevent_10 = prev10 if prev10 > 0 else None
        prevent_30 = prev30 if prev30 > 0 else None
        with st.expander("Ver variables PREVENT sincronizadas", expanded=False):
            render_prevent_sync_panel(show_editor=False)
        st.link_button("Abrir calculadora PREVENT oficial AHA", PREVENT_URL)
        st.subheader("Medicacion actual")
        estatina = st.selectbox("Estatina", ["Ninguna", "Atorvastatina", "Rosuvastatina", "Simvastatina", "Pravastatina", "Otra"], key="prev_estatina")
        dosis_estatina = st.text_input("Dosis de estatina", "")
        ezetimibe = st.checkbox("Ezetimibe")
        pcsk9 = st.checkbox("PCSK9 mAb")
        inclisiran = st.checkbox("Inclisiran")
        bempedoico = st.checkbox("Acido bempedoico")
        intolerancia_sams = st.checkbox("Intolerancia/SAMS")
        observaciones = st.text_area("Observaciones", "")

    p = Patient(paciente, dni, medico, matricula, edad, sexo, ldl_basal, ldl_actual, hdl, tg,
                colesterol_total, no_hdl, lpa_valor, lpa_unidad, apob, diabetes, ckd, egfr, presion_sistolica, hta,
                tratamiento_hta, tabaquismo, inflamacion_cronica, antecedente_familiar, menopausia_precoz, preeclampsia,
                ascvd, iam, acv, pad, revascularizacion, fh_sospecha, cac, prevent_10, prevent_30,
                estatina, dosis_estatina, ezetimibe, pcsk9, inclisiran, bempedoico, intolerancia_sams,
                observaciones)
    info = determinar_perfil(p)
    metas = metas_lipidicas(p)
    estado = estado_meta(p)
    decision = decidir_tratamiento_primaria(p)
    pots = potenciadores_riesgo(p)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        color = "red" if "Muy alto" in info["riesgo"] else "orange" if "Alto" in info["riesgo"] else "blue"
        resumen_card("Perfil de riesgo", info["perfil"], info["riesgo"], color, "Clasificacion clinica global")
    with c2:
        prev_txt = "No aplica" if p.prevent_10 is None else f"{p.prevent_10:.1f}%"
        prev_color, prev_cat, prev_ref = clasificar_prevent_color(p.prevent_10)
        resumen_card("Riesgo PREVENT 10 anos", prev_txt, prev_cat, prev_color, f"Referencia: {prev_ref}")
    with c3:
        resumen_card("LDL-C actual / meta", f"{p.ldl_actual:.0f} mg/dL", str(estado["texto"]),
                     str(estado["color"]), f"Meta recomendada: {metas['ldl']}")
    with c4:
        red = estado["reduccion"]
        red_txt = f"{red}%" if isinstance(red, float) else str(red)
        resumen_card("Reduccion LDL-C", red_txt, "Objetivo " + metas["reduccion"], "blue",
                     f"No-HDL-C meta: {metas['no_hdl']}")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Nota clinica", "Decision farmacologica", "PREVENT", "Plan completo", "Exportar / guardar"])

    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Nota clinica estructurada")
        st.markdown(f"**Paciente:** {p.paciente or 'No informado'}")
        st.markdown(f"**Perfil de riesgo:** {info['perfil']} - **{info['riesgo']}**.")
        if not p.ascvd:
            st.markdown(f"**Riesgo PREVENT:** 10 anos {p.prevent_10 if p.prevent_10 is not None else 'No informado'}% ({clasificar_prevent(p.prevent_10)}); 30 anos {p.prevent_30 if p.prevent_30 is not None else 'No informado'}%.")
        st.markdown(f"**Metas recomendadas:** LDL-C **{metas['ldl']}**, no-HDL-C **{metas['no_hdl']}**, reduccion **{metas['reduccion']}**.")
        st.markdown(f"**Estado actual:** LDL-C {p.ldl_actual:.0f} mg/dL, no-HDL-C {p.no_hdl:.0f} mg/dL, TG {p.tg:.0f} mg/dL.")
        st.markdown("**Potenciadores/reclasificadores:** " + (", ".join(pots) if pots else "no registrados"))
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        mostrar_panel_bioquimico(p, metas)
        st.markdown('</div>', unsafe_allow_html=True)
        if estado["texto"] == "En meta":
            st.markdown('<div class="alert-green">Paciente en meta lipidica. Sostener adherencia, estilo de vida y seguimiento.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="alert-red">Paciente fuera o cerca de meta. Revisar adherencia, intensidad de estatina y necesidad de terapia combinada.</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="rx-card">', unsafe_allow_html=True)
        st.markdown('<div class="rx-title">Decision farmacologica - PREVENCION PRIMARIA / SECUNDARIA</div>', unsafe_allow_html=True)
        if decision["requiere_farmaco"]:
            st.markdown(f'<div class="alert-red"><strong>SI requiere tratamiento farmacologico.</strong> Intensidad: {decision["intensidad"]}.</div>', unsafe_allow_html=True)
        else:
            color_alert = "green" if decision.get("color") == "green" else "orange"
            cls = "alert-green" if color_alert == "green" else "alert-orange"
            st.markdown(f'<div class="{cls}"><strong>NO requiere tratamiento farmacologico de inicio.</strong> {decision["intensidad"]}.</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="rx-drug">Droga sugerida: {decision["droga_principal"]}</div>', unsafe_allow_html=True)
        if "complementaria" in decision:
            st.markdown(f'<div class="rx-drug">Complemento: {decision["complementaria"]}</div>', unsafe_allow_html=True)
        st.markdown(f"**Justificacion:** {decision['justificacion']}")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Algoritmo de decision aplicado")
        st.markdown("""
- **Prevencion secundaria (ASCVD)**: estatina alta intensidad + escalado segun meta LDL.
- **LDL >=190 mg/dL o FH**: estatina alta intensidad sin necesidad de calcular riesgo.
- **Diabetes 40-75 anos con potenciadores o edad >=50**: estatina alta intensidad.
- **CKD eGFR 30-59**: estatina moderada-alta (preferir atorvastatina).
- **Edad <40 sin factores graves**: estilo de vida, reevaluar 3-5 anos.
- **PREVENT >=20%**: alta intensidad.
- **PREVENT 7.5-19.9%**: moderada/alta segun potenciadores y CAC. CAC=0 sin DM/FH/tabaquismo permite diferir.
- **PREVENT 5-7.4%**: estatina moderada si potenciadores o CAC>=100; sino estilo de vida.
- **PREVENT <5%**: estilo de vida.
        """)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Calculadora PREVENT oficial AHA")
        st.write("Algunos sitios oficiales bloquean el embebido. Use el boton para abrir en navegador.")
        st.link_button("Abrir PREVENT oficial", PREVENT_URL)
        try:
            components.iframe(PREVENT_URL, height=900, scrolling=True)
        except Exception as e:
            st.warning(f"No se pudo incrustar. Detalle: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tab4:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Plan farmacologico completo")
        plan = plan_farmacologico_completo(p)
        for item in plan:
            st.markdown(f"- {item}")
        st.markdown(f"**CAC:** {recomendaciones_cac(p)}")
        st.markdown('</div>', unsafe_allow_html=True)

    with tab5:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Exportar / guardar")
        nota = nota_clinica(p)
        row = make_row(p, decision)
        df_row = pd.DataFrame([row])

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Informe medico**")
            st.download_button("Descargar TXT (medico)", data=nota.encode("utf-8"),
                               file_name="informe_medico.txt", mime="text/plain")
            try:
                pdf_bytes = pdf_informe_medico(p)
                st.download_button("Descargar PDF graficado (medico)", data=pdf_bytes,
                                   file_name="informe_medico_lipidcare.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error generando PDF medico: {e}")
        with col_b:
            st.markdown("**Informe para paciente**")
            try:
                pdf_pac = pdf_informe_paciente(p)
                st.download_button("Descargar PDF graficado (paciente)", data=pdf_pac,
                                   file_name="informe_paciente_lipidcare.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error generando PDF paciente: {e}")

        st.markdown("---")
        st.markdown("**Datos del paciente actual (formato tabla)**")
        if EXCEL_ENGINE is not None:
            try:
                st.download_button("Descargar Excel del registro actual",
                                   data=excel_bytes_from_df(df_row, "Registro_actual"),
                                   file_name="registro_actual.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Excel: {e}")
        st.download_button("Descargar CSV del registro actual",
                           data=df_row.to_csv(index=False).encode("utf-8-sig"),
                           file_name="registro_actual.csv", mime="text/csv")

        st.markdown("---")
        if st.button("Guardar paciente en mi historial"):
            add_paciente_historial(st.session_state.username, row)
            st.success("Paciente guardado en su historial personal.")
        st.markdown('</div>', unsafe_allow_html=True)

def render_historial_propio():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Mi historial de pacientes")
    historial = load_historial()
    registros = historial.get(st.session_state.username, [])
    if not registros:
        st.info("Aun no tiene pacientes guardados. En 'Evaluacion clinica' use el boton 'Guardar paciente en mi historial'.")
        st.markdown('</div>', unsafe_allow_html=True)
        return
    df = pd.DataFrame(registros)
    st.caption(f"Total de evaluaciones guardadas: **{len(df)}**")
    st.dataframe(df, use_container_width=True)
    cols = st.columns(3)
    with cols[0]:
        if EXCEL_ENGINE is not None:
            try:
                st.download_button("Descargar mi historial Excel",
                                   data=excel_bytes_from_df(df, "Mi_historial"),
                                   file_name=f"historial_{st.session_state.username}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Excel: {e}")
        else:
            st.warning("Excel no disponible. Use CSV.")
    with cols[1]:
        st.download_button("Descargar mi historial CSV",
                           data=df.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"historial_{st.session_state.username}.csv", mime="text/csv")
    with cols[2]:
        if st.button("Borrar todo mi historial"):
            borrar_historial_usuario(st.session_state.username)
            st.warning("Historial borrado.")
            st.rerun()

    st.markdown("---")
    st.subheader("Borrar evaluacion individual")
    idx = st.number_input("Indice de fila a borrar (0-based)", 0, max(len(df) - 1, 0), 0)
    if st.button("Borrar fila seleccionada"):
        borrar_paciente(st.session_state.username, int(idx))
        st.warning(f"Fila {idx} borrada.")
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

def render_evolucion_paciente():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Evolucion temporal del paciente")
    st.caption("Seleccione un paciente para ver la evolucion del LDL-C y demas indicadores a lo largo del tiempo. Cada vez que guarda una nueva evaluacion del mismo paciente (mismo DNI), se agrega un punto a la serie.")
    pacientes = listar_pacientes_unicos(st.session_state.username)
    if not pacientes:
        st.info("Aun no tiene pacientes guardados. Guarde evaluaciones desde 'Evaluacion clinica'.")
        st.markdown('</div>', unsafe_allow_html=True)
        return
    # Resumen de pacientes
    df_resumen = pd.DataFrame([{
        "Paciente": p["paciente"] or "(sin nombre)",
        "DNI": p["dni"] or "(sin DNI)",
        "Evaluaciones": p["n"],
        "Primera": str(p["primera"])[:10],
        "Ultima": str(p["ultima"])[:10],
        "LDL inicial": p["ldl_inicial"],
        "LDL ultimo": p["ldl_ultimo"],
        "Cambio %": p["delta_pct"]
    } for p in pacientes])
    st.markdown("**Pacientes registrados:**")
    st.dataframe(df_resumen, use_container_width=True)

    # Selector
    opciones_disp = [f"{p['paciente'] or '(sin nombre)'} - DNI {p['dni'] or '-'} - {p['n']} eval." for p in pacientes]
    sel_idx = st.selectbox("Seleccione paciente para ver su evolucion:",
                           options=list(range(len(pacientes))),
                           format_func=lambda i: opciones_disp[i])
    g = pacientes[sel_idx]
    evals = g["evaluaciones"]
    df = evolucion_dataframe(evals)

    # Cards resumen
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        resumen_card("Evaluaciones", str(g["n"]), "Serie temporal", "blue", "Total guardadas")
    with c2:
        if g["ldl_inicial"] is not None:
            resumen_card("LDL-C inicial", f"{g['ldl_inicial']:.0f} mg/dL", "Primer registro", "gray", str(g["primera"])[:10])
    with c3:
        if g["ldl_ultimo"] is not None:
            color = "green" if (g["delta_pct"] or 0) >= 30 else "orange" if (g["delta_pct"] or 0) > 0 else "red"
            resumen_card("LDL-C ultimo", f"{g['ldl_ultimo']:.0f} mg/dL", "Ultimo registro", color, str(g["ultima"])[:10])
    with c4:
        if g["delta_pct"] is not None:
            color = "green" if g["delta_pct"] >= 50 else "orange" if g["delta_pct"] >= 30 else "red" if g["delta_pct"] >= 0 else "red"
            signo = "Reduccion" if g["delta_pct"] >= 0 else "Aumento"
            resumen_card("Cambio LDL-C", f"{abs(g['delta_pct']):.1f}%", signo, color, "Inicial vs ultimo")

    # Grafico de evolucion con altair (incluido en streamlit)
    if g["n"] >= 1:
        try:
            import altair as alt
            df_chart = df.copy()
            if "fecha" in df_chart.columns:
                df_chart["fecha"] = pd.to_datetime(df_chart["fecha"], errors="coerce")
            # Series para grafico
            series_map = {"LDL actual": "ldl_actual", "LDL basal": "ldl_basal",
                          "No-HDL-C": "no_hdl", "Trigliceridos": "tg", "HDL-C": "hdl"}
            seleccion = st.multiselect("Series a graficar:",
                                        options=list(series_map.keys()),
                                        default=["LDL actual", "LDL basal"])
            cols_keep = ["fecha"] + [series_map[k] for k in seleccion if series_map[k] in df_chart.columns]
            if "fecha" in cols_keep and len(cols_keep) > 1:
                long_df = df_chart[cols_keep].melt("fecha", var_name="Indicador", value_name="mg/dL")
                # remap nombres
                inv_map = {v: k for k, v in series_map.items()}
                long_df["Indicador"] = long_df["Indicador"].map(inv_map).fillna(long_df["Indicador"])
                # Meta horizontal: tomar de la ultima evaluacion
                meta_str = evals[-1].get("meta_ldl", "")
                meta_val = None
                for token in ["55", "70", "100", "116"]:
                    if token in str(meta_str):
                        meta_val = float(token); break
                base = alt.Chart(long_df).mark_line(point=alt.OverlayMarkDef(size=80, filled=True)).encode(
                    x=alt.X("fecha:T", title="Fecha de evaluacion"),
                    y=alt.Y("mg/dL:Q", title="mg/dL"),
                    color=alt.Color("Indicador:N",
                                    scale=alt.Scale(domain=list(series_map.keys()),
                                                     range=["#0B4F8A", "#94A3B8", "#7C3AED", "#EA580C", "#16A34A"])),
                    tooltip=["fecha:T", "Indicador:N", "mg/dL:Q"]
                ).properties(height=380, title=f"Evolucion temporal - {g['paciente']}")
                if meta_val is not None and "LDL actual" in seleccion:
                    rule = alt.Chart(pd.DataFrame({"y": [meta_val]})).mark_rule(
                        color="#16A34A", strokeDash=[6, 4], size=2
                    ).encode(y="y:Q")
                    text = alt.Chart(pd.DataFrame({"y": [meta_val], "label": [f"Meta LDL <{meta_val:.0f}"]})).mark_text(
                        align="left", dx=8, dy=-6, color="#16A34A", fontSize=11, fontWeight="bold"
                    ).encode(y="y:Q", text="label:N")
                    chart = (base + rule + text).resolve_scale(y="shared")
                else:
                    chart = base
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("Seleccione al menos una serie ademas de la fecha.")
        except ImportError:
            # fallback simple
            cols_show = [c for c in ["ldl_actual", "ldl_basal", "no_hdl", "tg", "hdl"] if c in df.columns]
            if "fecha" in df.columns and cols_show:
                tmp = df[["fecha"] + cols_show].copy()
                tmp = tmp.set_index("fecha")
                st.line_chart(tmp)

    # Tabla
    st.markdown("**Detalle de evaluaciones:**")
    st.dataframe(df, use_container_width=True)

    # Descargas
    cols_dl = st.columns(3)
    with cols_dl[0]:
        try:
            pdf_bytes = pdf_evolucion_paciente(st.session_state.username, g["key"])
            if pdf_bytes:
                fname = f"evolucion_{(g['dni'] or g['paciente'] or 'paciente').replace(' ','_')[:20]}.pdf"
                st.download_button("Descargar PDF de evolucion (con grafico)",
                                   data=pdf_bytes, file_name=fname, mime="application/pdf")
        except Exception as e:
            st.error(f"Error generando PDF: {e}")
    with cols_dl[1]:
        if EXCEL_ENGINE is not None:
            try:
                xlsx_b = excel_bytes_from_df(df, "Evolucion")
                st.download_button("Descargar evolucion Excel",
                                   data=xlsx_b,
                                   file_name=f"evolucion_{(g['dni'] or 'paciente')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Excel: {e}")
    with cols_dl[2]:
        st.download_button("Descargar evolucion CSV",
                           data=df.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"evolucion_{(g['dni'] or 'paciente')}.csv", mime="text/csv")
    st.markdown('</div>', unsafe_allow_html=True)

def render_admin_global():
    if st.session_state.user_data.get("rol") != "admin":
        st.error("Acceso restringido. Esta seccion es solo para administradores.")
        return
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Panel global de administracion")
    users = load_users()
    historial = load_historial()
    resumen_rows = []
    total_pac = 0
    for u, data in users.items():
        cant = len(historial.get(u, []))
        total_pac += cant
        resumen_rows.append({
            "usuario": u,
            "nombre": data.get("nombre", ""),
            "matricula": data.get("matricula", ""),
            "especialidad": data.get("especialidad", ""),
            "rol": data.get("rol", ""),
            "creado": data.get("creado", ""),
            "pacientes_guardados": cant
        })
    df_users = pd.DataFrame(resumen_rows)
    st.markdown(f"**Total de usuarios registrados:** {len(users)}    **Total de pacientes:** {total_pac}")
    st.dataframe(df_users, use_container_width=True)

    st.markdown("---")
    st.subheader("Exportar TODOS los pacientes de TODOS los usuarios")
    if EXCEL_ENGINE is not None:
        try:
            st.download_button("Descargar Excel multiusuario completo",
                               data=excel_historial_multiusuario_bytes(),
                               file_name="lipidcare_TODOS_los_usuarios.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel multiusuario: {e}")
    # CSV consolidado
    all_rows = []
    for u, regs in historial.items():
        for r in regs:
            row = dict(r)
            row["usuario_login"] = u
            all_rows.append(row)
    if all_rows:
        df_all = pd.DataFrame(all_rows)
        st.download_button("Descargar CSV consolidado (todos los pacientes)",
                           data=df_all.to_csv(index=False).encode("utf-8-sig"),
                           file_name="lipidcare_consolidado.csv", mime="text/csv")
        st.caption("Vista previa del consolidado:")
        st.dataframe(df_all.head(50), use_container_width=True)
    else:
        st.info("Aun no hay pacientes registrados por ningun usuario.")
    st.markdown('</div>', unsafe_allow_html=True)

def render_calculadora_prevent():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Calculadora PREVENT oficial AHA")
    st.markdown("""
Este modulo queda sincronizado con la barra vertical de **Evaluacion clinica**.
Las variables PREVENT se cargan una sola vez y se reutilizan para el calculo, el informe, la semaforizacion y la decision farmacologica.
""")
    render_prevent_sync_panel(show_editor=True)
    st.link_button("Abrir calculadora PREVENT oficial AHA", PREVENT_URL)
    try:
        components.iframe(PREVENT_URL, height=950, scrolling=True)
    except Exception as e:
        st.warning(f"No se pudo embeber. Detalle: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

def render_ayuda():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Como correr la app")
    st.markdown("""
Guarde este archivo como `app.py`. `requirements.txt` minimo:

```
streamlit>=1.28
pandas>=2.0
openpyxl>=3.1.2
xlsxwriter>=3.1.9
```

Ejecutar:
```
pip install -r requirements.txt
streamlit run app.py
```

**Datos persistentes**: se guardan en la carpeta `.lipidcare_data/` (usuarios y pacientes).
En Streamlit Cloud el filesystem es efimero; para produccion use una BBDD externa.

**Login por defecto del administrador:** usuario `admin`, contrasena `admin1234` - cambiela en produccion.

**Mejoras incluidas en esta version:**
- Sistema de login y registro con contrasena hasheada (PBKDF2-SHA256, 120k iteraciones).
- Historial PERSISTENTE por usuario en disco; cada medico ve solo sus pacientes.
- Exportacion total multiusuario (solo admin) en Excel y CSV consolidado.
- Algoritmo claro de decision farmacologica para prevencion primaria con droga y dosis sugerida.
- PDFs medicos y para paciente CON GRAFICOS SEMAFORIZADOS (barras de meta LDL, barra PREVENT, cajas coloreadas por indicador).
- Variables PREVENT sincronizadas con la barra vertical: edad, sexo, colesterol total, HDL, PAS, diabetes, tabaquismo, eGFR, tratamiento antihipertensivo y estatina.
""")
    st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# Router principal
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    render_login()
else:
    render_user_bar()
    st.markdown(f'''<div class="hero"><h1>🫀 {APP_NAME}</h1><p>Estratificacion de riesgo, decision farmacologica clara, metas lipidicas, semaforos y PDF didacticos.</p><p style="margin-top:8px;font-weight:800;">Autor: {AUTOR_APP}</p></div>''', unsafe_allow_html=True)
    with st.expander("Diagnostico tecnico"):
        st.write(f"Motor PDF: **{PDF_ENGINE}** (interno con graficos coloreados)")
        st.write(f"Motor Excel: **{EXCEL_ENGINE or 'no disponible'}**")
        if EXCEL_IMPORT_ERROR:
            st.caption(f"Detalle Excel: {EXCEL_IMPORT_ERROR}")
        st.write(f"Carpeta de datos: `{DATA_DIR.resolve()}`")

    opciones = ["Evaluacion clinica", "Mi historial", "Evolucion del paciente", "Calculadora PREVENT", "Ayuda"]
    if st.session_state.user_data.get("rol") == "admin":
        opciones.append("Admin: todos los usuarios")
    modo = st.sidebar.radio("Modulo", opciones)

    if modo == "Evaluacion clinica":
        render_evaluacion()
    elif modo == "Mi historial":
        render_historial_propio()
    elif modo == "Evolucion del paciente":
        render_evolucion_paciente()
    elif modo == "Calculadora PREVENT":
        render_calculadora_prevent()
    elif modo == "Admin: todos los usuarios":
        render_admin_global()
    elif modo == "Ayuda":
        render_ayuda()

st.markdown('<div class="footer">Uso profesional. Herramienta de soporte a la decision clinica. No sustituye el juicio medico ni las guias locales/regulatorias.</div>', unsafe_allow_html=True)
