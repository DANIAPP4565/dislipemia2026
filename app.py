from __future__ import annotations

import streamlit as st
import pandas as pd
import io
import textwrap
import zlib
from dataclasses import dataclass
from typing import Optional, Dict, List
from datetime import date, datetime
import streamlit.components.v1 as components

APP_NAME = "LipidCare 2026 Pro"
AUTOR_APP = "Ricardo Daniel Olano, Especialista en Cardiología y en Hipertensión Arterial"
PREVENT_URL = "https://professional.heart.org/en/guidelines-and-statements/prevent-calculator"

# =========================================================
# Motores PDF y Excel
# =========================================================
PDF_ENGINE = "interno_sin_dependencias"
PDF_IMPORT_ERROR = ""
try:
    from fpdf import FPDF  # opcional: si existe, se usa; si no, se usa motor interno
    PDF_ENGINE = "fpdf2"
except Exception as e:
    FPDF = None
    PDF_IMPORT_ERROR = repr(e)

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
# Estilo visual con contraste corregido
# =========================================================
st.markdown('''
<style>
html, body, [class*="css"] { color:#111827 !important; }
.main {background:#F8FAFC;}
.block-container {padding-top:1rem; padding-bottom:2rem;}
section[data-testid="stSidebar"] { background:#F1F5F9 !important; color:#111827 !important; }
section[data-testid="stSidebar"] * { color:#111827 !important; }
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] select { background:#FFFFFF !important; color:#111827 !important; }
.hero {
  background: linear-gradient(135deg,#0B4F8A 0%,#123C69 55%,#0F766E 100%);
  padding:28px 34px; border-radius:26px; color:white !important;
  box-shadow:0 14px 34px rgba(11,79,138,.25); margin-bottom:18px;
}
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
</style>
''', unsafe_allow_html=True)

# =========================================================
# Utilidades
# =========================================================
def safe_text(txt) -> str:
    if txt is None:
        return ""
    txt = str(txt)
    replacements = {
        "≥": ">=", "≤": "<=", "–": "-", "—": "-", "“": '"', "”": '"', "’": "'",
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N", "°": " grados ", "²": "2", "ü": "u", "Ü": "U"
    }
    for a, b in replacements.items():
        txt = txt.replace(a, b)
    return txt


def pct_reduccion(ldl_basal: float, ldl_actual: float) -> Optional[float]:
    if ldl_basal and ldl_basal > 0 and ldl_actual >= 0:
        return round((ldl_basal - ldl_actual) / ldl_basal * 100, 1)
    return None


def lpa_alta(valor: Optional[float], unidad: str) -> bool:
    if valor is None:
        return False
    return valor >= 50 if unidad == "mg/dL" else valor >= 125


def clasificar_prevent(riesgo: Optional[float]) -> str:
    if riesgo is None:
        return "No informado"
    if riesgo < 3:
        return "Bajo"
    if riesgo < 5:
        return "Limítrofe"
    if riesgo < 10:
        return "Intermedio"
    return "Alto"


def badge_html(texto: str, color: str = "blue") -> str:
    cls = {"green":"badge-green", "yellow":"badge-yellow", "orange":"badge-orange", "red":"badge-red", "blue":"badge-blue", "gray":"badge-gray"}.get(color, "badge-blue")
    return f'<span class="badge {cls}">{texto}</span>'


def resumen_card(titulo: str, valor: str, badge_texto: str = "", color: str = "blue", caption: str = ""):
    badge = badge_html(badge_texto, color) if badge_texto else ""
    st.markdown(f'''<div class="summary-card"><div class="summary-title">{titulo}</div><div class="summary-value">{valor}</div>{badge}<div class="summary-caption">{caption}</div></div>''', unsafe_allow_html=True)

# =========================================================
# PDF interno sin dependencias
# =========================================================
def _pdf_escape_text(text: str) -> str:
    text = safe_text(text)
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def generar_pdf_interno(texto: str, title: str = APP_NAME) -> bytes:
    width, height = 595, 842
    margin_x = 42
    y_start = 800
    line_height = 13
    font_size = 9
    pages, current = [], []
    y = y_start

    def add_page():
        nonlocal current, y
        if current:
            pages.append(current)
        current = []
        y = y_start

    current.append(f"BT /F1 15 Tf {margin_x} {y} Td ({_pdf_escape_text(title)}) Tj ET")
    y -= 20
    current.append(f"BT /F1 9 Tf {margin_x} {y} Td ({_pdf_escape_text(AUTOR_APP)}) Tj ET")
    y -= 22

    for raw in texto.splitlines():
        raw = safe_text(raw)
        wrapped = [""] if raw.strip() == "" else (textwrap.wrap(raw, width=92, replace_whitespace=False) or [""])
        for line in wrapped:
            if y < 45:
                add_page()
                current.append(f"BT /F1 10 Tf {margin_x} {y} Td ({_pdf_escape_text(APP_NAME)}) Tj ET")
                y -= 20
            size = 10 if line.isupper() and len(line) < 60 else font_size
            current.append(f"BT /F1 {size} Tf {margin_x} {y} Td ({_pdf_escape_text(line)}) Tj ET")
            y -= line_height
    if current:
        pages.append(current)

    kids_refs, page_content_objects = [], []
    next_obj_num = 4
    for page_lines in pages:
        content = "\n".join(page_lines).encode("latin-1", errors="replace")
        compressed = zlib.compress(content)
        content_obj_num = next_obj_num + 1
        page_obj = f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << /Font << /F1 3 0 R >> >> /Contents {content_obj_num} 0 R >>".encode()
        content_obj = b"<< /Length " + str(len(compressed)).encode() + b" /Filter /FlateDecode >>\nstream\n" + compressed + b"\nendstream"
        page_content_objects.extend([page_obj, content_obj])
        kids_refs.append(f"{next_obj_num} 0 R")
        next_obj_num += 2

    catalog_obj = b"<< /Type /Catalog /Pages 2 0 R >>"
    pages_obj = f"<< /Type /Pages /Kids [{' '.join(kids_refs)}] /Count {len(pages)} >>".encode()
    font_obj = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    final_objects = [catalog_obj, pages_obj, font_obj] + page_content_objects

    pdf = bytearray()
    pdf.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for i, obj in enumerate(final_objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{i} 0 obj\n".encode())
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")
    xref_pos = len(pdf)
    pdf.extend(f"xref\n0 {len(final_objects)+1}\n".encode())
    pdf.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf.extend(f"{off:010d} 00000 n \n".encode())
    pdf.extend(f"trailer\n<< /Size {len(final_objects)+1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF".encode())
    return bytes(pdf)


def generar_pdf_bytes(texto: str) -> bytes:
    if PDF_ENGINE == "fpdf2" and FPDF is not None:
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, safe_text(APP_NAME), ln=True)
            pdf.set_font("Arial", "", 9)
            pdf.multi_cell(0, 5, safe_text(AUTOR_APP))
            pdf.ln(4)
            pdf.set_font("Arial", "", 10)
            for line in safe_text(texto).split("\n"):
                pdf.multi_cell(0, 5, line)
            return pdf.output(dest="S").encode("latin-1", errors="replace")
        except Exception:
            pass
    return generar_pdf_interno(texto)

# =========================================================
# Modelo y lógica clínica
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
    hta: bool
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


def parse_meta_ldl(meta_txt: str) -> Optional[float]:
    if "55" in meta_txt: return 55
    if "70" in meta_txt: return 70
    if "100" in meta_txt: return 100
    return None


def determinar_perfil(p: Patient) -> Dict[str, str]:
    eventos = sum([p.iam, p.acv, p.pad, p.revascularizacion])
    biologico_alto = p.diabetes or p.ckd or p.fh_sospecha
    if p.ascvd:
        perfil = "Prevención secundaria"
        riesgo = "Muy alto riesgo" if eventos >= 1 or biologico_alto else "Alto riesgo secundario"
    else:
        perfil = "Prevención primaria"
        cat = clasificar_prevent(p.prevent_10)
        if p.fh_sospecha or p.diabetes or p.ckd or cat == "Alto": riesgo = "Alto riesgo primario"
        elif cat == "Intermedio": riesgo = "Riesgo intermedio"
        elif cat == "Limítrofe": riesgo = "Riesgo limítrofe"
        else: riesgo = "Bajo riesgo"
    return {"perfil": perfil, "riesgo": riesgo}


def metas_lipidicas(p: Patient) -> Dict[str, str]:
    info = determinar_perfil(p)
    if info["perfil"] == "Prevención secundaria" and "Muy alto" in info["riesgo"]:
        return {"ldl":"<55 mg/dL", "no_hdl":"<85 mg/dL", "reduccion":">=50%"}
    if info["perfil"] == "Prevención secundaria":
        return {"ldl":"<70 mg/dL", "no_hdl":"<100 mg/dL", "reduccion":">=50%"}
    if "Alto" in info["riesgo"]:
        return {"ldl":"<70 mg/dL", "no_hdl":"<100 mg/dL", "reduccion":">=50%"}
    if "Intermedio" in info["riesgo"]:
        return {"ldl":"<100 mg/dL", "no_hdl":"Individualizar", "reduccion":"30-49% o mayor si potenciadores"}
    if "limítrofe" in info["riesgo"].lower():
        return {"ldl":"<100 mg/dL", "no_hdl":"Individualizar", "reduccion":"según potenciadores/CAC"}
    return {"ldl":"<100 mg/dL", "no_hdl":"Individualizar", "reduccion":"estilo de vida"}


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


def clasificar_ldl_vs_meta(ldl: float, meta_txt: str):
    meta = parse_meta_ldl(meta_txt)
    if meta is None: return "gray", "Individualizar", "según perfil clínico"
    if ldl < meta: return "green", "En meta", f"<{meta:.0f} mg/dL"
    if ldl <= meta + 20: return "orange", "Cerca de meta", f"<{meta:.0f} mg/dL"
    return "red", "Fuera de meta", f"<{meta:.0f} mg/dL"


def clasificar_no_hdl(no_hdl: float, meta_txt: str):
    if "85" in meta_txt: meta = 85
    elif "100" in meta_txt: meta = 100
    else: return "gray", "Individualizar", "según perfil clínico"
    if no_hdl < meta: return "green", "En meta", f"<{meta} mg/dL"
    if no_hdl <= meta + 20: return "orange", "Cerca de meta", f"<{meta} mg/dL"
    return "red", "Fuera de meta", f"<{meta} mg/dL"


def clasificar_tg(tg: float):
    if tg < 150: return "green", "Normal", "<150 mg/dL"
    if tg < 175: return "yellow", "Limítrofe", "150-174 mg/dL"
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
    if valor is None: return "gray", "No medida", "útil si TG altos/diabetes/CKD"
    if valor < 90: return "green", "Óptima", "<90 mg/dL"
    if valor < 130: return "orange", "Elevada", "90-129 mg/dL"
    return "red", "Muy elevada", ">=130 mg/dL"


def clasificar_egfr(valor: Optional[float]):
    if valor is None: return "gray", "No aplica", "si CKD, cargar eGFR"
    if valor >= 60: return "green", "Preservado", ">=60"
    if valor >= 30: return "orange", "Disminuido", "30-59"
    return "red", "Muy disminuido", "<30"


def clasificar_cac_valor(valor: Optional[int]):
    if valor is None: return "gray", "No disponible", "usar si duda clínica"
    if valor == 0: return "green", "CAC 0", "0"
    if valor < 100: return "orange", "CAC positivo", "1-99"
    return "red", "CAC alto", ">=100"


def clasificar_prevent_color(riesgo: Optional[float]):
    cat = clasificar_prevent(riesgo)
    if cat == "Bajo": return "green", cat, "<3%"
    if cat == "Limítrofe": return "yellow", cat, "3-<5%"
    if cat == "Intermedio": return "orange", cat, "5-<10%"
    if cat == "Alto": return "red", cat, ">=10%"
    return "gray", cat, "30-79 años"


def potenciadores_riesgo(p: Patient) -> List[str]:
    out = []
    if lpa_alta(p.lpa_valor, p.lpa_unidad): out.append(f"Lp(a) elevada ({p.lpa_valor:.0f} {p.lpa_unidad})")
    if p.apob is not None and p.apob >= 130: out.append(f"ApoB elevada ({p.apob:.0f} mg/dL)")
    if p.diabetes: out.append("diabetes")
    if p.ckd: out.append("enfermedad renal crónica")
    if p.hta: out.append("hipertensión arterial")
    if p.tabaquismo: out.append("tabaquismo activo")
    if p.inflamacion_cronica: out.append("inflamación crónica")
    if p.antecedente_familiar: out.append("antecedente familiar de ASCVD prematura")
    if p.menopausia_precoz: out.append("menopausia precoz")
    if p.preeclampsia: out.append("antecedente de preeclampsia")
    if p.fh_sospecha: out.append("sospecha de hipercolesterolemia familiar")
    if p.cac is not None and p.cac > 0: out.append(f"CAC positivo ({p.cac})")
    if p.tg >= 175: out.append(f"hipertrigliceridemia ({p.tg:.0f} mg/dL)")
    return out


def recomendaciones_cac(p: Patient) -> str:
    if p.cac is None: return "CAC no disponible. Usar si hay incertidumbre en prevención primaria limítrofe/intermedia."
    if p.cac == 0: return "CAC = 0: puede apoyar diferir o reducir intensidad en escenarios seleccionados, excepto diabetes, tabaquismo, FH, ASCVD familiar prematura u otros riesgos mayores."
    if 1 <= p.cac < 100: return "CAC 1-99: evidencia aterosclerótica subclínica; favorece estatina, especialmente con edad >55 años o potenciadores."
    return "CAC >=100: favorece estatina e intensificación para alcanzar meta de LDL-C."


def plan_farmacologico(p: Patient) -> List[str]:
    info = determinar_perfil(p)
    estado = estado_meta(p)
    plan = []
    if p.intolerancia_sams:
        plan.append("Confirmar SAMS: evaluar temporalidad, CK si corresponde, interacciones, hipotiroidismo/deficiencia de vitamina D; realizar pausa y reexposición con otra estatina, dosis baja o días alternos.")
        plan.append("Usar ezetimibe 10 mg/día como primer no estatínico. Considerar ácido bempedoico si persiste LDL-C sobre meta o hay intolerancia significativa.")
    else:
        if info["perfil"] == "Prevención secundaria" or "Alto" in info["riesgo"]:
            plan.append("Indicar o sostener estatina de alta intensidad si es tolerada: atorvastatina 40-80 mg o rosuvastatina 20-40 mg.")
        elif "intermedio" in info["riesgo"].lower():
            plan.append("Considerar estatina de moderada a alta intensidad según PREVENT, potenciadores de riesgo, CAC y decisión compartida.")
        elif "limítrofe" in info["riesgo"].lower():
            plan.append("Considerar estatina si hay potenciadores de riesgo, Lp(a) elevada, ApoB elevada, antecedente familiar o CAC positivo.")
        else:
            plan.append("Priorizar intervención intensiva de estilo de vida; reevaluar riesgo y lípidos periódicamente.")
    if estado["texto"] != "En meta":
        if not p.ezetimibe:
            plan.append("Escalado 1: agregar ezetimibe 10 mg/día.")
        elif not (p.pcsk9 or p.inclisiran) and (info["perfil"] == "Prevención secundaria" or "Alto" in info["riesgo"]):
            plan.append("Escalado 2: si persiste sobre meta con estatina + ezetimibe, considerar inhibidor PCSK9 monoclonal. Inclisiran puede considerarse si se prioriza adherencia semestral.")
        if not p.bempedoico:
            plan.append("Ácido bempedoico: opción en intolerancia a estatinas o necesidad adicional de reducción de LDL-C.")
    if p.lpa_valor is None:
        plan.append("Solicitar Lp(a) al menos una vez en la vida.")
    elif lpa_alta(p.lpa_valor, p.lpa_unidad):
        plan.append("Lp(a) elevada: intensificar control de LDL-C y de todos los factores de riesgo; en prevención secundaria considerar PCSK9 si no alcanza meta.")
    if p.apob is None and (p.diabetes or p.ckd or p.tg >= 150):
        plan.append("Solicitar ApoB para evaluar riesgo residual por discordancia con LDL-C, especialmente en diabetes, CKD o hipertrigliceridemia.")
    if not p.ascvd and p.cac is None and clasificar_prevent(p.prevent_10) in ["Limítrofe", "Intermedio"]:
        plan.append("Considerar CAC para reclasificación si hay duda clínica sobre iniciar o intensificar estatinas.")
    return plan


def semaforo_item(nombre: str, valor, unidad: str, color: str, interpretacion: str, referencia: str):
    if valor is None: valor_txt = "No informado"
    elif isinstance(valor, float): valor_txt = f"{valor:.1f} {unidad}"
    else: valor_txt = f"{valor} {unidad}"
    st.markdown(f'''<div class="semaforo-card"><div class="semaforo-title">{nombre}</div><div class="semaforo-value">{valor_txt}</div>{badge_html(interpretacion, color)}<div class="semaforo-ref">Referencia: {referencia}</div></div>''', unsafe_allow_html=True)


def semaforo_items_data(p: Patient, metas: Dict[str, str]) -> List[Dict[str, str]]:
    base = [
        ("LDL-C actual", p.ldl_actual, "mg/dL", *clasificar_ldl_vs_meta(p.ldl_actual, metas["ldl"])),
        ("No-HDL-C", p.no_hdl, "mg/dL", *clasificar_no_hdl(p.no_hdl, metas["no_hdl"])),
        ("Triglicéridos", p.tg, "mg/dL", *clasificar_tg(p.tg)),
        ("HDL-C", p.hdl, "mg/dL", *clasificar_hdl(p.hdl, p.sexo)),
        ("Lp(a)", p.lpa_valor, p.lpa_unidad, *clasificar_lpa(p.lpa_valor, p.lpa_unidad)),
        ("ApoB", p.apob, "mg/dL", *clasificar_apob(p.apob)),
        ("eGFR", p.egfr, "ml/min/1.73m2", *clasificar_egfr(p.egfr)),
        ("CAC", p.cac, "Agatston", *clasificar_cac_valor(p.cac)),
    ]
    if p.prevent_10 is not None: base.append(("PREVENT 10 años", p.prevent_10, "%", *clasificar_prevent_color(p.prevent_10)))
    if p.prevent_30 is not None:
        color30 = "green" if p.prevent_30 < 15 else "orange" if p.prevent_30 < 30 else "red"
        interp30 = "Bajo largo plazo" if p.prevent_30 < 15 else "Intermedio largo plazo" if p.prevent_30 < 30 else "Alto largo plazo"
        base.append(("PREVENT 30 años", p.prevent_30, "%", color30, interp30, "orientativo"))
    items = []
    for nombre, valor, unidad, color, interp, ref in base:
        if valor is None: valor_txt = "No informado"
        elif isinstance(valor, float): valor_txt = f"{valor:.1f} {unidad}"
        else: valor_txt = f"{valor} {unidad}"
        items.append({"indicador": nombre, "valor": valor_txt, "color": color, "interpretacion": interp, "referencia": ref})
    return items


def mostrar_panel_bioquimico(p: Patient, metas: Dict[str, str]):
    st.subheader("Semaforización bioquímica y de riesgo")
    cols = st.columns(4)
    items = []
    items.append(("LDL-C actual", p.ldl_actual, "mg/dL", *clasificar_ldl_vs_meta(p.ldl_actual, metas["ldl"])))
    items.append(("No-HDL-C", p.no_hdl, "mg/dL", *clasificar_no_hdl(p.no_hdl, metas["no_hdl"])))
    items.append(("Triglicéridos", p.tg, "mg/dL", *clasificar_tg(p.tg)))
    items.append(("HDL-C", p.hdl, "mg/dL", *clasificar_hdl(p.hdl, p.sexo)))
    items.append(("Lp(a)", p.lpa_valor, p.lpa_unidad, *clasificar_lpa(p.lpa_valor, p.lpa_unidad)))
    items.append(("ApoB", p.apob, "mg/dL", *clasificar_apob(p.apob)))
    items.append(("eGFR", p.egfr, "ml/min/1.73m²", *clasificar_egfr(p.egfr)))
    items.append(("CAC", p.cac, "Agatston", *clasificar_cac_valor(p.cac)))
    if p.prevent_10 is not None: items.append(("PREVENT 10 años", p.prevent_10, "%", *clasificar_prevent_color(p.prevent_10)))
    if p.prevent_30 is not None:
        color30 = "green" if p.prevent_30 < 15 else "orange" if p.prevent_30 < 30 else "red"
        interp30 = "Bajo largo plazo" if p.prevent_30 < 15 else "Intermedio largo plazo" if p.prevent_30 < 30 else "Alto largo plazo"
        items.append(("PREVENT 30 años", p.prevent_30, "%", color30, interp30, "orientativo"))
    for i, it in enumerate(items):
        with cols[i % 4]:
            nombre, valor, unidad, color, interp, ref = it
            semaforo_item(nombre, valor, unidad, color, interp, ref)


def nota_clinica(p: Patient) -> str:
    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    pots, plan = potenciadores_riesgo(p), plan_farmacologico(p)
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
Riesgo PREVENT 10 años: {p.prevent_10 if p.prevent_10 is not None else 'No aplica/no informado'}% ({clasificar_prevent(p.prevent_10)}).
Riesgo PREVENT 30 años: {p.prevent_30 if p.prevent_30 is not None else 'No informado'}%.

LIPIDOS
Colesterol total: {p.colesterol_total:.0f} mg/dL.
LDL-C basal: {p.ldl_basal:.0f} mg/dL.
LDL-C actual: {p.ldl_actual:.0f} mg/dL.
Reducción LDL-C: {red if isinstance(red, str) else str(red) + '%'}.
HDL-C: {p.hdl:.0f} mg/dL.
No-HDL-C: {p.no_hdl:.0f} mg/dL.
Triglicéridos: {p.tg:.0f} mg/dL.
Lp(a): {str(p.lpa_valor) + ' ' + p.lpa_unidad if p.lpa_valor is not None else 'No informada'}.
ApoB: {str(p.apob) + ' mg/dL' if p.apob is not None else 'No informada'}.

METAS RECOMENDADAS
LDL-C: {metas['ldl']}.
No-HDL-C: {metas['no_hdl']}.
Reducción recomendada: {metas['reduccion']}.
Estado actual: {estado['texto']}.

SEMAFORIZACION BIOQUIMICA Y DE RIESGO
{semaforo_txt}

POTENCIADORES / RECLASIFICADORES
{', '.join(pots) if pots else 'No registrados'}.
CAC: {recomendaciones_cac(p)}

PLAN FARMACOLOGICO SUGERIDO
""" + "\n".join([f"- {x}" for x in plan]) + f"""

SEGUIMIENTO
Repetir perfil lipidico en 4-12 semanas luego de inicio o ajuste terapeutico. Luego controlar cada 3-12 meses segun estabilidad, adherencia y distancia a meta.

OBSERVACIONES
{p.observaciones or 'Sin observaciones adicionales.'}

Aviso: herramienta de soporte a la decision clinica. No sustituye juicio medico ni guias/regulaciones locales.
"""


def nota_paciente(p: Patient) -> str:
    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    return f"""
{APP_NAME}
Informe para paciente
Fecha: {date.today().isoformat()}

Paciente: {p.paciente or 'No informado'}

RESUMEN
Su perfil corresponde a: {perfil['perfil']} - {perfil['riesgo']}.
Su LDL actual es: {p.ldl_actual:.0f} mg/dL.
La meta sugerida para LDL es: {metas['ldl']}.
Estado actual: {estado['texto']}.

QUE SIGNIFICA
El LDL-C es una de las principales particulas relacionadas con aterosclerosis. Cuanto mayor es el riesgo cardiovascular, mas baja debe ser la meta de LDL-C.

PLAN GENERAL
- Mantener alimentacion cardioprotectora.
- Realizar actividad fisica segun indicacion medica.
- Evitar tabaco.
- Controlar presion arterial, peso, glucemia y otros factores de riesgo.
- Cumplir la medicacion indicada.
- Repetir laboratorio en 4 a 12 semanas si hubo inicio o cambio de tratamiento.

Este informe simplificado no reemplaza la explicacion medica personalizada.
"""


def make_row(p: Patient) -> Dict[str, object]:
    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    return {
        "fecha": datetime.now().isoformat(timespec="seconds"), "paciente": p.paciente, "dni": p.dni,
        "medico": p.medico, "matricula": p.matricula, "edad": p.edad, "sexo": p.sexo,
        "perfil": perfil["perfil"], "riesgo": perfil["riesgo"], "prevent_10": p.prevent_10, "prevent_30": p.prevent_30,
        "colesterol_total": p.colesterol_total, "hdl": p.hdl, "ldl_basal": p.ldl_basal, "ldl_actual": p.ldl_actual,
        "reduccion_ldl": estado["reduccion"], "meta_ldl": metas["ldl"], "estado_meta": estado["texto"],
        "no_hdl": p.no_hdl, "tg": p.tg, "lpa": p.lpa_valor, "lpa_unidad": p.lpa_unidad, "apob": p.apob,
        "diabetes": p.diabetes, "ckd": p.ckd, "egfr": p.egfr, "hta": p.hta, "tabaquismo": p.tabaquismo,
        "inflamacion_cronica": p.inflamacion_cronica, "antecedente_familiar": p.antecedente_familiar,
        "menopausia_precoz": p.menopausia_precoz, "preeclampsia": p.preeclampsia,
        "ascvd": p.ascvd, "iam": p.iam, "acv": p.acv, "pad": p.pad, "revascularizacion": p.revascularizacion,
        "fh_sospecha": p.fh_sospecha, "cac": p.cac, "estatina": p.estatina, "dosis_estatina": p.dosis_estatina,
        "ezetimibe": p.ezetimibe, "pcsk9": p.pcsk9, "inclisiran": p.inclisiran, "bempedoico": p.bempedoico,
        "sams": p.intolerancia_sams, "observaciones": p.observaciones,
    }


def get_user_key(medico: str, matricula: str) -> str:
    return f"{(medico or 'Usuario sin nombre').strip()} - {(matricula or 'Sin matricula').strip()}"


def asegurar_historial_usuario(user_key: str):
    if "historial_por_usuario" not in st.session_state:
        st.session_state.historial_por_usuario = {}
    if user_key not in st.session_state.historial_por_usuario:
        st.session_state.historial_por_usuario[user_key] = []


def dataframe_historial_usuario(user_key: str) -> pd.DataFrame:
    asegurar_historial_usuario(user_key)
    return pd.DataFrame(st.session_state.historial_por_usuario[user_key])


def excel_bytes_from_df(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    if EXCEL_ENGINE is None:
        raise RuntimeError("Motor Excel no disponible. Instale openpyxl o xlsxwriter desde requirements.txt.")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine=EXCEL_ENGINE) as writer:
        clean_sheet = sheet_name[:31] or "Datos"
        df.to_excel(writer, index=False, sheet_name=clean_sheet)
        worksheet = writer.sheets[clean_sheet]
        for i, col in enumerate(df.columns):
            width = min(max(len(str(col)) + 2, 12), 42)
            if EXCEL_ENGINE == "xlsxwriter":
                worksheet.set_column(i, i, width)
            else:
                from openpyxl.utils import get_column_letter
                worksheet.column_dimensions[get_column_letter(i + 1)].width = width
    return output.getvalue()


def excel_historial_multiusuario_bytes() -> bytes:
    if EXCEL_ENGINE is None:
        raise RuntimeError("Motor Excel no disponible.")
    historiales = st.session_state.get("historial_por_usuario", {})
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine=EXCEL_ENGINE) as writer:
        resumen = []
        for usuario, registros in historiales.items():
            if not registros: continue
            df = pd.DataFrame(registros)
            resumen.append({"usuario": usuario, "cantidad_pacientes/evaluaciones": len(df)})
            nombre_hoja = usuario.replace("/", "-").replace("\\", "-").replace("*", "-").replace("?", "-").replace(":", "-")[:31]
            df.to_excel(writer, index=False, sheet_name=nombre_hoja or "Usuario")
        if resumen:
            pd.DataFrame(resumen).to_excel(writer, index=False, sheet_name="Resumen")
        else:
            pd.DataFrame(columns=["usuario", "cantidad_pacientes/evaluaciones"]).to_excel(writer, index=False, sheet_name="Resumen")
    return output.getvalue()

# =========================================================
# Interfaz
# =========================================================
st.markdown(f'''<div class="hero"><h1>🫀 {APP_NAME}</h1><p>Herramienta clínica profesional para estratificación de riesgo, metas lipídicas, escalado terapéutico y seguimiento.</p><p style="margin-top:8px;font-weight:800;">Autor: {AUTOR_APP}</p></div>''', unsafe_allow_html=True)

with st.expander("Diagnóstico técnico de PDF y Excel"):
    st.write(f"Motor PDF: **{PDF_ENGINE}**")
    if PDF_IMPORT_ERROR:
        st.caption(f"fpdf2 no disponible: {PDF_IMPORT_ERROR}. La app usará PDF interno sin dependencias.")
    st.write(f"Motor Excel: **{EXCEL_ENGINE or 'no disponible'}**")
    if EXCEL_IMPORT_ERROR:
        st.caption(f"Detalle Excel: {EXCEL_IMPORT_ERROR}")
    st.info("Si Excel figura como no disponible, la app seguirá funcionando con CSV. Para Excel, requirements.txt debe incluir openpyxl y xlsxwriter.")

modo = st.sidebar.radio("Módulo", ["Evaluación clínica", "Calculadora PREVENT oficial", "Historial por usuario", "Ayuda de instalación"])

if modo == "Evaluación clínica":
    with st.sidebar:
        st.header("Ingreso clínico")
        paciente = st.text_input("Paciente", "")
        dni = st.text_input("DNI / ID", "")
        medico = st.text_input("Médico", "Dr. Ricardo Daniel Olano")
        matricula = st.text_input("Matrícula", "")
        st.subheader("Demográficos")
        edad = st.number_input("Edad", 18, 100, 55)
        sexo = st.selectbox("Sexo", ["Masculino", "Femenino"])
        st.subheader("Lípidos")
        colesterol_total = st.number_input("Colesterol total (mg/dL)", 0.0, 600.0, 220.0, 1.0)
        hdl = st.number_input("HDL-C (mg/dL)", 0.0, 150.0, 45.0, 1.0)
        tg = st.number_input("Triglicéridos (mg/dL)", 0.0, 1000.0, 150.0, 1.0)
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
        diabetes = st.checkbox("Diabetes")
        ckd = st.checkbox("Enfermedad renal crónica")
        egfr = st.number_input("eGFR ml/min/1.73m²", 0.0, 150.0, 75.0, 1.0) if ckd else None
        hta = st.checkbox("Hipertensión arterial")
        tabaquismo = st.checkbox("Tabaquismo activo")
        inflamacion_cronica = st.checkbox("Inflamación crónica")
        antecedente_familiar = st.checkbox("ASCVD prematura familiar")
        menopausia_precoz = st.checkbox("Menopausia precoz") if sexo == "Femenino" else False
        preeclampsia = st.checkbox("Antecedente de preeclampsia") if sexo == "Femenino" else False
        fh_sospecha = st.checkbox("Sospecha de hipercolesterolemia familiar")
        st.subheader("Historia cardiovascular")
        ascvd = st.checkbox("ASCVD clínica establecida")
        iam = st.checkbox("IAM previo") if ascvd else False
        acv = st.checkbox("ACV/AIT previo") if ascvd else False
        pad = st.checkbox("Enfermedad arterial periférica") if ascvd else False
        revascularizacion = st.checkbox("Revascularización previa") if ascvd else False
        st.subheader("Imagen y PREVENT")
        st.caption("Calcule PREVENT en la web oficial y transcriba aquí el resultado para incorporarlo al informe, PDF, Excel y CSV.")
        tiene_cac = st.checkbox("CAC disponible")
        cac = st.number_input("CAC Agatston", 0, 5000, 0, 1) if tiene_cac else None
        if ascvd or edad < 30 or edad > 79:
            st.warning("PREVENT está diseñado para prevención primaria en adultos de 30 a 79 años sin ASCVD clínica. Puede dejar los valores en 0 o cargarlos solo como dato orientativo.")
        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            prevent_10_input = st.number_input("PREVENT calculado 10 años (%)", 0.0, 100.0, 0.0, 0.1, key="prevent_10_manual")
        with col_prev2:
            prevent_30_input = st.number_input("PREVENT calculado 30 años (%)", 0.0, 100.0, 0.0, 0.1, key="prevent_30_manual")
        prevent_10 = prevent_10_input if prevent_10_input > 0 else None
        prevent_30 = prevent_30_input if prevent_30_input > 0 else None
        st.link_button("Abrir calculadora PREVENT oficial AHA", PREVENT_URL)
        st.subheader("Medicación actual")
        estatina = st.selectbox("Estatina", ["Ninguna", "Atorvastatina", "Rosuvastatina", "Simvastatina", "Pravastatina", "Otra"])
        dosis_estatina = st.text_input("Dosis de estatina", "")
        ezetimibe = st.checkbox("Ezetimibe")
        pcsk9 = st.checkbox("PCSK9 mAb")
        inclisiran = st.checkbox("Inclisirán")
        bempedoico = st.checkbox("Ácido bempedoico")
        intolerancia_sams = st.checkbox("Intolerancia/SAMS")
        observaciones = st.text_area("Observaciones", "")

    p = Patient(paciente, dni, medico, matricula, edad, sexo, ldl_basal, ldl_actual, hdl, tg, colesterol_total, no_hdl, lpa_valor, lpa_unidad, apob, diabetes, ckd, egfr, hta, tabaquismo, inflamacion_cronica, antecedente_familiar, menopausia_precoz, preeclampsia, ascvd, iam, acv, pad, revascularizacion, fh_sospecha, cac, prevent_10, prevent_30, estatina, dosis_estatina, ezetimibe, pcsk9, inclisiran, bempedoico, intolerancia_sams, observaciones)

    perfil, metas, estado = determinar_perfil(p), metas_lipidicas(p), estado_meta(p)
    pots, nota, nota_pac = potenciadores_riesgo(p), nota_clinica(p), nota_paciente(p)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        color = "red" if "Muy alto" in perfil["riesgo"] else "orange" if "Alto" in perfil["riesgo"] else "blue"
        resumen_card("Perfil de riesgo", perfil["perfil"], perfil["riesgo"], color, "Clasificación clínica global")
    with c2:
        prev_txt = "No aplica" if p.prevent_10 is None else f"{p.prevent_10:.1f}%"
        prev_color, prev_cat, prev_ref = clasificar_prevent_color(p.prevent_10)
        resumen_card("Riesgo PREVENT 10 años", prev_txt, prev_cat, prev_color, f"Referencia: {prev_ref}")
    with c3:
        resumen_card("LDL-C actual / meta", f"{p.ldl_actual:.0f} mg/dL", str(estado["texto"]), str(estado["color"]), f"Meta recomendada: {metas['ldl']}")
    with c4:
        red = estado["reduccion"]
        red_txt = f"{red}%" if isinstance(red, float) else str(red)
        resumen_card("Reducción LDL-C", red_txt, "Objetivo " + metas["reduccion"], "blue", f"No-HDL-C meta: {metas['no_hdl']}")

    tab1, tab2, tab3, tab4 = st.tabs(["Nota clínica", "CPR", "Tratamiento", "Exportar / guardar"])
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Nota clínica estructurada")
        st.markdown(f"**Paciente:** {p.paciente or 'No informado'}")
        st.markdown(f"**Perfil de riesgo:** {perfil['perfil']} - **{perfil['riesgo']}**.")
        if not p.ascvd:
            st.markdown(f"**Riesgo PREVENT:** 10 años {p.prevent_10 if p.prevent_10 is not None else 'No informado'}% ({clasificar_prevent(p.prevent_10)}); 30 años {p.prevent_30 if p.prevent_30 is not None else 'No informado'}%.")
        st.markdown(f"**Metas recomendadas:** LDL-C **{metas['ldl']}**, no-HDL-C **{metas['no_hdl']}**, reducción **{metas['reduccion']}**.")
        st.markdown(f"**Estado actual:** LDL-C {p.ldl_actual:.0f} mg/dL, no-HDL-C {p.no_hdl:.0f} mg/dL, TG {p.tg:.0f} mg/dL.")
        st.markdown("**Potenciadores/reclasificadores:** " + (", ".join(pots) if pots else "no registrados"))
        st.markdown("**Seguimiento:** reevaluar lípidos en 4-12 semanas luego de inicio o ajuste; luego cada 3-12 meses.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card">', unsafe_allow_html=True)
        mostrar_panel_bioquimico(p, metas)
        st.markdown('</div>', unsafe_allow_html=True)
        if estado["texto"] == "En meta":
            st.markdown('<div class="alert-green">Paciente en meta lipídica según el perfil ingresado. Sostener adherencia, estilo de vida y seguimiento.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="alert-red">Paciente fuera o cerca de meta. Revisar adherencia, intensidad de estatina y necesidad de terapia combinada.</div>', unsafe_allow_html=True)
    with tab2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Calculadora PREVENT oficial AHA")
        st.write("Use el botón para abrir la calculadora oficial. Algunas páginas oficiales bloquean la visualización embebida.")
        st.link_button("Abrir PREVENT oficial en navegador", PREVENT_URL)
        try:
            components.iframe(PREVENT_URL, height=900, scrolling=True)
        except Exception as e:
            st.warning(f"No se pudo incrustar la calculadora oficial. Abrir con el botón externo. Detalle: {e}")
        st.markdown('</div>', unsafe_allow_html=True)
    with tab3:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Plan farmacológico sugerido")
        for item in plan_farmacologico(p):
            st.markdown(f"- {item}")
        st.markdown('</div>', unsafe_allow_html=True)
    with tab4:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Exportar / guardar")
        row = make_row(p)
        df_row = pd.DataFrame([row])
        user_key = get_user_key(p.medico, p.matricula)
        asegurar_historial_usuario(user_key)
        col_a, col_b = st.columns(2)
        with col_a:
            st.download_button("Descargar informe médico TXT", data=nota.encode("utf-8"), file_name="informe_medico_lipidcare_2026.txt", mime="text/plain")
            st.download_button("Descargar informe médico PDF", data=generar_pdf_bytes(nota), file_name="informe_medico_lipidcare_2026.pdf", mime="application/pdf")
        with col_b:
            st.download_button("Descargar informe para paciente TXT", data=nota_pac.encode("utf-8"), file_name="informe_paciente_lipidcare_2026.txt", mime="text/plain")
            st.download_button("Descargar informe para paciente PDF", data=generar_pdf_bytes(nota_pac), file_name="informe_paciente_lipidcare_2026.pdf", mime="application/pdf")
        st.markdown("**Formato de datos:** variables en columnas y pacientes/evaluaciones en filas.")
        if EXCEL_ENGINE is not None:
            try:
                st.download_button("Descargar registro actual en Excel", data=excel_bytes_from_df(df_row, "Registro_actual"), file_name="registro_lipidcare_actual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"No se pudo generar Excel: {repr(e)}")
        else:
            st.warning("Excel no disponible en este entorno. Se habilita CSV. Para Excel, requirements.txt debe incluir openpyxl y xlsxwriter.")
        st.download_button("Descargar registro actual CSV", data=df_row.to_csv(index=False).encode("utf-8-sig"), file_name="registro_lipidcare_actual.csv", mime="text/csv")
        if st.button("Guardar paciente en historial del usuario"):
            st.session_state.historial_por_usuario[user_key].append(row)
            st.success(f"Paciente guardado en el historial de: {user_key}")
        df_hist_usuario = dataframe_historial_usuario(user_key)
        if not df_hist_usuario.empty:
            st.dataframe(df_hist_usuario, use_container_width=True)
            if EXCEL_ENGINE is not None:
                try:
                    st.download_button("Descargar historial Excel de este usuario", data=excel_bytes_from_df(df_hist_usuario, "Historial_usuario"), file_name="historial_lipidcare_usuario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"No se pudo generar historial Excel: {repr(e)}")
            st.download_button("Descargar historial CSV de este usuario", data=df_hist_usuario.to_csv(index=False).encode("utf-8-sig"), file_name="historial_lipidcare_usuario.csv", mime="text/csv")
        if st.session_state.get("historial_por_usuario") and EXCEL_ENGINE is not None:
            try:
                st.download_button("Descargar Excel completo multiusuario", data=excel_historial_multiusuario_bytes(), file_name="historial_lipidcare_multiusuario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"No se pudo generar Excel multiusuario: {repr(e)}")
        st.markdown('</div>', unsafe_allow_html=True)

elif modo == "Calculadora PREVENT oficial":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Calculadora PREVENT oficial AHA")
    st.markdown("""
Esta sección permite abrir la calculadora oficial PREVENT de la American Heart Association.

**Uso recomendado:**
1. Calcular el riesgo en la web oficial.
2. Volver a Evaluación clínica.
3. Transcribir PREVENT 10 años y 30 años.
4. Exportar informes PDF, Excel o CSV.
""")
    st.link_button("Abrir calculadora PREVENT oficial AHA", PREVENT_URL)
    try:
        components.iframe(PREVENT_URL, height=950, scrolling=True)
    except Exception as e:
        st.warning(f"No se pudo mostrar embebida. Use el botón superior. Detalle: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

elif modo == "Historial por usuario":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Historial por usuario/médico")
    historiales = st.session_state.get("historial_por_usuario", {})
    if not historiales:
        st.info("Aún no hay pacientes guardados. En Evaluación clínica use el botón: Guardar paciente en historial del usuario.")
    else:
        usuarios = list(historiales.keys())
        usuario_sel = st.selectbox("Seleccionar usuario", usuarios)
        df = pd.DataFrame(historiales[usuario_sel])
        st.caption("Formato: variables en columnas, pacientes/evaluaciones en filas.")
        st.dataframe(df, use_container_width=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            if EXCEL_ENGINE is not None:
                try:
                    st.download_button("Descargar Excel de este usuario", data=excel_bytes_from_df(df, "Historial_usuario"), file_name="historial_lipidcare_usuario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"No se pudo generar Excel: {repr(e)}")
            else:
                st.warning("Excel no disponible. Use CSV.")
        with c2:
            st.download_button("Descargar CSV de este usuario", data=df.to_csv(index=False).encode("utf-8-sig"), file_name="historial_lipidcare_usuario.csv", mime="text/csv")
        with c3:
            if st.button("Borrar historial del usuario seleccionado"):
                st.session_state.historial_por_usuario[usuario_sel] = []
                st.success("Historial del usuario seleccionado borrado.")
        if EXCEL_ENGINE is not None:
            try:
                st.download_button("Descargar Excel multiusuario", data=excel_historial_multiusuario_bytes(), file_name="historial_lipidcare_multiusuario.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"No se pudo generar Excel multiusuario: {repr(e)}")
    st.markdown('</div>', unsafe_allow_html=True)

elif modo == "Ayuda de instalación":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Cómo correr la app")
    st.markdown("""
Guardar este archivo como `app.py`.

Crear un archivo `requirements.txt` en la misma carpeta con:

```txt
streamlit>=1.28
pandas>=2.0
openpyxl>=3.1.2
xlsxwriter>=3.1.9
```

Esta versión **NO necesita `fpdf` ni `fpdf2`** para generar PDF.

Ejecutar localmente:

```bash
pip install -r requirements.txt
streamlit run app.py
```

En Streamlit Cloud, la estructura debe ser:

```txt
app.py
requirements.txt
```
""")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="footer">Uso profesional. Herramienta de soporte a la decisión clínica. No sustituye el juicio médico ni las guías locales/regulatorias.</div>', unsafe_allow_html=True)
