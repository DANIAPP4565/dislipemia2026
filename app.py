import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ============================================================
# IMPORT SEGURO DE PYPREVENT
# ============================================================

try:
    import pyprevent
    PYPREVENT_DISPONIBLE = True
except ModuleNotFoundError:
    pyprevent = None
    PYPREVENT_DISPONIBLE = False
except Exception:
    pyprevent = None
    PYPREVENT_DISPONIBLE = False


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Riesgo Cardiovascular Integrado",
    page_icon="❤️",
    layout="wide"
)


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>
    .main {
        background-color: #f8fafc;
    }

    .titulo-principal {
        background: linear-gradient(90deg, #0f172a, #1e3a8a);
        color: white;
        padding: 22px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 20px;
    }

    .card {
        background-color: white;
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 6px solid #1e3a8a;
    }

    .riesgo-bajo {
        background-color: #dcfce7;
        color: #14532d;
        padding: 12px;
        border-radius: 10px;
        font-weight: bold;
    }

    .riesgo-limitrofe {
        background-color: #fef9c3;
        color: #713f12;
        padding: 12px;
        border-radius: 10px;
        font-weight: bold;
    }

    .riesgo-intermedio {
        background-color: #fed7aa;
        color: #7c2d12;
        padding: 12px;
        border-radius: 10px;
        font-weight: bold;
    }

    .riesgo-alto {
        background-color: #fee2e2;
        color: #7f1d1d;
        padding: 12px;
        border-radius: 10px;
        font-weight: bold;
    }

    .riesgo-nocalculable {
        background-color: #e5e7eb;
        color: #111827;
        padding: 12px;
        border-radius: 10px;
        font-weight: bold;
    }

    .subtitulo {
        font-size: 22px;
        font-weight: 700;
        color: #0f172a;
        margin-top: 18px;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNCIONES UTILITARIAS
# ============================================================

def normalizar_si_no(valor):
    if isinstance(valor, bool):
        return valor
    if valor in ["Sí", "SI", "Si", "sí", "S", "s", "1", 1]:
        return True
    return False


def normalizar_sexo_para_prevent(sexo):
    """
    Intenta transformar la entrada de sexo a formatos compatibles.
    """
    if sexo == "Masculino":
        return "male"
    if sexo == "Femenino":
        return "female"
    return sexo


def safe_float(valor):
    try:
        if valor is None:
            return None
        return float(valor)
    except Exception:
        return None


def formatear_porcentaje(valor):
    if valor is None:
        return "No calculable"
    try:
        return f"{float(valor):.1f} %"
    except Exception:
        return "No calculable"


def clasificar_riesgo_prevent(valor):
    """
    Clasificación práctica usada para semaforización preventiva.
    """
    if valor is None:
        return "No calculable"

    try:
        valor = float(valor)
    except Exception:
        return "No calculable"

    if valor < 3:
        return "Riesgo bajo"
    elif valor < 5:
        return "Riesgo limítrofe"
    elif valor < 10:
        return "Riesgo intermedio"
    else:
        return "Riesgo alto"


def clase_css_riesgo(clasificacion):
    if clasificacion == "Riesgo bajo":
        return "riesgo-bajo"
    if clasificacion == "Riesgo limítrofe":
        return "riesgo-limitrofe"
    if clasificacion == "Riesgo intermedio":
        return "riesgo-intermedio"
    if clasificacion == "Riesgo alto":
        return "riesgo-alto"
    return "riesgo-nocalculable"


def interpretar_presion_arterial(pas, pad):
    pas = safe_float(pas)
    pad = safe_float(pad)

    if pas is None or pad is None:
        return "Presión arterial no evaluable"

    if pas < 120 and pad < 80:
        return "Presión arterial normal"
    elif 120 <= pas < 130 and pad < 80:
        return "Presión arterial elevada"
    elif 130 <= pas < 140 or 80 <= pad < 90:
        return "Hipertensión arterial grado 1"
    elif pas >= 140 or pad >= 90:
        return "Hipertensión arterial grado 2"
    else:
        return "Presión arterial no clasificable"


def interpretar_ldl(ldl):
    ldl = safe_float(ldl)

    if ldl is None:
        return "LDL no evaluable"

    if ldl < 70:
        return "LDL en rango intensivo"
    elif ldl < 100:
        return "LDL en rango aceptable"
    elif ldl < 130:
        return "LDL moderadamente elevado"
    elif ldl < 160:
        return "LDL elevado"
    elif ldl < 190:
        return "LDL muy elevado"
    else:
        return "LDL severamente elevado"


# ============================================================
# CÁLCULO PREVENT SEGURO
# ============================================================

def extraer_numero_de_resultado(obj, posibles_claves):
    """
    Busca valores en dict, objeto o tupla/lista.
    Esto permite tolerar pequeñas diferencias de API de pyprevent.
    """
    if obj is None:
        return None

    if isinstance(obj, dict):
        for clave in posibles_claves:
            if clave in obj:
                return safe_float(obj[clave])

    for clave in posibles_claves:
        if hasattr(obj, clave):
            return safe_float(getattr(obj, clave))

    return None


def intentar_calculo_prevent_con_pyprevent(datos):
    """
    Intenta ejecutar pyprevent usando varias firmas posibles.
    La librería puede cambiar nombres internos; por eso este bloque es defensivo.
    """

    edad = safe_float(datos.get("edad"))
    sexo = normalizar_sexo_para_prevent(datos.get("sexo"))
    colesterol_total = safe_float(datos.get("colesterol_total"))
    hdl = safe_float(datos.get("hdl"))
    pas = safe_float(datos.get("presion_sistolica"))
    diabetes = normalizar_si_no(datos.get("diabetes"))
    tabaquismo = normalizar_si_no(datos.get("tabaquismo"))
    tratamiento_hta = normalizar_si_no(datos.get("tratamiento_hta"))
    imc = safe_float(datos.get("imc"))
    egfr = safe_float(datos.get("egfr"))

    kwargs_base = {
        "age": edad,
        "sex": sexo,
        "total_cholesterol": colesterol_total,
        "hdl_cholesterol": hdl,
        "systolic_bp": pas,
        "diabetes": diabetes,
        "smoking": tabaquismo,
        "antihypertensive_treatment": tratamiento_hta,
        "bmi": imc,
        "egfr": egfr,
    }

    kwargs_alternativos = {
        "age": edad,
        "sex": sexo,
        "tc": colesterol_total,
        "hdl": hdl,
        "sbp": pas,
        "diabetes": diabetes,
        "smoker": tabaquismo,
        "anti_hypertensive": tratamiento_hta,
        "bmi": imc,
        "egfr": egfr,
    }

    funciones_posibles = [
        "calculate",
        "prevent",
        "calc_prevent",
        "calculate_prevent",
        "prevent_risk",
        "estimate_risk",
    ]

    ultimo_error = None

    for nombre_funcion in funciones_posibles:
        if hasattr(pyprevent, nombre_funcion):
            funcion = getattr(pyprevent, nombre_funcion)

            for kwargs in [kwargs_base, kwargs_alternativos]:
                try:
                    resultado = funcion(**kwargs)
                    return resultado
                except Exception as e:
                    ultimo_error = e

    # Si el paquete expone submódulos o clases, intentamos introspección simple
    try:
        for atributo in dir(pyprevent):
            if atributo.startswith("_"):
                continue

            obj = getattr(pyprevent, atributo)

            if callable(obj):
                nombre = atributo.lower()
                if "prevent" in nombre or "risk" in nombre or "calculate" in nombre:
                    for kwargs in [kwargs_base, kwargs_alternativos]:
                        try:
                            resultado = obj(**kwargs)
                            return resultado
                        except Exception as e:
                            ultimo_error = e
    except Exception as e:
        ultimo_error = e

    raise RuntimeError(f"No se pudo ejecutar pyprevent con la API disponible. Detalle: {repr(ultimo_error)}")


def calcular_prevent_seguro(datos):
    """
    Cálculo seguro de PREVENT.
    Si pyprevent falta o falla, la app continúa funcionando.
    """

    if not PYPREVENT_DISPONIBLE:
        return {
            "ok": False,
            "riesgo_ascvd_10": None,
            "riesgo_cvd_10": None,
            "riesgo_hf_10": None,
            "riesgo_ascvd_30": None,
            "riesgo_cvd_30": None,
            "riesgo_hf_30": None,
            "fuente": "No disponible",
            "mensaje": (
                "Para calcular PREVENT automáticamente instale "
                "pyprevent==0.1.5 y agréguelo a requirements.txt. "
                "Detalle: ModuleNotFoundError(\"No module named 'pyprevent'\")"
            ),
        }

    try:
        resultado = intentar_calculo_prevent_con_pyprevent(datos)

        riesgo_ascvd_10 = extraer_numero_de_resultado(
            resultado,
            [
                "ascvd_10_year",
                "ascvd_10",
                "risk_ascvd_10",
                "ASCVD_10",
                "ten_year_ascvd",
                "ascvd10",
            ]
        )

        riesgo_cvd_10 = extraer_numero_de_resultado(
            resultado,
            [
                "cvd_10_year",
                "cvd_10",
                "risk_cvd_10",
                "CVD_10",
                "ten_year_cvd",
                "cvd10",
            ]
        )

        riesgo_hf_10 = extraer_numero_de_resultado(
            resultado,
            [
                "hf_10_year",
                "hf_10",
                "risk_hf_10",
                "HF_10",
                "ten_year_hf",
                "hf10",
            ]
        )

        riesgo_ascvd_30 = extraer_numero_de_resultado(
            resultado,
            [
                "ascvd_30_year",
                "ascvd_30",
                "risk_ascvd_30",
                "ASCVD_30",
                "thirty_year_ascvd",
                "ascvd30",
            ]
        )

        riesgo_cvd_30 = extraer_numero_de_resultado(
            resultado,
            [
                "cvd_30_year",
                "cvd_30",
                "risk_cvd_30",
                "CVD_30",
                "thirty_year_cvd",
                "cvd30",
            ]
        )

        riesgo_hf_30 = extraer_numero_de_resultado(
            resultado,
            [
                "hf_30_year",
                "hf_30",
                "risk_hf_30",
                "HF_30",
                "thirty_year_hf",
                "hf30",
            ]
        )

        return {
            "ok": True,
            "riesgo_ascvd_10": riesgo_ascvd_10,
            "riesgo_cvd_10": riesgo_cvd_10,
            "riesgo_hf_10": riesgo_hf_10,
            "riesgo_ascvd_30": riesgo_ascvd_30,
            "riesgo_cvd_30": riesgo_cvd_30,
            "riesgo_hf_30": riesgo_hf_30,
            "fuente": "pyprevent",
            "mensaje": "PREVENT calculado correctamente.",
        }

    except Exception as e:
        return {
            "ok": False,
            "riesgo_ascvd_10": None,
            "riesgo_cvd_10": None,
            "riesgo_hf_10": None,
            "riesgo_ascvd_30": None,
            "riesgo_cvd_30": None,
            "riesgo_hf_30": None,
            "fuente": "pyprevent con error",
            "mensaje": f"No se pudo calcular PREVENT. Detalle: {repr(e)}",
        }


# ============================================================
# SCORE OPS / OMS SIMPLIFICADO
# ============================================================

def calcular_ops_oms_simplificado(datos):
    """
    Estimación orientativa tipo OPS/OMS.
    No reemplaza la tabla oficial por país/región, pero permite integrar
    semaforización clínica si no se implementó aún la tabla exacta.
    """

    edad = safe_float(datos.get("edad"))
    sexo = datos.get("sexo")
    pas = safe_float(datos.get("presion_sistolica"))
    colesterol_total = safe_float(datos.get("colesterol_total"))
    diabetes = normalizar_si_no(datos.get("diabetes"))
    tabaquismo = normalizar_si_no(datos.get("tabaquismo"))

    puntos = 0

    if edad is not None:
        if edad >= 70:
            puntos += 5
        elif edad >= 60:
            puntos += 4
        elif edad >= 50:
            puntos += 3
        elif edad >= 40:
            puntos += 2
        elif edad >= 30:
            puntos += 1

    if sexo == "Masculino":
        puntos += 1

    if pas is not None:
        if pas >= 180:
            puntos += 5
        elif pas >= 160:
            puntos += 4
        elif pas >= 140:
            puntos += 3
        elif pas >= 130:
            puntos += 2
        elif pas >= 120:
            puntos += 1

    if colesterol_total is not None:
        if colesterol_total >= 300:
            puntos += 4
        elif colesterol_total >= 240:
            puntos += 3
        elif colesterol_total >= 200:
            puntos += 2
        elif colesterol_total >= 180:
            puntos += 1

    if diabetes:
        puntos += 3

    if tabaquismo:
        puntos += 2

    if puntos <= 4:
        categoria = "Riesgo bajo"
        estimacion = "< 5 %"
    elif puntos <= 7:
        categoria = "Riesgo moderado"
        estimacion = "5 a < 10 %"
    elif puntos <= 10:
        categoria = "Riesgo alto"
        estimacion = "10 a < 20 %"
    else:
        categoria = "Riesgo muy alto"
        estimacion = "≥ 20 %"

    return {
        "puntos": puntos,
        "categoria": categoria,
        "estimacion": estimacion,
        "mensaje": (
            "Estimación OPS/OMS simplificada. Para uso clínico definitivo "
            "se recomienda integrar la tabla oficial por región OPS/OMS."
        ),
    }


# ============================================================
# INFORME MÉDICO INTEGRADO
# ============================================================

def generar_informe_medico_integrado(datos, prevent, ops):
    nombre = datos.get("paciente", "Paciente no identificado")
    documento = datos.get("documento", "No consignado")
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")

    edad = datos.get("edad", "No consignada")
    sexo = datos.get("sexo", "No consignado")
    pas = datos.get("presion_sistolica", "No consignada")
    pad = datos.get("presion_diastolica", "No consignada")
    colesterol_total = datos.get("colesterol_total", "No consignado")
    hdl = datos.get("hdl", "No consignado")
    ldl = datos.get("ldl", "No consignado")
    trigliceridos = datos.get("trigliceridos", "No consignado")
    imc = datos.get("imc", "No consignado")
    egfr = datos.get("egfr", "No consignado")
    diabetes = datos.get("diabetes", "No consignado")
    tabaquismo = datos.get("tabaquismo", "No consignado")
    tratamiento_hta = datos.get("tratamiento_hta", "No consignado")
    antecedente_cvd = datos.get("antecedente_cvd", "No consignado")
    enfermedad_renal = datos.get("enfermedad_renal", "No consignado")
    organo_blanco = datos.get("organo_blanco", "No consignado")

    interpretacion_pa = interpretar_presion_arterial(pas, pad)
    interpretacion_ldl = interpretar_ldl(ldl)

    riesgo_ascvd_10 = prevent.get("riesgo_ascvd_10")
    riesgo_cvd_10 = prevent.get("riesgo_cvd_10")
    riesgo_hf_10 = prevent.get("riesgo_hf_10")
    riesgo_ascvd_30 = prevent.get("riesgo_ascvd_30")
    riesgo_cvd_30 = prevent.get("riesgo_cvd_30")
    riesgo_hf_30 = prevent.get("riesgo_hf_30")

    clasificacion_prevent = clasificar_riesgo_prevent(riesgo_ascvd_10)

    recomendaciones = []

    if clasificacion_prevent == "Riesgo bajo":
        recomendaciones.append(
            "Riesgo global bajo: reforzar cambios de estilo de vida, control periódico y mantenimiento de objetivos preventivos."
        )
    elif clasificacion_prevent == "Riesgo limítrofe":
        recomendaciones.append(
            "Riesgo limítrofe: considerar factores modificadores de riesgo, antecedentes familiares, daño de órgano blanco y ateromatosis subclínica."
        )
    elif clasificacion_prevent == "Riesgo intermedio":
        recomendaciones.append(
            "Riesgo intermedio: valorar intensificación preventiva, optimización de presión arterial y tratamiento lipídico según LDL y modificadores de riesgo."
        )
    elif clasificacion_prevent == "Riesgo alto":
        recomendaciones.append(
            "Riesgo alto: se sugiere intervención preventiva intensiva, control estricto de factores de riesgo y evaluación integral de daño de órgano blanco."
        )
    else:
        recomendaciones.append(
            "PREVENT no calculable: completar variables requeridas o verificar instalación de pyprevent."
        )

    pas_float = safe_float(pas)
    pad_float = safe_float(pad)
    ldl_float = safe_float(ldl)

    if pas_float is not None and pas_float >= 140:
        recomendaciones.append(
            "Presión sistólica elevada: optimizar diagnóstico y tratamiento antihipertensivo según contexto clínico y mediciones fuera del consultorio."
        )

    if pad_float is not None and pad_float >= 90:
        recomendaciones.append(
            "Presión diastólica elevada: confirmar control tensional y evaluar adherencia, técnica de medición y eventual MAPA/MDPA."
        )

    if ldl_float is not None and ldl_float >= 160:
        recomendaciones.append(
            "LDL elevado o muy elevado: considerar intensificación del tratamiento hipolipemiante según riesgo global y guías vigentes."
        )

    if normalizar_si_no(diabetes):
        recomendaciones.append(
            "Diabetes presente: requiere abordaje cardiometabólico integral y metas preventivas más estrictas."
        )

    if normalizar_si_no(tabaquismo):
        recomendaciones.append(
            "Tabaquismo activo: indicar intervención intensiva para cesación tabáquica."
        )

    if normalizar_si_no(enfermedad_renal):
        recomendaciones.append(
            "Enfermedad renal crónica referida: integrar eGFR, albuminuria y riesgo cardiovascular aumentado."
        )

    if normalizar_si_no(organo_blanco):
        recomendaciones.append(
            "Daño de órgano blanco referido: reclasifica el riesgo y justifica estrategia preventiva de mayor intensidad."
        )

    recomendaciones_txt = "\n".join([f"- {r}" for r in recomendaciones])

    informe = f"""
INFORME MÉDICO INTEGRADO DE RIESGO CARDIOVASCULAR

Fecha de generación: {fecha}

1. DATOS DEL PACIENTE

Paciente: {nombre}
Documento: {documento}
Edad: {edad}
Sexo: {sexo}

2. VARIABLES CLÍNICAS INGRESADAS

Presión arterial sistólica: {pas} mmHg
Presión arterial diastólica: {pad} mmHg
Interpretación de presión arterial: {interpretacion_pa}

Colesterol total: {colesterol_total} mg/dL
HDL colesterol: {hdl} mg/dL
LDL colesterol: {ldl} mg/dL
Triglicéridos: {trigliceridos} mg/dL
Interpretación de LDL: {interpretacion_ldl}

Índice de masa corporal: {imc} kg/m²
Filtrado glomerular estimado/eGFR: {egfr} ml/min/1.73 m²

Diabetes: {diabetes}
Tabaquismo: {tabaquismo}
Tratamiento antihipertensivo: {tratamiento_hta}
Antecedente cardiovascular establecido: {antecedente_cvd}
Enfermedad renal crónica: {enfermedad_renal}
Daño de órgano blanco: {organo_blanco}

3. RESULTADO PREVENT

Estado del cálculo PREVENT: {prevent.get("mensaje", "No informado")}
Fuente: {prevent.get("fuente", "No informada")}

ASCVD a 10 años: {formatear_porcentaje(riesgo_ascvd_10)}
CVD total a 10 años: {formatear_porcentaje(riesgo_cvd_10)}
Insuficiencia cardíaca a 10 años: {formatear_porcentaje(riesgo_hf_10)}

ASCVD a 30 años: {formatear_porcentaje(riesgo_ascvd_30)}
CVD total a 30 años: {formatear_porcentaje(riesgo_cvd_30)}
Insuficiencia cardíaca a 30 años: {formatear_porcentaje(riesgo_hf_30)}

Clasificación clínica PREVENT: {clasificacion_prevent}

4. RESULTADO OPS/OMS

Categoría OPS/OMS: {ops.get("categoria", "No calculable")}
Estimación OPS/OMS: {ops.get("estimacion", "No calculable")}
Puntaje interno orientativo: {ops.get("puntos", "No calculable")}
Observación: {ops.get("mensaje", "")}

5. INTERPRETACIÓN MÉDICA INTEGRADA

El paciente presenta un perfil de riesgo cardiovascular que debe interpretarse integrando edad, sexo,
presión arterial, perfil lipídico, diabetes, tabaquismo, función renal, tratamiento antihipertensivo,
antecedentes cardiovasculares y presencia de daño de órgano blanco.

La estimación PREVENT permite valorar riesgo aterosclerótico, cardiovascular total e insuficiencia
cardíaca a 10 y 30 años. Su principal utilidad clínica es orientar la intensidad de las estrategias
preventivas, incluyendo control de presión arterial, intervención sobre lípidos, tratamiento de diabetes,
cesación tabáquica y eventual búsqueda de ateromatosis subclínica o daño de órgano blanco cuando
exista discordancia entre riesgo calculado y juicio clínico.

El score OPS/OMS se informa como complemento de estratificación poblacional y debe integrarse con
el juicio clínico individual.

6. CONDUCTA Y RECOMENDACIONES SUGERIDAS

{recomendaciones_txt}

7. CONCLUSIÓN

Clasificación final integrada: {clasificacion_prevent} por PREVENT, con categoría OPS/OMS {ops.get("categoria", "No calculable")}.

La conducta final debe individualizarse según contexto clínico, preferencias del paciente, presencia de
comorbilidades, daño de órgano blanco, historia familiar, ateromatosis subclínica, adherencia terapéutica
y objetivos preventivos definidos por el profesional tratante.

8. NOTA MÉDICA

Este informe es una herramienta de apoyo a la decisión clínica. No reemplaza la evaluación médica
integral ni el criterio del profesional tratante.
"""

    return informe.strip()


# ============================================================
# EXPORTACIÓN EXCEL
# ============================================================

def dataframe_a_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Riesgo_CV")
    output.seek(0)
    return output.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

st.markdown(
    """
    <div class="titulo-principal">
        <h1>Riesgo Cardiovascular Integrado</h1>
        <h3>PREVENT + OPS/OMS + Informe Médico Integrado</h3>
    </div>
    """,
    unsafe_allow_html=True
)

if not PYPREVENT_DISPONIBLE:
    st.warning(
        'PREVENT no está disponible porque falta instalar pyprevent==0.1.5. '
        'Agregue pyprevent==0.1.5 a requirements.txt y reinicie la app. '
        'La app continuará funcionando con informe integrado y OPS/OMS.'
    )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("Carga única de datos")

st.sidebar.markdown("### Datos del paciente")

paciente = st.sidebar.text_input("Paciente", value="")
documento = st.sidebar.text_input("Documento", value="")

edad = st.sidebar.number_input(
    "Edad",
    min_value=18,
    max_value=100,
    value=55,
    step=1
)

sexo = st.sidebar.selectbox(
    "Sexo",
    ["Masculino", "Femenino"]
)

st.sidebar.markdown("### Presión arterial")

presion_sistolica = st.sidebar.number_input(
    "Presión sistólica, PAS (mmHg)",
    min_value=70,
    max_value=260,
    value=130,
    step=1
)

presion_diastolica = st.sidebar.number_input(
    "Presión diastólica, PAD (mmHg)",
    min_value=40,
    max_value=160,
    value=80,
    step=1
)

tratamiento_hta = st.sidebar.selectbox(
    "Tratamiento antihipertensivo",
    ["No", "Sí"]
)

st.sidebar.markdown("### Perfil lipídico")

colesterol_total = st.sidebar.number_input(
    "Colesterol total (mg/dL)",
    min_value=80,
    max_value=500,
    value=200,
    step=1
)

hdl = st.sidebar.number_input(
    "HDL colesterol (mg/dL)",
    min_value=15,
    max_value=120,
    value=50,
    step=1
)

ldl = st.sidebar.number_input(
    "LDL colesterol (mg/dL)",
    min_value=20,
    max_value=400,
    value=120,
    step=1
)

trigliceridos = st.sidebar.number_input(
    "Triglicéridos (mg/dL)",
    min_value=30,
    max_value=1000,
    value=150,
    step=1
)

st.sidebar.markdown("### Metabolismo y riñón")

imc = st.sidebar.number_input(
    "IMC (kg/m²)",
    min_value=15.0,
    max_value=60.0,
    value=27.0,
    step=0.1
)

egfr = st.sidebar.number_input(
    "eGFR (ml/min/1.73 m²)",
    min_value=5,
    max_value=150,
    value=90,
    step=1
)

diabetes = st.sidebar.selectbox(
    "Diabetes",
    ["No", "Sí"]
)

st.sidebar.markdown("### Factores de riesgo")

tabaquismo = st.sidebar.selectbox(
    "Tabaquismo actual",
    ["No", "Sí"]
)

antecedente_cvd = st.sidebar.selectbox(
    "Antecedente cardiovascular establecido",
    ["No", "Sí"]
)

enfermedad_renal = st.sidebar.selectbox(
    "Enfermedad renal crónica",
    ["No", "Sí"]
)

organo_blanco = st.sidebar.selectbox(
    "Daño de órgano blanco",
    ["No", "Sí"]
)


# ============================================================
# DATOS ÚNICOS SINCRONIZADOS
# ============================================================

datos_paciente = {
    "paciente": paciente if paciente else "Paciente no identificado",
    "documento": documento if documento else "No consignado",
    "edad": edad,
    "sexo": sexo,
    "presion_sistolica": presion_sistolica,
    "presion_diastolica": presion_diastolica,
    "tratamiento_hta": tratamiento_hta,
    "colesterol_total": colesterol_total,
    "hdl": hdl,
    "ldl": ldl,
    "trigliceridos": trigliceridos,
    "imc": imc,
    "egfr": egfr,
    "diabetes": diabetes,
    "tabaquismo": tabaquismo,
    "antecedente_cvd": antecedente_cvd,
    "enfermedad_renal": enfermedad_renal,
    "organo_blanco": organo_blanco,
}


# ============================================================
# CÁLCULOS
# ============================================================

resultado_prevent = calcular_prevent_seguro(datos_paciente)
resultado_ops = calcular_ops_oms_simplificado(datos_paciente)
informe_integrado = generar_informe_medico_integrado(
    datos=datos_paciente,
    prevent=resultado_prevent,
    ops=resultado_ops
)

clasificacion_prevent = clasificar_riesgo_prevent(
    resultado_prevent.get("riesgo_ascvd_10")
)

css_prevent = clase_css_riesgo(clasificacion_prevent)


# ============================================================
# PANEL PRINCIPAL
# ============================================================

col_a, col_b, col_c = st.columns(3)

with col_a:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.metric("PAS", f"{presion_sistolica} mmHg")
    st.write(interpretar_presion_arterial(presion_sistolica, presion_diastolica))
    st.markdown("</div>", unsafe_allow_html=True)

with col_b:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.metric("LDL colesterol", f"{ldl} mg/dL")
    st.write(interpretar_ldl(ldl))
    st.markdown("</div>", unsafe_allow_html=True)

with col_c:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.metric("OPS/OMS", resultado_ops["categoria"])
    st.write(f"Estimación: {resultado_ops['estimacion']}")
    st.markdown("</div>", unsafe_allow_html=True)


st.markdown('<div class="subtitulo">Resultado PREVENT</div>', unsafe_allow_html=True)

if resultado_prevent["ok"]:
    st.success(resultado_prevent["mensaje"])
else:
    st.warning(resultado_prevent["mensaje"])

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "ASCVD 10 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_ascvd_10"))
    )

with col2:
    st.metric(
        "CVD total 10 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_cvd_10"))
    )

with col3:
    st.metric(
        "Insuficiencia cardíaca 10 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_hf_10"))
    )

col4, col5, col6 = st.columns(3)

with col4:
    st.metric(
        "ASCVD 30 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_ascvd_30"))
    )

with col5:
    st.metric(
        "CVD total 30 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_cvd_30"))
    )

with col6:
    st.metric(
        "Insuficiencia cardíaca 30 años",
        formatear_porcentaje(resultado_prevent.get("riesgo_hf_30"))
    )

st.markdown(
    f"""
    <div class="{css_prevent}">
        Clasificación PREVENT integrada: {clasificacion_prevent}
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# OPS/OMS
# ============================================================

st.markdown('<div class="subtitulo">Score OPS/OMS</div>', unsafe_allow_html=True)

col_ops1, col_ops2, col_ops3 = st.columns(3)

with col_ops1:
    st.metric("Categoría OPS/OMS", resultado_ops["categoria"])

with col_ops2:
    st.metric("Estimación", resultado_ops["estimacion"])

with col_ops3:
    st.metric("Puntaje orientativo", resultado_ops["puntos"])

st.info(resultado_ops["mensaje"])


# ============================================================
# TABLA DE DATOS SINCRONIZADOS
# ============================================================

st.markdown('<div class="subtitulo">Datos utilizados para cálculo simultáneo</div>', unsafe_allow_html=True)

df_datos = pd.DataFrame([{
    "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "Paciente": datos_paciente["paciente"],
    "Documento": datos_paciente["documento"],
    "Edad": edad,
    "Sexo": sexo,
    "PAS": presion_sistolica,
    "PAD": presion_diastolica,
    "Tratamiento_HTA": tratamiento_hta,
    "Colesterol_Total": colesterol_total,
    "HDL": hdl,
    "LDL": ldl,
    "Trigliceridos": trigliceridos,
    "IMC": imc,
    "eGFR": egfr,
    "Diabetes": diabetes,
    "Tabaquismo": tabaquismo,
    "Antecedente_CVD": antecedente_cvd,
    "ERC": enfermedad_renal,
    "Daño_Organo_Blanco": organo_blanco,
    "PREVENT_ASCVD_10": resultado_prevent.get("riesgo_ascvd_10"),
    "PREVENT_CVD_10": resultado_prevent.get("riesgo_cvd_10"),
    "PREVENT_HF_10": resultado_prevent.get("riesgo_hf_10"),
    "PREVENT_ASCVD_30": resultado_prevent.get("riesgo_ascvd_30"),
    "PREVENT_CVD_30": resultado_prevent.get("riesgo_cvd_30"),
    "PREVENT_HF_30": resultado_prevent.get("riesgo_hf_30"),
    "Clasificacion_PREVENT": clasificacion_prevent,
    "OPS_OMS_Categoria": resultado_ops.get("categoria"),
    "OPS_OMS_Estimacion": resultado_ops.get("estimacion"),
    "OPS_OMS_Puntos": resultado_ops.get("puntos"),
}])

st.dataframe(df_datos, use_container_width=True)


# ============================================================
# INFORME MÉDICO INTEGRADO
# ============================================================

st.markdown('<div class="subtitulo">Informe médico integrado</div>', unsafe_allow_html=True)

st.text_area(
    "Informe médico",
    value=informe_integrado,
    height=620
)

col_desc1, col_desc2 = st.columns(2)

with col_desc1:
    st.download_button(
        label="Descargar informe médico integrado TXT",
        data=informe_integrado.encode("utf-8"),
        file_name="informe_medico_integrado_riesgo_cardiovascular.txt",
        mime="text/plain"
    )

with col_desc2:
    excel_bytes = dataframe_a_excel_bytes(df_datos)
    st.download_button(
        label="Descargar base Excel",
        data=excel_bytes,
        file_name="riesgo_cardiovascular_integrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ============================================================
# PIE
# ============================================================

st.markdown("---")
st.caption(
    "App de apoyo a la decisión clínica. PREVENT depende de pyprevent==0.1.5. "
    "OPS/OMS se informa como estimación simplificada hasta integrar tablas oficiales regionales."
)
