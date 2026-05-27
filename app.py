def procesar_evaluacion_completa(p: Patient) -> Patient:
    # 1. Ejecución robusta y forzada de pyprevent (Alineado con ACC/AHA 2026)
    if PREVENT_AVAILABLE:
        try:
            # Forzamos la conversión de strings de género según la API de la AHA
            genero_py = "female" if str(p.sexo).strip().lower() in ["femenino", "female"] else "male"
            
            # CRUCIAL: Forzar enteros (int) y flotantes (float) estrictos para evitar excepciones internas
            res = pyprevent.calculate_risk(
                age=int(p.edad),
                sex=genero_py,
                sbp=int(round(p.presion_sistolica)),
                bp_med=1 if p.tratamiento_hta else 0,
                tot_chol=int(round(p.colesterol_total)),
                hdl_chol=int(round(p.hdl)),
                ldl_chol=int(round(p.ldl_actual)),
                diabetes=1 if p.diabetes else 0,
                smoker=1 if p.tabaquismo else 0,
                egfr=float(p.egfr) if p.egfr else 75.0
            )
            
            # Asignación segura de resultados extraídos del diccionario de pyprevent
            p.prevent_10 = round(res.get("10_yr_ascvd_risk", 0.0), 2) if res.get("10_yr_ascvd_risk") is not None else None
            p.prevent_30 = round(res.get("30_yr_ascvd_risk", 0.0), 2) if res.get("30_yr_ascvd_risk") is not None else None
            
        except Exception as e:
            # Si hay un error, lo imprimimos en la consola del servidor de Streamlit para diagnóstico rápido
            print(f" ERROR CRÍTICO EN PYPREVENT API: {repr(e)}")
            st.sidebar.error(f"Error en pyprevent: {e}")
            p.prevent_10, p.prevent_30 = None, None
    else:
        p.prevent_10, p.prevent_30 = None, None

    # 2. Estratificación OPS Hearts (Independiente del score PREVENT)
    p.ops_hearts_riesgo = calcular_ops_hearts(p)
    
    # 3. Correlación de Guías Terapéuticas e Indicaciones Farmacológicas
    if p.antecedente_infarto:
        p.categoria_riesgo_final = "Prevención Secundaria (Extremo / Muy Alto)"
        p.meta_ldl = "< 55 mg/dL"
        p.indicacion_tratamiento = "Evidencia Clase I. Estatinas de alta intensidad (Atorvastatina 40-80mg o Rosuvastatina 20-40mg) asociadas de ser necesario a Ezetimibe 10mg por máximo beneficio clínico."
    else:
        score_riesgo = p.prevent_10 if p.prevent_10 is not None else 0.0
        if score_riesgo >= 10.0 or p.ops_hearts_riesgo == "Alto / Muy Alto":
            p.categoria_riesgo_final = "Riesgo Alto"
            p.meta_ldl = "< 70 mg/dL"
            p.indicacion_tratamiento = "Indicación categórica de Estatinas de alta intensidad. Control lipídico institucional estricto a las 4-6 semanas."
        elif 5.0 <= score_riesgo < 10.0 or p.ops_hearts_riesgo == "Moderado":
            p.categoria_riesgo_final = "Riesgo Intermedio"
            p.meta_ldl = "< 100 mg/dL"
            p.indicacion_tratamiento = "Iniciar estatinas de moderada intensidad. Discutir la presencia de factores potenciadores de riesgo cardiovascular con el paciente."
        else:
            p.categoria_riesgo_final = "Riesgo Bajo"
            p.meta_ldl = "< 116 mg/dL"
            p.indicacion_tratamiento = "Modificación del estilo de vida, dieta cardioprotectora (mediterránea) y ejercicio regular. Reevaluación anual."
            
    return p
