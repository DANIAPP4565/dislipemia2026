# LipidCare Precision 2026

Aplicación Streamlit autocontenida para lipidología de precisión, cálculo automático PREVENT y apoyo farmacológico trazable.

## Funcionalidades

- PREVENT de 10 años para adultos de 30 a 79 años sin enfermedad cardiovascular conocida.
- PREVENT de 30 años para el rango de 30 a 59 años.
- Modelos base y completos, con UACR, HbA1c y SDI opcionales e indicadores de datos faltantes.
- Resultados para CVD total, ASCVD, insuficiencia cardíaca, enfermedad coronaria y ACV.
- Autoverificación numérica contra el ejemplo publicado en la Tabla S25.
- LDL-C directo o cálculo por Sampson/NIH; comparación con Friedewald.
- No-HDL-C, colesterol remanente estimado, ApoB, Lp(a), TG, hs-CRP y CAC.
- CKD-EPI 2021 a partir de creatinina.
- Estrategia farmacológica por prevención primaria/secundaria, hipercolesterolemia severa, grupos de alto riesgo, metas, intolerancia y escalamiento.
- Simulación de reducción de LDL-C con estatinas y terapias adicionales.
- Base de conocimiento versionada con identificadores de regla y trazabilidad.
- Exportación HTML, JSON y PDF.

## Ejecución local

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Community Cloud

Subir `app.py` y `requirements.txt` a la raíz del repositorio y seleccionar `app.py` como archivo principal.

## Validación incorporada

La pestaña PREVENT muestra un control automático. El caso publicado de mujer de 50 años, colesterol total 200 mg/dL, HDL-C 45 mg/dL, PAS tratada 160 mmHg, diabetes, no fumadora y eGFR 90 debe devolver:

- CVD total a 10 años: 14,683939%
- ASCVD a 10 años: 9,195090%
- Insuficiencia cardíaca a 10 años: 8,056097%

## Alcance clínico

Herramienta profesional de apoyo a la decisión. Las recomendaciones requieren confirmación clínica, revisión de contraindicaciones e interacciones, disponibilidad local, preferencias del paciente y seguimiento de laboratorio. PREVENT deriva de cohortes de Estados Unidos y requiere interpretación prudente en población argentina.
