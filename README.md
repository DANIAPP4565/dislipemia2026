# LipidCare Precision 2026

Aplicación Streamlit autocontenida para lipidología de precisión, importación bioquímica multiformato, cálculo automático PREVENT y apoyo farmacológico trazable.

## Funcionalidades principales

- PREVENT de 10 años para adultos de 30 a 79 años sin enfermedad cardiovascular conocida.
- PREVENT de 30 años para el rango de 30 a 59 años.
- Modelos base y completos, con UACR, HbA1c y SDI opcionales e indicadores de datos faltantes.
- Resultados para CVD total, ASCVD, insuficiencia cardíaca, enfermedad coronaria y ACV.
- Autoverificación numérica contra el ejemplo publicado en la Tabla S25.
- LDL-C informado por laboratorio o cálculo por Sampson/NIH; comparación con Friedewald.
- No-HDL-C, colesterol remanente estimado, ApoB, Lp(a), TG, hs-CRP y CAC.
- CKD-EPI 2021 a partir de creatinina.
- Estrategia farmacológica por prevención primaria/secundaria, hipercolesterolemia severa, grupos de alto riesgo, metas, intolerancia y escalamiento.
- Simulación de reducción de LDL-C con estatinas y terapias adicionales.
- Base de conocimiento versionada con identificadores de regla y trazabilidad.
- Exportación HTML, JSON y PDF.

## Importación de resultados bioquímicos

Formatos admitidos:

- PDF con texto digital.
- PDF escaneado mediante OCR.
- Imágenes PNG, JPG, JPEG, TIFF y WEBP.
- Excel XLSX, XLS y ODS.
- CSV y TSV, tanto en formato ancho como en formato `variable / resultado / unidad`.
- DOCX, JSON, XML, HL7, RTF y TXT.
- Texto pegado directamente desde portales o correos del laboratorio.

Analitos reconocidos:

- Colesterol total, HDL-C, LDL-C y triglicéridos.
- ApoB y Lp(a), conservando la unidad original de Lp(a) sin conversiones no validadas.
- Creatinina, eGFR, HbA1c, UACR y PCR ultrasensible.
- Glucosa, IMC, peso, talla y presión arterial sistólica.
- AST/GOT, ALT/GPT, CK/CPK y TSH como resultados auxiliares para revisión clínica.

El flujo es deliberadamente seguro:

1. Extrae texto digital y utiliza OCR sólo cuando es necesario.
2. Normaliza coma decimal y unidades frecuentes.
3. Convierte mmol/L a mg/dL para lípidos, µmol/L a mg/dL para creatinina, mg/mmol a mg/g para UACR y mg/dL a mg/L para PCR ultrasensible.
4. Muestra valor, unidad, fuente, línea detectada, confianza y posibles conflictos.
5. No modifica el caso hasta que el profesional confirme cada resultado.
6. Tras la confirmación, completa automáticamente los campos compatibles y recalcula PREVENT, métricas lipídicas y estrategia farmacológica.
7. Conserva trazabilidad por nombre de archivo, método de extracción y SHA-256, sin incorporar el texto clínico completo al JSON exportado.

La aplicación incluye una plantilla CSV descargable desde el importador.

## Ejecución local

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Para OCR local también debe estar instalado Tesseract con idioma español. En Ubuntu/Debian:

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-spa
```

## Despliegue en Streamlit Community Cloud

Subir a la raíz del repositorio:

- `app.py`
- `requirements.txt`
- `packages.txt`

Seleccionar `app.py` como archivo principal. `packages.txt` instala Tesseract y el idioma español para los PDF escaneados.

## Validación incorporada

La pestaña PREVENT muestra un control automático. El caso publicado de mujer de 50 años, colesterol total 200 mg/dL, HDL-C 45 mg/dL, PAS tratada 160 mmHg, diabetes, no fumadora y eGFR 90 debe devolver:

- CVD total a 10 años: 14,683939%.
- ASCVD a 10 años: 9,195090%.
- Insuficiencia cardíaca a 10 años: 8,056097%.

## Alcance clínico

Herramienta profesional de apoyo a la decisión. Los resultados importados requieren revisión y confirmación. Las recomendaciones requieren juicio clínico, verificación de contraindicaciones e interacciones, disponibilidad local, preferencias del paciente y seguimiento de laboratorio. PREVENT deriva de cohortes de Estados Unidos y requiere interpretación prudente en población argentina.
