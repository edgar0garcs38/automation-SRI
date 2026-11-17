# Auditor SRI - Conciliacion 103/104 vs Libro Mayor

Aplicacion de escritorio en Python para auditores tributarios en Ecuador. Incluye tres flujos:

- Formularios F103/F104 contra Libro Mayor.
- Formularios F103/F104 por sí solos, con extracción de todos los casilleros y validaciones internas.
- Formulario F101 contra la hoja `MAPEO DE LA DECLARACIÓN A1` del Anexo ICT.

## Instalacion y uso en 5 pasos

1. Crea y activa un entorno virtual con Python 3.10 o superior.

   ```bash
   cd auditorsri
   python3 -m venv .venv
   source .venv/bin/activate
   ```

2. Instala las dependencias.

   ```bash
   pip install -r requirements.txt
   ```

3. Ejecuta la aplicacion de escritorio.

   ```bash
   python desktop_app.py
   ```

   La version web anterior sigue disponible con `streamlit run app.py`.

4. En el menú inicial selecciona el análisis:

   - Para F103/F104, carga los PDFs y el Libro Mayor con columnas `CTA. CONTABLE`, `FECHA`, `CONCEPTO`, `DEBE`, `HABER`, `SALDO`.
   - Para revisar únicamente formularios F103/F104, carga uno o varios PDFs; no se requiere Libro Mayor.
   - Para F101, carga el PDF anual y el Anexo ICT que contenga la hoja `MAPEO DE LA DECLARACIÓN A1`.

5. Procesa los archivos, revisa las diferencias y exporta el reporte Excel. El detalle incluye número serial y fecha de recaudación.

## Regla de pago del Formulario 103

El total pagado del casillero `999` se compara contra la cuenta configurada del
Libro Mayor en el mes siguiente al periodo declarado. Por ejemplo, el F103 de
enero se cruza contra el movimiento contable registrado en febrero.

El reporte muestra por separado `PERIODO`, `PERIODO_PAGO`, `BASE IMPONIBLE`
(casillero 349 del F103) y `VALOR RETENIDO` (casillero 399). La hoja
`Casilleros Extraidos` conserva todos los códigos y valores detectados en los
PDFs, incluyendo número de página y texto de origen para trazabilidad.

## Archivos principales

- `app.py`: interfaz Streamlit.
- `desktop_app.py`: interfaz de escritorio con Tkinter.
- `parser_pdf.py`: extraccion de casilleros desde PDFs del SRI.
- `parser_ict.py`: lectura de la hoja A1 del Anexo ICT.
- `parser_libro_mayor.py`: limpieza y agrupacion del Libro Mayor.
- `comparador.py`: calculo de diferencias y estados.
- `comparador_101.py`: comparación F101 contra valores declarados del Anexo ICT.
- `mapeo_config.py`: lectura y persistencia de mapeos.
- `exportar.py`: generacion del reporte Excel.
- `mapeo_default.json`: mapeos iniciales editables.

## Notas

La aplicacion funciona offline. No llama APIs externas. Los PDFs deben tener texto seleccionable; si son escaneados, primero deben pasar por OCR.
