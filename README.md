# Procesador Exp Contable - SIAF

Esta app en **Streamlit** permite procesar un archivo Excel del SIAF y aplicar las siguientes reglas:

1. Se crea una columna `exp_contable` con la concatenación de:
   - `ano_eje`
   - `nro_not_exp`
   - `ciclo`
   - `fase`

2. Para cada `exp_contable`:
   - Si existe al menos una fila con `mayor = 1101`, se invierten los valores de `debe` y `haber` en nuevas columnas (`debe_adj`, `haber_adj`).
   - Caso contrario, se mantienen igual.

3. Se genera un archivo Excel con los resultados.

## 🚀 Cómo usar
```bash
pip install -r requirements.txt
streamlit run app.py
```

Sube tu archivo Excel y descarga el resultado procesado.
