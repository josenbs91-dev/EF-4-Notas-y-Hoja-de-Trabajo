# Procesador de Exp Contable - SIAF con Equivalencias

Esta aplicación en **Streamlit** procesa archivos de movimientos contables del SIAF y permite:

## Funcionalidades
1. Subir un archivo Excel principal con los movimientos contables.
2. Subir un archivo de **equivalencias** (con una hoja llamada `Hoja de Trabajo`) que contenga:
   - **Cuentas Contables**
   - **Rubros**
3. Crear un identificador `exp_contable` a partir de las columnas:
   - `ano_eje`, `nro_not_exp`, `ciclo`, `fase`.
4. Ajustar los montos de `debe` y `haber`:
   - Se invierten si el expediente **NO pertenece a mayor=1101**.
5. Generar equivalencias uniendo `mayor.sub_cta` con la columna `Cuentas Contables` y asignando el **Rubro** correspondiente.
6. Exportar a Excel con 3 hojas:
   - `Resultado_Completo` → todos los registros con equivalencias.
   - `TipoCTB1_1101` → solo registros con `tipo_ctb=1` y expedientes con `mayor=1101`.
   - `TipoCTB1_No1101` → solo registros con `tipo_ctb=1` y expedientes que **no** tienen `mayor=1101`.

## Instalación
1. Clonar el repositorio o copiar los archivos.
2. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
