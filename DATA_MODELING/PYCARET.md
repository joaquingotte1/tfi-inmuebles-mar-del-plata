# Comparación complementaria con PyCaret

El apartado 3.9 de `01_EDA.ipynb` ejecuta PyCaret en un entorno separado. PyCaret 3.3.2 usa versiones de pandas y scikit-learn diferentes de las del análisis principal.

Desde la raíz del proyecto, en Windows:

```powershell
py -3.11 -m venv DATA_MODELING/.venv_pycaret
DATA_MODELING/.venv_pycaret/Scripts/python.exe -m pip install -r DATA_MODELING/requirements-pycaret.txt
```

En Linux o macOS, crear el entorno con Python 3.11 y usar `DATA_MODELING/.venv_pycaret/bin/python` para instalar las dependencias.

Ejecutar el notebook en orden. El apartado 3.9 exporta `X_train`, `X_test`, `y_train` y `y_test` sin imputar a `.artifact_runtime/pycaret/` y llama a `pycaret_complementario.py`. El script conserva los índices y comprueba que la separación no cambie.

PyCaret compara 15 algoritmos con configuraciones predeterminadas, cinco pliegues aleatorios con semilla 42 y selección por MAE. Aplica imputación, codificación y estandarización en su propio pipeline; se fuerza One Hot Encoding también para barrio. No se transforma el precio ni se emplean precio por m² o valorización como predictores. Esta comparación complementa el ajuste de hiperparámetros del apartado 3.5.

Las tablas, predicciones y versiones se guardan en `DATA_MODELING/reportes/pycaret/`. El entorno y los archivos temporales están excluidos de Git.

LightGBM se excluyó por un error nativo de ejecución en este entorno Windows. En su lugar se incluyó Orthogonal Matching Pursuit; la comparación conserva 15 modelos.

La regresión lineal usa el solver LSQR con entrada dispersa, como el modelo principal. El solver denso de PyCaret produjo coeficientes inestables con las columnas redundantes de One Hot Encoding. Se mantiene la regresión sin regularización, los mismos predictores y los mismos pliegues; solo cambia la resolución numérica.

Referencia: [documentación de regresión de PyCaret](https://pycaret.readthedocs.io/en/stable/api/regression.html).
