"""Comparación complementaria de regresores con los datos separados en el notebook."""
import argparse
import json
import platform
from pathlib import Path

import numpy as np
import pandas as pd
import pycaret
import sklearn
from pycaret.regression import RegressionExperiment
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import KFold
from sklearn.linear_model import LinearRegression
from scipy.sparse import csr_matrix


class RegresionLinealDispersa(LinearRegression):
    """Usa LSQR, como la regresión principal, ante columnas codificadas redundantes."""

    def fit(self, X, y, sample_weight=None):
        return super().fit(csr_matrix(X), y, sample_weight=sample_weight)

    def predict(self, X):
        return super().predict(csr_matrix(X))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--entrada', type=Path, required=True)
    parser.add_argument('--salida', type=Path, required=True)
    args = parser.parse_args()
    config = json.loads((args.entrada / 'config.json').read_text(encoding='utf-8'))
    train = pd.read_json(args.entrada / 'entrenamiento.json', orient='split')
    test = pd.read_json(args.entrada / 'prueba.json', orient='split')
    assert train.index.is_unique and test.index.is_unique
    assert set(train.index).isdisjoint(test.index)
    assert train.columns.tolist() == test.columns.tolist()
    assert not {'precio_usd_m2', 'nivel_valorizacion_barrio'} & set(train.columns)
    for df in (train, test):
        df[config['numericas']] = df[config['numericas']].astype(float)
        for column in config['categoricas']:
            df[column] = df[column].astype(object).where(df[column].notna(), np.nan)

    folds = KFold(n_splits=5, shuffle=True, random_state=config['semilla'])
    exp = RegressionExperiment()
    exp.setup(data=train, test_data=test, target='precio', index=True,
        numeric_features=config['numericas'], categorical_features=config['categoricas'],
        numeric_imputation='median', categorical_imputation='constant',
        max_encoding_ohe=1000, normalize=True, normalize_method='zscore',
        fold_strategy=folds, session_id=config['semilla'], n_jobs=2,
        html=False, verbose=False, memory=False, system_log=False)
    assert exp.get_config('X_train').index.tolist() == train.index.tolist()
    assert exp.get_config('X_test').index.tolist() == test.index.tolist()
    assert len(exp.get_config('X_train')) == len(train)

    # Quince algoritmos: se comparan sus configuraciones predeterminadas.
    candidates = [RegresionLinealDispersa(), 'ridge', 'lasso', 'en', 'br', 'huber', 'knn', 'dt',
                  'rf', 'et', 'ada', 'gbr', 'omp', 'xgboost', 'dummy']
    best = exp.compare_models(include=candidates, sort='MAE', n_select=1,
        fold=folds, errors='raise', verbose=False)
    comparison = exp.pull().copy()
    comparison['Model'] = comparison['Model'].replace({
        'RegresionLinealDispersa': 'Regresión lineal (solver disperso)'})
    assert len(comparison) == 15
    predictions = exp.predict_model(best, data=test, verbose=False)
    assert predictions.index.tolist() == test.index.tolist()
    predicted = predictions['prediction_label'].to_numpy()
    assert np.isfinite(predicted).all()
    metrics = pd.DataFrame([{
        'modelo': comparison.iloc[0]['Model'],
        'MAE_test_USD': mean_absolute_error(test['precio'], predicted),
        'RMSE_test_USD': np.sqrt(mean_squared_error(test['precio'], predicted)),
        'R2_test': r2_score(test['precio'], predicted)
    }])
    metadata = {'modelo_seleccionado': comparison.iloc[0]['Model'],
        'modelo_id': str(comparison.index[0]), 'modelos_comparados': len(comparison),
        'entrenamiento': len(train), 'prueba': len(test), 'semilla': config['semilla'],
        'pliegues': 5, 'python': platform.python_version(), 'pycaret': pycaret.__version__,
        'sklearn': sklearn.__version__, 'pandas': pd.__version__, 'numpy': np.__version__}
    args.salida.mkdir(parents=True, exist_ok=True)
    comparison.to_csv(args.salida / 'comparacion_cv.csv', index_label='modelo_id')
    metrics.to_csv(args.salida / 'evaluacion_prueba.csv', index=False)
    pd.DataFrame({'precio_observado': test['precio'], 'precio_predicho': predicted}).to_csv(
        args.salida / 'predicciones_prueba.csv', index_label='fila_original')
    (args.salida / 'ejecucion.json').write_text(json.dumps(metadata, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps(metadata, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    main()
