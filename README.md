# Scraping de Mitula — Mar del Plata

Extrae exclusivamente casas, departamentos y chalets en venta publicados en Mitula para Mar del Plata y guarda el DataFrame directamente en Excel. El análisis exploratorio y el modelado se desarrollan en DATA_MODELING/01_EDA.ipynb.

## Instalación

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

El scraper utiliza exclusivamente Microsoft Edge y conserva su sesión en `data/.mitula_edge_profile`.

## Ejecución

Primera página, con un máximo de 30 propiedades:

```powershell
.venv\Scripts\python.exe scraper_mitula.py --max-paginas 1 --max-avisos 30
```

Cinco páginas, con un máximo de 150 propiedades:

```powershell
.venv\Scripts\python.exe scraper_mitula.py --max-paginas 5 --max-avisos 150
```

Prueba de dos páginas con reintentos y archivo parcial:

```powershell
.venv\Scripts\python.exe scraper_mitula.py --max-paginas 2 --max-avisos 60 --salida data/mitula_prueba_2_paginas.xlsx
```

Para una extracción extensa se puede indicar un número alto de páginas y desactivar el límite de avisos; el scraper se detiene si encuentra una página vacía o repetida:

```powershell
.venv\Scripts\python.exe scraper_mitula.py --max-paginas 200 --max-avisos 0 --salida data/mitula_venta_mar_del_plata.xlsx
```

La reanudación está activada de forma predeterminada. El archivo parcial se crea junto a la salida con el sufijo `_parcial.xlsx`. Para comenzar desde cero sin leerlo, use `--no-reanudar`. Los parámetros `--reintentos` y `--espera-reintento-ms` controlan los reintentos ante errores temporales.

Para continuar directamente desde una página conocida sin recorrer las anteriores, use `--pagina-inicial`. Por ejemplo:

```powershell
.venv\Scripts\python.exe scraper_mitula.py --pagina-inicial 11 --max-paginas 200 --max-avisos 0 --salida data/mitula_venta_mar_del_plata.xlsx
```

Si Mitula presenta un CAPTCHA durante una ejecución visible, el scraper mantiene Edge abierto y espera hasta 15 minutos para que se complete manualmente. El tiempo puede modificarse con `--esperar-captcha-minutos`. Después de resolverlo, la extracción continúa automáticamente desde el parcial.

La salida predeterminada es `data/mitula_venta_mar_del_plata.xlsx`.

El scraper recorre tanto el listado como la ficha interna de cada aviso. El archivo conserva las variables publicadas por Mitula y agrega trazabilidad de la captura: valores originales de precio y superficie, URL y página de origen, posición en el listado, fecha UTC del scraping, estado HTTP, éxito o error de la ficha y textos crudos del listado y del detalle. La única variable completada mediante inferencia es `barrio`, acompañada por `barrio_fuente` para indicar si procede del campo de la ficha o de una mención explícita en el título o descripción. Los registros se filtran para conservar solamente casas, departamentos y chalets de Mar del Plata y se eliminan duplicados mediante el ID del aviso (o la URL cuando no existe ID). `Chalet` se conserva como categoría original para normalizarla posteriormente durante el EDA.

También intenta conservar, cuando la publicación los informa, latitud, longitud, piso, unidad, cantidad de plantas, frente y fondo del terreno, estado, disposición, orientación, aptitud profesional, baulera, cloacas, tipo de gas, portal e ID de origen. Las coordenadas se leen de los metadatos de la ficha; el resto se obtiene únicamente de menciones o campos publicados. Si Mitula no expone una variable, la celda permanece vacía.

Cada XLSX contiene las hojas `Propiedades`, `Ejecucion` y `Cobertura`. Durante el recorrido se actualiza el archivo parcial al terminar cada página. La hoja `Ejecucion` registra estado, páginas, tarjetas, duplicados, fichas exitosas y errores; `Cobertura` resume completos y faltantes por variable.

El programa usa pausas moderadas y no intenta eludir verificaciones ni controles del sitio.

Las variables de presencia usan tres estados: `Sí` cuando la publicación afirma que la característica existe, `No` cuando la niega explícitamente y celda vacía cuando no aporta información. La ausencia de una mención no se interpreta como `No`.

## Trabajo Final Integrador y anexos

El documento final en Word y PDF, junto con los enlaces al notebook, las bases y los reportes, se encuentra en el [índice de anexos](ANEXOS.md).
