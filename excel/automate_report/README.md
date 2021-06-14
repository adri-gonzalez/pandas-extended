# AUTOMATE EXCEL REPORT

Si su flujo de trabajo es pesado en Excel, estoy seguro de que tiene libros de trabajo que se extienden en varias pestañas, muchas hojas de tablas dinámicas e incluso más fórmulas que hacen referencias cruzadas en cada hoja. Esto es genial ... si su objetivo es confundir a todos los que intentan entender sus hojas de cálculo.
Un beneficio de usar Python para resumir sus datos es la capacidad de almacenar sus tablas dinámicas y otros datos, como un DataFrame dentro de una sola variable. Hacer referencia de esta manera es más fácil de depurar que intentar rastrear números a través de múltiples pestañas en una hoja de cálculo.

## Pandas: 
este módulo es esencial para resumir con éxito sus datos.
Según el sitio de pandas:

"Pandas es una herramienta de manipulación y análisis de datos de código abierto rápida, potente, flexible y fácil de usar"
En resumen, pandas contiene funciones que harán todo el análisis de datos que normalmente haces en Excel. Aquí hay algunas de las funciones que encontrará interesantes provenientes de un fondo basado en Excel y la documentación está vinculada a cada una de ellas si desea obtener más información:

* value_counts()
* drop_duplicates()
* groupby()
* describe()
* pivot_table() (every Excel users favourite)
* plot()
* dtypes
* loc
* iloc

Estos son algunos ejemplos de la multitud de funciones de pandas disponibles para manipular o resumir rápidamente sus datos.

## Requirements

```python
import pandas as pd
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
```

### XLWINGS

xlwings es de código abierto y gratuito, viene preinstalado con Anaconda y WinPython, y funciona en Windows y macOS.

Automatice Excel a través de scripts de Python o cuadernos de Jupyter, llame a Python desde Excel a través de macros y escriba funciones definidas por el usuario (las UDF son solo para Windows).

> https://www.xlwings.org/

## REPORTE FINAL

![Alt Text](https://miro.medium.com/max/3760/1*9w0oAqIIwJCpGDWtfvPkKQ.gif)
