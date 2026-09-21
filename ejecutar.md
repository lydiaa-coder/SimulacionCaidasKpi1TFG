# Cómo ejecutar la simulación

## 1. Instalar Python

Si no lo tienes ya instalado, la forma más sencilla es instalar **Anaconda**
(incluye Python + las librerías que necesitamos, sin tener que instalar nada
más a mano):

- Descarga: https://www.anaconda.com/download
- Instálalo con las opciones por defecto.

## 2. Abrir una terminal

- **Windows**: busca "Anaconda Prompt" en el menú de inicio y ábrelo.
- **Mac**: abre "Terminal" (Cmd + Espacio, escribe "Terminal").

## 3. Descargar los archivos y guardarlos juntos en una misma carpeta

Del repositorio de GitHub del proyecto necesitas, como mínimo, estos dos
archivos guardados **en la misma carpeta**:

- `simulacion_caidas_centros_geriatricos.py` (el programa)
- `plantilla_datos_centro.xlsx` (los datos del centro)

Después, ve a esa carpeta desde la terminal. Por ejemplo, si la has guardado
en el Escritorio dentro de una carpeta llamada `simulacion`:

```
cd Desktop/simulacion
```

(en Windows a veces es `cd Desktop\simulacion` o `cd OneDrive\Desktop\simulacion`,
depende de tu equipo)

## 4. Comprobar que tienes las librerías necesarias

Copia y pega esto en la terminal:

```
pip install numpy pandas matplotlib openpyxl
```

Si usaste Anaconda, seguramente ya las tienes todas y este paso no hará nada
(te dirá "requirement already satisfied", es normal).

## 5. Adaptar los datos a tu propio centro (opcional)

Si quieres aplicar la simulación a un centro distinto de OnaCare, abre
`plantilla_datos_centro.xlsx` y edita únicamente las hojas
`Parametros_Generales`, `Zonas_Layout`, `Deteccion` y `Escenarios` con los
datos de tu centro. La hoja `Hipotesis` es solo informativa (documenta cada
supuesto del modelo) y la hoja `Leer` es la portada de la plantilla; ninguna
de las dos hace falta tocarla. No es necesario modificar el código para
esto.

## 6. Ejecutar la simulación

```
python simulacion_caidas_centros_geriatricos.py
```

Por defecto, el programa busca `plantilla_datos_centro.xlsx` en la misma
carpeta desde la que lo ejecutas. Si le has puesto otro nombre o está en otra
ubicación, indícaselo como argumento:

```
python simulacion_caidas_centros_geriatricos.py ruta/a/mi_plantilla.xlsx
```

## 7. Qué esperar

Con las 300 réplicas configuradas por defecto en la plantilla, tarda solo
unos segundos. Al terminar, verás en la terminal un resumen con el tiempo de
respuesta medio, el percentil 95, el máximo y el % de traslados de los
cuatro escenarios (Anual, Día tipo, Noche tipo y Pico nocturno), tanto en
As-Is como en To-Be.

Además, se habrá creado una carpeta nueva llamada `salidas_simulacion` con:

- **5 gráficas en formato `.png`**: boxplot general As-Is vs. To-Be,
  comparativa día vs. noche, % de traslados, comparativa de los tres
  escenarios y detalle del Escenario 3.
- **1 archivo `resultados_simulacion.xlsx`**, con los resultados numéricos
  organizados en varias hojas (resumen general, una hoja por escenario,
  comparativa de reducción As-Is→To-Be, muestras de datos crudos y las
  cinco gráficas insertadas).

Para saber qué significa cada columna del Excel de resultados y cómo
interpretar los números (percentil 95, intervalo de confianza, etc.),
consulta el apartado 4 ("Cómo leer el archivo de resultados") del Anexo 1
de la memoria.

## 8. Qué hacer si da un error

Si la terminal muestra algo en rojo empezando por `Traceback` o `Error`:

- Si menciona `No module named` seguido del nombre de una librería (por
  ejemplo `openpyxl`, `numpy`, `pandas` o `matplotlib`), repite el paso 4.
- Si menciona que no encuentra `plantilla_datos_centro.xlsx`, comprueba que
  el archivo está en la misma carpeta desde la que ejecutas el comando, o
  indica su ruta completa como en el paso 6.
- Para cualquier otro error, revisa que no hayas cambiado el nombre de las
  hojas o de las columnas de la plantilla — el programa las busca por su
  nombre exacto.
