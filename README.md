# Simulación DES – Respuesta ante caídas en centros geriátricos (ONACARE)

Simulación de eventos discretos (DES, *Discrete Event Simulation*) desarrollada como parte de un Trabajo de Fin de Grado (TFG) para modelar el tiempo de respuesta del personal ante caídas de residentes en un centro geriátrico, comparando un escenario actual (**As-Is**) frente a un escenario de mejora propuesto (**To-Be**).

El modelo tiene en cuenta:

- **Zonas y layout del centro**: el tiempo de desplazamiento depende de en qué zona ocurre la caída y de dónde está el personal disponible más cercano en ese momento (no se usa un tiempo medio fijo).
- **Modo de detección configurable**: ronda manual, botón fijo, sensor ambiental automático o dispositivo llevado por el residente, cada uno con su propio tiempo de detección.
- **Tres escenarios de análisis**: Día tipo, Noche tipo y Pico nocturno de incidencias, simulados y reportados por separado, además de un análisis anual conjunto.
- **Réplicas independientes**: cada configuración se simula varias veces (según la semilla y el número de réplicas indicado) para obtener resultados estadísticamente robustos (media, percentil 95, máximo, % de traslados hospitalarios, etc.).

Todas las hipótesis del modelo están documentadas en la hoja `Hipotesis` del Excel de entrada y se repiten como comentarios en el bloque `HIPOTESIS` del script.

## Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `simulacion_onacare.py` | Script principal de la simulación (versión limpia). |
| `simulacion_onacare_comentado.py` | Misma simulación, con comentarios línea a línea explicando cada bloque de código (pensada para quien quiera entender o modificar el modelo). |
| `plantilla_datos_centro.xlsx` | Plantilla Excel de entrada donde se configuran los parámetros del centro, las zonas, los modos de detección y los escenarios. |
| `requirements.txt` | Dependencias de Python necesarias para ejecutar el script. |

## Requisitos

- Python 3.9 o superior
- Las librerías listadas en `requirements.txt`

## Instalación

```bash
git clone https://github.com/<tu-usuario>/<nombre-del-repositorio>.git
cd <nombre-del-repositorio>
pip install -r requirements.txt
```

(Opcionalmente, usa un entorno virtual: `python -m venv venv` y actívalo antes de instalar.)

## Uso

1. Rellena la plantilla `plantilla_datos_centro.xlsx` con los datos de tu centro (parámetros generales, zonas, modos de detección y escenarios). Cada hoja incluye las columnas necesarias y, en la hoja `Hipotesis`, la justificación de cada supuesto del modelo.
2. Ejecuta el script indicando la ruta al Excel de entrada:

```bash
python simulacion_onacare.py plantilla_datos_centro.xlsx
```

Si no se indica ninguna ruta, el script buscará por defecto un archivo llamado `plantilla_datos_centro.xlsx` en la carpeta actual.

3. Al finalizar, se crea una carpeta `salidas_simulacion/` con:
   - Un Excel de resultados (`resultados_simulacion.xlsx`) con las tablas estadísticas de cada escenario (Anual, Día tipo, Noche tipo, Pico nocturno), la comparativa As-Is vs To-Be, muestras de los datos crudos generados y una hoja con todas las gráficas.
   - Las gráficas generadas también se guardan como imágenes individuales en esa misma carpeta.

## Estructura de la plantilla Excel

- **Parametros_Generales**: nº de residentes, tasa de caídas, velocidad de desplazamiento del personal, tiempos de atención, probabilidades de traslado hospitalario, días y réplicas a simular, semilla aleatoria, etc.
- **Zonas_Layout**: zonas del centro, residentes por zona, distancias As-Is / To-Be, personal asignado por turno.
- **Deteccion**: tiempos asociados a cada modo de detección (ronda manual, botón, sensor, dispositivo), para As-Is y To-Be.
- **Escenarios**: configuración de los tres escenarios de análisis (Día tipo, Noche tipo, Pico nocturno de incidencias).
- **Hipotesis**: documentación de todas las hipótesis y supuestos utilizados en el modelo.

## Licencia

Este proyecto se distribuye bajo la licencia MIT — consulta el archivo [LICENSE](LICENSE) para más detalles. Puedes usar, modificar y redistribuir el código libremente, citando la fuente.

## Contexto académico

Este código forma parte de un Trabajo de Fin de Grado centrado en la mejora del protocolo de respuesta ante caídas en centros geriátricos mediante simulación de eventos discretos.
