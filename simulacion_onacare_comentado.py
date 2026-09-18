"""
SIMULACIÓN DE EVENTOS DISCRETOS (DES) v4 - Respuesta ante caídas
==================================================================
[Este bloque de texto entre comillas triples es un "docstring": un comentario
largo que documenta de qué trata todo el archivo. Python lo ignora al
ejecutar el código, es solo para que un humano lo lea.]

Mejoras respecto a la v1:
  - Modelo de zonas/layout: el tiempo de desplazamiento depende de dónde
    ocurre la caída y de dónde está el personal disponible más cercano en
    ese momento (no es un tiempo medio fijo).
  - Modo de detección configurable: ronda manual / botón fijo / sensor
    ambiental automático / dispositivo llevado por el residente.
  - Los 3 escenarios de la memoria (Día tipo, Noche tipo, Pico nocturno) se
    simulan y se reportan por separado, además del análisis anual conjunto.

Todas las hipótesis del modelo están documentadas en la hoja 'Hipotesis'
del Excel de entrada, y se repiten en el bloque HIPOTESIS de este archivo.
"""

# ---- IMPORTACIONES: cargar "cajas de herramientas" que el programa usará ----
import sys                       # Permite leer argumentos que se pasan al ejecutar el script desde la terminal (ej. el nombre del Excel).
from pathlib import Path         # Herramienta moderna de Python para trabajar con rutas de carpetas/archivos de forma segura en cualquier sistema operativo.

import numpy as np               # NumPy: librería para cálculos numéricos rápidos, vectores, matrices y números aleatorios.
import openpyxl                  # Librería para leer y escribir archivos Excel (.xlsx) directamente.
import pandas as pd              # Pandas: librería para tablas de datos tipo "hoja de cálculo" (DataFrames), fácil de agregar/filtrar/exportar.
import matplotlib.pyplot as plt  # Matplotlib: librería para dibujar gráficas (barras, boxplots, dispersión, etc.).

# ==========================================================================
# 0. LECTURA DE PARÁMETROS DESDE EL EXCEL
# ==========================================================================

def leer_plantilla(path):
    # "def" define una función: un bloque de código con nombre que se puede llamar
    # muchas veces. "path" es el parámetro de entrada: la ruta al Excel de la plantilla.
    wb = openpyxl.load_workbook(path, data_only=True)
    # Abre el archivo Excel indicado en "path". data_only=True significa: si una celda
    # tiene una fórmula (ej. =A1+A2), leer el valor calculado, no el texto de la fórmula.

    # --- Parametros_Generales ---
    ws = wb["Parametros_Generales"]
    # Selecciona, dentro del Excel ya abierto, la hoja llamada "Parametros_Generales".
    g = {}
    # Crea un diccionario vacío "g" (una tabla de tipo "nombre del parámetro" -> "valor").
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Recorre la hoja fila por fila, empezando en la fila 2 (se salta la cabecera de la fila 1).
        # values_only=True significa: devolver solo los valores de las celdas, no objetos de Excel complejos.
        nombre, valor = row[0], row[1]
        # De cada fila se toma la primera columna (nombre del parámetro) y la segunda (su valor).
        if nombre is None:
            continue
            # Si la columna "nombre" está vacía (fila en blanco), se salta esta fila y se pasa a la siguiente.
        g[nombre] = valor
        # Guarda en el diccionario "g" la pareja nombre->valor, ej. g["Nº de residentes totales"] = 80.

    general = {
        # Se construye un segundo diccionario, "general", con nombres cortos en español
        # y convirtiendo cada valor al tipo de dato correcto:
        "n_residentes": float(g["Nº de residentes totales"]),
        # float(...) convierte el valor a número decimal. Nº total de residentes del centro.
        "tasa_caidas_anual": float(g["Tasa de caídas"]),
        # Nº medio de caídas por residente y año (ej. 1.5 caídas/residente/año).
        "velocidad_m_min": float(g["Velocidad de marcha del personal"]),
        # Velocidad a la que camina el personal, en metros por minuto.
        "variabilidad": float(g["Variabilidad del desplazamiento (±)"]),
        # Un porcentaje/fracción que indica cuánto puede variar al azar el tiempo de desplazamiento
        # (a veces se anda más deprisa o más despacio de lo normal).
        "atencion_min": float(g["Tiempo de atención in situ (mínimo)"]),
        "atencion_max": float(g["Tiempo de atención in situ (máximo)"]),
        # Rango (mínimo y máximo) de minutos que tarda el cuidador en atender al residente
        # una vez ha llegado junto a él.
        "p_traslado_asis": float(g["Probabilidad de traslado hospitalario - As-Is"]),
        "p_traslado_tobe": float(g["Probabilidad de traslado hospitalario - To-Be"]),
        # Probabilidad (entre 0 y 1) de que una caída acabe en traslado al hospital,
        # una para el escenario actual (As-Is) y otra para el escenario mejorado (To-Be).
        "sim_dias": int(g["Nº de días a simular"]),
        # int(...) convierte a número entero. Cuántos días de calendario se simulan (ej. 365).
        "n_replicas": int(g["Nº de réplicas independientes"]),
        # Cuántas veces se repite la simulación completa (cada repetición = una "réplica"),
        # para obtener promedios estadísticamente fiables.
        "seed": int(g["Semilla aleatoria"]),
        # Número de partida del generador de números aleatorios, para que los resultados
        # sean reproducibles.
        "p_ocupado_rutina": float(g.get("Probabilidad de ocupación en tarea no interrumpible", 0.20)),
        # g.get(clave, valor_por_defecto): busca esa clave en el diccionario; si no existe
        # (por ejemplo, si tu Excel es una versión antigua sin esa fila), usa 0.20 por defecto
        # en vez de dar un error. Es la probabilidad de que un cuidador "libre" esté en
        # realidad ocupado con otro residente en una tarea que no puede interrumpir.
    }

    # --- Zonas_Layout ---
    ws = wb["Zonas_Layout"]
    # Cambia a la hoja "Zonas_Layout" del mismo Excel.
    zonas = []
    # Lista vacía donde se irá guardando una "ficha" (diccionario) por cada zona del centro.
    for row in ws.iter_rows(min_row=3, values_only=True):
        # Recorre esta hoja empezando en la fila 3 (aquí hay dos filas de cabecera, no una).
        nombre = row[0]
        # Primera columna de la fila: nombre de la zona (ej. "Planta 1").
        if nombre is None or nombre == "TOTAL":
            break
            # "break" corta el bucle por completo: si la fila está vacía o dice "TOTAL"
            # (la fila de suma al final de la tabla), se deja de leer, porque ya no hay más zonas.
        # Distancia intra-zona (moverse DENTRO de la misma zona, no hasta el control).
        intra_asis = float(row[6]) if len(row) > 6 and row[6] is not None else float(row[2]) / 3.0
        intra_tobe = float(row[7]) if len(row) > 7 and row[7] is not None else float(row[3]) / 3.0
        # "len(row) > 6" comprueba que la fila tenga al menos 7 columnas (índice 6 = columna G),
        # por si se usa una plantilla antigua que no llega a tener esas columnas.

        zonas.append({
            # Añade a la lista "zonas" un nuevo diccionario con los datos de esta zona:
            "nombre": nombre,
            "n_residentes": float(row[1]),          # Nº de residentes que viven en esa zona.
            "dist_asis": float(row[2]),              # Distancia de referencia en el escenario actual (metros).
            "dist_tobe": float(row[3]),              # Distancia de referencia en el escenario mejorado (metros).
            "personal_dia": int(row[4]),             # Nº de cuidadores asignados a esa zona en turno de día.
            "personal_noche": int(row[5]),           # Nº de cuidadores asignados a esa zona en turno de noche.
            "intra_asis": intra_asis,                # Distancia media DENTRO de la zona, escenario As-Is (metros).
            "intra_tobe": intra_tobe,                # Distancia media DENTRO de la zona, escenario To-Be (metros).
        })

    # --- Deteccion ---
    ws = wb["Deteccion"]
    # Cambia a la hoja "Deteccion".
    d = {}
    # Diccionario donde cada clave es el nombre de un parámetro de detección y el valor
    # es una pareja (tupla) con el valor en As-Is y el valor en To-Be.
    for row in ws.iter_rows(min_row=3, values_only=True):
        nombre, v_asis, v_tobe = row[0], row[1], row[2]
        # Cada fila trae: nombre del parámetro, su valor en la columna As-Is, su valor en To-Be.
        if nombre is None:
            continue
            # Salta filas vacías.
        d[nombre] = (v_asis, v_tobe)
        # Guarda la pareja de valores. (v_asis, v_tobe) es una tupla: una lista fija de 2 elementos.

    # --- Detección acústica incidental (oír un grito/golpe sin ver la caída) ---
    # Parámetros nuevos y opcionales: si no existen en la plantilla, se usa 0.0
    # (nadie detecta por esta vía) y el modelo se comporta exactamente como antes.
    p_oida_dia = float(d.get("Probabilidad de detección acústica (oída) - turno día", (0.0, 0.0))[0] or 0.0)
    # d.get(clave, (0.0, 0.0)): si el parámetro no existe en el Excel, se usa la pareja (0.0, 0.0)
    # por defecto. [0] coge el primer elemento de esa pareja (el valor de As-Is).
    # "or 0.0" es un truco de Python: si el valor leído fuera None o vacío, usa 0.0 en su lugar.
    p_oida_noche = float(d.get("Probabilidad de detección acústica (oída) - turno noche", (0.0, 0.0))[0] or 0.0)
    # Igual que la línea anterior pero para el turno de noche.
    oida_min = float(d.get("Detección acústica (oída) - mínimo (min)", (1.0, 1.0))[0] or 1.0)
    oida_max = float(d.get("Detección acústica (oída) - máximo (min)", (5.0, 5.0))[0] or 5.0)
    # Rango mínimo/máximo (en minutos) de lo que tarda alguien en darse cuenta si oye el incidente.

    deteccion = {
        # Diccionario final con toda la configuración de detección, separado por escenario.
        "ASIS": {
            "modo": d["Modo de detección"][0],
            # Cómo se detectan las caídas en el escenario actual: por ejemplo "ronda_manual_o_boton".
            "p_no_presenciada": float(d["Probabilidad de caída no presenciada"][0]),
            # Probabilidad de que nadie vea la caída en el momento en que ocurre.
            "ronda_dia": float(d["Intervalo de ronda - turno día (min)"][0]),
            "ronda_noche": float(d["Intervalo de ronda - turno noche (min)"][0]),
            # Cada cuántos minutos el personal hace una ronda de comprobación, de día y de noche.
            "boton_min": float(d["Detección con botón fijo - mínimo (min)"][0]),
            "boton_max": float(d["Detección con botón fijo - máximo (min)"][0]),
            # Rango de minutos que tarda en detectarse si el residente logra pulsar un botón fijo.
            "p_oida_dia": p_oida_dia,
            "p_oida_noche": p_oida_noche,
            "oida_min": oida_min,
            "oida_max": oida_max,
            # Se añaden aquí los parámetros de detección acústica calculados arriba.
        },
        "TOBE": {
            "modo": d["Modo de detección"][1],
            # Modo de detección en el escenario mejorado, ej. "sensor_ambiental_auto".
            "p_no_presenciada": float(d["Probabilidad de caída no presenciada"][1])
                if not isinstance(d["Probabilidad de caída no presenciada"][1], str) else None,
            # Igual que antes, pero con una comprobación extra: si en esa celda del Excel
            # hay texto en vez de un número (isinstance(..., str) = "es una cadena de texto"),
            # se guarda None (vacío) en vez de dar un error al convertir a número.
            "sensor_min_seg": float(d["Detección sensor ambiental - mínimo (seg)"][1]),
            "sensor_max_seg": float((d.get("Detección sensor ambiental - máximo (seg)")
                                     or d["Detección sensor ambiental -máximo (seg)"])[1]),
            # d.get(clave): busca esta clave exacta en el diccionario; si no la encuentra
            # devuelve None, y "or" hace que en ese caso se use la clave alternativa (mismo
            # dato, escrito sin espacio antes de "máximo"). [1] coge el valor To-Be de esa
            # pareja. Rango de tiempo en segundos que tarda el sensor automático en detectar
            # una caída.
            "p_activar_manual": d["Prob. de poder activar dispositivo manual"][1],
            # Probabilidad de que el residente consiga activar su dispositivo manual tras caerse.
            # fallback si modo='dispositivo_manual' y no se puede activar: usa las mismas
            # rondas manuales definidas para As-Is
            "ronda_dia": float(d["Intervalo de ronda - turno día (min)"][0]),
            "ronda_noche": float(d["Intervalo de ronda - turno noche (min)"][0]),
            # Se reutilizan los intervalos de ronda de As-Is como "plan B" si el dispositivo falla.
            # Probabilidad de que el sensor NO detecte la caída (falso negativo).
            # Fallback en ese caso: misma vía de detección acústica incidental
            # que en As-Is (el entorno acústico del edificio no cambia por
            # instalar sensores), y en último caso la ronda manual.
            "p_fallo_sensor": float(d.get("Probabilidad de fallo del sensor (falso negativo)", (0.0, 0.0))[1] or 0.0),
            # Probabilidad de que el sensor falle (por defecto 0.0 = nunca falla, si no está en el Excel).
            "p_oida_dia": p_oida_dia,
            "p_oida_noche": p_oida_noche,
            "oida_min": oida_min,
            "oida_max": oida_max,
        },
    }

    # --- Escenarios ---
    # Lee el nº de incidencias objetivo del Escenario 3 (pico nocturno) de la
    # hoja 'Escenarios' del Excel.
    n_incidencias_e3 = 8
    # Valor de reserva: se usa si la hoja 'Escenarios' no existe en el Excel,
    # o si la fila del Escenario 3 no trae ese dato en su celda.
    if "Escenarios" in wb.sheetnames:
        # wb.sheetnames es la lista de nombres de todas las hojas del Excel;
        # se comprueba que la hoja exista antes de intentar abrirla.
        ws = wb["Escenarios"]
        for row in ws.iter_rows(min_row=3, values_only=True):
            # Recorre la hoja fila por fila a partir de la fila 3 (las dos primeras filas
            # son cabecera).
            nombre = row[0]
            if nombre and str(nombre).startswith("Escenario 3"):
                # Busca la fila cuyo nombre de escenario empieza por "Escenario 3", sin
                # exigir que coincida el texto completo de la celda.
                if len(row) > 4 and row[4] not in (None, "-"):
                    n_incidencias_e3 = int(row[4])
                    # Columna E (índice 4) de esa fila: nº de incidencias objetivo del brote.
                break
                # Una vez encontrada la fila del Escenario 3, no hace falta seguir recorriendo
                # el resto de la hoja.
    general["n_incidencias_e3"] = n_incidencias_e3
    # Se guarda en el diccionario general para que main() pueda pasárselo más
    # adelante a ejecutar_escenario3().

    return general, zonas, deteccion
    # "return" hace que la función entregue estos tres resultados a quien la haya llamado:
    # el diccionario general, la lista de zonas, y el diccionario de detección.


# ==========================================================================
# 1. MODELO DE ZONAS Y DISTANCIAS
# ==========================================================================

def construir_modelo_zonas(zonas, escenario, turno):
    """
    Devuelve:
      - probs: probabilidad de que la caída ocurra en cada zona (prop. a nº residentes)
      - home_zone_staff: array con la zona base de cada miembro del personal de ese turno
      - dist_matrix: matriz de distancias entre zonas (incluye intra-zona, propia de cada zona)
    """
    n_res = np.array([z["n_residentes"] for z in zonas], dtype=float)
    # Crea una lista con el nº de residentes de cada zona (recorriendo la lista "zonas"),
    # y la convierte en un "array" de NumPy (una lista optimizada para cálculo numérico).
    probs = n_res / n_res.sum()
    # Divide cada valor entre la suma total de residentes: así "probs" es una lista de
    # probabilidades que suman 1 (100%), proporcional al tamaño de cada zona.

    dist_col = "dist_asis" if escenario == "ASIS" else "dist_tobe"
    # Decide qué columna de distancia usar según el escenario: si es "ASIS" usa "dist_asis",
    # si no (es "TOBE") usa "dist_tobe". Es un "if" escrito en una sola línea.
    dist_control = np.array([z[dist_col] for z in zonas], dtype=float)
    # Lista con la distancia de referencia de cada zona, según el escenario elegido.

    intra_col = "intra_asis" if escenario == "ASIS" else "intra_tobe"
    # Igual que arriba, pero para elegir la columna de distancia Intra-zona del escenario correcto.
    dist_intra = np.array([z[intra_col] for z in zonas], dtype=float)
    # Lista con la distancia media dentro de cada zona.
    n_zonas = len(zonas)
    # Cuenta cuántas zonas hay en total.
    dist_matrix = np.zeros((n_zonas, n_zonas))
    # Crea una matriz (tabla cuadrada) de ceros, de tamaño nº_zonas x nº_zonas, para rellenarla después.
    for i in range(n_zonas):
        for j in range(n_zonas):
            # Recorre cada combinación posible de zona origen (i) y zona destino (j).
            if i == j:
                dist_matrix[i, j] = dist_intra[i]
                # Si origen y destino son la misma zona, se usa la distancia intra-zona propia de esa zona "i".
            else:
                dist_matrix[i, j] = dist_control[i] + dist_control[j]
                # Si son zonas distintas, se suman las distancias de control de ambas zonas
                # (como decir: "de la zona i al pasillo central" + "del pasillo central a la zona j").

    personal_col = "personal_dia" if turno == "DIA" else "personal_noche"
    # Igual que antes: elige la columna de personal según si es turno de día o de noche.
    home_zone_staff = []
    # Lista vacía que irá conteniendo, para cada cuidador, en qué zona trabaja habitualmente.
    for zi, z in enumerate(zonas):
        # enumerate(zonas) recorre la lista "zonas" dando también el índice zi (0, 1, 2...) de cada zona.
        home_zone_staff += [zi] * z[personal_col]
        # Si esa zona tiene, por ejemplo, 3 cuidadores, añade el número de zona "zi" tres veces seguidas
        # a la lista. Al final, home_zone_staff tiene tantos elementos como cuidadores totales,
        # y cada elemento dice a qué zona pertenece ese cuidador.
    home_zone_staff = np.array(home_zone_staff)
    # Convierte la lista final en un array de NumPy, para poder hacer cálculos rápidos con ella.

    return probs, home_zone_staff, dist_matrix
    # Devuelve las tres piezas del "mapa" del centro para este escenario y turno.


# ==========================================================================
# 2. DETECCIÓN
# ==========================================================================

def tiempo_deteccion(escenario, det_params, es_noche, rng):
    # Calcula cuántos minutos pasan hasta que se detecta una caída concreta.
    # det_params: el diccionario de parámetros de detección (ASIS o TOBE).
    # es_noche: True/False, si el incidente ocurre en turno de noche.
    # rng: el generador de números aleatorios (el "dado" de la simulación).
    modo = det_params["modo"]
    # Lee qué modo de detección está configurado para este escenario.
    ronda = det_params["ronda_noche"] if es_noche else det_params["ronda_dia"]
    # Elige el intervalo de ronda correspondiente al turno actual (noche o día).

    if modo == "sensor_ambiental_auto":
        # CASO 1: hay un sensor automático instalado en el centro (propuesta del TO-BE).
        p_fallo = det_params.get("p_fallo_sensor", 0.0)
        # Probabilidad de que el sensor falle (por defecto 0 si no está definida).
        if rng.random() < p_fallo:
            # rng.random() saca un número al azar entre 0 y 1. Si es menor que p_fallo,
            # "ha ocurrido" el fallo del sensor en este sorteo concreto.
            p_oida = det_params.get("p_oida_noche", 0.0) if es_noche else det_params.get("p_oida_dia", 0.0)
            # Si el sensor falló, se comprueba la probabilidad de detección por sonido.
            if rng.random() < p_oida:
                return rng.uniform(det_params.get("oida_min", 1.0), det_params.get("oida_max", 5.0))
                # Si se oye la caída, el tiempo de detección es un número al azar entre
                # oida_min y oida_max minutos. rng.uniform(a, b) saca un decimal al azar entre a y b.
                # "return" sale inmediatamente de la función con este resultado.
            return rng.uniform(0, ronda)
            # Si tampoco se oye, se asume que se descubre en algún momento de la próxima ronda:
            # un tiempo al azar entre 0 minutos (justo antes de la ronda) y la duración de la ronda.
        return rng.uniform(det_params["sensor_min_seg"], det_params["sensor_max_seg"]) / 60.0
        # Si el sensor NO falla (caso normal), se sortea el tiempo de detección en segundos
        # y se divide entre 60 para convertirlo a minutos.

    if modo == "dispositivo_manual":
        # CASO 2: el residente lleva un dispositivo (botón/pulsera) que puede activar él mismo.
        p_activar = det_params.get("p_activar_manual") or 0.5
        # Probabilidad de que consiga activarlo (por defecto 0.5 si no está definida o es None/0).
        if rng.random() < p_activar:
            return rng.uniform(det_params["sensor_min_seg"], det_params["sensor_max_seg"]) / 60.0
            # Si logra activarlo, la detección es casi inmediata (tiempo del propio dispositivo).
        else:
            return rng.uniform(0, ronda)
            # Si no logra activarlo, se detecta en algún momento de la siguiente ronda manual.

    # modo por defecto: ronda_manual_o_boton
    # CASO 3 (si no es ninguno de los anteriores): detección mediante rondas del personal o botón fijo.
    p_no_presenciada = det_params.get("p_no_presenciada", 0.6)
    # Probabilidad de que nadie presencie la caída en directo (por defecto 0.6 si falta el dato).
    if rng.random() < p_no_presenciada:
        # Se sortea si esta caída concreta es "no presenciada".
        p_oida = det_params.get("p_oida_noche", 0.0) if es_noche else det_params.get("p_oida_dia", 0.0)
        # Probabilidad de detección por sonido, según el turno.
        if rng.random() < p_oida:
            return rng.uniform(det_params.get("oida_min", 1.0), det_params.get("oida_max", 5.0))
            # Si se oye, tiempo de detección al azar entre oida_min y oida_max.
        return rng.uniform(0, ronda)
        # Si no se oye ni se presencia, se detecta en algún momento de la próxima ronda.
    else:
        return rng.uniform(det_params["boton_min"], det_params["boton_max"])
        # Si SÍ es presenciada (o el residente pulsa un botón fijo cercano), tiempo de
        # detección al azar entre boton_min y boton_max minutos (mucho más rápido).


# ==========================================================================
# 3. NÚCLEO: RESOLVER UNA LISTA DE INCIDENTES CON EL MODELO DE ZONAS
# ==========================================================================

def resolver_incidentes(t_llegadas, es_noche_flags, escenario, general, zonas, deteccion, rng):
    """
    t_llegadas: array de instantes de caída (minutos, dentro del horizonte simulado)
    es_noche_flags: array booleano paralelo, turno DIA/NOCHE de cada incidente
    Devuelve una lista de dicts, un registro por incidente.
    """
    det_params = deteccion[escenario]
    # Selecciona el diccionario de parámetros de detección correspondiente (ASIS o TOBE).
    filas = []
    # Lista vacía donde se irá guardando un resultado (fila) por cada caída procesada.

    # pre-construir modelo de zonas para turno día y turno noche (distinto personal)
    modelos = {
        "DIA": construir_modelo_zonas(zonas, escenario, "DIA"),
        "NOCHE": construir_modelo_zonas(zonas, escenario, "NOCHE"),
    }
    # Se construyen de antemano (una sola vez, no en cada caída) los "mapas" de zonas
    # para turno día y turno noche, porque el personal disponible cambia según el turno.
    free_time = {
        "DIA": np.zeros(len(modelos["DIA"][1])),
        "NOCHE": np.zeros(len(modelos["NOCHE"][1])),
    }
    # Para cada turno, se crea un array con un "0" por cada cuidador: representa el
    # instante (en minutos) en que cada cuidador queda libre de nuevo. Al empezar,
    # todos están libres desde el minuto 0. modelos["DIA"][1] es el "home_zone_staff"
    # de día, y su longitud (len) es el número de cuidadores de día.

    for t, es_noche in zip(t_llegadas, es_noche_flags):
        # zip empareja cada instante de caída "t" con su correspondiente "es_noche".
        # El bucle procesa las caídas una a una, en el orden en que llegan (cronológico,
        # porque t_llegadas viene ya ordenado de menor a mayor).
        turno = "NOCHE" if es_noche else "DIA"
        # Determina el turno de esta caída concreta.
        probs, home_zone_staff, dist_matrix = modelos[turno]
        # Recupera el "mapa" de zonas ya calculado para ese turno.
        ft = free_time[turno]
        # Recupera el array de "cuándo queda libre cada cuidador" para ese turno
        # (esta variable se irá actualizando a medida que se resuelven caídas).

        zona_incidente = rng.choice(len(zonas), p=probs)
        # Sortea en qué zona ocurre esta caída. rng.choice(n, p=probs) elige un número
        # entre 0 y n-1, con las probabilidades indicadas en "probs" (proporcionales
        # al nº de residentes de cada zona).

        deteccion_min = tiempo_deteccion(escenario, det_params, es_noche, rng)
        # Llama a la función del Bloque 2 para calcular cuánto tarda en detectarse esta caída.
        t_listo = t + deteccion_min
        # Instante en que la caída ya ha sido detectada (hora de la caída + tiempo de detección).

        # Personal "libre" según el histórico de caídas atendidas (no está
        # respondiendo a otra caída en este momento)
        libres_todos = np.where(ft <= t_listo)[0]
        # np.where(condición) devuelve los índices donde la condición es verdadera.
        # Aquí: qué cuidadores ya han quedado libres (su "ft") antes o en el momento
        # en que se detecta esta nueva caída. [0] extrae la lista de índices del resultado.

        # De ese personal libre, una parte puede estar en ese instante
        # ocupada en una tarea asistencial no interrumpible con otro
        # residente (aseo, alimentación, movilización con grúa,
        # administración de medicación) y no puede acudir de inmediato;
        # en ese caso debe ir otro compañero libre y disponible.
        # [SUPUESTO - ver parámetro 'p_ocupado_rutina' y su justificación]
        if len(libres_todos) > 0:
            # Si hay al menos un cuidador "libre" según el histórico...
            p_ocupado = general.get("p_ocupado_rutina", 0.0)
            # Probabilidad de que, de todas formas, esté ocupado en otra tarea rutinaria.
            disponible_mask = rng.random(len(libres_todos)) >= p_ocupado
            # Genera un número al azar para cada cuidador libre, y se queda con los que
            # superan la probabilidad de estar ocupado (es decir, los que sí están realmente
            # disponibles). Esto es un "filtro" booleano (True/False por cada elemento).
            libres = libres_todos[disponible_mask]
            # Aplica el filtro: "libres" son los cuidadores libres Y realmente disponibles ahora mismo.
        else:
            libres = libres_todos
            # Si no había ningún cuidador libre, "libres" se queda vacío igualmente.

        if len(libres) > 0:
            # CASO A: hay al menos un cuidador realmente disponible.
            distancias_libres = dist_matrix[home_zone_staff[libres], zona_incidente]
            # Calcula la distancia desde la zona base de cada cuidador disponible hasta
            # la zona donde ha ocurrido la caída, usando la matriz de distancias.
            idx = libres[np.argmin(distancias_libres)]
            # np.argmin encuentra la posición del valor más pequeño (la menor distancia).
            # "idx" es el índice del cuidador elegido: el más cercano de los disponibles.
            espera = 0.0
            # No hay espera extra: el cuidador puede salir inmediatamente.
        elif len(libres_todos) > 0:
            # CASO B: no hay nadie "realmente disponible", pero sí había cuidadores
            # libres según el histórico (todos resultaron estar ocupados en tareas rutinarias).
            # Todo el personal libre según el histórico de caídas resultó
            # (por puro azar del sorteo) estar ocupado en tarea rutinaria.
            # Como no se modela cuánto dura esa tarea, se opta por el más
            # cercano de ellos en vez de esperar a alguien que lleva más
            # tiempo inactivo (evita esperas negativas o irreales).
            distancias_libres = dist_matrix[home_zone_staff[libres_todos], zona_incidente]
            idx = libres_todos[np.argmin(distancias_libres)]
            # Se coge igualmente al más cercano de ese grupo, como simplificación razonada.
            espera = 0.0
        else:
            # CASO C: nadie está libre en absoluto (todos siguen atendiendo otras caídas).
            # Nadie realmente disponible: ni por estar ya respondiendo a
            # otra caída, ni por estar en una tarea rutinaria no
            # interrumpible. Se espera a que quede libre el primero según
            # el histórico de caídas (no se modela por separado cuándo
            # termina la tarea rutinaria - simplificación adicional,
            # documentada como tal en la memoria).
            idx = int(np.argmin(ft))
            # Se elige el cuidador que quedará libre ANTES que los demás (el mínimo de "ft").
            espera = ft[idx] - t_listo
            # El tiempo de espera es la diferencia entre cuándo quedará libre ese cuidador
            # y el momento en que se detectó la caída (t_listo).

        dist = dist_matrix[home_zone_staff[idx], zona_incidente]
        # Distancia real que debe recorrer el cuidador elegido, desde su zona base hasta el incidente.
        ruido = 1 + rng.uniform(-general["variabilidad"], general["variabilidad"])
        # Un factor multiplicador aleatorio alrededor de 1 (ej. entre 0.9 y 1.1 si variabilidad=0.1),
        # que simula que a veces se camina algo más rápido o más despacio de lo normal.
        desplazamiento = (dist / general["velocidad_m_min"]) * ruido
        # Tiempo de desplazamiento = distancia / velocidad de marcha, ajustado por el ruido aleatorio.

        atencion = rng.uniform(general["atencion_min"], general["atencion_max"])
        # Tiempo de atención in situ, sorteado entre el mínimo y el máximo configurados.

        tiempo_respuesta = deteccion_min + espera + desplazamiento
        # El "tiempo de respuesta" ante la caída es la suma de: tiempo hasta detectarla,
        # tiempo de espera (si no había nadie disponible) y tiempo de desplazamiento hasta llegar.
        # (El tiempo de atención NO se incluye aquí: es lo que ocurre DESPUÉS de haber llegado).
        ft[idx] = t_listo + espera + desplazamiento + atencion
        # Se actualiza cuándo quedará libre de nuevo el cuidador elegido: el momento en que
        # se detectó la caída, más la espera, más el desplazamiento, más el tiempo de atención.

        p_trasl = general["p_traslado_asis"] if escenario == "ASIS" else general["p_traslado_tobe"]
        # Elige la probabilidad de traslado hospitalario según el escenario.
        traslado = rng.random() < p_trasl
        # Sortea si esta caída concreta termina en traslado al hospital (True/False).

        filas.append({
            # Guarda un diccionario con todos los datos de esta caída resuelta:
            "t_min": t, "es_noche": bool(es_noche), "zona": zonas[zona_incidente]["nombre"],
            "deteccion_min": deteccion_min, "espera_min": espera,
            "desplazamiento_min": desplazamiento, "tiempo_respuesta_min": tiempo_respuesta,
            "traslado_hospital": traslado, "escenario": escenario,
        })

    return filas
    # Cuando ya se han procesado TODAS las caídas del bucle, se devuelve la lista completa de resultados.


# ==========================================================================
# 4. GENERADORES DE LLEGADAS - ESCENARIOS 1/2 (año completo) Y 3 (pico 4h)
# ==========================================================================

# Ventanas horarias de los escenarios (minutos desde medianoche), coherentes
# con la hoja 'Escenarios' del Excel: Día 08:00-20:00, Noche 20:00-08:00.
HORA_INICIO_NOCHE = 20 * 60
# Constante: el turno de noche empieza a las 20:00h, expresado en minutos desde medianoche (1200 min).
HORA_FIN_NOCHE = 8 * 60
# Constante: el turno de noche termina a las 8:00h (480 minutos desde medianoche).

# Reparto real día/noche de las caídas [FUENTE: Samper Lamenca et al.,
# Gerokomos 2016 - 48,1% turno mañana + 35% turno tarde = 83,1% día;
# 16,9% turno noche; n=160 caídas, residencia española real]
FRACCION_CAIDAS_DIA = 0.831
FRACCION_CAIDAS_NOCHE = 0.169
# Constantes tomadas de un estudio real: qué porcentaje de caídas ocurre de día y de noche.


def generar_llegadas_anual(general, rng):
    """
    Genera las caídas de un año simulado como dos procesos de Poisson
    independientes - turno día (08-20h) y turno noche (20-08h) - calibrados
    con el reparto real día/noche documentado en Samper Lamenca et al.
    (Gerokomos, 2016; n=160 caídas, residencia española real): 48,1% turno
    mañana + 35% turno tarde = 83,1% turno día; 16,9% turno noche.
    (Sustituye el supuesto de reparto uniforme por hora del día de
    versiones anteriores del modelo.)
    """
    lam_anual = general["tasa_caidas_anual"] * general["n_residentes"]
    # "lambda" (tasa media) de caídas al año para TODO el centro: tasa por residente
    # multiplicada por el número de residentes. Ej. 1.5 caídas/residente/año × 80 residentes = 120/año.
    sim_dias = general["sim_dias"]
    # Número de días que se van a simular (puede ser menos de 365, ej. para pruebas rápidas).

    lam_dia_periodo = lam_anual * FRACCION_CAIDAS_DIA * sim_dias / 365
    # Nº medio esperado de caídas DE DÍA en el periodo simulado: se toma la tasa anual,
    # se multiplica por la fracción que ocurre de día, y se ajusta proporcionalmente
    # a cuántos días se están simulando (si simulas medio año, esperas la mitad de caídas).
    lam_noche_periodo = lam_anual * FRACCION_CAIDAS_NOCHE * sim_dias / 365
    # Lo mismo pero para las caídas de noche.

    n_dia = rng.poisson(lam_dia_periodo)
    # Sortea, con una distribución de Poisson (la estándar para "cuántos eventos raros
    # ocurren en un periodo"), el NÚMERO EXACTO de caídas de día que habrá en esta réplica.
    n_noche = rng.poisson(lam_noche_periodo)
    # Igual para las caídas de noche.

    # Turno día: cada caída, un día aleatorio del horizonte simulado y un
    # instante uniforme dentro de la ventana 08:00-20:00 de ese día.
    dias_dia = rng.integers(0, sim_dias, n_dia)
    # Para cada una de las n_dia caídas, sortea en qué día del periodo simulado ocurre
    # (un número entero entre 0 y sim_dias-1).
    minutos_dia = rng.uniform(HORA_FIN_NOCHE, HORA_INICIO_NOCHE, n_dia)
    # Para cada caída, sortea a qué minuto del día ocurre, dentro de la ventana de 08:00 (480 min)
    # a 20:00 (1200 min).
    llegadas_dia = dias_dia * 1440 + minutos_dia
    # Convierte "día + minuto dentro del día" en un único número de minutos desde el inicio
    # de la simulación (1440 = minutos que tiene un día completo).

    # Turno noche: cada caída, una "noche" aleatoria (empieza a las 20:00
    # de un día y termina a las 8:00 del día siguiente) y un instante
    # uniforme dentro de esas 12h.
    dias_noche = rng.integers(0, sim_dias, n_noche)
    # Sortea, para cada caída nocturna, en qué día "empieza" la noche correspondiente.
    minutos_noche = rng.uniform(0, 720, n_noche)
    # Sortea el minuto dentro de la ventana nocturna, que dura 12h = 720 minutos.
    llegadas_noche = dias_noche * 1440 + HORA_INICIO_NOCHE + minutos_noche
    # Convierte a minutos absolutos: día × 1440 + hora de inicio de noche (20:00) + minuto sorteado.
    # Esto permite que una "noche" cruce la medianoche correctamente.

    llegadas = np.concatenate([llegadas_dia, llegadas_noche])
    # Junta en una sola lista los instantes de todas las caídas (día + noche).
    es_noche = np.concatenate([np.zeros(n_dia, dtype=bool), np.ones(n_noche, dtype=bool)])
    # Crea una lista paralela de True/False indicando si cada caída es de noche o de día
    # (np.zeros con dtype=bool da una lista de "False"; np.ones da una lista de "True").

    horizonte = sim_dias * 24 * 60
    # Duración total del periodo simulado, en minutos.
    mask = llegadas < horizonte
    # Filtro: algunas caídas nocturnas generadas en el último día podrían "caer" después
    # del final del horizonte simulado (porque la noche se extiende al día siguiente);
    # esta máscara descarta esas caídas que se saldrían del periodo.
    orden = np.argsort(llegadas[mask])
    # Calcula el orden cronológico (de más temprano a más tardío) de las caídas que sí quedan dentro.
    return llegadas[mask][orden], es_noche[mask][orden]
    # Devuelve las llegadas y sus indicadores de turno, ya filtrados y ordenados por tiempo.


def generar_llegadas_escenario3(n_incidencias_objetivo, ventana_min, rng):
    """Genera de forma independiente, réplica a réplica, un nº de incidencias
    de Poisson con media n_incidencias_objetivo, distribuidas uniformemente
    dentro de la ventana de 'ventana_min' minutos (p.ej. 4 horas de brote)."""
    n = rng.poisson(n_incidencias_objetivo)
    # Sortea cuántas caídas ocurren en este "brote" concentrado (con media n_incidencias_objetivo).
    n = max(n, 0)
    # Por seguridad, si por algún motivo saliera un número negativo (no debería con Poisson,
    # pero es una salvaguarda), se fuerza a que sea como mínimo 0.
    llegadas = np.sort(rng.uniform(0, ventana_min, n))
    # Sortea el instante (en minutos) de cada una de esas caídas dentro de la ventana del
    # brote (por defecto 240 minutos = 4 horas), y las ordena de menor a mayor.
    es_noche = np.ones(n, dtype=bool)  # todo el escenario 3 ocurre en turno de noche
    # Todas las caídas de este escenario se marcan como nocturnas.
    return llegadas, es_noche


# ==========================================================================
# 5. ORQUESTACIÓN: EJECUTAR RÉPLICAS PARA CADA ESCENARIO
# ==========================================================================

def ejecutar_anual(general, zonas, deteccion, escenario):
    # Ejecuta la simulación del año completo, repetida "n_replicas" veces, para un escenario dado.
    resultados = []
    # Lista donde se acumularán los resultados de TODAS las réplicas.
    for r in range(general["n_replicas"]):
        # Bucle que se repite una vez por cada réplica (r = 0, 1, 2, ..., n_replicas-1).
        rng = np.random.default_rng(general["seed"] + r)
        # Crea un generador aleatorio nuevo para esta réplica, con una semilla distinta
        # (semilla base + número de réplica), de forma que cada réplica sea independiente
        # pero el conjunto completo siga siendo reproducible.
        t_llegadas, es_noche = generar_llegadas_anual(general, rng)
        # Genera las caídas de esta réplica concreta (Bloque 4).
        filas = resolver_incidentes(t_llegadas, es_noche, escenario, general, zonas, deteccion, rng)
        # Resuelve cada una de esas caídas (Bloque 3): a quién se asigna, cuánto tarda, etc.
        for f in filas:
            f["replica"] = r
            # A cada fila de resultado se le añade a qué réplica pertenece, para poder
            # luego agrupar y comparar réplicas entre sí.
        resultados.extend(filas)
        # Añade todas las filas de esta réplica a la lista general de resultados.
    return pd.DataFrame(resultados)
    # Convierte la lista de diccionarios en una tabla de pandas (DataFrame), mucho más
    # cómoda para analizar, filtrar y exportar después.


def ejecutar_escenario3(general, zonas, deteccion, escenario, n_incidencias_objetivo=3,
                         ventana_min=240, n_replicas=None):
    # Igual que la función anterior, pero para el "brote" de 4 horas (Escenario 3).
    # n_incidencias_objetivo=3 y ventana_min=240 son valores por defecto (se pueden cambiar
    # al llamar a la función, pero si no se especifican, se usan estos).
    n_replicas = n_replicas or general["n_replicas"]
    # Si no se especifica un nº de réplicas distinto, se usa el mismo que en el resto del estudio.
    # offset de semilla distinto para no reutilizar la misma secuencia que el escenario anual
    resultados = []
    for r in range(n_replicas):
        rng = np.random.default_rng(general["seed"] + 100_000 + r)
        # Se usa una semilla desplazada en +100.000 respecto a la del escenario anual, para
        # asegurarse de que esta simulación usa una secuencia de números aleatorios distinta
        # (si no, se estarían repitiendo exactamente los mismos "sorteos" que en el otro escenario).
        t_llegadas, es_noche = generar_llegadas_escenario3(n_incidencias_objetivo, ventana_min, rng)
        # Genera las caídas del brote para esta réplica.
        if len(t_llegadas) == 0:
            continue
            # Si en esta réplica concreta no hubo ninguna caída (puede pasar, es al azar),
            # se pasa directamente a la siguiente réplica sin hacer nada más.
        filas = resolver_incidentes(t_llegadas, es_noche, escenario, general, zonas, deteccion, rng)
        for f in filas:
            f["replica"] = r
        resultados.extend(filas)
    return pd.DataFrame(resultados)


# ==========================================================================
# 6. AGREGACIÓN Y ESTADÍSTICOS
# ==========================================================================

def resumen_por_replica(df, n_replicas):
    # Recibe la tabla completa de resultados (todas las réplicas juntas) y calcula,
    # para CADA réplica por separado, un resumen de sus métricas principales.
    if len(df) == 0:
        # Si la tabla está completamente vacía (no hubo ninguna caída en ninguna réplica)...
        idx = pd.Index(range(n_replicas), name="replica")
        # Crea un índice del 0 al n_replicas-1, llamado "replica".
        return pd.DataFrame({"n_incidentes": 0, "tiempo_respuesta_medio": np.nan,
                              "pct_traslados": np.nan}, index=idx).reset_index()
        # Devuelve una tabla con 0 incidentes y valores "nan" (Not a Number = vacío/indefinido)
        # para las métricas que no se pueden calcular sin datos.
    g = df.groupby("replica").agg(
        # df.groupby("replica") agrupa todas las filas por el número de réplica a la que pertenecen.
        # .agg(...) calcula, dentro de cada grupo (cada réplica), las siguientes métricas:
        n_incidentes=("t_min", "size"),
        # Cuenta cuántas filas (caídas) hay en cada réplica, usando la columna "t_min" solo
        # como referencia para contar.
        tiempo_respuesta_medio=("tiempo_respuesta_min", "mean"),
        # Calcula la media del tiempo de respuesta dentro de cada réplica.
        pct_traslados=("traslado_hospital", "mean"),
        # Como "traslado_hospital" es True/False (1/0), su media es directamente
        # la PROPORCIÓN de caídas con traslado en esa réplica.
    ).reindex(range(n_replicas), fill_value=0).reset_index().rename(columns={"index": "replica"})
    # .reindex(...) se asegura de que aparezcan TODAS las réplicas del 0 al n_replicas-1,
    # incluso si alguna réplica no tuvo ninguna caída (fill_value=0 rellena esos casos con 0).
    # .reset_index() convierte el agrupamiento de vuelta en columnas normales de la tabla.
    # .rename(...) renombra la columna "index" (creada automáticamente) a "replica".
    return g
    # Devuelve la tabla resumen: una fila por réplica.


def ic_95(serie):
    # Calcula el intervalo de confianza del 95% de la media de una serie de valores.
    serie = serie.dropna()
    # Elimina los valores vacíos (NaN) de la serie antes de calcular nada.
    if len(serie) == 0:
        return np.nan, np.nan, np.nan
        # Si no queda ningún valor válido, devuelve "vacío" en los tres resultados.
    m = serie.mean()
    # Calcula la media de la serie.
    err = 1.96 * serie.std(ddof=1) / np.sqrt(len(serie)) if len(serie) > 1 else 0.0
    # Calcula el "margen de error" del intervalo de confianza al 95%: 1.96 es el valor
    # estándar (de la distribución normal) para un 95% de confianza; serie.std(ddof=1)
    # es la desviación típica (cuánto varían los datos); se divide entre la raíz cuadrada
    # del número de datos (a más datos, más precisión, menor margen de error).
    # Si solo hay 1 dato, no se puede calcular una desviación fiable, así que el margen es 0.
    return m, m - err, m + err
    # Devuelve: la media, el límite inferior del intervalo, y el límite superior.


def estadisticos_escenario(df, n_replicas):
    """Devuelve un dict con las métricas clave de un (escenario, ASIS/TOBE)."""
    resumen = resumen_por_replica(df, n_replicas)
    # Obtiene el resumen por réplica calculado antes.
    m, lo, hi = ic_95(resumen["tiempo_respuesta_medio"])
    # Calcula el intervalo de confianza del tiempo de respuesta medio, usando las medias
    # de cada réplica como "muestra" (esto es lo que da robustez estadística al resultado).
    mt, lot, hit = ic_95(resumen["pct_traslados"])
    # Igual, pero para el porcentaje de traslados.
    return {
        "n_incidentes_medio": resumen["n_incidentes"].mean() if len(df) else 0,
        # Nº medio de incidentes por réplica (0 si no hay datos en absoluto).
        "tiempo_respuesta_medio": m, "tr_ic95_lo": lo, "tr_ic95_hi": hi,
        # Media del tiempo de respuesta y su intervalo de confianza (límite inferior/superior).
        "tiempo_respuesta_mediana": df["tiempo_respuesta_min"].median() if len(df) else np.nan,
        # Mediana: el valor "del medio" de todos los tiempos de respuesta individuales
        # (calculada sobre TODAS las caídas de todas las réplicas juntas, no por réplica).
        "tiempo_respuesta_p95": df["tiempo_respuesta_min"].quantile(0.95) if len(df) else np.nan,
        # Percentil 95: el valor por debajo del cual está el 95% de los casos (útil para
        # ver qué tan malos son los "peores casos" sin que un único extremo distorsione la media).
        "tiempo_respuesta_max": df["tiempo_respuesta_min"].max() if len(df) else np.nan,
        # El tiempo de respuesta más alto registrado en toda la simulación.
        "pct_traslados": 100 * mt if not np.isnan(mt) else np.nan,
        # Porcentaje medio de traslados (se multiplica por 100 porque antes era una proporción
        # entre 0 y 1, y aquí se pasa a formato "porcentaje").
        "pct_traslados_ic95_lo": 100 * lot if not np.isnan(lot) else np.nan,
        "pct_traslados_ic95_hi": 100 * hit if not np.isnan(hit) else np.nan,
        # Límites del intervalo de confianza del porcentaje de traslados.
    }


# ==========================================================================
# 7. GRÁFICAS
# ==========================================================================

def generar_graficas(datos, carpeta):
    # Genera y guarda 5 gráficas en formato de imagen (.png), dentro de la carpeta indicada.
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    # Aplica un estilo visual a las gráficas (fondo con cuadrícula clara), si está disponible
    # en la versión de matplotlib instalada; si no, usa el estilo por defecto.
    carpeta = Path(carpeta)
    # Se asegura de que "carpeta" sea un objeto de tipo Path (ruta), no solo un texto.

    df_asis_anual = datos["anual"]["ASIS"]
    df_tobe_anual = datos["anual"]["TOBE"]
    # Extrae del diccionario "datos" las tablas de resultados anuales de ambos escenarios.

    # --- Figura 1: boxplot general (todo el año) ---
    fig, ax = plt.subplots(figsize=(7, 5))
    # Crea una figura nueva (fig) y un área de dibujo (ax) de 7x5 pulgadas.
    caja = [df_asis_anual["tiempo_respuesta_min"], df_tobe_anual["tiempo_respuesta_min"]]
    # Prepara los dos conjuntos de datos (tiempos de respuesta AS-IS y TO-BE) para el boxplot.
    try:
        ax.boxplot(caja, tick_labels=["AS-IS", "TO-BE"], showfliers=False)
        # Dibuja un "diagrama de caja" (boxplot), que muestra mediana, cuartiles y rango
        # de los tiempos de respuesta, sin mostrar los valores atípicos extremos (showfliers=False).
    except TypeError:
        ax.boxplot(caja, labels=["AS-IS", "TO-BE"], showfliers=False)
        # Alternativa por si la versión de matplotlib instalada usa el parámetro "labels"
        # en vez de "tick_labels" (diferencias entre versiones de la librería).
    ax.set_ylabel("Tiempo de respuesta ante caída (min)")
    # Pone el título del eje vertical.
    ax.set_title("Distribución del tiempo de respuesta - año completo\n(simulación DES por zonas)")
    # Pone el título de la gráfica ("\n" salta de línea).
    fig.tight_layout()
    # Ajusta automáticamente los márgenes para que no se corte ningún texto.
    fig.savefig(carpeta / "fig1_boxplot_general.png", dpi=200)
    # Guarda la imagen en la carpeta indicada, con una resolución de 200 puntos por pulgada.
    plt.close(fig)
    # Cierra la figura para liberar memoria (importante cuando se generan varias gráficas seguidas).

    # --- Figura 2: día (Escenario 1) vs noche (Escenario 2) ---
    fig, ax = plt.subplots(figsize=(7, 5))
    x = np.arange(2)
    # Crea las posiciones 0 y 1 en el eje X (una por cada grupo de barras).
    width = 0.35
    # Ancho de cada barra.
    med_asis = [datos["stats"]["Escenario 1 - Día tipo"]["ASIS"]["tiempo_respuesta_medio"],
                datos["stats"]["Escenario 2 - Noche tipo"]["ASIS"]["tiempo_respuesta_medio"]]
    # Recupera el tiempo de respuesta medio de AS-IS en los escenarios 1 y 2.
    med_tobe = [datos["stats"]["Escenario 1 - Día tipo"]["TOBE"]["tiempo_respuesta_medio"],
                datos["stats"]["Escenario 2 - Noche tipo"]["TOBE"]["tiempo_respuesta_medio"]]
    # Lo mismo para TO-BE.
    ax.bar(x - width / 2, med_asis, width, label="AS-IS")
    # Dibuja las barras de AS-IS, desplazadas un poco a la izquierda de cada posición.
    ax.bar(x + width / 2, med_tobe, width, label="TO-BE")
    # Dibuja las barras de TO-BE, desplazadas a la derecha (así quedan una junto a otra, no encima).
    ax.set_xticks(x)
    ax.set_xticklabels(["Escenario 1\nDía tipo", "Escenario 2\nNoche tipo"])
    # Pone las etiquetas de cada grupo de barras en el eje X.
    ax.set_ylabel("Tiempo de respuesta medio (min)")
    ax.set_title("Escenario 1 (día) vs Escenario 2 (noche)")
    ax.legend()
    # Muestra la leyenda (qué color/barra corresponde a AS-IS y cuál a TO-BE).
    fig.tight_layout()
    fig.savefig(carpeta / "fig2_dia_vs_noche.png", dpi=200)
    plt.close(fig)

    # --- Figura 3: % traslados anual ---
    fig, ax = plt.subplots(figsize=(6, 5))
    valores = [datos["stats"]["Anual"]["ASIS"]["pct_traslados"],
               datos["stats"]["Anual"]["TOBE"]["pct_traslados"]]
    # Recupera el % de traslados anual de ambos escenarios.
    ax.bar(["AS-IS", "TO-BE"], valores, color=["#c0392b", "#27ae60"])
    # Dibuja dos barras con colores concretos: rojo para AS-IS, verde para TO-BE
    # (códigos hexadecimales de color).
    ax.set_ylabel("% de caídas con traslado a hospital")
    ax.set_title("Tasa de traslados hospitalarios por caída (anual)")
    for i, v in enumerate(valores):
        ax.text(i, v + 0.5, f"{v:.1f}%", ha="center")
        # Para cada barra, escribe encima el valor exacto en texto (ej. "12.3%"),
        # un poco por encima de la barra (v + 0.5), centrado horizontalmente.
        # f"{v:.1f}%" es una "f-string": inserta el número v con 1 decimal, seguido de "%".
    fig.tight_layout()
    fig.savefig(carpeta / "fig3_pct_traslados.png", dpi=200)
    plt.close(fig)

    # --- Figura 4: comparativa de los 3 escenarios (cuál es más perjudicial) ---
    nombres_esc = ["Escenario 1\nDía tipo", "Escenario 2\nNoche tipo", "Escenario 3\nPico nocturno"]
    claves_esc = ["Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]
    # Dos listas paralelas: nombres "bonitos" para mostrar y claves reales del diccionario "datos".
    med_asis = [datos["stats"][k]["ASIS"]["tiempo_respuesta_medio"] for k in claves_esc]
    med_tobe = [datos["stats"][k]["TOBE"]["tiempo_respuesta_medio"] for k in claves_esc]
    # Recupera, para cada uno de los 3 escenarios, el tiempo de respuesta medio de AS-IS y TO-BE.
    # "[... for k in claves_esc]" es una "list comprehension": construye una lista recorriendo
    # cada elemento de claves_esc y calculando algo con él, en una sola línea.

    fig, ax = plt.subplots(figsize=(8.5, 5.5))
    x = np.arange(3)
    width = 0.35
    b1 = ax.bar(x - width / 2, med_asis, width, label="AS-IS", color="#c0392b")
    b2 = ax.bar(x + width / 2, med_tobe, width, label="TO-BE", color="#27ae60")
    # Dibuja las barras de los 3 escenarios, AS-IS en rojo y TO-BE en verde.
    # "b1" y "b2" guardan una referencia a las barras dibujadas, para poder etiquetarlas después.
    ax.set_xticks(x)
    ax.set_xticklabels(nombres_esc)
    ax.set_ylabel("Tiempo de respuesta medio (min)")
    ax.set_title("Comparativa de los 3 escenarios - ¿cuál es más perjudicial?")
    ax.legend()
    ax.bar_label(b1, fmt="%.1f", padding=3)
    ax.bar_label(b2, fmt="%.1f", padding=3)
    # Escribe automáticamente el valor numérico encima de cada barra, con 1 decimal.
    fig.tight_layout()
    fig.savefig(carpeta / "fig4_comparativa_escenarios.png", dpi=200)
    plt.close(fig)

    # --- Figura 5: detalle del Escenario 3 (pico nocturno) ---
    df3_asis = datos["escenario3"]["ASIS"]
    df3_tobe = datos["escenario3"]["TOBE"]
    fig, axes = plt.subplots(1, 2, figsize=(11, 5))
    # Crea una figura con 1 fila y 2 columnas de gráficas (una para AS-IS, otra para TO-BE),
    # devueltas en la lista "axes".
    for ax, df, nombre, color in zip(axes, [df3_asis, df3_tobe], ["AS-IS", "TO-BE"], ["#c0392b", "#27ae60"]):
        # Recorre a la vez: cada uno de los dos "ejes" (subgráficas), su tabla de datos
        # correspondiente, su nombre y su color.
        if len(df) > 0:
            ax.scatter(df["t_min"] / 60, df["tiempo_respuesta_min"], alpha=0.35, s=14, color=color)
            # Dibuja un gráfico de puntos (scatter): en el eje X, la hora dentro de la ventana
            # del brote (se divide entre 60 para pasar de minutos a horas); en el eje Y, el
            # tiempo de respuesta de cada caída individual. alpha=0.35 hace los puntos semitransparentes
            # (para ver mejor si se solapan muchos), s=14 es el tamaño de cada punto.
        ax.set_xlabel("Hora dentro de la ventana de 4h del brote")
        ax.set_ylabel("Tiempo de respuesta (min)")
        ax.set_title(f"Escenario 3 - {nombre}")
        ax.set_xlim(0, 4)
        # Fija el eje X entre 0 y 4 horas (la duración del brote), para que ambas subgráficas
        # se puedan comparar visualmente con la misma escala.
    fig.suptitle("Escenario 3 (pico de incidencias nocturnas): tiempo de respuesta según el momento del brote")
    # Título general para toda la figura (por encima de las dos subgráficas).
    fig.tight_layout()
    fig.savefig(carpeta / "fig5_escenario3_detalle.png", dpi=200)
    plt.close(fig)

    return [carpeta / f for f in ["fig1_boxplot_general.png", "fig2_dia_vs_noche.png",
                                    "fig3_pct_traslados.png", "fig4_comparativa_escenarios.png",
                                    "fig5_escenario3_detalle.png"]]
    # Devuelve la lista de rutas completas de las 5 imágenes generadas, para que se puedan
    # insertar después en el Excel de resultados.


# ==========================================================================
# 8. ESCRITURA DEL EXCEL DE RESULTADOS
# ==========================================================================

def _tabla_resumen(stats_por_escenario, orden):
    # El guion bajo delante del nombre ("_tabla_resumen") es una convención en Python para
    # indicar "esto es una función auxiliar interna, no pensada para usarse desde fuera".
    filas = []
    for nombre in orden:
        # Recorre los 4 bloques en el orden indicado: Anual, Escenario 1, 2 y 3.
        for esc in ["ASIS", "TOBE"]:
            # Para cada bloque, recorre los dos escenarios.
            s = stats_por_escenario[nombre][esc]
            # Recupera el diccionario de estadísticos correspondiente.
            filas.append({
                "Escenario": nombre, "Modelo": "AS-IS" if esc == "ASIS" else "TO-BE",
                "Nº incidentes (medio)": round(s["n_incidentes_medio"], 1),
                # round(valor, decimales) redondea el número a la cantidad de decimales indicada.
                "T. respuesta medio (min)": round(s["tiempo_respuesta_medio"], 2) if pd.notna(s["tiempo_respuesta_medio"]) else None,
                # pd.notna(...) comprueba que el valor NO sea vacío/NaN antes de redondearlo;
                # si es vacío, se deja como "None" (vacío) en la tabla final, en vez de dar error.
                "IC95% inf": round(s["tr_ic95_lo"], 2) if pd.notna(s["tr_ic95_lo"]) else None,
                "IC95% sup": round(s["tr_ic95_hi"], 2) if pd.notna(s["tr_ic95_hi"]) else None,
                "Mediana (min)": round(s["tiempo_respuesta_mediana"], 2) if pd.notna(s["tiempo_respuesta_mediana"]) else None,
                "P95 (min)": round(s["tiempo_respuesta_p95"], 2) if pd.notna(s["tiempo_respuesta_p95"]) else None,
                "Máximo (min)": round(s["tiempo_respuesta_max"], 2) if pd.notna(s["tiempo_respuesta_max"]) else None,
                "% traslados": round(s["pct_traslados"], 1) if pd.notna(s["pct_traslados"]) else None,
            })
            # Cada una de estas líneas construye una columna de la tabla final, redondeando
            # y protegiéndose de valores vacíos.
    return pd.DataFrame(filas)
    # Convierte la lista de diccionarios (una fila por escenario+modelo) en una tabla de pandas.


def escribir_excel_resultados(datos, imagenes, path_salida):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.drawing.image import Image as XLImage
    # Importaciones "locales" (dentro de la función): herramientas de openpyxl para dar
    # estilo a las celdas (Font=tipo de letra, PatternFill=color de fondo) y para insertar imágenes.

    orden = ["Anual", "Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]
    tabla = _tabla_resumen(datos["stats"], orden)
    # Construye la tabla resumen general con los 4 bloques en este orden fijo.

    with pd.ExcelWriter(path_salida, engine="openpyxl") as writer:
        # Abre un "escritor" de Excel: permite ir añadiendo varias hojas al mismo archivo
        # antes de guardarlo. El "with" se asegura de que el archivo se cierre correctamente
        # al terminar, incluso si algo falla a mitad de camino.
        tabla.to_excel(writer, sheet_name="Resumen_General", index=False)
        # Escribe la tabla resumen en una hoja llamada "Resumen_General".
        # index=False evita que se añada una columna extra con el número de fila de pandas.

        for nombre, hoja in [("Escenario 1 - Día tipo", "Escenario1_Dia"),
                              ("Escenario 2 - Noche tipo", "Escenario2_Noche"),
                              ("Escenario 3 - Pico nocturno", "Escenario3_PicoNocturno")]:
            # Recorre una lista de parejas (nombre completo, nombre corto de la hoja).
            sub = tabla[tabla["Escenario"] == nombre].reset_index(drop=True)
            # Filtra de la tabla general solo las filas de ese escenario concreto.
            sub.to_excel(writer, sheet_name=hoja, index=False, startrow=1)
            # Escribe esa sub-tabla en su propia hoja, dejando la fila 1 en blanco (startrow=1).

        # comparativa: reducción relativa por escenario, para ver cuál mejora más / cuál es más crítico en As-Is
        comp_filas = []
        for nombre in orden:
            s_asis = datos["stats"][nombre]["ASIS"]
            s_tobe = datos["stats"][nombre]["TOBE"]
            if pd.notna(s_asis["tiempo_respuesta_medio"]) and s_asis["tiempo_respuesta_medio"] > 0:
                reduccion = 100 * (s_asis["tiempo_respuesta_medio"] - s_tobe["tiempo_respuesta_medio"]) / s_asis["tiempo_respuesta_medio"]
                # Calcula el % de reducción del tiempo de respuesta al pasar de AS-IS a TO-BE:
                # (AS-IS - TO-BE) / AS-IS × 100. Si TO-BE es más rápido, el resultado es positivo.
            else:
                reduccion = None
                # Si no hay datos válidos (o AS-IS fuera 0), se deja en blanco para evitar
                # una división por cero o un resultado sin sentido.
            comp_filas.append({
                "Escenario": nombre,
                "T. respuesta AS-IS (min)": round(s_asis["tiempo_respuesta_medio"], 2) if pd.notna(s_asis["tiempo_respuesta_medio"]) else None,
                "T. respuesta TO-BE (min)": round(s_tobe["tiempo_respuesta_medio"], 2) if pd.notna(s_tobe["tiempo_respuesta_medio"]) else None,
                "Reducción (%)": round(reduccion, 1) if reduccion is not None else None,
                "P95 AS-IS (min)": round(s_asis["tiempo_respuesta_p95"], 2) if pd.notna(s_asis["tiempo_respuesta_p95"]) else None,
                "P95 TO-BE (min)": round(s_tobe["tiempo_respuesta_p95"], 2) if pd.notna(s_tobe["tiempo_respuesta_p95"]) else None,
            })
        comp = pd.DataFrame(comp_filas)
        # Convierte la lista de comparativas en una tabla.
        comp["Ranking severidad AS-IS"] = comp["T. respuesta AS-IS (min)"].rank(ascending=False).astype(int)
        # .rank(ascending=False) asigna un puesto (1º, 2º, 3º...) a cada escenario según su
        # tiempo de respuesta AS-IS, de mayor a menor (el más grave/lento es el nº 1).
        # .astype(int) convierte ese ranking a números enteros.
        comp.to_excel(writer, sheet_name="Comparativa_Escenarios", index=False)
        # Escribe esta tabla comparativa en su propia hoja.

        # muestras de datos crudos
        datos["anual"]["ASIS"].sample(min(3000, len(datos["anual"]["ASIS"])), random_state=1).to_excel(
            writer, sheet_name="Datos_Crudos_Anual_ASIS", index=False)
        # .sample(n, random_state=1) coge una muestra aleatoria de n filas de la tabla completa
        # (para no volcar potencialmente decenas de miles de filas en el Excel, lo que lo haría
        # enorme y lento de abrir). min(3000, len(...)) asegura no pedir más filas de las que hay.
        # random_state=1 fija la semilla de este muestreo, para que sea reproducible también.
        datos["anual"]["TOBE"].sample(min(3000, len(datos["anual"]["TOBE"])), random_state=1).to_excel(
            writer, sheet_name="Datos_Crudos_Anual_TOBE", index=False)
        # Igual, para el escenario TO-BE.
        datos["escenario3"]["ASIS"].to_excel(writer, sheet_name="Datos_Crudos_Escenario3_ASIS", index=False)
        datos["escenario3"]["TOBE"].to_excel(writer, sheet_name="Datos_Crudos_Escenario3_TOBE", index=False)
        # El Escenario 3 tiene muchas menos filas (es un brote corto), así que se vuelca entero,
        # sin necesidad de muestrear.

    # segunda pasada: formato + hoja de gráficas
    wb = openpyxl.load_workbook(path_salida)
    # Reabre el Excel que se acaba de guardar, para poder darle formato visual
    # (esto no se puede hacer directamente con pd.ExcelWriter en el mismo paso).
    cab_fill = PatternFill("solid", fgColor="1F4E78")
    # Define un color de fondo sólido (azul oscuro, código hexadecimal 1F4E78) para las cabeceras.
    cab_font = Font(name="Arial", color="FFFFFF", bold=True)
    # Define un estilo de letra: Arial, color blanco (FFFFFF), en negrita.
    for hoja in wb.sheetnames:
        # Recorre TODAS las hojas del Excel, una por una.
        ws = wb[hoja]
        for cell in ws[1]:
            # Recorre todas las celdas de la primera fila (la cabecera) de esa hoja.
            if cell.value is not None:
                cell.font = cab_font
                cell.fill = cab_fill
                # Si la celda tiene contenido, se le aplica el estilo de letra blanca en negrita
                # y el fondo azul definidos arriba.
        for col in ws.columns:
            # Recorre todas las columnas de la hoja.
            longitud = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            # Calcula la longitud del texto más largo de esa columna (para ajustar el ancho).
            # default=10 evita un error si la columna está completamente vacía.
            ws.column_dimensions[col[0].column_letter].width = min(max(longitud + 2, 10), 34)
            # Ajusta el ancho de la columna al contenido, con un mínimo de 10 y un máximo de 34
            # (para que ninguna columna quede ni demasiado estrecha ni excesivamente ancha).

    ws_g = wb.create_sheet("Graficas")
    # Crea una nueva hoja llamada "Graficas" al final del Excel.
    fila_img = 1
    # Variable que lleva la cuenta de en qué fila insertar la siguiente imagen.
    for img_path in imagenes:
        # Recorre la lista de rutas de imágenes generadas en el Bloque 7.
        img = XLImage(str(img_path))
        # Carga la imagen para poder insertarla en el Excel.
        img.width, img.height = img.width * 0.55, img.height * 0.55
        # Reduce el tamaño de la imagen al 55% de su tamaño original, para que quepa mejor en la hoja.
        ws_g.add_image(img, f"A{fila_img}")
        # Inserta la imagen en la hoja "Graficas", empezando en la celda "A" + el número de fila actual.
        fila_img += int(img.height / 15) + 2
        # Calcula cuánto "espacio" (en filas) ocupa la imagen insertada, y desplaza el punto
        # de inserción para la siguiente imagen, dejando además un pequeño margen (+2).

    wb.save(path_salida)
    # Guarda el Excel definitivo, ya con formato y con la hoja de gráficas incluida.


# ==========================================================================
# 9. PROGRAMA PRINCIPAL
# ==========================================================================

def main(ruta_plantilla="plantilla_datos_centro.xlsx", carpeta_salida="salidas_simulacion"):
    # Función principal: la que organiza y llama a todo lo demás en el orden correcto.
    # Tiene valores por defecto para sus dos parámetros, por si no se le indican otros.
    carpeta = Path(carpeta_salida)
    carpeta.mkdir(exist_ok=True)
    # Crea la carpeta de salida si no existe todavía. exist_ok=True evita un error
    # si la carpeta ya existía de una ejecución anterior.

    print(f"Leyendo parámetros de: {ruta_plantilla}")
    # Imprime un mensaje por pantalla, para que el usuario vea el progreso del programa.
    general, zonas, deteccion = leer_plantilla(ruta_plantilla)
    # Llama al Bloque 0 para leer todos los parámetros del Excel de entrada.
    print(f"  Centro: {general['n_residentes']:.0f} residentes, {len(zonas)} zonas, "
          f"{general['n_replicas']} réplicas x {general['sim_dias']} días")
    # Imprime un resumen de la configuración leída. ":.0f" formatea el número sin decimales.

    print("Simulando año completo (Escenarios 1 y 2)...")
    df_asis_anual = ejecutar_anual(general, zonas, deteccion, "ASIS")
    df_tobe_anual = ejecutar_anual(general, zonas, deteccion, "TOBE")
    # Ejecuta la simulación anual completa para ambos escenarios (Bloque 5).

    df_asis_e1 = df_asis_anual[~df_asis_anual["es_noche"]].reset_index(drop=True)
    # De la tabla anual de AS-IS, se queda solo con las filas donde "es_noche" es False
    # (el símbolo "~" invierte True/False, es decir, "NO es de noche" = "es de día").
    # Esto es el Escenario 1 (Día tipo).
    df_asis_e2 = df_asis_anual[df_asis_anual["es_noche"]].reset_index(drop=True)
    # Aquí se queda solo con las filas de noche: Escenario 2 (Noche tipo).
    df_tobe_e1 = df_tobe_anual[~df_tobe_anual["es_noche"]].reset_index(drop=True)
    df_tobe_e2 = df_tobe_anual[df_tobe_anual["es_noche"]].reset_index(drop=True)
    # Lo mismo, separando día/noche, para el escenario TO-BE.
    # .reset_index(drop=True) renumera las filas desde 0 después de filtrar, sin dejar "huecos".

    print("Simulando Escenario 3 (pico de incidencias nocturnas)...")
    df_asis_e3 = ejecutar_escenario3(general, zonas, deteccion, "ASIS",
                                      n_incidencias_objetivo=general["n_incidencias_e3"])
    df_tobe_e3 = ejecutar_escenario3(general, zonas, deteccion, "TOBE",
                                      n_incidencias_objetivo=general["n_incidencias_e3"])
    # Ejecuta la simulación del brote nocturno de 4 horas para ambos escenarios,
    # con el nº de incidencias objetivo leído de la hoja 'Escenarios' del Excel.

    n_rep = general["n_replicas"]
    stats = {
        "Anual": {"ASIS": estadisticos_escenario(df_asis_anual, n_rep),
                  "TOBE": estadisticos_escenario(df_tobe_anual, n_rep)},
        "Escenario 1 - Día tipo": {"ASIS": estadisticos_escenario(df_asis_e1, n_rep),
                                     "TOBE": estadisticos_escenario(df_tobe_e1, n_rep)},
        "Escenario 2 - Noche tipo": {"ASIS": estadisticos_escenario(df_asis_e2, n_rep),
                                       "TOBE": estadisticos_escenario(df_tobe_e2, n_rep)},
        "Escenario 3 - Pico nocturno": {"ASIS": estadisticos_escenario(df_asis_e3, n_rep),
                                          "TOBE": estadisticos_escenario(df_tobe_e3, n_rep)},
    }
    # Construye un diccionario con los estadísticos (Bloque 6) de los 4 bloques de escenarios,
    # cada uno con su versión AS-IS y TO-BE.

    datos = {
        "anual": {"ASIS": df_asis_anual, "TOBE": df_tobe_anual},
        "escenario1": {"ASIS": df_asis_e1, "TOBE": df_tobe_e1},
        "escenario2": {"ASIS": df_asis_e2, "TOBE": df_tobe_e2},
        "escenario3": {"ASIS": df_asis_e3, "TOBE": df_tobe_e3},
        "stats": stats,
    }
    # Junta TODO en un único diccionario "datos", que se pasará a las funciones de
    # gráficas y de escritura del Excel (para no tener que pasar 10 parámetros sueltos).

    print("\n" + "=" * 78)
    # Imprime una línea de 78 signos "=" como separador visual, precedida de un salto de línea.
    for nombre in ["Anual", "Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]:
        sa, st = stats[nombre]["ASIS"], stats[nombre]["TOBE"]
        print(f"{nombre}:")
        print(f"   AS-IS  -> media {sa['tiempo_respuesta_medio']:.2f} min "
              f"(P95 {sa['tiempo_respuesta_p95']:.2f}, max {sa['tiempo_respuesta_max']:.2f}) | "
              f"traslados {sa['pct_traslados']:.1f}%")
        print(f"   TO-BE  -> media {st['tiempo_respuesta_medio']:.2f} min "
              f"(P95 {st['tiempo_respuesta_p95']:.2f}, max {st['tiempo_respuesta_max']:.2f}) | "
              f"traslados {st['pct_traslados']:.1f}%")
        # Imprime por pantalla, para cada bloque de escenario, un resumen legible de AS-IS y TO-BE
        # con la media, el percentil 95, el máximo y el % de traslados, todo con 2 (o 1) decimales.
    print("=" * 78)

    print("\nGenerando gráficas...")
    imagenes = generar_graficas(datos, carpeta)
    # Llama al Bloque 7 para crear las 5 gráficas y guardarlas como imágenes.

    print("Escribiendo Excel de resultados...")
    escribir_excel_resultados(datos, imagenes, carpeta / "resultados_simulacion.xlsx")
    # Llama al Bloque 8 para volcar todo (tablas + gráficas) al Excel final de resultados.

    print(f"\nListo. Todo guardado en: {carpeta.resolve()}")
    # Mensaje final. carpeta.resolve() muestra la ruta completa y absoluta de la carpeta de salida.


if __name__ == "__main__":
    # Esta condición es un estándar en Python: el código de dentro SOLO se ejecuta si este
    # archivo se lanza directamente (ej. "python simulacion_onacare_v4.py"), y NO se ejecuta
    # si este archivo se importa como un módulo desde otro script de Python.
    ruta = sys.argv[1] if len(sys.argv) > 1 else "plantilla_datos_centro.xlsx"
    # sys.argv es la lista de "argumentos" que se pasan al ejecutar el script desde la terminal.
    # sys.argv[0] es siempre el propio nombre del script, así que sys.argv[1] sería el primer
    # argumento adicional (ej. el nombre del Excel a usar). Si no se ha pasado ninguno,
    # se usa por defecto "plantilla_datos_centro.xlsx".
    main(ruta)
    # Llama a la función principal con la ruta del Excel determinada, arrancando así todo el programa.
