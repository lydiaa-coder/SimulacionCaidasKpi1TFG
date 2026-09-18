"""
SIMULACIÓN DE EVENTOS DISCRETOS (DES) v4 - Respuesta ante caídas
==================================================================

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

import sys
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

# ==========================================================================
# 0. LECTURA DE PARÁMETROS DESDE EL EXCEL
# ==========================================================================

def leer_plantilla(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # --- Parametros_Generales ---
    ws = wb["Parametros_Generales"]
    g = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        nombre, valor = row[0], row[1]
        if nombre is None:
            continue
        g[nombre] = valor

    general = {
        "n_residentes": float(g["Nº de residentes totales"]),
        "tasa_caidas_anual": float(g["Tasa de caídas"]),
        "velocidad_m_min": float(g["Velocidad de marcha del personal"]),
        "variabilidad": float(g["Variabilidad del desplazamiento (±)"]),
        "atencion_min": float(g["Tiempo de atención in situ (mínimo)"]),
        "atencion_max": float(g["Tiempo de atención in situ (máximo)"]),
        "p_traslado_asis": float(g["Probabilidad de traslado hospitalario - As-Is"]),
        "p_traslado_tobe": float(g["Probabilidad de traslado hospitalario - To-Be"]),
        "sim_dias": int(g["Nº de días a simular"]),
        "n_replicas": int(g["Nº de réplicas independientes"]),
        "seed": int(g["Semilla aleatoria"]),
        "p_ocupado_rutina": float(g.get("Probabilidad de ocupación en tarea no interrumpible", 0.20)),
    }

    # --- Zonas_Layout ---
    ws = wb["Zonas_Layout"]
    zonas = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        nombre = row[0]
        if nombre is None or nombre == "TOTAL":
            break   # distancia intra-zona (moverse dentro de la misma zona)
        intra_asis = float(row[6]) if len(row) > 6 and row[6] is not None else float(row[2]) / 3.0
        intra_tobe = float(row[7]) if len(row) > 7 and row[7] is not None else float(row[3]) / 3.0
        zonas.append({
            "nombre": nombre,
            "n_residentes": float(row[1]),
            "dist_asis": float(row[2]),
            "dist_tobe": float(row[3]),
            "personal_dia": int(row[4]),
            "personal_noche": int(row[5]),
            "intra_asis": intra_asis,
            "intra_tobe": intra_tobe,
        })

    # --- Deteccion ---
    ws = wb["Deteccion"]
    d = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        nombre, v_asis, v_tobe = row[0], row[1], row[2]
        if nombre is None:
            continue
        d[nombre] = (v_asis, v_tobe)

    # Detección acústica incidental (oír un grito/golpe sin ver la caída)
    # si no existen en la plantilla, se usa 0.0 (nadie detecta por esta vía) 
    # y el modelo se comporta exactamente como antes.
    p_oida_dia = float(d.get("Probabilidad de detección acústica (oída) - turno día", (0.0, 0.0))[0] or 0.0)
    p_oida_noche = float(d.get("Probabilidad de detección acústica (oída) - turno noche", (0.0, 0.0))[0] or 0.0)
    oida_min = float(d.get("Detección acústica (oída) - mínimo (min)", (1.0, 1.0))[0] or 1.0)
    oida_max = float(d.get("Detección acústica (oída) - máximo (min)", (5.0, 5.0))[0] or 5.0)

    deteccion = {
        "ASIS": {
            "modo": d["Modo de detección"][0],
            "p_no_presenciada": float(d["Probabilidad de caída no presenciada"][0]),
            "ronda_dia": float(d["Intervalo de ronda - turno día (min)"][0]),
            "ronda_noche": float(d["Intervalo de ronda - turno noche (min)"][0]),
            "boton_min": float(d["Detección con botón fijo - mínimo (min)"][0]),
            "boton_max": float(d["Detección con botón fijo - máximo (min)"][0]),
            "p_oida_dia": p_oida_dia,
            "p_oida_noche": p_oida_noche,
            "oida_min": oida_min,
            "oida_max": oida_max,
        },
        "TOBE": {
            "modo": d["Modo de detección"][1],
            "p_no_presenciada": float(d["Probabilidad de caída no presenciada"][1])
                if not isinstance(d["Probabilidad de caída no presenciada"][1], str) else None,
            "sensor_min_seg": float(d["Detección sensor ambiental - mínimo (seg)"][1]),
            "sensor_max_seg": float((d.get("Detección sensor ambiental - máximo (seg)")
                                     or d["Detección sensor ambiental -máximo (seg)"])[1]),
            "p_activar_manual": d["Prob. de poder activar dispositivo manual"][1],
            # fallback si modo='dispositivo_manual' y no se puede activar: usa las mismas
            # rondas manuales definidas para As-Is
            "ronda_dia": float(d["Intervalo de ronda - turno día (min)"][0]),
            "ronda_noche": float(d["Intervalo de ronda - turno noche (min)"][0]),
            # Probabilidad de que el sensor NO detecte la caída (falso negativo).
            # Fallback en ese caso: misma vía de detección acústica incidental
            # que en As-Is (el entorno acústico del edificio no cambia por
            # instalar sensores), y en último caso la ronda manual.
            "p_fallo_sensor": float(d.get("Probabilidad de fallo del sensor (falso negativo)", (0.0, 0.0))[1] or 0.0),
            "p_oida_dia": p_oida_dia,
            "p_oida_noche": p_oida_noche,
            "oida_min": oida_min,
            "oida_max": oida_max,
        },
    }

    # --- Escenarios ---
    # Lee el nº de incidencias objetivo del Escenario 3 (pico nocturno) de la
    # hoja 'Escenarios'. Si la hoja no existe o la celda está vacía, se usa 8
    # como valor de reserva.
    n_incidencias_e3 = 8
    if "Escenarios" in wb.sheetnames:
        ws = wb["Escenarios"]
        for row in ws.iter_rows(min_row=3, values_only=True):
            nombre = row[0]
            if nombre and str(nombre).startswith("Escenario 3"):
                if len(row) > 4 and row[4] not in (None, "-"):
                    n_incidencias_e3 = int(row[4])
                break
    general["n_incidencias_e3"] = n_incidencias_e3

    return general, zonas, deteccion


# ==========================================================================
# 1. MODELO DE ZONAS Y DISTANCIAS
# ==========================================================================

def construir_modelo_zonas(zonas, escenario, turno):
    """
    Devuelve:
      - probs: probabilidad de que la caída ocurra en cada zona (prop. a nº residentes)
      - home_zone_staff: array con la zona base de cada miembro del personal de ese turno
      - dist_matrix: matriz de distancias entre zonas (incluye intra-zona, propia de cada zona - NUEVO)
    """
    n_res = np.array([z["n_residentes"] for z in zonas], dtype=float)
    probs = n_res / n_res.sum()

    dist_col = "dist_asis" if escenario == "ASIS" else "dist_tobe"
    dist_control = np.array([z[dist_col] for z in zonas], dtype=float)

    # --- NUEVO: distancia intra-zona propia de cada zona (ya no una constante única) ---
    intra_col = "intra_asis" if escenario == "ASIS" else "intra_tobe"
    dist_intra = np.array([z[intra_col] for z in zonas], dtype=float)
    # --- FIN NUEVO ---

    n_zonas = len(zonas)
    dist_matrix = np.zeros((n_zonas, n_zonas))
    for i in range(n_zonas):
        for j in range(n_zonas):
            if i == j:
                dist_matrix[i, j] = dist_intra[i]
            else:
                dist_matrix[i, j] = dist_control[i] + dist_control[j]

    personal_col = "personal_dia" if turno == "DIA" else "personal_noche"
    home_zone_staff = []
    for zi, z in enumerate(zonas):
        home_zone_staff += [zi] * z[personal_col]
    home_zone_staff = np.array(home_zone_staff)

    return probs, home_zone_staff, dist_matrix


# ==========================================================================
# 2. DETECCIÓN
# ==========================================================================

def tiempo_deteccion(escenario, det_params, es_noche, rng):
    modo = det_params["modo"]
    ronda = det_params["ronda_noche"] if es_noche else det_params["ronda_dia"]

    if modo == "sensor_ambiental_auto":
        # Ningún sensor tiene una precisión del 100%
        p_fallo = det_params.get("p_fallo_sensor", 0.0)
        if rng.random() < p_fallo:
            p_oida = det_params.get("p_oida_noche", 0.0) if es_noche else det_params.get("p_oida_dia", 0.0)
            if rng.random() < p_oida:
                return rng.uniform(det_params.get("oida_min", 1.0), det_params.get("oida_max", 5.0))
            return rng.uniform(0, ronda)
        return rng.uniform(det_params["sensor_min_seg"], det_params["sensor_max_seg"]) / 60.0

    if modo == "dispositivo_manual":
        p_activar = det_params.get("p_activar_manual") or 0.5
        if rng.random() < p_activar:
            return rng.uniform(det_params["sensor_min_seg"], det_params["sensor_max_seg"]) / 60.0
        else:
            return rng.uniform(0, ronda)

    # modo por defecto: ronda_manual_o_boton
    p_no_presenciada = det_params.get("p_no_presenciada", 0.6)
    if rng.random() < p_no_presenciada:
        # Dentro de las caídas "no presenciadas" (nadie la ve en el momento),
        # una parte se detecta igualmente pronto porque alguien oye un golpe
        # o un grito de ayuda sin haber visto la caída en sí 
        p_oida = det_params.get("p_oida_noche", 0.0) if es_noche else det_params.get("p_oida_dia", 0.0)
        if rng.random() < p_oida:
            return rng.uniform(det_params.get("oida_min", 1.0), det_params.get("oida_max", 5.0))
        return rng.uniform(0, ronda)
    else:
        return rng.uniform(det_params["boton_min"], det_params["boton_max"])


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
    filas = []

    # pre-construir modelo de zonas para turno día y turno noche (distinto personal)
    modelos = {
        "DIA": construir_modelo_zonas(zonas, escenario, "DIA"),
        "NOCHE": construir_modelo_zonas(zonas, escenario, "NOCHE"),
    }
    free_time = {
        "DIA": np.zeros(len(modelos["DIA"][1])),
        "NOCHE": np.zeros(len(modelos["NOCHE"][1])),
    }

    for t, es_noche in zip(t_llegadas, es_noche_flags):
        turno = "NOCHE" if es_noche else "DIA"
        probs, home_zone_staff, dist_matrix = modelos[turno]
        ft = free_time[turno]

        zona_incidente = rng.choice(len(zonas), p=probs)

        deteccion_min = tiempo_deteccion(escenario, det_params, es_noche, rng)
        t_listo = t + deteccion_min

        # Personal "libre" según el histórico de caídas atendidas
        # (no está respondiendo a otra caída en este momento)
        libres_todos = np.where(ft <= t_listo)[0]
        """
        De ese personal libre, una parte puede estar en ese instante
        ocupada en una tarea asistencial no interrumpible con otro
        residente (aseo, alimentación, movilización con grúa,
        administración de medicación) y no puede acudir de inmediato;
        en ese caso debe ir otro compañero libre y disponible.
        """
        # [Supuesto - ver parámetro 'p_ocupado_rutina' y su justificación]
        if len(libres_todos) > 0:
            p_ocupado = general.get("p_ocupado_rutina", 0.0)
            disponible_mask = rng.random(len(libres_todos)) >= p_ocupado
            libres = libres_todos[disponible_mask]
        else:
            libres = libres_todos

        if len(libres) > 0:
            distancias_libres = dist_matrix[home_zone_staff[libres], zona_incidente]
            idx = libres[np.argmin(distancias_libres)]
            espera = 0.0
        elif len(libres_todos) > 0:
            # Todo el personal libre según el histórico de caídas resultó
            # (por puro azar del sorteo) estar ocupado en tarea rutinaria.
            # Como no se modela cuánto dura esa tarea, se opta por el más
            # cercano de ellos en vez de esperar a alguien que lleva más
            # tiempo inactivo (evita esperas negativas o irreales).
            distancias_libres = dist_matrix[home_zone_staff[libres_todos], zona_incidente]
            idx = libres_todos[np.argmin(distancias_libres)]
            espera = 0.0
        else:
            # Nadie realmente disponible: ni por estar ya respondiendo a
            # otra caída, ni por estar en una tarea rutinaria no
            # interrumpible. Se espera a que quede libre el primero según
            # el histórico de caídas (no se modela por separado cuándo
            # termina la tarea rutinaria).
            idx = int(np.argmin(ft))
            espera = ft[idx] - t_listo

        dist = dist_matrix[home_zone_staff[idx], zona_incidente]
        ruido = 1 + rng.uniform(-general["variabilidad"], general["variabilidad"])
        desplazamiento = (dist / general["velocidad_m_min"]) * ruido

        atencion = rng.uniform(general["atencion_min"], general["atencion_max"])

        tiempo_respuesta = deteccion_min + espera + desplazamiento
        ft[idx] = t_listo + espera + desplazamiento + atencion

        p_trasl = general["p_traslado_asis"] if escenario == "ASIS" else general["p_traslado_tobe"]
        traslado = rng.random() < p_trasl

        filas.append({
            "t_min": t, "es_noche": bool(es_noche), "zona": zonas[zona_incidente]["nombre"],
            "deteccion_min": deteccion_min, "espera_min": espera,
            "desplazamiento_min": desplazamiento, "tiempo_respuesta_min": tiempo_respuesta,
            "traslado_hospital": traslado, "escenario": escenario,
        })

    return filas


# ==========================================================================
# 4. GENERADORES DE LLEGADAS - ESCENARIOS 1/2 (año completo) Y 3 (pico 4h)
# ==========================================================================

# Ventanas horarias de los escenarios (minutos desde medianoche).
HORA_INICIO_NOCHE = 20 * 60
HORA_FIN_NOCHE = 8 * 60

# Reparto real día/noche de las caídas [FUENTE: Samper Lamenca et al.,
# Gerokomos 2016 - 48,1% turno mañana + 35% turno tarde = 83,1% día;
# 16,9% turno noche; n=160 caídas, residencia española real]
FRACCION_CAIDAS_DIA = 0.831
FRACCION_CAIDAS_NOCHE = 0.169


def generar_llegadas_anual(general, rng):
    """
    Genera las caídas de un año simulado como dos procesos de Poisson
    independientes - turno día (08-20h) y turno noche (20-08h) - calibrados
    con el reparto real día/noche documentado en Samper Lamenca et al.
    (Gerokomos, 2016; n=160 caídas, residencia española real): 48,1% turno
    mañana + 35% turno tarde = 83,1% turno día; 16,9% turno noche.
    """
    lam_anual = general["tasa_caidas_anual"] * general["n_residentes"]
    sim_dias = general["sim_dias"]

    lam_dia_periodo = lam_anual * FRACCION_CAIDAS_DIA * sim_dias / 365
    lam_noche_periodo = lam_anual * FRACCION_CAIDAS_NOCHE * sim_dias / 365

    n_dia = rng.poisson(lam_dia_periodo)
    n_noche = rng.poisson(lam_noche_periodo)

    # Turno día: cada caída, un día aleatorio de los simulados y un
    # instante uniforme dentro de la ventana 08:00-20:00 de ese día.
    dias_dia = rng.integers(0, sim_dias, n_dia)
    minutos_dia = rng.uniform(HORA_FIN_NOCHE, HORA_INICIO_NOCHE, n_dia)
    llegadas_dia = dias_dia * 1440 + minutos_dia

    # Turno noche: cada caída, una noche aleatoria y un instante
    # uniforme dentro de esa ventana de horas.
    dias_noche = rng.integers(0, sim_dias, n_noche)
    minutos_noche = rng.uniform(0, 720, n_noche)
    llegadas_noche = dias_noche * 1440 + HORA_INICIO_NOCHE + minutos_noche

    llegadas = np.concatenate([llegadas_dia, llegadas_noche])
    es_noche = np.concatenate([np.zeros(n_dia, dtype=bool), np.ones(n_noche, dtype=bool)])

    horizonte = sim_dias * 24 * 60
    mask = llegadas < horizonte
    orden = np.argsort(llegadas[mask])
    return llegadas[mask][orden], es_noche[mask][orden]


def generar_llegadas_escenario3(n_incidencias_objetivo, ventana_min, rng):
    """Genera de forma independiente, réplica a réplica, un nº de incidencias
    de Poisson con media n_incidencias_objetivo, distribuidas uniformemente
    dentro de la ventana de 'ventana_min' minutos (p.ej. 4 horas de brote)."""
    n = rng.poisson(n_incidencias_objetivo)
    n = max(n, 0)
    llegadas = np.sort(rng.uniform(0, ventana_min, n))
    es_noche = np.ones(n, dtype=bool)  # todo el escenario 3 ocurre en turno de noche
    return llegadas, es_noche


# ==========================================================================
# 5. ORQUESTACIÓN: EJECUTAR RÉPLICAS PARA CADA ESCENARIO
# ==========================================================================

def ejecutar_anual(general, zonas, deteccion, escenario):
    resultados = []
    for r in range(general["n_replicas"]):
        rng = np.random.default_rng(general["seed"] + r)
        t_llegadas, es_noche = generar_llegadas_anual(general, rng)
        filas = resolver_incidentes(t_llegadas, es_noche, escenario, general, zonas, deteccion, rng)
        for f in filas:
            f["replica"] = r
        resultados.extend(filas)
    return pd.DataFrame(resultados)


def ejecutar_escenario3(general, zonas, deteccion, escenario, n_incidencias_objetivo=3,
                         ventana_min=240, n_replicas=None):
    n_replicas = n_replicas or general["n_replicas"]
    # offset de semilla distinto para no reutilizar la misma secuencia que el escenario anual
    resultados = []
    for r in range(n_replicas):
        rng = np.random.default_rng(general["seed"] + 100_000 + r)
        t_llegadas, es_noche = generar_llegadas_escenario3(n_incidencias_objetivo, ventana_min, rng)
        if len(t_llegadas) == 0:
            continue
        filas = resolver_incidentes(t_llegadas, es_noche, escenario, general, zonas, deteccion, rng)
        for f in filas:
            f["replica"] = r
        resultados.extend(filas)
    return pd.DataFrame(resultados)


# ==========================================================================
# 6. AGREGACIÓN Y ESTADÍSTICOS
# ==========================================================================

def resumen_por_replica(df, n_replicas):
    if len(df) == 0:
        idx = pd.Index(range(n_replicas), name="replica")
        return pd.DataFrame({"n_incidentes": 0, "tiempo_respuesta_medio": np.nan,
                              "pct_traslados": np.nan}, index=idx).reset_index()
    g = df.groupby("replica").agg(
        n_incidentes=("t_min", "size"),
        tiempo_respuesta_medio=("tiempo_respuesta_min", "mean"),
        pct_traslados=("traslado_hospital", "mean"),
    ).reindex(range(n_replicas), fill_value=0).reset_index().rename(columns={"index": "replica"})
    return g


def ic_95(serie):
    serie = serie.dropna()
    if len(serie) == 0:
        return np.nan, np.nan, np.nan
    m = serie.mean()
    err = 1.96 * serie.std(ddof=1) / np.sqrt(len(serie)) if len(serie) > 1 else 0.0
    return m, m - err, m + err


def estadisticos_escenario(df, n_replicas):
    """Devuelve un dict con las métricas clave de un (escenario, ASIS/TOBE)."""
    resumen = resumen_por_replica(df, n_replicas)
    m, lo, hi = ic_95(resumen["tiempo_respuesta_medio"])
    mt, lot, hit = ic_95(resumen["pct_traslados"])
    return {
        "n_incidentes_medio": resumen["n_incidentes"].mean() if len(df) else 0,
        "tiempo_respuesta_medio": m, "tr_ic95_lo": lo, "tr_ic95_hi": hi,
        "tiempo_respuesta_mediana": df["tiempo_respuesta_min"].median() if len(df) else np.nan,
        "tiempo_respuesta_p95": df["tiempo_respuesta_min"].quantile(0.95) if len(df) else np.nan,
        "tiempo_respuesta_max": df["tiempo_respuesta_min"].max() if len(df) else np.nan,
        "pct_traslados": 100 * mt if not np.isnan(mt) else np.nan,
        "pct_traslados_ic95_lo": 100 * lot if not np.isnan(lot) else np.nan,
        "pct_traslados_ic95_hi": 100 * hit if not np.isnan(hit) else np.nan,
    }


# ==========================================================================
# 7. GRÁFICAS
# ==========================================================================

def generar_graficas(datos, carpeta):
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    carpeta = Path(carpeta)

    df_asis_anual = datos["anual"]["ASIS"]
    df_tobe_anual = datos["anual"]["TOBE"]

    # --- Figura 1: boxplot general (todo el año) ---
    fig, ax = plt.subplots(figsize=(7, 5))
    caja = [df_asis_anual["tiempo_respuesta_min"], df_tobe_anual["tiempo_respuesta_min"]]
    try:
        ax.boxplot(caja, tick_labels=["AS-IS", "TO-BE"], showfliers=False)
    except TypeError:
        ax.boxplot(caja, labels=["AS-IS", "TO-BE"], showfliers=False)
    ax.set_ylabel("Tiempo de respuesta ante caída (min)")
    ax.set_title("Distribución del tiempo de respuesta - año completo\n(simulación DES por zonas)")
    fig.tight_layout()
    fig.savefig(carpeta / "fig1_boxplot_general.png", dpi=200)
    plt.close(fig)

    # --- Figura 2: día (Escenario 1) vs noche (Escenario 2) ---
    fig, ax = plt.subplots(figsize=(7, 5))
    x = np.arange(2)
    width = 0.35
    med_asis = [datos["stats"]["Escenario 1 - Día tipo"]["ASIS"]["tiempo_respuesta_medio"],
                datos["stats"]["Escenario 2 - Noche tipo"]["ASIS"]["tiempo_respuesta_medio"]]
    med_tobe = [datos["stats"]["Escenario 1 - Día tipo"]["TOBE"]["tiempo_respuesta_medio"],
                datos["stats"]["Escenario 2 - Noche tipo"]["TOBE"]["tiempo_respuesta_medio"]]
    ax.bar(x - width / 2, med_asis, width, label="AS-IS")
    ax.bar(x + width / 2, med_tobe, width, label="TO-BE")
    ax.set_xticks(x)
    ax.set_xticklabels(["Escenario 1\nDía tipo", "Escenario 2\nNoche tipo"])
    ax.set_ylabel("Tiempo de respuesta medio (min)")
    ax.set_title("Escenario 1 (día) vs Escenario 2 (noche)")
    ax.legend()
    fig.tight_layout()
    fig.savefig(carpeta / "fig2_dia_vs_noche.png", dpi=200)
    plt.close(fig)

    # --- Figura 3: % traslados anual ---
    fig, ax = plt.subplots(figsize=(6, 5))
    valores = [datos["stats"]["Anual"]["ASIS"]["pct_traslados"],
               datos["stats"]["Anual"]["TOBE"]["pct_traslados"]]
    ax.bar(["AS-IS", "TO-BE"], valores, color=["#c0392b", "#27ae60"])
    ax.set_ylabel("% de caídas con traslado a hospital")
    ax.set_title("Tasa de traslados hospitalarios por caída (anual)")
    for i, v in enumerate(valores):
        ax.text(i, v + 0.5, f"{v:.1f}%", ha="center")
    fig.tight_layout()
    fig.savefig(carpeta / "fig3_pct_traslados.png", dpi=200)
    plt.close(fig)

    # --- Figura 4: comparativa de los 3 escenarios (cuál es más perjudicial) ---
    nombres_esc = ["Escenario 1\nDía tipo", "Escenario 2\nNoche tipo", "Escenario 3\nPico nocturno"]
    claves_esc = ["Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]
    med_asis = [datos["stats"][k]["ASIS"]["tiempo_respuesta_medio"] for k in claves_esc]
    med_tobe = [datos["stats"][k]["TOBE"]["tiempo_respuesta_medio"] for k in claves_esc]

    fig, ax = plt.subplots(figsize=(8.5, 5.5))
    x = np.arange(3)
    width = 0.35
    b1 = ax.bar(x - width / 2, med_asis, width, label="AS-IS", color="#c0392b")
    b2 = ax.bar(x + width / 2, med_tobe, width, label="TO-BE", color="#27ae60")
    ax.set_xticks(x)
    ax.set_xticklabels(nombres_esc)
    ax.set_ylabel("Tiempo de respuesta medio (min)")
    ax.set_title("Comparativa de los 3 escenarios - ¿cuál es más perjudicial?")
    ax.legend()
    ax.bar_label(b1, fmt="%.1f", padding=3)
    ax.bar_label(b2, fmt="%.1f", padding=3)
    fig.tight_layout()
    fig.savefig(carpeta / "fig4_comparativa_escenarios.png", dpi=200)
    plt.close(fig)

    # --- Figura 5: detalle del Escenario 3 (pico nocturno) ---
    df3_asis = datos["escenario3"]["ASIS"]
    df3_tobe = datos["escenario3"]["TOBE"]
    fig, axes = plt.subplots(1, 2, figsize=(11, 5))
    for ax, df, nombre, color in zip(axes, [df3_asis, df3_tobe], ["AS-IS", "TO-BE"], ["#c0392b", "#27ae60"]):
        if len(df) > 0:
            ax.scatter(df["t_min"] / 60, df["tiempo_respuesta_min"], alpha=0.35, s=14, color=color)
        ax.set_xlabel("Hora dentro de la ventana de 4h del brote")
        ax.set_ylabel("Tiempo de respuesta (min)")
        ax.set_title(f"Escenario 3 - {nombre}")
        ax.set_xlim(0, 4)
    fig.suptitle("Escenario 3 (pico de incidencias nocturnas): tiempo de respuesta según el momento del brote")
    fig.tight_layout()
    fig.savefig(carpeta / "fig5_escenario3_detalle.png", dpi=200)
    plt.close(fig)

    return [carpeta / f for f in ["fig1_boxplot_general.png", "fig2_dia_vs_noche.png",
                                    "fig3_pct_traslados.png", "fig4_comparativa_escenarios.png",
                                    "fig5_escenario3_detalle.png"]]


# ==========================================================================
# 8. ESCRITURA DEL EXCEL DE RESULTADOS
# ==========================================================================

def _tabla_resumen(stats_por_escenario, orden):
    filas = []
    for nombre in orden:
        for esc in ["ASIS", "TOBE"]:
            s = stats_por_escenario[nombre][esc]
            filas.append({
                "Escenario": nombre, "Modelo": "AS-IS" if esc == "ASIS" else "TO-BE",
                "Nº incidentes (medio)": round(s["n_incidentes_medio"], 1),
                "T. respuesta medio (min)": round(s["tiempo_respuesta_medio"], 2) if pd.notna(s["tiempo_respuesta_medio"]) else None,
                "IC95% inf": round(s["tr_ic95_lo"], 2) if pd.notna(s["tr_ic95_lo"]) else None,
                "IC95% sup": round(s["tr_ic95_hi"], 2) if pd.notna(s["tr_ic95_hi"]) else None,
                "Mediana (min)": round(s["tiempo_respuesta_mediana"], 2) if pd.notna(s["tiempo_respuesta_mediana"]) else None,
                "P95 (min)": round(s["tiempo_respuesta_p95"], 2) if pd.notna(s["tiempo_respuesta_p95"]) else None,
                "Máximo (min)": round(s["tiempo_respuesta_max"], 2) if pd.notna(s["tiempo_respuesta_max"]) else None,
                "% traslados": round(s["pct_traslados"], 1) if pd.notna(s["pct_traslados"]) else None,
            })
    return pd.DataFrame(filas)


def escribir_excel_resultados(datos, imagenes, path_salida):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.drawing.image import Image as XLImage

    orden = ["Anual", "Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]
    tabla = _tabla_resumen(datos["stats"], orden)

    with pd.ExcelWriter(path_salida, engine="openpyxl") as writer:
        tabla.to_excel(writer, sheet_name="Resumen_General", index=False)

        for nombre, hoja in [("Escenario 1 - Día tipo", "Escenario1_Dia"),
                              ("Escenario 2 - Noche tipo", "Escenario2_Noche"),
                              ("Escenario 3 - Pico nocturno", "Escenario3_PicoNocturno")]:
            sub = tabla[tabla["Escenario"] == nombre].reset_index(drop=True)
            sub.to_excel(writer, sheet_name=hoja, index=False, startrow=1)

        # comparativa: reducción relativa por escenario, para ver cuál mejora más / cuál es más crítico en As-Is
        comp_filas = []
        for nombre in orden:
            s_asis = datos["stats"][nombre]["ASIS"]
            s_tobe = datos["stats"][nombre]["TOBE"]
            if pd.notna(s_asis["tiempo_respuesta_medio"]) and s_asis["tiempo_respuesta_medio"] > 0:
                reduccion = 100 * (s_asis["tiempo_respuesta_medio"] - s_tobe["tiempo_respuesta_medio"]) / s_asis["tiempo_respuesta_medio"]
            else:
                reduccion = None
            comp_filas.append({
                "Escenario": nombre,
                "T. respuesta AS-IS (min)": round(s_asis["tiempo_respuesta_medio"], 2) if pd.notna(s_asis["tiempo_respuesta_medio"]) else None,
                "T. respuesta TO-BE (min)": round(s_tobe["tiempo_respuesta_medio"], 2) if pd.notna(s_tobe["tiempo_respuesta_medio"]) else None,
                "Reducción (%)": round(reduccion, 1) if reduccion is not None else None,
                "P95 AS-IS (min)": round(s_asis["tiempo_respuesta_p95"], 2) if pd.notna(s_asis["tiempo_respuesta_p95"]) else None,
                "P95 TO-BE (min)": round(s_tobe["tiempo_respuesta_p95"], 2) if pd.notna(s_tobe["tiempo_respuesta_p95"]) else None,
            })
        comp = pd.DataFrame(comp_filas)
        comp["Ranking severidad AS-IS"] = comp["T. respuesta AS-IS (min)"].rank(ascending=False).astype(int)
        comp.to_excel(writer, sheet_name="Comparativa_Escenarios", index=False)

        # muestras de datos crudos
        datos["anual"]["ASIS"].sample(min(3000, len(datos["anual"]["ASIS"])), random_state=1).to_excel(
            writer, sheet_name="Datos_Crudos_Anual_ASIS", index=False)
        datos["anual"]["TOBE"].sample(min(3000, len(datos["anual"]["TOBE"])), random_state=1).to_excel(
            writer, sheet_name="Datos_Crudos_Anual_TOBE", index=False)
        datos["escenario3"]["ASIS"].to_excel(writer, sheet_name="Datos_Crudos_Escenario3_ASIS", index=False)
        datos["escenario3"]["TOBE"].to_excel(writer, sheet_name="Datos_Crudos_Escenario3_TOBE", index=False)

    # segunda pasada: formato + hoja de gráficas
    wb = openpyxl.load_workbook(path_salida)
    cab_fill = PatternFill("solid", fgColor="1F4E78")
    cab_font = Font(name="Arial", color="FFFFFF", bold=True)
    for hoja in wb.sheetnames:
        ws = wb[hoja]
        for cell in ws[1]:
            if cell.value is not None:
                cell.font = cab_font
                cell.fill = cab_fill
        for col in ws.columns:
            longitud = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max(longitud + 2, 10), 34)

    ws_g = wb.create_sheet("Graficas")
    fila_img = 1
    for img_path in imagenes:
        img = XLImage(str(img_path))
        img.width, img.height = img.width * 0.55, img.height * 0.55
        ws_g.add_image(img, f"A{fila_img}")
        fila_img += int(img.height / 15) + 2

    wb.save(path_salida)


# ==========================================================================
# 9. PROGRAMA PRINCIPAL
# ==========================================================================

def main(ruta_plantilla="plantilla_datos_centro.xlsx", carpeta_salida="salidas_simulacion"):
    carpeta = Path(carpeta_salida)
    carpeta.mkdir(exist_ok=True)

    print(f"Leyendo parámetros de: {ruta_plantilla}")
    general, zonas, deteccion = leer_plantilla(ruta_plantilla)
    print(f"  Centro: {general['n_residentes']:.0f} residentes, {len(zonas)} zonas, "
          f"{general['n_replicas']} réplicas x {general['sim_dias']} días")

    print("Simulando año completo (Escenarios 1 y 2)...")
    df_asis_anual = ejecutar_anual(general, zonas, deteccion, "ASIS")
    df_tobe_anual = ejecutar_anual(general, zonas, deteccion, "TOBE")

    df_asis_e1 = df_asis_anual[~df_asis_anual["es_noche"]].reset_index(drop=True)
    df_asis_e2 = df_asis_anual[df_asis_anual["es_noche"]].reset_index(drop=True)
    df_tobe_e1 = df_tobe_anual[~df_tobe_anual["es_noche"]].reset_index(drop=True)
    df_tobe_e2 = df_tobe_anual[df_tobe_anual["es_noche"]].reset_index(drop=True)

    print("Simulando Escenario 3 (pico de incidencias nocturnas)...")
    df_asis_e3 = ejecutar_escenario3(general, zonas, deteccion, "ASIS",
                                      n_incidencias_objetivo=general["n_incidencias_e3"])
    df_tobe_e3 = ejecutar_escenario3(general, zonas, deteccion, "TOBE",
                                      n_incidencias_objetivo=general["n_incidencias_e3"])

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

    datos = {
        "anual": {"ASIS": df_asis_anual, "TOBE": df_tobe_anual},
        "escenario1": {"ASIS": df_asis_e1, "TOBE": df_tobe_e1},
        "escenario2": {"ASIS": df_asis_e2, "TOBE": df_tobe_e2},
        "escenario3": {"ASIS": df_asis_e3, "TOBE": df_tobe_e3},
        "stats": stats,
    }

    print("\n" + "=" * 78)
    for nombre in ["Anual", "Escenario 1 - Día tipo", "Escenario 2 - Noche tipo", "Escenario 3 - Pico nocturno"]:
        sa, st = stats[nombre]["ASIS"], stats[nombre]["TOBE"]
        print(f"{nombre}:")
        print(f"   AS-IS  -> media {sa['tiempo_respuesta_medio']:.2f} min "
              f"(P95 {sa['tiempo_respuesta_p95']:.2f}, max {sa['tiempo_respuesta_max']:.2f}) | "
              f"traslados {sa['pct_traslados']:.1f}%")
        print(f"   TO-BE  -> media {st['tiempo_respuesta_medio']:.2f} min "
              f"(P95 {st['tiempo_respuesta_p95']:.2f}, max {st['tiempo_respuesta_max']:.2f}) | "
              f"traslados {st['pct_traslados']:.1f}%")
    print("=" * 78)

    print("\nGenerando gráficas...")
    imagenes = generar_graficas(datos, carpeta)

    print("Escribiendo Excel de resultados...")
    escribir_excel_resultados(datos, imagenes, carpeta / "resultados_simulacion.xlsx")

    print(f"\nListo. Todo guardado en: {carpeta.resolve()}")


if __name__ == "__main__":
    ruta = sys.argv[1] if len(sys.argv) > 1 else "plantilla_datos_centro.xlsx"
    main(ruta)
