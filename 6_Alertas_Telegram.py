"""
Script 6: Alertas de viento para quemas de caña – Sipacate.

Script INDEPENDIENTE: no modifica ni lee los CSV de Power BI. Descarga su propio
pronóstico de Open-Meteo (ECMWF IFS) para el área de intervención y envía un
mensaje por Telegram con:

  1. Ventana de quema (12:00–16:00): horas recomendables y no recomendables.
  2. Alertas de 06:00 a 17:00:
       - Viento medio > UMBRAL_VIENTO
       - Ráfagas      > UMBRAL_RAFAGA
       - Viento que va HACIA el sur, sureste o suroeste (riesgo de pavesa en salineras)
  3. Si el ICC emitió una alerta de frente frío (módulo de frentes), una línea que la recuerda.
     Este mensaje ya no calcula frentes por su cuenta, para no contradecir al módulo de frentes.

CONVENCIÓN DE DIRECCIÓN (importante):
  Open-Meteo entrega la dirección DE DONDE VIENE el viento (convención meteorológica).
    0°/360° = viene del norte → va HACIA el sur   → desfavorable (salineras)
    180°    = viene del sur   → va HACIA el norte → favorable
  En el mensaje siempre se indica hacia dónde VA el viento, que es lo que mueve la pavesa.

Variables de entorno (secrets de GitHub):
  TELEGRAM_TOKEN    token del bot (de @BotFather)
  TELEGRAM_CHAT_ID  id del chat o grupo; varios separados por coma
Si no existen, el mensaje solo se imprime en pantalla (modo prueba).
"""

import os
import html
import numpy as np
import pandas as pd

# =============================================================
# CONFIGURACIÓN (ajustar aquí)
# =============================================================
UMBRAL_VIENTO = 15.0        # km/h, viento medio (solo cuenta si va hacia el sur)
UMBRAL_RAFAGA = 15.0        # km/h, ráfagas (solo cuentan si van hacia el sur)
ALTURA = "10m"              # "10m" o "100m" para velocidad y dirección

HORA_MONITOREO = (6, 17)    # de 06:00 a 17:00 (inclusive)
HORA_QUEMA = (12, 16)       # ventana acordada en la mesa: 12:00 a 16:00 (inclusive)

# REGLA PRINCIPAL: manda la dirección (mismo límite que las franjas rojas del Power BI).
#   Viento HACIA el norte (entre el oeste y el este, pasando por el norte) → favorable.
#   Viento HACIA el sur (entre el este y el oeste, pasando por el sur, rumbo a las salineras)
#   → no favorable; solo ahí se reportan velocidad y ráfagas (⛔ si superan los umbrales).
# Se expresa como dirección DE DONDE VIENE (la que entrega Open-Meteo):
#   viene de 270° a 90° (del oeste, norte o este) = va hacia el sur.
DIR_DESFAVORABLE = (270.0, 90.0)

# TRANSICIÓN: si a una hora H el viento todavía es favorable pero en la hora siguiente ya entra a la
# zona no favorable, a partir de H se consideran vientos hacia condiciones no favorables y los
# MINUTOS_TRANSICION previos (H - 30 min a H) son la fase de transición (el viento se acerca al límite).
MINUTOS_TRANSICION = 30

# Criterios para "posible frente frío / norte" (heurístico, calibrar con INSIVUMEH)
FF_SECTOR_NORTE = (315.0, 45.0)  # viento que viene del NO, N o NE
FF_HORAS_NORTE = 6               # horas mínimas del día con viento del norte
FF_VIENTO_MIN = 12.0             # km/h: no cuenta la brisa terrestre débil de la madrugada
FF_SUBIDA_PRESION = 2.0          # hPa, subida de presión media respecto al día anterior
FF_BAJADA_TEMP = 1.5             # °C, bajada de temperatura media respecto al día anterior

SOLO_SI_HAY_ALERTA = False  # True = no envía nada en días sin alertas
# El mensaje diario va SOLO a los grupos de TELEGRAM_CHAT_ID (ids negativos).
# Los chats personales y los suscriptores ya no reciben copia (decisión del ICC, nov-2026).
ENVIAR_A_PERSONALES = False
LINK_PRONOSTICO = "https://app.powerbi.com/view?r=eyJrIjoiNGRkMzFlZjMtMzViZi00MjhiLWEzNDYtYzQ3MTEzNWRmNDhhIiwidCI6Ijg5MTkwMzQ4LTg5MjUtNDQxOS1hNTZhLTE3ZDIxOGE3YzgwYyIsImMiOjR9"  # Power BI; vacío = no se muestra
TZ = "America/Guatemala"

# =============================================================
# Utilidades
# =============================================================
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
PUNTOS = ["norte", "noreste", "este", "sureste", "sur", "suroeste", "oeste", "noroeste"]


def en_sector(direccion, sector):
    """True si la dirección (grados) cae en el sector (ini, fin), que puede cruzar 0°."""
    ini, fin = sector
    d = np.asarray(direccion) % 360
    if ini <= fin:
        return (d >= ini) & (d <= fin)
    return (d >= ini) | (d <= fin)


def nombre_punto(grados):
    return PUNTOS[int(((grados % 360) + 22.5) // 45) % 8]


def hacia(dir_desde):
    """Nombre de hacia dónde VA el viento. Cerca del este/oeste pero con componente al sur
    (ya en zona no favorable) se nombra este-sureste / oeste-suroeste."""
    va = (dir_desde + 180) % 360
    if 90 < va < 112.5:
        return "este-sureste"
    if 247.5 < va < 270:
        return "oeste-suroeste"
    return nombre_punto(va)


def fecha_txt(ts):
    return f"{DIAS[ts.weekday()]} {ts.day:02d}-{MESES[ts.month - 1]}"


def agrupar_horas(horas):
    """[9, 10, 11, 14] -> [(9, 11), (14, 14)]"""
    grupos = []
    for h in sorted(horas):
        if grupos and h == grupos[-1][1] + 1:
            grupos[-1][1] = h
        else:
            grupos.append([h, h])
    return [tuple(g) for g in grupos]


def rango_txt(h1, h2):
    return f"{h1:02d}:00" if h1 == h2 else f"{h1:02d}:00–{h2:02d}:00"


# =============================================================
# 1. Descarga (misma área y rejilla que 1_Median_and_average_conditions.py)
# =============================================================
def puntos_del_area(shapefile_path, res_deg=0.1):
    import geopandas as gpd
    from shapely.geometry import Point

    gdf = gpd.read_file(shapefile_path)
    if gdf.crs is None or gdf.crs.to_epsg() != 4326:
        gdf = gdf.to_crs(epsg=4326)
    try:
        poly = gdf.union_all()
    except AttributeError:
        poly = gdf.unary_union

    minx, miny, maxx, maxy = poly.bounds
    n_lon = max(3, int((maxx - minx) / res_deg) + 1)
    n_lat = max(3, int((maxy - miny) / res_deg) + 1)
    pts = [(lat, lon)
           for lat in np.linspace(miny, maxy, n_lat)
           for lon in np.linspace(minx, maxx, n_lon)
           if poly.contains(Point(lon, lat))]
    if not pts:
        raise RuntimeError("No se encontraron puntos dentro del polígono.")
    return pts


def descargar(puntos):
    import openmeteo_requests
    import requests_cache
    from retry_requests import retry

    session = retry(requests_cache.CachedSession(".cache", expire_after=3600),
                    retries=5, backoff_factor=0.2)
    client = openmeteo_requests.Client(session=session)

    variables = ["temperature_2m", "pressure_msl",
                 "wind_speed_10m", "wind_direction_10m",
                 "wind_speed_100m", "wind_direction_100m",
                 "wind_gusts_10m"]
    params = {
        "latitude": [p[0] for p in puntos],
        "longitude": [p[1] for p in puntos],
        "hourly": variables,
        "models": "ecmwf_ifs",
        "past_days": 1,          # ayer: referencia para detectar frentes fríos
        "forecast_days": 3,
        "timezone": TZ,
        "wind_speed_unit": "kmh",
    }
    # API pagada si existe el secret OPENMETEO_APIKEY; si no, la gratuita
    apikey = os.environ.get("OPENMETEO_APIKEY", "").strip()
    if apikey:
        url = "https://customer-api.open-meteo.com/v1/forecast"
        params["apikey"] = apikey
        print("Usando API comercial de Open-Meteo.")
    else:
        url = "https://api.open-meteo.com/v1/forecast"
        print("Usando API gratuita de Open-Meteo.")

    # Una sola petición para todos los puntos
    respuestas = client.weather_api(url, params=params)

    dfs = []
    for (lat, lon), r in zip(puntos, respuestas):
        h = r.Hourly()
        fechas = pd.date_range(
            start=pd.to_datetime(h.Time(), unit="s", utc=True).tz_convert(TZ),
            end=pd.to_datetime(h.TimeEnd(), unit="s", utc=True).tz_convert(TZ),
            freq=pd.Timedelta(seconds=h.Interval()), inclusive="left")
        d = {"date": fechas, "lat": lat, "lon": lon}
        for i, v in enumerate(variables):
            d[v] = h.Variables(i).ValuesAsNumpy()
        dfs.append(pd.DataFrame(d))
    return pd.concat(dfs, ignore_index=True)


# =============================================================
# 2. Resumen del área por hora (dirección con promedio vectorial)
# =============================================================
def resumir_area(df_all):
    vel, dirc = f"wind_speed_{ALTURA}", f"wind_direction_{ALTURA}"
    rad = np.deg2rad(df_all[dirc])
    df = df_all.assign(u=-df_all[vel] * np.sin(rad), v=-df_all[vel] * np.cos(rad))

    g = df.groupby("date")
    area = pd.DataFrame({
        "viento": g[vel].mean(),            # promedio del área
        "viento_max": g[vel].max(),         # punto más ventoso
        "rafaga": g["wind_gusts_10m"].max(),  # peor caso del área
        "temp": g["temperature_2m"].mean(),
        "presion": g["pressure_msl"].mean(),
        "u": g["u"].mean(), "v": g["v"].mean(),
    })
    # Dirección DE DONDE VIENE, a partir del vector medio (evita el error 350°+10° = 180°)
    area["direccion"] = (np.rad2deg(np.arctan2(-area["u"], -area["v"])) + 360) % 360
    area = area.drop(columns=["u", "v"]).reset_index()
    area["dia"] = area["date"].dt.normalize()
    area["hora"] = area["date"].dt.hour
    area["desfavorable"] = en_sector(area["direccion"], DIR_DESFAVORABLE)
    area["fuerte"] = area["viento"] > UMBRAL_VIENTO
    area["rafaga_fuerte"] = area["rafaga"] > UMBRAL_RAFAGA
    return area


CSV_POWERBI = "2_Results/Area_forecast_latest.csv"


def leer_area_powerbi(hoy):
    """Lee el MISMO archivo que usa el Power BI (lo genera el boletín ~4:00 am), para que el mensaje
    diga exactamente lo que muestra el tablero. Se toma la versión más reciente de GitHub y, si no se
    puede, la copia local. Devuelve None si el archivo no es de hoy (entonces se descarga el pronóstico)."""
    import io
    import requests
    txt = None
    tok, repo = os.environ.get("GITHUB_TOKEN"), os.environ.get("GITHUB_REPOSITORY")
    if tok and repo:
        try:
            r = requests.get(f"https://api.github.com/repos/{repo}/contents/{CSV_POWERBI}?ref=main",
                             headers={"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github.raw"},
                             timeout=30)
            if r.ok:
                txt = r.text
        except Exception as e:
            print(f"No se pudo leer el CSV de GitHub: {e}")
    if txt is None:
        ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), CSV_POWERBI)
        if os.path.exists(ruta):
            txt = open(ruta, encoding="utf-8").read()
    if not txt:
        return None
    df = pd.read_csv(io.StringIO(txt))
    df["date"] = pd.to_datetime(df["date"], utc=True).dt.tz_convert(TZ)
    if df["date"].min().normalize() != hoy or len(df) < 48:
        print(f"El CSV del Power BI no es de hoy (empieza {df['date'].min():%d/%m %H:%M}): se descarga el pronóstico.")
        return None
    area = pd.DataFrame({
        "date": df["date"],
        "viento": df[f"wind_speed_{ALTURA}"],
        "viento_max": df[f"wind_speed_{ALTURA}"],
        "rafaga": df["wind_gusts_10m"],
        "direccion": df[f"wind_direction_{ALTURA}"] % 360,
        "temp": np.nan, "presion": np.nan,
    })
    area["dia"] = area["date"].dt.normalize()
    area["hora"] = area["date"].dt.hour
    area["desfavorable"] = en_sector(area["direccion"], DIR_DESFAVORABLE)
    area["fuerte"] = area["viento"] > UMBRAL_VIENTO
    area["rafaga_fuerte"] = area["rafaga"] > UMBRAL_RAFAGA
    print(f"Datos: mismo CSV del Power BI ({len(area)} horas desde {area['date'].min():%d/%m %H:%M}).")
    return area


# =============================================================
# 3. Evaluación y mensaje
# =============================================================
def detectar_frente_frio(area):
    """Devuelve {dia: texto} con los días que muestran señales de frente frío."""
    diario = area.groupby("dia").agg(presion=("presion", "mean"), temp=("temp", "mean"))
    norte = area[en_sector(area["direccion"], FF_SECTOR_NORTE) & (area["viento"] >= FF_VIENTO_MIN)]
    diario["horas_norte"] = norte.groupby("dia").size()
    diario["horas_norte"] = diario["horas_norte"].fillna(0).astype(int)
    diario["dp"] = diario["presion"].diff()
    diario["dt"] = diario["temp"].diff()

    res = {}
    for dia, r in diario.iterrows():
        if pd.isna(r["dp"]):
            continue
        señales = []
        if r["dp"] >= FF_SUBIDA_PRESION:
            señales.append(f"presión sube {r['dp']:.1f} hPa")
        if r["dt"] <= -FF_BAJADA_TEMP:
            señales.append(f"temperatura baja {abs(r['dt']):.1f} °C")
        if r["horas_norte"] >= FF_HORAS_NORTE and señales:
            res[dia] = (f"{int(r['horas_norte'])} h con viento del norte (hacia el mar); "
                        + ", ".join(señales))
    return res


def transiciones(d):
    """Para un día: [(inicio, inicio_no_favorable, rumbo, hora_entrada)] cada vez que el viento pasa de
    favorable (hora H) a no favorable (hora H+1), con la transición empezando dentro del monitoreo
    (H hasta las 18:00: transición 17:30–18:00).
    inicio = H - MINUTOS_TRANSICION; inicio_no_favorable = H."""
    d = d.sort_values("date").reset_index(drop=True)
    res = []
    for i in range(1, len(d)):
        ant, act = d.iloc[i - 1], d.iloc[i]
        if not (HORA_MONITOREO[0] <= ant["hora"] <= HORA_MONITOREO[1] + 1):
            continue
        if act["desfavorable"] and not ant["desfavorable"] \
                and act["date"] - ant["date"] == pd.Timedelta(hours=1):
            res.append((ant["date"] - pd.Timedelta(minutes=MINUTOS_TRANSICION), ant["date"],
                        hacia(act["direccion"]), act["hora"]))
    return res


def rumbo_txt(b):
    """Hacia dónde va el viento en un bloque de horas (promedio vectorial)."""
    rad = np.deg2rad(b["direccion"])
    d = (np.rad2deg(np.arctan2(np.sin(rad).mean(), np.cos(rad).mean())) + 360) % 360
    return hacia(d)


def bloques(m, col):
    """Separa en grupos de horas seguidas con el mismo valor de `col`."""
    clave = (m["hora"].diff().ne(1) | m[col].ne(m[col].shift())).cumsum()
    return [b for _, b in m.groupby(clave)]


def alerta_txt(b):
    """Frase para un bloque de horas con viento hacia las salineras."""
    v, raf, rumbo = b["viento"].max(), b["rafaga"].max(), rumbo_txt(b)
    if b["fuerte"].any() or b["rafaga_fuerte"].any():
        return (f"⛔ Viento hacia el <b>{rumbo}</b> de {v:.0f} km/h con ráfagas de hasta "
                f"<b>{raf:.0f} km/h</b>: riesgo alto de pavesa en salineras")
    return (f"🧭 Viento hacia el <b>{rumbo}</b> ({v:.0f} km/h, ráfagas hasta {raf:.0f} km/h): "
            f"riesgo de pavesa en salineras")


def estado_frente_icc(horas=72):
    """¿El ICC emitió una alerta de frente frío al grupo en las últimas horas? Se lee de los logs del
    workflow de frentes (líneas '✓ alerta a -...'). Devuelve {'emitida', 'prob', 'desde'} o None."""
    import re
    import requests
    tok, repo = os.environ.get("GITHUB_TOKEN"), os.environ.get("GITHUB_REPOSITORY")
    if not tok or not repo:
        return None
    h = {"Authorization": f"Bearer {tok}", "Accept": "application/vnd.github+json"}
    desde = (pd.Timestamp.now(tz="UTC") - pd.Timedelta(hours=horas)).strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        runs = requests.get(f"https://api.github.com/repos/{repo}/actions/workflows/validacion_frentes.yml/runs",
                            headers=h, params={"created": f">={desde}", "per_page": 30}, timeout=30).json()
        for run in runs.get("workflow_runs", []):               # de la más reciente a la más antigua
            if run.get("conclusion") != "success":
                continue
            for j in requests.get(run["jobs_url"], headers=h, timeout=30).json().get("jobs", []):
                log = requests.get(f"https://api.github.com/repos/{repo}/actions/jobs/{j['id']}/logs",
                                   headers=h, timeout=60).text
                if re.search(r"✓ alerta a -\d+", log):
                    p = re.search(r"Probabilidad estimada: (\d+) %", log)
                    d = re.search(r"a partir del <b>([^<]+)</b>", log)
                    return {"emitida": pd.Timestamp(run["created_at"]).tz_convert(TZ),
                            "prob": p.group(1) if p else None, "desde": d.group(1) if d else None}
    except Exception as e:
        print(f"⚠ No se pudo revisar el estado de frentes fríos: {e!r}")
    return None


def construir_mensaje(area, hoy, frente_icc=None):
    area = area[area["dia"] >= hoy]
    lineas = ["🔥 <b>Perspectiva climática -Sipacate-</b>",
              "<i>Favorable: viento hacia el norte (entre el oeste y el este). "
              "No favorable: viento hacia el sur (sur, sureste y suroeste), rumbo a las salineras. "
              f"🟡 Transición: {MINUTOS_TRANSICION} min antes del cambio de dirección. "
              f"⛔ = desfavorable con viento &gt; {UMBRAL_VIENTO:.0f} km/h o ráfagas &gt; {UMBRAL_RAFAGA:.0f} km/h. "
              f"Viento a {ALTURA}.</i>"]
    hay_alerta = False
    if frente_icc:
        hay_alerta = True
        em = frente_icc["emitida"]
        fe = fecha_txt(em)
        txt = (f"🌬️ <b>Alerta de frente frío vigente</b>, emitida por el ICC el {fe[0].lower()}{fe[1:]} "
               f"a las {em:%H:%M}")
        if frente_icc.get("desde"):
            txt += f": posible llegada a partir del {frente_icc['desde']}"
        if frente_icc.get("prob"):
            txt += f" (probabilidad estimada {frente_icc['prob']} %)"
        lineas += ["", txt + ". Ver el aviso del ICC en este grupo."]

    for dia, d in area.groupby("dia"):
        lineas.append("")
        lineas.append(f"📅 <b>{fecha_txt(dia)}</b>")

        # --- Ventana de quema 12–16: solo importa la dirección ---
        q = d[d["hora"].between(*HORA_QUEMA)]
        ok = q[~q["desfavorable"]]
        trans = transiciones(d)
        inicio_nf = {t[1].hour: t[0] for t in trans}   # hora H -> inicio de la transición
        tramos, recortado = [], False
        for a, b in agrupar_horas(ok["hora"]):
            if b in inicio_nf:                         # el tramo termina donde empieza la transición
                recortado = True
                if b > a:
                    tramos.append(f"{a:02d}:00–{inicio_nf[b]:%H:%M}")
            else:
                tramos.append(rango_txt(a, b))
        if len(ok) == len(q) and not recortado:
            lineas.append(f"✅ <b>Quema: {rango_txt(*HORA_QUEMA)}</b>")
        elif not tramos:
            hay_alerta = True
            lineas.append(f"⛔ <b>Quema {rango_txt(*HORA_QUEMA)}: NO se recomienda quemar</b>")
        else:
            hay_alerta = True
            lineas.append(f"⚠️ <b>Quema: {', '.join(tramos)}</b>")

        # --- TARDE: transición cerca del final de la ventana de quema (lo que más importa a los ingenios) ---
        tarde = [t for t in trans if t[1].hour >= HORA_QUEMA[0]]
        manana = [t for t in trans if t[1].hour < HORA_QUEMA[0]]
        for ini, h_nf, rumbo, _ in tarde:
            hay_alerta = True
            fin_quema = d["date"].iloc[0].normalize() + pd.Timedelta(hours=HORA_QUEMA[1])
            reco = (f"finalizar las quemas antes de las <b>{ini:%H:%M}</b>." if ini <= fin_quema
                    else f"no extender las quemas después de las <b>{ini:%H:%M}</b>.")
            lineas += [f"🟠 <b>ATENCIÓN EN LA TARDE – fase de transición {ini:%H:%M}–{h_nf:%H:%M}</b>",
                       "El viento empieza a cambiar de dirección y se acerca a condiciones no favorables.",
                       f"A partir de las <b>{h_nf:%H:%M}</b> inician vientos no favorables.",
                       f"→ <b>Recomendación:</b> {reco}"]

        # --- 06–17: horas con viento hacia las salineras (y transiciones de la mañana, sin detalle) ---
        m = d[d["hora"].between(*HORA_MONITOREO)]
        bl = bloques(m, "desfavorable")
        malas = [b for b in bl if b["desfavorable"].iloc[0]]
        items = []                                     # (momento, texto) en orden cronológico
        for ini, h_nf, _, _ in manana:
            items.append((ini, f"🟡 Transición {ini:%H:%M}–{h_nf:%H:%M}"))
        for b in malas:
            items.append((b["date"].iloc[0],
                          f"• {rango_txt(b['hora'].iloc[0], b['hora'].iloc[-1])}: {alerta_txt(b)}"))
        lineas += [t for _, t in sorted(items, key=lambda x: x[0])]
        if malas or manana:
            hay_alerta = True
            lineas.append("→ <b>Recomendación:</b> No realizar quemas en esos horarios"
                          + (" ni durante la transición." if manana else "."))


    if LINK_PRONOSTICO:
        lineas += ["",
                   "📊 Para mayor información sobre las proyecciones de velocidad y dirección de viento, "
                   "consulte el siguiente enlace:",
                   html.escape(LINK_PRONOSTICO)]
    return "\n".join(lineas), hay_alerta


# =============================================================
# 4. Envío por Telegram
# =============================================================
# --- Lista de suscriptores (chat personal) -------------------------------
# Quien toca "Iniciar" en el bot queda registrado y recibe el mensaje en privado,
# siempre que siga siendo miembro del grupo. La lista se guarda CIFRADA en el
# repositorio (es público), con una clave derivada del TELEGRAM_TOKEN.
ARCHIVO_SUSCRIPTORES = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "3_Suscriptores", "suscriptores.enc")
ESTADOS_MIEMBRO = {"creator", "administrator", "member", "restricted"}


def _token():
    import re
    crudo = os.environ.get("TELEGRAM_TOKEN", "")
    # Tolera token pegado con comillas, la palabra "bot" delante o la URL completa
    m = re.search(r"\d{6,}:[A-Za-z0-9_-]{30,}", crudo)
    if crudo and not m:
        print(f"⚠ El TELEGRAM_TOKEN no tiene el formato esperado (número:clave). Longitud: {len(crudo)}")
    return m.group(0) if m else crudo.strip()


def _api(token):
    import requests

    def llamar(metodo, **params):
        r = requests.post(f"https://api.telegram.org/bot{token}/{metodo}", data=params, timeout=30)
        try:
            j = r.json()
        except Exception:
            j = {"ok": False, "description": r.text}
        j["_status"] = r.status_code
        return j
    return llamar


def _fernet(token):
    import base64
    import hashlib
    from cryptography.fernet import Fernet
    clave = hashlib.sha256(b"sipacate-suscriptores:" + token.encode()).digest()
    return Fernet(base64.urlsafe_b64encode(clave))


def cargar_suscriptores(token):
    import json
    vacio = {"offset": 0, "personas": {}}
    if not os.path.exists(ARCHIVO_SUSCRIPTORES):
        return vacio
    try:
        with open(ARCHIVO_SUSCRIPTORES, "rb") as f:
            return json.loads(_fernet(token).decrypt(f.read()))
    except Exception:
        print("⚠ No se pudo leer la lista de suscriptores (¿cambió el token?). Se empieza una lista nueva.")
        return vacio


def guardar_suscriptores(token, subs):
    import json
    os.makedirs(os.path.dirname(ARCHIVO_SUSCRIPTORES), exist_ok=True)
    datos = json.dumps(subs, ensure_ascii=False, sort_keys=True).encode()
    with open(ARCHIVO_SUSCRIPTORES, "wb") as f:
        f.write(_fernet(token).encrypt(datos))


def registrar_nuevos(api, subs):
    """Lee quién tocó Iniciar (o bloqueó el bot) desde la última revisión."""
    import json
    params = {"timeout": 0, "allowed_updates": json.dumps(["message", "my_chat_member"])}
    if subs.get("offset"):
        params["offset"] = subs["offset"] + 1
    r = api("getUpdates", **params)
    if not r.get("ok"):
        print(f"⚠ No se pudo leer getUpdates: {r.get('description')}")
        return []
    nuevos = []
    hoy = pd.Timestamp.now(tz=TZ).strftime("%Y-%m-%d")
    for u in r.get("result", []):
        subs["offset"] = max(subs.get("offset", 0), u["update_id"])
        msg = u.get("message") or {}
        chat = msg.get("chat") or {}
        if chat.get("type") == "private" and (msg.get("text") or "").startswith("/start"):
            cid = str(chat["id"])
            if cid not in subs["personas"]:
                subs["personas"][cid] = {"nombre": chat.get("first_name", ""), "desde": hoy}
                nuevos.append(cid)
        cm = u.get("my_chat_member") or {}
        if (cm.get("chat") or {}).get("type") == "private" and \
                (cm.get("new_chat_member") or {}).get("status") == "kicked":
            subs["personas"].pop(str(cm["chat"]["id"]), None)   # bloqueó el bot
    return nuevos


def enviar_telegram(texto):
    import re
    import json
    token = _token()
    fijos = re.findall(r"-?\d+", os.environ.get("TELEGRAM_CHAT_ID", ""))
    if not token or not fijos:
        print("⚠ TELEGRAM_TOKEN / TELEGRAM_CHAT_ID no configurados: solo se muestra el mensaje.")
        return
    api = _api(token)
    grupo = next((c for c in fijos if c.startswith("-")), None)

    def enviar(chat, txt):
        return api("sendMessage", chat_id=chat, text=txt, parse_mode="HTML",
                   disable_web_page_preview="true")

    # 1) Destinos: solo los grupos (salvo que se active ENVIAR_A_PERSONALES). Si fallan, el workflow falla
    destinos = fijos if ENVIAR_A_PERSONALES else [c for c in fijos if c.startswith("-")]
    errores = []
    for chat in destinos:
        r = enviar(chat, texto)
        if r.get("ok"):
            print(f"✓ Mensaje enviado a {chat}")
        else:
            print(f"✗ Error con {chat}: {r['_status']} {r.get('description')}")
            errores.append(chat)

    # 2) Suscriptores (copia en el chat personal): desactivado
    if ENVIAR_A_PERSONALES:
        try:
            enviar_suscriptores(api, token, texto, fijos, grupo, enviar)
        except Exception as e:
            print(f"⚠ Error en el envío a suscriptores (el grupo sí recibió el mensaje): {e!r}")

    if errores:
        raise RuntimeError(f"No se pudo enviar a: {', '.join(errores)}")


def enviar_suscriptores(api, token, texto, fijos, grupo, enviar):
    import json
    subs = cargar_suscriptores(token)
    antes = json.dumps(subs, sort_keys=True)
    nuevos = registrar_nuevos(api, subs)
    enviados = 0
    for cid in sorted(subs["personas"]):
        if cid in fijos:
            continue
        if grupo:
            m = api("getChatMember", chat_id=grupo, user_id=cid)
            if m.get("ok") and m["result"].get("status") not in ESTADOS_MIEMBRO:
                print(f"· {cid} ya no está en el grupo: no se le envía")
                continue
        if cid in nuevos:
            enviar(cid, "✅ Quedaste registrado en la <b>Perspectiva climática -Sipacate-</b>. "
                        "Desde hoy recibirás aquí el pronóstico diario.")
        r = enviar(cid, texto)
        if r.get("ok"):
            enviados += 1
        elif r["_status"] == 403:                      # bloqueó el bot o borró el chat
            print(f"· {cid} bloqueó el bot: se quita de la lista")
            subs["personas"].pop(cid, None)
        else:
            print(f"⚠ No se pudo enviar a suscriptor {cid}: {r['_status']} {r.get('description')}")
    print(f"✓ Suscriptores: {enviados} mensajes privados enviados, {len(nuevos)} nuevos, "
          f"{len(subs['personas'])} en la lista")

    if os.environ.get("GUARDAR_SUSCRIPTORES") == "1" and json.dumps(subs, sort_keys=True) != antes:
        guardar_suscriptores(token, subs)
        print("✓ Lista de suscriptores actualizada")


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    shp = os.path.join(base, "1_Shapefile", "Area_of_interest.shp")

    # Se espera primero a la hora de envío, para usar el CSV más reciente del boletín
    esperar_hasta(os.environ.get("ESPERAR_HASTA", ""))

    hoy = pd.Timestamp.now(tz=TZ).normalize()
    area = leer_area_powerbi(hoy)
    if area is None:                      # respaldo: si el boletín de hoy no está, se descarga
        puntos = puntos_del_area(shp)
        print(f"Usando {len(puntos)} puntos dentro del área.")
        area = resumir_area(descargar(puntos))

    texto, hay_alerta = construir_mensaje(area, hoy, estado_frente_icc())
    print(texto)

    if SOLO_SI_HAY_ALERTA and not hay_alerta:
        print("Sin alertas: no se envía mensaje.")
        return
    enviar_telegram(texto)


def esperar_hasta(hora_txt):
    """Si se indica (ej. '07:00'), espera hasta esa hora de Guatemala antes de enviar.
    Si ya pasó, envía de inmediato."""
    import time
    if not hora_txt.strip():
        return
    h, m = (int(x) for x in hora_txt.strip().split(":"))
    ahora = pd.Timestamp.now(tz=TZ)
    objetivo = ahora.normalize() + pd.Timedelta(hours=h, minutes=m)
    falta = (objetivo - ahora).total_seconds()
    if falta > 0:
        print(f"Esperando {falta / 60:.0f} min para enviar a las {hora_txt} (hora Guatemala)...")
        time.sleep(falta)
    else:
        print(f"Ya pasó de las {hora_txt}: se envía de inmediato.")


if __name__ == "__main__":
    main()
