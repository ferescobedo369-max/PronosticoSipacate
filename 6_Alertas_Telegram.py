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
  3. Posible influencia de frente frío (norte).

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

# REGLA PRINCIPAL: manda la dirección.
#   Viento HACIA el norte, noreste, noroeste, este u oeste → favorable (sin importar velocidad).
#   Viento HACIA el sur, sureste o suroeste → desfavorable; solo ahí se reportan
#   velocidad y ráfagas (⛔ si superan los umbrales).
# Se expresa como dirección DE DONDE VIENE (la que entrega Open-Meteo):
#   viene del NO, N o NE (292.5° a 67.5°) = va hacia SE, S o SO.
DIR_DESFAVORABLE = (292.5, 67.5)

# Criterios para "posible frente frío / norte" (heurístico, calibrar con INSIVUMEH)
FF_SECTOR_NORTE = (315.0, 45.0)  # viento que viene del NO, N o NE
FF_HORAS_NORTE = 6               # horas mínimas del día con viento del norte
FF_VIENTO_MIN = 12.0             # km/h: no cuenta la brisa terrestre débil de la madrugada
FF_SUBIDA_PRESION = 2.0          # hPa, subida de presión media respecto al día anterior
FF_BAJADA_TEMP = 1.5             # °C, bajada de temperatura media respecto al día anterior

SOLO_SI_HAY_ALERTA = False  # True = no envía nada en días sin alertas
LINK_PRONOSTICO = ""        # enlace al pronóstico interactivo (Power BI); vacío = no se muestra
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
    """Nombre de hacia dónde VA el viento."""
    return nombre_punto((dir_desde + 180) % 360)


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


def construir_mensaje(area, hoy):
    frentes = detectar_frente_frio(area)   # usa también el día anterior como referencia
    area = area[area["dia"] >= hoy]
    lineas = ["🔥 <b>Pronóstico de viento para quemas – Sipacate</b>",
              "<i>Favorable: viento hacia el norte, noreste, noroeste, este u oeste. "
              "Desfavorable: hacia el sur, sureste o suroeste. "
              f"⛔ = desfavorable con viento &gt; {UMBRAL_VIENTO:.0f} km/h o ráfagas &gt; {UMBRAL_RAFAGA:.0f} km/h. "
              f"Viento a {ALTURA}.</i>"]
    hay_alerta = False

    for dia, d in area.groupby("dia"):
        lineas.append("")
        lineas.append(f"📅 <b>{fecha_txt(dia)}</b>")

        # --- Ventana de quema 12–16: solo importa la dirección ---
        q = d[d["hora"].between(*HORA_QUEMA)]
        ok = q[~q["desfavorable"]]
        if len(ok) == len(q):
            lineas.append(f"✅ Quema {rango_txt(*HORA_QUEMA)}: <b>condiciones favorables</b> "
                          f"(viento hacia el {rumbo_txt(q)})")
        elif ok.empty:
            hay_alerta = True
            lineas.append(f"⛔ Quema {rango_txt(*HORA_QUEMA)}: <b>NO se recomienda quemar</b>")
        else:
            hay_alerta = True
            buenas = ", ".join(rango_txt(a, b) for a, b in agrupar_horas(ok["hora"]))
            lineas.append(f"⚠️ Quema: <b>solo recomendable {buenas}</b>")

        # --- 06–17: horas con viento hacia las salineras ---
        m = d[d["hora"].between(*HORA_MONITOREO)]
        bl = bloques(m, "desfavorable")
        malas = [b for b in bl if b["desfavorable"].iloc[0]]
        for b in malas:
            lineas.append(f"  • {rango_txt(b['hora'].iloc[0], b['hora'].iloc[-1])}: {alerta_txt(b)}")
        if malas:
            hay_alerta = True
            lineas.append("  → Se recomienda <b>no realizar ningún tipo de quema</b> en esos horarios.")

        # --- Cambios de dirección ---
        cambios = []
        for nxt in bl[1:]:
            h, r = nxt["hora"].iloc[0], hacia(nxt["direccion"].iloc[0])
            cambios.append(f"{h:02d}:00 gira hacia el {r} " + ("⚠️" if nxt["desfavorable"].iloc[0] else "✅"))
        if cambios:
            lineas.append("  🔄 Cambio de dirección: " + " · ".join(cambios))

        if dia in frentes:
            hay_alerta = True
            lineas.append(f"  🌬️ <b>Posible influencia de frente frío</b>: {html.escape(frentes[dia])}")

    lineas += ["",
               "🧭 hacia las salineras   ⛔ hacia las salineras con viento o ráfagas fuertes   🔄 cambio de dirección"]
    if LINK_PRONOSTICO:
        lineas.append(f'📊 <a href="{html.escape(LINK_PRONOSTICO)}">Consulte el pronóstico interactivo</a>')
    lineas += ["<i>Pronóstico ECMWF vía Open-Meteo; puede cambiar, revisar la actualización diaria.</i>"]
    return "\n".join(lineas), hay_alerta


# =============================================================
# 4. Envío por Telegram
# =============================================================
def enviar_telegram(texto):
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    chats = [c.strip() for c in os.environ.get("TELEGRAM_CHAT_ID", "").split(",") if c.strip()]
    if not token or not chats:
        print("⚠ TELEGRAM_TOKEN / TELEGRAM_CHAT_ID no configurados: solo se muestra el mensaje.")
        return
    import requests
    for chat in chats:
        r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                          data={"chat_id": chat, "text": texto, "parse_mode": "HTML",
                                "disable_web_page_preview": "true"}, timeout=30)
        if r.ok:
            print(f"✓ Mensaje enviado a {chat}")
        else:
            raise RuntimeError(f"Telegram respondió {r.status_code}: {r.text}")


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    shp = os.path.join(base, "1_Shapefile", "Area_of_interest.shp")

    puntos = puntos_del_area(shp)
    print(f"Usando {len(puntos)} puntos dentro del área.")
    area = resumir_area(descargar(puntos))

    hoy = pd.Timestamp.now(tz=TZ).normalize()
    texto, hay_alerta = construir_mensaje(area, hoy)
    print(texto)

    if SOLO_SI_HAY_ALERTA and not hay_alerta:
        print("Sin alertas: no se envía mensaje.")
        return
    enviar_telegram(texto)


if __name__ == "__main__":
    main()
