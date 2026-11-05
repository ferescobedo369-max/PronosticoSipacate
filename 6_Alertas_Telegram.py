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
FF_SECTOR_NORTE = (315.0, 45.0)  # viento que viene del NO, N o NE (va hacia SE, S o SO)
FF_VIENTO_MIN = 12.0             # km/h: no cuenta la brisa terrestre débil de la madrugada
FF_HORAS_SEGUIDAS = 4            # horas seguidas de viento del norte para marcar el ingreso
FF_HORAS_NORTE = 6               # horas totales mínimas de viento del norte durante el evento
FF_PAUSA_MAX = 12                # h: pausas más cortas no cortan el evento
FF_SUBIDA_PRESION = 2.0          # hPa que sube la presión respecto a las 24 h previas
FF_BAJADA_TEMP = 1.5             # °C que baja la temperatura mínima respecto a días anteriores
FF_DIAS_ADELANTE = 5             # días hacia adelante en que se busca el ingreso del frente

DIAS_MENSAJE = 3                 # días que muestra el mensaje diario

SOLO_SI_HAY_ALERTA = False  # True = no envía nada en días sin alertas
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

    variables = ["temperature_2m", "pressure_msl", "relative_humidity_2m",
                 "wind_speed_10m", "wind_direction_10m",
                 "wind_speed_100m", "wind_direction_100m",
                 "wind_gusts_10m"]
    params = {
        "latitude": [p[0] for p in puntos],
        "longitude": [p[1] for p in puntos],
        "hourly": variables,
        "models": "ecmwf_ifs",
        "past_days": 2,          # referencia para detectar frentes fríos
        "forecast_days": FF_DIAS_ADELANTE + 1,
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
    if "relative_humidity_2m" in df_all:
        area["humedad"] = g["relative_humidity_2m"].mean()
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
def detectar_frente_frio(area, ahora):
    """Busca el ingreso de un frente frío (desde ayer hasta FF_DIAS_ADELANTE días).
    Señales: varias horas seguidas de viento del norte (hacia el mar y las salineras)
    + subida de presión y/o baja de temperatura. Devuelve un dict o None."""
    a = area.sort_values("date").reset_index(drop=True)
    a["norte"] = en_sector(a["direccion"], FF_SECTOR_NORTE) & (a["viento"] >= FF_VIENTO_MIN)
    tramo = (a["norte"] != a["norte"].shift()).cumsum()
    norte = a[a["norte"]]
    limite = ahora + pd.Timedelta(days=FF_DIAS_ADELANTE)
    usados_hasta = None

    for _, b in norte.groupby(tramo[a["norte"]]):
        ini = b["date"].iloc[0]
        if len(b) < FF_HORAS_SEGUIDAS or ini > limite:
            continue
        if usados_hasta is not None and ini <= usados_hasta:
            continue
        # Fin del evento: se extiende mientras las pausas sin viento del norte sean cortas
        fin = b["date"].iloc[-1]
        for t in norte["date"][norte["date"] > fin]:
            if (t - fin) <= pd.Timedelta(hours=FF_PAUSA_MAX):
                fin = t
            else:
                break
        usados_hasta = fin
        if fin < ahora:
            continue                                  # ya terminó
        dur = a[(a["date"] >= ini) & (a["date"] <= fin)]
        horas_norte = dur[dur["norte"]]
        if len(horas_norte) < FF_HORAS_NORTE:
            continue
        antes = a[(a["date"] >= ini - pd.Timedelta(hours=24)) & (a["date"] < ini)]
        d0 = ini.normalize()
        ref = a[(a["dia"] < d0) & (a["dia"] >= d0 - pd.Timedelta(days=2))]
        if antes.empty or ref.empty:
            continue
        dp = dur["presion"].max() - antes["presion"].mean()
        tmin = dur["temp"].min()
        baja_t = ref.groupby("dia")["temp"].min().mean() - tmin
        if dp < FF_SUBIDA_PRESION and baja_t < FF_BAJADA_TEMP:
            continue                                  # viento del norte sin señales de frente
        desde = max(ini, ahora)
        horas = max(1, (fin - desde).total_seconds() / 3600)
        return {
            "inicio": ini, "fin": fin, "en_curso": ini <= ahora,
            "horas": int(np.ceil(horas / 24) * 24),
            "rumbo": rumbo_txt(horas_norte),
            "viento": horas_norte["viento"].mean(),
            "rafaga": dur["rafaga"].max(),
            "tmin": tmin, "baja_t": baja_t, "dp": dp,
            "humedad": dur["humedad"].min() if "humedad" in dur else None,
        }
    return None


def mensaje_frente_frio(ev, ahora):
    """Texto de la alerta de frente frío (revisado y aprobado por el ICC)."""
    ini = ev["inicio"]
    cuando = f"{fecha_txt(ini)[0].lower()}{fecha_txt(ini)[1:]}, alrededor de las {ini.hour:02d}:00"
    if ev["en_curso"]:
        titulo = "🌬️ <b>INFLUENCIA DE FRENTE FRÍO</b>"
        intro = (f"Se mantiene la influencia del frente frío que ingresó el <b>{cuando}</b>, con posibles "
                 "condiciones adversas para las quemas de caña de azúcar en Sipacate y su zona buffer.")
    else:
        titulo = "🌬️ <b>POSIBLE INGRESO DE FRENTE FRÍO</b>"
        intro = (f"Se prevé que a partir del <b>{cuando}</b>, ingrese un frente frío, con posibles "
                 "condiciones adversas para las quemas de caña de azúcar en Sipacate y su zona buffer.")
    pronto = ev["en_curso"] or (ini - ahora) <= pd.Timedelta(hours=24)
    durante = "durante las próximas" if pronto else "durante unas"

    l = [titulo, "<b>Perspectiva climática -Sipacate-</b>", "",
         "⛔ <b>Condiciones adversas para las quemas</b>", intro, "",
         "⚠️ <b>Alerta de condiciones</b>",
         f"• <b>Duración:</b> se estima que {durante} <b>{ev['horas']} horas</b> predominen vientos "
         f"con dirección <b>hacia el {ev['rumbo']}</b>.",
         f"• <b>Viento:</b> velocidad promedio de <b>{ev['viento']:.0f} km/h</b> y ráfagas de hasta "
         f"<b>{ev['rafaga']:.0f} km/h</b>."]
    t = f"• <b>Temperatura:</b> podría descender hasta <b>{ev['tmin']:.0f} °C</b>"
    if ev["baja_t"] >= 1:
        t += f", unos {ev['baja_t']:.0f} °C menos que en días anteriores"
    l.append(t + ".")
    if ev.get("humedad") is not None and not np.isnan(ev["humedad"]):
        l.append(f"• <b>Humedad relativa:</b> podría bajar hasta <b>{ev['humedad']:.0f} %</b>.")
    if ev["dp"] >= 1:
        l.append(f"• <b>Presión atmosférica:</b> aumento de <b>{ev['dp']:.0f} hPa</b>, "
                 "una señal típica del ingreso del frente.")
    l += ["",
          "📌 <b>Recomendación:</b> seguir las instrucciones emitidas por <b>ASAZGUA</b> al momento "
          "de realizar cualquier quema durante el día.", "",
          "🔎 El <b>ICC</b> dará seguimiento a este fenómeno y a la evolución de las condiciones "
          "atmosféricas. Manténgase atento a este grupo para las actualizaciones."]
    if LINK_PRONOSTICO:
        l += ["", "📊 Para mayor información sobre las proyecciones de velocidad y dirección de viento, "
                  "consulte el siguiente enlace:", html.escape(LINK_PRONOSTICO)]
    return "\n".join(l)


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


def construir_mensaje(area, hoy, frente=None):
    area = area[(area["dia"] >= hoy) & (area["dia"] < hoy + pd.Timedelta(days=DIAS_MENSAJE))]
    lineas = ["🔥 <b>Perspectiva climática -Sipacate-</b>",
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
            lineas.append("  → <b>Recomendación:</b> No realizar quemas en esos horarios.")

        # --- Cambios de dirección ---
        cambios = []
        for nxt in bl[1:]:
            h, r = nxt["hora"].iloc[0], hacia(nxt["direccion"].iloc[0])
            cambios.append(f"{h:02d}:00 gira hacia el {r} " + ("⚠️" if nxt["desfavorable"].iloc[0] else "✅"))
        if cambios:
            lineas.append("  🔄 Cambio de dirección de viento: " + " · ".join(cambios))

        if frente and frente["inicio"].normalize() <= dia <= frente["fin"].normalize():
            hay_alerta = True
            lineas.append("  🌬️ <b>Posible influencia de frente frío</b> (ver alerta aparte)")

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


def enviar_telegram(*textos):
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

    # 1) Destinos fijos (grupo y chats del secret): si fallan, el workflow falla
    errores = []
    for chat in fijos:
        for n, texto in enumerate(textos, 1):
            r = enviar(chat, texto)
            if r.get("ok"):
                print(f"✓ Mensaje {n} enviado a {chat}")
            else:
                print(f"✗ Error con {chat} (mensaje {n}): {r['_status']} {r.get('description')}")
                errores.append(chat)

    # 2) Suscriptores: quienes tocaron Iniciar y siguen en el grupo.
    #    Si algo falla aquí, no afecta los envíos fijos de arriba.
    try:
        enviar_suscriptores(api, token, textos, fijos, grupo, enviar)
    except Exception as e:
        print(f"⚠ Error en el envío a suscriptores (el grupo sí recibió el mensaje): {e!r}")

    if errores:
        raise RuntimeError(f"No se pudo enviar a: {', '.join(errores)}")


def enviar_suscriptores(api, token, textos, fijos, grupo, enviar):
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
        for texto in textos:
            r = enviar(cid, texto)
            if not r.get("ok"):
                break
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

    puntos = puntos_del_area(shp)
    print(f"Usando {len(puntos)} puntos dentro del área.")
    area = resumir_area(descargar(puntos))

    ahora = pd.Timestamp.now(tz=TZ)
    hoy = ahora.normalize()
    frente = detectar_frente_frio(area, ahora)
    texto, hay_alerta = construir_mensaje(area, hoy, frente)
    textos = [texto]
    print(texto)
    if frente:
        textos.append(mensaje_frente_frio(frente, ahora))
        print("\n" + "=" * 60 + "\n" + textos[-1])
    else:
        print("\n(Sin señales de frente frío en los próximos días)")

    if os.environ.get("SOLO_MOSTRAR") == "1":
        print("\nSOLO_MOSTRAR=1: no se envía nada (modo revisión).")
        return
    if SOLO_SI_HAY_ALERTA and not hay_alerta:
        print("Sin alertas: no se envía mensaje.")
        return
    esperar_hasta(os.environ.get("ESPERAR_HASTA", ""))
    enviar_telegram(*textos)


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
