"""
Script 8: VALIDACIÓN de alertas de frente frío – Sipacate (rama validacion-frentes).

NO envía nada al grupo y NO guarda nada en el repositorio. Cada día manda SOLO a los
chats personales del secret TELEGRAM_CHAT_ID (los ids positivos) un resumen para
comparar tres fuentes:

  1. Pronóstico ECMWF (Open-Meteo) en Sipacate: viento del norte, presión,
     temperatura y humedad para los próximos días, con una gráfica.
  2. NOAA – Tropical Weather Discussion (Atlántico y Pacífico Este): frases de los
     meteorólogos que mencionan frentes fríos cerca de la región.
  3. NOAA – Análisis de superficie (mapa con los frentes dibujados).

Si el modelo detecta un frente, también se muestra la alerta tal como la vería el grupo.
Reutiliza la detección de 6_Alertas_Telegram.py, así la validación prueba el mismo código.
"""

import os
import re
import io
import html
import importlib.util

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("alertas", os.path.join(BASE, "6_Alertas_Telegram.py"))
al = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(al)
TZ = al.TZ

URL_MAPA_NOAA = "https://www.nhc.noaa.gov/tafb_latest/USA_latest.gif"
URL_SAT_INSIVUMEH = "https://insivumeh.gob.gt/img/guate_truecolor.png"        # GOES-19 color verdadero
URL_BOLETIN_INSIVUMEH = "https://insivumeh.gob.gt/img/Papa_Bravo/Papa_iFoxtrot/Boletin_Informativo.pdf"
URL_PRONOSTICO_INSIVUMEH = "https://insivumeh.gob.gt/img/Papa_Bravo/Papa_Noviembre/pronostico_diario.pdf"
# El boletín PDF del INSIVUMEH (nov-2026) es del 23-abr y casi no tiene texto (es imagen):
# se deja apagado hasta encontrar una fuente de texto que se actualice.
USAR_BOLETIN_INSIVUMEH = False
CLAVES_INSIVUMEH = ["frente frío", "frente frio", "empuje de aire frío", "empuje de aire frio", "aire frío",
                    "vientos del norte", "viento del norte", "nortes", "alta presión", "alta presion"]
URLS_TWD = {
    "Atlántico / Golfo": ["https://tgftp.nws.noaa.gov/data/raw/ax/axnt20.knhc.twd.at.txt",
                          "https://www.nhc.noaa.gov/text/MIATWDAT.shtml"],
    "Pacífico Este": ["https://tgftp.nws.noaa.gov/data/raw/ax/axpz20.knhc.twd.ep.txt",
                      "https://www.nhc.noaa.gov/text/MIATWDEP.shtml"],
}
# Zonas que importan para Guatemala (la NOAA ahora usa "Gulf of America" para el Golfo de México)
# Zonas CERCANAS a Guatemala: si la NOAA menciona un frente aquí, se avisa para seguimiento
ZONAS_CERCA = ["yucatan", "campeche", "tehuantepec", "central america", "guatemala", "belize", "honduras",
               "sw gulf", "southwest gulf", "se gulf", "southeast gulf", "southern gulf", "south-central gulf",
               "nw caribbean", "northwest caribbean"]
PDF_INSIVUMEH = {}                       # copia del pronóstico del día (para enviarla al grupo si se emite)
ZONAS = ["gulf", "yucatan", "campeche", "tehuantepec", "central america", "guatemala",
         "mexico", "caribbean", "honduras", "belize"]
UA = {"User-Agent": "PronosticoSipacate-validacion (GitHub Actions)"}


# -------------------------------------------------------------------------
# 1. ECMWF: resumen diario de señales
# -------------------------------------------------------------------------
def resumen_diario(area, hoy):
    a = area.copy()
    a["norte"] = al.en_sector(a["direccion"], al.FF_SECTOR_NORTE) & (a["viento"] >= al.FF_VIENTO_MIN)
    agg = {"horas_norte": ("norte", "sum"), "viento_max": ("viento", "max"),
           "rafaga_max": ("rafaga", "max"), "presion": ("presion", "mean"), "tmin": ("temp", "min")}
    if "humedad" in a:
        agg["hr_min"] = ("humedad", "min")
    d = a.groupby("dia").agg(**agg)
    d["dp"] = d["presion"].diff()
    d["dtmin"] = d["tmin"].diff()
    return d[d.index >= hoy]


def texto_ecmwf(d, ev):
    l = ["1️⃣ <b>Modelo ECMWF – Sipacate</b> (viento del norte = va hacia el mar/salineras)"]
    for dia, r in d.iterrows():
        f = al.fecha_txt(dia)[:3] + " " + dia.strftime("%d")
        hr = f" · HR mín {r['hr_min']:.0f} %" if "hr_min" in r and not np.isnan(r["hr_min"]) else ""
        dp = f"{r['dp']:+.1f}" if not np.isnan(r["dp"]) else "–"
        dt = f"{r['dtmin']:+.1f}" if not np.isnan(r["dtmin"]) else "–"
        marca = "🌬️" if r["horas_norte"] >= al.FF_HORAS_NORTE else "▫️"
        l.append(f"{marca} <b>{f}</b>: norte {int(r['horas_norte'])} h · viento máx {r['viento_max']:.0f} · "
                 f"ráfaga {r['rafaga_max']:.0f} km/h · presión {dp} hPa · T mín {r['tmin']:.0f} °C ({dt}){hr}")
    if ev:
        l.append(f"→ <b>Señal de frente frío</b>: ingreso {al.fecha_txt(ev['inicio'])} "
                 f"{ev['inicio'].hour:02d}:00 (presión +{ev['dp']:.1f} hPa, T mín −{ev['baja_t']:.1f} °C)")
    else:
        l.append("→ Sin señal de frente frío según los criterios actuales")
    return "\n".join(l)


def grafica(area, ev, ruta):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates

    a = area[area["date"] >= area["date"].max() - pd.Timedelta(days=7)].copy()
    t = a["date"].dt.tz_localize(None)
    norte = al.en_sector(a["direccion"], al.FF_SECTOR_NORTE)
    paneles = [("presion", "Presión (hPa)"), ("temp", "Temperatura (°C)")]
    if "humedad" in a:
        paneles.append(("humedad", "Humedad relativa (%)"))
    fig, ax = plt.subplots(len(paneles) + 1, 1, figsize=(10, 2.3 * (len(paneles) + 1)), sharex=True)
    for k, (col, tit) in enumerate(paneles):
        ax[k].plot(t, a[col], color="#1f4e79", lw=1.6)
        ax[k].set_ylabel(tit, fontsize=9)
        ax[k].grid(alpha=.3)
    w = ax[-1]
    w.plot(t, a["viento"], color="#555", lw=1.2, label="Viento medio")
    w.plot(t, a["rafaga"], color="#999", lw=1, ls="--", label="Ráfaga máx.")
    w.scatter(t[norte], a["viento"][norte], color="#c0392b", s=12, zorder=3,
              label="Hacia el sur (viento del norte)")
    w.set_ylabel("Viento (km/h)", fontsize=9)
    w.legend(fontsize=8, loc="upper left")
    w.grid(alpha=.3)
    ahora = pd.Timestamp.now(tz=TZ).tz_localize(None)
    for x in ax:
        x.axvline(ahora, color="k", lw=.8, ls=":")
        if ev:
            x.axvspan(ev["inicio"].tz_localize(None), ev["fin"].tz_localize(None), color="#3498db", alpha=.12)
    w.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b\n%H:%M"))
    titulo = "ECMWF – Sipacate (área promedio). Línea punteada = ahora"
    if ev:
        titulo += ". Franja azul = frente detectado"
    fig.suptitle(titulo, fontsize=10)
    fig.tight_layout()
    fig.savefig(ruta, dpi=110)
    plt.close(fig)


# -------------------------------------------------------------------------
# 1b. PROBABILIDAD con el ensamble ECMWF (51 escenarios)
# -------------------------------------------------------------------------
URL_ENSAMBLE = "https://ensemble-api.open-meteo.com/v1/ensemble"
MODELOS_ENSAMBLE = ["ecmwf_ifs025", "gfs025"]          # si el ECMWF no responde, se usa el GFS
P_MIN_AVISO = 10           # % mínimo del ensamble para avisar al chat personal
P_CONFIRMADO = 80          # % del ensamble para considerar el frente confirmado (con fuente oficial)
HORAS_CONFIRMADO = 48      # el frente confirmado debe llegar dentro de estas horas (o ya estar)


def descargar_ensamble(lat, lon):
    """Devuelve (modelo, {miembro: DataFrame}) con viento, dirección, presión y temperatura."""
    import requests
    vars_ = ["wind_speed_10m", "wind_direction_10m", "pressure_msl", "temperature_2m"]
    for modelo in MODELOS_ENSAMBLE:
        try:
            r = requests.get(URL_ENSAMBLE, timeout=90, params={
                "latitude": lat, "longitude": lon, "hourly": ",".join(vars_), "models": modelo,
                "forecast_days": al.FF_DIAS_ADELANTE + 2, "past_days": 2, "timezone": TZ,
                "wind_speed_unit": "kmh"})
            if not r.ok:
                print(f"⚠ Ensamble {modelo}: {r.status_code} {r.text[:200]}")
                continue
            h = r.json()["hourly"]
            t = pd.to_datetime(h["time"]).tz_localize(TZ)
            miembros = {}
            for k, vals in h.items():
                m = re.match(r"(" + "|".join(vars_) + r")(?:_member(\d+))?(?:_" + modelo + r")?$", k)
                if m:
                    miembros.setdefault(int(m.group(2) or 0), {})[m.group(1)] = vals
            dfs = {}
            for mid, d in miembros.items():
                if len(d) == len(vars_):
                    df = pd.DataFrame(d, index=t).apply(pd.to_numeric, errors="coerce").dropna()
                    if len(df) > 24:
                        dfs[mid] = df
            if len(dfs) >= 10:
                print(f"✓ Ensamble {modelo}: {len(dfs)} escenarios, {len(t)} pasos de tiempo")
                return modelo, dfs
            print(f"⚠ Ensamble {modelo}: solo {len(dfs)} escenarios. Claves: {list(h)[:8]}")
        except Exception as e:
            print(f"⚠ Ensamble {modelo}: {e!r}")
    return None, {}


def dias_con_frente(df, hoy):
    """Días (de hoy a FF_DIAS_ADELANTE) en que este escenario muestra señales de frente frío."""
    idx = df.index
    paso = idx.to_series().diff().shift(-1).dt.total_seconds().div(3600)
    paso = paso.fillna(paso.median())
    norte = al.en_sector(df["wind_direction_10m"], al.FF_SECTOR_NORTE) & (df["wind_speed_10m"] >= al.FF_VIENTO_MIN)
    dia = idx.normalize()
    res = set()
    for k in range(al.FF_DIAS_ADELANTE + 1):
        d = hoy + pd.Timedelta(days=k)
        hoy_m, ayer_m = dia == d, dia == d - pd.Timedelta(days=1)
        if not hoy_m.any() or not ayer_m.any():
            continue
        horas_norte = paso[hoy_m & norte.values].sum()
        dp = df["pressure_msl"][hoy_m].max() - df["pressure_msl"][ayer_m].mean()
        baja_t = df["temperature_2m"][ayer_m].min() - df["temperature_2m"][hoy_m].min()
        if horas_norte >= al.FF_HORAS_NORTE and (dp >= al.FF_SUBIDA_PRESION or baja_t >= al.FF_BAJADA_TEMP):
            res.add(d)
    return res


def probabilidad_frente(puntos, hoy):
    """Probabilidad (%) total y por día = % de escenarios del ensamble con señales de frente frío."""
    lat = float(np.mean([p[0] for p in puntos]))
    lon = float(np.mean([p[1] for p in puntos]))
    modelo, dfs = descargar_ensamble(lat, lon)
    if not dfs:
        return None
    eventos = [dias_con_frente(df, hoy) for df in dfs.values()]
    n = len(eventos)
    dias = [hoy + pd.Timedelta(days=k) for k in range(al.FF_DIAS_ADELANTE + 1)]
    # Acumulada: % de escenarios en que el frente ya llegó a más tardar ese día (sube conforme se acerca)
    primera = [min(e) if e else None for e in eventos]
    por_dia = {d: 100 * sum(1 for x in primera if x is not None and x <= d) / n for d in dias}
    total = 100 * sum(x is not None for x in primera) / n
    return {"modelo": modelo, "n": n, "total": total, "por_dia": por_dia}


def texto_probabilidad(prob, mostrada):
    if not prob:
        return "📈 <b>Probabilidad</b>: no se pudo calcular hoy (ensamble no disponible)"
    nombre = "ECMWF" if prob["modelo"].startswith("ecmwf") else "GFS"
    l = [f"📈 <b>Probabilidad de frente frío en Sipacate</b> (próximos {al.FF_DIAS_ADELANTE} días): "
         f"<b>{mostrada:.0f} %</b> · {100 - mostrada:.0f} % de que no se presente"]
    l.append("   Acumulada (a más tardar ese día): " + " · ".join(f"{al.fecha_txt(d)[:3]} {d:%d}: {p:.0f} %"
                                          for d, p in prob["por_dia"].items()))
    l.append(f"   <i>% de los {prob['n']} escenarios del ensamble {nombre} que muestran viento del norte "
             "en Sipacate con subida de presión o baja de temperatura.</i>")
    return "\n".join(l)


def grafica_probabilidad(prob, mostrada, ruta):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    dias = list(prob["por_dia"])
    vals = [prob["por_dia"][d] for d in dias]
    etiquetas = [f"{al.fecha_txt(d)[:3]}\n{d:%d-%b}" for d in dias]
    colores = ["#2e7d32" if v < 10 else "#f9a825" if v < 40 else "#ef6c00" if v < 70 else "#c62828" for v in vals]
    fig, ax = plt.subplots(figsize=(8, 3.8))
    b = ax.bar(etiquetas, vals, color=colores)
    for r_, v in zip(b, vals):
        ax.text(r_.get_x() + r_.get_width() / 2, v + 2, f"{v:.0f} %", ha="center", fontsize=10, fontweight="bold")
    ax.set_ylim(0, 110)
    ax.set_ylabel("Probabilidad acumulada (%)")
    ax.axhline(P_CONFIRMADO, color="#c62828", lw=.8, ls="--")
    ax.text(-.45, P_CONFIRMADO + 2, f"umbral de confirmación ({P_CONFIRMADO} %)", ha="left", fontsize=8,
            color="#c62828", bbox=dict(facecolor="white", edgecolor="none", pad=1))
    ax.set_title(f"Probabilidad de frente frío en Sipacate: {mostrada:.0f} % sí · {100 - mostrada:.0f} % no\n"
                 "(probabilidad de que haya llegado a más tardar cada día)", fontsize=10)
    ax.grid(axis="y", alpha=.3)
    fig.tight_layout()
    fig.savefig(ruta, dpi=110)
    plt.close(fig)


# -------------------------------------------------------------------------
# 2. NOAA: texto de la discusión del trópico
# -------------------------------------------------------------------------
def leer_twd(urls):
    import requests
    for u in urls:
        try:
            r = requests.get(u, headers=UA, timeout=30)
            if not r.ok or len(r.text) < 500:
                continue
            txt = r.text
            if "<pre" in txt.lower():                       # página HTML del NHC
                m = re.search(r"<pre[^>]*>(.*?)</pre>", txt, re.S | re.I)
                txt = html.unescape(re.sub(r"<[^>]+>", "", m.group(1) if m else txt))
            return txt, u
        except Exception as e:
            print(f"⚠ No se pudo leer {u}: {e!r}")
    return None, None


def frases_frente(txt):
    """Frases que hablan de frentes fríos (o vientos del norte/Tehuantepec) en la región."""
    frases = []
    for parrafo in re.split(r"\n\s*\n", txt):              # el encabezado queda en su propio párrafo
        plano = re.sub(r"\s+", " ", parrafo).strip()
        frases += re.split(r"(?<=[.!?])\s+", plano)
    res = []
    for f in frases:
        fl = f.lower()
        frente = "cold front" in fl or "stationary front" in fl or "tehuantepec" in fl
        if "weather discussion" in fl or "hurricane center" in fl:
            continue                                         # encabezado del boletín
        if frente and any(z in fl for z in ZONAS) and 20 < len(f) < 400:
            if f not in res:
                res.append(f.strip())
    return res


def emision_twd(txt):
    for linea in txt.splitlines()[:15]:
        if re.search(r"\b(AM|PM|UTC)\b.*\b20\d\d\b", linea):
            return linea.strip()
    return ""


# Términos de la NOAA que la traducción automática suele dejar mal
GLOSARIO = [
    (r"Golfo de Am[eé]rica", "Golfo de México"),
    (r"vientos? de (brecha|hueco|espacio)", "vientos de paso (gap winds)"),
    (r"límite frontal estancado", "frente estancado"),
    (r"\bfrente frio\b", "frente frío"),
    (r"(la |el )?(mendiga|mango|sartén|asa|panhandle) de Florida", "el norte de Florida"),
    (r"Golfo del (Noroeste|Noreste|Suroeste|Sureste)", lambda m: m.group(1).lower() + " del Golfo"),
    (r"\bNE de México", "noreste de México"),
]


def traducir(frase):
    """Traducción automática completa (Google y, si falla, MyMemory). Devuelve None si no hay servicio."""
    try:
        from deep_translator import GoogleTranslator, MyMemoryTranslator
    except Exception:
        return None
    for nombre, tr in [("Google", lambda: GoogleTranslator(source="en", target="es")),
                       ("MyMemory", lambda: MyMemoryTranslator(source="en-US", target="es-ES"))]:
        try:
            t = tr().translate(frase)
            if t and t.strip().lower() != frase.strip().lower():
                for patron, rep in GLOSARIO:
                    t = re.sub(patron, rep, t, flags=re.I)
                return t
        except Exception as e:
            print(f"⚠ Traducción {nombre} no disponible: {e!r}")
    return None


# Resumen en español hecho por el script (no depende de servicios externos)
_FENOMENOS = [("tehuantepec", "Vientos del norte en el Golfo de Tehuantepec"),
              ("cold front", "Frente frío"), ("stationary front", "Frente estacionario"),
              ("stalled front", "Frente estancado"), ("frontal boundary", "Frente")]
_LUGARES = [("bay of campeche", "Bahía de Campeche"), ("yucatan", "Península de Yucatán"),
            ("sw gulf", "suroeste del Golfo de México"), ("southwest gulf", "suroeste del Golfo de México"),
            ("se gulf", "sureste del Golfo de México"), ("nw gulf", "noroeste del Golfo de México"),
            ("ne gulf", "noreste del Golfo de México"), ("northeastern gulf", "noreste del Golfo de México"),
            ("gulf of tehuantepec", None), ("gulf", "Golfo de México"),
            ("central america", "Centroamérica"), ("guatemala", "Guatemala"), ("honduras", "Honduras"),
            ("belize", "Belice"), ("caribbean", "Caribe"), ("mexico", "México"), ("florida", "Florida"),
            ("texas", "Texas"), ("louisiana", "Luisiana")]
_MOVIMIENTO = [("southeastward", "avanzando hacia el sureste"), ("southward", "avanzando hacia el sur"),
               ("move se", "avanzando hacia el sureste"), ("move s ", "avanzando hacia el sur"),
               ("push south", "avanzando hacia el sur"), ("stall", "casi sin moverse"),
               ("dissipate", "debilitándose"), ("weaken", "debilitándose")]
_FUERZA = [("gale", "con vientos de temporal (más de 60 km/h)"),
           ("fresh to strong", "con vientos moderados a fuertes"), ("strong", "con vientos fuertes"),
           ("fresh", "con vientos moderados a fuertes")]
_CUANDO = [("tonight", "esta noche"), ("this weekend", "este fin de semana"),
           ("early next week", "hasta inicios de la próxima semana"), ("next week", "la próxima semana"),
           ("monday", "lunes"), ("tuesday", "martes"), ("wednesday", "miércoles"), ("thursday", "jueves"),
           ("friday", "viernes"), ("saturday", "sábado"), ("sunday", "domingo"),
           ("tomorrow", "mañana"), ("today", "hoy")]


def resumen_es(frase):
    fl = " " + frase.lower() + " "
    fen = next((es for en, es in _FENOMENOS if en in fl), "Frente")
    lugares = []
    for en, es in _LUGARES:
        if en in fl:
            if es and es not in lugares and not any(es in x for x in lugares):
                lugares.append(es)
            fl = fl.replace(en, " ")
    fl2 = " " + frase.lower() + " "
    mov = next((es for en, es in _MOVIMIENTO if en in fl2), "")
    fuerza = next((es for en, es in _FUERZA if en in fl2), "")
    cuando = []
    for en, es in _CUANDO:
        if re.search(r"\b" + en + r"\b", fl2) and es not in cuando:
            cuando.append(es)
            fl2 = re.sub(r"\b" + en + r"\b", " ", fl2)
    lugares = [("el " + x) if x.split()[0] in ("noreste", "noroeste", "sureste", "suroeste") else x
               for x in lugares]
    txt = fen
    if lugares and "Tehuantepec" not in fen:
        txt += " en " + ", ".join(lugares[:3])
    elif lugares:
        txt += " (por un frente en " + ", ".join(lugares[:2]) + ")"
    for x in (mov, fuerza):
        if x:
            txt += ", " + x
    if cuando:
        txt += " — " + ", ".join(cuando[:3])
    return txt


def texto_noaa():
    l = ["2️⃣ <b>NOAA – Discusión del Trópico</b> (meteorólogos del Centro Nacional de Huracanes). "
         "Resumen en español; toca el recuadro para ver la traducción y el texto original."]
    hay = cerca = False
    for nombre, urls in URLS_TWD.items():
        txt, u = leer_twd(urls)
        if not txt:
            l.append(f"• <b>{nombre}</b>: no se pudo leer hoy")
            continue
        fr = frases_frente(txt)
        em = emision_twd(txt)
        l.append(f"• <b>{nombre}</b>" + (f" ({html.escape(em)})" if em else "") + ":")
        if fr:
            hay = True
            cerca = cerca or any(z in f.lower() for f in fr for z in ZONAS_CERCA)
            for f in fr[:4]:
                l.append(f"   ▸ <b>{html.escape(resumen_es(f))}</b>")
                es = traducir(f)
                detalle = (f"Traducción: {html.escape(es)}\n\n" if es else "") + f"Original: {html.escape(f)}"
                l.append(f"<blockquote expandable>{detalle}</blockquote>")
        else:
            l.append("   sin menciones de frentes fríos cerca de la región")
    if cerca:
        l.append("→ <b>La NOAA menciona un frente frío cerca de Guatemala</b> (sur del Golfo, Yucatán, "
                 "Tehuantepec o Centroamérica)")
    elif hay:
        l.append("→ La NOAA menciona frentes en el Golfo de México, todavía lejos de Guatemala")
    else:
        l.append("→ La NOAA no menciona frentes fríos en la región")
    return "\n".join(l), hay, cerca


def descargar_imagen(url, ruta):
    import requests
    from PIL import Image
    try:
        r = requests.get(url, headers=UA, timeout=60)
        if r.ok and len(r.content) > 10000:
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
            if max(img.size) > 2560:                         # Telegram limita el tamaño de las fotos
                img.thumbnail((2560, 2560))
            img.save(ruta, "PNG")
            print(f"✓ Imagen {url} ({img.size[0]}x{img.size[1]})")
            return True
        print(f"⚠ Imagen {url}: respuesta {r.status_code}, {len(r.content)} bytes")
    except Exception as e:
        print(f"⚠ No se pudo descargar {url}: {e!r}")
    return False


def texto_insivumeh():
    """Boletín informativo del INSIVUMEH (PDF): fecha y frases sobre frentes fríos o nortes."""
    import requests
    l = ["4️⃣ <b>INSIVUMEH – Boletín informativo</b>"]
    try:
        r = requests.get(URL_BOLETIN_INSIVUMEH, headers=UA, timeout=60)
        if not r.ok or not r.content.startswith(b"%PDF"):
            l.append(f"• No se pudo leer hoy (respuesta {r.status_code})")
            return "\n".join(l), False
        from pypdf import PdfReader
        txt = "\n".join((pg.extract_text() or "") for pg in PdfReader(io.BytesIO(r.content)).pages)
    except Exception as e:
        print(f"⚠ Boletín INSIVUMEH: {e!r}")
        l.append("• No se pudo leer hoy")
        return "\n".join(l), False
    plano = re.sub(r"\s+", " ", txt)
    m = re.search(r"\b\d{1,2} de [a-záéíóú]+ (de|del) 20\d\d\b", plano, re.I)
    l.append(f"• Fecha del boletín: <b>{m.group(0) if m else 'no indicada'}</b>")
    frases = [f.strip() for f in re.split(r"(?<=[.!?])\s+", plano)
              if any(c in f.lower() for c in CLAVES_INSIVUMEH) and 20 < len(f) < 400]
    if frases:
        for f in frases[:4]:
            l.append(f"   « {html.escape(f)} »")
    else:
        l.append("   sin menciones de frentes fríos o vientos del norte")
    print(f"Boletín INSIVUMEH: {len(txt)} caracteres de texto, {len(frases)} frases relevantes")
    return "\n".join(l), bool(frases)


def texto_pronostico_insivumeh():
    """Pronóstico diario oficial del INSIVUMEH (PDF con texto): número, fecha,
    menciones de frentes fríos / nortes y la región Pacífico (Escuintla)."""
    import time
    import requests
    l = ["🇬🇹 <b>INSIVUMEH – Pronóstico diario oficial</b>"]
    try:
        # La carpeta podría cambiar con el mes (Papa_Noviembre, Papa_Diciembre...): se prueba primero la del mes
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
                 "Octubre", "Noviembre", "Diciembre"]
        mes = meses[pd.Timestamp.now(tz=TZ).month - 1]
        urls = [URL_PRONOSTICO_INSIVUMEH.replace("Papa_Noviembre", f"Papa_{mes}"), URL_PRONOSTICO_INSIVUMEH]
        for u in dict.fromkeys(urls):
            r = requests.get(f"{u}?t={int(time.time() * 1000)}", headers=UA, timeout=60)
            if r.ok and r.content.startswith(b"%PDF"):
                break
        if not r.ok or not r.content.startswith(b"%PDF"):
            l.append(f"• ⚠️ No se pudo leer hoy (respuesta {r.status_code}). ¿Cambió el enlace?")
            return "\n".join(l), False
        print(f"Pronóstico INSIVUMEH: {r.url.split('?')[0]} · Last-Modified: {r.headers.get('Last-Modified')}")
        PDF_INSIVUMEH["bytes"] = r.content
        from pypdf import PdfReader
        txt = "\n".join((pg.extract_text() or "") for pg in PdfReader(io.BytesIO(r.content)).pages)
    except Exception as e:
        print(f"⚠ Pronóstico INSIVUMEH: {e!r}")
        l.append("• ⚠️ No se pudo leer hoy")
        return "\n".join(l), False
    if os.environ.get("SOLO_MOSTRAR") == "1":
        print("----- TEXTO PDF INSIVUMEH -----\n" + txt[:4000] + "\n----- FIN -----")
    plano = re.sub(r"\s+", " ", txt)
    num = re.search(r"\bBD\s?\d+-20\d\d\b", plano)
    fecha = re.search(r"\b\d{1,2} de [a-záéíóú]+ (de|del) 20\d\d\b", plano, re.I)
    l.append("• " + " · ".join(x for x in [num.group(0) if num else "",
                                           fecha.group(0).lower() if fecha else ""] if x))
    # ¿Está al día? El boletín de mañana se publica la tarde anterior: debe ser de hoy o de mañana
    MESES = {m: i for i, m in enumerate(["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
                                         "septiembre", "octubre", "noviembre", "diciembre"], 1)}
    if fecha:
        d, mes, a = re.match(r"(\d{1,2}) de ([a-záéíóú]+) (?:de|del) (\d{4})", fecha.group(0), re.I).groups()
        f_bol = pd.Timestamp(int(a), MESES.get(mes.lower(), 1), int(d), tz=TZ)
        if f_bol < pd.Timestamp.now(tz=TZ).normalize():
            l.append(f"• ⚠️ <b>El boletín es del {fecha.group(0).lower()}: no se ha actualizado.</b> "
                     "Revisar si el INSIVUMEH cambió el enlace.")
    PDF_INSIVUMEH["titulo"] = " · ".join(x for x in [num.group(0) if num else "",
                                                    fecha.group(0).lower() if fecha else ""] if x)
    persp = re.search(r"Condiciones y perspectiva:\s*(.+?[.])", plano)
    if persp:
        l.append(f"• Perspectiva: <i>{html.escape(persp.group(1))}</i>")
    frases = [f.strip() for f in re.split(r"(?<=[.!?])\s+", plano)
              if any(c in f.lower() for c in CLAVES_INSIVUMEH + ["norte moderado", "norte fuerte"])
              and 15 < len(f) < 400]
    if frases:
        for f in frases[:4]:
            l.append(f"   « {html.escape(f)} »")
    else:
        l.append("   sin menciones de frentes fríos, empujes de aire frío o vientos del norte fuertes")
    norte_costa = False
    for lugar in ESTACIONES_INSIVUMEH:
        d = fila_estacion(plano, lugar)
        if d:
            viento = d[2].lower()
            aviso = ""
            if re.match(r"(norte|noreste|noroeste|nor)", viento):
                norte_costa = True
                aviso = " ⚠️ <b>viento del norte (hacia el mar)</b>"
            l.append(f"   📍 <b>{lugar}</b>: mín {d[0]} °C · máx {d[1]} °C · viento {viento}{aviso}")
    return "\n".join(l), bool(frases) or norte_costa


ESTACIONES_INSIVUMEH = ["Puerto de San José", "Sabana Grande"]     # las más cercanas a Sipacate


def fila_estacion(plano, lugar):
    """De la tabla del PDF: 'Puerto de San José - Escuintla 20 a 22 31 a 33 Suroeste ligero'."""
    patron = (re.escape(lugar).replace("é", "[eé]") +
              r"\s*-\s*Escuintla\s+(\d+ a \d+)\s+(\d+ a \d+)\s+(.+?)"
              r"(?=\s+[A-ZÁÉÍÓÚ][a-záéíóú]+(?: [A-Za-záéíóú]+)*\s*-|\s+Región|$)")
    m = re.search(patron, plano)
    if not m:
        return None
    return m.group(1).replace(" a ", "–"), m.group(2).replace(" a ", "–"), m.group(3).strip()


# -------------------------------------------------------------------------
# 3. Envío solo a chats personales
# -------------------------------------------------------------------------
def enviar_validacion(textos, fotos):
    import requests
    token = al._token()
    chats = [c for c in re.findall(r"-?\d+", os.environ.get("TELEGRAM_CHAT_ID", "")) if not c.startswith("-")]
    if os.environ.get("SOLO_MOSTRAR") == "1" or not token or not chats:
        print("No se envía (SOLO_MOSTRAR o sin token/chat personal).")
        return
    url = f"https://api.telegram.org/bot{token}"
    for chat in chats:
        for t in textos:
            r = requests.post(f"{url}/sendMessage", data={"chat_id": chat, "text": t, "parse_mode": "HTML",
                                                          "disable_web_page_preview": "true"}, timeout=30)
            print(f"{'✓' if r.ok else '✗'} texto a {chat} {'' if r.ok else r.text}")
        for ruta, cap in fotos:
            with open(ruta, "rb") as f:
                r = requests.post(f"{url}/sendPhoto", data={"chat_id": chat, "caption": cap},
                                  files={"photo": f}, timeout=60)
            print(f"{'✓' if r.ok else '✗'} imagen a {chat} {'' if r.ok else r.text}")


def detectar_para_emitir(area, ahora):
    """Si el usuario decide emitir y el modelo no marcó el frente con los criterios normales,
    se buscan las condiciones con criterios más flexibles solo para llenar los datos de la alerta."""
    ev = al.detectar_frente_frio(area, ahora)
    if ev:
        return ev
    orig = (al.FF_SUBIDA_PRESION, al.FF_BAJADA_TEMP, al.FF_HORAS_SEGUIDAS, al.FF_HORAS_NORTE, al.FF_VIENTO_MIN)
    try:
        al.FF_SUBIDA_PRESION, al.FF_BAJADA_TEMP = -1e9, -1e9
        al.FF_HORAS_SEGUIDAS, al.FF_HORAS_NORTE, al.FF_VIENTO_MIN = 2, 3, 8.0
        return al.detectar_frente_frio(area, ahora)
    finally:
        (al.FF_SUBIDA_PRESION, al.FF_BAJADA_TEMP, al.FF_HORAS_SEGUIDAS,
         al.FF_HORAS_NORTE, al.FF_VIENTO_MIN) = orig


def emitir_al_grupo(area, ahora):
    """SOLO cuando el usuario lo autoriza (EMITIR_GRUPO=1): envía la alerta de frente frío al grupo,
    a los chats del secret y a los suscriptores, junto con el pronóstico diario del INSIVUMEH."""
    import requests
    ev = detectar_para_emitir(area, ahora)
    token = al._token()
    fijos = re.findall(r"-?\d+", os.environ.get("TELEGRAM_CHAT_ID", ""))
    personales = [c for c in fijos if not c.startswith("-")]
    url = f"https://api.telegram.org/bot{token}"
    solo_mostrar = os.environ.get("SOLO_MOSTRAR") == "1"

    if not ev:
        aviso = ("⚠️ <b>No se emitió la alerta</b>: el modelo no ve viento del norte en Sipacate en los "
                 "próximos días, así que no hay datos para llenar el mensaje. Revisemos juntos antes de emitir.")
        print(aviso)
        if not solo_mostrar:
            for c in personales:
                requests.post(f"{url}/sendMessage", data={"chat_id": c, "text": aviso, "parse_mode": "HTML"}, timeout=30)
        raise SystemExit(1)

    alerta = al.mensaje_frente_frio(ev, ahora)
    texto_pronostico_insivumeh()                       # descarga el PDF del día
    print("ALERTA A EMITIR:\n" + alerta)
    if solo_mostrar:
        print("SOLO_MOSTRAR=1: vista previa, no se envía nada.")
        return

    # Destinos: grupo + chats del secret + suscriptores registrados (lista cifrada guardada en main)
    destinos = list(dict.fromkeys(fijos))
    try:
        r = requests.get("https://raw.githubusercontent.com/ferescobedo369-max/PronosticoSipacate/main/"
                         "3_Suscriptores/suscriptores.enc", timeout=30)
        if r.ok:
            os.makedirs(os.path.dirname(al.ARCHIVO_SUSCRIPTORES), exist_ok=True)
            with open(al.ARCHIVO_SUSCRIPTORES, "wb") as f:
                f.write(r.content)
        subs = al.cargar_suscriptores(token)
        api = al._api(token)
        grupo = next((c for c in fijos if c.startswith("-")), None)
        for cid in sorted(subs["personas"]):
            if cid in destinos:
                continue
            if grupo:
                m = api("getChatMember", chat_id=grupo, user_id=cid)
                if m.get("ok") and m["result"].get("status") not in al.ESTADOS_MIEMBRO:
                    continue
            destinos.append(cid)
    except Exception as e:
        print(f"⚠ Suscriptores no disponibles: {e!r}")

    enviados = 0
    for c in destinos:
        r = requests.post(f"{url}/sendMessage", data={"chat_id": c, "text": alerta, "parse_mode": "HTML",
                                                      "disable_web_page_preview": "true"}, timeout=30)
        ok = r.ok
        if ok and PDF_INSIVUMEH.get("bytes"):
            cap = "📄 Pronóstico diario del INSIVUMEH" + (f" – {PDF_INSIVUMEH.get('titulo')}" if PDF_INSIVUMEH.get("titulo") else "")
            requests.post(f"{url}/sendDocument", data={"chat_id": c, "caption": cap},
                          files={"document": ("Pronostico_diario_INSIVUMEH.pdf", PDF_INSIVUMEH["bytes"],
                                              "application/pdf")}, timeout=60)
        enviados += ok
        print(f"{'✓' if ok else '✗'} alerta a {c}")
    for c in personales:
        requests.post(f"{url}/sendMessage", data={"chat_id": c, "parse_mode": "HTML",
                      "text": f"✅ <b>Alerta de frente frío emitida</b> a {enviados} destinos (grupo y chats personales)."},
                      timeout=30)


def main():
    ahora = pd.Timestamp.now(tz=TZ)
    hoy = ahora.normalize()
    puntos = al.puntos_del_area(os.path.join(BASE, "1_Shapefile", "Area_of_interest.shp"))
    area = al.resumir_area(al.descargar(puntos))

    if os.environ.get("EMITIR_GRUPO") == "1":
        emitir_al_grupo(area, ahora)
        return

    ev = al.detectar_frente_frio(area, ahora)
    t_ecmwf = texto_ecmwf(resumen_diario(area, hoy), ev)
    t_noaa, noaa_si, noaa_cerca = texto_noaa()
    t_pron, insiv_si = texto_pronostico_insivumeh()

    prob = probabilidad_frente(puntos, hoy)
    p_ens = prob["total"] if prob else 0.0

    # ¿Confirmado? Ensamble alto + modelo principal + fuente oficial + llega pronto
    pronto = bool(ev) and (ev["en_curso"] or (ev["inicio"] - ahora) <= pd.Timedelta(hours=HORAS_CONFIRMADO))
    confirmado = bool(ev) and pronto and p_ens >= P_CONFIRMADO and (insiv_si or noaa_cerca)
    mostrada = 100.0 if confirmado else min(95.0, float(round(p_ens)))
    if mostrada == 0 and (ev or noaa_cerca or insiv_si):
        mostrada = 5.0                                      # hay alguna señal, aunque el ensamble no la vea

    senales = {"Ensamble": p_ens >= P_MIN_AVISO, "Modelo ECMWF": bool(ev), "NOAA": noaa_cerca,
               "INSIVUMEH": insiv_si}
    t_prob = texto_probabilidad(prob, mostrada)
    if not any(senales.values()):
        print(f"Sin señales de frente frío (probabilidad {p_ens:.0f} %): no se envía nada hoy.")
        print(t_prob + "\n\n" + t_ecmwf + "\n\n" + t_noaa + "\n\n" + t_pron)
        return

    if confirmado:
        print(f"FRENTE FRÍO CONFIRMADO (ensamble {p_ens:.0f} %): se emite al grupo automáticamente.")
        emitir_al_grupo(area, ahora)

    marcas = " · ".join(f"{n} {'✅' if v else '❌'}" for n, v in senales.items())
    if confirmado:
        cab = (f"🔴 <b>FRENTE FRÍO CONFIRMADO – probabilidad 100 %</b>\n"
               f"<i>La alerta ya se envió al grupo automáticamente.</i>\nSeñales: {marcas}")
        pie = None
    else:
        nivel = "🟠" if mostrada >= 40 else "🟡"
        cab = (f"{nivel} <b>POSIBLE FRENTE FRÍO – probabilidad {mostrada:.0f} %</b> · favor de darle seguimiento\n"
               f"<i>Solo para ti: no se envió al grupo.</i>\nSeñales: {marcas}")
        pie = ("👉 Se enviará al grupo automáticamente si llega al 100 % (ensamble ≥ "
               f"{P_CONFIRMADO} %, el modelo lo detecta en las próximas {HORAS_CONFIRMADO} h y lo confirma la NOAA "
               "o el INSIVUMEH). Si quieres emitirlo antes, escríbeme en Claude <b>«emite la alerta de frente frío»</b>.")
    resumen = "\n\n".join([x for x in [cab, t_prob, t_ecmwf, t_noaa, t_pron, pie,
                            "🗺️ Abajo: gráfica de probabilidad, gráfica del modelo, mapa de la NOAA "
                            "(frente frío = línea azul con triángulos) e imagen satelital del INSIVUMEH."] if x])
    textos = [resumen]
    if ev and not confirmado:
        textos.append("👇 <b>Vista previa de la alerta para el grupo</b> (no enviada):\n\n"
                      + al.mensaje_frente_frio(ev, ahora))

    fotos = []
    if prob:
        gp = os.path.join(BASE, "_grafica_probabilidad.png")
        grafica_probabilidad(prob, mostrada, gp)
        fotos.append((gp, f"Probabilidad de frente frío en Sipacate: {mostrada:.0f} % sí · {100 - mostrada:.0f} % no"))
    g = os.path.join(BASE, "_grafica_validacion.png")
    grafica(area, ev, g)
    fotos.append((g, "Modelo ECMWF – Sipacate: presión, temperatura, humedad y viento"))
    m = os.path.join(BASE, "_mapa_noaa.png")
    if descargar_imagen(URL_MAPA_NOAA, m):
        fotos.append((m, "NOAA – Análisis de superficie más reciente (frente frío = línea azul con triángulos)"))
    sat = os.path.join(BASE, "_satelite_insivumeh.png")
    if descargar_imagen(URL_SAT_INSIVUMEH, sat):
        fotos.append((sat, "INSIVUMEH – GOES-19 color verdadero (Guatemala). Un frente suele verse como "
                           "una banda larga de nubes que baja desde el Golfo de México"))

    for t in textos:
        print(t + "\n" + "-" * 60)
    enviar_validacion(textos, fotos)


if __name__ == "__main__":
    main()
