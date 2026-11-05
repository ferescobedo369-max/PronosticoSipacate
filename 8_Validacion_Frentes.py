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
URLS_TWD = {
    "Atlántico / Golfo": ["https://tgftp.nws.noaa.gov/data/raw/ax/axnt20.knhc.twd.at.txt",
                          "https://www.nhc.noaa.gov/text/MIATWDAT.shtml"],
    "Pacífico Este": ["https://tgftp.nws.noaa.gov/data/raw/ax/axpz20.knhc.twd.ep.txt",
                      "https://www.nhc.noaa.gov/text/MIATWDEP.shtml"],
}
# Zonas que importan para Guatemala (la NOAA ahora usa "Gulf of America" para el Golfo de México)
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


def texto_noaa():
    l = ["2️⃣ <b>NOAA – Tropical Weather Discussion</b> (texto de meteorólogos, en inglés)"]
    hay = False
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
            for f in fr[:4]:
                l.append(f"   « {html.escape(f)} »")
        else:
            l.append("   sin menciones de frentes fríos cerca de la región")
    l.append("→ " + ("<b>La NOAA menciona frentes fríos cerca de la región</b>" if hay
                     else "La NOAA no menciona frentes fríos cerca de la región"))
    return "\n".join(l), hay


def mapa_noaa(ruta):
    import requests
    from PIL import Image
    try:
        r = requests.get(URL_MAPA_NOAA, headers=UA, timeout=60)
        if r.ok and len(r.content) > 10000:
            Image.open(io.BytesIO(r.content)).convert("RGB").save(ruta, "PNG")
            return True
        print(f"⚠ Mapa NOAA: respuesta {r.status_code}, {len(r.content)} bytes")
    except Exception as e:
        print(f"⚠ No se pudo descargar el mapa NOAA: {e!r}")
    return False


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


def main():
    ahora = pd.Timestamp.now(tz=TZ)
    hoy = ahora.normalize()
    puntos = al.puntos_del_area(os.path.join(BASE, "1_Shapefile", "Area_of_interest.shp"))
    area = al.resumir_area(al.descargar(puntos))
    ev = al.detectar_frente_frio(area, ahora)

    t_ecmwf = texto_ecmwf(resumen_diario(area, hoy), ev)
    t_noaa, noaa_si = texto_noaa()

    if ev and noaa_si:
        concl = "✅ <b>Coinciden</b>: el modelo y la NOAA indican frente frío."
    elif ev:
        concl = "⚠️ <b>Solo el modelo</b> detecta frente. Revisar el mapa: ¿posible falsa alarma o llega después?"
    elif noaa_si:
        concl = "⚠️ <b>Solo la NOAA</b> menciona un frente. ¿Llegará a Sipacate? Revisar si los criterios son muy estrictos."
    else:
        concl = "✅ Ninguna fuente indica frente frío cerca."

    cab = (f"🧪 <b>VALIDACIÓN FRENTES FRÍOS</b> – {al.fecha_txt(ahora)} {ahora:%H:%M}\n"
           "<i>Solo para revisión: esto NO se envió al grupo.</i>")
    resumen = "\n\n".join([cab, t_ecmwf, t_noaa, "3️⃣ " + concl,
                           "🗺️ Abajo: gráfica del modelo y mapa de superficie de la NOAA "
                           "(frente frío = línea azul con triángulos)."])
    textos = [resumen]
    if ev:
        textos.append("👇 <b>Vista previa de la alerta para el grupo</b> (no enviada):\n\n"
                      + al.mensaje_frente_frio(ev, ahora))

    fotos = []
    g = os.path.join(BASE, "_grafica_validacion.png")
    grafica(area, ev, g)
    fotos.append((g, "Modelo ECMWF – Sipacate: presión, temperatura, humedad y viento"))
    m = os.path.join(BASE, "_mapa_noaa.png")
    if mapa_noaa(m):
        fotos.append((m, "NOAA – Análisis de superficie más reciente (frente frío = línea azul con triángulos)"))

    for t in textos:
        print(t + "\n" + "-" * 60)
    enviar_validacion(textos, fotos)


if __name__ == "__main__":
    main()
