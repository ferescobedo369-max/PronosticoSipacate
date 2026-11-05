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
                l.append(f"   ▸ <b>{html.escape(resumen_es(f))}</b>")
                es = traducir(f)
                detalle = (f"Traducción: {html.escape(es)}\n\n" if es else "") + f"Original: {html.escape(f)}"
                l.append(f"<blockquote expandable>{detalle}</blockquote>")
        else:
            l.append("   sin menciones de frentes fríos cerca de la región")
    l.append("→ " + ("<b>La NOAA menciona frentes fríos cerca de la región</b>" if hay
                     else "La NOAA no menciona frentes fríos cerca de la región"))
    return "\n".join(l), hay


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
            l.append(f"• No se pudo leer hoy (respuesta {r.status_code})")
            return "\n".join(l), False
        from pypdf import PdfReader
        txt = "\n".join((pg.extract_text() or "") for pg in PdfReader(io.BytesIO(r.content)).pages)
    except Exception as e:
        print(f"⚠ Pronóstico INSIVUMEH: {e!r}")
        l.append("• No se pudo leer hoy")
        return "\n".join(l), False
    if os.environ.get("SOLO_MOSTRAR") == "1":
        print("----- TEXTO PDF INSIVUMEH -----\n" + txt[:4000] + "\n----- FIN -----")
    plano = re.sub(r"\s+", " ", txt)
    num = re.search(r"\bBD\s?\d+-20\d\d\b", plano)
    fecha = re.search(r"\b\d{1,2} de [a-záéíóú]+ (de|del) 20\d\d\b", plano, re.I)
    l.append("• " + " · ".join(x for x in [num.group(0) if num else "",
                                           fecha.group(0).lower() if fecha else ""] if x))
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


def main():
    ahora = pd.Timestamp.now(tz=TZ)
    hoy = ahora.normalize()
    puntos = al.puntos_del_area(os.path.join(BASE, "1_Shapefile", "Area_of_interest.shp"))
    area = al.resumir_area(al.descargar(puntos))
    ev = al.detectar_frente_frio(area, ahora)

    t_ecmwf = texto_ecmwf(resumen_diario(area, hoy), ev)
    t_noaa, noaa_si = texto_noaa()
    t_insiv = texto_insivumeh()[0] if USAR_BOLETIN_INSIVUMEH else None
    t_pron, insiv_si = texto_pronostico_insivumeh()

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
    concl += ("\n🇬🇹 El INSIVUMEH <b>sí</b> menciona frente frío o viento del norte en la costa sur."
              if insiv_si else "\n🇬🇹 El INSIVUMEH no menciona frente frío ni viento del norte en la costa sur.")
    resumen = "\n\n".join([x for x in [cab, t_ecmwf, t_noaa, t_pron, t_insiv, "3️⃣ " + concl,
                           "🗺️ Abajo: gráfica del modelo, mapa de superficie de la NOAA "
                           "(frente frío = línea azul con triángulos) e imagen satelital del INSIVUMEH."] if x])
    textos = [resumen]
    if ev:
        textos.append("👇 <b>Vista previa de la alerta para el grupo</b> (no enviada):\n\n"
                      + al.mensaje_frente_frio(ev, ahora))

    fotos = []
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
    if os.environ.get("SOLO_MOSTRAR") == "1":           # copia liviana de las imágenes en el log, para revisión
        import base64
        from PIL import Image
        for ruta, _ in fotos:
            img = Image.open(ruta).convert("RGB")
            img.thumbnail((1400, 1400))
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=70)
            print(f"IMG_B64 {os.path.basename(ruta)} {base64.b64encode(buf.getvalue()).decode()}")
    enviar_validacion(textos, fotos)


if __name__ == "__main__":
    main()
