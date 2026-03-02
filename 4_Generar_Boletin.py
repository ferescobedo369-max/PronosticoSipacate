"""
Script: Generador Automático de Boletín de Viento - Sipacate
- Lee Area_forecast_latest.csv generado por Script 1
- Genera análisis automático de dirección y velocidad
- Produce Boletin_Sipacate_BODDMMYY.docx usando la plantilla
- Compatible con GitHub Actions (sin rutas locales)
"""

import os
import shutil
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, timedelta

# ---------------------------------------------------------
# Rutas (relativas al repo en GitHub Actions)
# ---------------------------------------------------------
base_dir    = os.path.dirname(os.path.abspath(__file__))
results_dir = os.path.join(base_dir, "2_Results")
plantilla   = os.path.join(base_dir, "0_Archivos_Recibidos", "plantilla_de_boletin.docx")
csv_path    = os.path.join(results_dir, "Area_forecast_latest.csv")

# ---------------------------------------------------------
# Helpers de fechas en español
# ---------------------------------------------------------
DIAS_ES = {
    'Monday':'lunes','Tuesday':'martes','Wednesday':'miércoles',
    'Thursday':'jueves','Friday':'viernes','Saturday':'sábado','Sunday':'domingo'
}
MESES_ES = {
    'January':'enero','February':'febrero','March':'marzo','April':'abril',
    'May':'mayo','June':'junio','July':'julio','August':'agosto',
    'September':'septiembre','October':'octubre','November':'noviembre','December':'diciembre'
}

def dia_es(fecha):
    return DIAS_ES.get(fecha.strftime('%A'), fecha.strftime('%A'))

def mes_es(fecha):
    return MESES_ES.get(fecha.strftime('%B'), fecha.strftime('%B'))

# ---------------------------------------------------------
# Leer CSV
# ---------------------------------------------------------
def leer_datos(csv_path):
    df = pd.read_csv(csv_path)
    df['date'] = pd.to_datetime(df['date'], utc=True).dt.tz_convert('America/Guatemala')
    df['date'] = df['date'].dt.tz_localize(None)  # quitar tz para simplificar comparaciones
    return df

# ---------------------------------------------------------
# Análisis de DIRECCIÓN para un día
# ---------------------------------------------------------
def analizar_direccion(df_dia, dia_nombre, dia_num, mes_nombre):
    """
    Genera texto interpretativo de dirección del viento
    basado en los datos horarios de wind_direction_10m
    """
    dir_10m = df_dia['wind_direction_10m'].values
    horas   = df_dia['Hora'].values if 'Hora' in df_dia.columns else df_dia['date'].dt.hour.values

    def es_favorable(grados):
        return 90 <= grados <= 270

    favorable_mask = np.array([es_favorable(d) for d in dir_10m])
    horas_fav   = int(favorable_mask.sum())
    horas_nofav = int((~favorable_mask).sum())

    # Hora de inicio del bloque favorable principal
    inicio_fav = None
    fin_fav    = None
    max_bloque = 0
    bloque_actual = 0
    inicio_actual = None

    for i, (h, fav) in enumerate(zip(horas, favorable_mask)):
        if fav:
            if bloque_actual == 0:
                inicio_actual = h
            bloque_actual += 1
        else:
            if bloque_actual > max_bloque:
                max_bloque  = bloque_actual
                inicio_fav  = inicio_actual
                fin_fav     = horas[i-1] if i > 0 else h
            bloque_actual = 0

    if bloque_actual > max_bloque:
        max_bloque = bloque_actual
        inicio_fav = inicio_actual
        fin_fav    = horas[-1]

    # Madrugada (0-6h) — cuántas horas no favorables
    madrugada_nofav = sum(1 for h, fav in zip(horas, favorable_mask) if h < 7 and not fav)

    # Tarde (12-19h) — cuántas horas favorables
    tarde_fav = sum(1 for h, fav in zip(horas, favorable_mask) if 12 <= h <= 19 and fav)

    # Noche (20-23h) — cuántas horas no favorables
    noche_nofav = sum(1 for h, fav in zip(horas, favorable_mask) if h >= 20 and not fav)

    # ------- Construir texto -------
    partes = []

    # Comportamiento madrugada / mañana
    if madrugada_nofav >= 4:
        partes.append(
            "Se identifica una alta variabilidad del viento durante la madrugada y las "
            f"primeras horas de la mañana (00:00 - {min(horas[~favorable_mask & (horas < 10)], default=7):02d}:00), "
            "con el viento soplando frecuentemente en la zona no favorable, es decir hacia el sur."
        )
    else:
        partes.append(
            "Durante las primeras horas del día se presentan condiciones variables, "
            "con algunas horas mostrando viento hacia zonas no favorables."
        )

    # Consolidación favorable
    if inicio_fav is not None and max_bloque >= 4:
        partes.append(
            f"A partir de las {inicio_fav:02d}:00 horas, la dirección se consolida dentro de la "
            f"zona favorable (90° a 270°) en ambos niveles de altura (10 y 100 metros), "
            f"manteniendo estabilidad durante la tarde hasta las {min(fin_fav+1, 23):02d}:00 horas aproximadamente."
        )
    elif horas_fav >= 6:
        partes.append(
            "La dirección se mantiene dentro de la zona favorable durante buena parte del día, "
            "especialmente durante las horas de la tarde."
        )
    else:
        partes.append(
            "El viento presenta condiciones predominantemente no favorables durante la mayor "
            "parte del día, con pocas horas dentro de la zona verde."
        )

    # Comportamiento nocturno
    if noche_nofav >= 2:
        hora_cambio = next((h for h, fav in zip(horas, favorable_mask) if h >= 19 and not fav), 20)
        partes.append(
            f"El comportamiento vuelve a mostrar variaciones hacia la zona roja después de "
            f"las {hora_cambio:02d}:00 horas."
        )

    return " ".join(partes)


# ---------------------------------------------------------
# Análisis de VELOCIDAD para un día
# ---------------------------------------------------------
def analizar_velocidad(df_dia, dia_nombre, dia_num, mes_nombre):
    """
    Genera texto interpretativo de velocidad del viento
    """
    v10  = df_dia['wind_speed_10m'].values
    v100 = df_dia['wind_speed_100m'].values
    horas = df_dia['Hora'].values if 'Hora' in df_dia.columns else df_dia['date'].dt.hour.values

    prom_10  = float(np.mean(v10))
    prom_100 = float(np.mean(v100))
    max_10   = float(np.max(v10))
    hora_max = int(horas[np.argmax(v10)])

    # Horas sobre umbral 10 km/h
    sobre_umbral = int(np.sum(v10 >= 10))
    if sobre_umbral > 0:
        horas_sobre = horas[v10 >= 10]
        h_ini_umbral = int(horas_sobre[0])
        h_fin_umbral = int(horas_sobre[-1])
    else:
        h_ini_umbral = h_fin_umbral = 0

    # Velocidad mañana vs tarde
    v_manana = float(np.mean(v10[horas < 10])) if any(horas < 10) else 0
    v_tarde  = float(np.mean(v10[(horas >= 12) & (horas <= 19)])) if any((horas >= 12) & (horas <= 19)) else 0

    # Diferencia entre niveles
    diff_niveles = prom_100 - prom_10

    # ------- Construir texto -------
    partes = []

    # Velocidades promedio
    partes.append(
        f"Las velocidades promedio del viento se registran en {prom_10:.1f} km/h a 10 metros "
        f"de altura y {prom_100:.1f} km/h a 100 metros."
    )

    # Velocidad máxima
    intensidad = "moderado a fuerte" if max_10 >= 20 else "moderado" if max_10 >= 10 else "leve a moderado"
    partes.append(
        f"Las velocidades máximas alcanzan {max_10:.1f} km/h alrededor de las {hora_max:02d}:00 horas "
        f"a 10 metros de altura, indicando condiciones de viento {intensidad} durante este periodo."
    )

    # Horas sobre umbral
    if sobre_umbral > 0:
        partes.append(
            f"Se pronostican velocidades por encima de 10 km/h durante {sobre_umbral} horas del día, "
            f"principalmente entre las {h_ini_umbral:02d}:00 y {h_fin_umbral:02d}:00 horas."
        )
    else:
        partes.append(
            "Las velocidades se mantienen por debajo del umbral de 10 km/h durante todo el día."
        )

    # Incremento mañana → tarde
    if v_tarde > v_manana and v_tarde > 5:
        partes.append(
            f"Se observa un incremento significativo de las velocidades durante la tarde, "
            f"pasando de {v_manana:.1f} km/h en la mañana a {v_tarde:.1f} km/h en la tarde."
        )

    # Diferencia entre niveles
    if diff_niveles > 4:
        partes.append(
            f"Se observa una diferencia notable entre niveles, con velocidades {diff_niveles:.1f} km/h "
            f"más altas a 100 metros de altura, lo cual es característico del perfil vertical del viento."
        )

    return " ".join(partes)


# ---------------------------------------------------------
# FUNCIÓN PRINCIPAL: generar boletín
# ---------------------------------------------------------
def generar_boletin():

    # Verificar archivos necesarios
    for ruta, nombre in [(csv_path, "CSV de pronóstico"), (plantilla, "plantilla .docx")]:
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"No se encontró {nombre} en: {ruta}")

    # Leer datos
    print("📊 Leyendo datos de pronóstico...")
    df = leer_datos(csv_path)
    df['date_only'] = df['date'].dt.date
    df['Hora']      = df['date'].dt.hour

    fechas_unicas = sorted(df['date_only'].unique())
    n_dias = min(3, len(fechas_unicas))
    fechas = [datetime.combine(d, datetime.min.time()) for d in fechas_unicas[:n_dias]]

    if not fechas:
        raise RuntimeError("No hay datos de pronóstico disponibles.")

    fecha_inicio = fechas[0]
    correlativo  = f"BO{fecha_inicio.strftime('%d%m%y')}"
    print(f"📅 Correlativo: {correlativo} | Días: {n_dias}")

    # Abrir plantilla
    print("📄 Cargando plantilla...")
    doc = Document(plantilla)

    # Limpiar el párrafo vacío que trae la plantilla
    for p in doc.paragraphs:
        p.clear()

    def add_para(texto_bold=None, texto_normal=None, size_pt=None, space_after=6):
        p = doc.add_paragraph()
        if texto_bold:
            r = p.add_run(texto_bold)
            r.bold = True
            if size_pt:
                r.font.size = Pt(size_pt)
        if texto_normal:
            r2 = p.add_run(texto_normal)
            r2.bold = False
        p.paragraph_format.space_after  = Pt(space_after)
        p.paragraph_format.space_before = Pt(0)
        return p

    # ---- ENCABEZADO ----
    add_para(texto_bold=f"Correlativo {correlativo}", space_after=4)
    add_para(texto_bold="ICC: CeH y SSP", space_after=4)
    add_para(
        texto_bold="Condiciones de viento proyectadas para las zonas buffer prioritaria",
        size_pt=14, space_after=6
    )

    # Fechas del pronóstico
    dia1, dia2, dia3 = fechas[0], fechas[1] if n_dias>1 else fechas[0], fechas[2] if n_dias>2 else fechas[0]
    if n_dias == 3:
        texto_fechas = (
            f"Pronósticos de horarios indicados para el {dia1.day}, {dia2.day} de {mes_es(dia1)} "
            f"y {dia3.day} de {mes_es(dia3)} {dia1.year}."
        )
    elif n_dias == 2:
        texto_fechas = (
            f"Pronósticos de horarios indicados para el {dia1.day} y {dia2.day} de "
            f"{mes_es(dia1)} {dia1.year}."
        )
    else:
        texto_fechas = f"Pronóstico de horarios indicados para el {dia1.day} de {mes_es(dia1)} {dia1.year}."

    add_para(texto_normal=texto_fechas, space_after=10)

    # ---- INTERPRETACIÓN GENERAL ----
    p = doc.add_paragraph()
    p.add_run("Interpretación: ").bold = True
    p.add_run(
        "Para facilitar el análisis, la gráfica utiliza un sistema de colores: las áreas en rosado "
        "representan condiciones no favorables (no aptas para realizar la quema), mientras que las "
        "franjas en verde indican periodos favorables para la operación.\n"
        "Cada línea de color en el gráfico corresponde a una fecha específica del calendario. "
        "En la línea horizontal, se detallan las horas del día (de 0 a 23:59), donde cada cuadrícula "
        "representa un intervalo de 2 horas."
    )
    p.paragraph_format.space_after = Pt(6)

    p2 = doc.add_paragraph()
    p2.add_run("• Zonas no favorables (franjas rosadas): ").bold = True
    p2.add_run(
        "Representan los horarios donde el viento sopla predominantemente hacia el Sur "
        "(entre los 0° y 90°, y de 270° a 360°)."
    )
    p2.paragraph_format.space_after = Pt(4)

    p3 = doc.add_paragraph()
    p3.add_run("• Zonas favorables (franja verde): ").bold = True
    p3.add_run(
        "Corresponden a los momentos en que el viento mantiene una dirección hacia el Norte "
        "(entre los 90° y 270°). Estas condiciones son las adecuadas para realizar la actividad "
        "de quemas ya que se espera un desplazamiento de la pavesa al norte."
    )
    p3.paragraph_format.space_after = Pt(4)

    p4 = doc.add_paragraph()
    p4.add_run("Líneas de colores: ").bold = True
    p4.add_run("Significan una fecha en el calendario.")
    p4.paragraph_format.space_after = Pt(12)

    # ---- GRÁFICAS ----
    run_date_str = fecha_inicio.strftime("%Y%m%d")

    graficas_config = [
        ("Serie_ciclo_diario_wind_direction_10m",  "Dirección del viento a 10 metros"),
        ("Serie_ciclo_diario_wind_direction_100m", "Dirección del viento a 100 metros"),
        ("Serie_ciclo_diario_wind_speed_10m",      "Velocidad del viento a 10 metros"),
        ("Serie_ciclo_diario_wind_speed_100m",     "Velocidad del viento a 100 metros"),
    ]

    # Buscar imágenes en carpeta de fecha o raíz de results
    carpeta_fecha = os.path.join(results_dir, run_date_str)
    carpeta_buscar = carpeta_fecha if os.path.isdir(carpeta_fecha) else results_dir

    print("\n🖼️  Insertando gráficas...")
    for prefijo, titulo_grafica in graficas_config:
        # Buscar archivo
        ruta_img = None
        for archivo in os.listdir(carpeta_buscar):
            if prefijo in archivo and archivo.endswith(".png"):
                ruta_img = os.path.join(carpeta_buscar, archivo)
                break

        p_tit = doc.add_paragraph()
        r_tit = p_tit.add_run(titulo_grafica)
        r_tit.bold = True
        r_tit.font.size = Pt(11)
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tit.paragraph_format.space_before = Pt(6)
        p_tit.paragraph_format.space_after  = Pt(4)

        if ruta_img:
            doc.add_picture(ruta_img, width=Inches(6.0))
            print(f"   ✅ {titulo_grafica}")
        else:
            p_err = doc.add_paragraph()
            p_err.add_run(f"[Gráfica no disponible: {prefijo}]").italic = True
            print(f"   ⚠️  No encontrada: {prefijo}")

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # ---- ANÁLISIS DE DIRECCIÓN ----
    print("\n📋 Generando análisis de dirección...")
    p_dir_tit = doc.add_paragraph()
    p_dir_tit.add_run("Dirección del viento").bold = True
    p_dir_tit.runs[0].font.size = Pt(13)
    p_dir_tit.paragraph_format.space_before = Pt(12)
    p_dir_tit.paragraph_format.space_after  = Pt(8)

    for fecha in fechas:
        df_dia = df[df['date_only'] == fecha.date()].copy()
        if df_dia.empty:
            continue

        texto = analizar_direccion(df_dia, dia_es(fecha), fecha.day, mes_es(fecha))

        p = doc.add_paragraph()
        p.add_run(f"{dia_es(fecha).capitalize()} {fecha.day} de {mes_es(fecha)}: ").bold = True
        p.add_run(texto)
        p.paragraph_format.space_after = Pt(10)
        print(f"   ✅ Dirección: {dia_es(fecha)} {fecha.day}")

    # ---- ANÁLISIS DE VELOCIDAD ----
    print("\n🌬️  Generando análisis de velocidad...")
    p_vel_tit = doc.add_paragraph()
    p_vel_tit.add_run("Velocidad del viento").bold = True
    p_vel_tit.runs[0].font.size = Pt(13)
    p_vel_tit.paragraph_format.space_before = Pt(12)
    p_vel_tit.paragraph_format.space_after  = Pt(8)

    for fecha in fechas:
        df_dia = df[df['date_only'] == fecha.date()].copy()
        if df_dia.empty:
            continue

        texto = analizar_velocidad(df_dia, dia_es(fecha), fecha.day, mes_es(fecha))

        p = doc.add_paragraph()
        p.add_run(f"{dia_es(fecha).capitalize()} {fecha.day} de {mes_es(fecha)}: ").bold = True
        p.add_run(texto)
        p.paragraph_format.space_after = Pt(10)
        print(f"   ✅ Velocidad: {dia_es(fecha)} {fecha.day}")

    # ---- GUARDAR ----
    os.makedirs(carpeta_fecha, exist_ok=True)
    nombre_archivo = f"Boletin_Sipacate_{correlativo}.docx"
    ruta_salida    = os.path.join(carpeta_fecha, nombre_archivo)
    # También copia latest en raíz de results para fácil acceso
    ruta_latest    = os.path.join(results_dir, "Boletin_latest.docx")

    doc.save(ruta_salida)
    shutil.copy2(ruta_salida, ruta_latest)

    print(f"\n{'='*55}")
    print(f"✅ BOLETÍN GENERADO: {nombre_archivo}")
    print(f"   Carpeta fecha : {ruta_salida}")
    print(f"   Latest        : {ruta_latest}")
    print(f"{'='*55}\n")

    return ruta_salida


# ---- EJECUCIÓN ----
if __name__ == "__main__":
    generar_boletin()
