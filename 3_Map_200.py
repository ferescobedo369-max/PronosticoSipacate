import os
import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point

import openmeteo_requests
import requests_cache
from retry_requests import retry
import matplotlib.pyplot as plt
from scipy.interpolate import griddata
from scipy.spatial import QhullError

# ---------------------------------------------------------
# Rutas
# ---------------------------------------------------------
base_dir = r"C:\DATA\OneDrive - Asazgua\SSP\Información generada por Fernando\Zafra 25-26\Sipacate\Pronosticos\Wind_forecasts"

shapefile_path = os.path.join(base_dir, r"1_Shapefile", "Area_of_interest.shp")
results_dir = os.path.join(base_dir, r"2_Results")
os.makedirs(results_dir, exist_ok=True)

# ---------------------------------------------------------
# Leer shapefile principal (asegurar WGS84)
# ---------------------------------------------------------
gdf = gpd.read_file(shapefile_path)
if gdf.crs is None or gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

# Polígono principal (área de interés)
try:
    main_poly = gdf.union_all()
except AttributeError:
    main_poly = gdf.unary_union

# ---------------------------------------------------------
# Construir rejilla de puntos de muestreo dentro del polígono
# ---------------------------------------------------------
minx, miny, maxx, maxy = main_poly.bounds

target_res_deg = 0.1  # ~9–10 km
width_lon = maxx - minx
width_lat = maxy - miny

n_lon = max(3, int(width_lon / target_res_deg) + 1)
n_lat = max(3, int(width_lat / target_res_deg) + 1)

lon_vals = np.linspace(minx, maxx, n_lon)
lat_vals = np.linspace(miny, maxy, n_lat)

points = []
for lat in lat_vals:
    for lon in lon_vals:
        p = Point(lon, lat)
        if main_poly.contains(p):
            points.append((lat, lon))

if not points:
    raise RuntimeError(
        "No se encontraron puntos de muestreo dentro del polígono. "
        "Revise la extensión del shapefile o ajuste la rejilla."
    )

print(f"Usando {len(points)} puntos de muestreo dentro del polígono.")

# ---------------------------------------------------------
# Cliente de Open-Meteo
# ---------------------------------------------------------
cache_session = requests_cache.CachedSession('.cache', expire_after=3600)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)

url = "https://api.open-meteo.com/v1/forecast"

# ---------------------------------------------------------
# Variables
# ---------------------------------------------------------
# Velocidades y ráfagas (incluyendo 200 m)
speed_vars = [
    "wind_speed_10m",
    "wind_speed_100m",
    "wind_speed_200m",
    "wind_gusts_10m",
]

# Direcciones (incluyendo 200 m)
dir_vars = [
    "wind_direction_10m",
    "wind_direction_100m",
    "wind_direction_200m",
]

all_vars = speed_vars + dir_vars

# Etiquetas en español
# (primer elemento: título; segundo: etiqueta de eje / barra de color)
var_labels = {
    # Velocidad (km/h)
    "wind_speed_10m":   ("Velocidad del viento a 10 m",  "Velocidad del viento a 10 m (km/h)"),
    "wind_speed_100m":  ("Velocidad del viento a 100 m", "Velocidad del viento a 100 m (km/h)"),
    "wind_speed_200m":  ("Velocidad del viento a 200 m", "Velocidad del viento a 200 m (km/h)"),
    "wind_gusts_10m":   ("Ráfagas de viento a 10 m",     "Ráfagas de viento a 10 m (km/h)"),

    # Dirección (°)
    "wind_direction_10m":  ("Dirección del viento a 10 m",  "Dirección del viento a 10 m (°)"),
    "wind_direction_100m": ("Dirección del viento a 100 m", "Dirección del viento a 100 m (°)"),
    "wind_direction_200m": ("Dirección del viento a 200 m", "Dirección del viento a 200 m (°)"),
}

# ---------------------------------------------------------
# Descargar datos para cada punto de muestreo
# ---------------------------------------------------------
all_dfs = []

for (lat, lon) in points:
    params = {
        "latitude": lat,
        "longitude": lon,
        "hourly": all_vars,
        "models": "ecmwf_ifs",     # Nota: si 200 m no está disponible para este modelo,
                                   # puede ser necesario usar 'auto' o otro modelo.
        "forecast_days": 3,
        "timezone": "America/Guatemala",
        "windspeed_unit": "kmh"
    }

    responses = openmeteo.weather_api(url, params=params)
    response = responses[0]
    hourly = response.Hourly()

    vals = [hourly.Variables(i).ValuesAsNumpy() for i in range(len(all_vars))]

    date_index = pd.date_range(
        start=pd.to_datetime(hourly.Time(),     unit="s", utc=True).tz_convert("America/Guatemala"),
        end=pd.to_datetime(hourly.TimeEnd(),    unit="s", utc=True).tz_convert("America/Guatemala"),
        freq=pd.Timedelta(seconds=hourly.Interval()),
        inclusive="left"
    )

    data = {
        "date": date_index,
        "lat": lat,
        "lon": lon
    }
    for name, arr in zip(all_vars, vals):
        data[name] = arr

    df_point = pd.DataFrame(data)
    all_dfs.append(df_point)

df_all = pd.concat(all_dfs, ignore_index=True)

# ---------------------------------------------------------
# Preparar fechas
# ---------------------------------------------------------
df_all["date_only"] = df_all["date"].dt.date
unique_dates = sorted(df_all["date_only"].unique())
if len(unique_dates) < 3:
    print("Advertencia: hay menos de 3 días de pronóstico disponibles.")
n_days = min(3, len(unique_dates))

run_date_str = pd.Timestamp.now(tz="America/Guatemala").strftime("%Y%m%d")

# ---------------------------------------------------------
# Mapas: gradiente espacial suave + flechas de viento
# (ahora también para velocidad a 200 m)
# ---------------------------------------------------------
fine_res_deg = 0.02
n_lon_fine = max(40, int((maxx - minx) / fine_res_deg) + 1)
n_lat_fine = max(40, int((maxy - miny) / fine_res_deg) + 1)

lon_fine = np.linspace(minx, maxx, n_lon_fine)
lat_fine = np.linspace(miny, maxy, n_lat_fine)
Lon_fine, Lat_fine = np.meshgrid(lon_fine, lat_fine)

# Máscara del polígono en la rejilla fina
mask = np.zeros(Lon_fine.shape, dtype=bool)
for i in range(Lon_fine.shape[0]):
    for j in range(Lon_fine.shape[1]):
        if main_poly.contains(Point(Lon_fine[i, j], Lat_fine[i, j])):
            mask[i, j] = True

for day_idx in range(n_days):
    d = unique_dates[day_idx]
    d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
    df_day = df_all[df_all["date_only"] == d].copy()

    # Media diaria en cada punto de muestreo
    df_day_mean = df_day.groupby(["lat", "lon"])[all_vars].mean().reset_index()

    coords = np.vstack((df_day_mean["lon"].values, df_day_mean["lat"].values)).T

    # --- Mapas para cada variable de velocidad (10, 100, 200 m y ráfagas 10 m) ---
    for var in speed_vars:
        values = df_day_mean[var].values  # km/h

        # Interpolación del campo
        try:
            grid_fine = griddata(
                coords,
                values,
                (Lon_fine, Lat_fine),
                method="linear"
            )
        except QhullError:
            print(f"Interpolación 'linear' falló para {var}, día {d_str}. "
                  f"Se usa 'nearest'.")
            grid_fine = griddata(
                coords,
                values,
                (Lon_fine, Lat_fine),
                method="nearest"
            )

        grid_fine[~mask] = np.nan

        # Dirección para flechas
        if var == "wind_speed_10m":
            dir_col = "wind_direction_10m"
            mag_col = "wind_speed_10m"
        elif var == "wind_speed_100m":
            dir_col = "wind_direction_100m"
            mag_col = "wind_speed_100m"
        elif var == "wind_speed_200m":
            dir_col = "wind_direction_200m"
            mag_col = "wind_speed_200m"
        else:  # wind_gusts_10m
            dir_col = "wind_direction_10m"
            mag_col = "wind_gusts_10m"

        lon_ar = df_day_mean["lon"].values
        lat_ar = df_day_mean["lat"].values
        dir_deg = df_day_mean[dir_col].values
        mag = df_day_mean[mag_col].values  # km/h

        # Convertir dirección meteorológica a u,v (hacia dónde va el viento)
        dir_rad = np.deg2rad(dir_deg)
        u = -mag * np.sin(dir_rad)
        v = -mag * np.cos(dir_rad)

        # Escalado para que las flechas no sean tan largas
        u_plot = u * 0.5
        v_plot = v * 0.5

        # Etiquetas en español
        title_label, cbar_label = var_labels[var]

        # ----------------------------
        # Graficar
        # ----------------------------
        fig, ax = plt.subplots(figsize=(6, 6))

        pcm = ax.pcolormesh(Lon_fine, Lat_fine, grid_fine, shading="auto")
        cbar = plt.colorbar(pcm, ax=ax, label=cbar_label, fraction=0.035, pad=0.02)

        # Polígono principal
        gdf.boundary.plot(ax=ax, color="black", linewidth=1)

        # Flechas de viento
        ax.quiver(
            lon_ar,
            lat_ar,
            u_plot,
            v_plot,
            angles="xy",
            scale_units="xy",
            scale=120,
            width=0.0016,
            alpha=0.9,
            color="k",
            headwidth=12,
            headlength=20,
            headaxislength=14
        )

        ax.set_title(f"{title_label} – {d_str}")
        ax.set_xlabel("Longitud (°)")
        ax.set_ylabel("Latitud (°)")

        # Menos ticks en el eje de longitud (cada 0.1°)
        xticks = np.arange(
            np.floor(minx * 10) / 10,
            np.ceil(maxx * 10) / 10 + 0.0001,
            0.1
        )
        ax.set_xticks(xticks)

        out_png = os.path.join(
            results_dir,
            f"Mapa_{var}_{d_str.replace('-', '')}_{run_date_str}.png"
        )
        plt.tight_layout()
        plt.savefig(out_png, dpi=150)
        plt.close(fig)

        print(f"Guardado mapa (gradiente + flechas) para {var}, fecha {d_str}: {out_png}")

# ---------------------------------------------------------
# Series de tiempo – ciclo diario (0–23 h), promedio del polígono
# ---------------------------------------------------------

# 1) Velocidades y ráfagas (10, 100, 200 m + ráfagas 10 m)
df_mean_time = df_all.groupby("date")[speed_vars].mean().reset_index()

for var in speed_vars:
    title_label, y_label = var_labels[var]

    fig, ax = plt.subplots(figsize=(8, 4))

    for d in unique_dates[:n_days]:
        d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
        df_day = df_mean_time[df_mean_time["date"].dt.date == d].copy()
        df_day["hour"] = df_day["date"].dt.hour

        ax.plot(
            df_day["hour"],
            df_day[var],
            marker="o",
            linestyle="-",
            label=d_str
        )

    # Líneas de umbral específicas
    if var == "wind_gusts_10m":
        threshold = 20
        label_umbral = "Umbral 20 km/h"
    else:
        threshold = 10
        label_umbral = "Umbral 10 km/h"

    ax.axhline(
        y=threshold,
        color="red",
        linestyle="--",
        linewidth=1.5,
        label=label_umbral
    )

    ax.set_title("Ciclo diario (promedio horario)")
    ax.set_xlabel("Hora del día (hora local)")
    ax.set_ylabel(y_label)  # (km/h)

    ax.set_xticks(range(0, 24, 2))
    ax.grid(True)
    ax.legend()

    out_ts = os.path.join(
        results_dir,
        f"Serie_ciclo_diario_{var}_{run_date_str}.png"
    )
    plt.tight_layout()
    plt.savefig(out_ts, dpi=150)
    plt.close(fig)

    print(f"Guardada serie de tiempo para {var}: {out_ts}")

# 2) Direcciones (10, 100, 200 m)
#    Nota: aquí se usa promedio simple de grados. Si quieres, luego podemos cambiar a media circular.
df_dir_time = df_all.groupby("date")[dir_vars].mean().reset_index()

for var in dir_vars:
    title_label, y_label = var_labels[var]

    fig, ax = plt.subplots(figsize=(8, 4))

    # ----------------------------------------------------
    # Bandas de color por rango de dirección
    # ----------------------------------------------------
    # 0–90°: rojo (no favorable)
    ax.axhspan(0, 90,
               facecolor="#f8b6b6",  # rojo suave
               alpha=0.4,
               zorder=0)

    # 90–270°: verde (favorable)
    ax.axhspan(90, 270,
               facecolor="#e0f6e0",  # verde suave
               alpha=0.4,
               zorder=0)

    # 270–360°: rojo (no favorable)
    ax.axhspan(270, 360,
               facecolor="#f8b6b6",
               alpha=0.4,
               zorder=0)

    # ----------------------------------------------------
    # Series de tiempo de dirección
    # ----------------------------------------------------
    for d in unique_dates[:n_days]:
        d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
        df_day = df_dir_time[df_dir_time["date"].dt.date == d].copy()
        df_day["hour"] = df_day["date"].dt.hour

        ax.plot(
            df_day["hour"],
            df_day[var],
            marker="o",
            linestyle="-",
            label=d_str,
            zorder=5  # por encima de las bandas
        )

    ax.set_title("Ciclo diario (dirección, promedio horario)")
    ax.set_xlabel("Hora del día (hora local)")
    ax.set_ylabel(y_label)  # (°)

    ax.set_xticks(range(0, 24, 2))
    ax.set_yticks(range(0, 361, 45))
    ax.set_ylim(0, 360)      # asegura que se vean las tres bandas completas
    ax.grid(True, zorder=1)
    ax.legend()

    out_ts = os.path.join(
        results_dir,
        f"Serie_ciclo_diario_{var}_{run_date_str}.png"
    )
    plt.tight_layout()
    plt.savefig(out_ts, dpi=150)
    plt.close(fig)

    print(f"Guardada serie de tiempo para {var}: {out_ts}")
