"""
Script 2: Generación de mapas espaciales y series de tiempo.
- Lee datos descargados (incluyendo wind_speed_200m calculado por Ley de Potencia).
- Genera mapas de gradiente con vectores de viento para 10m, 100m, 200m y ráfagas.
- Genera series de tiempo con ciclo diario y bandas de favorabilidad.
- Organiza resultados en carpetas YYYYMMDD.
"""

import os
import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point

import openmeteo_requests
import requests_cache
from retry_requests import retry
import matplotlib
matplotlib.use("Agg")  # Backend sin pantalla (necesario en GitHub Actions)
import matplotlib.pyplot as plt
from scipy.interpolate import griddata
from scipy.spatial import QhullError

# ---------------------------------------------------------
# Rutas
# ---------------------------------------------------------
base_dir = os.path.dirname(os.path.abspath(__file__))
shapefile_path = os.path.join(base_dir, "1_Shapefile", "Area_of_interest.shp")
results_dir = os.path.join(base_dir, "2_Results")
os.makedirs(results_dir, exist_ok=True)

# Carpeta de la fecha actual
run_date_str = pd.Timestamp.now(tz="America/Guatemala").strftime("%Y%m%d")
daily_folder = os.path.join(results_dir, run_date_str)
os.makedirs(daily_folder, exist_ok=True)

# ---------------------------------------------------------
# Leer shapefile
# ---------------------------------------------------------
if not os.path.exists(shapefile_path):
    raise FileNotFoundError(f"Shapefile no encontrado en: {shapefile_path}")

gdf = gpd.read_file(shapefile_path)
if gdf.crs is None or gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

try:
    main_poly = gdf.union_all()
except AttributeError:
    main_poly = gdf.unary_union

# ---------------------------------------------------------
# Rejilla de puntos
# ---------------------------------------------------------
minx, miny, maxx, maxy = main_poly.bounds
target_res_deg = 0.1
n_lon = max(3, int((maxx - minx) / target_res_deg) + 1)
n_lat = max(3, int((maxy - miny) / target_res_deg) + 1)

lon_vals = np.linspace(minx, maxx, n_lon)
lat_vals = np.linspace(miny, maxy, n_lat)

points = []
for lat in lat_vals:
    for lon in lon_vals:
        if main_poly.contains(Point(lon, lat)):
            points.append((lat, lon))

if not points:
    raise RuntimeError("No se encontraron puntos dentro del polígono.")

print(f"Usando {len(points)} puntos de muestreo.")

# ---------------------------------------------------------
# Cliente Open-Meteo
# ---------------------------------------------------------
cache_session = requests_cache.CachedSession('.cache', expire_after=3600)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)
url = "https://api.open-meteo.com/v1/forecast"

ALPHA_DEFAULT = 0.143

# Variables nativas (sin 200m)
native_vars = [
    "wind_speed_10m",
    "wind_speed_100m",
    "wind_direction_10m",
    "wind_direction_100m",
    "wind_gusts_10m",
]

# ---------------------------------------------------------
# Descarga y cálculo de 200m por Ley de Potencia
# ---------------------------------------------------------
all_dfs = []

for (lat, lon) in points:
    params = {
        "latitude": lat,
        "longitude": lon,
        "hourly": native_vars,
        "models": "ecmwf_ifs",
        "forecast_days": 3,
        "timezone": "America/Guatemala",
        "windspeed_unit": "kmh"
    }

    responses = openmeteo.weather_api(url, params=params)
    response = responses[0]
    hourly = response.Hourly()
    vals = [hourly.Variables(i).ValuesAsNumpy() for i in range(len(native_vars))]

    date_index = pd.date_range(
        start=pd.to_datetime(hourly.Time(), unit="s", utc=True).tz_convert("America/Guatemala"),
        end=pd.to_datetime(hourly.TimeEnd(), unit="s", utc=True).tz_convert("America/Guatemala"),
        freq=pd.Timedelta(seconds=hourly.Interval()),
        inclusive="left"
    )

    data = {"date": date_index, "lat": lat, "lon": lon}
    for name, arr in zip(native_vars, vals):
        data[name] = arr

    df_point = pd.DataFrame(data)

    # Ley de Potencia → 200m
    v10  = df_point["wind_speed_10m"].values.copy()
    v100 = df_point["wind_speed_100m"].values.copy()
    alpha = np.full(len(v10), ALPHA_DEFAULT, dtype=float)

    valid = (v10 > 0.5) & (v100 > 0.5)
    ratio = np.where(valid, v100 / v10, np.nan)
    with np.errstate(divide='ignore', invalid='ignore'):
        alpha_dyn = np.log(ratio) / np.log(100.0 / 10.0)

    ok = valid & np.isfinite(alpha_dyn) & (alpha_dyn > 0) & (alpha_dyn < 0.6)
    alpha[ok] = alpha_dyn[ok]

    df_point["wind_speed_200m"]     = v100 * (200.0 / 100.0) ** alpha
    df_point["wind_direction_200m"] = df_point["wind_direction_100m"]

    all_dfs.append(df_point)

df_all = pd.concat(all_dfs, ignore_index=True)
df_all["date_only"] = df_all["date"].dt.date
unique_dates = sorted(df_all["date_only"].unique())
n_days = min(3, len(unique_dates))

# ---------------------------------------------------------
# Variables para graficar
# ---------------------------------------------------------
speed_vars = ["wind_speed_10m", "wind_speed_100m", "wind_speed_200m", "wind_gusts_10m"]
dir_vars   = ["wind_direction_10m", "wind_direction_100m", "wind_direction_200m"]
all_plot_vars = speed_vars + dir_vars

var_labels = {
    "wind_speed_10m":      ("Velocidad del viento a 10 m",  "Velocidad del viento a 10 m (km/h)"),
    "wind_speed_100m":     ("Velocidad del viento a 100 m", "Velocidad del viento a 100 m (km/h)"),
    "wind_speed_200m":     ("Velocidad del viento a 200 m", "Velocidad del viento a 200 m (km/h)"),
    "wind_gusts_10m":      ("Ráfagas de viento a 10 m",     "Ráfagas de viento a 10 m (km/h)"),
    "wind_direction_10m":  ("Dirección del viento a 10 m",  "Dirección del viento a 10 m (°)"),
    "wind_direction_100m": ("Dirección del viento a 100 m", "Dirección del viento a 100 m (°)"),
    "wind_direction_200m": ("Dirección del viento a 200 m", "Dirección del viento a 200 m (°)"),
}

# Correspondencia velocidad ↔ dirección para flechas en mapas
speed_dir_map = {
    "wind_speed_10m":  ("wind_direction_10m",  "wind_speed_10m"),
    "wind_speed_100m": ("wind_direction_100m",  "wind_speed_100m"),
    "wind_speed_200m": ("wind_direction_200m",  "wind_speed_200m"),
    "wind_gusts_10m":  ("wind_direction_10m",   "wind_gusts_10m"),
}

# ---------------------------------------------------------
# Rejilla fina para interpolación de mapas
# ---------------------------------------------------------
n_lon_fine = max(40, int((maxx - minx) / 0.02) + 1)
n_lat_fine = max(40, int((maxy - miny) / 0.02) + 1)
lon_fine = np.linspace(minx, maxx, n_lon_fine)
lat_fine = np.linspace(miny, maxy, n_lat_fine)
Lon_fine, Lat_fine = np.meshgrid(lon_fine, lat_fine)

mask = np.zeros(Lon_fine.shape, dtype=bool)
for i in range(Lon_fine.shape[0]):
    for j in range(Lon_fine.shape[1]):
        if main_poly.contains(Point(Lon_fine[i, j], Lat_fine[i, j])):
            mask[i, j] = True

# =============================================================
# MAPAS DE GRADIENTE ESPACIAL + VECTORES DE VIENTO
# =============================================================
print("\nGenerando mapas espaciales...")

for day_idx in range(n_days):
    d = unique_dates[day_idx]
    d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
    df_day = df_all[df_all["date_only"] == d].copy()
    df_day_mean = df_day.groupby(["lat", "lon"])[speed_vars + dir_vars].mean().reset_index()
    coords = np.vstack((df_day_mean["lon"].values, df_day_mean["lat"].values)).T

    for var in speed_vars:
        values = df_day_mean[var].values
        try:
            grid_fine = griddata(coords, values, (Lon_fine, Lat_fine), method="linear")
        except QhullError:
            grid_fine = griddata(coords, values, (Lon_fine, Lat_fine), method="nearest")

        grid_fine[~mask] = np.nan

        dir_col, mag_col = speed_dir_map[var]
        dir_deg = df_day_mean[dir_col].values
        mag     = df_day_mean[mag_col].values
        lon_ar  = df_day_mean["lon"].values
        lat_ar  = df_day_mean["lat"].values

        dir_rad = np.deg2rad(dir_deg)
        u = -mag * np.sin(dir_rad)
        v = -mag * np.cos(dir_rad)

        title_label, cbar_label = var_labels[var]

        fig, ax = plt.subplots(figsize=(6, 6))
        pcm = ax.pcolormesh(Lon_fine, Lat_fine, grid_fine, shading="auto", cmap="YlOrRd")
        cbar = plt.colorbar(pcm, ax=ax, label=cbar_label, fraction=0.035, pad=0.02)
        gdf.boundary.plot(ax=ax, color="black", linewidth=1)
        ax.quiver(
            lon_ar, lat_ar, u * 0.5, v * 0.5,
            angles="xy", scale_units="xy", scale=120,
            width=0.0016, alpha=0.9, color="k",
            headwidth=12, headlength=20, headaxislength=14
        )
        ax.set_title(f"{title_label} – {d_str}", fontsize=11, fontweight="bold")
        ax.set_xlabel("Longitud (°)")
        ax.set_ylabel("Latitud (°)")
        ax.set_xticks(np.arange(np.floor(minx * 10) / 10, np.ceil(maxx * 10) / 10 + 0.001, 0.1))
        ax.tick_params(axis='x', rotation=45)

        out_png = os.path.join(daily_folder, f"Mapa_{var}_{run_date_str}.png")
        plt.tight_layout()
        plt.savefig(out_png, dpi=150, bbox_inches="tight")
        plt.close(fig)
        print(f"  Mapa guardado: {out_png}")

# =============================================================
# SERIES DE TIEMPO – CICLO DIARIO (0–23 h)
# =============================================================
print("\nGenerando series de tiempo...")

COLORS = ["#1f77b4", "#ff7f0e", "#2ca02c"]

# ---- 1) Velocidades y ráfagas ----
df_speed_time = df_all.groupby("date")[speed_vars].mean().reset_index()

for var in speed_vars:
    title_label, y_label = var_labels[var]
    threshold = 20 if var == "wind_gusts_10m" else 10
    label_umbral = f"Umbral {threshold} km/h"

    fig, ax = plt.subplots(figsize=(10, 4.5))

    for idx, d in enumerate(unique_dates[:n_days]):
        d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
        df_day = df_speed_time[df_speed_time["date"].dt.date == d].copy()
        df_day["hour"] = df_day["date"].dt.hour
        ax.plot(df_day["hour"], df_day[var],
                marker="o", linestyle="-", color=COLORS[idx], label=d_str, linewidth=1.8)

    ax.axhline(y=threshold, color="red", linestyle="--", linewidth=1.5, label=label_umbral)

    ax.set_title(f"Ciclo diario – {title_label} (promedio horario)", fontsize=11, fontweight="bold")
    ax.set_xlabel("Hora del día (hora local)")
    ax.set_ylabel(y_label)
    ax.set_xticks(range(0, 24, 2))
    ax.grid(True, alpha=0.5)
    ax.legend(framealpha=0.9)

    out_ts = os.path.join(daily_folder, f"Serie_ciclo_diario_{var}_{run_date_str}.png")
    plt.tight_layout()
    plt.savefig(out_ts, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Serie guardada: {out_ts}")

# ---- 2) Direcciones (10, 100, 200 m) ----
df_dir_time = df_all.groupby("date")[dir_vars].mean().reset_index()

for var in dir_vars:
    title_label, y_label = var_labels[var]

    fig, ax = plt.subplots(figsize=(10, 4.5))

    # Bandas de favorabilidad
    ax.axhspan(0,   90,  facecolor="#f8b6b6", alpha=0.4, zorder=0, label="No favorable")
    ax.axhspan(90,  270, facecolor="#e0f6e0", alpha=0.4, zorder=0, label="Favorable (90°–270°)")
    ax.axhspan(270, 360, facecolor="#f8b6b6", alpha=0.4, zorder=0)

    for idx, d in enumerate(unique_dates[:n_days]):
        d_str = pd.to_datetime(d).strftime("%d-%m-%Y")
        df_day = df_dir_time[df_dir_time["date"].dt.date == d].copy()
        df_day["hour"] = df_day["date"].dt.hour
        ax.plot(df_day["hour"], df_day[var],
                marker="o", linestyle="-", color=COLORS[idx], label=d_str,
                linewidth=1.8, zorder=5)

    ax.set_title(f"Ciclo diario – {title_label} (promedio horario)", fontsize=11, fontweight="bold")
    ax.set_xlabel("Hora del día (hora local)")
    ax.set_ylabel(y_label)
    ax.set_xticks(range(0, 24, 2))
    ax.set_yticks(range(0, 361, 45))
    ax.set_ylim(0, 360)
    ax.grid(True, alpha=0.5, zorder=1)
    ax.legend(framealpha=0.9)

    out_ts = os.path.join(daily_folder, f"Serie_ciclo_diario_{var}_{run_date_str}.png")
    plt.tight_layout()
    plt.savefig(out_ts, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Serie guardada: {out_ts}")

print(f"\n✓ Todos los resultados guardados en: {daily_folder}")
