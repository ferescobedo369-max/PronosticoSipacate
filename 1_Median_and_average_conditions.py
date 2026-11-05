"""
Script 1: Extracción de datos meteorológicos y cálculo de estadísticas de área.
- Descarga datos horarios de Open-Meteo (ECMWF IFS).
- Calcula velocidad a 200m usando Ley de Potencia (alpha dinámico).
- Guarda DOS CSVs:
    1. Area_forecast_latest.csv              → compatible con Power BI (misma estructura que antes)
    2. Area_forecast_mean_median_YYYYMMDD.csv → historial completo con mean/median
"""

import os
import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point

import openmeteo_requests
import requests_cache
from retry_requests import retry

# ---------------------------------------------------------
# Rutas (compatibles con GitHub Actions)
# ---------------------------------------------------------
base_dir = os.path.dirname(os.path.abspath(__file__))
shapefile_path = os.path.join(base_dir, "1_Shapefile", "Area_of_interest.shp")
results_dir = os.path.join(base_dir, "2_Results")
os.makedirs(results_dir, exist_ok=True)

# ---------------------------------------------------------
# Leer shapefile
# ---------------------------------------------------------
if not os.path.exists(shapefile_path):
    raise FileNotFoundError(f"Shapefile no encontrado en: {shapefile_path}")

gdf = gpd.read_file(shapefile_path)
if gdf.crs is None or gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

try:
    polygon = gdf.union_all()
except AttributeError:
    polygon = gdf.unary_union

# ---------------------------------------------------------
# Rejilla de puntos dentro del polígono
# ---------------------------------------------------------
minx, miny, maxx, maxy = polygon.bounds

target_res_deg = 0.1
n_lon = max(3, int((maxx - minx) / target_res_deg) + 1)
n_lat = max(3, int((maxy - miny) / target_res_deg) + 1)

lon_vals = np.linspace(minx, maxx, n_lon)
lat_vals = np.linspace(miny, maxy, n_lat)

points = []
for lat in lat_vals:
    for lon in lon_vals:
        if polygon.contains(Point(lon, lat)):
            points.append((lat, lon))

if not points:
    raise RuntimeError("No se encontraron puntos dentro del polígono.")

print(f"Usando {len(points)} puntos de muestreo dentro del polígono.")

# ---------------------------------------------------------
# Cliente Open-Meteo
# ---------------------------------------------------------
cache_session = requests_cache.CachedSession('.cache', expire_after=3600)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)

url = "https://api.open-meteo.com/v1/forecast"

ALPHA_DEFAULT = 0.143

# Variables nativas (sin 200m — se calcula por Ley de Potencia)
hourly_vars = [
    "temperature_2m",
    "wind_speed_10m",
    "wind_speed_100m",
    "wind_direction_10m",
    "wind_direction_100m",
    "wind_gusts_10m",
    "rain"
]

# ---------------------------------------------------------
# Descarga de datos por punto
# ---------------------------------------------------------
all_dfs = []

for (lat, lon) in points:
    params = {
        "latitude": lat,
        "longitude": lon,
        "hourly": hourly_vars,
        "models": "ecmwf_ifs",
        "forecast_days": 3,
        "timezone": "America/Guatemala",
        "windspeed_unit": "kmh"
    }

    responses = openmeteo.weather_api(url, params=params)
    response = responses[0]
    hourly = response.Hourly()
    vals = [hourly.Variables(i).ValuesAsNumpy() for i in range(len(hourly_vars))]

    date_index = pd.date_range(
        start=pd.to_datetime(hourly.Time(), unit="s", utc=True).tz_convert("America/Guatemala"),
        end=pd.to_datetime(hourly.TimeEnd(), unit="s", utc=True).tz_convert("America/Guatemala"),
        freq=pd.Timedelta(seconds=hourly.Interval()),
        inclusive="left"
    )

    data = {"date": date_index, "lat": lat, "lon": lon}
    for name, arr in zip(hourly_vars, vals):
        data[name] = arr

    df_point = pd.DataFrame(data)

    # ---------------------------------------------------------
    # Ley de Potencia → velocidad a 200m
    # ---------------------------------------------------------
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
    df_point["alpha_shear"]         = alpha

    all_dfs.append(df_point)

df_all = pd.concat(all_dfs, ignore_index=True)

# =============================================================
# CSV 1: Area_forecast_latest.csv
# Estructura IDÉNTICA a la que Power BI ya conoce:
#   date, wind_speed_10m, wind_speed_100m, wind_speed_200m,
#   wind_direction_10m, wind_direction_100m, wind_direction_200m,
#   wind_gusts_10m, Zona_Favorable, Dia, Hora
# =============================================================

powerbi_cols = [
    "wind_speed_10m",
    "wind_speed_100m",
    "wind_speed_200m",
    "wind_direction_10m",
    "wind_direction_100m",
    "wind_direction_200m",
    "wind_gusts_10m",
]

# Promedio de área por hora (igual que hacía el script original de Power BI)
df_powerbi = df_all.groupby("date")[powerbi_cols].mean().reset_index()

# Columna Zona_Favorable (idéntica a tu script original)
df_powerbi["Zona_Favorable"] = df_powerbi["wind_direction_10m"].apply(
    lambda x: "Favorable" if 90 <= x <= 270 else "No Favorable"
)

# Columnas Dia y Hora separadas para facilitar filtros en Power BI
df_powerbi["Dia"]  = df_powerbi["date"].dt.strftime("%d-%m-%Y")
df_powerbi["Hora"] = df_powerbi["date"].dt.hour

# Guardar — siempre sobreescribe (Power BI apunta aquí)
latest_file = os.path.join(results_dir, "Area_forecast_latest.csv")
df_powerbi.to_csv(latest_file, index=False, encoding="utf-8")
print(f"✓ CSV Power BI guardado : {latest_file}")

# =============================================================
# CSV 2: Historial con mean/median en carpeta YYYYMMDD
# =============================================================
all_cols = hourly_vars + ["wind_speed_200m", "wind_direction_200m", "alpha_shear"]

agg = df_all.groupby("date")[all_cols].agg(["mean", "median"]).reset_index()
agg.columns = ["date"] + [f"{var}_{stat}" for var in all_cols for stat in ["mean", "median"]]

run_date_str = pd.Timestamp.now(tz="America/Guatemala").strftime("%Y%m%d")
daily_folder = os.path.join(results_dir, run_date_str)
os.makedirs(daily_folder, exist_ok=True)

hist_file = os.path.join(daily_folder, f"Area_forecast_mean_median_{run_date_str}.csv")
agg.to_csv(hist_file, index=False, encoding="utf-8")
print(f"✓ CSV historial guardado : {hist_file}")

print("\n✓ Proceso completado.")
