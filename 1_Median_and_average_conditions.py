import os
import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point

import openmeteo_requests
import requests_cache
from retry_requests import retry

# ---------------------------------------------------------
# Paths CORREGIDOS PARA GITHUB
# ---------------------------------------------------------
# Detecta la ubicación del script en el repositorio
base_dir = os.path.dirname(os.path.abspath(__file__))

# Las carpetas deben llamarse exactamente como las subiste a GitHub
shapefile_path = os.path.join(base_dir, "1_Shapefile", "Area_of_interest.shp")
results_dir = os.path.join(base_dir, "2_Results")

# Crear carpeta de resultados si no existe
os.makedirs(results_dir, exist_ok=True)

# ---------------------------------------------------------
# Read shapefile (ensure WGS84)
# ---------------------------------------------------------
if not os.path.exists(shapefile_path):
    raise FileNotFoundError(f"No se encontró el shapefile en: {shapefile_path}")

gdf = gpd.read_file(shapefile_path)
if gdf.crs is None or gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

# Unified polygon (in case of multipart geometries)
try:
    polygon = gdf.union_all()
except AttributeError:
    polygon = gdf.unary_union

# ---------------------------------------------------------
# Build grid of points inside polygon
# ---------------------------------------------------------
minx, miny, maxx, maxy = polygon.bounds

target_res_deg = 0.1
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
        if polygon.contains(p):
            points.append((lat, lon))

if not points:
    raise RuntimeError("No grid points fell inside the polygon.")

print(f"Using {len(points)} grid points inside polygon for areal statistics.")

# ---------------------------------------------------------
# Setup Open-Meteo client
# ---------------------------------------------------------
# El cache se guarda en una carpeta temporal en la nube
cache_session = requests_cache.CachedSession('.cache', expire_after=3600)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)

url = "https://api.open-meteo.com/v1/forecast"

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
# Loop over points, collect hourly data
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
    all_dfs.append(df_point)

df_all = pd.concat(all_dfs, ignore_index=True)

# ---------------------------------------------------------
# Areal mean & median per hour
# ---------------------------------------------------------
var_cols = hourly_vars
agg = df_all.groupby("date")[var_cols].agg(["mean", "median"]).reset_index()

agg.columns = ["date"] + [
    f"{var}_{stat}"
    for var in var_cols
    for stat in ["mean", "median"]
]

# ---------------------------------------------------------
# Save result
# ---------------------------------------------------------
run_date_str = pd.Timestamp.now(tz="America/Guatemala").strftime("%Y%m%d")

# NOTA: En GitHub, simplificaremos la ruta para que Power BI la encuentre siempre igual
output_file = os.path.join(results_dir, f"Area_forecast_latest.csv")

agg.to_csv(output_file, index=False, encoding="utf-8")

print(f"\nProceso completado. Datos guardados en: {output_file}")
