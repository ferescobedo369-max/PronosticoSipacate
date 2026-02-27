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
# Rutas CORREGIDOS PARA GITHUB
# ---------------------------------------------------------
base_dir = os.path.dirname(os.path.abspath(__file__))

shapefile_path = os.path.join(base_dir, "1_Shapefile", "Area_of_interest.shp")
results_dir = os.path.join(base_dir, "2_Results")
os.makedirs(results_dir, exist_ok=True)

# ---------------------------------------------------------
# Leer shapefile principal (asegurar WGS84)
# ---------------------------------------------------------
if not os.path.exists(shapefile_path):
    raise FileNotFoundError(f"No se encontró el shapefile en: {shapefile_path}")

gdf = gpd.read_file(shapefile_path)
if gdf.crs is None or gdf.crs.to_epsg() != 4326:
    gdf = gdf.to_crs(epsg=4326)

try:
    main_poly = gdf.union_all()
except AttributeError:
    main_poly = gdf.unary_union

# ---------------------------------------------------------
# Construir rejilla de puntos
# ---------------------------------------------------------
minx, miny, maxx, maxy = main_poly.bounds
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
        if main_poly.contains(p):
            points.append((lat, lon))

# ---------------------------------------------------------
# Cliente de Open-Meteo
# ---------------------------------------------------------
cache_session = requests_cache.CachedSession('.cache', expire_after=3600)
retry_session = retry(cache_session, retries=5, backoff_factor=0.2)
openmeteo = openmeteo_requests.Client(session=retry_session)
url = "https://api.open-meteo.com/v1/forecast"

# Variables
speed_vars = ["wind_speed_10m", "wind_speed_100m", "wind_speed_200m", "wind_gusts_10m"]
dir_vars = ["wind_direction_10m", "wind_direction_100m", "wind_direction_200m"]
all_vars = speed_vars + dir_vars

var_labels = {
    "wind_speed_10m": ("Velocidad 10m", "km/h"),
    "wind_speed_100m": ("Velocidad 100m", "km/h"),
    "wind_speed_200m": ("Velocidad 200m", "km/h"),
    "wind_gusts_10m": ("Ráfagas 10m", "km/h"),
    "wind_direction_10m": ("Dirección 10m", "°"),
    "wind_direction_100m": ("Dirección 100m", "°"),
    "wind_direction_200m": ("Dirección 200m", "°"),
}

# ---------------------------------------------------------
# Descarga de datos
# ---------------------------------------------------------
all_dfs = []
for (lat, lon) in points:
    params = {
        "latitude": lat, "longitude": lon, "hourly": all_vars,
        "models": "ecmwf_ifs", "forecast_days": 3,
        "timezone": "America/Guatemala", "windspeed_unit": "kmh"
    }
    responses = openmeteo.weather_api(url, params=params)
    response = responses[0]
    hourly = response.Hourly()
    vals = [hourly.Variables(i).ValuesAsNumpy() for i in range(len(all_vars))]
    date_index = pd.date_range(
        start=pd.to_datetime(hourly.Time(), unit="s", utc=True).tz_convert("America/Guatemala"),
        end=pd.to_datetime(hourly.TimeEnd(), unit="s", utc=True).tz_convert("America/Guatemala"),
        freq=pd.Timedelta(seconds=hourly.Interval()), inclusive="left"
    )
    data = {"date": date_index, "lat": lat, "lon": lon}
    for name, arr in zip(all_vars, vals): data[name] = arr
    all_dfs.append(pd.DataFrame(data))

df_all = pd.concat(all_dfs, ignore_index=True)
df_all["date_only"] = df_all["date"].dt.date
unique_dates = sorted(df_all["date_only"].unique())
n_days = min(3, len(unique_dates))

# ---------------------------------------------------------
# Generación de Mapas y Gráficas
# ---------------------------------------------------------
# (Mantenemos tu lógica de interpolación y gráficas igual, 
# solo ajustamos la salida de archivos)

fine_res_deg = 0.02
lon_fine = np.linspace(minx, maxx, 40)
lat_fine = np.linspace(miny, maxy, 40)
Lon_fine, Lat_fine = np.meshgrid(lon_fine, lat_fine)

mask = np.zeros(Lon_fine.shape, dtype=bool)
for i in range(Lon_fine.shape[0]):
    for j in range(Lon_fine.shape[1]):
        if main_poly.contains(Point(Lon_fine[i, j], Lat_fine[i, j])):
            mask[i, j] = True

for day_idx in range(n_days):
    d = unique_dates[day_idx]
    df_day = df_all[df_all["date_only"] == d].copy()
    df_day_mean = df_day.groupby(["lat", "lon"])[all_vars].mean().reset_index()
    coords = np.vstack((df_day_mean["lon"].values, df_day_mean["lat"].values)).T

    for var in speed_vars:
        values = df_day_mean[var].values
        try:
            grid_fine = griddata(coords, values, (Lon_fine, Lat_fine), method="linear")
        except:
            grid_fine = griddata(coords, values, (Lon_fine, Lat_fine), method="nearest")
        
        grid_fine[~mask] = np.nan
        fig, ax = plt.subplots(figsize=(6, 6))
        pcm = ax.pcolormesh(Lon_fine, Lat_fine, grid_fine, shading="auto")
        plt.colorbar(pcm, ax=ax, label=var_labels[var][1])
        gdf.boundary.plot(ax=ax, color="black")
        
        # Nombre de archivo simplificado para que sea fácil de encontrar
        out_png = os.path.join(results_dir, f"Mapa_{var}_Dia_{day_idx}.png")
        plt.savefig(out_png, dpi=100)
        plt.close(fig)

# ---------------------------------------------------------
# Series de Tiempo
# ---------------------------------------------------------
df_mean_time = df_all.groupby("date")[speed_vars].mean().reset_index()
for var in speed_vars:
    fig, ax = plt.subplots(figsize=(8, 4))
    for d in unique_dates[:n_days]:
        df_day = df_mean_time[df_mean_time["date"].dt.date == d].copy()
        ax.plot(df_day["date"].dt.hour, df_day[var], marker="o", label=str(d))
    
    ax.legend()
    out_ts = os.path.join(results_dir, f"Serie_{var}.png")
    plt.savefig(out_ts, dpi=100)
    plt.close(fig)

print("Mapas y Series generados con éxito en la carpeta 2_Results.")
