# -*- coding: utf-8 -*-
"""
fetch_climate.py
Runs daily via GitHub Actions.
Fetches last 24h of hourly climate data from IPMA (Guimarães station)
and appends new rows to data/climate_guimaraes.xlsx.
"""

import os
import json
import math
import datetime
import urllib.request

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("openpyxl not installed. Run: pip install openpyxl")

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
TARGET_LAT  = 41.4434
TARGET_LON  = -8.2938
TARGET_NAME = "Guimarães"

IPMA_STATIONS_URL = "https://api.ipma.pt/open-data/observation/meteorology/stations/stations.json"
IPMA_OBS_URL      = "https://api.ipma.pt/open-data/observation/meteorology/stations/observations.json"

OUTPUT_PATH = os.path.join("data", "climate_guimaraes.xlsx")

COLUMNS = [
    "Timestamp (UTC)",
    "Station",
    "Temperature (°C)",
    "Humidity (%)",
    "Precipitation (mm)",
    "Wind Speed (km/h)",
    "Wind Speed (m/s)",
    "Wind Direction",
    "Pressure (hPa)",
    "Radiation (W/m²)",
]

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def fetch_json(url):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 GitHubActions ClimateMonitor/1.0"}
    )
    response = urllib.request.urlopen(req, timeout=15)
    return json.loads(response.read().decode("utf-8"))

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    d_lat = math.radians(lat2 - lat1)
    d_lon = math.radians(lon2 - lon1)
    a = (math.sin(d_lat/2)**2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(d_lon/2)**2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def clean(val, decimals=1):
    if val is None:
        return None
    try:
        fval = float(val)
    except (ValueError, TypeError):
        return None
    return None if fval == -99.0 else round(fval, decimals)

WIND_DIR = {
    0: "Calm", 1: "N", 2: "NE", 3: "E", 4: "SE",
    5: "S",    6: "SW", 7: "W", 8: "NW", 9: "N"
}

# ---------------------------------------------------------------------------
# STEP 1 — Find nearest station
# ---------------------------------------------------------------------------
print("Finding nearest station to {}...".format(TARGET_NAME))
stations = fetch_json(IPMA_STATIONS_URL)

best_id, best_name, best_dist = None, None, float("inf")
for feature in stations:
    props  = feature.get("properties", {})
    coords = feature.get("geometry", {}).get("coordinates", [])
    if not coords or len(coords) < 2:
        continue
    dist = haversine(TARGET_LAT, TARGET_LON, float(coords[1]), float(coords[0]))
    if dist < best_dist:
        best_dist = dist
        best_id   = str(props.get("idEstacao", ""))
        best_name = props.get("localEstacao", best_id)

print("Nearest station: {} (ID: {}) — {:.1f} km".format(best_name, best_id, best_dist))

# ---------------------------------------------------------------------------
# STEP 2 — Fetch observations
# ---------------------------------------------------------------------------
print("Fetching observations...")
obs_data = fetch_json(IPMA_OBS_URL)

new_records = []
for ts in sorted(obs_data.keys()):
    sd = obs_data[ts].get(best_id)
    if sd is None:
        continue
    new_records.append([
        ts,
        best_name,
        clean(sd.get("temperatura")),
        clean(sd.get("humidade")),
        clean(sd.get("precAcumulada")),
        clean(sd.get("intensidadeVentoKM")),
        clean(sd.get("intensidadeVento")),
        WIND_DIR.get(sd.get("idDireccVento"), ""),
        clean(sd.get("pressao")),
        clean(sd.get("radiacao")),
    ])

print("{} records fetched.".format(len(new_records)))

# ---------------------------------------------------------------------------
# STEP 3 — Append to master Excel file
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", start_color="2F5496", end_color="2F5496")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DATA_FONT   = Font(name="Arial", size=10)
ALT_FILL    = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
CENTER      = Alignment(horizontal="center", vertical="center")
LEFT        = Alignment(horizontal="left",   vertical="center")
THIN_BORDER = Border(
    bottom=Side(style="thin", color="B8CCE4"),
    right =Side(style="thin", color="B8CCE4"),
)
COL_WIDTHS  = [20, 22, 16, 14, 18, 18, 16, 16, 14, 16]

if os.path.exists(OUTPUT_PATH):
    wb = openpyxl.load_workbook(OUTPUT_PATH)
    ws = wb["Hourly Data"]
    # Find existing timestamps to avoid duplicates
    existing_timestamps = set()
    for row in ws.iter_rows(min_row=4, max_col=1, values_only=True):
        if row[0]:
            existing_timestamps.add(str(row[0]))
    print("{} existing rows found.".format(len(existing_timestamps)))
else:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hourly Data"

    # Title
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value     = "IPMA Climate Data — {} — Master Log".format(TARGET_NAME)
    c.font      = Font(name="Arial", bold=True, size=13, color="1F3864")
    c.alignment = CENTER
    c.fill      = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
    ws.row_dimensions[1].height = 24

    # Subtitle
    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value     = "Source: IPMA open-data API  |  Station: {} (ID: {})  |  Auto-updated daily via GitHub Actions".format(
        best_name, best_id
    )
    c.font      = Font(name="Arial", size=9, italic=True, color="595959")
    c.alignment = LEFT

    # Headers
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        c = ws.cell(row=3, column=col_idx, value=col_name)
        c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER
    ws.row_dimensions[3].height = 20

    # Column widths
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"
    existing_timestamps = set()

# Append only new rows
added = 0
for record in new_records:
    if str(record[0]) in existing_timestamps:
        continue
    row_idx = ws.max_row + 1
    fill = ALT_FILL if (row_idx % 2 == 0) else None
    for col_idx, value in enumerate(record, start=1):
        c = ws.cell(row=row_idx, column=col_idx, value=value)
        c.font = DATA_FONT; c.alignment = CENTER; c.border = THIN_BORDER
        if fill: c.fill = fill
    added += 1

print("{} new rows added.".format(added))

os.makedirs("data", exist_ok=True)
wb.save(OUTPUT_PATH)
print("Saved to {}".format(OUTPUT_PATH))
