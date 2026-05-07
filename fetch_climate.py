# -*- coding: utf-8 -*-
"""
Daily IPMA climate data fetcher for GitHub Actions.
Appends the last 24h of hourly observations for the nearest
station to Guimarães to a master Excel file.
Runs headless — no pyrevit, no GUI, no ctypes.
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
    raise ImportError("Run: pip install openpyxl")

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
TARGET_LAT  = 41.4434
TARGET_LON  = -8.2938
TARGET_NAME = "Guimarães"

IPMA_STATIONS_URL = "https://api.ipma.pt/open-data/observation/meteorology/stations/stations.json"
IPMA_OBS_URL      = "https://api.ipma.pt/open-data/observation/meteorology/stations/observations.json"

DATA_DIR   = "data"
EXCEL_FILE = os.path.join(DATA_DIR, "climate_guimaraes.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

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

NO_DATA = -99.0
def clean(val, decimals=1):
    if val is None:
        return None
    try:
        fval = float(val)
    except (ValueError, TypeError):
        return None
    return None if fval == NO_DATA else round(fval, decimals)

WIND_DIR = {
    0: "Calm", 1: "N", 2: "NE", 3: "E", 4: "SE",
    5: "S",    6: "SW", 7: "W", 8: "NW", 9: "N"
}

COLUMNS = [
    "Timestamp (UTC)", "Timestamp (Lisbon)", "Station",
    "Temperature (°C)", "Humidity (%)", "Precipitation (mm)",
    "Wind Speed (km/h)", "Wind Speed (m/s)", "Wind Direction",
    "Pressure (hPa)", "Radiation (W/m²)"
]

def utc_to_lisbon(ts_str):
    """Convert IPMA UTC timestamp to Lisbon local time (handles DST)."""
    dt = datetime.datetime.strptime(ts_str, "%Y-%m-%dT%H:%M")
    year = dt.year
    march31 = datetime.datetime(year, 3, 31)
    dst_start = march31 - datetime.timedelta(days=(march31.weekday() + 1) % 7)
    oct31 = datetime.datetime(year, 10, 31)
    dst_end = oct31 - datetime.timedelta(days=(oct31.weekday() + 1) % 7)
    offset = 1 if dst_start <= dt < dst_end else 0
    return (dt + datetime.timedelta(hours=offset)).strftime("%Y-%m-%dT%H:%M")

# ---------------------------------------------------------------------------
# STEP 1 — Find nearest station
# ---------------------------------------------------------------------------
print("Finding nearest station to {}...".format(TARGET_NAME))
stations = fetch_json(IPMA_STATIONS_URL)

best_id   = None
best_name = None
best_dist = float("inf")

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

# Build list of all 24 expected hours from the API timestamps
all_timestamps = sorted(obs_data.keys())

new_records = []
missing_hours = []

for ts in all_timestamps:
    sd = obs_data[ts].get(best_id)
    if sd is None:
        missing_hours.append(ts)
        continue
    new_records.append({
        "Timestamp (UTC)":    ts,
        "Timestamp (Lisbon)": utc_to_lisbon(ts),
        "Station":            best_name,
        "Temperature (°C)":   clean(sd.get("temperatura")),
        "Humidity (%)":       clean(sd.get("humidade")),
        "Precipitation (mm)": clean(sd.get("precAcumulada")),
        "Wind Speed (km/h)":  clean(sd.get("intensidadeVentoKM")),
        "Wind Speed (m/s)":   clean(sd.get("intensidadeVento")),
        "Wind Direction":     WIND_DIR.get(sd.get("idDireccVento"), ""),
        "Pressure (hPa)":     clean(sd.get("pressao")),
        "Radiation (W/m²)":   clean(sd.get("radiacao")),
    })

print("{} records fetched.".format(len(new_records)))
if missing_hours:
    print("WARNING — {} hours with no station data (station gap): {}".format(
        len(missing_hours), ", ".join(missing_hours)
    ))

if not new_records:
    print("No data returned — skipping Excel update.")
    exit(0)

# ---------------------------------------------------------------------------
# STEP 3 — Append to (or create) master Excel file
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", start_color="2F5496", end_color="2F5496")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
DATA_FONT   = Font(name="Arial", size=10)
ALT_FILL    = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")
GAP_FILL    = PatternFill("solid", start_color="FFE699", end_color="FFE699")  # yellow for gaps
GAP_FONT    = Font(name="Arial", size=10, italic=True, color="7F6000")
CENTER      = Alignment(horizontal="center", vertical="center")
LEFT        = Alignment(horizontal="left",   vertical="center")
THIN_BORDER = Border(
    bottom=Side(style="thin", color="B8CCE4"),
    right =Side(style="thin", color="B8CCE4"),
)
COL_WIDTHS  = [20, 20, 22, 16, 14, 18, 18, 16, 16, 14, 16]

def style_row(ws, row_idx, record, is_gap=False):
    fill = GAP_FILL if is_gap else (ALT_FILL if (row_idx % 2 == 0) else None)
    font = GAP_FONT if is_gap else DATA_FONT
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        c = ws.cell(row=row_idx, column=col_idx, value=record.get(col_name))
        c.font = font; c.alignment = CENTER; c.border = THIN_BORDER
        if fill: c.fill = fill

if os.path.exists(EXCEL_FILE):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Hourly Data"]

    # Collect existing timestamps to avoid duplicates
    existing_ts = set()
    for row in ws.iter_rows(min_row=4, max_col=1, values_only=True):
        if row[0]:
            existing_ts.add(row[0])

    next_row = ws.max_row + 1
    appended = 0
    skipped  = 0

    # Append real records
    for record in new_records:
        if record["Timestamp (UTC)"] in existing_ts:
            skipped += 1
            continue
        style_row(ws, next_row, record)
        next_row += 1
        appended += 1

    # Append gap rows (yellow) for missing hours
    for ts in missing_hours:
        if ts in existing_ts:
            continue
        style_row(ws, next_row, {
            "Timestamp (UTC)": ts,
            "Station": best_name,
            "Temperature (°C)": "NO DATA",
        }, is_gap=True)
        next_row += 1

    print("{} rows appended, {} duplicates skipped.".format(appended, skipped))

else:
    print("Creating new Excel file...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hourly Data"

    # Title
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value     = "IPMA Climate Data — {} — Continuous Log".format(TARGET_NAME)
    c.font      = Font(name="Arial", bold=True, size=13, color="1F3864")
    c.alignment = CENTER
    c.fill      = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
    ws.row_dimensions[1].height = 24

    # Subtitle
    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value     = (
        "Source: IPMA open-data API  |  Station: {} (ID: {})  |  {:.1f} km from {}  |"
        "  Auto-updated daily via GitHub Actions  |  Yellow rows = station reported no data"
    ).format(best_name, best_id, best_dist, TARGET_NAME)
    c.font      = Font(name="Arial", size=9, italic=True, color="595959")
    c.alignment = LEFT

    # Headers
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        c = ws.cell(row=3, column=col_idx, value=col_name)
        c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER
    ws.row_dimensions[3].height = 20

    # Data rows
    for row_idx, record in enumerate(new_records, start=4):
        style_row(ws, row_idx, record)

    # Gap rows
    for ts in missing_hours:
        row_idx = ws.max_row + 1
        style_row(ws, row_idx, {
            "Timestamp (UTC)": ts,
            "Station": best_name,
            "Temperature (°C)": "NO DATA",
        }, is_gap=True)

    # Column widths
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"

wb.save(EXCEL_FILE)
print("Saved to {}".format(EXCEL_FILE))
print("Done!")
