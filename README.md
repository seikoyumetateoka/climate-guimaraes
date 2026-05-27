# Climate Data — Guimarães

Automated daily fetch of hourly climate data from IPMA for the nearest station to Guimarães (ID: 1210625).

## How it works

A GitHub Actions workflow runs every day at **midnight Lisbon time** and:
1. Fetches the last 24h of hourly observations from the IPMA open-data API
2. Appends new rows to `data/climate_guimaraes.xlsx`
3. Commits and pushes the updated file automatically

## Setup instructions

### 1. Create a new GitHub repository
- Go to github.com → New repository
- Name it e.g. `climate-guimaraes`
- Set it to **Private** (recommended for thesis data)
- Do **not** add a README (you'll push one)

### 2. Upload these files
Push this folder to your new repo:
```
climate-guimaraes/
├── .github/
│   └── workflows/
│       └── fetch_climate.yml   ← the automation schedule
├── data/                        ← Excel file will appear here
├── fetch_climate.py             ← the fetch script
└── README.md
```

### 3. Enable Actions permissions
- Go to your repo → Settings → Actions → General
- Under "Workflow permissions" → select **Read and write permissions**
- Click Save

### 4. Run it manually the first time
- Go to Actions tab → "Fetch Climate Data" → Run workflow
- After ~30 seconds, `data/climate_guimaraes.xlsx` will appear in your repo

## Data fields
| Column | Description |
|--------|-------------|
| Timestamp (UTC) | Hour of observation |
| Station | Station name |
| Temperature (°C) | Air temperature |
| Humidity (%) | Relative humidity |
| Precipitation (mm) | Accumulated precipitation |
| Wind Speed (km/h) | Wind speed |
| Wind Speed (m/s) | Wind speed (m/s) |
| Wind Direction | Cardinal direction |
| Pressure (hPa) | Atmospheric pressure |
| Radiation (W/m²) | Solar radiation |

## Source
Data provided by [IPMA](https://api.ipma.pt) — Instituto Português do Mar e da Atmosfera.
