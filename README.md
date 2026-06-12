# Climate Data — Guimarães
Automated daily fetch of hourly climate data from IPMA for the nearest station to Guimarães (ID: 1210625).
## How it works
A GitHub Actions workflow runs every day at **midnight Lisbon time** and:
1. Fetches the last 24h of hourly observations from the IPMA open-data API
2. Appends new rows to `data/climate_guimaraes.csv`
3. Commits and pushes the updated file automatically

On the **first run**, if no CSV exists yet but the old `data/climate_guimaraes.xlsx` is present, the script migrates all historical data from the xlsx into the new CSV so nothing is lost. After that, the xlsx is no longer updated and remains as a frozen backup.
## How the data feeds the Revit plugin
The CSV is read by the pyRevit "Counting Hours" button, which computes climate-risk counters per material and writes them into the Revit model. Pull the latest CSV (e.g. with GitHub Desktop) before running the button so it uses current data.
## Setup instructions
### 1. Create a new GitHub repository
- Go to github.com → New repository
- Name it e.g. `climate-guimaraes`
- Set it to **Private** (recommended for thesis data)
- Do **not** add a README (you'll push one)
### 2. Upload these files
Push this folder to your new repo:
### 3. Enable Actions permissions
- Go to your repo → Settings → Actions → General
- Under "Workflow permissions" → select **Read and write permissions**
- Click Save
### 4. Run it manually the first time
- Go to Actions tab → "Fetch Climate Data" → Run workflow
- After ~30 seconds, `data/climate_guimaraes.csv` will appear in your repo
- On this first run, the log will show `migrated N historical rows` if an existing xlsx was found
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

Hours where the station reported no data appear as rows marked **NO DATA**, so gaps stay visible in the record.
## Source
Data provided by [IPMA](https://api.ipma.pt) — Instituto Português do Mar e da Atmosfera.
