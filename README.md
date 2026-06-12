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
