# myair2fit

Import CPAP sleep data from a [myAir](https://myair.resmed.com/) export into [Fitbit](https://www.fitbit.com/) via the Sleep API.

## Prerequisites

1. **Python 3.10+**
2. **Fitbit Developer App** — register at https://dev.fitbit.com/apps/new with:
   - **OAuth 2.0 Application Type**: Personal
   - **Redirect URL**: `http://localhost:8080/callback`
   - **Default Access Type**: Read & Write

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env and set FITBIT_CLIENT_ID to your OAuth 2.0 Client ID
```

## Export myAir Data

1. Log in to https://myair.resmed.com/
2. Go to **Settings → Privacy → Export My Data**
3. Download the ZIP file

## Usage

```bash
# Dry run — preview what will be imported
python myair2fit.py --dry-run "C:\path\to\export.zip"

# Import all records
python myair2fit.py "C:\path\to\export.zip"

# Import only records in a date range
python myair2fit.py --start-date 2025-09-01 --end-date 2025-11-30 "C:\path\to\export.zip"

# Use a custom sleep start time (default: 22:30)
python myair2fit.py --start-time 23:00 "C:\path\to\export.zip"
```

On first run, a browser window will open for Fitbit authorization. After granting access, tokens are cached in `.fitbit_tokens.json` and refreshed automatically.

## How It Works

- Reads `SLEEP_RECORD.csv` from the myAir export (ZIP, directory, or CSV)
- For each session: `date` from `SESSION_DATE`, `duration` from `USAGE_HOURS × 3,600,000 ms`
- Since myAir doesn't export start times, defaults to 10:30 PM (override with `--start-time`)
- POSTs to `https://api.fitbit.com/1.2/user/-/sleep.json`
