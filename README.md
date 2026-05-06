# 3M Cartridge Service Life Tool v2.0

Automates the 3M SLS website to calculate cartridge service life and download PDF reports.

## Setup (one-time)

1. Install Python: https://www.python.org/downloads/
   - During install, check **"Add Python to PATH"**
2. Double-click `run.bat` — first run downloads Chromium (~200MB automatically)

## Fill in tasks.csv

Open `tasks.csv` with Excel or any text editor. Each row = one calculation job.

| Field | Required | Notes |
|-------|----------|-------|
| name | No | Display name for the chemical |
| cas | Yes | CAS number, used to search the website |
| exposure | Yes | Numeric value |
| unit | Yes | `ppm` or `mg/m³` |
| breakthrough | No | Leave blank if not needed |
| cartridge | Yes | e.g. `6001CN` |
| respirator_type | No | `可更换式` (default) or `PAPR` |
| humidity | Yes | `<65` or `>=65` |
| temp_c | Yes | Fixed options only: 10 / 15 / 20 / 25 / 30 / 35 / 40 |
| work_intensity | Yes | `light` / `moderate` / `heavy` |
| pressure_atm | No | Default: `1` |

## Run

Double-click `run.bat`

- Press **Enter** → silent background mode (screen not occupied)
- Type **v** then Enter → visual mode (browser visible, useful for first-run debugging)

## Output

All files saved in `output/` folder:
- One PDF per task
- `summary_YYYYMMDD_HHMMSS.csv` — service life values for all tasks
- `ERROR_xxx.png` — screenshot on failure (send to Claude for debugging)
