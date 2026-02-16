# MGM Data Wrangler

Parse Turkish Meteorological Service (MGM) daily Excel reports and combine them into one cleaned workbook with one sheet per station.

This README is for:

- `MGM_converter2_MGM Data Wrangler_Standalone Script.py`

---

## Download and run (any PC)

1. **Get the script**
   - Copy `MGM_converter2_MGM Data Wrangler_Standalone Script.py` into your project folder.

2. **Install Python**
   - Use Python 3.9 or newer.

3. **Install required packages**
   ```bash
   pip install pandas openpyxl
   ```

4. **Set your input folder**
   - Open the script and change:
   ```python
   DATA_DIR = r"C:\path\to\your\mgm_data_folder"
   ```

5. **Run**
   ```bash
   python "MGM_converter2_MGM Data Wrangler_Standalone Script.py"
   ```

The output workbook is created in `DATA_DIR` as:

- `MGM_Station_Data_Cleaned.xlsx`

---

## What the script does

- Scans `DATA_DIR` for MGM Excel files (`.xlsx`)
- Ignores temporary Excel lock files (`~$*.xlsx`)
- Detects parameter type from Turkish filename text using `PARAM_MAP`
- Parses all station/year blocks from sheet `Report`
- Merges parameters by `Date` for each station
- Splits maximum wind values into:
  - `Max_Wind_Dir_deg`
  - `Max_Wind_Speed_mps`
- Writes one summary sheet and one station sheet per station

---

## Expected input files

Place MGM daily parameter files (`.xlsx`) in the data folder.

Typical file naming patterns:

- `...Günlük Ortalama Sıcaklık (°C).xlsx`
- `...Günlük Toplam Yağış (mm=kg÷m²) OMGİ .xlsx`
- `...Günlük Maksimum Rüzgar Yönü (°) ve Hızı (m÷sn).xlsx`

If a file is not recognized by `PARAM_MAP`, the script skips it and prints a warning.

---

## Expected worksheet structure (inside each MGM file)

The parser reads sheet name **`Report`** and expects each station/year block to follow:

- Header row includes:
  - `Yıl: YYYY`
  - `İstasyon Adı/No: STATION_NAME/ID`
- Row +2: parameter title
- Row +3: month headers (`1..12`)
- Rows +4..+34: day rows (`1..31`)
- Values are read from month columns `2..13`

Invalid dates (for example 30 Feb) are skipped automatically.

---

## Output workbook format

Output file:

- `MGM_Station_Data_Cleaned.xlsx`

### `_SUMMARY` sheet columns

- `Station_Name`
- `Station_ID`
- `Sheet_Name`
- `Date_Start`
- `Date_End`
- `N_Records`
- `Parameters`

### Station sheets

- One sheet per station (Excel-safe name, max 31 chars)
- Column `Date` formatted as `YYYY-MM-DD`
- Remaining columns are available parameters for that station

---

## Output parameter columns

- `Precip_OMGI_mm`
- `Precip_Manuel_mm`
- `Tmax_C`
- `Tmin_C`
- `Tmean_C`
- `Cloud_Cover_Okta`
- `Storm_Day`
- `Strong_Wind_Storm_Day`
- `Max_Wind_Dir_deg`
- `Max_Wind_Speed_mps`
- `SWE_Manuel_mm`
- `Snow_Depth_Manuel_cm`
- `Snow_Depth_Mean_cm`
- `New_Snow_Mean_cm`
- `Open_Evap_mm`
- `ET_mm`

---

## Troubleshooting

- **No recognized files found**
  - Check `DATA_DIR` path and confirm files are `.xlsx`.
- **Unrecognized filename**
  - Add or adjust keyword mapping in `PARAM_MAP`.
- **Excel file locked**
  - Close open source files in Excel before running.
- **No data parsed from a file**
  - Confirm worksheet name is exactly `Report`.

---

## Notes

- The script normalizes Turkish filename text before matching.
- Output is designed for analysis-ready station-by-date tables.
- If MGM report naming/layout changes, update parsing rules and `PARAM_MAP`.
