# MGM Data Wrangler (`MGM_converter2.py`)

This script reads daily MGM report Excel files and combines them into one cleaned workbook:

- **Input folder**: `D:\Thesis\Thesis_data_examples`
- **Output file**: `D:\Thesis\Thesis_data_examples\MGM_Station_Data_Cleaned.xlsx`

---

## 1) What input files are expected

Place MGM daily parameter files (`.xlsx`) in the data folder.  
Typical filenames look like:

- `20250722794F-Günlük Ortalama Sıcaklık (°C).xlsx`
- `20250722794F-Günlük Toplam Yağış (mm=kg÷m²) OMGİ .xlsx`
- `20250722794F-Günlük Maksimum Rüzgar Yönü (°) ve Hızı  (m÷sn).xlsx`

The script recognizes parameter files by normalized Turkish filename text.  
Temporary Excel lock files (`~$*.xlsx`) are ignored automatically.

---

## 2) Installation

Use Python 3.9+ and install dependencies:

```bash
pip install pandas openpyxl
```

---

## 3) How to run

From terminal:

```bash
python "c:\Users\serha\Susurluk\Last_Script\MGM_converter2.py"
```

The script will:

1. Scan `DATA_DIR` for recognized MGM `.xlsx` files.
2. Parse all station/year blocks from each file.
3. Merge parameters by station and date.
4. Write the cleaned Excel workbook.

---

## 4) Input data layout (inside each MGM file)

The parser expects each data block in sheet **`Report`** with this structure:

- Header row: contains `Yıl: YYYY` and `İstasyon Adı/No: STATION/ID`
- Row +2: parameter title
- Row +3: month headers (`1..12`)
- Rows +4..+34: day rows (`1..31`)
- Month values are read from columns `2..13`

Invalid calendar dates (for example, 30 Feb) are skipped.

---

## 5) Output workbook format

Output file: `MGM_Station_Data_Cleaned.xlsx`

- Sheet **`_SUMMARY`**
  - Columns:
    - `Station_Name`
    - `Station_ID`
    - `Sheet_Name`
    - `Date_Start`
    - `Date_End`
    - `N_Records`
    - `Parameters`

- One sheet per station
  - Sheet name is Excel-safe and max 31 characters
  - First column: `Date` (`YYYY-MM-DD`)
  - Other columns: available parameters for that station

Wind parameter is split into:

- `Max_Wind_Dir_deg`
- `Max_Wind_Speed_mps`

---

## 6) Parameter name mapping in output

Recognized input files are mapped to these output columns:

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

## 7) Notes and troubleshooting

- If a file is shown as unrecognized, add or adjust its phrase in `PARAM_MAP` in `MGM_converter2.py`.
- Make sure source files are not open/locked by Excel while running.
- The script only parses sheet name `Report`; different sheet names are skipped by design.
