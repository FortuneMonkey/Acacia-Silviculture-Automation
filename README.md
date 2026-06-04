# Trial Plan Generator

A Streamlit dashboard to automatically generate Trial Plan Excel files from your Trial Register database.

---

## Quick Start

### 1. Prerequisites
- Python 3.9 or newer

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

Your browser will open at `http://localhost:8501`

---

## How to Use

1. **Upload** your database Excel file (must have a sheet named **Trial Register** with headers on row 3)
2. **Filter** — choose a specific year or "All" to include all years
3. **Preview** the list of trials that will be generated (only `Not yet Register` status)
4. **Click Generate** — a ZIP file is created and ready to download
5. **Extract** the ZIP — inside you'll find a `Trial Plan/` folder with one Excel per trial

---

## File Naming

Each generated Excel is named:
```
{Code Trial} ({Sector}{Comp}).xlsx
```
Example: `ACFER267 (BYSK086).xlsx`

The sheet inside the Excel is named the same way.

---

## Database Requirements

| Column | Index | Description |
|--------|-------|-------------|
| Code Trial | 3 | Trial code (used in filename) |
| Sector | 5 | Sector code |
| Comp | 6 | Compartment code |
| Year | 63 | Year (used for year filter) |
| Register | 70 | Must contain `Not yet Register` to be included |

Headers must be on **row 3**. Data starts from **row 4**.

---

## Coming Soon
- Schedule generator
- Treat generator
