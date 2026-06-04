<div align="center">

# 🌳 Acacia Silviculture — Trial Tools Automation

**Transform your Trial Register database into ready-to-use Trial, Schedule & Treatment Plan files — instantly, and beautifully.**

![Python](https://img.shields.io/badge/Python-3.9+-3776AB?style=flat&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=flat&logo=streamlit&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat&logo=pandas&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-22c55e?style=flat)

</div>

---

A **Streamlit** web application built for **Acacia Silviculture R&D** that automates the generation of trial planning documents directly from a Trial Register Excel database. Upload your file, filter by year, and download neatly bundled Excel files in seconds.

---

## ✨ Features

- 📤 **One-click upload** — Reads directly from your `Trial Register` Excel sheet
- 🔎 **Smart filtering** — Automatically detects records marked *"Not yet Register"* and filters by year
- 📊 **Live metrics** — Instant overview of total records, pending registrations, and matches
- 📋 **Trial Plan generator** — One Excel per trial with all plan fields mapped from the database
- 📅 **Schedule Plan generator** — Lists every activity with a scheduled date, per trial
- 🧪 **Treatment Plan generator** — Parses treatments from design notes into structured rows
- 📦 **Bundled downloads** — All generated files packaged into a single ZIP per tool
- 🎨 **Elegant UI** — Clean, modern, light-themed interface with smooth animations

---

## 🖥️ Demo

| Upload & Configure | Generate & Download |
|--------------------|---------------------|
| Upload your database, filter by year, preview matched trials | Generate Trial, Schedule, or Treatment plans and download as ZIP |

> 💡 _Tip: Add screenshots once deployed — create an `assets/` folder and reference images like `![Demo](assets/screenshot.png)`_

---

## 📂 Project Structure

```
trial_plan_generator/
├── .streamlit/
│   └── config.toml          # Theme & app configuration
├── app.py                   # Main Streamlit application (UI)
├── trial_plan.py            # Trial Plan logic & Excel builder
├── schedule_plan.py         # Schedule Plan logic & Excel builder
├── treatment_plan.py        # Treatment Plan logic & Excel builder
├── requirements.txt         # Python dependencies
├── .gitignore
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.9+**
- **pip** (Python package manager)

### Installation

**1. Clone the repository**
```bash
git clone https://github.com/FortuneMonkey/Acacia-Silviculture-Automation.git
cd Acacia-Silviculture-Automation
```

**2. (Optional) Create a virtual environment**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

**3. Install dependencies**
```bash
pip install -r requirements.txt
```

**4. Run the app**
```bash
streamlit run app.py
```

**5.** Open your browser at **`http://localhost:8501`** 🎉

---

## 📋 Usage

1. **Upload** your Trial Register Excel file via the sidebar.
   - The file **must** contain a sheet named **`Trial Register`**
   - Column **headers must be on row 3**
2. **Filter** records by year (or leave as *All*).
3. **Preview** the matched trials in the expandable table.
4. **Select a tab** (Trial / Schedule / Treatment Plan).
5. Click **Generate**, then **Download** the ZIP file.

> ⚠️ Only records with a **Register** status of *"Not yet Register"* / *"Not yet Registered"* are processed.

---

## 📑 Required Columns

Your `Trial Register` sheet must include the following columns (headers on **row 3**):

| Column | Description |
|--------|-------------|
| `Code Trial` | Unique trial code |
| `Sector` | Sector identifier |
| `Comp` | Compartment identifier |
| `Year` | Trial year |
| `Register` | Registration status |
| `Established by` | Person who established the trial |

Additional columns are resolved automatically by each module for plan fields, activities, and treatments.

---

## 🧩 How Each Tool Works

### 📋 Trial Plan
Generates one Excel file per trial, with all plan fields mapped automatically from your database columns.

### 📅 Schedule Plan
Lists only activities that have a scheduled date.

**Output columns:**
`activity_group` · `activity_code` · `plan_date` · `do_date` · `update_date` · `required` · `note` · `ass_program_by` · `ass_staff_by`

- `required` = `1` always
- `ass_staff_by` ← **Established by**
- Dates use Afrikaans locale `[$-0436]YYYY-MM-DD`

### 🧪 Treatment Plan
Parses treatments from the **Design And Treatment** column (falls back to **Title** if empty).

**Output columns:**
`treat_no` · `treat_lo` · `treat_name` · `treat_type` · `germplasmid` · `deployid` · `femaleid` · `maleid` · `provenance` · `species` · `bri` · `iri`

**Parsing rules:**

| Pattern | Action |
|---------|--------|
| Lines starting with `T1.` `T2.` … | Extracted as individual treatments (prefix stripped) |
| Lines starting with `-` | Extracted as individual treatments (dash stripped) |
| Neither pattern found | Entire cell text used as one treatment |
| Design And Treatment empty | Falls back to **Title** column |

- **Fixed values:** `treat_type` = `Treatment`
- **Species mapping:** `AC` → `ACRA` · `AH` → `AHYB`
- **Spacing parsing:** `3 x 2.3 m` → `bri = 3`, `iri = 2.3`

---

## 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| **[Streamlit](https://streamlit.io/)** | Web UI framework |
| **[openpyxl](https://openpyxl.readthedocs.io/)** | Excel reading & writing |
| **[pandas](https://pandas.pydata.org/)** | Data handling & previews |

---

## ☁️ Deployment (Streamlit Community Cloud)

1. Push this repository to GitHub (already done ✅)
2. Go to **[share.streamlit.io](https://share.streamlit.io/)**
3. Click **New app** → select this repository
4. Set the main file path to **`app.py`**
5. Click **Deploy** 🚀

Your app will be live at a public URL you can share with your team.

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!

1. Fork the project
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📜 License

This project is maintained for **Acacia Silviculture R&D**.
_(Add a `LICENSE` file if you wish to make it open-source — e.g., MIT.)_

---

## 👤 Author

**FortuneMonkey**
🔗 [GitHub Repository](https://github.com/FortuneMonkey/Acacia-Silviculture-Automation)

---

<div align="center">

**Trial Tools Automation** • Acacia Silviculture R&D 🌳

_Built with ❤️ and Streamlit_

</div>