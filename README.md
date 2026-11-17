# 🚶 GECK WALK 7.0 Dashboard

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python) ![Streamlit](https://img.shields.io/badge/Streamlit-1.x-FF4B4B?style=for-the-badge&logo=streamlit) ![Excel](https://img.shields.io/badge/Data-Workbook-217346?style=for-the-badge&logo=microsoftexcel) ![Status](https://img.shields.io/badge/Status-Ready-00C853?style=for-the-badge)

A sleek, dark-mode leaderboard for the GECK WALK 7.0 challenge — built to surface team momentum, daily totals, and individual step performance in one polished dashboard.

## ✨ What this dashboard shows

- Community-wide step totals across the challenge window
- Team leaderboard rankings by points and capped average steps
- Week 1 and Week 2 summary tables
- Daily totals and activity trends
- Individual walker rankings across all six days
- A presentation-ready layout with a modern sports-style design

## 🧭 Project structure

```text
dashboard/
├── main.py              # Streamlit dashboard app
├── requirements.txt     # Python dependencies
├── logo.jpeg            # Brand/logo image used in the UI
├── week1group.xlsx      # Source Excel workbook with challenge data
├── README.md            # Project documentation
└── .venv/               # Local virtual environment (optional)
```

## 🚀 Quick start

### 1) Create a virtual environment

```bash
cd dashboard
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

macOS/Linux:

```bash
source .venv/bin/activate
```

### 2) Install dependencies

```bash
pip install -r requirements.txt
```

### 3) Run the dashboard

```bash
streamlit run main.py
```

Then open the local URL shown in the terminal.

## 📊 Data requirements

The app expects the workbook file `week1group.xlsx` to be present in the same folder as `main.py` and to include these worksheets:

- `Total`
- `SW1`
- `SW2`
- `Sheet2`

It also validates required column names and headers before rendering the dashboard. If the workbook is missing or structurally different, the app shows a clear error instead of crashing.

## 🏆 Features in the app

- Read and validate Excel data automatically
- Rank teams by total points and capped daily averages
- Summarize team performance per week with daily averages and points
- Highlight the most active day and top-performing team
- Display a walker leaderboard by total steps
- Present everything in a dark, high-contrast theme for GitHub/demo-friendly viewing

## 🛠 Troubleshooting

- `Workbook not found` → confirm that `week1group.xlsx` exists in the `dashboard/` folder
- `Unable to load the walking workbook` → check that all required sheets and columns are present
- Missing packages → run `pip install -r requirements.txt`

## 📌 Notes

- The dashboard is designed for internal reporting and challenge presentations
- Team and walker calculations are driven directly from the workbook data
- Step totals use the workbook logic already implemented in the app, including capped average behavior

## ⚖ License

This project is intended for internal use and is not currently published as a public package.

---

Made for the GECK WALK community. Keep moving. Keep competing. Keep walking. 🚶‍♂️🚶‍♀️
