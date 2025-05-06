# 🧗 Climbing Training Plan Generator

This project creates a structured, multi-format training plan for climbers. It outputs:

- ✅ A **Google Calendar (.ics)** file with scheduled sessions  
- ✅ A clean, color-coded **Excel file** grouped by training type  
- ✅ A **stacked bar chart** of weekly training load by type  

---

## 📦 Features

- Plan spans multiple weeks
- Supports weekly types: `Training`, `Deload`, and `Send`
- Categorizes sessions into:
  - 🟢 Conditioning
  - 🟡 Lead
  - 🔴 Fingers
  - 🔵 Bouldering
- Color-coded Excel export
- Google Calendar `.ics` export
- Stacked bar plot of weekly load
- The current exmple is a 12 weeks plan aiming to send 8a/13b in Ceuse this summer 

---

## 📂 Output Files

| File                                  | Description                               |
|---------------------------------------|-------------------------------------------|
| `climbing_plan.ics`                   | Importable calendar file (Google/iCal)    |
| `climbing_training_plan_clean.xlsx`   | Formatted weekly Excel plan               |
| `training_load_chart.png`             | Visual bar chart of weekly load           |

---

## 📓 Jupyter Notebook

The notebook `climbing_training_plan.ipynb` is included for convenience.

You can use it to:

- 🧠 Explore or modify the plan interactively  
- 🧮 Recalculate training load week by week  
- 🗂️ Export all formats (`.ics`, `.xlsx`, `.png`) in one click

### ✅ Run it with:

- [Google Colab](https://colab.research.google.com) — no installation needed  
- JupyterLab or VS Code with Python 3.11+

---

## 🚀 Getting Started

### 1. Install dependencies

```bash
pip install pandas matplotlib xlsxwriter
