# 🧗‍♂️ Climbing Training Plan Generator

This project generates a **structured Excel training schedule** for climbing athletes. It helps you organize your weekly sessions with categories, load estimates, and formatting to visualize your training blocks.

## ✅ Features

- Weekly plan spanning **weeks 19–32**
- **Session types**: lead climbing, bouldering, finger training, conditioning
- **Training types**: training, deload, send
- **Color-coded rows** based on session category
- **Week header rows** with both the week number and starting date
- **Excel output** that is clean, formatted, and compatible with Google Sheets

## 📂 Output

- `training_schedule_matrix_layout.xlsx`: an Excel file showing:
  - Session names by week
  - Weekly load
  - Week type (training/deload/send)
  - Visual grouping by training category

## 📦 Requirements

- Python 3.8+
- `pandas`
- `xlsxwriter`

Install dependencies:
```bash
pip install pandas xlsxwriter
