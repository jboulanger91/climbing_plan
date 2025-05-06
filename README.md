# 🧗 Climbing Training Plan Generator

This project generates a structured weekly training plan for climbers. It outputs:

- ✅ A **Google Calendar (.ics)** file of scheduled sessions
- ✅ A clean **Excel plan** with colored categories (Conditioning, Lead, etc.)
- ✅ A **stacked bar chart** showing weekly load by training type

## 📦 Features

- Plan spans **Weeks 19–32**
- Supports **Training**, **Deload**, and **Send** weeks
- Categorizes sessions into:
  - 🟢 Conditioning
  - 🟡 Lead
  - 🔴 Fingers
  - 🔵 Bouldering
- Visualizes session distribution with colors and stacked plots
- Export formats:
  - `climbing_plan.ics`
  - `climbing_training_plan_clean.xlsx`
  - `training_load_chart.png`

## 🚀 Getting Started

### Install dependencies:
Tested with Python 3.11
pip install matplotlib xlsxwriter