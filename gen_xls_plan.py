from datetime import date, timedelta
import xlsxwriter

# --- Sessions ---
session_data = {
    'F':  ('Glutes + Core', 'Conditioning'),
    'R':  ('Run', 'Conditioning'),
    'V1': ('Max Route', 'Lead'),
    'V2': ('Doublettes', 'Lead'),
    'E1': ('Finger Endurance', 'Fingers'),
    'E2': ('Low Intensity Endurance', 'Fingers'),
    'D1': ('Max Finger Strength', 'Fingers'),
    'B1': ('Max Bouldering', 'Bouldering'),
    'B2': ('4X4', 'Bouldering'),
}

session_weights = {
    'F': 0.5, 'R': 1, 'V1': 4, 'V2': 2,
    'E1': 0.75, 'E2': 0.5, 'D1': 0.75, 'B1': 2, 'B2': 2,
}

bold_sessions = {'V1', 'D1'}

group_order = ['Conditioning', 'Lead', 'Fingers', 'Bouldering']
group_colors = {
    'Conditioning': '#D9EAD3',
    'Lead': '#FCE5CD',
    'Fingers': '#F4CCCC',
    'Bouldering': '#CFE2F3',
}

# --- Weekly sessions ---
weekly_sessions = {
    19: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'D1', 'R'],
    20: ['F', 'F', 'V1', 'B2', 'E1', 'E2', 'D1', 'R'],
    21: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'D1', 'R'],
    22: ['F', 'F', 'V1', 'B2', 'E1', 'E2', 'D1', 'R'],
    23: ['F', 'F', 'V1', 'B1', 'R'],
    24: ['F', 'F', 'V1', 'B1', 'R'],
    25: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'B1', 'R'],
    26: ['F', 'F', 'V1', 'B2', 'E1', 'E2', 'B1', 'R'],
    27: ['F', 'F', 'V1', 'B1', 'R'],
    28: ['F', 'F', 'V1', 'V2', 'E2', 'B1', 'R'],
    29: ['F', 'F', 'V1', 'B2', 'E2', 'B1', 'R'],
    30: ['F', 'F', 'V1', 'B1', 'R'],
    31: [],
    32: [],
}

# --- Week type coloring ---
week_types = {w: 'training' for w in range(19, 33)}
for w in [23, 24, 27, 30]: week_types[w] = 'deload'
for w in [31, 32]: week_types[w] = 'send'

week_type_colors = {
    'Training': '#FFFFFF',
    'Deload': '#FFE599',
    'Send': '#C9F3C2'
}

# --- Build headers ---
weeks = list(range(19, 33))
week_labels = [f"Week {w}" for w in weeks]
start_dates = [(date(2025, 5, 5) + timedelta(weeks=w - 19)).strftime('%-d-%b') for w in weeks]
week_types_list = [week_types[w].capitalize() for w in weeks]
loads = [f"{sum(session_weights.get(s, 0) for s in weekly_sessions[w]):.1f} hours" for w in weeks]

# --- Output Excel ---
workbook = xlsxwriter.Workbook("climbing_training_plan_clean.xlsx")
worksheet = workbook.add_worksheet("Plan")

def fmt(**kwargs):
    return workbook.add_format({**kwargs, 'font_name': 'Arial'})

# --- Title ---
worksheet.merge_range(0, 0, 0, len(weeks), "Climbing Training Plan", fmt(bold=True, font_size=14, align='center', bg_color='#DDEBF7'))

# --- Header info rows ---
info_rows = [
    ("Weeks", week_labels),
    ("Start date", start_dates),
    ("Type", week_types_list),
    ("Load", loads)
]

for i, (label, values) in enumerate(info_rows):
    worksheet.write(i+1, 0, label, fmt(bold=True))
    for j, val in enumerate(values):
        cell_fmt = fmt()
        if label == "Type":
            cell_fmt.set_bg_color(week_type_colors.get(val, "#FFFFFF"))
            cell_fmt.set_italic()
        worksheet.write(i+1, j+1, val, cell_fmt)

# --- Write session data ---
row = 5
for group in group_order:
    group_fmt = workbook.add_format({
        'font_name': 'Arial',
        'bold': True,
        'bg_color': group_colors[group],
        'bottom': 1  # thin bottom border
    })

    worksheet.write(row, 0, group, group_fmt)
    for col in range(1, len(weeks)+1):
        worksheet.write_blank(row, col, None, group_fmt)
    row += 1

    written = set()
    for code, (name, grp) in session_data.items():
        if grp != group or code in written:
            continue
        written.add(code)

        # Row label
        label_fmt = fmt(bg_color=group_colors[group])
        if code in bold_sessions:
            label_fmt.set_bold()
        worksheet.write(row, 0, name, label_fmt)

        # Per-week content
        for col, w in enumerate(weeks):
            if code in weekly_sessions[w]:
                cell_fmt = fmt(bg_color=group_colors[group])
                worksheet.write(row, col+1, name, cell_fmt)
        row += 1

# --- Format sheet ---
worksheet.set_column(0, 0, 35)
worksheet.set_column(1, len(weeks), 18)
worksheet.freeze_panes(5, 1)

workbook.close()
print("✅ File saved as: climbing_training_plan_clean.xlsx")
