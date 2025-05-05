from datetime import date, timedelta
import pandas as pd

# === Weekly sessions ===
weekly_sessions = {
    19: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'D1', 'R'],
    20: ['F', 'F', 'V1', 'V3', 'E1', 'E2', 'D1', 'R'],
    21: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'D1', 'R'],
    22: ['F', 'F', 'V1', 'V3', 'E1', 'E2', 'D1', 'R'],
    23: ['F', 'F', 'V1', 'B', 'R'],
    24: ['F', 'F', 'V1', 'B'],
    25: ['F', 'F', 'V1', 'V2', 'E1', 'E2', 'B', 'R'],
    26: ['F', 'F', 'V1', 'V3', 'E1', 'E2', 'B', 'R'],
    27: ['F', 'F', 'V1', 'B', 'R'],
    28: ['F', 'F', 'V1', 'V2', 'E2', 'B', 'R'],
    29: ['F', 'F', 'V1', 'V3', 'E2', 'B', 'R'],
    30: ['F', 'F', 'V1', 'B'],
    31: [],
    32: [],
}

# === Week types ===
week_types = {week: 'training' for week in range(19, 33)}
for w in [23, 24, 27, 30]:
    week_types[w] = 'deload'
for w in [31, 32]:
    week_types[w] = 'send'

# === Session metadata ===
session_names = {
    'F': 'Glutes + Core',
    'V1': 'Max Route',
    'V2': 'Doublettes',
    'V3': '4x4',
    'E1': 'Finger Endurance at Home',
    'E2': 'Base Intensity Finger Endurance',
    'D1': 'Max Finger Strength',
    'B': 'Max Bouldering',
    'R': 'Run'
}

session_weights = {
    'V1': 4,
    'V2': 2,
    'V3': 2,
    'B': 2,
    'E1': 0.75,
    'E2': 0.5,
    'D1': 0.75,
    'R': 1,
    'F': 0.5
}

session_groups = {
    'V1': 'lead',
    'V2': 'lead',
    'V3': 'lead',
    'B':  'boulder',
    'E1': 'fingers',
    'E2': 'fingers',
    'D1': 'fingers',
    'R':  'conditioning',
    'F':  'conditioning'
}

group_colors = {
    'lead': '#FCE4D6',
    'boulder': '#D9E1F2',
    'fingers': '#E2EFDA',
    'conditioning': '#FFF2CC'
}

# === Prepare DataFrame ===
cycles = {}
load = {}
session_rows = {session_names[s]: {} for s in session_names}

for week in range(19, 33):
    monday = date(2025, 5, 5) + timedelta(weeks=week - 19)
    col_label = f"Week {week} ({monday.strftime('%b %d')})"
    sessions = weekly_sessions[week]

    cycles[col_label] = week_types[week].capitalize()
    load[col_label] = sum(session_weights.get(s, 0) for s in sessions)

    for s in session_names:
        session_rows[session_names[s]][col_label] = session_names[s] if s in sessions else ""

# === Build DataFrame ===
df = pd.DataFrame([cycles, load], index=["CYCLES", "TOTAL LOAD"])
for session_label, values in session_rows.items():
    df.loc[session_label] = values

# === Export to Excel ===
output_file = "training_schedule_matrix_layout.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name="Schedule")
    workbook = writer.book
    worksheet = writer.sheets["Schedule"]

    # Write week info row (row 1)
    for col_num, w in enumerate(range(19, 33), start=1):
        monday = date(2025, 5, 5) + timedelta(weeks=w - 19)
        week_label = f"Week {w}\n{monday.strftime('%b %d')}"
        worksheet.write(1, col_num, week_label, workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial',
            'bold': True, 'bg_color': '#EEECE1'
        }))

    # Write week type row (row 2)
    for col_num, w in enumerate(range(19, 33), start=1):
        week_type = week_types[w].capitalize()
        worksheet.write(2, col_num, week_type, workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial',
            'italic': True, 'bg_color': '#DDDDEE'
        }))

    # Title row
    title = "CLIMBING TRAINING PLAN OVERVIEW"
    worksheet.write(0, 0, title, workbook.add_format({
        'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#DDEBF7', 'font_name': 'Arial'
    }))
    worksheet.merge_range(0, 0, 0, len(df.columns), title,
                          workbook.add_format({'bold': True, 'font_size': 14,
                                               'align': 'center', 'valign': 'vcenter',
                                               'bg_color': '#DDEBF7', 'font_name': 'Arial'}))

    # Freeze panes below header
    worksheet.freeze_panes(3, 1)

    # Set column widths
    for col_num in range(len(df.columns) + 1):
        worksheet.set_column(col_num, col_num, 22)

    # Format rows
    formats = {
        "CYCLES": workbook.add_format({'bold': True, 'bg_color': '#FFFF99', 'font_name': 'Arial'}),
        "TOTAL LOAD": workbook.add_format({'bg_color': '#F4A460', 'font_name': 'Arial'}),
    }

    for row_idx, row_label in enumerate(df.index.values, start=1):
        if row_label in formats:
            fmt = formats[row_label]
        else:
            code = next((k for k, v in session_names.items() if v == row_label), None)
            group = session_groups.get(code)
            bg = group_colors.get(group)
            fmt = workbook.add_format({'bg_color': bg, 'font_name': 'Arial'}) if bg else workbook.add_format({'font_name': 'Arial'})
        worksheet.set_row(row_idx + 2, None, fmt)  # +2 for title and week info rows

print(f"✅ Excel file saved as: {output_file}")