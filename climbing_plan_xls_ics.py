from datetime import datetime, timedelta, date
import matplotlib.pyplot as plt
import xlsxwriter

# === CONFIGURATION ===

# Set the date of the first Monday of week 19
week_start = date(2025, 5, 5)

# Define the range of weeks in the plan
weeks = list(range(19, 33))  # Week 19 to Week 32 inclusive

# Set week types for each week (training, deload, or send)
week_types = {w: 'training' for w in weeks}
for w in [23, 24, 27, 30]:
    week_types[w] = 'deload'
for w in [31, 32]:
    week_types[w] = 'send'

# Assign a background color for each week type (for Excel output)
week_type_colors = {
    'Training': '#FFFFFF',
    'Deload': '#FFE599',
    'Send': '#C9F3C2'
}

# Define all session types with metadata: label, group, days scheduled, time, and duration (min)
session_definitions = {
    'F':  {'name': 'Glutes + Core', 'group': 'Conditioning', 'days': ['Wednesday', 'Saturday'], 'time': '18:00', 'duration': 30},
    'R':  {'name': 'Run', 'group': 'Conditioning', 'days': ['Wednesday'], 'time': '07:00', 'duration': 60},
    'V1': {'name': 'Max Route', 'group': 'Lead', 'days': ['Sunday'], 'time': '10:00', 'duration': 240},
    'V2': {'name': 'Doublettes', 'group': 'Lead', 'days': ['Tuesday'], 'time': '18:00', 'duration': 120},
    'V3': {'name': '4X4', 'group': 'Bouldering', 'days': ['Tuesday'], 'time': '18:00', 'duration': 120},
    'E1': {'name': 'Finger Endurance', 'group': 'Fingers', 'days': ['Friday'], 'time': '18:00', 'duration': 45},
    'E2': {'name': 'Low Intensity Endurance', 'group': 'Fingers', 'days': ['Friday'], 'time': '19:00', 'duration': 30},
    'D1': {'name': 'Max Finger Strength', 'group': 'Fingers', 'days': ['Friday'], 'time': '20:00', 'duration': 45},
    'B1': {'name': 'Max Bouldering', 'group': 'Bouldering', 'days': ['Thursday'], 'time': '18:00', 'duration': 120},
}

# Convert durations (in minutes) to hours for plotting
session_weights = {k: v['duration'] / 60 for k, v in session_definitions.items()}

# Mark sessions that should appear bolded in Excel
bold_sessions = {'V1', 'D1'}

# Define session group ordering and colors for Excel formatting
group_order = ['Conditioning', 'Lead', 'Fingers', 'Bouldering']
group_colors = {
    'Conditioning': '#D9EAD3',
    'Lead': '#FCE5CD',
    'Fingers': '#F4CCCC',
    'Bouldering': '#CFE2F3'
}

# Define session assignments per week as text (comma-separated session codes)
weekly_sessions_text = {
    19: "F, F, V1, V2, E1, E2, D1, R",
    20: "F, F, V1, V3, E1, E2, D1, R",
    21: "F, F, V1, V2, E1, E2, D1, R",
    22: "F, F, V1, V3, E1, E2, D1, R",
    23: "F, F, V1, B1, R",
    24: "F, F, V1, B1, R",
    25: "F, F, V1, V2, E1, E2, B1, R",
    26: "F, F, V1, V3, E1, E2, B1, R",
    27: "F, F, V1, B1, R",
    28: "F, F, V1, V2, E2, B1, R",
    29: "F, F, V1, V3, E2, B1, R",
    30: "F, F, V1, B1, R",
    31: "",
    32: "",
}

# Map weekday names to numerical values for offsetting from Monday
weekday_map = {
    'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
    'Friday': 4, 'Saturday': 5, 'Sunday': 6
}

# === EXPORT FUNCTIONS ===

def parse_time(tstr):
    return datetime.strptime(tstr, "%H:%M").time()

def collect_events():
    events = []
    week_load = {}
    for week in weeks:
        week_load[week] = 0
        sessions = [s.strip() for s in weekly_sessions_text[week].split(",") if s.strip()]
        for s in sessions:
            if s not in session_definitions:
                continue
            info = session_definitions[s]
            for day in info['days']:
                event_date = week_start + timedelta(weeks=week-19, days=weekday_map[day])
                start = datetime.combine(event_date, parse_time(info['time']))
                end = start + timedelta(minutes=info['duration'])
                events.append({
                    'summary': info['name'],
                    'start': start,
                    'end': end,
                    'week': week,
                    'duration': info['duration'],
                    'type': week_types[week]
                })
                week_load[week] += info['duration']
    return events, week_load

def export_ics(events, filename="climbing_plan.ics"):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//ClimbingPlan//EN"]
    for e in events:
        lines += [
            "BEGIN:VEVENT",
            f"UID:{e['summary'].replace(' ', '')}-{e['start'].strftime('%Y%m%dT%H%M')}",
            f"DTSTAMP:{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART;TZID=Europe/Paris:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND;TZID=Europe/Paris:{e['end'].strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['summary']}",
            f"DESCRIPTION:Week {e['week']} - {e['type']}",
            "END:VEVENT"
        ]
    lines.append("END:VCALENDAR")
    with open(filename, "w") as f:
        f.write("\n".join(lines))
    print(f"✅ .ics calendar saved: {filename}")

def plot_weekly_load(week_load):
    # Calculate per-group contributions
    group_durations = {w: {g: 0 for g in group_order} for w in weeks}
    for week in weeks:
        sessions = [s.strip() for s in weekly_sessions_text[week].split(",") if s.strip()]
        for s in sessions:
            if s in session_definitions:
                group = session_definitions[s]['group']
                group_durations[week][group] += session_definitions[s]['duration'] / 60

    # Prepare stacked data
    labels = [f"W{w}" for w in weeks]
    fig, ax = plt.subplots(figsize=(10, 5))
    bottom = [0] * len(weeks)
    for group in group_order:
        values = [group_durations[w][group] for w in weeks]
        ax.bar(labels, values, bottom=bottom, label=group, color=group_colors[group])
        bottom = [bottom[i] + values[i] for i in range(len(weeks))]

    ax.set_title("Weekly Training Load (by group, hours)")
    ax.set_ylabel("Total Hours")
    ax.legend()
    ax.grid(axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig("training_load_chart.png")
    print("✅ Training load chart saved as: training_load_chart.png")


def export_excel(week_load):
    wb = xlsxwriter.Workbook("climbing_training_plan_clean.xlsx")
    ws = wb.add_worksheet("Plan")
    def fmt(**kwargs): return wb.add_format({**kwargs, 'font_name': 'Arial'})

    start_dates = [(week_start + timedelta(weeks=w - 19)).strftime('%-d-%b') for w in weeks]
    week_types_list = [week_types[w].capitalize() for w in weeks]
    loads = [f"{week_load[w]/60:.1f} hours" for w in weeks]

    ws.merge_range(0, 0, 0, len(weeks), "Climbing Training Plan", fmt(bold=True, font_size=14, align='center', bg_color='#DDEBF7'))

    info = [("Weeks", [f"Week {w}" for w in weeks]), ("Start date", start_dates), ("Type", week_types_list), ("Load", loads)]
    for i, (label, values) in enumerate(info, start=1):
        ws.write(i, 0, label, fmt(bold=True))
        for j, val in enumerate(values):
            style = fmt()
            if label == "Type":
                style = fmt(italic=True, bg_color=week_type_colors.get(val, "#FFFFFF"))
            ws.write(i, j+1, val, style)

    # Grouped data rows
    row = 5
    written = set()
    for group in group_order:
        color = group_colors[group]
        ws.write(row, 0, group, fmt(bold=True, bg_color=color, bottom=1))
        for col in range(1, len(weeks)+1):
            ws.write_blank(row, col, None, fmt(bg_color=color))
        row += 1
        for code, info in session_definitions.items():
            if info['group'] != group or code in written:
                continue
            written.add(code)
            bold = code in bold_sessions
            ws.write(row, 0, info['name'], fmt(bg_color=color, bold=bold))
            for j, w in enumerate(weeks):
                if code in weekly_sessions_text[w]:
                    ws.write(row, j+1, info['name'], fmt(bg_color=color))
            row += 1

    ws.set_column(0, 0, 35)
    ws.set_column(1, len(weeks), 18)
    ws.freeze_panes(5, 1)
    wb.close()
    print("✅ Excel file saved: climbing_plan.xlsx")

if __name__ == "__main__":
    events, week_load = collect_events()
    export_ics(events)
    plot_weekly_load(week_load)
    export_excel(week_load)