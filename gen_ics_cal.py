from datetime import datetime, timedelta, date
import pandas as pd
import matplotlib.pyplot as plt

# === CONFIGURATION ===

week_start = date(2025, 5, 5)
weeks = list(range(19, 33))
week_types = {w: 'training' for w in weeks}
for w in [23, 24, 27, 30]: week_types[w] = 'deload'
for w in [31, 32]: week_types[w] = 'send'

session_definitions = {
    'F':  {'name': 'Glutes + Core', 'days': ['Wednesday', 'Saturday'], 'time': '18:00', 'duration': 30, 'group': 'Conditioning'},
    'R':  {'name': 'Run', 'days': ['Wednesday'], 'time': '07:00', 'duration': 60, 'group': 'Conditioning'},
    'V1': {'name': 'Max Route', 'days': ['Sunday'], 'time': '10:00', 'duration': 240, 'group': 'Lead'},
    'V2': {'name': 'Doublettes', 'days': ['Tuesday'], 'time': '18:00', 'duration': 120, 'group': 'Lead'},
    'V3': {'name': '4X4', 'days': ['Tuesday'], 'time': '18:00', 'duration': 120, 'group': 'Bouldering'},
    'E1': {'name': 'Finger Endurance', 'days': ['Friday'], 'time': '18:00', 'duration': 45, 'group': 'Fingers'},
    'E2': {'name': 'Low Intensity Endurance', 'days': ['Friday'], 'time': '19:00', 'duration': 30, 'group': 'Fingers'},
    'D1': {'name': 'Max Finger Strength', 'days': ['Friday'], 'time': '20:00', 'duration': 45, 'group': 'Fingers'},
    'B1': {'name': 'Max Bouldering', 'days': ['Thursday'], 'time': '18:00', 'duration': 120, 'group': 'Bouldering'},
}

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

weekday_map = {
    'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3,
    'Friday': 4, 'Saturday': 5, 'Sunday': 6
}

# === COLLECT CALENDAR EVENTS AND WEEK LOAD ===

def parse_time(tstr): return datetime.strptime(tstr, "%H:%M").time()

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
            date_of_event = week_start + timedelta(weeks=week-19, days=weekday_map[day])
            start_dt = datetime.combine(date_of_event, parse_time(info['time']))
            end_dt = start_dt + timedelta(minutes=info['duration'])
            events.append({
                'summary': info['name'],
                'start': start_dt,
                'end': end_dt,
                'week': week,
                'duration': info['duration'],
                'type': week_types[week]
            })
            week_load[week] += info['duration']

# === EXPORT CALENDAR TO ICS ===

def create_ics(events, filename="climbing_plan.ics"):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//ClimbingPlan//EN"]
    for e in events:
        lines += [
            "BEGIN:VEVENT",
            f"UID:{e['summary'].replace(' ','')}-{e['start'].strftime('%Y%m%dT%H%M')}",
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

# === PLOT WEEKLY LOAD ===

def plot_weekly_load(week_load):
    plt.figure(figsize=(10, 5))
    labels = [f"W{w}" for w in week_load.keys()]
    loads = [v/60 for v in week_load.values()]  # convert to hours
    plt.bar(labels, loads, color='skyblue')
    plt.title("Weekly Training Load (hours)")
    plt.ylabel("Hours")
    plt.grid(True, axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig("training_load_chart.png")
    print("✅ Training load chart saved as: training_load_chart.png")

# === EXECUTE ===
if __name__ == "__main__":
    create_ics(events)
    plot_weekly_load(week_load)