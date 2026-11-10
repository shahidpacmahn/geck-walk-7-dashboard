from datetime import date, datetime
from html import escape
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


WORKBOOK = Path(__file__).with_name("week1group.xlsx")
LOGO = Path(__file__).with_name("logo.jpeg")
WEEK_DATES = (date(2026, 11, 6), date(2026, 11, 7), date(2026, 11, 8))
DATE_LABELS = ("06 Nov", "07 Nov", "08 Nov")
STEP_CAP = 15_000
TEAM_EMOJIS = {
    "BO10": "🦅",
    "HIIT": "⚡",
    "HMT": "🐯",
    "KS6": "🌟",
    "MH": "🏹",
    "SST": "🦩",
    "WW": "🐺",
    "MG": "🐉",
    "PDK": "🛡️",
    "SV": "🌸",
    "WLT": "🦋",
    "RFT": "🔥",
    "GOG": "🚀",
    "LONG": "🌿",
    "SMT": "🐢",
}


def normalize_code(value):
    return str(value or "").strip().upper()


def parse_date_header(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 30_000 <= value <= 60_000:
        return from_excel(value).date()
    if isinstance(value, str):
        for date_format in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(value.strip(), date_format).date()
            except ValueError:
                continue
    return None


def number_or_zero(value):
    if value is None:
        return 0
    try:
        return max(0, int(float(value)))
    except (TypeError, ValueError, OverflowError):
        return 0


def load_dashboard_data(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        member_sheet = workbook.worksheets[0]
        if "Sheet2" not in workbook.sheetnames:
            raise ValueError("The workbook is missing the Sheet2 team lookup.")
        team_sheet = workbook["Sheet2"]

        member_rows = member_sheet.iter_rows(values_only=True)
        member_headers = next(member_rows, None)
        if not member_headers:
            raise ValueError("The first worksheet is empty.")
        member_columns = {
            str(value).strip(): index
            for index, value in enumerate(member_headers)
            if value is not None
        }
        required_member_columns = {"First Name", "Last Name"}
        missing_member_columns = required_member_columns - member_columns.keys()
        if missing_member_columns:
            raise ValueError(
                "The first worksheet is missing required columns: "
                + ", ".join(sorted(missing_member_columns))
            )

        date_columns = {}
        for index, header in enumerate(member_headers):
            parsed_date = parse_date_header(header)
            if parsed_date in WEEK_DATES:
                date_columns[parsed_date] = index
        missing_dates = [day.strftime("%d/%m/%Y") for day in WEEK_DATES if day not in date_columns]
        if missing_dates:
            raise ValueError(
                "Could not find these daily step columns in the first worksheet: "
                + ", ".join(missing_dates)
            )

        team_rows = team_sheet.iter_rows(values_only=True)
        team_headers = next(team_rows, None)
        team_columns = {
            str(value).strip(): index
            for index, value in enumerate(team_headers or ())
            if value is not None
        }
        required_team_columns = {"Code", "Name", "Points"}
        missing_team_columns = required_team_columns - team_columns.keys()
        if missing_team_columns:
            raise ValueError(
                "Sheet2 is missing required columns: "
                + ", ".join(sorted(missing_team_columns))
            )

        teams_by_code = {}
        for row in team_rows:
            code = normalize_code(row[team_columns["Code"]])
            if not code:
                continue
            if code in teams_by_code:
                raise ValueError(f"Duplicate team code in Sheet2: {code}")
            try:
                points = float(row[team_columns["Points"]])
            except (TypeError, ValueError, IndexError):
                raise ValueError(f"Team {code} must have a numeric points value.") from None
            teams_by_code[code] = {
                "code": code,
                "name": str(row[team_columns["Name"]] or "").strip(),
                "points": points,
                "members": [],
            }

        for row in member_rows:
            name = str(row[member_columns["First Name"]] or "").strip()
            code = normalize_code(row[member_columns["Last Name"]])
            if not name or not code:
                continue
            if code not in teams_by_code:
                raise ValueError(f"Team code {code} from the first worksheet is missing in Sheet2.")
            steps = [number_or_zero(row[date_columns[day]]) for day in WEEK_DATES]
            member = {
                "name": name,
                "code": code,
                "steps": steps,
                "week_steps": sum(steps),
            }
            teams_by_code[code]["members"].append(member)

        members = []
        for team in teams_by_code.values():
            if not team["members"]:
                continue
            for member in team["members"]:
                member["team"] = team["name"]
                member["emoji"] = TEAM_EMOJIS.get(team["code"], "🥾")
                members.append(member)
            capped_daily_averages = [
                sum(min(member["steps"][day_index], STEP_CAP) for member in team["members"])
                / len(team["members"])
                for day_index in range(len(WEEK_DATES))
            ]
            team["members_count"] = len(team["members"])
            team["daily_averages"] = capped_daily_averages
            team["capped_average"] = sum(capped_daily_averages) / len(capped_daily_averages)
            team["emoji"] = TEAM_EMOJIS.get(team["code"], "🥾")

        return list(teams_by_code.values()), members
    finally:
        workbook.close()


def render_table(headers, rows, numeric_columns=()):
    header_html = "".join(f"<th>{escape(str(header))}</th>" for header in headers)
    body_rows = []
    for row in rows:
        cells = []
        for index, value in enumerate(row):
            css_class = ' class="number"' if index in numeric_columns else ""
            cells.append(f"<td{css_class}>{escape(str(value))}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    st.markdown(
        '<div class="table-wrap"><table class="leaderboard">'
        f"<thead><tr>{header_html}</tr></thead><tbody>{''.join(body_rows)}</tbody></table></div>",
        unsafe_allow_html=True,
    )


def daily_steps_chart(daily_totals):
    width, height = 720, 340
    left, right, top, bottom = 64, 24, 34, 54
    chart_height = height - top - bottom
    chart_width = width - left - right
    max_value = max(daily_totals, default=0) or 1
    scale_max = ((max_value + 4999) // 5000) * 5000 or 5000
    pieces = [
        f'<svg viewBox="0 0 {width} {height}" role="img" aria-label="Total community steps by day">'
    ]
    for tick in range(5):
        value = scale_max * tick / 4
        y = top + chart_height * (1 - tick / 4)
        pieces.append(
            f'<line x1="{left}" y1="{y:.1f}" x2="{width-right}" y2="{y:.1f}" class="grid"/>'
            f'<text x="{left-10}" y="{y+4:.1f}" class="axis" text-anchor="end">'
            f'{value/1000:,.0f}k</text>'
        )
    bar_width = min(112, chart_width / 5)
    colors = ("#ff9b54", "#2dd4a7", "#70a5ff")
    for index, value in enumerate(daily_totals):
        x = left + chart_width * (index + 0.5) / len(daily_totals) - bar_width / 2
        bar_height = chart_height * value / scale_max
        y = top + chart_height - bar_height
        pieces.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_width:.1f}" height="{bar_height:.1f}" '
            f'rx="9" fill="{colors[index]}"/>'
            f'<text x="{x+bar_width/2:.1f}" y="{max(top-8, y-10):.1f}" '
            f'class="value" text-anchor="middle">{value/1_000_000:.2f}m</text>'
            f'<text x="{x+bar_width/2:.1f}" y="{height-18}" class="label" '
            f'text-anchor="middle">{DATE_LABELS[index]}</text>'
        )
    pieces.append("</svg>")
    return "".join(pieces)


def team_daily_chart(team_board):
    top_teams = team_board[:6]
    width, height = 760, 360
    left, right, top, bottom = 122, 24, 62, 34
    chart_width = width - left - right
    row_height = (height - top - bottom) / max(len(top_teams), 1)
    colors = ("#ff9b54", "#2dd4a7", "#70a5ff")
    pieces = [
        f'<svg viewBox="0 0 {width} {height}" role="img" '
        'aria-label="Capped daily average steps for the top six teams">'
    ]
    for index, label in enumerate(DATE_LABELS):
        x = left + index * 190
        pieces.append(
            f'<rect x="{x}" y="22" width="10" height="10" rx="3" fill="{colors[index]}"/>'
            f'<text x="{x+16}" y="31" class="legend">{label}</text>'
        )
    for tick in (0, 5000, 10000, 15000):
        x = left + chart_width * tick / STEP_CAP
        pieces.append(
            f'<line x1="{x:.1f}" y1="{top-7}" x2="{x:.1f}" y2="{height-bottom}" class="grid"/>'
            f'<text x="{x:.1f}" y="{height-9}" class="axis" text-anchor="middle">'
            f'{tick/1000:.0f}k</text>'
        )
    for row_index, team in enumerate(top_teams):
        center_y = top + row_height * (row_index + 0.5)
        pieces.append(
            f'<text x="{left-12}" y="{center_y+4:.1f}" class="label" text-anchor="end">'
            f'{escape(team["code"])}</text>'
        )
        for day_index, value in enumerate(team["daily_averages"]):
            bar_y = center_y - 13 + day_index * 9
            bar_width = chart_width * value / STEP_CAP
            pieces.append(
                f'<rect x="{left}" y="{bar_y:.1f}" width="{bar_width:.1f}" height="6" '
                f'rx="3" fill="{colors[day_index]}"/>'
            )
    pieces.append("</svg>")
    return "".join(pieces)


st.set_page_config(page_title="GECK WALK 7.0", page_icon="🚶", layout="wide")
st.markdown(
    """
    <style>
    .stApp { background: radial-gradient(ellipse at 12% 0%, #142c2b 0%, #0c141c 40%, #080d13 100%); color: #e8f1f3; }
    [data-testid="stHeader"] { background: #080d13; }
    [data-testid="stToolbar"] { right: 1rem; }
    .block-container { max-width: 1260px; padding: 1rem 0.7rem 1.5rem; }
    h1, h2, h3, p, label { color: #e8f1f3; }
    [data-testid="stCaptionContainer"] p { color: #a4b7bc; }
    .hero {
        padding: 1rem 1rem; border-radius: 22px; color: white; min-height: unset;
        background: linear-gradient(115deg, #123b37, #15584c 58%, #18765f);
        border: 1px solid #347a67; box-shadow: 0 14px 36px #0000003a;
        margin-bottom: 0.9rem; display: flex; flex-direction: column; justify-content: center;
    }
    .hero h1 { margin: 0; font-size: 1.8rem; color: #f5fffb; }
    .hero p { margin: .4rem 0 0; color: #c1e8d9; font-size: 0.92rem; }
    .metric-card {
        background: linear-gradient(145deg, #16242d, #111b23); border: 1px solid #293a43;
        border-radius: 18px; padding: 0.9rem 0.95rem; box-shadow: 0 8px 22px #00000024;
        min-height: 0;
        margin-bottom: 0.75rem;
    }
    .metric-label { color: #a4b7bc; font-size: .8rem; font-weight: 600; }
    .metric-value { color: #f4fbf8; font-size: 1.2rem; font-weight: 750; margin-top: .2rem; }
    .chart-card {
        background: linear-gradient(145deg, #131f27, #101820); border: 1px solid #293a43;
        border-radius: 20px; padding: 0.8rem; box-shadow: 0 8px 22px #00000024;
        height: 100%;
        margin-bottom: 0.75rem;
    }
    .chart-title { color: #edf6f2; font-size: 1rem; font-weight: 700; margin-bottom: .15rem; }
    .chart-subtitle { color: #9eb1b6; font-size: .76rem; margin-bottom: .6rem; }
    .chart-card svg { display: block; width: 100%; height: auto; overflow: visible; }
    .chart-card .grid { stroke: #2a3941; stroke-width: 1; }
    .chart-card .axis { fill: #8fa4aa; font-size: 11px; }
    .chart-card .label { fill: #d7e5e3; font-size: 12px; font-weight: 600; }
    .chart-card .legend { fill: #b6c8c8; font-size: 12px; }
    .chart-card .value { fill: #f3fbf7; font-size: 12px; font-weight: 700; }
    .table-wrap {
        overflow-x: auto; border: 1px solid #293a43; border-radius: 16px;
        background: #111a22; box-shadow: 0 8px 22px #00000024;
    }
    table.leaderboard { width: 100%; border-collapse: collapse; }
    table.leaderboard th {
        background: #1b2b32; color: #bfe9d7; font-size: .84rem;
        text-align: left; padding: .7rem .8rem; white-space: nowrap;
    }
    table.leaderboard td {
        color: #dce8e7; padding: .7rem .8rem; border-top: 1px solid #25343c;
        white-space: nowrap;
    }
    table.leaderboard tr:nth-child(even) td { background: #142028; }
    table.leaderboard td.number { text-align: right; font-variant-numeric: tabular-nums; }
    div[data-testid="stTabs"] button { font-weight: 650; color: #b7c8ca; }
    div[data-testid="stTabs"] button[aria-selected="true"] { color: #57d9a8; }
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"] { background: #57d9a8; }
    hr { border-color: #293a43; }
    [data-testid="stHorizontalBlock"] { display: block !important; width: 100% !important; }
    [data-testid="stHorizontalBlock"] > div {
        width: 100% !important; max-width: 100% !important;
        padding-right: 0 !important; padding-left: 0 !important;
        margin-bottom: 0.7rem;
        text-align: center;
        display: flex;
        justify-content: center;
        align-items: center;
    }
    [data-testid="stImage"] { display: block; max-width: 120px; margin: 0 auto 0.5rem; }
    @media (min-width: 768px) {
        .block-container { padding-top: 1.6rem; padding-left: 1rem; padding-right: 1rem; }
        .hero { padding: 1.1rem 1.5rem; min-height: 120px; margin-bottom: 1.2rem; }
        .hero h1 { font-size: 2.3rem; }
        .hero p { font-size: 1.02rem; }
        .metric-card { min-height: 102px; margin-bottom: 0; }
        .metric-value { font-size: 1.45rem; }
        .chart-card { padding: 1.15rem 1.25rem; margin-bottom: 0; }
        .chart-title { font-size: 1.08rem; }
        .chart-subtitle { font-size: .82rem; }
        [data-testid="stHorizontalBlock"] {
            display: flex !important;
            width: 100% !important;
        }
        [data-testid="stHorizontalBlock"] > div {
            width: auto !important;
            max-width: none !important;
            padding-right: 0.5rem !important;
            padding-left: 0.5rem !important;
            margin-bottom: 0;
            text-align: left;
            display: block;
            justify-content: flex-start;
            align-items: stretch;
        }
        table.leaderboard th, table.leaderboard td { padding: .8rem 1rem; white-space: nowrap; }
        [data-testid="stImage"] { display: block; max-width: 170px; margin: 0; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
logo_column, title_column = st.columns([1.25, 5], vertical_alignment="center")
with logo_column:
    if LOGO.exists():
        st.image(str(LOGO), width=170)
    else:
        st.warning(f"Logo not found: {LOGO.name}")
with title_column:
    st.markdown(
        '<div class="hero"><h1>🚶 GECK WALK 7.0</h1>'
        '<p>One Month. One alumni community. Let’s see who’s walking the talk.</p></div>',
        unsafe_allow_html=True,
    )

if not WORKBOOK.exists():
    st.error(f"Workbook not found: {WORKBOOK}")
    st.stop()

try:
    teams, walkers = load_dashboard_data(WORKBOOK)
except (ValueError, KeyError, OSError) as exc:
    st.error(f"Unable to load the Week 1 workbook: {exc}")
    st.stop()

points_board = sorted(teams, key=lambda team: (-team["points"], -team["capped_average"]))
average_board = sorted(teams, key=lambda team: (-team["capped_average"], -team["points"]))
walker_board = sorted(walkers, key=lambda member: -member["week_steps"])
leader = points_board[0]
best_average = average_board[0]
daily_totals = [
    sum(member["steps"][day_index] for member in walkers)
    for day_index in range(len(WEEK_DATES))
]
best_day_index = max(range(len(daily_totals)), key=daily_totals.__getitem__)
average_per_walker_day = sum(daily_totals) / (len(walkers) * len(WEEK_DATES)) if walkers else 0

metric_columns = st.columns(4)
metrics = [
    ("👟 Total steps · 3 days", f"{sum(daily_totals):,}"),
    ("📍 Avg per walker / day", f"{average_per_walker_day:,.0f}"),
    ("🔥 Most active day", f"{DATE_LABELS[best_day_index]} · {daily_totals[best_day_index]:,}"),
    ("🏆 Top team · points", f"{leader['emoji']} {leader['code']} · {leader['points']:,.0f} pts"),
]
for column, (label, value) in zip(metric_columns, metrics):
    with column:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">{escape(label)}</div>'
            f'<div class="metric-value">{escape(value)}</div></div>',
            unsafe_allow_html=True,
        )

st.write("")
chart_columns = st.columns([1, 1.15])
with chart_columns[0]:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">🌍 Steps across the community</div>'
        '<div class="chart-subtitle">Total recorded steps each day</div>'
        f'{daily_steps_chart(daily_totals)}</div>',
        unsafe_allow_html=True,
    )
with chart_columns[1]:
    st.markdown(
        '<div class="chart-card"><div class="chart-title">📈 Team daily averages</div>'
        '<div class="chart-subtitle">Top 6 teams · each walker capped at 15,000 steps/day</div>'
        f'{team_daily_chart(average_board)}</div>',
        unsafe_allow_html=True,
    )

st.write("")
tabs = st.tabs(["🏆 Team points", "📊 Team average", "👟 Walker leaderboard"])

with tabs[0]:
    st.subheader(f"{leader['emoji']} {leader['name']} is setting the pace!")
    st.caption("Week 1 team points are taken from the Sheet2 team-code lookup.")
    rows = [
        [
            rank,
            f"{team['emoji']} {team['name']}",
            team["code"],
            f"{team['points']:,.0f}",
            team["members_count"],
            f"{team['capped_average']:,.0f} steps",
        ]
        for rank, team in enumerate(points_board, start=1)
    ]
    render_table(["Rank", "Team", "Code", "Points", "Walkers", "3-day capped avg"], rows, (0, 3, 4, 5))

with tabs[1]:
    st.subheader(f"{best_average['emoji']} {best_average['name']} leads on capped average")
    st.caption("Each walker’s daily steps are capped at 15,000 before averaging across the team.")
    rows = [
        [
            rank,
            f"{team['emoji']} {team['name']}",
            team["code"],
            f"{team['capped_average']:,.0f} steps",
            *(f"{average:,.0f}" for average in team["daily_averages"]),
        ]
        for rank, team in enumerate(average_board, start=1)
    ]
    render_table(["Rank", "Team", "Code", "3-day avg", *DATE_LABELS], rows, (0, 3, 4, 5, 6))

with tabs[2]:
    st.subheader("Every step counts")
    st.caption("Ranked by total uncapped steps across 06, 07, and 08 November.")
    rows = [
        [
            rank,
            member["name"],
            f"{member['emoji']} {member['team']}",
            f"{member['week_steps']:,} steps",
        ]
        for rank, member in enumerate(walker_board, start=1)
    ]
    render_table(["Rank", "Walker", "Team", "3-day steps"], rows, (0, 3))

st.caption("GECK WALK 7.0 · Week 1 · 06–08 November 2026")
