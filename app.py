"""PTPM Proposal Panel Appointment Dashboard.

The source adapter is separate from transformation so another source (such as a
private Google Sheet) can later replace Excel without changing the dashboard views.
"""

from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st


APP_TITLE = "PTPM Proposal Panel Appointment Dashboard"
DATA_SOURCES = {
    2025: {
        "path": Path(__file__).with_name("public_data_2025.xlsx"),
        "worksheet": "Nama Pelajar & Panel 2025.",
    },
    2026: {
        "path": Path(__file__).with_name("public_data_2026.xlsx"),
        "worksheet": "Nama Pelajar & Panel 2026",
    },
}
INTERNAL_CODES = ["WAJ", "CKT", "INU", "NAM", "NJ", "MBM", "SNR", "MTA", "MS", "NR", "NMB", "ABN"]
OPTIONAL_INTERNAL_CODES = ["BIE"]
CATEGORY_INTERNAL = "Internal PTPM"
CATEGORY_EXTERNAL = "External"
THEME_PURPLE = "#5B2C83"
THEME_PURPLE_DARK = "#3D1A5B"
THEME_PURPLE_LIGHT = "#F3ECF8"
THEME_ORANGE = "#E67E22"
THEME_ORANGE_LIGHT = "#FFF1E3"

# Panel recommendation configuration. These values can be adjusted without
# changing the recommendation functions or dashboard layout.
RECOMMENDATION_WEIGHTS = {
    "Overall Workload": 0.45,
    "Supervisor-Specific Use": 0.30,
    "Recent Workload": 0.15,
    "Supervisor Breadth": 0.10,
}
RECENT_APPOINTMENT_DAYS = 90
RECOMMENDATION_STATUS_THRESHOLDS = {
    "Highly recommended": 75,
    "Recommended": 55,
    "Consider": 35,
    "Lower priority": 0,
}
RECOMMENDATION_LECTURER_CODES = list(INTERNAL_CODES)
LECTURER_FULL_NAMES = {code: "" for code in RECOMMENDATION_LECTURER_CODES}


st.set_page_config(page_title=APP_TITLE, page_icon="🎓", layout="wide")


@st.cache_data(show_spinner=False)
def load_source_data(path: str, worksheet_name: str, reload_token: int = 0) -> tuple[pd.DataFrame, pd.Timestamp]:
    """Load the Excel grid. ``reload_token`` explicitly invalidates the cache."""
    del reload_token
    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(
            f"Excel data file '{source.name}' is missing from the project folder."
        )
    try:
        grid = pd.read_excel(
            source,
            sheet_name=worksheet_name,
            header=None,
            engine="openpyxl",
            dtype=object,
        )
    except ValueError as exc:
        raise ValueError(
            f"Worksheet '{worksheet_name}' was not found in '{source.name}'."
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Excel file '{source.name}' could not be read: {exc}") from exc
    return grid, pd.Timestamp.now()


def clean_label(value: object) -> str:
    """Normalize worksheet labels for reliable matching."""
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def locate_excel_layout(grid: pd.DataFrame) -> dict[str, int]:
    """Detect the two-row header and map required fields to column positions."""
    header_row = None
    for index in range(min(20, len(grid))):
        labels = {clean_label(value) for value in grid.iloc[index].tolist()}
        if {"NO.", "TARIKH", "NAMA PELAJAR"}.issubset(labels):
            header_row = index
            break
    if header_row is None or header_row + 1 >= len(grid):
        raise ValueError("Could not identify the Excel header rows containing No., Tarikh and Nama Pelajar.")

    upper = [clean_label(value) for value in grid.iloc[header_row].tolist()]
    lower = [clean_label(value) for value in grid.iloc[header_row + 1].tolist()]

    def column_for(label: str, rows: list[list[str]]) -> int:
        for row in rows:
            if label in row:
                return row.index(label)
        raise ValueError(f"Required Excel column '{label}' was not found.")

    supervisor_label = "SV" if "SV" in lower else "NAMA SV"
    layout = {
        "header_row": header_row,
        "data_start": header_row + 2,
        "number": column_for("NO.", [upper]),
        "date": column_for("TARIKH", [upper]),
        "student": column_for("NAMA PELAJAR", [upper]),
        "supervisor": column_for(supervisor_label, [lower, upper]),
        "panel_role": column_for("PANEL", [upper]),
        "external": column_for("LUAR", [lower]),
    }
    for code in INTERNAL_CODES:
        layout[code] = column_for(code, [lower])
    layout["internal_codes"] = list(INTERNAL_CODES)
    for code in OPTIONAL_INTERNAL_CODES:
        if code in lower:
            layout[code] = lower.index(code)
            layout["internal_codes"].append(code)
    return layout


def validate_source(df: pd.DataFrame) -> list[str]:
    """Return validation errors that prevent a reliable transformation."""
    errors = []
    if df.empty:
        errors.append("The Excel worksheet contains no records.")
    else:
        try:
            locate_excel_layout(df)
        except ValueError as exc:
            errors.append(str(exc))
    return errors


def is_appointment(value: object) -> bool:
    """Interpret common spreadsheet representations of an appointment flag."""
    return str(value).strip().lower() in {"1", "1.0", "yes", "y", "true"}


def transform_to_long(raw: pd.DataFrame, proposal_year: int) -> pd.DataFrame:
    """Convert repeated Excel student blocks into one row per appointment."""
    layout = locate_excel_layout(raw)
    source = raw.iloc[layout["data_start"]:].copy()
    source = source.map(
        lambda value: pd.NA if isinstance(value, str) and not value.strip() else value
    )
    data = pd.DataFrame({
        "Number": source.iloc[:, layout["number"]],
        "Date": source.iloc[:, layout["date"]],
        "Student": source.iloc[:, layout["student"]],
        "Supervisor": source.iloc[:, layout["supervisor"]],
        "Panel Role": source.iloc[:, layout["panel_role"]],
        "External": source.iloc[:, layout["external"]],
        **{code: source.iloc[:, layout[code]] for code in layout["internal_codes"]},
    })
    data = data.reset_index(drop=True)
    panel_mask = data["Panel Role"].map(clean_label) == "PANEL"
    # A session is anchored by its consecutive PANEL rows. This is more robust
    # than No. because merged cells occasionally place No. inside the panel run.
    panel_run = (panel_mask & ~panel_mask.shift(fill_value=False)).cumsum()
    records: list[dict] = []
    for run_id in panel_run[panel_mask].drop_duplicates():
        positions = data.index[panel_mask & (panel_run == run_id)]
        first_position, last_position = int(positions.min()), int(positions.max())
        block = data.iloc[max(0, first_position - 1): min(len(data), last_position + 2)]

        def first_value(column: str):
            values = block[column].dropna()
            return values.iloc[0] if not values.empty else pd.NA

        date_value = first_value("Date")
        parsed_date = pd.to_datetime(date_value, errors="coerce", dayfirst=True)
        base = {
            "Proposal Year": proposal_year,
            "Date": parsed_date,
            "Student": first_value("Student"),
            "Supervisor": first_value("Supervisor"),
        }
        panel_rows = data.loc[positions]
        for _, row in panel_rows.iterrows():
            for code in layout["internal_codes"]:
                if is_appointment(row.get(code)):
                    records.append({**base, "Panel Member": code, "Panel Category": CATEGORY_INTERNAL, "Panel Code": code})
            external = row.get("External")
            # LUAR sometimes contains a numeric frequency marker. Only text is
            # a panel name and therefore an external appointment.
            if isinstance(external, str) and external.strip():
                records.append({**base, "Panel Member": str(external).strip(), "Panel Category": CATEGORY_EXTERNAL, "Panel Code": pd.NA})

    columns = ["Proposal Year", "Date", "Student", "Supervisor", "Panel Member", "Panel Category", "Panel Code"]
    long_df = pd.DataFrame(records, columns=columns)
    if long_df.empty:
        return long_df
    long_df = long_df.dropna(subset=["Date", "Student", "Supervisor", "Panel Member"])
    return long_df.drop_duplicates().sort_values(["Proposal Year", "Date", "Student", "Panel Category"]).reset_index(drop=True)


def quality_checks(data: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Identify unusual panel counts and supervisor/panel conflicts."""
    if data.empty:
        return pd.DataFrame(), pd.DataFrame()
    counts = data.groupby(["Proposal Year", "Student", "Date"], as_index=False).size().rename(columns={"size": "Panel Appointments"})
    unusual = counts[counts["Panel Appointments"] != 3]
    conflicts = data[
        (data["Panel Category"] == CATEGORY_INTERNAL)
        & (data["Supervisor"].str.strip().str.upper() == data["Panel Member"].str.strip().str.upper())
    ][["Date", "Student", "Supervisor", "Panel Member"]]
    return unusual, conflicts


def multiselect_filter(label: str, options: Iterable[str], key: str) -> list[str]:
    clean = sorted(pd.Series(list(options), dtype="object").dropna().astype(str).unique())
    return st.multiselect(label, clean, placeholder="All", key=key)


def apply_filters(
    data: pd.DataFrame,
    date_range: tuple,
    years: list[int],
    supervisors: list[str],
    panel_members: list[str],
    categories: list[str],
    students: list[str],
) -> pd.DataFrame:
    """Apply all global filters consistently."""
    filtered = data.copy()
    if len(date_range) == 2:
        start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
        filtered = filtered[filtered["Date"].between(start, end)]
    if years:
        filtered = filtered[filtered["Proposal Year"].isin(years)]
    for column, selected in [
        ("Supervisor", supervisors), ("Panel Member", panel_members),
        ("Panel Category", categories), ("Student", students),
    ]:
        if selected:
            filtered = filtered[filtered[column].isin(selected)]
    return filtered


def most_common(data: pd.DataFrame, category: str | None = None) -> str:
    subset = data if category is None else data[data["Panel Category"] == category]
    if subset.empty:
        return "—"
    counts = subset["Panel Member"].value_counts()
    return f"{counts.index[0]} ({counts.iloc[0]})"


def frequency_table(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame(columns=["Panel Member", "Panel Category", "Appointments", "Percentage Share"])
    table = data.groupby(["Panel Member", "Panel Category"], as_index=False).size().rename(columns={"size": "Appointments"})
    table["Percentage Share"] = (100 * table["Appointments"] / table["Appointments"].sum()).round(1).astype(str) + "%"
    return table.sort_values("Appointments", ascending=False)


def supervisor_frequency_table(data: pd.DataFrame, supervisor: str) -> pd.DataFrame:
    """Show observed panels plus eligible internal members, including zeros."""
    table = frequency_table(data).copy()
    supervisor_code = clean_label(supervisor)
    table = table[
        ~(
            (table["Panel Category"] == CATEGORY_INTERNAL)
            & (table["Panel Member"].map(clean_label) == supervisor_code)
        )
    ].copy()
    internal_universe = [
        code for code in [*INTERNAL_CODES, *OPTIONAL_INTERNAL_CODES]
        if clean_label(code) != supervisor_code
    ]
    observed_internal = set(
        table.loc[table["Panel Category"] == CATEGORY_INTERNAL, "Panel Member"]
    )
    missing_rows = pd.DataFrame(
        [
            {
                "Panel Member": code,
                "Panel Category": CATEGORY_INTERNAL,
                "Appointments": 0,
                "Percentage Share": "0.0%",
            }
            for code in internal_universe
            if code not in observed_internal
        ]
    )
    if not missing_rows.empty:
        table = pd.concat([table, missing_rows], ignore_index=True)
    total_appointments = table["Appointments"].sum()
    table["Percentage Share"] = table["Appointments"].map(
        lambda count: f"{(100 * count / total_appointments):.1f}%"
        if total_appointments else "0.0%"
    )
    table["Appointment Status"] = table["Appointments"].map(
        lambda count: "Not appointed" if count == 0 else "Appointed"
    )
    return table.sort_values(
        ["Appointments", "Panel Category", "Panel Member"],
        ascending=[False, False, True],
    ).reset_index(drop=True)


def supervisor_share_chart(table: pd.DataFrame):
    """Create an appointment-share chart; zero rows remain table-only."""
    appointed = table[table["Appointments"] > 0]
    if appointed.empty:
        return None
    return px.pie(
        appointed,
        names="Panel Member",
        values="Appointments",
        color="Panel Category",
        hole=0.42,
        color_discrete_map={CATEGORY_INTERNAL: THEME_PURPLE, CATEGORY_EXTERNAL: THEME_ORANGE},
        hover_data=["Panel Category", "Appointments"],
    ).update_traces(
        textposition="inside",
        textinfo="percent+label",
    ).update_layout(
        legend_title_text="Panel category",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=520,
        margin=dict(l=20, r=20, t=20, b=20),
    )


def generate_eligible_lecturer_pool(
    data: pd.DataFrame,
    supervisor: str,
    manually_excluded: list[str] | None = None,
    existing_selected: list[str] | None = None,
) -> list[str]:
    """Return configured internal lecturers after applying all exclusions."""
    observed = set(
        data.loc[data["Panel Category"] == CATEGORY_INTERNAL, "Panel Member"]
        .dropna().astype(str)
    )
    # The configured list keeps zero-appointment lecturers eligible; observed
    # codes confirm which configured values are represented in the source.
    pool = list(dict.fromkeys([*RECOMMENDATION_LECTURER_CODES, *sorted(observed)]))
    pool = [code for code in pool if code in RECOMMENDATION_LECTURER_CODES]
    excluded = {
        clean_label(supervisor),
        *[clean_label(code) for code in (manually_excluded or [])],
        *[clean_label(code) for code in (existing_selected or [])],
    }
    return sorted(code for code in pool if clean_label(code) not in excluded)


def calculate_lecturer_statistics(
    data: pd.DataFrame,
    lecturers: list[str],
    supervisor: str,
    reference_date: pd.Timestamp,
) -> tuple[pd.DataFrame, bool]:
    """Calculate auditable historical workload measures for each lecturer."""
    internal = data[data["Panel Category"] == CATEGORY_INTERNAL].copy()
    internal["Date"] = pd.to_datetime(internal["Date"], errors="coerce")
    valid_dates = internal["Date"].notna().any()
    recent_start = reference_date - pd.Timedelta(days=RECENT_APPOINTMENT_DAYS)
    records = []
    for code in lecturers:
        lecturer_rows = internal[internal["Panel Member"] == code]
        supervisor_rows = lecturer_rows[lecturer_rows["Supervisor"] == supervisor]
        if valid_dates:
            recent_rows = lecturer_rows[
                lecturer_rows["Date"].between(recent_start, reference_date, inclusive="both")
            ]
            recent_count = len(recent_rows)
        else:
            recent_count = pd.NA
        records.append({
            "Lecturer Code": code,
            "Full Lecturer Name": LECTURER_FULL_NAMES.get(code) or "—",
            "Total Proposal Appointments": len(lecturer_rows),
            "Appointments for Selected Supervisor": len(supervisor_rows),
            f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days": recent_count,
            "Unique Students Assessed": lecturer_rows[["Proposal Year", "Student"]].drop_duplicates().shape[0],
            "Unique Supervisors Served": lecturer_rows["Supervisor"].nunique(),
        })
    return pd.DataFrame(records), valid_dates


def normalise_scores(values: pd.Series, higher_is_better: bool = False) -> pd.Series:
    """Normalise a metric to 0–100, with 100 representing better distribution."""
    numeric = pd.to_numeric(values, errors="coerce").fillna(0).astype(float)
    minimum, maximum = numeric.min(), numeric.max()
    if maximum == minimum:
        return pd.Series(100.0, index=values.index)
    score = (numeric - minimum) / (maximum - minimum) * 100
    return score if higher_is_better else 100 - score


def recommendation_status(score: float) -> str:
    for label, threshold in RECOMMENDATION_STATUS_THRESHOLDS.items():
        if score >= threshold:
            return label
    return "Lower priority"


def generate_recommendation_explanation(row: pd.Series) -> str:
    reasons = []
    if row["Overall Workload Score"] >= 67:
        reasons.append("has a relatively low overall appointment frequency")
    elif row["Overall Workload Score"] <= 33:
        reasons.append("has a comparatively high historical appointment frequency")
    else:
        reasons.append("has a moderate overall appointment frequency")
    if row["Appointments for Selected Supervisor"] == 0:
        reasons.append("has not previously served for this supervisor")
    elif row["Supervisor-Specific Use Score"] >= 67:
        reasons.append("has been used infrequently by this supervisor")
    else:
        reasons.append("has served this supervisor more frequently than some peers")
    recent_value = row.get(f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days")
    if pd.notna(recent_value):
        reasons.append("has a low recent workload" if row["Recent Workload Score"] >= 67 else "has recent panel activity")
    return "Recommended based on distribution because this lecturer " + ", ".join(reasons) + "."


def calculate_recommendation_scores(
    statistics: pd.DataFrame,
    recent_dates_available: bool,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calculate the transparent weighted recommendation score from 0 to 100."""
    if statistics.empty:
        return statistics.copy(), pd.DataFrame()
    scored = statistics.copy()
    scored["Overall Workload Score"] = normalise_scores(scored["Total Proposal Appointments"])
    scored["Supervisor-Specific Use Score"] = normalise_scores(scored["Appointments for Selected Supervisor"])
    scored["Supervisor Breadth Score"] = normalise_scores(scored["Unique Supervisors Served"], higher_is_better=True)
    recent_column = f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days"
    if recent_dates_available:
        scored["Recent Workload Score"] = normalise_scores(scored[recent_column])
        active_weights = dict(RECOMMENDATION_WEIGHTS)
    else:
        scored["Recent Workload Score"] = pd.NA
        active_weights = {key: value for key, value in RECOMMENDATION_WEIGHTS.items() if key != "Recent Workload"}
    weight_total = sum(active_weights.values())
    active_weights = {key: value / weight_total for key, value in active_weights.items()}
    component_columns = {
        "Overall Workload": "Overall Workload Score",
        "Supervisor-Specific Use": "Supervisor-Specific Use Score",
        "Recent Workload": "Recent Workload Score",
        "Supervisor Breadth": "Supervisor Breadth Score",
    }
    scored["Recommendation Score"] = sum(
        scored[component_columns[name]].astype(float) * weight
        for name, weight in active_weights.items()
    ).round(1)
    scored["Recommendation Status"] = scored["Recommendation Score"].map(recommendation_status)
    scored["Explanation"] = scored.apply(generate_recommendation_explanation, axis=1)
    scored = scored.sort_values(
        ["Recommendation Score", "Total Proposal Appointments", "Lecturer Code"],
        ascending=[False, True, True],
    ).reset_index(drop=True)
    scored.insert(0, "Rank", range(1, len(scored) + 1))
    component_table = pd.DataFrame([
        {
            "Score Component": name,
            "Configured Weight": f"{RECOMMENDATION_WEIGHTS[name] * 100:.0f}%",
            "Applied Weight": f"{active_weights.get(name, 0) * 100:.1f}%",
            "Direction": "Lower frequency receives a higher score" if name != "Supervisor Breadth" else "More supervisors served receives a higher score",
        }
        for name in RECOMMENDATION_WEIGHTS
    ])
    return scored, component_table


def validate_recommendation_inputs(
    data: pd.DataFrame,
    supervisor: str,
    eligible_pool: list[str],
    required_count: int,
) -> list[str]:
    warnings = []
    supervisor_sessions = data.loc[data["Supervisor"] == supervisor, ["Proposal Year", "Student"]].drop_duplicates().shape[0]
    if supervisor_sessions < 3:
        warnings.append("The selected supervisor has fewer than three historical students in the current analysis period.")
    if "Date" not in data or pd.to_datetime(data["Date"], errors="coerce").notna().sum() == 0:
        warnings.append("No valid appointment dates are available; the recent-workload component will be omitted.")
    elif pd.to_datetime(data["Date"], errors="coerce").isna().any():
        warnings.append("Some appointment dates are invalid or missing and are excluded from recent-workload calculations.")
    if not eligible_pool:
        warnings.append("No eligible internal lecturer codes could be identified after exclusions.")
    elif len(eligible_pool) < required_count:
        warnings.append("The eligible lecturer pool is smaller than the number of panel members required.")
    observed_codes = set(
        data.loc[data["Panel Category"] == CATEGORY_INTERNAL, "Panel Member"].dropna().astype(str)
    )
    if not observed_codes.intersection(RECOMMENDATION_LECTURER_CODES):
        warnings.append("Configured lecturer codes could not be matched to internal panel records in the dataset.")
    conflicts = data[
        (data["Panel Category"] == CATEGORY_INTERNAL)
        & (data["Supervisor"].map(clean_label) == data["Panel Member"].map(clean_label))
    ]
    if not conflicts.empty:
        warnings.append("The dataset contains a supervisor recorded as a panel member for their own student.")
    duplicates = data.duplicated(["Proposal Year", "Student", "Panel Member"], keep=False)
    if duplicates.any():
        warnings.append("Duplicate panel appointments appear for the same student and should be checked.")
    return warnings


def review_proposed_panel_selection(
    recommendations: pd.DataFrame,
    selected: list[str],
    supervisor: str,
    manually_excluded: list[str],
) -> dict:
    """Summarise workload balance and conflicts without blocking a selection."""
    review = recommendations[recommendations["Lecturer Code"].isin(selected)].copy()
    conflicts = []
    if supervisor in selected:
        conflicts.append(f"{supervisor} is the selected supervisor and should not be appointed for their own student.")
    overlap = sorted(set(selected) & set(manually_excluded))
    if overlap:
        conflicts.append("Manually excluded lecturer(s) selected: " + ", ".join(overlap))
    workloads = review["Total Proposal Appointments"].astype(float)
    if len(workloads) < 2 or workloads.mean() == 0:
        concentration = 0.0
    else:
        concentration = float(workloads.std(ddof=0) / workloads.mean())
    if concentration <= 0.25:
        indicator = "Balanced distribution"
    elif concentration <= 0.60:
        indicator = "Moderately concentrated distribution"
    else:
        indicator = "Highly concentrated distribution"
    return {"table": review, "conflicts": conflicts, "indicator": indicator, "concentration": concentration}


def frequency_chart(data: pd.DataFrame):
    table = frequency_table(data)
    if table.empty:
        return None
    table = table.sort_values("Appointments", ascending=True)
    return px.bar(
        table, x="Appointments", y="Panel Member", color="Panel Category", orientation="h",
        text="Appointments", color_discrete_map={CATEGORY_INTERNAL: THEME_PURPLE, CATEGORY_EXTERNAL: THEME_ORANGE},
        labels={"Panel Member": "Panel member", "Appointments": "Number of appointments"},
    ).update_layout(
        legend_title_text="Panel category",
        height=max(430, 30 * len(table)),
        margin=dict(l=20, r=20, t=20, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(gridcolor="#E8DDF0"),
    )


def matrix_chart(data: pd.DataFrame):
    if data.empty:
        return None
    matrix = pd.crosstab(data["Supervisor"], data["Panel Member"])
    return px.imshow(
        matrix, text_auto=True, aspect="auto",
        color_continuous_scale=["#FBF7FD", "#C9A7DD", THEME_PURPLE],
        labels={"x": "Panel member", "y": "Supervisor", "color": "Frequency"},
    ).update_layout(
        height=max(450, 38 * len(matrix)),
        margin=dict(l=20, r=20, t=20, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )


def show_empty() -> None:
    st.info("No records match the current filters. Adjust the filters to continue.")


def main() -> None:
    st.markdown("""
        <style>
        :root {
            --ptpm-purple: #5B2C83;
            --ptpm-purple-dark: #3D1A5B;
            --ptpm-purple-light: #F3ECF8;
            --ptpm-orange: #E67E22;
            --ptpm-orange-light: #FFF1E3;
        }
        .stApp {background: linear-gradient(180deg, #FCFAFD 0%, #FFFFFF 18rem);}
        .block-container {padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1500px;}
        [data-testid="stSidebar"] {background: linear-gradient(180deg, #F3ECF8 0%, #FFF9F3 100%); border-right: 1px solid #DCCBE7;}
        [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {color:var(--ptpm-purple-dark);}
        [data-testid="stMetric"] {
            background:#FFFFFF;
            border:1px solid #DCCBE7;
            border-top:4px solid var(--ptpm-purple);
            padding:16px;
            border-radius:10px;
            box-shadow:0 3px 12px rgba(61,26,91,.08);
        }
        [data-testid="stMetric"] [data-testid="stMetricValue"] {color:var(--ptpm-purple-dark);}
        h1 {color:var(--ptpm-purple-dark); border-bottom:3px solid var(--ptpm-orange); padding-bottom:.45rem;}
        h2, h3 {color:var(--ptpm-purple);}
        div[data-baseweb="tab-list"] {gap:.35rem; border-bottom:1px solid #DCCBE7;}
        button[data-baseweb="tab"] {background:#F7F1FA; border-radius:8px 8px 0 0; padding:.7rem 1rem;}
        button[data-baseweb="tab"][aria-selected="true"] {background:var(--ptpm-purple); color:#FFFFFF;}
        button[data-baseweb="tab"][aria-selected="true"] p {color:#FFFFFF;}
        .stButton > button, .stDownloadButton > button {
            background:var(--ptpm-purple);
            color:#FFFFFF;
            border:1px solid var(--ptpm-purple);
            border-radius:7px;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            background:var(--ptpm-orange);
            color:#FFFFFF;
            border-color:var(--ptpm-orange);
        }
        [data-testid="stDataFrame"] {border:1px solid #DCCBE7; border-radius:8px; overflow:hidden;}
        [data-testid="stAlert"] {border-left:5px solid var(--ptpm-orange); border-radius:7px;}
        hr {border-color:#E5D7ED;}
        </style>
    """, unsafe_allow_html=True)

    st.title(APP_TITLE)
    st.caption("Transparent monitoring of proposal panel appointment frequency and distribution")
    if "reload_token" not in st.session_state:
        st.session_state.reload_token = 0
    with st.sidebar:
        st.header("Data source")
        if st.button("Reload Data", width="stretch", help="Read the latest saved Excel workbook"):
            st.session_state.reload_token += 1
    frames = []
    reload_times = []
    try:
        for year, source_config in DATA_SOURCES.items():
            raw, loaded_at = load_source_data(
                str(source_config["path"]),
                str(source_config["worksheet"]),
                st.session_state.reload_token,
            )
            errors = validate_source(raw)
            if errors:
                raise ValueError(f"{year}: " + " ".join(errors))
            frames.append(transform_to_long(raw, year))
            reload_times.append(loaded_at)
    except Exception as exc:
        st.error(f"Unable to load the dashboard data. {exc}")
        st.stop()

    data = pd.concat(frames, ignore_index=True)
    refreshed_at = max(reload_times)
    if data.empty:
        st.error("No valid panel appointments were found in the source file.")
        st.stop()

    with st.sidebar:
        st.header("Dashboard filters")
        bounds = (data["Date"].min().date(), data["Date"].max().date())
        dates = st.date_input("Proposal date range", value=bounds, min_value=bounds[0], max_value=bounds[1])
        years = st.multiselect("Proposal year", sorted(data["Proposal Year"].unique()), placeholder="All years", key="global_year")
        supervisors = multiselect_filter("Supervisor", data["Supervisor"], "global_supervisor")
        panels = multiselect_filter("Panel member", data["Panel Member"], "global_panel")
        categories = multiselect_filter("Panel category", data["Panel Category"], "global_category")
        students = multiselect_filter("Student", data["Student"], "global_student")
        st.divider()
        for year, source_config in DATA_SOURCES.items():
            st.caption(f"{year}: {source_config['path'].name} · {source_config['worksheet']}")
        st.caption(f"Latest successful reload: {refreshed_at.strftime('%d %B %Y, %I:%M:%S %p')}")
        if st.button("Reset filters", width="stretch"):
            for key in ["global_year", "global_supervisor", "global_panel", "global_category", "global_student"]:
                st.session_state[key] = []
            st.rerun()

    filtered = apply_filters(data, dates, years, supervisors, panels, categories, students)
    unusual, conflicts = quality_checks(filtered)
    if not unusual.empty:
        with st.expander(f"Data quality warning: {len(unusual)} proposal session(s) do not have exactly three appointments"):
            st.dataframe(unusual, hide_index=True, width="stretch")
    if not conflicts.empty:
        with st.expander(f"Conflict warning: {len(conflicts)} supervisor-panel match(es) detected"):
            st.warning("A supervisor is recorded as a panel member for their own student. Please verify the source record.")
            st.dataframe(conflicts, hide_index=True, width="stretch")

    tabs = st.tabs(["Overview", "Panel Frequency Analysis", "Supervisor Analysis", "Supervisor–Panel Matrix", "Student Lookup", "Detailed Records", "Panel Recommendation"])

    with tabs[0]:
        if filtered.empty:
            show_empty()
        else:
            metrics = [
                ("Total proposal sessions", filtered[["Proposal Year", "Student", "Date"]].drop_duplicates().shape[0]),
                ("Total panel appointments", len(filtered)),
                ("Unique internal panel members", filtered.loc[filtered["Panel Category"] == CATEGORY_INTERNAL, "Panel Member"].nunique()),
                ("Unique external panel members", filtered.loc[filtered["Panel Category"] == CATEGORY_EXTERNAL, "Panel Member"].nunique()),
                ("Most frequently appointed panel member", most_common(filtered)),
                ("Most frequently appointed external panel member", most_common(filtered, CATEGORY_EXTERNAL)),
            ]
            for row in [metrics[:3], metrics[3:]]:
                cols = st.columns(3)
                for col, (label, value) in zip(cols, row):
                    col.metric(label, value)
            st.subheader("Appointment distribution")
            chart = frequency_chart(filtered)
            st.plotly_chart(chart, width="stretch", key="overview_frequency_chart")

    with tabs[1]:
        view = st.radio("Display panel categories", ["Both", CATEGORY_INTERNAL, CATEGORY_EXTERNAL], horizontal=True)
        analysis = filtered if view == "Both" else filtered[filtered["Panel Category"] == view]
        if analysis.empty:
            show_empty()
        else:
            st.plotly_chart(frequency_chart(analysis), width="stretch", key="panel_analysis_frequency_chart")
            selected_panel = st.selectbox("Select a panel member for appointment details", sorted(analysis["Panel Member"].unique()))
            detail = analysis[analysis["Panel Member"] == selected_panel]
            st.metric("Total appointments", len(detail))
            st.dataframe(detail[["Proposal Year", "Date", "Student", "Supervisor", "Panel Category"]], hide_index=True, width="stretch",
                         column_config={"Date": st.column_config.DateColumn("Appointment date", format="DD MMM YYYY")})

    with tabs[2]:
        available = sorted(filtered["Supervisor"].unique()) if not filtered.empty else []
        if not available:
            show_empty()
        else:
            supervisor = st.selectbox("Select a supervisor", available)
            subset = filtered[filtered["Supervisor"] == supervisor]
            cols = st.columns(4)
            cols[0].metric("Students supervised", subset[["Proposal Year", "Student"]].drop_duplicates().shape[0])
            cols[1].metric("Total panel appointments", len(subset))
            cols[2].metric("Unique panel members used", subset["Panel Member"].nunique())
            cols[3].metric("Most frequently appointed", most_common(subset))
            st.subheader("Panel member frequency and share")
            st.caption("All eligible internal PTPM panel members are listed except the selected supervisor; zero values indicate lecturers not appointed within the current filters.")
            supervisor_table = supervisor_frequency_table(subset, supervisor)
            st.dataframe(
                supervisor_table,
                hide_index=True,
                width="stretch",
                column_config={
                    "Appointments": st.column_config.NumberColumn("Appointments", format="%d"),
                },
            )
            st.subheader("Panel appointment share")
            st.caption("The chart shows appointed members only; zero-appointment internal members remain listed in the table above.")
            share_chart = supervisor_share_chart(supervisor_table)
            if share_chart is None:
                st.info("No appointments are available for the share chart under the current filters.")
            else:
                st.plotly_chart(share_chart, width="stretch", key="supervisor_panel_share_chart")
            st.subheader("Students and their panels")
            student_panels = subset.groupby(["Proposal Year", "Date", "Student"], as_index=False).agg(**{"Panel Members": ("Panel Member", lambda x: ", ".join(x))})
            st.dataframe(student_panels, hide_index=True, width="stretch",
                         column_config={"Date": st.column_config.DateColumn("Proposal date", format="DD MMM YYYY")})

    with tabs[3]:
        if filtered.empty:
            show_empty()
        else:
            st.plotly_chart(matrix_chart(filtered), width="stretch", key="supervisor_panel_matrix")

    with tabs[4]:
        available = (
            filtered[["Proposal Year", "Student"]]
            .drop_duplicates()
            .sort_values(["Proposal Year", "Student"])
            .apply(lambda row: f"{row['Proposal Year']} — {row['Student']}", axis=1)
            .tolist()
            if not filtered.empty else []
        )
        if not available:
            show_empty()
        else:
            student_selection = st.selectbox("Search or select a student", available)
            year_text, student = student_selection.split(" — ", 1)
            record = filtered[(filtered["Proposal Year"] == int(year_text)) & (filtered["Student"] == student)]
            first = record.iloc[0]
            cols = st.columns(4)
            cols[0].metric("Student", student)
            cols[1].metric("Proposal year", int(first["Proposal Year"]))
            cols[2].metric("Proposal date", first["Date"].strftime("%d %B %Y"))
            cols[3].metric("Supervisor", first["Supervisor"])
            st.dataframe(record[["Panel Member", "Panel Category", "Panel Code"]], hide_index=True, width="stretch")

    with tabs[5]:
        if filtered.empty:
            show_empty()
        else:
            query = st.text_input("Search detailed records", placeholder="Enter a student, supervisor or panel member")
            records = filtered.copy()
            if query:
                searchable = records[["Proposal Year", "Student", "Supervisor", "Panel Member", "Panel Category"]].astype(str).fillna("").agg(" ".join, axis=1)
                records = records[searchable.str.contains(query, case=False, regex=False)]
            display = records[["Proposal Year", "Date", "Student", "Supervisor", "Panel Member", "Panel Category"]]
            st.dataframe(display, hide_index=True, width="stretch",
                         column_config={"Date": st.column_config.DateColumn("Date", format="DD MMM YYYY")})
            st.download_button("Download filtered data as CSV", display.to_csv(index=False).encode("utf-8"),
                               "ptpm_filtered_panel_appointments.csv", "text/csv", width="stretch")

    with tabs[6]:
        st.subheader("Panel Recommendation")
        st.info(
            "Recommendations are based on historical appointment frequency and workload distribution. "
            "Panel expertise, suitability, availability, conflicts of interest, academic requirements "
            "and management judgement must be considered before any appointment is confirmed."
        )
        st.caption("This decision-support feature does not appoint panel members or determine academic suitability.")
        if filtered.empty:
            show_empty()
        else:
            input_left, input_right = st.columns(2)
            with input_left:
                recommendation_student = st.selectbox(
                    "Student name (optional)",
                    ["Not specified", *sorted(filtered["Student"].dropna().unique())],
                    key="recommendation_student",
                )
                recommendation_supervisor = st.selectbox(
                    "Supervisor name (required)",
                    sorted(filtered["Supervisor"].dropna().unique()),
                    key="recommendation_supervisor",
                )
                panel_members_required = st.number_input(
                    "Number of panel members required", min_value=1, max_value=5, value=3, step=1,
                    key="recommendation_count",
                )
            with input_right:
                valid_recommendation_dates = pd.to_datetime(filtered["Date"], errors="coerce").dropna()
                default_reference_date = (
                    valid_recommendation_dates.max().date()
                    if not valid_recommendation_dates.empty else pd.Timestamp.now().date()
                )
                reference_date = st.date_input(
                    "Appointment reference date", value=default_reference_date,
                    key="recommendation_reference_date",
                )
                manually_excluded = st.multiselect(
                    "Lecturers to exclude manually",
                    RECOMMENDATION_LECTURER_CODES,
                    key="recommendation_manual_exclusions",
                )
                existing_selected = st.multiselect(
                    "Existing selected panel members, if any",
                    RECOMMENDATION_LECTURER_CODES,
                    key="recommendation_existing_panel",
                )

            eligible_pool = generate_eligible_lecturer_pool(
                filtered,
                recommendation_supervisor,
                manually_excluded,
                existing_selected,
            )
            input_warnings = validate_recommendation_inputs(
                filtered,
                recommendation_supervisor,
                eligible_pool,
                int(panel_members_required),
            )
            for warning in input_warnings:
                st.warning(warning)

            statistics, recent_dates_available = calculate_lecturer_statistics(
                filtered,
                eligible_pool,
                recommendation_supervisor,
                pd.Timestamp(reference_date),
            )
            recommendations, component_table = calculate_recommendation_scores(
                statistics,
                recent_dates_available,
            )

            if recommendations.empty:
                st.info("No eligible lecturers remain after applying the selected exclusions.")
            else:
                st.subheader("Ranked recommendations")
                result_columns = [
                    "Rank", "Lecturer Code", "Full Lecturer Name", "Recommendation Score",
                    "Total Proposal Appointments", "Appointments for Selected Supervisor",
                    f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days",
                    "Unique Students Assessed", "Unique Supervisors Served",
                    "Recommendation Status", "Explanation",
                ]
                st.dataframe(
                    recommendations[result_columns],
                    hide_index=True,
                    width="stretch",
                    column_config={
                        "Recommendation Score": st.column_config.ProgressColumn(
                            "Recommendation Score", min_value=0, max_value=100, format="%.1f"
                        )
                    },
                )

                score_chart_data = recommendations.sort_values("Recommendation Score", ascending=True)
                score_chart = px.bar(
                    score_chart_data,
                    x="Recommendation Score",
                    y="Lecturer Code",
                    orientation="h",
                    text="Recommendation Score",
                    color="Recommendation Score",
                    color_continuous_scale=[THEME_ORANGE_LIGHT, THEME_ORANGE, THEME_PURPLE],
                    range_color=[0, 100],
                    labels={"Lecturer Code": "Lecturer", "Recommendation Score": "Score (0–100)"},
                ).update_layout(
                    coloraxis_showscale=False,
                    height=max(430, 34 * len(recommendations)),
                    margin=dict(l=20, r=20, t=20, b=20),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(range=[0, 100], gridcolor="#E8DDF0"),
                )
                st.plotly_chart(score_chart, width="stretch", key="recommendation_score_chart")

                workload_columns = [
                    "Total Proposal Appointments",
                    "Appointments for Selected Supervisor",
                    f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days",
                ]
                workload_chart_data = recommendations.melt(
                    id_vars="Lecturer Code",
                    value_vars=workload_columns,
                    var_name="Workload Measure",
                    value_name="Appointments",
                )
                workload_chart = px.bar(
                    workload_chart_data,
                    x="Lecturer Code",
                    y="Appointments",
                    color="Workload Measure",
                    barmode="group",
                    color_discrete_sequence=[THEME_PURPLE, THEME_ORANGE, "#A67BC3"],
                    labels={"Lecturer Code": "Lecturer"},
                ).update_layout(
                    height=460,
                    margin=dict(l=20, r=20, t=20, b=20),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    yaxis=dict(gridcolor="#E8DDF0"),
                    legend_title_text="Workload measure",
                )
                st.subheader("Appointment workload comparison")
                st.plotly_chart(workload_chart, width="stretch", key="recommendation_workload_chart")

                st.subheader("Score component calculation")
                st.dataframe(component_table, hide_index=True, width="stretch")
                with st.expander("How recommendations are calculated"):
                    st.markdown(
                        f"""
                        Each eligible lecturer receives four normalized component scores from 0 to 100.
                        Lower overall workload, lower use by the selected supervisor, and fewer appointments
                        in the previous {RECENT_APPOINTMENT_DAYS} days receive higher component scores. Serving
                        students across a broader range of supervisors receives a higher breadth score.

                        The configured weights are **45% overall workload**, **30% supervisor-specific use**,
                        **15% recent workload**, and **10% supervisor breadth**. If appointment dates are not
                        available, the recent component is omitted and its weight is redistributed
                        proportionally across the other components. Scores support workload distribution only;
                        they do not measure expertise, suitability, independence, or availability.
                        """
                    )

                st.subheader("Proposed panel selection")
                proposed_selection = st.multiselect(
                    "Select proposed panel members from the ranked list",
                    recommendations["Lecturer Code"].tolist(),
                    max_selections=int(panel_members_required),
                    key="recommendation_proposed_panel",
                )
                combined_selection = list(dict.fromkeys([*existing_selected, *proposed_selection]))
                if combined_selection:
                    review_statistics, _ = calculate_lecturer_statistics(
                        filtered,
                        combined_selection,
                        recommendation_supervisor,
                        pd.Timestamp(reference_date),
                    )
                    selection_review = review_proposed_panel_selection(
                        review_statistics,
                        combined_selection,
                        recommendation_supervisor,
                        manually_excluded,
                    )
                    st.subheader("Panel Selection Review")
                    review_col1, review_col2 = st.columns(2)
                    review_col1.metric("Selected lecturers", len(combined_selection))
                    review_col2.metric("Overall balance indicator", selection_review["indicator"])
                    review_display_columns = [
                        "Lecturer Code", "Total Proposal Appointments",
                        "Appointments for Selected Supervisor",
                        f"Appointments in Previous {RECENT_APPOINTMENT_DAYS} Days",
                    ]
                    st.dataframe(
                        selection_review["table"][review_display_columns],
                        hide_index=True,
                        width="stretch",
                    )
                    if selection_review["conflicts"]:
                        for conflict in selection_review["conflicts"]:
                            st.warning(conflict)
                    else:
                        st.success("No supervisor or manual-exclusion conflicts were detected in the proposed selection.")
                    eligible_median = statistics["Total Proposal Appointments"].median()
                    frequent = selection_review["table"][
                        selection_review["table"]["Total Proposal Appointments"] > eligible_median
                    ]["Lecturer Code"].tolist()
                    if frequent:
                        st.info(
                            "Historical pattern: " + ", ".join(frequent)
                            + " have appointment totals above the eligible-pool median. This does not prevent selection; consider expertise, availability and the overall panel balance."
                        )
                    st.caption(
                        f"Workload concentration coefficient: {selection_review['concentration']:.2f}. "
                        "This indicator describes distribution only and is not a suitability assessment."
                    )
                elif recommendation_student != "Not specified":
                    st.caption(f"No proposed lecturers have yet been selected for {recommendation_student}.")

    st.divider()
    st.info("Panel appointment frequency supports transparent and balanced decision-making. It does not determine a lecturer’s suitability, subject expertise, availability or independence for a particular proposal.")
    st.caption("Data sources: 2025 and 2026 Excel workbooks · Proposal sessions only")


if __name__ == "__main__":
    main()
