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

    tabs = st.tabs(["Overview", "Panel Frequency Analysis", "Supervisor Analysis", "Supervisor–Panel Matrix", "Student Lookup", "Detailed Records"])

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

    st.divider()
    st.info("Panel appointment frequency supports transparent and balanced decision-making. It does not determine a lecturer’s suitability, subject expertise, availability or independence for a particular proposal.")
    st.caption("Data sources: 2025 and 2026 Excel workbooks · Proposal sessions only")


if __name__ == "__main__":
    main()
