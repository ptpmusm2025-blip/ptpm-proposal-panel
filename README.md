# PTPM Proposal Panel Appointment Dashboard

A Streamlit dashboard for monitoring the frequency and distribution of lecturers appointed to PTPM postgraduate proposal panels. It reads separate 2025 and 2026 project workbooks without changing them.

## Features

- Overview indicators for sessions, appointments, and internal/external panel participation
- Panel frequency ranking with appointment-level drill-down
- Supervisor analysis with panel share and student-panel lists
- Interactive supervisor–panel heatmap
- Searchable student lookup and detailed records
- Consistent global filters and CSV download
- Proposal-year filtering for 2025, 2026, or a combined view
- Rule-based panel recommendations with transparent 0–100 workload-distribution scores
- Proposed-panel review with neutral workload-balance indicators and conflict checks
- Validation for malformed source data, unusual panel counts, and supervisor-panel conflicts

## Run locally

1. Create and activate a Python virtual environment.
2. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

3. Start the dashboard:

   ```powershell
   streamlit run app.py
   ```

Streamlit will display the local address, normally `http://localhost:8501`.

## Excel data structure

The public dashboard reads sanitized, anonymized copies:

- `public_data_2025.xlsx` → `Nama Pelajar & Panel 2025.`
- `public_data_2026.xlsx` → `Nama Pelajar & Panel 2026`

Each public workbook contains only its relevant proposal worksheet. Personal document metadata, comments, hyperlinks, hidden draft/history worksheets, and unrelated year sheets are excluded. Original source workbooks remain local and ignored by Git.

The 2025 header spans Excel rows 3–4 and the 2026 header spans rows 2–3. Student records use repeated blocks with merged cells and blank separator rows. The loader anchors each session on its consecutive `PANEL` rows, reads the surrounding date, student and supervisor metadata, and converts the block into appointment-level records with a `Proposal Year` field.

An internal lecturer code is treated as appointed when its cell contains `1`; a nonblank `LUAR` value is treated as an external panel appointment. The original workbook is read-only from the dashboard's perspective.

## Future private Google Sheet connection

The data source is isolated in `load_source_data()` in `app.py`. A future authenticated Google Sheets adapter can return the equivalent worksheet grid and refresh timestamp. Dashboard views do not need to be redesigned.

For a private production source, store credentials in Streamlit secrets or an approved institutional secret manager. Never commit credentials to the repository.

## Panel recommendation method

The `Panel Recommendation` tab ranks eligible internal PTPM lecturers using inspectable historical rules rather than machine learning. The initial score weights are 45% lower overall appointment frequency, 30% lower use by the selected supervisor, 15% lower workload in the previous 90 days, and 10% broader distribution across supervisors. If dates are unavailable, the recent-workload weight is redistributed proportionally.

Recommendations support workload-distribution decisions only. Expertise, suitability, availability, conflicts of interest, academic requirements, independence, and management judgement must be assessed before an appointment is confirmed. Nothing is written back to the Excel workbooks.

## Data expectations

Required source fields are `No.`, `Tarikh`, `Nama Pelajar`, `SV` or `Nama SV`, the internal lecturer codes (`WAJ`, `CKT`, `INU`, `NAM`, `NJ`, `MBM`, `SNR`, `MTA`, `MS`, `NR`, `NMB`, `ABN`, with `BIE` detected when present), and `LUAR`. Internal appointment cells should contain `1`; text names in `LUAR` are external appointments.
