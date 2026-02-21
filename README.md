# HRMS Attendance Dashboard

Simple browser-based HRMS dashboard for attendance.

## Features
- Upload Excel (`.xlsx/.xls`) with employee attendance data.
- Expected columns: `Employee ID`, `Name`, `Punch In`, `Punch Out`, `Total Time`.
- Configurable attendance threshold in hours (default `8`).
- Automatic Present/Absent classification.
- Summary cards and doughnut chart visualization.

## Run locally
```bash
python3 -m http.server 8000
```
Then open `http://localhost:8000`.
