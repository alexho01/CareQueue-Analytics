# Clinic Appointment No-Show Dashboard

> A Python data analysis project that identifies factors linked to patient no-shows using synthetic clinic appointment data, SQL queries, and an Excel dashboard.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Dataset](#dataset)
- [Project Structure](#project-structure)
- [How the Code Works](#how-the-code-works)
  - [Step 1 — generate_data.py](#step-1--generate_datapy)
  - [Step 2 — sql_analysis.py](#step-2--sql_analysispy)
  - [Step 3 — build_excel.py](#step-3--build_excelpy)
- [SQL Queries Explained](#sql-queries-explained)
- [Key Findings](#key-findings)
- [Recommendations](#recommendations)
- [Installation & Setup](#installation--setup)
- [Run Order](#run-order)
- [Technologies Used](#technologies-used)

---

## Project Overview

This project simulates a real-world healthcare analytics scenario. The goal is to answer:

**"What factors are linked to patients missing their clinic appointments?"**

The project generates a synthetic dataset of 5,000 appointments, stores it in a SQLite database, runs SQL analysis across multiple dimensions, and produces a fully formatted Excel dashboard with charts, tables, and recommendations.

---

## Dataset

The dataset contains **5,000 appointment records** spanning January 2023 to December 2024, with the following columns:

| Column | Description |
|---|---|
| `Appointment_ID` | Unique identifier for each appointment (e.g. APT10001) |
| `Patient_ID` | Unique patient identifier (e.g. PAT67455) |
| `Age` | Patient age (1–95), normally distributed around 42 |
| `Gender` | Male / Female / Other |
| `Appointment_Date` | The date of the actual appointment (YYYY-MM-DD) |
| `Scheduled_Date` | The date the patient booked the appointment |
| `Waiting_Days` | Days between booking and appointment date |
| `Appointment_Day` | Day of the week (Monday–Saturday) |
| `Appointment_Time` | Time of appointment (HH:MM, 08:00–17:45) |
| `Clinic_Type` | General Practice / Cardiology / Pediatrics / Orthopedics / Dermatology / Neurology |
| `Appointment_Type` | Consultation / Follow-up / Routine Check-up / Emergency / Procedure / Lab Test |
| `SMS_Reminder` | Whether an SMS reminder was sent (Yes / No) |
| `Previous_No_Shows` | How many times the patient has previously not shown up (0–5) |
| `Insurance_Type` | Private / Public / None / Medicare / Medicaid |
| `Neighbourhood` | Patient's local area (10 possible neighbourhoods) |
| `Show_Status` | **Target variable** — Showed / No-Show / Cancelled |

---

## Project Structure

```
clinic_project/
│
├── generate_data.py              # Step 1: Creates the dataset
├── sql_analysis.py               # Step 2: Runs SQL queries and saves results
├── build_excel.py                # Step 3: Builds the Excel dashboard
│
├── appointments.csv              # Auto-generated: raw dataset (CSV)
├── clinic.db                     # Auto-generated: SQLite database
├── query_results.pkl             # Auto-generated: saved query results (pickle)
└── Clinic_NoShow_Dashboard_v2.xlsx  # Auto-generated: final Excel report
```

> **Note:** The `.csv`, `.db`, `.pkl`, and `.xlsx` files are all generated automatically when you run the scripts. You only need the three `.py` files to start.

---

## How the Code Works

### Step 1 — `generate_data.py`

**Purpose:** Creates a realistic synthetic dataset and saves it to both CSV and SQLite.

#### What it does:

**1. Defines probability weights for no-shows**

Each clinic type, appointment type, and day of the week is assigned a base no-show probability based on realistic healthcare patterns:

```python
CLINIC_P = {
    "Dermatology": 0.33,      # highest risk
    "Orthopedics": 0.27,
    "General Practice": 0.24,
    ...
}
```

**2. Generates 5,000 patient records**

For each record it randomly assigns:
- A patient age (normal distribution, mean 42)
- A clinic, appointment type, insurance type, neighbourhood
- A scheduled date (random date in 2023–2024)
- A waiting period (exponential distribution — most waits are short, some are long)
- Whether an SMS reminder was sent (62% chance of Yes)
- How many previous no-shows the patient has (Poisson distribution)

**3. Calculates no-show probability per patient**

The base probability is adjusted up or down based on their specific combination of factors:

```python
p_ns = (CLINIC_P[clinic] + APTYPE_P[apt_type] + DAY_P[apt_day]) / 3

if wait_days > 14:   p_ns += 0.09   # longer wait → higher risk
if sms:              p_ns -= 0.11   # reminder sent → lower risk
if prev_ns >= 2:     p_ns += 0.12   # history of no-shows → higher risk
if age < 30:         p_ns += 0.05   # younger patients → higher risk
if insurance == "None": p_ns += 0.04
```

**4. Assigns Show_Status**

A random number is drawn. If it falls below the no-show probability, status is `No-Show`. If it falls in the next band, status is `Cancelled`. Otherwise, `Showed`.

**5. Saves output**

- `appointments.csv` — plain CSV for Excel/manual use
- `clinic.db` — SQLite database for SQL queries

---

### Step 2 — `sql_analysis.py`

**Purpose:** Connects to the SQLite database and runs 10 analytical SQL queries.

#### What it does:

**1. Connects to the database**
```python
conn = sqlite3.connect("clinic.db")
```

**2. Runs each query using pandas**
```python
df = pd.read_sql_query(sql, conn)
```
This executes the SQL and returns results as a pandas DataFrame — making it easy to display, save, and pass to the Excel builder.

**3. Saves all results to a pickle file**
```python
with open("query_results.pkl", "wb") as f:
    pickle.dump(results, f)
```
A pickle file serialises Python objects to disk. This lets `build_excel.py` load the query results without re-running the database queries.

#### The 10 queries cover:
- Overall no-show and cancellation rates
- No-show rate by age group
- No-show rate by waiting period
- No-show rate by SMS reminder
- No-show rate by clinic type
- No-show rate by appointment type
- No-show rate by day of week
- No-show rate by previous no-show history
- Appointment volume by month
- Top 10 highest-risk patient segments (cross-tabulation)

---

### Step 3 — `build_excel.py`

**Purpose:** Reads the query results and builds a fully formatted, multi-sheet Excel workbook.

#### What it does:

**1. Loads the saved query results**
```python
with open("query_results.pkl", "rb") as f:
    results = pickle.load(f)
```

**2. Defines a colour palette and style helpers**

Reusable functions handle fonts, fills, borders, and alignments so every table looks consistent:
```python
def ff(size=10, bold=False, color="1A1A1A"):
    return Font(name="Calibri", size=size, bold=bold, color=color)

def fill(hex_c):
    return PatternFill("solid", fgColor=hex_c)
```

**3. Writes 7 sheets using openpyxl**

| Sheet | Contents |
|---|---|
| Executive Summary | KPI cards + all key summary tables with colour-scale formatting |
| SQL Queries | All 10 SQL queries displayed as formatted code blocks |
| Charts | 6 embedded Excel charts (bar, column, line, pie) with underlying data |
| Analysis Tables | Full results from every query |
| High-Risk Segments | Top 10 patient-clinic-reminder combinations with highest no-show rates |
| Recommendations | 5 detailed action cards for reducing no-shows |
| Raw Data Sample | 1,000 styled rows from the appointments dataset |

**4. Applies conditional formatting**

The no-show rate columns use a green-yellow-red colour scale so high-risk rows stand out immediately:
```python
ws.conditional_formatting.add(range,
    ColorScaleRule(start_color="63BE7B",   # green = low risk
                   mid_color="FFEB84",     # yellow = medium
                   end_color="F8696B"))    # red = high risk
```

**5. Embeds Excel charts**

Charts are created using openpyxl's chart module and reference the data tables written to the sheet:
```python
ch = BarChart()
data = Reference(ws, min_col=3, min_row=1, max_row=6)
ch.add_data(data, titles_from_data=True)
ws.add_chart(ch, "E4")
```

---

## SQL Queries Explained

### Overall no-show rate
Counts total appointments and calculates the percentage that were no-shows using a `CASE WHEN` expression inside `SUM()` — a standard SQL technique for conditional counting.

```sql
SELECT 
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments;
```

### No-show rate by age group
Uses a `CASE WHEN` inside `GROUP BY` to bucket continuous age values into named groups before aggregating.

```sql
SELECT 
    CASE 
        WHEN age < 18 THEN '0-17'
        WHEN age BETWEEN 18 AND 30 THEN '18-30'
        WHEN age BETWEEN 31 AND 45 THEN '31-45'
        WHEN age BETWEEN 46 AND 60 THEN '46-60'
        ELSE '60+'
    END AS age_group,
    COUNT(*) AS total_appointments,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY age_group
ORDER BY no_show_rate DESC;
```

### No-show rate by waiting days
Same bucketing technique applied to waiting days — converts a numeric column into meaningful time bands.

```sql
SELECT 
    CASE 
        WHEN waiting_days = 0 THEN 'Same day'
        WHEN waiting_days BETWEEN 1 AND 7 THEN '1-7 days'
        WHEN waiting_days BETWEEN 8 AND 14 THEN '8-14 days'
        WHEN waiting_days BETWEEN 15 AND 30 THEN '15-30 days'
        ELSE '30+ days'
    END AS waiting_period,
    COUNT(*) AS total_appointments,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY waiting_period
ORDER BY no_show_rate DESC;
```

### No-show rate by SMS reminder
Simple `GROUP BY` on a binary column — directly compares reminded vs non-reminded patients.

```sql
SELECT 
    sms_reminder,
    COUNT(*) AS total_appointments,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY sms_reminder;
```

---

## Key Findings

| Factor | Finding |
|---|---|
| **SMS Reminder** | Patients without a reminder: **28.11%** no-show. With reminder: **18.11%**. That's a **10 percentage point** reduction. |
| **Previous No-Shows** | Patients with 3+ prior no-shows: **40.34%** no-show rate — nearly double the baseline. |
| **Waiting Days** | 15–30 day waits: **27.56%**. 30+ days: **29.43%**. Same day: only **19.42%**. |
| **Age Group** | 18–30 year olds have the highest no-show rate at **26.70%**. |
| **Day of Week** | Monday is the worst day at **23.01%**. Tuesday is the best at **19.02%**. |
| **Clinic Type** | Dermatology leads at **25.51%**, Neurology is lowest at **19.85%**. |
| **Appointment Type** | Lab Tests (**23.76%**) and Consultations (**23.73%**) are highest risk. |

---

## Recommendations

1. **Send SMS reminders to patients with longer wait times** — the reminder alone cuts no-shows by 10 percentage points. Full coverage (currently only 62%) is the single highest-ROI intervention.

2. **Prioritise follow-up reminders for patients with previous no-shows** — flag patients with 2+ prior no-shows in the booking system for a personal phone call, not just SMS.

3. **Avoid scheduling high-risk appointment types too far in advance** — Lab Tests and Consultations should be kept under 7 days where clinically safe.

4. **Add confirmation calls for appointments booked more than 14 days ahead** — no-show rate jumps sharply beyond the 14-day mark. A 7-day check-in call can reconfirm intent or free the slot.

5. **Monitor no-show trends by day of week for smarter scheduling** — consider deliberate Monday overbooking (10–15% buffer) and incentivised mid-week rescheduling.

---

## Installation & Setup

### Prerequisites
- Python 3.8 or higher
- PyCharm (Community Edition is fine) or any Python IDE

### Install required libraries

Open your terminal or PyCharm terminal and run:

```bash
pip install pandas numpy faker openpyxl
```

> `sqlite3` is built into Python — no installation needed.

### Clone or download the project

```bash
git clone https://github.com/your-username/clinic_project.git
cd clinic_project
```

---

## Run Order

The three scripts must be run **in this exact order**:

```bash
# Step 1 — Generate the dataset
python generate_data.py

# Step 2 — Run SQL analysis
python sql_analysis.py

# Step 3 — Build the Excel dashboard
python build_excel.py
```

After Step 3, open `Clinic_NoShow_Dashboard_v2.xlsx` to view the full dashboard.

To view the SQLite database, download **DB Browser for SQLite** from [sqlitebrowser.org](https://sqlitebrowser.org) — it's free and lets you browse tables and run SQL queries visually.

---

## Technologies Used

| Tool | Purpose |
|---|---|
| **Python 3** | Core programming language |
| **pandas** | Data manipulation and SQL result handling |
| **numpy** | Random number generation and probability simulation |
| **Faker** | Generating realistic fake patient IDs and names |
| **sqlite3** | Built-in Python library for SQLite database |
| **openpyxl** | Writing and formatting Excel (.xlsx) files |
| **SQLite** | Lightweight file-based database (no server needed) |

---

## Notes

- All data is **fully synthetic** — no real patient data was used
- The no-show probabilities are designed to reflect realistic healthcare patterns but are not based on any specific real-world dataset
- The project uses SQLite rather than MySQL or PostgreSQL so it runs entirely locally with no server setup required
- If you want to migrate to MySQL in the future, only the connection line needs to change — all SQL queries are standard and compatible

---

*Built as a healthcare data analytics portfolio project demonstrating SQL, Python, and Excel dashboard skills.*
