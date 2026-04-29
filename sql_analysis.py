"""
SQL Analysis — Clinic No-Show Dashboard
Runs all required queries and saves results for Excel/dashboard use
"""
import sqlite3, pickle
import pandas as pd

conn = sqlite3.connect("clinic.db")

QUERIES = {

"overall_noshow_rate": (
"Overall No-Show Rate",
"""SELECT 
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    SUM(CASE WHEN show_status = 'Cancelled' THEN 1 ELSE 0 END) AS cancellations,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate,
    ROUND(
        SUM(CASE WHEN show_status = 'Cancelled' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS cancellation_rate,
    ROUND(AVG(waiting_days), 1) AS avg_waiting_days
FROM appointments;"""),

"noshow_by_age_group": (
"No-Show Rate by Age Group",
"""SELECT 
    CASE 
        WHEN age < 18 THEN '0-17'
        WHEN age BETWEEN 18 AND 30 THEN '18-30'
        WHEN age BETWEEN 31 AND 45 THEN '31-45'
        WHEN age BETWEEN 46 AND 60 THEN '46-60'
        ELSE '60+'
    END AS age_group,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY age_group
ORDER BY no_show_rate DESC;"""),

"noshow_by_waiting_days": (
"No-Show Rate by Waiting Period",
"""SELECT 
    CASE 
        WHEN waiting_days = 0 THEN 'Same day'
        WHEN waiting_days BETWEEN 1 AND 7 THEN '1-7 days'
        WHEN waiting_days BETWEEN 8 AND 14 THEN '8-14 days'
        WHEN waiting_days BETWEEN 15 AND 30 THEN '15-30 days'
        ELSE '30+ days'
    END AS waiting_period,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY waiting_period
ORDER BY no_show_rate DESC;"""),

"noshow_by_sms": (
"No-Show Rate by SMS Reminder",
"""SELECT 
    sms_reminder,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY sms_reminder
ORDER BY no_show_rate DESC;"""),

"noshow_by_clinic": (
"No-Show Rate by Clinic Type",
"""SELECT 
    clinic_type,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY clinic_type
ORDER BY no_show_rate DESC;"""),

"noshow_by_appt_type": (
"No-Show Rate by Appointment Type",
"""SELECT 
    appointment_type,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY appointment_type
ORDER BY no_show_rate DESC;"""),

"noshow_by_day": (
"No-Show Rate by Day of Week",
"""SELECT 
    appointment_day,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY appointment_day
ORDER BY CASE appointment_day
    WHEN 'Monday' THEN 1 WHEN 'Tuesday' THEN 2 WHEN 'Wednesday' THEN 3
    WHEN 'Thursday' THEN 4 WHEN 'Friday' THEN 5 WHEN 'Saturday' THEN 6
END;"""),

"noshow_by_previous_ns": (
"No-Show Rate by Previous No-Show History",
"""SELECT 
    CASE 
        WHEN previous_no_shows = 0 THEN '0 (none)'
        WHEN previous_no_shows = 1 THEN '1'
        WHEN previous_no_shows = 2 THEN '2'
        ELSE '3+' 
    END AS prev_noshow_group,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY prev_noshow_group
ORDER BY no_show_rate DESC;"""),

"volume_by_month": (
"Appointment Volume by Month",
"""SELECT 
    SUBSTR(appointment_date, 1, 7) AS month,
    COUNT(*) AS total_appointments,
    SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) AS no_shows,
    SUM(CASE WHEN show_status = 'Cancelled' THEN 1 ELSE 0 END) AS cancellations,
    SUM(CASE WHEN show_status = 'Showed' THEN 1 ELSE 0 END) AS showed,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY month
ORDER BY month;"""),

"status_distribution": (
"Show Status Distribution",
"""SELECT 
    show_status,
    COUNT(*) AS count,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM appointments), 2) AS percentage
FROM appointments
GROUP BY show_status
ORDER BY count DESC;"""),

"high_risk_segments": (
"Top 10 Highest-Risk Segments",
"""SELECT 
    CASE 
        WHEN age < 18 THEN '0-17'
        WHEN age BETWEEN 18 AND 30 THEN '18-30'
        WHEN age BETWEEN 31 AND 45 THEN '31-45'
        WHEN age BETWEEN 46 AND 60 THEN '46-60'
        ELSE '60+'
    END AS age_group,
    clinic_type,
    sms_reminder,
    COUNT(*) AS appts,
    ROUND(
        SUM(CASE WHEN show_status = 'No-Show' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 2
    ) AS no_show_rate
FROM appointments
GROUP BY age_group, clinic_type, sms_reminder
HAVING appts >= 25
ORDER BY no_show_rate DESC
LIMIT 10;"""),
}

results = {}
for key, (title, sql) in QUERIES.items():
    df = pd.read_sql_query(sql, conn)
    results[key] = {"title": title, "sql": sql, "df": df}
    print(f"\n{'='*60}")
    print(f"  {title}")
    print('='*60)
    print(df.to_string(index=False))

conn.close()

with open("query_results.pkl", "wb") as f:
    pickle.dump(results, f)

print("\nAll SQL queries complete.")
