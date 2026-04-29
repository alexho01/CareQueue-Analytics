"""
Clinic No-Show Project — Data Generator
Produces appointments.csv + clinic.db (SQLite)
"""
import pandas as pd
import numpy as np
from faker import Faker
import sqlite3, os, random
from datetime import datetime, timedelta

fake = Faker()
np.random.seed(42)
random.seed(42)
N = 5000

CLINICS      = ["General Practice", "Cardiology", "Pediatrics", "Orthopedics", "Dermatology", "Neurology"]
APT_TYPES    = ["Consultation", "Follow-up", "Routine Check-up", "Emergency", "Procedure", "Lab Test"]
GENDERS      = ["Male", "Female", "Other"]
INSURANCE    = ["Private", "Public", "None", "Medicare", "Medicaid"]
NEIGHBOURHOODS = [
    "Westside", "Eastgate", "Northpark", "Southfield", "Midtown",
    "Riverside", "Hillcrest", "Oakwood", "Lakeview", "Greenhill"
]
DAYS_ORDER   = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]

# Base no-show probabilities per dimension
CLINIC_P     = {"General Practice":0.24,"Cardiology":0.17,"Pediatrics":0.20,
                "Orthopedics":0.27,"Dermatology":0.33,"Neurology":0.21}
APTYPE_P     = {"Consultation":0.26,"Follow-up":0.19,"Routine Check-up":0.22,
                "Emergency":0.10,"Procedure":0.23,"Lab Test":0.28}
DAY_P        = {"Monday":0.30,"Tuesday":0.23,"Wednesday":0.19,
                "Thursday":0.20,"Friday":0.26,"Saturday":0.24}

# Cancellation base rate (lower than no-show)
CANCEL_BASE  = 0.10

START_DATE   = datetime(2023, 1, 1)
END_DATE     = datetime(2024, 12, 31)

def rand_date(start, end):
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))

records = []
for i in range(N):
    age      = int(np.clip(np.random.normal(42, 18), 1, 95))
    gender   = random.choices(GENDERS, weights=[0.47, 0.50, 0.03])[0]
    clinic   = random.choice(CLINICS)
    apt_type = random.choice(APT_TYPES)
    insurance= random.choices(INSURANCE, weights=[0.30,0.25,0.15,0.20,0.10])[0]
    neighbourhood = random.choice(NEIGHBOURHOODS)
    sms      = random.random() < 0.62
    prev_ns  = int(np.clip(np.random.poisson(0.6), 0, 5))

    # Scheduled date (when patient booked)
    sched_date = rand_date(START_DATE, END_DATE - timedelta(days=60))
    # Wait days
    wait_days  = int(np.clip(np.random.exponential(11), 0, 60))
    apt_date   = sched_date + timedelta(days=wait_days)
    if apt_date > END_DATE:
        apt_date = END_DATE
        wait_days = (apt_date - sched_date).days

    apt_day  = apt_date.strftime("%A")
    if apt_day == "Sunday":   # no Sunday clinics
        apt_date += timedelta(days=1)
        apt_day = "Monday"
    apt_time = f"{random.randint(8,17):02d}:{random.choice(['00','15','30','45'])}"

    # No-show probability
    p_ns = (CLINIC_P[clinic] + APTYPE_P[apt_type] + DAY_P.get(apt_day, 0.22)) / 3
    if wait_days > 14: p_ns += 0.09
    if wait_days > 30: p_ns += 0.05
    if sms:            p_ns -= 0.11
    if prev_ns >= 2:   p_ns += 0.12
    if age < 30:       p_ns += 0.05
    if insurance == "None": p_ns += 0.04
    p_ns = float(np.clip(p_ns, 0.04, 0.88))

    # Determine status
    r = random.random()
    if r < p_ns:
        status = "No-Show"
    elif r < p_ns + CANCEL_BASE:
        status = "Cancelled"
    else:
        status = "Showed"

    records.append({
        "Appointment_ID":   f"APT{10000+i}",
        "Patient_ID":       f"PAT{fake.unique.random_int(min=1000, max=99999):05d}",
        "Age":              age,
        "Gender":           gender,
        "Appointment_Date": apt_date.strftime("%Y-%m-%d"),
        "Scheduled_Date":   sched_date.strftime("%Y-%m-%d"),
        "Waiting_Days":     wait_days,
        "Appointment_Day":  apt_day,
        "Appointment_Time": apt_time,
        "Clinic_Type":      clinic,
        "Appointment_Type": apt_type,
        "SMS_Reminder":     "Yes" if sms else "No",
        "Previous_No_Shows":prev_ns,
        "Insurance_Type":   insurance,
        "Neighbourhood":    neighbourhood,
        "Show_Status":      status,
    })

df = pd.DataFrame(records)

os.makedirs(".", exist_ok=True)
df.to_csv("appointments.csv", index=False)
conn = sqlite3.connect("clinic.db")
df.to_sql("appointments", conn, if_exists="replace", index=False)
conn.close()

ns  = (df["Show_Status"]=="No-Show").mean()
can = (df["Show_Status"]=="Cancelled").mean()
print(f"Generated {N} records")
print(f"  No-Show rate  : {ns:.1%}")
print(f"  Cancelled rate: {can:.1%}")
print(f"  Showed rate   : {1-ns-can:.1%}")
print(df.head(3).to_string())
