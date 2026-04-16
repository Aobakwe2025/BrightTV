"""
BrightTV Viewership Analytics - Full Analysis Script
Generates all charts and summary stats needed for the presentation.
"""

import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# 0. LOAD DATA (FIXED PATH HANDLING)
# ─────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Automatically find the dataset file
for file in os.listdir(BASE_DIR):
    if file.endswith(".xlsx"):
        FILE = os.path.join(BASE_DIR, file)
        break

print(f"📂 Using dataset: {FILE}")

df_users = pd.read_excel(FILE, sheet_name="User Profiles")
df_view  = pd.read_excel(FILE, sheet_name="Viewership")

# ─────────────────────────────────────────────
# 1. CLEAN & ENRICH
# ─────────────────────────────────────────────

df_view["RecordDate_SAST"] = pd.to_datetime(df_view["RecordDate2"]) + pd.Timedelta(hours=2)

def duration_to_seconds(val):
    try:
        if isinstance(val, str):
            parts = val.strip().split(":")
            if len(parts) == 3:
                return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        return val.hour*3600 + val.minute*60 + val.second
    except:
        return np.nan

df_view["duration_sec"] = df_view["Duration 2"].apply(duration_to_seconds)
df_view["duration_min"] = df_view["duration_sec"] / 60

df_view["hour"]        = df_view["RecordDate_SAST"].dt.hour
df_view["day_of_week"] = df_view["RecordDate_SAST"].dt.day_name()
df_view["week"]        = df_view["RecordDate_SAST"].dt.isocalendar().week.astype(int)
df_view["month"]       = df_view["RecordDate_SAST"].dt.month_name()
df_view["date"]        = df_view["RecordDate_SAST"].dt.date

df_users["AgeGroup"] = pd.cut(
    df_users["Age"].clip(0, 80),
    bins=[0, 17, 24, 34, 44, 54, 80],
    labels=["<18", "18-24", "25-34", "35-44", "45-54", "55+"]
)

df = df_view.merge(
    df_users[["UserID","Gender","Race","Age","Province","AgeGroup"]],
    on="UserID", how="left"
)

df["Gender"] = df["Gender"].str.strip().str.lower()
df["Race"]   = df["Race"].str.strip().str.lower()
df = df[df["Gender"].isin(["male","female"])]

print(f"✅ Data loaded: {len(df_view):,} records | {len(df_users):,} users")

# ─────────────────────────────────────────────
# SAVE FUNCTION (FIXED)
# ─────────────────────────────────────────────
CHART_DIR = os.path.join(BASE_DIR, "charts")
os.makedirs(CHART_DIR, exist_ok=True)

def save(fig, name):
    path = os.path.join(CHART_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"   💾 Saved {path}")

# ─────────────────────────────────────────────
# COLOURS
# ─────────────────────────────────────────────
PRIMARY   = "#1A73E8"
SECONDARY = "#EA4335"
ACCENT    = "#34A853"
WARN      = "#FBBC05"

# ─────────────────────────────────────────────
# CHART 1
# ─────────────────────────────────────────────
print("\n[1] Daily sessions trend...")
daily = df.groupby("date").agg(sessions=("UserID","count")).reset_index()
daily["date"] = pd.to_datetime(daily["date"])

fig, ax = plt.subplots()
ax.plot(daily["date"], daily["sessions"])
ax.set_title("Daily Sessions")
save(fig, "chart1_daily_sessions.png")

# ─────────────────────────────────────────────
# CHART 2
# ─────────────────────────────────────────────
print("[2] Sessions by day...")
dow = df.groupby("day_of_week")["UserID"].count()

fig, ax = plt.subplots()
ax.bar(dow.index, dow.values)
save(fig, "chart2_sessions_dow.png")

# ─────────────────────────────────────────────
# CHART 3
# ─────────────────────────────────────────────
print("[3] Heatmap...")
pivot = df.pivot_table(index="day_of_week", columns="hour", values="UserID", aggfunc="count", fill_value=0)

fig, ax = plt.subplots()
ax.imshow(pivot.values)
save(fig, "chart3_heatmap.png")

# ─────────────────────────────────────────────
# CHART 4
# ─────────────────────────────────────────────
print("[4] Channels...")
ch = df["Channel2"].value_counts().head(10)

fig, ax = plt.subplots()
ax.barh(ch.index, ch.values)
save(fig, "chart4_top_channels.png")

# ─────────────────────────────────────────────
# CHART 5
# ─────────────────────────────────────────────
print("[5] Demographics...")
g = df_users["Gender"].value_counts()

fig, ax = plt.subplots()
ax.pie(g.values, labels=g.index)
save(fig, "chart5_demographics.png")

# ─────────────────────────────────────────────
# CHART 6
# ─────────────────────────────────────────────
print("[6] Duration...")
dur = df.groupby("Channel2")["duration_min"].mean().head(10)

fig, ax = plt.subplots()
ax.bar(dur.index, dur.values)
save(fig, "chart6_avg_duration.png")

# ─────────────────────────────────────────────
# CHART 7
# ─────────────────────────────────────────────
print("[7] Consumption...")
age_g = df.groupby("AgeGroup")["duration_min"].sum()

fig, ax = plt.subplots()
ax.bar(age_g.index.astype(str), age_g.values)
save(fig, "chart7_consumption_demo.png")

# ─────────────────────────────────────────────
# CHART 9
# ─────────────────────────────────────────────
print("[8] Province vs Channel...")
prov = df.groupby("Province")["UserID"].count().head(8)

fig, ax = plt.subplots()
ax.barh(prov.index, prov.values)
save(fig, "chart9_province_channel.png")

# ─────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────
print("\nSUMMARY")
print(f"Total sessions: {len(df):,}")
print(f"Users: {df['UserID'].nunique():,}")
print(f"Avg duration: {df['duration_min'].mean():.1f} min")