"""
BrightTV Viewership Analytics - Full Analysis Script
Generates all charts and summary stats needed for the presentation.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# 0. LOAD DATA
# ─────────────────────────────────────────────
FILE = "/mnt/user-data/uploads/1774896982744_Bright_TV_-Dataset.xlsx"

df_users = pd.read_excel(FILE, sheet_name="User Profiles")
df_view  = pd.read_excel(FILE, sheet_name="Viewership")

# ─────────────────────────────────────────────
# 1. CLEAN & ENRICH
# ─────────────────────────────────────────────

# Convert UTC → SAST (UTC+2)
df_view["RecordDate_SAST"] = pd.to_datetime(df_view["RecordDate2"]) + pd.Timedelta(hours=2)

# Parse duration to seconds
def duration_to_seconds(val):
    try:
        if isinstance(val, str):
            parts = val.strip().split(":")
            if len(parts) == 3:
                return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        # datetime.time object
        return val.hour*3600 + val.minute*60 + val.second
    except:
        return np.nan

df_view["duration_sec"] = df_view["Duration 2"].apply(duration_to_seconds)
df_view["duration_min"] = df_view["duration_sec"] / 60

# Time features
df_view["hour"]       = df_view["RecordDate_SAST"].dt.hour
df_view["day_of_week"]= df_view["RecordDate_SAST"].dt.day_name()
df_view["week"]       = df_view["RecordDate_SAST"].dt.isocalendar().week.astype(int)
df_view["month"]      = df_view["RecordDate_SAST"].dt.month_name()
df_view["date"]       = df_view["RecordDate_SAST"].dt.date

# Age bins
df_users["AgeGroup"] = pd.cut(
    df_users["Age"].clip(0, 80),
    bins=[0, 17, 24, 34, 44, 54, 80],
    labels=["<18", "18-24", "25-34", "35-44", "45-54", "55+"]
)

# Merge
df = df_view.merge(df_users[["UserID","Gender","Race","Age","Province","AgeGroup"]],
                   on="UserID", how="left")

# Clean gender & race labels
df["Gender"] = df["Gender"].str.strip().str.lower()
df["Race"]   = df["Race"].str.strip().str.lower()
df = df[df["Gender"].isin(["male","female"])]

print(f"✅ Data loaded: {len(df_view):,} viewership records | {len(df_users):,} user profiles")
print(f"   Date range (SAST): {df['RecordDate_SAST'].min().date()} → {df['RecordDate_SAST'].max().date()}")
print(f"   Channels: {df['Channel2'].nunique()} | Unique users in views: {df['UserID'].nunique()}")

# ─────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────
PRIMARY   = "#1A73E8"
SECONDARY = "#EA4335"
ACCENT    = "#34A853"
WARN      = "#FBBC05"
BG        = "#F8F9FA"
PALETTE   = [PRIMARY, SECONDARY, ACCENT, WARN, "#9C27B0", "#00BCD4", "#FF5722"]

plt.rcParams.update({
    "figure.facecolor": BG,
    "axes.facecolor"  : "white",
    "axes.spines.top" : False,
    "axes.spines.right": False,
    "font.family"     : "DejaVu Sans",
    "axes.titlesize"  : 13,
    "axes.labelsize"  : 11,
})

def save(fig, name):
    fig.savefig(f"/home/claude/{name}", dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"   💾 Saved {name}")

# ─────────────────────────────────────────────
# CHART 1 – Sessions per day (time series)
# ─────────────────────────────────────────────
print("\n[1] Daily sessions trend...")
daily = df.groupby("date").agg(sessions=("UserID","count"), users=("UserID","nunique")).reset_index()
daily["date"] = pd.to_datetime(daily["date"])

fig, ax = plt.subplots(figsize=(12,4))
ax.fill_between(daily["date"], daily["sessions"], alpha=0.25, color=PRIMARY)
ax.plot(daily["date"], daily["sessions"], color=PRIMARY, linewidth=2)
ax.set_title("Daily Viewership Sessions (Jan–Mar 2016, SAST)")
ax.set_xlabel("Date"); ax.set_ylabel("Sessions")
# Annotate weekends lightly
for _, row in daily.iterrows():
    if row["date"].weekday() >= 5:
        ax.axvspan(row["date"]-pd.Timedelta(hours=12), row["date"]+pd.Timedelta(hours=12),
                   alpha=0.06, color=WARN)
ax.set_facecolor("white")
fig.tight_layout()
save(fig, "chart1_daily_sessions.png")

# ─────────────────────────────────────────────
# CHART 2 – Sessions by day of week
# ─────────────────────────────────────────────
print("[2] Sessions by day of week...")
dow_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
dow = df.groupby("day_of_week")["UserID"].count().reindex(dow_order)

fig, ax = plt.subplots(figsize=(9,4))
bars = ax.bar(dow.index, dow.values,
              color=[SECONDARY if d in ["Saturday","Sunday"] else PRIMARY for d in dow.index],
              edgecolor="white", width=0.6)
ax.bar_label(bars, fmt="{:,.0f}", padding=4, fontsize=9)
ax.set_title("Total Sessions by Day of Week")
ax.set_ylabel("Sessions")
ax.set_ylim(0, dow.max()*1.15)
save(fig, "chart2_sessions_dow.png")

# ─────────────────────────────────────────────
# CHART 3 – Hourly viewing pattern (heatmap by dow)
# ─────────────────────────────────────────────
print("[3] Hourly heatmap...")
pivot = df.pivot_table(index="day_of_week", columns="hour", values="UserID",
                       aggfunc="count", fill_value=0).reindex(dow_order)

fig, ax = plt.subplots(figsize=(14,5))
im = ax.imshow(pivot.values, aspect="auto", cmap="YlOrRd")
ax.set_xticks(range(24)); ax.set_xticklabels([f"{h:02d}:00" for h in range(24)], rotation=45, ha="right", fontsize=8)
ax.set_yticks(range(7));  ax.set_yticklabels(dow_order)
ax.set_title("Viewing Intensity – Day of Week × Hour (SAST)")
plt.colorbar(im, ax=ax, label="Sessions")
save(fig, "chart3_heatmap.png")

# ─────────────────────────────────────────────
# CHART 4 – Top channels by sessions & watch time
# ─────────────────────────────────────────────
print("[4] Top channels...")
ch = df.groupby("Channel2").agg(sessions=("UserID","count"),
                                 watch_min=("duration_min","sum")).sort_values("sessions", ascending=False).head(12)

fig, axes = plt.subplots(1,2, figsize=(14,5))
for ax, col, label, color in zip(axes,
        ["sessions","watch_min"], ["Sessions","Total Watch Time (min)"],
        [PRIMARY, ACCENT]):
    vals = ch[col].sort_values()
    ax.barh(vals.index, vals.values, color=color, edgecolor="white")
    ax.set_title(f"Top Channels by {label}")
    ax.set_xlabel(label)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
fig.tight_layout()
save(fig, "chart4_top_channels.png")

# ─────────────────────────────────────────────
# CHART 5 – User demographics (gender, age, province, race)
# ─────────────────────────────────────────────
print("[5] Demographics...")
fig, axes = plt.subplots(2, 2, figsize=(14,10))

# Gender
g = df_users["Gender"].str.strip().str.lower().value_counts()
g = g[g.index.isin(["male","female"])]
axes[0,0].pie(g.values, labels=g.index.str.title(), autopct="%1.1f%%",
              colors=[PRIMARY, SECONDARY], startangle=90)
axes[0,0].set_title("Subscribers by Gender")

# Age group
ag = df_users["AgeGroup"].value_counts().sort_index()
axes[0,1].bar(ag.index.astype(str), ag.values, color=ACCENT, edgecolor="white")
axes[0,1].set_title("Subscribers by Age Group")
axes[0,1].set_ylabel("Count")

# Province
prov = df_users["Province"].str.strip().value_counts().head(8)
axes[1,0].barh(prov.index, prov.values, color=PALETTE[4], edgecolor="white")
axes[1,0].set_title("Subscribers by Province")
axes[1,0].set_xlabel("Count")

# Race
race = df_users["Race"].str.strip().str.lower().value_counts()
race = race[race.index != ""]
axes[1,1].bar(race.index.str.title(), race.values,
              color=PALETTE[:len(race)], edgecolor="white")
axes[1,1].set_title("Subscribers by Race")
axes[1,1].set_ylabel("Count")
axes[1,1].tick_params(axis="x", rotation=20)

fig.suptitle("BrightTV – Subscriber Demographics", fontsize=15, fontweight="bold", y=1.01)
fig.tight_layout()
save(fig, "chart5_demographics.png")

# ─────────────────────────────────────────────
# CHART 6 – Avg session duration by channel
# ─────────────────────────────────────────────
print("[6] Avg duration by channel...")
dur = df.groupby("Channel2")["duration_min"].mean().sort_values(ascending=False).head(15)

fig, ax = plt.subplots(figsize=(10,5))
bars = ax.bar(dur.index, dur.values, color=WARN, edgecolor="white")
ax.set_title("Average Session Duration by Channel (minutes)")
ax.set_ylabel("Avg Minutes")
ax.tick_params(axis="x", rotation=45)
ax.bar_label(bars, fmt="{:.1f}", padding=4, fontsize=8)
fig.tight_layout()
save(fig, "chart6_avg_duration.png")

# ─────────────────────────────────────────────
# CHART 7 – Consumption by age group & gender
# ─────────────────────────────────────────────
print("[7] Consumption by demo...")
age_g = df.dropna(subset=["AgeGroup","Gender"]).groupby(["AgeGroup","Gender"])["duration_min"].sum().unstack(fill_value=0)

fig, ax = plt.subplots(figsize=(10,5))
age_g.plot(kind="bar", ax=ax, color=[PRIMARY, SECONDARY], edgecolor="white", width=0.7)
ax.set_title("Total Watch Time by Age Group & Gender (minutes)")
ax.set_ylabel("Total Minutes Watched")
ax.set_xlabel("Age Group")
ax.tick_params(axis="x", rotation=0)
ax.legend(title="Gender", title_fontsize=9)
fig.tight_layout()
save(fig, "chart7_consumption_demo.png")

# ─────────────────────────────────────────────
# CHART 8 – Weekly sessions trend
# ─────────────────────────────────────────────
print("[8] Weekly trend...")
weekly = df.groupby("week").agg(sessions=("UserID","count"), watch_min=("duration_min","sum")).reset_index()

fig, ax1 = plt.subplots(figsize=(10,4))
ax2 = ax1.twinx()
ax1.bar(weekly["week"], weekly["sessions"], color=PRIMARY, alpha=0.7, label="Sessions")
ax2.plot(weekly["week"], weekly["watch_min"]/60, color=SECONDARY, linewidth=2.5,
         marker="o", label="Watch Hours")
ax1.set_title("Weekly Sessions & Total Watch Hours")
ax1.set_xlabel("Week Number"); ax1.set_ylabel("Sessions", color=PRIMARY)
ax2.set_ylabel("Watch Hours", color=SECONDARY)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1+lines2, labels1+labels2, loc="upper left")
fig.tight_layout()
save(fig, "chart8_weekly_trend.png")

# ─────────────────────────────────────────────
# CHART 9 – Province vs channel preference (top 6)
# ─────────────────────────────────────────────
print("[9] Province × channel heatmap...")
top6_ch = df["Channel2"].value_counts().head(6).index
prov_ch = df[df["Channel2"].isin(top6_ch)].groupby(["Province","Channel2"])["UserID"].count().unstack(fill_value=0)
prov_ch = prov_ch.loc[prov_ch.sum(axis=1).sort_values(ascending=False).head(8).index]

fig, ax = plt.subplots(figsize=(12,5))
im = ax.imshow(prov_ch.values, aspect="auto", cmap="Blues")
ax.set_xticks(range(len(prov_ch.columns))); ax.set_xticklabels(prov_ch.columns, rotation=30, ha="right", fontsize=9)
ax.set_yticks(range(len(prov_ch.index)));   ax.set_yticklabels(prov_ch.index)
ax.set_title("Sessions by Province × Top Channels")
plt.colorbar(im, ax=ax, label="Sessions")
for i in range(len(prov_ch.index)):
    for j in range(len(prov_ch.columns)):
        ax.text(j, i, str(prov_ch.values[i,j]), ha="center", va="center", fontsize=8,
                color="white" if prov_ch.values[i,j] > prov_ch.values.max()*0.6 else "black")
save(fig, "chart9_province_channel.png")

# ─────────────────────────────────────────────
# PRINT SUMMARY STATS
# ─────────────────────────────────────────────
print("\n" + "="*55)
print("SUMMARY STATISTICS")
print("="*55)
print(f"Total sessions:          {len(df):,}")
print(f"Unique subscribers:      {df['UserID'].nunique():,}")
print(f"Total watch time:        {df['duration_min'].sum()/60:,.0f} hours")
print(f"Avg session duration:    {df['duration_min'].mean():.1f} min")
print(f"Channels available:      {df['Channel2'].nunique()}")
print(f"Top channel:             {df['Channel2'].value_counts().index[0]}")
print(f"Peak hour (SAST):        {df['hour'].value_counts().index[0]}:00")
print(f"Lowest session day:      {dow.idxmin()} ({dow.min():,} sessions)")
print(f"Highest session day:     {dow.idxmax()} ({dow.max():,} sessions)")
print("="*55)
