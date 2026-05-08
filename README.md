# OBX Marathon Training Plan

A Python script that generates a personalized 26-week marathon training plan as a formatted Excel spreadsheet.

## Goal
Sub-3:30 finish at the Outer Banks Marathon (November 8, 2025), targeting an 8:00/mile pace.

## What It Generates
- **26-week plan** broken into Base, Build, Peak, and Taper phases
- Daily workouts across 6 training days (5 runs + 2 lifting sessions)
- Color-coded Excel spreadsheet with weekly mileage, workout types, and lifting focus
- Separate pacing guide and training notes sheet

## Usage
```bash
pip install openpyxl
python marathon_plan.py
```
Open `OBX_Marathon_Training_Plan.xlsx` in Excel or Google Sheets.

## Training Structure
- **Weeks 1–6 (Base):** Rebuild aerobic base from ~12 to 38 miles/week
- **Weeks 7–13 (Build):** Introduce tempo runs and intervals, long runs to 18 miles
- **Weeks 14–21 (Peak):** Peak mileage (~55 mpw), two 20-milers, race-specific workouts
- **Weeks 22–26 (Taper):** Reduce volume, sharpen speed, arrive fresh on race day
