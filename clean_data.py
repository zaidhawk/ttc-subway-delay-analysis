"""
Cleans the TTC subway delay data and saves it to data/clean/.
Also saves a log of each cleaning step and how many rows it affected.
"""
import re
from pathlib import Path

import pandas as pd

RAW = Path("data/raw")
CLEAN = Path("data/clean")

# The first two letters of a delay code tell you its category
CATEGORIES = {
    "EU": "Equipment (Rolling Stock)",
    "PU": "Plant (Track, Signals, Power)",
    "TU": "Transportation (Operations)",
    "SU": "Security / Passenger",
    "MU": "Miscellaneous",
}

quality_log = []


def log(check, rows, action):
    quality_log.append({"check": check, "rows_affected": int(rows), "action": action})
    print(f"  {check:<45} {rows:>7,}  -> {action}")


def load_raw():
    df_2024 = pd.read_excel(RAW / "delays_2024.xlsx")
    df_2025 = pd.read_csv(RAW / "delays_since_2025.csv").drop(columns="_id")
    df = pd.concat([df_2024, df_2025], ignore_index=True)
    df["Date"] = pd.to_datetime(df["Date"])
    return df


def load_codes():
    codes = pd.read_csv(RAW / "delay_codes.csv", encoding="utf-8")
    # Some descriptions have broken dash characters, replace them with a normal dash
    codes["DESCRIPTION"] = (
        codes["DESCRIPTION"]
        .str.replace(r"\s*â[\u0080-\u009f]*\s*", " - ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip(" -")
    )
    return codes.rename(columns={"CODE": "Code", "DESCRIPTION": "Code Description"})[
        ["Code", "Code Description"]
    ]


def standardize_line(value):
    # The Line column is typed many different ways (YU/BD, YU / BD, YUS, LINE 1...)
    if pd.isna(value):
        return "Unknown"
    v = re.sub(r"[^A-Z0-9]", "", str(value).upper())
    has_yu = "YU" in v or "UNIVERSITY" in v or v.startswith("LINE1")
    has_bd = "BD" in v or "BLOOR" in v or v.startswith("LINE2")
    has_shp = "SHP" in v or "SHEP" in v or v.startswith("LINE4")
    if sum([has_yu, has_bd, has_shp]) > 1:
        return "Multiple Lines"
    if has_yu:
        return "Line 1 Yonge-University"
    if has_bd:
        return "Line 2 Bloor-Danforth"
    if has_shp:
        return "Line 4 Sheppard"
    return "Unknown"


def clean_station(name):
    # Make station names consistent and drop the word STATION at the end
    s = re.sub(r"\s+", " ", str(name).upper()).strip()
    return re.sub(r"\s+STATION$", "", s)


def main():
    print("Loading raw data...")
    df = load_raw()
    codes = load_codes()
    raw_rows = len(df)
    log("Raw records loaded", raw_rows, "2024 file + since-2025 file combined")

    print("\nData quality checks:")

    # Remove duplicate rows
    dupes = df.duplicated()
    log("Exact duplicate records", dupes.sum(), "removed")
    df = df[~dupes].copy()

    # Clean up line names
    raw_line_values = df["Line"].nunique(dropna=False)
    df["Line Clean"] = df["Line"].apply(standardize_line)
    log(f"Distinct raw Line spellings ({raw_line_values})", raw_line_values,
        f"standardized to {df['Line Clean'].nunique()} values")
    log("Missing Line", df["Line"].isna().sum(), "set to 'Unknown'")
    log("Line unrecognized (bus routes, SRT, notes)",
        ((df["Line Clean"] == "Unknown") & df["Line"].notna()).sum(), "set to 'Unknown'")

    # Fill in missing direction
    log("Missing Bound (direction)", df["Bound"].isna().sum(), "kept, filled 'Unknown'")
    df["Bound"] = df["Bound"].fillna("Unknown")

    # Add code descriptions, some codes are not in the official list
    df = df.merge(codes, on="Code", how="left")
    missing_code = df["Code Description"].isna()
    log("Code not in official code list", missing_code.sum(),
        "kept, described as 'Not in code list'")
    df["Code Description"] = df["Code Description"].fillna("Not in code list")

    # Get the category from the code
    df["Category"] = df["Code"].str[:2].map(CATEGORIES).fillna("Unknown")

    # Some incidents were logged but did not delay the train
    df["Caused Delay"] = df["Min Delay"] > 0
    log("Incidents with 0 min delay", (~df["Caused Delay"]).sum(),
        "kept, flagged Caused Delay = False")

    # Flag very long delays
    df["Long Delay Flag"] = df["Min Delay"] >= 60
    log("Delays >= 60 min", df["Long Delay Flag"].sum(), "kept, flagged for review")

    # Clean up station names
    raw_stations = df["Station"].nunique()
    df["Station Clean"] = df["Station"].apply(clean_station)
    log(f"Distinct raw station names ({raw_stations})", raw_stations,
        f"trimmed/upper-cased to {df['Station Clean'].nunique()}")

    # Add month, year and hour columns
    df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    df["Year"] = df["Date"].dt.year
    df["Hour"] = pd.to_numeric(df["Time"].str[:2], errors="coerce")

    log("Final clean records", len(df), f"{len(df) / raw_rows:.1%} of raw rows kept")

    out = df[[
        "Date", "Month", "Year", "Day", "Time", "Hour", "Line Clean", "Station Clean",
        "Bound", "Code", "Code Description", "Category", "Min Delay", "Min Gap",
        "Caused Delay", "Long Delay Flag", "Vehicle", "Line", "Station",
    ]].rename(columns={"Line Clean": "Line", "Station Clean": "Station",
                       "Line": "Line (raw)", "Station": "Station (raw)"})

    CLEAN.mkdir(parents=True, exist_ok=True)
    out.to_csv(CLEAN / "subway_delays_clean.csv", index=False)
    pd.DataFrame(quality_log).to_csv(CLEAN / "data_quality_log.csv", index=False)
    print(f"\nSaved {len(out):,} rows to {CLEAN / 'subway_delays_clean.csv'}")


if __name__ == "__main__":
    main()
