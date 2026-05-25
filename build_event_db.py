#!/usr/bin/env python3
"""
Pivot the tall CSVs in csv_export/ into a normalized event-keyed SQLite db.

Schema:
  event           — one row per (date, name); wide financial columns
  artist          — distinct artist registry
  event_artist    — many-to-many link
  event_metric    — overflow: any metric not folded into event columns

Usage:
  python build_event_db.py [--in-dir DIR] [--db PATH]
  python build_event_db.py --help

Options:
  --in-dir DIR   Where the *_tall.csv files live. Default: ./csv_export
  --db PATH      Output db path. Default: ./events.db
"""

import csv
import re
import sqlite3
import sys
from pathlib import Path

METRIC_MAP = {
    "Gross Ticket Sales":                  "gross_ticket_sales",
    "Merchandise Sales":                   "merchandise_sales",
    "CD Sales":                            "cd_sales",
    "Food Sales":                          "food_beverage",
    "Beverage Sales":                      "food_beverage",
    "Beverage  Sales":                     "food_beverage",
    "Liquor Sales":                        "food_beverage",
    "F&B Gross Revenue":                   "food_beverage",
    "Food & Beverage":                     "food_beverage",
    "Individual Donations":                "donations",
    "Program Donations":                   "donations",
    "Grant Underwriting":                  "grant_underwriting",
    "Venue Rental":                        "venue_rental",
    "Contract Sales":                      "other_revenue",
    "Streaming Sales":                     "other_revenue",
    "Other":                               "other_revenue",
    "Other Direct Show Revenue Subtotal":  "other_revenue",
    "Performance Fee":                     "performance_fee",
    "Performers Fee":                      "performance_fee",
    "Tech Support":                        "tech_support",
    "Security":                            "security",
    "Musicians Travel Expenses":           "musicians_travel",
    "Equipment Rental":                    "equipment_rental",
    "Advertising":                         "advertising",
    "Miscellaneous Event Expense":         "misc_expense",
    "Event Supplies":                      "misc_expense",
    "Music Purchasing":                    "misc_expense",
    "Meals and Hospitality":               "misc_expense",
    "Artist Commission (Painting)":        "misc_expense",
    "CD Commissions Paid":                 "misc_expense",
    "Total Merchant Fees":                 "merchant_fees",
    "Cost of Concessions Sold (25%)":                          "food_beverage_expense",
    "F&B COGS % from Tracking Summary (w/o Labor/Waste)":     "food_beverage_expense",
    "F&B COGS % from Tracking Summary (w/o Labor/waste)":     "food_beverage_expense",
    "F&B COGS % from Tracking Summary w/o Labor":             "food_beverage_expense",
    "F&B COGS from Tracking Summary w/o Labor":               "food_beverage_expense",
    "Total Concessions Expense":                               "food_beverage_expense",
    "Tickets Sold":                        "tickets_sold",
    "Full Price Tickets Sold":             "full_price_tickets",
    "Discounted Tickets Sold":             "discount_tickets",
    "Discount Tickets Sold":               "discount_tickets",
    "Total Tickets":                       "tickets_sold",
    "Total Tickets Sold":                  "tickets_sold",
    "Comp Tickets":                        "comp_tickets",
}

REVENUE_COLS = ["gross_ticket_sales", "food_beverage", "merchandise_sales",
                "cd_sales", "donations", "grant_underwriting",
                "venue_rental", "other_revenue"]
EXPENSE_COLS = ["performance_fee", "tech_support", "security", "musicians_travel",
                "equipment_rental", "advertising", "misc_expense", "merchant_fees",
                "food_beverage_expense"]
COUNT_COLS = ["tickets_sold", "full_price_tickets", "discount_tickets", "comp_tickets"]
ALL_NUMERIC_COLS = REVENUE_COLS + EXPENSE_COLS + COUNT_COLS

SCHEMA = """
DROP TABLE IF EXISTS event_artist;
DROP TABLE IF EXISTS event_metric;
DROP TABLE IF EXISTS event;
DROP TABLE IF EXISTS artist;

CREATE TABLE event (
    event_id              INTEGER PRIMARY KEY,
    date                  TEXT NOT NULL,
    name                  TEXT NOT NULL,
    series                TEXT,
    program               TEXT,
    fiscal_year           INTEGER,
    source_file           TEXT,
    sheet                 TEXT,
    notes                 TEXT,
""" + ",\n".join(f"    {c:25s} REAL DEFAULT 0" for c in REVENUE_COLS + EXPENSE_COLS) + ",\n" + \
       ",\n".join(f"    {c:25s} INTEGER DEFAULT 0" for c in COUNT_COLS) + """,
    UNIQUE(date, name, source_file)
);

CREATE TABLE artist (
    artist_id   INTEGER PRIMARY KEY,
    name        TEXT NOT NULL UNIQUE
);

CREATE TABLE event_artist (
    event_id    INTEGER NOT NULL REFERENCES event(event_id),
    artist_id   INTEGER NOT NULL REFERENCES artist(artist_id),
    role        TEXT,
    sort_order  INTEGER DEFAULT 0,
    share_pct   REAL DEFAULT 100,
    PRIMARY KEY (event_id, artist_id)
);

CREATE TABLE event_metric (
    event_id    INTEGER NOT NULL REFERENCES event(event_id),
    metric      TEXT NOT NULL,
    value       REAL,
    PRIMARY KEY (event_id, metric)
);

CREATE INDEX idx_event_date     ON event(date);
CREATE INDEX idx_event_series   ON event(series);
CREATE INDEX idx_ea_artist      ON event_artist(artist_id);
"""


def classify(source_file: str):
    s = source_file.lower()
    if "jam_sessions" in s:
        return ("Jam Sessions", "Performance")
    if "thursday_shows" in s:
        return ("Thursday Shows", "Performance")
    if "first_friday" in s:
        return ("First Friday", "Performance")
    if "education" in s:
        return (None, "Education")
    if "special_event" in s:
        return (None, "Special Event")
    if "performance" in s:
        return (None, "Performance")
    return (None, None)


def fiscal_year(date_str: str) -> int:
    y = int(date_str[:4])
    m = int(date_str[5:7])
    return y + 1 if m >= 7 else y


def main():
    if len(sys.argv) > 1 and any(a in ("-h", "--help") for a in sys.argv[1:]):
        print(__doc__)
        sys.exit(0)

    in_dir = Path("csv_export")
    db_path = Path("events.db")
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--in-dir" and i + 1 < len(args):
            in_dir = Path(args[i + 1]); i += 2
        elif args[i] == "--db" and i + 1 < len(args):
            db_path = Path(args[i + 1]); i += 2
        else:
            i += 1

    tall_files = sorted(in_dir.glob("*_tall.csv"))
    if not tall_files:
        print(f"No *_tall.csv files in {in_dir}/", file=sys.stderr)
        sys.exit(1)

    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)

    events = {}     # (date, name, source_file) -> dict of accumulated columns
    overflow = []   # (date, name, source_file, metric, value)

    for path in tall_files:
        with path.open() as f:
            reader = csv.DictReader(f)
            for row in reader:
                date, show, artist = row["date"], row["show"], row["artist"]
                metric = row["metric"]
                source_file = row["source_file"]
                sheet = row["sheet"]
                try:
                    value = float(row["value"])
                except (TypeError, ValueError):
                    continue

                key = (date, show, source_file)
                if key not in events:
                    series, program = classify(source_file)
                    events[key] = {
                        "date": date, "name": show, "series": series,
                        "program": program, "fiscal_year": fiscal_year(date),
                        "source_file": source_file, "sheet": sheet,
                        "artist": artist,
                    }
                    for c in ALL_NUMERIC_COLS:
                        events[key][c] = 0

                col = METRIC_MAP.get(metric)
                if col:
                    # food_beverage_expense can appear under multiple names in the
                    # same sheet (e.g. both "Total Concessions Expense" and an F&B
                    # COGS line) with identical values — take max to avoid doubling.
                    if col == "food_beverage_expense":
                        events[key][col] = max(events[key][col], value)
                    else:
                        events[key][col] += value
                else:
                    overflow.append((key, metric, value))

    cur = conn.cursor()
    key_to_id = {}
    for key, e in events.items():
        cols = ["date", "name", "series", "program", "fiscal_year",
                "source_file", "sheet"] + REVENUE_COLS + EXPENSE_COLS + COUNT_COLS
        placeholders = ",".join("?" * len(cols))
        cur.execute(
            f"INSERT INTO event ({','.join(cols)}) VALUES ({placeholders})",
            [e[c] for c in cols],
        )
        key_to_id[key] = cur.lastrowid

    artists = {}
    for key, e in events.items():
        a = (e["artist"] or e["name"] or "").strip()
        if not a:
            continue
        if a not in artists:
            cur.execute("INSERT INTO artist (name) VALUES (?)", [a])
            artists[a] = cur.lastrowid
        cur.execute(
            "INSERT INTO event_artist (event_id, artist_id, sort_order, share_pct) VALUES (?, ?, 0, 100)",
            [key_to_id[key], artists[a]],
        )

    for key, metric, value in overflow:
        eid = key_to_id.get(key)
        if not eid:
            continue
        cur.execute(
            "INSERT OR REPLACE INTO event_metric (event_id, metric, value) "
            "VALUES (?, ?, COALESCE((SELECT value FROM event_metric WHERE event_id=? AND metric=?),0)+?)",
            [eid, metric, eid, metric, value],
        )

    conn.commit()
    n_events = cur.execute("SELECT COUNT(*) FROM event").fetchone()[0]
    n_artists = cur.execute("SELECT COUNT(*) FROM artist").fetchone()[0]
    n_links = cur.execute("SELECT COUNT(*) FROM event_artist").fetchone()[0]
    n_overflow = cur.execute("SELECT COUNT(*) FROM event_metric").fetchone()[0]
    conn.close()

    print(f"Wrote {db_path}")
    print(f"  events:        {n_events}")
    print(f"  artists:       {n_artists}")
    print(f"  event_artist:  {n_links}")
    print(f"  event_metric:  {n_overflow} overflow rows")


if __name__ == "__main__":
    main()
