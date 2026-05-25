#!/usr/bin/env python3
"""
Local review + entry site for events.db.

Pages:
  /                  list & filter events
  /event/new         create a new event
  /event/<id>        event detail
  /event/<id>/edit   edit an event
  /artists           artist directory
  /artist/<id>       artist detail (per-artist totals reflect share_pct)

Usage:
  python app.py                    # serve on http://127.0.0.1:5050
  python app.py --port 8000
  python app.py --db other.db
  python app.py --help
"""

import base64
import datetime as _dt
import os
import secrets
import sqlite3
import sys
from functools import wraps
from pathlib import Path

import bcrypt
from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   session, url_for)

DB_PATH = Path(os.environ.get("NASH_DB_PATH", "events.db"))

# ── auth config ───────────────────────────────────────────────────────────
AUTH_USER = os.environ.get("NASH_USERNAME", "admin")
_raw_hash = os.environ.get("NASH_PASSWORD_HASH", "")
# b64: prefix lets callers store the bcrypt hash base64-encoded to avoid
# docker-compose interpolating the $ signs in $2b$12$<salt> hashes.
if _raw_hash.startswith("b64:"):
    _raw_hash = base64.b64decode(_raw_hash[4:]).decode()
AUTH_HASH = _raw_hash.encode()
SECRET_KEY = os.environ.get("SECRET_KEY", "")


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper

REVENUE_COLS = [
    ("gross_ticket_sales", "Ticket Sales"),
    ("food_beverage",      "F&B"),
    ("merchandise_sales",  "Merch"),
    ("cd_sales",           "CDs"),
    ("donations",          "Donations"),
    ("grant_underwriting", "Grants"),
    ("venue_rental",       "Venue Rental"),
    ("other_revenue",      "Other Rev"),
]
EXPENSE_COLS = [
    ("performance_fee",       "Performance Fee"),
    ("tech_support",          "Tech"),
    ("security",              "Security"),
    ("musicians_travel",      "Musicians Travel"),
    ("equipment_rental",      "Equipment Rental"),
    ("advertising",           "Advertising"),
    ("misc_expense",          "Misc Expense"),
    ("merchant_fees",         "Merchant Fees"),
    ("food_beverage_expense", "F&B Expense"),
]
COUNT_COLS = [
    ("full_price_tickets", "Full Price"),
    ("discount_tickets",   "Discount"),
    ("comp_tickets",       "Comp"),
]
PROGRAMS = ["Performance", "Special Event", "Education"]
ROLES = ["headliner", "support", "sit-in", "opener", "host"]

REV_SUM = " + ".join(c for c, _ in REVENUE_COLS)
EXP_SUM = " + ".join(c for c, _ in EXPENSE_COLS)


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def ensure_schema(db):
    """Create tables if absent; add columns that may be missing on older dbs."""
    db.executescript("""
        CREATE TABLE IF NOT EXISTS event (
            event_id              INTEGER PRIMARY KEY,
            date                  TEXT NOT NULL,
            name                  TEXT NOT NULL,
            series                TEXT,
            program               TEXT,
            fiscal_year           INTEGER,
            source_file           TEXT,
            sheet                 TEXT,
            notes                 TEXT,
            gross_ticket_sales    REAL DEFAULT 0,
            food_beverage         REAL DEFAULT 0,
            merchandise_sales     REAL DEFAULT 0,
            cd_sales              REAL DEFAULT 0,
            donations             REAL DEFAULT 0,
            grant_underwriting    REAL DEFAULT 0,
            venue_rental          REAL DEFAULT 0,
            other_revenue         REAL DEFAULT 0,
            performance_fee       REAL DEFAULT 0,
            tech_support          REAL DEFAULT 0,
            security              REAL DEFAULT 0,
            musicians_travel      REAL DEFAULT 0,
            equipment_rental      REAL DEFAULT 0,
            advertising           REAL DEFAULT 0,
            misc_expense          REAL DEFAULT 0,
            merchant_fees         REAL DEFAULT 0,
            food_beverage_expense REAL DEFAULT 0,
            tickets_sold          INTEGER DEFAULT 0,
            full_price_tickets    INTEGER DEFAULT 0,
            discount_tickets      INTEGER DEFAULT 0,
            comp_tickets          INTEGER DEFAULT 0,
            UNIQUE(date, name, source_file)
        );
        CREATE TABLE IF NOT EXISTS artist (
            artist_id INTEGER PRIMARY KEY,
            name      TEXT NOT NULL UNIQUE
        );
        CREATE TABLE IF NOT EXISTS event_artist (
            event_id  INTEGER NOT NULL REFERENCES event(event_id),
            artist_id INTEGER NOT NULL REFERENCES artist(artist_id),
            role      TEXT,
            sort_order INTEGER DEFAULT 0,
            share_pct  REAL DEFAULT 100,
            PRIMARY KEY (event_id, artist_id)
        );
        CREATE TABLE IF NOT EXISTS event_metric (
            event_id INTEGER NOT NULL REFERENCES event(event_id),
            metric   TEXT NOT NULL,
            value    REAL,
            PRIMARY KEY (event_id, metric)
        );
        CREATE INDEX IF NOT EXISTS idx_event_date   ON event(date);
        CREATE INDEX IF NOT EXISTS idx_event_series ON event(series);
        CREATE INDEX IF NOT EXISTS idx_ea_artist    ON event_artist(artist_id);
    """)
    # Add columns that may be missing on older dbs
    cols = {r[1] for r in db.execute("PRAGMA table_info(event_artist)")}
    if "share_pct" not in cols:
        db.execute("ALTER TABLE event_artist ADD COLUMN share_pct REAL DEFAULT 100")
        db.execute("UPDATE event_artist SET share_pct = 100 WHERE share_pct IS NULL")
    ev_cols = {r[1] for r in db.execute("PRAGMA table_info(event)")}
    if "notes" not in ev_cols:
        db.execute("ALTER TABLE event ADD COLUMN notes TEXT")
    if "food_beverage_expense" not in ev_cols:
        db.execute("ALTER TABLE event ADD COLUMN food_beverage_expense REAL DEFAULT 0")
    db.commit()


def fiscal_year(date_str: str) -> int:
    y = int(date_str[:4]); m = int(date_str[5:7])
    return y + 1 if m >= 7 else y


def parse_float(v, default=0.0):
    if v is None or v == "":
        return default
    try:
        return float(str(v).replace(",", "").replace("$", "").strip())
    except (TypeError, ValueError):
        return default


def parse_int(v, default=0):
    if v is None or v == "":
        return default
    try:
        return int(float(str(v).replace(",", "").strip()))
    except (TypeError, ValueError):
        return default


def create_app():
    app = Flask(__name__)
    app.secret_key = SECRET_KEY or "nash-local-dev-CHANGE-ME"
    app.permanent_session_lifetime = _dt.timedelta(days=14)

    with app.app_context():
        ensure_schema(get_db())

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if not AUTH_HASH:
            return ("Auth not configured. Set NASH_PASSWORD_HASH "
                    "(generate with: python make_password.py).", 500)
        if request.method == "POST":
            u = request.form.get("username", "").strip()
            p = request.form.get("password", "").encode()
            if u == AUTH_USER and bcrypt.checkpw(p, AUTH_HASH):
                session["user"] = u
                session.permanent = True
                nxt = request.args.get("next") or url_for("index")
                if not nxt.startswith("/"):
                    nxt = url_for("index")
                return redirect(nxt)
            return render_template("login.html", error="Invalid username or password.")
        return render_template("login.html")

    @app.route("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.teardown_appcontext
    def close(_e):
        db = g.pop("db", None)
        if db is not None:
            db.close()

    @app.template_filter("money")
    def money(v):
        if v is None:
            return ""
        try:
            v = float(v)
        except (TypeError, ValueError):
            return v
        if v == 0:
            return ""
        return f"${v:,.0f}"

    @app.template_filter("intf")
    def intf(v):
        if v is None or v == 0:
            return ""
        try:
            return f"{int(v):,}"
        except (TypeError, ValueError):
            return v

    # ── list ──────────────────────────────────────────────────────────────
    @app.route("/")
    @login_required
    def index():
        db = get_db()
        q = request.args.get("q", "").strip()
        series = request.args.get("series", "")
        program = request.args.get("program", "")
        fy = request.args.get("fy", "")
        sort = request.args.get("sort", "date")
        order = request.args.get("order", "desc").lower()
        if order not in ("asc", "desc"):
            order = "desc"
        allow = {"date", "name", "series", "program", "fiscal_year",
                 "revenue", "expense", "margin", "tickets_sold"}
        if sort not in allow:
            sort = "date"

        where, params = [], []
        if q:
            where.append("name LIKE ?"); params.append(f"%{q}%")
        if series:
            where.append("series = ?"); params.append(series)
        if program:
            where.append("program = ?"); params.append(program)
        if fy:
            where.append("fiscal_year = ?"); params.append(int(fy))
        where_sql = ("WHERE " + " AND ".join(where)) if where else ""

        sort_expr = {
            "revenue": REV_SUM, "expense": EXP_SUM,
            "margin": f"({REV_SUM}) - ({EXP_SUM})",
        }.get(sort, sort)

        name_filter = "name IS NOT NULL AND name != ''"
        base_where = f"({name_filter})" if not where else f"({name_filter}) AND {' AND '.join(where)}"
        where_sql = "WHERE " + base_where

        rows = db.execute(f"""
            SELECT event_id, date, name, series, program, fiscal_year,
                   ({REV_SUM}) AS revenue,
                   ({EXP_SUM}) AS expense,
                   ({REV_SUM}) - ({EXP_SUM}) AS margin,
                   CASE WHEN full_price_tickets + discount_tickets + comp_tickets > 0
                        THEN full_price_tickets + discount_tickets + comp_tickets
                        ELSE tickets_sold END AS tickets_sold
            FROM event {where_sql}
            ORDER BY {sort_expr} {order.upper()}, date DESC
            LIMIT 1000
        """, params).fetchall()

        totals = db.execute(f"""
            SELECT COUNT(*) AS n,
                   SUM({REV_SUM}) AS rev,
                   SUM({EXP_SUM}) AS exp,
                   SUM({REV_SUM} - ({EXP_SUM})) AS margin
            FROM event {where_sql}
        """, params).fetchone()

        series_list = [r[0] for r in db.execute(
            "SELECT DISTINCT series FROM event WHERE series IS NOT NULL ORDER BY series").fetchall()]
        program_list = [r[0] for r in db.execute(
            "SELECT DISTINCT program FROM event WHERE program IS NOT NULL ORDER BY program").fetchall()]
        fy_list = [r[0] for r in db.execute(
            "SELECT DISTINCT fiscal_year FROM event ORDER BY fiscal_year DESC").fetchall()]

        return render_template("index.html",
            rows=rows, totals=totals,
            q=q, series=series, program=program, fy=fy,
            sort=sort, order=order,
            series_list=series_list, program_list=program_list, fy_list=fy_list)

    # ── new / edit ────────────────────────────────────────────────────────
    @app.route("/event/new", methods=["GET", "POST"])
    @login_required
    def event_new():
        if request.method == "POST":
            return save_event(None)
        ev = empty_event()
        return render_form(ev, [])

    @app.route("/event/<int:event_id>/edit", methods=["GET", "POST"])
    @login_required
    def event_edit(event_id):
        db = get_db()
        if request.method == "POST":
            return save_event(event_id)
        ev = db.execute("SELECT * FROM event WHERE event_id = ?", [event_id]).fetchone()
        if not ev:
            abort(404)
        artists = db.execute("""
            SELECT a.name, ea.role, ea.share_pct
            FROM event_artist ea JOIN artist a USING (artist_id)
            WHERE ea.event_id = ? ORDER BY ea.sort_order
        """, [event_id]).fetchall()
        return render_form(dict(ev), [dict(a) for a in artists], event_id=event_id)

    def empty_event():
        ev = {c: 0 for c, _ in REVENUE_COLS + EXPENSE_COLS + COUNT_COLS}
        ev["tickets_sold"] = 0
        ev["date"] = _dt.date.today().isoformat()
        ev["name"] = ""
        ev["series"] = ""
        ev["program"] = ""
        ev["notes"] = ""
        return ev

    def render_form(ev, artists, event_id=None, error=None):
        db = get_db()
        all_artists = [r[0] for r in db.execute("SELECT name FROM artist ORDER BY name").fetchall()]
        series_list = [r[0] for r in db.execute(
            "SELECT DISTINCT series FROM event WHERE series IS NOT NULL ORDER BY series").fetchall()]
        return render_template("event_form.html",
            ev=ev, artists=artists, event_id=event_id, error=error,
            all_artists=all_artists, series_list=series_list,
            programs=PROGRAMS, roles=ROLES,
            revenue_cols=REVENUE_COLS, expense_cols=EXPENSE_COLS, count_cols=COUNT_COLS)

    def save_event(event_id):
        form = request.form
        date = (form.get("date") or "").strip()
        name = (form.get("name") or "").strip()

        # collect everything so we can re-render on validation failure
        ev = empty_event()
        ev.update({"date": date, "name": name,
                   "series": (form.get("series") or "").strip(),
                   "program": (form.get("program") or "").strip(),
                   "notes":   (form.get("notes") or "").strip()})
        for c, _ in REVENUE_COLS + EXPENSE_COLS:
            ev[c] = parse_float(form.get(c))
        for c, _ in COUNT_COLS:
            ev[c] = parse_int(form.get(c))

        artist_rows = []
        names = form.getlist("artist_name")
        roles = form.getlist("artist_role")
        shares = form.getlist("artist_share")
        for i in range(len(names)):
            n = (names[i] if i < len(names) else "").strip()
            if not n:
                continue
            r = (roles[i] if i < len(roles) else "").strip() or None
            s = parse_float(shares[i] if i < len(shares) else "", default=100.0)
            artist_rows.append({"name": n, "role": r, "share_pct": s})

        # validate
        try:
            _dt.date.fromisoformat(date)
        except ValueError:
            return render_form(ev, artist_rows, event_id, error="Date must be YYYY-MM-DD.")
        if not name:
            return render_form(ev, artist_rows, event_id, error="Name is required.")

        db = get_db()
        # auto-compute tickets_sold from components when a breakdown is provided
        component_sum = sum(ev[c] for c, _ in COUNT_COLS)
        if component_sum > 0:
            ev["tickets_sold"] = component_sum

        cols = ["date", "name", "series", "program", "fiscal_year", "notes",
                "tickets_sold"] + \
               [c for c, _ in REVENUE_COLS + EXPENSE_COLS] + [c for c, _ in COUNT_COLS]
        vals = [ev["date"], ev["name"], ev["series"] or None, ev["program"] or None,
                fiscal_year(ev["date"]), ev["notes"] or None, ev["tickets_sold"]] + \
               [ev[c] for c, _ in REVENUE_COLS + EXPENSE_COLS] + \
               [ev[c] for c, _ in COUNT_COLS]

        try:
            if event_id is None:
                placeholders = ", ".join("?" * len(cols))
                cur = db.execute(
                    f"INSERT INTO event ({', '.join(cols)}) VALUES ({placeholders})", vals)
                event_id = cur.lastrowid
            else:
                sets = ", ".join(f"{c} = ?" for c in cols)
                db.execute(
                    f"UPDATE event SET {sets} WHERE event_id = ?", vals + [event_id])
                db.execute("DELETE FROM event_artist WHERE event_id = ?", [event_id])
        except sqlite3.IntegrityError as e:
            return render_form(ev, artist_rows, event_id,
                               error=f"Could not save: {e}. Likely a duplicate (date, name).")

        for i, ar in enumerate(artist_rows):
            row = db.execute("SELECT artist_id FROM artist WHERE name = ?", [ar["name"]]).fetchone()
            if row:
                artist_id = row["artist_id"]
            else:
                cur = db.execute("INSERT INTO artist (name) VALUES (?)", [ar["name"]])
                artist_id = cur.lastrowid
            db.execute(
                "INSERT INTO event_artist (event_id, artist_id, role, sort_order, share_pct) "
                "VALUES (?, ?, ?, ?, ?)",
                [event_id, artist_id, ar["role"], i, ar["share_pct"]])

        db.commit()
        flash(f"Saved event #{event_id}.", "ok")
        return redirect(url_for("event_detail", event_id=event_id))

    # ── event detail ──────────────────────────────────────────────────────
    @app.route("/event/<int:event_id>")
    @login_required
    def event_detail(event_id):
        db = get_db()
        ev = db.execute("SELECT * FROM event WHERE event_id = ?", [event_id]).fetchone()
        if not ev:
            abort(404)
        artists = db.execute("""
            SELECT a.artist_id, a.name, ea.role, ea.share_pct
            FROM event_artist ea JOIN artist a USING (artist_id)
            WHERE ea.event_id = ? ORDER BY ea.sort_order
        """, [event_id]).fetchall()
        extras = db.execute("""
            SELECT metric, value FROM event_metric WHERE event_id = ?
            ORDER BY metric
        """, [event_id]).fetchall()
        revenue = sum((ev[c] or 0) for c, _ in REVENUE_COLS)
        expense = sum((ev[c] or 0) for c, _ in EXPENSE_COLS)
        return render_template("event.html",
            ev=ev, artists=artists, extras=extras,
            revenue=revenue, expense=expense, margin=revenue - expense,
            revenue_cols=REVENUE_COLS, expense_cols=EXPENSE_COLS, count_cols=COUNT_COLS)

    # ── duplicate ─────────────────────────────────────────────────────────
    @app.route("/event/<int:event_id>/duplicate", methods=["POST"])
    @login_required
    def event_duplicate(event_id):
        db = get_db()
        src = db.execute("SELECT * FROM event WHERE event_id = ?", [event_id]).fetchone()
        if not src:
            abort(404)
        # exclude event_id and source_file; clearing source_file avoids the
        # UNIQUE(date, name, source_file) constraint on the duplicate row.
        cols = [c for c in src.keys() if c not in ("event_id", "source_file", "sheet")]
        placeholders = ", ".join("?" * len(cols))
        cur = db.execute(
            f"INSERT INTO event ({', '.join(cols)}) VALUES ({placeholders})",
            [src[c] for c in cols])
        new_id = cur.lastrowid
        artists = db.execute(
            "SELECT artist_id, role, sort_order, share_pct FROM event_artist WHERE event_id = ?",
            [event_id]).fetchall()
        for a in artists:
            db.execute(
                "INSERT INTO event_artist (event_id, artist_id, role, sort_order, share_pct) "
                "VALUES (?, ?, ?, ?, ?)",
                [new_id, a["artist_id"], a["role"], a["sort_order"], a["share_pct"]])
        db.commit()
        flash("Duplicated event — adjust the details for this show.", "ok")
        return redirect(url_for("event_edit", event_id=new_id))

    # ── delete ────────────────────────────────────────────────────────────
    @app.route("/event/<int:event_id>/delete", methods=["POST"])
    @login_required
    def event_delete(event_id):
        db = get_db()
        ev = db.execute("SELECT name, date FROM event WHERE event_id = ?", [event_id]).fetchone()
        if not ev:
            abort(404)
        db.execute("DELETE FROM event_artist WHERE event_id = ?", [event_id])
        db.execute("DELETE FROM event_metric  WHERE event_id = ?", [event_id])
        db.execute("DELETE FROM event          WHERE event_id = ?", [event_id])
        db.commit()
        flash(f"Deleted event: {ev['name']} ({ev['date']}).", "ok")
        return redirect(url_for("index"))

    # ── artists ───────────────────────────────────────────────────────────
    @app.route("/artists")
    @login_required
    def artists():
        db = get_db()
        rows = db.execute(f"""
            SELECT a.artist_id, a.name,
                   COUNT(DISTINCT e.event_id) AS shows,
                   SUM(({REV_SUM}) * COALESCE(ea.share_pct, 100) / 100.0) AS revenue
            FROM artist a
            LEFT JOIN event_artist ea ON ea.artist_id = a.artist_id
            LEFT JOIN event e ON e.event_id = ea.event_id
            GROUP BY a.artist_id
            ORDER BY revenue DESC NULLS LAST, a.name
        """).fetchall()
        return render_template("artists.html", rows=rows)

    @app.route("/artist/<int:artist_id>")
    @login_required
    def artist_detail(artist_id):
        db = get_db()
        a = db.execute("SELECT * FROM artist WHERE artist_id = ?", [artist_id]).fetchone()
        if not a:
            abort(404)
        rows = db.execute(f"""
            SELECT e.event_id, e.date, e.name, e.series, e.fiscal_year,
                   ea.share_pct,
                   ({REV_SUM}) * ea.share_pct / 100.0 AS revenue,
                   ({EXP_SUM}) * ea.share_pct / 100.0 AS expense,
                   (({REV_SUM}) - ({EXP_SUM})) * ea.share_pct / 100.0 AS margin,
                   e.tickets_sold
            FROM event_artist ea JOIN event e USING (event_id)
            WHERE ea.artist_id = ?
            ORDER BY e.date DESC
        """, [artist_id]).fetchall()
        totals = db.execute(f"""
            SELECT COUNT(*) AS n,
                   SUM(({REV_SUM}) * ea.share_pct / 100.0) AS rev,
                   SUM(({EXP_SUM}) * ea.share_pct / 100.0) AS exp,
                   SUM((({REV_SUM}) - ({EXP_SUM})) * ea.share_pct / 100.0) AS margin
            FROM event_artist ea JOIN event e USING (event_id)
            WHERE ea.artist_id = ?
        """, [artist_id]).fetchone()
        return render_template("artist.html", artist=a, rows=rows, totals=totals)

    @app.route("/artist/<int:artist_id>/delete", methods=["POST"])
    @login_required
    def artist_delete(artist_id):
        db = get_db()
        a = db.execute("SELECT name FROM artist WHERE artist_id = ?", [artist_id]).fetchone()
        if not a:
            abort(404)
        db.execute("DELETE FROM event_artist WHERE artist_id = ?", [artist_id])
        db.execute("DELETE FROM artist WHERE artist_id = ?", [artist_id])
        db.commit()
        flash(f"Deleted artist: {a['name']}.", "ok")
        return redirect(url_for("artists"))

    return app


def main():
    if "--help" in sys.argv or "-h" in sys.argv:
        print(__doc__)
        sys.exit(0)

    port = 5050
    args = sys.argv[1:]
    i = 0
    global DB_PATH
    while i < len(args):
        if args[i] == "--port" and i + 1 < len(args):
            port = int(args[i + 1]); i += 2
        elif args[i] == "--db" and i + 1 < len(args):
            DB_PATH = Path(args[i + 1]); i += 2
        else:
            i += 1

    if not DB_PATH.exists():
        print(f"Database not found: {DB_PATH}", file=sys.stderr)
        print("Run: python build_event_db.py", file=sys.stderr)
        sys.exit(1)

    app = create_app()
    print(f"Serving {DB_PATH} on http://127.0.0.1:{port}/  (Ctrl-C to stop)")
    app.run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
