"""
Lease Accounting Reconciliation & Data Quality Checker
-------------------------------------------------------
Compares SOURCE data (lease system) with TARGET data (general-ledger postings)
using SQL, and reports data-quality defects across six dimensions:
completeness, accuracy, uniqueness, validity, consistency, timeliness,
plus a per-lease reconciliation.

All data is synthetic (generated below). Defects are seeded on purpose so every
check has something to catch. Uses only Python standard library + openpyxl.
"""
import random
import sqlite3
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font

random.seed(42)
con = sqlite3.connect(":memory:")
cur = con.cursor()

# ---------------------------------------------------------------- 1. SCHEMA
cur.executescript("""
CREATE TABLE src_leases (
    lease_id     INTEGER PRIMARY KEY,
    tenant       TEXT,
    monthly_rent REAL,
    start_date   TEXT,
    end_date     TEXT,
    status       TEXT
);
CREATE TABLE src_charges (
    charge_id   INTEGER PRIMARY KEY,
    lease_id    INTEGER,
    charge_date TEXT,
    amount      REAL
);
CREATE TABLE tgt_gl_postings (
    posting_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    charge_id    INTEGER,
    lease_id     INTEGER,
    posting_date TEXT,
    amount       REAL
);
""")

# ------------------------------------------------- 2. SYNTHETIC SOURCE DATA
for lease_id in range(1, 21):                       # 20 leases
    rent = random.choice([1200, 1500, 1800, 2200, 2500, 3100])
    cur.execute("INSERT INTO src_leases VALUES (?,?,?,?,?,?)",
                (lease_id, f"Tenant {lease_id:02d}", rent,
                 "2026-01-01", "2026-12-31", "ACTIVE"))

charge_id = 1
for lease_id, rent in cur.execute("SELECT lease_id, monthly_rent FROM src_leases").fetchall():
    for month in range(1, 7):                       # 6 monthly charges each
        cur.execute("INSERT INTO src_charges VALUES (?,?,?,?)",
                    (charge_id, lease_id, f"2026-{month:02d}-01", rent))
        charge_id += 1

# ----------------------------- 3. TARGET = copy of source, then seed defects
all_charges = cur.execute(
    "SELECT charge_id, lease_id, charge_date, amount FROM src_charges").fetchall()
ids = [c[0] for c in all_charges]
picked = random.sample(ids, 14)
missing, wrong_amt, duplicated, late = picked[:5], picked[5:9], picked[9:12], picked[12:14]

for cid, lid, cdate, amt in all_charges:
    if cid in missing:
        continue                                    # completeness defect
    post_date = date.fromisoformat(cdate) + timedelta(days=1)
    post_amt = amt + 150 if cid in wrong_amt else amt     # accuracy defect
    if cid in late:
        post_date = date.fromisoformat(cdate) + timedelta(days=9)  # timeliness defect
    cur.execute("INSERT INTO tgt_gl_postings (charge_id, lease_id, posting_date, amount) "
                "VALUES (?,?,?,?)", (cid, lid, post_date.isoformat(), post_amt))
    if cid in duplicated:                           # uniqueness defect
        cur.execute("INSERT INTO tgt_gl_postings (charge_id, lease_id, posting_date, amount) "
                    "VALUES (?,?,?,?)", (cid, lid, post_date.isoformat(), post_amt))

# consistency defects: ledger postings with no matching source charge
cur.execute("INSERT INTO tgt_gl_postings (charge_id, lease_id, posting_date, amount) "
            "VALUES (9001, 999, '2026-03-02', 2000)")
cur.execute("INSERT INTO tgt_gl_postings (charge_id, lease_id, posting_date, amount) "
            "VALUES (9002, 998, '2026-04-02', 1750)")

# validity defects in the source leases
cur.execute("UPDATE src_leases SET end_date = '2025-06-30' WHERE lease_id = 4")   # end < start
cur.execute("UPDATE src_leases SET monthly_rent = NULL    WHERE lease_id = 9")    # missing rent
cur.execute("UPDATE src_leases SET monthly_rent = -500    WHERE lease_id = 13")   # negative rent
con.commit()

# ----------------------------------------------------------- 4. SQL CHECKS
CHECKS = [
    ("C1 Completeness",
     "Source charges with NO posting in the ledger",
     """SELECT c.charge_id, c.lease_id, c.charge_date, c.amount
        FROM src_charges c
        LEFT JOIN tgt_gl_postings p ON p.charge_id = c.charge_id
        WHERE p.posting_id IS NULL
        ORDER BY c.charge_id"""),

    ("C2 Accuracy",
     "Posted amount differs from source amount",
     """SELECT c.charge_id, c.lease_id, c.amount AS src_amount,
               p.amount AS tgt_amount, ROUND(p.amount - c.amount, 2) AS difference
        FROM src_charges c
        JOIN tgt_gl_postings p ON p.charge_id = c.charge_id
        WHERE ABS(p.amount - c.amount) > 0.005
        ORDER BY c.charge_id"""),

    ("C3 Uniqueness",
     "Same charge posted more than once (double posting)",
     """SELECT charge_id, COUNT(*) AS times_posted, SUM(amount) AS total_posted
        FROM tgt_gl_postings
        GROUP BY charge_id
        HAVING COUNT(*) > 1
        ORDER BY charge_id"""),

    ("C4 Validity",
     "Leases with missing/non-positive rent or end date before start date",
     """SELECT lease_id, tenant, monthly_rent, start_date, end_date,
               CASE WHEN monthly_rent IS NULL THEN 'rent is NULL'
                    WHEN monthly_rent <= 0    THEN 'rent not positive'
                    WHEN end_date < start_date THEN 'end date before start date'
               END AS issue
        FROM src_leases
        WHERE monthly_rent IS NULL OR monthly_rent <= 0 OR end_date < start_date
        ORDER BY lease_id"""),

    ("C5 Consistency",
     "Ledger postings that have no matching source charge (orphans)",
     """SELECT p.posting_id, p.charge_id, p.lease_id, p.posting_date, p.amount
        FROM tgt_gl_postings p
        LEFT JOIN src_charges c ON c.charge_id = p.charge_id
        WHERE c.charge_id IS NULL
        ORDER BY p.posting_id"""),

    ("C6 Timeliness",
     "Posted more than 5 days after the charge date",
     """SELECT c.charge_id, c.lease_id, c.charge_date, p.posting_date,
               CAST(julianday(p.posting_date) - julianday(c.charge_date) AS INTEGER) AS days_late
        FROM src_charges c
        JOIN tgt_gl_postings p ON p.charge_id = c.charge_id
        WHERE julianday(p.posting_date) - julianday(c.charge_date) > 5
        ORDER BY days_late DESC"""),

    ("C7 Reconciliation",
     "Per-lease total: source vs ledger (variance != 0)",
     """SELECT s.lease_id, s.src_total, COALESCE(t.tgt_total, 0) AS tgt_total,
               ROUND(COALESCE(t.tgt_total, 0) - s.src_total, 2) AS variance
        FROM (SELECT lease_id, SUM(amount) AS src_total
              FROM src_charges GROUP BY lease_id) s
        LEFT JOIN (SELECT lease_id, SUM(amount) AS tgt_total
                   FROM tgt_gl_postings GROUP BY lease_id) t
               ON t.lease_id = s.lease_id
        WHERE ABS(COALESCE(t.tgt_total, 0) - s.src_total) > 0.005
        ORDER BY s.lease_id"""),
]

# ------------------------------------------------------- 5. RUN + REPORT
wb = Workbook()
summary = wb.active
summary.title = "Summary"
summary.append(["Check", "Description", "Defects found", "Status"])
for cell in summary[1]:
    cell.font = Font(bold=True)

print(f"{'CHECK':<20}{'DEFECTS':>8}  STATUS   DESCRIPTION")
print("-" * 90)
for name, desc, sql in CHECKS:
    rows = cur.execute(sql).fetchall()
    headers = [d[0] for d in cur.description]
    status = "PASS" if not rows else "FAIL"
    print(f"{name:<20}{len(rows):>8}  {status:<8} {desc}")
    summary.append([name, desc, len(rows), status])
    if rows:
        ws = wb.create_sheet(name.split()[0])
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for r in rows:
            ws.append(list(r))
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

for col, w in zip("ABCD", (20, 70, 15, 10)):
    summary.column_dimensions[col].width = w

wb.save("lease_recon_report.xlsx")
print("\nExcel report saved: lease_recon_report.xlsx")
