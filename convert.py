#!/usr/bin/env python3
"""
JobOps — Job Management & Operations System (local prototype)
convert.py : data utility for the prototype

Commands
  python3 convert.py serve [--port 8000]          Serve index.html + data.json locally (required: browsers block fetch() on file://)
  python3 convert.py demo  [--out data.json]      Re-generate the clearly-labelled DEMO dataset
  python3 convert.py empty [--out data.json]      Create a clean dataset (settings, roles, one admin user, no business records)
  python3 convert.py validate [data.json]         Referential integrity, duplicates, ranges + recomputed job financials
  python3 convert.py test                         Unit tests for the financial formulas (acceptance tests 4, 5, 13)
  python3 convert.py template [--out jobops_import_template.xlsx]   Excel template (one sheet per table)
  python3 convert.py from-excel <file.xlsx> [--out data.json] [--base data.json]
                                                   Import master data / jobs / cost lines from the Excel template
  python3 convert.py to-sql [data.json] [--out supabase_seed.sql]
                                                   PostgreSQL schema + seed inserts for the Supabase phase

Only the Python standard library is needed, except `template` / `from-excel` (openpyxl: pip install openpyxl).
"""
import argparse
import datetime as dt
import http.server
import json
import os
import random
import socketserver
import sys
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
SCHEMA_VERSION = 1

# ---------------------------------------------------------------------------
# Table specification — single source for template, import, validation, SQL.
# type: text | num | int | date | ts | bool | arr | json ; fk: referenced table
# ---------------------------------------------------------------------------
C = lambda name, typ="text", fk=None, req=False: {"name": name, "type": typ, "fk": fk, "req": req}

TABLES = {
    "customers": [C("id", req=True), C("code", req=True), C("name", req=True), C("contactPerson"), C("email"), C("phone"),
                  C("address"), C("trn"), C("paymentTerms"), C("creditLimit", "num"), C("zohoCustomerId"), C("notes"), C("status")],
    "endClients": [C("id", req=True), C("code", req=True), C("name", req=True), C("contactPerson"), C("email"), C("phone"),
                   C("address"), C("customerIds", "arr"), C("notes"), C("status")],
    "suppliers": [C("id", req=True), C("code", req=True), C("name", req=True), C("contactPerson"), C("email"), C("phone"),
                  C("address"), C("trn"), C("paymentTerms"), C("category"), C("notes"), C("status")],
    "employees": [C("id", req=True), C("code", req=True), C("name", req=True), C("type", req=True), C("phone"), C("email"),
                  C("title"), C("department"), C("joiningDate", "date"), C("status"), C("notes"), C("employmentType"),
                  C("stdHoursPerDay", "num"), C("costRate", "num"), C("skills"), C("supplierId", fk="suppliers"),
                  C("rateBasis"), C("agreedRate", "num"), C("capacityPct", "num"), C("contractStart", "date"), C("contractEnd", "date")],
    "leaves": [C("id", req=True), C("employeeId", fk="employees", req=True), C("fromDate", "date", req=True),
               C("toDate", "date", req=True), C("type"), C("notes")],
    "inventoryItems": [C("id", req=True), C("sku", req=True), C("name", req=True), C("uom"), C("category"), C("store"),
                       C("qtyOnHand", "num"), C("avgCost", "num"), C("status")],
    "stockReceipts": [C("id", req=True), C("itemId", fk="inventoryItems", req=True), C("receiptDate", "date"),
                      C("qty", "num"), C("unitCost", "num"), C("ref")],
    "jobs": [C("id", req=True), C("jobNo", req=True), C("type", req=True), C("openingDate", "date", req=True),
             C("customerId", fk="customers"), C("endClientId", fk="endClients"), C("zohoSo"), C("zohoSoId"), C("description"),
             C("status", req=True), C("completionPct", "num"), C("lpoDate", "date"), C("customerOrderNo"), C("poValue", "num"),
             C("actualPoValue", "num"), C("psDate", "date"), C("pfDate", "date"), C("actualStart", "date"),
             C("actualFinish", "date"), C("projectEngineerId", fk="employees"), C("serviceEngineerId", fk="employees"),
             C("comments"), C("locked", "bool"), C("closureDate", "date"), C("closureComments"), C("closedBy", fk="users"),
             C("closedAt", "ts")],
    "plannedManpower": [C("id", req=True), C("jobId", fk="jobs", req=True), C("employeeId", fk="employees"), C("resourceName"),
                        C("resourceType"), C("supplierId", fk="suppliers"), C("headcount", "num"), C("hoursPerPerson", "num"),
                        C("rate", "num"), C("startDate", "date"), C("endDate", "date"), C("comments"),
                        C("assignmentId", fk="assignments")],
    "actualManpower": [C("id", req=True), C("jobId", fk="jobs", req=True), C("employeeId", fk="employees"), C("resourceName"),
                       C("resourceType"), C("supplierId", fk="suppliers"), C("headcount", "num"), C("hoursPerPerson", "num"),
                       C("rate", "num"), C("dateFrom", "date"), C("dateTo", "date"), C("comments"),
                       C("assignmentId", fk="assignments")],
    "assignments": [C("id", req=True), C("employeeId", fk="employees", req=True), C("jobId", fk="jobs", req=True),
                    C("startDate", "date", req=True), C("endDate", "date", req=True), C("plannedHours", "num"),
                    C("allocationPct", "num"), C("status"), C("comments"), C("overrideBy", fk="users"), C("overrideReason")],
    "plannedMaterials": [C("id", req=True), C("jobId", fk="jobs", req=True), C("description"), C("category"),
                         C("supplierId", fk="suppliers"), C("qty", "num"), C("uom"), C("unitPrice", "num"),
                         C("expectedDate", "date"), C("comments")],
    "actualMaterials": [C("id", req=True), C("jobId", fk="jobs", req=True), C("description"), C("category"),
                        C("supplierId", fk="suppliers"), C("qty", "num"), C("uom"), C("unitPrice", "num"),
                        C("purchaseDate", "date"), C("supplierRef"), C("comments")],
    "inventoryTransfers": [C("id", req=True), C("jobId", fk="jobs", req=True), C("itemId", fk="inventoryItems", req=True),
                           C("qty", "num"), C("uom"), C("transferDate", "date"), C("issuedBy"), C("receivedBy"), C("store"),
                           C("costingMethod"), C("unitCost", "num"), C("costSource"), C("ref", req=True), C("status"),
                           C("comments")],
    "plannedMisc": [C("id", req=True), C("jobId", fk="jobs", req=True), C("category"), C("description"),
                    C("supplierId", fk="suppliers"), C("payee"), C("expenseDate", "date"), C("amount", "num"), C("comments")],
    "actualMisc": [C("id", req=True), C("jobId", fk="jobs", req=True), C("category"), C("description"),
                   C("supplierId", fk="suppliers"), C("payee"), C("expenseDate", "date"), C("amount", "num"),
                   C("paymentMode"), C("paymentRef"), C("attachmentName"), C("comments")],
    "miscPayments": [C("id", req=True), C("miscId", fk="actualMisc", req=True), C("payDate", "date"), C("amount", "num"),
                     C("mode"), C("ref")],
    "invoices": [C("id", req=True), C("jobId", fk="jobs", req=True), C("invoiceNo", req=True), C("invoiceDate", "date"),
                 C("dueDate", "date"), C("amount", "num"), C("status"), C("zohoInvoiceId"), C("notes")],
    "invoicePayments": [C("id", req=True), C("invoiceId", fk="invoices", req=True), C("payDate", "date"), C("amount", "num"),
                        C("mode"), C("ref"), C("zohoPaymentId")],
    "jobStatusHistory": [C("id", req=True), C("jobId", fk="jobs", req=True), C("fromStatus"), C("toStatus"), C("at", "ts"),
                         C("userId", fk="users"), C("note")],
    "poRevisions": [C("id", req=True), C("jobId", fk="jobs", req=True), C("field"), C("fromValue", "num"),
                    C("toValue", "num"), C("at", "ts"), C("userId", fk="users"), C("reason")],
    "jobComments": [C("id", req=True), C("jobId", fk="jobs", req=True), C("userId", fk="users"), C("at", "ts"), C("text")],
    "jobAttachments": [C("id", req=True), C("jobId", fk="jobs", req=True), C("name"), C("size", "int"), C("mime"),
                       C("at", "ts"), C("userId", fk="users"), C("note")],
    "closures": [C("id", req=True), C("jobId", fk="jobs", req=True), C("closureDate", "date"), C("comments"),
                 C("finalStatus"), C("financialsConfirmed", "bool"), C("pendingConfirmed", "bool"),
                 C("closedBy", fk="users"), C("closedAt", "ts"), C("reopenedBy", fk="users"), C("reopenedAt", "ts"),
                 C("reopenReason")],
    "roles": [C("id", req=True), C("name", req=True), C("system", "bool"), C("jobTypes", "arr"), C("permissions", "json")],
    "users": [C("id", req=True), C("name", req=True), C("email", req=True), C("roleId", fk="roles", req=True),
              C("department"), C("active", "bool")],
    "auditLog": [C("id", req=True), C("at", "ts"), C("userId", fk="users"), C("action"), C("entity"), C("recordId"),
                 C("before", "json"), C("after", "json"), C("reason")],
}
# Tables that carry createdAt/createdBy/updatedAt/version/archived
BUSINESS = [t for t in TABLES if t not in ("auditLog", "jobStatusHistory", "poRevisions", "roles")]
IMPORTABLE = ["customers", "endClients", "suppliers", "employees", "leaves", "inventoryItems", "jobs",
              "plannedManpower", "actualManpower", "plannedMaterials", "actualMaterials", "plannedMisc", "actualMisc",
              "invoices", "invoicePayments"]

PERMISSIONS = [
    "viewOverview", "viewJobs", "viewRoster", "viewMaster", "viewReports", "viewSettings",
    "jobCreate", "jobEdit", "jobDelete", "viewFinancials", "editFinancials", "viewCostRates",
    "manageManpower", "manageMaterials", "manageMisc", "changeInvoiceStatus", "recordPayments",
    "closeJobs", "reopenJobs", "editClosedJobs", "exportReports", "manageMaster", "rosterAssign",
    "overrideConflicts", "manageUsers", "manageSettings",
]
JOB_TYPES = ["Service", "Supply", "Supply/Services"]


def role(rid, name, perms, types=None):
    p = {k: (k in perms) for k in PERMISSIONS}
    return {"id": rid, "name": name, "system": True, "jobTypes": types or list(JOB_TYPES), "permissions": p}


ALL = set(PERMISSIONS)
VIEW_TABS = {"viewOverview", "viewJobs", "viewRoster", "viewMaster", "viewReports"}
DEFAULT_ROLES = [
    role("role_super", "Super Admin", ALL),
    role("role_admin", "Administrator", ALL - {"manageUsers"} | {"manageUsers"}),
    role("role_mgmt", "Management", VIEW_TABS | {"viewFinancials", "viewCostRates", "exportReports"}),
    role("role_fin", "Finance / Accounts", VIEW_TABS | {"viewFinancials", "editFinancials", "viewCostRates", "manageMisc",
                                                       "changeInvoiceStatus", "recordPayments", "closeJobs",
                                                       "exportReports"}),
    role("role_pm", "Project Manager", VIEW_TABS | {"jobCreate", "jobEdit", "viewFinancials", "editFinancials",
                                                    "manageManpower", "manageMaterials", "manageMisc", "rosterAssign",
                                                    "exportReports"}),
    role("role_ops", "Operations", VIEW_TABS - {"viewReports"} | {"jobEdit", "manageMaterials", "manageMisc"}),
    role("role_hr", "HR / Resource Planner", {"viewOverview", "viewJobs", "viewRoster", "viewMaster", "rosterAssign",
                                              "manageManpower", "manageMaster", "overrideConflicts"}),
    role("role_ro", "Read-Only User", {"viewOverview", "viewJobs", "viewRoster", "viewReports"}),
]

DEFAULT_SETTINGS = {
    "companyName": "Your Company LLC",
    "logoText": "JO",
    "currency": "AED",
    "precision": 2,
    "locale": "en-AE",
    "jobNumber": {"prefix": "JOB", "pattern": "{PREFIX}-{TYPE}-{YYYY}-{SEQ}", "padding": 4, "nextSeq": 1,
                  "typeCodes": {"Service": "SRV", "Supply": "SUP", "Supply/Services": "SSV"}},
    "jobStatuses": ["Open", "Planned", "In Progress", "On Hold", "Completed", "Closed", "Cancelled"],
    "expenseCategories": ["Transportation", "Accommodation", "Permits & Gate Passes", "Equipment Rental", "Delivery",
                          "Parking & Tolls", "Meals", "Other"],
    "materialCategories": ["Consumables", "Spare Parts", "Electrical", "Mechanical", "Piping & Valves", "Safety / PPE",
                           "Tools", "Other"],
    "paymentModes": ["Cash", "Bank Transfer", "Company Card", "Cheque", "Other"],
    "uoms": ["EA", "PCS", "SET", "M", "KG", "L", "BOX", "LOT"],
    "departments": ["Operations", "Projects", "Workshop", "Logistics", "Finance", "HR", "Management"],
    "workingDays": [1, 2, 3, 4, 5],
    "hoursPerDay": 8,
    "timezone": "Asia/Dubai",
    "fiscalYearStartMonth": 1,
    "revenueRule": "po_value",
    "inventoryCosting": "weighted_average",
    "suggest100OnCompleted": True,
    "gpWarnPct": 15,
    "attachmentLimitMb": 10,
    "notifyOnClose": True,
    "notifyOnOverbook": True,
    "exportRequiresPermission": True,
}

# ---------------------------------------------------------------------------
# Financial calculation policy (mirrors index.html → Calc)
# Each line total is rounded to the currency minor unit (2 dp, half-up), totals are sums of rounded lines.
# ---------------------------------------------------------------------------
TWO = Decimal("0.01")


def D(x):
    if x is None or x == "":
        return Decimal("0")
    return Decimal(str(x))


def r2(x):
    return D(x).quantize(TWO, rounding=ROUND_HALF_UP)


def line_mp(l):
    return r2(D(l.get("headcount")) * D(l.get("hoursPerPerson")) * D(l.get("rate")))


def line_qty(l, price="unitPrice"):
    return r2(D(l.get("qty")) * D(l.get(price)))


def job_financials(db, job):
    jid = job["id"]
    by = lambda t: [x for x in db.get(t, []) if x.get("jobId") == jid and not x.get("archived")]
    pm = sum((line_mp(l) for l in by("plannedManpower")), Decimal(0))
    am = sum((line_mp(l) for l in by("actualManpower")), Decimal(0))
    pmat = sum((line_qty(l) for l in by("plannedMaterials")), Decimal(0))
    amat_p = sum((line_qty(l) for l in by("actualMaterials")), Decimal(0))
    amat_t = sum((line_qty(l, "unitCost") for l in by("inventoryTransfers") if l.get("status", "Posted") == "Posted"),
                 Decimal(0))
    pmisc = sum((r2(l.get("amount")) for l in by("plannedMisc")), Decimal(0))
    amisc = sum((r2(l.get("amount")) for l in by("actualMisc")), Decimal(0))
    rule = db.get("settings", {}).get("revenueRule", "po_value")
    ppo = r2(job.get("poValue"))
    if rule == "invoiced":
        apo = sum((r2(i.get("amount")) for i in by("invoices")), Decimal(0))
    else:
        apo = r2(job["actualPoValue"]) if job.get("actualPoValue") not in (None, "") else ppo
    ptot = pm + pmat + pmisc
    atot = am + amat_p + amat_t + amisc
    pgp, agp = ppo - ptot, apo - atot
    pct = lambda gp, po: (gp / po * 100).quantize(TWO, rounding=ROUND_HALF_UP) if po > 0 else None
    return {"plannedPO": ppo, "actualPO": apo, "plannedManpower": pm, "actualManpower": am,
            "plannedMaterial": pmat, "actualMaterialPurchased": amat_p, "actualMaterialTransfer": amat_t,
            "actualMaterial": amat_p + amat_t, "plannedMisc": pmisc, "actualMisc": amisc,
            "plannedCost": ptot, "actualCost": atot, "plannedGP": pgp, "actualGP": agp,
            "plannedGPPct": pct(pgp, ppo), "actualGPPct": pct(agp, apo), "costVariance": atot - ptot,
            "gpVariance": agp - pgp}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def empty_db(demo=False):
    db = {"meta": {"schemaVersion": SCHEMA_VERSION, "generatedAt": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                   "demo": demo,
                   "note": ("DEMO DATA — fictional companies, people and figures for testing only." if demo
                            else "Clean dataset.")},
          "settings": json.loads(json.dumps(DEFAULT_SETTINGS)),
          "roles": json.loads(json.dumps(DEFAULT_ROLES))}
    for t in TABLES:
        if t not in ("roles",):
            db.setdefault(t, [])
    return db


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(db, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=1)
    print(f"wrote {path}  ({os.path.getsize(path):,} bytes)")


def stamp(rec, at, by="usr_super"):
    rec.setdefault("createdAt", at)
    rec.setdefault("createdBy", by)
    rec.setdefault("updatedAt", at)
    rec.setdefault("version", 1)
    rec.setdefault("archived", False)
    return rec


def job_no(settings, jtype, year, seq):
    jn = settings["jobNumber"]
    return (jn["pattern"].replace("{PREFIX}", jn["prefix"]).replace("{TYPE}", jn["typeCodes"][jtype])
            .replace("{YYYY}", str(year)).replace("{YY}", str(year)[2:]).replace("{SEQ}", str(seq).zfill(jn["padding"])))


# ---------------------------------------------------------------------------
# Demo data
# ---------------------------------------------------------------------------
def build_demo(today=None):
    rnd = random.Random(2026)
    today = today or dt.date(2026, 10, 9)
    db = empty_db(demo=True)
    S = db["settings"]
    iso = lambda d: d.isoformat() if d else None
    ts = lambda d: dt.datetime.combine(d, dt.time(9, 0)).isoformat() + "+04:00"

    db["users"] = [
        {"id": "usr_super", "name": "Sara Admin", "email": "superadmin@example.com", "roleId": "role_super", "department": "Management", "active": True},
        {"id": "usr_admin", "name": "Omar Haddad", "email": "admin@example.com", "roleId": "role_admin", "department": "Management", "active": True},
        {"id": "usr_mgmt", "name": "Layla Mansour", "email": "gm@example.com", "roleId": "role_mgmt", "department": "Management", "active": True},
        {"id": "usr_fin", "name": "Priya Nair", "email": "accounts@example.com", "roleId": "role_fin", "department": "Finance", "active": True},
        {"id": "usr_pm", "name": "Daniel Fernandes", "email": "pm@example.com", "roleId": "role_pm", "department": "Projects", "active": True},
        {"id": "usr_ops", "name": "Imran Qureshi", "email": "ops@example.com", "roleId": "role_ops", "department": "Operations", "active": True},
        {"id": "usr_hr", "name": "Mariam Saeed", "email": "hr@example.com", "roleId": "role_hr", "department": "HR", "active": True},
        {"id": "usr_ro", "name": "Guest Viewer", "email": "viewer@example.com", "roleId": "role_ro", "department": "Operations", "active": True},
    ]
    seed_at = ts(dt.date(2026, 1, 1))
    for u in db["users"]:
        stamp(u, seed_at)

    custs = [("Al Noor Industrial LLC", "Net 30"), ("Marina Facilities Management", "Net 45"),
             ("Desert Peak Contracting", "Net 60"), ("Gulfline Process Services", "Net 30"),
             ("Creekside Power Solutions", "Net 45"), ("Horizon Marine Works", "Net 30"),
             ("Sahara Utilities Maintenance", "Net 60"), ("Palm Coast Logistics", "Net 30")]
    for i, (n, terms) in enumerate(custs, 1):
        db["customers"].append(stamp({"id": f"cus_{i:03d}", "code": f"C-{100 + i}", "name": n,
                                      "contactPerson": rnd.choice(["Ahmed K.", "Rania S.", "John M.", "Vikram P.", "Fatima A."]),
                                      "email": f"procurement{i}@example.com", "phone": f"+971 4 {rnd.randint(200, 899)} {rnd.randint(1000, 9999)}",
                                      "address": rnd.choice(["Jebel Ali Free Zone, Dubai", "Al Quoz Industrial 3, Dubai",
                                                             "Mussafah M-12, Abu Dhabi", "Hamriyah Free Zone, Sharjah", "DIP 2, Dubai"]),
                                      "trn": f"100{rnd.randint(100000000, 999999999)}00003", "paymentTerms": terms,
                                      "creditLimit": rnd.choice([250000, 500000, 750000, 1000000]), "zohoCustomerId": None,
                                      "notes": "", "status": "Active" if i != 8 else "Inactive"}, seed_at))
    ecs = ["Northgate Refinery (demo)", "Bluewater Desalination Plant (demo)", "Oasis District Cooling (demo)",
           "Falcon Aluminium Smelter (demo)", "Crescent Gas Processing (demo)", "Seabreeze Power Station (demo)",
           "Mirage Mall Facilities (demo)", "Harbour Container Terminal (demo)"]
    for i, n in enumerate(ecs, 1):
        db["endClients"].append(stamp({"id": f"ecl_{i:03d}", "code": f"E-{200 + i}", "name": n,
                                       "contactPerson": rnd.choice(["Site Engineer", "Maintenance Manager", "Plant Superintendent"]),
                                       "email": f"site{i}@example.com", "phone": f"+971 2 {rnd.randint(200, 899)} {rnd.randint(1000, 9999)}",
                                       "address": rnd.choice(["Ruwais", "Jebel Ali", "Hamriyah", "Khalifa Port", "Taweelah"]),
                                       "customerIds": sorted(rnd.sample([c["id"] for c in db["customers"]], 2)),
                                       "notes": "", "status": "Active"}, seed_at))
    sups = [("Emirates Industrial Supplies", "Materials"), ("Al Barsha Fasteners Trading", "Materials"),
            ("ProValve Middle East", "Materials"), ("SkillForce Manpower Services", "Manpower"),
            ("Gulf Technical Staffing", "Manpower"), ("Rapid Crane & Equipment Rental", "Equipment"),
            ("Swift Freight & Delivery", "Logistics"), ("Safeguard PPE Trading", "Materials"),
            ("Metro Cable & Electrical", "Materials"), ("Oasis Hotel Apartments", "Accommodation")]
    for i, (n, cat) in enumerate(sups, 1):
        db["suppliers"].append(stamp({"id": f"sup_{i:03d}", "code": f"S-{300 + i}", "name": n,
                                      "contactPerson": rnd.choice(["Sales Desk", "Account Manager", "Operations"]),
                                      "email": f"sales{i}@example.com", "phone": f"+971 6 {rnd.randint(200, 899)} {rnd.randint(1000, 9999)}",
                                      "address": "UAE", "trn": f"100{rnd.randint(100000000, 999999999)}00003",
                                      "paymentTerms": rnd.choice(["Net 30", "Net 45", "Advance"]), "category": cat,
                                      "notes": "", "status": "Active"}, seed_at))

    inhouse = [("Rashid Al Hammadi", "Project Engineer", "Projects", 95), ("Arjun Menon", "Project Engineer", "Projects", 90),
               ("Kevin D'Souza", "Service Engineer", "Operations", 75), ("Hassan Raza", "Service Engineer", "Operations", 72),
               ("Nikhil Sharma", "Mechanical Technician", "Workshop", 48), ("Ramon Cruz", "Mechanical Technician", "Workshop", 45),
               ("Tariq Mahmood", "Electrical Technician", "Operations", 50), ("Joseph Thomas", "Instrument Technician", "Operations", 52),
               ("Ali Akbar", "Welder / Fabricator", "Workshop", 42), ("Suresh Kumar", "Rigger", "Operations", 35),
               ("Bilal Khan", "Helper", "Workshop", 25), ("Michael Santos", "Storekeeper", "Logistics", 30),
               ("Yousef Darwish", "QA/QC Inspector", "Projects", 70), ("Anil Varghese", "HSE Officer", "Operations", 60)]
    for i, (n, title, dep, rate) in enumerate(inhouse, 1):
        db["employees"].append(stamp({"id": f"emp_{i:03d}", "code": f"EMP-{i:03d}", "name": n, "type": "inhouse",
                                      "phone": f"+971 5{rnd.randint(0, 8)} {rnd.randint(100, 999)} {rnd.randint(1000, 9999)}",
                                      "email": f"{n.split()[0].lower()}@example.com", "title": title, "department": dep,
                                      "joiningDate": iso(dt.date(rnd.randint(2016, 2025), rnd.randint(1, 12), rnd.randint(1, 28))),
                                      "status": "Active", "notes": "", "employmentType": "Full-time", "stdHoursPerDay": 8,
                                      "costRate": rate, "skills": title, "supplierId": None, "rateBasis": "hour",
                                      "agreedRate": None, "capacityPct": 100, "contractStart": None, "contractEnd": None}, seed_at))
    outs = [("Outsourced Welder A", "sup_004", 38, "hour"), ("Outsourced Welder B", "sup_004", 38, "hour"),
            ("Outsourced Fitter A", "sup_004", 32, "hour"), ("Outsourced Electrician A", "sup_005", 40, "hour"),
            ("Outsourced Scaffolder A", "sup_005", 280, "day"), ("Outsourced Helper A", "sup_005", 20, "hour"),
            ("Outsourced Crane Operator", "sup_006", 520, "day"), ("Outsourced Helper B", "sup_004", 20, "hour")]
    for j, (n, sid, rate, basis) in enumerate(outs, 1):
        i = len(inhouse) + j
        db["employees"].append(stamp({"id": f"emp_{i:03d}", "code": f"OUT-{j:03d}", "name": n, "type": "outsourced",
                                      "phone": "", "email": "", "title": n.replace("Outsourced ", "")[:-2], "department": "Operations",
                                      "joiningDate": None, "status": "Active", "notes": "", "employmentType": "Contract",
                                      "stdHoursPerDay": 8 if basis == "hour" else 10, "costRate": None, "skills": "",
                                      "supplierId": sid, "rateBasis": basis, "agreedRate": rate,
                                      "capacityPct": 100 if j % 3 else 60,
                                      "contractStart": "2026-01-01", "contractEnd": "2026-12-31"}, seed_at))
    db["leaves"] = [stamp({"id": "lv_001", "employeeId": "emp_006", "fromDate": iso(today + dt.timedelta(days=3)),
                           "toDate": iso(today + dt.timedelta(days=16)), "type": "Annual leave", "notes": ""}, seed_at),
                    stamp({"id": "lv_002", "employeeId": "emp_011", "fromDate": iso(today - dt.timedelta(days=2)),
                           "toDate": iso(today + dt.timedelta(days=2)), "type": "Sick leave", "notes": ""}, seed_at)]

    items = [("CON-WR-001", "Welding rod E7018 3.2mm", "KG", "Consumables", 120, 18.5),
             ("CON-GD-002", "Grinding disc 7in", "PCS", "Consumables", 300, 6.75),
             ("PPE-GL-010", "Safety gloves (pair)", "PCS", "Safety / PPE", 500, 4.2),
             ("ELE-CB-100", "Cable 4C x 16mm2 armoured", "M", "Electrical", 900, 31.0),
             ("MEC-GS-220", "Spiral wound gasket 4in #300", "PCS", "Piping & Valves", 80, 42.0),
             ("MEC-BN-330", "Stud bolt B7 3/4in x 5in", "SET", "Mechanical", 400, 9.6)]
    for i, (sku, n, uom, cat, q, c) in enumerate(items, 1):
        db["inventoryItems"].append(stamp({"id": f"itm_{i:03d}", "sku": sku, "name": n, "uom": uom, "category": cat,
                                           "store": "Main Store — Al Quoz", "qtyOnHand": q, "avgCost": c,
                                           "status": "Active"}, seed_at))

    pe = ["emp_001", "emp_002"]
    se = ["emp_003", "emp_004"]
    techs = [e["id"] for e in db["employees"] if e["type"] == "inhouse"][4:11]
    outsrc = [e["id"] for e in db["employees"] if e["type"] == "outsourced"]
    emp = {e["id"]: e for e in db["employees"]}
    descs = {
        "Service": ["Shutdown maintenance — pump overhaul", "Valve servicing and recertification",
                    "Heat exchanger retubing", "Annual preventive maintenance contract call-out",
                    "Cooling tower fan alignment and balancing", "Electrical panel retrofit and testing"],
        "Supply": ["Supply of gate valves 4in #300", "Supply of armoured cables and glands",
                   "Supply of spiral wound gaskets and studs", "Supply of PPE kits for site crew",
                   "Supply of pump mechanical seals"],
        "Supply/Services": ["Supply and installation of control valves", "Supply, fabrication and erection of pipe spools",
                            "Supply and commissioning of booster pump set", "Supply and installation of cable trays"],
    }
    plan = [  # (type, month, status, closed)
        ("Service", 1, "Closed"), ("Supply", 1, "Closed"), ("Supply/Services", 2, "Closed"), ("Service", 2, "Closed"),
        ("Supply", 3, "Closed"), ("Service", 3, "Completed"), ("Supply/Services", 4, "Closed"), ("Service", 4, "Completed"),
        ("Supply", 5, "Completed"), ("Service", 5, "Closed"), ("Supply/Services", 6, "Completed"), ("Supply", 6, "Cancelled"),
        ("Service", 7, "Completed"), ("Supply", 7, "Completed"), ("Supply/Services", 7, "In Progress"),
        ("Service", 8, "In Progress"), ("Supply", 8, "In Progress"), ("Service", 8, "On Hold"),
        ("Supply/Services", 9, "In Progress"), ("Service", 9, "In Progress"), ("Supply", 9, "Planned"),
        ("Service", 9, "Planned"), ("Supply/Services", 10, "Planned"), ("Service", 10, "Open"), ("Supply", 10, "Open"),
        ("Service", 10, "Open"),
    ]
    seq = 0
    ids = {k: 0 for k in TABLES}

    def nid(t, p):
        ids[t] += 1
        return f"{p}_{ids[t]:04d}"

    for (jtype, month, status) in plan:
        seq += 1
        day = rnd.randint(1, 8 if month == 10 else 26)
        open_d = dt.date(2026, month, day)
        if open_d > today:
            open_d = today - dt.timedelta(days=1)
        at = ts(open_d)
        cust = rnd.choice([c for c in db["customers"] if c["status"] == "Active"])
        ecl = rnd.choice([e for e in db["endClients"] if cust["id"] in e["customerIds"]] or db["endClients"])
        base = {"Service": rnd.randint(18, 160), "Supply": rnd.randint(25, 220), "Supply/Services": rnd.randint(80, 420)}[jtype] * 1000
        po = float(base) if status not in ("Open",) or rnd.random() > .3 else 0.0
        lpo = open_d + dt.timedelta(days=rnd.randint(0, 6)) if po else None
        dur = rnd.randint(7, 40)
        ps = open_d + dt.timedelta(days=rnd.randint(3, 12))
        pf = ps + dt.timedelta(days=dur)
        started = status in ("In Progress", "Completed", "Closed", "On Hold")
        finished = status in ("Completed", "Closed")
        a_start = ps + dt.timedelta(days=rnd.randint(-1, 4)) if started else None
        a_fin = a_start + dt.timedelta(days=dur + rnd.randint(-4, 9)) if finished else None
        if a_fin and a_fin > today:
            a_fin = today - dt.timedelta(days=1)
        if status == "In Progress" and a_start and a_start > today:
            a_start = today - dt.timedelta(days=2)
        pct = {"Open": 0, "Planned": 0, "In Progress": rnd.choice([20, 35, 50, 65, 80]), "On Hold": 30,
               "Completed": 100, "Closed": 100, "Cancelled": 0}[status]
        jid = f"job_{seq:04d}"
        job = stamp({"id": jid, "jobNo": job_no(S, jtype, 2026, seq), "type": jtype, "openingDate": iso(open_d),
                     "customerId": cust["id"], "endClientId": ecl["id"], "zohoSo": f"SO-{26000 + seq * 7}" if po else "",
                     "zohoSoId": None, "description": rnd.choice(descs[jtype]), "status": status, "completionPct": pct,
                     "lpoDate": iso(lpo), "customerOrderNo": f"PO/{cust['code']}/{rnd.randint(1000, 9999)}" if po else "",
                     "poValue": po, "actualPoValue": None, "psDate": iso(ps), "pfDate": iso(pf), "actualStart": iso(a_start),
                     "actualFinish": iso(a_fin), "projectEngineerId": rnd.choice(pe),
                     "serviceEngineerId": rnd.choice(se) if jtype != "Supply" else None,
                     "comments": "", "locked": False, "closureDate": None, "closureComments": None, "closedBy": None,
                     "closedAt": None}, at, "usr_pm")
        if finished and rnd.random() < .3 and po:
            job["actualPoValue"] = round(po * rnd.choice([1.05, 1.08, 0.97]), 2)
            db["poRevisions"].append({"id": nid("poRevisions", "por"), "jobId": jid, "field": "actualPoValue",
                                      "fromValue": po, "toValue": job["actualPoValue"], "at": ts(a_fin), "userId": "usr_fin",
                                      "reason": "Variation order approved by customer"})
        db["jobs"].append(job)
        db["jobStatusHistory"].append({"id": nid("jobStatusHistory", "jsh"), "jobId": jid, "fromStatus": None,
                                       "toStatus": "Open", "at": at, "userId": "usr_pm", "note": "Job created"})
        if status != "Open":
            db["jobStatusHistory"].append({"id": nid("jobStatusHistory", "jsh"), "jobId": jid, "fromStatus": "Open",
                                           "toStatus": status, "at": ts(a_fin or a_start or open_d), "userId": "usr_pm",
                                           "note": ""})
        if not po or status == "Cancelled":
            continue
        drift = {"Closed": rnd.uniform(.88, 1.22), "Completed": rnd.uniform(.9, 1.18),
                 "In Progress": rnd.uniform(.35, .75), "On Hold": .3}.get(status, 0)
        target = po * rnd.uniform(.55, .78)  # planned cost
        shares = {"Service": (.55, .2, .25), "Supply": (.08, .82, .10), "Supply/Services": (.35, .5, .15)}[jtype]
        # planned manpower
        mp_budget = target * shares[0]
        crew = rnd.sample(techs, 2) + ([rnd.choice(outsrc)] if jtype != "Supply" else [])
        for e_id in crew:
            e = emp[e_id]
            rate = e["costRate"] if e["type"] == "inhouse" else (e["agreedRate"] / (10 if e["rateBasis"] == "day" else 1))
            hrs = max(8, round(mp_budget / len(crew) / rate / 8) * 8)
            pml = stamp({"id": nid("plannedManpower", "pmp"), "jobId": jid, "employeeId": e_id, "resourceName": e["name"],
                         "resourceType": e["type"], "supplierId": e.get("supplierId"), "headcount": 1, "hoursPerPerson": hrs,
                         "rate": rate, "startDate": iso(ps), "endDate": iso(pf), "comments": "", "assignmentId": None}, at, "usr_pm")
            if status in ("Planned", "In Progress", "On Hold"):
                aid = nid("assignments", "asg")
                days = max(1, (pf - ps).days)
                db["assignments"].append(stamp({"id": aid, "employeeId": e_id, "jobId": jid, "startDate": iso(ps),
                                                "endDate": iso(pf), "plannedHours": hrs,
                                                "allocationPct": min(100, max(25, round(hrs / (days * 8 * 5 / 7) * 100 / 25) * 25)),
                                                "status": "Assigned" if status != "Planned" else "Planned",
                                                "comments": "", "overrideBy": None, "overrideReason": None}, at, "usr_hr"))
                pml["assignmentId"] = aid
            db["plannedManpower"].append(pml)
            if drift:
                am_hrs = max(4, round(hrs * drift * rnd.uniform(.9, 1.1)))
                db["actualManpower"].append(stamp({"id": nid("actualManpower", "amp"), "jobId": jid, "employeeId": e_id,
                                                   "resourceName": e["name"], "resourceType": e["type"],
                                                   "supplierId": e.get("supplierId"), "headcount": 1,
                                                   "hoursPerPerson": am_hrs, "rate": rate, "dateFrom": iso(a_start),
                                                   "dateTo": iso(a_fin or today), "comments": "",
                                                   "assignmentId": pml["assignmentId"]}, at, "usr_pm"))
        # materials
        mat_budget = target * shares[1]
        n = 2 if jtype == "Service" else 3
        for k in range(n):
            desc, cat, uom = rnd.choice([("Gate valve 4in #300 RF", "Piping & Valves", "PCS"),
                                         ("Mechanical seal cartridge", "Spare Parts", "SET"),
                                         ("Armoured cable 4C x 25mm2", "Electrical", "M"),
                                         ("Control valve 3in with positioner", "Piping & Valves", "PCS"),
                                         ("Carbon steel pipe 6in SCH40", "Mechanical", "M"),
                                         ("Cable tray 300mm HDG", "Electrical", "M"),
                                         ("Bearing set SKF", "Spare Parts", "SET")])
            unit = round(rnd.uniform(40, 2400), 2)
            qty = max(1, round(mat_budget / n / unit))
            sid = rnd.choice(["sup_001", "sup_002", "sup_003", "sup_008", "sup_009"])
            db["plannedMaterials"].append(stamp({"id": nid("plannedMaterials", "pmt"), "jobId": jid, "description": desc,
                                                 "category": cat, "supplierId": sid, "qty": qty, "uom": uom, "unitPrice": unit,
                                                 "expectedDate": iso(ps), "comments": ""}, at, "usr_pm"))
            if drift:
                aq = qty if status in ("Closed", "Completed") else max(1, round(qty * drift))
                db["actualMaterials"].append(stamp({"id": nid("actualMaterials", "amt"), "jobId": jid, "description": desc,
                                                    "category": cat, "supplierId": sid, "qty": aq, "uom": uom,
                                                    "unitPrice": round(unit * rnd.uniform(.95, 1.12), 2),
                                                    "purchaseDate": iso(a_start), "supplierRef": f"INV-{rnd.randint(10000, 99999)}",
                                                    "comments": ""}, at, "usr_ops"))
        if drift and jtype != "Supply":
            it = rnd.choice(db["inventoryItems"])
            q = rnd.randint(5, 40)
            db["inventoryTransfers"].append(stamp({"id": nid("inventoryTransfers", "itr"), "jobId": jid, "itemId": it["id"],
                                                   "qty": q, "uom": it["uom"], "transferDate": iso(a_start),
                                                   "issuedBy": "Michael Santos", "receivedBy": emp[crew[0]]["name"],
                                                   "store": it["store"], "costingMethod": "Weighted average",
                                                   "unitCost": it["avgCost"], "costSource": "weighted_average",
                                                   "ref": f"ST-{2026}-{ids['inventoryTransfers']:04d}", "status": "Posted",
                                                   "comments": ""}, at, "usr_ops"))
        # misc
        misc_budget = target * shares[2]
        cats = rnd.sample(S["expenseCategories"][:6], 2)
        for cat in cats:
            amt = round(misc_budget / 2, -1)
            sid = {"Accommodation": "sup_010", "Equipment Rental": "sup_006", "Delivery": "sup_007",
                   "Transportation": "sup_007"}.get(cat)
            db["plannedMisc"].append(stamp({"id": nid("plannedMisc", "pms"), "jobId": jid, "category": cat,
                                            "description": f"{cat} — {jtype.lower()} job", "supplierId": sid,
                                            "payee": "" if sid else "Various", "expenseDate": iso(ps), "amount": amt,
                                            "comments": ""}, at, "usr_pm"))
            if drift:
                aamt = round(amt * drift * rnd.uniform(.85, 1.15), 2)
                mid = nid("actualMisc", "ams")
                mode = rnd.choice(S["paymentModes"][:4])
                db["actualMisc"].append(stamp({"id": mid, "jobId": jid, "category": cat,
                                               "description": f"{cat} — {jtype.lower()} job", "supplierId": sid,
                                               "payee": "" if sid else "Various", "expenseDate": iso(a_start), "amount": aamt,
                                               "paymentMode": mode, "paymentRef": f"PV-{rnd.randint(1000, 9999)}",
                                               "attachmentName": None, "comments": ""}, at, "usr_ops"))
                r = rnd.random()
                if status in ("Closed",) or r < .55:
                    db["miscPayments"].append(stamp({"id": nid("miscPayments", "mpy"), "miscId": mid, "payDate": iso(a_start),
                                                     "amount": aamt, "mode": mode, "ref": "Paid in full"}, at, "usr_fin"))
                elif r < .8:
                    db["miscPayments"].append(stamp({"id": nid("miscPayments", "mpy"), "miscId": mid, "payDate": iso(a_start),
                                                     "amount": round(aamt * .5, 2), "mode": mode, "ref": "Advance 50%"}, at, "usr_fin"))
        # invoices
        if status in ("Completed", "Closed") or (status == "In Progress" and rnd.random() < .5):
            inv_amt = float(job["actualPoValue"] or po) * (1 if finished else .4)
            inv_d = (a_fin or today - dt.timedelta(days=5)) + dt.timedelta(days=rnd.randint(1, 6))
            if inv_d > today:
                inv_d = today
            iid = nid("invoices", "inv")
            db["invoices"].append(stamp({"id": iid, "jobId": jid, "invoiceNo": f"INV-26-{1000 + ids['invoices']}",
                                         "invoiceDate": iso(inv_d), "dueDate": iso(inv_d + dt.timedelta(days=30)),
                                         "amount": round(inv_amt, 2), "status": "Unpaid", "zohoInvoiceId": None,
                                         "notes": "" if finished else "Progress invoice 40%"}, ts(inv_d), "usr_fin"))
            r = rnd.random()
            paid = 0
            if status == "Closed" and r < .7 or r < .3:
                paid = round(inv_amt, 2)
            elif r < .75:
                paid = round(inv_amt * rnd.choice([.3, .5, .6]), 2)
            if paid:
                pd_ = min(today, inv_d + dt.timedelta(days=rnd.randint(10, 45)))
                db["invoicePayments"].append(stamp({"id": nid("invoicePayments", "ipy"), "invoiceId": iid,
                                                    "payDate": iso(pd_), "amount": paid, "mode": "Bank Transfer",
                                                    "ref": f"TT-{rnd.randint(100000, 999999)}", "zohoPaymentId": None},
                                                   ts(pd_), "usr_fin"))
                db["invoices"][-1]["status"] = "Paid" if paid >= round(inv_amt, 2) else "Partially Paid"
        if status == "Closed":
            cd = min(today, (a_fin or today) + dt.timedelta(days=rnd.randint(5, 25)))
            job.update({"locked": True, "closureDate": iso(cd), "closureComments": "All deliverables accepted by client.",
                        "closedBy": "usr_fin", "closedAt": ts(cd)})
            db["closures"].append({"id": nid("closures", "cls"), "jobId": jid, "closureDate": iso(cd),
                                   "comments": "All deliverables accepted by client.", "finalStatus": "Closed",
                                   "financialsConfirmed": True, "pendingConfirmed": True, "closedBy": "usr_fin",
                                   "closedAt": ts(cd), "reopenedBy": None, "reopenedAt": None, "reopenReason": None})
        if rnd.random() < .4:
            db["jobComments"].append({"id": nid("jobComments", "cmt"), "jobId": jid, "userId": "usr_pm", "at": at,
                                      "text": rnd.choice(["Kick-off meeting held with client.", "Awaiting gate passes from site.",
                                                          "Material delivered to site store.", "Client requested weekend shift."])})
    S["jobNumber"]["nextSeq"] = seq + 1
    # a couple of roster assignments not yet in any job plan (future planning) — linked via plannedManpower for consistency
    db["auditLog"].append({"id": "aud_0001", "at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                           "userId": "usr_super", "action": "seed", "entity": "system", "recordId": "-",
                           "before": None, "after": {"note": "Demo dataset generated by convert.py demo"}, "reason": None})
    return db


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def validate(db, verbose=True):
    errors, warns = [], []
    idx = {t: {r.get("id"): r for r in db.get(t, [])} for t in TABLES}
    for t, cols in TABLES.items():
        rows = db.get(t, [])
        seen = set()
        for r in rows:
            rid = r.get("id")
            if rid in seen:
                errors.append(f"{t}: duplicate id {rid}")
            seen.add(rid)
            for c in cols:
                v = r.get(c["name"])
                if c["req"] and (v is None or v == ""):
                    errors.append(f"{t}.{c['name']} required (id={rid})")
                if c["fk"] and v not in (None, ""):
                    if v not in idx[c["fk"]]:
                        errors.append(f"{t}.{c['name']} → {c['fk']} missing '{v}' (id={rid})")
                if c["type"] == "date" and v:
                    try:
                        dt.date.fromisoformat(v)
                    except Exception:
                        errors.append(f"{t}.{c['name']} invalid date '{v}' (id={rid})")
                if c["type"] == "num" and v not in (None, ""):
                    try:
                        if D(v) < 0:
                            errors.append(f"{t}.{c['name']} negative ({v}) (id={rid})")
                    except Exception:
                        errors.append(f"{t}.{c['name']} not numeric '{v}' (id={rid})")
    nos = [j["jobNo"] for j in db.get("jobs", [])]
    for n in set(nos):
        if nos.count(n) > 1:
            errors.append(f"jobs: duplicate jobNo {n}")
    for j in db.get("jobs", []):
        if j.get("type") not in JOB_TYPES:
            errors.append(f"jobs: invalid type '{j.get('type')}' ({j['jobNo']})")
        p = j.get("completionPct")
        if p not in (None, "") and not (0 <= float(p) <= 100):
            errors.append(f"jobs: completionPct out of range ({j['jobNo']})")
        for a, b in (("psDate", "pfDate"), ("actualStart", "actualFinish")):
            if j.get(a) and j.get(b) and j[b] < j[a]:
                errors.append(f"jobs: {b} before {a} ({j['jobNo']})")
        if j.get("status") == "Closed" and not j.get("locked"):
            warns.append(f"jobs: {j['jobNo']} is Closed but not locked")
    refs = [t["ref"] for t in db.get("inventoryTransfers", []) if t.get("status") == "Posted"]
    for r in set(refs):
        if refs.count(r) > 1:
            errors.append(f"inventoryTransfers: transfer ref {r} posted twice (double charge)")
    for inv in db.get("invoices", []):
        paid = sum((r2(p["amount"]) for p in db.get("invoicePayments", []) if p["invoiceId"] == inv["id"]), Decimal(0))
        exp = "Paid" if paid >= r2(inv["amount"]) and paid > 0 else ("Partially Paid" if paid > 0 else "Unpaid")
        if inv.get("status") != exp:
            warns.append(f"invoices: {inv['invoiceNo']} status '{inv.get('status')}' but payments imply '{exp}'")
    if verbose:
        print(f"Tables: " + ", ".join(f"{t}={len(db.get(t, []))}" for t in TABLES if db.get(t)))
        for e in errors:
            print("  ERROR ", e)
        for w in warns:
            print("  WARN  ", w)
        print(f"{len(errors)} error(s), {len(warns)} warning(s)")
    return errors, warns


def print_financials(db):
    cur = db["settings"].get("currency", "AED")
    print(f"\n{'Job':<22}{'Status':<13}{'PO':>13}{'Plan cost':>13}{'Act cost':>13}{'Plan GP%':>9}{'Act GP%':>9}")
    for j in db["jobs"]:
        f = job_financials(db, j)
        fp = lambda x: "N/A" if x is None else f"{x:.1f}"
        print(f"{j['jobNo']:<22}{j['status']:<13}{f['plannedPO']:>13,.2f}{f['plannedCost']:>13,.2f}"
              f"{f['actualCost']:>13,.2f}{fp(f['plannedGPPct']):>9}{fp(f['actualGPPct']):>9}")
    print(f"(amounts in {cur})")


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
def run_tests():
    ok = 0
    fails = []

    def check(name, cond):
        nonlocal ok
        if cond:
            ok += 1
            print(f"  PASS  {name}")
        else:
            fails.append(name)
            print(f"  FAIL  {name}")

    db = empty_db()
    job = {"id": "j1", "poValue": 10000, "actualPoValue": None}
    db["jobs"] = [job]
    db["plannedMisc"] = [{"id": "a", "jobId": "j1", "amount": 6000}]
    f = job_financials(db, job)
    check("T4 planned GP = 4,000", f["plannedGP"] == Decimal("4000.00"))
    check("T4 planned GP% = 40%", f["plannedGPPct"] == Decimal("40.00"))
    db["actualMisc"] = [{"id": "b", "jobId": "j1", "amount": 7000}]
    f = job_financials(db, job)
    check("T5 actual GP = 3,000", f["actualGP"] == Decimal("3000.00"))
    check("T5 actual GP% = 30%", f["actualGPPct"] == Decimal("30.00"))
    check("T5 cost variance = +1,000 (unfavourable)", f["costVariance"] == Decimal("1000.00"))
    job["poValue"] = 0
    f = job_financials(db, job)
    check("T13 zero PO → GP% N/A (no division error)", f["plannedGPPct"] is None and f["actualGPPct"] is None)
    job["poValue"] = None
    check("T13 blank PO handled", job_financials(db, job)["plannedPO"] == Decimal("0.00"))
    check("Manpower line = headcount × hours/person × rate", line_mp({"headcount": 3, "hoursPerPerson": 8, "rate": 45.5}) == Decimal("1092.00"))
    check("Decimal amounts round half-up per line", line_qty({"qty": 3, "unitPrice": 0.335}) == Decimal("1.01"))
    db2 = empty_db()
    j2 = {"id": "j2", "poValue": 1000}
    db2["jobs"] = [j2]
    db2["inventoryTransfers"] = [{"id": "t1", "jobId": "j2", "qty": 10, "unitCost": 5, "status": "Posted", "ref": "A"},
                                 {"id": "t2", "jobId": "j2", "qty": 10, "unitCost": 5, "status": "Reversed", "ref": "B"}]
    db2["actualMaterials"] = [{"id": "m", "jobId": "j2", "qty": 2, "unitPrice": 100}]
    f2 = job_financials(db2, j2)
    check("Internal transfer costed (not free) and reversed transfer excluded",
          f2["actualMaterialTransfer"] == Decimal("50.00") and f2["actualMaterial"] == Decimal("250.00"))
    j2["actualPoValue"] = 1200
    check("Revised actual PO value used for actual GP", job_financials(db2, j2)["actualPO"] == Decimal("1200.00"))
    db2["settings"]["revenueRule"] = "invoiced"
    db2["invoices"] = [{"id": "i", "jobId": "j2", "amount": 900}]
    check("Revenue rule 'invoiced' uses invoice total", job_financials(db2, j2)["actualPO"] == Decimal("900.00"))
    demo = build_demo()
    e, _ = validate(demo, verbose=False)
    check("Demo dataset passes referential-integrity validation", not e)
    print(f"\n{ok} passed, {len(fails)} failed")
    return not fails


# ---------------------------------------------------------------------------
# Excel template / import
# ---------------------------------------------------------------------------
def need_openpyxl():
    try:
        import openpyxl  # noqa
        return openpyxl
    except ImportError:
        sys.exit("openpyxl is required for this command:  pip install openpyxl")


def make_template(out):
    ox = need_openpyxl()
    from openpyxl.styles import Font, PatternFill
    wb = ox.Workbook()
    ws = wb.active
    ws.title = "README"
    lines = ["JobOps import template", "",
             "One sheet per table. Row 1 = column keys (do not rename). Leave unknown cells blank.",
             "id: any unique text (e.g. cus_001). Foreign keys (…Id columns) must reference an id on the related sheet.",
             "Dates: YYYY-MM-DD.  Numbers: plain numbers, no currency symbols.  customerIds: comma-separated ids.",
             "jobs.type: Service | Supply | Supply/Services.   employees.type: inhouse | outsourced.",
             "Run:  python3 convert.py from-excel this_file.xlsx --base data.json --out data.json"]
    for i, l in enumerate(lines, 1):
        ws.cell(row=i, column=1, value=l)
    ws["A1"].font = Font(bold=True, size=14)
    for t in IMPORTABLE:
        sh = wb.create_sheet(t[:31])
        for ci, c in enumerate(TABLES[t], 1):
            cell = sh.cell(row=1, column=ci, value=c["name"])
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244DC7" if c["req"] else "4F607E")
            sh.column_dimensions[cell.column_letter].width = max(12, len(c["name"]) + 4)
        sh.freeze_panes = "A2"
    wb.save(out)
    print(f"wrote {out}  (dark-blue headers = required)")


def coerce(v, typ):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return [] if typ == "arr" else None
    if typ == "num":
        return float(v)
    if typ == "int":
        return int(float(v))
    if typ == "bool":
        return str(v).strip().lower() in ("1", "true", "yes", "y")
    if typ == "date":
        if isinstance(v, (dt.datetime, dt.date)):
            return (v.date() if isinstance(v, dt.datetime) else v).isoformat()
        return dt.date.fromisoformat(str(v).strip()[:10]).isoformat()
    if typ == "arr":
        return [s.strip() for s in str(v).split(",") if s.strip()]
    return str(v).strip()


def from_excel(path, base, out):
    ox = need_openpyxl()
    db = load(base) if base and os.path.exists(base) else empty_db()
    wb = ox.load_workbook(path, data_only=True)
    now = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    added = {}
    for t in IMPORTABLE:
        if t not in wb.sheetnames:
            continue
        sh = wb[t]
        rows = list(sh.iter_rows(values_only=True))
        if len(rows) < 2:
            continue
        hdr = [str(h).strip() if h else "" for h in rows[0]]
        types = {c["name"]: c["type"] for c in TABLES[t]}
        existing = {r["id"]: r for r in db.get(t, [])}
        for row in rows[1:]:
            if not any(v not in (None, "") for v in row):
                continue
            rec = {}
            for h, v in zip(hdr, row):
                if h in types:
                    rec[h] = coerce(v, types[h])
            if not rec.get("id"):
                continue
            stamp(rec, now)
            if rec["id"] in existing:
                existing[rec["id"]].update({k: v for k, v in rec.items() if k not in ("createdAt", "createdBy")})
            else:
                db.setdefault(t, []).append(rec)
            added[t] = added.get(t, 0) + 1
    if db.get("jobs"):
        db["meta"]["demo"] = db["meta"].get("demo", False) and not added.get("jobs")
        mx = 0
        for j in db["jobs"]:
            tail = "".join(ch for ch in j["jobNo"].split("-")[-1] if ch.isdigit())
            mx = max(mx, int(tail) if tail else 0)
        db["settings"]["jobNumber"]["nextSeq"] = max(db["settings"]["jobNumber"]["nextSeq"], mx + 1)
    print("imported: " + (", ".join(f"{k}={v}" for k, v in added.items()) or "nothing"))
    errors, _ = validate(db)
    if errors:
        sys.exit("Import NOT written — fix the errors above.")
    save(db, out)


# ---------------------------------------------------------------------------
# SQL export (Supabase / PostgreSQL)
# ---------------------------------------------------------------------------
SQL_T = {"text": "text", "num": "numeric(14,2)", "int": "integer", "date": "date", "ts": "timestamptz",
         "bool": "boolean", "arr": "text[]", "json": "jsonb"}


def snake(s):
    return "".join("_" + ch.lower() if ch.isupper() else ch for ch in s).lstrip("_")


def sql_lit(v, typ):
    if v is None or v == "":
        return "NULL"
    if typ in ("num", "int"):
        return str(v)
    if typ == "bool":
        return "true" if v else "false"
    if typ == "arr":
        return "ARRAY[" + ",".join("'" + str(x).replace("'", "''") + "'" for x in v) + "]::text[]" if v else "'{}'::text[]"
    if typ == "json":
        return "'" + json.dumps(v).replace("'", "''") + "'::jsonb"
    return "'" + str(v).replace("'", "''") + "'"


def to_sql(db, out):
    order = ["roles", "users", "customers", "endClients", "suppliers", "employees", "leaves", "inventoryItems",
             "stockReceipts", "jobs", "assignments", "plannedManpower", "actualManpower", "plannedMaterials",
             "actualMaterials", "inventoryTransfers", "plannedMisc", "actualMisc", "miscPayments", "invoices",
             "invoicePayments", "jobStatusHistory", "poRevisions", "jobComments", "jobAttachments", "closures", "auditLog"]
    L = ["-- JobOps schema + seed generated by convert.py to-sql", f"-- {dt.datetime.now().isoformat(timespec='seconds')}",
         "-- Phase-2 starting point: review, then add auth.users mapping, RLS policies and server-side functions.", "begin;", ""]
    L.append("create table if not exists app_settings (id int primary key default 1 check (id = 1), data jsonb not null);")
    for t in order:
        cols = []
        for c in TABLES[t]:
            d = f"  {snake(c['name'])} {SQL_T[c['type']]}"
            if c["name"] == "id":
                d += " primary key"
            elif c["req"]:
                d += " not null"
            cols.append(d)
        if t in BUSINESS:
            cols += ["  created_at timestamptz default now()", "  created_by text", "  updated_at timestamptz default now()",
                     "  version integer not null default 1", "  archived boolean not null default false"]
        L.append(f"create table if not exists {snake(t)} (\n" + ",\n".join(cols) + "\n);")
    L.append("")
    for t in order:
        for c in TABLES[t]:
            if c["fk"]:
                L.append(f"alter table {snake(t)} add constraint fk_{snake(t)}_{snake(c['name'])} foreign key "
                         f"({snake(c['name'])}) references {snake(c['fk'])}(id) on delete restrict;")
                L.append(f"create index if not exists ix_{snake(t)}_{snake(c['name'])} on {snake(t)}({snake(c['name'])});")
    L += ["alter table jobs add constraint uq_jobs_job_no unique (job_no);",
          "alter table jobs add constraint ck_jobs_type check (type in ('Service','Supply','Supply/Services'));",
          "alter table jobs add constraint ck_jobs_pct check (completion_pct between 0 and 100);",
          "alter table inventory_transfers add constraint uq_transfer_ref unique (ref);",
          "alter table invoices add constraint uq_invoice_no unique (invoice_no);", ""]
    L.append("insert into app_settings (id, data) values (1, " + sql_lit(db["settings"], "json") + ") on conflict (id) do update set data = excluded.data;")
    for t in order:
        rows = db.get(t, [])
        if not rows:
            continue
        names = [c["name"] for c in TABLES[t]]
        extra = ["createdAt", "createdBy", "updatedAt", "version", "archived"] if t in BUSINESS else []
        types = {c["name"]: c["type"] for c in TABLES[t]}
        types.update({"createdAt": "ts", "createdBy": "text", "updatedAt": "ts", "version": "int", "archived": "bool"})
        allc = names + extra
        L.append(f"\ninsert into {snake(t)} (" + ", ".join(snake(n) for n in allc) + ") values")
        vals = []
        for r in rows:
            vals.append("  (" + ", ".join(sql_lit(r.get(n, 1 if n == "version" else (False if n == "archived" else None)), types[n]) for n in allc) + ")")
        L.append(",\n".join(vals) + ";")
    L += ["", "commit;"]
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print(f"wrote {out}  ({os.path.getsize(out):,} bytes)")


# ---------------------------------------------------------------------------
# Server
# ---------------------------------------------------------------------------
def serve(port):
    os.chdir(HERE)

    class H(http.server.SimpleHTTPRequestHandler):
        def end_headers(self):
            self.send_header("Cache-Control", "no-store")
            super().end_headers()

        def log_message(self, *a):
            pass

    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", port), H) as httpd:
        print(f"JobOps running →  http://localhost:{port}/index.html   (Ctrl+C to stop)")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nstopped")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    p = sp.add_parser("serve"); p.add_argument("--port", type=int, default=8000)
    p = sp.add_parser("demo"); p.add_argument("--out", default=os.path.join(HERE, "data.json"))
    p = sp.add_parser("empty"); p.add_argument("--out", default=os.path.join(HERE, "data.json"))
    p = sp.add_parser("validate"); p.add_argument("file", nargs="?", default=os.path.join(HERE, "data.json"))
    sp.add_parser("test")
    p = sp.add_parser("template"); p.add_argument("--out", default="jobops_import_template.xlsx")
    p = sp.add_parser("from-excel"); p.add_argument("file"); p.add_argument("--out", default=os.path.join(HERE, "data.json"))
    p.add_argument("--base", default=os.path.join(HERE, "data.json"))
    p = sp.add_parser("to-sql"); p.add_argument("file", nargs="?", default=os.path.join(HERE, "data.json"))
    p.add_argument("--out", default="supabase_seed.sql")
    a = ap.parse_args()
    if a.cmd == "serve":
        serve(a.port)
    elif a.cmd == "demo":
        db = build_demo()
        validate(db)
        save(db, a.out)
    elif a.cmd == "empty":
        db = empty_db()
        db["users"] = [stamp({"id": "usr_super", "name": "System Administrator", "email": "admin@example.com",
                              "roleId": "role_super", "department": "Management", "active": True},
                             dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"))]
        save(db, a.out)
    elif a.cmd == "validate":
        db = load(a.file)
        errors, _ = validate(db)
        print_financials(db)
        sys.exit(1 if errors else 0)
    elif a.cmd == "test":
        sys.exit(0 if run_tests() else 1)
    elif a.cmd == "template":
        make_template(a.out)
    elif a.cmd == "from-excel":
        from_excel(a.file, a.base, a.out)
    elif a.cmd == "to-sql":
        to_sql(load(a.file), a.out)


if __name__ == "__main__":
    main()
