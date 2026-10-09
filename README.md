# JobOps — Job Management & Operations System (local prototype)

A working, single-page prototype of the full system in the brief: Overview, Jobs (Service / Supply / Supply-Services), job costing (manpower, materials, internal store transfers, miscellaneous), financial summary, invoicing, closure and locking, Roster Planner, Master Data, 13 reports, Settings with roles and permissions, and an audit trail.

The point of this release is to **check the workflows, screens and rules before building the Supabase version**. Every business rule runs through one data layer (`API.*` in `index.html`). That layer is written so it can move to Postgres functions and Row Level Security (RLS) later.

> **What this prototype is not:** it is not multi-user and it is not secure. Data is saved in your browser (localStorage), sign-in is a demo user switcher with no passwords, and permission checks run in the browser. All of that is replaced in the Supabase phase (see below). The included data is **fictional demo data**, and the app shows a "Demo data" badge while it is loaded.

---

## 1. Quick start

```bash
cd jobops                       # the folder with index.html, data.json, convert.py
python3 convert.py serve        # → http://localhost:8000/index.html
```

You need Python 3.8+ and no other packages. Browsers block reading `data.json` when the page is opened by double-click (`file://`), so start it with the server. If you do double-click it anyway, the page offers a "Choose data.json" button as a fallback.

**Data flow:** on first open the app loads `data.json` and keeps a working copy in that browser. Your edits survive a page refresh. To keep or move them, use **Export data.json** (top bar). To start over, use **Settings → Data & integrations → Discard browser copy & reload data.json**.

| File | Purpose |
|---|---|
| `index.html` | The whole application (HTML/CSS/JS, no build step, no frameworks, charts drawn as inline SVG) |
| `data.json` | Demo dataset: 26 jobs, 22 employees, 8 customers, 10 suppliers, 8 users/roles |
| `convert.py` | Serve, demo/empty data, validation, formula tests, Excel template and import, Supabase SQL export |
| `README.md` | This file |

---

## 2. Demo users (switch at the bottom of the sidebar)

| User | Role | What to try |
|---|---|---|
| Sara Admin | Super Admin | Everything, including Settings, roles and acceptance tests |
| Omar Haddad | Administrator | Full access |
| Layla Mansour | Management | Read-only with full financials and export |
| Priya Nair | Finance / Accounts | Invoices, payments and misc expenses; can close jobs but can't edit them; can change the invoice status of **closed** jobs |
| Daniel Fernandes | Project Manager | Creates and edits jobs and costs, assigns roster; **cannot see employee cost rates**; cannot override conflicts |
| Imran Qureshi | Operations | Edits jobs, materials and misc; **no financial visibility** (values masked •••) |
| Mariam Saeed | HR / Resource Planner | Roster and master data; can override roster conflicts; no financials |
| Guest Viewer | Read-Only | Views tabs only; no financials; no export |

---

## 3. How to test it

**In the app, under Settings → Acceptance tests → Run all tests.** This runs the brief's 15 acceptance workflows against a throw-away copy of the data, using the same data layer the screens use, and then restores your data. The current result is **15/15 pass**.

| # | What is verified |
|---|---|
| 1–2 | Jobs are created with unique sequential numbers, survive a reload, and appear in the right type section |
| 3 | Manpower, material, store-transfer and misc totals and variances; a duplicate transfer reference is blocked |
| 4–5 | PO 10,000 − cost 6,000 = GP 4,000 (40%); actual cost 7,000 gives GP 3,000 (30%) |
| 6–7 | A roster assignment appears in the calendar and the job's manpower plan; overbooking warns, and only users with override permission can force it, with a reason |
| 8–10 | Closed jobs are locked for ordinary users and against direct API calls; Finance can still update invoice status |
| 11 | A custom restricted role loses tabs, financial fields, export and job types |
| 12–15 | Changes reach the overview and reports; edge cases (zero PO, decimals, partial payments, double submit, bad dates); report export scope; stale concurrent saves are rejected |

**On the command line:**
```bash
python3 convert.py test       # 13 formula unit tests (Decimal-based, mirrors the JS engine)
python3 convert.py validate   # referential integrity + recomputed P&L per job
```

**Manual walk-through (about 10 minutes):**
1. Overview: click any KPI or chart bar to open the filtered job list.
2. Jobs → Service → **Create job**. Try a PF date before the PS date to see validation, then save.
3. On the new job, open Manpower, Materials (including **Post transfer**) and Miscellaneous, and add lines. Then check **Financial summary**.
4. Roster Planner → **Assign to job**. Book the same person twice on overlapping dates to see the conflict warning.
5. On the job, click **Close job**. Switch to *Imran (Operations)* and confirm nothing can be edited. Switch to *Priya (Finance)*, then go to Invoices → **Change status**.
6. Reports: open any report, filter it, and export CSV or Excel, or print to PDF.
7. Settings → Roles & permissions: create a custom role and assign it to a user.

---

## 4. Business rules and assumptions

The brief left some of these rules open. Each default below can be changed in Settings or in one place in code.

| Topic | Default used |
|---|---|
| Rounding | Each line is rounded to 2 dp (half-up). Totals are sums of rounded lines. Values are stored unrounded. Display precision is configurable. |
| Actual PO value | Equals the PO value unless a revised value is entered. A revision requires *Edit financial data* plus a reason, and is logged in PO revisions. An alternative rule (Settings) uses the total invoiced amount. |
| GP % | GP ÷ PO × 100. Shows **N/A** when the PO value is 0 or blank. |
| Cancelled jobs | Excluded from all financial totals and charts. They are still counted in job counts. |
| Overview "Actual GP" | Includes in-progress jobs, whose costs are still coming in. A separate **realised margin** figure (completed and closed jobs only) is shown alongside it so the comparison is fair. |
| Completion % vs status | These are independent. Setting the status to *Completed* only **offers** to set 100%; nothing is overwritten. Only **Close job** sets *Closed*. |
| Closed jobs | Locked for everyone except *Edit closed jobs*, and that requires a reason. Invoice status and payments stay editable for users with invoice permissions. *Reopen* requires a reason. |
| Invoice status | Uninvoiced, Unpaid, Partially Paid or Paid, **derived from recorded payments**. *Paid* is refused until the full amount is received. Every change is audited. |
| Misc. payment status | Derived from payments. A partial payment never marks an expense Paid. |
| Store transfers | Costed at the item's weighted-average cost, or at a confirmed manual cost (which needs *Edit financial data*). A reference can only be posted once. Reversal is used instead of deletion. Receiving stock re-averages the cost. |
| Manpower lines | Cost = headcount × hours **per person** × rate. The rate is copied onto the line, so later changes to a master rate don't alter history. |
| Roster ↔ job plan | The assignment is the single source of truth. Its linked planned-manpower line is created, updated and removed automatically. Completing an assignment can post actual hours. |
| Capacity | Working days (UAE default Mon–Fri) × hours per day × capacity % (outsourced resources can be below 100%), minus approved leave. Overbooking above capacity is blocked unless overridden. |
| Deletion | All business records are soft-deleted (archived) and the audit log is append-only. |
| Concurrency | Records carry a `version`. Saving an outdated copy is rejected with a "changed by X at T" message. Other browser tabs refresh automatically. |
| Job numbers | `{PREFIX}-{TYPE}-{YYYY}-{SEQ}` (e.g. JOB-SRV-2026-0027). They are never reused. Changing the format only affects new jobs, and Settings blocks a sequence that would collide with an existing number. |

---

## 5. Prototype limitations (and where they're solved)

| Prototype | Supabase phase |
|---|---|
| Data stored in one browser | Postgres, shared by every user and device |
| Demo user switcher | Supabase Auth (email or SSO), invitations, password reset |
| Permission checks in the browser | RLS policies plus `security definer` functions; the browser cannot bypass them |
| Attachments store the file name and size only | Supabase Storage with per-job buckets and type/size checks |
| Notification settings are stored but not sent | Email via an Edge Function |
| Zoho Books not connected (shown as such) | OAuth integration (§7) |

---

## 6. `convert.py` reference

```bash
python3 convert.py serve [--port 8000]
python3 convert.py demo   [--out data.json]          # regenerate demo data
python3 convert.py empty  [--out data.json]          # clean start: settings, roles, 1 admin
python3 convert.py validate [data.json]
python3 convert.py test
python3 convert.py template --out jobops_import_template.xlsx      # needs: pip install openpyxl
python3 convert.py from-excel your_file.xlsx --base data.json --out data.json
python3 convert.py to-sql data.json --out supabase_seed.sql
```

**Moving to your real data:** run `template`, fill the sheets (customers, end-clients, suppliers, employees, store items, jobs, cost lines, invoices), then run `from-excel`. The import is checked before it is written: it fails on missing references, duplicate job numbers, bad dates or negative amounts.

---

## 7. Data model (normalised, ready for Postgres)

```
customers ─┬─< jobs >── endClients (endClients.customerIds → customers, many-to-many)
           │     ├─< plannedManpower >── employees, suppliers, assignments (1:1 link)
suppliers ─┤     ├─< actualManpower  >── employees, suppliers
           │     ├─< assignments     >── employees          (roster)
employees ─┘     ├─< plannedMaterials / actualMaterials >── suppliers
   └─< leaves    ├─< inventoryTransfers >── inventoryItems ─< stockReceipts
                 ├─< plannedMisc / actualMisc ─< miscPayments
                 ├─< invoices ─< invoicePayments
                 ├─< jobStatusHistory, poRevisions, jobComments, jobAttachments, closures
roles ─< users                      auditLog (append-only)        app_settings (single row)
```
Every business table has `createdAt`, `createdBy`, `updatedAt`, `version` and `archived`. `convert.py to-sql` writes the DDL: tables, primary keys, foreign keys (`on delete restrict`), indexes, a unique `job_no`, a unique transfer `ref`, a unique `invoice_no`, and check constraints. It also writes seed inserts. The full column list is in `TABLES` at the top of `convert.py`.

## 8. Supabase phase plan

1. Run `to-sql` → apply it as migration `0001_init.sql` in Supabase. Map `users` to `auth.users`.
2. Move the `API.*` functions into Postgres RPC functions, one per action (`create_job`, `save_cost_line`, `close_job`, `set_invoice_status`, `save_assignment` …). Each function checks permissions, locks, versions and conflicts, then writes the audit row in the same transaction.
3. RLS: read policies use the role's `jobTypes` and permissions. Financial columns are exposed only through views gated by `viewFinancials`, and rates only by `viewCostRates`. Direct table writes are denied, so all writes go through the RPC functions. The audit table is insert-only.
4. Swap `Store` in `index.html` for the Supabase client, and keep the UI. Add realtime subscriptions to replace the cross-tab sync. Optionally port the UI to React + TypeScript, keeping the same screens.
5. Port the 15 acceptance tests to pgTAP and Playwright.
6. Backups: enable Supabase PITR (point-in-time recovery). Add a nightly `pg_dump` to separate storage, and the existing JSON export as a third copy. Test a restore each quarter.

**Zoho Books (later):** OAuth 2.0 tokens stay in an Edge Function and never reach the browser. Data comes *into* the dedicated ID fields (`zohoCustomerId`, `zohoSoId`, `zohoInvoiceId`, `zohoPaymentId`). The unique invoice number blocks duplicate imports. A sync-log table and a manual reconciliation screen handle mismatches. Job costing and GP stay in JobOps.
