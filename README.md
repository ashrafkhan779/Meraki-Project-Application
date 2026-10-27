# JobOps — Job Management & Operations System (live on Supabase)

Jobs (Service / Supply / Supply-Services), costing (manpower, materials, store transfers, miscellaneous), financial summary, invoicing, closure & locking, Roster Planner, Master Data, 13 reports, Settings with roles and permissions, audit trail — running on your Supabase project with sign-in, a shared Postgres database and live updates between users.

| File | Purpose |
|---|---|
| `index.html` | The application. Host it anywhere (or open it with `python3 convert.py serve`). Your project URL and publishable key are already inside it. |
| `supabase_schema.sql` | One-time database setup — paste into the Supabase SQL Editor and run. Safe to run again. |
| `convert.py` | Helper: local server, formula tests, blank Excel import template, offline check of an Excel file. |
| `README.md` | This file. |

There is **no demo data**. The system starts empty; the first person who signs in becomes Super Admin.

---

## 1. Go live — three steps (about 5 minutes)

**Step 1 — Create the database (once).**
Open your Supabase project → **SQL Editor** → **New query** → paste the whole of `supabase_schema.sql` → **Run**. It creates all 28 tables, the security rules, the write functions and switches on live updates. If you ever run it again nothing breaks.

**Step 2 — Auth settings (once).**
Supabase → **Authentication → URL Configuration** → set **Site URL** to the address where `index.html` is hosted (for example `https://yourcompany.github.io/jobops/` or `http://localhost:8000`). This is where password-reset and confirmation links send people.
Optional: **Authentication → Providers → Email → “Confirm email”**. Leave it on and new users must click the link in their e-mail before signing in; switch it off and they can sign in immediately. Either way works with the app.

**Step 3 — Open the app and create the first account.**
Host `index.html` (any static host: GitHub Pages, Netlify, Vercel, your own web server) or run `python3 convert.py serve` and open http://localhost:8000. Click **Create account**, enter your e-mail and a password. Because the system is empty, that account is registered automatically as **Super Admin** and the default settings and roles are created.

Then go to **Settings → Company & finance** to set your company name, currency and rules.

> Do not open `index.html` by double-clicking it (file://). Sign-in works but password-reset links cannot return to the page. Use the local server or a web host.

---

## 2. How it works day to day

- **Everything saves automatically.** Each action (create job, add a cost line, record a payment, close a job …) is written to Postgres in one transaction the moment you confirm it. There is no publish, save or sync button. The green **Live** badge in the top bar shows the connection; it turns to **Saving…** while a write is in flight and to **Not saved / Offline** with a Retry button if the server cannot be reached.
- **Everyone sees changes instantly.** Supabase Realtime pushes every committed row to all signed-in browsers. A table you are looking at updates within about a second; if you have a dialog open, the refresh waits until you close it so you never lose what you are typing.
- **Two people editing the same record.** Every record carries a version. If someone saved first, your save is refused with a clear message, the latest data is reloaded, and you re-apply your change. Nothing is silently overwritten.
- **If you lose the connection,** the app says so and refuses new changes rather than pretending to save them. When the connection returns it reloads the latest data.

## 3. Users and access

1. An administrator adds the person in **Settings → Users** (name, e-mail, role). The e-mail is what links them to their sign-in.
2. The person opens the app, clicks **Create account** with that same e-mail and chooses a password.
3. They sign in. If their e-mail is not registered (or is deactivated) they see *“Not registered yet”* and can read or change nothing — this is enforced by the database rules (Row Level Security), not just by the screen.

Roles and the permission matrix are in **Settings → Roles & permissions** (eight starting roles plus custom roles). Financial values are hidden from roles without *View financial data*; employee cost rates need *View employee cost rates*; closed jobs are locked; the audit log is append-only for every user.

Besides the action permissions, every role can switch individual **tabs** and **fields** on or off: *Job tabs* (Manpower, Materials, Miscellaneous, Financial summary, Invoices, Comments & files, Audit history), *Master data tabs*, each of the *Reports*, *Settings tabs*, and *Fields* (Zoho SO #, customer order # & LPO date, end-client, engineers, planned/actual dates, completion %, PO value, costs, GP & margin, invoicing amounts, employee contact, customer/supplier contact, credit limit & payment terms). A role that has never set one of these inherits its parent tab permission, so existing roles keep working unchanged; money fields still need *View financial data*. The role editor has *Select all / Clear* buttons per group.

**Stores** is its own tab (after Master Data). Stock is kept per store — **Main Store** and **AGPS Store** by default; the list is in **Settings → Statuses & categories → Stores**, and the first store is the receiving store for GRNs. The Items & stock table shows the quantity and average cost in each store plus the totals; **Stock transfer** moves items between stores (at the issuing store's average cost, the receiving store's average is re-weighted); the **Movements** tab lists every GRN, stock transfer and issue to a job. When posting a store transfer to a job you choose the issuing store, and that store's stock and cost are used. Access is controlled by *View Stores tab* and *Manage stores*; the standard **Stores** role has the Operations tabs plus store management, and is created automatically the first time an administrator signs in after this version.

Store items carry a **Group** and **Sub-group** (chosen in the item form; the sub-group list follows the group). The groups and their sub-groups are maintained in **Settings → Statuses & categories → Store item groups & sub-groups**, one group per line as `Group: Sub-group, Sub-group`; a group still used by items cannot be removed. Stock is received from **Stores → New GRN** (always into the first store): supplier, date received, GRN # (unique), LPO #, supplier invoice / delivery note and comments once, then every store item on that GRN (item, quantity, unit cost) in one table with totals and the resulting average cost. Saving updates each item's on-hand quantity, stock value and weighted-average cost together; the item's detail view lists all its goods receipts. Actual material purchases are entered per PO from **Job → Materials → Add purchase**: vendor, purchase date, PO #, supplier invoice/reference and comments once, then every item on that PO (material, category, quantity, unit, unit price) in one table with a running total — each item becomes its own costed line.

> Security model, honestly stated: the database allows registered, active users to read and write the application tables, and refuses everyone else. The finer per-role rules (who may close a job, see cost rates, change invoice status …) are enforced by the application layer that every screen goes through. Moving those checks into database functions is the next hardening step and does not change how the app is used.

## 4. Excel export and import — Settings → Data & integrations

**Export working copy as Excel** downloads one workbook with **all data**: a sheet per table (27) plus a Settings sheet and a README sheet. Row 1 of each sheet holds the column keys.

**Upload Excel** reads the same layout back:
1. Choose the file. The app checks it — required columns, dates, numbers, every reference (`jobs.customerId` must exist in `customers`, and so on), unique job numbers, invoice numbers and transfer references.
2. You see a preview per sheet: rows in the file, new, updated, unchanged. Problems are listed; nothing is imported until they are fixed.
3. Choose **Merge** (rows with a matching `id` are updated, rows without an `id` are added, everything else is untouched) or **Replace** (the sheets in the file replace those tables completely — a reason is recorded).
4. Click **Import**. It is written in one transaction and is live for everyone immediately.

Tips: keep the `id` column to update a record; leave it blank to create one; lists such as `customerIds` are separated with `;`; dates are `YYYY-MM-DD`. **Download blank import template** gives the empty layout. `python3 convert.py validate file.xlsx` checks a file offline (needs `pip install openpyxl`).

## 5. Business rules (unchanged from the prototype)

| Topic | Rule |
|---|---|
| Rounding | Each cost line is rounded to 2 dp (half-up); totals are sums of rounded lines; display precision is configurable. |
| Manpower cost | headcount × hours per person per day × days × hourly rate. Days default to the working days between the dates (Settings → Working calendar) and can be edited per line. |
| GP | GP = PO value − total cost; GP % = GP ÷ PO value × 100, **N/A** when PO value is 0. Cancelled jobs are excluded from financial totals. |
| Actual PO value | Equals the PO value unless a revised value is entered (requires *Edit financial data* and a reason; logged). Alternative rule: total invoiced amount. |
| Completion % vs status | Independent. *Completed* only offers to set 100 %. Only **Close job** sets *Closed*. |
| Closed jobs | Locked. Only *Edit closed jobs* can change them, with a reason. Invoice status and payments stay editable for invoice roles. |
| Invoice / expense status | Derived from recorded payments; *Paid* needs the full amount. |
| Store transfers | Costed at weighted-average store cost or a confirmed manual cost; a reference posts once; reversal instead of deletion. |
| Roster ↔ job plan | The assignment is the single source of truth; its planned-manpower line is kept in sync; overbooking needs an override permission and a reason. |
| Deletion | Soft (archived). Audit log is append-only. Job numbers are never reused. |
| To bill · 3 months | Unbilled PO value grouped into completed-not-billed, next three months by planned finish, and past-PF-date. |

## 6. Testing

- **In the app:** Settings → Acceptance tests → Run all tests — the 15 workflows from the brief run on an in-memory copy; nothing is sent to the database. Current result 15/15.
- **Formulas:** `python3 convert.py test` (14 unit tests).
- **Database rules:** the schema was exercised on PostgreSQL 16 with the Supabase auth model: first-user bootstrap, unregistered user blocked, anonymous blocked, stale-version write rejected, duplicate job number rejected, audit log immutable, replace-import that would lock you out rejected.
- **End to end:** the real `index.html` was driven in a browser with the real supabase-js client against these SQL functions: sign-up, job and cost creation, conflict handling, realtime updates, a second user with a restricted role, Excel export → edit → import, and session persistence.

## 7. Operations

- **Backups:** Supabase keeps daily backups (Dashboard → Database → Backups; point-in-time recovery is available on paid plans). The Excel export and JSON backup are your own copies — take one before a *Replace* import.
- **Restore:** import the Excel in *Replace* mode, or restore a Supabase backup.
- **Changing the key or project:** the two constants `SUPABASE_URL` and `SUPABASE_KEY` are at the top of the script in `index.html`. The publishable key is safe to ship in the page; the service-role key must never be put there.
- **Schema changes later:** add columns inside the JSON `data`; add generated columns/indexes in SQL when you need to filter on them. `convert.py` holds the column specification used by the Excel layout.

## 8. Zoho Books (later)

Not connected and never shown as connected. Dedicated ID fields already exist (`zohoCustomerId`, `zohoSoId`, `zohoInvoiceId`, `zohoPaymentId`); invoice numbers are unique so imports cannot duplicate. Plan: OAuth 2.0 inside a Supabase Edge Function (secrets stay server-side) → pull customers, sales orders, invoices and payments → sync log → reconciliation screen. Job costing and GP stay in JobOps.
