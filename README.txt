
**Move the working folder to somewhere better before you start** —
a dedicated `Working` folder next to the exe works fine. The tool
remembers wherever you put it after that.

### Keeping Excel and Outlook closed

Close Excel and Outlook before starting any run. The tools open their own
hidden Excel instances. A leftover "Save changes?" dialog will hang the
tool.

### Open Logs Folder

The launcher has an **Open Logs Folder** button. If anything goes wrong,
click it first — see [Logs and Troubleshooting](#logs-and-troubleshooting).

---

## The Four Tools

### 1. Packing List

**Purpose:** Read a filled-in packing list Excel file, search Outlook for
matching RITM emails, assign items to RITMs and locations, and generate
the output workbooks plus per-load folders.

**Basic use — step by step:**

1. Click **PackingList** on the launcher.
2. In the window, pick the folder the application will read files from
(this becomes the saved "last import folder").
3. Click **Open Input Sheet**. Fill out the sheet in Excel and save it.
- The first time, the copied template appears under
  `…\PackingList\_internal\`. **Move it somewhere stable** before
  filling it in.
4. Click **Rename Files** to rename files in the folder based on their
packing-list numbers. This makes it obvious which load each file
belongs to.
5. Click **Import Input Sheet** to load the information.
6. Wait until it says imported. (These steps read and scan Excel, so they
may take a moment.) Once the text above **Show Manual Entry** updates,
click **Scan Item List**.
- **Show Manual Entry** is for typing items by hand. It exists because
  the very first version of this tool worked that way, but filling
  the Excel sheet is easier and saves better.
7. A popup asks you to confirm **near matches**. See
[Near-match guidance](#near-match-guidance) below.
8. If any RITM or Address needs to be merged into an existing load or
pushed into a new one, enter it in the next popup.
- **Rescan for this RITM** re-searches Outlook for a RITM you typed
  by hand and pulls the PDF the same way a real match would.
9. On the results screen with both bottom boxes filled in, **double-check
the split**. If it looks wrong, or you need to wait for more
information, **this is the point where you can bail before anything
is written**. If it looks right, continue.
10. You can now click any email in the list (rows start with RITM, then
 Sender, etc.) and click **Reply to Selected Email**. The tool
 pre-fills a reply email with every item found on that load.
11. Click **Database…** to open the SQLite file. The file lives at:
12. Once you have sent all the emails, click **Export Folder…** and choose
the destination. The app creates:
- **Packing List** workbooks — one per load (A, B, C all included;
  even single-load FRNs get a workbook). Safe to delete after
  printing.
- **Summaries** folder containing:
  - `Reel Summary Report WE <date range>.xls`
  - If the packing list falls inside an existing summary's date range,
    it is added to that file instead of creating a new one.
- **`FRN…/` folder** per load containing:
  - The PDF pulled from the matching email (`Delivery FRN…`)
  - `Scrap Pick Up Request for Load FRN….xls`
13. **[OPTIONAL]** Once you have printed everything and moved it to your
scanned-docs folder, you can delete everything **except the
Summaries folder and its contents**. Those files are required by
Split Material.

#### Near-match guidance

When the scan shows a popup for a near miss, the general rule is: **if
the only difference is one digit and the rest of the description is
identical, treat it as the same**.

Common digit swaps that are safe to accept:

| Swap | Verdict |
|---|---|
| 7 ↔ 1 | same |
| 3 ↔ 8 | same |
| 11 ↔ 4 | same |
| 2 ↔ 3 | same |

A number that is one digit short, with everything else matching exactly,
is also the same. If in doubt, accept it; if you reject it, the item
stays in the unmatched pool and you'll have to assign it manually.

#### The database (optional)

The tool can log every item it exports and warn you if an item is already
on file under a different FRN — useful when the same reel shows up on
multiple loads by mistake.

- Set it via **Database…** (see step 11 above).
- **Repeat load** — replaces the current FRN's rows (all letter variants)
when you re-export.
- **Ignore database entirely** — no reads or writes.

Leave it unset and the tool works exactly the same, just without
duplicate protection.

---

### 2. Split Material

**Purpose:** Take a Reel Summary Report with a family of loads
(`FRN…-A`, `FRN…-B`, …) and split the material weights across those loads
so every material's grand total matches its budget.

**When to use it:** When the total weights for each material come in.
You need the **Summaries folder** (or its contents) from Packing List;
everything else can be discarded.

**Basic use:**

1. Point **Summary folder** at the folder holding your Reel Summary
Reports.
2. Leave **Split Material files go next to the Summary Report** checked
unless you want them in a separate folder.
3. Type the FRN (e.g. `FRN261008-2R`) and click **Load Family**.
4. Review the per-load tabs. The **Class** column is editable — correct
mis-typed materials here before running.
5. Switch to the **Materials** tab and enter the total weight for each
material (from Verizon). Leave a row blank to mark a material as
known-but-zero.
6. Set the **Step** (usually `1`) and **Min weight** (default `30`).
7. Click **▶ Run Solver && Write Files**.

**What it writes:**

- `Split Materials/Split Material Weighted <FRN>.xls` — one per load
- Updates the P tabs and the Summary tab in the Reel Summary Report
- If you edited Class cells, writes the corrections back into the R tabs

**Reading the solver output:** The tool prints a `=== SOLUTION ===`
block (every load with its material breakdown) and a `=== CHECK ===`
block (each material total vs. its budget).

**About "no solution" cases:**

If there is no solution that satisfies every load *and* every material
budget at the same time, the tool tries to **relax one load** (then
pairs, if needed) and prints a `NOTE:` line telling you which load it
relaxed.

If a valid answer truly does not exist, the tool does something like a
"last one takes the remainder" fallback:

- **Example:** loads A, B, C. A and B get sensible values. C just gets
whatever is left over. It will look wrong if you do the math, but this
**only happens when the numbers don't add up in the first place** —
for example, too much 99FO for it to fit.

Treat a relaxed run as a warning that something on the Verizon side or
on the Materials tab needs a second look.

---

### 3. Fill Payment Statement

**Purpose:** Collect every load for a week (from the JC statement, from
the Reel Summary Reports, and from loose files) and fill the Verizon BAU
Payment Statement twice — merged with JC's rows (email this) and
Rahway-only (print this).

**Basic use:**

1. Fill in the seven fields at the top of the window:
1. **JC Payment Statement** — its filename must contain
  `WE MM-DD-YYYY`; the tool reads that to determine the week.
2. **Receiving root for Reel Summaries** — the folder containing
  `YYYY/Month YYYY/Reel Summary/`.
3. **Receiving root for loose files** — the folder containing
  `YYYY/Month YYYY/`.
4. **Blank template** — the empty Verizon BAU statement.
5. **Scrap Pickup root** — the folder tree containing the
  `Scrap Pick Up Request for Load…` files.
6. **Export folder for MERGED file**.
7. **Export folder for RAHWAY-ONLY file**.
2. Click **Collect & sort…**
3. In the sort dialog: JC rows appear first; reel and loose rows follow
interleaved by date. **Drag rows to reorder.** The **Date** column is
editable. **Ctrl+R** re-sorts reels + loose by date (JC stays on top).
4. Click **Write both files**.

**What it writes:**

- `Verizon BAU Payment Statement WE <date> (merged).xls`
- `Verizon BAU NJ Payment Statement WE <date>.xls`
- Two sidecar audit files (`…_loads.log`) listing every load that went in

**Notes:**

- Tare for the statement is read from the per-FRN tab (`FRN…R`) via
Total Gross / Total Tare. Don't remove those tabs before running.
- Loose `.xls` files are matched by `FRN…-…L` in the filename; anything
else is skipped. The log lists every file considered.
- If an FRN lands in the wrong week, its encoded date is different from
what you expected — check the FRN.

---

### 4. PDF Merge

**Purpose:** After exporting, append scanned delivery slips to the
matching `Request Item for Load <FRN>.pdf` and file the scan inside the
load folder.

**Basic use:**

1. Put all scanned PDFs into one folder. Each filename must contain the
FRN identifier, e.g. `scan FRN261021-2R-A.pdf`.
2. Click **Choose Scans Folder…**.
3. Click **Choose Export Root…** and pick the folder that holds your
`FRN…/` load folders.
4. Click **Preview Matches**. Statuses show:
- **Ready** — will merge
- **No FRN found in filename** — rename the scan
- **No folder <FRN> in export root** — not exported yet, or letter
 missing
- **Missing Request Item for Load <FRN>.pdf** — no source PDF for
 that load (no RITM email matched during Packing List)
5. Optionally check **Overwrite an existing 'Scanned Delivery'** for
re-scanned loads.
6. Click **Merge & File Scans**.

**What it does for each Ready scan:**

- Appends the scan's pages after the pages of the Request Item PDF
(atomic — a failure leaves the original untouched)
- Moves the scan into the load folder as `Scanned Delivery <FRN>.pdf`
(or `… (2).pdf` if that name exists)

**Notes:**

- If a scan's filename lacks the trailing letter (`-A`/`-B`), the tool
looks for `FRN…R/`, not `FRN…R-A/`. Add the letter to fix.
- The scan stays in the load folder after merging for audit. Safe to
re-run.

---

## What You Still Need To Do By Hand

The tool creates **ready-to-print files**. It does **not** print them.

Still on you:

1. **Scan scale tickets.**
2. **Scan the other pages** that are needed.
3. **Print** the following (all made by the program):
- `Packing List FRN….xls` — safe to delete once printed
- Everything inside each `FRN…/` folder:
 - `Request Item for Load FRN….pdf`
 - `Scrap Pick Up Request for Load FRN….xls`
- The **Split Materials** workbooks

---

## The Database

- The database is **SQLite**, not a server database.
- It lives at: