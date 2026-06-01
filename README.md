# SA Edit vs Original — KPI Analysis

A browser-based dashboard for analysing how Sales Associates (SAs) edit AI-generated outreach messages in the AI360 system. No installation or server required — open the HTML file and run everything locally.

---

## How to Use

### 1. Open the tool

Open `index.html` in any modern browser (Chrome or Edge recommended). No setup needed.

You will land on the **home page**, which shows a list of saved monthly reports. On first use, it will be empty.

---

### 2. Run a new analysis

Click **+ New Analysis** in the top right.

On the upload screen:
- **Drop or select** your raw export file (`.xlsx` format, must contain a `Sheet1` with a `copy_to_chat_history` column)
- **Select the month and year** the data belongs to
- Click **Run Analysis**

---

### 3. Processing steps

The tool runs 4 steps automatically:

| Step | What happens |
|---|---|
| **1 — Filter Rows** | Scans all rows and identifies those where `is_change = true`. Displays the count (e.g. *181 edited rows out of 1,485 total*) before proceeding. |
| **2 — Translate** | Translates Chinese message text to English via Google Translate (for readable column output only — no KPI uses the translation). |
| **3 — Diffs & Summaries** | Computes character-level diffs between each AI message and sent message. Generates change summaries in Chinese. |
| **4 — Compute KPIs** | Aggregates all metrics across overall and per-use-case views. |

Once complete, the dashboard opens automatically and the report is saved to browser storage.

---

### 4. Reading the dashboard

The dashboard has four tabs on the left sidebar:

- **Overall** — all use cases combined
- **Emotional Bonding**
- **Task**
- **Product Storytelling**

The top banner shows key headline numbers: total messages generated, total edited, overall edit rate, and number of SAs.

KPI sections are listed in order:

1. **Edit Rate** — how many of all AI-generated messages were edited by SAs
2. **Location of Change** — where in the message edits occurred (opening / middle / closing)
3. **Length Change** — whether SAs shortened or expanded the message
4. **Content Type** — what categories of content were modified (greeting, CTA, product references, emoji)
5. **SA Distribution** — how edit rates are distributed across individual SAs

Click **Show evidence** next to any KPI row to expand a table of the individual rows or SAs that contributed to that number. Removed text is shown in red, added text in green.

> For the exact thresholds and definitions behind each KPI, see [KPI_REFERENCE.md](KPI_REFERENCE.md).

---

### 5. Returning to saved reports

Click the **AI360 SA Edits KPI** title at the top of any page to return to the home screen. All previously analysed months are stored there. Click any month card to reload its dashboard instantly.

---

### 6. Downloading results

At the bottom of the dashboard:

- **SA_edits_coloured.xlsx** — the filtered edited rows with original and translated message columns
- **KPI_Summary.xlsx** — the full KPI table exported as a spreadsheet

---

## Files

| File | Purpose |
|---|---|
| `index.html` | The entire tool — open this in a browser |
| `KPI_REFERENCE.md` | Full KPI definitions and thresholds |

## Notes

- All data stays in your browser (`localStorage`). Nothing is uploaded to any server.
- Storage is limited by the browser's localStorage quota (~5MB). Delete older reports from the home screen if needed.
- Translation uses Google's public endpoint and requires an internet connection for Step 2 only.
