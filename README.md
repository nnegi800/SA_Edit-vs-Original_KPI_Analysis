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

> For the exact thresholds and definitions behind each KPI, see [KPI_REFERENCE.md](documentation/KPI_REFERENCE.md).

---

### 5. Returning to saved reports

Click the **AI360 SA Edits KPI** title at the top of any page to return to the home screen. All previously analysed months are stored there. Click any month card to reload its dashboard instantly.

---

### 6. Downloading results

At the bottom of the dashboard:

- **SA_edits_translated.xlsx** — the filtered edited rows with both original Chinese and English translated columns (subject, description, AI message, sent message)
- **SA_edits_coloured.xlsx** — same rows with character-level diff highlighting (red = removed, green = added)
- **KPI_Summary.xlsx** — the full KPI table exported as a spreadsheet

---

## Files

| File | Purpose |
|---|---|
| `index.html` | The entire tool — open this in a browser |
| `KPI_REFERENCE.md` | Full KPI definitions and thresholds |

---

## Demo

A full walkthrough of the tool is available here:
[▶ Watch Loom demo](https://www.loom.com/share/75651eae16c9462b9fe4a243be0eea85)

---

## Notes

- All data stays in your browser (`localStorage`). Nothing is uploaded to any server.
- Storage is limited by the browser's localStorage quota (~5MB). Delete older reports from the home screen if needed.
- Translation uses Google's public endpoint and requires an internet connection for Step 2 only.

---

## Source File Structure

The tool expects a raw export file (`.xlsx`) with the following structure.

### Sheet1 — Main data

Each row represents one AI-generated outreach message. Key columns:

| Column | Description |
|---|---|
| `seller_id` | Unique identifier for the Sales Associate |
| `unify_id` | Client/customer identifier |
| `use_case` | One of three values: `ai360_action_plan_emotional_bonding`, `ai360_action_plan_task`, `ai360_action_plan_product_storytelling` |
| `created` | Timestamp the AI message was generated |
| `conversation_starter_subject` | Subject line of the AI message (Chinese) |
| `description` | Internal description of the use case (Chinese) |
| `conversation_starter_message` | The full AI-generated message text (Chinese) |
| `copy_to_chat_history` | JSON field containing the SA's edit, if any (see below) |

### `copy_to_chat_history` JSON structure

```json
{
  "latest_version": "v1",
  "v1": {
    "is_change": true,
    "send_time": "2024-11-15T10:32:00+08:00",
    "send_content": "The message the SA actually sent (Chinese)"
  }
}
```

- `is_change: true` — the SA edited the AI message before sending
- `is_change: false` — the SA sent the AI message without changes
- `send_content` — the final message sent to the client
- `send_time` — when the message was sent (timezone: +08:00)

Only rows where `is_change: true` are included in the analysis.

---

## Python Pipeline (Local / Excel Output)

For users who want to run the analysis locally and export results to Excel rather than using the browser tool.

### Setup

```bash
pip install -r python_pipeline/requirements.txt
```

### Scripts — run in this order from the `python_pipeline/` folder

| Step | Script | What it does |
|---|---|---|
| 1 | `translate_export.py` | Filters `is_change=true` rows, translates Chinese → English, exports `SA_edits_translated.xlsx` |
| 2 | `enrich_diffs.py` | Adds character-level diff highlighting (red = removed, green = added) and a `change_summary` column to the translated file |
| 3 | `rebuild_xlsx.py` | Rebuilds the file using `xlsxwriter` for full Excel compatibility (avoids XML corruption from openpyxl rich text) |
| 4 | `add_kpi_panel.py` | Reads the translated file + original file, computes all KPIs, and outputs `KPI_Summary.xlsx` |

### Adjusting thresholds

All thresholds are defined in `python_pipeline/add_kpi_panel.py`:

- **Length change threshold** — search for `25` in the `classify()` function (`>= 25` = Expanded, `<= -25` = Shortened)
- **SA distribution buckets** — search for `bucket_0`, `bucket_1_20` etc. in `_sa_distribution_kpis()`
- **Content type keywords** — `CN_GREET`, `CN_CTA`, `CN_PRODUCT`, `CN_SCARCITY`, `CN_CARE`, `CN_SEASON` lists in `classify()`
- **Location zones** — `0.3` and `0.7` thresholds in `classify()` (30% / 70% of character length)

