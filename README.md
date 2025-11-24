# POC-Simulation
A workflow that automates presales company matching and data enrichment. It uses scoring logic, confidence flags, and optional Python automation to identify the best company match per group and enrich each record with domains, industry tags, location, and online presence.
# 📊 Presales Data Matching & Company Enrichment

Automates presales company matching and data enrichment using scoring logic, confidence flags, and optional Python automation to select the best match per group and enrich each record with domains, industry, location, and online presence.

---

## 💡 Overview

This project streamlines the **presales research workflow** by turning a messy list of company records into a **clean, enriched, and quality-checked dataset**.

It combines:

- Excel-based logic (scoring, flags, filtering)
- Optional Python automation for enrichment
- A clear separation between **raw data**, **matching logic**, and **final outputs**

The main use case is:  
> “I have many possible company variants per group and I need to choose the best match, enrich it with reliable info, and quickly spot rows that need manual review.”

---

## 🔍 Core Features

### 1. Group-Based Company Matching

- Each company belongs to a **group** (e.g., 5 variants per group ID).
- A **Total Score** column ranks candidates within each group.
- Flags identify the outcome:
  - `Winner_flag`: `Matched` or `Review`
  - `tie_flag`: `Tie` or empty
  - `low_confidence`: `Poor` or empty

This allows you to automatically pick the most likely match per group.

---

### 2. Confidence & Review Workflow

The model separates rows into three categories:

- ✅ **High-confidence matches**  
  Automatically selected as the “winner” for each group.

- ⚖️ **Ties**  
  Multiple rows with similar scores. These go to a **Review** sheet for manual decision.

- 🛠 **Low-confidence / Poor matches**  
  Flagged as `Poor` so you can re-check them online or with additional data sources.

---

### 3. Data Enrichment

For each matched company, the workflow is designed to enrich with:

- Website URL, domain, TLD
- Basic description and business tags
- Country, region, city, postal code
- Industry classifications (NAICS / SIC / ISIC / NACE if available)
- Phone numbers and emails (where present)
- Social media links (LinkedIn, Facebook, etc.)
- Optional: automated checks to verify online presence matches the spreadsheet record

The goal is to reduce manual lookups while keeping a clear QC process.

---

## 🗂 Data & Files

### `data/`
- `presales_data_sample.xlsx`  
  Raw sample data you’re working with. Typically includes:
  - Group ID
  - Company name
  - Country / location fields
  - Scoring and flags (or space for them)


### `excel/` 
- `presales_matching_model.xlsx`  
  Main Excel file where formulas live:
  - Scoring logic
  - Conditional formatting
  - Summary sheets (winners, ties, poor matches)

### Top-level
- `README.md` – This documentation
- `requirements.txt` – Python dependencies 
- `.gitignore` – To ignore temporary files, virtual env, etc.

---

## 🚀 How to Use

### Option A — Excel-Only Workflow

1. **Open** `presales_matching_model.xlsx` (or your current Excel file).
2. Ensure it links or contains the raw data from `data/presales_data_sample.xlsx`.
3. Use formulas to calculate:
   - `Total Score` per row
   - `Winner_flag`
   - `tie_flag`
   - `low_confidence`
4. Create separate sheets:
   - `Matches_selected` → only high-confidence winners
   - `Review_needed` → rows where `tie_flag` or `low_confidence` is set
5. Export `Matches_selected` as your **final client-ready file**.

---

### Option B — With Python Automation

1. Create and activate a virtual environment (optional but recommended):

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
