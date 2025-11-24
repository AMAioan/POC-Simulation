
# 📊 Presales Data Matching & Company Enrichment

Automates presales company matching and data enrichment using scoring logic, confidence flags, and optional Python automation to select the best match per group and enrich each record with domains, industry, location, and online presence.

---

## 💡 Overview

This project streamlines the **presales research workflow** by turning a messy list of company records into a **clean, enriched, and quality-checked dataset**.

It combines:

- Excel-based logic (scoring, flags, filtering)
- Optional Python automation for enrichment
- A clear separation between:
  - **Raw data**
  - **Matching logic**
  - **Client-ready final output**

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

The enrichment process was completed using a combination of automated and manual methods.  
A custom Python script (developed with the help of ChatGPT) was used to extract and validate structured fields such as domains, website information, and standard formats. When automated methods were insufficient or information was missing, we supplemented the process with manual research on Google and other search engines.

This hybrid approach allowed us to:
- Automatically parse website domains, TLDs, and contact patterns
- Clean and standardize address fields when possible
- Identify missing or inconsistent country, city, or postcode values
- Validate whether a company’s online presence matched the dataset
- Manually fill gaps or confirm uncertain matches through targeted online searches

The result is a thoroughly enriched dataset that balances automation efficiency with human validation, ensuring accuracy for the final vendor-ready file.

---

## 🗂 Data & Files

### `data/`

- `presales_data_sample.xlsx`  
  Raw sample data used during development and testing. Typically includes:
  - Group ID
  - Company name
  - Country / location fields

- `ppresales_data_vendor.xlsx`  
  **Clean, client-ready file** generated after:
  - Scoring and matching are applied
  - Ties and poor matches are reviewed
  - Enrichment fields are completed
This is the version you would deliver to the client or use in downstream systems.

- `presales_matching_model.xlsx`  
  Main Excel file where formulas live:
  - Scoring logic
  - Conditional formatting
  - Summary sheets (winners, ties, poor matches)

### `src/`

- `presales_enrichment.py` 

 After reviewing the 592 companies, we used a combination of Wikidata, website scraping, and Google Places APIs to gather additional information and populate missing fields. All retrieved data was manually reviewed and corrected wherever inconsistencies were found.



### Top-level
- `README.md` – This documentation
- `requirements.txt` – Python dependencies 

---

## 🚀 How to Use

### Excel Workflow

1. **Open** `presales_matching_model.xlsx`
2. `!presales_data_match` represents the initial filtering stage. Here, we compare the original company information (Columns C–I) against all candidate fields to narrow down the possible matches.

2.1 🔢 Scoring Columns & Match Selection

For each input record (`row_key`), the model compares it to several candidate companies and scores each one using dedicated scoring columns:

- `addr_score`  (Column CD)
  Measures how well the **address information** matches (street, city, postcode).  
  Higher score = closer address match and fewer conflicts.

- `website_domain_score`  (Column CE)
  Compares the **website/domain** of the candidate with the expected or known domain.  
  A correct, consistent domain gets a high score; missing or conflicting domains get a lower score or zero.

- `contact_score`  (Column CF)
  Captures matches on **phone numbers and emails**.  
  Matching contact details is a strong signal that two records refer to the same company.

- `country_code_score`  (Column CG)
  Rewards an exact **country code** match and penalizes mismatches.  
  In practice, this works almost like a hard filter: wrong country usually means wrong company.

- `Total Score`  (Column CH)
  Overall score for the candidate, typically a weighted sum of the components above:  

  ```text
  Total Score = addr_score
              + website_domain_score
              + contact_score
              + country_code_score


3. Use formulas to calculate:
   - `Total Score` per row
   - `Winner_flag`
   - `tie_flag`
   - `low_confidence`
4. Create separate sheets:
   - `!Matches_selected` → only high-confidence winners
   - `!QC_Matches` → contains all rows marked as TRUE for `top_in_group` (Column CE). In this sheet, we review, correct, and standardize the data—either manually or with additional formulas—so it’s fully cleaned and ready for the vendor.
6. Copy values from `!QC_Matches`and exported `!CLean_Vendor_Master` as our **final client-ready file**.

---
