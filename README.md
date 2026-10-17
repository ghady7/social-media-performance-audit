# Multi-Channel Paid Social Performance Audit & Data Pipeline

* **Repository File:** `README.md` / `DATA_CLEANING_DOCUMENTATION.md`
* **Associated Workbook:** `Social Media Performance.xlsx`
* **Author:** Ghady Tayeh
* **Dataset Grain:** 10,000 Initial Records to 9,917 Validated Records
* **Architecture:** `raw_ad_data` | `cleaned_ad_data` | `Executive_Summary`

---

## 1. Project Overview & Architecture

### Business Objective
Audit, cleanse, enrich, and evaluate paid digital advertising campaign records across **Meta**, **TikTok**, and **LinkedIn** to identify marketing budget leaks and measure cross-platform Return on Ad Spend (ROAS).

### Worksheet Isolation Architecture
The workbook utilizes three isolated layers to guarantee complete data lineage and analytical auditability:
1. **`raw_ad_data`**: Preserved as an immutable source layer containing unmodified transactions and logging errors.
2. **`cleaned_ad_data`**: Engineered as an official Excel Table (`tbl_SocialAds`) with standardized categorical dimensions, imputed attributes, audit flags, and calculated KPI columns.
3. **`Executive_Summary`**: Contains dynamic Pivot Tables and an executive action plan summarizing channel efficiency, audience targeting metrics, and budget reallocation directives.

---

## 2. Summary of Issues Encountered & How They Were Resolved

Before running aggregations, a forensic audit of `raw_ad_data` uncovered multiple systemic data defects. Below is a high-level summary of every issue encountered and the exact resolution implemented:

| Issue Category | What Was Found in `raw_ad_data` (The Problem) | Business Impact | How It Was Passed / Resolved (The Solution) |
| :--- | :--- | :--- | :--- |
| **Record Duplication** | 83 exact duplicate campaign rows across all columns. | Artificially overstated total ad spend by ~\$43,000 and impressions by ~2.9 million. | Converted range to an Excel Table (`tbl_SocialAds`) and ran native deduplication (`Alt + A + M`), reducing records from 10,000 to 9,917. |
| **Mixed Date Formats** | 9,292 numeric Excel serial numbers (e.g., `45701`), 509 text dates (`DD-MM-YYYY`), and 199 blank cells. | Prevented chronological sorting and native date grouping in Pivot Tables. | Engineered a hybrid parsing formula using `ISNUMBER`, `INT`, `DATE`, `LEFT`, `MID`, and `RIGHT` to unify all dates into `YYYY-MM-DD`. |
| **Platform Typos & Trailing Spaces** | 9 dirty text variants (`'meta'`, `'META '`, `'Meta'`, `'tik_tok'`, `'TikTok'`, `'tiktok'`, `'LinkedIn'`, `'linkedin'`, `'LINKEDIN'`). | Split platform performance across separate rows in Pivot Tables. | Combined `TRIM`, `PROPER`, `IF`, and `SEARCH` to strip whitespace and consolidate into `Meta`, `TikTok`, and `LinkedIn`. |
| **Objective Casing Inconsistency** | 5 mixed-casing variants (`'BRAND TRAFFIC'`, `'Brand Traffic'`, `'Conversions'`, `'conversions'`, `'Lead Generation'`). | Fragmented campaign objective reporting. | Applied `IFS` + `SEARCH` substring matching to standardize into 3 clean objectives and derived macro `funnel_stage` (`Top of Funnel` vs. `BOFU`). |
| **Audience Semantic Inversion** | Reversed naming conventions (`'BOFU_Retargeting'` vs. `'Retargeting_BOFU'`) and casing issues (`'prospecting_cold'` vs. `'Prospecting_Cold'`). | Split identical audience segments into separate buckets. | Built an `IFS` + `OR` + `SEARCH` keyword matcher to merge variations into `Retargeting`, `Cold Prospecting`, and `Lookalike 1%`. |
| **Ad Format Variations** | Inconsistent naming across static images, video reels, and carousels. | Prevented clean creative format comparisons. | Standardized all creative labels into `Carousel`, `Static Image`, and `Video Reel` using `IFS` and `SEARCH`. |
| **Missing Device & Region Data** | 2,029 rows (20.3%) missing `device` and 2,021 rows (20.2%) missing `region`. | Deleting these rows would have erased ~20% of valid financial spend and revenue data. | Imputed blanks as `"Unknown"` while standardizing valid entries (`North America`, `APAC`, `EMEA`), preserving 100% of financial metrics. |
| **Negative & Blank Spend** | 95 negative spend values (e.g., `-$89.73`) caused by sign-entry errors, plus 93 empty records. | Understated total campaign investment and distorted ROAS. | Wrapped spend in `ABS()` to correct sign inversions and mapped blank cells to `0`. |
| **Tracking Outages & Ghost Clicks** | 90 rows with `-100` impressions (sentinel error code) and 43 rows with `0` impressions but positive clicks. | Corrupted CTR calculations and skewed impression totals. | Floored negative impressions to `0` and created a `data_quality_flag` column (`Tracking Outage`, `Ghost Clicks`, `Missing Date`, `Clean`) for full auditability. |
| **Anomalous Brand Traffic Revenue** | 83 `Brand Traffic` records logged \$6,888.45 in unexpected revenue despite zero leads/purchases. | Created misleading top-of-funnel attribution noise. | Isolated in Pivot Table analysis to expose the 0.00x effective ROAS and justify budget reallocation. |

---

## 3. Step-by-Step Data Cleaning & Pipeline Implementation

### Step 1: Work Environment Isolation & Table Conversion
* **Action:** Duplicated `raw_ad_data` into a new worksheet named `cleaned_ad_data` to ensure the raw transaction log remained untouched.
* **Action:** Converted the active data range into an official Excel Table (`Ctrl + T`) named `tbl_SocialAds` to enable structured column referencing and automatic formula propagation.

---

### Step 2: Record Deduplication
* **Issue:** 83 exact duplicate records inflating spend and delivery metrics.
* **Resolution:** Deployed native deduplication (`Alt + A + M`) across all columns in `tbl_SocialAds`, leaving **9,917 validated, unique records**.

---

### Step 3: Categorical Normalization & Text Parsing

Applied structured formula transformations to clean and standardize text fields across `tbl_SocialAds`:

#### 1. Platform Standardization (`platform`)
* **Issue:** 9 casing, underscore, and trailing-space variants across Meta, TikTok, and LinkedIn.
* **Formula:**
```excel
=TRIM(PROPER(IF(ISNUMBER(SEARCH("tik", [@platform])), "TikTok", 
 IF(ISNUMBER(SEARCH("link", [@platform])), "LinkedIn", 
 IF(ISNUMBER(SEARCH("meta", [@platform])), "Meta", [@platform])))))
```
* **Result:** Consolidated all 9 variations and stripped trailing whitespaces into `Meta`, `TikTok`, and `LinkedIn`.

#### 2. Objective Normalization (`objective`)
* **Issue:** 5 mixed-casing states splitting campaign objectives.
* **Formula:**
```excel
=IFS(
    ISNUMBER(SEARCH("brand", [@objective])), "Brand Traffic",
    ISNUMBER(SEARCH("lead", [@objective])), "Lead Generation",
    ISNUMBER(SEARCH("conv", [@objective])), "Conversions",
    TRUE, "Other"
)
```
* **Result:** Standardized casing variants into three distinct marketing objectives.

#### 3. Funnel Stage Derivation (`funnel_stage`)
* **Objective:** Group campaign objectives into macro-funnel stages for high-level performance comparison.
* **Formula:**
```excel
=IF([@objective]="Brand Traffic", "Top of Funnel", "BOFU")
```

#### 4. Audience Type Inversion Resolution (`audience_type`)
* **Issue:** Reversed word order (`BOFU_Retargeting` vs. `Retargeting_BOFU`) and inconsistent casing (`prospecting_cold` vs. `Prospecting_Cold`).
* **Formula:**
```excel
=IFS(
    OR(ISNUMBER(SEARCH("bofu", [@audience_type])), ISNUMBER(SEARCH("retargeting", [@audience_type]))), "Retargeting",
    ISNUMBER(SEARCH("lookalike", [@audience_type])), "Lookalike 1%",
    OR(ISNUMBER(SEARCH("prospecting", [@audience_type])), ISNUMBER(SEARCH("cold", [@audience_type]))), "Cold Prospecting",
    TRUE, "Other"
)
```
* **Result:** Merged reversed text values into unified targeting dimensions: `Retargeting`, `Cold Prospecting`, and `Lookalike 1%`.

#### 5. Ad Format Standardization (`ad_format`)
* **Issue:** Inconsistent naming conventions across ad creative formats.
* **Formula:**
```excel
=IFS(
    ISNUMBER(SEARCH("carousel", [@ad_format])), "Carousel",
    OR(ISNUMBER(SEARCH("static", [@ad_format])), ISNUMBER(SEARCH("image", [@ad_format]))), "Static Image",
    OR(ISNUMBER(SEARCH("video", [@ad_format])), ISNUMBER(SEARCH("reel", [@ad_format]))), "Video Reel",
    TRUE, "Other"
)
```
* **Result:** Unified ad format definitions into `Static Image`, `Video Reel`, and `Carousel`.

#### 6. Device & Region Missing Value Imputation (`device`, `region`)
* **Issue:** 2,029 records (20.3%) lacked device classifications and 2,021 records (20.2%) lacked regional identifiers.
* **Device Formula:**
```excel
=IF(OR(ISBLANK([@device]), [@device]=""), "Unknown", PROPER(TRIM([@device])))
```
* **Region Formula:**
```excel
=IF(OR(ISBLANK([@region]), [@region]=""), "Unknown", 
    IFS(
        ISNUMBER(SEARCH("north", [@region])), "North America",
        ISNUMBER(SEARCH("apac", [@region])), "APAC",
        ISNUMBER(SEARCH("emea", [@region])), "EMEA",
        TRUE, "Unknown"
    )
)
```
* **Result:** Imputed missing values to `"Unknown"` while cleaning valid text, avoiding the deletion of rows that contained valid financial performance data.

---

### Step 4: Hybrid Date Normalization
* **Issue:** The `date` field was fractured into 9,292 numeric Excel serial values (e.g., `45701`), 509 string-formatted dates (`DD-MM-YYYY`), and 199 blank cells.
* **Formula:**
```excel
=IF(OR(ISBLANK([@date]), [@date]=""), "", 
    IF(ISNUMBER([@date]), 
        INT([@date]), 
        DATE(RIGHT([@date], 4), MID([@date], 4, 2), LEFT([@date], 2))
    )
)
```
* **Result:**
  * Preserved existing serial integers using `INT()`.
  * Parsed string dates (`DD-MM-YYYY`) using `RIGHT()`, `MID()`, and `LEFT()`, and rebuilt valid serial dates using `DATE()`.
  * Formatted the entire column to the ISO standard date format (`YYYY-MM-DD`).

---

### Step 5: Numerical Sanitization & Data Quality Auditing

#### 1. Spend Correction (`spend`)
* **Issue:** 95 negative values (e.g., `-$89.73`) due to sign errors and 93 empty records.
* **Formula:**
```excel
=IF(OR(ISBLANK([@spend]), [@spend]=""), 0, ABS([@spend]))
```
* **Result:** Converted negative spend values to positive numbers using `ABS()` and defaulted blank records to `0`.

#### 2. Impression Correction (`impressions`)
* **Issue:** 90 tracking outage sentinel codes (`-100`).
* **Formula:**
```excel
=IF([@impressions] < 0, 0, [@impressions])
```
* **Result:** Replaced negative sentinel error codes (`-100`) with `0`.

#### 3. Data Quality Flag Creation (`data_quality_flag`)
* **Issue:** Needed a way to track corrupt telemetry (90 tracking outages, 43 ghost-click instances where `impressions = 0` but `clicks > 0`, and 199 missing dates) without silently deleting financial records.
* **Formula:**
```excel
=IFS(
    [@impressions] < 0, "Tracking Outage",
    AND([@impressions]=0, [@clicks]>0), "Ghost Clicks",
    ISBLANK([@date]), "Missing Date",
    TRUE, "Clean"
)
```
* **Result:** Maintained full audit transparency by tagging anomalous rows so analysts can filter between clean and corrupted telemetry on demand.

---

## 4. Calculated Performance Metrics (KPIs)

Constructed standard advertising efficiency KPIs across `tbl_SocialAds` with zero-division error handling:

| Metric | Excel Formula | Display Format |
| :--- | :--- | :--- |
| **Click-Through Rate (CTR)** | `=IF([@impressions] > 0, [@clicks] / [@impressions], 0)` | Percentage (`0.00%`) |
| **Cost Per Click (CPC)** | `=IF([@clicks] > 0, [@spend] / [@clicks], 0)` | Currency (`$0.00`) |
| **Cost Per Acquisition (CPA)** | `=IF([@purchases] > 0, [@spend] / [@purchases], 0)` | Currency (`$0.00`) |
| **Cost Per Lead (CPL)** | `=IF([@leads] > 0, [@spend] / [@leads], 0)` | Currency (`$0.00`) |
| **Return On Ad Spend (ROAS)** | `=IF([@spend] > 0, [@revenue] / [@spend], 0)` | Decimal with suffix (`0.00x`) |
| **Average Order Value (AOV)** | `=IF([@purchases] > 0, [@revenue] / [@purchases], 0)` | Currency (`$0.00`) |

---

## 5. Performance Insights & Pivot Table Analysis

The cleaned dataset was summarized on the `Executive_Summary` worksheet across **\$5,100,760** in total ad spend and **\$15,903,933** in gross revenue.

### Funnel Objective Breakdown
| Objective | Spend | Impressions | Clicks | Leads | Purchases | Revenue | ROAS |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Brand Traffic** | \$1,983,099 | 136,160,445 | 2,817,379 | 0 | 0 | \$6,888 | 0.00x |
| **Conversions** | \$2,059,022 | 140,617,110 | 2,876,859 | 0 | 97,421 | \$7,654,858 | 3.72x |
| **Lead Generation** | \$1,058,639 | 71,607,765 | 1,482,226 | 118,058 | 0 | \$8,242,187 | 7.79x |
| **Grand Total** | **\$5,100,760** | **348,385,320** | **7,176,464** | **118,058** | **97,421** | **\$15,903,933** | **3.12x** |

### Platform ROAS by Objective
| Platform | Brand Traffic | Conversions | Lead Generation | Total ROAS |
| :--- | :---: | :---: | :---: | :---: |
| **LinkedIn** | 0.00x | 3.71x | 8.14x | **3.24x** |
| **Meta** | 0.00x | 3.68x | 7.42x | **3.05x** |
| **TikTok** | 0.00x | 3.77x | 7.80x | **3.06x** |

### Channel Volume & Conversion Delivery
| Platform | Total Clicks | Total Purchases |
| :--- | :---: | :---: |
| **LinkedIn** | 2,409,607 | 32,347 |
| **Meta** | 2,388,523 | 32,322 |
| **TikTok** | 2,378,334 | 32,752 |
| **Grand Total** | **7,176,464** | **97,421** |

### Audience Targeting Efficiency
| Audience Segment | Spend | Revenue | ROAS |
| :--- | :---: | :---: | :---: |
| **Cold Prospecting** | \$2,025,739 | \$6,416,639 | **3.17x** |
| **Retargeting** | \$2,068,876 | \$6,484,130 | **3.13x** |
| **Lookalike 1%** | \$1,006,145 | \$3,003,164 | **2.98x** |

---

## 6. Strategic Recommendations & Action Plan

1. **Eliminate Brand Traffic Capital Waste:**
   * **Finding:** `Brand Traffic` consumed **\$1,983,099** (38.9% of total marketing budget) while producing zero qualified leads, zero direct purchases, and negligible revenue (\$6,888, resulting in a `0.00x` ROAS).
   * **Action:** Immediately reallocate **\$1,000,000** from `Brand Traffic` directly into `Lead Generation` campaigns.
2. **Scale High-Return Lead Generation:**
   * **Finding:** `Lead Generation` delivered **\$8,242,187** in revenue from **\$1,058,639** in spend, achieving a **7.79x ROAS** at an **\$8.97 CPL**. LinkedIn is the strongest channel, driving an **8.14x ROAS**.
   * **Action:** Prioritize LinkedIn Lead Generation for additional budget allocations.
3. **Prioritize TikTok for Direct Conversions:**
   * **Finding:** TikTok produced the highest `Conversions` ROAS at **3.77x**, outperforming both LinkedIn (3.71x) and Meta (3.68x), while delivering the highest total purchase volume (32,752 purchases).
   * **Action:** Weight e-commerce conversion ad spend toward TikTok.
4. **Rebalance Audience Targeting Mix:**
   * **Finding:** `Cold Prospecting` (3.17x ROAS) and `Retargeting` (3.13x ROAS) both outperformed `Lookalike 1%` targeting (2.98x ROAS), indicating lookalike audience saturation.
   * **Action:** Cap `Lookalike 1%` spend and reallocate budget to scale top-of-funnel `Cold Prospecting` acquisition.