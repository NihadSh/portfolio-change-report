# Daily Portfolio Change Report

Compares two daily exports of an insurance portfolio and reports **what changed since
yesterday**: new sales, cancellations, and the premium gained or lost. It also cleans the usual
data-quality problems in the export before reporting.

## What it does

- **New policies**: policies in today's file that were not in yesterday's.
- **Cancellations**: policies cancelled on the latest date. The lost premium is taken from the
  original policy in yesterday's file.
- **Data cleaning**:
  - removes duplicated invoice rows;
  - applies business corrections from `corrections.csv`, such as an outdated channel code, a
    seller moved to another department, or merged usage categories. These rules live in a CSV
    that the business maintains, not hard-coded in the script.
- **Report sheets**:

| Sheet | Content |
|---|---|
| Summary | headline numbers for the day |
| Changes | row-level list of every new / cancelled policy |
| By LOB | new vs. cancelled policies and GWP per line of business |
| By channel | Casco sales, cancellations, lost and net GWP per channel and client type |
| Casco by usage | new Casco business by client type and vehicle usage |
| Seller pivot | new business per channel → department → seller |

## Quick start

```bash
pip install -r requirements.txt
python generate_sample_data.py
python portfolio_change_report.py data/portfolio_2026-11-10.xlsx data/portfolio_2026-11-11.xlsx
```

```
                        metric      value
                   Report date 2026-11-11
            Policies yesterday       3000
                Policies today       3180
                  New policies        180
                       New GWP   91298.76
                 Cancellations         25
      Lost GWP (cancellations)   11762.54
Duplicate invoice rows removed          4
              Values corrected        314
```

## corrections.csv

| field | match_field | match_value | new_value |
|---|---|---|---|
| channel | channel | KK | SME |
| department | seller | Seller 07 | Department 3 |

Every row means: *where `match_field` equals `match_value`, set `field` to `new_value`.*

## Tech

Python · pandas · XlsxWriter

> The data in this repository is randomly generated. No real company data is included.
