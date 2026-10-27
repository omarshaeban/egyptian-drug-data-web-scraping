# Egyptian Drug Registration Data — Web Scraping

Python-based web data collection workflow for publicly accessible drug registration search pages from the Egyptian Drug Authority (EDA). The workflow uses Selenium to navigate search results and drug details, Beautiful Soup to parse HTML, pandas to structure and deduplicate records, and openpyxl to format Excel outputs.

## Repository contents

- `src/eda_scraper_pages.py` — original scraper source code.
- `data/eda_all_drugs_temp.xlsx` — **intermediate**, not necessarily complete, exported data.
- `requirements.txt` — Python dependencies.

## Dataset snapshot

The supplied intermediate workbook contains **797 data rows and 25 columns** on the `Drugs_Data` sheet. Fields include drug name, dosage form, strength, applicant, registration number, marketing and licensing status, source URL, and extraction date. Counts describe the supplied snapshot only, not the complete EDA registry.

## How the workflow works

1. Open the EDA registered-drug search page in Chrome via Selenium.
2. Enter one of the predefined search keywords.
3. **A human manually completes CAPTCHA and submits the search**; the program does not bypass CAPTCHA.
4. Iterate result pages, open detail records, and parse fields and HTML tables.
5. Deduplicate results using registration number/name when available.
6. Periodically save an intermediate Excel workbook and, after completion, a final workbook.

## Setup

```bash
python -m venv .venv
# Activate the virtual environment for your operating system
pip install -r requirements.txt
python src/eda_scraper_pages.py
```

Chrome and a compatible Selenium/ChromeDriver installation are required. The original script expects an interactive browser and terminal. The EDA website may change, and the scraper has **not been re-run or verified against the live site** during portfolio preparation.

## Important limitations

- Manual CAPTCHA entry and terminal confirmations are required for each keyword.
- Keyword-based search does not guarantee complete registry coverage.
- The workbook is a temporary snapshot, not a certified medical or regulatory reference.
- Validate record completeness, deduplication, website access rules, and source attribution before production use.
- Do not use this dataset for medical decisions.

## Tech stack

Python, Selenium, Beautiful Soup, pandas, openpyxl, lxml.
