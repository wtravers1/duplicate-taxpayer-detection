Delinquent_Account_Report_README: |
  # Delinquent Account Report Automation Tool

  ## Overview
  This automation tool logs into a county tax system and retrieves delinquent Real Estate (RES) and Personal Property (VPP) account reports.

  It identifies and exports:
  - Customers with both RES and VPP accounts
  - Probable matches between RES and VPP customers using fuzzy name and address matching
  - VPP customers with multiple active accounts

  The output is an Excel workbook with multiple sheets and formatted highlights.

  ## Features
  - Automated login and report retrieval via custom `tech_library`
  - Fuzzy matching for identifying probable matches
  - Excel output with:
    - Formatted headers
    - Auto-adjusted column widths
    - Conditional highlights for high street similarity
    - Alternating row colors for multi-account customers
  - Notifications on completion or errors

  ## Requirements
  - Python 3.12+
  - pandas
  - rapidfuzz
  - openpyxl
  - Custom automation library (`tech_library`)
  - Windows environment with compatible browser drivers

  ## Installation
  1. Clone the repository:
     ```bash
     git clone <repo_url>
     cd <repo_folder>
     ```
  2. Install dependencies:
     ```bash
     pip install pandas rapidfuzz openpyxl
     ```
  3. Update paths in the script:
     - `sys.path.append(r"path/to/your/automation/library")` → Path to your `tech_library`
     - `OUTPUT_FOLDER` → Path to folder for Excel output

  ## Configuration
  - `FUZZY_MATCH_THRESHOLD` – Minimum name similarity for fuzzy matches (default: 85)
  - `STREET_SIMILARITY_HIGHLIGHT_THRESHOLD` – Threshold to highlight similar streets (default: 80)

  ## Usage
  Run the report automation:
  ```bash
  python delinquent_account_report.py
