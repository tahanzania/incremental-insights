# Incremental Insights

A robust, local-first web application for analyzing campaign data and calculating incremental budget opportunities.

## Features

- **Local Processing**: All data remains on your device. No parsed data is sent to any server.
- **Excel Ingestion**: Drag and drop support for `.xlsx`, `.xls`, and `.csv`.
- **Dynamic Filtering**: Filter by Partner, Advertiser, Campaign, and Decisioned status.
- **Opportunity Logic**: Automatically identifies high-performing campaigns (Pacing ~100%, Score > 100) and computes value.
- **Email Builder**: Generates formatted email summaries for Partners or Advertisers to pitch incremental budgets.

## Quick Start
1.  Open the `index.html` file in any modern web browser.
2.  Drag and drop your campaign report Excel file.
    *   *Note: Ensure your file has columns like Partner, Advertiser, Campaign, Decisioned, Pacing, Budget, etc.*
3.  Use the dashboard to filter and analyze.
4.  Generate email templates in the Email Builder tab.
