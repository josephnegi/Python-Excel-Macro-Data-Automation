## Python Automation for Macroeconomic Data Collection and Reporting (Dissertation Project)

This repository contains Python code designed to automate the collection and processing of macroeconomic data, significantly reducing manual effort.

**Project Overview:**

This project automates the process of collecting and storing economic data relevant to financial markets and economic analysis. It integrates data from various sources, ultimately saving time spent on manual data entry and manipulation.

**Data Sources:**

* Publicly available Yahoo Finance API (S&P 500 data)
* Federal Reserve Economic Data (FRED) API (GDP data)
* Web scraping for:
    * US PMI and NMI data (Institute for Supply Management)
    * PMI Industry comments
    * Growth/contraction heatmap for PMI/NMI sub-indexes (18 industries)

**Data Points Collected (Monthly):**

* US PMI and NMI data (headline and sub-indexes: new orders, production, etc.)
* PMI Industry comments
* Heatmap of growth/contraction for all PMI/NMI sub-indexes (18 industries)

**Additional Data (Frequency):**

* S&P 500 data (monthly)
* GDP data (quarterly)

**Project Benefits:**

* Automates the retrieval and storage of economic data.
* Generates heatmaps for visualization of sub-index growth/contraction.
* Updates Excel spreadsheets with the collected data, eliminating manual work.
* Estimated time saved: 3 hours per month

**Getting Started:**

This project requires Python libraries for web scraping (e.g., BeautifulSoup) and data manipulation (e.g., Pandas). Configure API access for Yahoo Finance and potentially FRED (depending on data needs).

**Disclaimer:**

* This is a prototype and might require adjustments for specific use cases.
* Adapt the code for your preferred data formatting and spreadsheet structure.

**Link to Repository:**

https://github.com/josephnegi/Python-Excel-Macro-Data-Automation

**Feel free to explore the code and contribute to its improvement!**
