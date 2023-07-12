# Zotero and ESA Name Comparison Tool

This tool is designed to compare names extracted from Zotero API and ESA spreadsheets. It helps identify if any names from the ESA spreadsheet appear in the Zotero data, providing a quick way to check for citation matches.

## Functions

- `zotero_name_scraper`: The `zotero_name_scraper` function, which scrapes author names from the Zotero API and saves the data to a CSV file called `zotero_creators.csv`.

- `esa_name_combiner`: The `esa_name_combiner` function, which combines names from multiple sheets of an ESA spreadsheet, cleans the data, and saves it to a CSV file called `esa_names.csv`.

- `compare():` The `compare` function, which compares the names from `esa_names.csv` and `zotero_creators.csv` and saves the results in a CSV file called `citation_checker.csv`.
## Requirements
 Run `pip install -r requirements.txt`to get all python requirements necessary
Make sure to download the spreadsheet of ESA Professors, and have the filepath available to paste into the terminal when prompted. One is included in `ESA coneferences.xlsx`, however this is probably out of date by the time this is seen.

## Usage

1. Call the `zotero_name_scraper` function in `main.py` to scrape author names from Zotero and save the data to `zotero_creators.csv`.

2. Call the `esa_name_combiner` function in `main.py` and provide the file path for the ESA spreadsheet when prompted. The function will clean the names and save them to `esa_names.csv`.

3. Call the `compare` function in `compare.py` and provide the file paths for `esa_names.csv` and `zotero_creators.csv`. The function will compare the names and create `citation_checker.csv` with an additional 'Appeared' column indicating if a name appeared in both datasets.



