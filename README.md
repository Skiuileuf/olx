# OLX Scraper and Analyzer

This repository contains two Python scripts for scraping and analyzing data from OLX.ro, a popular Romanian classifieds website. The scripts allow you to collect data about specific product listings and then process that data into an Excel spreadsheet for further analysis.

## Scripts

1. `apiDownload.py`: Scrapes data from OLX.ro API
2. `apiProcess.py`: Processes the scraped data and generates an Excel spreadsheet

### apiDownload.py

This script fetches data from the OLX.ro API and saves it as JSON files.

#### Features:
- Fetches data in batches (default: 50 items per request)
- Allows customization of search query
- Saves raw API responses as JSON files

#### Usage:
1. Set the `query` variable to your desired search term (e.g., "macbook+air+m1")
2. Adjust `max_results` if needed (default: 300)
3. Run the script: `python apiDownload.py`

The script will create an `api` folder (if it doesn't exist) and save the JSON files there.

### olx_analyzer.py

This script processes the JSON files created by `apiDownload.py` and generates an Excel spreadsheet with the analyzed data.

#### Features:
- Reads all JSON files in the `api` folder
- Extracts relevant information from each listing
- Calculates additional metrics (e.g., distance from a target location, days since last refresh)
- Generates an Excel file with formatted data

#### Usage:
1. Ensure you have run `apiDownload.py` first to generate the JSON files
2. Adjust the `target_lat` and `target_lon` variables if needed
3. Run the script: `python apiProcess.py`

The script will generate an Excel file in the same directory.

## Requirements

- Python 3.6+
- Required Python packages:
  - requests
  - openpyxl

You can install the required packages using pip:

```
pip install requests openpyxl
```

## Note

Please be respectful of OLX.ro's terms of service and avoid making too many requests in a short period. Consider adding delays between requests if you're fetching large amounts of data.

## License

This project is open-source and available under the MIT License.
