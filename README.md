# URA GLS Sites Scraper
This repository contains a web scraper that monitors the Urban Redevelopment Authority (URA) Government Land Sales (GLS) site for awarded land parcels.

##Features

Scrapes the main URA GLS page daily
Identifies sites with "Awarded" status
Extracts details from individual project pages for awarded sites
Saves data to CSV format
Runs automatically via GitHub Actions

##Setup

Clone this repository to your GitHub account
Ensure GitHub Actions is enabled for the repository
The scraper will run automatically according to the schedule in the workflow file
You can also trigger the workflow manually from the Actions tab

33Files

ura_scraper.py: The main Python script that performs the web scraping
.github/workflows/scraper.yml: GitHub Actions workflow file to schedule and run the scraper
ura_land_sales.xlsx: Output Excel file containing the scraped data with multiple sheets
ura_scraper.log: Log file for debugging

33Requirements

Python 3.7+
Packages: requests, beautifulsoup4, pandas, openpyxl

33Customization
You might need to adjust the scraper based on the actual structure of the URA website. In particular:

CSS selectors in the parse_main_page() and fetch_project_details() methods
Change the cron schedule in the GitHub Actions workflow file to run at your preferred time
Modify the Excel formatting in the _format_excel_sheets() method if needed

33Notes

This scraper respects the website by using reasonable delays between requests
Make sure your repository has appropriate permissions to commit changes from GitHub Actions
Consider adding error notifications via email or other channels if needed
