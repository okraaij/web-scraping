# Leveraged product scrapers
Various Python scraping scripts that use selenium to obtain data for leveraged products from the most established providers.

## Overview

- This repository contains five different web scraping scripts that obtain data from five leveraged product providers and subsequently update that data in a database
- The scripts use SQL and the following (external) Python packages/dependencies, read the documentation for the correct use of the packages:
    - [selenium](https://selenium-python.readthedocs.io/) (web scraping package)
    - [pyodbc](https://github.com/mkleehammer/pyodbc/wiki) (database connection package)
- A separate price_correction.py script ensures that all prices for each stock are normalised by taking the mean price  
- The primary key in the database is the ISIN, where new ISINS are added as new rows and existing ISINs are updated. The script also updates whether the leveraged product is currently active. 
- Each script is designed to be triggered by a .bat file and start at a fixed period through Windows Task Scheduler and therefore logs its behavior in a logfile
- Data is collected from the following providers:
  - [ING](https://www.ingsprinters.nl/)
  - [BNP Paribas](https://www.bnpparibasmarkets.nl/producten/)
  - [Citigroup](https://nl.citifirst.com/)
  - [Commerzbank](https://www.beurs.commerzbank.com/)
  - [Goldman Sachs](https://www.gsmarkets.nl/en)
- Each script currently collects the data for the following stocks (with accompanying Bloomberg Terminal ticker):
  - Adyen (ADYEN NA Equity)
  - AMG (AMG NA Equity)
  - BAM (BAMNB NA Equity)
  - BE Semiconductor Industries (BESI NA Equity)
  - Fugro (FUR NA Equity)
  - Pharming (PHARM NA Equity)
  - Randstad (RAND NA Equity)
  - TomTom (TOM2 NA Equity)
  - Galapagos (GLPG NA Equity)
- For each stock, the following variables are scraped and written to a database:
  - ISIN
  - Name
  - Bid
  - Ask
  - Leverage
  - Stop loss level
  - Financing level
  - Reference level (latest underlying price)
  - Ratio
