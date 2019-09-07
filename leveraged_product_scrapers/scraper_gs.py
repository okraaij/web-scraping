##################################
#                                #
#   Leveraged product scrapers   #
#             oliverk1           #
#            July  2019          #
#                                #
##################################

# Import packages and setup time
from selenium import webdriver
import pandas as pd
import time, datetime, traceback, pyodbc, re
from statistics import mean

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("chromedriver.exe")

def goldman_sachs():
    driver.get('https://www.gsmarkets.nl/en/products?page=1')
    time.sleep(2)

    # Hit close on cookies
    driver.find_element_by_xpath("//span[.='Close']").click()

    # Setup ticker names and DataFrame
    gs_tickers = ['ADYEN NA', 'AMG NA', 'BAMNB NA', 'BESI NA', 'FUR NA', 'PHARM NA', 'PNL NA', 'RAND NA', 'TOM2 NA','GLPG NA']
    totalresult = pd.DataFrame(columns=['Productcode', 'ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'Strike', 'FinancingLevel', 'ReferenceLevel', 'Ratio'])

    # Iterate through every stock ticker
    for item in gs_tickers:

        # Load page for this ticker and set sleep to let page load
        driver.get('https://www.gsmarkets.nl/products?page=1&type=turbo&underlier=' + item)  
        time.sleep(2)

        # Obtain productcodes from URL
        elems = driver.find_elements_by_xpath("//a[@href]")
        productcodes = []
        for elem in elems:
            link = elem.get_attribute("href")
            if "/products/" in link:
                productcode = link.split("/")[-1]
                productcodes.append(productcode)
        productcodes = set(productcodes)

        # Check if Turbo products exist for this stock at this provider else pass
        if len(productcodes) == 0:
            pass

        # Scrape data if Turbo products exist
        else:

            productnames = []
            isins = []
            bids = []
            asks = []
            stoplosslevels = []
            leverages = []
            strikes = []
            referencelevels = []
            ratios = []

            # Extract variables for each stock
            for prodcode in productcodes:
                # "Try" command used because not all Turbo's are continuously traded
                try:
                    # Open page and expand page information
                    driver.get("https://www.gsmarkets.nl/en/products/" + prodcode)
                    time.sleep(1)
                    driver.find_element_by_xpath("//span[.='View all product terms']").click()

                    # Productname
                    productname = driver.find_element_by_xpath("//h1[@class='title']").text.replace(",",".")
                    productnames.append(productname)

                    # ISIN
                    isin = driver.find_element_by_class_name('ids').text.split('\n')[0][6:]
                    isins.append(isin)

                    # Bid and ask
                    pricing = driver.find_element_by_css_selector('div.price-data').text.split('\n')
                    bid = float(pricing[1].replace(',','.'))
                    bids.append(bid)
                    ask = float(pricing[8].replace(',','.'))
                    asks.append(ask)

                    # Key terms; leverage, stoploss and strike price
                    terms = driver.find_element_by_css_selector('div.key-terms').text.split('\n')

                    # Not all GS Turbo's are leveraged, append 0.0 if no leverage
                    try:
                        leverage = float(terms[1].replace("x","").replace(",","."))
                    except:
                        leverage = 0.0
                    leverages.append(leverage)

                    # Stoploss
                    stoploss = float(terms[3].replace(",","."))
                    stoplosslevels.append(stoploss)

                    # Strike price
                    strike = float(terms[7].replace(",","."))
                    strikes.append(strike)

                    # Reference level not always available, if not available append 0
                    try:
                        referencelevel = float(driver.find_element_by_css_selector('div.key-attributes').text.split('\n')[3].replace(',','.')[3:])
                    except:
                        referencelevel = 0.0
                    referencelevels.append(referencelevel)

                    # Ratio
                    ratio = float(driver.find_element_by_css_selector('div.react-slidedown').text.split('\n')[3].replace(',','.'))
                    ratios.append(ratio)
  
                except:
                    pass

            # Add elements to DataFrame and append
            result = pd.DataFrame({'ISIN': isins, 'Name': productnames, 'Bid': bids, 'Ask': asks, 'Leverage': leverages, 'StopLoss': stoplosslevels, 'Strike': strikes, 'FinancingLevel': stoplosslevels, 'ReferenceLevel': referencelevels, 'Ratio': ratios})
            totalresult = pd.concat([totalresult, result], axis=0, sort=False)
            
    # Format DataFrame, add IssuerId and return result
    totalresult = totalresult.reset_index(drop=True) 
    totalresult = totalresult[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'Strike', 'FinancingLevel', 'ReferenceLevel', 'Ratio']]
    gs_result = totalresult
    gs_result['IssuerId'] = int(2)
    
    return(gs_result)

# Setting up logfile
try:
    logf = open("logfile_goldmansachs.log", "a")
except:
    logf = open("logfile_goldmansachs.log", "w")
    
time_cur = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
logf.write(str(time_cur) + "\n")
logf.write("Starting the script:")
logf.write("\n")
    
text = ""

# Execute function
try:
    gs_result = goldman_sachs()
    text += "Data succesfully scraped for Goldman Sachs \n"
except Exception as e:
    gs_result = ""
    text += "Data unsuccesfully scraped for Goldman Sachs \n"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while running script for Goldman Sachs:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write("Scrape task completed on: " + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
logf.write("\n")
logf.write("\n")

# Set uniform UnderlyingId column and append to DataFrame
companies = []
for value in gs_result.Name:
    value = value.upper()
    if 'ADYEN' in value:
        companies.append(int(9))
    if 'AMG' in value:
        companies.append(int(4))
    if 'BAM' in value:
        companies.append(int(7))  
    if 'SEMICONDUCTOR' in value or 'BESI' in value:
        companies.append(int(11))
    if 'FUGRO' in value:
        companies.append(int(10))
    if 'PHARMING' in value:
        companies.append(int(3))
    if 'POST' in value:
        companies.append(int(12))
    if 'RANDSTAD' in value:
        companies.append(int(13))
    if 'TOM' in value:
        companies.append(int(2))
    if 'GALAPAGOS' in value:
        companies.append(int(5))
gs_result['UnderlyingId'] = companies

# Re-format column order
gs_result = gs_result[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio', 'UnderlyingId', 'IssuerId']]
driver.close()

# Update database
try:
    if len(gs_result) is not 0:
        # Prepare TurboInfo DataFrame by selecting headers and adding columns
        new_table = gs_result.loc[:, ['UnderlyingId', 'IssuerId', 'ISIN']]
        new_table['Currency'] = None
        new_table['IsLong'] = None
        new_table['IsActive'] = True
        new_table.columns = ['UnderlyingId', 'IssuerId', 'Isin', 'Currency','IsLong', 'IsActive']
        new_table = new_table.drop_duplicates("Isin").reset_index(drop=True)

        # Setup connection to SQL database
        
        # Connect to database and select data from the past 16 weeks
        conn = pyodbc.connect(
            "Driver={SQL Server Native Client 11.0};"
            "Server={PWSQL210};"
            "Database=***"
            "UID=***"
            "PWD=***"
        )
        cursor = conn.cursor()

        # Obtain old data
        old_table = pd.io.sql.read_sql("""SELECT * FROM dbo.TM_TurboInfo WHERE IssuerId = 2""", conn)

        # Find inactive ISINs (appear in the old data but not in the new)
        in_active_isins = list(old_table[~old_table['Isin'].isin(new_table['Isin'])]['Isin'])

        # Set inactive ISINs to inactive by updating in SQL 
        for value in in_active_isins:
            cursor.execute("""UPDATE dbo.TM_TurboInfo SET IsActive = 0 WHERE Isin=?""", value)
            conn.commit()

        # Set ISINs that were previously inactive to active
        newly_active_isins = list(pd.merge(old_table, new_table, how='inner', on='Isin')['Isin'])
        for value1 in newly_active_isins:
            cursor.execute("""UPDATE dbo.TM_TurboInfo SET IsActive = 1 WHERE Isin=?""", value1)
            conn.commit()

        # Find new ISINs (appear in new data but not in old)
        new_isins = list(new_table[~new_table['Isin'].isin(old_table['Isin'])]['Isin'])
        new_isins = new_table[new_table['Isin'].isin(new_isins)]

        # Append new ISINs to TurboInfo table
        for index, row in new_isins.iterrows():
            cursor.execute("INSERT INTO dbo.TM_TurboInfo(UnderlyingId, IssuerId, Isin, Currency, IsLong, IsActive) values (?,?,?,?,?,?)", row['UnderlyingId'], row['IssuerId'], row['Isin'], row['Currency'], row['IsLong'], row['IsActive'])
            conn.commit()

        # Extract assigned TurboIDs and related ISINs from TurboInfo table    
        turboinfo = pd.io.sql.read_sql("""SELECT TurboId, Isin FROM dbo.TM_TurboInfo""", conn)  
        turboinfo.columns = ['TurboId' , 'ISIN']

        # Merge TurboIDs to end_result DataFrame and add Timestamp and Creator
        turbodata = pd.merge(gs_result, turboinfo, how='left', on='ISIN')
        turbodata['CreatedOn'] = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%m"))
        turbodata['CreatedBy'] = "OKRA"
        turbodata = turbodata[['TurboId', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'ReferenceLevel', 'FinancingLevel', 'Ratio', 'CreatedOn', 'CreatedBy']]

        # Append updated TurboData to database table
        for index1, row1 in turbodata.iterrows():
            cursor.execute("INSERT INTO dbo.TM_TurboData(TurboId, Name, Bid, Ask, Leverage, StopLoss, ReferenceLevel, FinancingLevel, Ratio, CreatedOn, CreatedBy) values (?,?,?,?,?,?,?,?,?,?,?)", row1['TurboId'], row1['Name'], row1['Bid'], row1['Ask'], row1['Leverage'], row1['StopLoss'], row1['ReferenceLevel'], row1['FinancingLevel'], row1['Ratio'], row1['CreatedOn'], row1['CreatedBy'])
            conn.commit()

        text += "Data succesfully written to the database for Goldman Sachs\n"
        conn.close()
        
    # Write to logfile upon exception
    else:
        text += "Data has not been written to the database for Goldman Sachs"
        time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
        error = "Empty DataFrame \n"
        logf.write(error)
        logf.write('\n')
        logf.write('\n')
        
# Write to logfile upon exception        
except Exception as e:
    text += "Data has not been written to the database for Goldman Sachs"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while writing to the database for Goldman Sachs\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

# Store log to logfile and close file
logf.write(text)
logf.write("\n")
logf.write("#########################################################################")
logf.write("\n")   
logf.close()
