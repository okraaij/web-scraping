# Import packages and setup time
from selenium import webdriver
import pandas as pd
import time, datetime, traceback, pyodbc, re
from statistics import mean

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("C:/TurboScraper/Project TurboScraper/chromedriver.exe")

def citi():
    driver.set_window_size(1500,1920)
    driver.get('https://nl.citifirst.com/NL/Zoekopdracht')
    time.sleep(1)
    
    try:
        driver.find_element_by_xpath("//input[@value='Ik accepteer']").click()
    except:
        pass

    # Setup tickers and URLs
    citi_tickers = {'ADYEN NA': 'Adyen', 'AMG NA': 'AMG', 'BAMNB NA': 'BAM', 'BESI NA': 'BE Semiconductor', 'FUR NA': 'fugro', 'PHARM NA': 'pharming', 'PNL NA': 'post nl', 'RAND NA': 'randstad', 'TOM2 NA': 'tomtom', 'GLPG NA': 'galapagos'}
    all_urls = []
    all_isins = []

    # Search for ticker, see if Turbo's are present and obtain URLs
    for ticker in citi_tickers.values():
        element = driver.find_element_by_xpath("//input[@name='searchBox']")
        element.clear()
        element.send_keys(ticker + '\n')
        time.sleep(1)

        # Search for ticker in search box, if ticker does not match search result, no Turbo's are present for that ticker
        try: 
            if ticker.upper() in driver.find_element_by_xpath("//span[@class='suggGroupProductOnly']").text.upper():
                driver.find_element_by_xpath("//span[@class='suggGroupProductOnly']").click()
                time.sleep(2)

                # Find total number of products
                num_products = float(driver.find_element_by_xpath("//span[@class='ResultCount']").text.split(' ')[0])

                # Obtain ISINs and corresponding URLs
                while num_products > 0:            
                    all_isins += [item for item in driver.find_element_by_xpath("//div[@class='Results']").text.split('\n') if item.startswith('DE')]
                    all_urls += [item.get_attribute("href") for item in driver.find_elements_by_xpath("//a[@class='DetailLink']")]

                    # Go to next page
                    driver.find_element_by_xpath("//a[@class='Next']").click()
                    time.sleep(2)

                    num_products -= 25
        except:
            pass

    # Obtain all data
    productnames = []
    bids = []
    asks = []
    stoplosslevels = []
    financinglevels = []
    leverages = []
    strikes = []
    referencelevels = []
    ratios = [] 

    # Iterate through every product
    for url in all_urls:
        driver.get(url)
        data = driver.find_element_by_xpath("//div[@class='FullSizeControls']").text.split('\n')

        # Financing level
        financinglevel = float(data[4].split('/')[0].split(' ')[1])
        financinglevels.append(financinglevel)

        # Bid
        try:
            bid = float(data[7].replace("EUR","").replace(",","."))
        except:
            bid = 0
        bids.append(bid)

        # Ask
        try:
            ask = float(data[9].replace("EUR","").replace(",","."))
        except:
            ask = 0
        asks.append(ask)

        # Referencelevel
        referencelevel = float(data[14].replace("EUR","").replace(",","."))
        referencelevels.append(referencelevel)

        # Stoploss
        stoploss = float(data[4].split('/')[1].split(' ')[1:][1].replace(")",""))
        stoplosslevels.append(stoploss)

        # Productname
        productname = data[3] + " " + str(stoploss)
        productnames.append(productname)

        # Select different data table from page
        data = driver.find_elements_by_xpath("//div[@class='grid_4 omega column-Grid_4']")[1].text.split('\n')

        # Ratio
        ratio = float(data[4].split(' ')[1].split(':')[0].replace(",","."))
        ratios.append(ratio)

        # Leverage
        try:
            leverage = float(data[12].split(' ')[1].replace(",","."))
        except:
            leverage = 0
        leverages.append(leverage)

        # Strike
        strikes.append(0)

    # Format DataFrame, add IssuerId and return result
    result = pd.DataFrame({'ISIN': all_isins, 'Name': productnames, 'Bid': bids, 'Ask': asks, 'Leverage': leverages, 'StopLoss': stoplosslevels, 'Strike': strikes, 'FinancingLevel': financinglevels, 'ReferenceLevel': referencelevels, 'Ratio': ratios})
    citiresult = result
    citiresult['IssuerId'] = int(5)
    return(citiresult)

# Setting up logfile
try:
    logf = open("K:/Trading/Matlab Executables/TurboScraper/logfile_citi.log", "a")
except:
    logf = open("K:/Trading/Matlab Executables/TurboScraper/logfile_citi.log", "w")
    
time_cur = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
logf.write(str(time_cur) + "\n")
logf.write("Starting the script:")
logf.write("\n")
    
text = ""

try:
    citi_result = citi()
    text += "Data succesfully scraped for Citi \n"
except Exception as e:
    citi_result = ""
    text += "Data unsuccesfully scraped for Citi \n"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while running script for Citi:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n') 


logf.write("Scrape task completed on: " + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
logf.write("\n")
logf.write("\n")

# Set uniform UnderlyingId column and append to DataFrame
companies = []
for value in citi_result.Name:
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
citi_result['UnderlyingId'] = companies

# Re-format column order
citi_result = citi_result[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio', 'UnderlyingId', 'IssuerId']]
driver.close()

try:
    
    if len(citi_result) is not 0:
        # Prepare TurboInfo DataFrame by selecting headers and adding columns
        new_table = citi_result.loc[:, ['UnderlyingId', 'IssuerId', 'ISIN']]
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

        #########
        # Update TurboInfo table
        #########

        # Obtain old data
        old_table = pd.io.sql.read_sql("""SELECT * FROM dbo.TM_TurboInfo WHERE IssuerId = 5""", conn)

        # Find inactive ISINs (appear in the old data but not in the new)
        in_active_isins = list(old_table[~old_table['Isin'].isin(new_table['Isin'])]['Isin'])

        # # Set inactive ISINs to inactive by updating in SQL 
        for value in in_active_isins:
            cursor.execute("""UPDATE dbo.TM_TurboInfo SET IsActive = 0 WHERE Isin=?""", value)
            conn.commit()

        # # Set ISINs that were previously inactive to active
        newly_active_isins = list(pd.merge(old_table, new_table, how='inner', on='Isin')['Isin'])
        for value1 in newly_active_isins:
            cursor.execute("""UPDATE dbo.TM_TurboInfo SET IsActive = 1 WHERE Isin=?""", value1)
            conn.commit()

        # # Find new ISINs (appear in new data but not in old)
        new_isins = list(new_table[~new_table['Isin'].isin(old_table['Isin'])]['Isin'])
        new_isins = new_table[new_table['Isin'].isin(new_isins)]

        # # Append new ISINs to TurboInfo table
        for index, row in new_isins.iterrows():
            cursor.execute("INSERT INTO dbo.TM_TurboInfo(UnderlyingId, IssuerId, Isin, Currency, IsLong, IsActive) values (?,?,?,?,?,?)", row['UnderlyingId'], row['IssuerId'], row['Isin'], row['Currency'], row['IsLong'], row['IsActive'])
            conn.commit()

        # # Extract assigned TurboIDs and related ISINs from TurboInfo table    
        turboinfo = pd.io.sql.read_sql("""SELECT TurboId, Isin FROM dbo.TM_TurboInfo""", conn)  
        turboinfo.columns = ['TurboId' , 'ISIN']

        # # Merge TurboIDs to end_result DataFrame and add Timestamp and Creator
        turbodata = pd.merge(citi_result, turboinfo, how='left', on='ISIN')
        turbodata['CreatedOn'] = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%m"))
        turbodata['CreatedBy'] = "OKRA"
        turbodata = turbodata[['TurboId', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'ReferenceLevel', 'FinancingLevel', 'Ratio', 'CreatedOn', 'CreatedBy']]

        # # ##########
        # # # Update TurboData table
        # # ##########

        # # # Append updated TurboData to database table
        for index1, row1 in turbodata.iterrows():
            cursor.execute("INSERT INTO dbo.TM_TurboData(TurboId, Name, Bid, Ask, Leverage, StopLoss, ReferenceLevel, FinancingLevel, Ratio, CreatedOn, CreatedBy) values (?,?,?,?,?,?,?,?,?,?,?)", row1['TurboId'], row1['Name'], row1['Bid'], row1['Ask'], row1['Leverage'], row1['StopLoss'], row1['ReferenceLevel'], row1['FinancingLevel'], row1['Ratio'], row1['CreatedOn'], row1['CreatedBy'])
            conn.commit()
        text += "Data succesfully written to the database for Citi\n"
        conn.close()
        
    else:
        text += "Data has not been written to the database for Citi"
        time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
        error = "Empty DataFrame \n"
        logf.write(error)
        logf.write('\n')
        logf.write('\n')
        
except Exception as e:
    text += "Data has not been written to the database for Citi"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while writing to the database for Citi:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write(text)
logf.write("\n")
logf.write("#########################################################################")
logf.write("\n")   
logf.close()
