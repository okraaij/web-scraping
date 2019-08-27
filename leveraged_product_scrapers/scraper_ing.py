# Import packages and setup time
from selenium import webdriver
import pandas as pd
import time, datetime, traceback, pyodbc, re
from statistics import mean

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("C:/TurboScraper/Project TurboScraper/chromedriver.exe")

def ing():
       
    # Setup ticker names and DataFrame
    ing_tickers = ['adyen','amg','bam','besi','fugro','galapagos','pharming','post-nl','randstad','tomtom']
    totalresult = pd.DataFrame(columns=['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio'])

    # Accept cookies and disclaimer
    driver.get('https://www.ingsprinters.nl/markten/aandelen/' + ing_tickers[0] + "/page/1")
    time.sleep(1)
    driver.find_element_by_xpath("//*[contains(text(), 'Akkoord')]").click()
    time.sleep(1)
    buttons = driver.find_elements_by_xpath("//*[contains(text(), 'akkoord')]")
    buttons[1].click()
    time.sleep(1)
      
    for ticker in ing_tickers:
        driver.get('https://www.ingsprinters.nl/markten/aandelen/' + ticker + '/page/1')    
        
        # Make sure all products are shown by clicking a 
        show_alls = driver.find_elements_by_xpath("//span[@class='js-label']")
        for i in show_alls:
            i.click()
            time.sleep(1)
            
        # Obtain productcodes from page by selecting links that contain "NL00"
        elems = driver.find_elements_by_xpath("//a[@href]")
        productcodes = [] 
        for elem in elems:
            link = elem.get_attribute("href")
            if "NL00" in link:
                productcode = link.split("/")[-1]
                productcodes.append(productcode)
        productcodes = list(set(productcodes))        

        productnames = []
        isins = []
        bids = []
        asks = []
        stoplosslevels = []
        leverages = []
        referencelevels = []
        ratios = []
        financinglevels = []
        
        for code in productcodes:
            driver.get('https://www.ingsprinters.nl/markten/aandelen/' + ticker + '/' + code)

            # Append productname
            productnames.append(driver.find_element_by_xpath("//h1[@class='text-body']").text.split('\n')[0])
            
            # Append ISIN
            isins.append(code)
            
            top_row = driver.find_element_by_xpath("//div[@class='cell']").text.split('\n')
            time.sleep(0.5)
            
            # Append bid and ask
            bids.append(float(top_row[1].replace(",",".")))
            
            asks.append(float(top_row[4].replace(",",".")))
            
            # Leverage
            leverages.append(float(top_row[10].replace(",",".")))
            
            # Stoploss
            stoplosslevels.append(float(top_row[12].replace(",",".")[1:]))
            
            # Referencelevel
            if "€" in top_row[14]:
                referencelevels.append(float(top_row[14].split(" ")[1].replace(",",".")))
            else:
                referencelevels.append(float(top_row[14].split(" ")[0].replace(",",".")))
            
            side_row = driver.find_elements_by_xpath("//div[@class='card']")[3].text.split('\n')
            
            # FinancingLevel
            financinglevels.append(float(side_row[3][1:].replace(",",".")))
            
            # Ratio
            ratios.append(float(side_row[9].replace(",",".")))
            
        # Place results in DataFrame and concatenate results
        result = pd.DataFrame({'ISIN': isins, 'Name': productnames, 'Bid': bids, 'Ask': asks, 'Leverage': leverages, 'StopLoss': stoplosslevels, 'FinancingLevel': financinglevels, 'ReferenceLevel': referencelevels, 'Ratio': ratios})
        totalresult = pd.concat([totalresult, result], axis=0, sort=False)
        
    totalresult = totalresult.reset_index(drop=True)       
    ingresult = totalresult
    ingresult['IssuerId'] = int(1)
    return(ingresult)

# Setting up logfile
try:
    logf = open("logfile_ing.log", "a")
except:
    logf = open("logfile_ing.log", "w")
    
time_cur = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
logf.write(str(time_cur) + "\n")
logf.write("Starting the script:")
logf.write("\n")
    
text = ""

try:
    ing_result = ing()
    text += "Data succesfully scraped for ING \n"
except Exception as e:
    ing_result = ""
    text += "Data unsuccesfully scraped for ING \n"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while running script for ING:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n') 

logf.write("Scrape task completed on: " + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
logf.write("\n")
logf.write("\n")

# Set uniform UnderlyingId column and append to DataFrame
companies = []
for value in ing_result.Name:
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
ing_result['UnderlyingId'] = companies

# Re-format column order
ing_result = ing_result[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio', 'UnderlyingId', 'IssuerId']]
driver.close()

try:
    if len(ing_result) is not 0:
        # Prepare TurboInfo DataFrame by selecting headers and adding columns
        new_table = ing_result.loc[:, ['UnderlyingId', 'IssuerId', 'ISIN']]
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
        old_table = pd.io.sql.read_sql("""SELECT * FROM dbo.TM_TurboInfo WHERE IssuerId = 1""", conn)

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
        turbodata = pd.merge(ing_result, turboinfo, how='left', on='ISIN')
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

        text += "Data succesfully written to the database for ING\n"
        conn.close()
        
    else:
        text += "Data has not been written to the database for ING"
        time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
        error = "Empty DataFrame \n"
        logf.write(error)
        logf.write('\n')
        logf.write('\n')
        
except Exception as e:
    text += "Data has not been written to the database for ING"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while writing to the database for ING:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write(text)
logf.write("\n")
logf.write("#########################################################################")
logf.write("\n")   
logf.close()
