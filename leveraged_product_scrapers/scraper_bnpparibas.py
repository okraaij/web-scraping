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

def bnp_paribas():
    driver.set_window_size(1500,1920)
    driver.get('https://www.bnpparibasmarkets.nl/producten/?u=&cat=Turbo&lf=&lt=&slf=&slt=&bdf=&bdt=&sdf=&sdt=&ttm=0&irf=&irt=&s=20&sortby=&direction=Ascending')
    time.sleep(1)

    # Hit close on cookies
    try:
        driver.find_element_by_xpath("//input[@value='Ik accepteer']").click()
    except:
        pass
    time.sleep(2)

    # Setup tickers, DataFrame and lists
    bnpparibas_tickers = {'ADYEN NA': 11151, 'AMG NA': 1349, 'BAMNB NA': 407, 'BESI NA': 1629, 'FUR NA': 408, 'PHARM NA': 11164, 'PNL NA': 801, 'RAND NA': 708, 'TOM2 NA': 769, 'GLPG NA': 2718}
    totalresult = pd.DataFrame(columns=['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'Strike', 'FinancingLevel', 'ReferenceLevel', 'Ratio'])    

    productnames = []
    isins = []
    bids = []
    asks = []
    stoplosslevels = []
    financinglevels = []
    leverages = []
    strikes = []
    referencelevels = []
    ratios = []
    all_isins = [] 

    # Obtain URL and split in sections
    url = driver.current_url.split("=")

    for item in bnpparibas_tickers.values():

        # Insert productcode in URL
        newurl = url[0] + "=" + str(item) + "=".join(url[1:]) + "&p=" + str(1)
        driver.get(newurl)

        # Obtain number of products
        num_products = float(driver.find_element_by_xpath("//h2[@class='heading trailer-half']").text.split(' ')[0])

        isins = []
        page = 1

        # Obtain all ISINs
        while num_products > 0:
            # Navigate to ticker's productpage
            newurl1 = url[0] + "=" + str(item) + "=".join(url[1:]) + "&p=" + str(page)
            driver.get(newurl1)
            time.sleep(1)

            # Obtain data
            data = driver.find_element_by_xpath("//*[@class='table']").text.split('\n')[3:]

            # ISINs
            isins_update = [item for item in data if item.startswith("NL")]
            isins += isins_update

            num_products -= 20
            page += 1

        all_isins += isins
        
    print(len(all_isins))

    # Iterate through every ISIN
    for val in all_isins:
        driver.get('https://www.bnpparibasmarkets.nl/producten/' + val + "/")

        data = driver.find_element_by_xpath("//div[@class='ribbon']").text.split('\n')

        # Bid
        try: 
            bid = float(data[1].replace(",","."))
        except:
            bid = 0
        bids.append(bid)

        # Ask
        try:
            ask = float(data[4].replace(",","."))
        except:
            ask = 0
        asks.append(ask)            

        # Ratio
        try:
            ratio = float(data[-1])
        except:
            ratio = float(data[-1].replace(",","."))
        ratios.append(ratio)

        # Financing level
        financinglevel = float(data[10].split(" ")[2].replace(",","."))
        financinglevels.append(financinglevel)

        # Stoploss 
        stoplosslevel = float(data[9].replace(",","."))
        stoplosslevels.append(stoplosslevel)

        # Productname
        productname = driver.find_element_by_xpath("//h1[@class='heading trailer-reset']").text
        productname = productname + " " + str(stoplosslevel)
        productnames.append(productname)

        # Reference level
        referencelevel = float(data[12].replace(",","."))
        referencelevels.append(referencelevel)

        # Strikes
        strikes.append(float(0))

        # Leverage
        try:
            leverage = float(data[-3].replace(",","."))
        except:
            leverage = 0
        leverages.append(leverage)

    # Add elements to DataFrame and append
    result = pd.DataFrame({'ISIN': all_isins, 'Name': productnames, 'Bid': bids, 'Ask': asks, 'Leverage': leverages, 'StopLoss': stoplosslevels, 'Strike': strikes, 'FinancingLevel': financinglevels, 'ReferenceLevel': referencelevels, 'Ratio': ratios})
    totalresult = pd.concat([totalresult, result], axis=0, sort=False)

    # Format DataFrame, add IssuerId and return result
    bnp_result = totalresult
    bnp_result['IssuerId'] = int(4)
    return(bnp_result)

# Setting up logfile
try:
    logf = open("logfile_bnpparibas.log", "a")
except:
    logf = open("logfile_bnpparibas.log", "w")
    
time_cur = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
logf.write(str(time_cur) + "\n")
logf.write("Starting the script:")
logf.write("\n")
    
text = ""

try:
    bnp_result = bnp_paribas()
    text += "Data succesfully scraped for BNP Paribas \n"
except Exception as e:
    bnp_result = ""
    text += "Data unsuccesfully scraped for BNP Paribas \n"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while running script for BNP Paribas:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write("Scrape task completed on: " + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
logf.write("\n")
logf.write("\n")

# Set uniform UnderlyingId column and append to DataFrame
companies = []
for value in bnp_result.Name:
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
bnp_result['UnderlyingId'] = companies

# Re-format column order
bnp_result = bnp_result[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio', 'UnderlyingId', 'IssuerId']]
driver.close()

try:
    if len(bnp_result) is not 0:
        # Prepare TurboInfo DataFrame by selecting headers and adding columns
        new_table = bnp_result.loc[:, ['UnderlyingId', 'IssuerId', 'ISIN']]
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
        old_table = pd.io.sql.read_sql("""SELECT * FROM dbo.TM_TurboInfo WHERE IssuerId = 4""", conn)

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
        turbodata = pd.merge(bnp_result, turboinfo, how='left', on='ISIN')
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

        text += "Data succesfully written to the database for BNP Paribas\n"
        conn.close()
        
    else:
        text += "Data has not been written to the database for BNP Paribas"
        time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
        error = "Empty DataFrame \n"
        logf.write(error)
        logf.write('\n')
        logf.write('\n')
        
except Exception as e:
    text += "Data has not been written to the database for BNP Paribas"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while writing to the database for BNP Paribas:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write(text)
logf.write("\n")
logf.write("#########################################################################")
logf.write("\n")   
logf.close()
