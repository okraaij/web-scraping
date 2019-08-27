# Import packages and setup time
from selenium import webdriver
import pandas as pd
import time, datetime, traceback, pyodbc, re
from statistics import mean

start_time = time.time()

# Setup Chrome window
driver = webdriver.Chrome("chromedriver.exe")

def commerzbank():
    commerzbanktickers = {'ADYEN NA': 'adyen nv', 'AMG NA': 'amg', 'BAMNB NA': 'bam groep nv', 'BESI NA': 'be semiconductor industries nv', 'FUR NA': 'fugro', 'PHARM NA': 'pharming group nv', 'PNL NA': 'postnl n.v.', 'RAND NA': 'randstad nv', 'TOM2 NA': 'tomtom', 'GLPG NA': 'galapagos nv'}
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

    # Iterate through every ticker
    for item in commerzbanktickers.values():
        ticker = item.replace(' ','%20')
        driver.get("https://www.beurs.commerzbank.com/product-search/turbo's/all-assettypes/" + ticker + "?Page=" + str(1))
        time.sleep(12)

        # Append referencelevel if present
        try:
            referencelevel = float(driver.find_element_by_css_selector('div.value').text.replace(",","."))
        except:
            referencelevel = 0

        # Obtain number of products for this ticker
        
        num_products = driver.find_element_by_css_selector('span.product-search-totalCount').text
        print(num_products)
        num_products = float(re.sub("\D", "", num_products))
        print(num_products)
        page = 0

        while num_products > 0:
            page += 1
            driver.get("https://www.beurs.commerzbank.com/product-search/turbo's/all-assettypes/" + ticker + "?Page=" + str(page))
            time.sleep(6)

            data = driver.find_element_by_xpath("//*[@class='k-grid-content k-auto-scrollable']").text

            # Select data for every turbo seperately
            turbos = data.split('DE')[1:]

            for item in turbos:
                item = item.split('\n')

                # Reformat instance if "sold out"
                item[0] = "DE" + item[0]
                if item[1] == "KO":
                    item.remove('KO')
                    item[0:2] = [' '.join(item[0:2])]
                    item[2:4] = ['EUR 0,0 0,0 0,0 0,0', 'KV', '']
                    item = item[:-1]

                splits = item[0].split(' ')

                # ISIN
                isin = splits[0]
                isins.append(isin)

                # Productname
                productname = splits[1] + " Turbo " + splits[-3] + " " + item[1][4:]

                # Productname correction for BESI for later in the script
                if splits[1] == 'BE':
                    productname = splits[1] + " " + splits[2] + " Turbo " + splits[-3] + " " + item[1][4:]
                productnames.append(productname)

                # Ratio
                ratio = float(splits[-2].replace(':1',""))
                ratios.append(ratio)

                # Financinglevel
                financinglevel = float(splits[-1].replace(",","."))
                financinglevels.append(financinglevel)

                # Stoploss
                stoplosslevel = float(item[1][4:].replace(",","."))
                stoplosslevels.append(stoplosslevel)

                # Reference level
                referencelevels.append(referencelevel)

                # Strikes
                strikes.append(float(0))
                splits = item[2].split(' ')

                # Bids
                try: 
                    bid = float(splits[1].replace(',','.'))
                except:
                    bid = 0
                bids.append(bid)

                # Asks
                try:
                    ask = float(splits[2].replace(',','.'))
                except:
                    ask = 0
                asks.append(ask)

                # Leverage
                try:
                    leverage = float(splits[-1].replace(',','.'))
                except:
                    leverage = 0
                leverages.append(leverage)

            num_products -= 25

    # Format DataFrame, add IssuerId and return result
    result = pd.DataFrame({'ISIN': isins, 'Name': productnames, 'Bid': bids, 'Ask': asks, 'Leverage': leverages, 'StopLoss': stoplosslevels, 'Strike': strikes, 'FinancingLevel': financinglevels, 'ReferenceLevel': referencelevels, 'Ratio': ratios})    
    commerzbank_result = result
    commerzbank_result['IssuerId'] = int(3)
    return(commerzbank_result)

# Setting up logfile
try:
    logf = open("logfile_commerzbank.log", "a")
except:
    logf = open("logfile_commerzbank.log", "w")
    
time_cur = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
logf.write(str(time_cur) + "\n")
logf.write("Starting the script:")
logf.write("\n")
    
text = ""

try:
    commerzbank_result = commerzbank()
    text += "Data succesfully scraped for Commerzbank \n"
except Exception as e:
    print(e)
    commerzbank_result = ""
    text += "Data unsuccesfully scraped for Commerzbank \n"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while running script for Commerzbank:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write("Scrape task completed on: " + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'))
logf.write("\n")
logf.write("\n")


# Set uniform UnderlyingId column and append to DataFrame
companies = []
for value in commerzbank_result.Name:
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
commerzbank_result['UnderlyingId'] = companies

# Re-format column order
commerzbank_result = commerzbank_result[['ISIN', 'Name', 'Bid', 'Ask', 'Leverage', 'StopLoss', 'FinancingLevel', 'ReferenceLevel', 'Ratio', 'UnderlyingId', 'IssuerId']]
driver.close()

try:
    if len(commerzbank_result) is not 0:
        # Prepare TurboInfo DataFrame by selecting headers and adding columns
        new_table = commerzbank_result.loc[:, ['UnderlyingId', 'IssuerId', 'ISIN']]
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
        old_table = pd.io.sql.read_sql("""SELECT * FROM dbo.TM_TurboInfo WHERE IssuerId = 3""", conn)

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
        turbodata = pd.merge(commerzbank_result, turboinfo, how='left', on='ISIN')
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

        text += "Data succesfully written to the database for Commerzbank\n"
        conn.close()
        
    else:
        text += "Data has not been written to the database for Commerzbank"
        time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
        error = "Empty DataFrame \n"
        logf.write(error)
        logf.write('\n')
        logf.write('\n')
        
except Exception as e:
    text += "Data has not been written to the database for Commerzbank"
    time_cur = datetime.datetime.fromtimestamp(time.time()).strftime('%c')
    error = "Error while writing to the database for Commerzbank:\n" + str(time_cur) + "\n" + str(traceback.format_exc())
    logf.write(error)
    logf.write('\n')
    logf.write('\n')

logf.write(text)
logf.write("\n")
logf.write("#########################################################################")
logf.write("\n")   
logf.close()
