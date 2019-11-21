# Import packages
import pandas as pd
from selenium import webdriver

# Initiate values and driver
ov = 0
ovk = 0
og = 0
mon = "jan"

driver = webdriver.Chrome("C:/Users/***/chromedriver.exe")
result = pd.DataFrame()

# Iterate through pages
for i in range(1,18):
    url = "https://www.iex.nl/Forum/Topic/1357875/" + str(i) + "/vLK-2019.aspx"
    driver.get(url)
    commentcontainer = driver.find_elements_by_xpath("//div[@class='ForumPost']")
    
    # Count frequency certain texts per month
    for i in commentcontainer:
        date = i.text.split('2019')[0].split(' ')[-3:][1]
        if date == mon:            
            ov += i.text.count('overname')
            ovk += i.text.count('overnamekandidaat')
            og += i.text.count('overgenomen')
        else:
            resultsz = pd.concat([result, pd.DataFrame([mon, ov, ovk, og])], axis=1)
            ov = 0
            ovk = 0
            og = 0
            mon = date
            
# Concatenate results and transform DataFrame
result = pd.concat([resultsz, pd.DataFrame([mon, ov, ovk, og])], axis=1)    
result = result.T

# Set columns and month as index
result.columns = ['Month','Overname','Overnamekandidaat','Overgenomen']
result.set_index('Month')

display(result)
