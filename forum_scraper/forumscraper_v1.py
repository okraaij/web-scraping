######################################################
#       Import packages and ignore warnings          # 
######################################################

import warnings
warnings.filterwarnings('ignore')

from selenium import webdriver
import pandas as pd
import time, re
import urllib.request
from bs4 import BeautifulSoup
import requests

#####################################################
#     Collect all recent forum pages per ticker     #
#####################################################

# Setup driver
driver = webdriver.Chrome("C:/Users/Olivier.Kraaijeveld/Documents/chromedriver.exe")
driver.get('https://www.iex.nl/Forum/Default.aspx')

# Find all urls on most recent page for letter A
elems = driver.find_element_by_xpath("//table[@class='ContentTable WithBorder']")
lala = elems.find_elements_by_xpath(".//a[@href]")

# Find all stock forum pages
letter_urls = {}
for item in lala:
    letter_urls[item.text] = item.get_attribute("href")
    
######################
#  Requests version  #
######################

# t = time.time()

# for link in letter_urls:
#     r = requests.get(link, verify=False)
#     soup = BeautifulSoup(r.content)
#     for a in soup.find_all("td",  {"class": "FirstRow WithRightBorder"}):
#         #print(a.find_all(href=True)[0])
#         break

# print((time.time() - t)/60)

######################
#  Selenium version  #
######################

# # Find the latest stock forum page
# t = time.time()

# for link in letter_urls:
#     driver.get(link) 
    
#     # Find URL of top page
#     page = driver.find_element_by_xpath("//td[@class='FirstRow WithRightBorder']")
#     page1 = page.find_element_by_xpath(".//a[@href]")
#     url = page1.get_attribute("href")
#     #print(url)

# print((time.time() - t)/60)

######################
#   Urllib version   #
######################

companies_urls = {}
# Extract values from
for key, value in letter_urls.items():
    page = urllib.request.urlopen(value)
    soup = BeautifulSoup(page.read())
    for a in soup.find_all("td",  {"class": "FirstRow WithRightBorder"}):
        try:
            output = str(a.find_all(href=True)[0])
            url = "http://www.iex.nl/" + output.split('../../')[1]
            url = url.split(" ")[0]
            url = url[:-1]
            print(url)
        except:
            url = None
        companies_urls[key] = url


####################################################
# Iterate through every forum page to collect data #
####################################################

end_result = pd.DataFrame()

# Open every page and extract any URLs in the comments
for key1, url in companies_urls.items():
    
    if url != None:
        
        # Open page
        driver.get(url)

        # Setup variables
        allurls = []
        usernames = []
        dates = []
        times = []
        texts = []
        urltitles = []

        # Obtain page title
        titel = driver.find_element_by_tag_name('h1').text.split("«")[0]

        # Extract urls and related variables within each comment section
        commentcontainer = driver.find_elements_by_xpath("//div[@class='ForumPost']")
        for i in commentcontainer:
            try:
                urls1 = re.findall('(?:(?:https?|ftp):\/\/)?[\w/\-?=%.]+\.[\w/\-?=%.]+', i.text)
                if len(urls1) != 0:
                    for link in urls1:
                        try:
                            if "/" in link:
                                # Extract variables from CommentBox
                                allurls.append(link)
                                print(link)
                                usernames.append(" ".join(i.text.split('2019')[0].split(' ')[:-3]))
                                dates.append(" ".join(i.text.split('2019')[0].split(' ')[-3:]) + "2019")
                                times.append(i.text.split('2019')[1].split(" ")[2].split('\n')[0])
                                texts.append(" ".join(list(filter(None, i.text.split('\n')[2:]))))

                                # Obtain URL title
                                def findTitle(url):
                                    webpage = urllib.request.urlopen(url).read()
                                    title = str(webpage).split('<title>')[1].split('</title>')[0]
                                    print(title)
                                    #return title
                                urltitles.append(findTitle(link))
                        except:
                            pass
            except:
                pass
    
        # Append variables to DataFrame and append to larger DataFrame
        if len(allurls) != 0:
            results = [dates, times, usernames, allurls, urltitles, texts]
            results = pd.DataFrame(results).T
            results['Title'] = titel
            results.columns = ['Date', "Time", "Username", "URL", "URL Title", "Text", "Titel"]
            results = results[['Titel', 'Date', "Time", "Username", "URL", "URL Title", "Text"]]
            end_result = pd.concat([end_result, results], axis=0)
        else:
            pass
    else:
        pass

#pd.set_option('display.max_colwidth', -1)
#end_result

#names = pd.read_csv('fullname.csv', sep = ';')
#for title in end_result.Titel:
#    #tick = names.loc[names['Title'] == some_value]
#    tick = names[names['Name'].str.contains(title)]
#    print(tick)
