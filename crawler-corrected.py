from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
from datetime import datetime
import sys

### Configuration
dbCols = ['Title', 'Authors', 'Affiliations', 'Publisher', 'Date & Volume',
    'Keywords', 'Abstract', 'Link']
sleepTime= 0.5
searchTerm = "decision+support+systems"
browser = 'firefox' # 'firefox' || 'chrome'
pathToDriver = 'geckodriver'
url_= "https://www.sciencedirect.com/search?qs=" + searchTerm
### end Configuration

# Initialize main dataframe
df = pd.DataFrame(columns = dbCols)

# Setup browser driver
browser_driver = Service(pathToDriver)
if browser == 'firefox':
    driver = webdriver.Firefox(service=browser_driver)
elif browser == 'chrome':
    driver = webdriver.Chrome(service=browser_driver)
else:
    print("Invalid browser setting; please recheck.")
    sys.exit(100)

driver.get(url_)
driver.maximize_window()
#driver.minimize_window()

# Get date+time for filename
dateTimeObj = datetime.now()
timestampStr = dateTimeObj.strftime("%d-%m-%Y_%H-%M-%S") #fixed filename for windows
fileName = searchTerm + "_" + timestampStr + '.xlsx'

# Main loop
while True:
    time.sleep(sleepTime*2) # sleep for a bit
    df2 = pd.DataFrame() # setup temp dataframe
    results = driver.find_elements(By.CLASS_NAME, "result-list-title-link") # get search results
    links = [result.get_attribute('href') for result in results] # get link for each result
    for link in links: # iterate on links
        driver.get(link) # load link
        _ref = True
        time.sleep(sleepTime)
        title = driver.find_element(By.CLASS_NAME, "title-text").text # get title
        author_name = driver.find_elements(By.CLASS_NAME, "given-name") # get author's name
        author_surname = driver.find_elements(By.CLASS_NAME, "surname") # get author's surname

        try:
            author_ref = driver.find_elements(By.CLASS_NAME, "author-ref") # get references (if any)
            if author_ref == []:
                _ref = False
        except NoSuchElementException:
            _ref = False

        kw = ''
        all_affs = ''
        all_names = ''

        try:
            keywords = driver.find_elements(By.CLASS_NAME, "keyword") # get keywords
            for keyword in keywords:
                kw += keyword.text + ", "
        except NoSuchElementException:
            kw = "N/A"

        try:
            driver.find_element(By.ID, "show-more-btn").click() # click "show more"
            affiliations = driver.find_elements(By.CLASS_NAME, "affiliation") # get affiliations (if any)
            for aff in affiliations:
                all_affs += aff.text + "\r"
        except NoSuchElementException:
            all_affs = "N/A"
        all_affs = all_affs.strip()

        if _ref:
            for name, surname, ref in zip(author_name, author_surname, author_ref): # merge author's name string (if affiliations)
                all_names += name.text + " " + surname.text + " (" + ref.text + "), "
        else:
            for name, surname in zip(author_name, author_surname): # merge author's name string (if no affiliations)
                all_names += name.text + " " + surname.text + ", "

        try:
            mag = driver.find_element(By.CLASS_NAME, "publication-title-link").text # get magazine
        except NoSuchElementException:
            mag = "N/A"

        try:
            mag2 = driver.find_element(By.CLASS_NAME, "text-xs").text # get magazine's volume & date
        except NoSuchElementException:
            mag2 = "N/A"

        try:
            abstract = driver.find_element(By.CLASS_NAME, "Abstracts").text # get abstract
        except NoSuchElementException:
            abstract = "N/A"


        thisArticle = (title, all_names, all_affs, mag, mag2, kw, abstract, link)
        df2= pd.DataFrame(thisArticle, index = dbCols) # pass all data to temp df
        df = df.append(df2.T, ignore_index=True) # append to main df
        #print(df)
        time.sleep(sleepTime)
        driver.back() # go back

    time.sleep(2)
    try:
        print("Writing content to excel!")
        df.to_excel(fileName, sheet_name=searchTerm) # save to excel
        print("Done!")
        driver.find_element(By.LINK_TEXT, "next").click() # try to go to next page
    except NoSuchElementException:
        print("No more pages to crawl!")
        break # exit
