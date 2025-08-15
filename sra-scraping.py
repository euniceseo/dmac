# READ ME
# - this code takes a while to run because of the random times here and there, for a 400 sample set you can expect it to take 30-40 mins
# - this code works with selenium, meaning you will need to download a compatible chrome webdriver in order for selenium to be able to connect to your chrome browser.
#   more info on chrome webdrivers can be found here: https://chromedriver.chromium.org/downloads. if the webdriver is one or two versions behind that should (theoretically) be ok
# - index in line 103: sometimes the HTML structure of each SRA page is different. change the index accordingly
# - same with scroll number: the default value of 3 should be ok, but feel free to change accordingly
# - currently works on a 'input omero link through terminal' in the format of 'python sra_scraping.py "LINK HERE"'

# selenium imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# general imports
import time
from bs4 import BeautifulSoup
import argparse
import pandas as pd
import random
import re
import numpy as np

# initializing tthe web driver
driver = webdriver.Chrome()

# creating a parser object
parser = argparse.ArgumentParser(description='Scrape data from a web page.')

# adding an argument to the command line
parser.add_argument('url', type=str, help='The URL to scrape data from')

# parse the command line arguments
args = parser.parse_args()

# extract the url
url = args.url

# gets page from url
driver.get(url)

time.sleep(3)

# html list, keeps track of the data
all_accession_data = []

# sleep – stops the script from running for 10 seconds
time.sleep(10)

# scrolling, currently scrolls three times but you can change that. scrolling to find the drop down & set it to max number possible
scroll_number = 3
for _ in range(scroll_number): 
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.END)
    
# locates the options + clicks 500 to show max values. can also change that though
dropdown = driver.find_element(By.ID, 'options')
dropdown.send_keys('500')

# sleep
time.sleep(10)

# grabs all the SRA accession links and puts them into one list
sra_links = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/object/SRR"]')
sra_urls = [link.get_attribute('href') for link in sra_links]

# iterates through the sra_urls
for sra_url in sra_urls:
    
    before_click = random.randint(2, 5)
    time.sleep(before_click)
    
    # gets the page link
    driver.get(sra_url)
    
    after_click = random.randint(5, 10)
    time.sleep(after_click)

    # grabbing the page data + appending to the data list
    accession_data = driver.page_source  
    all_accession_data.append(accession_data)

    # go back to the main bioproject page to grab the next SRA accession
    driver.back()
    
# quit the driver when all sra accessions have been run through
driver.quit()

# if you want to print the accession data to see if it's giving the right data
# print(accession_data)

# putting all the data in a beautifulsoup container
soup_list = [BeautifulSoup(data, 'html.parser') for data in all_accession_data]
data = []

# iterating through each soup in the soup_list
for soup in soup_list:

    # find the sample element name labeled biosample
    sample_element = soup.find('object-relations-tree-item', label='BioSample')
    if sample_element is not None:
        p_elements = soup.select('p.tree-item-title')
        if len(p_elements) >= 3:
            
            # this index may need to be changed based on the project – HTML structure of each one is different. change the index to 1, 2, etc.
            index = 1
            element = p_elements[index].text.strip()
            sample_name = element
            

    # find the <strong> element, will give you the sra accession
    sra_element = soup.find('object-relations-tree-item', label='SRA')
    sra_accession = sra_element.find('strong').text.strip()
    
    # to grab the libraryID
    experiment_table = soup.find('table')
    if experiment_table:
        rows = experiment_table.find_all('tr')[1:]
        for row in rows:
            columns = row.find_all('td')
            library_id = columns[1].text.strip()
    
    # checksums
    # Find all the <dd> elements with class "field-combo-viewer-wrapper" that have a 32-character pattern
    checksum_pattern = re.compile(r'^[a-fA-F0-9]{32}$')
    checksum_elements = [dd for dd in soup.find_all('dd', class_='field-combo-viewer-wrapper') if checksum_pattern.match(dd.text.strip())]
    if len(checksum_elements) >= 2:
        checksum1 = checksum_elements[0].text.strip()
        checksum2 = checksum_elements[1].text.strip()

    data.append([sra_accession, library_id, sample_name, checksum1, checksum2])
    
# creating a dataframe with all the info in it
df = pd.DataFrame(data, columns=['SRA Accession', 'Library ID', 'Sample Name', 'Checksum 1', 'Checksum 2'])

# sets the display option in pandas to none so you can see the entire df
pd.set_option('display.max_rows', None)

# display the df if you want
# print(df)

# output to excel
output_filename = 'test.xlsx'

# writes the parent df and the child df to the excep sheet
with pd.ExcelWriter(output_filename) as writer:
    df.to_excel(writer, sheet_name='df', index=False)
    
# prints done when done
print("done")
    
