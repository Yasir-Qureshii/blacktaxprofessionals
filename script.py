import os
import re
import time
from datetime import datetime
import requests
import openpyxl 
import concurrent.futures
from openpyxl.styles import Font
from bs4 import BeautifulSoup


def generate_filename():
    now = datetime.now()
    xx = 'Output_' + str(now)[:-7] + '.xlsx'
    xx = xx.replace(' ', '_').replace(':', '-')
    return xx

filename = 'Output.xlsx'
if os.path.exists(filename):
    filename = generate_filename()
    
    
url = 'https://blacktaxprofessionals.com/taxprofessionals/category/taxprofessionals/page/'
urls = []

def scrape_blacktaxprofessionals(page):
    link = url + str(page)
    try:
        res = requests.get(link)
        soup = BeautifulSoup(res.text, 'lxml')
        titles = soup.find_all('h2', class_=["geodir-entry-title",  "h5", "text-truncate"])
        for  title in titles:
            try:
                href = title.find('a')['href']
                urls.append(href)
            except Exception as error:
                print(f'page: {page}; error in title: {error}')
    except Exception as error:
        print(f'page: {page}; error: {error}')
        


def find_pages():
    res = requests.get('https://blacktaxprofessionals.com/taxprofessionals/category/taxprofessionals')
    soup = BeautifulSoup(res.text, 'lxml')
    total_pages = soup.find('ul', 'pagination').find_all('a')[-2].text.strip()
    return int(total_pages)


pages = find_pages()
counter = []
i = 1
while i <= pages:
    counter.append(i)
    i += 1
    
print("Scraping the data...")
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(scrape_blacktaxprofessionals, counter)


excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Output'
sheet.append(['Busssines Name', 'Email', 'Website'])
excel.save(filename)


def scrape_blacktaxprofessionals_main(url, count):
    name = email = website = None
    try: 
        res = requests.get(url)
        soup = BeautifulSoup(res.text, 'lxml')
        try:
            name = soup.find('h1', class_=['entry-title' , 'main_title']).text.strip()
        except Exception as error:
            print(f'page: {url}; error in name: {error}')
            
        try:
            email = soup.find('div', class_=re.compile('geodir-field-email')).find('a').text.strip()
        except Exception as error:
            print(f'page: {url}; error in email: {error}')
            
        try:
            website = soup.find('div', class_=re.compile('geodir-field-website')).find('a')['href']
        except Exception as error:
            print(f'page: {url}; error in website: {error}')
        row = [name, email, website]
        sheet.append(row)
        if count % 30 == 0:
            excel.save(filename)
    except Exception as error:
        print(f'page: {url}; error: {error}')
 

counter = []
i = 1
while i <= 461:
    counter.append(i)
    i += 1
    
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(scrape_blacktaxprofessionals_main, urls, counter)
    
excel.save(filename)
