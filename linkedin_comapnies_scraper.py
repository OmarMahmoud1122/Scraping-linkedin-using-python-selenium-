import undetected_chromedriver as uc
from selenium.webdriver import Chrome,ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from playwright.sync_api import sync_playwright
import pandas as pd
import os
import threading

m = input('enter a new character or last character if not completed yet: ').upper()
t = float(input('enter time in hours for the scraper to be running ex: 4.0  '))
ex = None

def out():
    global ex
    ex = input('press enter to exit. ')
    
input_thread = threading.Thread(target=out)
input_thread.daemon = True
input_thread.start()
print('script is running.....')
    
headers = {
            "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "close",
            "Upgrade-Insecure-Requests": "1"
            }
if os.path.exists(f'linkedin_companies_{m}.csv') and os.path.exists(f'linkedin_inputs_{m}.txt'):
    df = pd.read_csv(f'linkedin_companies_{m}.csv')
    file = open(f'linkedin_inputs_{m}.txt','r')
    x = int(file.readline())
    y = int(file.readline())
else:
    x = 0
    y = 0
    df = pd.DataFrame(columns = ['company_name','company_url'])
with sync_playwright() as p:
    start = time.time()
    b = p.firefox.launch(headless=True)
    page = b.new_page(extra_http_headers=headers)
    page.goto('https://bing.com/')
    time.sleep(5)
    page.goto('https://www.linkedin.com/home')
    page.get_by_role('link',name='Companies').nth(1).click()
    time.sleep(5)
    page.locator('div.pagination-links').get_by_role('link',name=m).click()
    time.sleep(5)
    links = page.locator('ul.listings').locator('a.listings__entry-link').all()
    pages = page.locator('ol.flex-wrap').locator('a').all()
    for i in pages[x:]:
        if ex is not None or time.time() - start >= t * 3600:
            print('script is stopped.....')
            print(f'execution time is: {(time.time() - start)/ 3600}')
            break
        i.click()
        time.sleep(5)
        page_index = pages.index(i)
        print(f'page {i.text_content()} is starting.')
        for link in links[y:]:
            if ex is not None or time.time() - start >= t * 3600:
                file = open(f'linkedin_inputs_{m}.txt','w')
                file.write(str(page_index) + '\n')
                file.write(str(link_index))
                break
            df.loc[len(df)] = [link.text_content(),link.get_attribute('href')]
            print(df.loc[len(df) - 1])
            link_index = links.index(link)
        if ex is None:
            print(f'page {i.text_content()} is done.')
        elif ex is not None:
            print(f'still in page {i.text_content()}')
        y = 0
df.drop_duplicates(inplace=True)    
df.to_csv(f'linkedin_companies_{m}.csv',index = False)
        

            