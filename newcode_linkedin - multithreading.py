import csv
import pandas as pd
import openpyxl  #pip install openpyxl
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
from selenium.common import exceptions
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from threading import Thread

        
def scraper(rows):
    global flag
    global data
    global now
    try:     
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate",
            "DNT": "1",
            "Connection": "close",
            "Upgrade-Insecure-Requests": "1"
            }
        options = ChromeOptions()
        for x,y in headers.items():
            options.add_argument(f'--{x}={y}')
            
        options.add_argument('--headless')    #comment this line and try to run script for first time, if it runs without showing up captcha then you can uncomment it again and run script normally.
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--no-sandbox')
        options.add_argument('disable-dev-shm-usage')
        options.add_argument("--disable-webrtc")
        options.add_argument("--log-level=3")
        driver = Chrome(options=options)
        driver.get('https://www.linkedin.com/login')
        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys('omar58268@gmail.com')  #username
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys('NoAgeforanyempire_6') #password
        sign_in_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        sign_in_btn.click()
        time.sleep(10)

        for row in rows:
            if flag == False:
                driver.quit()
                break
            i = data[data['company_url'] == row].index[0]
            j = row
            link = j
            new_link = link + '/about' if 'showcase' not in link else link
        
            driver.get(new_link)
            time.sleep(4)
            
            try:
                data.loc[i,'logo'] = driver.find_element(By.XPATH,'//div[@class = "relative"]').find_element(By.TAG_NAME,'img').get_attribute('src')
            except Exception:
                data.loc[i,'logo'] = np.nan
            try:    
                data.loc[i,'about_us'] = driver.find_element(By.TAG_NAME,'p').text
            except Exception:
                data.loc[i,'about_us'] = np.nan
            try:
                data.loc[i,'phone'] = driver.find_element(By.XPATH,'//dt[contains(.,"Phone")]/following-sibling::dd').text if 'Phone' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'website'] = driver.find_element(By.XPATH,'//dt[contains(.,"Website")]/following-sibling::dd').text if 'Website' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'industry'] = driver.find_element(By.XPATH,'//dt[contains(.,"Industry")]/following-sibling::dd').text if 'Industry' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'company_size'] = driver.find_element(By.XPATH,'//dt[contains(.,"Company size")]/following-sibling::dd').text if 'Company size' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'headquarters'] = driver.find_element(By.XPATH,'//dt[contains(.,"Headquarters")]/following-sibling::dd').text if 'Headquarters' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'founded'] = driver.find_element(By.XPATH,'//dt[contains(.,"Founded")]/following-sibling::dd').text if 'Founded' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
                data.loc[i,'specialties'] = driver.find_element(By.XPATH,'//dt[contains(.,"Specialties")]/following-sibling::dd').text if 'Specialties' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
            except Exception:
                continue
            print(data.loc[i])
            if time.time() - now >= 60:
                break
    except Exception as e:
        print(e)

    

        


if __name__ == '__main__':
    flag = True
    now = time.time()
    try:
        data = pd.read_excel(r"C:\Users\omars\OneDrive\Desktop\files\linkedin_companies A_F.xlsx") #input path
        data = data.drop_duplicates(subset = 'company_url')
        if len(data.columns) == 2:
            data['logo'] = np.nan
            data['about_us'] = np.nan
            data['phone'] = np.nan
            data['website'] = np.nan
            data['industry'] = np.nan
            data['company_size'] = np.nan
            data['headquarters'] = np.nan
            data['founded'] = np.nan
            data['specialties'] = np.nan
            theindex = 0
        else:
            new_data = data.iloc[:,2:]
            theindex = new_data[new_data.isnull().all(axis=1)].index[0]
        d = list(data.loc[theindex:,'company_url'].values)
        number_of_threads = 3  #number of threads needed
        threads = []
        for i in range(number_of_threads):
            threads.append(Thread(target=scraper,args=(d[i::number_of_threads],)))
            threads[i].daemon = True
            
        for i in threads:
            i.start()
        for i in threads:
            i.join()
        
        
        
    except KeyboardInterrupt:
        flag = False
        
    finally:
        final_data = data.iloc[:,2:]
        final_index = final_data[final_data.notnull().any(axis=1)].index[-1]
        final_data = data.iloc[:final_index,2:]
        final_nulls = final_data.loc[final_data.isnull().all(axis=1)].index
        final_null_rows = list(data.loc[final_nulls,'company_url'].values)
        flag = True
        scraper(final_null_rows)
        data.to_excel(r"C:\Users\omars\OneDrive\Desktop\files\linkedin_companies1.xlsx",index=False) #output path
        
    