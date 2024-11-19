import pandas as pd
import openpyxl  #pip install openpyxl
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import numpy as np
try:
    headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
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
        
    options.add_argument('--headless')
    options.add_argument("--disable-webrtc")
    options.add_argument("--log-level=3")
    driver = Chrome(options=options)
    driver.get('https://www.linkedin.com/login')
    username_field = driver.find_element(By.ID, "username")
    username_field.send_keys('')  #uername
    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys('') #password
    sign_in_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
    sign_in_btn.click()
    time.sleep(10)

    data = pd.read_excel(r"C:\Users\omars\OneDrive\Desktop\files\linkedin_companies A_F.xlsx",dtype = str) #input path
    if len(data.columns) == 2:
        data['logo'] = np.nan
        data['about_us'] = np.nan
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
        
    for i,j in enumerate(data.loc[:,'company_url'].values[theindex:],start=theindex):
        try:
            link = j
            new_link = link + '/about' if 'showcase' not in link else link
            driver.get(new_link)
        except Exception:
            continue
        time.sleep(5)
        try:
            data.loc[i,'logo'] = driver.find_element(By.XPATH,'//div[@class = "relative"]').find_element(By.TAG_NAME,'img').get_attribute('src')
        except Exception:
            data.loc[i,'logo'] = ''
        try:    
            data.loc[i,'about_us'] = driver.find_element(By.TAG_NAME,'p').text
        except Exception:
            data.loc[i,'about_us'] = ''
        try:
            data.loc[i,'website'] = driver.find_element(By.XPATH,'//dt[contains(.,"Website")]/following-sibling::dd').text if 'Website' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else ''
       
            data.loc[i,'industry'] = driver.find_element(By.XPATH,'//dt[contains(.,"Industry")]/following-sibling::dd').text if 'Industry' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else ''
    
            data.loc[i,'company_size'] = driver.find_element(By.XPATH,'//dt[contains(.,"Company size")]/following-sibling::dd').text if 'Company size' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else ''
        
            data.loc[i,'headquarters'] = driver.find_element(By.XPATH,'//dt[contains(.,"Headquarters")]/following-sibling::dd').text if 'Headquarters' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else ''

            data.loc[i,'founded'] = driver.find_element(By.XPATH,'//dt[contains(.,"Founded")]/following-sibling::dd').text if 'Founded' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else np.nan
        
            data.loc[i,'specialties'] = driver.find_element(By.XPATH,'//dt[contains(.,"Specialties")]/following-sibling::dd').text if 'Specialties' in driver.find_element(By.CSS_SELECTOR,'dl.overflow-hidden').text else ''
        except Exception:
            continue
        
        print(data.loc[i])
            
except KeyboardInterrupt:
    print('The script is stopping now.')
    driver.quit()
finally:
    data.to_excel(r"C:\Users\omars\OneDrive\Desktop\files\linkedin_companies.xlsx",index=False) #output path
    print('script stopped')

