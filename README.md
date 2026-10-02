**to scrape companies names:**

  -run: _linkedin_comapnies_scraper.py_ to generate excel file with companies' names and links.
  
  -this file will be used in the other 2 scrappers to scrape companies details.

**to scrape companies details:**

  -Enter username and password in defined fields.

  -Enter path of companies' names excel file.

  -Run script and if script raises an errors, it means captcha appeared. so stop script and comment headless mode line( options.add_argument('--headless') ) and try to solve it manually as it appears.
