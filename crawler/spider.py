# import Libraries
import sys
import urllib.request
from bs4 import BeautifulSoup
from selenium import webdriver
from datetime import datetime
import time
import pandas as pd
from app_scraper import AppScraper
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Spider:
    # TODO: Implement missing
    def __init__(self, driver, data):

        # List for skipped apps
        missing = []

        # main_url = 'https://appsource.microsoft.com/en-us/marketplace/apps'

        # Main category URLs
        web_apps = "https://appsource.microsoft.com/en-us/marketplace/apps?product=web-apps"
        add_ins_dynamics_365 = "https://appsource.microsoft.com/en-us/marketplace/apps?product=dynamics-365%3Bdynamics-365-business-central%3Bdynamics-365-for-customer-services%3Bdynamics-365-for-field-services%3Bdynamics-365-for-finance-and-operations%3Bdynamics-365-for-project-service-automation%3Bdynamics-365-for-sales&page=1"
        add_ins_microsoft_365 = "https://appsource.microsoft.com/en-us/marketplace/apps?page=1&product=office%3Bexcel%3Bonenote%3Bpowerpoint%3Bproject%3Boutlook%3Bsharepoint%3Bteams%3Bword"
        add_ins_power_bi_apps = "https://appsource.microsoft.com/en-us/marketplace/apps?page=1&product=power-bi"
        add_ins_power_bi_visuals = "https://appsource.microsoft.com/en-us/marketplace/apps?page=1&product=power-bi-visuals"
        power_apps = "https://appsource.microsoft.com/en-us/marketplace/apps?page=1&product=powerapps"
        missing = []
        categories = [web_apps, add_ins_dynamics_365, add_ins_microsoft_365, add_ins_power_bi_apps, add_ins_power_bi_visuals, power_apps]

        for i in range(len(categories)):
            self.crawl_apps(categories[i], data, driver, missing, i)

#        For testing
#        self.crawl_apps(add_ins_power_bi_apps, data, driver, missing, 0)

        # TODO: IMPLEMENT MISSING
        # self.crawl_missing()


    def crawl_apps(self, url, data, driver, missing, option):

        driver.get(url)

        try:
                element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='spza_filteredTileContainer']//div[@class='spza_tileWrapper']"))
                )
        except:
                # TODO: Keep track of missing categories

                return

        on_last_page = False
        while(not on_last_page):
            pages = driver.find_elements_by_xpath("//div[@class='paginationContainer']//li//a")
            current_tab = driver.find_elements_by_xpath("//div[@class='paginationContainer']//li[@class='f-active']//a")[0]
            apps = driver.find_elements_by_xpath("//div[@class='spza_filteredTileContainer']//div[@class='spza_tileWrapper']//a")
            # Go through each app
            for app in apps:

                # Store the app URL
                app_url = app.get_attribute("href")

                # Open and switch to a new tab
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[1])
                time.sleep(1)
                # Scrape the app
                s = AppScraper(data, driver, app_url, missing, option)

                # Close the tab and switch back
                driver.close()
                driver.switch_to.window(driver.window_handles[0])


            # Check if on last page
            if(pages[len(pages) - 1] == current_tab):
                on_lat_page = True
                return

            # If not on last page, move on to the next review page
            else:
                for i in range(len(pages)):
                    if(pages[i] == current_tab):
                        pages[i+1].click()
                        time.sleep(1)
                        break


data = []
driver = webdriver.Firefox()
crawl_date = str(datetime.now()).split(" ")[0]
start_time = str(datetime.now()).split(" ")[1]
Spider(driver, data)

# Save to pandas dataframe
df = pd.DataFrame(data, columns=["app_name", "app_developer", "app_rating", "app_id", "app_url", "review_date", "review_rating", "reviewer_name", "review_header", "review_text", "crawl_date"])

df.to_csv(crawl_date, index=False, encoding='utf-8')
