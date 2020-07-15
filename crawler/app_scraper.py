# import Libraries
import urllib.request
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import pandas as pd


class AppScraper:

    def __init__(self, data, driver, app_page_url, missing, option = 0):

        # Get the app page
        driver.get(app_page_url)
        # Let the page load
        try:
              element = WebDriverWait(driver, 20).until(
                 EC.presence_of_element_located((By.XPATH, "//div[@class='tabContainer']//a[@class='defaultTab']"))
              )
        except:
            # Could not load the app, store the url in missing array and try to scrape it later
            missing.append(app_page_url)
            return
        # Check for the Products tab if this app has already been scraped
        if(self.option_check(option, driver)):
            has_reviews = True
            # Check if we have reviews to scrape
            check_reviews = driver.find_elements_by_xpath("//div[@class='appDetailHeader']//div[@class='ratingsCount']//a")[0].get_attribute("aria-label")
            if("0 User Reviews" in check_reviews):
                has_reviews = False
            # Go to the reviews tab
            review_tab_url = driver.find_elements_by_xpath("//div[@class='tabContainer']//a[@class='defaultTab']")[0].get_attribute("href")
            driver.get(review_tab_url)

            # Let the page load
            if(has_reviews):
                try:
#                     element = WebDriverWait(driver, 20).until(ExpectedConditions.or(
#                         EC.presence_of_element_located((By.XPATH, "//div[@class='tabContent']//div[@class='spza_appReviewContainer']//div[@class='reviewItem']")),
#                         EC.text_to_be_present_in_element((By.XPATH, "//div[@class='tabContent']//div[@class='spza_appReviewContainer']//div[3]"), "No reviews are available."))
#                     )
                    element = WebDriverWait(driver, 20).until(lambda driver: driver.find_elements_by_xpath("//div[@class='tabContent']//div[@class='spza_appReviewContainer']//div[@class='reviewItem']") or
                                               (driver.find_elements_by_xpath("//div[@class='tabContent']//div[@class='spza_appReviewContainer']//div[3]")[0].text == "No reviews are available."))
                    if driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']")[0].text.split("\n")[2] == "No reviews are available.":
                        has_reviews = False
                except:
                    # Could not load the reviews, check if reviews exist:
                    try:
                        no_reviews = driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']")[0].text.split("\n")[2]
                        if no_reviews == "No reviews are available.":
                            has_reviews = False
                    except:
                        # Browser timeout, save to try again later
                        missing.append(app_page_url)
                        return
                    has_reviews = False

            # Extract the meta data for the app
            try:
                app_name =  driver.find_elements_by_xpath("//div[@class='titleBlock']")[0].text
            except:
                app_name = "N/A"

            try:
                app_rating = driver.find_elements_by_xpath("//div[@class='appDetailHeader']//div[@class='detailsRating']//div")[0].get_attribute('aria-label')[0]
            except:
                app_rating = "N/A"

            try:
                app_developer = driver.find_elements_by_xpath("//div[@itemprop='publisher']//span")[0].text
            except:
                app_developer = "N/A"

            try:
                driver.execute_script("document.getElementById('correlationId').style.display='inline-block';")
                app_id = driver.find_elements_by_xpath('//*[@id="correlationId"]')[0].text
            except:
                app_id = "N/A"

            try:
                app_url = driver.find_elements_by_xpath("//div[@class='metaDetails']//a[@title='Support']")[0].get_attribute("href")
            except:
                app_url = "N/A"

            # Extract the review data for the app
            if(has_reviews):
                self.scrape_reviews(driver, review_tab_url, data, app_name, app_developer, app_rating, app_id, app_url)
            else:
                crawl_date = datetime.now()
                data.append((app_name, app_developer, app_rating, app_id, app_url, "N/A", "0", "N/A", "N/A", "N/A", crawl_date))
                return

    def scrape_reviews(self, driver, review_tab_url, data, app_name, app_developer, app_rating, app_id, app_url):

        on_last_page = False
        while(not on_last_page):

            # Get the elements for other review pages, current tab
            pages = driver.find_elements_by_xpath("//div[@class='reviewPagination']//li//a")
            try:
                current_tab = driver.find_elements_by_xpath("//div[@class='reviewPagination']//li[@class='f-active']//a")[0]
            except:
                # There is only 1 page of review
                on_last_page = True

            # Get the elements for necessary review contents
            reviews = driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']//div[@class='reviewItem']")
            if(len(reviews) == 0):
                return
            review_ratings = driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']//div[@class='reviewItem']//div[contains(@class, 'rating')]")
            left_content = driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']//div[@class='reviewItem']//div[@class='leftBar']")
            right_content = driver.find_elements_by_xpath("//div[@class='spza_appReviewContainer']//div[@class='reviewItem']//div[@class='rightContent']")

            # Scrape the review data
            for i in range(len(reviews)):

                review_rating = review_ratings[i].get_attribute('aria-label')[0]
                review_date = left_content[i].text.split("\n")[0]
                reviewer_name = left_content[i].text.split("\n")[1]

                try:
                    review_header = right_content[i].text.split("\n")[0]
                except:
                    review_header = "N/A"

                try:
                    review_text = right_content[i].text.split("\n")[1]
                except:
                    review_text = "N/A"

                if(review_text == "Report this review"):
                    review_text = "N/A"

                if(review_header == "Report this review"):
                    review_header = "N/A"

                # Append data
                crawl_date = datetime.now()
                data.append((app_name, app_developer, app_rating, app_id, app_url, review_date, review_rating, reviewer_name, review_header, review_text, crawl_date))

            if(not on_last_page):
                # Check if on last page
                if(pages[len(pages) - 1] == current_tab):
                    on_last_page = True
                    return

                # If not on last page, move on to the next review page
                else:
                    for i in range(len(pages)):
                        if(pages[i] == current_tab):
                            pages[i+1].click()
                            time.sleep(1)
                            break
        return

    def option_check(self, option, driver):
        meta_data = driver.find_elements_by_xpath("//div[@class='metadata']//div[@class='cell products']")[0].text

        if(option == 0):
            return True

        # If the app has "Web Apps" in the Products section skip scraping the app
        elif(option == 1):
            products_tags = ["Web Apps"]

        # If the app has Web Apps and any starting with "Dynamics 365" Products section, skip scraping the app
        elif(option == 2):
            products_tags = ["Web Apps", "Dynamics 365"]

        # If the app has Web Apps, any starting with "Dyanmics 365" or one of the following (Excel, OneNote, Outlook, Powerpoint, Project, SharePoint, Teams, Word) Products section, skip scraping the app
        elif(option == 3):
            products_tags = ["Web Apps", "Dynamics 365", "Excel", "OneNote", "Outlook", "PowerPoint", "Project", "SharePoint", "Teams", "Word"]

        # If the app has Web Apps, any starting with "Dyanmics 365", one of the following (Excel, OneNote, Outlook, Powerpoint, Project, SharePoint, Teams, Word) or Power BI apps Products section, skip scraping the app
        elif(option == 4):
            products_tags = ["Web Apps", "Dynamics 365", "Excel", "OneNote", "Outlook", "PowerPoint", "Project", "SharePoint", "Teams", "Word", "Power BI apps"]

        # If the app has Web Apps, any starting with "Dyanmics 365", one of the following (Excel, OneNote, Outlook, Powerpoint, Project, SharePoint, Teams, Word), Power BI apps or Power BI visuals Products section, skip scraping the app
        elif(option == 5):
            products_tags = ["Web Apps", "Dynamics 365", "Excel", "OneNote", "Outlook", "PowerPoint", "Project", "SharePoint", "Teams", "Word", "Power BI apps", "Power BI visuals"]

        return not any(matches in meta_data for matches in products_tags)



# # For unit testing
#
# data = []
# driver = webdriver.Firefox()
# # Specify the url
# urlpage = 'https://appsource.microsoft.com/en-us/product/office/SA000000007?tab=Overview'
# a = AppScraper(data, driver, urlpage, 0)
#
# # Save to pandas dataframe
# df = pd.DataFrame(data, columns=["app_name", "app_developer", "app_rating", "review_date", "review_rating", "reviewer_name", "review_header", "review_text", "crawl_date"])
# print(df)
