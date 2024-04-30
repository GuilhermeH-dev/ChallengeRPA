from RPA.Browser.Selenium import Selenium
from robocorp import workitems
import os
import logging
from datetime import datetime
import pandas as pd
import requests
from time import sleep
from robocorp.tasks import task
from pathlib import Path
import re


class NewsRobot:
    def __init__(self):
        self.browser = Selenium()
        self.time_execution = datetime.now().strftime("%Y%m%d%H%M%S")
        self.setup_logging()

    def setup_logging(self):
        log_name = f"robot.log_{self.time_execution}"
        log_path = Path("output") / log_name

        logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def open_site(self, url):
        try:
            logging.info(f'Opening the Website: {url}')
            self.browser.open_available_browser(url)
            self.browser.wait_until_element_is_visible("//button[@class='SearchOverlay-search-button']", timeout=55)
        except Exception as e:
            logging.error(f"It was not possible to load the website within the timeout period: {e}")
            raise

    def searching_news(self):

        for item in workitems.inputs:
            # =-=-==-=- Access the payload of the work item -=-=-=-=-=
            search_phrase = item.payload.get("search_phrase")
            # After processing, mark the item as done
            item.done()
        
        try:

            try:
                # =-=-==-=- Trying to close pop up if exists -=-=-=-=-=
                self.browser.wait_until_element_is_visible("//*[contains(@class, 'bx-element-')]/button[text()='Decline']", timeout=5)
                self.browser.click_element("//*[contains(@class, 'bx-element-')]/button[text()='Decline']")
            except:
                pass

            self.browser.wait_until_element_is_visible("//button[@class='SearchOverlay-search-button']", timeout=5)
            
            try:
                # =-=-==-=- Trying to click "I accept" if exists -=-=-=-=-=
                self.browser.click_element("//button[contains(., 'I Accept')]")
            except:
                pass

            # =-=-==-=- Define the screenshot file name and path -=-=-=-=-=
            screenshot_name = "screenshot.png"
            screenshot_path = Path("output") / screenshot_name
            
            self.browser.capture_page_screenshot(str(screenshot_path))

            # =-=-=-= Click on "Magnifier" to set the text =-=-=-= 
            self.browser.click_element("//button[@class='SearchOverlay-search-button']")
            logging.info('Clicking on "Magnifier" to set the text')
            
            # =-=-=-= Searching for the news =-=-=-=
            self.browser.input_text(f"//input[@class='SearchOverlay-search-input']", {search_phrase}, clear=True)
            logging.info(f'Searching for the news: "{search_phrase}"')

            # =-=-=-= Click on "Magnifier" icon to search =-=-=-= 
            self.browser.click_element("//button[@class='SearchOverlay-search-submit']")

            # =-=-=-= Waiting for page to load =-=-=-= 
            self.browser.wait_until_element_is_visible("//div[@class='SearchResultsModule-filters-title']", timeout=40)
            logging.info('Search completed successfully.')

            # =-=-=-= Selecting "Newest" =-=-=-= 
            logging.info('Selecting "Newest"')
            self.browser.select_from_list_by_label("//select[@name='s']", 'Newest')
    
            sleep(4)

            screenshot_name = "screenshotNewest.png"
            screenshot_path = Path("output") / screenshot_name
            
            self.browser.capture_page_screenshot(str(screenshot_path))
            
            self.browser.wait_until_element_is_visible("//div[@class='SearchFilter-heading']", timeout=15)
            # =-=-=-= Clicking on "Category" =-=-=-= 
            logging.info('Clicking on "Category"')
            self.browser.click_element("//div[@class='SearchFilter-heading']")
            sleep(4)

            screenshot_name = "screenshotStories.png"
            screenshot_path = Path("output") / screenshot_name
            
            self.browser.capture_page_screenshot(str(screenshot_path))

            # =-=-=-= Selecting the "Stories" category =-=-=-= 
            self.browser.click_element("//span[normalize-space()='Stories']")

            sleep(4)

            screenshot_name = "screenshotSelectedStories.png"
            screenshot_path = Path("output") / screenshot_name
            
            self.browser.capture_page_screenshot(str(screenshot_path))

        except Exception as error_search:
            logging.error(f'Error during search: {error_search}')
            raise

    def get_news(self):
        title_list = []
        descriptions_list = []
        dates = []
        stored_url_list = []
        count_titles_list = []
        count_descriptions_list = []
        result_list = []

        # =-=-=-= Getting the list elements =-=-=-= 
        web_elements = self.browser.get_webelements("//div[@class='SearchResultsModule-results']//div[@class='PagePromo-content']")
        
        # =-=-=-= Getting the quantity of items =-=-=-=
        quantity = len(web_elements)

        logging.info(f"Quantity of news: {quantity}")
       
        for item_index in range(1, quantity + 1):
            
            # =-=-=-= Getting news data =-=-=-=
            title = self.browser.find_element(f"(//div[@class='SearchResultsModule-results']//div[@class='PagePromo-title'])[{item_index}]").text
            description = self.browser.find_element(f"(//div[@class='PagePromo-description']/a/span)[{item_index}]").text
            date = self.browser.find_element(f"//div[@class='SearchResultsModule-results']//div[{item_index}]//div[1]//div[*]//div[2]").text
            
            
            logging.info(f"Tittle: {title}")
            logging.info(f"Description: {description}")
            logging.info(f"Date: {date}")
            logging.info("=-=--=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=")

            # =-=-=-= Creating Regex =-=-=-=
            pattern = r'(\$[\d,]+(?:\.\d+)?\b|\b\d+\s*(?:dollars|usd)\b)'
            
            title_lower = title.lower()
            description_lower = description.lower()

            if re.search(pattern,title_lower) or re.search(pattern,description_lower):
                result = "True"
            else:
                result = "False"

            has_url = False
            count_title = len(title)
            count_description = len(description)
            try:
                # =-=-==-=- Trying to get image URL -=-=-=-=-=                                    
                elements = self.browser.get_webelements(f"(//div[@class='PageList-items-item']//div[@class='PagePromo-media']//img)[{item_index}]")
                if elements:
                    image_name = f"downloaded_image{item_index}.jpg"
                    image_path = Path("output") / image_name

                    for element in elements:
                        url = self.browser.get_element_attribute(element, "src")
                        has_url = True
                    response = requests.get(url)
                    with open(image_path, 'wb') as image_file:
                        image_file.write(response.content)
                else:
                    logging.info(f"The news does not have an image")
            except Exception as error_image:
                logging.error(f"There was an error: {error_image}")

            title_list.append(title)
            descriptions_list.append(description)
            dates.append(date)
            count_titles_list.append(count_title)
            count_descriptions_list.append(count_description)
            result_list.append(result)
            if has_url == True:
                stored_url_list.append(image_path)
            else:
                stored_url_list.append("The news does not have an image")
        
        # =-=-==-=- Writing date to an Excel File -=-=-=-=-=     
        data_for_excel = {"Title": title_list, "Description": descriptions_list, "Update News": dates, "UrlPath": stored_url_list,"Title Count Phrases": count_titles_list, "Description Count Phrases": count_descriptions_list, "Amount of money": result_list}  
        df = pd.DataFrame(data_for_excel)
        
        excel_name = f"news_{self.time_execution}.xlsx"
        excel_path = Path("output") / excel_name

        df.to_excel(excel_path, index=False)
                
    def main_task(self):

        attempts = 0
        max_attempts = 3
        while attempts < max_attempts:
            try:
                self.open_site("https://apnews.com/")
                self.searching_news()
                self.get_news()
                logging.info('Robot executed successfully.')
                break
            except Exception as error_attempt:
                attempts +=1
                logging.error(f'Attempt {attempts} failed with error: {error_attempt}')
            finally:
                self.browser.close_all_browsers()

        if attempts == max_attempts:
            logging.error('Maximum attempts reached, robot execution failed.')


@task
def run_news_robot():
    robot = NewsRobot()
    robot.main_task()

