from RPA.Browser.Selenium import Selenium
from robocorp import workitems
import os
import logging
from datetime import datetime
import pandas as pd
import requests
from time import sleep
from robocorp.tasks import task

from robocorp.tasks import task
from robocorp import storage
from pathlib import Path
import re

from robocorp import workitems
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel

from robocorp.workitems import Inputs, Outputs

class NewsRobot:
    def __init__(self):
        self.browser = Selenium()
        self.username = self.get_username()
        self.time_execution = datetime.now().strftime("%Y%m%d%H%M%S")
        self.setup_logging()

    def get_username(self):
        try:
            return os.getlogin()
        except OSError:
            return os.environ.get('USERNAME') or os.environ.get('USER')
        
    
    def setup_logging(self):
        
        log_path = f'C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Output\\Logs\\robot.log_{self.time_execution}'
        logging.basicConfig(filename=log_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def open_site(self, url):
        try:
            logging.info(f'Opening the Website: {url}')
            self.browser.open_available_browser(url)
            self.browser.wait_until_element_is_visible('//*[@id="Page-header-trending-zephr"]/div[1]/div[3]/bsp-search-overlay/button', timeout=55)
        except Exception as e:
            logging.error(f"It was not possible to load the website within the timeout period: {e}")
            raise
    def create_work_items_from_excel(self):
        excel = Excel()
        excel.open_workbook(f"C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Input\\Config.xlsx")
        rows = excel.read_worksheet_as_table(header=True)
    
        for row in rows:
            # Assuming each row is a dictionary with 'key' and 'value' columns
            key = row["Key"]
            value = row["Value"]
            workitems.outputs.create(payload={key: value})

    def searching_news(self):

        for item in workitems.inputs:
            # =-=-==-=- Access the payload of the work item -=-=-=-=-=
            search_phrase = item.payload.get("search_phrase")
            # After processing, mark the item as done
            item.done()
        
        try:

            try:
                self.browser.wait_until_element_is_visible("//*[contains(@class, 'bx-element-')]/button[text()='Decline']", timeout=5)
                self.browser.click_element("//*[contains(@class, 'bx-element-')]/button[text()='Decline']")
            except:
                pass

            self.browser.wait_until_element_is_visible("//button[@class='SearchOverlay-search-button']", timeout=5)
            # =-=-=-= Click on "Magnifier" to set the text =-=-=-= 

            self.browser.press_key(None, "ESC")
            # Define the screenshot file name and path
            screenshot_name = "screenshot.png"
            screenshot_path = Path("output") / screenshot_name
            
            self.browser.capture_page_screenshot(str(screenshot_path))
            # Store the screenshot in the Control Room
            storage.set_file(screenshot_name, screenshot_path)
    
            # Optionally, retrieve the file path for verification or further operations
            path = storage.get_file(screenshot_name, screenshot_path, exist_ok=True)
            logging.info("Stored screenshot in Control Room:", path)

            self.browser.click_element("//button[@class='SearchOverlay-search-button']")
            logging.info('Clicking on "Magnifier" to set the text')
            
            # =-=-=-= Searching for the news =-=-=-=
            self.browser.input_text(f"//input[@class='SearchOverlay-search-input']", {search_phrase}, clear=True)
            logging.info(f'Searching for the news: "{search_phrase}"')

            # =-=-=-= Click on "Magnifier" icon to search =-=-=-= 
            self.browser.click_element("//button[@class='SearchOverlay-search-submit']")
            self.browser.wait_until_element_is_visible("//div[@class='SearchResultsModule-filters-title']", timeout=40)
            logging.info('Search completed successfully.')

            # =-=-=-= Selecting "Newest" =-=-=-= 
            logging.info('Selecting "Newest"')
            self.browser.select_from_list_by_label("//select[@name='s']", 'Newest')
    
            #sleep(4)
            self.browser.wait_until_element_is_visible("//div[@class='SearchFilter-heading']", timeout=15)
            # =-=-=-= Clicking on "Category" =-=-=-= 
            logging.info('Clicking on "Category"')
            
            self.browser.click_element("//div[@class='SearchFilter-heading']")
            sleep(4)

            # =-=-=-= Selecting the "Stories" category =-=-=-= 
            self.browser.select_checkbox("//input[@value='00000188-f942-d221-a78c-f9570e360000']")

            sleep(4)
        except Exception as ErrorSearch:
            logging.error(f'Error during search: {ErrorSearch}')
            raise

    def get_news(self):
        title_list = []
        descriptions_list = []
        dates = []
        stored_url_list = []
        count_titles_list = []
        count_descriptions_list = []
        result = False

        # =-=-=-= Getting the list elements =-=-=-= 
        web_elements = self.browser.get_webelements("//div[@class='SearchResultsModule-results']//div[@class='PagePromo-content']")
        
        # =-=-=-= Getting the quantity of items =-=-=-=
        Quantity = len(web_elements)

        logging.info(f"Quantity of news: {Quantity}")
       
        for item in range(1, Quantity + 1):
            
            Title = self.browser.find_element(f"(//div[@class='SearchResultsModule-results']//span[@class='PagePromoContentIcons-text'])[{item}]").text
            Description = self.browser.find_element(f"(//div[@class='PageList-items']//div[@class='PagePromo-description']//span[@class='PagePromoContentIcons-text'])[{item}]").text
            Date = self.browser.find_element(f"//div[@class='SearchResultsModule-results']//div[{item}]//div[1]//div[*]//div[2]").text
            
            logging.info(f"Tittle: {Title}")
            logging.info(f"Description: {Description}")
            logging.info(f"Date: {Date}")
            logging.info("=-=--=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=")

            # =-=-=-= Creating Regex =-=-=-=
            pattern = r'(\$[\d,]+(?:\.\d+)?\b|\b\d+\s*(?:dollars|usd)\b)'  # Exemplo: $10, $10.50, $111,111.11, 11 dollars, 11 USD
            
            title_lower = Title.lower()
            description_lower = Description.lower()

            if re.search(pattern,title_lower) or re.search(pattern,description_lower):
                result = True
            else:
                result = False 

            has_url = False
            count_title = len(Title)
            count_description = len(Description)
            try:
                # =-=-==-=- Trying to get image URL -=-=-=-=-=                                    
                elements = self.browser.get_webelements(f"(//div[@class='PageList-items-item']//div[@class='PagePromo-media']//img)[{item}]")
                if elements:
                    filename = f"C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Output\\downloaded_image{item}.jpg"
                    for element in elements:
                        url = self.browser.get_element_attribute(element, "src")
                        has_url = True
                    response = requests.get(url)
                    with open(filename, 'wb') as image_file:
                        image_file.write(response.content)
                else:
                    logging.info(f"The news does not have an image")
            except Exception as error_image:
                logging.error(f"There was an error: {error_image}")

            title_list.append(Title)
            descriptions_list.append(Description)
            dates.append(Date)
            count_titles_list.append(count_title)
            count_descriptions_list.append(count_description)
            if has_url == True:
                stored_url_list.append(filename)
            else:
                stored_url_list.append("The news does not have an image")
        
        # =-=-==-=- Writing date to an Excel File -=-=-=-=-=     
        date_for_excel = {"Title": title_list, "Description": descriptions_list, "Update News": dates, "UrlPath": stored_url_list,"Title Count Phrases": count_titles_list, "Description Count Phrases": count_descriptions_list, "Amount of money": result}  
        df = pd.DataFrame(date_for_excel)
        df.to_excel(f"C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Output\\News_{self.time_execution}.xlsx", index=False)

    def check_and_create_folder(self, path):
        if not os.path.exists(path):
            os.makedirs(path)

    def upload_artifacts(self):
        # Assuming 'output' is the directory where your files are saved
        output_dir = Path(f'C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Output\\')
        outputs = Outputs()

        for file_path in output_dir.iterdir():
            if file_path.is_file() and file_path.suffix in ['.jpg', '.xlsx']:
                # =-=-==-=- Construct a unique name for the artifact in the cloud -=-=-=-=-=
                artifact_name = f"artifact-{file_path.name}"
                # =-=-==-=- Create a new output work item (or use an existing one) -=-=-=-=-=
                output_item = outputs.create(save=False)
                # =-=-==-=- Attach the file to the output work item -=-=-=-=-=
                output_item.add_file(file_path, name=artifact_name)
                # Save the work item to upload the file and finalize changes
                output_item.save()
                

    def main_task(self):

        path_folder = f"C:\\Users\\{self.username}\\Documents\\Robots\\RPAChallenge\\Output\\Logs"
        self.check_and_create_folder(path_folder)

        Attempts = 0
        MaxAttempts = 3
        while Attempts < MaxAttempts:
            try:
                self.open_site("https://apnews.com/")
                self.searching_news()
                self.get_news()
                self.upload_artifacts()
                logging.info('Robot executed successfully.')
                break
            except Exception as ErrorAttempt:
                Attempts +=1
                logging.error(f'Attempt {Attempts} failed with error: {ErrorAttempt}')
            finally:
                self.browser.close_all_browsers()

        if Attempts == MaxAttempts:
            logging.error('Maximum attempts reached, robot execution failed.')


@task
def run_news_robot():
    robot = NewsRobot()
    robot.main_task()

