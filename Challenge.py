from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
import logging
from time import sleep
import requests
import pandas as pd
from datetime import datetime
from robocorp import workitems
from robocorp.tasks import task

browser = Selenium()

@task
def create_work_item():
    # =-=-==-=- Creating an output work item with a payload -=-=-=-=-=
    workitems.outputs.create(payload={"search_phrase": "Breaking News"})

@task
def open_site(url):
    try:
        # =-=-=-= Open Browser =-=-=-= 
        logging.info(f'Opening the Website: {url}')
        browser.open_available_browser(url)
        
        # =-=-=-= Waiting for browser to load =-=-=-= 
        browser.wait_until_element_is_visible('//*[@id="Page-header-trending-zephr"]/div[1]/div[3]/bsp-search-overlay/button', timeout=45)
    except Exception as e:
        print(f"It was not possible to load the website within the timeout period: {e}")
        raise
@task
def searching_news():
    # =-=-==-=- Loop through all available input work items -=-=-=-=-=
    for item in workitems.inputs:
        # =-=-==-=- Access the payload of the work item -=-=-=-=-=
        search_phrase = item.payload.get("search_phrase")
        print(f"Search phrase is: {search_phrase}")
        # After processing, mark the item as done
        item.done()
    try:
        # =-=-=-= Click on "Magnifier" to set the text =-=-=-= 
        browser.click_element('//*[@id="Page-header-trending-zephr"]/div[1]/div[3]/bsp-search-overlay/button')
        logging.info('Clicking on "Magnifier" to set the text')

        # =-=-=-= Searching for the news =-=-=-= 
        browser.input_text(f'//*[@id="Page-header-trending-zephr"]/div[1]/div[3]/bsp-search-overlay/div/form/label/input', {search_phrase}, clear=True)
        logging.info('Searching for the news: "Breaking News"')

        # =-=-=-= Click on "Magnifier" icon to search =-=-=-= 
        logging.info('Clicking on "Magnifier" icon to search')
        browser.click_element('//*[@id="Page-header-trending-zephr"]/div[1]/div[3]/bsp-search-overlay/div/form/label/button')

        # =-=-=-= Waiting for page to load =-=-=-= 
        #Colocar if pagina nao carregar, parar o robô
        browser.wait_until_element_is_visible("//div[@class='SearchResultsModule-filters-title']", timeout=40)
        

        # =-=-=-= Selecting "Newest" =-=-=-= 
        logging.info('Selecting "Newest"')
        browser.select_from_list_by_label("//select[@name='s']", 'Newest')
    
        sleep(4)

        # =-=-=-= Clicking on "Category" =-=-=-= 
        logging.info('Clicking on "Category"')
        browser.click_element("//div[@class='SearchFilter-heading']")
        sleep(1)

        # =-=-=-= Selecting the "Stories" category =-=-=-= 
        browser.select_checkbox("//input[@value='00000188-f942-d221-a78c-f9570e360000']")

        sleep(4)
    except Exception as ErrorSearch:
        print(f'Erro: {ErrorSearch}')
        raise
@task
def get_news():    
        title_list = []
        descriptions_list = []
        dates = []
        stored_url_list = []
        count_titles_list = []
        count_descriptions_list = []
        result = False

        # =-=-=-= Getting the list elements =-=-=-= 
        web_elements = browser.get_webelements("//div[@class='SearchResultsModule-results']//div[*]//div[1]//div[*]//bsp-custom-headline[1]//div[1]")
        
        # =-=-=-= Getting the quantity of items =-=-=-=
        Quantity = len(web_elements)

        print(f"Quantity of news", Quantity)
        logging.info(f"Quantity of news: {Quantity}")
       
        for item in range(1, Quantity + 1):
            
            Title = browser.find_element(f"//div[@class='SearchResultsModule-results']//div[{item}]//div[1]//div[*]//bsp-custom-headline[1]//div[1]").text
            Description = browser.find_element(f"//body/div[@class='SearchResultsPage-content']/bsp-search-results-module[@class='SearchResultsModule']/form[@class='SearchResultsModule-form']/div[@class='SearchResultsModule-ajax']/div[@class='SearchResultsModule-ajaxContent']/bsp-search-filters[@class='SearchResultsModule-filters']/div[@class='SearchResultsModule-wrapper']/main[@class='SearchResultsModule-main']/div[@class='SearchResultsModule-results']/bsp-list-loadmore[@class='PageListStandardD']/div[@class='PageList-items']/div[{item}]/div[1]/div[*]/div[1]").text
            Date = browser.find_element(f"//div[@class='SearchResultsModule-results']//div[{item}]//div[1]//div[*]//div[2]").text


            title_lower = Title.lower()
            description_lower = Description.lower()
            if "$" in title_lower or "dollars" in title_lower or "usd" in title_lower \
                or "$" in description_lower or "dollars" in description_lower or "usd" in description_lower:
                result = True
            else:
                result = False

            has_url = False
            count_title = len(Title)
            count_description = len(Description)
            try:
                # =-=-==-=- Trying to get image URL -=-=-=-=-=                                    
                elements = browser.get_webelements(f"//html[1]/body[1]/div[3]/bsp-search-results-module[1]/form[1]/div[2]/div[1]/bsp-search-filters[1]/div[1]/main[1]/div[3]/bsp-list-loadmore[1]/div[2]/div[{item}]/div[1]/div[1]/a[1]/picture[1]/img[1]")
                if elements:
                    filename = f"downloaded_image{item}.jpg"
                    for element in elements:
                        url = browser.get_element_attribute(element, "src")
                        print(f"The URL is: {url}")
                        has_url = True
                    response = requests.get(url)
                    with open(filename, 'wb') as image_file:
                        image_file.write(response.content)
                else:
                    print(f"The news does not have an image")
            except Exception as error_image:
                print(f"There was an error: {error_image}")

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
        df.to_excel(f"C:\\Users\\guilherme.florencio\\Documents\\Robots\\RPAChallenge\\Output\\News_{time_execution}.xlsx", index=False)
@task
def log():
    logging.basicConfig(filename=f'C:\\Users\\guilherme.florencio\\Documents\\Robots\\RPAChallenge\\Output\\Logs\\robot.log_({time_execution})', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

#search_phrase = "Breaking News"
# =-=-==-=- Getting current date -=-=-=-=-=
now = datetime.now() # current date and time
# =-=-==-=- Formatting current date -=-=-=-=-=
time_execution = datetime.now().strftime("%Y%m%d%H%M%S") # Obtaining the unique identifier of each execution "time_execution"
print(f"Time execution: {time_execution}")

# Call the function
log()
create_work_item()

@task
def main_task():
    Attempts = 0
    MaxAttempts = 3
    while Attempts < MaxAttempts:
        try:
            open_site("https://apnews.com/")
            searching_news()
            get_news()
            print("Robot executed successfully.")
            logging.info('Robot executed successfully.')
            break
        except Exception as ErrorAttempt:
            Attempts +=1
            logging.error(f'{ErrorAttempt}')
            logging.error(f"Attempt {Attempts} failed with error: {ErrorAttempt}")
            #browser.close_all_browsers()
        finally:
            browser.close_all_browsers()

    if Attempts == MaxAttempts:
        print("Maximum attempts reached, robot execution failed.")

