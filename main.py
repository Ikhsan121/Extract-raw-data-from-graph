import json
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from data_from_graph import create_excel
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging
from time import sleep
import requests
from selenium.webdriver.support.wait import WebDriverWait
import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class WebBrowser:
    def __enter__(self):
        """Start the browser when entering the context."""
        self.driver = self._start_browser()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Ensure the browser is closed when exiting the context."""
        self.quit()

    def _start_browser(self):
        """Set up the Chrome WebDriver."""
        options = Options()

        # Standard options for running in headless mode
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--headless")  # Remove or set to False for debugging
        # options.add_argument("--remote-debugging-port=9222")
        # options.add_argument("--disable-software-rasterizer")
        # options.add_argument("--window-size=1920x1080")
        # Enable performance logging (network traffic, etc.)
        # caps = DesiredCapabilities.CHROME
        # caps['goog:loggingPrefs'] = {'performance': 'ALL'}
        # # Set capabilities in options
        options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logger.info("Browser started successfully.")
        return driver


    def get_company_urls(self):
        """Navigate to a main page ."""
        logger.info(f"Navigating to the main page")
        links = []
        for i in range(20):
            logger.info(f"Navigating to the page {i+1}")
            try:
                self.driver.get(f"https://www.theaic.co.uk/aic/find-compare-investment-companies?sortid=Name&desc=false&page={i}")
                sleep(5)
                rows = self.driver.find_elements(By.CLASS_NAME, 'is-company-row')
                for row in rows:
                    element = row.find_element(By.CSS_SELECTOR, 'a.flex-1.text-brand-700.tour--click-fund')
                    href_attr =  element.get_attribute('href')
                    links.append(href_attr)
            except TimeoutException:
                logger.error(f"Timeout while trying to load https://www.theaic.co.uk/aic/find-compare-investment-companies?sortid=Name&desc=false&page={i}")
                raise
            except Exception as e:
                logger.error(f"Error loading https://www.theaic.co.uk/aic/find-compare-investment-companies?sortid=Name&desc=false&page={i}: {e}")
                raise
        return links

    def open_page(self, url: str):
        """Navigate to a given URL."""
        logger.info(f"Navigating to {url}")
        try:
            self.driver.get(url)
        except TimeoutException:
            logger.error(f"Timeout while trying to load {url}")
            raise
        except Exception as e:
            logger.error(f"Error loading {url}: {e}")
            raise

    def click_ten_years_button(self):
        """click ten years option on the page."""
        ten_years_button = self.driver.find_element(By.XPATH, '//button[@data-menuid="ten-year"]')
        self.driver.execute_script("arguments[0].click();", ten_years_button)
        logger.info(f"Element clicked: 10 years")


    def scroll_down(self):
        """scroll down to the bottom of the page"""
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # wait the graph to appear

    def fetch_graph_data(self,api):
        """fetch the data of the graph in json format"""
        nav_TR_API   = f"https://lt.morningstar.com/api/rest.svc/timeseries_cumulativereturn/fav18yujpm?applyTrackRecordExtension=true&currencyId=GBP&decPlaces=8&endDate={api['end_date']}&frequency=daily&id={api['ID']}&idType=Morningstar&outputType=json&performanceType=nav-cf&restructureDateOptions=ignore&startDate={api['start_date']}"
        share_price_TR_API = f"https://lt.morningstar.com/api/rest.svc/timeseries_cumulativereturn/fav18yujpm?applyTrackRecordExtension=true&currencyId=GBP&decPlaces=8&endDate={api['end_date']}&frequency=daily&id={api['ID']}&idType=Morningstar&outputType=json&performanceType=&restructureDateOptions=ignore&startDate={api['start_date']}"
        nav_API = f"https://lt.morningstar.com/api/rest.svc/timeseries_price/fav18yujpm?applyTrackRecordExtension=true&currencyId=GBP&decPlaces=8&endDate={api['end_date']}&forwardFill=true&frequency=daily&id={api['ID']}&idType=Morningstar&outputType=json&priceType=nav-cf&startDate={api['start_date']}"
        price_API = f"https://lt.morningstar.com/api/rest.svc/timeseries_price/fav18yujpm?applyTrackRecordExtension=true&currencyId=GBP&decPlaces=8&endDate={api['end_date']}&forwardFill=true&frequency=daily&id={api['ID']}&idType=Morningstar&outputType=json&priceType=price&startDate={api['start_date']}"
        nav_TR_API_json = requests.get(nav_TR_API).json()["TimeSeries"]['Security'][0]["CumulativeReturnSeries"][0]["HistoryDetail"]
        share_price_TR_API_json = requests.get(share_price_TR_API).json()["TimeSeries"]['Security'][0]["CumulativeReturnSeries"][0]["HistoryDetail"]
        nav_API_json =requests.get(nav_API).json()["TimeSeries"]['Security'][0]['HistoryDetail']
        price_API_json = requests.get(price_API).json()["TimeSeries"]['Security'][0]['HistoryDetail']
        return {
            'nav_TR_API_json': nav_TR_API_json,
            'share_price_TR_API_json': share_price_TR_API_json,
            'nav_API_json': nav_API_json,
            'price_API_json': price_API_json
        }


    def extract_network_log(self):
        """Extract network logs and return the ID, start date, and end date of the data"""
        logs = self.driver.get_log('performance')

        # Parse logs for network requests
        for log in logs:
            log_message = log['message']
            if 'Network.responseReceived' in log_message:
                if 'application/json' in log_message:
                    log_message_dict = json.loads(log_message)
                    # Check if 'response' key exists before accessing it
                    params = log_message_dict.get('message', {}).get('params', {})
                    response = params.get('response', {})
                    # If 'response' exists, access the URL
                    if response:
                        url_api = response.get('url')
                        if url_api:
                            # Parse the URL
                            parsed_url = urlparse(url_api)
                            # Parse the query parameters
                            query_params = parse_qs(parsed_url.query)
                            # Extract and validate the specific values
                            id_value = query_params.get('id', [None])[0]
                            start_date_str = query_params.get('startDate', [None])[0]
                            end_date = query_params.get('endDate', [None])[0]

                            # Convert the start_date to a datetime object
                            if start_date_str:
                                start_date = datetime.strptime(start_date_str,
                                                               "%Y-%m-%d")  # Adjust format as needed

                                # Calculate the date 10 years ago from today
                                current_year = datetime.now().year
                                ten_years_ago_year = current_year - 10

                                # Check if the start date is exactly 10 years from today
                                if start_date.year != ten_years_ago_year:
                                    continue  # Skip this log if the condition is not met

                                api = {
                                    'ID': id_value,
                                    'start_date': start_date_str,
                                    'end_date': end_date
                                }
        return api

    def quit(self):
        """Close the browser."""
        if self.driver:
            logger.info("Closing the browser...")
            self.driver.quit()


# Example Usage
if __name__ == "__main__":
    # Using a context manager to handle setup and teardown
    df = pd.read_csv("urls.csv")
    urls_list = df['URL'].to_list()

    with WebBrowser() as browser:
        # urls = browser.get_company_urls()
        for url in urls_list:
            try:
                browser.open_page(url)
                fname = url.split("/")[-2]
                sleep(2)
                # Scroll down to the bottom of the page
                browser.scroll_down()
                sleep(10)
                browser.click_ten_years_button()
                sleep(5)
                api_graph = browser.extract_network_log()
                # fetch the graph
                r = browser.fetch_graph_data(api=api_graph)
                create_excel(filename=fname, response=r)
                logger.info(f"{fname}.xlsx created successfully.")
            except TimeoutException as e:
                logger.info(f"This company '{fname}' is not currently a member of the AIC. We are therefore unable to provide full company information at this time.")