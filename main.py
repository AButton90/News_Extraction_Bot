from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from decouple import config

from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

import requests

import pandas as pd


class NewsExtractionBot:
    def __init__(self):
        self.site = "https://www.nytimes.com/"
        self.phrase = config("SEARCH_PHRASE")
        self.category = config("CATEGORY")
        self.determine_date_filters()

    def determine_date_filters(self):
        """Function that take the specified period and returns the start and end dates for the search"""

        period = config("PERIOD", cast=int)
        today = datetime.today()

        # Format the dates to the format used by the site mm/dd/yyyy and change the day to th first day of the month for start date and the last day of the month for end date
        self.start_date = (today - relativedelta(months=(period - 1))).strftime("%m/01/%Y")
        last_day = (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        self.end_date = last_day.strftime("%m/%d/%Y")

    def open_site(self):
        """Function that creates the webdriver and navigates to the news site"""

        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        self.driver.get(self.site)
        self.driver.maximize_window()

    def search_news(self):
        """Function that searches for the news and applies the applicable filters"""

        # Search for the news phrase specified
        WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/button"))).click()
        WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/div/form/div/input"))).send_keys(self.phrase)
        WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/div/form/button"))).click()

        # Filter for the 'Section' and 'Date Range'
        # Date Range
        WebDriverWait(driver=self.driver, timeout=30).until(method=EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/button"))).click()
        date_filter_options = WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_all_elements_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/ul/li")))

        for option in date_filter_options:
            if option.text == "Specific Dates":
                option.click()
                break

        WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/div[2]/div/label[1]/input"))).send_keys(self.start_date)
        to_date = WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/div[2]/div/label[2]/input"))).send_keys(self.end_date + Keys.ENTER)
        # to_date.send_keys(Keys.ENTER)
        # Section
        WebDriverWait(driver=self.driver, timeout=30).until(method=EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/button"))).click()
        sections = WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_all_elements_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/div/ul/li")))

        section_found = False
        sec_num = 0
        for section in sections:
            print(section.text)
            if self.category.lower() in section.text.lower():
                section_found = True
                break
            sec_num += 1

        if section_found:
            WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, f"/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/div/ul/li[{sec_num}]"))).click()
            WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_element_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/button"))).click()
        else:
            raise Exception("The specified section is not available")

    def extract_news(self):
        """Function that extracts the news and saves them in a dictionary"""

        # Extract the news
        news = WebDriverWait(driver=self.driver, timeout=30).until(EC.presence_of_all_elements_located(locator=(By.XPATH, "/html/body/div/div[2]/main/div/div[2]/div[2]/ol/li")))
        news_dict = {}

        for article in news:
            print(article.text)

        for i in range(len(news)):
            news_dict[i] = {}
            news_dict[i]["date"] = news[i].find_element_by_xpath(".//div/span").text
            news_dict[i]["title"] = news[i].find_element_by_xpath(".//div/div/div/a/h4").text
            news_dict[i]["description"] = news[i].find_element_by_xpath(".//div/div/div/a/p").text
            news_dict[i]["figure_name"] = news[i].find_element_by_xpath(".//div/div/figure/div/img").get_attribute("alt")
            news_dict[i]["figure_url"] = news[i].find_element_by_xpath(".//div/div/figure/div/img").get_attribute("src")
            news_dict[i]["phrase_count"] = 0
            news_dict[i]["money"] = False

        self.news_dict = news_dict

    def download_files(self):
        """Function that downloads the images of the news"""

        for key, value in self.news_dict.items():

            response = requests.get(value["figure_url"])

            with open(f"images_{key}.jpg", "wb") as file:
                file.write(response.content)

    def save_to_excel(self):
        """Function that saves the news dictionary to an excel file"""

        df = pd.DataFrame(self.news_dict).T
        df.to_excel("news.xlsx")

    def close_site(self):
        """Function that closes the webdriver"""

        self.driver.close()

    def run(self):
        """Function that runs the news extraction bot"""

        self.open_site()
        self.search_news()
        self.extract_news()
        self.close_site()
        self.download_files()
        self.save_to_excel()


if __name__ == "__main__":
    bot = NewsExtractionBot()
    bot.run()