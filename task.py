from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files

from datetime import datetime
from dateutil.relativedelta import relativedelta
import time
import re
import requests
import logging
import os
import sys


class InfoFilter(logging.Filter):
    def filter(self, rec):
        return rec.levelno < 30


class NewsExtractionBot:
    def __init__(self):
        self.site = "https://www.nytimes.com/"
        self.phrase = os.getenv(key="PHRASE")
        self.category = os.getenv(key="CATEGORY")
        self.period = int(os.getenv(key="PERIOD"))
        self.now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        self.determine_date_filters()
        self.set_logger()

        self.logger.info(msg=f"Initalized NewsExtractionBot for {self.site} with the following parameters: phrase={self.phrase}, category={self.category}, start_date={self.start_date}, end_date={self.end_date}")

    def set_logger(self):
        self.logger = logging.getLogger("__name__")
        self.logger.setLevel(logging.DEBUG)

        h1 = logging.StreamHandler(sys.stdout)
        h1.setLevel(logging.INFO)
        h1.addFilter(InfoFilter())
        h2 = logging.StreamHandler()
        h2.setLevel(logging.WARNING)

        self.logger.addHandler(h1)
        self.logger.addHandler(h2)

    def determine_date_filters(self):
        """Function that take the specified period and returns the start and end dates for the search"""

        today = datetime.today()

        self.start_date = (today - relativedelta(months=(self.period - 1))).strftime("%m/01/%Y")
        self.end_date = today.strftime("%m/%d/%Y")

    def open_site(self):
        """Function that creates the webdriver and navigates to the news site"""

        self.driver = Selenium()
        self.driver.open_available_browser(url=self.site, maximized=True)

    def search_news(self):
        """Function that searches for the news and applies the applicable filters"""

        # Search for the news phrase specified
        search_button = self.driver.wait_until_element_is_visible("xpath://html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/button", timeout=60)
        self.driver.click_element(search_button)
        self.driver.input_text("xpath://html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/div/form/div/input", self.phrase)
        self.driver.click_element("xpath://html/body/div[1]/div[2]/div[2]/header/section[1]/div[1]/div[2]/div/form/button")

        # Filter for the 'Section' and 'Date Range'
        # Date Range
        self.logger.info(msg=f"Searching for news from {self.start_date} to {self.end_date}.")
        self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/button")
        date_filter_options = self.driver.get_webelements("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/ul/li")

        for option in date_filter_options:
            if option.text == "Specific Dates":
                self.driver.click_element(option)
                break

        self.driver.input_text("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/div[2]/div/label[1]/input", self.start_date)
        self.driver.input_text("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/div[2]/div/label[2]/input", self.end_date)
        self.driver.press_keys("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[1]/div/div/div/div[2]/div/label[2]/input", "RETURN")

        # Section
        self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/button")
        sections = self.driver.get_webelements("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/div/ul/li")

        section_found = False
        sec_num = 1
        for section in sections:
            if self.category.lower() in section.text.lower():
                section_found = True
                break
            sec_num += 1

        if section_found:
            self.driver.click_element(f"xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/div/ul/li[{sec_num}]")
            self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/button")
            time.sleep(3)
        else:
            self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/div/ul/li")
            self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[1]/div[2]/div/div/div[2]/div/div/button")
            time.sleep(3)
            self.logger.warning(f"Section '{self.category}' not found. Searching for all sections instead.")

    def extract_news(self):
        """Function that extracts the news and saves them in a dictionary"""

        # Expand all results
        more_pages = True
        while more_pages:
            try:
                self.driver.wait_until_element_is_visible("xpath://html/body/div/div[2]/main/div/div[2]/div[3]/div/button", timeout=10)
                self.driver.click_element("xpath://html/body/div/div[2]/main/div/div[2]/div[3]/div/button")
            except Exception:
                more_pages = False
        # Extract the news
        index = 1
        news = self.driver.get_webelements("xpath://html/body/div/div[2]/main/div/div[2]/div[2]/ol/li")

        news_data = []

        for article in news:
            article_txt_list = article.text.split("\n")
            if article_txt_list[0].lower() != "advertisement":
                news_dict = {}
                news_dict["title"] = article_txt_list[2]
                news_dict["date"] = article_txt_list[0]
                try:
                    news_dict["description"] = article_txt_list[3]
                except IndexError:
                    news_dict["description"] = None
                news_dict["figure_name"], news_dict["figure_url"] = self.get_figure_info(index=index)
                news_dict["title_phrase_count"] = article_txt_list[2].lower().count(self.phrase.lower())
                news_dict["money"] = self.check_money_ref(article_name_description=f"{article_txt_list[2]} {article_txt_list[3]}")

                if news_dict not in news_data:
                    news_data.append(news_dict)

            index += 1

        self.logger.info(msg=f"Found {len(news_data)} news articles.")

        self.news_data = news_data

    def get_figure_info(self, index: int):
        """Function that gets the figure info of the news article"""

        fig_xpat = f"xpath://html/body/div/div[2]/main/div/div[2]/div[2]/ol/li[{index}]/div/div/figure/div/img"
        try:
            figure_name = self.driver.get_webelement(fig_xpat).get_attribute("alt")
            figure_url = self.driver.get_webelement(fig_xpat).get_attribute("src")
        except Exception:
            figure_name = None
            figure_url = None

        return figure_name, figure_url

    def check_money_ref(self, article_name_description: str):
        """Function that checks if any money reference is made in the news article, formats can be : $11.1 | $111,111.11 | 11 dollars | 11 USD"""

        if re.search(r"\$\d{1,3}(,\d{3})*(\.\d{2})?", article_name_description):
            return True
        else:
            return False

    def download_files(self):
        """Function that downloads the images of the news"""

        # Check if the folder exists
        if not os.path.exists("output/images"):
            os.mkdir("output/images")

        # Download the images
        for article in self.news_data:

            if article["figure_url"]:
                response = requests.get(article["figure_url"])

                file_name = re.sub(r"[^a-zA-Z0-9]+", "_", article["figure_name"])

                with open(f"output/images/{file_name}.jpg", "wb") as file:
                    file.write(response.content)

    def save_to_excel(self):
        """Function that saves the news dictionary to an excel file"""

        file_lib = Files()
        try:
            file_lib.open_workbook(path="output/news.xlsx")
        except FileNotFoundError:
            file_lib.create_workbook()
            file_lib.save_workbook("output/news.xlsx")

        news_data_no_url = []
        for article in self.news_data:
            del article["figure_url"]
            news_data_no_url.append(article)

        file_lib.create_worksheet(name=f"{self.phrase}_{self.category}_{self.now}", content=news_data_no_url, header=True)
        file_lib.save_workbook("output/news.xlsx")

    def close_site(self):
        """Function that closes the webdriver"""

        self.driver.close_all_browsers()

    def run(self):
        """Function that runs the news extraction bot"""

        try:
            self.open_site()
            self.search_news()
            self.extract_news()
            self.download_files()
            self.save_to_excel()

        except Exception as e:
            self.logger.error(msg=f"Error: {e}")

        finally:
            self.close_site()


if __name__ == "__main__":
    bot = NewsExtractionBot()
    bot.run()
