from ast import While
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import datetime
import tkinter as tk
from tkinter import FIRST, messagebox



class WebScraper:
    def __init__(self):
        # Initialize the webdriver
        self.firefox_options = Options()
        self.firefox_options.headless = True
        self.driver = webdriver.Firefox(options=self.firefox_options)
        self.actions = ActionChains(self.driver)
        self.data_export = []

    def close_browser(self):
        self.driver.quit()

    def navigate_to_url(self, url):
        self.driver.get(url)
        sleep(30)  # Adjust sleep as needed

    def Start(self):
        # Now that the content is loaded, get the HTML
        print('Render the dynamic content to static HTML')
        html = self.driver.page_source       
        print(' Parse the static HTML')
        soup = BeautifulSoup(html, "html.parser")
        
        # Find all occurrences of 'div' with class "content p-3"
        all_divs = soup.find_all("div", {"class": "content p-3"})
        
        # Process each 'div' separately
        for item_div in all_divs:
            # Extracting information
            first_name = item_div.find('h3', class_='person__name').strong.get_text(strip=True)
            last_name = item_div.find('span', class_='last default-order').get_text(strip=True)
            title = 'Loan Officer'
            nmls = item_div.find('span', class_='profile-caption').get_text().split(':')[-1]
            email = item_div.find('a', class_='email tippy').get('href').split(':')[-1]
            phone = item_div.find('a', class_='phonee tippy').get('href').split(':')[-1]

            # Append the data to the list
            self.data_export.append({'First Name': first_name, 'Last Name': last_name, 'Title': title, 'NMLS #': nmls, 'Email': email, 'Phone #': phone})
            self.export_to_excel()
            print("Add to excel")


    def export_to_excel(self, file_name='output.xlsx'):
        # Create a DataFrame from the list of dictionaries
        df = pd.DataFrame(self.data_export)

        # Export DataFrame to Excel
        df.to_excel(file_name, index=False)
        print(f'Data exported to {file_name}')


# Example usage:
if __name__ == "__main__":
    url = "https://edgehomefinance.com/our-team/"
    scraper = WebScraper()

    try:
        scraper.navigate_to_url(url)
        scraper.Start()
        scraper.export_to_excel()

    finally:
        scraper.close_browser()
