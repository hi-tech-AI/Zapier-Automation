from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from time import sleep
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from selenium.webdriver.support.ui import Select

def main():
    # Development mode ----------------------
    service = Service(executable_path=r"C:\chromedriver-win64\chromedriver.exe")   
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9030")
    driver = webdriver.Chrome(service=service, options=options)
    driver.get('https://app.boldsign.com/document/new')
    
    sleep(2)
    
    template_button = driver.find_element(By.ID, 'templateButton')
    driver.execute_script("arguments[0].click();", template_button)
    sleep(1)

    select_template = driver.find_element(By.XPATH, '//*[@id="template-grid_content_table"]/tbody/tr[1]')
    driver.execute_script("arguments[0].click();", select_template)
    
    sleep(1)
    select_button = driver.find_element(By.CLASS_NAME, 'bs-mergeTemplate')
    driver.execute_script("arguments[0].click();", select_button)

if __name__ == "__main__":
    main()