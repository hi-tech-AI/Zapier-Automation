from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from time import sleep
import pyautogui

def Send_Keys(element : WebElement, content : str):
    element.clear()
    for i in content:
        element.send_keys(i)
        sleep(0.01)

def Click(x, y):
    pyautogui.moveTo(x, y)
    sleep(1)
    pyautogui.leftClick(x, y)
    sleep(1)

def boldsign(personal_data, payment_method, pdf_file_path):
    client_name = personal_data[0]['<Client_Name>']
    client_email = personal_data[0]['<Client_Email>']
    attorney_name = personal_data[0]['<Lawyer_Name>']
    attorney_email = personal_data[0]['<Lawyer_Email>']

    service = Service(executable_path=r"C:\chromedriver-win64\chromedriver.exe")   
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9030")
    driver = webdriver.Chrome(service=service, options=options)

    driver.get('https://app.boldsign.com/document/new')
    
    sleep(5)
    
    # Click template selection button
    template_button = driver.find_element(By.ID, 'templateButton')
    driver.execute_script("arguments[0].click();", template_button)
    sleep(3)

    # Choose template
    if payment_method == "Zelle (instant)":
        select_template = driver.find_element(By.CLASS_NAME, 'scroll-bar').find_element(By.ID, 'template-grid_content_table').find_elements(By.TAG_NAME, 'tr')[1].find_elements(By.TAG_NAME, 'td')[0]
        driver.execute_script("arguments[0].click();", select_template)
        sleep(2)
    else:
        select_template = driver.find_element(By.CLASS_NAME, 'scroll-bar').find_element(By.ID, 'template-grid_content_table').find_elements(By.TAG_NAME, 'tr')[0].find_elements(By.TAG_NAME, 'td')[0]
        driver.execute_script("arguments[0].click();", select_template)
        sleep(2)
    
    # Click use button
    use_button = driver.find_element(By.ID, 'merge-Template-section').find_element(By.TAG_NAME, 'button')
    driver.execute_script('arguments[0].click();', use_button)
    sleep(10)

    Click(1600, 535)
    pyautogui.typewrite(pdf_file_path)
    pyautogui.press('enter')
    sleep(10)

    input_client_name = driver.find_elements(By.CLASS_NAME, 'e-autocomplete')[0]
    Send_Keys(input_client_name, client_name)
    sleep(1)

    input_client_email = driver.find_elements(By.CLASS_NAME, 'bs-signer-email')[0].find_element(By.TAG_NAME, 'input')
    input_client_email.clear()
    Send_Keys(input_client_email, client_email)
    sleep(1)

    input_attorney_name = driver.find_elements(By.CLASS_NAME, 'e-autocomplete')[1]
    Send_Keys(input_attorney_name, attorney_name)
    sleep(1)

    input_attorney_email = driver.find_elements(By.CLASS_NAME, 'bs-signer-email')[1].find_element(By.TAG_NAME, 'input')
    input_attorney_email.clear()
    Send_Keys(input_attorney_email, attorney_email)
    sleep(2)

    driver.get('https://app.boldsign.com/document/prepare')
    sleep(0.5)
    pyautogui.press('enter')
    sleep(10)

    submit_sign_button = driver.find_element(By.ID, 'submit-sign')
    driver.execute_script('arguments[0].click();', submit_sign_button)
    sleep(3)

    pyautogui.press('enter')

    driver.delete_all_cookies()

    driver.get('https://app.boldsign.com/dashboard')

if __name__ == "__main__":
    sleep(5)
    Click(1600, 535)
