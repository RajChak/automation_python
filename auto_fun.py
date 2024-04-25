from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import time
#from selenium.common.exceptions import NoSuchElementException
import win32com.client as win32
import json
import shutil
import datetime

def load_config():
    with open('Details.json') as config_file:
        data = json.load(config_file)
    return data

def login(username,password,data):
    try:
    #print(f'Inside login Username is {username},pass is {password}')
        service=Service(executable_part='chromedriver.exe')
        driver=webdriver.Chrome(service=service)
        driver.maximize_window()
        driver.get(data['url'])
        wait=WebDriverWait(driver,1000)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'#j_username'))).send_keys(username)
        driver.find_element(By.CSS_SELECTOR,'#logOnFormSubmit > div').click()
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,'#password'))).send_keys(password)
        driver.execute_script("window.scrollBy(0, 500)")
        time.sleep(3)

        element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#root > main > div:nth-child(1) > div > div > div > div.uid-newbox__content > div.no-tracking.uid-autologin__container > div > div > div > div > form > div > div:nth-child(2)')))
        element.click()
        
        return driver, wait
    except Exception as e:
        print(f'Error {e} occured for user: {username}')

def navigate_to_case(wait, driver,username):
    try:
        time.sleep(3)
        # wait for busy indictor to disappear
        wait.until(EC.invisibility_of_element_located((By.ID, '__xmlview15--mainPage-busyIndicator')))
        wait.until(EC.invisibility_of_element_located((By.ID, '__xmlview4--mainPage-busyIndicator')))
        # click on case button
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'#shell-component---dashboard-43741043--iconTabBar--header-3-text'))).click()
        time.sleep(5)
        # Wait until the busy indicator disappears
        wait.until(EC.invisibility_of_element_located((By.ID, '__xmlview15--mainPage-busyIndicator')))
        #select the filter
        wait.until(EC.invisibility_of_element_located((By.ID, '__xmlview4--mainPage-busyIndicator')))
        target = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#__xmlview15--status_mcb-arrow')))
        target.click()
        time.sleep(5)
        # wait for busy indictor to disappear
        wait.until(EC.invisibility_of_element_located((By.ID, '__xmlview15--mainPage-busyIndicator')))
        # Check if checkbox is selected
        checkbox = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#__toolbar36 div[role="checkbox"]')))
        if checkbox.get_attribute('aria-checked') == 'false':
            checkbox.click()        
        time.sleep(5)
        driver.execute_script("window.scrollBy(0, 500)")
        #click on go button
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'#__xmlview15--caseListFilterBar-btnGo-BDI-content'))).click()
    except Exception as e:
        print(f'Error {e} occured for user: {username}')

def download(wait, driver, data,username):
    time.sleep(5)
    WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, '__xmlview15--caseListTable-busyIndicator')))
    #wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Download"]'))).click()
    target_button = driver.find_element(By.XPATH, '//button[@aria-label="Download"]')
    target_button.click()
    download_file=os.path.join(data['download_path'],'caseList.xlsx')
    while not os.path.exists(download_file): 
        time.sleep(1)
    while os.path.getsize(download_file) == 0: 
        time.sleep(1)
    #time.sleep(5)
    driver.quit()
    return download_file

def process_file(download_file, data,username):
    shutil.move(download_file,data['case_folder'])
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    case_file=os.path.join(data['case_folder'],'caseList.xlsx')
    case_df=pd.read_excel(case_file)
    case_updated_df=case_df[['CASE', 'SUBJECT', 'STATUS', 'PRIORITY','PRIORITY','REPORTER','CREATED ON (UTC)','UPDATED ON (UTC)']]
    now_str = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    update_case_file=os.path.join(data['case_folder'],f"updated_Case_{now_str}.xlsx")
    case_updated_df.to_excel(update_case_file,index=False)
    new_file_path = os.path.join(data['case_folder'],f'caseList_{now_str}.xlsx')
    os.rename(case_file, new_file_path)
    return now_str
def send_mail(data, now_str,username):
    receiver_email = ";".join(data['receiver_email'])
    subject = data['subject']
    olApp = win32.Dispatch('Outlook.Application')
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = subject + str(now_str)
    mailItem.BodyFormat = 2  # 2 is olFormatHTML
    attachment_path = f"{data['case_folder']}\\updated_Case_{now_str}.xlsx"    
    # Read the Excel file and convert it to HTML
    df = pd.read_excel(attachment_path)
    html = df.to_html(index=False)   
    # Add the HTML to the email body
    mailItem.HTMLBody = f"This is the updated case detail from {username} sheet:<br>{html}"    
    mailItem.To = receiver_email
    mailItem.Attachments.Add(attachment_path)
    mailItem.Display()
    mailItem.Save()
    mailItem.Send()

def main():
    try:
        data = load_config()
        for username,password in data['credential'].items():
            #print(f'Username is {username},pass is {password}')
            driver, wait = login(username,password,data)
            navigate_to_case(wait, driver,username)
            downloaded_file = download(wait, driver, data,username)
            now_str=process_file(downloaded_file, data,username)
            send_mail(data,now_str,username)
    except Exception as e:
        print('An eror ocurred:',e)

if __name__ == "__main__":
    main()

