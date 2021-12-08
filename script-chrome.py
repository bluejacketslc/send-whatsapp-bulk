from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas
import urllib.parse

import os
import platform

if platform.system() == 'Darwin':
    # MACOS Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver'
else:
    # Windows Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver.exe'

driver = None
wait = None
chrome_options = None
Link = "https://web.whatsapp.com/"
usr_path = '--user-data-dir='+os.getcwd()+'\\User_Data'

def chrome_driver():
    options2 = webdriver.ChromeOptions()
    options2.add_argument(usr_path)
    driver = webdriver.Chrome(executable_path=chrome_default_path, options=options2)
    return driver

def whatsapp_QR():
    driver = chrome_driver()
    user_agent = driver.execute_script("return navigator.userAgent;")
    driver.get("https://web.whatsapp.com/")

    try:
        WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.CLASS_NAME, '_1G3Wr')))
    except Exception as e:
        print(e)
        exit(1)
    else:
        driver.close()
        return user_agent

def whatsapp_login():
    global wait, driver, Link, chrome_options
    chrome_options = Options()
    chrome_options.add_argument(usr_path)
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--hide-scrollbars')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument('--user-agent={}'.format(whatsapp_QR())) # User agent for validation
    chrome_options.add_argument(usr_path) #apdata user profile, to by pass QR code
    driver = webdriver.Chrome(executable_path=chrome_default_path, options=chrome_options)
    wait = WebDriverWait(driver, 600)
    driver.get(Link)
    print("QR scanned")

excel_data = pandas.read_excel('Recipients data.xlsx', sheet_name='Recipients')

count = 0

whatsapp_login()

try:
    WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.CLASS_NAME, '_1G3Wr')))
except Exception as e:
    print(e)
else:
    for column in excel_data['Contact'].tolist():
        try:
            url = 'https://web.whatsapp.com/send?phone=' + str(excel_data['Contact'][count]) + '&text=' + urllib.parse.quote(excel_data['Message'][count])
            driver.get(url)

            try:
                click_btn = WebDriverWait(driver, 120).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, '_4sWnG')))
            except Exception as e:
                print("Sorry message could not sent to " + str(excel_data['Contact'][count]))
            else:
                sleep(2)
                click_btn.click()
                WebDriverWait(driver, 120).until_not(EC.element_to_be_clickable((By.XPATH, "//span[@aria-label=' Pending ']")))

                list_media = excel_data['Media'][count].split(";")

                for media in list_media:
                    try:
                        clipButton = WebDriverWait(driver, 120).until(
                        EC.element_to_be_clickable((By.CLASS_NAME, '_2jitM')))
                        clipButton.click()
                        sleep(3)
                    except:
                        print("There was a problem while open clip button for "+str(excel_data['Contact'][count]))
                    else:
                        try:
                            attachImageButton = WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")))
                            attachImageButton.send_keys(os.getcwd() + "\\Media\\" + media)
                            sleep(3)
                        except:
                            print("There was a problem while attach image button for "+str(excel_data['Contact'][count]))
                        else:
                            try:
                                sendButton = WebDriverWait(driver, 120).until(
                                EC.element_to_be_clickable((By.CLASS_NAME, '_165_h')))
                                sendButton.click()
                                sleep(3)
                            except:
                                print("There was a problem while click send button for "+str(excel_data['Contact'][count]))
                            else:
                                WebDriverWait(driver, 120).until_not(EC.element_to_be_clickable((By.XPATH, "//span[@aria-label=' Pending ']")))

                list_document = excel_data['Document'][count].split(";")

                for document in list_document:
                    try:
                        clipButton = WebDriverWait(driver, 120).until(
                        EC.element_to_be_clickable((By.CLASS_NAME, '_2jitM')))
                        clipButton.click()
                        sleep(3)
                    except:
                        print("There was a problem while open clip button for "+str(excel_data['Contact'][count]))
                    else:
                        try:
                            attachDocumentButton = WebDriverWait(driver, 120).until(
                            EC.presence_of_element_located((By.XPATH, "//input[@accept='*']")))
                            attachDocumentButton.send_keys(os.getcwd() + "\\Document\\" + media)
                            sleep(3)
                        except:
                            print("There was a problem while attach document button for "+str(excel_data['Contact'][count]))
                        else:
                            try:
                                sendButton = WebDriverWait(driver, 120).until(
                                EC.element_to_be_clickable((By.CLASS_NAME, '_165_h')))
                                sendButton.click()
                                sleep(3)
                            except:
                                print("There was a problem while click send button for "+str(excel_data['Contact'][count]))
                            else:
                                WebDriverWait(driver, 120).until_not(EC.element_to_be_clickable((By.XPATH, "//span[@aria-label=' Pending ']")))

                print('Message sent to: ' + str(excel_data['Contact'][count]))
                    
            count = count + 1
        except Exception as e:
            print('Failed to send message to ' + str(excel_data['Contact'][count]) + str(e))
    driver.quit()
    print("The script executed successfully.")
