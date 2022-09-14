from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions as selenium_exception
import time
import socket

# selenium browser file location: file path must change whenever any new update to browser.
browser_location = r"C:\Users\iamke\OneDrive\Documents_Old\Drivers\msedgedriver.exe"  # edge

# edge profile location
profile_directory = "C:\\Users\\iamke\\AppData\\Local\\Microsoft\\Edge\\User Data\\Profile 1"  # profile 1

chat_list = "//*[@id='app']/div/div/div[3]/header"
search_bar = "//*[@id='side']/div[1]/div/div/div[2]/div/div[2]"
search_button = "//*[@id='side']/div[1]/div/div/button/div[2]/span"
message_bar = "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]"


def internet_connected():
    try:
        sock = socket.create_connection(("www.google.com", 80))
        if sock is not None:
            sock.close()
        return True
    except OSError:
        pass
    return False


def check_whatsapp_logged_in(driver):
    try:
        wait = WebDriverWait(driver, 100)
        wait.until(EC.visibility_of_element_located((By.XPATH, search_button))).click()
        driver.find_element(By.XPATH, chat_list).is_displayed()
        logged_in = True
    except selenium_exception.TimeoutException:
        logged_in = False
    return logged_in


def whatsapp_communicate(numbers, names, messages):

    intrnt_connected = False
    success_numbers = []
    failed_numbers = []
    logged_in = False
    
    browser_configuration = webdriver.EdgeOptions()
    browser_configuration.add_argument("user-data-dir="+profile_directory)
    browser_configuration.add_experimental_option("detach", True)
    driver = webdriver.Edge(executable_path=browser_location, options=browser_configuration)

    if internet_connected():
        intrnt_connected = True
        driver.get('https://web.whatsapp.com')
        driver.implicitly_wait(10)
        logged_in = check_whatsapp_logged_in(driver)
        if logged_in:
            driver.find_element(By.XPATH, search_button).click()
            for i in range(len(numbers)):
                if internet_connected():
                    driver.find_element(By.XPATH, search_bar).send_keys(numbers[i]+Keys.ENTER)
                    time.sleep(1)
                    try:
                        driver.find_element(By.XPATH, message_bar).send_keys((messages[i]))
                        driver.find_element(By.XPATH, message_bar).send_keys(Keys.ENTER)
                        print(i+1, 'Message Sent to', numbers[i])
                        success_numbers.append((i, numbers[i], messages[i]))
                        time.sleep(2)
                    except selenium_exception.NoSuchElementException:
                        try:
                            driver.find_element(By.XPATH, search_bar).clear()
                            driver.find_element(By.XPATH, search_bar).send_keys(names[i] + Keys.ENTER)
                            time.sleep(1)
                            driver.find_element(By.XPATH, message_bar).send_keys((messages[i]))
                            driver.find_element(By.XPATH, message_bar).send_keys(Keys.ENTER)
                            print(i + 1, 'Message Sent to', numbers[i])
                            success_numbers.append((i, numbers[i], messages[i]))
                            time.sleep(2)
                        except selenium_exception.NoSuchElementException:
                            failed_numbers.append((i, numbers[i], messages[i]))
                            print('Message Not Sent. Contact:', numbers[i], 'Not Found.')
                        
                else:
                    failed_numbers.append((i, numbers[i], messages[i]))
            driver.close()
            return success_numbers, failed_numbers, intrnt_connected, logged_in
        else:
            driver.close()
            return success_numbers, failed_numbers, intrnt_connected, logged_in
    else:
        failed_numbers = [(i, numbers[i], messages[i]) for i in range(len(numbers))]
        driver.close()
        return success_numbers, failed_numbers, intrnt_connected, logged_in
