from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
from pathlib import Path
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from dotenv import load_dotenv

load_dotenv()

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")

# LOGIN
driver = webdriver.Firefox()
driver.get("https://outlook.live.com")

wait = WebDriverWait(driver, 5)

enter = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//a[@id='c-shellmenu_custom_outline_signin_bhvr100_right']")
))
enter.click()

email_input = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
email_input.send_keys(EMAIL + Keys.ENTER)

# TODO: microsoft added a middle screen here

password_input = wait.until(EC.visibility_of_element_located((By.ID, "passwordEntry")))
password_input.send_keys(PASSWORD + Keys.ENTER)

no_button = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//button[@data-testid='secondaryButton' and text()='Não']")
))
no_button.click()

# SINGLE E-MAIL
first_email = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//div[@aria-label='Message list']//div[@role='option'][1]"
)))

actions = ActionChains(driver)
actions.context_click(first_email).perform()

download_option = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//span[contains(@class, 'fui-MenuItem__content') and text()='Download']"
)))
download_option.click()

eml_option = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//span[contains(@class, 'fui-MenuItem__content') and text()='Download as EML']"
)))
eml_option.click()

actions.context_click(first_email).perform()

move_option = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//span[contains(@class, 'fui-MenuItem__content') and text()='Move']"
)))
move_option.click()

archive_option = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//span[contains(@class, 'fui-MenuItem__content') and text()='Archive']"
)))
archive_option.click()

# MULTIPLE
while True:
    try:
        first_email = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[@aria-label='Message list']//div[@role='option'][1]"
        )))

        actions = ActionChains(driver)
        actions.context_click(first_email).perform()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Download']"
        ))).click()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Download as EML']"
        ))).click()

        first_email = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//div[@aria-label='Message list']//div[@role='option'][1]"
        )))
        actions.context_click(first_email).perform()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Move']"
        ))).click()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Archive']"
        ))).click()

    except:
        print("No more emails to process.")
        break

# V2
import time
from selenium.common.exceptions import StaleElementReferenceException

while True:
    try:
        emails = driver.find_elements(
            By.XPATH,
            "//div[@aria-label='Message list']//div[@role='option']"
        )

        if not emails:
            print("No emails. Waiting 3 seconds...")
            time.sleep(3)
            continue

        first_email = emails[0]

        actions = ActionChains(driver)
        actions.context_click(first_email).perform()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Download']"
        ))).click()

        wait.until(EC.element_to_be_clickable((
            By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Download as EML']"
        ))).click()

        emails = driver.find_elements(
            By.XPATH,
            "//div[@aria-label='Message list']//div[@role='option']"
        )

        if emails:
            actions.context_click(emails[0]).perform()

            wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Move']"
            ))).click()

            wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[contains(@class,'fui-MenuItem__content') and text()='Archive']"
            ))).click()

        print("Processed one email.")

    except StaleElementReferenceException:
        continue

    except Exception as e:
        print("Unexpected error:", e)
        time.sleep(3)