import smtplib
from datetime import datetime
from email.message import EmailMessage
from time import sleep

import undetected_chromedriver as uc
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

login_email = ''
login_passwd = ''
url_link = ''
hourlyTime = 8


def email_notify(subject, body):
    msg = EmailMessage()
    msg.set_content(body)

    msg['subject'] = subject
    msg['to'] = ''

    user = ''
    msg['from'] = user

    password = ''

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(user, password)
    server.send_message(msg)

    server.quit()


def time_entry():
    timeStamp = datetime.today().strftime('%Y-%m-%d')

    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.XPATH, f"""//*[@id="{timeStamp}"]/div[1]/div[1]/div"""))).click()

    sleep(3)

    # activate  time fields
    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.XPATH, '//*[@id="ext-gen318"]/div/table/tbody/tr[1]/td[2]'))).click()
    sleep(1)
    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.XPATH, f"""//*[@id="ext-gen318"]/div/table/tbody/tr[1]/td[2]/div"""))).click()
    sleep(1)

    # fill time monday - friday
    for i in range(2, 7):
        time_field_activate = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.XPATH, f"""//*[@id="timefield{i}"]""")))
        time_field_activate.click()
        sleep(1)

        time_field_activate.send_keys(hourlyTime)
        sleep(1)

    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.CSS_SELECTOR, '#ext-gen179'))).click()
    sleep(3)

    # save time sheet fields
    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.CSS_SELECTOR, '#ext-gen525'))).click()
    sleep(3)

    # confirm time sheet
    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.CSS_SELECTOR, '#ext-gen567'))).click()

    sleep(8)


def main():
    try:
        driver.get(url_link)

        email = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div['
                       '2]/div[2]/div/input[1]')))
        email.click()
        sleep(1)
        email.send_keys(login_email)
        sleep(1)

        nxtButton = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.CSS_SELECTOR, '#idSIButton9')))
        nxtButton.click()
        sleep(3)

        password = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div['
                       '3]/div/div[2]/input')))
        password.click()
        sleep(1)
        password.send_keys(login_passwd)
        sleep(3)

        nxtButton = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.CSS_SELECTOR, '#idSIButton9')))
        nxtButton.click()
        sleep(5)

        WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.CSS_SELECTOR, '#idBtn_Back'))).click()
        sleep(5)

        time_entry()

        email_notify("Time Sheet", "Time sheet was successfully submitted.")

    except TimeoutException or ElementClickInterceptedException:

        email_notify("[ALERT] Time Sheet", "Time sheet was NOT submitted.")


if __name__ == '__main__':
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--window-size=1920,1080')
    driver = uc.Chrome(options=chrome_options)

    main()
