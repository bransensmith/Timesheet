import datetime
import email
import imaplib
import smtplib
from email.message import EmailMessage
from time import sleep
from datetime import datetime as dt

import undetected_chromedriver as uc
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

gmail_username = ''
gmail_app_pass = ''

imap = imaplib.IMAP4_SSL('imap.gmail.com')
imap.login(gmail_username, gmail_app_pass)


def email_evoke(subject_line):
    str(subject_line)
    email_payload = ''
    time_sheet_provoke = ['Continue', 'True', 'Yes', 'Normal']

    imap.select('INBOX')

    status, info = imap.search(None, f"""(FROM "" SUBJECT "{subject_line}" UNSEEN)""")

    mail_ids = []

    for block in info:
        mail_ids += block.split()

        if mail_ids:
            # Get most recent Email from List filter
            mail_ids = mail_ids[-1]

            status, info = imap.fetch(mail_ids, '(RFC822)')

            for response_part in info:

                if isinstance(response_part, tuple):

                    message = email.message_from_bytes(response_part[1])

                    if message.is_multipart():
                        mail_content = ''

                        for part in message.get_payload():

                            if part.get_content_type() == 'text/plain':
                                mail_content += part.get_payload()
                    else:

                        mail_content = message.get_payload()

                    email_payload = mail_content
    # if string is empty
    if not email_payload:
        return None

    # if string has content
    elif email_payload:

        if any(word in email_payload for word in time_sheet_provoke):
            return True

        else:
            return False


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
    hourlyTime = 8

    timeStamp = dt.today().strftime('%Y-%m-%d')

    # must run on saturdays
    WebDriverWait(driver, 10).until(ec.presence_of_element_located(
        (By.XPATH, f"""//*[@id="{timeStamp}"]/div[1]/div[1]/div"""))).click()

    sleep(5)

    # Activate time fields

    for i in range(2):
        WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.XPATH, '//*[@id="ext-gen319"]/div/table/tbody/tr[1]/td[2]/div'))).click()
        sleep(1)

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

    sleep(10)


def main():
    login_email = ''
    login_passwd = ''

    url_link = ''

    try:
        driver.get(url_link)

        email_name = WebDriverWait(driver, 10).until(ec.presence_of_element_located(
            (By.XPATH, '/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div['
                       '2]/div[2]/div/input[1]')))
        email_name.click()
        sleep(1)
        email_name.send_keys(login_email)
        sleep(3)

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

        try:
            # try clicking send verify by call box
            WebDriverWait(driver, 10).until(ec.presence_of_element_located(
                (By.XPATH, '//*[@id="idDiv_SAOTCS_Proofs"]/div[2]/div/div/div[2]'))).click()

            sleep(3)

            WebDriverWait(driver, 25).until(ec.presence_of_element_located(
                (By.CSS_SELECTOR, '#idBtn_Back'))).click()

        except TimeoutException:

            WebDriverWait(driver, 10).until(ec.presence_of_element_located(
                (By.CSS_SELECTOR, '#idBtn_Back'))).click()

        sleep(5)

        time_entry()

        email_notify("Time Sheet", "[UPDATE]" + "\n" + "Time sheet was successfully submitted.")

    except TimeoutException or ElementClickInterceptedException:

        email_notify("Time Sheet", "[ALERT]" + "\n" + "Time sheet was NOT submitted.")


if __name__ == '__main__':
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--window-size=1920,1080')
    driver = uc.Chrome(options=chrome_options)

    # wait time 30 minutes for response
    time_out_limit = 30

    sleep(1)

    # check for start email
    email_status_first = email_evoke("Time Sheet")

    if email_status_first is None:

        sleep(15)

        email_notify("Time Sheet",
                     "[ALERT]" + "\n" + "Input required to process or decline auto-processing. Please "
                                        "respond to "
                                        "this email within " + str(time_out_limit) + " minutes; Failure to "
                                                                                     "respond "
                                                                                     "will result in "
                                                                                     "cancellation.")

        endTime = datetime.datetime.now() + datetime.timedelta(minutes=time_out_limit)
        while True:

            email_status_second = email_evoke("Time Sheet")

            if email_status_second is False:
                email_notify("Time Sheet",
                             "[NOTICE]" + "\n" + "Cancel request received; Auto-fill has been suspended.")
                break

            elif email_status_second is True:
                main()
                break

            elif datetime.datetime.now() >= endTime:
                email_notify("Time Sheet", "[NOTICE]" + "\n" + "Time limit has exceeded " + str(
                    time_out_limit) + " minutes; Auto-fill has been "
                                      "suspended.")
                break
            else:
                sleep(15)

    elif email_status_first is False:
        email_notify("Time Sheet", "[NOTICE]" + "\n" + "Cancel request received; Auto-fill has been "
                                                       "suspended.")

    else:
        main()

    imap.close()
    imap.logout()
