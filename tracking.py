import utils
import os
import numpy as np
import re
import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

def find_lasers(emails):
    lasers = np.array([])
    serial_numbers = np.array([])

    # Find laser serial numbers
    for email in emails:
        # Search email subject lines for all types of laser serial numbers
        subject = email.subject
        serial_numbers = re.findall(r"[LPS]\d{4}[A-Z]\d{4}", subject)
        other_serial_numbers = re.findall(r"[S]\d{1}[C]\d{2}[E]\d{4}", subject)
        tgbl_lasers = re.findall(r"SH\d{8}", subject)
        
        if serial_numbers or other_serial_numbers or tgbl_lasers:
            lasers = np.append(lasers, email)

    return lasers

def get_sn(email):
    # Getting serial numbers seems like it would be an easy task, but TGBL lasers are special.
    # They come from a TRUMPF subsidiary in the UK, and they don't have serial numbers
    # until they come here. We have to check the laser type.
    subject = email.subject
    serial_numbers = re.findall(r"[LPS]\d{4}[A-Z]\d{4}|[S]\d{1}[C]\d{2}[E]\d{4}", subject)
    tgbl_lasers = re.findall(r"SH\d{8}|UKS\d{5}", subject)
    if serial_numbers:
        sn = serial_numbers
    # If it's a TGBL laser, we need to check if it has an attachment containing
    # the batch number. That batch number will be used as the new serial numeber
    elif tgbl_lasers:
        if email.Attachments.Count > 0:
            attachments = email.Attachments
            for attachment in attachments:
                content = get_pdf_content(attachment)
                batch_num = re.findall(r"Batch No. : \d{6}", content)
                batch_num = ', '.join(batch_num)
                batch_num = re.findall(r"\d{6}", batch_num)
                sn = ['XPI' + str(num) for num in batch_num]
    # Not all emails for TGBL lasers will have that attachment with the serial number
    # So it's possible that the serial number will be empty. We ignore these until later.
    else:
        sn = []
    
    return sn

def find_carrier(email):
    carrier = []
    # For UPS, we check the domain of the sender email. Some emails pertaining to
    # lasers shipped through UPS are not from UPS senders. These will be ignored until later.
    sender_email = email.SenderEmailAddress
    if "ups.com" in sender_email:
        carrier = "UPS"
    # FedEx does not send emails to this inbox directly, so we need another approach.
    # Only our tgbl lasers are shipped through FedEx, so if it is a tgbl laser,
    # we say its carrier is FedEx
    subject = email.Subject
    tgbl_laser = re.findall(r"SH\d{8}|UKS\d{5}", subject)
    if tgbl_laser:
        carrier = "FedEx"
    return carrier

def find_fedex_tracking(email):
    tracking_num = []
    # Sometimes FedEx just puts the tracking number in the body of the email.
    # We check there first.
    if email.Attachments.Count == 0:
        tracking_num = re.findall(r"\d{12}", tracking_num)
        if tracking_num:
            return tracking_num
    
    # It may also be found in an attched commercial invoice along with some
    # other files. They do not have a standardized naming convention,
    # which makes it difficult to systematically find the right file.
    # We instead search them all until we find what we need.
    else:
        attachments = email.Attachments
        for attachment in attachments:
            content = get_pdf_content(attachment)
            # There is more than one instance of 12 consecutive digits, so 
            # we find the instance with "Tracking No:    " preceeding it
            # If they include more than one tracking number, we join the 
            # list so we can then just get the numbers
            tracking_num = re.findall(r"Tracking No :   \d{12}", content)
            tracking_num = ', '.join(tracking_num) # This joins the list into one string
            tracking_num = re.findall(r"\d{12}", tracking_num) # Then search again for just the numbers
            if tracking_num:
                return tracking_num
    return tracking_num

def find_ups_tracking(email):
    hawb_no = []
    # Sometimes UPS includes the housebill number in the body of the email.
    # We check there first
    if email.Attachments.Count == 0:
        hawb_no = re.findall(r"\d{10}", email.Body)
        if hawb_no:
            return hawb_no
    # It could also be in an attachment. We read the pdf to get it. There are
    # no other instances of 10 consecutive digits in their files.
    else:
        attachments = email.Attachments
        for attachment in attachments:
            content = get_pdf_content(attachment)
            hawb_no = re.findall(r"\d{10}", content)
            if hawb_no:
                return hawb_no
    return hawb_no

def get_fedex_link(tracking_num):
    # FedEx redirects Selenium webdriver to an error page, but when you input that error link into your
    # normal browser, it redirects you to the right page. Since we use a hyperlink, the error link
    # is not noticable.

    # First we set up our webdriver
    new_directory = "C:\Temp"
    os.chdir(new_directory)
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.67'
    edge_driver_path = os.path.join(os.getcwd(), 'msedgedriver.exe')
    edge_service = Service(edge_driver_path)
    edge_options = Options()
    edge_options.add_argument(f'user-agent={user_agent}')
    browser = webdriver.Edge(service=edge_service, options=edge_options)

    # Browser is now set up. Go to the FedEx site
    browser.get('https://www.fedex.com/en-us/tracking.html')

    # Wait a bit for the page to load, then find the html element for inputting field based on its ID
    WebDriverWait(browser, 15)
    input_field = browser.find_element(By.CSS_SELECTOR, '.fdx-c-form-group__input')
    input_field.clear()
    input_field.send_keys(tracking_num)
    # And find the html element for submitting the form
    submit_button = browser.find_element(By.XPATH, "//button[@class='fdx-c-button fdx-c-button--primary fdx-c-button--responsive fdx-u-flex-justify-content--center']")
    submit_button.click()

    # wait for URL to change and get url
    WebDriverWait(browser, 15).until(EC.url_changes(browser.current_url))
    redirected_url = browser.current_url
    browser.quit()

    return redirected_url

def get_ups_link(tracking_num):
    # Selenium not compatible with UPS site, so this was my workaround
    link = 'https://www.ups.com/track?loc=en_US&requester=QUIC&tracknum=' + tracking_num + '/trackdetails'
    return link
