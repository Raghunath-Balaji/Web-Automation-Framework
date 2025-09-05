import time
from cmath import phase

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from features.steps.test import HEADER_SUBSTRING
from utilities.Utilities import Utilities
from utilities.ExcelReader import ExcelReader
from utilities.DriverEngine import DriverEngine
from utilities.PageObjectManager import PageObjectManager
from features.page_object.demoNocucu_page import demoNocucu_page
from behave import given, then
from utilities.ExcelReader import ExcelReader

FILE_PATH = r"C:\Users\raghunath_b\PycharmProjects\Web_Automation_Framework\Sample Book.xlsx"

@given('the URL is opened')
def step_open_url(context):
    driver_engine = DriverEngine()

    context.driver = driver_engine.get_chrome_driver()
    context.page = demoNocucu_page(driver=context.driver)
    context.utils = Utilities()
    context.xlread = ExcelReader()
    context.driver.get("https://testautomationpractice.blogspot.com/")

@then('Click gender')
def step_then(context):
    context.utils.tap_element_simple(context.driver, context.page.gender_male, 100)

@then('Enter your full name {string}')
def step_then(context, string):
    HEADER_SUBSTRING = "Name"
    name = context.xlread.get_data_as_string(FILE_PATH, string[1:-1], HEADER_SUBSTRING)
    context.page.input_name.send_keys(name)

@then('Enter your Email {string}')
def step_then(context, string):
    HEADER_SUBSTRING = "Email"
    email = context.xlread.get_data_as_string(FILE_PATH, string[1:-1], HEADER_SUBSTRING)
    mail = context.page.input_email
    mail.send_keys(email)

@then('Enter your phonenumber {string}')
def step_then(context, string):
    HEADER_SUBSTRING = "Phone"
    txt = context.xlread.get_data_as_string(FILE_PATH, string[1:-1], HEADER_SUBSTRING)
    phone = context.page.input_phone
    phone.send_keys(txt)

@then('Scroll until visible and enter name in section 1')
def step_scroll_and_enter_section_name(context):
    HEADER_SUBSTRING = "Text"
    txt = context.xlread.get_data_as_string(FILE_PATH, "Sample sheet 1", HEADER_SUBSTRING)
    context.utils.web_scroll_to_element(context.page.input_misc)
    input_misc = context.page.input_misc
    context.utils.tap_element_simple(context.driver, input_misc, 10)
    input_misc.send_keys(txt)

@then('close the driver')
def step_close_driver(context):
    context.driver.close()








