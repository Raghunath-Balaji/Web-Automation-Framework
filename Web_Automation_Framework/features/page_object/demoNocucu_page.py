from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver

class demoNocucu_page:

    def __init__(self, driver: WebDriver):
        self.driver = driver

    @property
    def input_name(self):
        return self.driver.find_element(By.ID, "name")

    @property
    def input_email(self):
        return self.driver.find_element(By.ID, "email")

    @property
    def input_phone(self):
        return self.driver.find_element(By.ID, "phone")

    @property
    def input_misc(self):
        return self.driver.find_element(By.ID, "input1")

    @property
    def gender_male(self):
        return self.driver.find_element(By.ID, "male")