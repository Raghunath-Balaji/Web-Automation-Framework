from selenium.webdriver.remote.webdriver import WebDriver
from features.page_object.demoNocucu_page import demoNocucu_page


class PageObjectManager:


    def __init__(self, driver: WebDriver):
        self.driver = driver
        self._sign_in_page = None
        self._welcome_page = None
        self._demoNocucu_page = None


    def get_demoNocucu(self):
        if self._demoNocucu_page is None:
            self._demoNocucu_page = demoNocucu_page(self.driver)
        return self._demoNocucu_page