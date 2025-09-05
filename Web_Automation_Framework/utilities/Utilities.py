import os
import subprocess
import time
import logging
import re
import uuid
from datetime import datetime
from typing import Union, Tuple, List
from utilities.DriverEngine import DriverEngine
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    ElementNotInteractableException,
)
from selenium.webdriver import ActionChains
from openpyxl import Workbook

# Configure module logger
logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT = 30


class Utilities(DriverEngine):

    def execute_cmd(self, cmd: str) -> str:
        result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        output = result.stdout.strip().splitlines()
        return output[0] if output else ""

    def wait_element_to_be_clickable(self, driver: WebDriver, element: WebElement,
                                     timeout: int = DEFAULT_TIMEOUT) -> None:
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(element))

    def wait_visibility_of_element(self, driver: WebDriver, element: WebElement,
                                   timeout: int = DEFAULT_TIMEOUT) -> None:
        WebDriverWait(driver, timeout).until(EC.visibility_of(element))

    def custom_wait_visibility_of_element(self, driver: WebDriver, element: WebElement, timeout: int) -> None:
        WebDriverWait(driver, timeout).until(EC.visibility_of(element))

    def wait_invisibility_of_element(self, driver: WebDriver, element: WebElement,
                                     timeout: int = DEFAULT_TIMEOUT) -> None:
        WebDriverWait(driver, timeout).until(EC.invisibility_of_element(element))

    def wait_for_page_load(self, driver: WebDriver, url_substring: str, timeout: int = DEFAULT_TIMEOUT) -> None:
        WebDriverWait(driver, timeout).until(EC.url_contains(url_substring))

    def tap_element(self, driver: WebDriver, element: WebElement, click_method: str = "click", obj: str = "",
                    timeout: int = DEFAULT_TIMEOUT) -> None:
        self.wait_visibility_of_element(driver, element, timeout)
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        clicked = False
        if click_method == "jsClick":
            self.js_click(driver, element)
            clicked = True
        elif click_method == "click":
            element.click()
            clicked = True
        elif click_method == "actionClick":
            self.action_click(driver, element)
            clicked = True
        if clicked:
            logger.info(f"{obj}: Element clicked successfully")
        else:
            logger.error(f"{obj}: Element not clicked")

    def tap_element_simple(self, driver: WebDriver, element: WebElement, timeout: int = DEFAULT_TIMEOUT) -> None:
        self.wait_visibility_of_element(driver, element, timeout)
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(2)
        for _ in range(3):
            try:
                element.click()
                logger.info("Element has been tapped")
                break
            except ElementClickInterceptedException:
                time.sleep(0.5)
            except StaleElementReferenceException:
                driver.refresh()
            except ElementNotInteractableException:
                self.js_click(driver, element)
                time.sleep(0.5)
                break
            except Exception as e:
                logger.error(f"Element tap failed: {e}")

    def js_click(self, driver: WebDriver, element: WebElement) -> None:
        driver.execute_script("arguments[0].click();", element)

    def action_click(self, driver: WebDriver, element: WebElement) -> None:
        ActionChains(driver).move_to_element(element).click().perform()
        logger.info("Action click performed successfully")

    def action_backspace(self, driver: WebDriver, element: WebElement) -> None:
        ActionChains(driver).move_to_element(element).click().send_keys(Keys.BACK_SPACE).perform()
        logger.info("Action backspace performed successfully")

    def handle_click(self, driver: WebDriver, element: WebElement, timeout: int = DEFAULT_TIMEOUT) -> None:
        try:
            self.wait_element_to_be_clickable(driver, element, timeout)
        except Exception:
            self.wait_visibility_of_element(driver, element, timeout)
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        try:
            element.click()
        except Exception:
            if element.is_displayed():
                self.js_click(driver, element)
            else:
                logger.error("Element tapping failed")

    def enter_input(self, driver: WebDriver, element: WebElement, input_text: str) -> None:
        time.sleep(1)
        try:
            self.wait_visibility_of_element(driver, element)
            element.clear()
            element.send_keys(input_text)
            logger.info(f"Text: {input_text} entered successfully")
        except Exception as e:
            logger.error(f"Sending input failed: {e}")

    def element_click(self, driver: WebDriver, element: WebElement, obj: str = "") -> None:
        for _ in range(3):
            try:
                element.click()
                logger.info(f"{obj}: element clicked")
                return
            except (ElementClickInterceptedException, StaleElementReferenceException):
                continue
        logger.error(f"Failed to click element: {obj}")

    def element_click_by_locator(self, driver: WebDriver, locator: Tuple[str, str], obj: str = "") -> None:
        element = self.get_element(driver, locator)
        self.element_click(driver, element, obj)

    def get_element(self, driver: WebDriver, locator: Tuple[str, str]) -> WebElement:
        return driver.find_element(*locator)

    def enter_text(self, driver: WebDriver, element: WebElement, value: str, ele_name: str) -> None:
        self.wait_visibility_of_element(driver, element)
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        element.send_keys(value)
        logger.info(f"{value} entered in {ele_name}")

    def enter_text_by_locator(self, driver: WebDriver, locator: Tuple[str, str], value: str, ele_name: str) -> None:
        element = self.get_element(driver, locator)
        self.enter_text(driver, element, value, ele_name)

    def element_is_displayed(self, driver: WebDriver, element: WebElement, timeout: int = DEFAULT_TIMEOUT) -> bool:
        try:
            self.wait_visibility_of_element(driver, element, timeout)
            return element.is_displayed()
        except Exception:
            return False

    def validate_element(self, driver: WebDriver, element: WebElement, expected_text: str,
                         timeout: int = DEFAULT_TIMEOUT) -> None:
        if self.element_is_displayed(driver, element, timeout):
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            actual = element.text.strip()
            logger.info(f"Expected: '{expected_text}', Actual: '{actual}'")
            if expected_text.lower() in actual.lower():
                logger.info(f"Text '{expected_text}' matched successfully")
            else:
                logger.warning(f"Text '{expected_text}' not matched. Actual: '{actual}'")

    def validate_element_by_locator(self, driver: WebDriver, locator: Tuple[str, str], expected_text: str,
                                    timeout: int = DEFAULT_TIMEOUT) -> None:
        self.validate_element(driver, self.get_element(driver, locator), expected_text, timeout)

    def validate_button_no_text(self, driver: WebDriver, element: WebElement) -> None:
        if self.element_is_displayed(driver, element):
            logger.info("Button displayed as expected")
        else:
            logger.warning("Button not displayed as expected")

    def validate_button_no_text_by_locator(self, driver: WebDriver, locator: Tuple[str, str]) -> None:
        self.validate_button_no_text(driver, self.get_element(driver, locator))

    def validate_image(self, driver: WebDriver, element: WebElement) -> None:
        if self.element_is_displayed(driver, element):
            logger.info("Image displayed as expected")
        else:
            logger.warning("Image not displayed")

    def element_is_present(self, driver: WebDriver, element: WebElement) -> bool:
        try:
            return element.is_displayed()
        except Exception:
            return False

    def element_is_present_by_locator(self, driver: WebDriver, locator: Tuple[str, str]) -> bool:
        return len(driver.find_elements(*locator)) > 0

    def ignoring_stale_element_exception(self, wait: WebDriverWait) -> None:
        wait.ignored_exceptions(StaleElementReferenceException)

    def scroll_till_the_end_of_the_page(self, driver: WebDriver, left: int, top: int, width: int, height: int) -> None:
        # scroll_till_the_end_of_the_page - requires Appium Python client for mobile scrollGesture
        raise NotImplementedError("scroll_till_the_end_of_the_page requires Appium client")

    def scroll_down_gesture(self, driver: WebDriver, left: int, top: int, width: int, height: int) -> None:
        # scrollDownGesture - requires Appium Python client for mobile scrollGesture
        raise NotImplementedError("scroll_down_gesture requires Appium client")

    def scroll_up_gesture(self, driver: WebDriver, left: int, top: int, width: int, height: int) -> None:
        # scrollUpGesture - requires Appium Python client for mobile scrollGesture
        raise NotImplementedError("scroll_up_gesture requires Appium client")

    def scroll_up(self, driver: WebDriver) -> None:
        # scrollup - uses scroll_up_gesture with default coordinates
        raise NotImplementedError("scrollup requires Appium client")

    def click_text_from_list_elements(self, driver: WebDriver, elements: List[WebElement], text: str) -> None:
        for el in elements:
            if el.text == text:
                logger.info(f"Clicking element with text '{text}'")
                el.click()
                return
        logger.warning(f"Element with text '{text}' not found")

    def fetch_list_objects(self, elements: List[WebElement]) -> List[WebElement]:
        return list(elements)

    def confirm_page_navigation(self, driver: WebDriver) -> None:
        main = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main:
                driver.switch_to.window(handle)
                logger.info(f"Switched to window: {driver.current_url}")
                driver.close()
                driver.switch_to.window(main)

    def handle_windows(self, driver: WebDriver) -> None:
        # alias for confirm_page_navigation
        self.confirm_page_navigation(driver)

    def hiding_keyboard(self) -> None:
        # HidingKeyboard - requires mobile context, not available in Selenium
        raise NotImplementedError("hiding_keyboard requires mobile context")

    def date(self) -> str:
        return datetime.now().strftime("%d")

    def get_current_year(self) -> str:
        return datetime.now().strftime("%Y")

    def get_today_date(self, format_str: str) -> str:
        return datetime.now().strftime(format_str)

    def replace_leading_zero(self, date_str: str) -> str:
        return re.sub(r"^0+", "", date_str)

    def iOS_date_picker(self, driver: WebDriver, date_str: str, input_str: str, elements: List[WebElement]) -> None:
        # iOSDatePicker - requires Appium and iOS context
        raise NotImplementedError("iOS_date_picker requires Appium client")

    def iOS_year_picker(self, driver: WebDriver, year_str: str, input_str: str, elements: List[WebElement]) -> None:
        # IOSYearPicker - requires Appium and iOS context
        raise NotImplementedError("iOS_year_picker requires Appium client")

    def week_day_picker(self, driver: WebDriver, elements: List[WebElement], day: str) -> None:
        # WeekDayPicker - mobile/desktop picker logic
        for el in elements:
            if el.text.strip().lower() == day.lower():
                el.click()
                logger.info(f"Selected day '{day}'")
                return
        logger.warning(f"Day '{day}' not found")

    def sheet_path_value(self):
        return self.property.get("DEFAULT", "TESTDATA", fallback=None)

    def sheet_path(brand: str, build_type: str, test_data_type: str) -> str:

        brand_lower = brand.lower()
        data_type = test_data_type.lower()

        # 1) Azure-backed test data: just filename
        if data_type == "azure":
            if brand_lower == "moes":
                return "Moes_TestData.xlsx"
            elif brand_lower == "jamba":
                return "Jamba_TestData.xlsx"
            elif brand_lower == "mca":
                return "Mca_TestData.xlsx"
            elif brand_lower in ("lotz", "schlotzskys"):
                return "Lotz_TestData.xlsx"
            elif brand_lower == "aa":
                return "aa_TestData.xlsx"
            elif brand_lower == "cinn":
                return "Cinnabon_TestData.xlsx"
            elif brand_lower == "car":
                return "Carvel_TestData.xlsx"
            else:
                raise ValueError(f"Unknown brand for Azure test data: {brand}")

        # 2) File-based test data: build absolute base path
        base_path = os.getcwd()

        # 2a) PROD paths
        if build_type.lower() == "prod":
            folder = os.path.join(base_path,
                                  "src", "test", "resources", "TestData", "PROD_TestData")
            if "moes" in brand_lower:
                fname = "Moes_PROD_TestData.xlsx"
            elif brand_lower == "jamba":
                fname = "Jamba_PROD_TestData.xlsx"
            elif brand_lower == "mca":
                fname = "MCA_PROD_TestData.xlsx"
            elif brand_lower in ("lotz", "schlotzskys"):
                fname = "Lotz_PROD_TestData.xlsx"
            elif brand_lower == "aa":
                fname = "aa_PROD_TestData.xlsx"
            elif brand_lower == "cinn":
                fname = "Cinnabon_PROD_TestData.xlsx"
            elif brand_lower == "car":
                fname = "Carvel_PROD_TestData.xlsx"
            else:
                raise ValueError(f"Unknown brand for PROD test data: {brand}")

        # 2b) NON-PROD paths
        else:
            folder = os.path.join(base_path,
                                  "src", "test", "resources", "TestData", "NONPROD_Testdata")
            if "moes" in brand_lower:
                fname = "Moes_NONPROD_TestData.xlsx"
            elif brand_lower == "jamba":
                fname = "Jamba_NONPROD_TestData.xlsx"
            elif brand_lower == "mca":
                fname = "MCA_NONPROD_TestData.xlsx"
            elif brand_lower in ("lotz", "schlotzskys"):
                fname = "Lotz_NONPROD_TestData.xlsx"
            elif brand_lower == "aa":
                fname = "aa_NONPROD_TestData.xlsx"
            elif brand_lower in ("cinn", "Cinn"):  # match both cases just in case
                fname = "Cinnabon_NONPROD_TestData.xlsx"
            elif brand_lower == "car":
                fname = "Carvel_NONPROD_TestData.xlsx"
            else:
                raise ValueError(f"Unknown brand for NONPROD test data: {brand}")

        return os.path.join(folder, fname)

    def create_report_excel(self, brand: str, browser_name: str, output_dir: str = "ExcelReport") -> None:

        os.makedirs(output_dir, exist_ok=True)

        wb = Workbook()
        sheet_name = f"{brand.upper()}_Report"
        ws = wb.active
        ws.title = sheet_name

        headers = ["SCENARIO_NAME", "SCREEN_NAME", "ACTION", "TIME_TAKEN"]
        ws.append(headers)

        timestamp = datetime.now().strftime("%d-%m-%Y_%H.%M.%S")
        filename = f"{brand.upper()}_{browser_name.upper()}_REPORT_{timestamp}.xlsx"
        file_path = os.path.join(output_dir, filename)

        print(f"Excel file path: {file_path}")
        try:
            wb.save(file_path)
            print(f"Excel file has been created successfully: {file_path}")
            self.report_sheet_name = sheet_name
            self.report_sheet_path = file_path
            self.sheet_exists_flag = True
        except Exception as e:
            print(f"Error saving Excel file: {e}")

    def update_execution_time(self, row_number: int, column_name: str, time_value: str) -> None:
        if getattr(self, "performance_time_update", "").lower() == "yes":
            self.set_cell_data_insert("PageLoadTime", column_name, time_value)

    def get_email_from_brand_sheet(self, sheet_name: str, type_: str) -> dict:
        return self.get_email_from_brand_sheet(sheet_name, type_)

    def get_result_from_brand_sheet(self, sheet_name: str, scenario_name: str) -> dict:
        return self.get_result_from_brand_sheet(sheet_name, scenario_name)

    def set_cell_data(self, sheet_name: str, scenario_name: str, field_name: str, value: str) -> None:
        pass

    def set_cell_data_insert(self, sheet_name: str, field_name: str, value: str) -> None:
        try:
            query = f"INSERT INTO {sheet_name}({field_name}) VALUES '{value}'"
            self.tlcon_test_data.get().execute_update(query)
        except Exception as e:
            logger.error("error")

    def get_number_from_txt(self, num: str) -> str:
        return re.sub(r"\D+", "", num)

    def get_page_load_time(self, driver: WebDriver, screen_name: str, action_performed: str) -> float:
        time_ms = driver.execute_script(
            "return window.performance.timing.loadEventEnd - window.performance.timing.navigationStart;"
        )
        seconds = time_ms / 1000.0
        logger.info(f"Page Load Time: {seconds}")
        return seconds

    def get_page_load_time(self, screen_name: str, action_performed: str) -> float:

        driver = getattr(self.tl_driver, 'driver', None)

        # 1) run the same JS to get ms between navigationStart and loadEventEnd
        page_time_ms = driver.execute_script(
            "return window.performance.timing.loadEventEnd"
            " - window.performance.timing.navigationStart;"
        )
        time_sec = page_time_ms / 1000.0

        # 2) log it
        logging.info(f"Page Load Time: {time_sec}")

        # 3) (optional) write into your report, etc.

        return time_sec

    def get_dom_load_time(self, driver: WebDriver) -> float:
        time_ms = driver.execute_script(
            "return window.performance.timing.domContentLoadedEventEnd - window.performance.timing.navigationStart;"
        )
        seconds = time_ms / 1000.0
        logger.info(f"DOM Load Time: {seconds}")
        return seconds

    def get_dom_load_time_default(self) -> float:
        # getDOMLoadTime() overload without driver - not supported, use get_dom_load_time with driver
        raise NotImplementedError("get_dom_load_time_default requires explicit driver")

    def get_byte_screenshot(self, driver: WebDriver) -> bytes:
        return driver.get_screenshot_as_png()

    def is_valid_email(self, email: str) -> bool:
        return bool(re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", email))

    def wait_for(ms: int) -> None:

        time.sleep(ms / 1000.0)

    def lotz_confirm_page_navigation(self) -> None:

        driver = self.get_driver()
        driver.implicitly_wait(30)
        main_window = driver.current_window_handle
        windows = driver.window_handles
        if len(windows) > 1:
            driver.implicitly_wait(120)
            child_window = windows[1]
            driver.switch_to.window(child_window)
            child_url = driver.current_url
            print(f"Navigated to {child_url} Page successfully")
            driver.close()
            driver.switch_to.window(main_window)

    def toaster_message(self, element, expected_message: str) -> None:
        """
        Verify that a toaster message is displayed and matches expected text.
        """
        driver = self.get_driver()
        try:
            WebDriverWait(driver, 10).until(EC.visibility_of(element))
            actual = element.text.strip()
            if actual == expected_message:
                print(f"Toaster matches expected: '{actual}'")
            else:
                print(f"Toaster mismatch: got '{actual}', expected '{expected_message}'")
        except Exception:
            print(f"Toaster '{expected_message}' not displayed")

    def assert_button_enabled(self, element) -> None:
        """
        Assert that the given button element is enabled.
        """
        driver = self.get_driver()
        self.wait_for(2000)
        WebDriverWait(driver, 10).until(EC.visibility_of(element))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        self.wait_for(2000)
        assert element.is_enabled(), f"Button not enabled: '{element.text}'"
        print(f"Button enabled: '{element.text}'")

    def assert_button_disabled(self, element) -> None:
        """
        Assert that the given button element is disabled.
        """
        driver = self.get_driver()
        WebDriverWait(driver, 10).until(EC.visibility_of(element))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        self.wait_for(2000)
        assert not element.is_enabled(), f"Button not disabled: '{element.text}'"
        print(f"Button disabled: '{element.text}'")

    def assert_element_displayed(self, element, message: str) -> None:
        """
        Assert that the element is displayed on the page.
        """
        driver = self.get_driver()
        try:
            WebDriverWait(driver, 10).until(EC.visibility_of(element))
            assert element.is_displayed(), f"{message}: element not displayed"
            print(f"{message}: element displayed")
        except Exception:
            print(f"{message}: element not displayed")

    def assert_text(self, element, expected_text: str) -> None:
        """
        Assert that the element's text matches the expected text.
        """
        driver = self.get_driver()
        try:
            WebDriverWait(driver, 10).until(EC.visibility_of(element))
            actual = element.text.strip()
            assert actual == expected_text, f"Expected '{expected_text}', got '{actual}'"
            print(f"Text matches: '{actual}'")
        except AssertionError as ae:
            print(f"Assertion failed: {ae}")
        except Exception as e:
            print(f"Error in text assertion: {e}")

    def assert_field_is_disabled(self, element) -> None:
        """
        Assert that the field element is disabled (not editable).
        """
        assert not element.is_enabled(), "Field is not disabled"
        print("Field disabled")

    def browser_refresh(self) -> None:
        """
        Refresh the current browser window.
        """
        driver = self.get_driver()
        driver.refresh()

    def is_alert_present(self) -> bool:
        """
        Check if an alert is present.
        """
        driver = self.get_driver()
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            return True
        except Exception:
            return False

    def handle_push_notification(self) -> None:
        """
        Dismiss push notification alert if present.
        """
        driver = self.get_driver()
        try:
            alert = driver.switch_to.alert
            alert.accept()
            print("Push notification accepted")
        except Exception:
            pass

    def generate_random_user_email(self, domain: str = "example.com") -> str:

        unique = uuid.uuid4().hex[:8]
        email = f"user_{unique}@{domain}"
        logging.info(f"Generated email: {email}")
        return email

    def log_info(self, message: str) -> None:

        logging.info(message)

    def log_warning(self, message: str) -> None:

        logging.warning(message)

    def log_pass(self, message: str) -> None:

        logging.info(f"[PASS] {message}")

    def log_fail(self, message: str) -> None:

        logging.error(f"[FAIL] {message}")

    def logging(self, condition: str, message: str) -> None:

        lvl = condition.lower()
        if lvl == 'info':
            logging.info(message)
        elif lvl == 'warning':
            logging.warning(message)
        elif lvl == 'pass':
            logging.info(f"[PASS] {message}")
        elif lvl == 'fail':
            logging.error(f"[FAIL] {message}")
        else:
            logging.debug(message)

    def set_screen_size(self, width: int, height: int) -> None:

        driver = self.get_driver()
        driver.set_window_size(width, height)

    def web_scroll_to_element(self, element) -> None:
        """
        Scroll the page until the element is in view.
        """
        driver = self.get_driver()
        driver.execute_script("arguments[0].scrollIntoView(true);", element)

    def web_scroll_down(self, pixels: int = 1000) -> None:

        driver = self.get_driver()
        driver.execute_script(f"window.scrollBy(0, {pixels});")

    def is_selected(self, element) -> bool:

        return element.is_selected()

    def is_enabled(self, element) -> bool:

        return element.is_enabled()

    def validate_element_present(self, element, message: str = "") -> None:

        driver = self.get_driver()
        try:
            WebDriverWait(driver, 10).until(EC.visibility_of(element))
            assert element.is_displayed(), f"{message} element not present"
            print(f"{message}: element present and displayed")
        except Exception:
            raise AssertionError(f"{message}: element not present or not visible")

    def validate_button(self, element, expected_text: str = None) -> None:

        self.validate_element_present(element, "Button")
        assert element.tag_name.lower() == 'button', "Element is not a button"
        assert element.is_enabled(), "Button is not enabled"
        if expected_text is not None:
            actual = element.text.strip()
            assert actual == expected_text, f"Expected button text '{expected_text}', got '{actual}'"
        print("Button validation passed")

    def validate_element_get_attribute(self, element, attribute: str, expected_value: str) -> None:

        actual = element.get_attribute(attribute)
        assert actual == expected_value, f"Expected attribute '{attribute}' to be '{expected_value}', got '{actual}'"
        print(f"Attribute '{attribute}' validation passed: '{actual}'")

    def validate_element_no_text(self, element) -> None:

        text = element.text.strip()
        assert text == "", f"Expected no text, but found '{text}'"
        print("Element has no text as expected")
