# # hooks
#
# # Imports & Class Definition
# # from behave import before_all, after_all, before, after
# from behave.model import Scenario
#
# import os
# import io
# from openpyxl import load_workbook
# from openpyxl.utils.exceptions import InvalidFileException
#
# from selenium import webdriver
# from selenium.webdriver.remote.remote_connection import RemoteConnection
#
# from utilities import Utilities
#
#
# class Hooks(Utilities):
#     """
#     Cucumber-style hooks for Behave test automation.
#     """
#
#     def before_all(self):
#         """
#         Suite-level setup: initialize execution time and open configuration sheet.
#         """
#         import datetime
#         # Capture the current execution time
#         self.execution_time = datetime.datetime.now().strftime("%d-%b")
#         try:
#             # Open the Fillo-based config connection
#             self.con_config = self.open_fillo_connection(self.ConfigSheetPath)
#             # Load inputs from config
#             self.load_inputs()
#             # Initialize sheet path
#             self.sheet_path()
#             self.logging(True, "Suite-level setup complete.")
#         except Exception as e:
#             self.logging(False, f"Suite-level setup failed: {e}")
#
#     def before_scenario(self, scenario: Scenario):
#         """
#         Scenario-level setup: initialize test data sheet, WebDriver, and reporting.
#         """
#         # Store the scenario name
#         self.scenario_name = scenario.name
#         # Build a unique run name
#         run_parts = [
#             self.Brand,
#             self.Platform,
#             self.Environment,
#             self.Browser,
#             self.execution_time
#         ]
#         self.run_name = "_".join(run_parts)
#         try:
#             # Open the test-data sheet connection
#             self.con_testdata = self.open_fillo_connection(self.TestSheetPath)
#             # Initialize the WebDriver and wait
#             self.driver = self.initialize_driver()
#             self.wait = self.initialize_wait(self.driver)
#             # Initialize Excel reporting if needed
#             if hasattr(self, 'ReportType') and self.ReportType.lower() == 'excel':
#                 self.init_excel_report()
#             self.logging(True, f"Scenario setup complete: {self.run_name}")
#         except Exception as e:
#             self.logging(False, f"Scenario setup failed: {e}")
#
#     def after_scenario(self, scenario: Scenario):
#         """
#         Scenario-level teardown: take screenshots on failure, mark BrowserStack sessions, and clean up.
#         """
#         try:
#             # On failure, capture screenshot and attach to report
#             if scenario.status == 'failed':
#                 screenshot = self.driver.get_screenshot_as_png()
#                 # Attach screenshot to scenario (if using allure or other)
#                 try:
#                     import allure
#                     allure.attach(screenshot, name=scenario.name, attachment_type=allure.attachment_type.PNG)
#                 except ImportError:
#                     pass
#             # BrowserStack session reporting
#             if getattr(self, 'BrowserStack', '').lower() == 'bs':
#                 session_id = self.driver.session_id
#                 status = 'passed' if scenario.status == 'passed' else 'failed'
#                 self.driver.execute_script(
#                     f"browserstack_executor: {{'action': 'setSessionStatus', 'arguments': {{'status':'{status}','reason':'{scenario.status}'}}}}")
#                 self.logging(True, f"BrowserStack session {session_id} marked as {status}.")
#             # Clean up thread-local resources
#             self.driver.quit()
#             if hasattr(self, 'con_testdata') and self.con_testdata:
#                 self.close_fillo_connection(self.con_testdata)
#             self.clear_self()
#             self.logging(True, f"Scenario teardown complete: {scenario.name}")
#         except Exception as e:
#             self.logging(False, f"Scenario teardown failed: {e}")
#
#     def after_all(self):
#         """
#         Suite-level teardown: close configuration sheet connection.
#         """
#         try:
#             if hasattr(self, 'con_config') and self.con_config:
#                 self.close_fillo_connection(self.con_config)
#                 self.logging(True, "Suite-level teardown complete.")
#         except Exception as e:
#             self.logging(False, f"Suite-level teardown failed: {e}")
