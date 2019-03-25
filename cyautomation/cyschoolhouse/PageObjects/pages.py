# -*- coding: utf-8 -*-
from time import sleep

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from . import locators as loc
from . import elements as elem

SF_URL = ('https://cityyear.okta.com/home/salesforce/'
          '0oa19u4wnhzgPqjtw0h8/46?fromHome=true')
SF_SB_URL = ('https://cityyear.okta.com/home/salesforce/'
             '0oa1g66my5rkEPiEn0h8/46?fromHome=true')

class BasePage(object):
    """Base class to initialize the base page that will be called from all pages"""

    def __init__(self, driver):
        self.driver = driver

    def wait_for_page_to_load(self):
        WebDriverWait(self.driver, 100).until(EC.title_contains(self.title))

    def page_is_loaded(self):
        self.wait_for_page_to_load()
        return self.title in self.driver.title

class OktaLoginPage(BasePage):
    """Actions for the Okta Login Page"""

    #title used to ensure the page has loaded before acting
    title = 'Sign In'
    #Declare variables corresponding to the username and password box elements
    username = elem.OktaUsernameElement()
    password = elem.OktaPasswordElement()

    def click_login_button(self):
        element = self.driver.find_element(*loc.OktaLoginPageLocators.GO_BUTTON)
        element.click()

class OktaHomePage(BasePage):
    """Actions for the Okta Homepage"""

    #title used to ensure the page has loaded before acting
    title = 'My Applications'

    #Functions for launching various applications.  It's important that we
    def launch_cyschoolhouse(self):
        self.driver.get(SF_URL)

    def launch_cyschoolhouse_sandbox(self):
        self.driver.get(SF_SB_URL)

class CyshHomePage(BasePage):
    """Actions for the cyschoolhouse Home Page"""

    title = 'Salesforce - Unlimited Edition'
    search_bar = elem.CyshSearchElement()

    def set_search_filter(self, filter):
        """Changes to the filter type on the search function"""
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'sen')))
        filters = Select(self.driver.find_element(*loc.CyshHomePageLocators.SEARCH_FILTER))
        filters.select_by_visible_text(filter)

    def click_search_button(self):
        button = self.driver.find_element(*loc.CyshHomePageLocators.SEACH_BUTTON)
        button.click()

class CyshIndicatorAreas(BasePage):

    def wait_for_page_to_load(self):
        WebDriverWait(self.driver, 100)\
            .until(EC.presence_of_all_elements_located(loc.IndicatorAreaLocators.PAGE_TITLE))

    def select_school(self, school_name):
        """Updates the school selector to the given school name"""
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable(loc.IndicatorAreaLocators.SCHOOL_SELECT)
        )
        selector = Select(self.driver.find_element(*loc.IndicatorAreaLocators.SCHOOL_SELECT))
        selector.select_by_visible_text(school_name)
        sleep(0.75)

    def select_grade(self, grade):
        """Updates the school selector to the given school name"""
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(loc.IndicatorAreaLocators.GRADE_SELECT))
        selector = Select(self.driver.find_element(*loc.IndicatorAreaLocators.GRADE_SELECT))
        selector.select_by_visible_text(grade)
        sleep(0.75)

    def select_first_page(self):
        """Start from the first page"""
        first_page = self.driver.find_elements(By.LINK_TEXT, 'First')
        if first_page:
            first_page[0].click()
            sleep(0.5)

    def select_student(self, student_id):
        """Selects a visible student using their Salesforce Id

        These selectors were pretty annoying.  Originally I used XPATH to select the appropriate
        row, but the resulting student object wouldn't register the clicks.  I fixed the problem
        by switching to css selectors and selecting the first cell of the chosen row instead of selecting
        the row itself.
        """
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//table[@id='StudentsTable']"))
        )
        try:
            WebDriverWait(self.driver, 1.5).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, f"tr[id ^= '{student_id}'] td"))
            )
        except TimeoutException:
            try:
                self.driver.find_element(By.LINK_TEXT, 'Next').click()
                sleep(0.5)
            except NoSuchElementException:
                raise Exception('Student not found')

            return self.select_student(student_id)

        self.driver.find_element(By.CSS_SELECTOR, f"tr[id ^= '{student_id}'] td").click()
        sleep(0.5)
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable(loc.IndicatorAreaLocators.ADD_TO_IA_BUTTON)
        )
        self.driver.find_element(*loc.IndicatorAreaLocators.ADD_TO_IA_BUTTON).click()

    def assign_indicator_area(self, ia):
        ia_dict = {
            'Attendance': loc.IndicatorAreaLocators.ATTENDANCE,
            'Behavior': loc.IndicatorAreaLocators.BEHAVIOR,
            'ELA/Literacy': loc.IndicatorAreaLocators.ELA,
            'Math': loc.IndicatorAreaLocators.MATH,
        }
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ia_dict[ia]))
        self.driver.find_element(*ia_dict[ia]).click()
        sleep(0.25)
        self.driver.find_element(*loc.IndicatorAreaLocators.ADD_IA_BUTTON).click()
        WebDriverWait(self.driver, 10).until(EC.invisibility_of_element_located(loc.IndicatorAreaLocators.IA_WINDOW))

    def save(self):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(loc.IndicatorAreaLocators.SAVE))
        self.driver.find_element(*loc.IndicatorAreaLocators.SAVE).click()
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(loc.IndicatorAreaLocators.SUCCESS_MSG))
