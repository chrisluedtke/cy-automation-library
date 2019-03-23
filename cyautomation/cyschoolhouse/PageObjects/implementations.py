# -*- coding: utf-8 -*-
from pathlib import Path
import sys
from time import sleep

from seleniumrequests import Firefox
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException
)
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile

from . import pages as page
from ..config import (
    INPUT_PATH, OKTA_USER, OKTA_PASS, SF_USER, SF_PASS, TEMP_PATH
)

GECKO_PATH = str(Path(__file__).parents[3] / 'geckodriver/geckodriver.exe')


class BaseImplementation(object):
    """Base Implementation object

    It's an important feature of every implementation that it either take an
    existing driver or be capable of starting one
    """
    def __init__(self, driver=None):
        if driver == None:
            self.driver  = self.get_driver()
        else:
            self.driver = driver

    def get_driver(self):
        profile = FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2)
        profile.set_preference('browser.download.manager.showWhenStarting', False)
        profile.set_preference('browser.download.dir', TEMP_PATH)
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk',
                               ('application/csv,text/csv,application/vnd.ms-excel,'
                                'application/x-msexcel,application/excel,'
                                'application/x-excel,text/comma-separated-values'))

        driver = Firefox(firefox_profile=profile,
                         executable_path=GECKO_PATH)
        return driver


class Okta(BaseImplementation):
    """Object for handling Okta

    Wraps all processes for logging into Okta and navigation.
    """
    def __init__(self):
        super().__init__()
        self.user = OKTA_USER
        self.pwd = OKTA_PASS

    def set_up(self):
        """Navigate to City Year Okta login"""
        self.driver.get("https://cityyear.okta.com")

    def enter_credentials(self):
        """Enter credentials for Okta Login"""
        login_page = page.OktaLoginPage(self.driver)
        assert login_page.page_is_loaded()
        login_page.username = self.user
        login_page.password = self.pwd
        login_page.click_login_button()

    def check_logged_in(self):
        """Confirm login"""
        homepage = page.OktaHomePage(self.driver)
        assert homepage.page_is_loaded()

    def login(self):
        """Runs all steps to login to Okta"""
        self.set_up()
        self.enter_credentials()
        self.check_logged_in()

    def launch_cyschoolhouse(self):
        """Script for logging into Okta and cyschoolhouse"""
        # Login via okta
        self.login()
        # Nav from Okta home to cyschoolhouse
        Okta = page.OktaHomePage(self.driver)
        assert Okta.page_is_loaded()
        Okta.launch_cyschoolhouse()


class IndicatorAreaEnrollment(Okta):
    """Implementation object for Indicator Area Enrollment"""

    def __init__(self):
        from pandas import read_excel
        super().__init__()
        xl_path = str(Path(INPUT_PATH) / 'indicator_area_roster.xlsx')
        self.data = read_excel(xl_path)
        self.student_list = self.data['Student: Student ID'].unique()

    def nav_to_form(self):
        """Initial setup script for IA enrollment.

        Goes through login for Okta as well as navigating to the form and
        ensuring the page is loaded appropriately
        """
        self.launch_cyschoolhouse()
        cysh_home = page.CyshHomePage(self.driver)
        assert cysh_home.page_is_loaded()

        self.driver.get("https://c.na24.visual.force.com/apex/IM_Indicator_Areas")
        ia_form = page.CyshIndicatorAreas(self.driver)
        ia_form.wait_for_page_to_load()

    def get_student_details(self, student_id):
        """Returns a students details including school, grade, name, and ia list given their id"""
        student_records = self.data[self.data['Student: Student ID'] == student_id]
        school = student_records.School.unique()[0]
        grade = student_records['Student: Grade'].unique()[0]
        name = student_records['Student: Student Last Name'].unique()[0]
        ia_list = student_records['Indicator Area'].values
        return school, grade, name, ia_list

    def enroll_student(self, student_id):
        """Handles the enrollment process of all IAs for a single student"""
        ia_form = page.CyshIndicatorAreas(self.driver)
        ia_form.wait_for_page_to_load()
        school, grade, name, ia_list = self.get_student_details(student_id)
        ia_form.select_school(school)
        sleep(1)
        ia_form.select_grade(str(grade))
        sleep(1)
        for ia in ia_list:
            ia_form.name_search = name
            sleep(3)
            self.assign_ias(student_id, ia)
            sleep(2)
        ia_form.save()

    def assign_ias(self, student_id, ia):
        """Assign an IA for a single student and indicator"""
        ia_form = page.CyshIndicatorAreas(self.driver)
        ia_form.wait_for_page_to_load()
        ia_form.select_student(student_id)
        sleep(3)
        ia_form.assign_indicator_area(ia)

    def enroll_all_students(self):
        """Executes the full IA enrollment"""
        self.nav_to_form()
        self.error_count = 0
        for student_id in self.student_list:
            if self.error_count == 5:
                self.error_count = 0
                self.driver.quit()
                self.driver = self.get_driver()
                self.nav_to_form()

            try:
                self.enroll_student(student_id)
            except KeyboardInterrupt:
                return None

            except TimeoutException:
                print(f"Timeout Failure on student: {student_id}")
                self.error_count += 1

            except StaleElementReferenceException:
                print(f"Stale Element failure on student: {student_id}")
                self.error_count += 1

            except Exception as e:
                print(f"Caught a {e} error on student: {student_id}")
                self.error_count += 1
