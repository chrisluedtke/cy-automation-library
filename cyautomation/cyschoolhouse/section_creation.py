# -*- coding: utf-8 -*-
"""Automated Section Creation
Script for the automatic creation of sections in cyschoolhouse. Please ensure
you have all dependencies, and have set up the input file called "section-creator-input.xlsx"
in the input files folder.
"""
import logging
import os
from pathlib import Path
from time import sleep

import pandas as pd
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from .config import SF_URL
from .cyschoolhousesuite import get_driver, open_cyschoolhouse
from .simple_cysh import get_object_df, in_str, execute_query
from .utils import validate_date


class Section:
    def __init__(self, school, corps_member, program, in_after_sch, start_date,
                 end_date, nickname=""):
        self.school = school
        self.corps_member = corps_member
        self.program = program
        self.in_after_sch = in_after_sch
        self.start_date = start_date
        validate_date(start_date)
        self.end_date = end_date
        validate_date(end_date)
        self.nickname = nickname

    def create(self, driver=None):
        """Creates a single section"""
        exists_as_id = self.check_exists()
        if exists_as_id:
            logging.info(
                f"{self.program} section already exists for "
                f"{self.corps_member}: {exists_as_id}"
            )
            return exists_as_id

        if driver is None:
            driver = get_driver()
            open_cyschoolhouse(driver)

        driver.get(f'{SF_URL}/apex/IM_AddStudentsToPrograms')

        self._set_school(driver)
        self._set_program(driver)
        self._set_corps_member(driver)
        self._set_start_date(driver)
        self._set_end_date(driver)
        self._set_in_after_sch(driver)
        sleep(1)
        self._save_section(driver)
        logging.info(f"Created {self.program} section for {self.corps_member}")

        if self.nickname:
            self._set_nickname(driver)

    def check_exists(self):
        inputs = [self.program, self.corps_member, self.school]
        clean_inputs = [s.replace("'", "\\'") for s in inputs]

        query = """\
        SELECT Id
        FROM Section__c
        WHERE Program__r.Name = '{}'
          AND Intervention_Primary_Staff__r.Name = '{}'
          AND School__r.Name = '{}'
        """.format(*clean_inputs)

        result = execute_query(query)

        if result['records']:
            return result['records'][0]['Id']
        else:
            return False

    def _set_school(self, driver):
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "j_id0:j_id1:school-selector"))
        )
        dropdown = Select(driver.find_element_by_id("j_id0:j_id1:school-selector"))
        dropdown.select_by_visible_text(self.school)
        sleep(2)

    def _set_program(self, driver):
        """Selects the section type.
        """
        driver.find_element_by_xpath(f"//label[contains(text(), '{self.program}')]").click()

        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@value='Proceed']"))
            )
            driver.find_element_by_xpath("//input[@value='Proceed']").click()
        except TimeoutException:
            logging.warning("May have failed to choose subject")

        sleep(2.5)

    def _set_corps_member(self, driver):
        """Selects the staff name from the drop down
        """
        dropdown = Select(driver.find_element_by_id("j_id0:j_id1:staffID"))
        dropdown.select_by_visible_text(self.corps_member)

    def _set_start_date(self, driver):
        driver.find_element_by_id("j_id0:j_id1:startDateID").send_keys(self.start_date)

    def _set_end_date(self, driver):
        driver.find_element_by_id("j_id0:j_id1:endDateID").send_keys(self.end_date)

    def _set_in_after_sch(self, driver):
        dropdown = Select(driver.find_element_by_id("j_id0:j_id1:inAfterID"))
        dropdown.select_by_visible_text(self.in_after_sch)

    def _save_section(self, driver):
        """Saves the section.
        """
        driver.find_element_by_css_selector('input.black_btn:nth-child(2)').click()
        x_path = (
            '/html/body/div[1]/div[3]/table/tbody/tr/td[2]/div[4]/div[1]/'
            'table/tbody/tr/td[2]/input[5]'
        )
        condition = EC.presence_of_element_located((By.XPATH, x_path))
        WebDriverWait(driver, 10).until(condition)

    def _set_nickname(self, driver):
        driver.find_element_by_css_selector('#topButtonRow > input:nth-child(3)').click()
        sleep(2)
        driver.find_element_by_id("00N1a000006Syte").send_keys(self.nickname)
        driver.find_element_by_xpath("//input[@value=' Save ']").click()
        sleep(2)


def create_all_sections(data=pd.DataFrame(), driver=None):
    """Loads sections to create from the
    spreadsheet at 'input_files/section-creator-input.xlsx'.
    """
    if data.empty:
        data = pd.read_excel(os.path.join(os.path.dirname(__file__),
                             'input_files/section-creator-input.xlsx'))

    logging.info(f'Creating {len(data)} sections')

    data['Start_Date'] = pd.to_datetime(data['Start_Date']).dt.strftime('%m/%d/%Y')
    data['End_Date'] = pd.to_datetime(data['End_Date']).dt.strftime('%m/%d/%Y')
    data = data.fillna('').replace('NaT', '')

    if driver is None:
        driver = get_driver()
        open_cyschoolhouse(driver=driver)

    for _, row in data.iterrows():
        try:
            section = Section(
                school=row['School'],
                corps_member=row['ACM'],
                program=row['SectionName'],
                in_after_sch=row['In_School_or_Extended_Learning'],
                start_date=row['Start_Date'],
                end_date=row['End_Date'],
            )
            section.create(driver)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            logging.error(f"Section creation failed for {row['ACM']}, {row['SectionName']}: {e}")
            driver.get(SF_URL)
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                driver.switch_to.alert.accept()
                sleep(2)
            except TimeoutException:
                pass

    driver.quit()


if __name__ == '__main__':
    create_all_sections()
