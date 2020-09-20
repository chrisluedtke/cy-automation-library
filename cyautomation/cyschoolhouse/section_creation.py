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


def input_staff_name(driver, staff_name):
    """ Selects the staff name from the drop down
    """
    dropdown = Select(driver.find_element_by_id("j_id0:j_id1:staffID"))
    dropdown.select_by_visible_text(staff_name)


def fill_static_elements(driver, insch_extlrn, start_date, end_date, target_dosage):
    """Fills in the static fields in the section creation form.

    Includes the start/end dat, days of week, if time is in or out of school,
    and the estimated amount of time for that section.
    """
    driver.find_element_by_id("j_id0:j_id1:startDateID").send_keys(start_date)
    driver.find_element_by_id("j_id0:j_id1:endDateID").send_keys(end_date)
    driver.find_element_by_id("j_id0:j_id1:freqID:1").click()
    driver.find_element_by_id("j_id0:j_id1:freqID:2").click()
    driver.find_element_by_id("j_id0:j_id1:freqID:3").click()
    driver.find_element_by_id("j_id0:j_id1:freqID:4").click()
    driver.find_element_by_id("j_id0:j_id1:freqID:5").click()
    dropdown = Select(driver.find_element_by_id("j_id0:j_id1:inAfterID"))
    dropdown.select_by_visible_text(insch_extlrn)
    driver.find_element_by_id("j_id0:j_id1:totalDosageID").send_keys(str(target_dosage))


def select_school(driver, school):
    """Selects the school name from section creation form.
    """
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "j_id0:j_id1:school-selector")))
    dropdown = Select(driver.find_element_by_id("j_id0:j_id1:school-selector"))
    dropdown.select_by_visible_text(school)


def select_subject(driver, section_name):
    """Selects the section type.
    """
    driver.find_element_by_xpath("//label[contains(text(), '"+ section_name +"')]").click()

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Proceed']")))
        driver.find_element_by_xpath("//input[@value='Proceed']").click()
    except TimeoutException:
        logging.warning("May have failed to choose subject")


def save_section(driver):
    """Saves the section.
    """
    driver.find_element_by_css_selector('input.black_btn:nth-child(2)').click()
    x_path = ('/html/body/div[1]/div[3]/table/tbody/tr/td[2]/div[4]/div[1]/'
              'table/tbody/tr/td[2]/input[5]')
    condition = EC.presence_of_element_located((By.XPATH, x_path))
    WebDriverWait(driver, 10).until(condition)


def update_nickname(driver, nickname):
    driver.find_element_by_css_selector('#topButtonRow > input:nth-child(3)').click()
    sleep(2)
    driver.find_element_by_id("00N1a000006Syte").send_keys(nickname)
    driver.find_element_by_xpath("//input[@value=' Save ']").click()
    sleep(2)


def create_single_section(school, acm, sectionname, insch_extlrn, start_date,
                          end_date, target_dosage, nickname="", driver=None):
    """ Creates one single section"""
    if driver is None:
        driver = get_driver()
        open_cyschoolhouse(driver)
    driver.get(f'{SF_URL}/apex/IM_AddStudentsToPrograms')
    select_school(driver, school)
    sleep(2)
    select_subject(driver, sectionname)
    sleep(2.5)
    input_staff_name(driver, acm)
    fill_static_elements(driver, insch_extlrn, start_date, end_date,
                         target_dosage)
    sleep(1)
    save_section(driver)
    logging.info(f"Created {sectionname} section for {acm}")
    if nickname:
        update_nickname(driver, nickname)


def create_all_sections(data=pd.DataFrame(), driver=None):
    """Runs the entire script. Loads sections to create from the 
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
            create_single_section(
                school=row['School'], acm=row['ACM'],
                sectionname=row['SectionName'],
                insch_extlrn=row['In_School_or_Extended_Learning'],
                start_date=row['Start_Date'], end_date=row['End_Date'],
                target_dosage=row['Target_Dosage'], driver=driver
            )
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
