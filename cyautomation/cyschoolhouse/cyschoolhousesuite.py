# -*- coding: utf-8 -*-
"""cyschoolhouse Suite
This suite is a set of helper functions which address the broader task of 
accessing data on cyschoolhouse. This will include navigation functions, logins,
and other common tasks we can antipicate needing to do for multiple products. 
"""

from pathlib import Path
from time import time, sleep
import getpass
import logging
from seleniumrequests import Firefox
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
import pandas as pd
import io

"""Configuration Variables

The below variables are used to configure some machine specific details. Unfortunately
I don't know how we can avoid needing some of these variables for the potentially
arbitrary machine and file system we might get.  
"""
# Location of the file with my SSO username and password
key_file_path = str(Path.home() / 'Desktop/keyfile.txt')
gecko_path = str(Path.cwd().parents[1] / 'geckodriver/geckodriver.exe')
log_path = str(Path.cwd().parents[1] / 'cyautomation/cyschoolhouse/log')
temp_path = str(Path.cwd().parents[1] / 'cyautomation/cyschoolhouse/temp')
templates_path =str(Path.cwd().parents[1] / 'cyautomation/cyschoolhouse/templates')

def extract_key():
    """Extract SSO Information from keyfile
    
    One practice I use here is to store my SSO in a keyfile instead of either
    entering it in the cmd prompt on each run or hardcoding it in the program. 
    Keyfile should be formatted like so:
        'descript:cityyear sso/user:aperusse/pass:p@ssw0rd'
    Simply replace the username and password with your credentials and then save
    it as keyfile.txt
    """
    with open(key_file_path) as file:
        keys = file.read()
    split_line = keys.split("/")
    entries = [item.split(":")[1] for item in split_line]
    desc, user, pwd = entries
    return user, pwd

def request_key():
    """Requests user credentials
    
    Added as a handy way to capture the SSO credentials from the user.  However, the interaction
    kills the fully automated aspect of the code.  
    """
    print('Please enter your City Year Okta credential below.')
    print('It is used for sign in only and is not stored in any way after the script closes.')
    user = input('Username:')
    # If you run from the cmd prompt, it won't show your password.  It will in an interactive
    # console though, so keep that in mind. 
    pwd = getpass.getpass()
    return user, pwd

def get_driver():
    """Get Firefox driver
    
    Returns# the Firefox driver object and handles the path. 
    """
    configure_log(log_path)
    
    profile = FirefoxProfile()
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference('browser.download.dir', temp_path)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ('application/csv,text/csv,application/vnd.ms-excel,application/x-msexcel,application/excel,application/x-excel,text/comma-separated-values'))
    return Firefox(firefox_profile=profile, executable_path=gecko_path)

def standard_login(driver):
    """# Login to a form using the standard element names "username" and "password"
    
    User preference on how login is collected. It's probably a little more secure
    to enter the user/pass every time, however that will require that you be
    there to initialize the script.  If you intend to run on a schedule then
    be prepared to store your credentials in a file.  Ideally, we would have 
    slightly better security around that filer.
    """

    user, pwd = extract_key()
    #user, pwd = request_key()
    driver.find_element_by_name("username").send_keys(user)
    driver.find_element_by_name("password").send_keys(pwd + Keys.RETURN)
    return driver

def open_okta(driver):
    """Logs into Okta via the front door.
    """
    # Open Okta login
    logging.info(":".join([str(time()), "Opening Okta"]))
    driver.get("https://cityyear.okta.com")
    # Input login
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "username")))
    driver = standard_login(driver)
    # Wait for next page to load
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "app-link")))
    # Check that we aren't in the login page anymore
    assert 'login' not in driver.current_url
    return driver

def open_cyschoolhouse18(driver):
    """Opens the 2018 cyschoolhouse instance
    
    """
    #driver.implicitly_wait(10)
    driver = open_okta(driver)
    driver.get("https://cityyear.okta.com/home/salesforce/0oao08nxmQCYJQHUHCBU/46?fromHome=true")
    # Wait for next page to load
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tsidLabel")))
    assert 'salesforce' in driver.current_url
    return driver

def open_cyschoolhouse18_sb(driver):
    """Opens the 208 cyschoolhouse sandbox instance
    
    """
    driver.implicitly_wait(10)
    driver = open_okta(driver)
    logging.debug(":".join([str(time()), "Opening cyschoolhouse FY18 sandbox"]))
    driver.get("https://cityyear.okta.com/home/salesforce/0oa1dt5ae7mOkRt3O0h8/46?fromHome=true")
    # Wait for next page to load
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tsidLabel")))
    assert 'salesforce' in driver.current_url
    return driver

def configure_log(log_folder):
    """Configures the log for update cycle
    
    Generates a log file for this session. Uses a timestamp from the time library to as a part
    of the name. 
    """
    timestamp = str(time()).split(".")[0]
    logging.basicConfig(filename="".join([log_folder, '/update_', timestamp, '.log']), level=logging.INFO)

def get_report(report_key):
    driver = get_driver()
    open_cyschoolhouse18(driver)
    url = 'https://na24.salesforce.com/' + report_key + '?export=1&enc=UTF-8&xf=csv'
    response = driver.request('GET', url)
    df = pd.read_csv(io.StringIO(response.content.decode('utf-8')))
    driver.quit()
    return df

def delete_folder(pth):
    for sub in pth.iterdir():
        if sub.is_dir():
            delete_folder(sub)
        else:
            sub.unlink()
    pth.rmdir()
    
def fancy_box_wait(driver, waittime=10):
    WebDriverWait(driver, waittime).until(EC.presence_of_element_located((By.XPATH, ".//div[contains(@id, 'fancybox-wrap')]")))
    WebDriverWait(driver, (waittime+30)).until(EC.invisibility_of_element_located((By.XPATH, ".//div[contains(@id, 'fancybox-wrap')]")))
    sleep(2)
    return driver
    
#if __name__ == '__main__':
#    driver = get_driver()
#    driver = open_cyschoolhouse18_sb(driver)
#    driver.quit()