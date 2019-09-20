# -*- coding: utf-8 -*-
"""cyschoolhouse Suite
This suite is a set of helper functions which address the broader task of
accessing data on cyschoolhouse. This will include navigation functions, logins,
and other common tasks we can antipicate needing to do for multiple products.
"""

import getpass
import io
import pickle
from pathlib import Path
from time import sleep, time

import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from seleniumrequests import Firefox

from .config import LOG_PATH, SF_PASS, SF_URL, SF_USER, TEMP_PATH, set_logger

logger = set_logger(name=Path(__file__).stem)
COOKIES_PATH = Path(__file__).parent / 'cookies.pkl'
GECKO_PATH = str(Path(__file__).parents[2] / 'geckodriver/geckodriver.exe')


def get_login_credentials(prompt_user_pass=False):
    """Extract login information from credentials.ini

    Optionally, set prompt_user_pass to `True` and supply credentials interactively.
    It's a little more secure to enter the user/pass every time, but it requires the user to
    be present at script initialization. If you intend to run on a schedule, then
    be prepared to store your credentials in a file.
    """

    if prompt_user_pass == False:
        user = SF_USER
        pwd = SF_PASS
    else:
        print('Please enter your City Year Okta credential below.')
        print('It is used for sign in only and is not stored in any way after the script closes.')
        user = input('Username:')
        # If you run from the cmd prompt, it won't show your password.  It will in an interactive
        # console though, so keep that in mind.
        pwd = getpass.getpass()

    return user, pwd


def get_driver():
    """Get Firefox driver

    Returns the Firefox driver object and handles the path.
    """
    profile = FirefoxProfile()
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference('browser.download.dir', TEMP_PATH)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk',
                           ('application/csv,text/csv,application/vnd.ms-excel,'
                            'application/x-msexcel,application/excel,'
                            'application/x-excel,text/comma-separated-values'))
    return Firefox(firefox_profile=profile, executable_path=GECKO_PATH)


def standard_login(driver, prompt_user_pass=False):
    """ Login to salesforce using the standard element names "username" and "password"
    """
    user, pwd = get_login_credentials(prompt_user_pass)

    driver.find_element_by_name("username").send_keys(user)
    driver.find_element_by_name("pw").send_keys(pwd + Keys.RETURN)
    return driver


def open_cyschoolhouse(driver=None, prompt_user_pass=False):
    """Opens the cyschoolhouse instance

    You will need to monitor your email inbox at this point to copy+paste an
    authentication code.
    """
    if driver is None:
        driver = get_driver()

    driver.get(SF_URL)

    # if cookies exist, load them and reload salesforce
    if COOKIES_PATH.exists():
        for cookie in pickle.load(open(COOKIES_PATH, "rb")):
            driver.add_cookie(cookie)
        driver.get(SF_URL)
    
    if driver.find_elements_by_name("username"):
        driver = standard_login(driver, prompt_user_pass)

    # Wait for next page to load. User may need to supply 2FA here.
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "tsidLabel")))
    assert 'salesforce' in driver.current_url

    # On successful login, save cookies. This should reduce 2FA.
    pickle.dump(driver.get_cookies() , open(COOKIES_PATH, "wb"))
    return driver


def get_report(report_key):
    driver = get_driver()
    open_cyschoolhouse(driver)
    url = f'{SF_URL}/{report_key}/?export=1&enc=UTF-8&xf=csv'
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


if __name__ == '__main__':
    driver = get_driver()
    driver = open_cyschoolhouse(driver)
    driver.quit()
