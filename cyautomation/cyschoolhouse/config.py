from configparser import ConfigParser
import logging
import os

from pathlib import Path
import pandas as pd

__all__ = [
    'USER_SITE',
    'SF_URL',
    'SF_USER',
    'SF_PASS',
    'SF_TOKN',
    'INPUT_PATH',
    'LOG_PATH',
    'TEMP_PATH',
    'TEMPLATES_PATH',
    'get_sch_ref_df',
]

USER_SITE = 'Chicago'
SANDBOX = False
INPUT_PATH = str(Path(__file__).parent / 'input_files')
LOG_PATH = str(Path(__file__).parent / 'log')
TEMP_PATH = str(Path(__file__).parent / 'temp')
TEMPLATES_PATH = str(Path(__file__).parent / 'templates')
SCH_REF_PATH = ('Z:/ChiPrivate/Chicago Data and Evaluation/'
                'SY19/SY19 School Reference.xlsx')

if not os.path.exists(TEMP_PATH):
    os.mkdir(TEMP_PATH)

creds_path = str(Path(__file__).parent / 'credentials.ini')
config = ConfigParser()

try:
    open(creds_path)
except FileNotFoundError:
    raise FileNotFoundError('Before you can use this library, you must create '
                            'a credentials.ini file. See the README for '
                            'details.')
else:
    config.readfp(open(creds_path))

if SANDBOX == False:
    SF_URL = "https://na82.salesforce.com"
    sf_creds = config['Salesforce']
elif SANDBOX == True:
    SF_URL = "https://cs59.salesforce.com"
    sf_creds = config['Salesforce Sandbox']

okta_creds = config['Single Sign On']

SF_USER = sf_creds['username']
SF_PASS = sf_creds['password']
SF_TOKN = sf_creds['security_token']
OKTA_USER = okta_creds['username']
OKTA_PASS = okta_creds['password']

def set_logger(name):
    logger = logging.getLogger(name)
    logger.setLevel('DEBUG')

    format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    formatter = logging.Formatter(format)

    file_log_handler = logging.FileHandler(str(Path(LOG_PATH) / f"{name}.log"))
    file_log_handler.setFormatter(formatter)
    logger.addHandler(file_log_handler)

    stderr_log_handler = logging.StreamHandler()
    stderr_log_handler.setFormatter(formatter)
    logger.addHandler(stderr_log_handler)

    return logger

def get_sch_ref_df(sch_df_path=SCH_REF_PATH):
    df = pd.read_excel(sch_df_path)
    df = df.loc[~df['Informal Name'].isin(['CE', 'Onboarding'])]

    return df
