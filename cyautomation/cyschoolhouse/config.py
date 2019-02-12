from configparser import ConfigParser
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

creds_path = str(Path(__file__).parent / 'credentials.ini')
config = ConfigParser()

try:
    open(creds_path)
except FileNotFoundError:
    raise FileNotFoundError('Before you can use this pacakge, you must create '
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

SF_USER = sf_creds['username']
SF_PASS = sf_creds['password']
SF_TOKN = sf_creds['security_token']

INPUT_PATH = str(Path(__file__).parent / 'input')
LOG_PATH = str(Path(__file__).parent / 'log')
TEMP_PATH = str(Path(__file__).parent / 'temp')
TEMPLATES_PATH = str(Path(__file__).parent / 'templates')
SCH_REF_PATH = ('Z:/ChiPrivate/Chicago Data and Evaluation/'
                'SY19/SY19 School Reference.xlsx')

def get_sch_ref_df(sch_df_path=SCH_REF_PATH):
    sch_ref_df = pd.read_excel(sch_df_path)
    sch_ref_df = sch_ref_df.loc[~sch_ref_df['Informal Name'].isin(['CE', 'Onboarding'])]

    return sch_ref_df