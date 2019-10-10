import logging
import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

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

assert Path('Z:').exists(), ("Failed to connect to Z: drive. You must map "
                             "cyconnect (Sharepoint) to this drive location. "
                             "You may need to login to cyconnect on Internet "
                             "Explorer. See the README.md.")

# configuration from .env
YEAR = os.environ['YEAR']
USER_SITE = os.environ['USER_SITE']

SF_URL_DICT = {
    'SY19': "https://cs59.salesforce.com",
    'SY19_SB': "https://na82.salesforce.com",
    'SY20': "https://na90.salesforce.com",
    # 'SY_20_SB': "",
}

if os.getenv('SF_SANDBOX') == 'True':
    SF_URL = SF_URL_DICT[YEAR + '_SB']
    SF_USER = os.environ['SF_SB_USER']
    SF_PASS = os.environ['SF_SB_PASS']
    SF_TOKN = os.environ['SF_SB_TOKEN']
else:
    SF_URL = SF_URL_DICT[YEAR]
    SF_USER = os.environ['SF_USER']
    SF_PASS = os.environ['SF_PASS']
    SF_TOKN = os.environ['SF_TOKEN']

OKTA_USER = os.getenv('OKTA_USER')
OKTA_PASS = os.getenv('OKTA_PASS')

# configuration
INPUT_PATH = str(Path(__file__).parent / 'input_files')
LOG_PATH = str(Path(__file__).parent / 'log')
TEMP_PATH = str(Path(__file__).parent / 'temp')
TEMPLATES_PATH = f"Z:/ChiPrivate/Chicago Data and Evaluation/{YEAR}/Templates/"
SCH_REF_PATH = ('Z:/ChiPrivate/Chicago Data and Evaluation/'
                f'{YEAR}/{YEAR} School Reference.xlsx')

for path in [LOG_PATH, TEMP_PATH]:
    Path(path).mkdir(exist_ok=True)


def set_logger(name, filename='log.log'):
    logger = logging.getLogger(name)
    logger.setLevel('DEBUG')

    format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    formatter = logging.Formatter(format)

    file_log_handler = logging.FileHandler(str(Path(LOG_PATH) / filename))
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
