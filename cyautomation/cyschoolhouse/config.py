import logging
import os
from pathlib import Path

from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

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
]

# configuration from .env
YEAR = os.environ['YEAR']
USER_SITE = os.environ['USER_SITE']
SF_URL = os.environ['SF_URL']
SF_USER = os.environ['SF_USER']
SF_PASS = os.environ['SF_PASS']
SF_TOKN = os.environ['SF_TOKEN']
EXCEL_PROTECTION_PWD = os.environ['EXCEL_PROTECTION_PWD']

OKTA_USER = os.getenv('OKTA_USER')
OKTA_PASS = os.getenv('OKTA_PASS')

# configuration
INPUT_PATH = str(Path(__file__).parent / 'input_files')
LOG_PATH = str(Path(__file__).parents[2] / 'logs')
TEMP_PATH = str(Path(__file__).parents[2] / 'test')
TEMPLATES_PATH = Path(f"Z:/ChiPrivate/Chicago Data and Evaluation/{YEAR}/Templates/")
SCH_REF_PATH = ('Z:/ChiPrivate/Chicago Data and Evaluation/'
                f'{YEAR}/{YEAR} School Reference.xlsx')

for path in [LOG_PATH, TEMP_PATH]:
    Path(path).mkdir(exist_ok=True)

# logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(filename)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("log.log"),
        logging.StreamHandler()
    ]
)
