import os
import subprocess
import sys

from .config import SCH_REF_PATH

import pandas as pd


def map_sharepoint_drive():
    try:
        subprocess.run(
            r'net use z: /del /Y',
            shell=True,
            check=True,
            capture_output=True
        )
    except subprocess.CalledProcessError:
        pass

    try:
        subprocess.run(
            f'net use z: {os.environ["SHAREPOINT_URL"]}',
            shell=True,
            check=True,
            capture_output=True
        )
    except subprocess.CalledProcessError as e:
        print(e.stderr)
        sys.exit(1)


def get_sch_ref_df(sch_df_path=SCH_REF_PATH):
    df = pd.read_excel(sch_df_path)
    df = df.loc[~df['Informal Name'].isin(['CE', 'Onboarding'])]

    return df
