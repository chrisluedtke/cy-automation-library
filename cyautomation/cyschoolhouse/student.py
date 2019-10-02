from pathlib import Path
from time import sleep

import numpy as np
import pandas as pd
from selenium.webdriver.support.ui import Select

from . import simple_cysh as cysh
from .cyschoolhousesuite import get_driver, open_cyschoolhouse
from .config import INPUT_PATH, YEAR, SF_URL, TEMP_PATH
from .sendemail import send_email


def _sf_api_approach(xlsx_path):
    """ This is how the task would be accomplished via salesforce API, 
    if we could edit the fields:
    """
    df = pd.read_excel(xlsx_path)
    school_df = get_cysh_df('Account', ['Id', 'Name'])
    df = df.merge(school_df, how='left', left_on='School', right_on='Name')

    drop_ids = []
    for index, row in df.iterrows():
        search_result = cysh.sf.query('SELECT Id FROM Student__c WHERE '
                                      f"Local_Student_ID__c = '{row['Student CPS ID']}'")
        if len(search_result['records']) > 0:
            drop_ids.append(row['Student CPS ID'])
    df = df.loc[~df['Student CPS ID'].isin(drop_ids)]

    for index, row in df.iterrows():
        stu_dict = {
            'Local_Student_ID__c':str(row['Student CPS ID']),
            'School__c':row['Id'],
            'Name':(row['Student First Name'] + ' ' + row['Student Last Name']),
            'Student_Last_Name__c':row['Student Last Name'],
            'Grade__c':str(row['Student Grade Level']),
            #'School_Name__c':row['Name_y'],
         }

    return None


def upload_all(enrollment_date, xlsx_dir=INPUT_PATH,
               xlsx_name='New Students for cyschoolhouse.xlsx', sf=cysh.sf):
    """ Runs the entire student upload process.
    """
    xlsx_path = str(Path(xlsx_dir) / xlsx_name)

    sdnt_df = import_parameters(xlsx_path, enrollment_date)
    sdnt_df = remove_extant_students(sdnt_df)
    sdnt_df = sdnt_df.rename(columns={'School': 'Informal Name'})
    sch_ref_df = cysh.get_sch_ref_df()
    sdnt_df = sdnt_df.merge(sch_ref_df[['School', 'Informal Name']],
                            how='left', on='Informal Name')

    setup_df = cysh.get_object_df('Setup__c', ['Id', 'School__c'],
                                  rename_id=True, rename_name=True)
    school_df = cysh.get_object_df('Account', ['Id', 'Name'])
    setup_df = setup_df.merge(school_df, how='left', left_on='School__c',
                              right_on='Id')
    setup_df = setup_df.loc[~setup_df['Id'].isnull()]

    if len(sdnt_df) == 0:
        print(f'No new students to upload.')
        return None

    driver = get_driver()
    open_cyschoolhouse(driver)

    for school_name, df in sdnt_df.groupby('School'):
        # Write csv
        path_to_csv = Path(TEMP_PATH) / f"{YEAR} New Students for CYSH - {school_name}.csv"

        (df.drop(columns=["School"])
           .to_csv(path_to_csv, index=False, date_format='%m/%d/%Y'))

        # Navigatge to student enrollment page
        setup_id = setup_df.loc[setup_df['Name']==school_name, 'Setup__c'].values[0]
        driver.get(f'{SF_URL}/apex/CT_core_LoadCsvData_v2?setupId={setup_id}'
                   '&OldSideBar=true&type=Student')
        sleep(2)

        input_file(driver, path_to_csv)
        sleep(2)

        insert_data(driver)
        sleep(2)

        # Publish
        # Seems to work, but not completely sure if script
        # pauses until upload is complete, both for the "Insert Data"
        # phase, and the "Publish Staff/Student Records" phase.

        driver.get(f'{SF_URL}/apex/schoolsetup_staff?setupId={setup_id}')
        driver.find_element_by_css_selector('input.red_btn').click()
        sleep(3)

        print(f"Uploaded {len(df)} students")

        path_to_csv.unlink()

    # Email school manager to inform of successful student upload
    staff_df = cysh.get_staff_df()
    staff_df = staff_df.loc[staff_df['Role__c'].str.lower()=='impact manager']

    to_addrs = staff_df.loc[staff_df['School'].isin(sdnt_df['School']), 'Email__c']
    to_addrs = to_addrs.unique().tolist()

    send_email(
        to_addrs = to_addrs,
        subject = 'New students now in cyschoolhouse',
        body = ('The students you submitted have been successfully uploaded '
                'to cyschoolhouse.')
    )

    driver.quit()


def import_parameters(xlsx_path, enrollment_date):
    """Imports input data from xlsx

    `enrollment_date` in the format 'MM/DD/YYYY'
    """
    df = pd.read_excel(xlsx_path, converters={'*REQ* Grade':int})

    column_rename = {
        'Student CPS ID':'*REQ* Local Student ID',
        'Student First Name':'*REQ* First Name',
        'Student Last Name':'*REQ* Last Name',
        'Student Grade Level':'*REQ* Grade',
    }

    df.rename(columns=column_rename, inplace=True)

    df["*REQ* Student Id"] = df['*REQ* Local Student ID']
    df["*REQ* Type"] = 'Student'

    if "*REQ* Entry Date" not in df.columns:
        df["*REQ* Entry Date"] = enrollment_date

    for col in ["Date of Birth", "Gender", "Ethnicity", "Disability Flag", "ELL"]:
        if col not in df.columns:
            df[col] = np.nan

    col_order = [
        'School',
        '*REQ* Student Id',
        '*REQ* Local Student ID',
        '*REQ* First Name',
        '*REQ* Last Name',
        '*REQ* Grade',
        'Date of Birth',
        'Gender',
        'Ethnicity',
        'Disability Flag',
        'ELL',
        '*REQ* Entry Date',
        '*REQ* Type',
    ]

    df = df[col_order]

    return df


def remove_extant_students(df):
    student_df = cysh.get_student_df()
    df = df.loc[
        ~df['*REQ* Local Student ID'].isin(student_df['Local_Student_ID__c'])
    ]
    return df


def input_file(driver, path_to_csv):
    driver.find_element_by_xpath('//*[@id="selectedFile"]').send_keys(path_to_csv)
    driver.find_element_by_xpath('//*[@id="j_id0:j_id42"]/div[3]/div[1]/div[6]/input[2]').click()


def insert_data(driver):
    driver.find_element_by_xpath('//*[@id="startBatchButton"]').click()


def update_student_External_Id(prefix='CPS_'):
    """ Updates 'External_Id__c' field to 'CPS_' + 'Local_Student_ID__c'. 
    Triggers external integrations at HQ.
    """
    student_df = cysh.get_student_df()

    if len(student_df['Local_Student_ID__c'].duplicated()) > 0:
        raise ValueError(f'Error: Duplicates exist on Local_Student_ID__c.')

    student_df = student_df.loc[
        student_df['External_Id__c'].isnull() & 
        (student_df['Local_Student_ID__c'].str.len()==8)
    ]

    if len(student_df) == 0:
        print(f'No students to fix IDs for.')
        return None

    results = []
    for index, row in student_df.iterrows():
        result = cysh.sf.Student__c.update(
            row['Id'],
            {'External_Id__c': (prefix + row['Local_Student_ID__c'])}
        )
        results.append(result)

    return results
