"""
Script used to share City Year data with Thrive.

Requires the following environment variables:
    - THRIVE_HOST
    - THRIVE_USER
    - THRIVE_PASS
"""

import datetime
import os
from pathlib import Path

from dotenv import load_dotenv
import pandas as pd
import pysftp
import numpy as np

from . import simple_cysh as cysh
from .config import set_logger

logger = set_logger(name=Path(__file__).stem)
load_dotenv()

BASE_DIR = Path('Z:/ChiPrivate/Chicago Data and Evaluation/Thrive/')


def load_omni_df():
    # local school reference with various school attributes
    sch_ref_df = cysh.get_sch_ref_df()
    sch_ref_df = sch_ref_df[['CYSH ID', 'CPS ID', 'School', 'Portfolio']]
    sch_ref_df['CPS ID'] = sch_ref_df['CPS ID'].astype(int)

    # Pull Salesforce data
    ISR_df = cysh.get_object_df(
        'Intervention_Session_Result__c', 
        ['Student_Section__c', 'Amount_of_Time__c', 
        'Intervention_Session_Date__c', 'Primary_Skill__c'], 
        rename_id=True
        )
    school_df = cysh.get_object_df(
        'Account', 
        ['Id', 'Name']
        )
    student_df = cysh.get_object_df(
        'Student__c',
        ['Id', 'Local_Student_ID__c', 'Student_Id__c', 'Date_of_Birth__c', 
        'Student_First_Name__c', 'Student_Last_Name__c', 'Grade__c'],
        where=f"School__c IN ({str(school_df['Id'].tolist())[1:-1]})",
        rename_id=True
        )
    student_df['Date_of_Birth__c'] = dt_to_date(student_df['Date_of_Birth__c'])

    stu_sec_df = cysh.get_object_df(
        'Student_Section__c', 
        ['Id', 'Name', 'Active__c', 'Section__c', 'Student__c',
        'Student_Grade__c', 'Intervention_Enrollment_Start_Date__c', 
        'Enrollment_End_Date__c', 'Section_Exit_Reason__c'], 
        rename_id=True, rename_name=True
        )
    stu_sec_df = stu_sec_df.rename(
        columns={'Active__c': 'Student_Section_Active__c'}
        )
    cols = ['Intervention_Enrollment_Start_Date__c', 'Enrollment_End_Date__c']
    for col in cols:
        stu_sec_df[col] = dt_to_date(stu_sec_df[col])

    program_df = cysh.get_object_df(
        'Program__c', 
        ['Id', 'Name'], 
        rename_id=True
        )
    program_df = program_df.rename(columns={'Id': 'Program__c',
    'Name': 'Program'})

    section_df = cysh.get_object_df(
        'Section__c', 
        ['Id', 'Name', 'Active__c', 'School__c', 'Program__c', 
        'Intervention_Primary_Staff__c', 'In_After_School__c',
        'Target_Dosage_Section_Goal__c'], 
        rename_id=True, rename_name=True
        )
    section_df = section_df.rename(columns={'Active__c':'Section_Active__c'})
    col = 'Target_Dosage_Section_Goal__c'
    section_df[col] = section_df[col].replace({0: np.nan})

    account_df = cysh.get_object_df('Account', ['Id', 'Name'])
    account_df = account_df.rename(columns={'Id': 'School__c', 
                                            'Name': 'School'})

    # merge tables
    all_df = (stu_sec_df.merge(section_df, on='Section__c', how='inner')
                        .merge(program_df, on='Program__c', how='left')
                        .merge(student_df, on='Student__c', how='left')
                        .merge(account_df, on='School__c', how='left')
                        .merge(sch_ref_df, on='School', how='left')
                        .merge(ISR_df, on='Student_Section__c', how='left'))

    # filter for sections of interest
    sections = ['Coaching: Attendance', 
                'Tutoring: Math',
                'Tutoring: Literacy', 
                'Homework Assistance', 
                'SEL Check In Check Out']
    all_df = all_df.loc[all_df['Program'].isin(sections)]

    return all_df


def parse_omni_df(all_df):
    data_dict_path = (BASE_DIR / 
                      f"{os.environ['YEAR']} Thrive Program Data Layout.xlsx")

    data_dict = pd.read_excel(data_dict_path,
                              sheet_name='Program Data Elements')
    data_dict = data_dict[['PROGRAM DATA FILE', 'DATA ELEMENTS', 
                           'CY COLUMN NAME', 'CY COLUMN VALUES']]

    # Program ~ Section (Active only?)
    data_file='PROGRAM'
    program_df = convert_table(df=all_df, data_file=data_file, 
                                  data_dict=data_dict)

    # reduce to one section per row
    program_df = program_df.drop_duplicates('PROGRAM_SYSTEM_ID')

    # convert floats to ints, fill NaNs with 0
    for col in ['DELIVERY_OVERALL_DURATION', 'DELIVERY_WEEKS']:
        program_df[col] = program_df[col].fillna(0.0).astype(int)

    # Fill
    # PROGRAM_INTERVENTION_LEVEL   Multiple: Tier2, Tier1 (Homework Assistance)
    program_df['PROGRAM_INTERVENTION_LEVEL'] = "Tier2"
    condition = program_df['PROGRAM_GROUP'] == 'Homework Assistance'
    program_df.loc[condition, 'PROGRAM_INTERVENTION_LEVEL'] = "Tier1"

    # DELIVERY_WEEKS    Multiple: 8 (SEL/Attendance), blank (all others)
    condition = \
        program_df['PROGRAM_GROUP'].str.contains('SEL|Attendance') == True
    program_df.loc[condition, 'DELIVERY_WEEKS'] = 8

    # Attendance ~ ISR
    data_file = 'ATTENDANCE'

    attend_df = convert_table(df=all_df, data_file=data_file, 
                                 data_dict=data_dict)
    attend_df = attend_df.loc[attend_df['ATTENDANCE_DATE'].notna()]
    attend_df = attend_df.drop_duplicates()

    # MEMBERSHIP ~ Student Section
    data_file = 'MEMBERSHIP'

    member_df = convert_table(df=all_df, data_file=data_file, 
                                 data_dict=data_dict)
    member_df = member_df.drop_duplicates('PROGRAM_MEMBERSHIP_SYSTEM_ID')
    member_df['MEMBERSHIP_EXIT_REASONS'] = \
        member_df['MEMBERSHIP_EXIT_REASONS'].str.slice(0, 50)

    # PARTICIPANT ~ Student
    data_file = 'PARTICIPANT'

    partic_df = convert_table(df=all_df, data_file=data_file, 
                                 data_dict=data_dict)
    partic_df = partic_df.drop_duplicates('PARTICIPANT_SYSTEM_ID')

    # FACILITY ~ School
    data_file = 'FACILITY'

    facility_df = convert_table(df=all_df, data_file=data_file, 
                                data_dict=data_dict)
    facility_df = facility_df.dropna().drop_duplicates('FACILITY_SYSTEM_ID')

    return (program_df, attend_df, member_df, partic_df, facility_df)


def write_tables_to_cyconnect(cy_export_dir):
    logger.info(f'Writing Thrive tables to cyconnect at: {cy_export_dir}')

    all_df = load_omni_df()

    program_df, attend_df, member_df, partic_df, facility_df = \
        parse_omni_df(all_df)

    program_df.to_csv(cy_export_dir / 'PROGRAM.csv', index=False)
    attend_df.to_csv(cy_export_dir / 'ATTENDANCE.csv', index=False)
    member_df.to_csv(cy_export_dir / 'MEMBERSHIP.csv', index=False)
    partic_df.to_csv(cy_export_dir / 'PARTICIPANT.csv', index=False)
    facility_df.to_csv(cy_export_dir / 'FACILITY.csv', index=False)

    return


def dt_to_date(series):
    series = pd.to_datetime(series)
    series = pd.to_datetime(series.dt.date)
    return series


def convert_table(df, data_file, data_dict):
    df = df.copy()
    
    data_dict = data_dict.loc[data_dict['PROGRAM DATA FILE']==data_file]

    # Map CY columns to Thrive columns
    rename_dict = {}
    for i, r in data_dict.loc[data_dict['CY COLUMN NAME'].notna()].iterrows():
         rename_dict[r['CY COLUMN NAME']] = r['DATA ELEMENTS']

    df = df.rename(columns=rename_dict)
    
    if 'PROGRAM_MEMBERSHIP_SYSTEM_ID' in rename_dict.values():
        df['PROGRAM_MEMBERSHIP_SYSTEM_ID'] = (df['PROGRAM_SYSTEM_ID'] + "_" +
                                              df['PARTICIPANT_SYSTEM_ID'])
    
    df = df[rename_dict.values()]
    
    # Fill constant values
    for i, r in data_dict.loc[data_dict['CY COLUMN VALUES'].notna()].iterrows():
        if r['CY COLUMN VALUES'].startswith('All: '):
            df[r['DATA ELEMENTS']] = r['CY COLUMN VALUES'].replace('All: ', '')

    return df


def write_tables_to_thrive_sftp(files_dir):
    logger.info('Writing Thrive tables to Thrive SFTP')
    srv = get_srv(host=os.environ['THRIVE_HOST'],
                  username=os.environ['THRIVE_USER'],
                  password=os.environ['THRIVE_PASS'])

    for p in files_dir.iterdir():
        srv.put(p, remotepath=f'./salesforce/{p.name}')


def get_srv(host, username, password):
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None  # returns error without this line
    srv = pysftp.Connection(host=host,
                            username=username,
                            password=password,
                            cnopts=cnopts)
    return srv


def run():
    cy_export_dir = (BASE_DIR / 'exports' / 
                     datetime.datetime.now().strftime(r'%Y.%m.%d'))
    cy_export_dir.mkdir(exist_ok=True)
    write_tables_to_cyconnect(cy_export_dir)
    write_tables_to_thrive_sftp(cy_export_dir)


if __name__ == "__main__":
    run()
