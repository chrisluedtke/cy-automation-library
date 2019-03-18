import functools
from pathlib import Path

import pandas as pd
from simple_salesforce import (
    Salesforce,
    SalesforceExpiredSession,
    SalesforceMalformedRequest
)

from .config import *

__all__ = [
    'init_sf_session',
    'get_object_df',
    'get_object_fields',
    'get_section_df',
    'get_student_section_staff_df',
    'get_staff_df',
    'sf'
]

def init_sf_session():
    sf = Salesforce(
        instance_url=SF_URL,
        password=SF_PASS,
        username=SF_USER,
        security_token=SF_TOKN
    )
    return sf

sf = init_sf_session()

def check_sf_session(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        global sf
        try:
            return func(*args, **kwargs)
        except SalesforceExpiredSession:
            sf = init_sf_session()
            return func(*args, **kwargs)

    return wrapper

@check_sf_session
def get_object_fields(object_name):
    one_id = sf.query(f"SELECT Id FROM {object_name} LIMIT 1")['records']
    if one_id:
        one_id = one_id[0]['Id']
    else:
        raise ValueError('Could not pull a record Id in order to identify '
                         f'fields of object: {object_name}')
    response = getattr(sf, object_name).get(one_id)
    fields = sorted(list(response.keys()))
    fields.remove('attributes')

    return fields

@check_sf_session
def get_object_df(object_name, field_list=None, where=None, rename_id=False,
                  rename_name=False, year='SY19'):
    if year=='SY19':
        if not field_list:
            field_list = get_object_fields(object_name)

        querystring = f"SELECT {', '.join(field_list)} FROM {object_name}"

        if where:
            querystring += f" WHERE {where}"

        query_return = sf.query_all(querystring)

        query_list = []
        for row in query_return['records']:
            record = []
            for column in field_list:
                col_data = row[column]
                record.append(col_data)
            query_list.append(record)

        df = pd.DataFrame(query_list, columns=field_list)

    else:
        valid_years = ['SY17', 'SY18', 'SY19']
        if year not in valid_years:
            raise ValueError(f"The year provided ({year}) must be one of:"
                             f"{', '.join(valid_years)}.")

        df = pd.read_csv('Z:/ChiPrivate/Chicago Data and Evaluation/'
                         'Whole Site End of Year Data/Salesforce Objects/'
                         f'{year}/{object_name}.csv')
        if field_list is not None:
            df = df[field_list]

    if rename_id==True:
        df = df.rename(columns={'Id':object_name})
    if rename_name==True:
        df = df.rename(columns={'Name':(object_name+'_Name')})

    return df


def get_section_df(sections_of_interest):
    if type(sections_of_interest)==str:
        sections_of_interest = list(sections_of_interest)

    program_df = get_object_df(
        'Program__c', ['Id', 'Name'],
        where=f"Name IN ({str(sections_of_interest)[1:-1]})",
        rename_id=True, rename_name=True)

    section_df = get_object_df(
        'Section__c',
        ['Id', 'Name', 'Intervention_Primary_Staff__c', 'Program__c'],
        rename_id=True, rename_name=True,
        where=f"Program__c IN ({str(program_df['Program__c'].tolist())[1:-1]})",
    )

    df = section_df.merge(program_df, how='left', on='Program__c')

    return df


def get_student_section_staff_df(sections_of_interest):
    if type(sections_of_interest) == str:
        sections_of_interest = [sections_of_interest]

    # load salesforce tables
    program_df = get_object_df(
        'Program__c',
        ['Id', 'Name'],
        where=f"Name IN ({str(sections_of_interest)[1:-1]})",
        rename_id=True, rename_name=True
    )

    stu_sect_cols = [
        'Id', 'Name', 'Student_Program__c', 'Program__c', 'Section__c',
        'Active__c', 'Enrollment_End_Date__c', 'Student__c',
        'Student_Name__c', 'Dosage_to_Date__c', 'School_Reference_Id__c',
        'Student_Grade__c', 'School__c'
    ]
    stu_sect_df = get_object_df(
        'Student_Section__c', stu_sect_cols,
        where=f"Program__c IN ({str(program_df['Program__c'].tolist())[1:-1]})",
        rename_id=True, rename_name=True
    )

    section_df = get_object_df(
        'Section__c',
        ['Id', 'Intervention_Primary_Staff__c'],
        where=f"Program__c IN ({str(program_df['Program__c'].tolist())[1:-1]})",
        rename_id=True
    )
    staff_df = get_object_df('Staff__c', ['Id', 'Name'], rename_id=True,
                             rename_name=True)

    # merge salesforce tables
    df = (stu_sect_df.merge(section_df, how='left', on='Section__c')
                     .merge(staff_df, how='left',
                            left_on='Intervention_Primary_Staff__c',
                            right_on='Staff__c')
                     .merge(program_df, how='left', on='Program__c'))

    return df


def get_staff_df():
    school_df = get_object_df('Account', ['Id', 'Name'])
    school_df = school_df.rename(columns={'Id':'Organization__c',
                                          'Name':'School'})

    staff_cols = ['Id', 'Individual__c', 'Name', 'First_Name_Staff__c',
                  'Staff_Last_Name__c', 'Role__c', 'Email__c',
                  'Organization__c']
    schools_q = str(school_df['Organization__c'].tolist())[1:-1]
    staff_df = get_object_df(
        'Staff__c',
        staff_cols,
        where=f"Organization__c IN ({schools_q})",
        rename_name=True,
        rename_id=True
    )

    staff_df = staff_df.merge(school_df, how='left', on='Organization__c')

    return staff_df


@check_sf_session
def object_reference():
    sf_describe = sf.describe()
    return {object['name']:object['label'] for object in result['sobjects']}
