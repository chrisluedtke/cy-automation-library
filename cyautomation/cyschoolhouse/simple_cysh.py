import logging
import functools
from pathlib import Path

import pandas as pd
from simple_salesforce import (Salesforce, SalesforceExpiredSession,
                               SalesforceMalformedRequest)

from .config import SF_PASS, SF_TOKN, SF_URL, SF_USER, YEAR
from .utils import get_sch_ref_df


def init_sf_session():
    sf = Salesforce(
        instance_url=SF_URL,
        password=SF_PASS,
        username=SF_USER,
        security_token=SF_TOKN
    )
    return sf


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
                  rename_name=False, archive_year=None):
    if archive_year:
        archive_year = archive_year.upper()
        archive_years = ['SY17', 'SY18', 'SY19']
        if archive_year not in archive_years:
            raise ValueError(f"Invalid archive_year. Try one of: "
                             f"{', '.join(archive_years)}.")
        print(f'Loading {object_name} from {archive_year} archive')
        df = pd.read_csv('Z:/ChiPrivate/Chicago Data and Evaluation/'
                         'Whole Site End of Year Data/Salesforce Objects/'
                         f'{archive_year}/{object_name}.csv')
        if field_list:
            df = df[field_list]
    else:
        if not field_list:
            field_list = get_object_fields(object_name)

        querystring = f"SELECT {', '.join(field_list)} FROM {object_name}"

        if where:
            querystring += f" WHERE {where}"

        if querystring.endswith('()'):
            query_return = {'records': []}
        else:
            query_return = sf.query_all(querystring)

        if query_return['records']:
            df = pd.DataFrame(query_return['records'])
            df = df[field_list]
        else:
            logging.warn(f'No records found for query:\n  {querystring}')
            df = pd.DataFrame(columns=field_list)

    if rename_id:
        df = df.rename(columns={'Id':object_name})
    if rename_name:
        df = df.rename(columns={'Name':(object_name+'_Name')})

    return df


def get_section_df(sections_of_interest):
    if isinstance(sections_of_interest, str):
        sections_of_interest = [sections_of_interest]

    program_df = get_object_df(
        'Program__c', ['Id', 'Name'],
        where=f"Name IN {in_str(sections_of_interest)}",
        rename_id=True, rename_name=True)

    section_df = get_object_df(
        'Section__c',
        ['Id', 'Name', 'Intervention_Primary_Staff__c', 'Program__c'],
        rename_id=True, rename_name=True,
        where=f"Program__c IN {in_str(program_df['Program__c'])}",
    )

    df = section_df.merge(program_df, how='left', on='Program__c')

    return df


def get_student_section_staff_df(sections_of_interest, schools=None):
    if isinstance(sections_of_interest, str):
        sections_of_interest = [sections_of_interest]
    if schools and isinstance(schools, str):
        schools = [schools]

    # load salesforce tables
    program_df = get_object_df(
        'Program__c',
        ['Id', 'Name'],
        where=f"Name IN {in_str(sections_of_interest)}",
        rename_id=True, rename_name=True
    )

    stu_sect_cols = [
        'Id', 'Name', 'Student_Program__c', 'Program__c', 'Section__c',
        'Active__c', 'Enrollment_End_Date__c', 'Student__c',
        'Student_Name__c', 'Dosage_to_Date__c', 'School_Reference_Id__c',
        'Student_Grade__c', 'School__c'
    ]
    where = f"Program__c IN {in_str(program_df['Program__c'])}"
    if schools:
        where = f"({where} AND School__c IN {in_str(schools)})"
    stu_sect_df = get_object_df(
        'Student_Section__c', stu_sect_cols,
        where=where,
        rename_id=True, rename_name=True
    )

    section_df = get_object_df(  # TODO: filter this table by schools as well
        'Section__c',
        ['Id', 'Intervention_Primary_Staff__c'],
        where=f"Program__c IN {in_str(program_df['Program__c'])}",
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


def get_staff_df(schools=None, roles=None):
    """
    schools: List of schools as named in salesforce
    roles: List of roles as named in salesforce
    """
    if schools and type(schools) == str:
        schools = [schools]

    if roles and type(roles) == str:
        roles = [roles]

    where = f"Name IN {in_str(schools)}" if schools else None
    school_df = get_object_df('Account', ['Id', 'Name'], where=where)
    school_df = school_df.rename(columns={'Id': 'Organization__c',
                                          'Name': 'School'})

    staff_cols = ['Id', 'Individual__c', 'Name', 'First_Name_Staff__c',
                  'Staff_Last_Name__c', 'Role__c', 'Email__c',
                  'Organization__c']
    where = f"Organization__c IN {in_str(school_df['Organization__c'])}"

    if roles:
        where = f"({where} AND Role__c IN {in_str(roles)})"
    staff_df = get_object_df('Staff__c', staff_cols, where=where,
                             rename_name=True, rename_id=True)

    staff_df = staff_df.merge(school_df, how='left', on='Organization__c')

    return staff_df


def get_student_df(schools=None):
    """
    Args:
        - schools: List of schools as named in salesforce
    """
    if type(schools) == str:
        schools = [schools]

    where = f"Name IN {in_str(schools)}" if schools else None
    school_df = get_object_df('Account', ['Id', 'Name'], where=where)
    school_df = school_df.rename(columns={'Id': 'Organization__c',
                                          'Name': 'School'})

    student_df = get_object_df(
        'Student__c',
        ['Id', 'Local_Student_ID__c', 'External_Id__c'],
        where=f"School__c IN {in_str(school_df['Organization__c'])}",
        rename_name=True,
        rename_id=True
    )

    # student_df = student_df.merge(school_df, how='left', on='Organization__c')

    return student_df


@check_sf_session
def object_reference():
    result = sf.describe()
    return {obj['name']:obj['label'] for obj in result['sobjects']}


def in_str(ls):
    """ Formats a list to pass into a SOQL "WHERE ... IN ..." statement
    """
    if not isinstance(ls, list):
        ls = list(ls)

    for i in range(len(ls)):
        if isinstance(ls[i], str):
           ls[i] = ls[i].replace("'", "*")

    return f"({str(ls)[1:-1]})".replace("*", "\\'")


sf = init_sf_session()
