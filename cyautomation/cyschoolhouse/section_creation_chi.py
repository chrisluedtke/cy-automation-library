import os

import pandas as pd

from .config import get_sch_ref_df
from . import simple_cysh as cysh


sch_ref_df = get_sch_ref_df()


def academic_sections_to_create():
    """
    Gather ACM deployment docs to determine which 'Tutoring: Math'
    and 'Tutoring: Literacy' sections to make. Store sections to make in a 
    spreadsheet at 'input_files/section-creator-input.xlsx'.
    """
    xl = pd.ExcelFile(r'Z:/Impact Analytics Team/SY19 ACM Deployment.xlsx')

    df_list = []
    for sheet in xl.sheet_names:
        if sheet != 'Sample Deployment':
            df = xl.parse(sheet).iloc[:,0:6]
            df['Informal Name'] = sheet
            df_list.append(df)
            del df
    acm_dep_df = pd.concat(df_list)

    acm_dep_df.rename(columns={
        'ACM Name':'ACM',
        'Related IA (ELA/Math)':'SectionName'
    }, inplace=True)

    acm_dep_df = acm_dep_df.loc[~acm_dep_df['ACM'].isnull() &
                                ~acm_dep_df['SectionName'].isnull()]

    acm_dep_df['ACM'] = acm_dep_df['ACM'].str.strip()
    acm_dep_df['SectionName'] = acm_dep_df['SectionName'].str.strip().str.upper()

    acm_dep_df.loc[acm_dep_df['SectionName'].str.contains('MATH'), 'SectionName_MATH'] = 'Tutoring: Math'
    acm_dep_df.loc[acm_dep_df['SectionName'].str.contains('ELA'), 'SectionName_ELA'] = 'Tutoring: Literacy'

    acm_dep_df = acm_dep_df.fillna('')

    acm_dep_df = acm_dep_df[['ACM', 'Informal Name', 'SectionName_MATH', 'SectionName_ELA']].groupby(['ACM', 'Informal Name']).agg(lambda x: ''.join(x.unique()))
    acm_dep_df.reset_index(inplace=True)

    acm_dep_df = pd.melt(
        acm_dep_df,
        id_vars=['ACM', 'Informal Name'],
        value_vars=['SectionName_MATH', 'SectionName_ELA'],
        value_name='SectionName'
    )
    acm_dep_df = acm_dep_df.loc[~acm_dep_df['SectionName'].isnull() & (acm_dep_df['SectionName'] != '')]
    acm_dep_df = acm_dep_df.sort_values('ACM')
    acm_dep_df['key'] = acm_dep_df['ACM'] + acm_dep_df['SectionName']

    section_df = cysh.get_section_df(sections_of_interest=['Tutoring: Literacy', 'Tutoring: Math'])
    staff_df = cysh.get_staff_df()
    section_df = section_df.merge(staff_df, how='left', left_on='Intervention_Primary_Staff__c', right_on='Staff__c')
    section_df['key'] = section_df['Staff__c_Name'] + section_df['Program__c_Name']

    acm_dep_df = acm_dep_df.loc[
        acm_dep_df['ACM'].isin(staff_df['Staff__c_Name']) &
        ~acm_dep_df['key'].isin(section_df['key'])
    ]

    df = acm_dep_df.merge(sch_ref_df[['School', 'Informal Name']],
                          how='left', on='Informal Name')

    df = format_and_write_xl(df)

    return df


def non_CP_sections_to_create(sections_of_interest=['Coaching: Attendance', 'SEL Check In Check Out']):
    """
    Produce table of sections to create, with the assumption that all 'Corps Member' roles should have 1 of each section.
    """
    section_df = cysh.get_section_df(sections_of_interest)
    section_df['key'] = section_df['Intervention_Primary_Staff__c'] + section_df['Program__c_Name']

    staff_df = cysh.get_object_df('Staff__c', ['Id', 'Name', 'Role__c', 'Organization__c'], where="Site__c='Chicago'", rename_name=True)
    school_df = cysh.get_object_df('Account', ['Id', 'Name'])
    school_df.rename(columns={'Id':'School__c', 'Name':'School'}, inplace=True)
    staff_df = staff_df.merge(school_df, how='left', left_on='Organization__c', right_on='School__c')

    acm_df = staff_df.loc[staff_df['Role__c'].str.contains('Corps Member')==True].copy()
    acm_df['key'] = 1

    section_deployment = pd.DataFrame.from_dict({'SectionName': sections_of_interest})
    section_deployment['key'] = 1

    acm_df = acm_df.merge(section_deployment, on='key')
    acm_df['key'] = acm_df['Id'] + acm_df['SectionName']

    df = acm_df.loc[~acm_df['key'].isin(section_df['key'])]

    df = df.rename(columns={'Staff__c_Name':'ACM'})

    df = format_and_write_xl(df)

    return df


def MIRI_sections_to_create():
    """
    Produce table of ACM 'Math Inventory' and 'Reading Inventory' sections to 
    make, only relevant to high schools (in Chicago)
    """
    program_df = cysh.get_object_df('Program__c', ['Id', 'Name'], rename_id=True, rename_name=True)

    school_df = cysh.get_object_df('Account', ['Id', 'Name'])
    school_df.rename(columns={'Id':'School__c', 'Name':'School'}, inplace=True)

    staff_df = cysh.get_object_df('Staff__c', ['Id', 'Name'], where="Site__c='Chicago'", rename_name=True)

    section_cols = ['Id', 'Name', 'Intervention_Primary_Staff__c', 'School__c',
                    'Program__c']
    section_df = cysh.get_object_df('Section__c', section_cols, rename_id=True, rename_name=True)
    section_df = section_df.merge(school_df, how='left', on='School__c')
    section_df = section_df.merge(program_df, how='left', on='Program__c')
    section_df = section_df.merge(staff_df, how='left', left_on='Intervention_Primary_Staff__c', right_on='Id')

    # TODO: don't store data like this here, put it in the school reference
    highschools = [
        'Tilden Career Community Academy High School',
        'Gage Park High School',
        'Collins Academy High School',
        'Schurz High School',
        'Sullivan High School',
        'Chicago Academy High School',
        'Roberto Clemente Community Academy',
        'Wendell Phillips Academy',
    ]

    section_df = section_df.loc[section_df['School'].isin(highschools)]

    condition = section_df['Program__c_Name'].str.contains('Inventory')==True
    miri_section_df = section_df.loc[condition]

    condition = section_df['Program__c_Name'].str.contains('Tutoring')==True
    section_df = section_df.loc[condition]

    section_df['Program__c_Name'] = section_df['Program__c_Name'].map({
        'Tutoring: Literacy':'Reading Inventory',
        'Tutoring: Math':'Math Inventory'
    })

    for df in [section_df, miri_section_df]:
        df['key'] = df['Staff__c_Name'] + df['Program__c_Name']

    df = (section_df.loc[~section_df['key'].isin(miri_section_df['key'])]
                    .rename(columns={
                        'Staff__c_Name':'ACM',
                        'Program__c_Name':'SectionName'
                    }))

    df = format_and_write_xl(df)

    return df


def format_and_write_xl(df, start_date='09/04/2018', end_date='06/07/2019',
                        in_sch_ext_lrn='In School', target_dosage=0):

    df = pd.DataFrame({
        'School':df.School,
        'ACM':df.ACM,
        'SectionName':df.SectionName,
        'In_School_or_Extended_Learning':in_sch_ext_lrn,
        'Start_Date':start_date,
        'End_Date':end_date,
        'Target_Dosage':target_dosage
    })

    write_path = os.path.join(os.path.dirname(__file__),
                              'input_files/section-creator-input.xlsx')
    df.to_excel(write_path, index=False)

    return df


def deactivate_all_sections(section_type):
    """
    This is necessary due to a bug in section creation. When section creation fails,
    a `50 Acts of Greatness` section is made, as the default section type selection.
    We don't provide this programming in Chicago, so we can safely deactivate all.
    """
    section_cols = ['Id', 'Name', 'Intervention_Primary_Staff__c',
                    'School__c', 'Program__c', 'Active__c']
    section_df = cysh.get_object_df('Section__c', section_cols, rename_id=True, rename_name=True)
    program_df = cysh.get_object_df('Program__c', ['Id', 'Name'], rename_id=True, rename_name=True)

    df = section_df.merge(program_df, how='left', on='Program__c')

    df = df.loc[
        (df['Program__c_Name']==section_type) &
        (section_df['Active__c']==True),
        'Section__c'
    ]

    print(f"{len(df)} {section_type} sections to de-activate.")
    user_input = input("Are you sure? (yes/y to continue): ").lower()

    if user_input in ['yes', 'y']:
        for section_id in df:
            cysh.sf.Section__c.update(section_id, {'Active__c':False})
        return True
    else:
        return False
