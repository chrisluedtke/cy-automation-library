import os

import pandas as pd

from .config import get_sch_ref_df
from . import simple_cysh as cysh


ACM_DEPLOY_PATH = r'Z:/Impact Analytics Team/SY20/SY20 ACM Deployment.xlsx'

def academic_sections_to_create(start_date, end_date):
    """ Reads ACM deployment spreadsheet to determine which 'Tutoring: Math'
    and 'Tutoring: Literacy' sections to make.
    """
    acm_dep_df = pd.read_excel(ACM_DEPLOY_PATH)

    acm_dep_df = acm_dep_df.rename(columns={
        'ACM Name (First Last)': 'Staff__c_Name',
        'Related IA (ELA/Math)': 'SectionName',
        'ACM ID': 'Staff__c'
    })

    acm_dep_df = acm_dep_df.loc[acm_dep_df['Staff__c_Name'].notna()]

    acm_dep_df['Staff__c_Name'] = acm_dep_df['Staff__c_Name'].str.strip()
    acm_dep_df['SectionName'] = acm_dep_df['SectionName'].str.strip().str.upper()

    acm_dep_df.loc[acm_dep_df['SectionName'].str.contains('MATH'), 
                   'SectionName_MATH'] = 'Tutoring: Math'
    acm_dep_df.loc[acm_dep_df['SectionName'].str.contains('ELA'), 
                   'SectionName_ELA'] = 'Tutoring: Literacy'

    acm_dep_df = pd.melt(
        acm_dep_df,
        id_vars=['Staff__c_Name', 'Staff__c'],
        value_vars=['SectionName_MATH', 'SectionName_ELA'],
        value_name='Program__c_Name'
    )
    acm_dep_df = acm_dep_df.loc[acm_dep_df['Program__c_Name'].notna()]

    # Filter out existing sections
    acm_dep_df['key'] = (acm_dep_df['Staff__c'] + '_' +  
                         acm_dep_df['Program__c_Name'])

    sections_of_interest=['Tutoring: Literacy', 'Tutoring: Math']
    section_df = cysh.get_section_df(sections_of_interest)
    section_df['key'] = (section_df['Intervention_Primary_Staff__c'] + '_' + 
                         section_df['Program__c_Name'])

    acm_dep_df = acm_dep_df.loc[~acm_dep_df['key'].isin(section_df['key'])]

    # inner-join on staff name to merge in School
    staff_df = cysh.get_staff_df()
    acm_dep_df = acm_dep_df.merge(staff_df[['Staff__c_Name', 'School']],
                                  on='Staff__c_Name')

    acm_dep_df = format_df(acm_dep_df, start_date=start_date, 
                           end_date=end_date, 
                           in_sch_ext_lrn='In School', 
                           target_dosage=900)

    return acm_dep_df


def non_academic_sections_to_create(start_date, end_date):
    """ Produces table of sections to create, with the assumption that all 
    'Corps Member' roles should have 1 of each non-academic section.
    """
    sections_of_interest = ['Coaching: Attendance', 'SEL Check In Check Out']

    # Assign all Corps Member staff to all sections of interest
    staff_df = cysh.get_staff_df()
    staff_df = staff_df.loc[
        staff_df['Role__c'].str.contains('Corps Member')==True
    ]
    sections_to_make = pd.DataFrame({'SectionName': sections_of_interest})
    staff_df = (staff_df.assign(key=1)
                        .merge(sections_to_make.assign(key=1), on='key')
                        .drop(columns='key'))

    # Filter out existing sections
    staff_df['key'] = staff_df['Staff__c'] + staff_df['SectionName']

    section_df = cysh.get_section_df(sections_of_interest)
    section_df['key'] = (section_df['Intervention_Primary_Staff__c'] + 
                         section_df['Program__c_Name'])

    staff_df = staff_df.loc[~staff_df['key'].isin(section_df['key'])]

    staff_df = format_df(staff_df, start_date=start_date, end_date=end_date)

    return staff_df


def MIRI_sections_to_create(start_date, end_date):
    """ Produces table of ACM 'Math Inventory' and 'Reading Inventory' sections
    to make. Only relevant to high schools.
    """
    # Cross reference course performane and MIRI sections to determine which
    # MIRI sections to make
    cp_sections = ['Tutoring: Math', 'Tutoring: Literacy']
    miri_secitons = ['Math Inventory', 'Reading Inventory']

    sect_df = cysh.get_section_df(cp_sections + miri_secitons)

    cp_sect_df = sect_df.loc[sect_df['Program__c_Name'].isin(cp_sections)]
    cp_sect_df.loc[:, 'Program__c_Name'] = cp_sect_df['Program__c_Name'].map({
        'Tutoring: Literacy': 'Math Inventory',
        'Tutoring: Math': 'Math Inventory'
    })

    miri_sect_df = sect_df.loc[sect_df['Program__c_Name'].isin(miri_secitons)]

    for df in [cp_sect_df, miri_sect_df]:
        df.loc[:, 'key'] = (df['Intervention_Primary_Staff__c'] + '_' +
                            df['Program__c_Name'])

    df = cp_sect_df.loc[~cp_sect_df['key'].isin(miri_sect_df['key'])]

    # Merge in staff name and school
    staff_df = cysh.get_staff_df()
    df = df.merge(staff_df[['Staff__c', 'Staff__c_Name', 'School']], 
                  how='left', left_on='Intervention_Primary_Staff__c', 
                  right_on='Staff__c')

    # Filter to only high schools
    sch_ref_df = get_sch_ref_df()
    highschools = sch_ref_df.loc[sch_ref_df['GradeLevel']=="High", "School"]
    df = df.loc[df['School'].isin(highschools)]

    df = format_df(df, start_date=start_date, end_date=end_date)

    return df


def format_df(df, start_date, end_date, in_sch_ext_lrn='In School', 
              target_dosage=0):
    assert in_sch_ext_lrn in {'In School', 'Extended Learning', 'Curriculum'}

    df = df.rename(columns={
        'Staff__c_Name': 'ACM',
        'Program__c_Name': 'SectionName'
    })

    df = df[['School', 'ACM', 'SectionName']]
    df['In_School_or_Extended_Learning'] = in_sch_ext_lrn
    df['Start_Date'] = start_date
    df['End_Date'] = end_date
    df['Target_Dosage'] = target_dosage

    return df


def deactivate_all_sections(section_type):
    """
    This is necessary due to a bug in section creation. When section creation fails,
    a `50 Acts of Greatness` section is made, as the default section type selection.
    We don't provide this programming in Chicago, so we can safely deactivate all.
    """
    section_df = cysh.get_object_df(
        'Section__c', 
        ['Id', 'Program__c', 'Active__c'], 
        rename_id=True, 
        rename_name=True
    )
    program_df = cysh.get_object_df(
        'Program__c', 
        ['Id', 'Name'], 
        rename_id=True,
        rename_name=True
    )
    section_df = section_df.merge(program_df, how='left', on='Program__c')

    section_df = section_df.loc[
        (section_df['Program__c_Name']==section_type) &
        (section_df['Active__c']==True),
        'Section__c'
    ]

    print(f"{len(section_df)} {section_type} sections to de-activate.")
    user_input = input("Are you sure? (yes/y to continue): ").lower()

    if user_input in {'yes', 'y'}:
        for section_id in section_df:
            cysh.sf.Section__c.update(section_id, {'Active__c':False})
        return True
    else:
        return False
