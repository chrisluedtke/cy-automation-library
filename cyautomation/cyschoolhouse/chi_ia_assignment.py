from pathlib import Path

import pandas as pd
from datetime import datetime

from . import simple_cysh as cysh
from .config import INPUT_PATH
from .PageObjects.implementations import IndicatorAreaEnrollment

__all__ = [
    'get_ia_to_assign'
]

SECTION_IA_DICT = {
    'Coaching: Attendance': 'Attendance',
    'SEL Check In Check Out': 'Behavior',
    'Tutoring: Math': 'Math',
    'Tutoring: Literacy': 'ELA/Literacy'
}

def get_ia_to_assign():
    """
    Writes an Excel template that contains indicator areas to be assigned.
    That template is then processed by a separate function to actually
    make those assignments.
    """
    df = get_student_enrollment_details()
    df2 = get_ia_enrollment_details()
    # sort by active in order to keep active row if it exists
    df = (df.merge(df2, how='left', on='Student_Program')
            .sort_values('Active__c', ascending=False)
            .drop_duplicates('Student_Program'))

    assmt_df = get_assessment_details()

    df = assign_ia_col(df,
                       section_type='Coaching: Attendance',
                       assmt_df=assmt_df,
                       assmt_name='Reporting Period ADA Tracker - ATTENDANCE',
                       min_days_active=56)

    df = assign_ia_col(df,
                       section_type='SEL Check In Check Out',
                       assmt_df=assmt_df,
                       assmt_name='DESSA 40',
                       min_days_active=56)

    df = assign_ia_col(df,
                       section_type='Tutoring: Math',
                       assmt_df=assmt_df,
                       assmt_name='NWEA - MATH',
                       assmt_prior_to='2019-10-01',
                       min_time=1,
                       is_active=True)

    df = assign_ia_col(df,
                       section_type='Tutoring: Literacy',
                       assmt_df=assmt_df,
                       assmt_name='NWEA - ELA',
                       assmt_prior_to='2019-10-01',
                       min_time=1,
                       is_active=True)

    #remove rows where indicator area already assigned
    df = (df.loc[df['Indicator_Area_Type__c'].isnull()
                 & ~df['Assign Indicator Area'].isnull()]
            .rename(columns={
                    'Student__c':'Student: Student ID',
                    'Grade__c':'Student: Grade',
                    'Student_Last_Name__c':'Student: Student Last Name',
                    'Assign Indicator Area':'Indicator Area',
            }))

    cols =['School', 'Student: Student ID', 'Student: Grade',
           'Student: Student Last Name', 'Indicator Area']
    df = df[cols].sort_values(cols)

    write_path = str(Path(INPUT_PATH) / 'indicator_area_roster.xlsx')
    df.to_excel(write_path, index=False)

    return df


def get_student_enrollment_details():
    sects = ['Coaching: Attendance', 'Tutoring: Literacy',
             'Tutoring: Math', 'SEL Check In Check Out']
    section_df = cysh.get_section_df(sects)

    program_df = cysh.get_object_df(
        'Program__c', ['Id', 'Name'],
        where=f"Name IN ({str(sects)[1:-1]})",
        rename_id=True, rename_name=True
    )

    stu_sec_df = cysh.get_object_df(
        'Student_Section__c',
        ['Id', 'Active__c', 'Section__c', 'Student__c', 'Amount_of_Time__c',
         'Intervention_Enrollment_Start_Date__c', 'Enrollment_End_Date__c'],
        rename_id=True,
        where=f"Program__c IN ({str(program_df['Program__c'].tolist())[1:-1]})",
    )
    stu_sec_df.loc[:, 'Intervention_Enrollment_Start_Date__c'] = \
        pd.to_datetime(stu_sec_df['Intervention_Enrollment_Start_Date__c'])
    stu_sec_df.loc[:,'Enrollment_End_Date__c'] = \
        (pd.to_datetime(stu_sec_df['Enrollment_End_Date__c'])
           .fillna(pd.to_datetime(str(datetime.now()))))

    school_df = cysh.get_object_df(
        'Account',
        ['Id', 'Name']
    )
    school_df = school_df.rename(columns={'Id':'School__c',
                                          'Name':'School'})

    student_df = cysh.get_object_df(
        'Student__c',
        ['Id', 'Name', 'Student_First_Name__c', 'Student_Last_Name__c',
         'School__c', 'Grade__c'],
        where=f"School__c IN ({str(school_df['School__c'].tolist())[1:-1]})",
        rename_id=True, rename_name=True
    )

    df = (stu_sec_df.merge(section_df, how='left', on='Section__c')
                    .merge(student_df, how='left', on='Student__c')
                    .merge(school_df, how='left', on='School__c'))

    df['Student_Program'] = df['Student__c'] + "_" + df['Program__c']
    df = df.set_index('Student_Program')

    df_aggs = (df.groupby('Student_Program')
                 .agg({'Intervention_Enrollment_Start_Date__c':'min',
                       'Enrollment_End_Date__c':'max',
                       'Amount_of_Time__c':'sum'}))
    df.update(df_aggs)
    df = df.reset_index()

    df['Days Active'] = (df['Enrollment_End_Date__c'] -
                         df['Intervention_Enrollment_Start_Date__c']).dt.days

    return df


def get_ia_enrollment_details():
    # Get IA's for each student, and include Program to match with sections
    stud_ia_df = cysh.get_object_df(
        'Indicator_Area_Student__c',
        ['Id', 'Student__c', 'Indicator_Area__c'],
        rename_id=True,
    #    where="Active__c = True"
    )

    ia_df = cysh.get_object_df(
        'Indicator_Area__c',
        ['Id', 'Indicator_Area_Type__c'],
        rename_id=True
    )

    ia_section_dict = {v:k for k, v in SECTION_IA_DICT.items()}
    ia_df['Indicator_Area_Type__c'] = \
        ia_df['Indicator_Area_Type__c'].map(ia_section_dict)

    program_df = cysh.get_object_df(
        'Program__c',
        ['Id', 'Name'],
        rename_id=True,
        rename_name=True
    )

    cols = ['Student_Program', 'Indicator_Area__c', 'Indicator_Area_Type__c']
    df = (stud_ia_df.merge(ia_df, on='Indicator_Area__c', how='left')
                    .merge(program_df, left_on='Indicator_Area_Type__c',
                           right_on='Program__c_Name', how='left')
                    .assign(Student_Program = lambda x: (
                                x['Student__c'] + "_" + x['Program__c']))
                    .loc[:, cols])

    return df


def get_assessment_details():
    df = cysh.get_object_df(
        'Assesment__c',
        ['Id', 'Type__c', 'Date_Administered__c',
         'X0_to_300_Scaled_Score__c', 'Student__c',
         'Average_Daily_Attendance__c', 'SEL_Composite_T_Score__c'],
        rename_id=True
    )

    assmt_types = cysh.get_object_df('Picklist_Value__c', ['Id', 'Name'])
    assmt_types = assmt_types.rename(columns={'Id':'Type__c',
                                              'Name':'Assessment Type'})

    score_cols = ['X0_to_300_Scaled_Score__c', 'Average_Daily_Attendance__c',
                  'SEL_Composite_T_Score__c']
    df = (df.merge(assmt_types, how='left', on='Type__c')
            .drop(columns=['Type__c'])
            .assign(Score = lambda x: x[score_cols].sum(axis=1)))

    df = df.loc[df['Score'] > 0]

    return df


def assign_ia_col(df, section_type, assmt_df, assmt_name, assmt_prior_to=None,
                  is_active=False, min_time=None, min_days_active=None):
    """Fills a column with IAs to assign for a given student based on Chicago's rules"""
    mask = assmt_df['Assessment Type'] == assmt_name

    if assmt_prior_to:
        mask = mask & (assmt_df['Date_Administered__c'] < assmt_prior_to)

    stu_with_assmt = assmt_df.loc[mask, 'Student__c']

    mask = (df['Program__c_Name'] == section_type)

    if 'Tutoring:' in section_type:
        mask = mask & (df['Student__c'].isin(stu_with_assmt) |
                       (df['Grade__c'].astype(int) > 8))
    else:
        mask = mask & df['Student__c'].isin(stu_with_assmt)

    if min_days_active:
        mask = mask & (df['Days Active'] > min_days_active)

    if min_time:
        mask = mask & (df['Amount_of_Time__c'] >= min_time)

    if is_active:
        mask = mask & (df['Active__c'] == True)

    df.loc[mask, 'Assign Indicator Area'] = SECTION_IA_DICT[section_type]

    return df
