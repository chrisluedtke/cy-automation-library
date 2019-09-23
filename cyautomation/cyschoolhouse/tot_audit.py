import os

import pandas as pd
# from StyleFrame import StyleFrame, Styler, utils

from .config import get_sch_ref_df
from . import simple_cysh as cysh


def fix_T1T2ELT(sf=cysh.sf):
    """ Standardize common spellings of "T1" "T2" and "ELT"
    """

    typo_map = {
        r'([Tt]([Ii][Ee]|[Ee][Ii])[Rr] ?|t)(1|[Oo]ne)':'T1',
        r'([Tt]([Ii][Ee]|[Ee][Ii])[Rr] ?|t)(2|[Tt]wo)':'T2',
        #r'([Aa]fter ?[Ss]chool|ASP)':'ELT',
    }

    all_typos = '|'.join(list(typo_map.keys()))

    df = cysh.get_object_df('Intervention_Session__c', ['Id', 'Comments__c'],
                            rename_id=True)
    df['Comments__c'].fillna('', inplace=True)
    df = df.loc[df['Comments__c'].str.contains(all_typos)]

    df['Comments__c'].replace(typo_map, regex=True, inplace=True)

    print(f"Found {len(df)} T1, T2, or ELT labels that can be fixed")

    results = []
    for index, row in df.iterrows():
        result = sf.Intervention_Session__c.update(row.Intervention_Session__c,
                                                   {'Comments__c':row['Comments__c']})
        results.append(results)

    return results

def get_error_table():
    ISR_df = cysh.get_object_df('Intervention_Session_Result__c',
                                ['Amount_of_Time__c', 'IsDeleted',
                                'Intervention_Session_Date__c',
                                'Related_Student_s_Name__c',
                                'Intervention_Session__c', 'CreatedDate'])
    IS_df = cysh.get_object_df('Intervention_Session__c',
                               ['Id', 'Name', 'Comments__c', 'Section__c'],
                               rename_id=True, rename_name=True)
    section_df = cysh.get_object_df('Section__c', ['Id', 'School__c', 'Intervention_Primary_Staff__c', 'Program__c'], rename_id=True)

    school_df = cysh.get_object_df('Account', ['Id', 'Name'])
    school_df.rename(columns={'Id':'School__c', 'Name':'School_Name__c'}, inplace=True)

    staff_df = cysh.get_object_df('Staff__c', ['Id', 'Name'], where="Site__c = 'Chicago'", rename_id=True, rename_name=True)

    program_df = cysh.get_object_df('Program__c', ['Id', 'Name'], rename_id=True, rename_name=True)

    df = ISR_df.merge(IS_df, how='left', on='Intervention_Session__c'); del df['Intervention_Session__c']
    df = df.merge(section_df, how='left', on='Section__c'); del df['Section__c']
    df = df.merge(school_df, how='left', on='School__c'); del df['School__c']
    df = df.merge(staff_df, how='left', left_on='Intervention_Primary_Staff__c', right_on='Staff__c'); del df['Intervention_Primary_Staff__c'], df['Staff__c']
    df = df.merge(program_df, how='left', on='Program__c'); del df['Program__c']

    df['Intervention_Session_Date__c'] = pd.to_datetime(df['Intervention_Session_Date__c']).dt.date
    df['CreatedDate'] = pd.to_datetime(df['CreatedDate']).dt.date
    df['Comments__c'].fillna('', inplace=True)

    df.loc[df['Program__c_Name'].str.contains('Tutoring')
           & ~df['Comments__c'].str.contains('T1|T2'), 'Missing T1/T2 Code'] = 'Missing T1/T2 Code'

    df.loc[df['Program__c_Name'].str.contains('Tutoring')
           & df['Comments__c'].str.contains('T1')
           & df['Comments__c'].str.contains('T2'), 'Listed T1 and T2'] = 'Listed T1 and T2'

    df.loc[df['Program__c_Name'].str.contains('Tutoring')
           & (df['Amount_of_Time__c'] < 10), '<10 Minutes'] = '<10 Minutes'

    df.loc[df['Program__c_Name'].str.contains('Tutoring')
           & (df['Amount_of_Time__c'] > 120), '>120 Minutes'] = '>120 Minutes'

    df.loc[df['Intervention_Session_Date__c'] > df['CreatedDate'],
           'Logged in Future'] = 'Logged in Future'

    df.loc[df['Program__c_Name'].isin(['DESSA', 'Math Inventory', 'Reading Inventory']), 'Wrong Section'] = 'Wrong Section'

    error_cols = ['Missing T1/T2 Code', 'Listed T1 and T2', '<10 Minutes', '>120 Minutes', 'Logged in Future', 'Wrong Section']

    df['Error'] = df[error_cols].apply(lambda x: x.str.cat(sep=' & '), axis=1)

    accepted_errors_df = pd.read_excel((
        "Z:/ChiPrivate/Chicago Data and Evaluation/SY19/"
        "SY19 ToT Audit Accepted Errors.xlsx"
    ))

    df = df.loc[
        (df['Error'] != '') &
        ~df['Intervention_Session__c_Name'].isin(accepted_errors_df['SESSION_ID'])
    ]

    col_friendly_names = {
        'School_Name__c':'School',
        'Staff__c_Name':'ACM',
        'Program__c_Name':'Program',
        'Intervention_Session__c_Name':'Session ID',
        'Related_Student_s_Name__c':'Student',
        'CreatedDate':'Submission Date',
        'Intervention_Session_Date__c':'Session Date',
        'Amount_of_Time__c':'ToT',
        #'Comments__c':'Comment',
        'Error':'Error',
    }

    df = df.rename(columns=col_friendly_names)
    df = df.sort_values(list(col_friendly_names.values()))
    df = df[list(col_friendly_names.values())]

    return df

def write_error_tables_to_cyconnect(df):
    sch_ref_df = get_sch_ref_df()
    for index, row in sch_ref_df.iterrows():
        school_error_df = df.loc[df['School'] == row['School']].copy()
        del school_error_df['School']

        write_path = (f"Z:/{row['Informal Name']} Team Documents/SY19 ToT Audit "
                      f"Errors - {row['Informal Name']}.xlsx")

        if os.path.exists(write_path):
            os.remove(write_path)

        school_error_df.to_excel(write_path, index=False)

        # excel_writer = StyleFrame.ExcelWriter(write_path)
        # sfr = StyleFrame(school_error_df)
        # sfr.apply_column_style(
        #     cols_to_style=list(school_error_df),
        #     styler_obj=Styler(
        #         horizontal_alignment=utils.horizontal_alignments.left,
        #         #vertical_alignment=utils.vertical_alignments.top
        #     ),
        #     width=25,
        #     style_header=False,
        # )
        #
        # if len(school_error_df) > 0:
        #     freeze = 'A2'
        # else:
        #     freeze = 'A1'
        #
        # sfr.to_excel(
        #     excel_writer=excel_writer,
        #     row_to_add_filters=0,
        #     columns_and_rows_to_freeze=freeze,
        # )
        # excel_writer.save()
