import logging
from pathlib import Path

import pandas as pd

from . import simple_cysh as cysh
from .config import TEMP_PATH, YEAR
from .utils import get_sch_ref_df


class ToTAudit:
    def __init__(self, kind='ToT Audit Errors', folder='Team Documents',
                 filetype='.xlsx', test=False):
        self.kind = kind
        self.sch_ref_df = get_sch_ref_df()
        self.test = test

        if self.test:
            root_dir = TEMP_PATH
        else:
            root_dir = 'Z:/'

        self.sch_ref_df['tracker_path'] = self.sch_ref_df.apply(
            lambda df: (
                Path(root_dir) / f"{df['Informal Name']} {folder}" /
                f"{YEAR} {self.kind} - {df['Informal Name']}{filetype}"
            ),
            axis=1
        )
        self.sch_ref_df = self.sch_ref_df.set_index('School')

    def deploy_all(self):
        """ Fixes typos and distributes audit for all school teams.

        Returns a summary dataframe.
        """
        self.fix_T1T2ELT_typos()
        errors_df = self.get_errors_df()

        for school in self.sch_ref_df.index:
            write_path = self.sch_ref_df.loc[school, 'tracker_path']

            if not write_path.parent.exists() and self.test:
                write_path.parent.mkdir(parents=True)

            (errors_df.loc[errors_df['School'] == school]
                      .drop(columns='School')
                      .to_excel(write_path, index=False))

        # write aggregate table
        counts = (errors_df.groupby(['School', 'Error'])['ACM']
                           .count()
                           .reset_index(level='Error')
                           .rename(columns={'ACM':'Count'}))
        counts.to_excel(f'Z:/Impact Analytics Team/{YEAR} '
                        'ToT Audit Error Counts.xlsx')
        return counts

    def fix_T1T2ELT_typos(self):
        df = self.get_T1T2ELT_typo_fixes_df()

        if df.empty:
            return None

        logging.info(f"Fixing {len(df)} T1, T2, or ELT typos")

        for _, row in df.iterrows():
            result = cysh.sf.Intervention_Session__c.update(
                row.Intervention_Session__c,
                {'Comments__c':row['Comments__c_fixed']}
            )

            if not isinstance(result, int) or result != 204:
                logging.warning(
                    f'T1, T2, ELT fix failed for '
                    f'{row.Intervention_Session__c}: {result}'
                )

    @staticmethod
    def get_T1T2ELT_typo_fixes_df():
        """ Standardize common spellings of "T1" "T2" and "ELT"
        """
        typo_map = {
            'T1': r'[Tt](?:[Ii][Ee]|[Ee][Ii])[Rr] ?(?:1|[Oo][Nn][Ee])|t1',
            'T2': r'[Tt](?:[Ii][Ee]|[Ee][Ii])[Rr] ?(?:2|[Tt][Ww][Oo])|t2',
        #     'ELT': r'([Aa]fter ?[Ss]chool|ASP)',
        }

        all_typos = '|'.join(list(typo_map.values()))

        df = cysh.get_object_df(
            'Intervention_Session__c',
            ['Id', 'Comments__c'],
            rename_id=True
        )
        df = df.loc[df['Comments__c'].str.contains(all_typos)==True]

        df['Comments__c_fixed'] = df['Comments__c']
        for k, v in typo_map.items():
            df.loc[:, 'Comments__c_fixed'] = \
                df['Comments__c_fixed'].replace(regex=v, value=k)

        return df

    @staticmethod
    def get_errors_df():
        ISR_df = cysh.get_object_df(
            'Intervention_Session_Result__c',
            ['Amount_of_Time__c', 'IsDeleted', 'Intervention_Session_Date__c',
             'Related_Student_s_Name__c', 'Intervention_Session__c',
             'CreatedDate']
        )
        IS_df = cysh.get_object_df(
            'Intervention_Session__c',
            ['Id', 'Name', 'Comments__c', 'Section__c'],
            rename_id=True, rename_name=True
        )
        section_df = cysh.get_object_df(
            'Section__c',
            ['Id', 'School__c', 'Intervention_Primary_Staff__c', 'Program__c'],
            rename_id=True
        )
        school_df = cysh.get_object_df('Account', ['Id', 'Name'])
        school_df = school_df.rename(columns={
            'Id': 'School__c',
            'Name': 'School_Name__c'}
        )
        staff_df = cysh.get_object_df(
            'Staff__c',
            ['Id', 'Name'],
            where="Site__c = 'Chicago'",
            rename_id=True, rename_name=True
        )
        program_df = cysh.get_object_df(
            'Program__c',
            ['Id', 'Name'],
            rename_id=True, rename_name=True
        )

        df = (ISR_df.merge(IS_df, how='left', on='Intervention_Session__c')
                    .drop(columns=['Intervention_Session__c'])
                    .merge(section_df, how='left', on='Section__c')
                    .drop(columns=['Section__c'])
                    .merge(school_df, how='left', on='School__c')
                    .drop(columns=['School__c'])
                    .merge(staff_df, how='left',
                        left_on='Intervention_Primary_Staff__c',
                        right_on='Staff__c')
                    .drop(columns=['Intervention_Primary_Staff__c',
                                   'Staff__c'])
                    .merge(program_df, how='left', on='Program__c')
                    .drop(columns=['Program__c']))

        df['Intervention_Session_Date__c'] = \
            pd.to_datetime(df['Intervention_Session_Date__c']).dt.date
        df['CreatedDate'] = pd.to_datetime(df['CreatedDate']).dt.date
        df['Comments__c'] = df['Comments__c'].fillna('')

        error_masks = {
            'Missing T1/T2 Code': (
                df['Program__c_Name'].str.contains('Tutoring')
                & ~df['Comments__c'].str.contains('T1|T2')
            ),
            'Listed T1 and T2': (
                df['Program__c_Name'].str.contains('Tutoring')
                & df['Comments__c'].str.contains('T1')
                & df['Comments__c'].str.contains('T2')
            ),
            '<10 Minutes': (
                df['Program__c_Name'].isin([
                    'SEL Check In Check Out',
                    'Coaching: Attendance',
                    'Tutoring: Math',
                    'Tutoring: Literacy'
                ])
                & (df['Amount_of_Time__c'] < 10)
            ),
            '>120 Minutes': (
                df['Program__c_Name'].str.contains('Tutoring')
                & (df['Amount_of_Time__c'] > 120)
            ),
            'Logged in Future': (
                df['Intervention_Session_Date__c'] > df['CreatedDate']
            ),
            'Wrong Section': (
                df['Program__c_Name'].isin(['DESSA', 'Math Inventory',
                                            'Reading Inventory'])
            ),
        }

        for error_name, mask in error_masks.items():
            df.loc[mask, error_name] = error_name

        error_cols = list(error_masks.keys())

        df['Error'] = \
            df[error_cols].apply(lambda x: x.str.cat(sep=' & '), axis=1)

        accepted_errors_df = pd.read_excel((
            f"Z:/ChiPrivate/Chicago Data and Evaluation/{YEAR}/"
            f"{YEAR} ToT Audit Accepted Errors.xlsx"
        ))

        df = df.loc[(df['Error'] != '') &
                    ~df['Intervention_Session__c_Name'].isin(
                        accepted_errors_df['SESSION_ID'])]

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

        df = (df.rename(columns=col_friendly_names)
                .sort_values(list(col_friendly_names.values())))

        return df[list(col_friendly_names.values())]
