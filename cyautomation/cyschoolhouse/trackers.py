import logging
import traceback
from pathlib import Path

import pandas as pd
import xlwings as xw
from PyPDF2 import PdfFileMerger

from . import simple_cysh as cysh
from .config import TEMP_PATH, TEMPLATES_PATH, YEAR
from .utils import get_sch_ref_df


class Tracker:  # class used only for inheritance
    def __init__(self, kind, folder, filetype, test=False):
        self.kind = kind
        self.template_path = str(
            TEMPLATES_PATH / f"{YEAR} {self.kind} Template.xlsx"
        )
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
        self.sch_ref_df = self.sch_ref_df.set_index('Informal Name')


class ExcelTracker(Tracker):  # class used only for inheritance
    def __init__(self, kind, folder, filetype, test):
        super().__init__(kind=kind, folder=folder, filetype=filetype, 
                         test=test)

    def deploy_all(self):
        """ Distributes tracker for all school teams. Run only at the 
        start of the year.
        """
        resp = input(f'This will overwrite {self.kind}s. Are you sure? y/n: ')
        if resp.lower() != 'y':
            return None

        app = xw.App()
        for school_informal in self.sch_ref_df.index:
            wb = app.books.open(self.template_path)
            self.deploy_one(school_informal, wb, warn=False)

        app.kill()

    def deploy_one(self, school_informal, wb=None, warn=True):
        """ Distributes tracker for just one school team.

        school_informal: school as named in the 'Informal Name' column of the 
                         school reference dataframe
        wb: optional template workbook (loads automatically by default)
        """
        if warn:
            resp = input(f'This will overwrite {self.kind}s. '
                          'Are you sure? y/n: ')
            if resp.lower() != 'y':
                return None

        write_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
        logging.info(f"Deploying {write_path.stem}")
        
        if not wb:
            wb = xw.Book(self.template_path)

        sheet_names = [x.name for x in wb.sheets]

        title = write_path.stem
        sht = wb.sheets[0]
        sht.range('A1').clear_contents()
        sht.range('A1').value = title

        if 'ACM Validation' in sheet_names:
            self.update_one_acm_validation_sheet(school_informal, wb,
                                                 save_and_close=False)
        if 'Student Validation' in sheet_names:
            self.update_one_stdnt_validation_sheet(school_informal, wb,
                                                   save_and_close=False)

        if not write_path.parent.exists() and self.test:
            write_path.parent.mkdir(parents=True)

        wb_save_and_close(wb, write_path)

    def update_one_acm_validation_sheet(self, school_informal: str, wb=None, 
                                        save_and_close=True):
        if not wb:
            wb_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
            wb = xw.Book(wb_path)

        staff_df = self._get_staff_df(school_informal)

        sht = wb.sheets['ACM Validation']
        sht.clear_contents()
        sht.range('A1').options(index=False, header=False).value = staff_df

        if save_and_close:
            wb_save_and_close(wb, wb_path)

    def update_one_stdnt_validation_sheet(self, school_informal: str, wb=None, 
                                          save_and_close=True):
        if not wb:
            wb_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
            wb = xw.Book(wb_path)

        school_formal = self.sch_ref_df.loc[school_informal, 'School']
        
        if self.kind == 'Attendance Tracker':
            sections_of_interest = 'Coaching: Attendance'
        else:
            sections_of_interest = ''

        stdnt_df = cysh.get_student_section_staff_df(
            sections_of_interest=sections_of_interest,
            schools=[school_formal]
        )
        stdnt_df = stdnt_df.sort_values('Student_Name__c')

        sht = wb.sheets['Student Validation']
        sht.clear_contents()
        sht.range('A1').options(index=False, header=False).value = \
            stdnt_df[['Student_Name__c', 'Student__c']]

        if save_and_close:
            wb_save_and_close(wb, wb_path)

    def update_all_acm_stdnt_validation_sheets(self):
        """ Iterates through trackers and updates the ACM and Student names for
        dropdown validations.
        """
        app = xw.App()
        # app.display_alerts = False        
        for school_informal, row in self.sch_ref_df.iterrows():
            wb_path = row['tracker_path']
            logging.info(f'Updating {wb_path.stem}')

            wb = app.books.open(wb_path)

            sheet_names = [x.name for x in wb.sheets]

            if 'ACM Validation' in sheet_names:
                self.update_one_acm_validation_sheet(
                    school_informal, wb, save_and_close=False)

            if 'Student Validation' in sheet_names:
                self.update_one_stdnt_validation_sheet(
                    school_informal, wb, save_and_close=False)

            wb_save_and_close(wb, wb_path)

        app.kill()

    def _get_staff_df(self, school_informal):
        school_formal = self.sch_ref_df.loc[school_informal, 'School']

        roles = ['Corps Member', 'Second Year Corps Member', 
                 'Senior Corps Team Leader']
        staff_df = cysh.get_staff_df(schools=[school_formal], roles=roles)

        staff_df['First_Name_Staff__c'] = (
            staff_df['First_Name_Staff__c'] + " " +
            staff_df['Staff_Last_Name__c'].astype(str).str[0] + "."
        )

        staff_df.sort_values('First_Name_Staff__c', inplace=True)

        staff_df = staff_df[['Individual__c', 'First_Name_Staff__c', 
                             'Staff__c_Name']]

        return staff_df


class AttendanceTracker(ExcelTracker):
    def __init__(self, kind='Attendance Tracker', 
                 folder='Team Documents', filetype='.xlsx',
                 test=False):
        super().__init__(kind=kind, folder=folder, filetype=filetype, 
                         test=test)


class LeadershipTracker(ExcelTracker):
    def __init__(self, kind='Leadership Tracker', 
                 folder='Leadership Team Documents', filetype='.xlsx',
                 test=False):
        super().__init__(kind=kind, folder=folder, filetype=filetype, 
                         test=test)


class CoachingLog(ExcelTracker):
    def __init__(self, kind='Coaching Log',
                 folder='Leadership Team Documents', filetype='.xlsx',
                 test=False):
        super().__init__(kind=kind, folder=folder, filetype=filetype, 
                         test=test)

    def update_one_acm_validation_sheet(self, school_informal: str, wb=None, 
                                        save_and_close=True):
        if not wb:
            wb_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
            wb = xw.Book(wb_path)

        staff_df = self._get_staff_df(school_informal)

        sht = wb.sheets['ACM Validation']
        sht.clear_contents()
        sht.range('A1').options(index=False, header=False).value = staff_df

        sht = wb.sheets['ACM Template']
        sht.range('A1,A3:J300').clear_contents()
        sht.api.Visible = False

        sheet_names_lower = [x.name.lower() for x in wb.sheets]

        for i, r in staff_df.iterrows():
            if r['First_Name_Staff__c'].lower() not in sheet_names_lower:
                sht.api.Copy(Before=wb.sheets['Dev Map'].api)
                acm_sheet = wb.sheets['ACM Template (2)']
                acm_sheet.name = r['First_Name_Staff__c']
                acm_sheet.range('A1').value = r['Staff__c_Name']
                acm_sheet.api.Visible = True

        self._fill_acm_rollup_sheet(wb)

        if save_and_close:
            wb_save_and_close(wb, wb_path)

    @staticmethod
    def _fill_acm_rollup_sheet(wb):
        sht = wb.sheets['ACM Rollup']
        sht.range('A:N').clear_contents()

        # fill headers
        sht.range('A1').value = [
            'Individual__c', 'ACM', 'Date', 'Role', 'Coaching Cycle',
            'Subject', 'Focus', 'Strategy/Skill', 'Notes', 'Action Steps',
            'Completed?', 'Followed Up'
        ]
        # fill column A
        sht.range('A2:A3002').value = (
            "=INDEX('ACM Validation'!$A:$A, "
            "MATCH($B2, 'ACM Validation'!$C:$C, 0))"
        )
        # fill columns B:L
        skip_sheets = set(
            ['Dev Tracker', 'Dev Map', 'ACM Template', 'ACM Validation', 
             'ACM Rollup', 'Calendar Validation', 'Log Validation'] + 
            [f'Sheet{_}' for _ in range(1,9)]
        )
        acm_sheets = [x.name for x in wb.sheets if x.name not in skip_sheets]

        row = 2
        for sheet_name in acm_sheets:
            sheet_name = sheet_name.replace("'", "''")
            sht.range(f'B{row}:B{row + 300}').value = f"='{sheet_name}'!$A$1"
            sht.range(f'C{row}:L{row + 300}').value = f"='{sheet_name}'!A3"
            row += 300


class WeeklyServiceTracker(Tracker):
    def __init__(self, kind='Weekly Service Tracker',
                 folder='Team Documents', filetype='.pdf',
                 test=False):
        super().__init__(kind=kind, folder=folder, filetype=filetype, 
                         test=test)

    def deploy_all(self):
        """ Runs the entire Service Tracker publishing process
        """
        app = xw.App()
        for school_informal in self.sch_ref_df.index:
            wb = app.books.open(self.template_path)
            self.deploy_one(school_informal, wb)

        app.kill()

    def deploy_one(self, school_informal, wb=None):
        school_formal = self.sch_ref_df.loc[school_informal, 'School']
        write_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
        logging.info(f"Deploying {write_path.stem}")

        if not wb:
            wb = xw.Book(self.template_path)

        stu_sec_df = cysh.get_student_section_staff_df(
            sections_of_interest = ['Coaching: Attendance',
                                    'SEL Check In Check Out',
                                    'Tutoring: Literacy',
                                    'Tutoring: Math'], 
            schools=school_formal
        )
        stu_sec_df = self._process_section_enrollment_table(stu_sec_df)

        # Ensure `temp` folder is empty
        for path in Path(TEMP_PATH).iterdir():
            if path.is_dir():
                rm_tree(path)
            else:
                path.unlink()

        for acm_name, acm_df in stu_sec_df.groupby('Staff__c_Name'):
            try:
                self._fill_one_acm_wb(acm_df, acm_name, wb)
                pdf_path = f"{TEMP_PATH}/{acm_name}.pdf"
                wb.sheets['Service Tracker'].api.ExportAsFixedFormat(0, pdf_path)
            except (KeyboardInterrupt, SystemExit):
                raise
            except Exception as e:
                logging.error(f"Failed for {acm_name}: {e}")

        if not write_path.parent.exists() and self.test:
            write_path.parent.mkdir(parents=True)

        self._merge_and_save_one_school_pdf(school_informal, write_path)

        return None

    @staticmethod
    def _process_section_enrollment_table(df):
        # group by Student_Program__c, then sum ToT
        dosage_df = df.groupby('Student_Program__c')['Dosage_to_Date__c'].sum()
        df = df.join(dosage_df, how='left', on='Student_Program__c', 
                     rsuffix='_r')

        # filter out inactive students
        df = df.loc[
            (df['Active__c']==True) & 
            df['Enrollment_End_Date__c'].isnull()
        ]

        # clean program names and set zeros
        df['Program__c_Name'] = df['Program__c_Name'].replace({
            'Tutoring: Math': 'Math', 
            'Tutoring: Literacy': 'ELA'
        })
        df['Dosage_to_Date__c_r'] = (df['Dosage_to_Date__c_r'].fillna(value=0)
                                                              .astype(int))

        df['Dosage_to_Write'] = (df['Dosage_to_Date__c_r'].astype(str) + 
                                 "\r\n" + df['Program__c_Name'])

        df = df.sort_values(by=[
            'School_Reference_Id__c', 'Staff__c_Name',
            'Program__c_Name', 'Student_Grade__c',
            'Student_Name__c',
        ])

        df = df[['School_Reference_Id__c', 'Staff__c_Name', 'Program__c_Name', 
                 'Student_Name__c', 'Dosage_to_Write']]

        return df

    @staticmethod
    def _fill_one_acm_wb(acm_df, acm_name, wb):
        # Write header
        sht = wb.sheets['Header']
        sht.range('A1').options(index=False, header=False).value = acm_name

        # Write Course Performance
        df_acm_CP = acm_df.loc[
            acm_df['Program__c_Name'].isin(['Math', 'ELA'])
        ].copy()

        if len(df_acm_CP) > 12:
            logging.warning(f"More than 12 Math/ELA students for {acm_name}\n")

        sht = wb.sheets['Course Performance']
        sht.range('B4:C15').clear_contents()
        sht.range('B4').options(index=False, header=False).value = \
            df_acm_CP[['Student_Name__c', 'Dosage_to_Write']][0:12]

        # Write SEL
        df_acm_SEL = acm_df.loc[
            acm_df['Program__c_Name'].str.contains("SEL")
        ].copy()

        if len(df_acm_SEL) > 6:
            logging.warning(f"More than 6 SEL students for {acm_name}\n")

        sht = wb.sheets['SEL']
        sht.range('B5:B10').clear_contents()
        sht.range('B5').options(index=False, header=False).value = \
            df_acm_SEL['Student_Name__c'][0:6]

        # Write Attendance
        df_acm_attendance = acm_df.loc[
            acm_df['Program__c_Name'].str.contains("Attendance")
        ].copy()

        if len(df_acm_attendance['Student_Name__c']) > 6:
            logging.warning(f"More than 6 Attendance students for {acm_name}\n")

        sht = wb.sheets['Attendance CICO']
        sht.range('B4:B6, F4:F6').clear_contents()
        sht.range('B4').options(index=False, header=False).value = \
            df_acm_attendance['Student_Name__c'][0:3]
        sht.range('F4').options(index=False, header=False).value = \
            df_acm_attendance['Student_Name__c'][3:6]

        return None

    @staticmethod
    def _merge_and_save_one_school_pdf(school_informal_name, write_path):
        # Merge team PDFs
        merger = PdfFileMerger()
        for filepath in Path(TEMP_PATH).iterdir():
            if filepath.name.endswith('.pdf'):
                merger.append(str(filepath))

        merger.write(str(write_path))
        merger.close()

        return None


def wb_save_and_close(wb, write_path):
    wb.sheets[0].activate()  # sets focus on first sheet of document

    try:
        wb.save(write_path)
    except (KeyboardInterrupt, SystemExit):
        raise
    except Exception as e:
        logging.error(f"Failed to save {write_path.name}: {e}")
        traceback.print_exc()
    finally:
        wb.close()


def rm_tree(path: Path):
    for child in path.glob('*'):
        if child.is_file():
            child.unlink()
        else:
            rm_tree(child)
    path.rmdir()
