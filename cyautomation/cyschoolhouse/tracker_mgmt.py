import os
from pathlib import Path
import traceback

import pandas as pd
import xlwings as xw

from .config import get_sch_ref_df, YEAR, TEMP_PATH
from . import simple_cysh as cysh


TRACKER_DIRS = {
    'Service Tracker': 'Team Documents',
}

# TODO:
# * intervention trackers will require more manual setup
# * test deploy 
# * coaching logs
# * ToT Audit trackers
# * Weekly service trackers
# * SY20 calendars

class ExcelTracker:  # class used only for inheritance
    def __init__(self, kind, folder, test=False):
        self.kind = kind
        self.folder = folder
        self.template_path = (
            f"Z:/ChiPrivate/Chicago Data and Evaluation/{YEAR}/Templates/"
            f"{YEAR} {self.kind} Template.xlsx"
        )
        self.sch_ref_df = get_sch_ref_df()

        if test:
            root_dir = TEMP_PATH
        else:
            root_dir = 'Z:/'

        self.sch_ref_df['tracker_path'] = self.sch_ref_df.apply(
            lambda df: (Path(root_dir) / 
                        f"{df['Informal Name']} {self.folder}" / 
                        f"{YEAR} {self.kind} - {df['Informal Name']}.xlsx"),
            axis=1
        )
        self.sch_ref_df = self.sch_ref_df.set_index('Informal Name')

    def deploy_one(self, school_informal, wb=None, close_after=True):
        """ Distributes tracker for just one school team.

        school_informal: school as named in the 'Informal Name' column of the 
                         school reference dataframe
        close_after: optional, closes the workbook after deploying
        wb: optional template workbook (loads automatically by default)
        """
        if not wb:
            wb = xw.Book(self.template_path)

        write_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
        if not write_path.parent.exists():
            write_path.parent.mkdir(parents=True)

        sheet_names = [x.name for x in wb.sheets]

        title = write_path.stem
        sht = wb.sheets['Tracker']
        sht.range('A1').clear_contents()
        sht.range('A1').value = title

        if 'ACM Validation' in sheet_names:
            self.update_one_acm_validation_sheet(school_informal, wb,
                                                 close_after=False)
        if 'Student Validation' in sheet_names:
            self.update_one_stdnt_validation_sheet(school_informal, wb,
                                                 close_after=False)

        try:
            wb.save(write_path)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            print(f"Save failed {school_informal}: {e}")
            pass
    
        if close_after:
            wb.close()

    def deploy_all(self):
        """ Distributes tracker for all school teams. Run only at the 
        start of the year.
        """
        app = xw.App()
        wb = app.books.open(self.template_path)

        for school_informal in self.sch_ref_df.index:
            self.deploy_one(school_informal, wb, close_after=False)

        wb.close()
        app.kill()

    def update_one_acm_validation_sheet(self, school_informal: str, wb=None, 
                                        close_after=True):
        if not wb:
            wb_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
            wb = wb = xw.Book(wb_path)

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

        sht = wb.sheets['ACM Validation']
        sht.clear_contents()
        sht.range('A1').options(index=False, header=False).value = staff_df

        if close_after:
            wb.close()

        return staff_df

    def update_one_stdnt_validation_sheet(self, school_informal: str, wb=None, 
                                          close_after=True):
        if not wb:
            wb_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
            wb = xw.Book(wb_path)

        school_formal = self.sch_ref_df.loc[school_informal, 'School']
        
        if self.kind == 'Attendance Tracker':
            sections_of_interest = 'Coaching: Attendance'
        else:
            sections_of_interest = ''

        stdnt_df = cysh.get_student_section_staff_df(  # TODO: accept schools
            schools=[school_formal], sections_of_interest=sections_of_interest
        )
        stdnt_df = stdnt_df.sort_values('Student_Name__c')

        sht = wb.sheets['Student Validation']
        sht.clear_contents()
        (sht.range('A1')
            .options(index=False, header=False)
            .value) = stdnt_df[['Student_Name__c', 'Student__c']]

        if close_after:
            wb.close()

    def update_all_acm_stdnt_validation_sheets(self):
        """
        Iterates through trackers and updates the ACM and Student names for 
        dropdown validations.
        """
        app = xw.App()
        # app.display_alerts = False        
        for school_informal, row in self.sch_ref_df.iterrows():
            xlsx_path = row['tracker_path']

            try:
                wb = app.books.open(xlsx_path)

                sheet_names = [x.name for x in wb.sheets]
                if 'ACM Validation' in sheet_names:
                    self.update_one_acm_validation_sheet(school_informal, wb,
                                                         close_after=False)
                if 'Student Validation' in sheet_names:
                    self.update_one_stdnt_validation_sheet(school_informal, wb,
                                                           close_after=False)

                wb.save(xlsx_path)
            except (KeyboardInterrupt, SystemExit):
                raise
            except Exception:
                print(f"Failed to process {xlsx_path.name}:")
                traceback.print_exc()
            finally:
                if len(app.books) > 1:
                    app.books[-1].close()

        app.kill()

        return None


class AttendanceTracker(ExcelTracker):
    def __init__(self, kind='Attendance Tracker', 
                 folder='Team Documents', 
                 test=False):
        super().__init__(kind=kind, folder=folder, test=test)


class LeadershipTracker(ExcelTracker):
    def __init__(self, kind='Leadership Tracker', 
                 folder='Team Leadership Documents', 
                 test=False):
        super().__init__(kind=kind, folder=folder, test=test)


# class InterventionTracker(ExcelTracker):
#     def __init__(self, kind='Intervention Tracker', 
#                  folder='Team Documents', 
#                  test=False):
#         super().__init__(kind=kind, folder=folder, test=test)


class CoachingLog(ExcelTracker):
    def __init__(self, kind='Coaching Log', folder='Team Leadership Documents', 
                 test=False):
        super().__init__(kind=kind, folder=folder, test=test)

    def deploy_one(self, school_informal, wb=None, close_after=True):
        if not wb:
            wb = xw.Book(self.template_path)

        write_path = self.sch_ref_df.loc[school_informal, 'tracker_path']
        if not write_path.parent.exists():
            write_path.parent.mkdir(parents=True)

        title = write_path.stem
        sht = wb.sheets['Dev Tracker']
        sht.range('A1,A5:L104').clear_contents()
        sht.range('A1').value = title

        staff_df = self.update_one_acm_validation_sheet(
            school_informal, wb, close_after=False)

        pos = 1
        acm_sheets = []
        for _, row in staff_df.iterrows():
            sht = wb.sheets[f'ACM{pos}']
            sht.name = row['First_Name_Staff__c']
            sht.range('A1').value = row['Staff__c_Name']
            acm_sheets.append(sht.name)
            pos += 1

        try:
            wb.save(write_path)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            print(f"Save failed {school_informal}: {e}")
            pass

        # Reset Sheets
        pos = 1
        for acm_sheet in acm_sheets:
            wb.sheets[acm_sheet].name = f'ACM{pos}'
            pos += 1

        if close_after:
            wb.close()

    def deploy_all(self):
        """ Distributes tracker for all school teams. Run only at the 
        start of the year.
        """
        app = xw.App()
        wb = app.books.open(self.template_path)
        self.duplicate_acm_sheets(wb)
        self.fill_acm_rollup_sheet(wb)

        for school_informal in self.sch_ref_df.index:
            self.deploy_one(school_informal, wb, close_after=False)

        wb.close()
        app.kill()

    @staticmethod
    def duplicate_acm_sheets(wb):
        wb.sheets['Dev Tracker'].range('A1,A5:L104').clear_contents()

        # Create ACM Copies
        sht = wb.sheets['ACM1']
        sht.range('A1,A3:J300').clear_contents()

        for x in range(2,11):
            sht.api.Copy(Before=wb.sheets['Dev Map'].api)
            wb.sheets['ACM1 (2)'].name = f'ACM{x}'

        return None

    @staticmethod
    def fill_acm_rollup_sheet(wb):
        sht = wb.sheets['ACM Rollup']
        sht.range('A:N').clear_contents()

        cols = ['Individual__c', 'ACM', 'Date', 'Role', 'Coaching Cycle', 
                'Subject', 'Focus', 'Strategy/Skill', 'Notes', 'Action Steps', 
                'Completed?', 'Followed Up']
        # fill headers
        sht.range('A1').value = cols
        # fill column A
        sht.range('A2:A3002').value = (
            "=INDEX('ACM Validation'!$A:$A, "
            "MATCH($B2, 'ACM Validation'!$C:$C, 0))"
        )
        # fill columns B:L
        skip_sheets = set(
            ['Dev Tracker', 'Dev Map', 'ACM Validation', 'ACM Rollup', 
             'Calendar Validation', 'Log Validation'] + 
            [f'Sheet{_}' for _ in range(1,9)]
        )
        acm_sheets = [x.name for x in wb.sheets if x.name not in skip_sheets]

        row = 2
        for sheet_name in acm_sheets:
            sheet_name = sheet_name.replace("'", "''")
            sht.range(f'B{row}:B{row + 300}').value = f"='{sheet_name}'!$A$1"
            sht.range(f'C{row}:L{row + 300}').value = f"='{sheet_name}'!A3"
            row += 300

        return None


def unprotect_sheets(resource_type: str, containing_folder: str,
                     sheet_name: str):
    """Unprotects a given sheet in a given excel resource across all schools"""
    import win32com.client

    app = win32com.client.Dispatch('Excel.Application')
    app.visible = True
    app.DisplayAlerts = False
    
    sch_ref_df = get_sch_ref_df()
    for _, row in sch_ref_df.iterrows():
        xlsx_path = (f"Z:/{row['Informal Name']} {containing_folder}/"
                     f"{resource_type} - {row['Informal Name']}.xlsx")
        wb = app.workbooks.open(xlsx_path)
        wb.Sheets(sheet_name).Unprotect()
        wb.Close(True, xlsx_path)

    app.Quit()

    return None


# def update_coach_log_validation(sheet_name='Log Validation',
#                                 resource_type='SY19 Coaching Log',
#                                 containing_folder='Leadership Team Documents'):
#     """
#     Overwrites 'Log Validation' sheet of all schools' coaching logs
#
#     This is an unused function. In SY19, it was necessary to modify
#     a named range after deployment. Similar code could be useful in the
#     future.
#     """
#     # {MetaName: {NameOfRange: [RangeItems] } }
#     names = {
#         'Focus': {
#             "ELA": [
#                 "Comprehension",
#                 "Fluency",
#                 "Phonemic.Awareness.or.Phonics",
#                 "Student.Behavior.Management",
#                 "Vocabulary",
#                 "Writing",
#                 "Other",
#             ],
#             "Math":[
#                 "Adaptive.Reasoning",
#                 "Conceptual.Understanding",
#                 "Procedural.Fluency",
#                 "Strategic.Competence",
#                 "Student.Behavior.Management",
#                 "Other",
#             ],
#         },
#     }

#     app = xw.App()
#     for index, row in sch_ref_df.iterrows():
#         xlsx_path = (f"Z:/{row['Informal Name']} {containing_folder}/"
#                      f"{resource_type} - {row['Informal Name']}.xlsx")

#         wb = app.books.open(xlsx_path)

#         for i, item in enumerate(names):
#             col_chr = chr(65 + i)
#             sht = wb.sheets[sheet_name]
#             sht.range(f"${col_chr}$1:${col_chr}$300").clear_contents()
#             xw.Range(f"'{sheet_name}'!${col_chr}$1").value = item.upper()

#             row_n = 1
#             for sub_item in names[item]:
#                 row_n += 2
#                 xl_range = f"'{sheet_name}'!${col_chr}${row_n}"
#                 xw.Range(xl_range).value = sub_item.upper()

#                 row_n += 1
#                 row_end = row_n + len(names[item][sub_item]) - 1
#                 xl_range = f"'{sheet_name}'!${col_chr}${row_n}:${col_chr}${row_end}"
#                 xw.Range(xl_range).name = sub_item
#                 xw.Range(xl_range).value = [[_] for _ in names[item][sub_item]]
#                 row_n = row_end

#         wb.save(xlsx_path)

#         for _ in app.books:
#             _.close()

#     for _ in xw.apps:
#         _.kill()

#     return True
