import os
from pathlib import Path

import pandas as pd
import xlwings as xw

from .config import get_sch_ref_df
from . import simple_cysh as cysh


sch_ref_df = get_sch_ref_df()

def fill_one_coaching_log_acm_rollup(wb):
    sht = wb.sheets['ACM Rollup']
    sht.range('A:N').clear_contents()

    cols = ['Individual__c', 'ACM', 'Date', 'Role', 'Coaching Cycle', 'Subject',
            'Focus', 'Strategy/Skill', 'Notes', 'Action Steps', 'Completed?',
            'Followed Up']

    sht.range('A1').value = cols
    sht.range('A2:A3002').value = ("=INDEX('ACM Validation'!$A:$A, "
                                   "MATCH($B2, 'ACM Validation'!$C:$C, 0))")

    skip_sheets = ['Dev Tracker', 'Dev Map', 'ACM Validation', 'ACM Rollup',
                   'Calendar Validation', 'Log Validation']
    skip_sheets = set(skip_sheets + [f'Sheet{_}' for _ in range(1,9)])
    acm_sheets = sorted(list(set([x.name for x in wb.sheets]) - skip_sheets))

    row = 2
    for sheet_name in acm_sheets:
        sheet_name = sheet_name.replace("'", "''")
        sht.range(f'B{row}:B{row + 300}').value = f"='{sheet_name}'!$A$1"
        sht.range(f'C{row}:L{row + 300}').value = f"='{sheet_name}'!A3"
        row += 300

    return None


def fill_all_coaching_log_acm_rollup(sch_ref_df=sch_ref_df):
    app = xw.App()
    # app.display_alerts = False
    for index, row in sch_ref_df.iterrows():
        try:
            xlsx_path = (f"Z:/{row['Informal Name']} Leadership Team Documents/"
                         f"SY19 Coaching Log - {row['Informal Name']}.xlsx")
            wb = app.books.open(xlsx_path)
            fill_one_coaching_log_acm_rollup(wb)
            wb.save(xlsx_path)
            print(f"{row['Informal Name']} Coaching Log ACM Rollup updated")
        except (KeyboardInterrupt, SystemExit):
            raise
        except:
            print(f"{row['Informal Name']} failed to generate ACM Rollup sheet")
        finally:
            for _ in app.books:
                _.close()

    app.kill()

    return None


def prep_coaching_log(wb):
    wb.sheets['Dev Tracker'].range('A1,A5:L104').clear_contents()

    # Create ACM Copies
    sht = wb.sheets['ACM1']
    sht.range('A1,A3:J300').clear_contents()

    for x in range(2,11):
        sht.api.Copy(Before=wb.sheets['Dev Map'].api)
        wb.sheets['ACM1 (2)'].name = f'ACM{x}'

    return None


def deploy_choaching_logs(wb, staff_df, sch_ref_df=sch_ref_df):
    """Multiply by Schools
    """
    for index, row in sch_ref_df.iterrows():
        condition = ((staff_df['School']==row['School']) &
                     staff_df['Role__c'].str.contains('Corps Member'))
        school_staff = staff_df.loc[condition].copy()

        sht = wb.sheets['Dev Tracker']
        sht.range('A1,A5:L104').clear_contents()
        sht.range('A1').value = ("SY19 Coaching Log - "
                                 f"{row['Informal Name']}")

        sht = wb.sheets['ACM Validation']
        sht.range('A:N').clear()
        sht.range('A1').options(index=False, header=False).value = school_staff

        pos = 0
        sheets_added = []
        for staff_index, staff_row in school_staff.iterrows():
            pos += 1
            sht = wb.sheets[f'ACM{pos}']
            sht.name = staff_row['First_Name_Staff__c']
            sht.range('A1').value = staff_row['Name']
            sheets_added.append(sht.name)

        try:
            wb.save(f"Z:/{row['Informal Name']} Leadership Team Documents/"
                    f"SY19 Coaching Log - {row['Informal Name']}.xlsx")
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            print(f"Save failed {row['Informal Name']}: {e}")
            pass

        # Reset Sheets
        pos = 0
        for sheet_name in sheets_added:
            pos += 1
            wb.sheets[sheet_name].name = f'ACM{pos}'

    return None


def deploy_tracker(resource_type: str, containing_folder: str):
    """Distributes Excel tracker template to school team folders

    resource_type as 'SY19 Attendance Tracker' or 'SY19 Leadership Tracker'
    containing_folder as 'Team Documents' or 'Leadership Team Documents'
    """
    template_path = ('Z:/ChiPrivate/Chicago Data and Evaluation/SY19/'
                     f"Templates/{resource_type} Template.xlsx")
    wb = xw.Book(template_path)

    for index, row in sch_ref_df.iterrows():
        sht = wb.sheets['Tracker']
        sht.range('A1').clear_contents()
        sht.range('A1').value = f"{resource_type} - {row['Informal Name']}"

        try:
            wb.save(f"Z:/{row['Informal Name']} {containing_folder}/"
                    f"{resource_type} - {row['Informal Name']}.xlsx")
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            print(f"Save failed {row['Informal Name']}: {e}")
            continue

    wb.close()


def unprotect_sheets(resource_type: str, containing_folder: str,
                     sheet_name: str, sch_ref_df=sch_ref_df):
    """Unprotects a given sheet in a given excel resource across all schools"""
    import win32com.client

    xcl = win32com.client.Dispatch('Excel.Application')
    xcl.visible = True
    xcl.DisplayAlerts = False

    for index, row in sch_ref_df.iterrows():
        xlsx_path = (f"Z:/{row['Informal Name']} {containing_folder}/"
                     f"{resource_type} - {row['Informal Name']}.xlsx")
        wb = xcl.workbooks.open(xlsx_path)
        wb.Sheets(sheet_name).Unprotect()
        wb.Close(True, xlsx_path)

    xcl.Quit()

    return None


def update_acm_stdnt_validation_sheets(resource_type: str, containing_folder: str,
                                       sch_ref_df=sch_ref_df):
    """
    Updates the ACM and Student names referenced
    in all schools' Excel trackers of a given type
    """
    staff_df = cysh.get_staff_df()
    staff_df['First_Name_Staff__c'] = (
        staff_df['First_Name_Staff__c'] + " " +
        staff_df['Staff_Last_Name__c'].astype(str).str[0] + "."
    )
    staff_df.sort_values('First_Name_Staff__c', inplace=True)
    condition = (
        staff_df['Role__c'].str.contains('Corps Member|Team Leader') == True
    )
    staff_df = staff_df.loc[condition]

    if 'Attendance' in resource_type:
        att_df = cysh.get_student_section_staff_df(
            sections_of_interest='Coaching: Attendance'
        )
        att_df = att_df.sort_values('Student_Name__c')

    app = xw.App()
    # app.display_alerts = False

    for _, row in sch_ref_df.iterrows():
        xlsx_path = (f"Z:/{row['Informal Name']} {containing_folder}/"
                     f"{resource_type} - {row['Informal Name']}.xlsx")
        try:
            wb = app.books.open(xlsx_path)

            sheet_names = [x.name for x in wb.sheets]

            cols = ['Individual__c', 'First_Name_Staff__c', 'Staff__c_Name']
            school_staff = (staff_df.loc[staff_df['School'] == row['School'], cols]
                                    .copy())

            sht = wb.sheets['ACM Validation']
            sht.clear_contents()
            sht.range('A1').options(index=False, header=False).value = school_staff

            if 'Student Validation' in sheet_names and 'Attendance' in resource_type:
                school_att_df = att_df.loc[att_df['School__c']==row['School']].copy()
                sht = wb.sheets['Student Validation']
                sht.clear_contents()
                sht.range('A1').options(index=False, header=False).value = (
                    school_att_df[['Student_Name__c', 'Student__c']]
                )

            wb.save(xlsx_path)
        except (KeyboardInterrupt, SystemExit):
            raise
        except Exception as e:
            print(f"Failed to update {resource_type} sheets for "
                  f"{row['Informal Name']}: {e}")
        finally:
            wb.close()

    app.kill()

    return None


def update_coach_log_validation(sheet_name='Log Validation',
                                resource_type='SY19 Coaching Log',
                                containing_folder='Leadership Team Documents',
                                sch_ref_df=sch_ref_df):
    """
    Overwrites 'Log Validation' sheet of all schools' coaching logs
    """
    # {MetaName: {NameOfRange: [RangeItems] } }
    names = {
        'Focus': {
            "ELA": [
                "Comprehension",
                "Fluency",
                "Phonemic.Awareness.or.Phonics",
                "Student.Behavior.Management",
                "Vocabulary",
                "Writing",
                "Other",
            ],
            "Math":[
                "Adaptive.Reasoning",
                "Conceptual.Understanding",
                "Procedural.Fluency",
                "Strategic.Competence",
                "Student.Behavior.Management",
                "Other",
            ],
        },
    }

    app = xw.App()
    for index, row in sch_ref_df.iterrows():
        xlsx_path = (f"Z:/{row['Informal Name']} {containing_folder}/"
                     f"{resource_type} - {row['Informal Name']}.xlsx")

        wb = app.books.open(xlsx_path)

        for i, item in enumerate(names):
            col_chr = chr(65 + i)
            sht = wb.sheets[sheet_name]
            sht.range(f"${col_chr}$1:${col_chr}$300").clear_contents()
            xw.Range(f"'{sheet_name}'!${col_chr}$1").value = item.upper()

            row_n = 1
            for sub_item in names[item]:
                row_n += 2
                xl_range = f"'{sheet_name}'!${col_chr}${row_n}"
                xw.Range(xl_range).value = sub_item.upper()

                row_n += 1
                row_end = row_n + len(names[item][sub_item]) - 1
                xl_range = f"'{sheet_name}'!${col_chr}${row_n}:${col_chr}${row_end}"
                xw.Range(xl_range).name = sub_item
                xw.Range(xl_range).value = [[_] for _ in names[item][sub_item]]
                row_n = row_end

        wb.save(xlsx_path)

        for _ in app.books:
            _.close()

    for _ in xw.apps:
        _.kill()

    return True
