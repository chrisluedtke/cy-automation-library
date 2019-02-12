from pathlib import Path
from time import sleep

import pandas as pd
from PyPDF2 import PdfFileMerger
import xlwings as xw

from .config import get_sch_ref_df, LOG_PATH, TEMP_PATH, TEMPLATES_PATH
from . import simple_cysh as cysh


sch_ref_df = get_sch_ref_df()


def get_section_enrollment_table(sections_of_interest):
    """
    """
    df = cysh.get_student_section_staff_df(sections_of_interest)

    # group by Student_Program__c, then sum ToT
    df_dosage_sum = df.groupby('Student_Program__c')['Dosage_to_Date__c'].sum()
    df = df.join(df_dosage_sum, how='left', on='Student_Program__c', rsuffix='_r')

    # filter out inactive students
    df = df.loc[(df['Active__c']==True) & df['Enrollment_End_Date__c'].isnull()]

    # clean program names and set zeros
    df['Program__c_Name'] = df['Program__c_Name'].replace({
        'Tutoring: Math':'Math', 'Tutoring: Literacy':'ELA'
    })
    df['Dosage_to_Date__c_r'] = df['Dosage_to_Date__c_r'].fillna(value=0).astype(int)

    df['Dosage_to_Write'] = df['Dosage_to_Date__c_r'].astype(str) + "\r\n" + df['Program__c_Name']

    df.sort_values(by=[
        'School_Reference_Id__c', 'Staff__c_Name',
        'Program__c_Name', 'Student_Grade__c',
        'Student_Name__c',
    ], inplace = True)

    df = df[['School_Reference_Id__c', 'Staff__c_Name', 'Program__c_Name', 'Student_Name__c', 'Dosage_to_Write']]

    return df


def fill_one_acm_wb(acm_df, acm_name, wb, logf):
    # Write header
    sht = wb.sheets['Header']
    sht.range('A1').options(index=False, header=False).value = acm_name

    # Write Course Performance
    df_acm_CP = acm_df.loc[acm_df['Program__c_Name'].isin(['Math', 'ELA'])].copy()
    if len(df_acm_CP) > 12:
        logf.write(f"Warning: More than 12 Math/ELA students for {acm_name}\n")

    sht = wb.sheets['Course Performance']
    sht.range('B4:C15').clear_contents()
    sht.range('B4').options(index=False, header=False).value = df_acm_CP[
        ['Student_Name__c', 'Dosage_to_Write']
        ][0:12]

    # Write SEL
    df_acm_SEL = acm_df.loc[acm_df['Program__c_Name'].str.contains("SEL")].copy()
    if len(df_acm_SEL) > 6:
        logf.write(f"Warning: More than 6 SEL students for {acm_name}\n")

    sht = wb.sheets['SEL']
    sht.range('B5:B10').clear_contents()
    sht.range('B5').options(index=False, header=False).value = df_acm_SEL[
        'Student_Name__c'
        ][0:6]

    # Write Attendance
    df_acm_attendance = acm_df.loc[acm_df['Program__c_Name'].str.contains("Attendance")].copy()
    if len(df_acm_attendance['Student_Name__c']) > 6:
        logf.write(f"Warning: More than 6 Attendance students for {acm_name}\n")

    sht = wb.sheets['Attendance CICO']
    sht.range('B4:B6, F4:F6').clear_contents()
    sht.range('B4').options(index=False, header=False).value = df_acm_attendance['Student_Name__c'][0:3]
    sht.range('F4').options(index=False, header=False).value = df_acm_attendance['Student_Name__c'][3:6]

    return None


def merge_and_save_one_school_pdf(school_informal_name):
    # Merge team PDFs
    merger = PdfFileMerger()
    for filepath in Path(TEMP_PATH).iterdir():
        if '.pdf' in str(filepath):
            merger.append(str(filepath))

    # Edit this write path to match your Sharepoint file structure
    merger.write(f"Z:\\{school_informal_name} Team Documents\\SY19 Weekly Service Trackers - {school_informal_name}.pdf")
    merger.close()

    return None


def update_service_trackers():
    """ Runs the entire Service Tracker publishing process
    """
    logf = open(f"{LOG_PATH}/Service Tracker Log.log", "w")

    student_section_df = get_section_enrollment_table(
        sections_of_interest= [
            'Coaching: Attendance',
            'SEL Check In Check Out',
            'Tutoring: Literacy',
            'Tutoring: Math',
        ]
    )

    # Open Excel Template and define sheet references
    xlsx_path = f"{TEMPLATES_PATH}/Service Tracker Template.xlsx"

    # Iterate through school names to build Service Tracker PDFs
    for school in student_section_df['School_Reference_Id__c'].unique():
        wb = xw.Book(xlsx_path)

        # Ensure `temp` folder is empty
        for filepath in Path(TEMP_PATH).iterdir():
            filepath.unlink()

        logf.write(f"Writing Service Tracker: {school}")
        print(f"Writing Service Tracker: {school}\n")

        df_school = student_section_df.loc[
            student_section_df['School_Reference_Id__c'] == school
            ].copy()

        for acm_name in df_school['Staff__c_Name'].unique():
            acm_df = df_school.loc[df_school['Staff__c_Name']==acm_name].copy()

            pdf_path = f"{TEMP_PATH}/{acm_name}.pdf"

            try:
                fill_one_acm_wb(acm_df, acm_name, wb, logf)
                wb.sheets['Service Tracker'].api.ExportAsFixedFormat(0, pdf_path)
            except Exception as e:
                text = ('Error filling template or saving pdf for '
                        f"{acm_name}: {e}")
                logf.write(text)
                print(text)

        xw.apps.active.kill()

        school_informal_name = sch_ref_df.loc[
            sch_ref_df['School'] == school, 'Informal Name'
            ].values[0]

        merge_and_save_one_school_pdf(school_informal_name)

    logf.write("Completed script 'Weekly Service Tracker Update'\n")

    return