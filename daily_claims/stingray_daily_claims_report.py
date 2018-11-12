from os.path import join
from xlwt import *
from datetime import date, timedelta, datetime, time
import time as t
import pyodbc
import util
from configparser import ConfigParser

config = ConfigParser()
config.read('./config.ini')

smtp_host = config.get('email', 'smtp_host')
smtp_user = config.get('email', 'smtp_user')
smtp_pwd = config.get('email', 'smtp_pwd')
from_address = config.get('email', 'from_address')
to_list = [] if config.get('email', 'to_list').split(',') == [''] else config.get('email', 'to_list').split(',')
cc_list = [] if config.get('email', 'cc_list').split(',') == [''] else config.get('email', 'cc_list').split(',')
dev_list = [] if config.get('email', 'dev_list').split(',') == [''] else config.get('email', 'dev_list').split(',')

output_folder = config.get('folders', 'output_folder')
log_folder = config.get('folders', 'log_folder')
log_file = join(log_folder, "daily_claims_" + date.today().strftime("%m%d%Y") + ".log")

def run_report(conn):

    f = open('./query.sql', 'r')
    query = [line.rstrip('\n') for line in f] #f.readlines()
    query = ''.join(query)
    f.close()

    cursor = conn.cursor()
    cursor.execute(query)
    result_set = cursor.fetchall()
    cursor.close()

    wb = Workbook()
    ws0 = wb.add_sheet('Sheet 1')

    style_row = easyxf("font: bold off, height 220, name Calibri")
    style_header = easyxf("font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; " \
                          "borders: top_color black, bottom_color black, right_color black, left_color black, " \
                          "left thin, right thin, top thin, bottom thin;")

    # Add headers to the excel workbook
    ws0.write(0, 0, 'Claim Number', style_header)
    ws0.write(0, 1, 'Date Reported', style_header)
    ws0.write(0, 2, 'Date of Loss', style_header)
    ws0.write(0, 3, 'Claim Status', style_header)
    ws0.write(0, 4, 'Date Reopened', style_header)
    ws0.write(0, 5, 'PA/Attorney assigned', style_header)
    ws0.write(0, 6, 'AOB assigned', style_header)
    ws0.write(0, 7, 'Adjuster assigned', style_header)
    ws0.write(0, 8, 'Cat Loss 1', style_header)
    ws0.write(0, 9, 'Cat Loss 2', style_header)
    ws0.write(0, 10, 'Last Close Date', style_header)
    ws0.write(0, 11, 'Cause of Loss', style_header)
    ws0.write(0, 12, 'Peril', style_header)
    ws0.write(0, 13, 'Form Type', style_header)
    ws0.write(0, 14, 'Total Indemnity Reserve Amount', style_header)
    ws0.write(0, 15, 'Total Indemnity Paid Amount', style_header)
    ws0.write(0, 16, 'Total Expense Reserve Amount', style_header)
    ws0.write(0, 17, 'Total Expense Paid Amount', style_header)
    ws0.write(0, 18, 'Insured Name', style_header)
    ws0.write(0, 19, 'Street Address', style_header)
    ws0.write(0, 20, 'City', style_header)
    ws0.write(0, 21, 'State', style_header)
    ws0.write(0, 22, 'ZIP', style_header)

    for i in range(21):
        ws0.col(i).width = 256 * 20

    for row_id, row in enumerate(result_set, start=1):
        for col_id, item in enumerate(row):
            ws0.write(row_id, col_id, str(item), style_row)

    report_name = 'Stingray Daily Claims report - ' + (date.today() - timedelta(1)).strftime("%m%d%Y") + '.xls'
    wb.save(join(output_folder, report_name))
    send_report_email(join(output_folder, report_name))

def send_report_email(file):
    subject_suffix = (date.today() - timedelta(1)).strftime("%m/%d/%Y")
    util.send_mail(smtp_host, smtp_user, smtp_pwd,
        from_address, to_list, cc_list,
        "Stingray Daily Claims report - " + subject_suffix,
        "Good morning,\n\n" \
            + "Please find attached daily open/closed/reopen claims report from Stingray for " \
            + subject_suffix + ".\n\nThank you, Stingray IT",
        [file]
    )

def send_email_dwh_not_current(file=None):
    util.send_mail(smtp_host, smtp_user, smtp_pwd,
        from_address, dev_list, [],
        "DWH is not current to send Daily Claims Report",
        "Good morning,\n\n" \
            + "Note: DWH is not current with latest data to send daily claims report. " \
            + "The process will check in 60 mins and send the report if the DB is refreshed\n\n" \
            + "Thank you, Stingray IT",
        [file]
    )

def is_dwh_current(conn):

    try:
        cursor = conn.cursor()
        query = "select top 1 cast(DTS_LastRunDate as datetime) from DTS_ScheduleTask order by 1 desc"
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        dwh_last_entry_date = result[0][0]

        if dwh_last_entry_date < datetime.combine(datetime.now().date() + timedelta(days=-1), time(13,00)):
            return False
    except Exception as e:
        return False

    return True

if __name__ == "__main__":

    holiday_list = ['2018-11-22', '2018-11-23']

    util.log(log_file, "Process started", False)
    connect_string = util.get_db_connection_string(config.get('database', 'db_host'), config.get('database', 'db_name'),
                                                   config.get('database', 'trusted'), config.get('database', 'db_user'),
                                                   config.get('database', 'db_pwd'))
    success = False
    while not success:
        try:
            conn = pyodbc.connect(connect_string)
            if not is_dwh_current(conn):
                raise Exception("DWH is not current")
            run_report(conn)
            util.log(log_file, "Report sent successfully", False)
            success = True
        except Exception as e:
            util.log(log_file, repr(e), False)
            send_email_dwh_not_current(log_file)
            t.sleep(int(config.get('report', 'retry')))
            success = False

    util.log(log_file, "Process ended", False)