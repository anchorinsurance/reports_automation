import pyodbc

import xlwt
from xlrd import open_workbook
from datetime import date, timedelta
import os
import glob
from datetime import date, timedelta
from xlutils.copy import copy  # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook  # http://pypi.python.org/pypi/xlrd
from os.path import join
import shutil

def run_report(inputdate):
    connection = pyodbc.connect(
        driver='{iSeries Access ODBC Driver}',
        system='192.168.168.51',
        uid='tbobbili',
        pwd='anchor2017')
    c1 = connection.cursor()

    q1 = """SELECT TRIM( MONTH(T06.PMISSUDTE) || '/' ||DAY(T06.PMISSUDTE) || '/' || YEAR(T06.PMISSUDTE) ), T01.NTCONR, CONCAT(T01.NTPRFX, T01.NTPLNR), T01.NTPLTY, \
                CASE SUBSTR(T02.AGTER#,1,3)       WHEN 'TXE' THEN 'Kyndra Wallace'       WHEN 'TXN' THEN 'Robyn Dallas'       WHEN 'TXS' THEN 'Victor Garcia'       WHEN 'TXW' THEN 'Chris Coker'       END, \
                CONCAT( CONCAT( CONCAT(T01.NTAGL3,' '), CONCAT(T01.NTAGL2,' ')), T01.NTAGL1), T02.AGMNAM, \
                CHAR(T06.PMTEFFDTE, USA), T05.NICITY, T05.NISTAT,  T01.NTCNTY,  T07.COUNTYNAME, T01.NTVALU, T03.PSTMTP, T04.HOYEAR, T04.HOPRC \
                FROM JHOWARD.NEWBUSREP T01 CROSS JOIN SISPRDD250.AGAGTN01 T02 CROSS JOIN SISPRDD250.ALPUPS05 T03 \
                CROSS JOIN SISPRDD250.ALRTHO01 T04 CROSS JOIN SISPRDD250.CINMAD22 T05 CROSS JOIN SISPRDD250.CIPOMF08 T06 LEFT JOIN SISPRDD250.TXLOCC T07 ON T07.POLLOCC =  T01.NTCNTY \
                WHERE NTAGL3 = T02.AGGEN# AND NTAGL2 = T02.AGSUB# AND NTAGL1 = T02.AGPRD# AND NTOREF = PSREFN \
                AND PSREFN = HOREFN AND NTCLID = NICLID AND NTCONR = COMP# AND NTPRFX = PMPRFX AND NTPLNR = PMPLNR """

    c1.execute(q1)
    result_set2 = c1.fetchall()

    archive = "C:/Reports/Daily/SIS/NB Production report/archive/"
    file_path = "C:/Reports/Daily/SIS/NB Production report/"

    # Input file name
    os.chdir(file_path)
    inputfile_result = glob.glob('New Business Production 2018*.xls')
    print(inputfile_result)

    # Set output file variables
    b = inputdate - timedelta(1)
    newfilename = b.strftime("%m.%d.%Y")

    # Open workbook
    rb = open_workbook(file_path + inputfile_result[0], formatting_info=True)

    # read only copy to introspect the file
    r_sheet = rb.sheet_by_index(10)  # Need to change the sheet number for every month
    # a writable copy (I can't read values out of this, only write to it)
    wb = copy(rb)

    # the sheet to write to within the writable copy
    ws0 = wb.get_sheet(10)  # Need to change the sheet number for every month
    # ws0 = wb.add_sheet('Daily Activity')

    START_ROW = 97  # 0 based (subtract 1 from excel row number)
    col_age_november = 1
    col_summer1 = 2
    col_fall1 = 3

    style_string = "font: bold off, height 220, name Calibri"
    style = xlwt.easyxf(style_string)

    style_string1 = "font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; borders: top_color black, bottom_color black, right_color black, left_color black,\
                                      left thin, right thin, top thin, bottom thin;"
    style1 = xlwt.easyxf(style_string1)

    # Added 04/18/2018 ( To change the date format while writing data to m/d/yy )
    style_string2 = "font: bold off, height 220, name Calibri"
    style2 = xlwt.easyxf(style_string2)
    style2.num_format_str = 'm/d/yy'

    last_row = r_sheet.nrows
    print(last_row)

    row_number = last_row

    for row in result_set2:
        column_num = 0
        for item in row:  # i.e. for each field in that row
            if (column_num == 0 or column_num == 7):  # Date columns
                # print(item)
                month, day, year = item.split("/")
                expires_on = date(int(year), int(month), int(day))
                ws0.write(row_number, column_num, (expires_on),
                          style2)  # For two date columns write data after changing date format and style 2
            elif (column_num == 10):
                ws0.write(row_number, column_num, int(item), style)  # convert row 10 to integer before writing to excel
            else:
                ws0.write(row_number, column_num, (item), style)
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1

    # Save the excel file
    fname = 'New Business Production 2018 - ' + newfilename + '.xls'

    wb.save(join(file_path, file_path + fname))

    shutil.move(file_path+inputfile_result[0],archive + inputfile_result[0])

    # SMTPLIB Code to send email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "New Business report SIS - " + newfilename
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['sisdlynb@relyonanchor.com','tbobbili@relyonanchor.com', 'shimpy@relyonanchor.com', 'jhoward@relyonanchor.com']
    msg["To"] = "sisdlynb@relyonanchor.com"
    msg["CC"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>; Janie Howard <jhoward@relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText(
        "Good morning,\n\nPlease find attached New Business report from SIS for "+newfilename+".\n\nThanks.",
        'plain', 'utf-8'))
    # Add attachment
    att1 = MIMEBase('application', "octet-stream")
    att1.set_payload(open(file_path + fname, 'rb').read())
    encoders.encode_base64(att1)
    att1.add_header('Content-Disposition', 'attachment; filename= ' + fname + '')
    msg.attach(att1)
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("StingrayDev@relyonanchor.com", to_address, msg.as_string())
    s.quit()


if __name__ == "__main__":
    todaydate = date.today()

    weekno = todaydate.weekday()

    holiday_list = ['2018-11-22', '2018-11-23']

    if todaydate.strftime('%Y-%m-%d') in holiday_list or weekno >= 5:
        print('Report not run due to company holiday')
    else:
        print('Running script')
        run_report(todaydate)



"""

# Code to Send email from Outlook with File Attachments

import win32com.client
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "New Business report SIS - "+newfilename#time.strftime("%m/")+str((int(time.strftime("%d"))-1))+time.strftime("/%Y")
newMail.Body = "Good morning,\n\nPlease find attached New Business report from SIS for "+newfilename+".\n\nThanks,\nTeja"
#time.strftime("%m/%d/%Y")+
#newMail.To = "tbobbili@relyonanchor.com"
newMail.To = "Michael Farrell <mfarrell@relyonanchor.com>; Bruce Storie <bstorie@relyonanchor.com>; Victor Garcia <vgarcia@relyonanchor.com>; Lisa Coates <lcoates@relyonanchor.com>; Justin Afkhami <jafkhami@relyonanchor.com>; Nancy Rodriguez <nrodriguez@relyonanchor.com>; Michael Fletcher <mfletcher@relyonanchor.com>; Kyndra Wallace <kwallace@relyonanchor.com>; Robyn Dallas <rdallas@relyonanchor.com>; Christina Romero <cromero@relyonanchor.com>; Brian Travnicek <btravnicek@relyonanchor.com>; Yan Wang <ywang@relyonanchor.com>; Chris Coker <ccoker@relyonanchor.com>; Christopher Ferguson <cferguson@relyonanchor.com>; James 'Vinny' Flaherty III <jflaherty@relyonanchor.com>; Paul Guijarro <pguijarro@Relyonanchor.com>"
newMail.CC = "Kevin Pawlowski <kpawlowski@relyonanchor.com>; Janie Howard <jhoward@relyonanchor.com>; Rads Mydam <rmydam@relyonanchor.com>; Nitish Settipalli <nsettipalli@relyonanchor.com>"
#newMail.BCC = "tbobbili@relyonanchor.com"
attachment1 = ("C:/Users/tbobbili/Desktop/Task/New Business Report - SIS/New Business Production 2018 - "+newfilename+".xls")
#attachment2 = ("C:/Users/tbobbili/Desktop/DailyActivity - 12042017.xls")
#attachment2 = "Path to attachment no. 2"
newMail.Attachments.Add(attachment1)
#newMail.Attachments.Add(attachment2)
newMail.display()
#newMail.Send()
print ("New Business report email sent successfully!!")
"""