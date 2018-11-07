import pyodbc
import xlwt
from os.path import join
from xlwt import *
from datetime import date, timedelta

def run_report(inputdate):

    # Connection to the IBM i-series (AS400)
    connection = pyodbc.connect(
        driver='{iSeries Access ODBC Driver}',
        system='192.168.168.51',
        uid='tbobbili',
        pwd='anchor2017')
    c1 = connection.cursor()

    # Query for the NB Inspection report
    q1 = """SELECT  CONCAT( T01.NTPRFX, DIGITS(T01.NTPLNR)) AS POLICY, T03.PSTTID, T06.PMTEFFDTE, T05.NINAML, T05.NINMFR,  \
            CONCAT( CONCAT( '(' , CONCAT( DIGITS(T05.NIPHA1),') ')) , CONCAT( DIGITS(T05.NIPHP1), CONCAT('-',DIGITS(T05.NIPHN1))  )) AS INST, \
            CONCAT(  NIBLNR , NISTNM )  as InsAddress,  T05.NICITY, T05.NISTAT, CONCAT( SUBSTR(T05.NIZIPC,1,5), CONCAT( '-', SUBSTR(T05.NIZIPC,6,4)) ) , T04.HOYEAR, T01.NTVALU, T02.AGMNAM, CONCAT(T01.NTAGL3, CONCAT(T01.NTAGL2, T01.NTAGL1)) , \
            CONCAT( CONCAT('(', CONCAT( SUBSTR(T02.AGMTEL,1,3),') ')) , CONCAT( SUBSTR(T02.AGMTEL,4,3), CONCAT( '-', SUBSTR(T02.AGMTEL,7,4))  ) ), T02.AGEMAL, T05.NILOTY ,T06.PMISSUDTE, T01.NTPLTY, T02.AGTER#, T03.PSTMTP, T04.HOPRC, T01.NTUNIT , '' as Test\
            FROM JHOWARD.NEWBUSINP T01 CROSS JOIN SISPRDD250.AGAGTN01 T02 CROSS JOIN SISPRDD250.ALPUPS05 T03 CROSS JOIN \
            SISPRDD250.ALRTHO01 T04 CROSS JOIN SISPRDD250.CINMAD22 T05 CROSS JOIN SISPRDD250.CIPOMF08 T06  \
            WHERE NTAGL3 = T02.AGGEN# AND NTAGL2 = T02.AGSUB# AND NTAGL1 = T02.AGPRD# \
            AND NTOREF = PSREFN AND PSREFN = HOREFN AND NTCLID = NICLID AND NTCONR = COMP#  \
            AND NTPRFX = PMPRFX AND NTPLNR = PMPLNR AND NTUNIT = HOUNIT AND HOUNIT = NIPLLC """
    # Execute the query
    c1.execute(q1)
    # Save the results in result_set2 variable
    result_set2 = c1.fetchall()

    # Create a Excel Workbook
    wb = Workbook()
    # Add a sheet SIS NB Inspection
    ws0 = wb.add_sheet('SIS NB Inspection')
    # Set Style_string for headers & rows
    style_string = "font: bold off, height 220, name Calibri"
    style = xlwt.easyxf(style_string)
    style_string1 = "font: bold on, height 220, name Calibri; pattern: pattern solid, fore_colour gray25; borders: top_color black, bottom_color black, right_color black, left_color black,\
                                  left thin, right thin, top thin, bottom thin;"
    style1 = xlwt.easyxf(style_string1)
    # Add headers to the excel workbook
    ws0.write(0, 0, 'Policy Number', style1)
    ws0.write(0, 1, 'Inspection Type', style1)
    ws0.write(0, 2, 'Policy Effective Date', style1)
    ws0.write(0, 3, 'Insured Last Name', style1)
    ws0.write(0, 4, 'Insured First Name', style1)
    ws0.write(0, 5, 'Home Phone', style1)
    ws0.write(0, 6, 'Property Street Address', style1)
    ws0.write(0, 7, 'Property City', style1)
    ws0.write(0, 8, 'Property State', style1)
    ws0.write(0, 9, 'Property ZIP Code', style1)
    ws0.write(0, 10, 'Year Built', style1)
    ws0.write(0, 11, 'Coverage A', style1)
    ws0.write(0, 12, 'Agency Name', style1)
    ws0.write(0, 13, 'Agency ID', style1)
    ws0.write(0, 14, 'Agent Phone', style1)
    ws0.write(0, 15, 'Agent Email', style1)
    ws0.write(0, 16, 'Underwriter', style1)
    ws0.write(0, 17, 'Policy Issued Date', style1)
    ws0.write(0, 18, 'PolicyType', style1)
    ws0.write(0, 19, 'AGTER#', style1)
    ws0.write(0, 20, 'Premium', style1)
    ws0.write(0, 21, 'ProtectionClass', style1)
    ws0.write(0, 22, 'Unit#', style1)

    row_number = 1
    # Paste the results in the work book
    for row in result_set2:
        column_num = 0
        for item in row:  # i.e. for each field in that row
            ws0.write(row_number, column_num, str(item),
                      style)  # ,wb.get_sheet(0).cell(0,0).style)  #write excel cell from the cursor at row 1
            column_num = column_num + 1  # increment the column to get the next field

        row_number = row_number + 1
    # Set column width for the sheet
    for i in range(24):
        if i == 1:
            ws0.col(i).width = 256 * 18
        else:
            ws0.col(i).width = 256 * 12
        i = i + 1


    #Path where the reports are stored
    test_files = "C:/Reports/Daily/SIS/NB Inspection report/"
    #Set filenames and subject name for email
    a = inputdate - timedelta(1)
    filename = a.strftime("%m%d%Y")
    subjectname = a.strftime("%m/%d/%Y")

    fname = 'NB Inspection SIS-' + filename + '.xls'
    # Save the excel file
    wb.save(join(test_files, fname))

    # SMTPLIB Code to send email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "NB Inspection report SIS - " + subjectname
    msg["From"] = "StingrayDev@relyonanchor.com"
    to_address = ['sisdlynbinspn@relyonanchor.com','tbobbili@relyonanchor.com', 'shimpy@relyonanchor.com', 'jhoward@relyonanchor.com']
    msg["To"] = "sisdlynbinspn@relyonanchor.com"
    msg["CC"] = "tbobbili@relyonanchor.com; Shimpy <shimpy@relyonanchor.com>; Janie Howard <jhoward@relyonanchor.com>"
    # Add Body
    msg.attach(MIMEText("Good Morning,\n\nPlease find attached NB Inspection report from SIS for " + subjectname + ".\n\nThanks.",'plain', 'utf-8'))
    # Add attachment

    att1 = MIMEBase('application', "octet-stream")
    att1.set_payload(open(test_files + fname, 'rb').read())
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

