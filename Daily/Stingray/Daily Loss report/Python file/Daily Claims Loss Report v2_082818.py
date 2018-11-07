import pyodbc
from datetime import date, timedelta

def run_report(inputdate):
    cnxn = pyodbc.connect(
        'Driver={SQL Server Native Client 11.0};Server=dwh.relyonanchor.com;Database=AIH_Insurance;Trusted_Connection=yes;')
    cursor = cnxn.cursor()

    # Begin Query and execution
    q1 = "IF OBJECT_ID('tempdb..#Claims_ClaimData') IS NOT NULL    DROP TABLE #Claims_ClaimData; "
    q2 = "IF OBJECT_ID('tempdb..#MaxCD_ID') IS NOT NULL    DROP TABLE #MaxCD_ID; "
    q3 = "IF OBJECT_ID('tempdb..#tbl_Claim_List_detail_CAT_Column1') IS NOT NULL    DROP TABLE #tbl_Claim_List_detail_CAT_Column1; "
    q4 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Status_detail_CAT_Column1') IS NOT NULL    DROP TABLE #tbl_ClaimList_Status_detail_CAT_Column1; "
    q5 = "IF OBJECT_ID('tempdb..#tbl_ExpPaid_basic_Column1') IS NOT NULL    DROP TABLE #tbl_ExpPaid_basic_Column1; "
    q6 = "IF OBJECT_ID('tempdb..#tbl_IndemPaid_basic_Column1') IS NOT NULL    DROP TABLE #tbl_IndemReserve_basic_Column1; "
    q7 = "IF OBJECT_ID('tempdb..#tbl_LAEReserve_basic_Column1') IS NOT NULL    DROP TABLE #tbl_LAEReserve_basic_Column1; "
    q8 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Incurred_CAT_Column1') IS NOT NULL    DROP TABLE #tbl_ClaimList_Incurred_CAT_Column1; "
    q9 = "IF OBJECT_ID('tempdb..#ClaimLoss_Table') IS NOT NULL    DROP TABLE #ClaimLoss_Table; "
    q10 = "IF OBJECT_ID('tempdb..#tbl_Claim_List_detail_CAT_Column2') IS NOT NULL    DROP TABLE #tbl_Claim_List_detail_CAT_Column2; "
    q11 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Status_detail_CAT_Column2') IS NOT NULL    DROP TABLE #tbl_ClaimList_Status_detail_CAT_Column2; "
    q12 = "IF OBJECT_ID('tempdb..#tbl_ExpPaid_basic_Column2') IS NOT NULL    DROP TABLE #tbl_ExpPaid_basic_Column2; "
    q13 = "IF OBJECT_ID('tempdb..#tbl_IndemPaid_basic_Column2') IS NOT NULL    DROP TABLE #tbl_IndemPaid_basic_Column2; "
    q14 = "IF OBJECT_ID('tempdb..#tbl_IndemReserve_basic_Column2') IS NOT NULL    DROP TABLE #tbl_IndemReserve_basic_Column2; "
    q15 = "IF OBJECT_ID('tempdb..#tbl_LAEReserve_basic_Column2') IS NOT NULL    DROP TABLE #tbl_LAEReserve_basic_Column2; "
    q16 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Incurred_CAT_Column2') IS NOT NULL    DROP TABLE #tbl_ClaimList_Incurred_CAT_Column2; "
    q17 = "IF OBJECT_ID('tempdb..#CatDescriptions') IS NOT NULL    DROP TABLE #CatDescriptions; "
    q17_1 = "IF OBJECT_ID('tempdb..#tbl_Claim_List_detail_CAT_Column3') IS NOT NULL    DROP TABLE #tbl_Claim_List_detail_CAT_Column3; "
    q17_2 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Status_detail_CAT_Column3') IS NOT NULL    DROP TABLE #tbl_ClaimList_Status_detail_CAT_Column3; "
    q17_3 = "IF OBJECT_ID('tempdb..#tbl_ExpPaid_basic_Column3') IS NOT NULL    DROP TABLE #tbl_ExpPaid_basic_Column3; "
    q17_4 = "IF OBJECT_ID('tempdb..#tbl_IndemPaid_basic_Column3') IS NOT NULL    DROP TABLE #tbl_IndemPaid_basic_Column3; "
    q17_5 = "IF OBJECT_ID('tempdb..#tbl_IndemReserve_basic_Column3') IS NOT NULL    DROP TABLE #tbl_IndemReserve_basic_Column3; "
    q17_6 = "IF OBJECT_ID('tempdb..#tbl_LAEReserve_basic_Column3') IS NOT NULL    DROP TABLE #tbl_LAEReserve_basic_Column3; "
    q17_7 = "IF OBJECT_ID('tempdb..#tbl_ClaimList_Incurred_CAT_Column3') IS NOT NULL    DROP TABLE #tbl_ClaimList_Incurred_CAT_Column3; "
    q17_8 = "IF OBJECT_ID('tempdb..#ClaimLoss_Table_Input') IS NOT NULL    DROP TABLE #ClaimLoss_Table_Input; "

    cursor.execute(q1)
    cursor.execute(q2)
    cursor.execute(q3)
    cursor.execute(q4)
    cursor.execute(q5)
    cursor.execute(q6)
    cursor.execute(q7)
    cursor.execute(q8)
    cursor.execute(q9)
    cursor.execute(q10)
    cursor.execute(q11)
    cursor.execute(q12)
    cursor.execute(q13)
    cursor.execute(q14)
    cursor.execute(q15)
    cursor.execute(q16)
    cursor.execute(q17)
    cursor.execute(q17_1)
    cursor.execute(q17_2)
    cursor.execute(q17_3)
    cursor.execute(q17_4)
    cursor.execute(q17_5)
    cursor.execute(q17_6)
    cursor.execute(q17_7)
    cursor.execute(q17_8)

    q18 = "Declare @D1 date; Set @D1 = CONVERT(char (10),DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1),126); select @D1 as D1;"
    q19 = "Declare @D5 date; Set @D5 = CONVERT(char (10),DATEADD(DAY,0,GETDATE()),126); select @D5 as D5;"
    q20 = "Declare @D6 date; Set @D6 = CONVERT(char (10),DATEADD(DAY,-2, GETDATE()),126); select @D6 as D6;"

    cursor.execute(q18)
    result_set1 = cursor.fetchall()
    d1 = str(result_set1[0][0])
    print(d1)

    cursor.execute(q19)
    result_set2 = cursor.fetchall()
    d5 = str(result_set2[0][0])
    print(d5)

    cursor.execute(q20)
    result_set3 = cursor.fetchall()
    d6 = str(result_set3[0][0])
    print(d6)

    q21 = "Declare @D7 date; Set @D7 = CONVERT(char (10),DATEADD(month, DATEDIFF(month, 0, '%s'), 0),126); select @D7 as D7;" % (
        d5)

    cursor.execute(q21)
    result_set4 = cursor.fetchall()
    d7 = str(result_set4[0][0])
    print(d7)

    q22 = "Declare @D8 date; Set @D8 = CONVERT(char (10),DATEADD(DAY,-1,GETDATE()),126); select @D8 as D8;"

    cursor.execute(q22)
    result_set4_1 = cursor.fetchall()
    d8 = str(result_set4_1[0][0])
    print(d8)

    L1 = '1'
    L2 = '2'
    L3 = '3'

    q23 = "CREATE TABLE #CatDescriptions (StingrayCode int,CatLabel varchar(255) );"
    q24 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('-1','Non-Cat');"
    q25 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('3','Hermine');"
    q26 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('4','Matthew');"
    q27 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('5','Irma');"
    q28 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('6','Nate');"
    q29 = "Insert into #CatDescriptions (StingrayCode,CatLabel) values ('7','Alberto');"

    cursor.execute(q23)
    cursor.execute(q24)
    cursor.execute(q25)
    cursor.execute(q26)
    cursor.execute(q27)
    cursor.execute(q28)
    cursor.execute(q29)

    q30 = "SELECT MAX(CD_ID) AS CD_ID, CD_ClaimCode INTO #MaxCD_ID FROM   dbo.Claims_ClaimData A GROUP BY CD_ClaimCode;"
    q31 = "Select dbo.System_CompanyTree.* INTO #System_CompanyTree FROM dbo.System_CompanyTree;"
    q32 = "UPDATE #System_CompanyTree SET sCT_CODE = left(sCT_CODE,8);"
    # q33 = "SELECT MAX(CD_ID) AS CD_ID, CD_ClaimCode INTO #MaxCD_ID FROM   dbo.#Claims_ClaimData A GROUP BY CD_ClaimCode;"

    cursor.execute(q29)
    cursor.execute(q30)
    cursor.execute(q31)
    cursor.execute(q32)
    # cursor.execute(q33)

    q33 = ("""SELECT        A.CD_CloseDate AS CD_CloseDate, A.CD_ReOpenDate AS CD_ReOpenDate, A.CD_ClaimCode, A.sLOB_ID, A.CD_ID, A.PD_PolicyCode, \
                A.CD_DOL AS CD_DOL, A.CD_LossReportDate, A.PD_Agency, C.CCL_Loss AS CCL_Loss, A.CD_statusDate AS CD_statusDate, A.CD_Zip, A.CD_LossCat1,sCT_Code \
                INTO              #tbl_Claim_List_detail_CAT_Column1 \
                FROM          Claims_ClaimData A  INNER JOIN dbo.Claims_CauseOfLoss C ON A.CCL_ID = C.CCL_ID \
                INNER JOIN #MaxCD_ID D on D.CD_ID = A.CD_ID left join  #system_companytree E on E.sCT_ID= A.PD_Agency \
                WHERE (A.CD_LossReportDate <  '%s' ); """ % (d5))
    q34 = """SELECT     B.CD_ClaimCode,  B.sCT_CODE, B.sLOB_ID, B.CD_ID,B.PD_PolicyCode,B.CD_DOL,B.CD_LossReportDate,B.PD_Agency,B.CCL_Loss, B.CD_statusDate, C.CS_Status, (isnull(A.CD_CloseDate,0)) as CD_CloseDate, (isnull(A.CD_ReOpenDate,0)) as CD_ReOpenDate, A.CD_LossDesc, A.CD_LossCat1  \
             INTO #tbl_ClaimList_Status_detail_CAT_Column1 \
             FROM dbo.Claims_ClaimData A INNER JOIN dbo.Claims_Status C ON A.CD_StatusID = C.CS_ID INNER JOIN #tbl_Claim_List_detail_CAT_Column1 B ON B.CD_ID = A.CD_ID ;"""
    q35 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) - SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0)) AS [LAE PAID] \
                INTO              #tbl_ExpPaid_basic_Column1 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'AEP') and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d5))
    q36 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM PAID] \
                INTO                #tbl_IndemPaid_basic_Column1 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'LP')  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d5))
    q37 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) as decimal)- cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM RSRV] \
                INTO              #tbl_IndemReserve_basic_Column1 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        ((dbo.Claims_TransactionList.CTT_PP_Code = 'LR') or (dbo.Claims_TransactionList.CTT_PP_Code = 'RES'))   and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d5))
    q38 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast(SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [LAE RSRV] \
                INTO               #tbl_LAEReserve_basic_Column1 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE       ((dbo.Claims_TransactionList.CTT_ID > '4') and (dbo.Claims_TransactionList.CTT_ID < '9'))  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d5))

    q39 = ("""SELECT        A.CD_ClaimCode, A.sCT_COde,  (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) as 'TOTAL_LOSS', \
                case when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) >'50000' then '50000' + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0))  when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) <= '50000' THEN (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) END   as 'CAPPED_LOSS', A.CD_LossCat1, A.CD_DOL, A.CCL_Loss \
                INTO              #tbl_ClaimList_Incurred_CAT_Column1 \
                FROM            #tbl_ClaimList_Status_detail_CAT_Column1 A LEFT OUTER JOIN #tbl_ExpPaid_basic_Column1 B ON A.CD_ClaimCode = B.CD_ClaimCode LEFT OUTER JOIN #tbl_IndemPaid_basic_Column1 C ON A.CD_ClaimCode = C.CD_ClaimCode \
                LEFT OUTER JOIN #tbl_IndemReserve_basic_Column1 D ON A.CD_ClaimCode = D.CD_ClaimCode LEFT OUTER JOIN #tbl_LAEReserve_basic_Column1 E ON A.CD_ClaimCode = E.CD_ClaimCode; """)

    cursor.execute(q33)
    cursor.execute(q34)
    cursor.execute(q35)
    cursor.execute(q36)
    cursor.execute(q37)
    cursor.execute(q38)
    cursor.execute(q39)

    q40 = ("""SELECT        A.CD_CloseDate AS CD_CloseDate, A.CD_ReOpenDate AS CD_ReOpenDate, A.CD_ClaimCode, A.sLOB_ID, A.CD_ID, A.PD_PolicyCode, \
                A.CD_DOL AS CD_DOL, A.CD_LossReportDate, A.PD_Agency, C.CCL_Loss AS CCL_Loss, A.CD_statusDate AS CD_statusDate, A.CD_Zip, A.CD_LossCat1,sCT_Code \
                INTO              #tbl_Claim_List_detail_CAT_Column2 \
                FROM          Claims_ClaimData A  INNER JOIN dbo.Claims_CauseOfLoss C ON A.CCL_ID = C.CCL_ID INNER JOIN  #MaxCD_ID D on D.CD_ID = A.CD_ID left join  #system_companytree E on E.sCT_ID= A.PD_Agency \
                WHERE        (A.CD_LossReportDate <  '%s' ); """ % (d8))

    q41 = """SELECT     B.CD_ClaimCode,  B.sCT_CODE, B.sLOB_ID, B.CD_ID,B.PD_PolicyCode,B.CD_DOL,B.CD_LossReportDate,B.PD_Agency,B.CCL_Loss, B.CD_statusDate, C.CS_Status, (isnull(A.CD_CloseDate,0)) as CD_CloseDate, (isnull(A.CD_ReOpenDate,0)) as CD_ReOpenDate, A.CD_LossDesc, A.CD_LossCat1 \
             INTO                #tbl_ClaimList_Status_detail_CAT_Column2 \
             FROM            dbo.Claims_ClaimData A INNER JOIN dbo.Claims_Status C ON A.CD_StatusID = C.CS_ID INNER JOIN #tbl_Claim_List_detail_CAT_Column2 B ON  B.CD_ID = A.CD_ID ;"""

    q42 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) - SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0)) AS [LAE PAID] \
                INTO              #tbl_ExpPaid_basic_Column2 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'AEP') and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d8))

    q43 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM PAID] \
                INTO                #tbl_IndemPaid_basic_Column2 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'LP')  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d8))

    q44 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) as decimal)- cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM RSRV] \
                INTO              #tbl_IndemReserve_basic_Column2 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        ((dbo.Claims_TransactionList.CTT_PP_Code = 'LR') or (dbo.Claims_TransactionList.CTT_PP_Code = 'RES'))   and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d8))

    q45 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast(SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [LAE RSRV] \
                INTO               #tbl_LAEReserve_basic_Column2 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE       ((dbo.Claims_TransactionList.CTT_ID > '4') and (dbo.Claims_TransactionList.CTT_ID < '9'))  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d8))

    q46 = ("""SELECT        A.CD_ClaimCode, A.sCT_COde,  (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) as 'TOTAL_LOSS', \
                case when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) >'50000' then '50000' + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0))  when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) <= '50000' THEN (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) END   as 'CAPPED_LOSS', A.CD_LossCat1, A.CD_DOL, A.CCL_Loss \
                INTO              #tbl_ClaimList_Incurred_CAT_Column2 \
                FROM            #tbl_ClaimList_Status_detail_CAT_Column2 A LEFT OUTER JOIN #tbl_ExpPaid_basic_Column2 B ON A.CD_ClaimCode = B.CD_ClaimCode LEFT OUTER JOIN #tbl_IndemPaid_basic_Column2 C ON A.CD_ClaimCode = C.CD_ClaimCode  \
                LEFT OUTER JOIN #tbl_IndemReserve_basic_Column2 D ON A.CD_ClaimCode = D.CD_ClaimCode LEFT OUTER JOIN #tbl_LAEReserve_basic_Column2 E ON A.CD_ClaimCode = E.CD_ClaimCode; """)

    # ---------------------------------------------------------------

    q47 = ("""SELECT        A.CD_CloseDate AS CD_CloseDate, A.CD_ReOpenDate AS CD_ReOpenDate, A.CD_ClaimCode, A.sLOB_ID, A.CD_ID, A.PD_PolicyCode, \
                A.CD_DOL AS CD_DOL, A.CD_LossReportDate, A.PD_Agency, C.CCL_Loss AS CCL_Loss, A.CD_statusDate AS CD_statusDate, A.CD_Zip, A.CD_LossCat1,sCT_Code \
                INTO              #tbl_Claim_List_detail_CAT_Column3 \
                FROM          Claims_ClaimData A  INNER JOIN dbo.Claims_CauseOfLoss C ON A.CCL_ID = C.CCL_ID INNER JOIN  #MaxCD_ID D on D.CD_ID = A.CD_ID left join  #system_companytree E on E.sCT_ID= A.PD_Agency \
                WHERE        (A.CD_LossReportDate <  '%s' ); """ % (d7))

    q48 = """SELECT     B.CD_ClaimCode,  B.sCT_CODE, B.sLOB_ID, B.CD_ID,B.PD_PolicyCode,B.CD_DOL,B.CD_LossReportDate,B.PD_Agency,B.CCL_Loss, B.CD_statusDate, C.CS_Status, (isnull(A.CD_CloseDate,0)) as CD_CloseDate, (isnull(A.CD_ReOpenDate,0)) as CD_ReOpenDate, A.CD_LossDesc, A.CD_LossCat1 \
             INTO                #tbl_ClaimList_Status_detail_CAT_Column3 \
             FROM            dbo.Claims_ClaimData A INNER JOIN dbo.Claims_Status C ON A.CD_StatusID = C.CS_ID INNER JOIN #tbl_Claim_List_detail_CAT_Column3 B ON  B.CD_ID = A.CD_ID ;"""

    q49 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) - SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0)) AS [LAE PAID] \
                INTO              #tbl_ExpPaid_basic_Column3 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'AEP') and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d7))

    q50 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM PAID] \
                INTO                #tbl_IndemPaid_basic_Column3 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        (dbo.Claims_TransactionList.CTT_PP_Code = 'LP')  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d7))

    q51 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0)) as decimal)- cast (SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [INDEM RSRV] \
                INTO              #tbl_IndemReserve_basic_Column3 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE        ((dbo.Claims_TransactionList.CTT_PP_Code = 'LR') or (dbo.Claims_TransactionList.CTT_PP_Code = 'RES'))   and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d7))

    q52 = ("""SELECT        dbo.Claims_Transactions.CD_ClaimCode, cast(SUM(isnull(dbo.Claims_Transactions.CTR_DebitAmt,0))as decimal) - cast(SUM(isnull(dbo.Claims_Transactions.CTR_CreditAmt,0))as decimal) AS [LAE RSRV] \
                INTO               #tbl_LAEReserve_basic_Column3 \
                FROM            dbo.Claims_Transactions INNER JOIN dbo.Claims_TransactionList ON dbo.Claims_Transactions.CTT_ID = dbo.Claims_TransactionList.CTT_ID \
                WHERE       ((dbo.Claims_TransactionList.CTT_ID > '4') and (dbo.Claims_TransactionList.CTT_ID < '9'))  and (dbo.Claims_Transactions.CTR_EntryDate <  '%s' ) \
                GROUP BY dbo.Claims_Transactions.CD_ClaimCode; """ % (d7))

    q53 = ("""SELECT        A.CD_ClaimCode, A.sCT_COde,  (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) as 'TOTAL_LOSS', \
                case when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) >'50000' then '50000' + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0))  when (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) <= '50000' THEN (isnull(C.[INDEM PAID],0)) + (isnull(D.[INDEM RSRV],0)) + (isnull(B.[LAE PAID],0)) + (isnull(E.[LAE RSRV],0)) END   as 'CAPPED_LOSS', A.CD_LossCat1, A.CD_DOL, A.CCL_Loss \
                INTO              #tbl_ClaimList_Incurred_CAT_Column3 \
                FROM            #tbl_ClaimList_Status_detail_CAT_Column3 A LEFT OUTER JOIN #tbl_ExpPaid_basic_Column3 B ON A.CD_ClaimCode = B.CD_ClaimCode LEFT OUTER JOIN #tbl_IndemPaid_basic_Column3 C ON A.CD_ClaimCode = C.CD_ClaimCode  \
                LEFT OUTER JOIN #tbl_IndemReserve_basic_Column3 D ON A.CD_ClaimCode = D.CD_ClaimCode LEFT OUTER JOIN #tbl_LAEReserve_basic_Column3 E ON A.CD_ClaimCode = E.CD_ClaimCode; """)

    cursor.execute(q40)
    cursor.execute(q41)
    cursor.execute(q42)
    cursor.execute(q43)
    cursor.execute(q44)
    cursor.execute(q45)
    cursor.execute(q46)
    cursor.execute(q47)
    cursor.execute(q48)
    cursor.execute(q49)
    cursor.execute(q50)
    cursor.execute(q51)
    cursor.execute(q52)
    cursor.execute(q53)

    q54 = ("""select A.sCT_Code, ((ISNULL(A.TOTAL_LOSS,0)- ISNULL(B.TOTAL_LOSS,0))) as TOTAL_LOSS, ((ISNULL(A.CAPPED_LOSS,0) -ISNULL(B.CAPPED_LOSS,0))) as CAPPED_LOSS,A.CD_LossCat1, '%s' as Report_Column, A.CD_ClaimCode,A.CD_DOL, A.CCL_Loss \
                INTO #ClaimLoss_Table \
                from #tbl_ClaimList_Incurred_CAT_Column1 A LEFT JOIN #tbl_ClaimList_Incurred_CAT_Column2 B on A.CD_ClaimCode =B.CD_ClaimCode; """ % (
        L1))
    q55 = ("""INSERT INTO #ClaimLoss_Table ( sCT_Code,TOTAL_LOSS,CAPPED_LOSS, CD_LossCat1,Report_Column, CD_ClaimCode , CD_DOL, CCL_Loss) \
                select A.sCT_Code, ((ISNULL(A.TOTAL_LOSS,0)- ISNULL(B.TOTAL_LOSS,0))) , ((ISNULL(A.CAPPED_LOSS,0) -ISNULL(B.CAPPED_LOSS,0))) ,A.CD_LossCat1, '%s' as Report_Column, A.CD_ClaimCode, A.CD_DOL, A.CCL_Loss \
                from #tbl_ClaimList_Incurred_CAT_Column1 A LEFT JOIN #tbl_ClaimList_Incurred_CAT_Column3 B on A.CD_ClaimCode =B.CD_ClaimCode; """ % (
        L2))
    q56 = ("""INSERT INTO #ClaimLoss_Table ( sCT_Code,TOTAL_LOSS,CAPPED_LOSS, CD_LossCat1,Report_Column, CD_ClaimCode, CD_DOL, CCL_Loss) \
                select A.sCT_Code, ((ISNULL(A.TOTAL_LOSS,0))) , ((ISNULL(A.CAPPED_LOSS,0))) ,A.CD_LossCat1, '%s' as Report_Column, A.CD_ClaimCode, A.CD_DOL, A.CCL_Loss \
                from #tbl_ClaimList_Incurred_CAT_Column1 A; """ % (L3))

    cursor.execute(q54)
    cursor.execute(q55)
    cursor.execute(q56)

    q57 = ("""select '%s' as MonthBeg, '%s' as DataAsOf,CD_CLAIMCODE, round (TOTAL_LOSS,2) as 'Incurred_Loss',round(CAPPED_LOSS,2) as 'Capped_Loss', Report_Column, \
                case when CD_LossCat1 ='-1' then 'Non-CAT' when CD_LossCat1 ='3' then 'Hermine' when CD_LossCat1 ='4' then 'Matthew' when CD_LossCat1 ='5' then 'Irma' when CD_LossCat1 ='6' then 'Nate' when CD_LossCat1 ='7' then 'Alberto' when CD_LossCat1 ='8' then 'HailEvent' else 'Unknown' END as CD_LossCat1 ,  CONVERT(char (10),CD_DOL,126) as CD_DOL , CCL_Loss  \
                    FROM #ClaimLoss_Table ; """ % (d7, d8))

    cursor.execute(q57)
    col1 = [desc[0] for desc in cursor.description]
    final_result_set1 = cursor.fetchall()

    q58 = ("""select '%s' as MonthBeg, '%s' as DataAsOf,CD_CLAIMCODE, round (TOTAL_LOSS,2) as 'Incurred_Loss',round(CAPPED_LOSS,2) as 'Capped_Loss', Report_Column, \
                case when CD_LossCat1 ='-1' then 'Non-CAT' when CD_LossCat1 ='3' then 'Hermine' when CD_LossCat1 ='4' then 'Matthew' when CD_LossCat1 ='5' then 'Irma' when CD_LossCat1 ='6' then 'Nate' when CD_LossCat1 ='7' then 'Alberto' when CD_LossCat1 ='8' then 'HailEvent' else 'Unknown' END as CD_LossCat1 ,  CONVERT(char (10),CD_DOL,126) as CD_DOL , CCL_Loss  \
                 INTO #ClaimLoss_Table_Input FROM #ClaimLoss_Table ; """ % (d7, d8))

    q59 = ("""ALTER TABLE #ClaimLoss_Table_Input ADD FLAG varchar(5) ; """)

    q60 = (
        """UPDATE #ClaimLoss_Table_Input SET FLAG = '1' FROM #ClaimLoss_Table_Input A inner join #tbl_ClaimList_Incurred_CAT_Column3 B on A.CD_CLAIMCODE =B.CD_CLAIMCODE ; """)

    q61 = ("""UPDATE #ClaimLoss_Table_Input SET FLAG = '0' WHERE FLAG is null ; """)

    cursor.execute(q58)
    cursor.execute(q59)
    cursor.execute(q60)
    cursor.execute(q61)

    q62 = ("""select CD_ClaimCode,round (Incurred_Loss,2) as 'Incurred_Loss', CD_LossCat1, CD_DOL,CCL_Loss \
                    FROM #ClaimLoss_Table_Input where ([Incurred_Loss] > '1' or [Incurred_Loss] < '-1') and (CD_ClaimCode like '%V%' and CD_ClaimCode like '%FL%' and Report_Column ='1') and FLAG ='0' ; """)

    cursor.execute(q62)
    col2 = [desc[0] for desc in cursor.description]
    final_result_set2 = cursor.fetchall()

    q63 = ("""select CD_ClaimCode,round (Incurred_Loss,2) as 'Incurred_Loss', CD_LossCat1, CD_DOL,CCL_Loss \
                    FROM #ClaimLoss_Table_Input where ([Incurred_Loss] > '1' or [Incurred_Loss] < '-1') and (CD_ClaimCode like '%T%' and CD_ClaimCode like '%FL%' and Report_Column ='1') and FLAG ='0' ; """)

    cursor.execute(q63)
    col3 = [desc[0] for desc in cursor.description]
    final_result_set3 = cursor.fetchall()

    q64 = ("""select CD_ClaimCode,round (Incurred_Loss,2) as 'Incurred_Loss', CD_LossCat1, CD_DOL,CCL_Loss \
                    FROM #ClaimLoss_Table_Input where ([Incurred_Loss] > '1' or [Incurred_Loss] < '-1') and (CD_ClaimCode like '%V%' and CD_ClaimCode like '%LA%' and Report_Column ='1') and FLAG ='0' ; """)

    cursor.execute(q64)
    col4 = [desc[0] for desc in cursor.description]
    final_result_set4 = cursor.fetchall()
    # print (col)
    # print (result_set5)
    cursor.close()
    cnxn.close()

    import pandas as pd
    df1 = pd.DataFrame.from_records((final_result_set1))  # , columns=col)
    df1.columns = col1

    df2 = pd.DataFrame.from_records((final_result_set2))  # , columns=col)
    # df2.columns = col2

    df3 = pd.DataFrame.from_records((final_result_set3))  # , columns=col)
    # df3.columns = col3

    df4 = pd.DataFrame.from_records((final_result_set4))  # , columns=col)
    # df4.columns = col4
    # print (df.head())

    from datetime import date, time, timedelta
    todaydate = date.today()
    newfilename = todaydate.strftime("%m.%d.%Y")

    rep_path = "C:/Reports/Daily/Stingray/Daily Loss report/"
    fname = 'Daily loss report - ' + newfilename + '.xlsx'

    from shutil import copyfile
    copyfile(rep_path + 'Template.xlsx', rep_path + fname)

    import xlwings as xw
    wb = xw.Book(rep_path + fname)
    app = xw.apps.active
    sheet1 = wb.sheets['Input']

    sheet1.range('A1').value = col1
    sheet1.range('A2').value = df1.values

    # wb.sheets['Input'].Visible = False
    # wb.sheets['Input'].hidden= True

    sheet2 = wb.sheets['1.FLVOL']

    sheet2.range('A1').value = col2
    sheet2.range('A2').value = df2.values

    sheet3 = wb.sheets['2.FLTO']

    sheet3.range('A1').value = col3
    sheet3.range('A2').value = df3.values

    sheet4 = wb.sheets['3.LA']

    sheet4.range('A1').value = col4
    sheet4.range('A2').value = df4.values

    wb.save()
    wb.close()
    app.kill()

    # SMTPLIB Code to send email
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email import encoders
    # Define from and to address
    msg = MIMEMultipart()
    msg["Subject"] = "Daily loss report - " + newfilename
    msg["From"] = "tbobbili@relyonanchor.com"
    msg["To"] = "tbobbili@relyonanchor.com"
    # Add Body
    msg.attach(MIMEText(
        "Good morning,\n\nPlease find attached daily loss report from Stingray for " + newfilename + ".\n\nThanks.",
        'plain', 'utf-8'))
    # Add attachment
    att1 = MIMEBase('application', "octet-stream")
    att1.set_payload(open(rep_path + fname, 'rb').read())
    encoders.encode_base64(att1)
    att1.add_header('Content-Disposition', 'attachment; filename= ' + fname + '')
    msg.attach(att1)
    # Send Email
    s = smtplib.SMTP("relyonanchor-com.mail.protection.outlook.com")
    s.sendmail("tbobbili@relyonanchor.com", "tbobbili@relyonanchor.com", msg.as_string())
    s.quit()



"""

# Send email from Outlook
import win32com.client
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Daily loss report - "+newfilename
newMail.Body = "Good morning,\n\nPlease find attached daily loss report from Stingray for "+newfilename+".\n\nThanks,\nTeja"
#time.strftime("%m/%d/%Y")+
#newMail.To = "tbobbili@relyonanchor.com"
newMail.To = "Michael Fletcher <mfletcher@Relyonanchor.com>; John Rohloff <jrohloff@Relyonanchor.com>; Tyler Andrejack <tandrejack@Relyonanchor.com>; Nia Patel <npatel@Relyonanchor.com>; Howard Cosner <hcosner@relyonanchor.com>; William Mayo <wmayo@relyonanchor.com>; Tracy Clark <tclark@relyonanchor.com>; Anthony Satira <asatira@Relyonanchor.com>"
newMail.CC = "Michael Farrell <MFarrell@Relyonanchor.com>; Brian Travnicek <btravnicek@Relyonanchor.com>; Rads Mydam <rmydam@relyonanchor.com>; Paul Guijarro <pguijarro@Relyonanchor.com>; Michael Terry <MTerry@Relyonanchor.com>; Justin Afkhami <jafkhami@Relyonanchor.com>; Kevin Pawlowski <kpawlowski@Relyonanchor.com>"
#newMail.BCC = "tbobbili@relyonanchor.com"
attachment1 = ("C:/Users/tbobbili/Desktop/Claims Report/Daily loss report - "+newfilename+".xlsx")
newMail.Attachments.Add(attachment1)
newMail.display()
"""

todaydate = date.today()

weekno = todaydate.weekday()

holiday_list = ['2018-08-01', '2018-08-11','2018-09-03', '2018-11-20']

if todaydate.strftime('%Y-%m-%d') in holiday_list or weekno >= 5:
    print ('Report not run due to company holiday')
else:
    print ('Running script')
    run_report(todaydate)
