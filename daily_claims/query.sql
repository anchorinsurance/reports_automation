select distinct CD.CD_ClaimCode as 'Claim Number', cast(CD.CD_LossReportDate as date) as 'Date Reported'
            , cast(CD.CD_DOL as date) as 'Date of Loss'
            , CS.CS_Status as 'Claims Status'
            , REPLACE(ISNULL(cast(CD.CD_ReOpenDate as date),''),'1900-01-01', '') as 'Date Reopened'
            , ISNULL(CV1.CV_Name,'') as 'PA / Attorney assigned'
            , ISNULL(CV2.CV_Name,'') as 'AOB assigned'
            , SU.sUSR_FirstName + ' ' + SU.sUSR_LastName as 'Adjuster assigned'
            , case when CC1.CA_ID IS NULL then '' else CC1.CA_Catastrophe end as 'Cat Loss 1'
            , case when CC2.CA_ID IS NULL then '' else CC2.CA_Catastrophe end as 'Cat Loss 2'
            , REPLACE(ISNULL(cast(CD_CloseDate as date),''),'1900-01-01', '') as 'Last Close Date'
            , CCL.CCL_Loss as 'Cause of Loss'
            , SPR.sPER_Peril as 'Peril'
            , SF.sFRM_FormType as 'Form Type'
            , ROUND(SUM(case when CTT_ResTrans = 1 and CTT_IndTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Indemnity Reserve Amount'
            , ROUND(SUM(case when CTT_ResTrans = 0 and CTT_IndTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Indemnity Paid Amount'
            , ROUND(SUM(case when CTT_ResTrans = 1 and CTT_ExpTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Expense Reserve Amount'
            , ROUND(SUM(case when CTT_ResTrans = 0 and CTT_ExpTrans = 1 then isnull((isnull(CT.CTR_DebitAmt, 0)) - (isnull(CT.CTR_CreditAmt,0)), 0) else 0 end),2) as 'Total Expense Paid Amount'
            , case when CD_IsClaimOnly = 1 then CO_FirstName + ' ' + CO_LastName else PC_FirstName + ' ' + PC_LastName end as InsuredName
            , case when CD_IsClaimOnly = 1 then CO_Address1 + ' ' + ISNULL(CO_Address2,'') else PDLI_Address1 + ' ' + ISNULL(PDLI_Address2,'') end as StreetAddress
            , case when CD_IsClaimOnly = 1 then CO_City else PDLI_City end as City
            , case when CD_IsClaimOnly = 1 then CO_State else PDLI_State end as State
            , case when CD_IsClaimOnly = 1 then CO_Zip else PDLI_Zip end as Zip
        from Claims_ClaimData CD
        left join Claims_Vendors CV1 ON CD.CV_ID_PAA = CV1.CV_ID
        left join Claims_Vendors CV2 ON CD.CV_ID_AOB = CV2.CV_ID
        left join System_Users SU ON CD.CD_ClaimAdjuster = SU.sUSR_ID
        left join Claims_Status CS ON CD.CD_StatusID = CS.CS_ID
        left join Claims_Catastrophes CC1 ON CD.CD_LossCat1 = CC1.CA_ID
        left join Claims_Catastrophes CC2 ON CD.CD_LossCat2 = CC2.CA_ID
        left join System_Perils SPR ON CD.CD_PerilID = SPR.sPER_ID
        left join System_Forms SF ON CD.sFRM_ID = SF.sFRM_ID
        left join Claims_CauseOfLoss CCL ON CD.CCL_ID = CCL.CCL_ID
        left JOIN Claims_ClaimsOnly AS CCO ON CCO.CO_ID = CD.PD_ID
        left JOIN PolicyData AS PD ON PD.PD_PolicyCode = CD.PD_PolicyCode and PD.PD_CurrentRecord=1
        left JOIN Policy_Contacts AS PC ON PD.PC_ID = PC.PC_ID
        left JOIN PolicyData_LocationInfo AS PDLI ON PDLI.PD_ID = PD.PD_ID
        left join Claims_Transactions CT ON CD.CD_ClaimCode = CT.CD_ClaimCode
        left join Claims_TransactionList as CTL on CTL.CTT_ID = CT.CTT_ID
        where CD_CurrentRecord = 1 and CD.CD_ClaimCode is not null
        group by CD.CD_ClaimCode	, cast(CD.CD_LossReportDate as date)	, cast(CD.CD_DOL as date)	, CS.CS_Status	, cast(CD.CD_ReOpenDate as date)
            , ISNULL(CV1.CV_Name,'')	, ISNULL(CV2.CV_Name,'')	, SU.sUSR_FirstName + ' ' + SU.sUSR_LastName	, CC1.CA_ID	, CC2.CA_ID
            , CC1.CA_Catastrophe	, CC2.CA_Catastrophe	, cast(CD_CloseDate as date)	, CCL.CCL_Loss	, SPR.sPER_Peril	, SF.sFRM_FormType
            , CD_IsClaimOnly	, CO_FirstName + ' ' + CO_LastName	, CO_Address1 + ' ' + ISNULL(CO_Address2,'')	, CO_City	, CO_State	, CO_Zip
            , PC_FirstName + ' ' + PC_LastName	, PDLI_Address1 + ' ' + ISNULL(PDLI_Address2,'')	, PDLI_City	, PDLI_State	, PDLI_Zip
        order by CD.CD_ClaimCode;