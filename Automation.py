# Martha Hause
# This documentation incorporates multiple different sources being joined together to get one central area for the projects expense savings to be represented.
# The projects in this file is Docuflash, COI, Self-service, Loss Run, and Session Time Reduced.
import pyodbc as p
import pandas as pd

# python "u:/Github Training/trialfile.py"        <- run this to get the code to work. for some reason my code doesn't run unless you add this.

# Bringing in Loss Run
print("Loss Run Information")

LR_connect = p.connect("DSN=MSS-P1-PCA-06")

LR_Table = """
Select top 100
      FORMAT(EmailSentDate, 'MM-yy') As 'Date'
     ,Sum(Case When EmailSubject = 'Progressive Agent Contact Us' Then 1 Else 0 End) As 'Auto - Form'
     ,Sum(Case When EmailSubject Like 'PGR_FAO_LossRunRequest%' Then 1 Else 0 End)   As 'Auto - Widget'
     ,Sum( Case
             When Not(
                     EmailSubject = 'Progressive Agent Contact Us'
                     Or EmailSubject Like 'PGR_FAO_LossRunRequest%'
                     ) Then 1 Else 0 End)  As 'Auto - Email'
     ,Count(*)  As 'Auto - Total'
From  PCAStage.dbo.vwEmailClassifierProcessed
Where
      EmailSentDate > '2022-01-01' AND ProcessStatus = '1'
Group By FORMAT(EmailSentDate, 'MM-yy')
"""

LR_df = pd.read_sql(LR_Table, LR_connect)
# print(LR_df)

LR_connect.close()

# Bringing in Docuflash
print("Docuflash Information")


# Bringing in COI
print("COI Information")

COI_connect = p.connect("DSN=MSS-P1-PCA-06")

COI_Table = """
SELECT count(distinct(INFOID)) as dis_INFOID, FORMAT(ClickDateTime, 'MM-yy') as date
FROM FSScoreCard.CLCQ.CLEXP_MLTIPRDTCOIData
GROUP BY FORMAT(ClickDateTime, 'MM-yy');
"""
# Top 100 count(INFOID) as dis_INFOID,  distinct FORMAT(ClickDateTime, 'MM-yy') as date where ClickDateTime > '02-27-23'
# execute the statement
COI_df = pd.read_sql(COI_Table, COI_connect)
# print(COI_df)

# close connection
COI_connect.close()


# Bringing in Self-Service
print("Self-Service Information")

SelfService_connect = p.connect("DSN=MSS-P1-PCA-06")

SelfService_Table = """
Select A.eventmmyy,
a.pagename,
--a.pagetype,
a.applicantgroups,
A.PRODUCT_GROUP,
SUM(a.pagenamecount) as pagehits FROM
(SELECT
	--CONVERT(date,DateAdd(hour,-5,[eventTimestamp])) as eventdate,
	[EventDate] as eventmmyy,
	CASE WHEN [pagetype] = 'BASE_PAGE' THEN concat([pagename],'@') ELSE [pagename] end AS [pagename],
	--[pagetype],

	CASE WHEN [userdata userroletxt] LIKE '%Agent%' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata AgentCode] <> 'null' then 'Agent'
               WHEN [userdata userroletxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata AgentCode] = 'null' AND [pagename] LIKE '%foragentsonly%' then 'Agent'
               WHEN [userdata UserRoleTxt] <> 'Agent' and  [userdata PartyLogonId] IN ('holder', 'proxy') AND [geocarrier] <> 'Progressive Casualty Insurance Companies' then 'Customer'
               WHEN [userdata userroletxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' AND [userdata PartyLogonId] = 'null' AND [pagename] LIKE '%progressivecommercial%' then 'Customer'
               WHEN  [userdata UserRoleTxt] = 'null' AND [geocarrier] <>  'Progressive Casualty Insurance Companies' then 'Unknown'
               WHEN [geocarrier] = 'Progressive Casualty Insurance Companies' then 'Internal'
              ELSE  'Internal' END AS applicantgroups,

	--CASE WHEN [pagename] IN ('businessownerinfo', 'enterrequestdetails', 'reissue', 'cancel', 'reinstate', 'nonrenew') THEN 'BEGINNING TRANS'
	--WHEN [pagename] IN ('quotesoldconfirm','endorseconfirm', 'ReinstateConfirm', 'NonRenewRBS', 'makepayment') THEN 'ENDPOINTS'
	--ELSE 'NA' END AS TRANS_TYPES,

	CASE WHEN  [userdata ProductCode] = 'CV' then 'Quoting'
	WHEN [userdata ProductCode] = 'CA' then 'Servicing'
	ELSE 'NA' END AS PRODUCT_GROUP,

	count([pagename]) as pagenamecount
	FROM [CLExperience].[dbo].[APPDCSV]
	WHERE
  --CONVERT(datetime,DateAdd(hour,-5,[eventTimestamp])) >  '2022-11-10 00:00:00.000'
  -- Oct 3rd-Oct 29th VS Nov 3rd-Dec 1st
	Eventdate between '2022-12-31' and '2023-02-03'
	--Eventdate between '2022-11-03' and '2022-12-01'
	AND [pagename] IN ('EndorseRBS', 'VehicleGaragingInformation', 'VehicleGaragingInformation-Reload')
	AND [userdata ProductCode] = 'CA'
	Group By
	--CONVERT(date,DateAdd(hour,-5,[eventTimestamp])),
	[Eventdate],
	[pagename],
	[pagetype],
	[userdata userroletxt], [geocarrier], [userdata AgentCode], [pagename], [userdata PartyLogonId], [userdata ProductCode]) A

Group by A.eventmmyy,
a.pagename,
--a.pagetype,
a.applicantgroups,

A.PRODUCT_GROUP
Order By A.eventmmyy
"""


SelfService_df = pd.read_sql(SelfService_Table, SelfService_connect)
# print(SelfService_df)

# join dimdate table in FSScoreCard to get Prog Month for Tableau

ProgMonths_connect = p.connect("DSN=MSS-P1-PCA-06")

ProgMonths_Table = """
SELECT DT_VAL, ACCT_CCYYMM
FROM [FSScoreCard].[CLCQ].[DimDate]
  where ACCT_CCYY >= 2022
"""

ProgMonths_df = pd.read_sql(ProgMonths_Table, ProgMonths_connect)
# print(ProgMonths_df)

# join the two tables above to get the Prog Months Correctly
print("")
print("Prog Self Service With Prog Month")
ProgSelfService_df = pd.merge(
    ProgMonths_df, SelfService_df, how="inner", left_on="DT_VAL", right_on="eventmmyy"
)

print(
    ProgSelfService_df.groupby(
        [
            "ACCT_CCYYMM",
            "pagename",
            "applicantgroups",
            "PRODUCT_GROUP",
        ]
    )["pagehits"].sum()
)

# Prog_Self = """
# SELECT ACCT_CCYYMM, PAGENAME, APPLICANTGROUPS, PRODUCT_GROUP, SUM(pagehits) as PageHiters
# FROM ProgSelfService_df
# Group by ACCT_CCYYMM, PAGENAME, APPLICANTGROUPS, PRODUCT_GROUP
# """

# print(ProgSelfService_df)
print("")
# print(Prog_Self)
# close connection
SelfService_connect.close()
ProgMonths_connect.close()

print("")
# Bringing in Session Time Reduced
print("Session Time Reduced Information")
